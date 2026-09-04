import { describe, expect, it } from 'vitest'
import { SessionError, SessionStore } from '../src/session-store'

describe('SessionStore', () => {
  it('keeps a lease alive while its owner transfers document bytes between polls', () => {
    let now = 0
    const store = new SessionStore(() => now)
    const session = store.create('pdf')
    store.poll(session.id, undefined, 'owner')
    now = 29_000
    store.assertView(session.id, 'owner')
    now = 31_000
    expect(() => store.poll(session.id, undefined, 'new-mount')).toThrow(/already connected/)
  })

  it('allows a cold mount to reclaim only an expired idle lease, never a live or in-flight one', () => {
    let now = 0
    const store = new SessionStore(() => now)
    const session = store.create('pptx')
    store.poll(session.id, undefined, 'view/mount-1')
    expect(() => store.poll(session.id, undefined, 'view/mount-2')).toThrow(/already connected/)
    now = 31_000
    store.poll(session.id, undefined, 'view/mount-2')
    expect(() => store.assertView(session.id, 'view/mount-1')).toThrow(/does not own/)
    store.enqueue(session.id, 0, 'pptx.slide.add', {})
    now += 31_000
    expect(() => store.poll(session.id, undefined, 'view/mount-3')).toThrow(/already connected/)
  })

  it('requires the shared editor to be connected before queuing a mutation', () => {
    const store = new SessionStore()
    const session = store.create()
    expect(() => store.enqueue(session.id, 0, 'insert_text', { text: 'hello' })).toThrowError(
      new SessionError('editor_offline', 'Open the TandemFolio editor before editing.'),
    )
  })

  it('queues a command against one revision and advances only after editor acknowledgement', () => {
    const store = new SessionStore()
    const session = store.create()
    store.connect(session.id, { fileName: 'demo.docx' })
    const command = store.enqueue(session.id, 0, 'insert_text', { text: 'hello' })
    expect(store.poll(session.id)).toEqual([command])
    expect(store.get(session.id).revision).toBe(0)
    store.acknowledge(session.id, command.commandId, 1, { dirty: true })
    expect(store.get(session.id)).toMatchObject({ revision: 1, dirty: true })
  })

  it('ends a bounded waiting poll with no commands after its timeout', async () => {
    const store = new SessionStore()
    const session = store.create()

    await expect(store.waitForPoll(session.id, undefined, 1)).resolves.toEqual([])
    expect(store.get(session.id).connected).toBe(true)
  })

  it('releases a waiting poll when its editor disconnects', async () => {
    const store = new SessionStore()
    const session = store.create()
    const waitingPoll = store.waitForPoll(session.id, undefined, 1_000)

    store.disconnect(session.id)

    await expect(waitingPoll).resolves.toEqual([])
    expect(store.get(session.id).connected).toBe(false)
  })

  it('supersedes a lost waiting poll and wakes the newest waiter on enqueue', async () => {
    const store = new SessionStore()
    const session = store.create()
    const lostPoll = store.waitForPoll(session.id, undefined, 1_000)
    const currentPoll = store.waitForPoll(session.id, undefined, 1_000)

    await expect(lostPoll).resolves.toEqual([])
    const command = store.enqueue(session.id, 0, 'insert_text', { text: 'wake now' })
    await expect(currentPoll).resolves.toEqual([command])
  })

  it('rejects a duplicate view without interrupting the owner waiting poll', async () => {
    const store = new SessionStore()
    const session = store.create()
    store.poll(session.id, undefined, 'owner-view')
    const ownerPoll = store.waitForPoll(session.id, undefined, 1_000, 'owner-view')

    expect(() => store.waitForPoll(session.id, undefined, 1_000, 'duplicate-view')).toThrowError(
      /already connected to another mounted editor view/,
    )
    const command = store.enqueue(session.id, 0, 'insert_text', { text: 'owner' })

    await expect(ownerPoll).resolves.toEqual([command])
  })

  it('rejects a second mutation until the first command is acknowledged', () => {
    const store = new SessionStore()
    const session = store.create()
    store.connect(session.id)
    store.enqueue(session.id, 0, 'insert_text', { text: 'first' })

    expect(() => store.enqueue(session.id, 0, 'insert_text', { text: 'second' })).toThrowError(
      new SessionError(
        'command_in_flight',
        'Wait for the active editor command to finish before submitting another mutation.',
      ),
    )
  })

  it('keeps the command in flight after the editor polls it', () => {
    const store = new SessionStore()
    const session = store.create()
    store.connect(session.id)
    store.enqueue(session.id, 0, 'insert_text', { text: 'first' })
    store.poll(session.id)

    expect(() => store.enqueue(session.id, 0, 'insert_text', { text: 'second' })).toThrowError(
      /Wait for the active editor command/,
    )
  })

  it('rejects an acknowledgement for a command that is not active', () => {
    const store = new SessionStore()
    const session = store.create()
    store.connect(session.id)
    store.enqueue(session.id, 0, 'insert_text', { text: 'first' })
    store.poll(session.id)

    expect(() => store.acknowledge(session.id, 'wrong-command', 1, {})).toThrowError(
      new SessionError(
        'command_not_found',
        'Command wrong-command is not the active command for this session.',
      ),
    )
    expect(() => store.enqueue(session.id, 0, 'insert_text', { text: 'second' })).toThrow(
      /active editor command/,
    )
  })

  it('resolves a command only after the editor acknowledges its next revision', async () => {
    const store = new SessionStore()
    const session = store.create()
    store.connect(session.id)
    const command = store.enqueue(session.id, 0, 'insert_text', { text: 'hello' })
    const completed = store.waitForCommand(session.id, command.commandId, 1_000)

    store.poll(session.id)
    store.acknowledge(session.id, command.commandId, 1, { dirty: true })

    await expect(completed).resolves.toEqual({
      commandId: command.commandId,
      ok: true,
      revision: 1,
    })
  })

  it('rejects commands created from stale context', () => {
    const store = new SessionStore()
    const session = store.create()
    store.connect(session.id)
    const command = store.enqueue(session.id, 0, 'insert_text', { text: 'first' })
    store.poll(session.id)
    store.acknowledge(session.id, command.commandId, 1, {})
    expect(() => store.enqueue(session.id, 0, 'insert_text', {})).toThrow(/Expected revision 1/)
  })
})
