import { mkdtemp, rm } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, describe, expect, it } from 'vitest'
import { RecoveryStore } from '../src/recovery-store'

const temporaryDirectories: string[] = []

afterEach(async () => {
  await Promise.all(
    temporaryDirectories
      .splice(0)
      .map((directory) => rm(directory, { recursive: true, force: true })),
  )
})

describe('RecoveryStore', () => {
  it('keeps recovery snapshots isolated by live editing session', async () => {
    const directory = await mkdtemp(join(tmpdir(), 'tandemfolio-recovery-'))
    temporaryDirectories.push(directory)
    const store = new RecoveryStore(directory)

    await store.save({
      sessionId: 'session-a',
      format: 'markdown',
      fileName: 'a.md',
      data: Buffer.from('session a'),
    })
    await store.save({
      sessionId: 'session-b',
      format: 'markdown',
      fileName: 'b.md',
      data: Buffer.from('session b'),
    })

    await expect(store.latest('markdown', 'session-a')).resolves.toMatchObject({
      sessionId: 'session-a',
      fileName: 'a.md',
      data: Buffer.from('session a'),
    })
    await expect(store.latest('markdown', 'session-b')).resolves.toMatchObject({
      sessionId: 'session-b',
      fileName: 'b.md',
      data: Buffer.from('session b'),
    })
  })

  it('loads the latest renderer snapshot after the broker is recreated', async () => {
    const directory = await mkdtemp(join(tmpdir(), 'tandemfolio-recovery-'))
    temporaryDirectories.push(directory)
    const data = Buffer.from('PK\u0003\u0004durable-docx')

    await new RecoveryStore(directory).save({
      sessionId: 'durable-session',
      format: 'docx',
      fileName: 'draft.docx',
      data,
    })

    await expect(new RecoveryStore(directory).latest('docx')).resolves.toMatchObject({
      sessionId: 'durable-session',
      format: 'docx',
      fileName: 'draft.docx',
      data,
    })
  })

  it('expires stale snapshots and removes both payload and metadata', async () => {
    const directory = await mkdtemp(join(tmpdir(), 'tandemfolio-recovery-'))
    temporaryDirectories.push(directory)
    let now = Date.parse('2026-08-01T00:00:00.000Z')
    const store = new RecoveryStore(directory, { now: () => now, retentionMs: 7 * 86_400_000 })
    await store.save({
      sessionId: 'stale-session',
      format: 'xlsx',
      fileName: 'draft.xlsx',
      data: Buffer.from('xlsx'),
    })
    now += 8 * 86_400_000

    await expect(store.latest('xlsx')).resolves.toBeNull()
    await expect(store.latest('xlsx')).resolves.toBeNull()
  })

  it('rejects snapshots larger than the bounded recovery allowance', async () => {
    const directory = await mkdtemp(join(tmpdir(), 'tandemfolio-recovery-'))
    temporaryDirectories.push(directory)
    const store = new RecoveryStore(directory, { maxBytes: 3 })
    await expect(
      store.save({
        sessionId: 'large-session',
        format: 'pdf',
        fileName: 'large.pdf',
        data: Buffer.from('four'),
      }),
    ).rejects.toMatchObject({ code: 'invalid_arguments' })
  })
})
