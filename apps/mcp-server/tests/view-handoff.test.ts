import { mkdtemp, rm } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { Client } from '@modelcontextprotocol/sdk/client/index.js'
import { StdioClientTransport } from '@modelcontextprotocol/sdk/client/stdio.js'
import { afterEach, expect, it } from 'vitest'

const cleanups: (() => Promise<void>)[] = []
afterEach(async () => {
  for (const cleanup of cleanups.splice(0)) await cleanup()
})

async function editorSession() {
  const directory = await mkdtemp(join(tmpdir(), 'tandemfolio-handoff-'))
  const client = new Client({ name: 'handoff-regression', version: '1' })
  cleanups.push(async () => {
    await client.close()
    await rm(directory, { recursive: true, force: true })
  })
  await client.connect(
    new StdioClientTransport({
      command: process.execPath,
      args: [
        fileURLToPath(new URL('../../../node_modules/tsx/dist/cli.mjs', import.meta.url)),
        fileURLToPath(new URL('../src/server.ts', import.meta.url)),
      ],
      env: {
        ...process.env,
        TANDEMFOLIO_STATE_DIR: directory,
        TANDEMFOLIO_OUTPUT_DIR: join(directory, 'outputs'),
      },
    }),
  )
  const call = async (name: string, args: Record<string, unknown>) => {
    const response = await client.callTool({ name, arguments: args })
    return response as { isError?: boolean; structuredContent: Record<string, unknown> }
  }
  const created = await call('office_create_session', { format: 'markdown' })
  const sessionId = (created.structuredContent as { session: { id: string } }).session.id
  const shown = await call('office_show_markdown_editor', { sessionId })
  const viewId = shown.structuredContent!.viewId
  const owner = { sessionId, viewId, mountId: 'owner' }
  const candidate = { sessionId, viewId, mountId: 'candidate' }
  await call('office_editor_poll', owner)
  return { call, owner, candidate }
}

it('routes explicit continuation to the chosen replica instead of an older automatic waiter', async () => {
  const { call, owner, candidate } = await editorSession()
  await call('office_editor_poll', { ...candidate, active: true })
  const oldRequest = (await call('office_editor_poll', owner)).structuredContent.handoffRequest
  const chosen = { ...candidate, mountId: 'chosen-replica', active: true, activateView: true }
  await call('office_editor_poll', chosen)
  const requested = (await call('office_editor_poll', owner)).structuredContent
  expect(requested.handoffRequestedByUser).toBe(true)
  expect(requested.handoffRequest).not.toBe(oldRequest)
  // An automatic waiter cannot redirect the user's selection.
  await call('office_editor_poll', { ...candidate, active: true })
  expect((await call('office_editor_poll', owner)).structuredContent.handoffRequest).toBe(
    requested.handoffRequest,
  )
})

it('does not accept a new document save while the owner is checkpointing a view handoff', async () => {
  const { call, owner, candidate } = await editorSession()
  await call('office_editor_poll', { ...candidate, active: true })
  const polled = await call('office_editor_poll', owner)
  const handoffId = polled.structuredContent!.handoffRequest
  expect(handoffId).toEqual(expect.any(String))
  const prepared = await call('office_editor_handoff', { ...owner, handoffId, action: 'prepare' })
  expect(prepared.isError).not.toBe(true)
  const save = await call('office_editor_begin_document_save', {
    ...owner,
    fileName: 'unsafe.md',
    size: 1,
    mode: 'save',
  })
  expect(save.structuredContent).toMatchObject({ ok: false, error: 'command_in_flight' })
})

it('does not retarget an already prepared handoff to another explicit candidate', async () => {
  const { call, owner, candidate } = await editorSession()
  await call('office_editor_poll', { ...candidate, active: true, activateView: true })
  const request = (await call('office_editor_poll', owner)).structuredContent.handoffRequest
  await call('office_editor_handoff', { ...owner, handoffId: request, action: 'prepare' })
  await call('office_editor_poll', {
    ...candidate,
    mountId: 'another-choice',
    active: true,
    activateView: true,
  })
  expect((await call('office_editor_poll', owner)).structuredContent.handoffRequest).toBe(request)
  expect(
    (await call('office_editor_handoff', { ...owner, handoffId: request, action: 'abort' }))
      .isError,
  ).not.toBe(true)
})

it.each(['document_save', 'recovery'])(
  'keeps the owner when a %s upload is still in progress before handoff',
  async (uploadKind) => {
    const { call, owner, candidate } = await editorSession()
    await call(`office_editor_begin_${uploadKind}`, {
      ...owner,
      fileName: 'pending.md',
      size: 1,
      ...(uploadKind === 'document_save' ? { mode: 'save' } : {}),
    })
    await call('office_editor_poll', { ...candidate, active: true })
    const polled = await call('office_editor_poll', owner)
    const prepared = await call('office_editor_handoff', {
      ...owner,
      handoffId: polled.structuredContent!.handoffRequest,
      action: 'prepare',
    })
    expect(prepared.structuredContent).toMatchObject({ ok: false, error: 'command_in_flight' })
  },
)

it('blocks Agent mutation between committed handoff and the returning renderer restore', async () => {
  const { call, owner, candidate } = await editorSession()
  await call('office_editor_poll', { ...candidate, active: true })
  const polled = await call('office_editor_poll', owner)
  const handoffId = polled.structuredContent!.handoffRequest
  await call('office_editor_handoff', { ...owner, handoffId, action: 'prepare' })
  const bytes = Buffer.from('# Unsaved handoff content')
  const begun = await call('office_editor_begin_recovery', {
    ...owner,
    fileName: 'Retained.md',
    size: bytes.length,
  })
  const uploadId = begun.structuredContent!.uploadId
  await call('office_editor_write_recovery_chunk', {
    ...owner,
    uploadId,
    offset: 0,
    data: bytes.toString('base64'),
  })
  await call('office_editor_commit_recovery', { ...owner, uploadId })
  expect(
    (await call('office_editor_handoff', { ...owner, handoffId, action: 'commit' })).isError,
  ).not.toBe(true)
  const context = await call('office_get_context', { sessionId: owner.sessionId })
  expect(context.structuredContent!.session).toMatchObject({ connected: false })
  const mutation = await call('office_execute', {
    sessionId: owner.sessionId,
    baseRevision: 0,
    requestId: 'handoff-gap',
    operations: [{ id: 'markdown.history.undo', arguments: {} }],
  })
  expect(mutation.structuredContent).toMatchObject({ ok: false, error: 'editor_offline' })
  const resumed = await call('office_editor_poll', { ...candidate, active: true })
  const payload = resumed.structuredContent as {
    restoreCommandId: string
    commands: { commandId: string; arguments: { blobId: string } }[]
  }
  expect(payload.commands[0].commandId).toBe(payload.restoreCommandId)
  const chunk = await call('office_editor_read_file_chunk', {
    ...candidate,
    blobId: payload.commands[0].arguments.blobId,
    offset: 0,
    length: bytes.length,
  })
  expect(chunk.structuredContent!.data).toBe(bytes.toString('base64'))
  expect(
    (await call('office_editor_begin_document_save', { ...owner, fileName: 'stale.md', size: 1 }))
      .structuredContent,
  ).toMatchObject({ ok: false, error: 'editor_view_conflict' })
})

it('cannot commit a handoff from an older checkpoint or a different Session/view', async () => {
  const { call, owner, candidate } = await editorSession()
  const begun = await call('office_editor_begin_recovery', {
    ...owner,
    fileName: 'older.md',
    size: 0,
  })
  await call('office_editor_commit_recovery', {
    ...owner,
    uploadId: begun.structuredContent!.uploadId,
  })
  const other = await call('office_editor_poll', {
    ...candidate,
    viewId: 'another-view',
    active: true,
    activateView: true,
  })
  expect(other.structuredContent).toMatchObject({ ok: false, error: 'editor_view_conflict' })
  expect(other.structuredContent!.retryAfterMs).toBeUndefined()
  await call('office_editor_poll', { ...candidate, active: true })
  const polled = await call('office_editor_poll', owner)
  const handoffId = polled.structuredContent!.handoffRequest
  const created = await call('office_create_session', { format: 'markdown' })
  const anotherId = (created.structuredContent!.session as { id: string }).id
  const crossSession = await call('office_editor_handoff', {
    ...owner,
    sessionId: anotherId,
    handoffId,
    action: 'prepare',
  })
  expect(crossSession.structuredContent).toMatchObject({ ok: false, error: 'editor_view_conflict' })
  await call('office_editor_handoff', { ...owner, handoffId, action: 'prepare' })
  const commit = await call('office_editor_handoff', { ...owner, handoffId, action: 'commit' })
  expect(commit.structuredContent).toMatchObject({ ok: false, error: 'document_restore_failed' })
  expect(
    (await call('office_editor_poll', { ...candidate, active: true })).structuredContent,
  ).toMatchObject({ ok: false, error: 'editor_view_conflict' })
  expect(
    (await call('office_editor_handoff', { ...owner, handoffId, action: 'abort' })).isError,
  ).not.toBe(true)
  expect((await call('office_editor_poll', owner)).isError).not.toBe(true)
})

it('does not hand off while an Agent command awaits renderer acknowledgement', async () => {
  const { call, owner, candidate } = await editorSession()
  const editing = call('office_execute', {
    sessionId: owner.sessionId,
    baseRevision: 0,
    requestId: 'before-handoff',
    operations: [{ id: 'markdown.history.undo', arguments: {} }],
  })
  const delivery = await call('office_editor_poll', { ...owner, waitMs: 1000 })
  const [command] = delivery.structuredContent!.commands as { commandId: string }[]
  expect(command).toBeDefined()
  await call('office_editor_poll', { ...candidate, active: true })
  const polled = await call('office_editor_poll', owner)
  const prepared = await call('office_editor_handoff', {
    ...owner,
    handoffId: polled.structuredContent!.handoffRequest,
    action: 'prepare',
  })
  expect(prepared.structuredContent).toMatchObject({ ok: false, error: 'command_in_flight' })
  await call('office_editor_acknowledge', {
    ...owner,
    commandId: command.commandId,
    ok: true,
    revision: 1,
  })
  expect((await editing).isError).not.toBe(true)
})
