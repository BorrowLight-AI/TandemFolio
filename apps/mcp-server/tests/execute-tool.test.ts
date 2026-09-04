import { mkdir, mkdtemp, readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { Client } from '@modelcontextprotocol/sdk/client/index.js'
import { StdioClientTransport } from '@modelcontextprotocol/sdk/client/stdio.js'
import { afterEach, describe, expect, it } from 'vitest'

const clients: Client[] = []
const temporaryDirectories: string[] = []

afterEach(async () => {
  await Promise.all(clients.splice(0).map((client) => client.close()))
  await Promise.all(
    temporaryDirectories
      .splice(0)
      .map((directory) => rm(directory, { recursive: true, force: true })),
  )
})

async function connectClient(
  stateDirectory?: string,
  environmentOverrides: Record<string, string> = {},
): Promise<Client> {
  if (!stateDirectory) {
    stateDirectory = await mkdtemp(join(tmpdir(), 'tandemfolio-test-state-'))
    temporaryDirectories.push(stateDirectory)
  }
  const client = new Client({ name: 'tandemfolio-test', version: '0.1.0' })
  const environment = Object.fromEntries(
    Object.entries(process.env).filter(
      (entry): entry is [string, string] => entry[1] !== undefined,
    ),
  )
  if (stateDirectory) environment.TANDEMFOLIO_STATE_DIR = stateDirectory
  Object.assign(environment, environmentOverrides)
  const transport = new StdioClientTransport({
    command: process.execPath,
    args: [
      fileURLToPath(new URL('../../../node_modules/tsx/dist/cli.mjs', import.meta.url)),
      fileURLToPath(new URL('../src/server.ts', import.meta.url)),
    ],
    env: environment,
  })
  await client.connect(transport)
  clients.push(client)
  return client
}

async function readAllCapabilitySchemas(
  client: Client,
  format: 'docx' | 'markdown' | 'xlsx' | 'pptx' | 'pdf',
): Promise<{
  structuredContent: {
    ok: true
    capabilities: Record<string, unknown> & { operations: Record<string, unknown> }
  }
}> {
  let cursor: string | null = null
  let metadata: Record<string, unknown> | undefined
  const operations: Record<string, unknown> = {}

  do {
    const summary = await client.callTool({
      name: 'office_get_capabilities',
      arguments: { format, view: 'summary', limit: 20, cursor },
    })
    const capabilities = (
      summary.structuredContent as {
        capabilities: Record<string, unknown> & {
          discovery: {
            operations: Array<{ id: string }>
            pagination: { nextCursor: string | null }
          }
        }
      }
    ).capabilities
    if (!metadata) {
      const { discovery: _discovery, ...base } = capabilities
      metadata = base
    }
    const descriptors = await Promise.all(
      capabilities.discovery.operations.map(({ id }) =>
        client.callTool({
          name: 'office_get_capabilities',
          arguments: { format, view: 'detail', operation: id },
        }),
      ),
    )
    for (const response of descriptors) {
      const descriptor = (
        response.structuredContent as {
          capabilities: {
            discovery: { operation: { id: string; inputSchema: unknown } }
          }
        }
      ).capabilities.discovery.operation
      operations[descriptor.id] = descriptor.inputSchema
    }
    cursor = capabilities.discovery.pagination.nextCursor
  } while (cursor)

  return {
    structuredContent: {
      ok: true,
      capabilities: { ...metadata, operations },
    },
  }
}

let operationRequestSequence = 0

function executeOperation(
  client: Client,
  sessionId: string,
  baseRevision: number,
  operation: string,
  arguments_: Record<string, unknown>,
) {
  operationRequestSequence += 1
  return client.callTool({
    name: 'office_execute',
    arguments: {
      sessionId,
      baseRevision,
      requestId: `test-operation-${operationRequestSequence}`,
      operations: [{ id: operation, arguments: arguments_ }],
    },
  })
}

async function executeAndAcknowledgeDocxOperation(
  operation: string,
  arguments_: Record<string, unknown>,
  output: Record<string, unknown>,
): Promise<void> {
  const client = await connectClient()
  const created = await client.callTool({
    name: 'office_create_session',
    arguments: { format: 'docx' },
  })
  const sessionId = (created.structuredContent as { session: { id: string } }).session.id
  await client.callTool({
    name: 'office_editor_poll',
    arguments: { sessionId, selection: { from: 1, to: 1 } },
  })

  const execution = executeOperation(client, sessionId, 0, operation, arguments_)

  let command:
    | {
        commandId: string
        operation: string
        arguments: Record<string, unknown>
      }
    | undefined
  for (let attempt = 0; attempt < 20 && !command; attempt += 1) {
    const polled = await client.callTool({
      name: 'office_editor_poll',
      arguments: { sessionId },
    })
    command = (
      polled.structuredContent as {
        commands: Array<{
          commandId: string
          operation: string
          arguments: Record<string, unknown>
        }>
      }
    ).commands[0]
    if (!command) await new Promise((resolve) => setTimeout(resolve, 10))
  }
  expect(command).toMatchObject({ operation, arguments: arguments_ })

  await client.callTool({
    name: 'office_editor_acknowledge',
    arguments: {
      sessionId,
      commandId: command!.commandId,
      revision: 1,
      dirty: true,
      output,
    },
  })
  await expect(execution).resolves.toMatchObject({
    structuredContent: {
      ok: true,
      result: { revision: 1, operations: [{ id: operation, result: output }] },
    },
  })
}

async function executeAndAcknowledgeMarkdownOperation(
  operation: string,
  arguments_: Record<string, unknown>,
  output: Record<string, unknown>,
): Promise<void> {
  const client = await connectClient()
  const created = await client.callTool({
    name: 'office_create_session',
    arguments: { format: 'markdown' },
  })
  const sessionId = (created.structuredContent as { session: { id: string } }).session.id
  await client.callTool({
    name: 'office_editor_poll',
    arguments: { sessionId, selection: { from: 1, to: 1 } },
  })

  const execution = executeOperation(client, sessionId, 0, operation, arguments_)
  let command:
    { commandId: string; operation: string; arguments: Record<string, unknown> } | undefined
  for (let attempt = 0; attempt < 20 && !command; attempt += 1) {
    const polled = await client.callTool({
      name: 'office_editor_poll',
      arguments: { sessionId },
    })
    command = (
      polled.structuredContent as {
        commands: Array<{
          commandId: string
          operation: string
          arguments: Record<string, unknown>
        }>
      }
    ).commands[0]
    if (!command) await new Promise((resolve) => setTimeout(resolve, 10))
  }
  expect(command).toMatchObject({ operation, arguments: arguments_ })

  await client.callTool({
    name: 'office_editor_acknowledge',
    arguments: {
      sessionId,
      commandId: command!.commandId,
      revision: 1,
      dirty: true,
      output,
    },
  })
  await expect(execution).resolves.toMatchObject({
    structuredContent: {
      ok: true,
      result: { revision: 1, operations: [{ id: operation, result: output }] },
    },
  })
}

async function expectMarkdownOperationRejectedBeforeEnqueue(
  operation: string,
  arguments_: Record<string, unknown>,
  message: string,
): Promise<void> {
  const client = await connectClient()
  const created = await client.callTool({
    name: 'office_create_session',
    arguments: { format: 'markdown' },
  })
  const sessionId = (created.structuredContent as { session: { id: string } }).session.id
  await client.callTool({
    name: 'office_editor_poll',
    arguments: { sessionId, selection: { from: 1, to: 1 } },
  })

  const execution = await executeOperation(client, sessionId, 0, operation, arguments_)
  expect(execution).toMatchObject({
    isError: true,
    structuredContent: { ok: false, error: 'operation_schema_invalid', message },
  })
  const polled = await client.callTool({
    name: 'office_editor_poll',
    arguments: { sessionId },
  })
  expect(
    (polled.structuredContent as { commands: Array<{ commandId: string }> }).commands[0],
  ).toBeUndefined()
  const context = await client.callTool({ name: 'office_get_context', arguments: { sessionId } })
  expect(context.structuredContent).toMatchObject({ session: { revision: 0 } })
}

async function expectDocxOperationRejectedBeforeEnqueue(
  operation: string,
  arguments_: Record<string, unknown>,
  message: string,
): Promise<void> {
  const client = await connectClient()
  const created = await client.callTool({
    name: 'office_create_session',
    arguments: { format: 'docx' },
  })
  const sessionId = (created.structuredContent as { session: { id: string } }).session.id
  await client.callTool({
    name: 'office_editor_poll',
    arguments: { sessionId, selection: { from: 1, to: 1 } },
  })

  const execution = await executeOperation(client, sessionId, 0, operation, arguments_)
  expect(execution).toMatchObject({
    isError: true,
    structuredContent: { ok: false, error: 'operation_schema_invalid', message },
  })

  const polled = await client.callTool({
    name: 'office_editor_poll',
    arguments: { sessionId },
  })
  expect(
    (polled.structuredContent as { commands: Array<{ commandId: string }> }).commands[0],
  ).toBeUndefined()

  const context = await client.callTool({ name: 'office_get_context', arguments: { sessionId } })
  expect(context.structuredContent).toMatchObject({ session: { revision: 0 } })
}

async function executeAndAcknowledgeXlsxOperation(
  operation: string,
  arguments_: Record<string, unknown>,
  output: Record<string, unknown>,
): Promise<void> {
  const client = await connectClient()
  const created = await client.callTool({
    name: 'office_create_session',
    arguments: { format: 'xlsx' },
  })
  const sessionId = (created.structuredContent as { session: { id: string } }).session.id
  await client.callTool({
    name: 'office_editor_poll',
    arguments: { sessionId, selection: { from: 1, to: 1 } },
  })

  const execution = executeOperation(client, sessionId, 0, operation, arguments_)

  let command:
    | {
        commandId: string
        operation: string
        arguments: Record<string, unknown>
      }
    | undefined
  for (let attempt = 0; attempt < 20 && !command; attempt += 1) {
    const polled = await client.callTool({
      name: 'office_editor_poll',
      arguments: { sessionId },
    })
    command = (
      polled.structuredContent as {
        commands: Array<{
          commandId: string
          operation: string
          arguments: Record<string, unknown>
        }>
      }
    ).commands[0]
    if (!command) await new Promise((resolve) => setTimeout(resolve, 10))
  }
  if (!command) {
    await execution
    expect(command).toBeDefined()
    return
  }
  await client.callTool({
    name: 'office_editor_acknowledge',
    arguments: {
      sessionId,
      commandId: command.commandId,
      revision: 1,
      dirty: true,
      output,
    },
  })
  expect(command).toMatchObject({ operation, arguments: arguments_ })
  await expect(execution).resolves.toMatchObject({
    structuredContent: {
      ok: true,
      result: { revision: 1, operations: [{ id: operation, result: output }] },
    },
  })
}

async function expectXlsxOperationRejectedBeforeEnqueue(
  operation: string,
  arguments_: Record<string, unknown>,
  message: string,
): Promise<void> {
  const client = await connectClient()
  const created = await client.callTool({
    name: 'office_create_session',
    arguments: { format: 'xlsx' },
  })
  const sessionId = (created.structuredContent as { session: { id: string } }).session.id
  await client.callTool({
    name: 'office_editor_poll',
    arguments: { sessionId, selection: { from: 1, to: 1 } },
  })

  const execution = await executeOperation(client, sessionId, 0, operation, arguments_)
  expect(execution).toMatchObject({
    isError: true,
    structuredContent: { ok: false, error: 'operation_schema_invalid', message },
  })

  const polled = await client.callTool({
    name: 'office_editor_poll',
    arguments: { sessionId },
  })
  expect(
    (polled.structuredContent as { commands: Array<{ commandId: string }> }).commands[0],
  ).toBeUndefined()

  const context = await client.callTool({ name: 'office_get_context', arguments: { sessionId } })
  expect(context.structuredContent).toMatchObject({ session: { revision: 0 } })
}

async function executeAndAcknowledgePdfOperation(
  operation: string,
  arguments_: Record<string, unknown>,
  output: Record<string, unknown>,
): Promise<void> {
  const client = await connectClient()
  const created = await client.callTool({
    name: 'office_create_session',
    arguments: { format: 'pdf' },
  })
  const sessionId = (created.structuredContent as { session: { id: string } }).session.id
  await client.callTool({
    name: 'office_editor_poll',
    arguments: { sessionId, selection: { from: 1, to: 1 } },
  })

  const execution = executeOperation(client, sessionId, 0, operation, arguments_)

  let command:
    | {
        commandId: string
        operation: string
        arguments: Record<string, unknown>
      }
    | undefined
  for (let attempt = 0; attempt < 20 && !command; attempt += 1) {
    const polled = await client.callTool({
      name: 'office_editor_poll',
      arguments: { sessionId },
    })
    command = (
      polled.structuredContent as {
        commands: Array<{
          commandId: string
          operation: string
          arguments: Record<string, unknown>
        }>
      }
    ).commands[0]
    if (!command) await new Promise((resolve) => setTimeout(resolve, 10))
  }
  expect(command).toMatchObject({ operation, arguments: arguments_ })

  await client.callTool({
    name: 'office_editor_acknowledge',
    arguments: {
      sessionId,
      commandId: command!.commandId,
      revision: 1,
      dirty: operation !== 'pdf.document.save',
      output,
    },
  })
  await expect(execution).resolves.toMatchObject({
    structuredContent: {
      ok: true,
      result: { revision: 1, operations: [{ id: operation, result: output }] },
    },
  })
}

async function expectPdfOperationRejectedBeforeEnqueue(
  operation: string,
  arguments_: Record<string, unknown>,
  message: string,
): Promise<void> {
  const client = await connectClient()
  const created = await client.callTool({
    name: 'office_create_session',
    arguments: { format: 'pdf' },
  })
  const sessionId = (created.structuredContent as { session: { id: string } }).session.id
  await client.callTool({
    name: 'office_editor_poll',
    arguments: { sessionId, selection: { from: 1, to: 1 } },
  })

  const execution = await executeOperation(client, sessionId, 0, operation, arguments_)
  expect(execution).toMatchObject({
    isError: true,
    structuredContent: { ok: false, error: 'operation_schema_invalid', message },
  })

  const polled = await client.callTool({
    name: 'office_editor_poll',
    arguments: { sessionId },
  })
  expect(
    (polled.structuredContent as { commands: Array<{ commandId: string }> }).commands[0],
  ).toBeUndefined()
  const context = await client.callTool({ name: 'office_get_context', arguments: { sessionId } })
  expect(context.structuredContent).toMatchObject({ session: { revision: 0 } })
}

describe('office_execute', () => {
  it.each([
    ['docx', 'office_show_editor'],
    ['markdown', 'office_show_markdown_editor'],
    ['xlsx', 'office_show_xlsx_editor'],
    ['pptx', 'office_show_pptx_editor'],
    ['pdf', 'office_show_pdf_editor'],
  ] as const)(
    'keeps one exclusive mounted %s editor view per session',
    async (format, showTool) => {
      const client = await connectClient()
      const created = await client.callTool({
        name: 'office_create_session',
        arguments: { format, resume: 'none' },
      })
      const sessionId = (created.structuredContent as { session: { id: string } }).session.id
      const firstShown = await client.callTool({ name: showTool, arguments: { sessionId } })
      const firstViewId = (firstShown.structuredContent as { viewId: string }).viewId

      expect(firstViewId).toEqual(expect.any(String))
      await expect(
        client.callTool({
          name: 'office_editor_poll',
          arguments: { sessionId, viewId: firstViewId },
        }),
      ).resolves.toMatchObject({ structuredContent: { ok: true } })

      const duplicateShown = await client.callTool({ name: showTool, arguments: { sessionId } })
      const duplicateViewId = (duplicateShown.structuredContent as { viewId: string }).viewId
      expect(duplicateViewId).not.toBe(firstViewId)

      const duplicatePoll = await client.callTool({
        name: 'office_editor_poll',
        arguments: { sessionId, viewId: duplicateViewId },
      })
      expect(duplicatePoll).toMatchObject({
        isError: true,
        structuredContent: { ok: false, error: 'editor_view_conflict' },
      })

      const duplicateDisconnect = await client.callTool({
        name: 'office_editor_disconnect',
        arguments: { sessionId, viewId: duplicateViewId },
      })
      expect(duplicateDisconnect).toMatchObject({
        isError: true,
        structuredContent: { ok: false, error: 'editor_view_conflict' },
      })
      await expect(
        client.callTool({ name: 'office_get_context', arguments: { sessionId } }),
      ).resolves.toMatchObject({ structuredContent: { session: { connected: true } } })

      await expect(
        client.callTool({
          name: 'office_editor_disconnect',
          arguments: { sessionId, viewId: firstViewId },
        }),
      ).resolves.toMatchObject({ structuredContent: { ok: true } })
      await expect(
        client.callTool({ name: 'office_get_context', arguments: { sessionId } }),
      ).resolves.toMatchObject({ structuredContent: { session: { connected: false } } })
    },
  )

  it('rejects acknowledgement, recovery, and document saves from a non-owner editor view', async () => {
    const client = await connectClient()
    const created = await client.callTool({
      name: 'office_create_session',
      arguments: { format: 'markdown', resume: 'none' },
    })
    const sessionId = (created.structuredContent as { session: { id: string } }).session.id
    const firstShown = await client.callTool({
      name: 'office_show_markdown_editor',
      arguments: { sessionId },
    })
    const firstViewId = (firstShown.structuredContent as { viewId: string }).viewId
    await client.callTool({
      name: 'office_editor_poll',
      arguments: { sessionId, viewId: firstViewId, selection: { from: 1, to: 1 } },
    })
    const duplicateShown = await client.callTool({
      name: 'office_show_markdown_editor',
      arguments: { sessionId },
    })
    const duplicateViewId = (duplicateShown.structuredContent as { viewId: string }).viewId

    const execution = executeOperation(client, sessionId, 0, 'markdown.text.insert', {
      text: 'owned',
    })
    const polled = await client.callTool({
      name: 'office_editor_poll',
      arguments: { sessionId, viewId: firstViewId, waitMs: 1_000 },
    })
    const command = (polled.structuredContent as { commands: Array<{ commandId: string }> })
      .commands[0]

    const duplicateAcknowledgement = await client.callTool({
      name: 'office_editor_acknowledge',
      arguments: {
        sessionId,
        viewId: duplicateViewId,
        commandId: command.commandId,
        revision: 1,
        dirty: true,
      },
    })
    expect(duplicateAcknowledgement).toMatchObject({
      isError: true,
      structuredContent: { ok: false, error: 'editor_view_conflict' },
    })
    await expect(
      client.callTool({ name: 'office_get_context', arguments: { sessionId } }),
    ).resolves.toMatchObject({ structuredContent: { session: { revision: 0 } } })

    const duplicateRecovery = await client.callTool({
      name: 'office_editor_begin_recovery',
      arguments: {
        sessionId,
        viewId: duplicateViewId,
        fileName: 'wrong.md',
        size: 1,
      },
    })
    expect(duplicateRecovery).toMatchObject({
      isError: true,
      structuredContent: { ok: false, error: 'editor_view_conflict' },
    })

    const duplicateSave = await client.callTool({
      name: 'office_editor_begin_document_save',
      arguments: {
        sessionId,
        viewId: duplicateViewId,
        fileName: 'wrong.md',
        size: 1,
        mode: 'save',
      },
    })
    expect(duplicateSave).toMatchObject({
      isError: true,
      structuredContent: { ok: false, error: 'editor_view_conflict' },
    })

    await client.callTool({
      name: 'office_editor_acknowledge',
      arguments: {
        sessionId,
        viewId: firstViewId,
        commandId: command.commandId,
        revision: 1,
        dirty: true,
      },
    })
    await expect(execution).resolves.toMatchObject({
      structuredContent: { ok: true, result: { revision: 1 } },
    })
  })

  it('executes one revision-guarded operation transaction and returns its identity', async () => {
    const client = await connectClient()
    const created = await client.callTool({
      name: 'office_create_session',
      arguments: { format: 'markdown' },
    })
    const sessionId = (created.structuredContent as { session: { id: string } }).session.id
    await client.callTool({
      name: 'office_editor_poll',
      arguments: { sessionId, selection: { from: 1, to: 1 } },
    })

    const execution = client.callTool({
      name: 'office_execute',
      arguments: {
        sessionId,
        baseRevision: 0,
        requestId: 'request-single-transaction',
        operations: [{ id: 'markdown.text.insert', arguments: { text: 'hello' } }],
      },
    })
    const polled = await client.callTool({
      name: 'office_editor_poll',
      arguments: { sessionId, waitMs: 1_000 },
    })
    const command = (
      polled.structuredContent as {
        commands: Array<{ commandId: string; operation: string; arguments: unknown }>
      }
    ).commands[0]

    expect(command).toMatchObject({
      operation: 'markdown.text.insert',
      arguments: { text: 'hello' },
    })
    await client.callTool({
      name: 'office_editor_acknowledge',
      arguments: {
        sessionId,
        commandId: command.commandId,
        revision: 1,
        dirty: true,
        output: { inserted: true },
      },
    })

    await expect(execution).resolves.toMatchObject({
      structuredContent: {
        ok: true,
        transaction: {
          transactionId: expect.any(String),
          requestId: 'request-single-transaction',
          baseRevision: 0,
        },
        result: {
          revision: 1,
          operations: [{ id: 'markdown.text.insert', result: { inserted: true } }],
        },
      },
    })
    const context = await client.callTool({
      name: 'office_get_context',
      arguments: { sessionId },
    })
    expect(context.structuredContent).toMatchObject({ session: { revision: 1 } })
  })

  it('replays a completed operation transaction without dispatching it again', async () => {
    const client = await connectClient()
    const created = await client.callTool({
      name: 'office_create_session',
      arguments: { format: 'markdown' },
    })
    const sessionId = (created.structuredContent as { session: { id: string } }).session.id
    await client.callTool({
      name: 'office_editor_poll',
      arguments: { sessionId, selection: { from: 1, to: 1 } },
    })
    const transaction = {
      sessionId,
      baseRevision: 0,
      requestId: 'request-completed-replay',
      operations: [{ id: 'markdown.text.insert', arguments: { text: 'once' } }],
    }

    const firstExecution = client.callTool({ name: 'office_execute', arguments: transaction })
    const polled = await client.callTool({
      name: 'office_editor_poll',
      arguments: { sessionId, waitMs: 1_000 },
    })
    const commandId = (polled.structuredContent as { commands: Array<{ commandId: string }> })
      .commands[0]!.commandId
    await client.callTool({
      name: 'office_editor_acknowledge',
      arguments: {
        sessionId,
        commandId,
        revision: 1,
        dirty: true,
        output: { inserted: true },
      },
    })
    const first = await firstExecution

    const replay = await client.callTool({ name: 'office_execute', arguments: transaction })
    expect(replay.structuredContent).toEqual(first.structuredContent)
    const afterReplay = await client.callTool({
      name: 'office_editor_poll',
      arguments: { sessionId, waitMs: 0 },
    })
    expect(afterReplay.structuredContent).toMatchObject({ commands: [] })
    const context = await client.callTool({
      name: 'office_get_context',
      arguments: { sessionId },
    })
    expect(context.structuredContent).toMatchObject({ session: { revision: 1 } })
  })

  it('joins an in-flight replay to the same operation transaction', async () => {
    const client = await connectClient()
    const created = await client.callTool({
      name: 'office_create_session',
      arguments: { format: 'markdown' },
    })
    const sessionId = (created.structuredContent as { session: { id: string } }).session.id
    await client.callTool({
      name: 'office_editor_poll',
      arguments: { sessionId, selection: { from: 1, to: 1 } },
    })
    const transaction = {
      sessionId,
      baseRevision: 0,
      requestId: 'request-in-flight-replay',
      operations: [{ id: 'markdown.text.insert', arguments: { text: 'once' } }],
    }

    const firstExecution = client.callTool({ name: 'office_execute', arguments: transaction })
    const polled = await client.callTool({
      name: 'office_editor_poll',
      arguments: { sessionId, waitMs: 1_000 },
    })
    const commandId = (polled.structuredContent as { commands: Array<{ commandId: string }> })
      .commands[0]!.commandId
    const replayExecution = client.callTool({ name: 'office_execute', arguments: transaction })
    await new Promise((resolve) => setTimeout(resolve, 25))
    const duringReplay = await client.callTool({
      name: 'office_editor_poll',
      arguments: { sessionId, waitMs: 0 },
    })
    expect(duringReplay.structuredContent).toMatchObject({ commands: [] })

    await client.callTool({
      name: 'office_editor_acknowledge',
      arguments: {
        sessionId,
        commandId,
        revision: 1,
        dirty: true,
        output: { inserted: true },
      },
    })
    const [first, replay] = await Promise.all([firstExecution, replayExecution])
    expect(replay.structuredContent).toEqual(first.structuredContent)
  })

  it('replays an acknowledged transaction after an earlier caller timed out', async () => {
    const client = await connectClient(undefined, {
      TANDEMFOLIO_COMMAND_TIMEOUT_MS: '25',
    })
    const created = await client.callTool({
      name: 'office_create_session',
      arguments: { format: 'markdown' },
    })
    const sessionId = (created.structuredContent as { session: { id: string } }).session.id
    await client.callTool({
      name: 'office_editor_poll',
      arguments: { sessionId, selection: { from: 1, to: 1 } },
    })
    const transaction = {
      sessionId,
      baseRevision: 0,
      requestId: 'request-replay-after-timeout',
      operations: [{ id: 'markdown.text.insert', arguments: { text: 'eventual' } }],
    }

    const firstExecution = client.callTool({ name: 'office_execute', arguments: transaction })
    const polled = await client.callTool({
      name: 'office_editor_poll',
      arguments: { sessionId, waitMs: 1_000 },
    })
    const commandId = (polled.structuredContent as { commands: Array<{ commandId: string }> })
      .commands[0]!.commandId
    await expect(firstExecution).resolves.toMatchObject({
      isError: true,
      structuredContent: { ok: false, error: 'command_timeout' },
    })

    await client.callTool({
      name: 'office_editor_acknowledge',
      arguments: {
        sessionId,
        commandId,
        revision: 1,
        dirty: true,
        output: { inserted: true },
      },
    })
    const replay = await client.callTool({ name: 'office_execute', arguments: transaction })
    expect(replay).toMatchObject({
      structuredContent: {
        ok: true,
        transaction: { requestId: 'request-replay-after-timeout' },
        result: {
          revision: 1,
          operations: [{ id: 'markdown.text.insert', result: { inserted: true } }],
        },
      },
    })
    const afterReplay = await client.callTool({
      name: 'office_editor_poll',
      arguments: { sessionId, waitMs: 0 },
    })
    expect(afterReplay.structuredContent).toMatchObject({ commands: [] })
  })

  it('keeps staged transaction bytes available until a timed-out command is acknowledged', async () => {
    const directory = await mkdtemp(join(tmpdir(), 'tandemfolio-transaction-image-'))
    temporaryDirectories.push(directory)
    const path = join(directory, 'logo.png')
    const bytes = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a])
    await writeFile(path, bytes)

    const client = await connectClient(undefined, {
      TANDEMFOLIO_COMMAND_TIMEOUT_MS: '25',
    })
    const created = await client.callTool({
      name: 'office_create_session',
      arguments: { format: 'markdown' },
    })
    const sessionId = (created.structuredContent as { session: { id: string } }).session.id
    await client.callTool({
      name: 'office_editor_poll',
      arguments: { sessionId, selection: { from: 1, to: 1 } },
    })
    const transaction = {
      sessionId,
      baseRevision: 0,
      requestId: 'request-staged-bytes-after-timeout',
      operations: [
        {
          id: 'markdown.image.insert',
          arguments: { path, position: 2, alt: 'Logo', title: null },
        },
      ],
    }

    const firstExecution = client.callTool({ name: 'office_execute', arguments: transaction })
    const polled = await client.callTool({
      name: 'office_editor_poll',
      arguments: { sessionId, waitMs: 1_000 },
    })
    const command = (
      polled.structuredContent as {
        commands: Array<{
          commandId: string
          arguments: { blobId: string }
        }>
      }
    ).commands[0]!
    await expect(firstExecution).resolves.toMatchObject({
      isError: true,
      structuredContent: { ok: false, error: 'command_timeout' },
    })

    const chunk = await client.callTool({
      name: 'office_editor_read_file_chunk',
      arguments: {
        sessionId,
        blobId: command.arguments.blobId,
        offset: 0,
        length: 4,
      },
    })
    expect(chunk.structuredContent).toMatchObject({
      ok: true,
      data: bytes.subarray(0, 4).toString('base64'),
      eof: false,
    })

    await client.callTool({
      name: 'office_editor_acknowledge',
      arguments: {
        sessionId,
        commandId: command.commandId,
        revision: 1,
        dirty: true,
        output: { inserted: true },
      },
    })
    const replay = await client.callTool({ name: 'office_execute', arguments: transaction })
    expect(replay).toMatchObject({
      structuredContent: {
        ok: true,
        transaction: { requestId: 'request-staged-bytes-after-timeout' },
        result: { revision: 1 },
      },
    })
  })

  it('rejects reuse of a transaction request id with a different payload', async () => {
    const client = await connectClient()
    const created = await client.callTool({
      name: 'office_create_session',
      arguments: { format: 'markdown' },
    })
    const sessionId = (created.structuredContent as { session: { id: string } }).session.id
    await client.callTool({
      name: 'office_editor_poll',
      arguments: { sessionId, selection: { from: 1, to: 1 } },
    })
    const firstExecution = client.callTool({
      name: 'office_execute',
      arguments: {
        sessionId,
        baseRevision: 0,
        requestId: 'request-reused-with-different-payload',
        operations: [{ id: 'markdown.text.insert', arguments: { text: 'first' } }],
      },
    })
    const polled = await client.callTool({
      name: 'office_editor_poll',
      arguments: { sessionId, waitMs: 1_000 },
    })
    const commandId = (polled.structuredContent as { commands: Array<{ commandId: string }> })
      .commands[0]!.commandId
    await client.callTool({
      name: 'office_editor_acknowledge',
      arguments: { sessionId, commandId, revision: 1, dirty: true },
    })
    await firstExecution

    const reused = await client.callTool({
      name: 'office_execute',
      arguments: {
        sessionId,
        baseRevision: 0,
        requestId: 'request-reused-with-different-payload',
        operations: [{ id: 'markdown.text.insert', arguments: { text: 'second' } }],
      },
    })
    expect(reused).toMatchObject({
      isError: true,
      structuredContent: { ok: false, error: 'request_reused' },
    })
    const afterReuse = await client.callTool({
      name: 'office_editor_poll',
      arguments: { sessionId, waitMs: 0 },
    })
    expect(afterReuse.structuredContent).toMatchObject({ commands: [] })
    const context = await client.callTool({
      name: 'office_get_context',
      arguments: { sessionId },
    })
    expect(context.structuredContent).toMatchObject({ session: { revision: 1 } })
  })

  it('treats JSON object key order as the same transaction payload', async () => {
    const client = await connectClient()
    const created = await client.callTool({
      name: 'office_create_session',
      arguments: { format: 'xlsx' },
    })
    const sessionId = (created.structuredContent as { session: { id: string } }).session.id
    await client.callTool({ name: 'office_editor_poll', arguments: { sessionId } })

    const firstExecution = client.callTool({
      name: 'office_execute',
      arguments: {
        sessionId,
        baseRevision: 0,
        requestId: 'request-stable-json-identity',
        operations: [
          {
            id: 'xlsx.range.set_values',
            arguments: { sheet: 'Sheet1', range: 'A1', values: [[1]] },
          },
        ],
      },
    })
    const polled = await client.callTool({
      name: 'office_editor_poll',
      arguments: { sessionId, waitMs: 1_000 },
    })
    const commandId = (polled.structuredContent as { commands: Array<{ commandId: string }> })
      .commands[0]!.commandId
    await client.callTool({
      name: 'office_editor_acknowledge',
      arguments: {
        sessionId,
        commandId,
        revision: 1,
        dirty: true,
        output: { updatedCells: 1 },
      },
    })
    const first = await firstExecution

    const replay = await client.callTool({
      name: 'office_execute',
      arguments: {
        sessionId,
        baseRevision: 0,
        requestId: 'request-stable-json-identity',
        operations: [
          {
            id: 'xlsx.range.set_values',
            arguments: { values: [[1]], range: 'A1', sheet: 'Sheet1' },
          },
        ],
      },
    })
    expect(replay.structuredContent).toEqual(first.structuredContent)
  })

  it('guards a new transaction by revision without reserving a rejected request id', async () => {
    const client = await connectClient()
    const created = await client.callTool({
      name: 'office_create_session',
      arguments: { format: 'markdown' },
    })
    const sessionId = (created.structuredContent as { session: { id: string } }).session.id
    await client.callTool({
      name: 'office_editor_poll',
      arguments: { sessionId, selection: { from: 1, to: 1 } },
    })

    const firstExecution = client.callTool({
      name: 'office_execute',
      arguments: {
        sessionId,
        baseRevision: 0,
        requestId: 'request-advance-revision',
        operations: [{ id: 'markdown.text.insert', arguments: { text: 'first' } }],
      },
    })
    const firstPoll = await client.callTool({
      name: 'office_editor_poll',
      arguments: { sessionId, waitMs: 1_000 },
    })
    const firstCommandId = (
      firstPoll.structuredContent as { commands: Array<{ commandId: string }> }
    ).commands[0]!.commandId
    await client.callTool({
      name: 'office_editor_acknowledge',
      arguments: { sessionId, commandId: firstCommandId, revision: 1, dirty: true },
    })
    await firstExecution

    const stale = await client.callTool({
      name: 'office_execute',
      arguments: {
        sessionId,
        baseRevision: 0,
        requestId: 'request-correctable-after-revision-conflict',
        operations: [{ id: 'markdown.text.insert', arguments: { text: 'second' } }],
      },
    })
    expect(stale).toMatchObject({
      isError: true,
      structuredContent: { ok: false, error: 'revision_conflict' },
    })

    const correctedExecution = client.callTool({
      name: 'office_execute',
      arguments: {
        sessionId,
        baseRevision: 1,
        requestId: 'request-correctable-after-revision-conflict',
        operations: [{ id: 'markdown.text.insert', arguments: { text: 'second' } }],
      },
    })
    const correctedPoll = await client.callTool({
      name: 'office_editor_poll',
      arguments: { sessionId, waitMs: 1_000 },
    })
    const correctedCommandId = (
      correctedPoll.structuredContent as { commands: Array<{ commandId: string }> }
    ).commands[0]!.commandId
    await client.callTool({
      name: 'office_editor_acknowledge',
      arguments: { sessionId, commandId: correctedCommandId, revision: 2, dirty: true },
    })
    await expect(correctedExecution).resolves.toMatchObject({
      structuredContent: {
        ok: true,
        transaction: { requestId: 'request-correctable-after-revision-conflict' },
        result: { revision: 2 },
      },
    })
  })

  it('rejects a multi-operation transaction before renderer dispatch', async () => {
    const client = await connectClient()
    const created = await client.callTool({
      name: 'office_create_session',
      arguments: { format: 'markdown' },
    })
    const sessionId = (created.structuredContent as { session: { id: string } }).session.id
    await client.callTool({ name: 'office_editor_poll', arguments: { sessionId } })

    const response = await client.callTool({
      name: 'office_execute',
      arguments: {
        sessionId,
        baseRevision: 0,
        requestId: 'request-non-atomic-group',
        operations: [
          { id: 'markdown.text.insert', arguments: { text: 'first' } },
          { id: 'markdown.document.save', arguments: {} },
        ],
      },
    })

    expect(response).toMatchObject({
      isError: true,
      structuredContent: { ok: false, error: 'transaction_not_atomic' },
    })
    const polled = await client.callTool({
      name: 'office_editor_poll',
      arguments: { sessionId, waitMs: 0 },
    })
    expect(polled.structuredContent).toMatchObject({ commands: [] })
  })

  it('rejects the retired operation and arguments envelope before renderer dispatch', async () => {
    const client = await connectClient()
    const created = await client.callTool({
      name: 'office_create_session',
      arguments: { format: 'markdown' },
    })
    const sessionId = (created.structuredContent as { session: { id: string } }).session.id
    await client.callTool({ name: 'office_editor_poll', arguments: { sessionId } })

    const execution = client.callTool({
      name: 'office_execute',
      arguments: {
        sessionId,
        baseRevision: 0,
        operation: 'markdown.document.save',
        arguments: {},
      },
    })
    const polled = await client.callTool({
      name: 'office_editor_poll',
      arguments: { sessionId, waitMs: 50 },
    })
    const command = (polled.structuredContent as { commands: Array<{ commandId: string }> })
      .commands[0]
    if (command) {
      await client.callTool({
        name: 'office_editor_acknowledge',
        arguments: { sessionId, commandId: command.commandId, revision: 1, dirty: false },
      })
    }
    const response = await execution

    expect(response).toMatchObject({ isError: true })
    expect(response.content).toEqual([
      expect.objectContaining({
        type: 'text',
        text: expect.stringContaining('Invalid arguments for tool office_execute'),
      }),
    ])
    expect(polled.structuredContent).toMatchObject({ commands: [] })
  })

  it('accepts only canonical Agent-visible operation ids in a transaction', async () => {
    const client = await connectClient()
    const created = await client.callTool({
      name: 'office_create_session',
      arguments: { format: 'markdown' },
    })
    const sessionId = (created.structuredContent as { session: { id: string } }).session.id
    await client.callTool({ name: 'office_editor_poll', arguments: { sessionId } })

    for (const [requestId, id] of [
      ['request-internal-operation', 'markdown.document.load_staged'],
      ['request-compatibility-alias', 'insert_text'],
    ]) {
      const response = await client.callTool({
        name: 'office_execute',
        arguments: {
          sessionId,
          baseRevision: 0,
          requestId,
          operations: [{ id, arguments: {} }],
        },
      })
      expect(response).toMatchObject({
        isError: true,
        structuredContent: { ok: false, error: 'operation_not_found' },
      })
    }
    const polled = await client.callTool({
      name: 'office_editor_poll',
      arguments: { sessionId, waitMs: 0 },
    })
    expect(polled.structuredContent).toMatchObject({ commands: [] })
  })

  it('rejects transaction arguments outside the exact operation schema', async () => {
    const client = await connectClient()
    const created = await client.callTool({
      name: 'office_create_session',
      arguments: { format: 'markdown' },
    })
    const sessionId = (created.structuredContent as { session: { id: string } }).session.id
    await client.callTool({ name: 'office_editor_poll', arguments: { sessionId } })

    const response = await client.callTool({
      name: 'office_execute',
      arguments: {
        sessionId,
        baseRevision: 0,
        requestId: 'request-schema-invalid',
        operations: [{ id: 'markdown.text.insert', arguments: {} }],
      },
    })

    expect(response).toMatchObject({
      isError: true,
      structuredContent: { ok: false, error: 'operation_schema_invalid' },
    })
    const polled = await client.callTool({
      name: 'office_editor_poll',
      arguments: { sessionId, waitMs: 0 },
    })
    expect(polled.structuredContent).toMatchObject({ commands: [] })
  })

  it('rejects a transaction whose required live selection is unavailable', async () => {
    const client = await connectClient()
    const created = await client.callTool({
      name: 'office_create_session',
      arguments: { format: 'markdown' },
    })
    const sessionId = (created.structuredContent as { session: { id: string } }).session.id
    await client.callTool({
      name: 'office_editor_poll',
      arguments: { sessionId, selection: null },
    })

    const execution = client.callTool({
      name: 'office_execute',
      arguments: {
        sessionId,
        baseRevision: 0,
        requestId: 'request-selection-unavailable',
        operations: [{ id: 'markdown.text.replace_selection', arguments: { text: 'replacement' } }],
      },
    })
    const polled = await client.callTool({
      name: 'office_editor_poll',
      arguments: { sessionId, waitMs: 50 },
    })
    const command = (polled.structuredContent as { commands: Array<{ commandId: string }> })
      .commands[0]
    if (command) {
      await client.callTool({
        name: 'office_editor_acknowledge',
        arguments: { sessionId, commandId: command.commandId, revision: 1, dirty: true },
      })
    }
    const response = await execution

    expect(response).toMatchObject({
      isError: true,
      structuredContent: { ok: false, error: 'operation_unavailable' },
    })
    expect(polled.structuredContent).toMatchObject({ commands: [] })
  })

  it('accepts only the exact bounded XLSX cold-start trace on the app-only poll', async () => {
    const client = await connectClient()
    const created = await client.callTool({
      name: 'office_create_session',
      arguments: { format: 'xlsx' },
    })
    const sessionId = (created.structuredContent as { session: { id: string } }).session.id
    const trace = {
      operation: 'xlsx.editor.cold_start',
      phases: {
        bootstrapMs: 600,
        univerCreateMs: 8,
        worksheetInstallMs: 14,
        firstCommitMs: 60,
      },
      bootstrapPhases: {
        resourceReceiveMs: 25,
        moduleGraphReadyMs: 550,
        reactMountMs: 25,
      },
    }

    const accepted = await client.callTool({
      name: 'office_editor_poll',
      arguments: { sessionId, startupTrace: trace },
    })
    expect(accepted.structuredContent).toMatchObject({ ok: true, commands: [] })

    const negative = await client.callTool({
      name: 'office_editor_poll',
      arguments: {
        sessionId,
        startupTrace: {
          ...trace,
          phases: { ...trace.phases, bootstrapMs: -1 },
        },
      },
    })
    expect(negative.isError).toBe(true)

    const extra = await client.callTool({
      name: 'office_editor_poll',
      arguments: {
        sessionId,
        startupTrace: { ...trace, documentData: 'must-not-cross-the-boundary' },
      },
    })
    expect(extra.isError).toBe(true)
  })

  it('serves allowlisted bundled font assets to the mounted editor in bounded chunks', async () => {
    const client = await connectClient()
    const response = await client.callTool({
      name: 'office_editor_read_font_chunk',
      arguments: {
        fileName: 'Carlito-Regular.ttf',
        offset: 0,
        length: 4,
      },
    })

    expect(response.structuredContent).toMatchObject({
      ok: true,
      fileName: 'Carlito-Regular.ttf',
      offset: 0,
      data: Buffer.from([0, 1, 0, 0]).toString('base64'),
      nextOffset: 4,
      eof: false,
    })

    const rejected = await client.callTool({
      name: 'office_editor_read_font_chunk',
      arguments: { fileName: '../server.ts', offset: 0, length: 4 },
    })
    expect(rejected).toMatchObject({
      isError: true,
      structuredContent: { ok: false, error: 'invalid_arguments' },
    })
  })

  it('returns a bounded schema-free capability summary by default', async () => {
    const client = await connectClient()
    const response = await client.callTool({
      name: 'office_get_capabilities',
      arguments: { format: 'xlsx' },
    })
    const capabilities = (
      response.structuredContent as {
        capabilities: {
          format: string
          operations?: unknown
          discovery: {
            view: string
            operations: Array<Record<string, unknown>>
            pagination: { limit: number; total: number; nextCursor: string | null }
          }
        }
      }
    ).capabilities

    expect(capabilities.format).toBe('xlsx')
    expect(capabilities.operations).toBeUndefined()
    expect(capabilities.discovery.view).toBe('summary')
    expect(capabilities.discovery.operations).toHaveLength(20)
    expect(capabilities.discovery.pagination).toEqual({
      limit: 20,
      total: 112,
      nextCursor: expect.any(String),
    })
    for (const operation of capabilities.discovery.operations) {
      expect(Object.keys(operation).sort()).toEqual([
        'context',
        'effects',
        'family',
        'id',
        'risk',
        'summary',
      ])
    }
  })

  it('returns one exact operation descriptor only when detail is requested', async () => {
    const client = await connectClient()
    const response = await client.callTool({
      name: 'office_get_capabilities',
      arguments: {
        format: 'xlsx',
        view: 'detail',
        operation: 'xlsx.range.set_values',
      },
    })

    expect(response.structuredContent).toMatchObject({
      ok: true,
      capabilities: {
        format: 'xlsx',
        discovery: {
          view: 'detail',
          operation: {
            id: 'xlsx.range.set_values',
            format: 'xlsx',
            family: 'range',
            visibility: 'agent',
            inputSchema: {
              type: 'object',
              required: ['sheet', 'range', 'values'],
              additionalProperties: false,
            },
            outputSchema: { type: 'object' },
          },
        },
      },
    })
    const discovery = (
      response.structuredContent as {
        capabilities: { discovery: Record<string, unknown> }
      }
    ).capabilities.discovery
    expect(discovery).not.toHaveProperty('operations')
    expect(discovery).not.toHaveProperty('pagination')
  })

  it('projects selection availability without hiding the requested operation', async () => {
    const client = await connectClient()
    const created = await client.callTool({
      name: 'office_create_session',
      arguments: { format: 'markdown' },
    })
    const sessionId = (created.structuredContent as { session: { id: string } }).session.id
    await client.callTool({
      name: 'office_editor_poll',
      arguments: { sessionId, selection: null },
    })

    const unavailable = await client.callTool({
      name: 'office_get_capabilities',
      arguments: {
        format: 'markdown',
        view: 'detail',
        operation: 'markdown.text.replace_selection',
        sessionId,
      },
    })
    expect(unavailable.structuredContent).toMatchObject({
      ok: true,
      capabilities: {
        discovery: {
          operation: {
            id: 'markdown.text.replace_selection',
            availability: { available: false, reason: 'selection_required' },
          },
        },
      },
    })

    await client.callTool({
      name: 'office_editor_poll',
      arguments: { sessionId, selection: { from: 1, to: 1 } },
    })
    const available = await client.callTool({
      name: 'office_get_capabilities',
      arguments: {
        format: 'markdown',
        view: 'detail',
        operation: 'markdown.text.replace_selection',
        sessionId,
      },
    })
    expect(available.structuredContent).toMatchObject({
      capabilities: {
        discovery: {
          operation: {
            id: 'markdown.text.replace_selection',
            availability: { available: true, reason: null },
          },
        },
      },
    })
  })

  it('reports an Agent operation unavailable while its editor is offline', async () => {
    const client = await connectClient()
    const created = await client.callTool({
      name: 'office_create_session',
      arguments: { format: 'xlsx' },
    })
    const sessionId = (created.structuredContent as { session: { id: string } }).session.id

    const response = await client.callTool({
      name: 'office_get_capabilities',
      arguments: {
        format: 'xlsx',
        view: 'detail',
        operation: 'xlsx.document.save',
        sessionId,
      },
    })

    expect(response.structuredContent).toMatchObject({
      ok: true,
      capabilities: {
        discovery: {
          operation: {
            id: 'xlsx.document.save',
            availability: { available: false, reason: 'editor_offline' },
          },
        },
      },
    })
  })

  it('reports a format mismatch for availability projected from another format session', async () => {
    const client = await connectClient()
    const created = await client.callTool({
      name: 'office_create_session',
      arguments: { format: 'markdown' },
    })
    const sessionId = (created.structuredContent as { session: { id: string } }).session.id
    await client.callTool({ name: 'office_editor_poll', arguments: { sessionId } })

    const response = await client.callTool({
      name: 'office_get_capabilities',
      arguments: {
        format: 'xlsx',
        view: 'detail',
        operation: 'xlsx.document.save',
        sessionId,
      },
    })

    expect(response.structuredContent).toMatchObject({
      ok: true,
      capabilities: {
        discovery: {
          operation: {
            id: 'xlsx.document.save',
            availability: { available: false, reason: 'format_mismatch' },
          },
        },
      },
    })
  })

  it('projects availability for every operation in a session-aware summary page', async () => {
    const client = await connectClient()
    const created = await client.callTool({
      name: 'office_create_session',
      arguments: { format: 'markdown' },
    })
    const sessionId = (created.structuredContent as { session: { id: string } }).session.id
    await client.callTool({
      name: 'office_editor_poll',
      arguments: { sessionId, selection: null },
    })

    const response = await client.callTool({
      name: 'office_get_capabilities',
      arguments: { format: 'markdown', sessionId },
    })
    const operations = (
      response.structuredContent as {
        capabilities: {
          discovery: {
            operations: Array<{
              id: string
              availability?: { available: boolean; reason: string | null }
            }>
          }
        }
      }
    ).capabilities.discovery.operations

    expect(operations.find(({ id }) => id === 'markdown.document.save')).toMatchObject({
      availability: { available: true, reason: null },
    })
    expect(operations.find(({ id }) => id === 'markdown.text.replace_selection')).toMatchObject({
      availability: { available: false, reason: 'selection_required' },
    })
    expect(operations.every(({ availability }) => availability !== undefined)).toBe(true)
  })

  it('filters capability summaries by family before applying the requested page limit', async () => {
    const client = await connectClient()
    const response = await client.callTool({
      name: 'office_get_capabilities',
      arguments: { format: 'xlsx', family: 'range', limit: 5 },
    })
    const discovery = (
      response.structuredContent as {
        capabilities: {
          discovery: {
            operations: Array<{ family: string }>
            pagination: { limit: number; total: number; nextCursor: string | null }
          }
        }
      }
    ).capabilities.discovery

    expect(discovery.operations).toHaveLength(5)
    expect(discovery.operations.every(({ family }) => family === 'range')).toBe(true)
    expect(discovery.pagination).toEqual({
      limit: 5,
      total: 36,
      nextCursor: expect.any(String),
    })
  })

  it('continues a filtered capability summary after its returned cursor', async () => {
    const client = await connectClient()
    const first = await client.callTool({
      name: 'office_get_capabilities',
      arguments: { format: 'xlsx', family: 'range', limit: 3 },
    })
    const firstDiscovery = (
      first.structuredContent as {
        capabilities: {
          discovery: {
            operations: Array<{ id: string }>
            pagination: { total: number; nextCursor: string }
          }
        }
      }
    ).capabilities.discovery
    const second = await client.callTool({
      name: 'office_get_capabilities',
      arguments: {
        format: 'xlsx',
        family: 'range',
        limit: 3,
        cursor: firstDiscovery.pagination.nextCursor,
      },
    })
    const secondDiscovery = (
      second.structuredContent as {
        capabilities: {
          discovery: {
            operations: Array<{ id: string }>
            pagination: { total: number; nextCursor: string | null }
          }
        }
      }
    ).capabilities.discovery

    expect(secondDiscovery.operations).toHaveLength(3)
    expect(secondDiscovery.pagination.total).toBe(36)
    expect(secondDiscovery.operations.map(({ id }) => id)).not.toEqual(
      firstDiscovery.operations.map(({ id }) => id),
    )
    expect(
      secondDiscovery.operations.some(({ id }) =>
        firstDiscovery.operations.some((firstOperation) => firstOperation.id === id),
      ),
    ).toBe(false)
  })

  it('rejects a capability cursor outside the active format and family filter', async () => {
    const client = await connectClient()
    const response = await client.callTool({
      name: 'office_get_capabilities',
      arguments: {
        format: 'xlsx',
        family: 'range',
        cursor: 'xlsx.sheet.add',
      },
    })

    expect(response).toMatchObject({
      isError: true,
      structuredContent: { ok: false, error: 'invalid_arguments' },
    })
  })

  it('rejects a detail operation selector in summary discovery mode', async () => {
    const client = await connectClient()
    const response = await client.callTool({
      name: 'office_get_capabilities',
      arguments: {
        format: 'xlsx',
        view: 'summary',
        operation: 'xlsx.range.set_values',
      },
    })

    expect(response).toMatchObject({
      isError: true,
      structuredContent: { ok: false, error: 'invalid_arguments' },
    })
  })

  it('rejects summary pagination selectors in detail discovery mode', async () => {
    const client = await connectClient()

    for (const selector of [
      { family: 'range' },
      { cursor: 'xlsx.range.set_values' },
      { limit: 5 },
    ]) {
      const response = await client.callTool({
        name: 'office_get_capabilities',
        arguments: {
          format: 'xlsx',
          view: 'detail',
          operation: 'xlsx.range.set_values',
          ...selector,
        },
      })
      expect(response).toMatchObject({
        isError: true,
        structuredContent: { ok: false, error: 'invalid_arguments' },
      })
    }
  })

  it('resolves detail only for a canonical Agent-visible operation in the requested format', async () => {
    const client = await connectClient()

    for (const operation of ['save', 'xlsx.document.load_staged', 'docx.document.save']) {
      const response = await client.callTool({
        name: 'office_get_capabilities',
        arguments: { format: 'xlsx', view: 'detail', operation },
      })
      expect(response).toMatchObject({
        isError: true,
        structuredContent: { ok: false, error: 'invalid_arguments' },
      })
    }
  })

  it('keeps every maximum summary page inside an 8 KiB MCP context budget', async () => {
    const client = await connectClient()

    for (const format of ['docx', 'markdown', 'xlsx', 'pptx', 'pdf']) {
      const response = await client.callTool({
        name: 'office_get_capabilities',
        arguments: { format, view: 'summary', limit: 20 },
      })
      expect(Buffer.byteLength(JSON.stringify(response.structuredContent))).toBeLessThanOrEqual(
        8_192,
      )
    }
  })

  it('reports registry capabilities with evidence-gated readiness', async () => {
    const projection = JSON.parse(
      await readFile(new URL('../src/generated/release-readiness.json', import.meta.url), 'utf8'),
    ) as { ready: boolean }
    // ADR 0005 applies one all-format release decision. A development build may
    // expose the complete registry without claiming that release evidence passed.
    const expectedReady = projection.ready
    const client = await connectClient()
    const docx = await readAllCapabilitySchemas(client, 'docx')
    const markdown = await readAllCapabilitySchemas(client, 'markdown')
    const xlsx = await readAllCapabilitySchemas(client, 'xlsx')
    const pptx = await readAllCapabilitySchemas(client, 'pptx')
    const pdf = await readAllCapabilitySchemas(client, 'pdf')

    expect(docx.structuredContent).toMatchObject({
      ok: true,
      capabilities: {
        format: 'docx',
        ready: expectedReady,
        displayModes: ['inline', 'fullscreen'],
        defaultDisplayMode: 'fullscreen',
        localFile: { picker: true, path: true, recovery: true },
        operations: {
          'docx.block.delete': { required: ['target'] },
          'docx.block.move': { required: ['blockIndexes', 'afterBlockIndex'] },
          'docx.document.save': { required: [] },
          'docx.image.update': { required: ['target', 'properties', 'fields'] },
          'docx.list.apply': { required: ['target', 'kind'] },
          'docx.list.remove': { required: ['target'] },
          'docx.paragraph.set_heading_level': { required: ['target', 'level'] },
          'docx.paragraph.set_style': { required: ['target', 'style', 'fields'] },
          'docx.text.insert': { required: ['text'] },
          'docx.text.replace_all': { required: ['containsText', 'replaceText'] },
          'docx.text.replace_selection': { required: ['text'] },
          'docx.text.set_style': { required: ['target', 'style', 'fields'] },
          'docx.toc.insert': { required: ['afterBlockIndex'] },
        },
      },
    })
    expect(
      (docx.structuredContent as { capabilities: { operations: Record<string, unknown> } })
        .capabilities.operations,
    ).not.toHaveProperty('batch_update')
    expect(markdown.structuredContent).toMatchObject({
      ok: true,
      capabilities: {
        format: 'markdown',
        ready: expectedReady,
        showTool: 'office_show_markdown_editor',
        localFile: { picker: true, path: true, recovery: true },
        operations: {
          'markdown.text.insert': { required: ['text'] },
          'markdown.text.replace_selection': { required: ['text'] },
          'markdown.document.save': { required: [] },
        },
      },
    })
    expect(xlsx.structuredContent).toMatchObject({
      ok: true,
      capabilities: {
        format: 'xlsx',
        ready: expectedReady,
        showTool: 'office_show_xlsx_editor',
        displayModes: ['inline', 'fullscreen'],
        defaultDisplayMode: 'fullscreen',
        localFile: { picker: true, path: true, recovery: true },
        operations: {
          'xlsx.cell.set_value': { required: ['sheet', 'address', 'value'] },
          'xlsx.column.copy_widths': {
            required: [
              'sourceSheet',
              'sourceColumn',
              'destinationSheet',
              'destinationColumn',
              'count',
            ],
          },
          'xlsx.column.set_width': {
            required: ['sheet', 'column', 'count', 'widthCharacters'],
          },
          'xlsx.formula.insert_aggregate': { required: ['sheet', 'range', 'function'] },
          'xlsx.history.redo': { required: [] },
          'xlsx.history.undo': { required: [] },
          'xlsx.hyperlink.remove': { required: ['sheet', 'address'] },
          'xlsx.hyperlink.set': { required: ['sheet', 'address', 'target'] },
          'xlsx.range.apply_cell_style': { required: ['sheet', 'range', 'preset'] },
          'xlsx.range.clear': { required: ['sheet', 'range', 'scope'] },
          'xlsx.range.clear_filter_criteria': { required: ['sheet', 'range'] },
          'xlsx.range.copy_formats': {
            required: ['sourceSheet', 'sourceRange', 'destinationSheet', 'destinationRange'],
          },
          'xlsx.range.copy_formulas': {
            required: ['sourceSheet', 'sourceRange', 'destinationSheet', 'destinationRange'],
          },
          'xlsx.range.copy_without_borders': {
            required: ['sourceSheet', 'sourceRange', 'destinationSheet', 'destinationRange'],
          },
          'xlsx.range.copy_values': {
            required: ['sourceSheet', 'sourceRange', 'destinationSheet', 'destinationRange'],
          },
          'xlsx.range.flash_fill': { required: ['sheet', 'range'] },
          'xlsx.range.fill': { required: ['sheet', 'range', 'direction'] },
          'xlsx.range.merge': { required: ['sheet', 'range', 'mode'] },
          'xlsx.range.remove_duplicates': { required: ['sheet', 'range', 'hasHeader'] },
          'xlsx.range.set_alignment': { required: ['sheet', 'range', 'alignment', 'fields'] },
          'xlsx.range.set_border': { required: ['sheet', 'range', 'border'] },
          'xlsx.range.set_custom_filter': {
            required: ['sheet', 'range', 'column', 'conjunction', 'conditions'],
          },
          'xlsx.range.set_filter': { required: ['sheet', 'range', 'enabled'] },
          'xlsx.range.set_filter_values': {
            required: ['sheet', 'range', 'column', 'values', 'includeBlank'],
          },
          'xlsx.range.set_fill': { required: ['sheet', 'range', 'color'] },
          'xlsx.range.set_font': { required: ['sheet', 'range', 'font', 'fields'] },
          'xlsx.range.set_number_format': { required: ['sheet', 'range', 'pattern'] },
          'xlsx.range.set_protection': {
            required: ['sheet', 'range', 'protection', 'fields'],
          },
          'xlsx.range.set_text_style': { required: ['sheet', 'range', 'style', 'fields'] },
          'xlsx.range.set_values': { required: ['sheet', 'range', 'values'] },
          'xlsx.range.sort': { required: ['sheet', 'range', 'direction'] },
          'xlsx.range.sort_custom': { required: ['sheet', 'range', 'keys', 'hasHeader'] },
          'xlsx.range.text_to_columns': { required: ['sheet', 'range', 'delimiter'] },
          'xlsx.row.set_height': { required: ['sheet', 'row', 'count', 'heightPoints'] },
          'xlsx.row.insert': { required: ['sheet', 'row', 'count'] },
          'xlsx.row.delete': { required: ['sheet', 'row', 'count'] },
          'xlsx.column.delete': { required: ['sheet', 'column', 'count'] },
          'xlsx.column.insert': { required: ['sheet', 'column', 'count'] },
          'xlsx.sheet.add': { required: ['name'] },
          'xlsx.sheet.delete': { required: ['sheet'] },
          'xlsx.sheet.move': { required: ['sheet', 'position'] },
          'xlsx.sheet.rename': { required: ['sheet', 'name'] },
          'xlsx.sheet.set_fit_to_pages': {
            required: ['sheet', 'widthPages', 'heightPages'],
          },
          'xlsx.sheet.set_freeze': {
            required: ['sheet', 'frozenRows', 'frozenColumns'],
          },
          'xlsx.sheet.set_formula_view': { required: ['sheet', 'enabled'] },
          'xlsx.sheet.set_gridlines': { required: ['sheet', 'visible'] },
          'xlsx.sheet.set_page_margins': { required: ['sheet', 'margins'] },
          'xlsx.sheet.set_page_orientation': { required: ['sheet', 'orientation'] },
          'xlsx.sheet.set_paper_size': { required: ['sheet', 'paperSize'] },
          'xlsx.sheet.set_print_area': { required: ['sheet', 'range'] },
          'xlsx.sheet.set_print_gridlines': { required: ['sheet', 'enabled'] },
          'xlsx.sheet.set_print_headings': { required: ['sheet', 'enabled'] },
          'xlsx.sheet.set_print_scale': { required: ['sheet', 'scalePercent'] },
          'xlsx.sheet.set_print_titles': { required: ['sheet', 'rows'] },
          'xlsx.sheet.set_protection': { required: ['sheet', 'protected'] },
          'xlsx.table.add': { required: ['sheet', 'range', 'style'] },
          'xlsx.document.save': { required: [] },
        },
      },
    })
    expect(
      (xlsx.structuredContent as { capabilities: { operations: Record<string, unknown> } })
        .capabilities.operations,
    ).not.toHaveProperty('save')
    expect(
      (xlsx.structuredContent as { capabilities: { operations: Record<string, unknown> } })
        .capabilities.operations,
    ).not.toHaveProperty('set_cell_value')
    expect(
      (xlsx.structuredContent as { capabilities: { operations: Record<string, unknown> } })
        .capabilities.operations,
    ).not.toHaveProperty('set_range_values')
    expect(
      (xlsx.structuredContent as { capabilities: { operations: Record<string, unknown> } })
        .capabilities.operations,
    ).not.toHaveProperty('ribbon_command')
    expect(
      (xlsx.structuredContent as { capabilities: { operations: Record<string, unknown> } })
        .capabilities.operations,
    ).not.toHaveProperty('insert_rows')
    expect(
      (xlsx.structuredContent as { capabilities: { operations: Record<string, unknown> } })
        .capabilities.operations,
    ).not.toHaveProperty('delete_rows')
    expect(
      (xlsx.structuredContent as { capabilities: { operations: Record<string, unknown> } })
        .capabilities.operations,
    ).not.toHaveProperty('insert_columns')
    expect(
      (xlsx.structuredContent as { capabilities: { operations: Record<string, unknown> } })
        .capabilities.operations,
    ).not.toHaveProperty('delete_columns')
    expect(
      (xlsx.structuredContent as { capabilities: { operations: Record<string, unknown> } })
        .capabilities.operations,
    ).not.toHaveProperty('add_sheet')
    expect(
      (xlsx.structuredContent as { capabilities: { operations: Record<string, unknown> } })
        .capabilities.operations,
    ).not.toHaveProperty('rename_sheet')
    expect(
      (xlsx.structuredContent as { capabilities: { operations: Record<string, unknown> } })
        .capabilities.operations,
    ).not.toHaveProperty('delete_sheet')
    expect(
      (xlsx.structuredContent as { capabilities: { operations: Record<string, unknown> } })
        .capabilities.operations,
    ).not.toHaveProperty('move_sheet')
    expect(
      (xlsx.structuredContent as { capabilities: { operations: Record<string, unknown> } })
        .capabilities.operations,
    ).not.toHaveProperty('xlsx.document.load_staged')
    expect(pptx.structuredContent).toMatchObject({
      ok: true,
      capabilities: {
        format: 'pptx',
        ready: expectedReady,
        showTool: 'office_show_pptx_editor',
        defaultDisplayMode: 'fullscreen',
        localFile: { picker: true, path: true, recovery: true },
        operations: {
          'pptx.document.create_blank': { required: [] },
          'pptx.document.save': { required: [] },
          'pptx.history.undo': { required: [] },
          'pptx.object.move_selection': { required: ['deltaXEmu', 'deltaYEmu'] },
          'pptx.selection.set': { required: ['slideIndex', 'objectIds'] },
          'pptx.text.replace_selection': { required: ['text'] },
        },
      },
    })
    expect(
      (pptx.structuredContent as { capabilities: { operations: Record<string, unknown> } })
        .capabilities.operations,
    ).not.toHaveProperty('save')
    expect(
      (pptx.structuredContent as { capabilities: { operations: Record<string, unknown> } })
        .capabilities.operations,
    ).not.toHaveProperty('pptx.document.load_staged')
    expect(
      (pptx.structuredContent as { capabilities: { operations: Record<string, unknown> } })
        .capabilities.operations,
    ).not.toHaveProperty('select_objects')
    expect(
      (pptx.structuredContent as { capabilities: { operations: Record<string, unknown> } })
        .capabilities.operations,
    ).not.toHaveProperty('replace_selected_text')
    expect(
      (pptx.structuredContent as { capabilities: { operations: Record<string, unknown> } })
        .capabilities.operations,
    ).not.toHaveProperty('move_selected_objects')
    expect(pdf.structuredContent).toMatchObject({
      ok: true,
      capabilities: {
        format: 'pdf',
        ready: expectedReady,
        showTool: 'office_show_pdf_editor',
        defaultDisplayMode: 'fullscreen',
        localFile: { picker: true, path: true, recovery: true },
        operations: {
          'pdf.annotation.delete_saved': {
            required: ['pageIndex', 'objNum', 'subtype', 'rect'],
          },
          'pdf.document.save': { required: [] },
          'pdf.document.set_metadata': {
            required: ['title', 'author', 'subject', 'keywords'],
          },
          'pdf.page.insert': { required: ['path', 'afterPageIndex'] },
          'pdf.text.replace': {
            required: ['pageIndex', 'rect', 'oldText', 'newText', 'fontSize'],
          },
          'pdf.history.redo': { required: [] },
          'pdf.history.undo': { required: [] },
        },
      },
    })
    expect(
      Object.keys(
        (pdf.structuredContent as { capabilities: { operations: Record<string, unknown> } })
          .capabilities.operations,
      ),
    ).toEqual([
      'pdf.annotation.delete_saved',
      'pdf.document.save',
      'pdf.document.set_metadata',
      'pdf.drawing.add',
      'pdf.drawing.update',
      'pdf.form.set_value',
      'pdf.history.redo',
      'pdf.history.undo',
      'pdf.image.delete',
      'pdf.image.insert',
      'pdf.image.replace',
      'pdf.image.transform',
      'pdf.markup.add',
      'pdf.page.delete',
      'pdf.page.insert',
      'pdf.page.reorder',
      'pdf.page.set_rotation',
      'pdf.pending.delete',
      'pdf.stamp.set',
      'pdf.static_form.set',
      'pdf.text.insert',
      'pdf.text.replace',
      'pdf.text.update_inserted',
    ])
    expect(
      (pdf.structuredContent as { capabilities: { operations: Record<string, unknown> } })
        .capabilities.operations,
    ).not.toHaveProperty('pdf.document.load_staged')
    expect(
      (pdf.structuredContent as { capabilities: { operations: Record<string, unknown> } })
        .capabilities.operations,
    ).not.toHaveProperty('pdf.page.insert_staged')
  })

  it('projects canonical Markdown operations from the generated manifest', async () => {
    const client = await connectClient()
    const response = await readAllCapabilitySchemas(client, 'markdown')
    const operations = (
      response.structuredContent as {
        capabilities: { operations: Record<string, unknown> }
      }
    ).capabilities.operations

    expect(operations).toMatchObject({
      'markdown.text.insert': {
        type: 'object',
        required: ['text'],
        properties: { text: { type: 'string' } },
        additionalProperties: false,
      },
      'markdown.text.replace_selection': {
        type: 'object',
        required: ['text'],
        properties: { text: { type: 'string' } },
        additionalProperties: false,
      },
      'markdown.selection.set': {
        type: 'object',
        required: ['from', 'to'],
        properties: {
          from: { type: 'integer', minimum: 1, maximum: 100000000 },
          to: { type: 'integer', minimum: 1, maximum: 100000000 },
        },
        additionalProperties: false,
      },
      'markdown.document.save': {
        type: 'object',
        required: [],
        properties: {},
        additionalProperties: false,
      },
      'markdown.document.save_as': {
        type: 'object',
        required: [],
        properties: {},
        additionalProperties: false,
      },
      'markdown.document.export_docx': {
        type: 'object',
        required: [],
        properties: {},
        additionalProperties: false,
      },
      'markdown.document.open_print_dialog': {
        type: 'object',
        required: [],
        properties: {},
        additionalProperties: false,
      },
      'markdown.document.set_auto_save': {
        type: 'object',
        required: ['enabled'],
        properties: { enabled: { type: 'boolean' } },
        additionalProperties: false,
      },
      'markdown.history.undo': {
        type: 'object',
        required: [],
        properties: {},
        additionalProperties: false,
      },
      'markdown.history.redo': {
        type: 'object',
        required: [],
        properties: {},
        additionalProperties: false,
      },
      'markdown.block.set_type': {
        type: 'object',
        required: ['textBlockIndex', 'type'],
        properties: {
          textBlockIndex: { type: 'integer', minimum: 0, maximum: 1000000 },
          type: {
            type: 'string',
            enum: [
              'paragraph',
              'heading_1',
              'heading_2',
              'heading_3',
              'heading_4',
              'heading_5',
              'heading_6',
              'quote',
              'code_block',
            ],
          },
        },
        additionalProperties: false,
      },
      'markdown.block.update': {
        type: 'object',
        required: ['blockIndex', 'action', 'afterBlockIndex', 'content'],
        properties: {
          blockIndex: { type: 'integer', minimum: 0, maximum: 1000000 },
          action: { type: 'string', enum: ['duplicate', 'delete', 'add_below', 'move'] },
          afterBlockIndex: { type: ['integer', 'null'], minimum: -1, maximum: 1000000 },
          content: { type: ['string', 'null'], maxLength: 65536 },
        },
        additionalProperties: false,
      },
      'markdown.text.set_marks': {
        type: 'object',
        required: ['from', 'to', 'marks'],
        properties: {
          from: { type: 'integer', minimum: 1, maximum: 100000000 },
          to: { type: 'integer', minimum: 1, maximum: 100000000 },
          marks: {
            type: 'object',
            required: ['bold', 'italic', 'strike', 'code', 'link'],
            properties: {
              bold: { type: 'boolean' },
              italic: { type: 'boolean' },
              strike: { type: 'boolean' },
              code: { type: 'boolean' },
              link: { type: ['string', 'null'], minLength: 1, maxLength: 4096 },
            },
            additionalProperties: false,
          },
        },
        additionalProperties: false,
      },
      'markdown.list.set_type': {
        type: 'object',
        required: ['textBlockIndex', 'type'],
        properties: {
          textBlockIndex: { type: 'integer', minimum: 0, maximum: 1000000 },
          type: { type: 'string', enum: ['none', 'bullet', 'ordered', 'task'] },
        },
        additionalProperties: false,
      },
      'markdown.table.insert': {
        type: 'object',
        required: ['position', 'rows', 'columns', 'headerRow'],
        properties: {
          position: { type: 'integer', minimum: 1, maximum: 100000000 },
          rows: { type: 'integer', minimum: 1, maximum: 100 },
          columns: { type: 'integer', minimum: 1, maximum: 100 },
          headerRow: { type: 'boolean' },
        },
        additionalProperties: false,
      },
      'markdown.table.update': {
        type: 'object',
        required: ['position', 'action', 'headerRow'],
        properties: {
          position: { type: 'integer', minimum: 1, maximum: 100000000 },
          action: {
            type: 'string',
            enum: [
              'add_row_before',
              'add_row_after',
              'delete_row',
              'add_column_before',
              'add_column_after',
              'delete_column',
              'set_header_row',
              'delete_table',
            ],
          },
          headerRow: { type: ['boolean', 'null'] },
        },
        additionalProperties: false,
      },
      'markdown.divider.insert': {
        type: 'object',
        required: ['position'],
        properties: {
          position: { type: 'integer', minimum: 1, maximum: 100000000 },
        },
        additionalProperties: false,
      },
      'markdown.image.insert': {
        type: 'object',
        required: ['path', 'position', 'alt', 'title'],
        properties: {
          path: { type: 'string', minLength: 1, maxLength: 1024 },
          position: { type: 'integer', minimum: 1, maximum: 100000000 },
          alt: { type: 'string', maxLength: 1024 },
          title: { type: ['string', 'null'], minLength: 1, maxLength: 1024 },
        },
        additionalProperties: false,
      },
      'markdown.frontmatter.set': {
        type: 'object',
        required: ['yaml'],
        properties: {
          yaml: { type: 'string', maxLength: 1048576 },
        },
        additionalProperties: false,
      },
      'markdown.code_block.set_language': {
        type: 'object',
        required: ['textBlockIndex', 'language'],
        properties: {
          textBlockIndex: { type: 'integer', minimum: 0, maximum: 1000000 },
          language: {
            type: 'string',
            enum: [
              'plaintext',
              'bash',
              'c',
              'cpp',
              'csharp',
              'css',
              'diff',
              'dockerfile',
              'go',
              'graphql',
              'html',
              'java',
              'javascript',
              'json',
              'kotlin',
              'lua',
              'markdown',
              'objectivec',
              'php',
              'python',
              'r',
              'ruby',
              'rust',
              'scala',
              'scss',
              'sql',
              'swift',
              'typescript',
              'xml',
              'yaml',
            ],
          },
        },
        additionalProperties: false,
      },
    })
    expect(operations).not.toHaveProperty('insert_text')
    expect(operations).not.toHaveProperty('replace_selection')
    expect(operations).not.toHaveProperty('save')
    expect(operations).not.toHaveProperty('markdown.document.load_staged')
    expect(operations).not.toHaveProperty('markdown.image.insert_staged')
  })

  it.each(['markdown.history.undo', 'markdown.history.redo'])(
    'queues and acknowledges %s through the live Markdown session',
    async (operation) => {
      await executeAndAcknowledgeMarkdownOperation(operation, {}, {})
    },
  )

  it('queues and acknowledges markdown.selection.set through the live Markdown session', async () => {
    await executeAndAcknowledgeMarkdownOperation(
      'markdown.selection.set',
      { from: 3, to: 8 },
      { from: 3, to: 8 },
    )
  })

  it('rejects a zero Markdown selection position before enqueueing', async () => {
    await expectMarkdownOperationRejectedBeforeEnqueue(
      'markdown.selection.set',
      { from: 0, to: 1 },
      'Invalid arguments for markdown.selection.set: $.from must be greater than or equal to 1.',
    )
  })

  it('queues and acknowledges markdown.document.save_as through the live Markdown session', async () => {
    await executeAndAcknowledgeMarkdownOperation(
      'markdown.document.save_as',
      {},
      { saved: true, fileName: 'copy.md' },
    )
  })

  it('queues and acknowledges markdown.document.export_docx through the live Markdown session', async () => {
    await executeAndAcknowledgeMarkdownOperation(
      'markdown.document.export_docx',
      {},
      { exported: true, fileName: 'notes.docx' },
    )
  })

  it('queues and acknowledges markdown.document.open_print_dialog through the live Markdown session', async () => {
    await executeAndAcknowledgeMarkdownOperation(
      'markdown.document.open_print_dialog',
      {},
      { opened: true },
    )
  })

  it('queues and acknowledges markdown.document.set_auto_save through the live Markdown session', async () => {
    await executeAndAcknowledgeMarkdownOperation(
      'markdown.document.set_auto_save',
      { enabled: true },
      { enabled: true },
    )
  })

  it('rejects a non-boolean Markdown autosave state before enqueueing', async () => {
    await expectMarkdownOperationRejectedBeforeEnqueue(
      'markdown.document.set_auto_save',
      { enabled: 'yes' },
      'Invalid arguments for markdown.document.set_auto_save: $.enabled must be a boolean.',
    )
  })

  it('queues and acknowledges markdown.block.set_type through the live Markdown session', async () => {
    await executeAndAcknowledgeMarkdownOperation(
      'markdown.block.set_type',
      { textBlockIndex: 2, type: 'heading_3' },
      {},
    )
  })

  it('rejects an unknown Markdown block type before enqueueing', async () => {
    await expectMarkdownOperationRejectedBeforeEnqueue(
      'markdown.block.set_type',
      { textBlockIndex: 0, type: 'title' },
      'Invalid arguments for markdown.block.set_type: $.type must be one of "paragraph", "heading_1", "heading_2", "heading_3", "heading_4", "heading_5", "heading_6", "quote", "code_block".',
    )
  })

  it('queues and acknowledges markdown.text.set_marks through the live Markdown session', async () => {
    await executeAndAcknowledgeMarkdownOperation(
      'markdown.text.set_marks',
      {
        from: 1,
        to: 6,
        marks: {
          bold: true,
          italic: false,
          strike: false,
          code: false,
          link: 'https://example.com',
        },
      },
      {},
    )
  })

  it('rejects an empty Markdown link before enqueueing', async () => {
    await expectMarkdownOperationRejectedBeforeEnqueue(
      'markdown.text.set_marks',
      {
        from: 1,
        to: 6,
        marks: {
          bold: false,
          italic: false,
          strike: false,
          code: false,
          link: '',
        },
      },
      'Invalid arguments for markdown.text.set_marks: $.marks.link must contain at least 1 character.',
    )
  })

  it('queues and acknowledges markdown.list.set_type through the live Markdown session', async () => {
    await executeAndAcknowledgeMarkdownOperation(
      'markdown.list.set_type',
      { textBlockIndex: 1, type: 'ordered' },
      {},
    )
  })

  it('rejects an unknown Markdown list type before enqueueing', async () => {
    await expectMarkdownOperationRejectedBeforeEnqueue(
      'markdown.list.set_type',
      { textBlockIndex: 0, type: 'toggle' },
      'Invalid arguments for markdown.list.set_type: $.type must be one of "none", "bullet", "ordered", "task".',
    )
  })

  it.each([
    ['markdown.table.insert', { position: 2, rows: 3, columns: 4, headerRow: true }],
    ['markdown.divider.insert', { position: 2 }],
  ])('queues and acknowledges %s through the live Markdown session', async (operation, args) => {
    await executeAndAcknowledgeMarkdownOperation(operation, args, {})
  })

  it('rejects a zero-row Markdown table before enqueueing', async () => {
    await expectMarkdownOperationRejectedBeforeEnqueue(
      'markdown.table.insert',
      { position: 1, rows: 0, columns: 3, headerRow: true },
      'Invalid arguments for markdown.table.insert: $.rows must be greater than or equal to 1.',
    )
  })

  it('queues and acknowledges markdown.frontmatter.set through the live Markdown session', async () => {
    await executeAndAcknowledgeMarkdownOperation(
      'markdown.frontmatter.set',
      { yaml: 'title: Typed\ntags:\n  - registry' },
      {},
    )
  })

  it('rejects oversized Markdown frontmatter before enqueueing', async () => {
    await expectMarkdownOperationRejectedBeforeEnqueue(
      'markdown.frontmatter.set',
      { yaml: 'x'.repeat(1_048_577) },
      'Invalid arguments for markdown.frontmatter.set: $.yaml must contain at most 1048576 characters.',
    )
  })

  it('queues and acknowledges markdown.table.update through the live Markdown session', async () => {
    await executeAndAcknowledgeMarkdownOperation(
      'markdown.table.update',
      { position: 4, action: 'add_column_after', headerRow: null },
      {},
    )
  })

  it('rejects an unknown Markdown table action before enqueueing', async () => {
    await expectMarkdownOperationRejectedBeforeEnqueue(
      'markdown.table.update',
      { position: 4, action: 'merge_cells', headerRow: null },
      'Invalid arguments for markdown.table.update: $.action must be one of "add_row_before", "add_row_after", "delete_row", "add_column_before", "add_column_after", "delete_column", "set_header_row", "delete_table".',
    )
  })

  it('queues and acknowledges markdown.block.update through the live Markdown session', async () => {
    await executeAndAcknowledgeMarkdownOperation(
      'markdown.block.update',
      {
        blockIndex: 3,
        action: 'move',
        afterBlockIndex: 0,
        content: null,
      },
      {},
    )
  })

  it('rejects an unknown Markdown block action before enqueueing', async () => {
    await expectMarkdownOperationRejectedBeforeEnqueue(
      'markdown.block.update',
      {
        blockIndex: 0,
        action: 'drag',
        afterBlockIndex: null,
        content: null,
      },
      'Invalid arguments for markdown.block.update: $.action must be one of "duplicate", "delete", "add_below", "move".',
    )
  })

  it('queues and acknowledges markdown.code_block.set_language through the live Markdown session', async () => {
    await executeAndAcknowledgeMarkdownOperation(
      'markdown.code_block.set_language',
      { textBlockIndex: 2, language: 'rust' },
      {},
    )
  })

  it('rejects an unknown Markdown code-block language before enqueueing', async () => {
    await expectMarkdownOperationRejectedBeforeEnqueue(
      'markdown.code_block.set_language',
      { textBlockIndex: 0, language: 'brainfuck' },
      'Invalid arguments for markdown.code_block.set_language: $.language must be one of "plaintext", "bash", "c", "cpp", "csharp", "css", "diff", "dockerfile", "go", "graphql", "html", "java", "javascript", "json", "kotlin", "lua", "markdown", "objectivec", "php", "python", "r", "ruby", "rust", "scala", "scss", "sql", "swift", "typescript", "xml", "yaml".',
    )
  })

  it('projects canonical DOCX operations from the generated manifest', async () => {
    const client = await connectClient()
    const response = await readAllCapabilitySchemas(client, 'docx')
    const operations = (
      response.structuredContent as {
        capabilities: { operations: Record<string, unknown> }
      }
    ).capabilities.operations

    expect(operations).toMatchObject({
      'docx.block.delete': {
        type: 'object',
        required: ['target'],
        properties: {
          target: {
            type: 'object',
            required: [],
            properties: {
              nodeType: {
                type: 'string',
                enum: ['docHeading', 'docParagraph', 'docListItem', 'image'],
              },
              headingLevel: { type: 'integer', minimum: 1, maximum: 6 },
              containsText: { type: 'string' },
              matchCase: { type: 'boolean' },
              blockIndexes: {
                type: 'array',
                items: { type: 'integer', minimum: 0 },
              },
              scope: { type: 'string', enum: ['selection', 'document'] },
            },
            additionalProperties: false,
          },
        },
        additionalProperties: false,
      },
      'docx.block.move': {
        type: 'object',
        required: ['blockIndexes', 'afterBlockIndex'],
        properties: {
          blockIndexes: {
            type: 'array',
            minItems: 1,
            items: { type: 'integer', minimum: 0 },
          },
          afterBlockIndex: { type: 'integer', minimum: -1 },
        },
        additionalProperties: false,
      },
      'docx.image.update': {
        type: 'object',
        required: ['target', 'properties', 'fields'],
        properties: {
          properties: {
            type: 'object',
            properties: {
              widthPx: { type: ['number', 'null'], minimum: 1 },
              heightPx: { type: ['number', 'null'], minimum: 1 },
              align: { type: ['string', 'null'], enum: ['left', 'center', 'right', null] },
            },
            required: [],
            additionalProperties: false,
          },
          fields: {
            type: 'array',
            minItems: 1,
            items: { type: 'string', enum: ['widthPx', 'heightPx', 'align'] },
          },
        },
        additionalProperties: false,
      },
      'docx.list.apply': {
        type: 'object',
        required: ['target', 'kind'],
        properties: {
          target: {
            type: 'object',
            required: [],
            properties: {
              nodeType: {
                type: 'string',
                enum: ['docHeading', 'docParagraph', 'docListItem', 'image'],
              },
              headingLevel: { type: 'integer', minimum: 1, maximum: 6 },
              containsText: { type: 'string' },
              matchCase: { type: 'boolean' },
              blockIndexes: {
                type: 'array',
                items: { type: 'integer', minimum: 0 },
              },
              scope: { type: 'string', enum: ['selection', 'document'] },
            },
            additionalProperties: false,
          },
          kind: { type: 'string', enum: ['bullet', 'ordered'] },
        },
        additionalProperties: false,
      },
      'docx.list.remove': {
        type: 'object',
        required: ['target'],
        properties: {
          target: {
            type: 'object',
            required: [],
            properties: {
              nodeType: {
                type: 'string',
                enum: ['docHeading', 'docParagraph', 'docListItem', 'image'],
              },
              headingLevel: { type: 'integer', minimum: 1, maximum: 6 },
              containsText: { type: 'string' },
              matchCase: { type: 'boolean' },
              blockIndexes: {
                type: 'array',
                items: { type: 'integer', minimum: 0 },
              },
              scope: { type: 'string', enum: ['selection', 'document'] },
            },
            additionalProperties: false,
          },
        },
        additionalProperties: false,
      },
      'docx.paragraph.set_heading_level': {
        type: 'object',
        required: ['target', 'level'],
        properties: {
          target: {
            type: 'object',
            required: [],
            properties: {
              nodeType: {
                type: 'string',
                enum: ['docHeading', 'docParagraph', 'docListItem', 'image'],
              },
              headingLevel: { type: 'integer', minimum: 1, maximum: 6 },
              containsText: { type: 'string' },
              matchCase: { type: 'boolean' },
              blockIndexes: {
                type: 'array',
                items: { type: 'integer', minimum: 0 },
              },
              scope: { type: 'string', enum: ['selection', 'document'] },
            },
            additionalProperties: false,
          },
          level: { type: 'integer', minimum: 0, maximum: 6 },
        },
        additionalProperties: false,
      },
      'docx.paragraph.set_style': {
        type: 'object',
        required: ['target', 'style', 'fields'],
        properties: {
          style: {
            type: 'object',
            properties: {
              align: {
                type: ['string', 'null'],
                enum: ['left', 'center', 'right', 'justify', null],
              },
              lineSpacing: { type: ['number', 'null'], minimum: 0.06 },
              indentFirstLine: { type: ['number', 'null'] },
              pageBreakBefore: { type: 'boolean' },
            },
            required: [],
            additionalProperties: false,
          },
          fields: { type: 'array', minItems: 1 },
        },
        additionalProperties: false,
      },
      'docx.text.insert': {
        type: 'object',
        required: ['text'],
        properties: { text: { type: 'string', minLength: 1, maxLength: 65_536 } },
        additionalProperties: false,
      },
      'docx.text.replace_selection': {
        type: 'object',
        required: ['text'],
        properties: { text: { type: 'string' } },
        additionalProperties: false,
      },
      'docx.text.replace_all': {
        type: 'object',
        required: ['containsText', 'replaceText'],
        properties: {
          containsText: { type: 'string' },
          replaceText: { type: 'string' },
          matchCase: { type: 'boolean' },
        },
        additionalProperties: false,
      },
      'docx.text.set_style': {
        type: 'object',
        required: ['target', 'style', 'fields'],
        properties: {
          style: {
            type: 'object',
            properties: {
              color: { type: ['string', 'null'] },
              bold: { type: 'boolean' },
              baselineOffset: {
                type: ['string', 'null'],
                enum: ['SUPERSCRIPT', 'SUBSCRIPT', 'NONE', null],
              },
            },
            required: [],
            additionalProperties: false,
          },
          fields: { type: 'array', minItems: 1 },
        },
        additionalProperties: false,
      },
      'docx.text.set_link': {
        type: 'object',
        required: ['range', 'href', 'text'],
        properties: {
          range: {
            type: 'object',
            required: ['from', 'to'],
            properties: {
              from: { type: 'integer', minimum: 1 },
              to: { type: 'integer', minimum: 1 },
            },
            additionalProperties: false,
          },
          href: { type: ['string', 'null'], minLength: 1, maxLength: 4096 },
          text: { type: ['string', 'null'], minLength: 1, maxLength: 65_536 },
        },
        additionalProperties: false,
      },
      'docx.bibliography.insert': {
        type: 'object',
        required: ['afterBlockIndex', 'heading', 'entries'],
        properties: {
          afterBlockIndex: { type: 'integer', minimum: -1, maximum: 2_147_483_647 },
          heading: { type: 'string', minLength: 1, maxLength: 4096 },
          entries: {
            type: 'array',
            minItems: 1,
            maxItems: 1024,
            items: {
              type: 'object',
              required: ['sourceTag', 'text'],
              properties: {
                sourceTag: { type: 'string', minLength: 1, maxLength: 255 },
                text: { type: 'string', minLength: 1, maxLength: 4096 },
              },
              additionalProperties: false,
            },
          },
        },
        additionalProperties: false,
      },
      'docx.bookmark.set': {
        type: 'object',
        required: ['blockIndex', 'name', 'enabled'],
        properties: {
          blockIndex: { type: 'integer', minimum: 0, maximum: 2_147_483_647 },
          name: {
            type: 'string',
            minLength: 1,
            maxLength: 40,
            pattern: '^[A-Za-z_一-鿿][A-Za-z0-9_一-鿿]*$',
          },
          enabled: { type: 'boolean' },
        },
        additionalProperties: false,
      },
      'docx.caption.insert': {
        type: 'object',
        required: ['afterBlockIndex', 'label', 'number', 'text'],
        properties: {
          afterBlockIndex: { type: 'integer', minimum: -1, maximum: 2_147_483_647 },
          label: { type: 'string', minLength: 1, maxLength: 255 },
          number: { type: 'integer', minimum: 1, maximum: 2_147_483_647 },
          text: { type: 'string', maxLength: 4096 },
        },
        additionalProperties: false,
      },
      'docx.comment.add': {
        type: 'object',
        required: ['range', 'comment'],
        properties: {
          range: {
            type: 'object',
            required: ['from', 'to'],
            properties: {
              from: { type: 'integer', minimum: 1 },
              to: { type: 'integer', minimum: 1 },
            },
            additionalProperties: false,
          },
          comment: {
            type: 'object',
            required: ['id', 'author', 'initials', 'date', 'text'],
            properties: {
              id: { type: 'string', minLength: 1, maxLength: 10, pattern: '^[0-9]+$' },
              author: { type: 'string', minLength: 1, maxLength: 255 },
              initials: { type: ['string', 'null'], maxLength: 16 },
              date: {
                type: 'string',
                minLength: 20,
                maxLength: 32,
                pattern: '^\\d{4}-\\d{2}-\\d{2}T\\d{2}:\\d{2}:\\d{2}(?:\\.\\d{1,3})?Z$',
              },
              text: { type: 'string', minLength: 1, maxLength: 65_536 },
            },
            additionalProperties: false,
          },
        },
        additionalProperties: false,
      },
      'docx.comment.delete': {
        type: 'object',
        required: ['id'],
        properties: {
          id: { type: 'string', minLength: 1, maxLength: 10, pattern: '^[0-9]+$' },
        },
        additionalProperties: false,
      },
      'docx.comment.reply': {
        type: 'object',
        required: ['parentId', 'comment'],
        properties: {
          parentId: { type: 'string', minLength: 1, maxLength: 10, pattern: '^[0-9]+$' },
          comment: {
            type: 'object',
            required: ['id', 'author', 'initials', 'date', 'text'],
            properties: {
              id: { type: 'string', minLength: 1, maxLength: 10, pattern: '^[0-9]+$' },
              author: { type: 'string', minLength: 1, maxLength: 255 },
              initials: { type: ['string', 'null'], maxLength: 16 },
              date: {
                type: 'string',
                minLength: 20,
                maxLength: 32,
                pattern: '^\\d{4}-\\d{2}-\\d{2}T\\d{2}:\\d{2}:\\d{2}(?:\\.\\d{1,3})?Z$',
              },
              text: { type: 'string', minLength: 1, maxLength: 65_536 },
            },
            additionalProperties: false,
          },
        },
        additionalProperties: false,
      },
      'docx.comment.set_resolved': {
        type: 'object',
        required: ['id', 'resolved'],
        properties: {
          id: { type: 'string', minLength: 1, maxLength: 10, pattern: '^[0-9]+$' },
          resolved: { type: 'boolean' },
        },
        additionalProperties: false,
      },
      'docx.index.insert': {
        type: 'object',
        required: ['afterBlockIndex', 'label', 'terms'],
        properties: {
          afterBlockIndex: { type: 'integer', minimum: -1, maximum: 2_147_483_647 },
          label: { type: 'string', minLength: 1, maxLength: 255 },
          terms: {
            type: 'array',
            minItems: 1,
            maxItems: 1024,
            items: { type: 'string', minLength: 1, maxLength: 4096 },
          },
        },
        additionalProperties: false,
      },
      'docx.index.mark': {
        type: 'object',
        required: ['range', 'term'],
        properties: {
          range: {
            type: 'object',
            required: ['from', 'to'],
            properties: {
              from: { type: 'integer', minimum: 1 },
              to: { type: 'integer', minimum: 1 },
            },
            additionalProperties: false,
          },
          term: {
            type: 'string',
            minLength: 1,
            maxLength: 4096,
            pattern: '^[^"\\u0000-\\u001F]+$',
          },
        },
        additionalProperties: false,
      },
      'docx.note.insert': {
        type: 'object',
        required: ['range', 'kind', 'noteId', 'text'],
        properties: {
          range: {
            type: 'object',
            required: ['from', 'to'],
            properties: {
              from: { type: 'integer', minimum: 1 },
              to: { type: 'integer', minimum: 1 },
            },
            additionalProperties: false,
          },
          kind: { type: 'string', enum: ['footnote', 'endnote'] },
          noteId: { type: 'integer', minimum: 1, maximum: 2_147_483_647 },
          text: { type: 'string', minLength: 1, maxLength: 65_536 },
        },
        additionalProperties: false,
      },
      'docx.citation.insert': {
        type: 'object',
        required: ['range', 'sourceTag', 'displayText'],
        properties: {
          range: {
            type: 'object',
            required: ['from', 'to'],
            properties: {
              from: { type: 'integer', minimum: 1 },
              to: { type: 'integer', minimum: 1 },
            },
            additionalProperties: false,
          },
          sourceTag: { type: 'string', minLength: 1, maxLength: 255 },
          displayText: { type: 'string', minLength: 1, maxLength: 4096 },
        },
        additionalProperties: false,
      },
      'docx.note.delete': {
        type: 'object',
        required: ['kind', 'noteId'],
        properties: {
          kind: { type: 'string', enum: ['footnote', 'endnote'] },
          noteId: { type: 'integer', minimum: 1, maximum: 2_147_483_647 },
        },
        additionalProperties: false,
      },
      'docx.note.update': {
        type: 'object',
        required: ['kind', 'noteId', 'text'],
        properties: {
          kind: { type: 'string', enum: ['footnote', 'endnote'] },
          noteId: { type: 'integer', minimum: 1, maximum: 2_147_483_647 },
          text: { type: 'string', minLength: 1, maxLength: 65_536 },
        },
        additionalProperties: false,
      },
      'docx.revision.set_tracking': {
        type: 'object',
        required: ['enabled'],
        properties: {
          enabled: { type: 'boolean' },
        },
        additionalProperties: false,
      },
      'docx.revision.apply_decision': {
        type: 'object',
        required: ['decision', 'scope'],
        properties: {
          decision: { type: 'string', enum: ['accept', 'reject'] },
          scope: { type: 'string', enum: ['current', 'all'] },
        },
        additionalProperties: false,
      },
      'docx.document.set_protection': {
        type: 'object',
        required: ['enabled', 'password'],
        properties: {
          enabled: { type: 'boolean' },
          password: { type: ['string', 'null'], minLength: 1, maxLength: 255 },
        },
        additionalProperties: false,
      },
      'docx.wordart.insert': {
        type: 'object',
        required: ['afterBlockIndex', 'preset', 'text', 'widthEmu', 'heightEmu', 'drawingId'],
        properties: {
          afterBlockIndex: { type: 'integer', minimum: -1, maximum: 2_147_483_647 },
          preset: {
            type: 'string',
            enum: [
              'blue',
              'gold',
              'red',
              'purple',
              'green-italic',
              'white-orange',
              'white-red',
              'gold-brown',
              'sky-navy',
              'navy-white',
              'black-gold',
              'silver-dark',
            ],
          },
          text: { type: 'string', minLength: 1, maxLength: 4096 },
          drawingId: { type: 'integer', minimum: 1, maximum: 2_147_483_647 },
        },
        additionalProperties: false,
      },
      'docx.paragraph.set_drop_cap': {
        type: 'object',
        required: ['blockIndex', 'mode', 'lines'],
        properties: {
          blockIndex: { type: 'integer', minimum: 0, maximum: 2_147_483_647 },
          mode: { type: 'string', enum: ['none', 'drop', 'margin'] },
          lines: { type: ['integer', 'null'], minimum: 2, maximum: 10 },
        },
        additionalProperties: false,
      },
      'docx.cover_page.insert': {
        type: 'object',
        required: ['preset', 'title', 'subtitle', 'author', 'date', 'year'],
        properties: {
          preset: {
            type: 'string',
            enum: [
              'classic',
              'banded',
              'boxed',
              'sideline',
              'modern',
              'elegant',
              'minimal',
              'dark',
              'accent',
              'badge',
              'facet',
              'annual',
            ],
          },
          year: { type: 'integer', minimum: 1900, maximum: 9999 },
        },
        additionalProperties: false,
      },
      'docx.document.set_design': {
        type: 'object',
        required: ['fields'],
        properties: {
          fields: {
            type: 'array',
            minItems: 1,
            maxItems: 4,
            items: {
              type: 'string',
              enum: ['pageColor', 'watermark', 'themeFonts', 'themeColors'],
            },
          },
          pageColor: {
            type: ['string', 'null'],
            minLength: 6,
            maxLength: 6,
            pattern: '^[0-9A-F]{6}$',
          },
        },
        additionalProperties: false,
      },
      'docx.document.compare': {
        type: 'object',
        required: ['path'],
        properties: {
          path: { type: 'string', minLength: 1, maxLength: 1024 },
        },
        additionalProperties: false,
      },
      'docx.ink.apply': {
        type: 'object',
        required: ['action'],
        properties: {
          action: { type: 'string', enum: ['add', 'delete', 'clear'] },
          annotation: {
            type: 'object',
            required: ['id', 'anchorIndex', 'tool', 'color', 'width', 'points'],
            properties: {
              id: {
                type: 'string',
                minLength: 1,
                maxLength: 128,
                pattern: '^[A-Za-z0-9_.:-]+$',
              },
              anchorIndex: { type: 'integer', minimum: 0, maximum: 2_147_483_647 },
              tool: { type: 'string', enum: ['pen', 'highlighter'] },
              color: { type: 'string', pattern: '^[0-9A-Fa-f]{6}$' },
              width: { type: 'number', minimum: 0.1, maximum: 100 },
              points: {
                type: 'array',
                minItems: 1,
                maxItems: 4096,
                items: {
                  type: 'object',
                  required: ['x', 'y'],
                  properties: {
                    x: { type: 'number', minimum: -1_000_000, maximum: 1_000_000 },
                    y: { type: 'number', minimum: -1_000_000, maximum: 1_000_000 },
                  },
                  additionalProperties: false,
                },
              },
            },
            additionalProperties: false,
          },
          ids: {
            type: 'array',
            minItems: 1,
            maxItems: 1024,
            items: {
              type: 'string',
              minLength: 1,
              maxLength: 128,
              pattern: '^[A-Za-z0-9_.:-]+$',
            },
          },
        },
        additionalProperties: false,
      },
      'docx.source.upsert': {
        type: 'object',
        required: ['source'],
        properties: {
          source: {
            type: 'object',
            required: ['tag', 'type', 'author', 'title', 'year', 'publisher', 'url'],
            properties: {
              tag: {
                type: 'string',
                minLength: 1,
                maxLength: 255,
                pattern: '^[A-Za-z0-9_.-]+$',
              },
              type: {
                type: 'string',
                enum: ['Book', 'JournalArticle', 'InternetSite', 'Report', 'Misc'],
              },
              author: { type: 'string', maxLength: 4096 },
              title: { type: 'string', minLength: 1, maxLength: 4096 },
              year: { type: 'string', maxLength: 32 },
              publisher: { type: ['string', 'null'], maxLength: 4096 },
              url: { type: ['string', 'null'], maxLength: 4096 },
            },
            additionalProperties: false,
          },
        },
        additionalProperties: false,
      },
      'docx.cross_reference.insert': {
        type: 'object',
        required: ['range', 'bookmarkName', 'displayText'],
        properties: {
          range: {
            type: 'object',
            required: ['from', 'to'],
            properties: {
              from: { type: 'integer', minimum: 1 },
              to: { type: 'integer', minimum: 1 },
            },
            additionalProperties: false,
          },
          bookmarkName: {
            type: 'string',
            minLength: 1,
            maxLength: 40,
            pattern: '^[A-Za-z_一-鿿][A-Za-z0-9_一-鿿]*$',
          },
          displayText: { type: 'string', minLength: 1, maxLength: 65_536 },
        },
        additionalProperties: false,
      },
      'docx.field.insert': {
        type: 'object',
        required: ['range', 'instruction', 'displayText'],
        properties: {
          range: {
            type: 'object',
            required: ['from', 'to'],
            properties: {
              from: { type: 'integer', minimum: 1 },
              to: { type: 'integer', minimum: 1 },
            },
            additionalProperties: false,
          },
          instruction: {
            type: 'string',
            enum: ['DATE', 'TIME', 'PAGE', 'NUMPAGES', 'FILENAME'],
          },
          displayText: { type: 'string', minLength: 1, maxLength: 65_536 },
        },
        additionalProperties: false,
      },
      'docx.field.update': {
        type: 'object',
        required: ['updates'],
        properties: {
          updates: {
            type: 'array',
            minItems: 1,
            maxItems: 1024,
            items: {
              type: 'object',
              required: ['range', 'instruction', 'displayText'],
              properties: {
                range: {
                  type: 'object',
                  required: ['from', 'to'],
                  properties: {
                    from: { type: 'integer', minimum: 1 },
                    to: { type: 'integer', minimum: 1 },
                  },
                  additionalProperties: false,
                },
                instruction: {
                  type: 'string',
                  minLength: 1,
                  maxLength: 512,
                  pattern: '^[A-Za-z]+(?:\\s.*)?$',
                },
                displayText: { type: 'string', minLength: 1, maxLength: 65_536 },
              },
              additionalProperties: false,
            },
          },
        },
        additionalProperties: false,
      },
      'docx.toc.insert': {
        type: 'object',
        required: ['afterBlockIndex'],
        properties: { afterBlockIndex: { type: 'integer', minimum: -1 } },
        additionalProperties: false,
      },
      'docx.toc.refresh': {
        type: 'object',
        required: ['tocBlockIndex', 'entries'],
        properties: {
          tocBlockIndex: { type: 'integer', minimum: 0, maximum: 2_147_483_647 },
          entries: {
            type: 'array',
            minItems: 1,
            maxItems: 1024,
            items: {
              type: 'object',
              required: ['level', 'text', 'pageNumber'],
              properties: {
                level: { type: 'integer', minimum: 1, maximum: 9 },
                text: { type: 'string', minLength: 1, maxLength: 4096 },
                pageNumber: {
                  type: ['integer', 'null'],
                  minimum: 1,
                  maximum: 2_147_483_647,
                },
              },
              additionalProperties: false,
            },
          },
        },
        additionalProperties: false,
      },
      'docx.document.save': {
        type: 'object',
        required: [],
        properties: {},
        additionalProperties: false,
      },
    })
    expect(operations).not.toHaveProperty('insert_text')
    expect(operations).not.toHaveProperty('replace_selection')
    expect(operations).not.toHaveProperty('save')
    expect(operations).not.toHaveProperty('docx.document.load_staged')
    expect(operations).not.toHaveProperty('docx.document.compare_staged')
    expect(operations).not.toHaveProperty('docx.image.insert_staged')
    expect(operations).not.toHaveProperty('docx.image.replace_staged')
    expect(
      (
        operations['docx.text.set_style'] as {
          properties: { style: { properties: Record<string, unknown> } }
        }
      ).properties.style.properties,
    ).not.toHaveProperty('link')
    expect(Object.keys(operations)).toEqual([
      'docx.bibliography.insert',
      'docx.block.delete',
      'docx.block.move',
      'docx.bookmark.set',
      'docx.caption.insert',
      'docx.chart.insert',
      'docx.chart.update',
      'docx.citation.insert',
      'docx.comment.add',
      'docx.comment.delete',
      'docx.comment.reply',
      'docx.comment.set_resolved',
      'docx.cover_page.insert',
      'docx.cross_reference.insert',
      'docx.document.compare',
      'docx.document.insert_page_break',
      'docx.document.save',
      'docx.document.set_design',
      'docx.document.set_different_odd_even_pages',
      'docx.document.set_protection',
      'docx.equation.insert',
      'docx.equation.update',
      'docx.field.insert',
      'docx.field.update',
      'docx.header_footer.set_page_number',
      'docx.header_footer.set_paragraphs',
      'docx.header_footer.set_text',
      'docx.history.redo',
      'docx.history.undo',
      'docx.image.insert',
      'docx.image.remove',
      'docx.image.replace',
      'docx.image.set_crop',
      'docx.image.set_margin_position',
      'docx.image.set_offset_position',
      'docx.image.set_transform',
      'docx.image.set_wrap',
      'docx.image.update',
      'docx.index.insert',
      'docx.index.mark',
      'docx.ink.apply',
      'docx.line.insert',
      'docx.list.apply',
      'docx.list.apply_preset',
      'docx.list.continue',
      'docx.list.remove',
      'docx.list.restart',
      'docx.list.set_level',
      'docx.note.delete',
      'docx.note.insert',
      'docx.note.update',
      'docx.object.remove',
      'docx.object.set_offset_position',
      'docx.object.set_size',
      'docx.object.set_style',
      'docx.paragraph.set_direction',
      'docx.paragraph.set_drop_cap',
      'docx.paragraph.set_heading_level',
      'docx.paragraph.set_style',
      'docx.revision.apply_decision',
      'docx.revision.set_tracking',
      'docx.section.insert_break',
      'docx.section.set_columns',
      'docx.section.set_different_first_page',
      'docx.section.set_margins',
      'docx.section.set_orientation',
      'docx.section.set_page_border',
      'docx.section.set_page_numbering',
      'docx.section.set_page_size',
      'docx.shape.insert',
      'docx.source.upsert',
      'docx.table.delete',
      'docx.table.delete_columns',
      'docx.table.delete_rows',
      'docx.table.insert',
      'docx.table.insert_columns',
      'docx.table.insert_rows',
      'docx.table.merge_cells',
      'docx.table.set_cell_borders',
      'docx.table.set_cell_format',
      'docx.table.set_column_widths',
      'docx.table.set_row_height',
      'docx.table.set_style',
      'docx.table.split_cell',
      'docx.text.clear_character_format',
      'docx.text.insert',
      'docx.text.replace_all',
      'docx.text.replace_selection',
      'docx.text.set_character_format',
      'docx.text.set_character_style',
      'docx.text.set_link',
      'docx.text.set_style',
      'docx.text.transform_case',
      'docx.textbox.insert',
      'docx.textbox.set_content',
      'docx.toc.insert',
      'docx.toc.refresh',
      'docx.wordart.insert',
    ])
  })

  it.each([
    [
      'docx.wordart.insert',
      {
        afterBlockIndex: 0,
        preset: 'blue',
        text: 'Registry WordArt',
        widthEmu: 2_700_000,
        heightEmu: 720_000,
        drawingId: 42,
      },
      {
        summary: 'Inserted DOCX blue WordArt at block 1',
        blockIndex: 1,
        preset: 'blue',
        drawingId: 42,
        changed: true,
      },
    ],
    [
      'docx.paragraph.set_drop_cap',
      { blockIndex: 0, mode: 'drop', lines: 3 },
      {
        summary: 'Set DOCX drop cap drop on block 0',
        blockIndex: 0,
        mode: 'drop',
        lines: 3,
        changed: true,
      },
    ],
    [
      'docx.cover_page.insert',
      {
        preset: 'classic',
        title: 'Registry Architecture',
        subtitle: 'Typed operations',
        author: 'Agent',
        date: '30 August 2026',
        year: 2026,
      },
      {
        summary: 'Inserted DOCX classic cover page with 6 blocks',
        preset: 'classic',
        insertedBlocks: 6,
        changed: true,
      },
    ],
    [
      'docx.document.set_design',
      { fields: ['pageColor', 'watermark'], pageColor: 'FFF9E6', watermark: 'DRAFT' },
      {
        summary: 'Set DOCX document design fields: pageColor, watermark',
        fields: ['pageColor', 'watermark'],
        changedFields: ['pageColor', 'watermark'],
        changed: true,
      },
    ],
    [
      'docx.ink.apply',
      {
        action: 'add',
        annotation: {
          id: 'ink-agent-1',
          anchorIndex: 0,
          tool: 'pen',
          color: 'C00000',
          width: 2,
          points: [{ x: 1, y: 2 }],
        },
      },
      {
        summary: 'Applied DOCX ink add: 1 added, 0 deleted',
        action: 'add',
        added: 1,
        deleted: 0,
        count: 1,
        changed: true,
      },
    ],
    [
      'docx.document.set_protection',
      { enabled: true, password: null },
      {
        summary: 'Enabled DOCX document protection',
        enabled: true,
        passwordProtected: false,
        changed: true,
      },
    ],
    [
      'docx.revision.apply_decision',
      { decision: 'accept', scope: 'all' },
      {
        summary: 'Accepted 1 DOCX revision(s) in all scope',
        decision: 'accept',
        scope: 'all',
        matched: 1,
        remaining: 0,
        changed: true,
      },
    ],
    [
      'docx.revision.set_tracking',
      { enabled: true },
      {
        summary: 'Enabled DOCX revision tracking',
        enabled: true,
        changed: true,
      },
    ],
    [
      'docx.comment.delete',
      { id: '1' },
      {
        summary: 'Deleted 1 DOCX comment record(s) for 1',
        id: '1',
        deleted: 1,
        anchors: 1,
        changed: true,
      },
    ],
    [
      'docx.comment.set_resolved',
      { id: '1', resolved: true },
      {
        summary: 'Resolved DOCX comment thread 1',
        id: '1',
        resolved: true,
        affected: 1,
        changed: true,
      },
    ],
    [
      'docx.comment.reply',
      {
        parentId: '1',
        comment: {
          id: '2',
          author: 'Agent',
          initials: null,
          date: '2026-08-30T02:01:00Z',
          text: 'Reply',
        },
      },
      {
        summary: 'Added DOCX comment reply 2 to parent 1',
        id: '2',
        parentId: '1',
        references: 1,
        changed: true,
      },
    ],
    [
      'docx.comment.add',
      {
        range: { from: 1, to: 6 },
        comment: {
          id: '1',
          author: 'Agent',
          initials: null,
          date: '2026-08-30T02:00:00Z',
          text: 'Review this text.',
        },
      },
      {
        summary: 'Added DOCX comment 1 at range 1..6',
        id: '1',
        from: 1,
        to: 6,
        changed: true,
      },
    ],
    [
      'docx.index.insert',
      { afterBlockIndex: 0, label: 'Index', terms: ['Beta', 'Alpha'] },
      {
        summary: 'Inserted DOCX index with 2 entries after block 0',
        afterBlockIndex: 0,
        entries: 2,
        insertedBlocks: 2,
      },
    ],
    [
      'docx.index.mark',
      { range: { from: 1, to: 6 }, term: 'Alpha' },
      {
        summary: 'Marked DOCX index term Alpha at range 1..6',
        from: 1,
        to: 6,
        term: 'Alpha',
        changed: true,
      },
    ],
    [
      'docx.caption.insert',
      {
        afterBlockIndex: 0,
        label: 'Figure',
        number: 2,
        text: 'Registry architecture',
      },
      {
        summary: 'Inserted DOCX caption Figure 2 after block 0',
        afterBlockIndex: 0,
        label: 'Figure',
        number: 2,
        changed: true,
      },
    ],
    [
      'docx.bibliography.insert',
      {
        afterBlockIndex: 0,
        heading: 'Bibliography',
        entries: [{ sourceTag: 'Wang2026', text: 'Wang, Wei. (2026). Registry.' }],
      },
      {
        summary: 'Inserted 1 DOCX bibliography entries after block 0',
        afterBlockIndex: 0,
        entries: 1,
        insertedBlocks: 2,
      },
    ],
    [
      'docx.citation.insert',
      {
        range: { from: 6, to: 6 },
        sourceTag: 'Wang2026',
        displayText: '(Wang, Wei, 2026)',
      },
      {
        summary: 'Inserted DOCX citation for source Wang2026',
        from: 6,
        to: 6,
        sourceTag: 'Wang2026',
        changed: true,
      },
    ],
    [
      'docx.source.upsert',
      {
        source: {
          tag: 'Wang2026',
          type: 'Book',
          author: 'Wang, Wei',
          title: 'Registry Architecture',
          year: '2026',
          publisher: null,
          url: null,
        },
      },
      {
        summary: 'Added DOCX source Wang2026',
        tag: 'Wang2026',
        created: true,
        changed: true,
      },
    ],
    [
      'docx.note.delete',
      { kind: 'endnote', noteId: 2 },
      {
        summary: 'Deleted DOCX endnote 2',
        kind: 'endnote',
        noteId: '2',
        references: 1,
        renumbered: 0,
        changed: true,
      },
    ],
    [
      'docx.note.update',
      { kind: 'footnote', noteId: 2, text: 'Updated Broker note' },
      {
        summary: 'Updated DOCX footnote 2',
        kind: 'footnote',
        noteId: '2',
        references: 1,
        changed: true,
      },
    ],
    [
      'docx.note.insert',
      {
        range: { from: 6, to: 6 },
        kind: 'footnote',
        noteId: 2,
        text: 'Broker note',
      },
      {
        summary: 'Inserted DOCX footnote 2 at range 6..6',
        from: 6,
        to: 6,
        kind: 'footnote',
        noteId: '2',
        number: 1,
        changed: true,
      },
    ],
    [
      'docx.text.set_style',
      {
        target: { blockIndexes: [0] },
        style: { color: null, bold: true },
        fields: ['color', 'bold'],
      },
      {
        summary: 'Updated text style in 1 block(s)',
        matched: 1,
        changed: 1,
        skippedProtected: 0,
      },
    ],
    [
      'docx.paragraph.set_style',
      {
        target: { blockIndexes: [1] },
        style: { align: 'center' },
        fields: ['align'],
      },
      {
        summary: 'Updated paragraph style in 1 block(s)',
        matched: 1,
        changed: 1,
        skippedProtected: 0,
      },
    ],
    [
      'docx.block.move',
      { blockIndexes: [1, 2], afterBlockIndex: 3 },
      { summary: 'Moved 2 block(s)', matched: 2, changed: 2, skippedProtected: 0 },
    ],
    [
      'docx.image.update',
      {
        target: { nodeType: 'image' },
        properties: { widthPx: 200 },
        fields: ['widthPx'],
      },
      { summary: 'Updated 1 image(s)', matched: 1, changed: 1, skippedProtected: 0 },
    ],
    [
      'docx.toc.insert',
      { afterBlockIndex: -1 },
      {
        summary: 'Inserted a table of contents with 2 entries',
        matched: 2,
        changed: 2,
        skippedProtected: 0,
      },
    ],
  ] as const)(
    'queues %s through the generated Broker contract',
    async (operation, args, output) => {
      await executeAndAcknowledgeDocxOperation(operation, args, output)
    },
  )

  it.each([
    [
      'docx.wordart.insert',
      {
        afterBlockIndex: 0,
        preset: 'custom',
        text: 'Registry WordArt',
        widthEmu: 2_700_000,
        heightEmu: 720_000,
        drawingId: 42,
      },
      'Invalid arguments for docx.wordart.insert: $.preset must be one of "blue", "gold", "red", "purple", "green-italic", "white-orange", "white-red", "gold-brown", "sky-navy", "navy-white", "black-gold", "silver-dark".',
    ],
    [
      'docx.paragraph.set_drop_cap',
      { blockIndex: 0, mode: 'drop', lines: 1 },
      'Invalid arguments for docx.paragraph.set_drop_cap: $.lines must be greater than or equal to 2.',
    ],
    [
      'docx.cover_page.insert',
      {
        preset: 'custom',
        title: 'Registry Architecture',
        subtitle: 'Typed operations',
        author: 'Agent',
        date: '30 August 2026',
        year: 2026,
      },
      'Invalid arguments for docx.cover_page.insert: $.preset must be one of "classic", "banded", "boxed", "sideline", "modern", "elegant", "minimal", "dark", "accent", "badge", "facet", "annual".',
    ],
    [
      'docx.document.set_design',
      { fields: ['pageColor'], pageColor: 'fff9e6' },
      'Invalid arguments for docx.document.set_design: $.pageColor must match pattern ^[0-9A-F]{6}$.',
    ],
    [
      'docx.document.compare',
      { path: '' },
      'Invalid arguments for docx.document.compare: $.path must contain at least 1 character.',
    ],
    [
      'docx.ink.apply',
      {
        action: 'add',
        annotation: {
          id: 'ink-agent-1',
          anchorIndex: 0,
          tool: 'pen',
          color: 'red',
          width: 2,
          points: [{ x: 1, y: 2 }],
        },
      },
      'Invalid arguments for docx.ink.apply: $.annotation.color must match pattern ^[0-9A-Fa-f]{6}$.',
    ],
    [
      'docx.document.set_protection',
      { enabled: true, password: '' },
      'Invalid arguments for docx.document.set_protection: $.password must contain at least 1 character.',
    ],
    [
      'docx.revision.apply_decision',
      { decision: 'keep', scope: 'all' },
      'Invalid arguments for docx.revision.apply_decision: $.decision must be one of "accept", "reject".',
    ],
    [
      'docx.revision.set_tracking',
      { enabled: 'yes' },
      'Invalid arguments for docx.revision.set_tracking: $.enabled must be a boolean.',
    ],
    [
      'docx.comment.delete',
      { id: 'bad-id' },
      'Invalid arguments for docx.comment.delete: $.id must match pattern ^[0-9]+$.',
    ],
    [
      'docx.comment.set_resolved',
      { id: '1', resolved: 'yes' },
      'Invalid arguments for docx.comment.set_resolved: $.resolved must be a boolean.',
    ],
    [
      'docx.comment.reply',
      {
        parentId: '1',
        comment: {
          id: '2',
          author: 'Agent',
          initials: null,
          date: '2026-08-30T02:01:00Z',
          text: '',
        },
      },
      'Invalid arguments for docx.comment.reply: $.comment.text must contain at least 1 character.',
    ],
    [
      'docx.comment.add',
      {
        range: { from: 1, to: 6 },
        comment: {
          id: 'bad-id',
          author: 'Agent',
          initials: null,
          date: '2026-08-30T02:00:00Z',
          text: 'Invalid',
        },
      },
      'Invalid arguments for docx.comment.add: $.comment.id must match pattern ^[0-9]+$.',
    ],
    [
      'docx.index.insert',
      { afterBlockIndex: 0, label: 'Index', terms: [] },
      'Invalid arguments for docx.index.insert: $.terms must contain at least 1 item.',
    ],
    [
      'docx.index.mark',
      { range: { from: 1, to: 1 }, term: '"Alpha"' },
      'Invalid arguments for docx.index.mark: $.term must match pattern ^[^"\\u0000-\\u001F]+$.',
    ],
    [
      'docx.caption.insert',
      { afterBlockIndex: 0, label: 'Figure', number: 0, text: '' },
      'Invalid arguments for docx.caption.insert: $.number must be greater than or equal to 1.',
    ],
    [
      'docx.bibliography.insert',
      { afterBlockIndex: 0, heading: 'Bibliography', entries: [] },
      'Invalid arguments for docx.bibliography.insert: $.entries must contain at least 1 item.',
    ],
    [
      'docx.citation.insert',
      { range: { from: 1, to: 1 }, sourceTag: 'Wang2026', displayText: '' },
      'Invalid arguments for docx.citation.insert: $.displayText must contain at least 1 character.',
    ],
    [
      'docx.source.upsert',
      {
        source: {
          tag: 'bad tag',
          type: 'Book',
          author: '',
          title: 'Invalid',
          year: '',
          publisher: null,
          url: null,
        },
      },
      'Invalid arguments for docx.source.upsert: $.source.tag must match pattern ^[A-Za-z0-9_.-]+$.',
    ],
    [
      'docx.note.delete',
      { kind: 'endnote', noteId: 2, text: 'unexpected' },
      'Invalid arguments for docx.note.delete: $.text is not allowed.',
    ],
    [
      'docx.note.update',
      { kind: 'comment', noteId: 2, text: 'Invalid' },
      'Invalid arguments for docx.note.update: $.kind must be one of "footnote", "endnote".',
    ],
    [
      'docx.note.insert',
      { range: { from: 1, to: 1 }, kind: 'footnote', noteId: 0, text: 'Invalid' },
      'Invalid arguments for docx.note.insert: $.noteId must be greater than or equal to 1.',
    ],
    [
      'docx.toc.refresh',
      { tocBlockIndex: 0, entries: [] },
      'Invalid arguments for docx.toc.refresh: $.entries must contain at least 1 item.',
    ],
    [
      'docx.text.insert',
      { text: '' },
      'Invalid arguments for docx.text.insert: $.text must contain at least 1 character.',
    ],
    [
      'docx.field.update',
      { updates: [] },
      'Invalid arguments for docx.field.update: $.updates must contain at least 1 item.',
    ],
    [
      'docx.field.insert',
      { range: { from: 1, to: 1 }, instruction: 'FORMULA', displayText: '1' },
      'Invalid arguments for docx.field.insert: $.instruction must be one of "DATE", "TIME", "PAGE", "NUMPAGES", "FILENAME".',
    ],
    [
      'docx.cross_reference.insert',
      { range: { from: 1, to: 1 }, bookmarkName: 'Anchor', displayText: '' },
      'Invalid arguments for docx.cross_reference.insert: $.displayText must contain at least 1 character.',
    ],
    [
      'docx.bookmark.set',
      { blockIndex: 0, name: '1invalid', enabled: true },
      'Invalid arguments for docx.bookmark.set: $.name must match pattern ^[A-Za-z_一-鿿][A-Za-z0-9_一-鿿]*$.',
    ],
    [
      'docx.text.set_link',
      { range: { from: 1, to: 2 }, href: '', text: null },
      'Invalid arguments for docx.text.set_link: $.href must contain at least 1 character.',
    ],
    [
      'docx.paragraph.set_style',
      {
        target: { blockIndexes: [0] },
        style: { align: 'diagonal' },
        fields: ['align'],
      },
      'Invalid arguments for docx.paragraph.set_style: $.style.align must be one of "left", "center", "right", "justify", null.',
    ],
    [
      'docx.block.move',
      { blockIndexes: [], afterBlockIndex: -1 },
      'Invalid arguments for docx.block.move: $.blockIndexes must contain at least 1 item.',
    ],
    [
      'docx.image.update',
      {
        target: { nodeType: 'image' },
        properties: { widthPx: 0 },
        fields: ['widthPx'],
      },
      'Invalid arguments for docx.image.update: $.properties.widthPx must be greater than or equal to 1.',
    ],
    [
      'docx.toc.insert',
      { afterBlockIndex: -2 },
      'Invalid arguments for docx.toc.insert: $.afterBlockIndex must be greater than or equal to -1.',
    ],
  ] as const)(
    'rejects malformed %s arguments before enqueueing',
    async (operation, args, message) => {
      await expectDocxOperationRejectedBeforeEnqueue(operation, args, message)
    },
  )

  it.each([
    ['Markdown', 'markdown', 'markdown.document.load_staged'],
    ['Markdown image', 'markdown', 'markdown.image.insert_staged'],
    ['DOCX', 'docx', 'docx.document.load_staged'],
    ['XLSX', 'xlsx', 'xlsx.document.load_staged'],
    ['PPTX', 'pptx', 'pptx.document.load_staged'],
    ['PDF', 'pdf', 'pdf.document.load_staged'],
  ] as const)(
    'rejects the internal %s staged operation from office_execute',
    async (_label, format, operation) => {
      const client = await connectClient()
      const created = await client.callTool({
        name: 'office_create_session',
        arguments: { format },
      })
      const sessionId = (created.structuredContent as { session: { id: string } }).session.id
      await client.callTool({ name: 'office_editor_poll', arguments: { sessionId } })

      const execution = executeOperation(client, sessionId, 0, operation, {})

      await new Promise((resolve) => setTimeout(resolve, 20))
      const polled = await client.callTool({
        name: 'office_editor_poll',
        arguments: { sessionId },
      })
      const command = (
        polled.structuredContent as {
          commands: Array<{ commandId: string }>
        }
      ).commands[0]
      if (command) {
        await client.callTool({
          name: 'office_editor_acknowledge',
          arguments: {
            sessionId,
            commandId: command.commandId,
            ok: false,
            error: 'operation_not_found',
            message: `Operation ${operation} is not a canonical Agent operation for ${format}.`,
          },
        })
      }

      await expect(execution).resolves.toMatchObject({
        isError: true,
        structuredContent: {
          ok: false,
          error: 'operation_not_found',
          message: `Operation ${operation} is not a canonical Agent operation for ${format}.`,
        },
      })
      expect(command).toBeUndefined()
    },
  )

  it('queues canonical Markdown text insertion', async () => {
    await executeAndAcknowledgeMarkdownOperation('markdown.text.insert', { text: 'hello' }, {})
  })

  it('queues canonical DOCX text insertion', async () => {
    await executeAndAcknowledgeDocxOperation('docx.text.insert', { text: 'hello' }, {})
  })

  it('queues canonical DOCX selection replacement', async () => {
    await executeAndAcknowledgeDocxOperation('docx.text.replace_selection', { text: 'hello' }, {})
  })

  it('queues DOCX replace-all with its canonical arguments and returns the renderer output', async () => {
    const client = await connectClient()
    const created = await client.callTool({
      name: 'office_create_session',
      arguments: { format: 'docx' },
    })
    const sessionId = (created.structuredContent as { session: { id: string } }).session.id
    await client.callTool({ name: 'office_editor_poll', arguments: { sessionId } })

    const execution = executeOperation(client, sessionId, 0, 'docx.text.replace_all', {
      containsText: 'acme',
      replaceText: 'Codex',
      matchCase: false,
    })

    let command:
      { commandId: string; operation: string; arguments: Record<string, unknown> } | undefined
    for (let attempt = 0; attempt < 20 && !command; attempt += 1) {
      const polled = await client.callTool({
        name: 'office_editor_poll',
        arguments: { sessionId },
      })
      command = (
        polled.structuredContent as {
          commands: Array<{
            commandId: string
            operation: string
            arguments: Record<string, unknown>
          }>
        }
      ).commands[0]
      if (!command) await new Promise((resolve) => setTimeout(resolve, 10))
    }
    expect(command).toMatchObject({
      operation: 'docx.text.replace_all',
      arguments: { containsText: 'acme', replaceText: 'Codex', matchCase: false },
    })

    await client.callTool({
      name: 'office_editor_acknowledge',
      arguments: {
        sessionId,
        commandId: command!.commandId,
        revision: 1,
        dirty: true,
        output: {
          summary: 'replaced 2 occurrence(s) in 1 block(s)',
          matched: 1,
          changed: 1,
          skippedProtected: 0,
          skippedDeleted: 0,
        },
      },
    })
    await expect(execution).resolves.toMatchObject({
      structuredContent: {
        ok: true,
        result: {
          revision: 1,
          operations: [
            {
              id: 'docx.text.replace_all',
              result: {
                summary: 'replaced 2 occurrence(s) in 1 block(s)',
                matched: 1,
                changed: 1,
                skippedProtected: 0,
                skippedDeleted: 0,
              },
            },
          ],
        },
      },
    })
  })

  it('queues DOCX heading-level changes with canonical target arguments and returns the renderer output', async () => {
    const client = await connectClient()
    const created = await client.callTool({
      name: 'office_create_session',
      arguments: { format: 'docx' },
    })
    const sessionId = (created.structuredContent as { session: { id: string } }).session.id
    await client.callTool({
      name: 'office_editor_poll',
      arguments: { sessionId, selection: { from: 1, to: 1 } },
    })

    const execution = executeOperation(client, sessionId, 0, 'docx.paragraph.set_heading_level', {
      target: { blockIndexes: [1] },
      level: 3,
    })

    let command:
      { commandId: string; operation: string; arguments: Record<string, unknown> } | undefined
    for (let attempt = 0; attempt < 20 && !command; attempt += 1) {
      const polled = await client.callTool({
        name: 'office_editor_poll',
        arguments: { sessionId },
      })
      command = (
        polled.structuredContent as {
          commands: Array<{
            commandId: string
            operation: string
            arguments: Record<string, unknown>
          }>
        }
      ).commands[0]
      if (!command) await new Promise((resolve) => setTimeout(resolve, 10))
    }
    expect(command).toMatchObject({
      operation: 'docx.paragraph.set_heading_level',
      arguments: { target: { blockIndexes: [1] }, level: 3 },
    })

    await client.callTool({
      name: 'office_editor_acknowledge',
      arguments: {
        sessionId,
        commandId: command!.commandId,
        revision: 1,
        dirty: true,
        output: {
          summary: 'Updated heading level in 1 block(s)',
          matched: 1,
          changed: 1,
          skippedProtected: 0,
        },
      },
    })
    await expect(execution).resolves.toMatchObject({
      structuredContent: {
        ok: true,
        result: {
          revision: 1,
          operations: [
            {
              id: 'docx.paragraph.set_heading_level',
              result: {
                summary: 'Updated heading level in 1 block(s)',
                matched: 1,
                changed: 1,
                skippedProtected: 0,
              },
            },
          ],
        },
      },
    })
  })

  it('queues DOCX block deletion with canonical target arguments and returns the renderer output', async () => {
    const client = await connectClient()
    const created = await client.callTool({
      name: 'office_create_session',
      arguments: { format: 'docx' },
    })
    const sessionId = (created.structuredContent as { session: { id: string } }).session.id
    await client.callTool({
      name: 'office_editor_poll',
      arguments: { sessionId, selection: { from: 1, to: 1 } },
    })

    const execution = executeOperation(client, sessionId, 0, 'docx.block.delete', {
      target: { blockIndexes: [2] },
    })

    let command:
      { commandId: string; operation: string; arguments: Record<string, unknown> } | undefined
    for (let attempt = 0; attempt < 20 && !command; attempt += 1) {
      const polled = await client.callTool({
        name: 'office_editor_poll',
        arguments: { sessionId },
      })
      command = (
        polled.structuredContent as {
          commands: Array<{
            commandId: string
            operation: string
            arguments: Record<string, unknown>
          }>
        }
      ).commands[0]
      if (!command) await new Promise((resolve) => setTimeout(resolve, 10))
    }
    expect(command).toMatchObject({
      operation: 'docx.block.delete',
      arguments: { target: { blockIndexes: [2] } },
    })

    await client.callTool({
      name: 'office_editor_acknowledge',
      arguments: {
        sessionId,
        commandId: command!.commandId,
        revision: 1,
        dirty: true,
        output: {
          summary: 'Deleted 1 block(s)',
          matched: 1,
          changed: 1,
          skippedProtected: 0,
          skippedDeleted: 0,
        },
      },
    })
    await expect(execution).resolves.toMatchObject({
      structuredContent: {
        ok: true,
        result: {
          revision: 1,
          operations: [
            {
              id: 'docx.block.delete',
              result: {
                summary: 'Deleted 1 block(s)',
                matched: 1,
                changed: 1,
                skippedProtected: 0,
                skippedDeleted: 0,
              },
            },
          ],
        },
      },
    })
  })

  it('queues DOCX list application with canonical target arguments and returns the renderer output', async () => {
    const client = await connectClient()
    const created = await client.callTool({
      name: 'office_create_session',
      arguments: { format: 'docx' },
    })
    const sessionId = (created.structuredContent as { session: { id: string } }).session.id
    await client.callTool({
      name: 'office_editor_poll',
      arguments: { sessionId, selection: { from: 1, to: 1 } },
    })

    const execution = executeOperation(client, sessionId, 0, 'docx.list.apply', {
      target: { blockIndexes: [2] },
      kind: 'bullet',
    })

    let command:
      { commandId: string; operation: string; arguments: Record<string, unknown> } | undefined
    for (let attempt = 0; attempt < 20 && !command; attempt += 1) {
      const polled = await client.callTool({
        name: 'office_editor_poll',
        arguments: { sessionId },
      })
      command = (
        polled.structuredContent as {
          commands: Array<{
            commandId: string
            operation: string
            arguments: Record<string, unknown>
          }>
        }
      ).commands[0]
      if (!command) await new Promise((resolve) => setTimeout(resolve, 10))
    }
    expect(command).toMatchObject({
      operation: 'docx.list.apply',
      arguments: { target: { blockIndexes: [2] }, kind: 'bullet' },
    })

    await client.callTool({
      name: 'office_editor_acknowledge',
      arguments: {
        sessionId,
        commandId: command!.commandId,
        revision: 1,
        dirty: true,
        output: {
          summary: 'Converted 1 block(s) to list items',
          matched: 1,
          changed: 1,
          skippedProtected: 0,
        },
      },
    })
    await expect(execution).resolves.toMatchObject({
      structuredContent: {
        ok: true,
        result: {
          revision: 1,
          operations: [
            {
              id: 'docx.list.apply',
              result: {
                summary: 'Converted 1 block(s) to list items',
                matched: 1,
                changed: 1,
                skippedProtected: 0,
              },
            },
          ],
        },
      },
    })
  })

  it('queues DOCX list removal with canonical target arguments and returns the renderer output', async () => {
    const client = await connectClient()
    const created = await client.callTool({
      name: 'office_create_session',
      arguments: { format: 'docx' },
    })
    const sessionId = (created.structuredContent as { session: { id: string } }).session.id
    await client.callTool({
      name: 'office_editor_poll',
      arguments: { sessionId, selection: { from: 1, to: 1 } },
    })

    const execution = executeOperation(client, sessionId, 0, 'docx.list.remove', {
      target: { blockIndexes: [2] },
    })

    let command:
      { commandId: string; operation: string; arguments: Record<string, unknown> } | undefined
    for (let attempt = 0; attempt < 20 && !command; attempt += 1) {
      const polled = await client.callTool({
        name: 'office_editor_poll',
        arguments: { sessionId },
      })
      command = (
        polled.structuredContent as {
          commands: Array<{
            commandId: string
            operation: string
            arguments: Record<string, unknown>
          }>
        }
      ).commands[0]
      if (!command) await new Promise((resolve) => setTimeout(resolve, 10))
    }
    expect(command).toMatchObject({
      operation: 'docx.list.remove',
      arguments: { target: { blockIndexes: [2] } },
    })

    await client.callTool({
      name: 'office_editor_acknowledge',
      arguments: {
        sessionId,
        commandId: command!.commandId,
        revision: 1,
        dirty: true,
        output: {
          summary: 'Converted 1 list item(s) to body text',
          matched: 1,
          changed: 1,
          skippedProtected: 0,
        },
      },
    })
    await expect(execution).resolves.toMatchObject({
      structuredContent: {
        ok: true,
        result: {
          revision: 1,
          operations: [
            {
              id: 'docx.list.remove',
              result: {
                summary: 'Converted 1 list item(s) to body text',
                matched: 1,
                changed: 1,
                skippedProtected: 0,
              },
            },
          ],
        },
      },
    })
  })

  it('queues canonical Markdown selection replacement', async () => {
    await executeAndAcknowledgeMarkdownOperation(
      'markdown.text.replace_selection',
      { text: 'hello' },
      {},
    )
  })

  it('queues canonical Markdown save', async () => {
    await executeAndAcknowledgeMarkdownOperation(
      'markdown.document.save',
      {},
      { saved: true, fileName: 'notes.md' },
    )
  })

  it('queues canonical DOCX save, returns the persisted identity, and clears recovery', async () => {
    const stateDirectory = await mkdtemp(join(tmpdir(), 'tandemfolio-docx-save-'))
    temporaryDirectories.push(stateDirectory)
    const client = await connectClient(stateDirectory)
    const created = await client.callTool({
      name: 'office_create_session',
      arguments: { format: 'docx', resume: 'none' },
    })
    const sessionId = (created.structuredContent as { session: { id: string } }).session.id
    await client.callTool({ name: 'office_editor_poll', arguments: { sessionId } })
    const recoveryBytes = Buffer.from('PK\u0003\u0004unsaved-docx')
    const begun = await client.callTool({
      name: 'office_editor_begin_recovery',
      arguments: { sessionId, fileName: 'draft.docx', size: recoveryBytes.length },
    })
    const uploadId = (begun.structuredContent as { uploadId: string }).uploadId
    await client.callTool({
      name: 'office_editor_write_recovery_chunk',
      arguments: {
        sessionId,
        uploadId,
        offset: 0,
        data: recoveryBytes.toString('base64'),
      },
    })
    await client.callTool({
      name: 'office_editor_commit_recovery',
      arguments: { sessionId, uploadId },
    })
    const execution = executeOperation(client, sessionId, 0, 'docx.document.save', {})

    let command: { commandId: string; operation: string } | undefined
    for (let attempt = 0; attempt < 20 && !command; attempt += 1) {
      const polled = await client.callTool({
        name: 'office_editor_poll',
        arguments: { sessionId },
      })
      command = (
        polled.structuredContent as {
          commands: Array<{ commandId: string; operation: string }>
        }
      ).commands[0]
      if (!command) await new Promise((resolve) => setTimeout(resolve, 10))
    }
    expect(command).toBeDefined()

    await client.callTool({
      name: 'office_editor_acknowledge',
      arguments: {
        sessionId,
        commandId: command!.commandId,
        revision: 1,
        fileName: 'draft.docx',
        dirty: false,
        output: { saved: true, fileName: 'draft.docx' },
      },
    })
    await expect(execution).resolves.toMatchObject({
      structuredContent: {
        ok: true,
        result: {
          revision: 1,
          operations: [
            {
              id: 'docx.document.save',
              result: { saved: true, fileName: 'draft.docx' },
            },
          ],
        },
      },
    })
    expect(command!.operation).toBe('docx.document.save')

    const resumed = await client.callTool({
      name: 'office_create_session',
      arguments: { format: 'docx', resume: 'latest' },
    })
    expect(resumed.structuredContent).toMatchObject({ recoveryAvailable: false })
  })

  it('queues canonical XLSX save and returns the persisted workbook identity', async () => {
    const client = await connectClient()
    const created = await client.callTool({
      name: 'office_create_session',
      arguments: { format: 'xlsx' },
    })
    const sessionId = (created.structuredContent as { session: { id: string } }).session.id
    await client.callTool({ name: 'office_editor_poll', arguments: { sessionId } })

    const execution = executeOperation(client, sessionId, 0, 'xlsx.document.save', {})

    let command: { commandId: string; operation: string } | undefined
    for (let attempt = 0; attempt < 20 && !command; attempt += 1) {
      const polled = await client.callTool({
        name: 'office_editor_poll',
        arguments: { sessionId },
      })
      command = (
        polled.structuredContent as {
          commands: Array<{ commandId: string; operation: string }>
        }
      ).commands[0]
      if (!command) await new Promise((resolve) => setTimeout(resolve, 10))
    }
    expect(command).toBeDefined()

    await client.callTool({
      name: 'office_editor_acknowledge',
      arguments: {
        sessionId,
        commandId: command!.commandId,
        revision: 1,
        fileName: 'budget.xlsx',
        dirty: false,
        output: { saved: true, fileName: 'budget.xlsx' },
      },
    })
    await expect(execution).resolves.toMatchObject({
      structuredContent: {
        ok: true,
        result: {
          revision: 1,
          operations: [
            {
              id: 'xlsx.document.save',
              result: { saved: true, fileName: 'budget.xlsx' },
            },
          ],
        },
      },
    })
    expect(command!.operation).toBe('xlsx.document.save')
  })

  it('queues canonical PPTX save and returns the persisted presentation identity', async () => {
    const client = await connectClient()
    const created = await client.callTool({
      name: 'office_create_session',
      arguments: { format: 'pptx' },
    })
    const sessionId = (created.structuredContent as { session: { id: string } }).session.id
    await client.callTool({ name: 'office_editor_poll', arguments: { sessionId } })

    const execution = executeOperation(client, sessionId, 0, 'pptx.document.save', {})

    let command: { commandId: string; operation: string } | undefined
    for (let attempt = 0; attempt < 20 && !command; attempt += 1) {
      const polled = await client.callTool({
        name: 'office_editor_poll',
        arguments: { sessionId },
      })
      command = (
        polled.structuredContent as {
          commands: Array<{ commandId: string; operation: string }>
        }
      ).commands[0]
      if (!command) await new Promise((resolve) => setTimeout(resolve, 10))
    }
    expect(command).toBeDefined()

    await client.callTool({
      name: 'office_editor_acknowledge',
      arguments: {
        sessionId,
        commandId: command!.commandId,
        revision: 1,
        fileName: 'deck.pptx',
        dirty: false,
        output: { saved: true, fileName: 'deck.pptx' },
      },
    })
    await expect(execution).resolves.toMatchObject({
      structuredContent: {
        ok: true,
        result: {
          revision: 1,
          operations: [
            {
              id: 'pptx.document.save',
              result: { saved: true, fileName: 'deck.pptx' },
            },
          ],
        },
      },
    })
    expect(command!.operation).toBe('pptx.document.save')
  })

  it('queues canonical PDF save and returns the mounted renderer result', async () => {
    await executeAndAcknowledgePdfOperation('pdf.document.save', {}, { saved: true })
  })

  it('queues canonical PDF saved-annotation deletion', async () => {
    await executeAndAcknowledgePdfOperation(
      'pdf.annotation.delete_saved',
      {
        pageIndex: 0,
        objNum: 17,
        subtype: 'highlight',
        rect: [40, 220, 180, 242],
      },
      { deleted: 17 },
    )
  })

  it('queues canonical PDF undo', async () => {
    await executeAndAcknowledgePdfOperation('pdf.history.undo', {}, { undone: true })
  })

  it('rejects extra PDF save arguments before enqueueing', async () => {
    await expectPdfOperationRejectedBeforeEnqueue(
      'pdf.document.save',
      { fileName: 'invented.pdf' },
      'Invalid arguments for pdf.document.save: $.fileName is not allowed.',
    )
  })

  it('rejects an unsupported PDF annotation subtype before enqueueing', async () => {
    await expectPdfOperationRejectedBeforeEnqueue(
      'pdf.annotation.delete_saved',
      {
        pageIndex: 0,
        objNum: 17,
        subtype: 'square',
        rect: [40, 220, 180, 242],
      },
      'Invalid arguments for pdf.annotation.delete_saved: $.subtype must be one of "highlight", "underline", "strikeout".',
    )
  })

  it('queues canonical XLSX xlsx.cell.set_value', async () => {
    await executeAndAcknowledgeXlsxOperation(
      'xlsx.cell.set_value',
      { sheet: 'Budget', address: 'B2', value: 42 },
      { changed: 1, sheet: 'Budget', range: 'B2' },
    )
  })

  it('queues canonical XLSX xlsx.range.set_values', async () => {
    await executeAndAcknowledgeXlsxOperation(
      'xlsx.range.set_values',
      {
        sheet: 'Budget',
        range: 'A1:B2',
        values: [
          ['North', 10],
          ['South', 20],
        ],
      },
      { changed: 4, sheet: 'Budget', range: 'A1:B2' },
    )
  })

  it('queues XLSX xlsx.range.set_text_style with an explicit italic value', async () => {
    await executeAndAcknowledgeXlsxOperation(
      'xlsx.range.set_text_style',
      {
        sheet: 'Budget',
        range: 'B2:C3',
        style: { italic: true },
        fields: ['italic'],
      },
      { sheet: 'Budget', range: 'B2:C3', fields: ['italic'] },
    )
  })

  it('queues XLSX xlsx.range.set_text_style with an explicit bold value', async () => {
    await executeAndAcknowledgeXlsxOperation(
      'xlsx.range.set_text_style',
      {
        sheet: 'Budget',
        range: 'A1',
        style: { bold: false },
        fields: ['bold'],
      },
      { sheet: 'Budget', range: 'A1', fields: ['bold'] },
    )
  })

  it('queues XLSX xlsx.range.set_text_style with an explicit underline value', async () => {
    await executeAndAcknowledgeXlsxOperation(
      'xlsx.range.set_text_style',
      {
        sheet: 'Budget',
        range: 'A1:B2',
        style: { underline: 'double' },
        fields: ['underline'],
      },
      { sheet: 'Budget', range: 'A1:B2', fields: ['underline'] },
    )
  })

  it('queues XLSX xlsx.range.set_text_style with an explicit strike value', async () => {
    await executeAndAcknowledgeXlsxOperation(
      'xlsx.range.set_text_style',
      {
        sheet: 'Budget',
        range: 'A1',
        style: { strike: true },
        fields: ['strike'],
      },
      { sheet: 'Budget', range: 'A1', fields: ['strike'] },
    )
  })

  it.each([
    [
      'xlsx.range.set_alignment',
      {
        sheet: 'Budget',
        range: 'A1:B2',
        alignment: { horizontal: 'center', wrap: true, rotation: { mode: 'angle', degrees: 45 } },
        fields: ['horizontal', 'wrap', 'rotation'],
      },
      { sheet: 'Budget', range: 'A1:B2', fields: ['horizontal', 'wrap', 'rotation'] },
    ],
    [
      'xlsx.range.set_font',
      {
        sheet: 'Budget',
        range: 'A1:B2',
        font: { family: 'Aptos', size: 14, color: '#44546A' },
        fields: ['family', 'size', 'color'],
      },
      { sheet: 'Budget', range: 'A1:B2', fields: ['family', 'size', 'color'] },
    ],
    [
      'xlsx.range.set_fill',
      { sheet: 'Budget', range: 'A1:B2', color: '#DDEBF7' },
      { sheet: 'Budget', range: 'A1:B2' },
    ],
    [
      'xlsx.range.set_border',
      {
        sheet: 'Budget',
        range: 'A1:B2',
        border: { preset: 'outer', lineStyle: 'medium', color: '#4472C4' },
      },
      { sheet: 'Budget', range: 'A1:B2', preset: 'outer' },
    ],
    [
      'xlsx.range.apply_cell_style',
      { sheet: 'Budget', range: 'A1:B2', preset: 'input' },
      { sheet: 'Budget', range: 'A1:B2', preset: 'input' },
    ],
    [
      'xlsx.range.set_number_format',
      { sheet: 'Budget', range: 'A1:B2', pattern: '0.00%' },
      { sheet: 'Budget', range: 'A1:B2', pattern: '0.00%' },
    ],
    [
      'xlsx.range.merge',
      { sheet: 'Budget', range: 'A1:B2', mode: 'center' },
      { sheet: 'Budget', range: 'A1:B2', mode: 'center' },
    ],
    [
      'xlsx.range.clear',
      { sheet: 'Budget', range: 'A1:B2', scope: 'formats' },
      { sheet: 'Budget', range: 'A1:B2', scope: 'formats' },
    ],
    [
      'xlsx.range.fill',
      { sheet: 'Budget', range: 'A1:A3', direction: 'down' },
      { sheet: 'Budget', range: 'A1:A3', direction: 'down' },
    ],
    [
      'xlsx.range.sort',
      { sheet: 'Budget', range: 'A1:B4', direction: 'asc' },
      { sheet: 'Budget', range: 'A1:B4', direction: 'asc' },
    ],
    [
      'xlsx.range.sort_custom',
      {
        sheet: 'Budget',
        range: 'A1:C4',
        keys: [
          { column: 'C', direction: 'desc' },
          { column: 'A', direction: 'asc' },
        ],
        hasHeader: true,
      },
      {
        sheet: 'Budget',
        range: 'A1:C4',
        keys: [
          { column: 'C', direction: 'desc' },
          { column: 'A', direction: 'asc' },
        ],
        hasHeader: true,
      },
    ],
    [
      'xlsx.range.remove_duplicates',
      { sheet: 'Budget', range: 'A1:C4', hasHeader: true },
      { sheet: 'Budget', range: 'A1:C4', removed: 1 },
    ],
    [
      'xlsx.formula.insert_aggregate',
      { sheet: 'Budget', range: 'B2:C4', function: 'MAX' },
      {
        sheet: 'Budget',
        range: 'B2:C4',
        targetRange: 'B5:C5',
        function: 'MAX',
        inserted: 2,
      },
    ],
    ['xlsx.history.undo', {}, { undone: true }],
    ['xlsx.history.redo', {}, { redone: true }],
    [
      'xlsx.range.copy_values',
      {
        sourceSheet: 'Budget',
        sourceRange: 'A1:B2',
        destinationSheet: 'Archive',
        destinationRange: 'D4:E5',
      },
      {
        sourceSheet: 'Budget',
        sourceRange: 'A1:B2',
        destinationSheet: 'Archive',
        destinationRange: 'D4:E5',
        copied: 4,
      },
    ],
    [
      'xlsx.range.copy_formulas',
      {
        sourceSheet: 'Budget',
        sourceRange: 'A1:B2',
        destinationSheet: 'Archive',
        destinationRange: 'D4:E5',
      },
      {
        sourceSheet: 'Budget',
        sourceRange: 'A1:B2',
        destinationSheet: 'Archive',
        destinationRange: 'D4:E5',
        copied: 4,
      },
    ],
    [
      'xlsx.range.copy_formats',
      {
        sourceSheet: 'Budget',
        sourceRange: 'A1:B2',
        destinationSheet: 'Archive',
        destinationRange: 'D4:E5',
      },
      {
        sourceSheet: 'Budget',
        sourceRange: 'A1:B2',
        destinationSheet: 'Archive',
        destinationRange: 'D4:E5',
        copied: 4,
      },
    ],
    [
      'xlsx.range.copy_without_borders',
      {
        sourceSheet: 'Budget',
        sourceRange: 'A1:B2',
        destinationSheet: 'Archive',
        destinationRange: 'D4:E5',
      },
      {
        sourceSheet: 'Budget',
        sourceRange: 'A1:B2',
        destinationSheet: 'Archive',
        destinationRange: 'D4:E5',
        copied: 4,
      },
    ],
    [
      'xlsx.column.copy_widths',
      {
        sourceSheet: 'Budget',
        sourceColumn: 'A',
        destinationSheet: 'Archive',
        destinationColumn: 'D',
        count: 2,
      },
      {
        sourceSheet: 'Budget',
        sourceColumn: 'A',
        destinationSheet: 'Archive',
        destinationColumn: 'D',
        count: 2,
      },
    ],
    [
      'xlsx.range.flash_fill',
      { sheet: 'Budget', range: 'C1:C4' },
      { sheet: 'Budget', range: 'C1:C4', targetRange: 'C1:C4', filled: 2 },
    ],
    [
      'xlsx.range.text_to_columns',
      { sheet: 'Budget', range: 'A1:A4', delimiter: 'semicolon' },
      { sheet: 'Budget', range: 'A1:A4', delimiter: 'semicolon' },
    ],
    [
      'xlsx.row.set_height',
      { sheet: 'Budget', row: 3, count: 2, heightPoints: 24.75 },
      { sheet: 'Budget', row: 3, count: 2, heightPoints: 24.75 },
    ],
    [
      'xlsx.column.set_width',
      { sheet: 'Budget', column: 'C', count: 2, widthCharacters: 12.5 },
      { sheet: 'Budget', column: 'C', count: 2, widthCharacters: 12.4296875 },
    ],
    [
      'xlsx.hyperlink.set',
      { sheet: 'Budget', address: 'B2', target: 'https://example.com/report' },
      { sheet: 'Budget', address: 'B2', target: 'https://example.com/report' },
    ],
    [
      'xlsx.hyperlink.remove',
      { sheet: 'Budget', address: 'B2' },
      { sheet: 'Budget', address: 'B2', removed: true },
    ],
    [
      'xlsx.table.add',
      { sheet: 'Budget', range: 'A1:B3', style: 'TableStyleMedium4' },
      {
        sheet: 'Budget',
        range: 'A1:B3',
        name: 'Table1',
        style: 'TableStyleMedium4',
      },
    ],
    [
      'xlsx.range.set_protection',
      {
        sheet: 'Budget',
        range: 'C2:C3',
        protection: { locked: false, hidden: true },
        fields: ['locked', 'hidden'],
      },
      { sheet: 'Budget', range: 'C2:C3', fields: ['locked', 'hidden'] },
    ],
    [
      'xlsx.sheet.set_protection',
      { sheet: 'Budget', protected: true },
      { sheet: 'Budget', protected: true },
    ],
    [
      'xlsx.sheet.set_freeze',
      { sheet: 'Budget', frozenRows: 2, frozenColumns: 1 },
      { sheet: 'Budget', frozenRows: 2, frozenColumns: 1 },
    ],
    [
      'xlsx.sheet.set_fit_to_pages',
      { sheet: 'Budget', widthPages: 2, heightPages: 3 },
      { sheet: 'Budget', widthPages: 2, heightPages: 3 },
    ],
    [
      'xlsx.range.set_filter',
      { sheet: 'Budget', range: 'A1:C4', enabled: true },
      { sheet: 'Budget', range: 'A1:C4', enabled: true },
    ],
    [
      'xlsx.range.clear_filter_criteria',
      { sheet: 'Budget', range: 'A1:C4' },
      { sheet: 'Budget', range: 'A1:C4', cleared: true },
    ],
    [
      'xlsx.range.set_filter_values',
      {
        sheet: 'Budget',
        range: 'A1:C4',
        column: 'B',
        values: ['East', 'South'],
        includeBlank: false,
      },
      {
        sheet: 'Budget',
        range: 'A1:C4',
        column: 'B',
        selectedValues: 2,
        includeBlank: false,
      },
    ],
    [
      'xlsx.range.set_custom_filter',
      {
        sheet: 'Budget',
        range: 'A1:C21',
        column: 'B',
        conjunction: 'and',
        conditions: [
          { operator: 'greaterThanOrEqual', value: '10' },
          { operator: 'lessThan', value: '20' },
        ],
      },
      {
        sheet: 'Budget',
        range: 'A1:C21',
        column: 'B',
        conjunction: 'and',
        conditions: 2,
      },
    ],
    [
      'xlsx.sheet.set_formula_view',
      { sheet: 'Budget', enabled: true },
      { sheet: 'Budget', enabled: true },
    ],
    [
      'xlsx.sheet.set_gridlines',
      { sheet: 'Budget', visible: false },
      { sheet: 'Budget', visible: false },
    ],
    [
      'xlsx.sheet.set_page_margins',
      { sheet: 'Budget', margins: 'wide' },
      { sheet: 'Budget', margins: 'wide' },
    ],
    [
      'xlsx.sheet.set_page_orientation',
      { sheet: 'Budget', orientation: 'landscape' },
      { sheet: 'Budget', orientation: 'landscape' },
    ],
    [
      'xlsx.sheet.set_paper_size',
      { sheet: 'Budget', paperSize: 9 },
      { sheet: 'Budget', paperSize: 9 },
    ],
    [
      'xlsx.sheet.set_print_gridlines',
      { sheet: 'Budget', enabled: true },
      { sheet: 'Budget', enabled: true },
    ],
    [
      'xlsx.sheet.set_print_headings',
      { sheet: 'Budget', enabled: true },
      { sheet: 'Budget', enabled: true },
    ],
    [
      'xlsx.sheet.set_print_scale',
      { sheet: 'Budget', scalePercent: 80 },
      { sheet: 'Budget', scalePercent: 80 },
    ],
    [
      'xlsx.sheet.set_print_area',
      { sheet: 'Budget', range: 'B2:D8' },
      { sheet: 'Budget', range: 'B2:D8' },
    ],
    [
      'xlsx.sheet.set_print_area',
      { sheet: 'Budget', range: null },
      { sheet: 'Budget', range: null },
    ],
    [
      'xlsx.sheet.set_print_titles',
      { sheet: 'Budget', rows: '2:8' },
      { sheet: 'Budget', rows: '2:8' },
    ],
    [
      'xlsx.sheet.set_print_titles',
      { sheet: 'Budget', rows: null },
      { sheet: 'Budget', rows: null },
    ],
  ] as const)(
    'queues XLSX %s through the generated Broker contract',
    async (operation, args, output) => {
      await executeAndAcknowledgeXlsxOperation(operation, args, output)
    },
  )

  it('queues canonical XLSX xlsx.row.insert', async () => {
    await executeAndAcknowledgeXlsxOperation(
      'xlsx.row.insert',
      { sheet: 'Budget', row: 3, count: 2 },
      { sheet: 'Budget', row: 3, count: 2 },
    )
  })

  it('queues canonical XLSX xlsx.row.delete', async () => {
    await executeAndAcknowledgeXlsxOperation(
      'xlsx.row.delete',
      { sheet: 'Budget', row: 4, count: 3 },
      { sheet: 'Budget', row: 4, count: 3 },
    )
  })

  it('queues canonical XLSX xlsx.column.insert', async () => {
    await executeAndAcknowledgeXlsxOperation(
      'xlsx.column.insert',
      { sheet: 'Budget', column: 'C', count: 2 },
      { sheet: 'Budget', column: 'C', count: 2 },
    )
  })

  it('queues canonical XLSX xlsx.column.delete', async () => {
    await executeAndAcknowledgeXlsxOperation(
      'xlsx.column.delete',
      { sheet: 'Budget', column: 'D', count: 2 },
      { sheet: 'Budget', column: 'D', count: 2 },
    )
  })

  it('queues canonical XLSX xlsx.sheet.add', async () => {
    await executeAndAcknowledgeXlsxOperation(
      'xlsx.sheet.add',
      { name: 'Forecast' },
      { name: 'Forecast' },
    )
  })

  it('queues canonical XLSX xlsx.sheet.rename', async () => {
    await executeAndAcknowledgeXlsxOperation(
      'xlsx.sheet.rename',
      { sheet: 'Budget', name: 'Plan' },
      { sheet: 'Budget', name: 'Plan' },
    )
  })

  it('queues canonical XLSX xlsx.sheet.delete', async () => {
    await executeAndAcknowledgeXlsxOperation(
      'xlsx.sheet.delete',
      { sheet: 'Forecast' },
      { sheet: 'Forecast' },
    )
  })

  it('queues canonical XLSX xlsx.sheet.move', async () => {
    await executeAndAcknowledgeXlsxOperation(
      'xlsx.sheet.move',
      { sheet: 'Forecast', position: 1 },
      { sheet: 'Forecast', position: 1 },
    )
  })

  it.each([
    [
      'xlsx.cell.set_value',
      { sheet: 'Budget', address: 'A1', value: { formula: '=1+1' } },
      'Invalid arguments for xlsx.cell.set_value: $.value must be string or number or boolean or null.',
    ],
    [
      'xlsx.range.set_values',
      { sheet: 'Budget', range: 'A1', values: [[]] },
      'Invalid arguments for xlsx.range.set_values: $.values[0] must contain at least 1 item.',
    ],
    [
      'xlsx.range.set_text_style',
      {
        sheet: 'Budget',
        range: 'A1',
        style: { underline: 'triple' },
        fields: ['underline'],
      },
      'Invalid arguments for xlsx.range.set_text_style: $.style.underline must be one of "none", "single", "double".',
    ],
    [
      'xlsx.range.set_alignment',
      {
        sheet: 'Budget',
        range: 'A1',
        alignment: { indent: 251 },
        fields: ['indent'],
      },
      'Invalid arguments for xlsx.range.set_alignment: $.alignment.indent must be less than or equal to 250.',
    ],
    [
      'xlsx.range.set_font',
      { sheet: 'Budget', range: 'A1', font: { size: 410 }, fields: ['size'] },
      'Invalid arguments for xlsx.range.set_font: $.font.size must be less than or equal to 409.',
    ],
    [
      'xlsx.range.set_fill',
      { sheet: 'Budget', range: 'A1', color: 42 },
      'Invalid arguments for xlsx.range.set_fill: $.color must be string or null.',
    ],
    [
      'xlsx.range.set_border',
      { sheet: 'Budget', range: 'A1', border: { preset: 'diagonal' } },
      'Invalid arguments for xlsx.range.set_border: $.border.preset must be one of "none", "all", "outer", "thick-outer", "top", "bottom", "left", "right".',
    ],
    [
      'xlsx.range.apply_cell_style',
      { sheet: 'Budget', range: 'A1', preset: 'unknown' },
      'Invalid arguments for xlsx.range.apply_cell_style: $.preset must be one of "good", "bad", "neutral", "input", "output", "calculation", "warning-text", "title", "heading-1", "heading-2", "total", "accent1-20", "accent1-40", "accent1".',
    ],
    [
      'xlsx.range.set_number_format',
      { sheet: 'Budget', range: 'A1', pattern: 42 },
      'Invalid arguments for xlsx.range.set_number_format: $.pattern must be a string.',
    ],
    [
      'xlsx.range.merge',
      { sheet: 'Budget', range: 'A1:B2', mode: 'vertical' },
      'Invalid arguments for xlsx.range.merge: $.mode must be one of "cells", "across", "center", "unmerge".',
    ],
    [
      'xlsx.range.clear',
      { sheet: 'Budget', range: 'A1:B2', scope: 'values' },
      'Invalid arguments for xlsx.range.clear: $.scope must be one of "contents", "formats", "all".',
    ],
    [
      'xlsx.range.fill',
      { sheet: 'Budget', range: 'A1:B2', direction: 'up' },
      'Invalid arguments for xlsx.range.fill: $.direction must be one of "down", "right".',
    ],
    [
      'xlsx.range.sort',
      { sheet: 'Budget', range: 'A1:B2', direction: 'sideways' },
      'Invalid arguments for xlsx.range.sort: $.direction must be one of "asc", "desc".',
    ],
    [
      'xlsx.range.sort_custom',
      {
        sheet: 'Budget',
        range: 'A1:B2',
        keys: [{ column: 'A', direction: 'sideways' }],
        hasHeader: false,
      },
      'Invalid arguments for xlsx.range.sort_custom: $.keys[0].direction must be one of "asc", "desc".',
    ],
    [
      'xlsx.range.remove_duplicates',
      { sheet: 'Budget', range: 'A1:B2', hasHeader: 'yes' },
      'Invalid arguments for xlsx.range.remove_duplicates: $.hasHeader must be a boolean.',
    ],
    [
      'xlsx.formula.insert_aggregate',
      { sheet: 'Budget', range: 'A1:B2', function: 'MEDIAN' },
      'Invalid arguments for xlsx.formula.insert_aggregate: $.function must be one of "SUM", "AVERAGE", "COUNT", "MAX", "MIN".',
    ],
    [
      'xlsx.history.undo',
      { force: true },
      'Invalid arguments for xlsx.history.undo: $.force is not allowed.',
    ],
    [
      'xlsx.history.redo',
      { force: true },
      'Invalid arguments for xlsx.history.redo: $.force is not allowed.',
    ],
    [
      'xlsx.range.copy_values',
      {
        sourceSheet: 'Budget',
        sourceRange: 'A1:B2',
        destinationSheet: 'Archive',
        destinationRange: 42,
      },
      'Invalid arguments for xlsx.range.copy_values: $.destinationRange must be a string.',
    ],
    [
      'xlsx.range.copy_formulas',
      {
        sourceSheet: 'Budget',
        sourceRange: 'A1:B2',
        destinationSheet: 'Archive',
        destinationRange: 42,
      },
      'Invalid arguments for xlsx.range.copy_formulas: $.destinationRange must be a string.',
    ],
    [
      'xlsx.range.copy_formats',
      {
        sourceSheet: 'Budget',
        sourceRange: 'A1:B2',
        destinationSheet: 'Archive',
        destinationRange: 42,
      },
      'Invalid arguments for xlsx.range.copy_formats: $.destinationRange must be a string.',
    ],
    [
      'xlsx.range.copy_without_borders',
      {
        sourceSheet: 'Budget',
        sourceRange: 'A1:B2',
        destinationSheet: 'Archive',
        destinationRange: 42,
      },
      'Invalid arguments for xlsx.range.copy_without_borders: $.destinationRange must be a string.',
    ],
    [
      'xlsx.column.copy_widths',
      {
        sourceSheet: 'Budget',
        sourceColumn: 'A',
        destinationSheet: 'Archive',
        destinationColumn: 'D',
        count: 10_001,
      },
      'Invalid arguments for xlsx.column.copy_widths: $.count must be less than or equal to 10000.',
    ],
    [
      'xlsx.range.flash_fill',
      { sheet: 'Budget', range: 'C1:C4', mode: 'series' },
      'Invalid arguments for xlsx.range.flash_fill: $.mode is not allowed.',
    ],
    [
      'xlsx.range.text_to_columns',
      { sheet: 'Budget', range: 'A1:A4', delimiter: 'pipe' },
      'Invalid arguments for xlsx.range.text_to_columns: $.delimiter must be one of "tab", "comma", "semicolon", "space".',
    ],
    [
      'xlsx.row.set_height',
      { sheet: 'Budget', row: 1, count: 1, heightPoints: 410 },
      'Invalid arguments for xlsx.row.set_height: $.heightPoints must be less than or equal to 409.5.',
    ],
    [
      'xlsx.column.set_width',
      { sheet: 'Budget', column: 'A', count: 1, widthCharacters: 256 },
      'Invalid arguments for xlsx.column.set_width: $.widthCharacters must be less than or equal to 255.',
    ],
    [
      'xlsx.table.add',
      { sheet: 'Budget', range: 'A1:B2', style: 'TableStyleRainbow99' },
      'Invalid arguments for xlsx.table.add: $.style must be one of "TableStyleLight1", "TableStyleLight9", "TableStyleMedium2", "TableStyleMedium4", "TableStyleMedium7", "TableStyleDark2".',
    ],
    [
      'xlsx.sheet.set_protection',
      { sheet: 'Budget', protected: 'yes' },
      'Invalid arguments for xlsx.sheet.set_protection: $.protected must be a boolean.',
    ],
    [
      'xlsx.sheet.set_freeze',
      { sheet: 'Budget', frozenRows: 1, frozenColumns: 16_384 },
      'Invalid arguments for xlsx.sheet.set_freeze: $.frozenColumns must be less than or equal to 16383.',
    ],
    [
      'xlsx.sheet.set_fit_to_pages',
      { sheet: 'Budget', widthPages: -1, heightPages: 1 },
      'Invalid arguments for xlsx.sheet.set_fit_to_pages: $.widthPages must be greater than or equal to 0.',
    ],
    [
      'xlsx.sheet.set_formula_view',
      { sheet: 'Budget', enabled: 'yes' },
      'Invalid arguments for xlsx.sheet.set_formula_view: $.enabled must be a boolean.',
    ],
    [
      'xlsx.range.set_filter',
      { sheet: 'Budget', range: 'A1:C4', enabled: 'yes' },
      'Invalid arguments for xlsx.range.set_filter: $.enabled must be a boolean.',
    ],
    [
      'xlsx.range.set_custom_filter',
      {
        sheet: 'Budget',
        range: 'A1:C4',
        column: 'B',
        conjunction: 'or',
        conditions: [
          { operator: 'equal', value: 'A' },
          { operator: 'equal', value: 'B' },
          { operator: 'equal', value: 'C' },
        ],
      },
      'Invalid arguments for xlsx.range.set_custom_filter: $.conditions must contain at most 2 items.',
    ],
    [
      'xlsx.sheet.set_gridlines',
      { sheet: 'Budget', visible: 'hidden' },
      'Invalid arguments for xlsx.sheet.set_gridlines: $.visible must be a boolean.',
    ],
    [
      'xlsx.sheet.set_page_margins',
      { sheet: 'Budget', margins: 'custom' },
      'Invalid arguments for xlsx.sheet.set_page_margins: $.margins must be one of "normal", "wide", "narrow".',
    ],
    [
      'xlsx.sheet.set_page_orientation',
      { sheet: 'Budget', orientation: 'diagonal' },
      'Invalid arguments for xlsx.sheet.set_page_orientation: $.orientation must be one of "portrait", "landscape".',
    ],
    [
      'xlsx.sheet.set_paper_size',
      { sheet: 'Budget', paperSize: 10 },
      'Invalid arguments for xlsx.sheet.set_paper_size: $.paperSize must be one of 1, 3, 5, 7, 8, 9, 11.',
    ],
    [
      'xlsx.sheet.set_print_gridlines',
      { sheet: 'Budget', enabled: 'yes' },
      'Invalid arguments for xlsx.sheet.set_print_gridlines: $.enabled must be a boolean.',
    ],
    [
      'xlsx.sheet.set_print_headings',
      { sheet: 'Budget', enabled: 'yes' },
      'Invalid arguments for xlsx.sheet.set_print_headings: $.enabled must be a boolean.',
    ],
    [
      'xlsx.sheet.set_print_scale',
      { sheet: 'Budget', scalePercent: 9 },
      'Invalid arguments for xlsx.sheet.set_print_scale: $.scalePercent must be greater than or equal to 10.',
    ],
    [
      'xlsx.sheet.set_print_area',
      { sheet: 'Budget', range: false },
      'Invalid arguments for xlsx.sheet.set_print_area: $.range must be string or null.',
    ],
    [
      'xlsx.sheet.set_print_titles',
      { sheet: 'Budget', rows: false },
      'Invalid arguments for xlsx.sheet.set_print_titles: $.rows must be string or null.',
    ],
    [
      'xlsx.row.insert',
      { sheet: 'Budget', row: 0, count: 1 },
      'Invalid arguments for xlsx.row.insert: $.row must be greater than or equal to 1.',
    ],
    [
      'xlsx.row.delete',
      { sheet: 'Budget', row: 1, count: 10_001 },
      'Invalid arguments for xlsx.row.delete: $.count must be less than or equal to 10000.',
    ],
    [
      'xlsx.column.insert',
      { sheet: 'Budget', column: 'A', count: 0 },
      'Invalid arguments for xlsx.column.insert: $.count must be greater than or equal to 1.',
    ],
    [
      'xlsx.column.delete',
      { sheet: 'Budget', column: 'A', count: 0 },
      'Invalid arguments for xlsx.column.delete: $.count must be greater than or equal to 1.',
    ],
    [
      'xlsx.sheet.add',
      { name: 42 },
      'Invalid arguments for xlsx.sheet.add: $.name must be a string.',
    ],
    [
      'xlsx.sheet.rename',
      { sheet: 'Budget', name: 42 },
      'Invalid arguments for xlsx.sheet.rename: $.name must be a string.',
    ],
    [
      'xlsx.sheet.delete',
      { sheet: 42 },
      'Invalid arguments for xlsx.sheet.delete: $.sheet must be a string.',
    ],
    [
      'xlsx.sheet.move',
      { sheet: 'Budget', position: 0 },
      'Invalid arguments for xlsx.sheet.move: $.position must be greater than or equal to 1.',
    ],
  ] as const)(
    'rejects malformed %s arguments before enqueueing',
    async (operation, args, message) => {
      await expectXlsxOperationRejectedBeforeEnqueue(operation, args, message)
    },
  )

  it('rejects invalid DOCX save arguments before enqueueing', async () => {
    const client = await connectClient()
    const created = await client.callTool({
      name: 'office_create_session',
      arguments: { format: 'docx' },
    })
    const sessionId = (created.structuredContent as { session: { id: string } }).session.id
    await client.callTool({ name: 'office_editor_poll', arguments: { sessionId } })

    const execution = executeOperation(client, sessionId, 0, 'docx.document.save', {
      force: true,
    })

    await new Promise((resolve) => setTimeout(resolve, 20))
    const polled = await client.callTool({
      name: 'office_editor_poll',
      arguments: { sessionId },
    })
    const command = (
      polled.structuredContent as {
        commands: Array<{ commandId: string }>
      }
    ).commands[0]

    await expect(execution).resolves.toMatchObject({
      isError: true,
      structuredContent: {
        ok: false,
        error: 'operation_schema_invalid',
        message: 'Invalid arguments for docx.document.save: $.force is not allowed.',
      },
    })
    expect(command).toBeUndefined()
  })

  it('rejects invalid XLSX save arguments before enqueueing', async () => {
    const client = await connectClient()
    const created = await client.callTool({
      name: 'office_create_session',
      arguments: { format: 'xlsx' },
    })
    const sessionId = (created.structuredContent as { session: { id: string } }).session.id
    await client.callTool({ name: 'office_editor_poll', arguments: { sessionId } })

    const execution = await executeOperation(client, sessionId, 0, 'xlsx.document.save', {
      force: true,
    })

    expect(execution).toMatchObject({
      isError: true,
      structuredContent: {
        ok: false,
        error: 'operation_schema_invalid',
        message: 'Invalid arguments for xlsx.document.save: $.force is not allowed.',
      },
    })

    const polled = await client.callTool({
      name: 'office_editor_poll',
      arguments: { sessionId },
    })
    expect(
      (polled.structuredContent as { commands: Array<{ commandId: string }> }).commands[0],
    ).toBeUndefined()

    const context = await client.callTool({
      name: 'office_get_context',
      arguments: { sessionId },
    })
    expect(context.structuredContent).toMatchObject({ session: { revision: 0 } })
  })

  it('rejects invalid PPTX save arguments before enqueueing', async () => {
    const client = await connectClient()
    const created = await client.callTool({
      name: 'office_create_session',
      arguments: { format: 'pptx' },
    })
    const sessionId = (created.structuredContent as { session: { id: string } }).session.id
    await client.callTool({ name: 'office_editor_poll', arguments: { sessionId } })

    const execution = await executeOperation(client, sessionId, 0, 'pptx.document.save', {
      force: true,
    })

    expect(execution).toMatchObject({
      isError: true,
      structuredContent: {
        ok: false,
        error: 'operation_schema_invalid',
        message: 'Invalid arguments for pptx.document.save: $.force is not allowed.',
      },
    })

    const polled = await client.callTool({
      name: 'office_editor_poll',
      arguments: { sessionId },
    })
    expect(
      (polled.structuredContent as { commands: Array<{ commandId: string }> }).commands[0],
    ).toBeUndefined()

    const context = await client.callTool({
      name: 'office_get_context',
      arguments: { sessionId },
    })
    expect(context.structuredContent).toMatchObject({ session: { revision: 0 } })
  })

  it('rejects invalid Markdown insert arguments before enqueueing', async () => {
    const client = await connectClient()
    const created = await client.callTool({
      name: 'office_create_session',
      arguments: { format: 'markdown' },
    })
    const sessionId = (created.structuredContent as { session: { id: string } }).session.id
    await client.callTool({ name: 'office_editor_poll', arguments: { sessionId } })

    const execution = executeOperation(client, sessionId, 0, 'markdown.text.insert', { text: 42 })

    await new Promise((resolve) => setTimeout(resolve, 20))
    const polled = await client.callTool({
      name: 'office_editor_poll',
      arguments: { sessionId },
    })
    const command = (
      polled.structuredContent as {
        commands: Array<{ commandId: string }>
      }
    ).commands[0]
    if (command) {
      await client.callTool({
        name: 'office_editor_acknowledge',
        arguments: {
          sessionId,
          commandId: command.commandId,
          revision: 0,
          ok: false,
          error: 'invalid_arguments',
          message: 'Invalid arguments for markdown.text.insert: $.text must be a string.',
        },
      })
    }

    await expect(execution).resolves.toMatchObject({
      isError: true,
      structuredContent: {
        ok: false,
        error: 'operation_schema_invalid',
        message: 'Invalid arguments for markdown.text.insert: $.text must be a string.',
      },
    })
    expect(command).toBeUndefined()

    const context = await client.callTool({
      name: 'office_get_context',
      arguments: { sessionId },
    })
    expect(context.structuredContent).toMatchObject({
      session: { revision: 0 },
    })
  })

  it('rejects invalid DOCX insert arguments before enqueueing', async () => {
    const client = await connectClient()
    const created = await client.callTool({
      name: 'office_create_session',
      arguments: { format: 'docx' },
    })
    const sessionId = (created.structuredContent as { session: { id: string } }).session.id
    await client.callTool({ name: 'office_editor_poll', arguments: { sessionId } })

    const execution = executeOperation(client, sessionId, 0, 'docx.text.insert', { text: 42 })

    await new Promise((resolve) => setTimeout(resolve, 20))
    const polled = await client.callTool({
      name: 'office_editor_poll',
      arguments: { sessionId },
    })
    const command = (
      polled.structuredContent as {
        commands: Array<{ commandId: string }>
      }
    ).commands[0]
    if (command) {
      await client.callTool({
        name: 'office_editor_acknowledge',
        arguments: {
          sessionId,
          commandId: command.commandId,
          revision: 0,
          ok: false,
          error: 'invalid_arguments',
          message: 'Invalid arguments for docx.text.insert: $.text must be a string.',
        },
      })
    }

    await expect(execution).resolves.toMatchObject({
      isError: true,
      structuredContent: {
        ok: false,
        error: 'operation_schema_invalid',
        message: 'Invalid arguments for docx.text.insert: $.text must be a string.',
      },
    })
    expect(command).toBeUndefined()

    const context = await client.callTool({
      name: 'office_get_context',
      arguments: { sessionId },
    })
    expect(context.structuredContent).toMatchObject({
      session: { revision: 0 },
    })
  })

  it('rejects invalid DOCX replacement arguments before enqueueing', async () => {
    const client = await connectClient()
    const created = await client.callTool({
      name: 'office_create_session',
      arguments: { format: 'docx' },
    })
    const sessionId = (created.structuredContent as { session: { id: string } }).session.id
    await client.callTool({ name: 'office_editor_poll', arguments: { sessionId } })

    const execution = executeOperation(client, sessionId, 0, 'docx.text.replace_selection', {
      text: 42,
    })

    await new Promise((resolve) => setTimeout(resolve, 20))
    const polled = await client.callTool({
      name: 'office_editor_poll',
      arguments: { sessionId },
    })
    const command = (
      polled.structuredContent as {
        commands: Array<{ commandId: string }>
      }
    ).commands[0]
    if (command) {
      await client.callTool({
        name: 'office_editor_acknowledge',
        arguments: {
          sessionId,
          commandId: command.commandId,
          revision: 0,
          ok: false,
          error: 'invalid_arguments',
          message: 'Invalid arguments for docx.text.replace_selection: $.text must be a string.',
        },
      })
    }

    await expect(execution).resolves.toMatchObject({
      isError: true,
      structuredContent: {
        ok: false,
        error: 'operation_schema_invalid',
        message: 'Invalid arguments for docx.text.replace_selection: $.text must be a string.',
      },
    })
    expect(command).toBeUndefined()

    const context = await client.callTool({
      name: 'office_get_context',
      arguments: { sessionId },
    })
    expect(context.structuredContent).toMatchObject({
      session: { revision: 0 },
    })
  })

  it('rejects invalid DOCX replace-all arguments before enqueueing', async () => {
    const client = await connectClient()
    const created = await client.callTool({
      name: 'office_create_session',
      arguments: { format: 'docx' },
    })
    const sessionId = (created.structuredContent as { session: { id: string } }).session.id
    await client.callTool({ name: 'office_editor_poll', arguments: { sessionId } })

    const execution = executeOperation(client, sessionId, 0, 'docx.text.replace_all', {
      containsText: 'old',
      replaceText: 'new',
      scope: 'document',
    })

    await new Promise((resolve) => setTimeout(resolve, 20))
    const polled = await client.callTool({
      name: 'office_editor_poll',
      arguments: { sessionId },
    })
    const command = (
      polled.structuredContent as {
        commands: Array<{ commandId: string }>
      }
    ).commands[0]

    await expect(execution).resolves.toMatchObject({
      isError: true,
      structuredContent: {
        ok: false,
        error: 'operation_schema_invalid',
        message: 'Invalid arguments for docx.text.replace_all: $.scope is not allowed.',
      },
    })
    expect(command).toBeUndefined()

    const context = await client.callTool({
      name: 'office_get_context',
      arguments: { sessionId },
    })
    expect(context.structuredContent).toMatchObject({
      session: { revision: 0 },
    })
  })

  it('rejects an invalid DOCX heading target enum before enqueueing', async () => {
    const client = await connectClient()
    const created = await client.callTool({
      name: 'office_create_session',
      arguments: { format: 'docx' },
    })
    const sessionId = (created.structuredContent as { session: { id: string } }).session.id
    await client.callTool({ name: 'office_editor_poll', arguments: { sessionId } })

    const execution = executeOperation(client, sessionId, 0, 'docx.paragraph.set_heading_level', {
      target: { nodeType: 'table' },
      level: 2,
    })

    await new Promise((resolve) => setTimeout(resolve, 20))
    const polled = await client.callTool({
      name: 'office_editor_poll',
      arguments: { sessionId },
    })
    const command = (
      polled.structuredContent as {
        commands: Array<{ commandId: string }>
      }
    ).commands[0]

    await expect(execution).resolves.toMatchObject({
      isError: true,
      structuredContent: {
        ok: false,
        error: 'operation_schema_invalid',
        message:
          'Invalid arguments for docx.paragraph.set_heading_level: $.target.nodeType must be one of "docHeading", "docParagraph", "docListItem", "image".',
      },
    })
    expect(command).toBeUndefined()

    const context = await client.callTool({
      name: 'office_get_context',
      arguments: { sessionId },
    })
    expect(context.structuredContent).toMatchObject({ session: { revision: 0 } })
  })

  it('rejects an unknown DOCX block-delete target field before enqueueing', async () => {
    const client = await connectClient()
    const created = await client.callTool({
      name: 'office_create_session',
      arguments: { format: 'docx' },
    })
    const sessionId = (created.structuredContent as { session: { id: string } }).session.id
    await client.callTool({ name: 'office_editor_poll', arguments: { sessionId } })

    const execution = executeOperation(client, sessionId, 0, 'docx.block.delete', {
      target: { blockIndexes: [0], all: true },
    })

    await new Promise((resolve) => setTimeout(resolve, 20))
    const polled = await client.callTool({
      name: 'office_editor_poll',
      arguments: { sessionId },
    })
    const command = (
      polled.structuredContent as {
        commands: Array<{ commandId: string }>
      }
    ).commands[0]

    await expect(execution).resolves.toMatchObject({
      isError: true,
      structuredContent: {
        ok: false,
        error: 'operation_schema_invalid',
        message: 'Invalid arguments for docx.block.delete: $.target.all is not allowed.',
      },
    })
    expect(command).toBeUndefined()

    const context = await client.callTool({
      name: 'office_get_context',
      arguments: { sessionId },
    })
    expect(context.structuredContent).toMatchObject({ session: { revision: 0 } })
  })

  it('rejects an invalid DOCX list kind before enqueueing', async () => {
    const client = await connectClient()
    const created = await client.callTool({
      name: 'office_create_session',
      arguments: { format: 'docx' },
    })
    const sessionId = (created.structuredContent as { session: { id: string } }).session.id
    await client.callTool({ name: 'office_editor_poll', arguments: { sessionId } })

    const execution = executeOperation(client, sessionId, 0, 'docx.list.apply', {
      target: { blockIndexes: [0] },
      kind: 'NUMBERED_DECIMAL_ALPHA_ROMAN',
    })

    await new Promise((resolve) => setTimeout(resolve, 20))
    const polled = await client.callTool({
      name: 'office_editor_poll',
      arguments: { sessionId },
    })
    const command = (
      polled.structuredContent as {
        commands: Array<{ commandId: string }>
      }
    ).commands[0]

    await expect(execution).resolves.toMatchObject({
      isError: true,
      structuredContent: {
        ok: false,
        error: 'operation_schema_invalid',
        message:
          'Invalid arguments for docx.list.apply: $.kind must be one of "bullet", "ordered".',
      },
    })
    expect(command).toBeUndefined()

    const context = await client.callTool({
      name: 'office_get_context',
      arguments: { sessionId },
    })
    expect(context.structuredContent).toMatchObject({ session: { revision: 0 } })
  })

  it('rejects an invalid DOCX list-removal block index before enqueueing', async () => {
    const client = await connectClient()
    const created = await client.callTool({
      name: 'office_create_session',
      arguments: { format: 'docx' },
    })
    const sessionId = (created.structuredContent as { session: { id: string } }).session.id
    await client.callTool({ name: 'office_editor_poll', arguments: { sessionId } })

    const execution = executeOperation(client, sessionId, 0, 'docx.list.remove', {
      target: { blockIndexes: [0, -1] },
    })

    await new Promise((resolve) => setTimeout(resolve, 20))
    const polled = await client.callTool({
      name: 'office_editor_poll',
      arguments: { sessionId },
    })
    const command = (
      polled.structuredContent as {
        commands: Array<{ commandId: string }>
      }
    ).commands[0]

    await expect(execution).resolves.toMatchObject({
      isError: true,
      structuredContent: {
        ok: false,
        error: 'operation_schema_invalid',
        message:
          'Invalid arguments for docx.list.remove: $.target.blockIndexes[1] must be greater than or equal to 0.',
      },
    })
    expect(command).toBeUndefined()

    const context = await client.callTool({
      name: 'office_get_context',
      arguments: { sessionId },
    })
    expect(context.structuredContent).toMatchObject({ session: { revision: 0 } })
  })

  it.each([
    [
      'block-index item',
      { target: { blockIndexes: [0, -1] }, level: 2 },
      '$.target.blockIndexes[1] must be greater than or equal to 0.',
    ],
    [
      'heading-level range',
      { target: { containsText: 'Risk Notes' }, level: 7 },
      '$.level must be less than or equal to 6.',
    ],
  ])('rejects an invalid DOCX %s before enqueueing', async (_case, arguments_, detail) => {
    const client = await connectClient()
    const created = await client.callTool({
      name: 'office_create_session',
      arguments: { format: 'docx' },
    })
    const sessionId = (created.structuredContent as { session: { id: string } }).session.id
    await client.callTool({ name: 'office_editor_poll', arguments: { sessionId } })

    const execution = executeOperation(
      client,
      sessionId,
      0,
      'docx.paragraph.set_heading_level',
      arguments_,
    )

    await new Promise((resolve) => setTimeout(resolve, 20))
    const polled = await client.callTool({
      name: 'office_editor_poll',
      arguments: { sessionId },
    })
    const command = (
      polled.structuredContent as {
        commands: Array<{ commandId: string }>
      }
    ).commands[0]

    await expect(execution).resolves.toMatchObject({
      isError: true,
      structuredContent: {
        ok: false,
        error: 'operation_schema_invalid',
        message: `Invalid arguments for docx.paragraph.set_heading_level: ${detail}`,
      },
    })
    expect(command).toBeUndefined()

    const context = await client.callTool({
      name: 'office_get_context',
      arguments: { sessionId },
    })
    expect(context.structuredContent).toMatchObject({ session: { revision: 0 } })
  })

  it('returns only after the live editor acknowledges the applied revision', async () => {
    const client = await connectClient()
    const created = await client.callTool({
      name: 'office_create_session',
      arguments: { format: 'docx' },
    })
    const sessionId = (created.structuredContent as { session: { id: string } }).session.id
    await client.callTool({
      name: 'office_editor_poll',
      arguments: { sessionId, selection: { from: 1, to: 1 } },
    })

    const execution = executeOperation(client, sessionId, 0, 'docx.text.insert', {
      text: 'hello',
    })

    let command: { commandId: string } | undefined
    for (let attempt = 0; attempt < 20 && !command; attempt += 1) {
      const polled = await client.callTool({
        name: 'office_editor_poll',
        arguments: { sessionId },
      })
      command = (polled.structuredContent as { commands: Array<{ commandId: string }> }).commands[0]
      if (!command) await new Promise((resolve) => setTimeout(resolve, 10))
    }
    expect(command).toBeDefined()

    const acknowledgement = await client.callTool({
      name: 'office_editor_acknowledge',
      arguments: {
        sessionId,
        commandId: command!.commandId,
        revision: 1,
        dirty: true,
        output: { summary: 'Inserted text', changed: 1 },
        timing: {
          hydrateMs: 0.25,
          executeMs: 2.5,
          trace: {
            operation: 'markdown.document.load_staged',
            phases: {
              decodeMs: 0.1,
              parseMs: 0.8,
              tiptapStateInstallMs: 0.4,
              reactCommitMs: 0.2,
            },
          },
        },
      },
    })

    expect(acknowledgement.structuredContent).toMatchObject({
      ok: true,
      timing: {
        hydrateMs: 0.25,
        executeMs: 2.5,
        trace: {
          operation: 'markdown.document.load_staged',
          phases: {
            decodeMs: 0.1,
            parseMs: 0.8,
            tiptapStateInstallMs: 0.4,
            reactCommitMs: 0.2,
          },
        },
      },
    })

    await expect(execution).resolves.toMatchObject({
      structuredContent: {
        ok: true,
        result: {
          revision: 1,
          operations: [
            {
              id: 'docx.text.insert',
              result: { summary: 'Inserted text', changed: 1 },
            },
          ],
        },
      },
    })
  })

  it('wakes a waiting editor poll as soon as a live command is enqueued', async () => {
    const client = await connectClient()
    const created = await client.callTool({
      name: 'office_create_session',
      arguments: { format: 'docx' },
    })
    const sessionId = (created.structuredContent as { session: { id: string } }).session.id
    await client.callTool({
      name: 'office_editor_poll',
      arguments: { sessionId, selection: { from: 1, to: 1 } },
    })

    const waitingPoll = client.callTool({
      name: 'office_editor_poll',
      arguments: { sessionId, waitMs: 1_000 },
    })
    await new Promise((resolve) => setTimeout(resolve, 20))

    const execution = executeOperation(client, sessionId, 0, 'docx.text.insert', {
      text: 'wake now',
    })

    const polled = await waitingPoll
    const command = (
      polled.structuredContent as {
        commands: Array<{ commandId: string; operation: string; arguments: unknown }>
      }
    ).commands[0]
    expect(command).toMatchObject({
      operation: 'docx.text.insert',
      arguments: { text: 'wake now' },
    })

    await client.callTool({
      name: 'office_editor_acknowledge',
      arguments: {
        sessionId,
        commandId: command.commandId,
        revision: 1,
        dirty: true,
      },
    })
    await expect(execution).resolves.toMatchObject({
      structuredContent: { ok: true, result: { revision: 1 } },
    })
  })

  it('returns a renderer failure for a known operation instead of reporting success', async () => {
    const client = await connectClient()
    const created = await client.callTool({
      name: 'office_create_session',
      arguments: { format: 'docx' },
    })
    const sessionId = (created.structuredContent as { session: { id: string } }).session.id
    await client.callTool({
      name: 'office_editor_poll',
      arguments: { sessionId, selection: { from: 1, to: 1 } },
    })

    const execution = executeOperation(client, sessionId, 0, 'docx.text.insert', {
      text: 'rejected by renderer',
    })

    let command: { commandId: string } | undefined
    for (let attempt = 0; attempt < 20 && !command; attempt += 1) {
      const polled = await client.callTool({
        name: 'office_editor_poll',
        arguments: { sessionId },
      })
      command = (polled.structuredContent as { commands: Array<{ commandId: string }> }).commands[0]
      if (!command) await new Promise((resolve) => setTimeout(resolve, 10))
    }
    expect(command).toBeDefined()

    await client.callTool({
      name: 'office_editor_acknowledge',
      arguments: {
        sessionId,
        commandId: command!.commandId,
        revision: 1,
        ok: false,
        error: 'execution_failed',
        message: 'The renderer rejected the insertion.',
      },
    })

    await expect(execution).resolves.toMatchObject({
      isError: true,
      structuredContent: {
        ok: false,
        error: 'execution_failed',
        message: 'The renderer rejected the insertion.',
      },
    })
  })

  it('queues a DOCX local file as docx.document.load_staged and waits for the editor', async () => {
    const directory = await mkdtemp(join(tmpdir(), 'tandemfolio-open-'))
    temporaryDirectories.push(directory)
    const path = join(directory, 'local.docx')
    const bytes = Buffer.from('PK\u0003\u0004docx-test-bytes')
    await writeFile(path, bytes)

    const client = await connectClient()
    const created = await client.callTool({
      name: 'office_create_session',
      arguments: { format: 'docx' },
    })
    const sessionId = (created.structuredContent as { session: { id: string } }).session.id
    await client.callTool({ name: 'office_editor_poll', arguments: { sessionId } })

    const opening = client.callTool({
      name: 'office_open_local_file',
      arguments: { sessionId, baseRevision: 0, path },
    })

    let command:
      | {
          commandId: string
          operation: string
          arguments: { blobId: string; name: string; size: number }
        }
      | undefined
    for (let attempt = 0; attempt < 20 && !command; attempt += 1) {
      const polled = await client.callTool({
        name: 'office_editor_poll',
        arguments: { sessionId },
      })
      command = (
        polled.structuredContent as {
          commands: Array<{
            commandId: string
            operation: string
            arguments: { blobId: string; name: string; size: number }
          }>
        }
      ).commands[0]
      if (!command) await new Promise((resolve) => setTimeout(resolve, 10))
    }

    expect(command).toMatchObject({
      operation: 'docx.document.load_staged',
      arguments: { name: 'local.docx', size: bytes.length },
    })
    expect(command!.arguments).not.toHaveProperty('data')

    const chunk = await client.callTool({
      name: 'office_editor_read_file_chunk',
      arguments: {
        sessionId,
        blobId: command!.arguments.blobId,
        offset: 0,
        length: 8,
      },
    })
    expect(chunk.structuredContent).toMatchObject({
      ok: true,
      offset: 0,
      data: bytes.subarray(0, 8).toString('base64'),
      nextOffset: 8,
      eof: false,
    })

    await client.callTool({
      name: 'office_editor_acknowledge',
      arguments: {
        sessionId,
        commandId: command!.commandId,
        revision: 1,
        fileName: 'local.docx',
        dirty: false,
      },
    })

    await expect(opening).resolves.toMatchObject({
      structuredContent: {
        ok: true,
        result: { commandId: command!.commandId, ok: true, revision: 1 },
      },
    })
  })

  it('queues a Markdown local file as markdown.document.load_staged', async () => {
    const directory = await mkdtemp(join(tmpdir(), 'tandemfolio-markdown-open-'))
    temporaryDirectories.push(directory)
    const path = join(directory, 'local.md')
    const bytes = Buffer.from('# Local Markdown')
    await writeFile(path, bytes)
    await mkdir(join(directory, 'assets'))
    const png = Buffer.from(
      'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M/wHwAF/gL+Xw4bWQAAAABJRU5ErkJggg==',
      'base64',
    )
    await writeFile(join(directory, 'assets', 'pixel.png'), png)

    const client = await connectClient()
    const created = await client.callTool({
      name: 'office_create_session',
      arguments: { format: 'markdown' },
    })
    const sessionId = (created.structuredContent as { session: { id: string } }).session.id
    await client.callTool({ name: 'office_editor_poll', arguments: { sessionId } })

    const opening = client.callTool({
      name: 'office_open_local_file',
      arguments: { sessionId, baseRevision: 0, path },
    })

    let command:
      | {
          commandId: string
          operation: string
          arguments: { blobId: string; name: string; size: number; assetRootId: string }
        }
      | undefined
    for (let attempt = 0; attempt < 20 && !command; attempt += 1) {
      const polled = await client.callTool({
        name: 'office_editor_poll',
        arguments: { sessionId },
      })
      command = (
        polled.structuredContent as {
          commands: Array<{
            commandId: string
            operation: string
            arguments: { blobId: string; name: string; size: number; assetRootId: string }
          }>
        }
      ).commands[0]
      if (!command) await new Promise((resolve) => setTimeout(resolve, 10))
    }

    expect(command).toMatchObject({
      operation: 'markdown.document.load_staged',
      arguments: { name: 'local.md', size: bytes.length, assetRootId: expect.any(String) },
    })
    expect(command!.arguments).not.toHaveProperty('data')

    const asset = await client.callTool({
      name: 'office_editor_read_local_asset_chunk',
      arguments: {
        sessionId,
        rootId: command!.arguments.assetRootId,
        path: 'assets/pixel.png',
        offset: 0,
        length: 8,
      },
    })
    expect(asset.structuredContent).toMatchObject({
      ok: true,
      size: png.length,
      mime: 'image/png',
      data: png.subarray(0, 8).toString('base64'),
      nextOffset: 8,
      eof: false,
    })

    await client.callTool({
      name: 'office_editor_acknowledge',
      arguments: {
        sessionId,
        commandId: command!.commandId,
        revision: 1,
        fileName: 'local.md',
        dirty: false,
      },
    })
    await expect(opening).resolves.toMatchObject({
      structuredContent: {
        ok: true,
        result: { commandId: command!.commandId, ok: true, revision: 1 },
      },
    })
  })

  it('stages markdown.image.insert paths as internal bytes without inline payloads', async () => {
    const directory = await mkdtemp(join(tmpdir(), 'tandemfolio-markdown-image-'))
    temporaryDirectories.push(directory)
    const path = join(directory, 'logo.png')
    const bytes = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a])
    await writeFile(path, bytes)

    const client = await connectClient()
    const created = await client.callTool({
      name: 'office_create_session',
      arguments: { format: 'markdown' },
    })
    const sessionId = (created.structuredContent as { session: { id: string } }).session.id
    await client.callTool({
      name: 'office_editor_poll',
      arguments: { sessionId, selection: { from: 1, to: 1 } },
    })

    const inserting = executeOperation(client, sessionId, 0, 'markdown.image.insert', {
      path,
      position: 2,
      alt: 'Logo',
      title: null,
    })
    let command:
      | {
          commandId: string
          operation: string
          arguments: {
            blobId: string
            name: string
            size: number
            position: number
            alt: string
            title: string | null
          }
        }
      | undefined
    for (let attempt = 0; attempt < 20 && !command; attempt += 1) {
      const polled = await client.callTool({
        name: 'office_editor_poll',
        arguments: { sessionId },
      })
      command = (
        polled.structuredContent as {
          commands: Array<typeof command>
        }
      ).commands[0]
      if (!command) await new Promise((resolve) => setTimeout(resolve, 10))
    }
    expect(command).toMatchObject({
      operation: 'markdown.image.insert_staged',
      arguments: {
        name: 'logo.png',
        size: bytes.length,
        position: 2,
        alt: 'Logo',
        title: null,
      },
    })
    expect(command!.arguments).not.toHaveProperty('path')
    expect(command!.arguments).not.toHaveProperty('data')

    await client.callTool({
      name: 'office_editor_acknowledge',
      arguments: {
        sessionId,
        commandId: command!.commandId,
        revision: 1,
        dirty: true,
        output: {},
      },
    })
    await expect(inserting).resolves.toMatchObject({
      structuredContent: {
        ok: true,
        result: {
          revision: 1,
          operations: [{ id: 'markdown.image.insert', result: {} }],
        },
      },
    })
  })

  it('queues an XLSX local file as xlsx.document.load_staged', async () => {
    const directory = await mkdtemp(join(tmpdir(), 'tandemfolio-xlsx-open-'))
    temporaryDirectories.push(directory)
    const path = join(directory, 'local.xlsx')
    const bytes = Buffer.from('PK\u0003\u0004xlsx-test-bytes')
    await writeFile(path, bytes)

    const client = await connectClient()
    const created = await client.callTool({
      name: 'office_create_session',
      arguments: { format: 'xlsx' },
    })
    const sessionId = (created.structuredContent as { session: { id: string } }).session.id
    await client.callTool({ name: 'office_editor_poll', arguments: { sessionId } })

    const opening = client.callTool({
      name: 'office_open_local_file',
      arguments: { sessionId, baseRevision: 0, path },
    })

    let command:
      | {
          commandId: string
          operation: string
          arguments: { blobId: string; name: string; size: number }
        }
      | undefined
    for (let attempt = 0; attempt < 20 && !command; attempt += 1) {
      const polled = await client.callTool({
        name: 'office_editor_poll',
        arguments: { sessionId },
      })
      command = (
        polled.structuredContent as {
          commands: Array<{
            commandId: string
            operation: string
            arguments: { blobId: string; name: string; size: number }
          }>
        }
      ).commands[0]
      if (!command) await new Promise((resolve) => setTimeout(resolve, 10))
    }

    expect(command).toMatchObject({
      operation: 'xlsx.document.load_staged',
      arguments: { name: 'local.xlsx', size: bytes.length },
    })
    expect(command!.arguments).not.toHaveProperty('data')

    await client.callTool({
      name: 'office_editor_acknowledge',
      arguments: {
        sessionId,
        commandId: command!.commandId,
        revision: 1,
        fileName: 'local.xlsx',
        dirty: false,
        output: { opened: true, fileName: 'local.xlsx' },
      },
    })
    await expect(opening).resolves.toMatchObject({
      structuredContent: {
        ok: true,
        result: {
          commandId: command!.commandId,
          ok: true,
          revision: 1,
          output: { opened: true, fileName: 'local.xlsx' },
        },
      },
    })
  })

  it('stages xlsx.image.add paths as an internal image operation without inline bytes', async () => {
    const directory = await mkdtemp(join(tmpdir(), 'tandemfolio-xlsx-image-'))
    temporaryDirectories.push(directory)
    const path = join(directory, 'logo.png')
    const bytes = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a])
    await writeFile(path, bytes)

    const client = await connectClient()
    const created = await client.callTool({
      name: 'office_create_session',
      arguments: { format: 'xlsx' },
    })
    const sessionId = (created.structuredContent as { session: { id: string } }).session.id
    await client.callTool({ name: 'office_editor_poll', arguments: { sessionId } })

    const inserting = executeOperation(client, sessionId, 0, 'xlsx.image.add', {
      sheet: 'Sheet1',
      path,
      anchorCell: 'C3',
    })
    let command:
      | {
          commandId: string
          operation: string
          arguments: {
            blobId: string
            name: string
            size: number
            sheet: string
            anchorCell: string
          }
        }
      | undefined
    for (let attempt = 0; attempt < 20 && !command; attempt += 1) {
      const polled = await client.callTool({
        name: 'office_editor_poll',
        arguments: { sessionId },
      })
      command = (
        polled.structuredContent as {
          commands: Array<typeof command>
        }
      ).commands[0]
      if (!command) await new Promise((resolve) => setTimeout(resolve, 10))
    }
    expect(command).toMatchObject({
      operation: 'xlsx.image.add_staged',
      arguments: {
        name: 'logo.png',
        size: bytes.length,
        sheet: 'Sheet1',
        anchorCell: 'C3',
      },
    })
    expect(command!.arguments).not.toHaveProperty('data')

    await client.callTool({
      name: 'office_editor_acknowledge',
      arguments: {
        sessionId,
        commandId: command!.commandId,
        revision: 1,
        dirty: true,
        output: { sheet: 'Sheet1', anchorCell: 'C3', visualId: 'added-image-1' },
      },
    })
    await expect(inserting).resolves.toMatchObject({
      structuredContent: {
        ok: true,
        result: {
          revision: 1,
          operations: [
            {
              id: 'xlsx.image.add',
              result: { sheet: 'Sheet1', anchorCell: 'C3', visualId: 'added-image-1' },
            },
          ],
        },
      },
    })
  })

  it('stages docx.document.compare paths as internal DOCX bytes without exposing them inline', async () => {
    const directory = await mkdtemp(join(tmpdir(), 'tandemfolio-docx-compare-'))
    temporaryDirectories.push(directory)
    const path = join(directory, 'other.docx')
    const bytes = Buffer.from('PK\u0003\u0004comparison-fixture')
    await writeFile(path, bytes)

    const client = await connectClient()
    const created = await client.callTool({
      name: 'office_create_session',
      arguments: { format: 'docx' },
    })
    const sessionId = (created.structuredContent as { session: { id: string } }).session.id
    await client.callTool({ name: 'office_editor_poll', arguments: { sessionId } })

    const comparing = executeOperation(client, sessionId, 0, 'docx.document.compare', { path })
    let command:
      { commandId: string; operation: string; arguments: Record<string, unknown> } | undefined
    for (let attempt = 0; attempt < 20 && !command; attempt += 1) {
      const polled = await client.callTool({
        name: 'office_editor_poll',
        arguments: { sessionId },
      })
      command = (polled.structuredContent as { commands: Array<typeof command> }).commands[0]
      if (!command) await new Promise((resolve) => setTimeout(resolve, 10))
    }
    expect(command).toMatchObject({
      operation: 'docx.document.compare_staged',
      arguments: { name: 'other.docx', size: bytes.length },
    })
    expect(command!.arguments).not.toHaveProperty('path')
    expect(command!.arguments).not.toHaveProperty('data')

    await client.callTool({
      name: 'office_editor_acknowledge',
      arguments: {
        sessionId,
        commandId: command!.commandId,
        revision: 1,
        dirty: false,
        output: {
          summary: 'Compared DOCX with other.docx: 0 added, 0 removed, 1 changed',
          otherName: 'other.docx',
          added: 0,
          removed: 0,
          changed: 1,
          identical: false,
        },
      },
    })
    await expect(comparing).resolves.toMatchObject({
      structuredContent: {
        ok: true,
        result: {
          revision: 1,
          operations: [
            {
              id: 'docx.document.compare',
              result: { changed: 1, identical: false },
            },
          ],
        },
      },
    })
  })

  it('stages docx.image.insert paths as an internal image operation without inline bytes', async () => {
    const directory = await mkdtemp(join(tmpdir(), 'tandemfolio-docx-image-'))
    temporaryDirectories.push(directory)
    const path = join(directory, 'logo.gif')
    const bytes = Buffer.from('GIF87a')
    await writeFile(path, bytes)

    const client = await connectClient()
    const created = await client.callTool({
      name: 'office_create_session',
      arguments: { format: 'docx' },
    })
    const sessionId = (created.structuredContent as { session: { id: string } }).session.id
    await client.callTool({ name: 'office_editor_poll', arguments: { sessionId } })

    const inserting = executeOperation(client, sessionId, 0, 'docx.image.insert', {
      path,
      afterBlockIndex: 2,
      widthPx: 320,
      heightPx: 180,
      alignment: 'center',
    })
    let command:
      | {
          commandId: string
          operation: string
          arguments: Record<string, unknown>
        }
      | undefined
    for (let attempt = 0; attempt < 20 && !command; attempt += 1) {
      const polled = await client.callTool({
        name: 'office_editor_poll',
        arguments: { sessionId },
      })
      command = (polled.structuredContent as { commands: Array<typeof command> }).commands[0]
      if (!command) await new Promise((resolve) => setTimeout(resolve, 10))
    }
    expect(command).toMatchObject({
      operation: 'docx.image.insert_staged',
      arguments: {
        name: 'logo.gif',
        size: bytes.length,
        afterBlockIndex: 2,
        widthPx: 320,
        heightPx: 180,
        alignment: 'center',
      },
    })
    expect(command!.arguments).not.toHaveProperty('path')
    expect(command!.arguments).not.toHaveProperty('data')

    await client.callTool({
      name: 'office_editor_acknowledge',
      arguments: {
        sessionId,
        commandId: command!.commandId,
        revision: 1,
        dirty: true,
        output: { insertedBlockIndex: 3 },
      },
    })
    await expect(inserting).resolves.toMatchObject({
      structuredContent: {
        ok: true,
        result: {
          revision: 1,
          operations: [{ id: 'docx.image.insert', result: { insertedBlockIndex: 3 } }],
        },
      },
    })
  })

  it('stages docx.image.replace paths as an internal replacement without inline bytes', async () => {
    const directory = await mkdtemp(join(tmpdir(), 'tandemfolio-docx-image-replace-'))
    temporaryDirectories.push(directory)
    const path = join(directory, 'replacement.gif')
    const bytes = Buffer.from('GIF87a')
    await writeFile(path, bytes)

    const client = await connectClient()
    const created = await client.callTool({
      name: 'office_create_session',
      arguments: { format: 'docx' },
    })
    const sessionId = (created.structuredContent as { session: { id: string } }).session.id
    await client.callTool({ name: 'office_editor_poll', arguments: { sessionId } })

    const replacing = executeOperation(client, sessionId, 0, 'docx.image.replace', {
      path,
      imageBlockIndex: 3,
      widthPx: 400,
      heightPx: 225,
    })
    let command:
      | {
          commandId: string
          operation: string
          arguments: Record<string, unknown>
        }
      | undefined
    for (let attempt = 0; attempt < 20 && !command; attempt += 1) {
      const polled = await client.callTool({
        name: 'office_editor_poll',
        arguments: { sessionId },
      })
      command = (polled.structuredContent as { commands: Array<typeof command> }).commands[0]
      if (!command) await new Promise((resolve) => setTimeout(resolve, 10))
    }
    expect(command).toMatchObject({
      operation: 'docx.image.replace_staged',
      arguments: {
        name: 'replacement.gif',
        size: bytes.length,
        imageBlockIndex: 3,
        widthPx: 400,
        heightPx: 225,
      },
    })
    expect(command!.arguments).not.toHaveProperty('path')
    expect(command!.arguments).not.toHaveProperty('data')

    await client.callTool({
      name: 'office_editor_acknowledge',
      arguments: {
        sessionId,
        commandId: command!.commandId,
        revision: 1,
        dirty: true,
        output: { imageBlockIndex: 3, widthPx: 400, heightPx: 225 },
      },
    })
    await expect(replacing).resolves.toMatchObject({
      structuredContent: {
        ok: true,
        result: {
          revision: 1,
          operations: [{ id: 'docx.image.replace', result: { imageBlockIndex: 3 } }],
        },
      },
    })
  })

  it('queues a PPTX local file as pptx.document.load_staged', async () => {
    const directory = await mkdtemp(join(tmpdir(), 'tandemfolio-pptx-open-'))
    temporaryDirectories.push(directory)
    const path = join(directory, 'local.pptx')
    const bytes = Buffer.from('PK\u0003\u0004pptx-test-bytes')
    await writeFile(path, bytes)

    const client = await connectClient()
    const created = await client.callTool({
      name: 'office_create_session',
      arguments: { format: 'pptx' },
    })
    const sessionId = (created.structuredContent as { session: { id: string } }).session.id
    await client.callTool({ name: 'office_editor_poll', arguments: { sessionId } })

    const opening = client.callTool({
      name: 'office_open_local_file',
      arguments: { sessionId, baseRevision: 0, path },
    })

    let command:
      | {
          commandId: string
          operation: string
          arguments: { blobId: string; name: string; size: number }
        }
      | undefined
    for (let attempt = 0; attempt < 20 && !command; attempt += 1) {
      const polled = await client.callTool({
        name: 'office_editor_poll',
        arguments: { sessionId },
      })
      command = (
        polled.structuredContent as {
          commands: Array<{
            commandId: string
            operation: string
            arguments: { blobId: string; name: string; size: number }
          }>
        }
      ).commands[0]
      if (!command) await new Promise((resolve) => setTimeout(resolve, 10))
    }

    expect(command).toMatchObject({
      operation: 'pptx.document.load_staged',
      arguments: { name: 'local.pptx', size: bytes.length },
    })
    expect(command!.arguments).not.toHaveProperty('data')

    await client.callTool({
      name: 'office_editor_acknowledge',
      arguments: {
        sessionId,
        commandId: command!.commandId,
        revision: 1,
        fileName: 'local.pptx',
        dirty: false,
        output: { opened: true, fileName: 'local.pptx' },
      },
    })
    await expect(opening).resolves.toMatchObject({
      structuredContent: {
        ok: true,
        result: {
          commandId: command!.commandId,
          ok: true,
          revision: 1,
          output: { opened: true, fileName: 'local.pptx' },
        },
      },
    })
  })

  it('queues a PDF local file as pdf.document.load_staged', async () => {
    const directory = await mkdtemp(join(tmpdir(), 'tandemfolio-pdf-open-'))
    temporaryDirectories.push(directory)
    const path = join(directory, 'local.pdf')
    const bytes = Buffer.from('%PDF-1.7\npdf-test-bytes')
    await writeFile(path, bytes)

    const client = await connectClient()
    const created = await client.callTool({
      name: 'office_create_session',
      arguments: { format: 'pdf' },
    })
    const sessionId = (created.structuredContent as { session: { id: string } }).session.id
    await client.callTool({ name: 'office_editor_poll', arguments: { sessionId } })

    const opening = client.callTool({
      name: 'office_open_local_file',
      arguments: { sessionId, baseRevision: 0, path },
    })

    let command:
      | {
          commandId: string
          operation: string
          arguments: { blobId: string; name: string; size: number }
        }
      | undefined
    for (let attempt = 0; attempt < 20 && !command; attempt += 1) {
      const polled = await client.callTool({
        name: 'office_editor_poll',
        arguments: { sessionId },
      })
      command = (
        polled.structuredContent as {
          commands: Array<{
            commandId: string
            operation: string
            arguments: { blobId: string; name: string; size: number }
          }>
        }
      ).commands[0]
      if (!command) await new Promise((resolve) => setTimeout(resolve, 10))
    }

    expect(command).toMatchObject({
      operation: 'pdf.document.load_staged',
      arguments: { name: 'local.pdf', size: bytes.length },
    })
    expect(command!.arguments).not.toHaveProperty('data')

    await client.callTool({
      name: 'office_editor_acknowledge',
      arguments: {
        sessionId,
        commandId: command!.commandId,
        revision: 1,
        fileName: 'local.pdf',
        dirty: false,
        output: { opened: 'local.pdf' },
      },
    })
    await expect(opening).resolves.toMatchObject({
      structuredContent: {
        ok: true,
        result: {
          commandId: command!.commandId,
          ok: true,
          revision: 1,
          output: { opened: 'local.pdf' },
        },
      },
    })
  })

  it('stages pdf.page.insert paths as internal PDF bytes without inline payloads', async () => {
    const directory = await mkdtemp(join(tmpdir(), 'tandemfolio-pdf-insert-'))
    temporaryDirectories.push(directory)
    const path = join(directory, 'insert.pdf')
    const bytes = Buffer.from('%PDF-1.7\ninsert-test-bytes')
    await writeFile(path, bytes)

    const client = await connectClient()
    const created = await client.callTool({
      name: 'office_create_session',
      arguments: { format: 'pdf' },
    })
    const sessionId = (created.structuredContent as { session: { id: string } }).session.id
    await client.callTool({ name: 'office_editor_poll', arguments: { sessionId } })

    const insertion = executeOperation(client, sessionId, 0, 'pdf.page.insert', {
      path,
      afterPageIndex: 0,
    })

    let command:
      | {
          commandId: string
          operation: string
          arguments: { blobId: string; name: string; size: number; afterPageIndex: number }
        }
      | undefined
    for (let attempt = 0; attempt < 20 && !command; attempt += 1) {
      const polled = await client.callTool({ name: 'office_editor_poll', arguments: { sessionId } })
      command = (
        polled.structuredContent as {
          commands: Array<{
            commandId: string
            operation: string
            arguments: { blobId: string; name: string; size: number; afterPageIndex: number }
          }>
        }
      ).commands[0]
      if (!command) await new Promise((resolve) => setTimeout(resolve, 10))
    }

    expect(command).toMatchObject({
      operation: 'pdf.page.insert_staged',
      arguments: { name: 'insert.pdf', size: bytes.length, afterPageIndex: 0 },
    })
    expect(command!.arguments).not.toHaveProperty('path')
    expect(command!.arguments).not.toHaveProperty('data')

    await client.callTool({
      name: 'office_editor_acknowledge',
      arguments: {
        sessionId,
        commandId: command!.commandId,
        revision: 1,
        dirty: false,
        output: { insertedCount: 2 },
      },
    })
    await expect(insertion).resolves.toMatchObject({
      structuredContent: {
        ok: true,
        result: {
          revision: 1,
          operations: [{ id: 'pdf.page.insert', result: { insertedCount: 2 } }],
        },
      },
    })
  })

  it('restores the latest renderer checkpoint after the MCP process restarts', async () => {
    const stateDirectory = await mkdtemp(join(tmpdir(), 'tandemfolio-state-'))
    temporaryDirectories.push(stateDirectory)
    const bytes = Buffer.from('PK\u0003\u0004recovered-docx')

    const firstClient = await connectClient(stateDirectory)
    const firstCreated = await firstClient.callTool({
      name: 'office_create_session',
      arguments: { format: 'docx', resume: 'none' },
    })
    const firstSessionId = (firstCreated.structuredContent as { session: { id: string } }).session
      .id
    const begun = await firstClient.callTool({
      name: 'office_editor_begin_recovery',
      arguments: { sessionId: firstSessionId, fileName: 'draft.docx', size: bytes.length },
    })
    const uploadId = (begun.structuredContent as { uploadId: string }).uploadId
    await firstClient.callTool({
      name: 'office_editor_write_recovery_chunk',
      arguments: { sessionId: firstSessionId, uploadId, offset: 0, data: bytes.toString('base64') },
    })
    await firstClient.callTool({
      name: 'office_editor_commit_recovery',
      arguments: { sessionId: firstSessionId, uploadId },
    })
    await firstClient.close()
    clients.splice(clients.indexOf(firstClient), 1)

    const secondClient = await connectClient(stateDirectory)
    const secondCreated = await secondClient.callTool({
      name: 'office_create_session',
      arguments: { format: 'docx', resume: 'latest' },
    })
    const secondSessionId = (secondCreated.structuredContent as { session: { id: string } }).session
      .id
    const polled = await secondClient.callTool({
      name: 'office_editor_poll',
      arguments: { sessionId: secondSessionId },
    })
    const [command] = (
      polled.structuredContent as {
        commands: Array<{
          operation: string
          arguments: { blobId: string; name: string; size: number }
        }>
      }
    ).commands

    expect(command).toMatchObject({
      operation: 'docx.document.load_staged',
      arguments: { name: 'draft.docx', size: bytes.length },
    })
    const recovered = await secondClient.callTool({
      name: 'office_editor_read_file_chunk',
      arguments: {
        sessionId: secondSessionId,
        blobId: command.arguments.blobId,
        offset: 0,
        length: bytes.length,
      },
    })
    expect(recovered.structuredContent).toMatchObject({
      ok: true,
      data: bytes.toString('base64'),
      eof: true,
    })
  })

  it('creates an isolated blank session by default even when another task has a recovery', async () => {
    const stateDirectory = await mkdtemp(join(tmpdir(), 'tandemfolio-state-'))
    temporaryDirectories.push(stateDirectory)
    const bytes = Buffer.from('PK\u0003\u0004other-task-pptx')

    const firstClient = await connectClient(stateDirectory)
    const firstCreated = await firstClient.callTool({
      name: 'office_create_session',
      arguments: { format: 'pptx', resume: 'none' },
    })
    const firstSessionId = (firstCreated.structuredContent as { session: { id: string } }).session
      .id
    const begun = await firstClient.callTool({
      name: 'office_editor_begin_recovery',
      arguments: { sessionId: firstSessionId, fileName: 'other-task.pptx', size: bytes.length },
    })
    const uploadId = (begun.structuredContent as { uploadId: string }).uploadId
    await firstClient.callTool({
      name: 'office_editor_write_recovery_chunk',
      arguments: { sessionId: firstSessionId, uploadId, offset: 0, data: bytes.toString('base64') },
    })
    await firstClient.callTool({
      name: 'office_editor_commit_recovery',
      arguments: { sessionId: firstSessionId, uploadId },
    })
    await firstClient.close()
    clients.splice(clients.indexOf(firstClient), 1)

    const secondClient = await connectClient(stateDirectory)
    const secondCreated = await secondClient.callTool({
      name: 'office_create_session',
      arguments: { format: 'pptx' },
    })
    const secondSessionId = (secondCreated.structuredContent as { session: { id: string } }).session
      .id
    const polled = await secondClient.callTool({
      name: 'office_editor_poll',
      arguments: { sessionId: secondSessionId },
    })

    expect(secondCreated.structuredContent).toMatchObject({ recoveryAvailable: false })
    expect(polled.structuredContent).toMatchObject({ ok: true, commands: [] })
  })

  it.each([
    ['docx', 'task.docx', 'docx.document.load_staged'],
    ['markdown', 'task.md', 'markdown.document.load_staged'],
    ['xlsx', 'task.xlsx', 'xlsx.document.load_staged'],
    ['pptx', 'task.pptx', 'pptx.document.load_staged'],
    ['pdf', 'task.pdf', 'pdf.document.load_staged'],
  ] as const)(
    'reopens the exact prior %s session instead of selecting another recovery',
    async (format, fileName, operation) => {
      const stateDirectory = await mkdtemp(join(tmpdir(), 'tandemfolio-state-'))
      temporaryDirectories.push(stateDirectory)
      const bytes = Buffer.from(`exact-${format}-session`)

      const firstClient = await connectClient(stateDirectory)
      const firstCreated = await firstClient.callTool({
        name: 'office_create_session',
        arguments: { format, resume: 'none' },
      })
      const sessionId = (firstCreated.structuredContent as { session: { id: string } }).session.id
      const begun = await firstClient.callTool({
        name: 'office_editor_begin_recovery',
        arguments: { sessionId, fileName, size: bytes.length },
      })
      const uploadId = (begun.structuredContent as { uploadId: string }).uploadId
      await firstClient.callTool({
        name: 'office_editor_write_recovery_chunk',
        arguments: { sessionId, uploadId, offset: 0, data: bytes.toString('base64') },
      })
      await firstClient.callTool({
        name: 'office_editor_commit_recovery',
        arguments: { sessionId, uploadId },
      })
      await firstClient.close()
      clients.splice(clients.indexOf(firstClient), 1)

      const secondClient = await connectClient(stateDirectory)
      const reopened = await secondClient.callTool({
        name: 'office_create_session',
        arguments: { format, resume: 'exact', sessionId },
      })
      expect(reopened.structuredContent).toMatchObject({
        ok: true,
        recoveryAvailable: true,
        reused: false,
        session: { id: sessionId, format },
      })

      const polled = await secondClient.callTool({
        name: 'office_editor_poll',
        arguments: { sessionId },
      })
      expect(polled.structuredContent).toMatchObject({
        ok: true,
        commands: [{ operation, arguments: { name: fileName, size: bytes.length } }],
      })
    },
  )

  it('restores the disconnected session checkpoint when that editing session reconnects', async () => {
    const client = await connectClient()
    const created = await client.callTool({
      name: 'office_create_session',
      arguments: { format: 'markdown', resume: 'none' },
    })
    const sessionId = (created.structuredContent as { session: { id: string } }).session.id
    await client.callTool({ name: 'office_editor_poll', arguments: { sessionId } })
    const bytes = Buffer.from('# Session-specific draft')
    const begun = await client.callTool({
      name: 'office_editor_begin_recovery',
      arguments: { sessionId, fileName: 'task-draft.md', size: bytes.length },
    })
    const uploadId = (begun.structuredContent as { uploadId: string }).uploadId
    await client.callTool({
      name: 'office_editor_write_recovery_chunk',
      arguments: { sessionId, uploadId, offset: 0, data: bytes.toString('base64') },
    })
    await client.callTool({
      name: 'office_editor_commit_recovery',
      arguments: { sessionId, uploadId },
    })
    await client.callTool({ name: 'office_editor_disconnect', arguments: { sessionId } })

    const reconnected = await client.callTool({
      name: 'office_editor_poll',
      arguments: { sessionId },
    })
    const [command] = (
      reconnected.structuredContent as {
        commands: Array<{ operation: string; arguments: { name: string; size: number } }>
      }
    ).commands

    expect(command).toMatchObject({
      operation: 'markdown.document.load_staged',
      arguments: { name: 'task-draft.md', size: bytes.length },
    })
  })

  it.each([
    ['docx', 'generated.docx'],
    ['markdown', 'generated.md'],
    ['xlsx', 'generated.xlsx'],
    ['pptx', 'generated.pptx'],
    ['pdf', 'generated.pdf'],
  ] as const)(
    'persists renderer-produced %s bytes through the app-only save protocol',
    async (format, fileName) => {
      const root = await mkdtemp(join(tmpdir(), 'tandemfolio-document-save-protocol-'))
      temporaryDirectories.push(root)
      const client = await connectClient(join(root, 'state'), {
        TANDEMFOLIO_OUTPUT_DIR: join(root, 'outputs'),
      })
      const created = await client.callTool({
        name: 'office_create_session',
        arguments: { format },
      })
      const sessionId = (created.structuredContent as { session: { id: string } }).session.id
      const viewId = `view-${format}`
      await client.callTool({ name: 'office_editor_poll', arguments: { sessionId, viewId } })
      const data = Buffer.from(`renderer-produced-${format}`)

      const begun = await client.callTool({
        name: 'office_editor_begin_document_save',
        arguments: { sessionId, viewId, fileName, size: data.length, mode: 'save' },
      })
      const { uploadId, path } = begun.structuredContent as { uploadId: string; path: string }
      const written = await client.callTool({
        name: 'office_editor_write_document_save_chunk',
        arguments: { sessionId, viewId, uploadId, offset: 0, data: data.toString('base64') },
      })
      const committed = await client.callTool({
        name: 'office_editor_commit_document_save',
        arguments: { sessionId, viewId, uploadId },
      })

      expect(written.structuredContent).toMatchObject({ ok: true, nextOffset: data.length })
      expect(committed.structuredContent).toMatchObject({ ok: true, path })
      const context = await client.callTool({
        name: 'office_get_context',
        arguments: { sessionId },
      })
      expect(context.structuredContent).toMatchObject({ session: { filePath: path } })
      await expect(readFile(path)).resolves.toEqual(data)
    },
  )

  it.each([
    ['docx', 'saved.docx'],
    ['markdown', 'saved.md'],
    ['xlsx', 'saved.xlsx'],
    ['pptx', 'saved.pptx'],
    ['pdf', 'saved.pdf'],
  ] as const)(
    'restores saved %s bytes after remount and broker restart without a recovery snapshot',
    async (format, fileName) => {
      const root = await mkdtemp(join(tmpdir(), 'tandemfolio-saved-resume-'))
      temporaryDirectories.push(root)
      const environment = { TANDEMFOLIO_OUTPUT_DIR: join(root, 'outputs') }
      let client = await connectClient(join(root, 'state'), environment)
      const created = await client.callTool({
        name: 'office_create_session',
        arguments: { format },
      })
      const sessionId = (created.structuredContent as { session: { id: string } }).session.id
      let viewId = 'initial-view'
      await client.callTool({ name: 'office_editor_poll', arguments: { sessionId, viewId } })
      const data = Buffer.from(`saved-by-${format}-renderer`)
      const begun = await client.callTool({
        name: 'office_editor_begin_document_save',
        arguments: {
          sessionId,
          viewId,
          fileName,
          size: data.length,
        },
      })
      const { uploadId, path } = begun.structuredContent as { uploadId: string; path: string }
      await client.callTool({
        name: 'office_editor_write_document_save_chunk',
        arguments: {
          sessionId,
          viewId,
          uploadId,
          offset: 0,
          data: data.toString('base64'),
        },
      })
      await client.callTool({
        name: 'office_editor_commit_document_save',
        arguments: { sessionId, viewId, uploadId },
      })
      for (const restart of [false, true, 'automatic'] as const) {
        await client.callTool({
          name: 'office_editor_disconnect',
          arguments: { sessionId, viewId },
        })
        if (restart) {
          await client.close()
          clients.splice(clients.indexOf(client), 1)
          client = await connectClient(join(root, 'state'), environment)
          if (restart !== 'automatic') {
            const resumed = await client.callTool({
              name: 'office_create_session',
              arguments: { format, sessionId, resume: 'exact' },
            })
            expect(resumed.structuredContent).toMatchObject({
              ok: true,
              session: { id: sessionId, filePath: path },
            })
          }
        }
        viewId = `resumed-${restart}`
        const response = await client.callTool({
          name: 'office_editor_poll',
          arguments: {
            sessionId,
            viewId,
            format,
            coldStart: true,
            fileName: 'Untitled',
            dirty: false,
          },
        })
        expect(response.structuredContent).toMatchObject({
          ok: true,
          commands: [
            { operation: `${format}.document.load_staged`, arguments: { name: fileName } },
          ],
        })
        const command = (
          response.structuredContent as {
            commands: Array<{
              commandId: string
              baseRevision: number
              arguments: { blobId: string }
            }>
          }
        ).commands[0]
        const chunk = await client.callTool({
          name: 'office_editor_read_file_chunk',
          arguments: {
            sessionId,
            viewId,
            blobId: command.arguments.blobId,
            offset: 0,
            length: 262144,
          },
        })
        expect(chunk.structuredContent).toMatchObject({ ok: true, data: data.toString('base64') })
        await client.callTool({
          name: 'office_editor_acknowledge',
          arguments: {
            sessionId,
            viewId,
            commandId: command.commandId,
            revision: command.baseRevision + 1,
            fileName,
            dirty: false,
          },
        })
      }
      await client.callTool({ name: 'office_editor_disconnect', arguments: { sessionId, viewId } })
      await rm(path)
      const missing = await client.callTool({
        name: 'office_editor_poll',
        arguments: { sessionId, viewId: 'missing-file', format, coldStart: true },
      })
      expect(missing.structuredContent).toMatchObject({
        ok: false,
        error: 'document_restore_failed',
      })
      const retried = await client.callTool({
        name: 'office_editor_poll',
        arguments: { sessionId, viewId: 'missing-file', format, coldStart: true },
      })
      expect(retried.structuredContent).toMatchObject({
        ok: false,
        error: 'document_restore_failed',
      })
    },
  )

  it('rejects a duplicate mount replaying the same show identity without changing document context', async () => {
    const client = await connectClient()
    const created = await client.callTool({
      name: 'office_create_session',
      arguments: { format: 'pptx' },
    })
    const sessionId = (created.structuredContent as { session: { id: string } }).session.id
    await client.callTool({
      name: 'office_editor_poll',
      arguments: {
        sessionId,
        viewId: 'same-show',
        mountId: 'owner',
        fileName: 'Owner.pptx',
        dirty: true,
      },
    })
    const duplicate = await client.callTool({
      name: 'office_editor_poll',
      arguments: {
        sessionId,
        viewId: 'same-show',
        mountId: 'duplicate',
        fileName: 'Untitled.pptx',
        dirty: false,
      },
    })
    expect(duplicate.structuredContent).toMatchObject({ ok: false, error: 'editor_view_conflict' })
    const context = await client.callTool({ name: 'office_get_context', arguments: { sessionId } })
    expect(context.structuredContent).toMatchObject({
      session: { fileName: 'Owner.pptx', dirty: true },
    })
  })

  it.each([
    ['docx', 'docx'],
    ['markdown', 'md'],
    ['xlsx', 'xlsx'],
    ['pptx', 'pptx'],
    ['pdf', 'pdf'],
  ] as const)(
    'browser replacement in %s cannot overwrite the previous document, even with the same name',
    async (format, extension) => {
      const root = await mkdtemp(join(tmpdir(), 'tandemfolio-replace-'))
      temporaryDirectories.push(root)
      const client = await connectClient(join(root, 'state'), {
        TANDEMFOLIO_OUTPUT_DIR: join(root, 'outputs'),
      })
      const created = await client.callTool({
        name: 'office_create_session',
        arguments: { format },
      })
      const sessionId = (created.structuredContent as { session: { id: string } }).session.id
      const viewId = 'replace-view'
      await client.callTool({ name: 'office_editor_poll', arguments: { sessionId, viewId } })
      const save = async (text: string) => {
        const data = Buffer.from(text)
        const begun = await client.callTool({
          name: 'office_editor_begin_document_save',
          arguments: { sessionId, viewId, fileName: `same.${extension}`, size: data.length },
        })
        const { uploadId } = begun.structuredContent as { uploadId: string }
        await client.callTool({
          name: 'office_editor_write_document_save_chunk',
          arguments: { sessionId, viewId, uploadId, offset: 0, data: data.toString('base64') },
        })
        const committed = await client.callTool({
          name: 'office_editor_commit_document_save',
          arguments: { sessionId, viewId, uploadId },
        })
        return (committed.structuredContent as { path: string }).path
      }
      const firstPath = await save('first-document')
      const reset = await client.callTool({
        name: 'office_editor_reset_document',
        arguments: { sessionId, viewId },
      })
      expect(reset.structuredContent).toMatchObject({ ok: true, filePath: null })
      const secondPath = await save('second-document')
      expect(secondPath).not.toBe(firstPath)
      await expect(readFile(firstPath, 'utf8')).resolves.toBe('first-document')
      await expect(readFile(secondPath, 'utf8')).resolves.toBe('second-document')
    },
  )

  it('retires pre-save recovery and rejects a late checkpoint so a remount loads the committed file', async () => {
    const root = await mkdtemp(join(tmpdir(), 'tandemfolio-save-recovery-'))
    temporaryDirectories.push(root)
    const client = await connectClient(join(root, 'state'), {
      TANDEMFOLIO_OUTPUT_DIR: join(root, 'outputs'),
    })
    const created = await client.callTool({
      name: 'office_create_session',
      arguments: { format: 'markdown' },
    })
    const sessionId = (created.structuredContent as { session: { id: string } }).session.id
    const viewId = 'save-recovery'
    await client.callTool({ name: 'office_editor_poll', arguments: { sessionId, viewId } })
    const draft = Buffer.from('old draft')
    const beginRecovery = async () => {
      const begun = await client.callTool({
        name: 'office_editor_begin_recovery',
        arguments: { sessionId, viewId, fileName: 'draft.md', size: draft.length },
      })
      const { uploadId } = begun.structuredContent as { uploadId: string }
      await client.callTool({
        name: 'office_editor_write_recovery_chunk',
        arguments: { sessionId, viewId, uploadId, offset: 0, data: draft.toString('base64') },
      })
      return uploadId
    }
    await client.callTool({
      name: 'office_editor_commit_recovery',
      arguments: { sessionId, viewId, uploadId: await beginRecovery() },
    })
    const lateRecovery = await beginRecovery()
    const saved = Buffer.from('latest saved contents')
    const begun = await client.callTool({
      name: 'office_editor_begin_document_save',
      arguments: { sessionId, viewId, fileName: 'saved.md', size: saved.length },
    })
    const { uploadId } = begun.structuredContent as { uploadId: string }
    await client.callTool({
      name: 'office_editor_write_document_save_chunk',
      arguments: { sessionId, viewId, uploadId, offset: 0, data: saved.toString('base64') },
    })
    await client.callTool({
      name: 'office_editor_commit_document_save',
      arguments: { sessionId, viewId, uploadId },
    })
    const late = await client.callTool({
      name: 'office_editor_commit_recovery',
      arguments: { sessionId, viewId, uploadId: lateRecovery },
    })
    expect(late.structuredContent).toMatchObject({ ok: false, error: 'command_not_found' })
    await client.callTool({ name: 'office_editor_disconnect', arguments: { sessionId, viewId } })
    const resumed = await client.callTool({
      name: 'office_editor_poll',
      arguments: { sessionId, viewId: 'new-view' },
    })
    expect(resumed.structuredContent).toMatchObject({
      commands: [{ arguments: { name: 'saved.md', size: saved.length } }],
    })
  })

  it('binds a successfully opened file so Save atomically overwrites that exact path', async () => {
    const root = await mkdtemp(join(tmpdir(), 'tandemfolio-open-file-save-'))
    temporaryDirectories.push(root)
    const sourcePath = join(root, 'opened.docx')
    await writeFile(sourcePath, 'original')
    const client = await connectClient(join(root, 'state'), {
      TANDEMFOLIO_OUTPUT_DIR: join(root, 'outputs'),
    })
    const created = await client.callTool({
      name: 'office_create_session',
      arguments: { format: 'docx' },
    })
    const sessionId = (created.structuredContent as { session: { id: string } }).session.id
    const viewId = 'opened-file-view'
    await client.callTool({ name: 'office_editor_poll', arguments: { sessionId, viewId } })

    const opening = client.callTool({
      name: 'office_open_local_file',
      arguments: { sessionId, baseRevision: 0, path: sourcePath },
    })
    const polled = await client.callTool({
      name: 'office_editor_poll',
      arguments: { sessionId, viewId, waitMs: 2_000 },
    })
    const [command] = (polled.structuredContent as { commands: Array<{ commandId: string }> })
      .commands
    await client.callTool({
      name: 'office_editor_acknowledge',
      arguments: {
        sessionId,
        viewId,
        commandId: command.commandId,
        ok: true,
        revision: 1,
        output: { loaded: true },
      },
    })
    await opening

    const data = Buffer.from('edited')
    const begun = await client.callTool({
      name: 'office_editor_begin_document_save',
      arguments: {
        sessionId,
        viewId,
        fileName: 'mcp-local:renderer-blob',
        size: data.length,
        mode: 'save',
      },
    })
    const { uploadId } = begun.structuredContent as { uploadId: string }
    await client.callTool({
      name: 'office_editor_write_document_save_chunk',
      arguments: { sessionId, viewId, uploadId, offset: 0, data: data.toString('base64') },
    })
    const committed = await client.callTool({
      name: 'office_editor_commit_document_save',
      arguments: { sessionId, viewId, uploadId },
    })

    expect(committed.structuredContent).toMatchObject({ ok: true, path: sourcePath })
    const context = await client.callTool({
      name: 'office_get_context',
      arguments: { sessionId },
    })
    expect(context.structuredContent).toMatchObject({ session: { filePath: sourcePath } })
    await expect(readFile(sourcePath, 'utf8')).resolves.toBe('edited')
  })
})
