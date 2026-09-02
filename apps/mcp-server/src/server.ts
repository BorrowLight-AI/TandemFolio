import { readFile } from 'node:fs/promises'
import { randomUUID } from 'node:crypto'
import { homedir } from 'node:os'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js'
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js'
import { validateJsonSchemaValue } from '@tandemfolio/operation-contract'
import { z } from 'zod'
import { LocalFileStore } from './local-file-store'
import { FontAssetStore } from './font-asset-store'
import { DocumentSaveStore } from './document-save-store'
import { getFormatCapabilities } from './capabilities'
import { getCanonicalRegisteredOperation, resolveRegisteredOperation } from './operation-manifest'
import { RecoveryStore, RecoveryUploadStore, type RecoverySnapshot } from './recovery-store'
import { SessionError, SessionStore, type LiveSession } from './session-store'
import { TransactionJournal, type TransactionJournalStart } from './transaction-journal'

const RESOURCE_URI = 'ui://tandemfolio/editor.html'
const XLSX_RESOURCE_URI = 'ui://tandemfolio/xlsx.html'
const PPTX_RESOURCE_URI = 'ui://tandemfolio/pptx.html'
const PDF_RESOURCE_URI = 'ui://tandemfolio/pdf.html'
const MARKDOWN_RESOURCE_URI = 'ui://tandemfolio/markdown.html'
const RESOURCE_MIME_TYPE = 'text/html;profile=mcp-app'
const store = new SessionStore()
const localFiles = new LocalFileStore()
const fontAssets = new FontAssetStore()
const stateDirectory =
  process.env.TANDEMFOLIO_STATE_DIR ?? join(homedir(), '.tandemfolio', 'recovery')
const recoveryStore = new RecoveryStore(stateDirectory)
const recoveryUploads = new RecoveryUploadStore(recoveryStore)
const documentSaves = new DocumentSaveStore(
  process.env.TANDEMFOLIO_OUTPUT_DIR ?? join(homedir(), 'Documents', 'TandemFolio'),
  join(stateDirectory, 'document-bindings'),
)
const pendingRecoveries = new Map<string, RecoverySnapshot>()
const configuredCommandTimeoutMs = Number(process.env.TANDEMFOLIO_COMMAND_TIMEOUT_MS)
const COMMAND_TIMEOUT_MS =
  Number.isInteger(configuredCommandTimeoutMs) &&
  configuredCommandTimeoutMs >= 1 &&
  configuredCommandTimeoutMs <= 60_000
    ? configuredCommandTimeoutMs
    : 15_000

function result(value: Record<string, unknown>) {
  return {
    content: [{ type: 'text' as const, text: JSON.stringify(value) }],
    structuredContent: value,
  }
}

type ToolResponse = ReturnType<typeof result> | ReturnType<typeof failure>
type NewTransaction = Extract<TransactionJournalStart<ToolResponse>, { kind: 'new' }>

const transactionJournal = new TransactionJournal<ToolResponse>()

function failure(error: unknown) {
  const code = error instanceof SessionError ? error.code : 'internal_error'
  const message = error instanceof Error ? error.message : String(error)
  return {
    isError: true,
    content: [{ type: 'text' as const, text: `${code}: ${message}` }],
    structuredContent: { ok: false, error: code, message },
  }
}

function waitForTransactionResponse(
  promise: Promise<ToolResponse>,
  requestId: string,
): Promise<ToolResponse> {
  return new Promise<ToolResponse>((resolve, reject) => {
    const timer = setTimeout(() => {
      reject(
        new SessionError(
          'command_timeout',
          `Transaction request ${requestId} was not acknowledged within ${COMMAND_TIMEOUT_MS} ms.`,
        ),
      )
    }, COMMAND_TIMEOUT_MS)
    promise.then(
      (response) => {
        clearTimeout(timer)
        resolve(response)
      },
      (error) => {
        clearTimeout(timer)
        reject(error)
      },
    )
  })
}

function deferStagedFileRelease(transaction: NewTransaction, blobId: string): undefined {
  void transaction.promise.then(() => localFiles.release(blobId))
  return undefined
}

const STAGED_FILE_OPERATIONS = {
  docx: 'docx.document.load_staged',
  markdown: 'markdown.document.load_staged',
  xlsx: 'xlsx.document.load_staged',
  pptx: 'pptx.document.load_staged',
  pdf: 'pdf.document.load_staged',
} as const satisfies Record<LiveSession['format'], string>

function stagedFileOperation(format: LiveSession['format']): string {
  return STAGED_FILE_OPERATIONS[format]
}

function stagedImageOperation(
  format: 'docx' | 'markdown' | 'xlsx',
  sourceOperation: string,
): string {
  const operation =
    format === 'docx'
      ? sourceOperation === 'docx.image.replace'
        ? 'docx.image.replace_staged'
        : 'docx.image.insert_staged'
      : format === 'markdown'
        ? 'markdown.image.insert_staged'
        : 'xlsx.image.add_staged'
  return resolveRegisteredOperation(format, operation, 'internal')?.id ?? operation
}

function stagedCompareOperation(): string {
  return (
    resolveRegisteredOperation('docx', 'docx.document.compare_staged', 'internal')?.id ??
    'docx.document.compare_staged'
  )
}

function stagedPdfPageInsertOperation(): string {
  return (
    resolveRegisteredOperation('pdf', 'pdf.page.insert_staged', 'internal')?.id ??
    'pdf.page.insert_staged'
  )
}

async function editorHtml(
  format: 'docx' | 'markdown' | 'xlsx' | 'pptx' | 'pdf' = 'docx',
): Promise<string> {
  const candidates =
    format === 'docx'
      ? [
          new URL('../assets/editor/index.html', import.meta.url),
          new URL('../../../plugins/tandemfolio/assets/editor/index.html', import.meta.url),
        ]
      : format === 'xlsx'
        ? [
            new URL('../assets/editors/xlsx/index.html', import.meta.url),
            new URL(
              '../../../plugins/tandemfolio/assets/editors/xlsx/index.html',
              import.meta.url,
            ),
          ]
        : format === 'pptx'
          ? [
              new URL('../assets/editors/pptx/index.html', import.meta.url),
              new URL(
                '../../../plugins/tandemfolio/assets/editors/pptx/index.html',
                import.meta.url,
              ),
            ]
          : format === 'pdf'
            ? [
                new URL('../assets/editors/pdf/index.html', import.meta.url),
                new URL(
                  '../../../plugins/tandemfolio/assets/editors/pdf/index.html',
                  import.meta.url,
                ),
              ]
            : [
                new URL('../assets/editors/markdown/index.html', import.meta.url),
                new URL(
                  '../../../plugins/tandemfolio/assets/editors/markdown/index.html',
                  import.meta.url,
                ),
              ]

  for (const candidate of candidates) {
    try {
      return await readFile(fileURLToPath(candidate), 'utf8')
    } catch (error) {
      if ((error as NodeJS.ErrnoException).code !== 'ENOENT') {
        throw error
      }
    }
  }

  return '<!doctype html><html><body><main><h1>TandemFolio</h1><p>Run npm run build to package the editor.</p></main></body></html>'
}

const server = new McpServer({ name: 'tandemfolio', version: '0.1.0' })

server.registerResource(
  'tandemfolio-editor',
  RESOURCE_URI,
  {
    title: 'TandemFolio editor',
    description: 'Persistent live Office editing surface.',
    mimeType: RESOURCE_MIME_TYPE,
  },
  async () => ({
    contents: [
      {
        uri: RESOURCE_URI,
        mimeType: RESOURCE_MIME_TYPE,
        text: await editorHtml(),
        _meta: {
          ui: {
            prefersBorder: false,
            csp: { connectDomains: [], resourceDomains: [] },
          },
        },
      },
    ],
  }),
)

server.registerResource(
  'tandemfolio-xlsx-editor',
  XLSX_RESOURCE_URI,
  {
    title: 'TandemFolio XLSX editor',
    description: 'Persistent live spreadsheet editing surface.',
    mimeType: RESOURCE_MIME_TYPE,
  },
  async () => ({
    contents: [
      {
        uri: XLSX_RESOURCE_URI,
        mimeType: RESOURCE_MIME_TYPE,
        text: await editorHtml('xlsx'),
        _meta: {
          ui: {
            prefersBorder: false,
            csp: { connectDomains: [], resourceDomains: [] },
          },
        },
      },
    ],
  }),
)

server.registerResource(
  'tandemfolio-pptx-editor',
  PPTX_RESOURCE_URI,
  {
    title: 'TandemFolio PPTX editor',
    description: 'Persistent live presentation editing surface.',
    mimeType: RESOURCE_MIME_TYPE,
  },
  async () => ({
    contents: [
      {
        uri: PPTX_RESOURCE_URI,
        mimeType: RESOURCE_MIME_TYPE,
        text: await editorHtml('pptx'),
        _meta: { ui: { prefersBorder: false, csp: { connectDomains: [], resourceDomains: [] } } },
      },
    ],
  }),
)

server.registerResource(
  'tandemfolio-pdf-editor',
  PDF_RESOURCE_URI,
  {
    title: 'TandemFolio PDF editor',
    description: 'Persistent local PDF viewing and safe editing surface.',
    mimeType: RESOURCE_MIME_TYPE,
  },
  async () => ({
    contents: [
      {
        uri: PDF_RESOURCE_URI,
        mimeType: RESOURCE_MIME_TYPE,
        text: await editorHtml('pdf'),
        _meta: { ui: { prefersBorder: false, csp: { connectDomains: [], resourceDomains: [] } } },
      },
    ],
  }),
)

server.registerResource(
  'tandemfolio-markdown-editor',
  MARKDOWN_RESOURCE_URI,
  {
    title: 'TandemFolio Markdown editor',
    description: 'Persistent live Markdown editing surface.',
    mimeType: RESOURCE_MIME_TYPE,
  },
  async () => ({
    contents: [
      {
        uri: MARKDOWN_RESOURCE_URI,
        mimeType: RESOURCE_MIME_TYPE,
        text: await editorHtml('markdown'),
        _meta: { ui: { prefersBorder: false, csp: { connectDomains: [], resourceDomains: [] } } },
      },
    ],
  }),
)

server.registerTool(
  'office_create_session',
  {
    title: 'Create Office editing session',
    description: 'Create a live TandemFolio session before opening its editor.',
    inputSchema: {
      format: z.enum(['docx', 'markdown', 'xlsx', 'pptx', 'pdf']).default('docx'),
      resume: z.enum(['latest', 'exact', 'none']).default('none'),
      sessionId: z.string().min(1).optional(),
    },
  },
  async ({ format, resume, sessionId }) => {
    try {
      if (resume === 'exact') {
        if (!sessionId) {
          throw new SessionError(
            'invalid_arguments',
            'resume: exact requires the prior editing sessionId.',
          )
        }
        try {
          const session = store.get(sessionId)
          if (session.format !== format) {
            throw new SessionError(
              'invalid_arguments',
              `Session ${sessionId} belongs to ${session.format}, not ${format}.`,
            )
          }
          return result({ ok: true, session, recoveryAvailable: false, reused: true })
        } catch (error) {
          if (!(error instanceof SessionError && error.code === 'session_not_found')) throw error
        }
        const recovery = await recoveryStore.latest(format, sessionId)
        if (!recovery) {
          throw new SessionError(
            'session_not_found',
            `No live session or exact recovery exists for ${sessionId}.`,
          )
        }
        const session = store.restore(sessionId, format)
        pendingRecoveries.set(session.id, recovery)
        return result({ ok: true, session, recoveryAvailable: true, reused: false })
      }
      if (sessionId) {
        throw new SessionError(
          'invalid_arguments',
          'sessionId is accepted only with resume: exact.',
        )
      }
      const session = store.create(format)
      if (resume === 'latest') {
        const recovery = await recoveryStore.latest(format)
        if (recovery) pendingRecoveries.set(session.id, recovery)
      }
      return result({
        ok: true,
        session,
        recoveryAvailable: pendingRecoveries.has(session.id),
        reused: false,
      })
    } catch (error) {
      return failure(error)
    }
  },
)

server.registerTool(
  'office_show_editor',
  {
    title: 'Show Office editor',
    description: 'Open the persistent TandemFolio editor for an existing session.',
    inputSchema: { sessionId: z.string().min(1) },
    _meta: { ui: { resourceUri: RESOURCE_URI } },
  },
  async ({ sessionId }) => {
    try {
      const session = store.get(sessionId)
      if (session.format !== 'docx') {
        throw new SessionError('invalid_arguments', 'office_show_editor requires a DOCX session.')
      }
      return result({
        ok: true,
        sessionId,
        viewId: randomUUID(),
        format: session.format,
        revision: session.revision,
      })
    } catch (error) {
      return failure(error)
    }
  },
)

server.registerTool(
  'office_show_xlsx_editor',
  {
    title: 'Show XLSX editor',
    description: 'Open the persistent TandemFolio XLSX editor for an existing XLSX session.',
    inputSchema: { sessionId: z.string().min(1) },
    _meta: { ui: { resourceUri: XLSX_RESOURCE_URI } },
  },
  async ({ sessionId }) => {
    try {
      const session = store.get(sessionId)
      if (session.format !== 'xlsx') {
        throw new SessionError(
          'invalid_arguments',
          'office_show_xlsx_editor requires an XLSX session.',
        )
      }
      return result({
        ok: true,
        sessionId,
        viewId: randomUUID(),
        format: session.format,
        revision: session.revision,
      })
    } catch (error) {
      return failure(error)
    }
  },
)

server.registerTool(
  'office_show_pptx_editor',
  {
    title: 'Show PPTX editor',
    description: 'Open the persistent TandemFolio PPTX editor for an existing PPTX session.',
    inputSchema: { sessionId: z.string().min(1) },
    _meta: { ui: { resourceUri: PPTX_RESOURCE_URI } },
  },
  async ({ sessionId }) => {
    try {
      const session = store.get(sessionId)
      if (session.format !== 'pptx')
        throw new SessionError(
          'invalid_arguments',
          'office_show_pptx_editor requires a PPTX session.',
        )
      return result({
        ok: true,
        sessionId,
        viewId: randomUUID(),
        format: session.format,
        revision: session.revision,
      })
    } catch (error) {
      return failure(error)
    }
  },
)

server.registerTool(
  'office_show_pdf_editor',
  {
    title: 'Show PDF editor',
    description: 'Open the persistent TandemFolio PDF editor for an existing PDF session.',
    inputSchema: { sessionId: z.string().min(1) },
    _meta: { ui: { resourceUri: PDF_RESOURCE_URI } },
  },
  async ({ sessionId }) => {
    try {
      const session = store.get(sessionId)
      if (session.format !== 'pdf')
        throw new SessionError(
          'invalid_arguments',
          'office_show_pdf_editor requires a PDF session.',
        )
      return result({
        ok: true,
        sessionId,
        viewId: randomUUID(),
        format: session.format,
        revision: session.revision,
      })
    } catch (error) {
      return failure(error)
    }
  },
)

server.registerTool(
  'office_show_markdown_editor',
  {
    title: 'Show Markdown editor',
    description:
      'Open the persistent TandemFolio Markdown editor for an existing Markdown session.',
    inputSchema: { sessionId: z.string().min(1) },
    _meta: { ui: { resourceUri: MARKDOWN_RESOURCE_URI } },
  },
  async ({ sessionId }) => {
    try {
      const session = store.get(sessionId)
      if (session.format !== 'markdown') {
        throw new SessionError(
          'invalid_arguments',
          'office_show_markdown_editor requires a Markdown session.',
        )
      }
      return result({
        ok: true,
        sessionId,
        viewId: randomUUID(),
        format: session.format,
        revision: session.revision,
      })
    } catch (error) {
      return failure(error)
    }
  },
)

server.registerTool(
  'office_get_capabilities',
  {
    title: 'Get Office format capabilities',
    description: 'Read bounded operation summaries or one exact operation descriptor for a format.',
    inputSchema: {
      format: z.enum(['docx', 'markdown', 'xlsx', 'pptx', 'pdf']).default('docx'),
      view: z.enum(['summary', 'detail']).default('summary'),
      operation: z.string().min(1).max(256).optional(),
      sessionId: z.string().min(1).optional(),
      family: z.string().min(1).max(64).optional(),
      limit: z.number().int().min(1).max(20).optional(),
      cursor: z.string().min(1).max(256).nullable().optional(),
    },
  },
  async ({ format, view, operation, sessionId, family, limit, cursor }) => {
    try {
      const capabilities = getFormatCapabilities(format, {
        view,
        operation,
        ...(sessionId ? { session: store.get(sessionId) } : {}),
        family,
        limit,
        cursor,
      })
      if (!capabilities) {
        throw new SessionError('invalid_arguments', `Invalid ${format} capability discovery query.`)
      }
      return result({ ok: true, capabilities })
    } catch (error) {
      return failure(error)
    }
  },
)

server.registerTool(
  'office_get_context',
  {
    title: 'Get Office editor context',
    description: 'Read the current revision, active file, dirty state, and selection.',
    inputSchema: { sessionId: z.string().min(1) },
  },
  async ({ sessionId }) => {
    try {
      const session = store.get(sessionId)
      session.filePath ??= await documentSaves.boundPath(sessionId, session.format)
      return result({ ok: true, session })
    } catch (error) {
      return failure(error)
    }
  },
)

server.registerTool(
  'office_execute',
  {
    title: 'Execute Office editor operation',
    description:
      'Execute a revision-guarded, idempotently replayable live editor transaction and wait for acknowledgement.',
    inputSchema: {
      sessionId: z.string().min(1),
      baseRevision: z.number().int().nonnegative(),
      requestId: z.string().min(1).max(128),
      operations: z
        .array(
          z.object({
            id: z.string().min(1).max(256),
            arguments: z.record(z.string(), z.unknown()).default({}),
          }),
        )
        .min(1)
        .max(32),
    },
  },
  async ({ sessionId, baseRevision, requestId, operations }) => {
    let transaction: NewTransaction | undefined
    try {
      if (operations.length > 1) {
        throw new SessionError(
          'transaction_not_atomic',
          'No current format registry declares a multi-operation atomic transaction.',
        )
      }
      const requestedOperation = operations[0]!.id
      const args = operations[0]!.arguments
      const started = transactionJournal.start(sessionId, requestId, {
        baseRevision,
        operations,
      })
      if (started.kind === 'replay') {
        return await waitForTransactionResponse(started.promise, requestId)
      }
      if (started.kind === 'conflict') {
        throw new SessionError(
          'request_reused',
          `Request ${requestId} was already used with a different transaction payload.`,
        )
      }
      transaction = started
      const activeTransaction = started
      const enqueue = (queuedOperation: string, queuedArguments: Record<string, unknown>) => {
        const command = store.enqueue(sessionId, baseRevision, queuedOperation, queuedArguments)
        activeTransaction.markAccepted()
        return command
      }
      const executionResult = (
        completion: { revision: number; output?: Record<string, unknown> },
        canonicalOperation: string,
      ) =>
        result({
          ok: true,
          transaction: {
            transactionId: activeTransaction.transactionId,
            requestId,
            baseRevision,
          },
          result: {
            revision: completion.revision,
            operations: [{ id: canonicalOperation, result: completion.output ?? {} }],
          },
        })
      const waitForExecution = async (
        command: { commandId: string; baseRevision: number },
        canonicalOperation: string,
      ): Promise<ToolResponse> => {
        void store
          .waitForCommand(sessionId, command.commandId, null)
          .then(async (completion) => {
            if (completion.output?.saved === true) {
              await recoveryStore.clear(store.get(sessionId).format, sessionId)
            }
            activeTransaction.complete(executionResult(completion, canonicalOperation))
          })
          .catch((error) => activeTransaction.fail(failure(error)))
        return await waitForTransactionResponse(activeTransaction.promise, requestId)
      }
      const session = store.assertCanEnqueue(sessionId, baseRevision)
      const descriptor = getCanonicalRegisteredOperation(session.format, requestedOperation)
      if (!descriptor) {
        throw new SessionError(
          'operation_not_found',
          `Operation ${requestedOperation} is not a canonical Agent operation for ${session.format}.`,
        )
      }
      const validation = validateJsonSchemaValue(descriptor.inputSchema, args)
      if (!validation.ok) {
        throw new SessionError(
          'operation_schema_invalid',
          `Invalid arguments for ${descriptor.id}: ${validation.error.message}`,
        )
      }
      if (descriptor.context.includes('selection') && session.selection === null) {
        throw new SessionError(
          'operation_unavailable',
          `Operation ${descriptor.id} requires a current editor selection.`,
        )
      }
      if (
        (session.format === 'xlsx' && descriptor.id === 'xlsx.image.add') ||
        (session.format === 'markdown' && descriptor.id === 'markdown.image.insert') ||
        (session.format === 'docx' &&
          (descriptor.id === 'docx.image.insert' || descriptor.id === 'docx.image.replace'))
      ) {
        let blobId: string | undefined
        try {
          const staged = await localFiles.stageImage(sessionId, args.path as string)
          blobId = staged.blobId
          blobId = deferStagedFileRelease(activeTransaction, blobId)
          const { path: _path, ...parameters } = args
          const command = enqueue(stagedImageOperation(session.format, descriptor.id), {
            ...staged,
            ...parameters,
          })
          return await waitForExecution(command, descriptor.id)
        } finally {
          if (blobId) localFiles.release(blobId)
        }
      }
      if (session.format === 'docx' && descriptor.id === 'docx.document.compare') {
        let blobId: string | undefined
        try {
          const staged = await localFiles.stage(sessionId, 'docx', args.path as string)
          blobId = staged.blobId
          blobId = deferStagedFileRelease(activeTransaction, blobId)
          const command = enqueue(stagedCompareOperation(), {
            ...staged,
          })
          return await waitForExecution(command, descriptor.id)
        } finally {
          if (blobId) localFiles.release(blobId)
        }
      }
      if (session.format === 'pdf' && descriptor.id === 'pdf.page.insert') {
        let blobId: string | undefined
        try {
          const staged = await localFiles.stage(sessionId, 'pdf', args.path as string)
          blobId = staged.blobId
          blobId = deferStagedFileRelease(activeTransaction, blobId)
          const command = enqueue(stagedPdfPageInsertOperation(), {
            ...staged,
            afterPageIndex: args.afterPageIndex,
          })
          return await waitForExecution(command, descriptor.id)
        } finally {
          if (blobId) localFiles.release(blobId)
        }
      }
      const command = enqueue(descriptor.id, args)
      return await waitForExecution(command, descriptor.id)
    } catch (error) {
      const response = failure(error)
      if (!(error instanceof SessionError && error.code === 'command_timeout')) {
        transaction?.fail(response)
      }
      return response
    }
  },
)

server.registerTool(
  'office_open_local_file',
  {
    title: 'Open local Office file',
    description:
      'Open an absolute local file path in the mounted live editor and wait for it to load.',
    inputSchema: {
      sessionId: z.string().min(1),
      baseRevision: z.number().int().nonnegative(),
      path: z.string().min(1),
    },
  },
  async ({ sessionId, baseRevision, path }) => {
    let blobId: string | undefined
    try {
      const session = store.get(sessionId)
      const staged = await localFiles.stage(sessionId, session.format, path)
      blobId = staged.blobId
      const assetRootId =
        session.format === 'markdown' ? localFiles.registerAssetRoot(sessionId, path) : undefined
      const command = store.enqueue(sessionId, baseRevision, stagedFileOperation(session.format), {
        ...staged,
        ...(assetRootId ? { assetRootId } : {}),
      })
      const completion = await store.waitForCommand(sessionId, command.commandId)
      await documentSaves.bind(sessionId, session.format, path)
      session.filePath = path
      return result({ ok: true, command, result: completion })
    } catch (error) {
      return failure(error)
    } finally {
      if (blobId) localFiles.release(blobId)
    }
  },
)

server.registerTool(
  'office_editor_poll',
  {
    title: 'Poll editor commands',
    description: 'Editor-only transport endpoint for receiving queued commands.',
    inputSchema: {
      sessionId: z.string().min(1),
      viewId: z.string().min(1).optional(),
      fileName: z.string().nullable().optional(),
      dirty: z.boolean().optional(),
      selection: z.record(z.string(), z.unknown()).nullable().optional(),
      startupTrace: z
        .object({
          operation: z.literal('xlsx.editor.cold_start'),
          phases: z
            .object({
              bootstrapMs: z.number().finite().min(0).max(3_600_000),
              univerCreateMs: z.number().finite().min(0).max(3_600_000),
              worksheetInstallMs: z.number().finite().min(0).max(3_600_000),
              firstCommitMs: z.number().finite().min(0).max(3_600_000),
            })
            .strict(),
          bootstrapPhases: z
            .object({
              resourceReceiveMs: z.number().finite().min(0).max(3_600_000),
              moduleGraphReadyMs: z.number().finite().min(0).max(3_600_000),
              reactMountMs: z.number().finite().min(0).max(3_600_000),
            })
            .strict(),
        })
        .strict()
        .optional(),
      waitMs: z.number().int().min(0).max(10_000).default(0),
    },
    _meta: { ui: { visibility: ['app'] } },
  },
  async ({ sessionId, viewId, fileName, dirty, selection, waitMs }) => {
    try {
      const sessionBeforePoll = store.get(sessionId)
      const reconnecting = !sessionBeforePoll.connected
      const context = {
        ...(fileName !== undefined ? { fileName } : {}),
        ...(dirty !== undefined ? { dirty } : {}),
        ...(selection !== undefined ? { selection } : {}),
      }
      let commands = store.poll(sessionId, context, viewId)
      const recovery =
        pendingRecoveries.get(sessionId) ??
        (reconnecting ? await recoveryStore.latest(sessionBeforePoll.format, sessionId) : null)
      if (commands.length === 0 && recovery) {
        const staged = localFiles.stageBuffer(sessionId, recovery.fileName, recovery.data)
        const session = store.get(sessionId)
        store.enqueue(sessionId, session.revision, stagedFileOperation(session.format), {
          ...staged,
        })
        commands = store.poll(sessionId, context, viewId)
        pendingRecoveries.delete(sessionId)
      }
      if (commands.length === 0 && waitMs > 0) {
        commands = await store.waitForPoll(sessionId, context, waitMs, viewId)
      }
      return result({
        ok: true,
        commands,
      })
    } catch (error) {
      return failure(error)
    }
  },
)

server.registerTool(
  'office_editor_begin_document_save',
  {
    title: 'Begin local document save',
    description: 'Editor-only endpoint for beginning a bounded, session-owned local file save.',
    inputSchema: {
      sessionId: z.string().min(1),
      viewId: z.string().min(1),
      fileName: z.string().min(1).max(240),
      size: z.number().int().nonnegative().max(268_435_456),
      mode: z.enum(['save', 'save-as', 'export-copy']).default('save'),
    },
    _meta: { ui: { visibility: ['app'] } },
  },
  async ({ sessionId, viewId, fileName, size, mode }) => {
    try {
      const session = store.assertView(sessionId, viewId)
      const begun = await documentSaves.begin(sessionId, session.format, fileName, size, mode)
      return result({ ok: true, ...begun })
    } catch (error) {
      return failure(error)
    }
  },
)

server.registerTool(
  'office_editor_write_document_save_chunk',
  {
    title: 'Write local document save chunk',
    description: 'Editor-only endpoint for appending one ordered document save chunk.',
    inputSchema: {
      sessionId: z.string().min(1),
      viewId: z.string().min(1),
      uploadId: z.string().min(1),
      offset: z.number().int().nonnegative(),
      data: z.string().max(262_144),
    },
    _meta: { ui: { visibility: ['app'] } },
  },
  async ({ sessionId, viewId, uploadId, offset, data }) => {
    try {
      store.assertView(sessionId, viewId)
      const nextOffset = await documentSaves.write(sessionId, uploadId, offset, data)
      return result({ ok: true, nextOffset })
    } catch (error) {
      return failure(error)
    }
  },
)

server.registerTool(
  'office_editor_commit_document_save',
  {
    title: 'Commit local document save',
    description: 'Editor-only endpoint for atomically committing a complete local document save.',
    inputSchema: {
      sessionId: z.string().min(1),
      viewId: z.string().min(1),
      uploadId: z.string().min(1),
    },
    _meta: { ui: { visibility: ['app'] } },
  },
  async ({ sessionId, viewId, uploadId }) => {
    try {
      const session = store.assertView(sessionId, viewId)
      const committed = await documentSaves.commit(sessionId, uploadId)
      if (committed.bound) session.filePath = committed.path
      return result({ ok: true, uploadId, ...committed })
    } catch (error) {
      return failure(error)
    }
  },
)

server.registerTool(
  'office_editor_abort_document_save',
  {
    title: 'Abort local document save',
    description: 'Editor-only endpoint for discarding an incomplete local document save.',
    inputSchema: {
      sessionId: z.string().min(1),
      viewId: z.string().min(1),
      uploadId: z.string().min(1),
    },
    _meta: { ui: { visibility: ['app'] } },
  },
  async ({ sessionId, viewId, uploadId }) => {
    try {
      store.assertView(sessionId, viewId)
      await documentSaves.abort(sessionId, uploadId)
      return result({ ok: true, uploadId })
    } catch (error) {
      return failure(error)
    }
  },
)

server.registerTool(
  'office_editor_begin_recovery',
  {
    title: 'Begin editor recovery checkpoint',
    description: 'Editor-only endpoint for beginning a bounded local recovery upload.',
    inputSchema: {
      sessionId: z.string().min(1),
      viewId: z.string().min(1).optional(),
      fileName: z.string().min(1),
      size: z.number().int().nonnegative().max(268_435_456),
    },
    _meta: { ui: { visibility: ['app'] } },
  },
  async ({ sessionId, viewId, fileName, size }) => {
    try {
      const session = store.assertView(sessionId, viewId)
      const uploadId = recoveryUploads.begin(sessionId, session.format, fileName, size)
      return result({ ok: true, uploadId })
    } catch (error) {
      return failure(error)
    }
  },
)

server.registerTool(
  'office_editor_write_recovery_chunk',
  {
    title: 'Write editor recovery chunk',
    description: 'Editor-only endpoint for appending a recovery checkpoint chunk.',
    inputSchema: {
      sessionId: z.string().min(1),
      viewId: z.string().min(1).optional(),
      uploadId: z.string().min(1),
      offset: z.number().int().nonnegative(),
      data: z.string(),
    },
    _meta: { ui: { visibility: ['app'] } },
  },
  async ({ sessionId, viewId, uploadId, offset, data }) => {
    try {
      store.assertView(sessionId, viewId)
      return result({
        ok: true,
        nextOffset: recoveryUploads.write(sessionId, uploadId, offset, data),
      })
    } catch (error) {
      return failure(error)
    }
  },
)

server.registerTool(
  'office_editor_commit_recovery',
  {
    title: 'Commit editor recovery checkpoint',
    description: 'Editor-only endpoint for atomically committing a complete local recovery upload.',
    inputSchema: {
      sessionId: z.string().min(1),
      viewId: z.string().min(1).optional(),
      uploadId: z.string().min(1),
    },
    _meta: { ui: { visibility: ['app'] } },
  },
  async ({ sessionId, viewId, uploadId }) => {
    try {
      store.assertView(sessionId, viewId)
      await recoveryUploads.commit(sessionId, uploadId)
      return result({ ok: true, uploadId })
    } catch (error) {
      return failure(error)
    }
  },
)

server.registerTool(
  'office_editor_disconnect',
  {
    title: 'Disconnect editor session',
    description: 'Editor-only transport endpoint for marking a mounted session offline.',
    inputSchema: {
      sessionId: z.string().min(1),
      viewId: z.string().min(1).optional(),
    },
    _meta: { ui: { visibility: ['app'] } },
  },
  async ({ sessionId, viewId }) => {
    try {
      store.disconnect(sessionId, viewId)
      localFiles.releaseAssetRoots(sessionId)
      return result({ ok: true, sessionId })
    } catch (error) {
      return failure(error)
    }
  },
)

server.registerTool(
  'office_editor_read_local_asset_chunk',
  {
    title: 'Read local editor asset chunk',
    description: 'Editor-only transport endpoint for reading a session-bound local asset.',
    inputSchema: {
      sessionId: z.string().min(1),
      viewId: z.string().min(1).optional(),
      rootId: z.string().min(1),
      path: z.string().min(1).max(4096),
      offset: z.number().int().nonnegative(),
      length: z.number().int().positive().max(262_144),
    },
    _meta: { ui: { visibility: ['app'] } },
  },
  async ({ sessionId, viewId, rootId, path, offset, length }) => {
    try {
      store.assertView(sessionId, viewId)
      return result({
        ok: true,
        ...(await localFiles.readLocalAsset(sessionId, rootId, path, offset, length)),
      })
    } catch (error) {
      return failure(error)
    }
  },
)

server.registerTool(
  'office_editor_read_file_chunk',
  {
    title: 'Read staged local file chunk',
    description:
      'Editor-only transport endpoint for reading a staged local file without inline payloads.',
    inputSchema: {
      sessionId: z.string().min(1),
      viewId: z.string().min(1).optional(),
      blobId: z.string().min(1),
      offset: z.number().int().nonnegative(),
      length: z.number().int().positive().max(262_144),
    },
    _meta: { ui: { visibility: ['app'] } },
  },
  async ({ sessionId, viewId, blobId, offset, length }) => {
    try {
      store.assertView(sessionId, viewId)
      return result({ ok: true, ...localFiles.read(sessionId, blobId, offset, length) })
    } catch (error) {
      return failure(error)
    }
  },
)

server.registerTool(
  'office_editor_read_font_chunk',
  {
    title: 'Read bundled font asset chunk',
    description: 'Editor-only endpoint for lazily reading an allowlisted bundled font.',
    inputSchema: {
      fileName: z.string().min(1),
      offset: z.number().int().nonnegative(),
      length: z.number().int().positive().max(262_144),
    },
    _meta: { ui: { visibility: ['app'] } },
  },
  async ({ fileName, offset, length }) => {
    try {
      return result({ ok: true, ...(await fontAssets.read(fileName, offset, length)) })
    } catch (error) {
      return failure(error)
    }
  },
)

server.registerTool(
  'office_editor_acknowledge',
  {
    title: 'Acknowledge editor command',
    description: 'Editor-only transport endpoint for advancing the shared revision.',
    inputSchema: {
      sessionId: z.string().min(1),
      viewId: z.string().min(1).optional(),
      commandId: z.string().min(1),
      revision: z.number().int().nonnegative().optional(),
      ok: z.boolean().default(true),
      error: z.enum(['unsupported_operation', 'invalid_arguments', 'execution_failed']).optional(),
      message: z.string().optional(),
      output: z.record(z.string(), z.unknown()).optional(),
      timing: z
        .object({
          hydrateMs: z.number().nonnegative().max(3_600_000),
          executeMs: z.number().nonnegative().max(3_600_000),
          trace: z
            .object({
              operation: z.literal('markdown.document.load_staged'),
              phases: z
                .object({
                  decodeMs: z.number().nonnegative().max(3_600_000),
                  parseMs: z.number().nonnegative().max(3_600_000),
                  tiptapStateInstallMs: z.number().nonnegative().max(3_600_000),
                  reactCommitMs: z.number().nonnegative().max(3_600_000),
                })
                .strict(),
            })
            .strict()
            .optional(),
        })
        .strict()
        .optional(),
      fileName: z.string().nullable().optional(),
      dirty: z.boolean().optional(),
      selection: z.record(z.string(), z.unknown()).nullable().optional(),
    },
    _meta: { ui: { visibility: ['app'] } },
  },
  async ({
    sessionId,
    viewId,
    commandId,
    revision,
    ok,
    error,
    message,
    output,
    timing,
    fileName,
    dirty,
    selection,
  }) => {
    try {
      store.assertView(sessionId, viewId)
      if (!ok) {
        const session = store.reject(
          sessionId,
          commandId,
          error ?? 'execution_failed',
          message ?? 'The editor could not apply the command.',
        )
        return result({ ok: true, session })
      }
      if (revision === undefined) {
        throw new SessionError(
          'invalid_arguments',
          'A successful editor acknowledgement requires a revision.',
        )
      }
      return result({
        ok: true,
        ...(timing ? { timing } : {}),
        session: store.acknowledge(
          sessionId,
          commandId,
          revision,
          {
            ...(fileName !== undefined ? { fileName } : {}),
            ...(dirty !== undefined ? { dirty } : {}),
            ...(selection !== undefined ? { selection } : {}),
          },
          output,
        ),
      })
    } catch (error) {
      return failure(error)
    }
  },
)

const transport = new StdioServerTransport()
await server.connect(transport)
