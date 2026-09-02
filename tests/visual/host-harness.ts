import { AppBridge, PostMessageTransport } from '@modelcontextprotocol/ext-apps/app-bridge'

type DisplayMode = 'inline' | 'fullscreen'

interface MarkdownStagedLoadTrace {
  operation: 'markdown.document.load_staged'
  phases: {
    decodeMs: number
    parseMs: number
    tiptapStateInstallMs: number
    reactCommitMs: number
  }
}

interface XlsxColdStartTrace {
  operation: 'xlsx.editor.cold_start'
  phases: {
    bootstrapMs: number
    univerCreateMs: number
    worksheetInstallMs: number
    firstCommitMs: number
  }
  bootstrapPhases: {
    resourceReceiveMs: number
    moduleGraphReadyMs: number
    reactMountMs: number
  }
}

interface VisualHostState {
  acknowledgementSnapshots: Record<
    string,
    {
      editorText: string
      fileName: string | null
      ready: boolean
    }
  >
  acknowledgements: Array<Record<string, unknown>>
  commands: Array<Record<string, unknown>>
  displayMode: DisplayMode
  downloads: Array<{ fileName: string; mimeType: string; size: number }>
  editorLoads: number
  errors: string[]
  events: string[]
  initialized: boolean
  initializedAt: number | null
  firstPollAt: number | null
  firstPollArguments: (Record<string, unknown> & { startupTrace?: XlsxColdStartTrace }) | null
  firstPollSnapshot: {
    appShellVisible: boolean
    canvasCount: number
    activeSheet: string | null
  } | null
  lastPollArguments: Record<string, unknown> | null
  polls: number
  recoveryCommits: number
  sizeNotifications: Array<{ width?: number; height?: number }>
  stagedFiles: Record<string, string>
  startedAt: number
  commandTimings: Record<
    string,
    {
      enqueuedAt: number
      polledAt?: number
      acknowledgedAt?: number
      pollWaitMs?: number
      hydrateMs?: number
      executeMs?: number
      ackTransportMs?: number
      totalMs?: number
      trace?: MarkdownStagedLoadTrace
    }
  >
  enqueueCommand: (
    command: Record<string, unknown>,
    stagedFile?: { blobId: string; base64: string },
  ) => void
}

declare global {
  interface Window {
    __codexVisualHost: VisualHostState
  }
}

const params = new URLSearchParams(window.location.search)
const format = params.get('format') ?? 'docx'
const width = Number(params.get('width') ?? 420)
const height = Number(params.get('height') ?? 900)
const acceptFullscreen = params.get('acceptFullscreen') === 'true'
const fullscreenWidth = Number(params.get('fullscreenWidth') ?? 1332)
const fullscreenHeight = Number(params.get('fullscreenHeight') ?? 1280)
const iframe = document.querySelector<HTMLIFrameElement>('#editor-frame')!
iframe.style.width = `${width}px`
iframe.style.height = `${height}px`

const state: VisualHostState = {
  acknowledgementSnapshots: {},
  acknowledgements: [],
  commands: [],
  displayMode: 'inline',
  downloads: [],
  editorLoads: 0,
  errors: [],
  events: [],
  initialized: false,
  initializedAt: null,
  firstPollAt: null,
  firstPollArguments: null,
  firstPollSnapshot: null,
  lastPollArguments: null,
  polls: 0,
  recoveryCommits: 0,
  sizeNotifications: [],
  stagedFiles: {},
  startedAt: performance.now(),
  commandTimings: {},
  enqueueCommand(command, stagedFile) {
    const commandId = String(command.commandId ?? '')
    if (!commandId) throw new Error('A benchmark command requires commandId.')
    if (stagedFile) state.stagedFiles[stagedFile.blobId] = stagedFile.base64
    state.commandTimings[commandId] = { enqueuedAt: performance.now() }
    state.commands.push(command)
  },
}
window.__codexVisualHost = state
window.addEventListener('error', (event) => state.errors.push(event.message))
window.addEventListener('unhandledrejection', (event) => state.errors.push(String(event.reason)))

let hostContext = {
  theme: 'light' as const,
  displayMode: 'inline' as DisplayMode,
  availableDisplayModes: ['inline', 'fullscreen'] as const,
  containerDimensions: { width, height },
  locale: 'zh-CN',
  platform: 'desktop' as const,
}
const bridge = new AppBridge(
  null,
  { name: 'Codex visual host', version: '1.0.0' },
  {},
  { hostContext },
)

function toolResult(structuredContent: Record<string, unknown>) {
  return {
    content: [{ type: 'text' as const, text: JSON.stringify(structuredContent) }],
    structuredContent,
  }
}

type PendingPoll = {
  resolve: (result: ReturnType<typeof toolResult>) => void
  timer: number
}

let pendingPoll: PendingPoll | null = null
let nextRecoveryUpload = 0
const recoveryOffsets = new Map<string, number>()
let nextDocumentSaveUpload = 0
const documentSaveUploads = new Map<
  string,
  { fileName: string; bytes: Uint8Array; offset: number }
>()

const formatMimeTypes: Record<string, string> = {
  docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  markdown: 'text/markdown;charset=utf-8',
  xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  pptx: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  pdf: 'application/pdf',
}

function drainCommands(now = performance.now()): Array<Record<string, unknown>> {
  const commands = state.commands.splice(0)
  for (const command of commands) {
    const timing = state.commandTimings[String(command.commandId ?? '')]
    if (timing) timing.polledAt = now
  }
  return commands
}

function finishPendingPoll(commands: Array<Record<string, unknown>>): void {
  if (!pendingPoll) return
  const { resolve, timer } = pendingPoll
  pendingPoll = null
  window.clearTimeout(timer)
  resolve(toolResult({ ok: true, commands }))
}

const pushCommands = state.commands.push.bind(state.commands)
state.commands.push = (...commands) => {
  const length = pushCommands(...commands)
  if (pendingPoll) finishPendingPoll(drainCommands())
  return length
}

bridge.oncalltool = async ({ name, arguments: arguments_ }) => {
  state.events.push(`tool:${name}`)
  if (name === 'office_editor_begin_document_save') {
    const uploadId = `visual-document-save-${++nextDocumentSaveUpload}`
    const fileName = String(arguments_?.fileName ?? '')
    const size = Number(arguments_?.size ?? -1)
    if (!fileName || !Number.isInteger(size) || size < 0) {
      return toolResult({ ok: false, error: 'invalid_document_save' })
    }
    documentSaveUploads.set(uploadId, { fileName, bytes: new Uint8Array(size), offset: 0 })
    return toolResult({ ok: true, uploadId, path: `/visual-output/${fileName}` })
  }
  if (name === 'office_editor_write_document_save_chunk') {
    const uploadId = String(arguments_?.uploadId ?? '')
    const upload = documentSaveUploads.get(uploadId)
    const offset = Number(arguments_?.offset ?? -1)
    if (!upload || offset !== upload.offset) {
      return toolResult({ ok: false, error: 'invalid_document_save_offset' })
    }
    const source = String(arguments_?.data ?? '')
    const chunk = Uint8Array.from(atob(source), (character) => character.charCodeAt(0))
    if (offset + chunk.length > upload.bytes.length) {
      return toolResult({ ok: false, error: 'document_save_overflow' })
    }
    upload.bytes.set(chunk, offset)
    upload.offset += chunk.length
    return toolResult({ ok: true, nextOffset: upload.offset })
  }
  if (name === 'office_editor_commit_document_save') {
    const uploadId = String(arguments_?.uploadId ?? '')
    const upload = documentSaveUploads.get(uploadId)
    if (!upload || upload.offset !== upload.bytes.length) {
      return toolResult({ ok: false, error: 'incomplete_document_save' })
    }
    documentSaveUploads.delete(uploadId)
    const mimeType = formatMimeTypes[format] ?? 'application/octet-stream'
    if (format === 'xlsx' && iframe.contentWindow) {
      ;(
        iframe.contentWindow as Window & { __savedWorkbookBytes?: ArrayBuffer }
      ).__savedWorkbookBytes = upload.bytes.buffer.slice(0)
    }
    state.downloads.push({ fileName: upload.fileName, mimeType, size: upload.bytes.length })
    const url = URL.createObjectURL(new Blob([upload.bytes], { type: mimeType }))
    const anchor = document.createElement('a')
    anchor.href = url
    anchor.download = upload.fileName
    anchor.click()
    setTimeout(() => URL.revokeObjectURL(url), 0)
    return toolResult({
      ok: true,
      uploadId,
      path: `/visual-output/${upload.fileName}`,
      bound: true,
    })
  }
  if (name === 'office_editor_abort_document_save') {
    documentSaveUploads.delete(String(arguments_?.uploadId ?? ''))
    return toolResult({ ok: true })
  }
  if (name === 'office_editor_begin_recovery') {
    const uploadId = `visual-recovery-${++nextRecoveryUpload}`
    recoveryOffsets.set(uploadId, 0)
    return toolResult({ ok: true, uploadId })
  }
  if (name === 'office_editor_write_recovery_chunk') {
    const uploadId = String(arguments_?.uploadId ?? '')
    const offset = Number(arguments_?.offset ?? 0)
    const source = String(arguments_?.data ?? '')
    if (!recoveryOffsets.has(uploadId) || recoveryOffsets.get(uploadId) !== offset) {
      return toolResult({ ok: false, error: 'invalid_recovery_offset' })
    }
    const nextOffset = offset + atob(source).length
    recoveryOffsets.set(uploadId, nextOffset)
    return toolResult({ ok: true, nextOffset })
  }
  if (name === 'office_editor_commit_recovery') {
    const uploadId = String(arguments_?.uploadId ?? '')
    if (!recoveryOffsets.delete(uploadId)) {
      return toolResult({ ok: false, error: 'unknown_recovery_upload' })
    }
    state.recoveryCommits += 1
    return toolResult({ ok: true })
  }
  if (name === 'office_editor_poll') {
    const now = performance.now()
    if (state.firstPollAt === null) {
      const editorDocument = iframe.contentDocument
      const selection = arguments_?.selection as { sheet?: unknown } | null | undefined
      state.firstPollAt = now
      state.firstPollArguments = arguments_ ?? null
      state.firstPollSnapshot = {
        appShellVisible: editorDocument?.querySelector<HTMLElement>('.app-shell')?.hidden === false,
        canvasCount: editorDocument?.querySelectorAll('#univer-container canvas').length ?? 0,
        activeSheet: typeof selection?.sheet === 'string' ? selection.sheet : null,
      }
    }
    state.polls += 1
    state.lastPollArguments = arguments_ ?? null
    const commands = drainCommands(now)
    const waitMs = Number(arguments_?.waitMs ?? 0)
    if (commands.length > 0 || waitMs <= 0) return toolResult({ ok: true, commands })

    finishPendingPoll([])
    return new Promise((resolve) => {
      const timer = window.setTimeout(() => finishPendingPoll([]), waitMs)
      pendingPoll = { resolve, timer }
    })
  }
  if (name === 'office_editor_acknowledge') {
    state.acknowledgements.push(arguments_ ?? {})
    const commandId = String(arguments_?.commandId ?? '')
    const editorDocument = iframe.contentDocument
    state.acknowledgementSnapshots[commandId] = {
      editorText: editorDocument?.querySelector('.doc-editor')?.textContent ?? '',
      fileName: editorDocument?.querySelector('.status-file')?.textContent ?? null,
      ready: editorDocument?.querySelector<HTMLElement>('.app-main')?.hidden === false,
    }
    const timing = state.commandTimings[commandId]
    if (timing) {
      const acknowledgedAt = performance.now()
      const rendererTiming = arguments_?.timing as
        { hydrateMs?: unknown; executeMs?: unknown; trace?: unknown } | undefined
      const hydrateMs = typeof rendererTiming?.hydrateMs === 'number' ? rendererTiming.hydrateMs : 0
      const executeMs = typeof rendererTiming?.executeMs === 'number' ? rendererTiming.executeMs : 0
      const pollWaitMs = (timing.polledAt ?? acknowledgedAt) - timing.enqueuedAt
      const totalMs = acknowledgedAt - timing.enqueuedAt
      Object.assign(timing, {
        acknowledgedAt,
        pollWaitMs,
        hydrateMs,
        executeMs,
        ackTransportMs: totalMs - pollWaitMs - hydrateMs - executeMs,
        totalMs,
        ...(rendererTiming?.trace
          ? { trace: rendererTiming.trace as MarkdownStagedLoadTrace }
          : {}),
      })
    }
    return toolResult({ ok: true })
  }
  if (name === 'office_editor_read_font_chunk') {
    const query = new URLSearchParams({
      name: String(arguments_?.fileName ?? ''),
      offset: String(arguments_?.offset ?? 0),
      length: String(arguments_?.length ?? 262_144),
    })
    const response = await fetch(`/font?${query}`)
    return toolResult((await response.json()) as Record<string, unknown>)
  }
  if (name === 'office_editor_read_file_chunk') {
    const blobId = String(arguments_?.blobId ?? '')
    const offset = Number(arguments_?.offset ?? 0)
    const length = Number(arguments_?.length ?? 262_144)
    const source = state.stagedFiles[blobId]
    if (!source) return toolResult({ ok: false, error: 'file_not_found' })
    const bytes = Uint8Array.from(atob(source), (character) => character.charCodeAt(0))
    const chunk = bytes.subarray(offset, Math.min(offset + length, bytes.length))
    let binary = ''
    for (const byte of chunk) binary += String.fromCharCode(byte)
    return toolResult({ ok: true, data: btoa(binary), nextOffset: offset + chunk.length })
  }
  return toolResult({ ok: true })
}

async function applyDisplayMode(mode: DisplayMode): Promise<void> {
  const dimensions =
    mode === 'fullscreen' ? { width: fullscreenWidth, height: fullscreenHeight } : { width, height }
  state.displayMode = mode
  iframe.style.width = `${dimensions.width}px`
  iframe.style.height = `${dimensions.height}px`
  hostContext = { ...hostContext, displayMode: mode, containerDimensions: dimensions }
  bridge.setHostContext(hostContext)
  await new Promise<void>((resolve) => requestAnimationFrame(() => resolve()))
  await new Promise<void>((resolve) => requestAnimationFrame(() => resolve()))
}

bridge.onrequestdisplaymode = async ({ mode }) => {
  state.events.push(`display-request:${mode}`)
  if (acceptFullscreen && (mode === 'inline' || mode === 'fullscreen')) {
    await applyDisplayMode(mode)
  }
  return { mode: state.displayMode }
}
bridge.onsizechange = (size) => {
  state.sizeNotifications.push(size)
  state.events.push(`size:${size.width ?? 'auto'}x${size.height ?? 'auto'}`)
}
bridge.ondownloadfile = async () => {
  throw new Error('The editor unexpectedly used the retired host download boundary.')
}
bridge.oninitialized = () => {
  installVisualWorkbookSavePicker()
  state.initialized = true
  state.initializedAt = performance.now()
  state.events.push('initialized')
  void (async () => {
    await bridge.sendToolInput({ arguments: { sessionId: 'visual-session' } })
    await bridge.sendToolResult(
      toolResult({
        ok: true,
        sessionId: 'visual-session',
        viewId: 'visual-view',
        format,
        revision: 0,
      }),
    )
  })()
}

function installVisualWorkbookSavePicker(): void {
  if (format !== 'xlsx' || !iframe.contentWindow || !iframe.contentDocument) return
  const editorWindow = iframe.contentWindow
  const hostWindow = window as Window & {
    __codexVisualWorkbookSavePicker?: (
      options?: Record<string, unknown>,
    ) => Promise<FileSystemFileHandle>
  }
  hostWindow.__codexVisualWorkbookSavePicker = async (options = {}) => {
    const name = typeof options.suggestedName === 'string' ? options.suggestedName : 'workbook.xlsx'
    let persisted = new ArrayBuffer(0)
    return {
      kind: 'file',
      name,
      createWritable: async () =>
        ({
          write: async (data: FileSystemWriteChunkType) => {
            if (Object.prototype.toString.call(data) === '[object ArrayBuffer]') {
              persisted = (data as ArrayBuffer).slice(0)
            } else if (ArrayBuffer.isView(data)) {
              persisted = data.buffer.slice(
                data.byteOffset,
                data.byteOffset + data.byteLength,
              ) as ArrayBuffer
            } else if (data instanceof Blob) persisted = await data.arrayBuffer()
            else throw new Error('Unexpected XLSX write payload.')
          },
          close: async () => {
            const url = editorWindow.URL.createObjectURL(
              new Blob([persisted], {
                type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
              }),
            )
            const anchor = editorWindow.document.createElement('a')
            anchor.href = url
            anchor.download = name
            anchor.click()
            editorWindow.setTimeout(() => editorWindow.URL.revokeObjectURL(url), 0)
          },
        }) as FileSystemWritableFileStream,
      getFile: async () =>
        new editorWindow.File([persisted], name, {
          type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        }),
    } as FileSystemFileHandle
  }
  const installer = iframe.contentDocument.createElement('script')
  installer.textContent =
    'window.showSaveFilePicker = (...args) => parent.__codexVisualWorkbookSavePicker(...args)'
  iframe.contentDocument.documentElement.append(installer)
  installer.remove()
}

async function mountEditor(): Promise<void> {
  const editorHtml = await fetch(`/editor.html?format=${encodeURIComponent(format)}`).then(
    (response) => response.text(),
  )
  await bridge.connect(new PostMessageTransport(iframe.contentWindow!, iframe.contentWindow!))
  iframe.addEventListener('load', () => {
    state.editorLoads += 1
  })
  iframe.srcdoc = editorHtml
}

void mountEditor().catch((error) =>
  state.errors.push(error instanceof Error ? error.message : String(error)),
)

export {}
