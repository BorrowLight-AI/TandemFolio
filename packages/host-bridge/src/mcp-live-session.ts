import { App as McpApp } from '@modelcontextprotocol/ext-apps'
import {
  createMcpDisplayModeController,
  type McpDisplayMode,
  type McpDisplayModeState,
} from './display-mode'
import { createWakeablePollLoop } from './wakeable-poll-loop'
import { createLiveSessionStatus } from './live-session-status'

export interface LiveEditorSnapshot {
  revision: number
  fileName: string | null
  dirty: boolean
  selection: Record<string, unknown> | null
}

export interface LiveEditorDisplayModeState extends McpDisplayModeState {
  readonly embedded: boolean
}

export interface LiveEditorCommand {
  commandId: string
  baseRevision: number
  operation: string
  arguments: Record<string, unknown>
}

export interface LiveEditorExecutionTrace {
  readonly operation: 'markdown.document.load_staged'
  readonly phases: {
    readonly decodeMs: number
    readonly parseMs: number
    readonly tiptapStateInstallMs: number
    readonly reactCommitMs: number
  }
}

export interface LiveEditorStartupTrace {
  readonly operation: 'xlsx.editor.cold_start'
  readonly phases: {
    readonly bootstrapMs: number
    readonly univerCreateMs: number
    readonly worksheetInstallMs: number
    readonly firstCommitMs: number
  }
  readonly bootstrapPhases: {
    readonly resourceReceiveMs: number
    readonly moduleGraphReadyMs: number
    readonly reactMountMs: number
  }
}

export type LiveEditorExecution =
  | {
      ok: true
      output?: Record<string, unknown>
      recovery?: { fileName: string; data: ArrayBuffer }
      trace?: LiveEditorExecutionTrace
    }
  | {
      ok: false
      error: 'unsupported_operation' | 'invalid_arguments' | 'execution_failed'
      message: string
    }

export interface LiveEditorAdapter {
  execute(command: LiveEditorCommand): Promise<LiveEditorExecution>
  snapshot(revision: number): LiveEditorSnapshot
  recoverySnapshot?(force?: boolean): Promise<{ fileName: string; data: ArrayBuffer } | null>
  recoveryVersion?(): string | number
  startupTrace?(): Promise<LiveEditorStartupTrace>
}

const displayController = createMcpDisplayModeController()
let displayState: LiveEditorDisplayModeState = {
  embedded: typeof window !== 'undefined' && window.parent !== window,
  ...displayController.getState(),
}
const displayListeners = new Set<() => void>()
const activityListeners = new Set<() => void>()
let liveEditorActivity = true
const MAX_LOCAL_ASSET_BYTES = 20 * 1024 * 1024
const MCP_CONNECT_TIMEOUT_MS = 1_000
const MCP_CONNECT_RETRY_MS = 250

export interface LiveEditorLocalAsset {
  readonly mime: 'image/png' | 'image/jpeg' | 'image/gif'
  readonly data: ArrayBuffer
}

export interface LiveEditorFileSave {
  readonly fileName: string
  readonly data: ArrayBuffer
  readonly mode?: 'save' | 'save-as' | 'export-copy'
}

export type LiveEditorFileSaveResult =
  { readonly ok: true; readonly path: string } | { readonly ok: false; readonly error: string }

let localAssetReader:
  ((rootId: string, path: string) => Promise<LiveEditorLocalAsset | null>) | null = null
let bundledFontReader: ((fileName: string) => Promise<ArrayBuffer>) | null = null
let liveFileSaver: ((file: LiveEditorFileSave) => Promise<LiveEditorFileSaveResult>) | null = null
let documentReplacer: (<T>(replace: () => Promise<T>) => Promise<T>) | null = null

export async function replaceLiveEditorDocument<T>(replace: () => Promise<T>): Promise<T> {
  if (window.parent === window) return replace()
  if (!documentReplacer) throw new Error('The document session is not connected.')
  return documentReplacer(replace)
}

export async function saveLiveEditorFile(
  file: LiveEditorFileSave,
): Promise<LiveEditorFileSaveResult> {
  return liveFileSaver
    ? liveFileSaver(file)
    : { ok: false, error: 'The local document save bridge is not connected.' }
}

export async function readLiveEditorLocalAsset(
  rootId: string,
  path: string,
): Promise<LiveEditorLocalAsset | null> {
  if (!localAssetReader) throw new Error('The live-editor local asset bridge is not connected.')
  return localAssetReader(rootId, path)
}

export async function readLiveEditorBundledFontAsset(fileName: string): Promise<ArrayBuffer> {
  if (!bundledFontReader) throw new Error('The live-editor font bridge is not connected.')
  return bundledFontReader(fileName)
}

displayController.subscribe(() => {
  displayState = {
    embedded: typeof window !== 'undefined' && window.parent !== window,
    ...displayController.getState(),
  }
  for (const listener of displayListeners) listener()
})

export function getLiveEditorDisplayMode(): LiveEditorDisplayModeState {
  const embedded = window.parent !== window
  if (displayState.embedded !== embedded) {
    displayState = { ...displayState, embedded }
  }
  return displayState
}

export function getLiveEditorActivity(): boolean {
  return liveEditorActivity
}

export function subscribeLiveEditorActivity(listener: () => void): () => void {
  activityListeners.add(listener)
  return () => activityListeners.delete(listener)
}

function publishLiveEditorActivity(active: boolean): void {
  if (liveEditorActivity === active) return
  liveEditorActivity = active
  for (const listener of activityListeners) listener()
}

export function subscribeLiveEditorDisplayMode(listener: () => void): () => void {
  displayListeners.add(listener)
  return () => displayListeners.delete(listener)
}

export async function toggleLiveEditorFullscreen(): Promise<void> {
  await displayController.toggle()
}

function displayContext(app: McpApp): {
  availableDisplayModes?: readonly McpDisplayMode[]
  displayMode?: McpDisplayMode
} {
  const context = app.getHostContext()
  return {
    ...(context?.availableDisplayModes
      ? { availableDisplayModes: context.availableDisplayModes }
      : {}),
    ...(context?.displayMode ? { displayMode: context.displayMode } : {}),
  }
}

function decodeBase64(data: string): Uint8Array {
  const binary = atob(data)
  const bytes = new Uint8Array(binary.length)
  for (let index = 0; index < binary.length; index += 1) bytes[index] = binary.charCodeAt(index)
  return bytes
}

function encodeBase64(bytes: Uint8Array): string {
  let binary = ''
  for (let offset = 0; offset < bytes.length; offset += 32_768) {
    binary += String.fromCharCode(...bytes.subarray(offset, offset + 32_768))
  }
  return btoa(binary)
}

async function readStagedFile(
  app: McpApp,
  sessionId: string,
  viewId: string,
  descriptor: Record<string, unknown>,
): Promise<ArrayBuffer> {
  const blobId = descriptor.blobId
  const size = descriptor.size
  if (typeof blobId !== 'string' || !Number.isInteger(size) || (size as number) < 0) {
    throw new Error('Invalid staged file descriptor.')
  }
  const output = new Uint8Array(size as number)
  let offset = 0
  while (offset < output.length) {
    const response = await app.callServerTool({
      name: 'office_editor_read_file_chunk',
      arguments: {
        sessionId,
        viewId,
        blobId,
        offset,
        length: Math.min(262_144, output.length - offset),
      },
    })
    const payload = response.structuredContent as
      { ok?: boolean; data?: string; nextOffset?: number } | undefined
    if (
      !payload?.ok ||
      typeof payload.data !== 'string' ||
      typeof payload.nextOffset !== 'number'
    ) {
      throw new Error('The staged file chunk could not be read.')
    }
    const chunk = decodeBase64(payload.data)
    if (payload.nextOffset !== offset + chunk.length || chunk.length === 0) {
      throw new Error('The staged file chunk returned invalid offsets.')
    }
    output.set(chunk, offset)
    offset = payload.nextOffset
  }
  return output.buffer
}

function isStagedFileOperation(operation: string): boolean {
  return (
    operation === 'open_local_file' ||
    operation.endsWith('.document.load_staged') ||
    operation === 'docx.document.compare_staged' ||
    operation === 'pdf.page.insert_staged' ||
    operation === 'xlsx.image.add_staged' ||
    (operation.startsWith('docx.image.') && operation.endsWith('_staged'))
  )
}

async function storeRecovery(
  app: McpApp,
  sessionId: string,
  viewId: string,
  recovery: { fileName: string; data: ArrayBuffer },
): Promise<void> {
  const begun = await app.callServerTool({
    name: 'office_editor_begin_recovery',
    arguments: { sessionId, viewId, fileName: recovery.fileName, size: recovery.data.byteLength },
  })
  const uploadId = (begun.structuredContent as { uploadId?: unknown } | undefined)?.uploadId
  if (typeof uploadId !== 'string') throw new Error('The recovery upload could not begin.')
  const bytes = new Uint8Array(recovery.data)
  for (let offset = 0; offset < bytes.length; offset += 196_608) {
    const chunk = bytes.subarray(offset, Math.min(offset + 196_608, bytes.length))
    const response = await app.callServerTool({
      name: 'office_editor_write_recovery_chunk',
      arguments: { sessionId, viewId, uploadId, offset, data: encodeBase64(chunk) },
    })
    const nextOffset = (response.structuredContent as { nextOffset?: unknown } | undefined)
      ?.nextOffset
    if (nextOffset !== offset + chunk.length) throw new Error('Invalid recovery offset.')
  }
  const committed = await app.callServerTool({
    name: 'office_editor_commit_recovery',
    arguments: { sessionId, viewId, uploadId },
  })
  if (!(committed.structuredContent as { ok?: unknown } | undefined)?.ok) {
    throw new Error('The recovery upload could not be committed.')
  }
}

export function attachMcpLiveSession(adapter: LiveEditorAdapter): () => void {
  if (window.parent === window) return () => undefined
  const mountId = crypto.randomUUID()
  const createApp = (): McpApp => {
    const target = new McpApp(
      { name: 'tandemfolio-editor', version: '0.1.0' },
      { availableDisplayModes: ['inline', 'fullscreen'] },
    )
    const call = target.callServerTool.bind(target)
    target.callServerTool = (request, options) =>
      call(
        {
          ...request,
          arguments: { ...request.arguments, ...(request.arguments?.sessionId ? { mountId } : {}) },
        },
        options,
      )
    return target
  }
  let app = createApp()
  let stopped = false
  let terminalError = ''
  let editingReady = false
  let suspended = false
  let retryableConflict = false
  let resumeOnVisible = false
  let retryHandoff = false
  let activateView = false
  // A user-selected handoff outlives stale visibility hints, but only until restore/timeout.
  let continueHerePending = false
  let handoffInProgress = false
  let uncertainHandoff: string | null = null
  let resumeTimer = 0
  let resumeDelay = 1_000
  let waitTimer = 0
  let waitExpired = false
  let sessionId = new URLSearchParams(window.location.search).get('sessionId') ?? ''
  let viewId = new URLSearchParams(window.location.search).get('viewId') ?? ''
  let brokerSessionId = ''
  let sessionFormat: string | undefined
  let coldStart = true
  let bindingVersion = 0
  let boundPath: string | null = null
  let pathKnown = false
  let connectRetryTimer = 0
  let recoveryTimer = 0
  let recoveryTask: Promise<void> | null = null
  let documentChanging = false
  let replacementBarrier: Promise<void> | null = null
  let executingCommandId: string | undefined
  let saveTask: Promise<LiveEditorFileSaveResult> | null = null
  const storedRecoveryVersions = new Map<string, string | number>()
  let appConnected = false
  let displayStarted = false
  let startupTrace: LiveEditorStartupTrace | undefined
  let startupTracePending = false
  let editorIntersectionKnown = false
  let editorIntersecting = true
  let editorContentSkipped = false
  const retryConnection = (activate = false): void => {
    if (stopped) return
    if (uncertainHandoff) {
      void reconcileUncertainHandoff()
      return
    }
    // A retry can already have an app-only poll in flight while the continuation
    // action remains visible. Preserve the later, explicit user intent so the
    // next conflict retry upgrades the existing handoff instead of dropping it.
    if (!terminalError) {
      if (!activate) return
      retryHandoff = true
      activateView = true
      continueHerePending = true
      showContinuationPending()
      return
    }
    clearWait()
    resumeDelay = 1_000
    retryHandoff = true
    activateView = activate
    continueHerePending = activate
    resumeConnection()
    if (activate) showContinuationPending()
  }
  const status = createLiveSessionStatus(
    () => retryConnection(),
    () => retryConnection(true),
  )
  const showContinuationPending = (): void => {
    status.activation(true)
    status.actionsPending(true)
    status.block('已请求在此继续编辑，正在等待原编辑器安全交接…')
    status.error('已请求在此继续编辑，正在等待原编辑器安全交接…', true)
  }
  status.block('正在连接并恢复原文件…')
  function clearWait(): void {
    window.clearTimeout(waitTimer)
    waitTimer = 0
    waitExpired = false
    continueHerePending = false
    status.actionsPending(false)
  }
  function beginWait(): void {
    if (waitTimer || waitExpired) return
    waitTimer = window.setTimeout(() => {
      waitTimer = 0
      waitExpired = true
      continueHerePending = false
      status.actionsPending(false)
      resumeOnVisible = false
      window.clearTimeout(resumeTimer)
      resumeTimer = 0
      pollLoop.stop()
      status.block('等待已超时，原文件尚未恢复。')
      status.error(
        '等待已超时（30 秒），已停止自动重试。可点击“在此继续编辑”请求原编辑器安全交接；未确认前不会开放编辑。',
        true,
      )
    }, 30_000)
  }
  function resumeConnection(): void {
    window.clearTimeout(resumeTimer)
    resumeTimer = 0
    if (
      stopped ||
      handoffInProgress ||
      waitExpired ||
      (!continueHerePending && !isEditorWorkActive())
    )
      return
    terminalError = ''
    suspended = false
    status.error('正在重新连接…')
    pollLoop.stop()
    pollLoop = createWakeablePollLoop(poll)
    pollLoop.start()
  }
  function scheduleResume(): void {
    if (resumeTimer || stopped || waitExpired || (!continueHerePending && !isEditorWorkActive()))
      return
    resumeTimer = window.setTimeout(resumeConnection, resumeDelay)
    resumeDelay = Math.min(resumeDelay * 2, 5_000)
  }
  const requestLocalSave = async (file: LiveEditorFileSave): Promise<LiveEditorFileSaveResult> => {
    let uploadId = ''
    try {
      if (terminalError || !editingReady)
        return { ok: false, error: terminalError || 'The document is still restoring.' }
      if (!sessionId || !viewId) {
        return { ok: false, error: 'The editor is not bound to a local save session.' }
      }
      const begun = await app.callServerTool({
        name: 'office_editor_begin_document_save',
        arguments: {
          sessionId,
          viewId,
          fileName: file.fileName,
          size: file.data.byteLength,
          mode: file.mode ?? 'save',
        },
      })
      const begunPayload = begun.structuredContent as
        { ok?: unknown; uploadId?: unknown; message?: unknown } | undefined
      if (begunPayload?.ok !== true || typeof begunPayload.uploadId !== 'string') {
        throw new Error(
          typeof begunPayload?.message === 'string'
            ? begunPayload.message
            : 'The local document save upload could not begin.',
        )
      }
      uploadId = begunPayload.uploadId
      const bytes = new Uint8Array(file.data)
      for (let offset = 0; offset < bytes.length; offset += 196_608) {
        const chunk = bytes.subarray(offset, Math.min(offset + 196_608, bytes.length))
        const response = await app.callServerTool({
          name: 'office_editor_write_document_save_chunk',
          arguments: {
            sessionId,
            viewId,
            uploadId,
            offset,
            data: encodeBase64(chunk),
          },
        })
        const payload = response.structuredContent as
          { ok?: unknown; nextOffset?: unknown; message?: unknown } | undefined
        if (payload?.ok !== true || payload.nextOffset !== offset + chunk.length) {
          throw new Error(
            typeof payload?.message === 'string'
              ? payload.message
              : 'The local document save upload returned an invalid offset.',
          )
        }
      }
      const committed = await app.callServerTool({
        name: 'office_editor_commit_document_save',
        arguments: { sessionId, viewId, uploadId },
      })
      const committedPayload = committed.structuredContent as
        { ok?: unknown; path?: unknown; message?: unknown } | undefined
      if (committedPayload?.ok !== true || typeof committedPayload.path !== 'string') {
        throw new Error(
          typeof committedPayload?.message === 'string'
            ? committedPayload.message
            : 'The local document save upload could not be committed.',
        )
      }
      uploadId = ''
      bindingVersion += 1
      if (file.mode !== 'export-copy') boundPath = committedPayload.path
      pathKnown = true
      status.path(committedPayload.path, file.mode === 'export-copy')
      status.error(null)
      return { ok: true, path: committedPayload.path }
    } catch (error) {
      if (uploadId && sessionId && viewId) {
        await app
          .callServerTool({
            name: 'office_editor_abort_document_save',
            arguments: { sessionId, viewId, uploadId },
          })
          .catch(() => undefined)
      }
      const message = error instanceof Error ? error.message : String(error)
      status.error(`保存失败：${message}`)
      return { ok: false, error: message }
    }
  }
  const saveFile = async (file: LiveEditorFileSave): Promise<LiveEditorFileSaveResult> => {
    if (documentChanging || saveTask)
      return { ok: false, error: 'A document operation is already in progress.' }
    const task = (async () => {
      await recoveryTask
      return requestLocalSave(file)
    })()
    saveTask = task
    try {
      return await task
    } finally {
      if (saveTask === task) saveTask = null
    }
  }
  liveFileSaver = saveFile
  const replaceDocument = async <T>(replace: () => Promise<T>): Promise<T> => {
    if (terminalError || !editingReady || !brokerSessionId || documentChanging)
      throw new Error(terminalError || 'The document session is not ready.')
    documentChanging = true
    let releaseReplacement!: () => void
    replacementBarrier = new Promise<void>((resolve) => {
      releaseReplacement = resolve
    })
    stopRecoveryTimer()
    pollLoop.pause()
    try {
      await saveTask
      await recoveryTask
      const response = await app.callServerTool({
        name: 'office_editor_reset_document',
        arguments: {
          sessionId,
          viewId,
          ...(executingCommandId ? { commandId: executingCommandId } : {}),
        },
      })
      const payload = response.structuredContent as { ok?: boolean; message?: string } | undefined
      if (response.isError || payload?.ok !== true)
        throw new Error(payload?.message ?? 'Cannot detach the previous document.')
      bindingVersion += 1
      boundPath = null
      pathKnown = true
      status.path(null)
      storedRecoveryVersions.delete(sessionId)
      const result = await replace()
      await new Promise<void>((resolve) => window.requestAnimationFrame(() => resolve()))
      await checkpoint(sessionId, true)
      return result
    } finally {
      documentChanging = false
      replacementBarrier = null
      releaseReplacement()
      startRecoveryTimer()
      pollLoop.resume()
    }
  }
  documentReplacer = replaceDocument
  const isEditorVisible = (): boolean => document.visibilityState !== 'hidden'
  const isEditorPaintActive = (): boolean =>
    editorIntersectionKnown ? editorIntersecting : isEditorVisible()
  const isEditorWorkActive = (): boolean => isEditorPaintActive() && !editorContentSkipped
  let editorWorkActive = isEditorWorkActive()
  const syncEditorPaintActivity = (): void => {
    const nextPaintActive = isEditorPaintActive()
    const nextWorkActive = nextPaintActive && !editorContentSkipped
    document.documentElement.dataset.liveEditorActive = String(nextPaintActive)
    publishLiveEditorActivity(nextWorkActive)
    if (nextWorkActive === editorWorkActive) return
    editorWorkActive = nextWorkActive
    if (nextWorkActive && resumeOnVisible) resumeConnection()
    if (!nextWorkActive && !continueHerePending) {
      window.clearTimeout(resumeTimer)
      resumeTimer = 0
    }
    if (!displayStarted) return
    if (nextWorkActive) {
      startRecoveryTimer()
      return
    }
    stopRecoveryTimer()
    if (sessionId) void checkpoint(sessionId)
  }
  syncEditorPaintActivity()
  const editorRoot = document.getElementById('root')
  const onContentVisibilityAutoStateChange = (event: Event): void => {
    editorContentSkipped = (event as Event & { skipped?: unknown }).skipped === true
    syncEditorPaintActivity()
  }
  editorRoot?.addEventListener(
    'contentvisibilityautostatechange',
    onContentVisibilityAutoStateChange,
  )
  const editorIntersectionObserver =
    typeof IntersectionObserver === 'function'
      ? new IntersectionObserver((entries) => {
          const entry = entries.at(-1)
          if (!entry) return
          editorIntersectionKnown = true
          editorIntersecting = entry.isIntersecting && entry.intersectionRatio > 0
          syncEditorPaintActivity()
        })
      : null
  editorIntersectionObserver?.observe(document.documentElement)

  const readLocalAsset = async (
    rootId: string,
    path: string,
  ): Promise<LiveEditorLocalAsset | null> => {
    if (!sessionId) throw new Error('The live-editor session is not connected.')
    let output: Uint8Array | null = null
    let mime: LiveEditorLocalAsset['mime'] | null = null
    let offset = 0
    while (output === null || offset < output.length) {
      const response = await app.callServerTool({
        name: 'office_editor_read_local_asset_chunk',
        arguments: {
          sessionId,
          viewId,
          rootId,
          path,
          offset,
          length: Math.min(262_144, output === null ? 262_144 : output.length - offset),
        },
      })
      const payload = response.structuredContent as
        | {
            ok?: boolean
            size?: number
            mime?: string
            data?: string
            nextOffset?: number
          }
        | undefined
      if (!payload?.ok) return null
      if (
        !Number.isInteger(payload.size) ||
        (payload.size as number) < 1 ||
        (payload.size as number) > MAX_LOCAL_ASSET_BYTES ||
        !['image/png', 'image/jpeg', 'image/gif'].includes(payload.mime ?? '') ||
        typeof payload.data !== 'string' ||
        typeof payload.nextOffset !== 'number'
      ) {
        throw new Error('The local asset bridge returned an invalid descriptor.')
      }
      if (output === null) {
        output = new Uint8Array(payload.size as number)
        mime = payload.mime as LiveEditorLocalAsset['mime']
      } else if (payload.size !== output.length || payload.mime !== mime) {
        throw new Error('The local asset changed while it was being read.')
      }
      const chunk = decodeBase64(payload.data)
      if (chunk.length === 0 || payload.nextOffset !== offset + chunk.length) {
        throw new Error('The local asset bridge returned invalid offsets.')
      }
      output.set(chunk, offset)
      offset = payload.nextOffset
    }
    return {
      mime: mime!,
      data: output.buffer.slice(
        output.byteOffset,
        output.byteOffset + output.byteLength,
      ) as ArrayBuffer,
    }
  }
  localAssetReader = readLocalAsset

  const readBundledFont = async (fileName: string): Promise<ArrayBuffer> => {
    const chunks: Uint8Array[] = []
    let size = 0
    let offset = 0
    let eof = false
    while (!eof) {
      const response = await app.callServerTool({
        name: 'office_editor_read_font_chunk',
        arguments: { fileName, offset, length: 262_144 },
      })
      const payload = response.structuredContent as
        { ok?: boolean; data?: string; nextOffset?: number; eof?: boolean } | undefined
      if (
        !payload?.ok ||
        typeof payload.data !== 'string' ||
        typeof payload.nextOffset !== 'number' ||
        payload.nextOffset < offset
      ) {
        throw new Error(`The bundled font asset could not be read: ${fileName}`)
      }
      const chunk = decodeBase64(payload.data)
      if (chunk.length !== payload.nextOffset - offset || (chunk.length === 0 && !payload.eof)) {
        throw new Error(`The bundled font asset returned invalid offsets: ${fileName}`)
      }
      chunks.push(chunk)
      size += chunk.length
      if (size > MAX_LOCAL_ASSET_BYTES) {
        throw new Error(`The bundled font asset exceeds the browser limit: ${fileName}`)
      }
      offset = payload.nextOffset
      eof = payload.eof === true
    }
    const data = new Uint8Array(size)
    let cursor = 0
    for (const chunk of chunks) {
      data.set(chunk, cursor)
      cursor += chunk.length
    }
    return data.buffer
  }
  bundledFontReader = readBundledFont

  const bindSession = (requestedSessionId: unknown, requestedViewId: unknown): void => {
    if (
      typeof requestedSessionId !== 'string' ||
      requestedSessionId.length === 0 ||
      typeof requestedViewId !== 'string' ||
      requestedViewId.length === 0
    )
      return
    if (sessionId && (requestedSessionId !== sessionId || requestedViewId !== viewId)) return
    if (requestedSessionId === sessionId && requestedViewId === viewId) return
    sessionId = requestedSessionId
    viewId = requestedViewId
    if (appConnected) pollLoop.wake()
  }
  const configureApp = (target: McpApp): void => {
    target.ontoolinput = () => undefined
    target.ontoolresult = ({ structuredContent }) => {
      if ((structuredContent as { ok?: unknown } | undefined)?.ok !== true) return
      if (!sessionId && typeof structuredContent?.format === 'string')
        sessionFormat = structuredContent.format
      bindSession(
        (structuredContent as { sessionId?: unknown } | undefined)?.sessionId,
        (structuredContent as { viewId?: unknown } | undefined)?.viewId,
      )
    }
    target.onhostcontextchanged = () => displayController.sync(displayContext(target))
  }

  const poll = async (waitMs: number): Promise<void> => {
    if (stopped) return
    if (!sessionId || !viewId) return
    const polledSessionId = sessionId
    const polledBindingVersion = bindingVersion
    const requestedActivation = activateView
    const { fileName, dirty, selection } = adapter.snapshot(0)
    const response = await app.callServerTool({
      name: 'office_editor_poll',
      arguments: {
        sessionId: polledSessionId,
        viewId,
        ...(sessionFormat ? { format: sessionFormat, coldStart } : {}),
        active: continueHerePending || isEditorWorkActive(),
        ...(retryHandoff ? { retryHandoff: true } : {}),
        ...(requestedActivation ? { activateView: true } : {}),
        fileName,
        dirty,
        selection,
        ...(startupTracePending ? { startupTrace } : {}),
        waitMs,
      },
    })
    if (stopped) return
    retryHandoff = false
    if (requestedActivation) activateView = false
    const payload = response.structuredContent as
      | {
          ok?: unknown
          error?: string
          message?: string
          filePath?: string | null
          restoreCommandId?: string
          handoffRequest?: string
          handoffRequestedByUser?: boolean
          retryAfterMs?: number
          handoffFailed?: boolean
          commands?: LiveEditorCommand[]
        }
      | undefined
    if (
      polledBindingVersion === bindingVersion &&
      payload?.filePath !== undefined &&
      (!pathKnown || payload.filePath !== boundPath)
    ) {
      pathKnown = true
      boundPath = payload.filePath
      status.path(boundPath)
    }
    if (response.isError || payload?.ok !== true) {
      if (
        payload?.error === 'editor_view_conflict' ||
        payload?.error === 'session_not_found' ||
        payload?.error === 'document_restore_failed'
      ) {
        terminalError = payload.error
        editingReady = false
        if (waitExpired) return
        retryableConflict =
          terminalError === 'editor_view_conflict' && Boolean(payload.retryAfterMs)
        resumeOnVisible = terminalError === 'editor_view_conflict' && !payload.handoffFailed
        if (
          terminalError === 'editor_view_conflict' &&
          continueHerePending &&
          payload.handoffFailed !== true
        ) {
          showContinuationPending()
        } else {
          status.activation(retryableConflict || payload.handoffFailed === true)
          status.block(
            retryableConflict
              ? '正在等待原编辑器安全交接，原文件尚未恢复…'
              : '原文件尚未恢复，请查看连接状态。',
          )
          status.error(
            terminalError === 'editor_view_conflict'
              ? payload.handoffFailed
                ? '交接失败，原编辑器及内容已保留。请排查保存或连接问题后重试连接。'
                : '正在等待原编辑器安全交接；若原编辑器仍可见，请继续使用它。'
              : (payload.message ?? '未找到此会话的文件或恢复快照。请重新打开原文件。'),
            true,
          )
        }
        document.documentElement.dataset.liveEditorConnection = terminalError
        pollLoop.stop()
        stopRecoveryTimer()
        if (retryableConflict) {
          beginWait()
          scheduleResume()
        } else clearWait()
        return
      }
      throw new Error(payload?.message ?? 'The editor poll did not succeed.')
    }
    clearWait()
    if (document.documentElement.dataset.liveEditorConnection !== 'connected') status.error(null)
    startupTracePending = false
    if (sessionId === polledSessionId) brokerSessionId = polledSessionId
    const commands = payload.commands ?? []
    for (const command of commands) {
      await replacementBarrier
      if (stopped) return
      const stopFailedRestore = (message: string): boolean => {
        if (payload.restoreCommandId !== command.commandId) return false
        terminalError = 'document_restore_failed'
        editingReady = false
        status.block('原文件恢复失败，已禁止编辑空白副本。')
        status.error(`恢复失败：${message}`, true)
        document.documentElement.dataset.liveEditorConnection = terminalError
        pollLoop.stop()
        stopRecoveryTimer()
        return true
      }
      let hydrated = command
      let hydrateMs = 0
      if (isStagedFileOperation(command.operation)) {
        const hydrateStartedAt = performance.now()
        try {
          hydrated = {
            ...command,
            arguments: {
              ...command.arguments,
              data: await readStagedFile(app, polledSessionId, viewId, command.arguments),
            },
          }
          hydrateMs = performance.now() - hydrateStartedAt
        } catch (error) {
          const message = error instanceof Error ? error.message : String(error)
          const restoreFailed = stopFailedRestore(message)
          await app.callServerTool({
            name: 'office_editor_acknowledge',
            arguments: {
              sessionId: polledSessionId,
              viewId,
              commandId: command.commandId,
              ok: false,
              error: 'execution_failed',
              message,
            },
          })
          if (restoreFailed) return
          continue
        }
      }
      const executeStartedAt = performance.now()
      executingCommandId = command.commandId
      let execution: LiveEditorExecution
      try {
        execution = await adapter.execute(hydrated)
      } catch (error) {
        execution = {
          ok: false,
          error: 'execution_failed',
          message: error instanceof Error ? error.message : String(error),
        }
      } finally {
        executingCommandId = undefined
      }
      const executeMs = performance.now() - executeStartedAt
      if (!execution.ok) {
        const restoreFailed = stopFailedRestore(execution.message)
        await app.callServerTool({
          name: 'office_editor_acknowledge',
          arguments: {
            sessionId: polledSessionId,
            viewId,
            commandId: command.commandId,
            ...execution,
          },
        })
        if (restoreFailed) return
        continue
      }
      if (execution.recovery) {
        const recoveryVersion = adapter.recoveryVersion?.()
        try {
          await storeRecovery(app, polledSessionId, viewId, execution.recovery)
          if (recoveryVersion !== undefined)
            storedRecoveryVersions.set(polledSessionId, recoveryVersion)
        } catch {
          // Best effort: the mounted renderer remains authoritative.
        }
      }
      await app.callServerTool({
        name: 'office_editor_acknowledge',
        arguments: {
          sessionId: polledSessionId,
          viewId,
          commandId: command.commandId,
          ok: true,
          ...(execution.output ? { output: execution.output } : {}),
          timing: {
            hydrateMs,
            executeMs,
            ...(execution.trace ? { trace: execution.trace } : {}),
          },
          ...adapter.snapshot(command.baseRevision + 1),
        },
      })
    }
    editingReady = true
    retryableConflict = false
    resumeOnVisible = false
    suspended = false
    resumeDelay = 1_000
    coldStart = false
    status.block(null)
    status.activation(false)
    document.documentElement.dataset.liveEditorConnection = 'connected'
    if (payload.handoffRequest && (payload.handoffRequestedByUser || !isEditorWorkActive())) {
      await handoff(payload.handoffRequest, payload.handoffRequestedByUser === true)
      if (suspended) return
    }
    if (!displayStarted) {
      displayStarted = true
      startRecoveryTimer()
      await new Promise<void>((resolve) => window.requestAnimationFrame(() => resolve()))
      if (stopped) return
      void displayController.connect(
        { requestDisplayMode: async (mode) => app.requestDisplayMode({ mode }) },
        displayContext(app),
      )
    }
  }

  async function handoff(handoffId: string, userActivated = false): Promise<void> {
    if (handoffInProgress || documentChanging || executingCommandId || !adapter.recoverySnapshot)
      return
    handoffInProgress = true
    editingReady = false
    status.block('正在保护当前内容并交接编辑器…')
    stopRecoveryTimer()
    const transfer = async (action: 'prepare' | 'commit' | 'abort') => {
      const response = await app.callServerTool({
        name: 'office_editor_handoff',
        arguments: { sessionId, viewId, handoffId, action },
      })
      const value = response.structuredContent as { ok?: boolean; message?: string } | undefined
      if (response.isError || !value?.ok)
        throw new Error(value?.message ?? 'Editor handoff failed.')
    }
    try {
      await saveTask
      await recoveryTask
      if (!userActivated && isEditorWorkActive()) return
      await transfer('prepare')
      const recovery = await adapter.recoverySnapshot(true)
      if (!recovery) throw new Error('无法生成恢复快照；已保留原编辑器。')
      await storeRecovery(app, sessionId, viewId, recovery)
      await transfer('commit')
      suspended = true
      resumeOnVisible = false
      coldStart = true
      brokerSessionId = ''
      terminalError = 'suspended'
      document.documentElement.dataset.liveEditorConnection = 'suspended'
      status.error('此视图已暂停；如需在这里继续，请点击“在此继续编辑”。')
      status.activation(true)
      status.block('此视图已暂停，内容已安全交接。')
      pollLoop.stop()
    } catch (error) {
      // A failed transport response may follow a successful prepare or commit. Only a
      // lease-checked abort confirms that this mount still owns an editable document.
      const aborted = await transfer('abort').then(
        () => true,
        () => false,
      )
      if (aborted) {
        status.error(
          `交接失败，已保留原编辑器：${error instanceof Error ? error.message : String(error)}`,
          true,
        )
      } else {
        uncertainHandoff = handoffId
        brokerSessionId = ''
        suspended = true
        resumeOnVisible = false
        terminalError = 'handoff_uncertain'
        document.documentElement.dataset.liveEditorConnection = terminalError
        status.block('交接结果尚未确认，已锁定此视图并保留内容。')
        status.error(
          '交接结果尚未确认，不能安全恢复编辑。请保留此视图，重试连接以确认编辑权。',
          true,
        )
        status.activation(false)
        pollLoop.stop()
      }
    } finally {
      handoffInProgress = false
      if (!suspended) {
        editingReady = true
        status.block(null)
        startRecoveryTimer()
      }
    }
  }

  async function reconcileUncertainHandoff(): Promise<void> {
    if (!uncertainHandoff || handoffInProgress || stopped) return
    handoffInProgress = true
    let stillOwner = false
    try {
      const response = await app.callServerTool({
        name: 'office_editor_handoff',
        arguments: { sessionId, viewId, handoffId: uncertainHandoff, action: 'abort' },
      })
      if (stopped) return
      const payload = response.structuredContent as
        { ok?: boolean; error?: string; message?: string } | undefined
      if (!response.isError && payload?.ok) {
        uncertainHandoff = null
        terminalError = ''
        suspended = false
        editingReady = true
        stillOwner = true
        status.block(null)
        status.error(null)
      } else if (payload?.error === 'editor_view_conflict') {
        uncertainHandoff = null
        coldStart = true
        terminalError = 'suspended'
        document.documentElement.dataset.liveEditorConnection = terminalError
        status.block('此视图已暂停，内容已安全交接。')
        status.error('已确认编辑权在另一个实例。如需在这里继续，请点击“在此继续编辑”。')
        status.activation(true)
      } else throw new Error(payload?.message ?? '无法确认编辑权')
    } catch (error) {
      if (!stopped)
        status.error(
          `交接结果尚未确认：${error instanceof Error ? error.message : String(error)}`,
          true,
        )
    } finally {
      handoffInProgress = false
      if (stillOwner && !stopped) resumeConnection()
    }
  }

  let pollLoop = createWakeablePollLoop(poll)

  async function checkpoint(targetSessionId: string, force = false): Promise<void> {
    if (stopped || terminalError || !targetSessionId || !adapter.recoverySnapshot) return
    if (!force && (documentChanging || saveTask || !adapter.snapshot(0).dirty)) return
    const recoveryVersion = adapter.recoveryVersion?.()
    if (
      !force &&
      recoveryVersion !== undefined &&
      storedRecoveryVersions.get(targetSessionId) === recoveryVersion
    )
      return
    if (recoveryTask) return recoveryTask
    const task = (async () => {
      try {
        const recovery = await adapter.recoverySnapshot!(force)
        if (recovery) {
          await storeRecovery(app, targetSessionId, viewId, recovery)
          if (recoveryVersion !== undefined)
            storedRecoveryVersions.set(targetSessionId, recoveryVersion)
        }
      } catch {
        // Best effort: the mounted renderer remains authoritative.
      }
    })()
    recoveryTask = task
    try {
      await task
    } finally {
      if (recoveryTask === task) recoveryTask = null
    }
  }

  function stopRecoveryTimer(): void {
    window.clearInterval(recoveryTimer)
    recoveryTimer = 0
  }

  function startRecoveryTimer(): void {
    if (recoveryTimer || stopped || terminalError || !isEditorWorkActive()) return
    recoveryTimer = window.setInterval(() => void checkpoint(sessionId), 2_000)
  }

  const onVisibilityChange = (): void => {
    syncEditorPaintActivity()
  }
  document.addEventListener('visibilitychange', onVisibilityChange)

  const connectApp = async (): Promise<void> => {
    const target = app
    configureApp(target)
    try {
      await target.connect(undefined, { timeout: MCP_CONNECT_TIMEOUT_MS })
      if (stopped || app !== target) return
      const startupTracePromise = adapter.startupTrace?.()
      await new Promise<void>((resolve) => window.requestAnimationFrame(() => resolve()))
      if (stopped || app !== target) return
      startupTrace = await startupTracePromise
      startupTracePending = startupTrace !== undefined
      if (stopped || app !== target) return
      appConnected = true
      if (sessionId && viewId) pollLoop.start()
    } catch {
      if (stopped || app !== target) return
      await target.close().catch(() => undefined)
      if (stopped || app !== target) return
      app = createApp()
      connectRetryTimer = window.setTimeout(() => {
        connectRetryTimer = 0
        void connectApp()
      }, MCP_CONNECT_RETRY_MS)
    }
  }
  void connectApp()

  return () => {
    stopped = true
    pollLoop.stop()
    window.clearTimeout(resumeTimer)
    clearWait()
    window.clearTimeout(connectRetryTimer)
    connectRetryTimer = 0
    stopRecoveryTimer()
    document.removeEventListener('visibilitychange', onVisibilityChange)
    editorRoot?.removeEventListener(
      'contentvisibilityautostatechange',
      onContentVisibilityAutoStateChange,
    )
    editorIntersectionObserver?.disconnect()
    delete document.documentElement.dataset.liveEditorActive
    delete document.documentElement.dataset.liveEditorConnection
    publishLiveEditorActivity(true)
    if (sessionId && viewId && brokerSessionId === sessionId) {
      void app.callServerTool({
        name: 'office_editor_disconnect',
        arguments: { sessionId, viewId },
      })
    }
    displayController.disconnect()
    if (localAssetReader === readLocalAsset) localAssetReader = null
    if (bundledFontReader === readBundledFont) bundledFontReader = null
    if (liveFileSaver === saveFile) liveFileSaver = null
    if (documentReplacer === replaceDocument) documentReplacer = null
    status.dispose()
    void app.close().catch(() => undefined)
  }
}
