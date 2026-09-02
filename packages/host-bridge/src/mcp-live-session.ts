import { App as McpApp } from '@modelcontextprotocol/ext-apps'
import {
  createMcpDisplayModeController,
  type McpDisplayMode,
  type McpDisplayModeState,
} from './display-mode'
import { createWakeablePollLoop } from './wakeable-poll-loop'

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
  recoverySnapshot?(): Promise<{ fileName: string; data: ArrayBuffer } | null>
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
  const createApp = (): McpApp =>
    new McpApp(
      { name: 'tandemfolio-editor', version: '0.1.0' },
      { availableDisplayModes: ['inline', 'fullscreen'] },
    )
  let app = createApp()
  let stopped = false
  let sessionId = new URLSearchParams(window.location.search).get('sessionId') ?? ''
  let viewId = new URLSearchParams(window.location.search).get('viewId') ?? ''
  let brokerSessionId = ''
  let connectRetryTimer = 0
  let recoveryTimer = 0
  let recoveryTask: Promise<void> | null = null
  const storedRecoveryVersions = new Map<string, string | number>()
  let appConnected = false
  let displayStarted = false
  let startupTrace: LiveEditorStartupTrace | undefined
  let startupTracePending = false
  let editorIntersectionKnown = false
  let editorIntersecting = true
  let editorContentSkipped = false
  const requestLocalSave = async (file: LiveEditorFileSave): Promise<LiveEditorFileSaveResult> => {
    let uploadId = ''
    try {
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
      return { ok: false, error: error instanceof Error ? error.message : String(error) }
    }
  }
  liveFileSaver = requestLocalSave
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
    const { fileName, dirty, selection } = adapter.snapshot(0)
    const response = await app.callServerTool({
      name: 'office_editor_poll',
      arguments: {
        sessionId: polledSessionId,
        viewId,
        fileName,
        dirty,
        selection,
        ...(startupTracePending ? { startupTrace } : {}),
        waitMs,
      },
    })
    if ((response.structuredContent as { ok?: unknown } | undefined)?.ok === true) {
      startupTracePending = false
      if (sessionId === polledSessionId) brokerSessionId = polledSessionId
    }
    const commands =
      (response.structuredContent as { commands?: LiveEditorCommand[] } | undefined)?.commands ?? []
    for (const command of commands) {
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
          await app.callServerTool({
            name: 'office_editor_acknowledge',
            arguments: {
              sessionId: polledSessionId,
              viewId,
              commandId: command.commandId,
              ok: false,
              error: 'execution_failed',
              message: error instanceof Error ? error.message : String(error),
            },
          })
          continue
        }
      }
      const executeStartedAt = performance.now()
      const execution = await adapter.execute(hydrated)
      const executeMs = performance.now() - executeStartedAt
      if (!execution.ok) {
        await app.callServerTool({
          name: 'office_editor_acknowledge',
          arguments: {
            sessionId: polledSessionId,
            viewId,
            commandId: command.commandId,
            ...execution,
          },
        })
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

  const pollLoop = createWakeablePollLoop(poll)

  async function checkpoint(targetSessionId: string): Promise<void> {
    if (stopped || !targetSessionId || !adapter.recoverySnapshot) return
    if (!adapter.snapshot(0).dirty) return
    const recoveryVersion = adapter.recoveryVersion?.()
    if (
      recoveryVersion !== undefined &&
      storedRecoveryVersions.get(targetSessionId) === recoveryVersion
    )
      return
    if (recoveryTask) return recoveryTask
    const task = (async () => {
      try {
        const recovery = await adapter.recoverySnapshot!()
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
    if (recoveryTimer || stopped || !isEditorWorkActive()) return
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
    if (liveFileSaver === requestLocalSave) liveFileSaver = null
    void app.close().catch(() => undefined)
  }
}
