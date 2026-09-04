export type EditorHostKind = 'browser' | 'mcp-app'

export interface EditorOperation {
  name: string
  arguments: Record<string, unknown>
}

export interface EditorCommand {
  commandId: string
  sessionId: string
  baseRevision: number
  operation: EditorOperation
}

export interface EditorContext {
  sessionId: string
  revision: number
  format: 'docx' | 'markdown' | 'xlsx' | 'pptx' | 'pdf'
  fileName: string | null
  dirty: boolean
  selection: Record<string, unknown> | null
}

export interface CommandResult {
  commandId: string
  revision: number
  ok: boolean
  error?: 'revision_conflict' | 'unsupported_operation' | 'invalid_arguments'
  message?: string
}

export function detectEditorHost(win: Window = window): EditorHostKind {
  return win.parent !== win ? 'mcp-app' : 'browser'
}

export function nextRevision(current: number, command: EditorCommand): number {
  if (command.baseRevision !== current) {
    throw new Error(`revision_conflict:${command.baseRevision}:${current}`)
  }
  return current + 1
}

export {
  createMcpDisplayModeController,
  type McpDisplayMode,
  type McpDisplayModeContext,
  type McpDisplayModeController,
  type McpDisplayModeHost,
  type McpDisplayModeState,
} from './display-mode'
export {
  attachMcpLiveSession,
  getLiveEditorActivity,
  getLiveEditorDisplayMode,
  readLiveEditorBundledFontAsset,
  readLiveEditorLocalAsset,
  saveLiveEditorFile,
  replaceLiveEditorDocument,
  subscribeLiveEditorDisplayMode,
  subscribeLiveEditorActivity,
  toggleLiveEditorFullscreen,
  type LiveEditorAdapter,
  type LiveEditorCommand,
  type LiveEditorDisplayModeState,
  type LiveEditorExecution,
  type LiveEditorExecutionTrace,
  type LiveEditorFileSave,
  type LiveEditorFileSaveResult,
  type LiveEditorLocalAsset,
  type LiveEditorSnapshot,
  type LiveEditorStartupTrace,
} from './mcp-live-session'
export {
  createWakeablePollLoop,
  LIVE_POLL_RETRY_MS,
  LIVE_POLL_WAIT_MS,
  type WakeablePollLoop,
} from './wakeable-poll-loop'
