import type { LiveEditorExecution } from '@tandemfolio/host-bridge'
import { executePdfOperation, type PdfOperationServices } from '../operations/registry'

export interface PdfCommunityController extends Required<
  Omit<PdfOperationServices, 'loadStaged' | 'insertPagesStaged'>
> {
  recoverySnapshot(): Promise<{ fileName: string; data: ArrayBuffer } | null>
  recoveryVersion(): number
}

export interface PdfCommunityCommandBridge {
  execute(operation: string, arguments_: Record<string, unknown>): Promise<LiveEditorExecution>
  recoverySnapshot(): Promise<{ fileName: string; data: ArrayBuffer } | null>
  recoveryVersion(): number
  register(controller: PdfCommunityController): () => void
}

export function createPdfCommunityCommandBridge(): PdfCommunityCommandBridge {
  let controller: PdfCommunityController | null = null
  return {
    register(next) {
      controller = next
      return () => {
        if (controller === next) controller = null
      }
    },
    async execute(operation, arguments_) {
      if (!controller) {
        return {
          ok: false,
          error: 'execution_failed',
          message: 'The mounted PDF community renderer is not ready.',
        }
      }
      const activeController = controller
      const registered = await executePdfOperation(
        { operation, arguments: arguments_ },
        activeController,
      )
      if (registered.handled) {
        return registered.ok
          ? { ok: true, output: { ...registered.output } }
          : { ok: false, error: registered.error, message: registered.message }
      }
      return {
        ok: false,
        error: 'unsupported_operation',
        message: `PDF community operation is not connected yet: ${operation}`,
      }
    },
    async recoverySnapshot() {
      return controller?.recoverySnapshot() ?? null
    },
    recoveryVersion() {
      return controller?.recoveryVersion() ?? 0
    },
  }
}
