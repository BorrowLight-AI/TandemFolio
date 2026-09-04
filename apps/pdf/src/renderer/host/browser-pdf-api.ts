import type {
  ExportImagesRequest,
  ExtractPagesResult,
  InsertPdfResult,
  PdfApi,
  SavePdfResult,
  TextEditValidation,
} from '../../shared/ipc'
import {
  saveLiveEditorFile,
  type LiveEditorAdapter,
  type LiveEditorCommand,
  type LiveEditorExecution,
  type LiveEditorSnapshot,
} from '@tandemfolio/host-bridge'
import {
  applySaveRequest,
  extractPagesBytes,
  insertPdfBytes,
  readStaticFormFills,
} from '../../domain/save-pdf'
import {
  listBrowserPageImages,
  renderBrowserImagePng,
  renderBrowserPagePreviewPng,
} from '../../domain/browser-image-edit'
import { validateBrowserTextEdits } from '../../domain/browser-text-edit'
import type { PdfCommunityCommandBridge } from './community-command-bridge'
import { executePdfOperation, resolvePdfOperationId } from '../operations/registry'

export const PDF_OPEN_EVENT = 'tandemfolio:pdf-open'

function cloneBuffer(data: ArrayBuffer): ArrayBuffer {
  return data.slice(0)
}

type PdfPersistMode = 'save' | 'export-copy'

async function download(
  fileName: string,
  data: ArrayBuffer,
  mode: PdfPersistMode,
): Promise<boolean> {
  if (window.parent !== window) {
    return (await saveLiveEditorFile({ fileName, data, mode })).ok
  }
  const url = URL.createObjectURL(new Blob([data], { type: 'application/pdf' }))
  const anchor = document.createElement('a')
  anchor.href = url
  anchor.download = fileName
  anchor.click()
  URL.revokeObjectURL(url)
  return true
}

function pickPdfFile(): Promise<File | null> {
  return new Promise((resolve) => {
    const input = document.createElement('input')
    input.type = 'file'
    input.accept = '.pdf,application/pdf'
    input.onchange = () => resolve(input.files?.[0] ?? null)
    input.addEventListener('cancel', () => resolve(null), { once: true })
    input.click()
  })
}

export interface BrowserPdfHost {
  api: PdfApi
  adapter: LiveEditorAdapter
  stageFile(fileName: string, data: ArrayBuffer): Promise<void>
  completeOpen(path: string, error?: unknown): void
}

export interface BrowserPdfHostOptions {
  commandBridge?: PdfCommunityCommandBridge
  download?: (fileName: string, data: ArrayBuffer) => boolean | void | Promise<boolean | void>
}

export function createBrowserPdfHost(options: BrowserPdfHostOptions = {}): BrowserPdfHost {
  const files = new Map<string, ArrayBuffer>()
  const pending: string[] = []
  let consumeStarted = false
  let activeFile: string | null = null
  let dirty = false
  const openWaiters = new Map<string, { resolve: () => void; reject: (error: unknown) => void }>()
  const persist = (
    fileName: string,
    data: ArrayBuffer,
    mode: PdfPersistMode,
  ): boolean | void | Promise<boolean | void> =>
    options.download ? options.download(fileName, data) : download(fileName, data, mode)

  const stageFile = (fileName: string, data: ArrayBuffer): Promise<void> => {
    const path = `browser://${crypto.randomUUID()}/${encodeURIComponent(fileName)}`
    files.set(path, cloneBuffer(data))
    activeFile = path
    return new Promise<void>((resolveOpen, rejectOpen) => {
      openWaiters.set(path, { resolve: resolveOpen, reject: rejectOpen })
      if (consumeStarted) {
        window.dispatchEvent(new CustomEvent(PDF_OPEN_EVENT, { detail: { path } }))
      } else {
        pending.push(path)
      }
    })
  }

  const completeOpen = (path: string, error?: unknown): void => {
    const waiter = openWaiters.get(path)
    if (!waiter) return
    openWaiters.delete(path)
    if (error === undefined) waiter.resolve()
    else waiter.reject(error)
  }

  const insertPages = async (data: ArrayBuffer, afterPageIndex: number): Promise<number> => {
    if (!activeFile) throw new Error('No PDF document is open.')
    const current = files.get(activeFile)
    if (!current) throw new Error('The selected PDF is no longer available.')
    const { merged, count } = await insertPdfBytes(
      new Uint8Array(current),
      new Uint8Array(data),
      afterPageIndex,
    )
    if (count < 1) throw new Error('The selected PDF contains no pages.')
    const mergedBuffer = merged.buffer.slice(
      merged.byteOffset,
      merged.byteOffset + merged.byteLength,
    ) as ArrayBuffer
    files.set(activeFile, mergedBuffer)
    const fileName = decodeURIComponent(activeFile.split('/').pop() ?? 'document.pdf')
    if ((await persist(fileName, mergedBuffer, 'save')) === false) {
      throw new Error('The local PDF save failed.')
    }
    window.dispatchEvent(new CustomEvent(PDF_OPEN_EVENT, { detail: { path: activeFile } }))
    dirty = false
    return count
  }

  const api: PdfApi = {
    consumePending: async () => {
      consumeStarted = true
      return pending.shift() ?? null
    },
    readFile: async (path) => {
      const data = files.get(path)
      if (!data) throw new Error('The selected PDF is no longer available.')
      return cloneBuffer(data)
    },
    save: async (request): Promise<SavePdfResult> => {
      const data = files.get(request.path)
      if (!data) return { ok: false, error: 'The selected PDF is no longer available.' }
      try {
        const saved = await applySaveRequest(new Uint8Array(data), request)
        const savedBuffer = saved.bytes.buffer.slice(
          saved.bytes.byteOffset,
          saved.bytes.byteOffset + saved.bytes.byteLength,
        ) as ArrayBuffer
        files.set(request.path, savedBuffer)
        const fileName = decodeURIComponent(
          (request.targetPath ?? request.path).split(/[\\/]/).pop() ?? 'document.pdf',
        )
        if (
          (await persist(fileName, savedBuffer, request.targetPath ? 'export-copy' : 'save')) ===
          false
        ) {
          return { ok: false, error: 'The local PDF save failed.' }
        }
        dirty = false
        return {
          ok: true,
          skippedTextEdits: saved.skippedTextEdits,
          skippedTextInserts: saved.skippedTextInserts,
          skippedImageEdits: saved.skippedImageEdits,
        }
      } catch (error) {
        return {
          ok: false,
          error: error instanceof Error ? error.message : String(error),
        }
      }
    },
    createRecovery: async (request) => {
      const data = files.get(request.path)
      if (!data) throw new Error('The selected PDF is no longer available.')
      const saved = await applySaveRequest(new Uint8Array(data), request)
      return saved.bytes.buffer.slice(
        saved.bytes.byteOffset,
        saved.bytes.byteOffset + saved.bytes.byteLength,
      ) as ArrayBuffer
    },
    validateTextEdits: async (request) => {
      const data = files.get(request.path)
      if (!data) return request.edits.map((): TextEditValidation => ({ reason: 'PDF unavailable' }))
      return validateBrowserTextEdits(new Uint8Array(data), request.edits)
    },
    listEditFonts: async () => ['arial', 'times', 'courier'],
    listPageImages: async (path) => {
      const data = files.get(path)
      return data ? listBrowserPageImages(new Uint8Array(data)) : []
    },
    listStaticFormFills: async (path) => {
      const data = files.get(path)
      return data ? readStaticFormFills(new Uint8Array(data)) : []
    },
    pageImagePng: async (request) => {
      const data = files.get(request.path)
      return data
        ? renderBrowserImagePng(
            new Uint8Array(data),
            request.pageIndex,
            request.rect,
            request.scale,
          )
        : null
    },
    pagePreviewPng: async ({ path, ...request }) => {
      const data = files.get(path)
      return data ? renderBrowserPagePreviewPng(new Uint8Array(data), request) : null
    },
    extractPages: async (request): Promise<ExtractPagesResult> => {
      const data = files.get(request.path)
      if (!data) return { ok: false, error: 'The selected PDF is no longer available.' }
      try {
        const extracted = await extractPagesBytes(new Uint8Array(data), request.pages)
        const buffer = extracted.buffer.slice(
          extracted.byteOffset,
          extracted.byteOffset + extracted.byteLength,
        ) as ArrayBuffer
        if ((await persist(request.suggestedName, buffer, 'export-copy')) === false) {
          return { ok: false, error: 'The local PDF export failed.' }
        }
        return { ok: true, savedPath: request.suggestedName }
      } catch (error) {
        return { ok: false, error: error instanceof Error ? error.message : String(error) }
      }
    },
    insertPdf: async (request): Promise<InsertPdfResult> => {
      try {
        const file = await pickPdfFile()
        if (!file) return { ok: true, canceled: true }
        return {
          ok: true,
          insertedCount: await insertPages(await file.arrayBuffer(), request.afterPageIndex),
        }
      } catch (error) {
        return { ok: false, error: error instanceof Error ? error.message : String(error) }
      }
    },
    exportImages: async (request: ExportImagesRequest) => {
      for (let index = 0; index < request.images.length; index += 1) {
        const binary = atob(request.images[index] ?? '')
        const bytes = Uint8Array.from(binary, (character) => character.charCodeAt(0))
        const url = URL.createObjectURL(new Blob([bytes], { type: 'image/png' }))
        const anchor = document.createElement('a')
        anchor.href = url
        anchor.download = `${request.baseName}-${request.pageNumbers[index] ?? index + 1}.png`
        anchor.click()
        URL.revokeObjectURL(url)
      }
      return { ok: true, savedDir: 'browser-downloads', count: request.images.length }
    },
    setDirty: (next) => {
      dirty = next
    },
    onCloseSaveRequest: () => () => undefined,
    sendCloseSaveResult: () => undefined,
    onSaveAsRequest: () => () => undefined,
    sendSaveAsResult: () => undefined,
    onSaveAsFlow: () => () => undefined,
    getLanguage: async () => 'zh',
    onLanguageChanged: () => () => undefined,
    getTheme: async () => 'system',
    onThemeChanged: () => () => undefined,
  }

  const execute = async (command: LiveEditorCommand): Promise<LiveEditorExecution> => {
    if (resolvePdfOperationId(command.operation) === 'pdf.document.load_staged') {
      const registered = await executePdfOperation(
        { operation: command.operation, arguments: command.arguments },
        { loadStaged: ({ name, data }) => stageFile(name, data) },
      )
      if (registered.handled) {
        return registered.ok
          ? { ok: true, output: { ...registered.output } }
          : { ok: false, error: registered.error, message: registered.message }
      }
    }
    if (resolvePdfOperationId(command.operation) === 'pdf.page.insert_staged') {
      const registered = await executePdfOperation(
        { operation: command.operation, arguments: command.arguments },
        {
          insertPagesStaged: ({ data, afterPageIndex }) => insertPages(data, afterPageIndex),
        },
      )
      if (registered.handled) {
        return registered.ok
          ? { ok: true, output: { ...registered.output } }
          : { ok: false, error: registered.error, message: registered.message }
      }
    }
    return (
      options.commandBridge?.execute(command.operation, command.arguments) ?? {
        ok: false,
        error: 'unsupported_operation',
        message: `PDF community operation is not connected yet: ${command.operation}`,
      }
    )
  }

  const snapshot = (revision: number): LiveEditorSnapshot => ({
    revision,
    fileName: activeFile ? decodeURIComponent(activeFile.split('/').pop() ?? '') : null,
    dirty,
    selection: null,
  })

  const recoverySnapshot = async (force = false) => {
    if ((!force && !dirty) || !options.commandBridge) return null
    return options.commandBridge.recoverySnapshot(force)
  }

  return {
    api,
    adapter: {
      execute,
      snapshot,
      recoverySnapshot,
      recoveryVersion: () => options.commandBridge?.recoveryVersion() ?? 0,
    },
    stageFile,
    completeOpen,
  }
}
