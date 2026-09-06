import type {
  ExportImagesRequest,
  ExtractPagesResult,
  InsertPdfResult,
  PdfApi,
  SavedSignature,
  SavePdfResult,
  SignatureData,
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
  cropPagesBytes,
  extractPagesBytes,
  insertBlankPageBytes,
  insertPdfBytes,
  mergePdfBytes,
  mergePagesBytes,
  replacePagesBytes,
  readStaticFormFills,
  setPageSizeBytes,
  splitPdfBytes,
  splitPagesBytes,
} from '../../domain/save-pdf'
import {
  listBrowserPageImages,
  renderBrowserImagePng,
  renderBrowserPagePreviewPng,
} from '../../domain/browser-image-edit'
import { canBrowserDrawText, validateBrowserTextEdits } from '../../domain/browser-text-edit'
import { blankPdfBytes } from '../../domain/blank-pdf'
import type { PdfCommunityCommandBridge } from './community-command-bridge'
import { executePdfOperation, resolvePdfOperationId } from '../operations/registry'

export const PDF_OPEN_EVENT = 'tandemfolio:pdf-open'
const SIGNATURE_LIBRARY_KEY = 'genoffice-lite:pdf-signatures'
const SIGNATURE_LIBRARY_LIMIT = 12
const DOCUMENT_HISTORY_LIMIT = 32

function loadSignatures(): SavedSignature[] {
  try {
    const parsed: unknown = JSON.parse(localStorage.getItem(SIGNATURE_LIBRARY_KEY) ?? '[]')
    return Array.isArray(parsed) ? (parsed as SavedSignature[]).slice(0, SIGNATURE_LIBRARY_LIMIT) : []
  } catch {
    return []
  }
}

function saveSignatures(signatures: SavedSignature[]): SavedSignature[] {
  const bounded = signatures.slice(0, SIGNATURE_LIBRARY_LIMIT)
  localStorage.setItem(SIGNATURE_LIBRARY_KEY, JSON.stringify(bounded))
  return bounded
}

function signatureKey(data: SignatureData): string {
  return JSON.stringify(data)
}

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

function pickPdfFiles(): Promise<File[]> {
  return new Promise((resolve) => {
    const input = document.createElement('input')
    input.type = 'file'
    input.accept = '.pdf,application/pdf'
    input.multiple = true
    input.onchange = () => resolve(Array.from(input.files ?? []))
    input.addEventListener('cancel', () => resolve([]), { once: true })
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
  pickPdfFiles?: () => Promise<{ name: string; data: ArrayBuffer }[]>
}

export function createBrowserPdfHost(options: BrowserPdfHostOptions = {}): BrowserPdfHost {
  const files = new Map<string, ArrayBuffer>()
  const documentHistory = new Map<string, { undo: ArrayBuffer[]; redo: ArrayBuffer[] }>()
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

  const historyFor = (path: string) => {
    const existing = documentHistory.get(path)
    if (existing) return existing
    const created = { undo: [] as ArrayBuffer[], redo: [] as ArrayBuffer[] }
    documentHistory.set(path, created)
    return created
  }

  const commitPageMutation = async (
    path: string,
    before: ArrayBuffer,
    next: ArrayBuffer,
  ): Promise<void> => {
    const fileName = decodeURIComponent(path.split('/').pop() ?? 'document.pdf')
    if ((await persist(fileName, next, 'save')) === false) {
      throw new Error('The local PDF save failed.')
    }
    files.set(path, next)
    const history = historyFor(path)
    history.undo.push(cloneBuffer(before))
    if (history.undo.length > DOCUMENT_HISTORY_LIMIT) history.undo.shift()
    history.redo = []
    window.dispatchEvent(new CustomEvent(PDF_OPEN_EVENT, { detail: { path } }))
    dirty = false
  }

  const stageFile = (fileName: string, data: ArrayBuffer): Promise<void> => {
    const path = `browser://${crypto.randomUUID()}/${encodeURIComponent(fileName)}`
    files.set(path, cloneBuffer(data))
    documentHistory.set(path, { undo: [], redo: [] })
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
    await commitPageMutation(activeFile, current, mergedBuffer)
    return count
  }

  const replacePages = async (
    data: ArrayBuffer,
    pages: number[],
  ): Promise<{ removed: number; inserted: number }> => {
    if (!activeFile) throw new Error('No PDF document is open.')
    const current = files.get(activeFile)
    if (!current) throw new Error('The selected PDF is no longer available.')
    const result = await replacePagesBytes(
      new Uint8Array(current),
      new Uint8Array(data),
      pages,
    )
    const buffer = result.merged.buffer.slice(
      result.merged.byteOffset,
      result.merged.byteOffset + result.merged.byteLength,
    ) as ArrayBuffer
    await commitPageMutation(activeFile, current, buffer)
    return { removed: result.removed, inserted: result.inserted }
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
    canDrawText: (text, font, bold, italic) => canBrowserDrawText(text, font, bold, italic),
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
        const picked = options.pickPdfFiles ? (await options.pickPdfFiles())[0] : undefined
        const file = picked ? null : await pickPdfFile()
        if (!picked && !file) return { ok: true, canceled: true }
        return {
          ok: true,
          insertedCount: await insertPages(
            picked?.data ?? (await file!.arrayBuffer()),
            request.afterPageIndex,
          ),
        }
      } catch (error) {
        return { ok: false, error: error instanceof Error ? error.message : String(error) }
      }
    },
    insertBlankPage: async (request) => {
      const data = files.get(request.path)
      if (!data) return { ok: false, error: 'The selected PDF is no longer available.' }
      try {
        const bytes = await insertBlankPageBytes(
          new Uint8Array(data),
          request.afterPageIndex,
        )
        const buffer = bytes.buffer.slice(
          bytes.byteOffset,
          bytes.byteOffset + bytes.byteLength,
        ) as ArrayBuffer
        await commitPageMutation(request.path, data, buffer)
        return { ok: true }
      } catch (error) {
        return { ok: false, error: error instanceof Error ? error.message : String(error) }
      }
    },
    splitPdf: async (request) => {
      const data = files.get(request.path)
      if (!data) return { ok: false, error: 'The selected PDF is no longer available.' }
      try {
        const parts = await splitPdfBytes(new Uint8Array(data), request.chunkSize)
        for (let index = 0; index < parts.length; index += 1) {
          const bytes = parts[index]!
          const buffer = bytes.buffer.slice(
            bytes.byteOffset,
            bytes.byteOffset + bytes.byteLength,
          ) as ArrayBuffer
          if ((await persist(`${request.baseName}-${index + 1}.pdf`, buffer, 'export-copy')) === false) {
            return { ok: false, error: 'The local PDF export failed.' }
          }
        }
        return { ok: true, savedDir: 'browser-downloads', count: parts.length }
      } catch (error) {
        return { ok: false, error: error instanceof Error ? error.message : String(error) }
      }
    },
    mergePages: async (request) => {
      const data = files.get(request.path)
      if (!data) return { ok: false, error: 'The selected PDF is no longer available.' }
      try {
        const bytes = await mergePagesBytes(new Uint8Array(data), request)
        const buffer = bytes.buffer.slice(
          bytes.byteOffset,
          bytes.byteOffset + bytes.byteLength,
        ) as ArrayBuffer
        if ((await persist(request.suggestedName, buffer, 'export-copy')) === false) {
          return { ok: false, error: 'The local PDF export failed.' }
        }
        return { ok: true, savedPath: request.suggestedName }
      } catch (error) {
        return { ok: false, error: error instanceof Error ? error.message : String(error) }
      }
    },
    splitPages: async (request) => {
      const data = files.get(request.path)
      if (!data) return { ok: false, error: 'The selected PDF is no longer available.' }
      try {
        const bytes = await splitPagesBytes(new Uint8Array(data), request.perPage)
        const buffer = bytes.buffer.slice(
          bytes.byteOffset,
          bytes.byteOffset + bytes.byteLength,
        ) as ArrayBuffer
        if ((await persist(request.suggestedName, buffer, 'export-copy')) === false) {
          return { ok: false, error: 'The local PDF export failed.' }
        }
        return { ok: true, savedPath: request.suggestedName }
      } catch (error) {
        return { ok: false, error: error instanceof Error ? error.message : String(error) }
      }
    },
    setPageSize: async (request) => {
      const data = files.get(request.path)
      if (!data) return { ok: false, error: 'The selected PDF is no longer available.' }
      try {
        const bytes = await setPageSizeBytes(
          new Uint8Array(data),
          request.width,
          request.height,
        )
        const buffer = bytes.buffer.slice(
          bytes.byteOffset,
          bytes.byteOffset + bytes.byteLength,
        ) as ArrayBuffer
        await commitPageMutation(request.path, data, buffer)
        return { ok: true }
      } catch (error) {
        return { ok: false, error: error instanceof Error ? error.message : String(error) }
      }
    },
    cropPages: async (request) => {
      const data = files.get(request.path)
      if (!data) return { ok: false, error: 'The selected PDF is no longer available.' }
      try {
        const bytes = await cropPagesBytes(new Uint8Array(data), request.pages, request.rect)
        const buffer = bytes.buffer.slice(
          bytes.byteOffset,
          bytes.byteOffset + bytes.byteLength,
        ) as ArrayBuffer
        await commitPageMutation(request.path, data, buffer)
        return { ok: true }
      } catch (error) {
        return { ok: false, error: error instanceof Error ? error.message : String(error) }
      }
    },
    replacePages: async (request) => {
      const data = files.get(request.path)
      if (!data) return { ok: false, error: 'The selected PDF is no longer available.' }
      try {
        const picked = options.pickPdfFiles ? (await options.pickPdfFiles())[0] : undefined
        const file = picked ? null : await pickPdfFile()
        if (!picked && !file) return { ok: true, canceled: true }
        const replacement = picked?.data ?? (await file!.arrayBuffer())
        const result = await replacePagesBytes(
          new Uint8Array(data),
          new Uint8Array(replacement),
          request.pages,
        )
        const buffer = result.merged.buffer.slice(
          result.merged.byteOffset,
          result.merged.byteOffset + result.merged.byteLength,
        ) as ArrayBuffer
        await commitPageMutation(request.path, data, buffer)
        return { ok: true, removed: result.removed, inserted: result.inserted }
      } catch (error) {
        return { ok: false, error: error instanceof Error ? error.message : String(error) }
      }
    },
    mergePdf: async (request) => {
      const data = files.get(request.path)
      if (!data) return { ok: false, error: 'The selected PDF is no longer available.' }
      try {
        const picked = options.pickPdfFiles
          ? await options.pickPdfFiles()
          : await Promise.all(
              (await pickPdfFiles()).map(async (file) => ({
                name: file.name,
                data: await file.arrayBuffer(),
              })),
            )
        if (picked.length === 0) return { ok: true, canceled: true }
        const result = await mergePdfBytes(
          new Uint8Array(data),
          picked.map((file) => new Uint8Array(file.data)),
        )
        const buffer = result.merged.buffer.slice(
          result.merged.byteOffset,
          result.merged.byteOffset + result.merged.byteLength,
        ) as ArrayBuffer
        if ((await persist(request.suggestedName, buffer, 'export-copy')) === false) {
          return { ok: false, error: 'The local PDF export failed.' }
        }
        return {
          ok: true,
          savedPath: request.suggestedName,
          appendedCount: result.appended,
        }
      } catch (error) {
        return { ok: false, error: error instanceof Error ? error.message : String(error) }
      }
    },
    documentHistoryState: async (path) => {
      const history = historyFor(path)
      return { canUndo: history.undo.length > 0, canRedo: history.redo.length > 0 }
    },
    undoDocumentMutation: async (path) => {
      const current = files.get(path)
      const history = historyFor(path)
      const previous = history.undo.at(-1)
      if (!current || !previous) return { ok: false, error: 'There is no PDF edit to undo.' }
      try {
        const fileName = decodeURIComponent(path.split('/').pop() ?? 'document.pdf')
        if ((await persist(fileName, previous, 'save')) === false) {
          return { ok: false, error: 'The local PDF save failed.' }
        }
        history.undo.pop()
        history.redo.push(cloneBuffer(current))
        files.set(path, cloneBuffer(previous))
        window.dispatchEvent(new CustomEvent(PDF_OPEN_EVENT, { detail: { path } }))
        dirty = false
        return { ok: true }
      } catch (error) {
        return { ok: false, error: error instanceof Error ? error.message : String(error) }
      }
    },
    redoDocumentMutation: async (path) => {
      const current = files.get(path)
      const history = historyFor(path)
      const next = history.redo.at(-1)
      if (!current || !next) return { ok: false, error: 'There is no PDF edit to redo.' }
      try {
        const fileName = decodeURIComponent(path.split('/').pop() ?? 'document.pdf')
        if ((await persist(fileName, next, 'save')) === false) {
          return { ok: false, error: 'The local PDF save failed.' }
        }
        history.redo.pop()
        history.undo.push(cloneBuffer(current))
        files.set(path, cloneBuffer(next))
        window.dispatchEvent(new CustomEvent(PDF_OPEN_EVENT, { detail: { path } }))
        dirty = false
        return { ok: true }
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
    listSavedSignatures: async () => loadSignatures(),
    addSavedSignature: async (data) => {
      const key = signatureKey(data)
      const retained = loadSignatures().filter((entry) => signatureKey(entry.data) !== key)
      return saveSignatures([
        { id: crypto.randomUUID(), createdAt: Date.now(), data },
        ...retained,
      ])
    },
    removeSavedSignature: async (id) =>
      saveSignatures(loadSignatures().filter((entry) => entry.id !== id)),
    setDirty: (next) => {
      dirty = next
    },
    onCloseSaveRequest: () => () => undefined,
    sendCloseSaveResult: () => undefined,
    onSaveAsRequest: () => () => undefined,
    sendSaveAsResult: () => undefined,
    onSaveAsFlow: () => () => undefined,
    onPrintRequest: () => () => undefined,
    getLanguage: async () => 'zh',
    onLanguageChanged: () => () => undefined,
    getTheme: async () => 'system',
    onThemeChanged: () => () => undefined,
  }

  const execute = async (command: LiveEditorCommand): Promise<LiveEditorExecution> => {
    if (resolvePdfOperationId(command.operation) === 'pdf.document.create_blank') {
      const registered = await executePdfOperation(
        { operation: command.operation, arguments: command.arguments },
        {
          createBlank: async ({ confirmReplace }) => {
            if (!confirmReplace && (activeFile !== null || dirty)) {
              throw new Error(
                'Replacing the current PDF requires explicit user confirmation (confirmReplace: true).',
              )
            }
            const bytes = await blankPdfBytes()
            const data = bytes.buffer.slice(
              bytes.byteOffset,
              bytes.byteOffset + bytes.byteLength,
            ) as ArrayBuffer
            await stageFile('Untitled.pdf', data)
            dirty = false
            return { fileName: 'Untitled.pdf', pageCount: 1 }
          },
        },
      )
      if (registered.handled) {
        return registered.ok
          ? { ok: true, output: { ...registered.output } }
          : { ok: false, error: registered.error, message: registered.message }
      }
    }
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
    if (resolvePdfOperationId(command.operation) === 'pdf.page.replace_staged') {
      const registered = await executePdfOperation(
        { operation: command.operation, arguments: command.arguments },
        {
          replacePagesStaged: ({ data, pages }) => replacePages(data, pages),
        },
      )
      if (registered.handled) {
        return registered.ok
          ? { ok: true, output: { ...registered.output } }
          : { ok: false, error: registered.error, message: registered.message }
      }
    }
    if (resolvePdfOperationId(command.operation) === 'pdf.page.insert_blank') {
      const registered = await executePdfOperation(
        { operation: command.operation, arguments: command.arguments },
        {
          insertBlankPage: async (afterPageIndex) => {
            if (!activeFile) throw new Error('No PDF document is open.')
            const result = await api.insertBlankPage({ path: activeFile, afterPageIndex })
            if (!result.ok) throw new Error(result.error)
          },
        },
      )
      if (registered.handled) {
        return registered.ok
          ? { ok: true, output: { ...registered.output } }
          : { ok: false, error: registered.error, message: registered.message }
      }
    }
    if (resolvePdfOperationId(command.operation) === 'pdf.page.set_size') {
      const registered = await executePdfOperation(
        { operation: command.operation, arguments: command.arguments },
        {
          setPageSize: async (width, height) => {
            if (!activeFile) throw new Error('No PDF document is open.')
            const result = await api.setPageSize({ path: activeFile, width, height })
            if (!result.ok) throw new Error(result.error)
          },
        },
      )
      if (registered.handled) {
        return registered.ok
          ? { ok: true, output: { ...registered.output } }
          : { ok: false, error: registered.error, message: registered.message }
      }
    }
    if (resolvePdfOperationId(command.operation) === 'pdf.page.crop') {
      const registered = await executePdfOperation(
        { operation: command.operation, arguments: command.arguments },
        {
          cropPages: async (pages, rect) => {
            if (!activeFile) throw new Error('No PDF document is open.')
            const result = await api.cropPages({ path: activeFile, pages, rect })
            if (!result.ok) throw new Error(result.error)
          },
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
