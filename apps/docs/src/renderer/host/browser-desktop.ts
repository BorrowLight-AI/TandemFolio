import type {
  DesktopApi,
  MenuCommand,
  OpenFileResult,
  PickImageResult,
} from '../../shared/lite-api'
import { saveLiveEditorFile } from '@tandemfolio/host-bridge'

const DOCX_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'

type PickerWindow = Window &
  typeof globalThis & {
    showOpenFilePicker?: (options: Record<string, unknown>) => Promise<FileSystemFileHandle[]>
    showSaveFilePicker?: (options: Record<string, unknown>) => Promise<FileSystemFileHandle>
  }

const handles = new Map<string, FileSystemFileHandle>()
const recent: string[] = []

function noopSubscription(): () => void {
  return () => undefined
}

async function sha256(data: ArrayBuffer): Promise<string> {
  const digest = await crypto.subtle.digest('SHA-256', data)
  return Array.from(new Uint8Array(digest), (byte) => byte.toString(16).padStart(2, '0')).join('')
}

async function openWithInput(accept: string): Promise<File | null> {
  return await new Promise((resolve) => {
    const input = document.createElement('input')
    input.type = 'file'
    input.accept = accept
    input.addEventListener('change', () => resolve(input.files?.[0] ?? null), { once: true })
    input.click()
  })
}

async function toOpenResult(file: File, handle?: FileSystemFileHandle): Promise<OpenFileResult> {
  const data = await file.arrayBuffer()
  const path = `browser:${file.name}:${file.lastModified}`
  if (handle) handles.set(path, handle)
  if (!recent.includes(path)) recent.unshift(path)
  return { path, name: file.name, data, hash: await sha256(data) }
}

async function openDocx(): Promise<OpenFileResult | null> {
  const picker = window as PickerWindow
  try {
    if (picker.showOpenFilePicker) {
      const [handle] = await picker.showOpenFilePicker({
        multiple: false,
        types: [
          {
            description: 'Word document',
            accept: {
              'application/vnd.openxmlformats-officedocument.wordprocessingml.document': ['.docx'],
            },
          },
        ],
      })
      return handle ? await toOpenResult(await handle.getFile(), handle) : null
    }
    const file = await openWithInput('.docx')
    return file ? await toOpenResult(file) : null
  } catch (error) {
    if (error instanceof DOMException && error.name === 'AbortError') return null
    throw error
  }
}

function download(name: string, data: ArrayBuffer): void {
  const url = URL.createObjectURL(
    new Blob([data], {
      type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    }),
  )
  const anchor = document.createElement('a')
  anchor.href = url
  anchor.download = name.endsWith('.docx') ? name : `${name}.docx`
  anchor.click()
  setTimeout(() => URL.revokeObjectURL(url), 0)
}

async function writeHandle(handle: FileSystemFileHandle, data: ArrayBuffer): Promise<void> {
  const writable = await handle.createWritable()
  await writable.write(data)
  await writable.close()
}

async function saveAs(defaultName: string, data: ArrayBuffer) {
  const picker = window as PickerWindow
  const fileName = defaultName.endsWith('.docx') ? defaultName : `${defaultName}.docx`
  try {
    if (window.parent === window && picker.showSaveFilePicker) {
      const handle = await picker.showSaveFilePicker({
        suggestedName: fileName,
        types: [{ description: 'Word document', accept: { [DOCX_MIME]: ['.docx'] } }],
      })
      await writeHandle(handle, data)
      const file = await handle.getFile()
      const path = `browser:${file.name}:${file.lastModified}`
      handles.set(path, handle)
      return { ok: true, path }
    }
    if (window.parent !== window) {
      return saveLiveEditorFile({ fileName, data, mode: 'save-as' })
    }
    download(fileName, data)
    return { ok: true, path: `download:${fileName}` }
  } catch (error) {
    if (error instanceof DOMException && error.name === 'AbortError')
      return { ok: false, error: 'cancelled' }
    return { ok: false, error: String(error) }
  }
}

async function pickImage(): Promise<PickImageResult | null> {
  const file = await openWithInput('image/png,image/jpeg,image/gif')
  if (!file) return null
  const mime = file.type as PickImageResult['mime']
  if (!['image/png', 'image/jpeg', 'image/gif'].includes(mime)) return null
  const bytes = new Uint8Array(await file.arrayBuffer())
  let binary = ''
  for (const byte of bytes) binary += String.fromCharCode(byte)
  return { base64: btoa(binary), mime, name: file.name }
}

export function installBrowserDesktop(): void {
  if (window.desktop) return
  const api: DesktopApi = {
    getLanguage: async () => 'zh',
    onLanguageChanged: noopSubscription,
    getTheme: async () => 'system',
    onThemeChanged: noopSubscription,
    openDocx,
    openDocxPath: async () => null,
    consumePendingOpenDocx: async () => null,
    consumeNewBlankDoc: async () => true,
    onOpenDocx: noopSubscription,
    onRenamedDocx: noopSubscription,
    saveDocx: async (path, data) => {
      try {
        const handle = handles.get(path)
        if (handle) await writeHandle(handle, data)
        else {
          const fileName = path.split(/[\\/]/).pop() || 'document.docx'
          if (window.parent !== window) {
            const persisted = await saveLiveEditorFile({
              fileName,
              data,
              mode: 'save',
            })
            if (!persisted.ok) return persisted
          } else download(fileName, data)
        }
        return { ok: true }
      } catch (error) {
        return { ok: false, error: String(error) }
      }
    },
    writeRecoveryCopy: async () => ({ ok: false }),
    onTeardown: noopSubscription,
    saveDocxAs: saveAs,
    saveDocxNew: saveAs,
    getRecentFiles: async () => [...recent],
    pickImage,
    fontMetrics: async () => null,
    print: async () => window.print(),
    exportPdf: async () => ({
      ok: false,
      error: 'PDF export is not available in the browser host yet.',
    }),
    printPdfBuffer: async () => ({
      ok: false,
      error: 'PDF export is not available in the browser host yet.',
    }),
    saveMergedPdf: async () => ({
      ok: false,
      error: 'PDF export is not available in the browser host yet.',
    }),
    openNewTab: async () => void window.open(window.location.href, '_blank', 'noopener'),
    listDocsTabs: async () => [{ id: 'browser', title: document.title, focused: true }],
    focusDocsTab: async () => window.focus(),
    onMenuCommand: (handler) => {
      const listener = (event: Event) => {
        const detail = (event as CustomEvent<{ command: MenuCommand; payload?: string }>).detail
        handler(detail.command, detail.payload)
      }
      window.addEventListener('tandemfolio-menu', listener)
      return () => window.removeEventListener('tandemfolio-menu', listener)
    },
    onCloseCheck: noopSubscription,
    reportCloseCheck: () => undefined,
    onCloseSaveRequest: noopSubscription,
    reportCloseSaveResult: () => undefined,
    reportViewMenuState: () => undefined,
  }
  window.desktop = api
}
