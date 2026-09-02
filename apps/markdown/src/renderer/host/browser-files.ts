import { saveLiveEditorFile } from '@tandemfolio/host-bridge'

export interface LoadedMarkdown {
  fileName: string
  text: string
  assetBasePath?: string
  assetFiles?: ReadonlyMap<string, File>
}

const MAX_BROWSER_ASSET_BYTES = 20 * 1024 * 1024

function normalizeSelectedAssetPath(basePath: string, authoredPath: string): string | null {
  if (/^[a-z][a-z0-9+.-]*:/i.test(authoredPath) || authoredPath.startsWith('/')) return null
  let decoded: string
  try {
    decoded = decodeURIComponent(authoredPath.split(/[?#]/, 1)[0] ?? '')
  } catch {
    return null
  }
  const segments: string[] = []
  for (const segment of `${basePath}/${decoded}`.replace(/\\/g, '/').split('/')) {
    if (!segment || segment === '.') continue
    if (segment === '..') {
      if (segments.length === 0) return null
      segments.pop()
    } else {
      segments.push(segment)
    }
  }
  return segments.join('/')
}

function browserImageMime(
  file: File,
  bytes: Uint8Array,
): 'image/png' | 'image/jpeg' | 'image/gif' | null {
  const extension = file.name.split('.').pop()?.toLowerCase()
  if (
    extension === 'png' &&
    bytes.length >= 8 &&
    bytes[0] === 0x89 &&
    bytes[1] === 0x50 &&
    bytes[2] === 0x4e &&
    bytes[3] === 0x47
  ) {
    return 'image/png'
  }
  if (
    (extension === 'jpg' || extension === 'jpeg') &&
    bytes.length >= 3 &&
    bytes[0] === 0xff &&
    bytes[1] === 0xd8 &&
    bytes[2] === 0xff
  ) {
    return 'image/jpeg'
  }
  const gif = String.fromCharCode(...bytes.subarray(0, 6))
  return extension === 'gif' && (gif === 'GIF87a' || gif === 'GIF89a') ? 'image/gif' : null
}

export async function readLoadedMarkdownAsset(
  loaded: LoadedMarkdown,
  authoredPath: string,
): Promise<{
  readonly mime: 'image/png' | 'image/jpeg' | 'image/gif'
  readonly data: ArrayBuffer
} | null> {
  const path = normalizeSelectedAssetPath(loaded.assetBasePath ?? '', authoredPath)
  const file = path ? loaded.assetFiles?.get(path) : undefined
  if (!file || file.size < 1 || file.size > MAX_BROWSER_ASSET_BYTES) return null
  const bytes = new Uint8Array(await file.arrayBuffer())
  const mime = browserImageMime(file, bytes)
  if (!mime) return null
  return {
    mime,
    data: bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer,
  }
}

let activeHandle: FileSystemFileHandle | null = null

interface MarkdownFileCandidate {
  readonly relativePath: string
  readonly file: File
  readonly handle: FileSystemFileHandle | null
}

function pickerTypes() {
  return [
    {
      description: 'Markdown',
      accept: { 'text/markdown': ['.md', '.markdown'] },
    },
  ]
}

function chooseMarkdownCandidate(
  candidates: readonly MarkdownFileCandidate[],
): MarkdownFileCandidate | null {
  if (candidates.length === 0) return null
  if (candidates.length === 1) return candidates[0]!
  const selection = window.prompt(
    `Choose a Markdown file:\n${candidates
      .map((candidate, index) => `${index + 1}. ${candidate.relativePath}`)
      .join('\n')}`,
    '1',
  )
  if (selection === null) return null
  const index = Number(selection) - 1
  return Number.isInteger(index) && candidates[index] ? candidates[index]! : null
}

function loadedMarkdownFromFiles(
  files: readonly MarkdownFileCandidate[],
): Promise<LoadedMarkdown | null> {
  const selected = chooseMarkdownCandidate(
    files.filter((candidate) => /\.(md|markdown)$/i.test(candidate.file.name)),
  )
  if (!selected) return Promise.resolve(null)
  activeHandle = selected.handle
  const slash = selected.relativePath.lastIndexOf('/')
  const assetBasePath = slash < 0 ? '' : selected.relativePath.slice(0, slash)
  const assetFiles = new Map(
    files
      .filter((candidate) => /\.(png|jpe?g|gif)$/i.test(candidate.file.name))
      .map((candidate) => [candidate.relativePath, candidate.file] as const),
  )
  return selected.file.text().then((text) => ({
    fileName: selected.file.name,
    text,
    assetBasePath,
    assetFiles,
  }))
}

async function collectDirectoryFiles(
  directory: FileSystemDirectoryHandle,
  prefix = '',
  depth = 0,
  output: MarkdownFileCandidate[] = [],
): Promise<MarkdownFileCandidate[]> {
  if (depth > 8 || output.length >= 2_000) return output
  for await (const entry of directory.values()) {
    if (output.length >= 2_000) break
    const relativePath = prefix ? `${prefix}/${entry.name}` : entry.name
    if (entry.kind === 'directory') {
      await collectDirectoryFiles(
        entry as FileSystemDirectoryHandle,
        relativePath,
        depth + 1,
        output,
      )
    } else {
      const handle = entry as FileSystemFileHandle
      output.push({ relativePath, file: await handle.getFile(), handle })
    }
  }
  return output
}

async function chooseWithInput(): Promise<LoadedMarkdown | null> {
  return new Promise((resolve) => {
    const input = document.createElement('input')
    input.type = 'file'
    input.multiple = true
    input.setAttribute('webkitdirectory', '')
    input.accept = '.md,.markdown,.png,.jpg,.jpeg,.gif,text/markdown,text/plain,image/*'
    input.addEventListener(
      'change',
      () => {
        const files = Array.from(input.files ?? [])
        if (files.length === 0) {
          resolve(null)
          return
        }
        void loadedMarkdownFromFiles(
          files.map((file) => ({
            relativePath: file.webkitRelativePath || file.name,
            file,
            handle: null,
          })),
        ).then(resolve)
      },
      { once: true },
    )
    input.click()
  })
}

export async function openMarkdownFile(): Promise<LoadedMarkdown | null> {
  if (window.showDirectoryPicker) {
    try {
      const directory = await window.showDirectoryPicker()
      return loadedMarkdownFromFiles(await collectDirectoryFiles(directory))
    } catch (error) {
      if (error instanceof DOMException && error.name === 'AbortError') return null
      throw error
    }
  }
  if (window.showOpenFilePicker) {
    try {
      const [handle] = await window.showOpenFilePicker({ multiple: false, types: pickerTypes() })
      if (!handle) return null
      activeHandle = handle
      const file = await handle.getFile()
      return { fileName: file.name, text: await file.text() }
    } catch (error) {
      if (error instanceof DOMException && error.name === 'AbortError') return null
      throw error
    }
  }
  activeHandle = null
  return chooseWithInput()
}

export function detachMarkdownFileHandle(): void {
  activeHandle = null
}

export function canOverwriteMarkdownFile(): boolean {
  return activeHandle !== null
}

function downloadBlob(blob: Blob, fileName: string): void {
  const url = URL.createObjectURL(blob)
  const anchor = document.createElement('a')
  anchor.href = url
  anchor.download = fileName
  anchor.click()
  queueMicrotask(() => URL.revokeObjectURL(url))
}

export async function saveMarkdownFile(
  text: string,
  suggestedName: string,
  saveAs = false,
): Promise<{ ok: true; fileName: string } | { ok: false }> {
  let handle = saveAs ? null : activeHandle
  if (window.parent !== window) {
    const bytes = new TextEncoder().encode(text)
    const persisted = await saveLiveEditorFile({
      fileName: suggestedName,
      data: bytes.buffer.slice(
        bytes.byteOffset,
        bytes.byteOffset + bytes.byteLength,
      ) as ArrayBuffer,
      mode: saveAs ? 'save-as' : 'save',
    })
    return persisted.ok ? { ok: true, fileName: suggestedName } : { ok: false }
  }
  if (!handle && window.showSaveFilePicker) {
    try {
      handle = await window.showSaveFilePicker({
        suggestedName,
        types: pickerTypes(),
      })
    } catch (error) {
      if (error instanceof DOMException && error.name === 'AbortError') return { ok: false }
      throw error
    }
  }
  if (handle) {
    const writable = await handle.createWritable()
    await writable.write(text)
    await writable.close()
    activeHandle = handle
    return { ok: true, fileName: handle.name }
  }
  downloadBlob(new Blob([text], { type: 'text/markdown;charset=utf-8' }), suggestedName)
  return { ok: true, fileName: suggestedName }
}

export function downloadMarkdownExport(data: BlobPart, fileName: string, type: string): void {
  downloadBlob(new Blob([data], { type }), fileName)
}

export async function pickImageDataUrl(): Promise<{ src: string; alt: string } | null> {
  return new Promise((resolve) => {
    const input = document.createElement('input')
    input.type = 'file'
    input.accept = 'image/png,image/jpeg,image/gif,image/webp,image/svg+xml'
    input.addEventListener(
      'change',
      () => {
        const file = input.files?.[0]
        if (!file) {
          resolve(null)
          return
        }
        const reader = new FileReader()
        reader.addEventListener(
          'load',
          () =>
            resolve({
              src: String(reader.result),
              alt: file.name.replace(/\.[a-z0-9]+$/i, ''),
            }),
          { once: true },
        )
        reader.readAsDataURL(file)
      },
      { once: true },
    )
    input.click()
  })
}
