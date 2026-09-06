import { saveLiveEditorFile } from '@tandemfolio/host-bridge'

export interface LoadedMarkdown {
  fileName: string
  text: string
  assetBasePath?: string
  assetFiles?: ReadonlyMap<string, File>
}

export interface MarkdownSaveCompanion {
  readonly relativePath: string
  readonly data: ArrayBuffer
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
let activeDirectory: FileSystemDirectoryHandle | null = null
let ownedCompanions = new Map<string, ArrayBuffer>()

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
      const loaded = await loadedMarkdownFromFiles(await collectDirectoryFiles(directory))
      if (loaded) {
        activeDirectory = directory
        ownedCompanions = new Map()
      }
      return loaded
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
      activeDirectory = null
      ownedCompanions = new Map()
      const file = await handle.getFile()
      return { fileName: file.name, text: await file.text() }
    } catch (error) {
      if (error instanceof DOMException && error.name === 'AbortError') return null
      throw error
    }
  }
  activeHandle = null
  activeDirectory = null
  ownedCompanions = new Map()
  return chooseWithInput()
}

export function detachMarkdownFileHandle(): void {
  activeHandle = null
  activeDirectory = null
  ownedCompanions = new Map()
}

export function canOverwriteMarkdownFile(): boolean {
  return activeHandle !== null
}

export function canPersistMarkdownCompanions(saveAs: boolean): boolean {
  if (window.parent !== window) return true
  return saveAs ? Boolean(window.showDirectoryPicker) : activeDirectory !== null
}

async function directoryForRelativePath(
  root: FileSystemDirectoryHandle,
  relativePath: string,
): Promise<{ directory: FileSystemDirectoryHandle; name: string }> {
  const segments = relativePath.split('/')
  if (
    segments.length < 2 ||
    segments.some((segment) => !segment || segment === '.' || segment === '..')
  ) {
    throw new Error('Markdown companion paths must be safe relative paths.')
  }
  let directory = root
  for (const segment of segments.slice(0, -1)) {
    directory = await directory.getDirectoryHandle(segment, { create: true })
  }
  return { directory, name: segments.at(-1)! }
}

async function equalFileBytes(file: File, data: ArrayBuffer): Promise<boolean> {
  if (file.size !== data.byteLength) return false
  const left = new Uint8Array(await file.arrayBuffer())
  const right = new Uint8Array(data)
  return left.every((byte, index) => byte === right[index])
}

async function writeDirectoryBundle(
  directory: FileSystemDirectoryHandle,
  documentHandle: FileSystemFileHandle,
  text: string,
  companionFiles: readonly MarkdownSaveCompanion[],
  removedCompanions: ReadonlyMap<string, ArrayBuffer>,
): Promise<void> {
  const created: Array<{ directory: FileSystemDirectoryHandle; name: string }> = []
  try {
    for (const companion of companionFiles) {
      const target = await directoryForRelativePath(directory, companion.relativePath)
      let handle: FileSystemFileHandle
      try {
        handle = await target.directory.getFileHandle(target.name)
        if (!(await equalFileBytes(await handle.getFile(), companion.data))) {
          throw new Error(
            `Markdown companion target already contains different bytes: ${companion.relativePath}`,
          )
        }
        continue
      } catch (error) {
        if (!(error instanceof DOMException) || error.name !== 'NotFoundError') throw error
        handle = await target.directory.getFileHandle(target.name, { create: true })
        created.push(target)
      }
      const writable = await handle.createWritable()
      await writable.write(companion.data)
      await writable.close()
    }
    const writable = await documentHandle.createWritable()
    await writable.write(text)
    await writable.close()
    for (const [relativePath, previousData] of removedCompanions) {
      try {
        const target = await directoryForRelativePath(directory, relativePath)
        const handle = await target.directory.getFileHandle(target.name)
        if (await equalFileBytes(await handle.getFile(), previousData)) {
          await target.directory.removeEntry(target.name)
        }
      } catch {
        // Preserve missing or user-modified files and relinquish ownership.
      }
    }
  } catch (error) {
    await Promise.all(
      created.map((target) => target.directory.removeEntry(target.name).catch(() => undefined)),
    )
    throw error
  }
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
  companionFiles: readonly MarkdownSaveCompanion[] = [],
): Promise<{ ok: true; fileName: string } | { ok: false }> {
  let handle = saveAs ? null : activeHandle
  const currentCompanions = new Map(
    companionFiles.map((companion) => [companion.relativePath, companion.data] as const),
  )
  const removedCompanions = saveAs
    ? new Map<string, ArrayBuffer>()
    : new Map([...ownedCompanions].filter(([path]) => !currentCompanions.has(path)))
  if (window.parent !== window) {
    const bytes = new TextEncoder().encode(text)
    const persisted = await saveLiveEditorFile({
      fileName: suggestedName,
      data: bytes.buffer.slice(
        bytes.byteOffset,
        bytes.byteOffset + bytes.byteLength,
      ) as ArrayBuffer,
      mode: saveAs ? 'save-as' : 'save',
      ...(companionFiles.length ? { companionFiles } : {}),
      ...(removedCompanions.size ? { removeCompanionPaths: [...removedCompanions.keys()] } : {}),
    })
    if (!persisted.ok) return { ok: false }
    ownedCompanions = currentCompanions
    return { ok: true, fileName: suggestedName }
  }
  if (companionFiles.length > 0 || removedCompanions.size > 0) {
    let directory = saveAs ? null : activeDirectory
    if (!directory && window.showDirectoryPicker) {
      try {
        directory = await window.showDirectoryPicker()
      } catch (error) {
        if (error instanceof DOMException && error.name === 'AbortError') return { ok: false }
        throw error
      }
    }
    if (!directory) return { ok: false }
    const documentHandle = saveAs
      ? await directory.getFileHandle(suggestedName, { create: true })
      : activeHandle
    if (!documentHandle) return { ok: false }
    await writeDirectoryBundle(directory, documentHandle, text, companionFiles, removedCompanions)
    activeDirectory = directory
    activeHandle = documentHandle
    ownedCompanions = currentCompanions
    return { ok: true, fileName: documentHandle.name }
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
    if (saveAs) ownedCompanions = new Map()
    return { ok: true, fileName: handle.name }
  }
  downloadBlob(new Blob([text], { type: 'text/markdown;charset=utf-8' }), suggestedName)
  if (saveAs) ownedCompanions = new Map()
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
