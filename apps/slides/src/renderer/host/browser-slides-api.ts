import type { RenderSlide } from '@genoffice/pptx-render'
import type {
  LiveEditorAdapter,
  LiveEditorCommand,
  LiveEditorExecution,
  LiveEditorSnapshot,
} from '@tandemfolio/host-bridge'
import { saveLiveEditorFile } from '@tandemfolio/host-bridge'
import { BrowserPresentation } from './browser-presentation'
import type {
  AnimationItem,
  AudienceNavAction,
  GetLayoutsResult,
  OpenResult,
  SectionInfo,
  ShapeKey,
  ShowInkEvent,
  ShowSyncState,
  SlideComment,
  SlidesApi,
  TransitionKind,
} from '../../shared/ipc'
import { executePptxOperation } from '../operations/registry'
import { THEME_PRESETS } from '../themes'

const PPTX_MIME = 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
const CHART_COLOR_SCHEMES = [
  { key: 'default', label: 'Default', colors: [] },
  {
    key: 'blue',
    label: 'Blue',
    colors: ['#2E75B6', '#4472C4', '#5B9BD5', '#70AD47', '#ED7D31'],
  },
  { key: 'warm', label: 'Warm', colors: ['#ED7D31', '#FFC000', '#FF0000', '#C55A11', '#833C00'] },
  { key: 'cool', label: 'Cool', colors: ['#0070C0', '#00B0F0', '#00B0A0', '#7030A0', '#2E75B6'] },
  {
    key: 'mono',
    label: 'Monochrome',
    colors: ['#404040', '#666666', '#888888', '#AAAAAA', '#CCCCCC'],
  },
] as const
const IMAGE_EXTENSIONS = ['png', 'jpg', 'jpeg', 'gif', 'bmp', 'webp', 'tif', 'tiff'] as const
type ImageExtension = (typeof IMAGE_EXTENSIONS)[number]

function isImageExtension(value: string): value is ImageExtension {
  return (IMAGE_EXTENSIONS as readonly string[]).includes(value.toLowerCase())
}
const MEDIA_EXTENSIONS = [
  'mp4',
  'm4v',
  'mov',
  'webm',
  'avi',
  'mp3',
  'wav',
  'm4a',
  'aac',
  'ogg',
] as const
type MediaExtension = (typeof MEDIA_EXTENSIONS)[number]

function isMediaExtension(value: string): value is MediaExtension {
  return (MEDIA_EXTENSIONS as readonly string[]).includes(value.toLowerCase())
}

async function chooseMedia(kind: 'video' | 'audio'): Promise<File | null> {
  return new Promise<File | null>((resolve) => {
    const input = document.createElement('input')
    input.type = 'file'
    input.accept = kind === 'video' ? '.mp4,.m4v,.mov,.webm,.avi' : '.mp3,.wav,.m4a,.aac,.ogg'
    input.hidden = true
    input.addEventListener(
      'change',
      () => {
        const file = input.files?.[0] ?? null
        input.remove()
        resolve(file)
      },
      { once: true },
    )
    document.body.append(input)
    input.click()
  })
}

async function chooseModel3d(): Promise<File | null> {
  return new Promise<File | null>((resolve) => {
    const input = document.createElement('input')
    input.type = 'file'
    input.accept = '.glb,.gltf'
    input.hidden = true
    input.addEventListener(
      'change',
      () => {
        const file = input.files?.[0] ?? null
        input.remove()
        resolve(file)
      },
      { once: true },
    )
    document.body.append(input)
    input.click()
  })
}

async function chooseImage(): Promise<File | null> {
  return new Promise<File | null>((resolve) => {
    const input = document.createElement('input')
    input.type = 'file'
    input.accept = IMAGE_EXTENSIONS.map((extension) => `.${extension}`).join(',')
    input.hidden = true
    input.addEventListener(
      'change',
      () => {
        const file = input.files?.[0] ?? null
        input.remove()
        resolve(file)
      },
      { once: true },
    )
    document.body.append(input)
    input.click()
  })
}

async function fileBase64(file: File): Promise<string> {
  const bytes = new Uint8Array(await file.arrayBuffer())
  let binary = ''
  for (let offset = 0; offset < bytes.length; offset += 32_768) {
    binary += String.fromCharCode(...bytes.subarray(offset, offset + 32_768))
  }
  return btoa(binary)
}

async function imageNaturalSize(file: File): Promise<{ width: number; height: number }> {
  if (typeof createImageBitmap !== 'function') return { width: 4, height: 3 }
  try {
    const bitmap = await createImageBitmap(file)
    const size = { width: bitmap.width, height: bitmap.height }
    bitmap.close()
    return size.width > 0 && size.height > 0 ? size : { width: 4, height: 3 }
  } catch {
    return { width: 4, height: 3 }
  }
}

type FilePickerWindow = Window & {
  showOpenFilePicker?: (options: {
    multiple: boolean
    types: Array<{ description: string; accept: Record<string, string[]> }>
  }) => Promise<Array<{ getFile(): Promise<File> }>>
  showSaveFilePicker?: (options: {
    suggestedName: string
    types: Array<{ description: string; accept: Record<string, string[]> }>
  }) => Promise<FileSystemFileHandle>
}

function toArrayBuffer(bytes: Uint8Array): ArrayBuffer {
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer
}

function downloadBlob(fileName: string, blob: Blob): void {
  const url = URL.createObjectURL(blob)
  const anchor = document.createElement('a')
  anchor.href = url
  anchor.download = fileName
  anchor.click()
  queueMicrotask(() => URL.revokeObjectURL(url))
}

function download(fileName: string, bytes: Uint8Array): void {
  downloadBlob(fileName, new Blob([toArrayBuffer(bytes)], { type: PPTX_MIME }))
}

function base64Bytes(data: string): Uint8Array {
  const binary = atob(data)
  const bytes = new Uint8Array(binary.length)
  for (let index = 0; index < binary.length; index += 1) bytes[index] = binary.charCodeAt(index)
  return bytes
}

function escapeHtml(value: string): string {
  return value
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#39;')
}

function printableSlidesHtml(input: {
  readonly pngsBase64: readonly string[]
  readonly widthPx: number
  readonly heightPx: number
  readonly layout?: 'full' | 'handout2' | 'handout3' | 'handout6' | 'notes'
  readonly notes?: readonly string[]
}): string {
  const layout = input.layout ?? 'full'
  const ratio = `${Math.max(1, input.widthPx)} / ${Math.max(1, input.heightPx)}`
  const cards = input.pngsBase64.map(
    (png, index) =>
      `<article class="card"><img alt="Slide ${index + 1}" src="data:image/png;base64,${png}">${
        layout === 'notes' ? `<pre>${escapeHtml(input.notes?.[index] ?? '')}</pre>` : ''
      }</article>`,
  )
  const perPage =
    layout === 'handout2' ? 2 : layout === 'handout3' ? 3 : layout === 'handout6' ? 6 : 1
  const sheets: string[] = []
  for (let index = 0; index < cards.length; index += perPage) {
    sheets.push(`<section class="sheet">${cards.slice(index, index + perPage).join('')}</section>`)
  }
  const columns = layout === 'handout6' ? 2 : 1
  return `<!doctype html><html><head><meta charset="utf-8"><title>Print slides</title><style>
@page{size:landscape;margin:8mm}*{box-sizing:border-box}html,body{margin:0;font-family:system-ui,sans-serif}.sheet{break-after:page;min-height:95vh;display:grid;grid-template-columns:repeat(${columns},minmax(0,1fr));gap:8mm;align-content:center}.sheet:last-child{break-after:auto}.card{break-inside:avoid;display:grid;gap:5mm;align-content:center}.card img{display:block;width:100%;aspect-ratio:${ratio};object-fit:contain}.card pre{font:14px/1.5 system-ui,sans-serif;white-space:pre-wrap;margin:0;padding:0 8mm}.sheet:has(.card:only-child){grid-template-columns:1fr}.sheet:has(.card:only-child) .card img{max-height:88vh}
</style></head><body>${sheets.join('')}</body></html>`
}

type BrowserDirectoryHandle = {
  readonly name: string
  getFileHandle(
    name: string,
    options: { create: true },
  ): Promise<{
    createWritable(): Promise<{ write(blob: Blob): Promise<void>; close(): Promise<void> }>
  }>
}

type DirectoryPickerWindow = Window & {
  showDirectoryPicker?: () => Promise<BrowserDirectoryHandle>
}

async function choosePptx(): Promise<File | null> {
  const picker = (window as FilePickerWindow).showOpenFilePicker
  // File System Access pickers are not consistently exposed/forwarded from MCP srcdoc iframes.
  // Use the standards-based file input there; standalone browser tabs can use a persistent handle.
  if (window.parent === window && picker) {
    try {
      const [handle] = await picker({
        multiple: false,
        types: [{ description: 'PowerPoint presentation', accept: { [PPTX_MIME]: ['.pptx'] } }],
      })
      return handle ? await handle.getFile() : null
    } catch (error) {
      if (error instanceof DOMException && error.name === 'AbortError') return null
      throw error
    }
  }

  return new Promise<File | null>((resolve) => {
    const input = document.createElement('input')
    input.type = 'file'
    input.accept = '.pptx'
    input.hidden = true
    const finish = (file: File | null) => {
      input.remove()
      resolve(file)
    }
    input.addEventListener('change', () => finish(input.files?.[0] ?? null), { once: true })
    document.body.append(input)
    input.click()
  })
}

export interface BrowserSlidesHost {
  api: SlidesApi
  adapter: LiveEditorAdapter
}

type PresenterMessage =
  | {
      readonly type: 'snapshot'
      readonly slides: RenderSlide[]
      readonly animations: AnimationItem[][]
      readonly transitions: TransitionKind[]
      readonly shapeKeys: ShapeKey[][]
      readonly sync: ShowSyncState | null
    }
  | { readonly type: 'sync'; readonly state: ShowSyncState }
  | { readonly type: 'ink'; readonly event: ShowInkEvent }
  | { readonly type: 'nav'; readonly action: AudienceNavAction }
  | { readonly type: 'ready' }

export function createBrowserSlidesHost(): BrowserSlidesHost {
  let presentation: BrowserPresentation | null = null
  let activeFileName = ''
  let activeSaveHandle: FileSystemFileHandle | null = null
  let latestClipboardKind: 'slide' | 'internal' | null = null
  let activeMasterPartPath: string | null = null
  let latestContext: Record<string, unknown> | null = null
  let recoveryVersion = 0
  let fitWidthPx = 1200
  let exportDirectory: BrowserDirectoryHandle | null = null
  let presenterChannel: BroadcastChannel | null = null
  let audienceWindow: Window | null = null
  let latestShowSync: ShowSyncState | null = null
  let audienceSlides: RenderSlide[] | null = null
  let audienceAnimations: AnimationItem[][] = []
  let audienceTransitions: TransitionKind[] = []
  let audienceShapeKeys: ShapeKey[][] = []
  const openedListeners = new Set<(result: OpenResult) => void>()
  const historyListeners = new Set<(state: { canUndo: boolean; canRedo: boolean }) => void>()
  const showSyncListeners = new Set<(state: ShowSyncState) => void>()
  const showInkListeners = new Set<(event: ShowInkEvent) => void>()
  const audienceNavListeners = new Set<(action: AudienceNavAction) => void>()

  const historyState = () => ({
    canUndo: presentation?.canUndo ?? false,
    canRedo: presentation?.canRedo ?? false,
  })
  const publishHistory = () => {
    if (presentation?.dirty) recoveryVersion += 1
    const state = historyState()
    for (const listener of historyListeners) listener(state)
  }
  const result = (width: number): OpenResult => {
    if (!presentation) throw new Error('No presentation is open.')
    return {
      path: activeFileName,
      slides: presentation.renderAll(width),
      size: presentation.size,
    }
  }
  const refreshContext = async (): Promise<void> => {
    latestContext = presentation ? await presentation.context() : null
  }
  const setPresentation = async (next: BrowserPresentation, width: number): Promise<OpenResult> => {
    presentation = next
    activeFileName = next.name
    activeSaveHandle = null
    fitWidthPx = width
    await refreshContext()
    const opened = result(width)
    publishHistory()
    for (const listener of openedListeners) listener(opened)
    return opened
  }
  const openFile = async (file: File, width: number): Promise<OpenResult> => {
    if (!file.name.toLowerCase().endsWith('.pptx')) throw new Error('请选择 .pptx 文件。')
    return setPresentation(
      await BrowserPresentation.open(file.name, await file.arrayBuffer()),
      width,
    )
  }
  const createBlankPresentation = async (width: number): Promise<OpenResult> =>
    setPresentation(await BrowserPresentation.blank(), width)
  const saveCurrent = async (
    fileName?: string,
    forceSaveAs = false,
  ): Promise<{ ok: true; path: string } | { ok: false; error?: string }> => {
    if (!presentation) return { ok: false, error: 'No presentation is open.' }
    const name = fileName || activeFileName || 'Untitled.pptx'
    try {
      const bytes = await presentation.save()
      let persistedPath: string | undefined
      let targetHandle = forceSaveAs ? null : activeSaveHandle
      if (
        !targetHandle &&
        window.parent === window &&
        (window as FilePickerWindow).showSaveFilePicker
      ) {
        targetHandle = await (window as FilePickerWindow).showSaveFilePicker!({
          suggestedName: name,
          types: [{ description: 'PowerPoint presentation', accept: { [PPTX_MIME]: ['.pptx'] } }],
        })
      }
      if (targetHandle) {
        const writable = await targetHandle.createWritable()
        await writable.write(toArrayBuffer(bytes))
        await writable.close()
        activeSaveHandle = targetHandle
        activeFileName = targetHandle.name
      } else if (window.parent !== window) {
        const persisted = await saveLiveEditorFile({
          fileName: name,
          data: toArrayBuffer(bytes),
          mode: forceSaveAs ? 'save-as' : 'save',
        })
        if (!persisted.ok) return persisted
        persistedPath = persisted.path
        activeFileName = name
      } else {
        download(name, bytes)
        activeFileName = name
      }
      publishHistory()
      return { ok: true, path: targetHandle?.name ?? persistedPath ?? name }
    } catch (error) {
      if (error instanceof DOMException && error.name === 'AbortError') return { ok: false }
      return { ok: false, error: error instanceof Error ? error.message : String(error) }
    }
  }
  const addBlankSlideAt = async (sourceIndex: number, width: number) => {
    if (!presentation) throw new Error('Open a PPTX presentation first.')
    const index = await presentation.addBlankSlide(sourceIndex)
    fitWidthPx = width
    await refreshContext()
    publishHistory()
    return { slides: presentation.renderAll(width), index }
  }
  const duplicateSlideAt = async (sourceIndex: number, clearText: boolean, width: number) => {
    if (!presentation) throw new Error('Open a PPTX presentation first.')
    const index = await presentation.duplicateActiveSlide(sourceIndex, clearText)
    fitWidthPx = width
    await refreshContext()
    publishHistory()
    return { slides: presentation.renderAll(width), index }
  }
  const deleteSlideAt = async (slideIndex: number) => {
    if (!presentation) throw new Error('Open a PPTX presentation first.')
    const index = await presentation.deleteActiveSlide(slideIndex)
    await refreshContext()
    publishHistory()
    return { slides: presentation.renderAll(fitWidthPx), index }
  }
  const moveSlideTo = async (fromIndex: number, toIndex: number) => {
    if (!presentation) throw new Error('Open a PPTX presentation first.')
    const index = await presentation.moveSlide(fromIndex, toIndex)
    await refreshContext()
    publishHistory()
    return {
      slides: presentation.renderAll(fitWidthPx),
      sections: await presentation.sections(),
      index,
    }
  }
  const restoreHistory = async (direction: 'undo' | 'redo') => {
    if (!presentation || !(await presentation[direction]())) return null
    await refreshContext()
    publishHistory()
    return {
      slides: presentation.renderAll(fitWidthPx),
      index: presentation.activeSlide,
    }
  }
  const deleteObjectAt = async (slideIndex: number, objectId: string) => {
    if (!presentation) throw new Error('Open a PPTX presentation first.')
    const rendered = await presentation.deleteElement(
      { slideIndex, sourceId: objectId },
      fitWidthPx,
    )
    if (rendered) {
      await refreshContext()
      publishHistory()
    }
    return rendered
  }
  const replaceAllText = async (op: Parameters<SlidesApi['findReplace']>[0]) => {
    if (!presentation) throw new Error('Open a PPTX presentation first.')
    const replaced = await presentation.findReplace(op, fitWidthPx)
    if (replaced.count) {
      await refreshContext()
      publishHistory()
    }
    return replaced
  }
  const setTextParagraphs = async (
    slideIndex: number,
    objectId: string,
    paragraphs: Parameters<SlidesApi['editText']>[0]['paragraphs'],
    groupId?: string,
  ) => {
    if (!presentation) throw new Error('Open a PPTX presentation first.')
    const rendered = await presentation.editText(
      { slideIndex, sourceId: objectId, ...(groupId ? { groupId } : {}), paragraphs },
      fitWidthPx,
    )
    if (rendered) {
      await refreshContext()
      publishHistory()
    }
    return rendered
  }
  const addElement = async (op: Parameters<SlidesApi['addElement']>[0]) => {
    if (!presentation) throw new Error('Open a PPTX presentation first.')
    const added = await presentation.addElement(op)
    if (added) {
      await refreshContext()
      publishHistory()
    }
    return added ? { slide: added.slide, sourceId: added.objectId } : null
  }

  const sendPresenterMessage = (message: PresenterMessage): void => {
    presenterChannel?.postMessage(message)
  }
  const publishPresenterSnapshot = async (): Promise<void> => {
    if (!presentation) return
    const slides = presentation.renderAll(fitWidthPx)
    const indexes = slides.map((_, index) => index)
    sendPresenterMessage({
      type: 'snapshot',
      slides,
      animations: await Promise.all(indexes.map((index) => presentation!.animations(index))),
      transitions: await Promise.all(indexes.map((index) => presentation!.transition(index))),
      shapeKeys: await Promise.all(indexes.map((index) => presentation!.shapeKeys(index))),
      sync: latestShowSync,
    })
  }
  const receivePresenterMessage = (event: MessageEvent<PresenterMessage>): void => {
    const message = event.data
    if (!message || typeof message !== 'object') return
    if (message.type === 'snapshot') {
      audienceSlides = message.slides
      audienceAnimations = message.animations
      audienceTransitions = message.transitions
      audienceShapeKeys = message.shapeKeys
      if (message.sync) {
        latestShowSync = message.sync
        for (const listener of showSyncListeners) listener(message.sync)
      }
    } else if (message.type === 'sync') {
      latestShowSync = message.state
      for (const listener of showSyncListeners) listener(message.state)
    } else if (message.type === 'ink') {
      for (const listener of showInkListeners) listener(message.event)
    } else if (message.type === 'nav') {
      for (const listener of audienceNavListeners) listener(message.action)
    } else if (message.type === 'ready') {
      void publishPresenterSnapshot()
    }
  }
  const connectPresenterChannel = (sessionId: string): boolean => {
    if (typeof BroadcastChannel !== 'function') return false
    presenterChannel?.close()
    presenterChannel = new BroadcastChannel(`genoffice-pptx-presenter-${sessionId}`)
    presenterChannel.onmessage = receivePresenterMessage
    return true
  }
  if (typeof window !== 'undefined') {
    const current = new URL(window.location.href)
    const sessionId = current.searchParams.get('session')
    if (current.searchParams.get('mode') === 'audience' && sessionId) {
      connectPresenterChannel(sessionId)
    }
  }

  const implemented: SlidesApi = {
    getLanguage: async () => 'zh',
    onLanguageChanged: () => () => undefined,
    getTheme: async () => 'system',
    onThemeChanged: () => () => undefined,
    openPptx: async (fitWidthPx) => {
      const file = await choosePptx()
      return file ? openFile(file, fitWidthPx) : null
    },
    openPptxPath: async () => {
      throw new Error('Browser-hosted PPTX cannot open an arbitrary path without a file handle.')
    },
    consumePendingOpen: async () => null,
    newBlank: createBlankPresentation,
    getLayouts: async (): Promise<GetLayoutsResult | null> =>
      presentation ? presentation.layouts() : null,
    getSlideSize: async () => presentation?.size ?? null,
    getTransition: async (slideIndex): Promise<TransitionKind> =>
      presentation
        ? presentation.transition(slideIndex)
        : (audienceTransitions[slideIndex] ?? 'none'),
    getRenderSlides: async () => presentation?.renderAll(fitWidthPx) ?? audienceSlides,
    getShapeKeys: async (slideIndex) => {
      if (!presentation) return audienceShapeKeys[slideIndex] ?? []
      return presentation.shapeKeys(slideIndex)
    },
    getAnimations: async (slideIndex): Promise<AnimationItem[]> =>
      presentation ? presentation.animations(slideIndex) : (audienceAnimations[slideIndex] ?? []),
    getNotes: async (slideIndex): Promise<string> =>
      presentation ? presentation.notes(slideIndex) : '',
    getComments: async (slideIndex): Promise<SlideComment[]> =>
      presentation ? presentation.comments(slideIndex) : [],
    getSections: async (): Promise<SectionInfo[]> => (presentation ? presentation.sections() : []),
    masterEnter: async (width) => {
      if (!presentation) return null
      const items = await presentation.masterParts(width)
      activeMasterPartPath = items[0]?.partPath ?? null
      return items.length ? { items } : null
    },
    masterOpen: async (partPath) => {
      if (!presentation) return null
      const rendered = await presentation.masterPart(partPath, fitWidthPx)
      if (rendered) activeMasterPartPath = partPath
      return rendered
    },
    masterClose: async () => {
      activeMasterPartPath = null
      return presentation?.renderAll(fitWidthPx) ?? null
    },
    masterDeleteElement: async (op) => {
      if (!presentation || !activeMasterPartPath) return null
      const rendered = await presentation.deleteMasterObject(
        { partPath: activeMasterPartPath, objectId: op.sourceId },
        fitWidthPx,
      )
      if (rendered) {
        await refreshContext()
        publishHistory()
      }
      return rendered
    },
    masterEditText: async (op) => {
      if (!presentation || !activeMasterPartPath) return null
      const rendered = await presentation.setMasterText(
        { partPath: activeMasterPartPath, objectId: op.sourceId, paragraphs: op.paragraphs },
        fitWidthPx,
      )
      if (rendered) {
        await refreshContext()
        publishHistory()
      }
      return rendered
    },
    masterEditTransform: async (op) => {
      if (!presentation || !activeMasterPartPath) return null
      const scale = op.fitWidthPx / (presentation.size.cx / 9_525)
      const toEmu = (px: number) => Math.round((px / scale) * 9_525)
      const rendered = await presentation.setMasterTransform(
        {
          partPath: activeMasterPartPath,
          objectId: op.sourceId,
          xEmu: toEmu(op.xPx),
          yEmu: toEmu(op.yPx),
          widthEmu: toEmu(op.wPx),
          heightEmu: toEmu(op.hPx),
          rotationDegrees: op.rotationDeg,
          ...(op.preview !== undefined ? { preview: op.preview } : {}),
        },
        op.fitWidthPx,
      )
      if (rendered && !op.preview) {
        await refreshContext()
        publishHistory()
      }
      return rendered
    },
    masterEditFill: async (op) => {
      if (!presentation || !activeMasterPartPath) return null
      const fill =
        typeof op.fill === 'string'
          ? op.fill === 'none'
            ? ({ kind: 'none' } as const)
            : ({ kind: 'solid', color: op.fill } as const)
          : ({
              kind: 'gradient' as const,
              from: op.fill.gradient.from,
              to: op.fill.gradient.to,
              ...(op.fill.gradient.radial
                ? { radial: true }
                : { angleDegrees: op.fill.gradient.angleDeg ?? 0 }),
            } as const)
      const rendered = await presentation.setMasterFill(
        { partPath: activeMasterPartPath, objectId: op.sourceId, fill },
        fitWidthPx,
      )
      if (rendered) {
        await refreshContext()
        publishHistory()
      }
      return rendered
    },
    masterEditStroke: async (op) => {
      if (!presentation || !activeMasterPartPath) return null
      const rendered = await presentation.setMasterStroke(
        {
          partPath: activeMasterPartPath,
          objectId: op.sourceId,
          stroke: op.stroke
            ? { color: op.stroke.color, widthEmu: Math.round(op.stroke.widthPt * 12_700) }
            : null,
        },
        fitWidthPx,
      )
      if (rendered) {
        await refreshContext()
        publishHistory()
      }
      return rendered
    },
    editText: async (op) =>
      presentation && !op.groupId
        ? setTextParagraphs(op.slideIndex, op.sourceId, op.paragraphs)
        : presentation
          ? presentation.editText(op, fitWidthPx)
          : null,
    setElementFont: async (op) => {
      if (!presentation) return null
      const rendered = await presentation.setElementFont(op, fitWidthPx)
      if (rendered) {
        await refreshContext()
        publishHistory()
      }
      return rendered
    },
    setElementParagraphFormat: async (op) => {
      if (!presentation) return null
      const rendered = await presentation.setElementParagraphFormat(op, fitWidthPx)
      if (rendered) {
        await refreshContext()
        publishHistory()
      }
      return rendered
    },
    deleteElement: async (op) => (presentation ? deleteObjectAt(op.slideIndex, op.sourceId) : null),
    editTransform: async (op) => {
      if (!presentation) return null
      const rendered = await presentation.editTransform(op)
      if (rendered) {
        await refreshContext()
        publishHistory()
      }
      return rendered
    },
    findReplace: async (op) => (presentation ? replaceAllText(op) : null),
    addElement: async (op) => (presentation ? addElement(op) : null),
    flipElements: async (op) => {
      if (!presentation) return null
      const rendered = await presentation.toggleElements(op, fitWidthPx)
      if (rendered) {
        await refreshContext()
        publishHistory()
      }
      return rendered
    },
    editFill: async (op) => {
      if (!presentation) return null
      const rendered = await presentation.setFill(op, fitWidthPx)
      if (rendered) {
        await refreshContext()
        publishHistory()
      }
      return rendered
    },
    editStroke: async (op) => {
      if (!presentation) return null
      const rendered = await presentation.setStroke(op, fitWidthPx)
      if (rendered) {
        await refreshContext()
        publishHistory()
      }
      return rendered
    },
    editBackground: async (op) => {
      if (!presentation) return null
      const result = await presentation.setBackground(
        {
          scope: op.slideIndex === -1 ? 'all' : 'slide',
          ...(op.slideIndex === -1 ? {} : { slideIndex: op.slideIndex }),
          color: op.color,
        },
        op.fitWidthPx,
      )
      if (result) {
        await refreshContext()
        publishHistory()
      }
      return result?.slides ?? null
    },
    groupElements: async (op) => {
      if (!presentation) return null
      const result = await presentation.groupObjects(op.slideIndex, op.sourceIds, fitWidthPx)
      if (result) {
        await refreshContext()
        publishHistory()
      }
      return result
    },
    ungroupElement: async (op) => {
      if (!presentation) return null
      const result = await presentation.ungroupObject(op.slideIndex, op.sourceId, fitWidthPx)
      if (result) {
        await refreshContext()
        publishHistory()
      }
      return result?.slide ?? null
    },
    reorderElement: async (op) => {
      if (!presentation) return null
      const rendered = await presentation.reorderObject(
        op.slideIndex,
        op.sourceId,
        op.dir,
        fitWidthPx,
      )
      if (rendered) {
        await refreshContext()
        publishHistory()
      }
      return rendered
    },
    batchEditTransform: async (op) => {
      if (!presentation) return null
      const scale = op.fitWidthPx / (presentation.size.cx / 9_525)
      const toEmu = (px: number) => Math.round((px / scale) * 9_525)
      const result = await presentation.setTransformsEmu(
        op.slideIndex,
        op.items.map((item) => ({
          objectId: item.sourceId,
          xEmu: toEmu(item.xPx),
          yEmu: toEmu(item.yPx),
          widthEmu: toEmu(item.wPx),
          heightEmu: toEmu(item.hPx),
          rotationDegrees: item.rotationDeg,
        })),
        op.fitWidthPx,
      )
      if (result) {
        await refreshContext()
        publishHistory()
      }
      return result?.slide ?? null
    },
    editConnectorEndpoints: async (op) => {
      if (!presentation) return null
      const scale = op.fitWidthPx / (presentation.size.cx / 9_525)
      const toEmu = (px: number) => Math.round((px / scale) * 9_525)
      const rendered = await presentation.setConnectorEndpoints(
        {
          slideIndex: op.slideIndex,
          connectorId: op.sourceId,
          start: {
            xEmu: toEmu(op.x1Px),
            yEmu: toEmu(op.y1Px),
            ...(op.start !== undefined
              ? {
                  attachment:
                    op.start === null
                      ? null
                      : { objectId: op.start.targetId, connectionPoint: op.start.idx },
                }
              : {}),
          },
          end: {
            xEmu: toEmu(op.x2Px),
            yEmu: toEmu(op.y2Px),
            ...(op.end !== undefined
              ? {
                  attachment:
                    op.end === null
                      ? null
                      : { objectId: op.end.targetId, connectionPoint: op.end.idx },
                }
              : {}),
          },
        },
        op.fitWidthPx,
      )
      if (rendered) {
        await refreshContext()
        publishHistory()
      }
      return rendered
    },
    editPictureSrcRect: async (op) => {
      if (!presentation) return null
      let frame: { xEmu: number; yEmu: number; widthEmu: number; heightEmu: number } | undefined
      if (op.boxPx && op.fitWidthPx) {
        const scale = op.fitWidthPx / (presentation.size.cx / 9_525)
        const toEmu = (px: number) => Math.round((px / scale) * 9_525)
        frame = {
          xEmu: toEmu(op.boxPx.x),
          yEmu: toEmu(op.boxPx.y),
          widthEmu: toEmu(op.boxPx.w),
          heightEmu: toEmu(op.boxPx.h),
        }
      }
      const rendered = await presentation.setPictureCrop(
        {
          slideIndex: op.slideIndex,
          pictureId: op.sourceId,
          crop: op.srcRect
            ? { left: op.srcRect.l, top: op.srcRect.t, right: op.srcRect.r, bottom: op.srcRect.b }
            : null,
          ...(frame ? { frame } : {}),
        },
        fitWidthPx,
      )
      if (rendered) {
        await refreshContext()
        publishHistory()
      }
      return rendered
    },
    editPictureOpacity: async (op) => {
      if (!presentation) return null
      const rendered = await presentation.setPictureOpacity(
        op.slideIndex,
        op.sourceId,
        op.opacity,
        fitWidthPx,
      )
      if (rendered) {
        await refreshContext()
        publishHistory()
      }
      return rendered
    },
    setTextAnchor: async (op) => {
      if (!presentation) return null
      const rendered = await presentation.setTextVerticalAnchor(
        op.slideIndex,
        op.sourceId,
        op.anchor,
        fitWidthPx,
      )
      if (rendered) {
        await refreshContext()
        publishHistory()
      }
      return rendered
    },
    setSlideLayout: async (op) => {
      if (!presentation) return null
      const rendered = await presentation.setSlideLayout(
        op.slideIndex,
        op.layoutPath ?? null,
        fitWidthPx,
      )
      if (rendered) {
        await refreshContext()
        publishHistory()
      }
      return rendered
    },
    setSlideSize: async (op) => {
      if (!presentation) return null
      const slides = await presentation.setSlideSize(op.cx, op.cy, fitWidthPx)
      if (slides) {
        await refreshContext()
        publishHistory()
      }
      return slides
    },
    setTransition: async (op) => {
      if (!presentation) return false
      const result = await presentation.setSlideTransition(
        {
          scope: op.slideIndex === -1 ? 'all' : 'slide',
          ...(op.slideIndex === -1 ? {} : { slideIndex: op.slideIndex }),
          transition: op.kind,
        },
        fitWidthPx,
      )
      if (result) {
        await refreshContext()
        publishHistory()
      }
      return Boolean(result)
    },
    setSlideHidden: async (op) => {
      if (!presentation) return null
      const rendered = await presentation.setSlideHidden(op.slideIndex, op.hidden, fitWidthPx)
      if (rendered) {
        await refreshContext()
        publishHistory()
      }
      return rendered
    },
    setAdvanceTimes: async (op) => {
      if (!presentation) return false
      const result = await presentation.setSlideAdvanceTimes(
        op.times.map((item) => ({ slideIndex: item.slideIndex, milliseconds: item.ms })),
        fitWidthPx,
      )
      if (result) {
        await refreshContext()
        publishHistory()
      }
      return Boolean(result)
    },
    setAnimations: async (op) => {
      if (!presentation) return false
      const count = await presentation.setAnimations(
        op.slideIndex,
        op.items.map((item) => ({ objectId: item.sourceId, ...item })),
      )
      if (count >= 0) {
        await refreshContext()
        publishHistory()
      }
      return count >= 0
    },
    setLink: async (op) => {
      if (!presentation) return null
      const result = await presentation.setHyperlink(
        op.slideIndex,
        op.sourceId,
        op.target,
        fitWidthPx,
      )
      if (result) {
        await refreshContext()
        publishHistory()
      }
      return result?.slide ?? null
    },
    getLink: async (slideIndex, sourceId) => presentation?.hyperlink(slideIndex, sourceId) ?? null,
    getSlideLinks: async (slideIndex) => presentation?.slideLinks(slideIndex) ?? [],
    getRunLinks: async (slideIndex) => presentation?.runLinks(slideIndex) ?? [],
    applyHeaderFooter: async (op) => {
      if (!presentation) return null
      const result = await presentation.applyHeaderFooter(
        {
          footer: op.footer ?? null,
          slideNumber: op.slideNum ?? false,
          date: op.date ?? null,
          automaticDate: op.dateAuto ?? false,
        },
        op.fitWidthPx,
      )
      if (result) {
        await refreshContext()
        publishHistory()
      }
      return result?.slides ?? null
    },
    getHeaderFooter: async (slideIndex) =>
      presentation?.headerFooter(slideIndex) ?? { footer: null, slideNum: false, date: null },
    applyTheme: async (op) => {
      if (!presentation) return null
      try {
        const slides = await presentation.applyThemeSpec(
          {
            name: op.name,
            colors: op.colors,
            ...(op.majorFont ? { majorFont: op.majorFont } : {}),
            ...(op.minorFont ? { minorFont: op.minorFont } : {}),
          },
          op.fitWidthPx,
        )
        if (slides) {
          await refreshContext()
          publishHistory()
        }
        return slides
      } catch (error) {
        return { error: error instanceof Error ? error.message : String(error) }
      }
    },
    addSlideWithLayout: async (op) => {
      if (!presentation) return null
      const result = await presentation.addSlideWithLayout(
        op.sourceIndex,
        op.layoutPath,
        op.fitWidthPx,
      )
      if (result) {
        await refreshContext()
        publishHistory()
      }
      return result ? { slides: result.slides, index: result.slideIndex } : null
    },
    copySlide: async (slideIndex, pngBase64) => {
      const copied = presentation
        ? await presentation.copySlideForPaste(slideIndex, pngBase64)
        : false
      if (copied) latestClipboardKind = 'slide'
      return copied
    },
    hasSlideClipboard: async () => presentation?.hasSlideClipboard ?? false,
    pasteSlide: async (op) => {
      if (!presentation) return null
      const pasted = await presentation.pasteCopiedSlide(
        op.afterIndex,
        op.mode ?? 'theme',
        op.fitWidthPx,
      )
      if (pasted) {
        await refreshContext()
        publishHistory()
      }
      return pasted
        ? {
            slides: pasted.slides,
            index: pasted.slideIndex,
            ...(pasted.objectId ? { sourceId: pasted.objectId } : {}),
          }
        : null
    },
    repasteSlide: async (op) => {
      if (!presentation) return null
      const pasted = await presentation.repasteCopiedSlide(op.mode, op.fitWidthPx)
      if (pasted) {
        await refreshContext()
        publishHistory()
      }
      return pasted
        ? {
            slides: pasted.slides,
            index: pasted.slideIndex,
            ...(pasted.objectId ? { sourceId: pasted.objectId } : {}),
          }
        : null
    },
    setNotes: async (op) => {
      if (!presentation) return false
      const ok = await presentation.setNotes(op.slideIndex, op.text)
      if (ok) {
        await refreshContext()
        publishHistory()
      }
      return ok
    },
    addComment: async (op) => {
      if (!presentation) return null
      const added = await presentation.addComment(op.slideIndex, 'TandemFolio', op.text)
      if (!added) return null
      await refreshContext()
      publishHistory()
      return presentation.comments(op.slideIndex)
    },
    deleteComment: async (op) => {
      if (!presentation) return null
      const deleted = await presentation.deleteComment(op.slideIndex, op.authorId, op.idx)
      if (!deleted) return null
      await refreshContext()
      publishHistory()
      return presentation.comments(op.slideIndex)
    },
    addSection: async (op) => {
      if (!presentation) return null
      const sections = await presentation.addSection(op.atSlideIndex, op.name)
      if (sections) {
        await refreshContext()
        publishHistory()
      }
      return sections
    },
    renameSection: async (op) => {
      if (!presentation) return null
      const sections = await presentation.renameSection(op.id, op.name)
      if (sections) {
        await refreshContext()
        publishHistory()
      }
      return sections
    },
    removeSection: async (op) => {
      if (!presentation) return null
      const sections = await presentation.removeSection(op.id)
      if (sections) {
        await refreshContext()
        publishHistory()
      }
      return sections
    },
    moveSection: async (op) => {
      if (!presentation) return null
      const result = await presentation.moveSection(op.id, op.dir, fitWidthPx)
      if (result) {
        await refreshContext()
        publishHistory()
      }
      return result
    },
    addTable: async (op) => {
      if (!presentation) return null
      const scale = op.fitWidthPx / (presentation.size.cx / 9_525)
      const toEmu = (px: number) => Math.round((px / scale) * 9_525)
      const result = await presentation.addTableEmu(
        {
          slideIndex: op.slideIndex,
          rows: op.rows,
          columns: op.cols,
          xEmu: toEmu(op.xPx),
          yEmu: toEmu(op.yPx),
          widthEmu: toEmu(op.wPx),
          heightEmu: toEmu(op.hPx),
        },
        op.fitWidthPx,
      )
      if (result) {
        await refreshContext()
        publishHistory()
      }
      return result ? { slide: result.slide, sourceId: result.tableId } : null
    },
    editTableCell: async (op) => {
      if (!presentation) return null
      const rendered = await presentation.setTableCellContent(
        {
          slideIndex: op.slideIndex,
          tableId: op.sourceId,
          row: op.row,
          column: op.col,
          paragraphs: op.paragraphs,
        },
        fitWidthPx,
      )
      if (rendered) {
        await refreshContext()
        publishHistory()
      }
      return rendered
    },
    tableStructure: async (op) => {
      if (!presentation) return null
      const result = await presentation.editTableStructure(
        {
          slideIndex: op.slideIndex,
          tableId: op.sourceId,
          action: op.kind,
          index: op.index,
          ...(op.before !== undefined ? { before: op.before } : {}),
        },
        fitWidthPx,
      )
      if (result) {
        await refreshContext()
        publishHistory()
      }
      return result ? { slide: result.slide, sourceId: result.tableId } : null
    },
    tableMerge: async (op) => {
      if (!presentation) return null
      const result = await presentation.mergeTableCells(
        {
          slideIndex: op.slideIndex,
          tableId: op.sourceId,
          action: op.kind,
          row: op.row,
          column: op.col,
        },
        fitWidthPx,
      )
      if (result) {
        await refreshContext()
        publishHistory()
      }
      return result ? { slide: result.slide, sourceId: result.tableId } : null
    },
    setTableColWidth: async (op) => {
      if (!presentation) return null
      const scale = op.fitWidthPx / (presentation.size.cx / 9_525)
      const widthEmu = (op.wPx / scale) * 9_525
      const rendered = await presentation.setTableColumnWidth(
        op.slideIndex,
        op.sourceId,
        op.col,
        widthEmu,
        op.fitWidthPx,
      )
      if (rendered) {
        await refreshContext()
        publishHistory()
      }
      return rendered
    },
    setTableRowHeight: async (op) => {
      if (!presentation) return null
      const scale = op.fitWidthPx / (presentation.size.cx / 9_525)
      const heightEmu = (op.hPx / scale) * 9_525
      const rendered = await presentation.setTableRowHeight(
        op.slideIndex,
        op.sourceId,
        op.row,
        heightEmu,
        op.fitWidthPx,
      )
      if (rendered) {
        await refreshContext()
        publishHistory()
      }
      return rendered
    },
    setTableCellAnchor: async (op) => {
      if (!presentation) return null
      const rendered = await presentation.setTableCellAnchor(
        op.slideIndex,
        op.sourceId,
        op.row,
        op.col,
        op.anchor,
        fitWidthPx,
      )
      if (rendered) {
        await refreshContext()
        publishHistory()
      }
      return rendered
    },
    editTableStyle: async (op) => {
      if (!presentation) return null
      const result = await presentation.setTableStyle(
        {
          slideIndex: op.slideIndex,
          tableId: op.sourceId,
          ...(op.styleName
            ? {
                styleName: op.styleName as
                  | 'none'
                  | 'lightGrid'
                  | 'zebraBlue'
                  | 'zebraGray'
                  | 'headerDarkBlue'
                  | 'headerOrange'
                  | 'noBorder'
                  | 'fullBorder',
              }
            : {}),
          ...(op.firstRow !== undefined ? { firstRow: op.firstRow } : {}),
          ...(op.bandRow !== undefined ? { bandedRows: op.bandRow } : {}),
          ...(op.shadingColor !== undefined ? { shadingColor: op.shadingColor } : {}),
          ...(op.borderColor ? { borderColor: op.borderColor } : {}),
          ...(op.borderWidthPt != null ? { borderWidthPt: op.borderWidthPt } : {}),
          ...(op.borderPreset ? { borderPreset: op.borderPreset } : {}),
          ...(op.cells
            ? { cells: op.cells.map((cell) => ({ row: cell.row, column: cell.col })) }
            : {}),
        },
        fitWidthPx,
      )
      if (result) {
        await refreshContext()
        publishHistory()
      }
      return result ? { slide: result.slide, sourceId: result.tableId } : null
    },
    addChart: async (op) => {
      if (!presentation) return null
      const scale = op.fitWidthPx / (presentation.size.cx / 9_525)
      const toEmu = (px: number) => Math.round((px / scale) * 9_525)
      const added = await presentation.addChartEmu(
        {
          slideIndex: op.slideIndex,
          kind: op.kind,
          ...(op.title !== undefined ? { title: op.title } : {}),
          categories: op.categories,
          series: op.series,
          xEmu: toEmu(op.xPx),
          yEmu: toEmu(op.yPx),
          widthEmu: toEmu(op.wPx),
          heightEmu: toEmu(op.hPx),
        },
        op.fitWidthPx,
      )
      if (added) {
        await refreshContext()
        publishHistory()
      }
      return added ? { slide: added.slide, sourceId: added.chartId } : null
    },
    addSmartArt: async (op) => {
      if (!presentation) return null
      const scale = op.fitWidthPx / (presentation.size.cx / 9_525)
      const toEmu = (px: number) => Math.round((px / scale) * 9_525)
      const added = await presentation.addSmartArtEmu(
        {
          slideIndex: op.slideIndex,
          layout: op.layout,
          items: op.items,
          xEmu: toEmu(op.xPx),
          yEmu: toEmu(op.yPx),
          widthEmu: toEmu(op.wPx),
          heightEmu: toEmu(op.hPx),
        },
        op.fitWidthPx,
      )
      if (added) {
        await refreshContext()
        publishHistory()
      }
      return added ? { slide: added.slide, sourceId: added.objectId } : null
    },
    addImageBytes: async (op) => {
      if (!presentation) return null
      const extension = op.ext.toLowerCase()
      if (!isImageExtension(extension)) return { error: 'unsupported', ext: op.ext }
      const scale = op.fitWidthPx / (presentation.size.cx / 9_525)
      const toEmu = (px: number) => Math.round((px / scale) * 9_525)
      const added = await presentation.addImageEmu(
        {
          slideIndex: op.slideIndex,
          data: op.base64,
          extension,
          xEmu: toEmu(op.xPx),
          yEmu: toEmu(op.yPx),
          widthEmu: toEmu(op.wPx),
          heightEmu: toEmu(op.hPx),
          ...(op.name ? { name: op.name } : {}),
        },
        op.fitWidthPx,
      )
      if (added) {
        await refreshContext()
        publishHistory()
      }
      return added ? { slide: added.slide, sourceId: added.objectId } : null
    },
    insertImage: async (slideIndex, width) => {
      if (!presentation) return null
      const file = await chooseImage()
      if (!file) return null
      const extension = file.name.split('.').pop()?.toLowerCase() ?? ''
      if (!isImageExtension(extension)) return { error: 'unsupported', ext: extension }
      const natural = await imageNaturalSize(file)
      const deck = presentation.size
      const scale = Math.min(deck.cx / 2 / natural.width, deck.cy / 2 / natural.height)
      const widthEmu = Math.max(1, Math.round(natural.width * scale))
      const heightEmu = Math.max(1, Math.round(natural.height * scale))
      const added = await presentation.addImageEmu(
        {
          slideIndex,
          data: await fileBase64(file),
          extension,
          xEmu: Math.round((deck.cx - widthEmu) / 2),
          yEmu: Math.round((deck.cy - heightEmu) / 2),
          widthEmu,
          heightEmu,
          name: file.name,
        },
        width,
      )
      if (added) {
        await refreshContext()
        publishHistory()
      }
      return added ? { slide: added.slide, sourceId: added.objectId } : null
    },
    editImageFill: async ({ slideIndex, sourceId }) => {
      if (!presentation) return null
      const file = await chooseImage()
      if (!file) return null
      const extension = file.name.split('.').pop()?.toLowerCase() ?? ''
      if (!isImageExtension(extension)) return null
      const rendered = await presentation.setImageFill(
        {
          slideIndex,
          objectId: sourceId,
          data: await fileBase64(file),
          extension,
        },
        fitWidthPx,
      )
      if (rendered) {
        await refreshContext()
        publishHistory()
      }
      return rendered
    },
    replacePictureBytes: async (op) => {
      if (!presentation) return null
      const extension = op.ext.toLowerCase()
      if (!isImageExtension(extension)) return { error: 'unsupported', ext: op.ext }
      const rendered = await presentation.replaceImageBytes(
        {
          slideIndex: op.slideIndex,
          pictureId: op.sourceId,
          data: op.base64,
          extension,
          preserveCrop: op.keepSrcRect ?? false,
        },
        fitWidthPx,
      )
      if (rendered) {
        await refreshContext()
        publishHistory()
      }
      return rendered
    },
    addInk: async (op) => {
      if (!presentation) return null
      const scale = op.fitWidthPx / (presentation.size.cx / 9_525)
      const toEmu = (px: number) => Math.max(1, Math.round((px / scale) * 9_525))
      const added = await presentation.addInkEmu(
        {
          slideIndex: op.slideIndex,
          data: op.base64,
          payload: op.payload,
          xEmu: Math.round((op.xPx / scale) * 9_525),
          yEmu: Math.round((op.yPx / scale) * 9_525),
          widthEmu: toEmu(op.wPx),
          heightEmu: toEmu(op.hPx),
        },
        op.fitWidthPx,
      )
      if (added) {
        await refreshContext()
        publishHistory()
      }
      return added ? { slide: added.slide, sourceId: added.objectId } : null
    },
    duplicateElements: async (op) => {
      if (!presentation) return null
      const scale = op.fitWidthPx / (presentation.size.cx / 9_525)
      const toEmu = (px: number) => Math.round((px / scale) * 9_525)
      const duplicated = await presentation.duplicateObjects(
        {
          slideIndex: op.slideIndex,
          objectIds: op.sourceIds,
          deltaXEmu: toEmu(op.dxPx),
          deltaYEmu: toEmu(op.dyPx),
        },
        op.fitWidthPx,
      )
      if (duplicated) {
        await refreshContext()
        publishHistory()
      }
      return duplicated ? { slide: duplicated.slide, sourceIds: duplicated.objectIds } : null
    },
    copyElements: async (op) => {
      const copied = presentation ? await presentation.copyObjects(op.slideIndex, op.sourceIds) : 0
      if (copied) latestClipboardKind = 'internal'
      return copied
    },
    clipboardExternal: async () => {
      if (latestClipboardKind) return { kind: latestClipboardKind }
      const clipboard = globalThis.navigator?.clipboard
      if (!clipboard) return { kind: 'none' }
      try {
        if ('read' in clipboard && typeof clipboard.read === 'function') {
          const items = await clipboard.read()
          for (const item of items) {
            const imageType = item.types.find((type) => type.startsWith('image/'))
            if (imageType) {
              const blob = await item.getType(imageType)
              const dataUrl = await new Promise<string>((resolve, reject) => {
                const reader = new FileReader()
                reader.addEventListener('load', () => resolve(String(reader.result)), {
                  once: true,
                })
                reader.addEventListener('error', () => reject(reader.error), { once: true })
                reader.readAsDataURL(blob)
              })
              return {
                kind: 'image',
                base64: dataUrl.split(',')[1] ?? '',
                ext: imageType.split('/')[1] ?? 'png',
              }
            }
            if (item.types.includes('text/plain')) {
              const text = await (await item.getType('text/plain')).text()
              if (text) return { kind: 'text', text }
            }
          }
        }
        if ('readText' in clipboard && typeof clipboard.readText === 'function') {
          const text = await clipboard.readText()
          if (text) return { kind: 'text', text }
        }
      } catch {
        return { kind: 'none' }
      }
      return { kind: 'none' }
    },
    pasteElements: async (op) => {
      if (!presentation) return null
      const scale = op.fitWidthPx / (presentation.size.cx / 9_525)
      const shiftStepEmu = Math.round((16 / scale) * 9_525)
      const pasted = await presentation.pasteCopiedObjects(
        op.slideIndex,
        shiftStepEmu,
        op.fitWidthPx,
      )
      if (pasted) {
        await refreshContext()
        publishHistory()
      }
      return pasted ? { slide: pasted.slide, sourceIds: pasted.objectIds } : null
    },
    addMediaBytes: async (op) => {
      if (!presentation) return null
      const extension = op.ext.toLowerCase()
      if (!isMediaExtension(extension)) return null
      const deck = presentation.size
      const widthEmu = Math.round(deck.cx * 0.6)
      const heightEmu = Math.round((widthEmu * 9) / 16)
      const added = await presentation.addMediaBytes(
        {
          slideIndex: op.slideIndex,
          kind: op.kind,
          data: op.base64,
          extension,
          xEmu: Math.round((deck.cx - widthEmu) / 2),
          yEmu: Math.round((deck.cy - heightEmu) / 2),
          widthEmu,
          heightEmu,
          ...(op.name ? { name: op.name } : {}),
        },
        op.fitWidthPx,
      )
      if (added) {
        await refreshContext()
        publishHistory()
      }
      return added ? { slide: added.slide, sourceId: added.objectId } : null
    },
    insertMedia: async (slideIndex, kind, width) => {
      if (!presentation) return null
      const file = await chooseMedia(kind)
      if (!file) return null
      const extension = file.name.split('.').pop()?.toLowerCase() ?? ''
      if (!isMediaExtension(extension)) return null
      const deck = presentation.size
      const widthEmu = Math.round(deck.cx * (kind === 'video' ? 0.6 : 0.24))
      const heightEmu = Math.round(kind === 'video' ? (widthEmu * 9) / 16 : deck.cy * 0.09)
      const added = await presentation.addMediaBytes(
        {
          slideIndex,
          kind,
          data: await fileBase64(file),
          extension,
          xEmu: Math.round((deck.cx - widthEmu) / 2),
          yEmu: Math.round((deck.cy - heightEmu) / 2),
          widthEmu,
          heightEmu,
          name: file.name,
        },
        width,
      )
      if (added) {
        await refreshContext()
        publishHistory()
      }
      return added ? { slide: added.slide, sourceId: added.objectId } : null
    },
    getMediaData: async (slideIndex, sourceId) =>
      presentation?.mediaData(slideIndex, sourceId) ?? null,
    insertModel3d: async (slideIndex, width) => {
      if (!presentation) return null
      const file = await chooseModel3d()
      if (!file) return null
      const extension = file.name.split('.').pop()?.toLowerCase()
      if (extension !== 'glb' && extension !== 'gltf') return null
      const deck = presentation.size
      const widthEmu = Math.round(deck.cx * 0.5)
      const heightEmu = Math.round(deck.cy * 0.5)
      const added = await presentation.addModel3dBytes(
        {
          slideIndex,
          data: await fileBase64(file),
          extension,
          xEmu: Math.round((deck.cx - widthEmu) / 2),
          yEmu: Math.round((deck.cy - heightEmu) / 2),
          widthEmu,
          heightEmu,
          name: file.name,
        },
        width,
      )
      if (added) {
        await refreshContext()
        publishHistory()
      }
      return added ? { slide: added.slide, sourceId: added.objectId } : null
    },
    editChart: async (op) => {
      if (!presentation) return null
      const scheme = op.colorScheme
        ? CHART_COLOR_SCHEMES.find((candidate) => candidate.key === op.colorScheme)
        : undefined
      const pointColors = op.pointColors
        ? Object.entries(op.pointColors).flatMap(([seriesIndex, points]) =>
            Object.entries(points).map(([pointIndex, color]) => ({
              seriesIndex: Number(seriesIndex),
              pointIndex: Number(pointIndex),
              color,
            })),
          )
        : undefined
      const input = {
        slideIndex: op.slideIndex,
        chartId: op.sourceId,
        ...(op.kind !== undefined ? { kind: op.kind } : {}),
        ...(op.title !== undefined ? { title: op.title } : {}),
        ...(op.categories ? { categories: op.categories } : {}),
        ...(op.series ? { series: op.series } : {}),
        ...(scheme?.colors.length ? { colors: [...scheme.colors] } : {}),
        ...(op.legendPos
          ? {
              legendPosition: {
                b: 'bottom',
                t: 'top',
                r: 'right',
                l: 'left',
                none: 'none',
              }[op.legendPos] as 'bottom' | 'top' | 'right' | 'left' | 'none',
            }
          : {}),
        ...(op.dataLabels !== undefined ? { dataLabels: op.dataLabels } : {}),
        ...(op.gridlines !== undefined ? { gridlines: op.gridlines } : {}),
        ...(op.catAxisTitle !== undefined ? { categoryAxisTitle: op.catAxisTitle } : {}),
        ...(op.valAxisTitle !== undefined ? { valueAxisTitle: op.valAxisTitle } : {}),
        ...(op.gapWidthPct !== undefined ? { gapWidthPercent: op.gapWidthPct } : {}),
        ...(op.switchRowCol !== undefined ? { switchRowsAndColumns: op.switchRowCol } : {}),
        ...(pointColors ? { pointColors } : {}),
      }
      let updated = await presentation.updateChart(
        { ...input, allowImportedSimplification: false },
        fitWidthPx,
      )
      if (
        !updated &&
        window.confirm(
          'Editing an imported chart can simplify unsupported PowerPoint formatting. Continue?',
        )
      ) {
        updated = await presentation.updateChart(
          { ...input, allowImportedSimplification: true },
          fitWidthPx,
        )
      }
      if (updated) {
        await refreshContext()
        publishHistory()
      }
      return updated ? { slide: updated.slide, sourceId: updated.chartId } : null
    },
    getChartColorSchemes: async () =>
      CHART_COLOR_SCHEMES.map((scheme) => ({ ...scheme, colors: [...scheme.colors] })),
    getChartData: async (slideIndex, sourceId) =>
      presentation ? presentation.chartData(slideIndex, sourceId) : null,
    addBlankSlide: async ({ sourceIndex, fitWidthPx: width }) =>
      presentation ? addBlankSlideAt(sourceIndex, width) : null,
    addSlide: async ({ sourceIndex, clearText = false, fitWidthPx: width }) =>
      presentation ? duplicateSlideAt(sourceIndex, clearText, width) : null,
    deleteSlide: async (slideIndex) =>
      presentation && presentation.slideCount > 1 ? (await deleteSlideAt(slideIndex)).slides : null,
    moveSlide: async ({ fromIndex, toIndex }) =>
      presentation ? moveSlideTo(fromIndex, toIndex) : null,
    getRecentFiles: async () => [],
    pickExportDir: async () => {
      const picker =
        typeof window !== 'undefined'
          ? (window as DirectoryPickerWindow).showDirectoryPicker
          : undefined
      if (picker && window.parent === window) {
        try {
          exportDirectory = await picker()
          return exportDirectory.name
        } catch (error) {
          if (error instanceof DOMException && error.name === 'AbortError') return null
          throw error
        }
      }
      exportDirectory = null
      return 'Downloads'
    },
    exportImages: async ({ dir, baseName, pngsBase64 }) => {
      try {
        const paths: string[] = []
        const digits = Math.max(2, String(pngsBase64.length).length)
        for (const [index, png] of pngsBase64.entries()) {
          const name = `${baseName}-${String(index + 1).padStart(digits, '0')}.png`
          const blob = new Blob([toArrayBuffer(base64Bytes(png))], { type: 'image/png' })
          if (exportDirectory) {
            const handle = await exportDirectory.getFileHandle(name, { create: true })
            const writable = await handle.createWritable()
            await writable.write(blob)
            await writable.close()
          } else {
            downloadBlob(name, blob)
          }
          paths.push(`${dir.replace(/[\\/]$/, '')}/${name}`)
        }
        return { ok: true, paths }
      } catch (error) {
        return { ok: false, error: String(error) }
      }
    },
    pickExportPdfPath: async (defaultName) => defaultName,
    exportPdf: async ({ filePath, pngsBase64, widthPx, heightPx }) => {
      try {
        const { PDFDocument } = await import('pdf-lib')
        const pdf = await PDFDocument.create()
        const widthPt = Math.max(1, widthPx * 0.75)
        const heightPt = Math.max(1, heightPx * 0.75)
        for (const png of pngsBase64) {
          const image = await pdf.embedPng(base64Bytes(png))
          const page = pdf.addPage([widthPt, heightPt])
          page.drawImage(image, { x: 0, y: 0, width: widthPt, height: heightPt })
        }
        const bytes = await pdf.save()
        const name = filePath.split(/[\\/]/).pop() || 'presentation.pdf'
        downloadBlob(name, new Blob([toArrayBuffer(bytes)], { type: 'application/pdf' }))
        return { ok: true, path: filePath }
      } catch (error) {
        return { ok: false, error: String(error) }
      }
    },
    printSlides: async (input) => {
      try {
        const frame = document.createElement('iframe')
        frame.style.cssText =
          'position:fixed;right:0;bottom:0;width:1px;height:1px;border:0;visibility:hidden;'
        frame.srcdoc = printableSlidesHtml(input)
        await new Promise<void>((resolve, reject) => {
          frame.onload = () => resolve()
          frame.onerror = () => reject(new Error('The print frame could not be loaded.'))
          document.body.append(frame)
        })
        frame.contentWindow?.focus()
        frame.contentWindow?.print()
        frame.remove()
        return { ok: true }
      } catch (error) {
        return { ok: false, error: String(error) }
      }
    },
    presenterStart: async () => {
      if (typeof window === 'undefined') return { audience: false }
      const sessionId =
        globalThis.crypto?.randomUUID?.() ??
        `${Date.now().toString(36)}-${Math.random().toString(36).slice(2)}`
      if (!connectPresenterChannel(sessionId)) return { audience: false }
      const url = new URL(window.location.href)
      url.searchParams.set('mode', 'audience')
      url.searchParams.set('session', sessionId)
      audienceWindow = window.open(url.toString(), 'genoffice-pptx-audience', 'popup')
      if (!audienceWindow) {
        presenterChannel?.close()
        presenterChannel = null
        return { audience: false }
      }
      await publishPresenterSnapshot()
      return { audience: true }
    },
    presenterSync: (state) => {
      latestShowSync = state
      for (const listener of showSyncListeners) listener(state)
      sendPresenterMessage({ type: 'sync', state })
    },
    presenterInk: (event) => {
      for (const listener of showInkListeners) listener(event)
      sendPresenterMessage({ type: 'ink', event })
    },
    presenterSwap: async () => false,
    presenterEnd: async () => {
      if (audienceWindow && !audienceWindow.closed) audienceWindow.close()
      audienceWindow = null
      presenterChannel?.close()
      presenterChannel = null
    },
    audienceReady: async () => {
      sendPresenterMessage({ type: 'ready' })
      return latestShowSync
    },
    audienceNav: (action) => {
      for (const listener of audienceNavListeners) listener(action)
      sendPresenterMessage({ type: 'nav', action })
    },
    onShowSync: (listener) => {
      showSyncListeners.add(listener)
      if (latestShowSync) listener(latestShowSync)
      return () => showSyncListeners.delete(listener)
    },
    onShowInk: (listener) => {
      showInkListeners.add(listener)
      return () => showInkListeners.delete(listener)
    },
    onAudienceNav: (listener) => {
      audienceNavListeners.add(listener)
      return () => audienceNavListeners.delete(listener)
    },
    isDirty: async () => presentation?.dirty ?? false,
    save: () => saveCurrent(),
    saveAs: (defaultName) => saveCurrent(defaultName, true),
    undo: async (): Promise<RenderSlide[] | null> => (await restoreHistory('undo'))?.slides ?? null,
    redo: async (): Promise<RenderSlide[] | null> => (await restoreHistory('redo'))?.slides ?? null,
    nativeClipboard: async (operation) => {
      document.execCommand(operation)
    },
    setAutoSavePref: () => undefined,
    reportCloseSaveResult: () => undefined,
    onCloseSaveRequest: () => () => undefined,
    onHistoryChanged: (listener) => {
      historyListeners.add(listener)
      listener(historyState())
      return () => historyListeners.delete(listener)
    },
    onMenuCommand: () => () => undefined,
    onOpened: (listener) => {
      openedListeners.add(listener)
      return () => openedListeners.delete(listener)
    },
    onRenamed: () => () => undefined,
  }

  const api = implemented

  const recoverySnapshot = async (): Promise<{ fileName: string; data: ArrayBuffer } | null> => {
    if (!presentation?.dirty) return null
    return {
      fileName: activeFileName || 'Untitled.pptx',
      data: toArrayBuffer(await presentation.checkpoint()),
    }
  }
  const execute = async (command: LiveEditorCommand): Promise<LiveEditorExecution> => {
    try {
      const registered = await executePptxOperation(command, {
        createBlank: async () => {
          const created = await createBlankPresentation(fitWidthPx)
          return { fileName: created.path || 'Untitled.pptx', slideCount: created.slides.length }
        },
        loadStaged: async ({ name, data }) => {
          await setPresentation(await BrowserPresentation.open(name, data), fitWidthPx)
        },
        save: async () => {
          const saved = await saveCurrent()
          return saved.ok
            ? { ok: true, fileName: saved.path }
            : { ok: false, message: saved.error ?? 'PPTX save was cancelled.' }
        },
        saveAs: async ({ fileName }) => {
          const saved = await saveCurrent(fileName, true)
          return saved.ok
            ? { ok: true, fileName: saved.path }
            : { ok: false, message: saved.error ?? 'PPTX save-as was cancelled.' }
        },
        setSelection: ({ slideIndex, objectIds }) =>
          presentation?.select(slideIndex, objectIds) ?? 0,
        replaceSelectedText: ({ text }) => presentation?.replaceSelectedText(text) ?? 0,
        moveSelectedObjects: ({ deltaXEmu, deltaYEmu }) =>
          presentation?.moveSelected(deltaXEmu, deltaYEmu) ?? 0,
        addBlankSlide: async ({ afterSlideIndex }) => {
          const added = await addBlankSlideAt(afterSlideIndex, fitWidthPx)
          return { slideIndex: added.index, slideCount: added.slides.length }
        },
        duplicateSlide: async ({ slideIndex, clearText }) => {
          const duplicated = await duplicateSlideAt(slideIndex, clearText, fitWidthPx)
          return { slideIndex: duplicated.index, slideCount: duplicated.slides.length }
        },
        deleteSlide: async ({ slideIndex }) => {
          const deleted = await deleteSlideAt(slideIndex)
          return { slideIndex: deleted.index, slideCount: deleted.slides.length }
        },
        moveSlide: async ({ fromIndex, toIndex }) => {
          const moved = await moveSlideTo(fromIndex, toIndex)
          return { slideIndex: moved.index, slideCount: moved.slides.length }
        },
        undoHistory: async () => {
          const restored = await restoreHistory('undo')
          return restored
            ? { slideIndex: restored.index, slideCount: restored.slides.length }
            : null
        },
        redoHistory: async () => {
          const restored = await restoreHistory('redo')
          return restored
            ? { slideIndex: restored.index, slideCount: restored.slides.length }
            : null
        },
        deleteObject: async ({ slideIndex, objectId }) =>
          Boolean(await deleteObjectAt(slideIndex, objectId)),
        setObjectTransform: async (input) => {
          if (!presentation) return false
          const rendered = await presentation.setTransformEmu(input, fitWidthPx)
          if (rendered) {
            await refreshContext()
            publishHistory()
          }
          return Boolean(rendered)
        },
        setElementFont: async ({ slideIndex, objectIds, ...patch }) => {
          if (!presentation) return 0
          const result = await presentation.setElementFontWithCount(
            { slideIndex, sourceIds: [...objectIds], ...patch },
            fitWidthPx,
          )
          if (result) {
            await refreshContext()
            publishHistory()
          }
          return result?.changed ?? 0
        },
        setParagraphFormat: async ({ slideIndex, objectIds, ...patch }) => {
          if (!presentation) return 0
          const result = await presentation.setElementParagraphFormatWithCount(
            { slideIndex, sourceIds: [...objectIds], ...patch },
            fitWidthPx,
          )
          if (result) {
            await refreshContext()
            publishHistory()
          }
          return result?.changed ?? 0
        },
        replaceAllText: async ({ objectId, ...input }) =>
          (
            await replaceAllText({
              ...input,
              ...(objectId !== undefined ? { elementId: objectId } : {}),
            })
          ).count,
        setTextParagraphs: async ({ slideIndex, objectId, groupId, paragraphs }) =>
          Boolean(await setTextParagraphs(slideIndex, objectId, [...paragraphs], groupId)),
        addObject: async (input) => {
          if (!presentation) throw new Error('Open a PPTX presentation first.')
          const { paragraphs, ...rest } = input
          const added = await presentation.addElementEmu(
            {
              ...rest,
              ...(paragraphs ? { paragraphs: [...paragraphs] } : {}),
            },
            fitWidthPx,
          )
          if (!added) throw new Error('The requested PPTX slide does not exist.')
          await refreshContext()
          publishHistory()
          return added.objectId
        },
        setObjectFlip: async (input) => {
          if (!presentation) return 0
          const result = await presentation.setFlipState(input, fitWidthPx)
          if (result?.changed) {
            await refreshContext()
            publishHistory()
          }
          return result?.changed ?? 0
        },
        setObjectFill: async ({ slideIndex, objectId, groupId, fill }) => {
          if (!presentation) return false
          const mapped =
            fill.kind === 'none'
              ? 'none'
              : fill.kind === 'solid'
                ? (fill.color as string)
                : {
                    gradient: {
                      from: fill.from as string,
                      to: fill.to as string,
                      ...(fill.radial === true
                        ? { radial: true }
                        : { angleDeg: (fill.angleDegrees as number | undefined) ?? 0 }),
                    },
                  }
          const rendered = await presentation.setFill(
            { slideIndex, sourceId: objectId, ...(groupId ? { groupId } : {}), fill: mapped },
            fitWidthPx,
          )
          if (rendered) {
            await refreshContext()
            publishHistory()
          }
          return Boolean(rendered)
        },
        setObjectImageFill: async (input) => {
          if (!presentation) return false
          const rendered = await presentation.setImageFill(input, fitWidthPx)
          if (rendered) {
            await refreshContext()
            publishHistory()
          }
          return Boolean(rendered)
        },
        setObjectStroke: async ({ slideIndex, objectId, groupId, stroke }) => {
          if (!presentation) return false
          const rendered = await presentation.setStroke(
            { slideIndex, sourceId: objectId, ...(groupId ? { groupId } : {}), stroke },
            fitWidthPx,
          )
          if (rendered) {
            await refreshContext()
            publishHistory()
          }
          return Boolean(rendered)
        },
        setSlideBackground: async (input) => {
          if (!presentation) return 0
          const result = await presentation.setBackground(input, fitWidthPx)
          if (result) {
            await refreshContext()
            publishHistory()
          }
          return result?.changed ?? 0
        },
        groupObjects: async ({ slideIndex, objectIds }) => {
          if (!presentation) return ''
          const result = await presentation.groupObjects(slideIndex, objectIds, fitWidthPx)
          if (result) {
            await refreshContext()
            publishHistory()
          }
          return result?.groupId ?? ''
        },
        ungroupObject: async ({ slideIndex, groupId }) => {
          if (!presentation) return 0
          const result = await presentation.ungroupObject(slideIndex, groupId, fitWidthPx)
          if (result) {
            await refreshContext()
            publishHistory()
          }
          return result?.childCount ?? 0
        },
        reorderObject: async ({ slideIndex, objectId, position }) => {
          if (!presentation) return false
          const result = await presentation.reorderObject(
            slideIndex,
            objectId,
            position,
            fitWidthPx,
          )
          if (result) {
            await refreshContext()
            publishHistory()
          }
          return Boolean(result)
        },
        setObjectTransforms: async ({ slideIndex, objects }) => {
          if (!presentation) return 0
          const result = await presentation.setTransformsEmu(slideIndex, objects, fitWidthPx)
          if (result) {
            await refreshContext()
            publishHistory()
          }
          return result?.changed ?? 0
        },
        setConnectorEndpoints: async (input) => {
          if (!presentation) return false
          const result = await presentation.setConnectorEndpoints(input, fitWidthPx)
          if (result) {
            await refreshContext()
            publishHistory()
          }
          return Boolean(result)
        },
        setPictureCrop: async (input) => {
          if (!presentation) return false
          const result = await presentation.setPictureCrop(input, fitWidthPx)
          if (result) {
            await refreshContext()
            publishHistory()
          }
          return Boolean(result)
        },
        setPictureOpacity: async ({ slideIndex, pictureId, opacity }) => {
          if (!presentation) return false
          const result = await presentation.setPictureOpacity(
            slideIndex,
            pictureId,
            opacity,
            fitWidthPx,
          )
          if (result) {
            await refreshContext()
            publishHistory()
          }
          return Boolean(result)
        },
        setTextVerticalAnchor: async ({ slideIndex, objectId, anchor }) => {
          if (!presentation) return false
          const result = await presentation.setTextVerticalAnchor(
            slideIndex,
            objectId,
            anchor,
            fitWidthPx,
          )
          if (result) {
            await refreshContext()
            publishHistory()
          }
          return Boolean(result)
        },
        addTable: async (input) => {
          if (!presentation) return ''
          const result = await presentation.addTableEmu(input, fitWidthPx)
          if (result) {
            await refreshContext()
            publishHistory()
          }
          return result?.tableId ?? ''
        },
        setTableCellContent: async (input) => {
          if (!presentation) return false
          const result = await presentation.setTableCellContent(input, fitWidthPx)
          if (result) {
            await refreshContext()
            publishHistory()
          }
          return Boolean(result)
        },
        editTableStructure: async (input) => {
          if (!presentation) return ''
          const result = await presentation.editTableStructure(input, fitWidthPx)
          if (result) {
            await refreshContext()
            publishHistory()
          }
          return result?.tableId ?? ''
        },
        mergeTableCells: async (input) => {
          if (!presentation) return ''
          const result = await presentation.mergeTableCells(input, fitWidthPx)
          if (result) {
            await refreshContext()
            publishHistory()
          }
          return result?.tableId ?? ''
        },
        setTableColumnWidth: async ({ slideIndex, tableId, column, widthEmu }) => {
          if (!presentation) return false
          const result = await presentation.setTableColumnWidth(
            slideIndex,
            tableId,
            column,
            widthEmu,
            fitWidthPx,
          )
          if (result) {
            await refreshContext()
            publishHistory()
          }
          return Boolean(result)
        },
        setTableRowHeight: async ({ slideIndex, tableId, row, heightEmu }) => {
          if (!presentation) return false
          const result = await presentation.setTableRowHeight(
            slideIndex,
            tableId,
            row,
            heightEmu,
            fitWidthPx,
          )
          if (result) {
            await refreshContext()
            publishHistory()
          }
          return Boolean(result)
        },
        setTableCellAnchor: async ({ slideIndex, tableId, row, column, anchor }) => {
          if (!presentation) return false
          const result = await presentation.setTableCellAnchor(
            slideIndex,
            tableId,
            row,
            column,
            anchor,
            fitWidthPx,
          )
          if (result) {
            await refreshContext()
            publishHistory()
          }
          return Boolean(result)
        },
        setTableStyle: async (input) => {
          if (!presentation) return ''
          const result = await presentation.setTableStyle(input, fitWidthPx)
          if (result) {
            await refreshContext()
            publishHistory()
          }
          return result?.tableId ?? ''
        },
        addChart: async (input) => {
          if (!presentation) return ''
          const result = await presentation.addChartEmu(input, fitWidthPx)
          if (result) {
            await refreshContext()
            publishHistory()
          }
          return result?.chartId ?? ''
        },
        addSmartArt: async (input) => {
          if (!presentation) return ''
          const result = await presentation.addSmartArtEmu(input, fitWidthPx)
          if (result) {
            await refreshContext()
            publishHistory()
          }
          return result?.objectId ?? ''
        },
        addImage: async (input) => {
          if (!presentation) return ''
          const result = await presentation.addImageEmu(input, fitWidthPx)
          if (result) {
            await refreshContext()
            publishHistory()
          }
          return result?.objectId ?? ''
        },
        replaceImage: async (input) => {
          if (!presentation) return false
          const result = await presentation.replaceImageBytes(input, fitWidthPx)
          if (result) {
            await refreshContext()
            publishHistory()
          }
          return Boolean(result)
        },
        addInk: async (input) => {
          if (!presentation) return ''
          const result = await presentation.addInkEmu(input, fitWidthPx)
          if (result) {
            await refreshContext()
            publishHistory()
          }
          return result?.objectId ?? ''
        },
        duplicateObjects: async (input) => {
          if (!presentation) return []
          const result = await presentation.duplicateObjects(input, fitWidthPx)
          if (result) {
            await refreshContext()
            publishHistory()
          }
          return result?.objectIds ?? []
        },
        copyObjectsTo: async (input) => {
          if (!presentation) return []
          const result = await presentation.copyObjectsTo(input, fitWidthPx)
          if (result) {
            await refreshContext()
            publishHistory()
          }
          return result?.objectIds ?? []
        },
        copySlideTo: async (input) => {
          if (!presentation) return null
          const result = await presentation.copySlideTo(input, fitWidthPx)
          if (result) {
            await refreshContext()
            publishHistory()
          }
          return result
            ? {
                slideIndex: result.slideIndex,
                slideCount: result.slideCount,
                ...(result.objectId ? { objectId: result.objectId } : {}),
              }
            : null
        },
        deleteMasterObject: async (input) => {
          if (!presentation) return false
          const result = await presentation.deleteMasterObject(input, fitWidthPx)
          if (result) {
            await refreshContext()
            publishHistory()
          }
          return Boolean(result)
        },
        setMasterText: async (input) => {
          if (!presentation) return false
          const result = await presentation.setMasterText(input, fitWidthPx)
          if (result) {
            await refreshContext()
            publishHistory()
          }
          return Boolean(result)
        },
        setMasterTransform: async (input) => {
          if (!presentation) return false
          const result = await presentation.setMasterTransform(input, fitWidthPx)
          if (result) {
            await refreshContext()
            publishHistory()
          }
          return Boolean(result)
        },
        setMasterFill: async (input) => {
          if (!presentation) return false
          const result = await presentation.setMasterFill(input, fitWidthPx)
          if (result) {
            await refreshContext()
            publishHistory()
          }
          return Boolean(result)
        },
        setMasterStroke: async (input) => {
          if (!presentation) return false
          const result = await presentation.setMasterStroke(input, fitWidthPx)
          if (result) {
            await refreshContext()
            publishHistory()
          }
          return Boolean(result)
        },
        addMedia: async (input) => {
          if (!presentation) return ''
          const result = await presentation.addMediaBytes(input, fitWidthPx)
          if (result) {
            await refreshContext()
            publishHistory()
          }
          return result?.objectId ?? ''
        },
        addModel3d: async (input) => {
          if (!presentation) return ''
          const result = await presentation.addModel3dBytes(input, fitWidthPx)
          if (result) {
            await refreshContext()
            publishHistory()
          }
          return result?.objectId ?? ''
        },
        applyTheme: async ({ preset }) => {
          if (!presentation) return 0
          const theme = THEME_PRESETS.find((candidate) => candidate.id === preset)
          if (!theme) return 0
          const result = await presentation.applyThemeSpec(theme, fitWidthPx)
          if (result) {
            await refreshContext()
            publishHistory()
          }
          return result?.length ?? 0
        },
        updateChart: async (input) => {
          if (!presentation) return ''
          const result = await presentation.updateChart(input, fitWidthPx)
          if (result) {
            await refreshContext()
            publishHistory()
          }
          return result?.chartId ?? ''
        },
        setSlideLayout: async ({ slideIndex, layoutPath }) => {
          if (!presentation) return false
          const result = await presentation.setSlideLayout(slideIndex, layoutPath, fitWidthPx)
          if (result) {
            await refreshContext()
            publishHistory()
          }
          return Boolean(result)
        },
        setSlideSize: async ({ widthEmu, heightEmu }) => {
          if (!presentation) return 0
          const result = await presentation.setSlideSize(widthEmu, heightEmu, fitWidthPx)
          if (result) {
            await refreshContext()
            publishHistory()
          }
          return result?.length ?? 0
        },
        setSlideTransition: async (input) => {
          if (!presentation) return 0
          const result = await presentation.setSlideTransition(input, fitWidthPx)
          if (result) {
            await refreshContext()
            publishHistory()
          }
          return result?.changed ?? 0
        },
        setSlideHidden: async ({ slideIndex, hidden }) => {
          if (!presentation) return false
          const result = await presentation.setSlideHidden(slideIndex, hidden, fitWidthPx)
          if (result) {
            await refreshContext()
            publishHistory()
          }
          return Boolean(result)
        },
        setSlideAdvanceTimes: async ({ slides }) => {
          if (!presentation) return 0
          const result = await presentation.setSlideAdvanceTimes(slides, fitWidthPx)
          if (result) {
            await refreshContext()
            publishHistory()
          }
          return result?.changed ?? 0
        },
        setAnimations: async ({ slideIndex, animations }) => {
          if (!presentation) return -1
          const result = await presentation.setAnimations(slideIndex, animations)
          if (result >= 0) {
            await refreshContext()
            publishHistory()
          }
          return result
        },
        setHyperlink: async ({ slideIndex, objectId, target }) => {
          if (!presentation) return ''
          const result = await presentation.setHyperlink(slideIndex, objectId, target, fitWidthPx)
          if (result) {
            await refreshContext()
            publishHistory()
          }
          return result?.objectId ?? ''
        },
        applyHeaderFooter: async (input) => {
          if (!presentation) return 0
          const result = await presentation.applyHeaderFooter(input, fitWidthPx)
          if (result) {
            await refreshContext()
            publishHistory()
          }
          return result?.changed ?? 0
        },
        addSlideWithLayout: async ({ afterSlideIndex, layoutPath }) => {
          if (!presentation) throw new Error('Open a PPTX presentation first.')
          const result = await presentation.addSlideWithLayout(
            afterSlideIndex,
            layoutPath,
            fitWidthPx,
          )
          if (!result) throw new Error('The requested PPTX slide or layout does not exist.')
          await refreshContext()
          publishHistory()
          return { slideIndex: result.slideIndex, slideCount: result.slides.length }
        },
        setNotes: async ({ slideIndex, text }) => {
          if (!presentation) return false
          const result = await presentation.setNotes(slideIndex, text)
          if (result) {
            await refreshContext()
            publishHistory()
          }
          return result
        },
        addComment: async ({ slideIndex, author, text }) => {
          if (!presentation) throw new Error('Open a PPTX presentation first.')
          const result = await presentation.addComment(slideIndex, author, text)
          if (!result) throw new Error('The requested PPTX slide does not exist.')
          await refreshContext()
          publishHistory()
          return { authorId: result.authorId, index: result.idx }
        },
        deleteComment: async ({ slideIndex, authorId, index }) => {
          if (!presentation) return false
          const result = await presentation.deleteComment(slideIndex, authorId, index)
          if (result) {
            await refreshContext()
            publishHistory()
          }
          return result
        },
        addSection: async ({ beforeSlideIndex, name }) => {
          if (!presentation) return 0
          const result = await presentation.addSection(beforeSlideIndex, name)
          if (result) {
            await refreshContext()
            publishHistory()
          }
          return result?.length ?? 0
        },
        renameSection: async ({ sectionId, name }) => {
          if (!presentation) return 0
          const result = await presentation.renameSection(sectionId, name)
          if (result) {
            await refreshContext()
            publishHistory()
          }
          return result?.length ?? 0
        },
        removeSection: async ({ sectionId }) => {
          if (!presentation) return -1
          const result = await presentation.removeSection(sectionId)
          if (result) {
            await refreshContext()
            publishHistory()
          }
          return result?.length ?? -1
        },
        moveSection: async ({ sectionId, direction }) => {
          if (!presentation) return 0
          const result = await presentation.moveSection(sectionId, direction, fitWidthPx)
          if (result) {
            await refreshContext()
            publishHistory()
          }
          return result?.sections.length ?? 0
        },
      })
      if (registered.handled) {
        if (!registered.ok) {
          return {
            ok: false,
            error: registered.error,
            message: registered.message,
          }
        }
        await refreshContext()
        const output = { output: registered.output }
        if (registered.refreshDocument) {
          for (const listener of openedListeners) listener(result(fitWidthPx))
        }
        if (registered.checkpointRecovery === false || !presentation?.dirty) {
          return { ok: true, ...output }
        }
        if (!registered.refreshDocument) {
          for (const listener of openedListeners) listener(result(fitWidthPx))
        }
        const recovery = await recoverySnapshot()
        return { ok: true, ...output, ...(recovery ? { recovery } : {}) }
      }
      if (!presentation) {
        return { ok: false, error: 'execution_failed', message: 'Open a PPTX presentation first.' }
      }
      return {
        ok: false,
        error: 'unsupported_operation',
        message: `Unsupported PPTX operation: ${command.operation}`,
      }
    } catch (error) {
      return {
        ok: false,
        error: 'execution_failed',
        message: error instanceof Error ? error.message : String(error),
      }
    }
  }
  const adapter: LiveEditorAdapter = {
    execute,
    snapshot(revision: number): LiveEditorSnapshot {
      return {
        revision,
        fileName: presentation ? activeFileName || null : null,
        dirty: presentation?.dirty ?? false,
        selection: latestContext,
      }
    },
    recoverySnapshot,
    recoveryVersion: () => recoveryVersion,
  }

  return { api, adapter }
}
