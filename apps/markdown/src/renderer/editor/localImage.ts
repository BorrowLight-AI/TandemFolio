import { Image } from '@tiptap/extension-image'
import type { Editor } from '@tiptap/core'
import { Plugin, PluginKey } from '@tiptap/pm/state'
import { insertMarkdownImage, markdownImageDataUrl } from './image-actions'

/** Directory of the open .md file; relative image paths resolve against it for display */
let imageBaseDir: string | null = null

export function setImageBaseDir(dir: string | null): void {
  imageBaseDir = dir
}

/**
 * Map an authored image src to a displayable URL. Markdown keeps the authored
 * value (usually a path relative to the .md file); the editor DOM loads it via
 * the main process's md-asset:// handler — a plain file:// subresource would be
 * blocked when the renderer page itself is served over http (dev server).
 */
export function resolveImageSrc(src: string, baseDir: string | null = imageBaseDir): string {
  if (!src) return src
  // ':' is legal in URL path segments (RFC 3986) — restore it after encoding so
  // Windows drive prefixes stay `C:` instead of the `C%3A` Chromium rejects
  const encodeSegment = (seg: string) => encodeURIComponent(seg).replace(/%3A/gi, ':')
  const toAssetUrl = (path: string) =>
    `md-asset://${path.startsWith('/') ? '' : '/'}${path.replace(/\\/g, '/').split('/').map(encodeSegment).join('/')}`
  // a Windows drive path would also match the URL-scheme regex — check it first
  if (src.startsWith('/') || /^[a-zA-Z]:[\\/]/.test(src)) return toAssetUrl(src)
  if (/^[a-z][a-z0-9+.-]*:/i.test(src)) return src
  if (!baseDir) return src
  return toAssetUrl(`${baseDir.replace(/\\/g, '/').replace(/\/$/, '')}/${src}`)
}

/**
 * Reverse of {@link resolveImageSrc}: map a display URL (md-asset://) back to
 * the authored path so DOM-parsed content (copy/paste inside the editor) never
 * bakes display URLs into the stored document / serialized markdown.
 */
export function unresolveImageSrc(src: string, baseDir: string | null = imageBaseDir): string {
  if (!src.startsWith('md-asset://')) return src
  let path = decodeURIComponent(src.slice('md-asset://'.length))
  // Windows drive paths were prefixed with '/' to form a valid URL path
  if (/^\/[a-zA-Z]:\//.test(path)) path = path.slice(1)
  if (baseDir) {
    const base = `${baseDir.replace(/\\/g, '/').replace(/\/$/, '')}/`
    if (path.startsWith(base)) return path.slice(base.length)
  }
  return path
}

const EXT_BY_MIME: Record<string, string> = {
  'image/png': 'png',
  'image/jpeg': 'jpg',
  'image/gif': 'gif',
}

function imageFileIn(data: DataTransfer | null): File | null {
  for (const file of data?.files ?? []) {
    if (EXT_BY_MIME[file.type]) return file
  }
  return null
}

export type MarkdownLocalAssetReader = (
  rootId: string,
  path: string,
) => Promise<{
  readonly mime: 'image/png' | 'image/jpeg' | 'image/gif'
  readonly data: ArrayBuffer
} | null>

function isAuthoredLocalImagePath(src: string): boolean {
  return Boolean(src) && !/^[a-z][a-z0-9+.-]*:/i.test(src)
}

/** Hydrate authored filesystem paths for display/export without changing Markdown serialization. */
export async function hydrateMarkdownLocalImages(
  editor: Editor,
  rootId: string,
  readAsset: MarkdownLocalAssetReader,
): Promise<number> {
  const pending: Array<{ position: number; authoredSrc: string }> = []
  editor.state.doc.descendants((node, position) => {
    if (node.type.name !== 'image') return
    const authoredSrc = String(node.attrs.authoredSrc ?? node.attrs.src ?? '')
    if (isAuthoredLocalImagePath(authoredSrc)) pending.push({ position, authoredSrc })
  })
  const hydrated = await Promise.all(
    pending.map(async (item) => ({ ...item, asset: await readAsset(rootId, item.authoredSrc) })),
  )
  let transaction = editor.state.tr
  let changed = 0
  for (const item of hydrated) {
    if (!item.asset) continue
    const node = transaction.doc.nodeAt(item.position)
    if (node?.type.name !== 'image') continue
    transaction = transaction.setNodeMarkup(item.position, undefined, {
      ...node.attrs,
      src: markdownImageDataUrl(item.asset.mime, item.asset.data),
      authoredSrc: item.authoredSrc,
    })
    changed += 1
  }
  if (changed > 0) {
    editor.view.dispatch(transaction.setMeta('addToHistory', false).setMeta('uiOnly', true))
  }
  return changed
}

async function persistAndInsert(
  editor: import('@tiptap/core').Editor,
  file: File,
  pos: number,
): Promise<void> {
  const bytes = new Uint8Array(await file.arrayBuffer())
  const alt = file.name.replace(/\.[a-z0-9]+$/i, '')
  insertMarkdownImage(editor, {
    position: pos,
    src: markdownImageDataUrl(
      file.type,
      bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer,
    ),
    alt,
    title: null,
  })
}

/**
 * Image node whose DOM src is display-resolved while the stored attribute (and
 * therefore the serialized markdown) keeps the authored path untouched.
 * In the browser/MCP host, pasted and dropped image files are embedded as data
 * URLs so the mounted renderer never requires filesystem or custom-protocol access.
 */
export const LocalImage = Image.extend({
  addAttributes() {
    return {
      ...this.parent?.(),
      // resolved at the attribute level so the renderer displays the
      // md-asset:// URL while the stored value keeps the authored path
      src: {
        default: null,
        parseHTML: (element) => unresolveImageSrc(element.getAttribute('src') ?? ''),
        renderHTML: (attrs) => ({ src: resolveImageSrc(String(attrs.src ?? '')) }),
      },
      authoredSrc: {
        default: null,
        parseHTML: () => null,
        renderHTML: () => ({}),
      },
    }
  },

  renderMarkdown: (node) => {
    const src = String(node.attrs?.authoredSrc ?? node.attrs?.src ?? '')
    const alt = String(node.attrs?.alt ?? '')
    const title = String(node.attrs?.title ?? '')
    return title ? `![${alt}](${src} "${title}")` : `![${alt}](${src})`
  },

  addProseMirrorPlugins() {
    const editor = this.editor
    return [
      new Plugin({
        key: new PluginKey('localImageUpload'),
        props: {
          handlePaste(view, event) {
            const file = imageFileIn(event.clipboardData)
            if (!file) return false
            void persistAndInsert(editor, file, view.state.selection.from)
            return true
          },
          handleDrop(view, event, _slice, moved) {
            if (moved) return false
            const file = imageFileIn(event.dataTransfer)
            if (!file) return false
            const pos =
              view.posAtCoords({ left: event.clientX, top: event.clientY })?.pos ??
              view.state.selection.from
            void persistAndInsert(editor, file, pos)
            return true
          },
        },
      }),
    ]
  },
})
