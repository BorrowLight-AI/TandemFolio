import type { JSONContent } from '@tiptap/core'
import {
  BLANK_BULLET_NUM_ID,
  TABLE_HEADER_FILL,
  buildBlankDocx,
  generateTableModelXml,
  parseDocx,
  saveDocx,
} from '@genoffice/docx-engine'
import type {
  GeneratedBlock,
  NewImage,
  ParaFormat,
  Run,
  SaveBlock,
  SaveOptions,
  TableCell,
  TableModel,
  TableParagraph,
} from '@genoffice/docx-engine'

/** Resolve an authored image src to embeddable bytes; null → fall back to alt text */
export type ImageLoader = (src: string) => Promise<NewImage | null>

/** widest image that fits the A4 text column */
export const DOCX_MAX_IMAGE_PX = 620
export const DOCX_MAX_IMAGE_BYTES = 20_971_520

export interface DocxMapping {
  blocks: SaveBlock[]
  options: SaveOptions
}

type DocxImageMime = NewImage['mime']

function bytesToBase64(bytes: Uint8Array): string {
  let binary = ''
  for (let offset = 0; offset < bytes.length; offset += 0x8000) {
    binary += String.fromCharCode(...bytes.subarray(offset, offset + 0x8000))
  }
  return btoa(binary)
}

function base64ToBytes(base64: string): Uint8Array | null {
  try {
    const binary = atob(base64)
    if (binary.length > DOCX_MAX_IMAGE_BYTES) return null
    return Uint8Array.from(binary, (character) => character.charCodeAt(0))
  } catch {
    return null
  }
}

function matchesImageMime(bytes: Uint8Array, mime: DocxImageMime): boolean {
  if (mime === 'image/png') {
    return (
      bytes.length >= 24 &&
      bytes[0] === 0x89 &&
      bytes[1] === 0x50 &&
      bytes[2] === 0x4e &&
      bytes[3] === 0x47 &&
      bytes[4] === 0x0d &&
      bytes[5] === 0x0a &&
      bytes[6] === 0x1a &&
      bytes[7] === 0x0a
    )
  }
  if (mime === 'image/gif') {
    const signature = String.fromCharCode(...bytes.subarray(0, 6))
    return bytes.length >= 10 && (signature === 'GIF87a' || signature === 'GIF89a')
  }
  return bytes.length >= 4 && bytes[0] === 0xff && bytes[1] === 0xd8
}

function imageDimensions(
  bytes: Uint8Array,
  mime: DocxImageMime,
): { width: number; height: number } | null {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  if (mime === 'image/png') {
    return { width: view.getUint32(16), height: view.getUint32(20) }
  }
  if (mime === 'image/gif') {
    return { width: view.getUint16(6, true), height: view.getUint16(8, true) }
  }
  let offset = 2
  const startOfFrame = new Set([
    0xc0, 0xc1, 0xc2, 0xc3, 0xc5, 0xc6, 0xc7, 0xc9, 0xca, 0xcb, 0xcd, 0xce, 0xcf,
  ])
  while (offset + 8 < bytes.length) {
    while (offset < bytes.length && bytes[offset] !== 0xff) offset += 1
    while (offset < bytes.length && bytes[offset] === 0xff) offset += 1
    if (offset >= bytes.length) break
    const marker = bytes[offset++]!
    if (marker === 0xd8 || marker === 0xd9) continue
    if (offset + 2 > bytes.length) break
    const length = view.getUint16(offset)
    if (length < 2 || offset + length > bytes.length) break
    if (startOfFrame.has(marker) && length >= 7) {
      return { width: view.getUint16(offset + 5), height: view.getUint16(offset + 3) }
    }
    offset += length
  }
  return null
}

async function boundedResponseBytes(response: Response): Promise<Uint8Array | null> {
  if (!response.ok) return null
  const declaredLength = Number(response.headers.get('content-length') ?? '0')
  if (declaredLength > DOCX_MAX_IMAGE_BYTES) return null
  const bytes = new Uint8Array(await response.arrayBuffer())
  return bytes.length <= DOCX_MAX_IMAGE_BYTES ? bytes : null
}

/** Resolve browser-safe Markdown image URLs into bounded DOCX image payloads. */
export async function loadMarkdownImageForDocx(src: string): Promise<NewImage | null> {
  let mime: DocxImageMime
  let bytes: Uint8Array | null
  const data = /^data:(image\/(?:png|jpeg|gif));base64,([a-z0-9+/=\s]+)$/i.exec(src)
  if (data) {
    mime = data[1]!.toLowerCase() as DocxImageMime
    bytes = base64ToBytes(data[2]!.replace(/\s/g, ''))
  } else {
    if (/^md-asset:/i.test(src) || src.length > 4096) return null
    const response = await fetch(src)
    const contentType = response.headers.get('content-type')?.split(';', 1)[0]?.toLowerCase()
    if (!contentType || !['image/png', 'image/jpeg', 'image/gif'].includes(contentType)) return null
    mime = contentType as DocxImageMime
    bytes = await boundedResponseBytes(response)
  }
  if (!bytes || !matchesImageMime(bytes, mime)) return null
  const natural = imageDimensions(bytes, mime)
  if (!natural || natural.width < 1 || natural.height < 1) return null
  const scale = Math.min(1, DOCX_MAX_IMAGE_PX / natural.width)
  return {
    base64: bytesToBase64(bytes),
    mime,
    widthPx: Math.max(1, Math.round(natural.width * scale)),
    heightPx: Math.max(1, Math.round(natural.height * scale)),
  }
}

const MAX_LIST_LEVEL = 4
const INDENT_STEP = 360
/** decimal abstractNum of the blank template (numbering.xml: 0 = bullet, 1 = decimal) */
const DECIMAL_ABSTRACT_NUM_ID = '1'
const CODE_FONT = 'Consolas'
const CODE_FILL = 'F2F3F5'

// ── inline content → Run[] ──

function runsFromInline(content: JSONContent[] | undefined): Run[] {
  const runs: Run[] = []
  for (const child of content ?? []) {
    if (child.type === 'hardBreak') {
      runs.push({ text: '\n' })
      continue
    }
    if (child.type !== 'text' || !child.text) continue
    const run: Run = { text: child.text }
    for (const mark of child.marks ?? []) {
      if (mark.type === 'bold') run.bold = true
      else if (mark.type === 'italic') run.italic = true
      else if (mark.type === 'strike') run.strike = true
      else if (mark.type === 'code') run.font = CODE_FONT
      else if (mark.type === 'link' && typeof mark.attrs?.href === 'string') {
        run.link = { href: mark.attrs.href }
      }
    }
    runs.push(run)
  }
  return runs
}

function plainText(node: JSONContent): string {
  if (node.type === 'text') return node.text ?? ''
  return (node.content ?? []).map(plainText).join('')
}

// ── mapping walker ──

interface WalkContext {
  blocks: SaveBlock[]
  restartNums: NonNullable<NonNullable<SaveOptions['numbering']>['restartNums']>
  nextOrderedNumId: number
  loadImage: ImageLoader
  pendingImages: Array<{ index: number; src: string; alt: string }>
}

function mergeFormat(base: ParaFormat | undefined, extra: ParaFormat): ParaFormat {
  return { ...base, ...extra, indentLeft: (base?.indentLeft ?? 0) + (extra.indentLeft ?? 0) }
}

function pushParagraph(ctx: WalkContext, block: GeneratedBlock): void {
  ctx.blocks.push({ kind: 'generated', block })
}

function walkList(
  ctx: WalkContext,
  node: JSONContent,
  kind: 'bullet' | 'ordered',
  ilvl: number,
  numId: string,
  base?: ParaFormat,
): void {
  for (const item of node.content ?? []) {
    if (item.type !== 'listItem' && item.type !== 'taskItem') continue
    let firstPara = true
    for (const child of item.content ?? []) {
      if (child.type === 'paragraph') {
        const runs = runsFromInline(child.content)
        if (item.type === 'taskItem') {
          runs.unshift({ text: item.attrs?.checked ? '☑ ' : '☐ ' })
        }
        if (firstPara && item.type !== 'taskItem') {
          pushParagraph(ctx, {
            type: 'listItem',
            list: { kind, numId, ilvl: Math.min(ilvl, MAX_LIST_LEVEL) },
            runs,
            format: base,
          })
        } else {
          // task items and continuation paragraphs indent under the marker
          pushParagraph(ctx, {
            type: 'paragraph',
            runs,
            format: mergeFormat(base, { indentLeft: INDENT_STEP * (ilvl + 1) }),
          })
        }
        firstPara = false
      } else if (child.type === 'bulletList' || child.type === 'taskList') {
        walkList(ctx, child, 'bullet', ilvl + 1, BLANK_BULLET_NUM_ID, base)
      } else if (child.type === 'orderedList') {
        walkList(ctx, child, 'ordered', ilvl + 1, allocOrderedNumId(ctx), base)
      } else {
        walkBlock(ctx, child, mergeFormat(base, { indentLeft: INDENT_STEP * (ilvl + 1) }))
      }
    }
  }
}

function allocOrderedNumId(ctx: WalkContext): string {
  const numId = String(ctx.nextOrderedNumId++)
  ctx.restartNums.push({
    numId,
    abstractNumId: DECIMAL_ABSTRACT_NUM_ID,
    startOverrides: { 0: 1, 1: 1, 2: 1, 3: 1, 4: 1 },
  })
  return numId
}

function tableParagraphs(cell: JSONContent): TableParagraph[] {
  const paras: TableParagraph[] = []
  for (const child of cell.content ?? []) {
    if (child.type === 'paragraph') paras.push({ runs: runsFromInline(child.content) })
  }
  return paras.length > 0 ? paras : [{ runs: [] }]
}

function mapTable(node: JSONContent): TableModel {
  const rows: TableCell[][] = []
  for (const row of node.content ?? []) {
    if (row.type !== 'tableRow') continue
    const cells: TableCell[] = []
    for (const cell of row.content ?? []) {
      if (cell.type !== 'tableCell' && cell.type !== 'tableHeader') continue
      const isHeader = cell.type === 'tableHeader'
      const rich = tableParagraphs(cell)
      cells.push({
        paras: rich.map((p) => p.runs.map((r) => r.text).join('')),
        richParas: rich,
        ...(Number(cell.attrs?.colspan) > 1 ? { colSpan: Number(cell.attrs?.colspan) } : {}),
        ...(isHeader ? { bold: true, fill: TABLE_HEADER_FILL } : {}),
      })
    }
    if (cells.length > 0) rows.push(cells)
  }
  return { rows }
}

function walkBlock(ctx: WalkContext, node: JSONContent, base?: ParaFormat): void {
  switch (node.type) {
    case 'paragraph':
      pushParagraph(ctx, {
        type: 'paragraph',
        runs: runsFromInline(node.content),
        format: base,
      })
      break
    case 'heading': {
      const level = Math.min(Math.max(Number(node.attrs?.level) || 1, 1), 6) as number
      pushParagraph(ctx, {
        type: 'heading',
        level,
        runs: runsFromInline(node.content),
        format: base,
      })
      break
    }
    case 'bulletList':
      walkList(ctx, node, 'bullet', 0, BLANK_BULLET_NUM_ID, base)
      break
    case 'orderedList':
      walkList(ctx, node, 'ordered', 0, allocOrderedNumId(ctx), base)
      break
    case 'taskList':
      walkList(ctx, node, 'bullet', 0, BLANK_BULLET_NUM_ID, base)
      break
    case 'blockquote':
      for (const child of node.content ?? []) {
        walkBlock(ctx, child, mergeFormat(base, { indentLeft: INDENT_STEP, borders: 'l' }))
      }
      break
    case 'codeBlock':
      pushParagraph(ctx, {
        type: 'paragraph',
        runs: [{ text: plainText(node), font: CODE_FONT, sizeHalfPoints: 19 }],
        format: mergeFormat(base, { shadingFill: CODE_FILL }),
      })
      break
    case 'horizontalRule':
      pushParagraph(ctx, {
        type: 'paragraph',
        runs: [],
        format: mergeFormat(base, { borders: 'b' }),
      })
      break
    case 'image': {
      const src = String(node.attrs?.src ?? '')
      const alt = String(node.attrs?.alt ?? '')
      // placeholder now, replaced by the loaded image (or alt text) after the async pass
      ctx.pendingImages.push({ index: ctx.blocks.length, src, alt })
      ctx.blocks.push({ kind: 'generated', block: { type: 'paragraph', runs: [] } })
      break
    }
    case 'table':
      ctx.blocks.push({ kind: 'xml', xml: generateTableModelXml(mapTable(node)) })
      break
    default: {
      // unknown block: keep its text so nothing silently disappears
      const text = plainText(node)
      if (text.trim()) pushParagraph(ctx, { type: 'paragraph', runs: [{ text }], format: base })
    }
  }
}

/** Map a ProseMirror document to docx-engine SaveBlocks (pure except image loading) */
export async function mapDocToSaveBlocks(
  doc: JSONContent,
  loadImage: ImageLoader,
): Promise<DocxMapping> {
  const ctx: WalkContext = {
    blocks: [],
    restartNums: [],
    nextOrderedNumId: 100,
    loadImage,
    pendingImages: [],
  }
  for (const node of doc.content ?? []) walkBlock(ctx, node)

  for (const pending of ctx.pendingImages) {
    const image = await loadImage(pending.src).catch(() => null)
    if (image) {
      ctx.blocks[pending.index] = { kind: 'image', image }
    } else {
      ctx.blocks[pending.index] = {
        kind: 'generated',
        block: {
          type: 'paragraph',
          runs: [{ text: `[${pending.alt || pending.src}]`, italic: true, color: '888888' }],
        },
      }
    }
  }

  const options: SaveOptions =
    ctx.restartNums.length > 0 ? { numbering: { restartNums: ctx.restartNums } } : {}
  return { blocks: ctx.blocks, options }
}

/** md document (PM JSON) → brand-new .docx bytes, fully local */
export async function exportDocxBytes(
  doc: JSONContent,
  loadImage: ImageLoader,
): Promise<Uint8Array> {
  const mapping = await mapDocToSaveBlocks(doc, loadImage)
  const parsed = await parseDocx(await buildBlankDocx())
  return saveDocx(parsed, mapping.blocks, mapping.options)
}
