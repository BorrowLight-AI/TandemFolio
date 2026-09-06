import {
  PDFArray,
  PDFBool,
  PDFDict,
  PDFDocument,
  PDFDropdown,
  PDFHexString,
  PDFName,
  PDFNumber,
  PDFOptionList,
  PDFRef,
  PDFString,
  degrees,
  rgb,
} from 'pdf-lib'
import type { PDFPage } from 'pdf-lib'
import { VISUAL_SIGNATURE_CONTENT_PREFIX } from '../shared/ipc'
import { applyAnnotDeletes } from './annot-delete'
import { applyBrowserImageEdits } from './browser-image-edit'
import { applyBrowserTextEdits, applyBrowserTextInserts } from './browser-text-edit'
import type {
  DrawingInput,
  FormValueInput,
  ImageEditFailure,
  MarkupInput,
  MetadataInput,
  SavePdfRequest,
  StampInput,
  StaticFormFillRecord,
  TextEditFailure,
  TextInsertFailure,
  NoteReplyTarget,
} from '../shared/ipc'

const num = (v: number) => Math.round(v * 100) / 100
const STATIC_FORM_FILLS_KEY = PDFName.of('GenOfficeStaticFormFills')
const GENERATED_STAMP_KEY = PDFName.of('GenOfficeGeneratedStamp')

function pdfDateString(ms: number): string {
  const date = new Date(ms)
  const padded = (value: number) => String(value).padStart(2, '0')
  const offset = -date.getTimezoneOffset()
  const sign = offset >= 0 ? '+' : '-'
  const absoluteOffset = Math.abs(offset)
  return (
    `D:${date.getFullYear()}${padded(date.getMonth() + 1)}${padded(date.getDate())}` +
    `${padded(date.getHours())}${padded(date.getMinutes())}${padded(date.getSeconds())}` +
    `${sign}${padded(Math.floor(absoluteOffset / 60))}'${padded(absoluteOffset % 60)}'`
  )
}

const NOTE_RECT_TOLERANCE = 2

function noteRectsClose(left: readonly number[], right: readonly number[]): boolean {
  return (
    Math.abs(Math.min(left[0]!, left[2]!) - Math.min(right[0]!, right[2]!)) <=
      NOTE_RECT_TOLERANCE &&
    Math.abs(Math.max(left[1]!, left[3]!) - Math.max(right[1]!, right[3]!)) <=
      NOTE_RECT_TOLERANCE
  )
}

function findNoteAnnotationRef(
  document: PDFDocument,
  page: PDFPage,
  target: NoteReplyTarget,
): PDFRef | null {
  const annotations = page.node.lookupMaybe(PDFName.of('Annots'), PDFArray)
  if (!annotations) return null
  const matches: PDFRef[] = []
  for (let index = 0; index < annotations.size(); index += 1) {
    const reference = annotations.get(index)
    if (!(reference instanceof PDFRef)) continue
    const annotation = document.context.lookupMaybe(reference, PDFDict)
    if (
      !annotation ||
      annotation.lookupMaybe(PDFName.of('Subtype'), PDFName) !== PDFName.of('Text')
    ) {
      continue
    }
    const rectangle = annotation.lookupMaybe(PDFName.of('Rect'), PDFArray)
    if (!rectangle || rectangle.size() !== 4) continue
    const rect = Array.from({ length: 4 }, (_, item) =>
      rectangle.lookup(item, PDFNumber).asNumber(),
    )
    if (!noteRectsClose(rect, target.rect)) continue
    const contents = annotation.lookup(PDFName.of('Contents'))
    const text =
      contents instanceof PDFString || contents instanceof PDFHexString
        ? contents.decodeText()
        : ''
    if (text !== target.contents) continue
    if (reference.objectNumber === target.objNum) return reference
    matches.push(reference)
  }
  return matches[0] ?? null
}

function validStaticFormFill(value: unknown): value is StaticFormFillRecord {
  if (!value || typeof value !== 'object') return false
  const record = value as Partial<StaticFormFillRecord>
  return (
    typeof record.id === 'string' &&
    (record.kind === 'text' || record.kind === 'check' || record.kind === 'cross') &&
    Number.isInteger(record.pageIndex) &&
    Array.isArray(record.rect) &&
    record.rect.length === 4 &&
    record.rect.every((entry) => typeof entry === 'number' && Number.isFinite(entry))
  )
}

export async function readStaticFormFills(bytes: Uint8Array): Promise<StaticFormFillRecord[]> {
  const pdfDoc = await PDFDocument.load(bytes, { updateMetadata: false })
  const value = pdfDoc.catalog.get(STATIC_FORM_FILLS_KEY)
  if (!(value instanceof PDFHexString)) return []
  try {
    const parsed: unknown = JSON.parse(value.decodeText())
    return Array.isArray(parsed) ? parsed.filter(validStaticFormFill) : []
  } catch {
    return []
  }
}

function resultingStaticFormFills(
  request: SavePdfRequest,
  pageCount: number,
): StaticFormFillRecord[] | undefined {
  if (request.staticFormFills === undefined) return undefined
  const deleted = new Set(request.deletedPages ?? [])
  const remaining =
    request.pageOrder?.filter((pageIndex) => !deleted.has(pageIndex)) ??
    Array.from({ length: pageCount }, (_, pageIndex) => pageIndex).filter(
      (pageIndex) => !deleted.has(pageIndex),
    )
  const newPageIndex = new Map(remaining.map((oldPageIndex, index) => [oldPageIndex, index]))
  return request.staticFormFills.flatMap((record) => {
    const pageIndex = newPageIndex.get(record.pageIndex)
    return pageIndex === undefined ? [] : [{ ...record, pageIndex }]
  })
}

function setVisualSignatureMetadata(annot: PDFDict, fieldName: string | undefined): void {
  if (!fieldName) return
  annot.set(PDFName.of('GenOfficeFormField'), PDFHexString.fromText(fieldName))
  annot.set(
    PDFName.of('Contents'),
    PDFHexString.fromText(`${VISUAL_SIGNATURE_CONTENT_PREFIX}${fieldName}`),
  )
}

const quadBounds = (q: number[]) => {
  const xs = [q[0]!, q[2]!, q[4]!, q[6]!]
  const ys = [q[1]!, q[3]!, q[5]!, q[7]!]
  return [Math.min(...xs), Math.min(...ys), Math.max(...xs), Math.max(...ys)] as const
}

/**
 * Hand-written appearance stream (/AP /N), so viewers don't self-draw from QuadPoints —
 * Acrobat/Preview/pdfjs all render from AP for consistent results.
 * Highlight uses Multiply blending to mimic a highlighter; underline/strikeout are stroked
 * segments drawn along the "visual bottom edge" (pageRot is the page's final /Rotate;
 * at 90/270 line height runs along the x axis).
 */
function markupAppearance(
  pdfDoc: PDFDocument,
  m: MarkupInput,
  rect: number[],
  pageRot: number,
): ReturnType<typeof pdfDoc.context.stream> {
  const [r, g, b] = m.color
  const ops: string[] = []
  if (m.type === 'highlight') {
    ops.push('/GsM gs', `${r} ${g} ${b} rg`)
    for (const q of m.quads) {
      const [x1, y1, x2, y2] = quadBounds(q)
      ops.push(`${num(x1)} ${num(y1)} ${num(x2 - x1)} ${num(y2 - y1)} re f`)
    }
  } else {
    const t = m.type === 'underline' ? 0.08 : 0.46
    ops.push(`${r} ${g} ${b} RG`)
    for (const q of m.quads) {
      const [x1, y1, x2, y2] = quadBounds(q)
      const h = pageRot % 180 === 0 ? y2 - y1 : x2 - x1
      ops.push(`${Math.max(0.8, num(h * 0.06))} w`)
      if (pageRot === 90) {
        const x = x2 - h * t
        ops.push(`${num(x)} ${num(y1)} m ${num(x)} ${num(y2)} l S`)
      } else if (pageRot === 270) {
        const x = x1 + h * t
        ops.push(`${num(x)} ${num(y1)} m ${num(x)} ${num(y2)} l S`)
      } else {
        const y = pageRot === 180 ? y2 - h * t : y1 + h * t
        ops.push(`${num(x1)} ${num(y)} m ${num(x2)} ${num(y)} l S`)
      }
    }
  }
  return pdfDoc.context.stream(ops.join('\n'), {
    Type: 'XObject',
    Subtype: 'Form',
    BBox: rect,
    Resources:
      m.type === 'highlight'
        ? { ExtGState: { GsM: { Type: 'ExtGState', BM: 'Multiply', ca: 1 } } }
        : {},
  })
}

const SUBTYPE: Record<MarkupInput['type'], string> = {
  highlight: 'Highlight',
  underline: 'Underline',
  strikeout: 'StrikeOut',
}

function addMarkup(pdfDoc: PDFDocument, page: PDFPage, m: MarkupInput): void {
  const xs = m.quads.flatMap((q) => [q[0]!, q[2]!, q[4]!, q[6]!])
  const ys = m.quads.flatMap((q) => [q[1]!, q[3]!, q[5]!, q[7]!])
  const rect = [Math.min(...xs), Math.min(...ys), Math.max(...xs), Math.max(...ys)]
  const pageRot = ((page.getRotation().angle % 360) + 360) % 360
  const apRef = pdfDoc.context.register(markupAppearance(pdfDoc, m, rect, pageRot))
  const annot = pdfDoc.context.obj({
    Type: 'Annot',
    Subtype: SUBTYPE[m.type],
    Rect: rect,
    QuadPoints: m.quads.flat(),
    C: m.color,
    F: 4, // print
    T: 'GenOffice',
    P: page.ref,
    AP: { N: apRef },
  })
  appendAnnot(pdfDoc, page, pdfDoc.context.register(annot))
}

function appendAnnot(pdfDoc: PDFDocument, page: PDFPage, annotRef: PDFRef): void {
  const existing = page.node.lookupMaybe(PDFName.of('Annots'), PDFArray)
  if (existing) {
    existing.push(annotRef)
  } else {
    page.node.set(PDFName.of('Annots'), pdfDoc.context.obj([annotRef]))
  }
}

/** 4-segment Bezier approximation of an ellipse */
function ellipseOps(x1: number, y1: number, x2: number, y2: number): string[] {
  const k = 0.5522847
  const cx = (x1 + x2) / 2
  const cy = (y1 + y2) / 2
  const rx = Math.abs(x2 - x1) / 2
  const ry = Math.abs(y2 - y1) / 2
  const n = num
  return [
    `${n(cx + rx)} ${n(cy)} m`,
    `${n(cx + rx)} ${n(cy + k * ry)} ${n(cx + k * rx)} ${n(cy + ry)} ${n(cx)} ${n(cy + ry)} c`,
    `${n(cx - k * rx)} ${n(cy + ry)} ${n(cx - rx)} ${n(cy + k * ry)} ${n(cx - rx)} ${n(cy)} c`,
    `${n(cx - rx)} ${n(cy - k * ry)} ${n(cx - k * rx)} ${n(cy - ry)} ${n(cx)} ${n(cy - ry)} c`,
    `${n(cx + k * rx)} ${n(cy - ry)} ${n(cx + rx)} ${n(cy - k * ry)} ${n(cx + rx)} ${n(cy)} c`,
    'S',
  ]
}

/**
 * Image signature/stamp: a Stamp annotation whose appearance stream draws the embedded PNG.
 * The image is counter-rotated against the page's final /Rotate (viewers rotate annotation
 * appearances with the page), so it displays upright — matching the renderer preview.
 */
async function addImageStamp(
  pdfDoc: PDFDocument,
  page: PDFPage,
  d: Extract<DrawingInput, { kind: 'image' }>,
): Promise<void> {
  const png = await pdfDoc.embedPng(d.image)
  const [x1, y1, x2, y2] = d.rect
  const rw = x2 - x1
  const rh = y2 - y1
  const rot = ((page.getRotation().angle % 360) + 360) % 360
  // cm matrix mapping the image unit square into the BBox, pre-counter-rotated for the page
  const cm =
    rot === 90
      ? `0 ${num(rh)} ${num(-rw)} 0 ${num(rw)} 0`
      : rot === 180
        ? `${num(-rw)} 0 0 ${num(-rh)} ${num(rw)} ${num(rh)}`
        : rot === 270
          ? `0 ${num(-rh)} ${num(rw)} 0 0 ${num(rh)}`
          : `${num(rw)} 0 0 ${num(rh)} 0 0`
  const ap = pdfDoc.context.stream(`q ${cm} cm /Im0 Do Q`, {
    Type: 'XObject',
    Subtype: 'Form',
    BBox: [0, 0, num(rw), num(rh)],
    Resources: { XObject: { Im0: png.ref } },
  })
  const annot = pdfDoc.context.obj({
    Type: 'Annot',
    Subtype: 'Stamp',
    Rect: [num(x1), num(y1), num(x2), num(y2)],
    F: 4,
    P: page.ref,
    AP: { N: pdfDoc.context.register(ap) },
  })
  annot.set(PDFName.of('T'), PDFHexString.fromText('GenOffice'))
  setVisualSignatureMetadata(annot, d.formFieldName)
  appendAnnot(pdfDoc, page, pdfDoc.context.register(annot))
}

async function addGeneratedStamp(
  pdfDoc: PDFDocument,
  page: PDFPage,
  stamp: StampInput,
): Promise<void> {
  const png = await pdfDoc.embedPng(stamp.image)
  const [x1, y1, x2, y2] = stamp.rect
  const width = x2 - x1
  const height = y2 - y1
  const opacity = Math.min(1, Math.max(0, stamp.opacity ?? 1))
  const appearance = pdfDoc.context.stream(
    `q /Gs0 gs ${num(width)} 0 0 ${num(height)} 0 0 cm /Im0 Do Q`,
    {
      Type: 'XObject',
      Subtype: 'Form',
      BBox: [0, 0, num(width), num(height)],
      Resources: {
        XObject: { Im0: png.ref },
        ExtGState: { Gs0: { Type: 'ExtGState', ca: opacity, CA: opacity } },
      },
    },
  )
  const annotation = pdfDoc.context.obj({
    Type: 'Annot',
    Subtype: 'Stamp',
    Rect: [num(x1), num(y1), num(x2), num(y2)],
    F: 4,
    P: page.ref,
    AP: { N: pdfDoc.context.register(appearance) },
  })
  annotation.set(PDFName.of('T'), PDFHexString.fromText('GenOffice'))
  annotation.set(GENERATED_STAMP_KEY, PDFBool.True)
  appendAnnot(pdfDoc, page, pdfDoc.context.register(annotation))
}

function removeGeneratedStamps(pdfDoc: PDFDocument): void {
  for (const page of pdfDoc.getPages()) {
    const annotations = page.node.lookupMaybe(PDFName.of('Annots'), PDFArray)
    if (!annotations) continue
    for (let index = annotations.size() - 1; index >= 0; index -= 1) {
      const annotation = annotations.lookupMaybe(index, PDFDict)
      if (annotation?.lookupMaybe(GENERATED_STAMP_KEY, PDFBool)?.asBoolean()) {
        annotations.remove(index)
      }
    }
  }
}

/** Drawing annots: hand-written AP for Ink/Square/Circle/Line; notes are standard Text annots (viewer draws the icon) */
function addDrawing(
  pdfDoc: PDFDocument,
  page: PDFPage,
  d: DrawingInput,
  noteRefs?: Map<string, PDFRef>,
): void {
  if (d.kind === 'image') return // handled by addImageStamp (needs async embed)
  const [r, g, b] = d.color

  if (d.kind === 'note') {
    const [x, y] = d.at
    const annot = pdfDoc.context.obj({
      Type: 'Annot',
      Subtype: 'Text',
      Rect: [num(x), num(y - 18), num(x + 20), num(y)],
      Name: 'Comment',
      C: d.color,
      F: 4,
      P: page.ref,
    })
    annot.set(PDFName.of('Contents'), PDFHexString.fromText(d.contents))
    annot.set(PDFName.of('T'), PDFHexString.fromText(d.author || 'GenOffice'))
    const when = pdfDateString(d.createdMs ?? Date.now())
    annot.set(PDFName.of('CreationDate'), PDFString.of(when))
    annot.set(PDFName.of('M'), PDFString.of(when))
    const parentRef = d.replyToLocalId
      ? (noteRefs?.get(d.replyToLocalId) ?? null)
      : d.replyToSaved
        ? findNoteAnnotationRef(pdfDoc, page, d.replyToSaved)
        : null
    if (parentRef) {
      annot.set(PDFName.of('IRT'), parentRef)
      annot.set(PDFName.of('RT'), PDFName.of('R'))
    }
    const reference = pdfDoc.context.register(annot)
    if (d.localId) noteRefs?.set(d.localId, reference)
    appendAnnot(pdfDoc, page, reference)
    return
  }

  const ops: string[] = [`${num(d.width)} w 1 J 1 j ${r} ${g} ${b} RG`]
  let xs: number[] = []
  let ys: number[] = []
  let subtype: string

  if (d.kind === 'ink') {
    subtype = 'Ink'
    for (const path of d.paths) {
      if (path.length < 4) continue
      ops.push(`${num(path[0]!)} ${num(path[1]!)} m`)
      for (let i = 2; i < path.length; i += 2) ops.push(`${num(path[i]!)} ${num(path[i + 1]!)} l`)
      ops.push('S')
      for (let i = 0; i < path.length; i += 2) {
        xs.push(path[i]!)
        ys.push(path[i + 1]!)
      }
    }
  } else if (d.kind === 'rect' || d.kind === 'ellipse') {
    const [x1, y1, x2, y2] = d.rect
    subtype = d.kind === 'rect' ? 'Square' : 'Circle'
    if (d.kind === 'rect') ops.push(`${num(x1)} ${num(y1)} ${num(x2 - x1)} ${num(y2 - y1)} re S`)
    else ops.push(...ellipseOps(x1, y1, x2, y2))
    xs = [x1, x2]
    ys = [y1, y2]
  } else {
    const [fx, fy] = d.from
    const [tx, ty] = d.to
    subtype = 'Line'
    ops.push(`${num(fx)} ${num(fy)} m ${num(tx)} ${num(ty)} l S`)
    xs = [fx, tx]
    ys = [fy, ty]
    if (d.kind === 'arrow') {
      const ang = Math.atan2(ty - fy, tx - fx)
      const len = Math.max(9, d.width * 4.5)
      for (const off of [-0.45, 0.45]) {
        const hx = tx - len * Math.cos(ang + off)
        const hy = ty - len * Math.sin(ang + off)
        ops.push(`${num(tx)} ${num(ty)} m ${num(hx)} ${num(hy)} l S`)
        xs.push(hx)
        ys.push(hy)
      }
    }
  }

  const pad = d.width + 2
  const rect = [
    Math.min(...xs) - pad,
    Math.min(...ys) - pad,
    Math.max(...xs) + pad,
    Math.max(...ys) + pad,
  ]
  const ap = pdfDoc.context.stream(ops.join('\n'), { Type: 'XObject', Subtype: 'Form', BBox: rect })
  const annot = pdfDoc.context.obj({
    Type: 'Annot',
    Subtype: subtype,
    Rect: rect,
    C: d.color,
    F: 4,
    P: page.ref,
    BS: { W: d.width },
    AP: { N: pdfDoc.context.register(ap) },
  })
  if (d.kind === 'ink') annot.set(PDFName.of('InkList'), pdfDoc.context.obj(d.paths))
  if (d.kind === 'line' || d.kind === 'arrow') {
    annot.set(PDFName.of('L'), pdfDoc.context.obj([...d.from, ...d.to]))
  }
  annot.set(PDFName.of('T'), PDFHexString.fromText('GenOffice'))
  if (d.kind === 'ink') setVisualSignatureMetadata(annot, d.formFieldName)
  appendAnnot(pdfDoc, page, pdfDoc.context.register(annot))
}

function applyFormValues(pdfDoc: PDFDocument, values: FormValueInput[]): void {
  const form = pdfDoc.getForm()
  for (const v of values) {
    if (v.kind === 'text') {
      form.getTextField(v.name).setText(v.value ?? '')
    } else if (v.kind === 'radio') {
      const rg = form.getRadioGroup(v.name)
      if (v.value) rg.select(v.value)
      else rg.clear()
    } else if (v.kind === 'choice') {
      const f = form.getField(v.name)
      if (f instanceof PDFDropdown || f instanceof PDFOptionList) {
        if (v.value) f.select(v.value)
        else f.clear()
      }
    } else {
      const cb = form.getCheckBox(v.name)
      if (v.checked) cb.check()
      else cb.uncheck()
    }
  }
}

/** Extract the given pages (original indices) into bytes of a new PDF */
export async function extractPagesBytes(bytes: Uint8Array, pages: number[]): Promise<Uint8Array> {
  const src = await PDFDocument.load(bytes, { updateMetadata: false })
  const out = await PDFDocument.create()
  const valid = pages.filter((p) => p >= 0 && p < src.getPageCount())
  const copied = await out.copyPages(src, valid)
  for (const p of copied) out.addPage(p)
  return out.save({ useObjectStreams: false })
}

/** Insert all pages of another PDF after afterPageIndex (-1 = front); returns merged bytes and inserted page count */
export async function insertPdfBytes(
  bytes: Uint8Array,
  otherBytes: Uint8Array,
  afterPageIndex: number,
): Promise<{ merged: Uint8Array; count: number }> {
  const dst = await PDFDocument.load(bytes, { updateMetadata: false })
  const src = await PDFDocument.load(otherBytes, { updateMetadata: false })
  const copied = await dst.copyPages(src, src.getPageIndices())
  let at = Math.min(Math.max(afterPageIndex + 1, 0), dst.getPageCount())
  for (const p of copied) dst.insertPage(at++, p)
  return { merged: await dst.save({ useObjectStreams: false }), count: copied.length }
}

/** Insert one blank page after the given page, matching its size and rotation. */
export async function insertBlankPageBytes(
  bytes: Uint8Array,
  afterPageIndex: number,
): Promise<Uint8Array> {
  const document = await PDFDocument.load(bytes, { updateMetadata: false })
  const insertAt = Math.min(Math.max(afterPageIndex + 1, 0), document.getPageCount())
  const neighbor = document.getPage(
    Math.min(Math.max(afterPageIndex, 0), document.getPageCount() - 1),
  )
  const page = document.insertPage(insertAt, [neighbor.getWidth(), neighbor.getHeight()])
  page.setRotation(neighbor.getRotation())
  return document.save({ useObjectStreams: false })
}

/** Split into consecutive page chunks while preserving source page order. */
export async function splitPdfBytes(
  bytes: Uint8Array,
  chunkSize: number,
): Promise<Uint8Array[]> {
  const source = await PDFDocument.load(bytes, { updateMetadata: false })
  const size = Math.max(1, Math.floor(chunkSize))
  const parts: Uint8Array[] = []
  for (let start = 0; start < source.getPageCount(); start += size) {
    const output = await PDFDocument.create()
    const count = Math.min(size, source.getPageCount() - start)
    const pages = await output.copyPages(
      source,
      Array.from({ length: count }, (_, offset) => start + offset),
    )
    for (const page of pages) output.addPage(page)
    parts.push(await output.save({ useObjectStreams: false }))
  }
  return parts
}

export interface MergePagesOptions {
  perSheet: number
  direction: 'horizontal' | 'vertical'
  separator: boolean
}

export function mergeGrid(perSheet: number): { cols: number; rows: number } {
  const count = Math.min(Math.max(Math.floor(perSheet), 2), 16)
  if (count === 2) return { cols: 2, rows: 1 }
  const cols = Math.ceil(Math.sqrt(count))
  return { cols, rows: Math.ceil(count / cols) }
}

/** Impose consecutive pages on sheets while retaining their vector content. */
export async function mergePagesBytes(
  bytes: Uint8Array,
  options: MergePagesOptions,
): Promise<Uint8Array> {
  const perSheet = Math.min(Math.max(Math.floor(options.perSheet), 2), 16)
  const source = await PDFDocument.load(bytes, { updateMetadata: false })
  const output = await PDFDocument.create()
  const first = source.getPage(0)
  const { cols, rows } = mergeGrid(perSheet)
  const sheetWidth = perSheet === 2 ? first.getHeight() : first.getWidth()
  const sheetHeight = perSheet === 2 ? first.getWidth() : first.getHeight()
  for (const page of source.getPages()) {
    if (!page.node.Contents()) {
      page.node.set(PDFName.of('Contents'), source.context.register(source.context.stream('')))
    }
  }
  const embedded = await output.embedPages(source.getPages())
  const cellWidth = sheetWidth / cols
  const cellHeight = sheetHeight / rows
  for (let start = 0; start < embedded.length; start += perSheet) {
    const sheet = output.addPage([sheetWidth, sheetHeight])
    for (let item = 0; item < perSheet && start + item < embedded.length; item += 1) {
      const page = embedded[start + item]!
      const scale = Math.min(cellWidth / page.width, cellHeight / page.height)
      const width = page.width * scale
      const height = page.height * scale
      const col = options.direction === 'vertical' ? Math.floor(item / rows) : item % cols
      const row = options.direction === 'vertical' ? item % rows : Math.floor(item / cols)
      sheet.drawPage(page, {
        x: col * cellWidth + (cellWidth - width) / 2,
        y: sheetHeight - (row + 1) * cellHeight + (cellHeight - height) / 2,
        width,
        height,
      })
    }
    if (options.separator) {
      const line = { thickness: 0.75, color: rgb(0.62, 0.62, 0.62) }
      for (let col = 1; col < cols; col += 1) {
        sheet.drawLine({
          start: { x: col * cellWidth, y: 0 },
          end: { x: col * cellWidth, y: sheetHeight },
          ...line,
        })
      }
      for (let row = 1; row < rows; row += 1) {
        sheet.drawLine({
          start: { x: 0, y: row * cellHeight },
          end: { x: sheetWidth, y: row * cellHeight },
          ...line,
        })
      }
    }
  }
  return output.save({ useObjectStreams: false })
}

/** Resize all pages to one paper size, scaling content and annotations together. */
export async function setPageSizeBytes(
  bytes: Uint8Array,
  targetWidth: number,
  targetHeight: number,
): Promise<Uint8Array> {
  const document = await PDFDocument.load(bytes, { updateMetadata: false })
  for (const page of document.getPages()) {
    const rotation = ((page.getRotation().angle % 360) + 360) % 360
    const width = rotation === 90 || rotation === 270 ? targetHeight : targetWidth
    const height = rotation === 90 || rotation === 270 ? targetWidth : targetHeight
    const media = page.getMediaBox()
    if (media.width === width && media.height === height) continue
    const scale = Math.min(width / media.width, height / media.height)
    page.scaleContent(scale, scale)
    page.scaleAnnotations(scale, scale)
    const x = media.x * scale - (width - media.width * scale) / 2
    const y = media.y * scale - (height - media.height * scale) / 2
    page.setMediaBox(x, y, width, height)
    page.setCropBox(x, y, width, height)
  }
  return document.save({ useObjectStreams: false })
}

function displayFractionToUserRect(
  rotation: number,
  box: { x: number; y: number; width: number; height: number },
  left: number,
  top: number,
  right: number,
  bottom: number,
): { x: number; y: number; width: number; height: number } {
  const { x, y, width, height } = box
  if (rotation === 90) {
    return {
      x: x + top * width,
      y: y + left * height,
      width: (bottom - top) * width,
      height: (right - left) * height,
    }
  }
  if (rotation === 180) {
    return {
      x: x + (1 - right) * width,
      y: y + top * height,
      width: (right - left) * width,
      height: (bottom - top) * height,
    }
  }
  if (rotation === 270) {
    return {
      x: x + (1 - bottom) * width,
      y: y + (1 - right) * height,
      width: (bottom - top) * width,
      height: (right - left) * height,
    }
  }
  return {
    x: x + left * width,
    y: y + (1 - bottom) * height,
    width: (right - left) * width,
    height: (bottom - top) * height,
  }
}

export interface CropFractionsRect {
  l: number
  t: number
  r: number
  b: number
}

/** Apply a displayed-page crop rectangle to selected pages without rewriting their content. */
export async function cropPagesBytes(
  bytes: Uint8Array,
  pages: number[],
  fraction: CropFractionsRect,
): Promise<Uint8Array> {
  const document = await PDFDocument.load(bytes, { updateMetadata: false })
  const left = Math.min(Math.max(fraction.l, 0), 1)
  const top = Math.min(Math.max(fraction.t, 0), 1)
  const right = Math.min(Math.max(fraction.r, left), 1)
  const bottom = Math.min(Math.max(fraction.b, top), 1)
  if (right - left <= 0 || bottom - top <= 0) throw new Error('cropPages: empty crop rect')
  for (const pageIndex of pages) {
    if (pageIndex < 0 || pageIndex >= document.getPageCount()) continue
    const page = document.getPage(pageIndex)
    const rotation = ((page.getRotation().angle % 360) + 360) % 360
    const rect = displayFractionToUserRect(
      rotation,
      page.getCropBox(),
      left,
      top,
      right,
      bottom,
    )
    page.setCropBox(rect.x, rect.y, rect.width, rect.height)
  }
  return document.save({ useObjectStreams: false })
}

/** Split every source page into a displayed grid, preserving vector page content. */
export async function splitPagesBytes(
  bytes: Uint8Array,
  perPage: 2 | 4 | 9,
): Promise<Uint8Array> {
  const source = await PDFDocument.load(bytes, { updateMetadata: false })
  const output = await PDFDocument.create()
  const { cols, rows } = mergeGrid(perPage)
  for (let sourceIndex = 0; sourceIndex < source.getPageCount(); sourceIndex += 1) {
    const copies = await output.copyPages(
      source,
      Array.from({ length: perPage }, () => sourceIndex),
    )
    for (let cell = 0; cell < perPage; cell += 1) {
      const page = copies[cell]!
      const rotation = ((page.getRotation().angle % 360) + 360) % 360
      const col = cell % cols
      const row = Math.floor(cell / cols)
      const rect = displayFractionToUserRect(
        rotation,
        page.getCropBox(),
        col / cols,
        row / rows,
        (col + 1) / cols,
        (row + 1) / rows,
      )
      page.setMediaBox(rect.x, rect.y, rect.width, rect.height)
      page.setCropBox(rect.x, rect.y, rect.width, rect.height)
      output.addPage(page)
    }
  }
  return output.save({ useObjectStreams: false })
}

/** Replace selected pages with all pages from another PDF at the first selected position. */
export async function replacePagesBytes(
  bytes: Uint8Array,
  otherBytes: Uint8Array,
  pages: number[],
): Promise<{ merged: Uint8Array; removed: number; inserted: number }> {
  const destination = await PDFDocument.load(bytes, { updateMetadata: false })
  const source = await PDFDocument.load(otherBytes, { updateMetadata: false })
  const valid = [...new Set(pages.filter((page) => page >= 0 && page < destination.getPageCount()))]
    .sort((left, right) => left - right)
  if (valid.length === 0) throw new Error('replacePages: no valid pages to replace')
  const insertAt = valid[0]!
  for (const page of [...valid].reverse()) destination.removePage(page)
  const copied = await destination.copyPages(source, source.getPageIndices())
  let index = Math.min(insertAt, destination.getPageCount())
  for (const page of copied) destination.insertPage(index++, page)
  return {
    merged: await destination.save({ useObjectStreams: false }),
    removed: valid.length,
    inserted: copied.length,
  }
}

/** Append all pages from the supplied PDFs in source order. */
export async function mergePdfBytes(
  first: Uint8Array,
  others: Uint8Array[],
): Promise<{ merged: Uint8Array; appended: number }> {
  const destination = await PDFDocument.load(first, { updateMetadata: false })
  let appended = 0
  for (const bytes of others) {
    const source = await PDFDocument.load(bytes, { updateMetadata: false })
    const pages = await destination.copyPages(source, source.getPageIndices())
    for (const page of pages) destination.addPage(page)
    appended += pages.length
  }
  return {
    merged: await destination.save({ useObjectStreams: false }),
    appended,
  }
}

function applyMetadata(pdfDoc: PDFDocument, meta: MetadataInput): void {
  if (meta.title !== undefined) pdfDoc.setTitle(meta.title)
  if (meta.author !== undefined) pdfDoc.setAuthor(meta.author)
  if (meta.subject !== undefined) pdfDoc.setSubject(meta.subject)
  if (meta.keywords !== undefined) {
    pdfDoc.setKeywords(
      meta.keywords
        .split(/[,，;；]/)
        .map((k) => k.trim())
        .filter(Boolean),
    )
  }
  pdfDoc.setModificationDate(new Date())
}

/**
 * Apply the request to the PDF at sourcePath and atomically write the result to targetPath
 * (temp file next to the target + rename, so a mid-write crash can't corrupt it).
 * The source file is only ever read: Save As (targetPath !== sourcePath) must never mutate
 * the original document, and a failed or cancelled save leaves both paths untouched.
 * In-place Save passes targetPath === sourcePath.
 * Returns the text edits that no longer matched the document and were skipped.
 */
export interface SavePdfSkips {
  skippedTextEdits: TextEditFailure[]
  skippedTextInserts: TextInsertFailure[]
  skippedImageEdits: ImageEditFailure[]
}

export interface AppliedSaveRequest {
  bytes: Uint8Array
  /** Text edits that could not be matched to the document; the rest of the request is in `bytes` */
  skippedTextEdits: TextEditFailure[]
  skippedTextInserts: TextInsertFailure[]
  /** Same, for content-stream image operations */
  skippedImageEdits: ImageEditFailure[]
}

/** Apply markups + form values + page ops, returning new bytes. Original objects are not reordered (pdf-lib keeps untouched objects). */
export async function applySaveRequest(
  bytes: Uint8Array,
  request: SavePdfRequest,
): Promise<AppliedSaveRequest> {
  let skippedTextEdits: TextEditFailure[] = []
  let skippedTextInserts: TextInsertFailure[] = []
  let skippedImageEdits: ImageEditFailure[] = []
  if (request.annotDeletes && request.annotDeletes.length > 0) {
    // First stage: the object numbers address the on-disk bytes; later pdfium
    // rewrites (text/image edits) may renumber objects
    bytes = await applyAnnotDeletes(bytes, request.annotDeletes)
  }
  if (request.textEdits && request.textEdits.length > 0) {
    const result = await applyBrowserTextEdits(bytes, request.textEdits)
    bytes = result.bytes
    skippedTextEdits = [...result.skipped]
  }
  if (request.textInserts && request.textInserts.length > 0) {
    const result = await applyBrowserTextInserts(bytes, request.textInserts)
    bytes = result.bytes
    skippedTextInserts = [...result.skipped]
  }
  if (request.imageEdits && request.imageEdits.length > 0) {
    const result = await applyBrowserImageEdits(bytes, request.imageEdits)
    bytes = result.bytes
    skippedImageEdits = [...result.skipped]
  }
  const pdfDoc = await PDFDocument.load(bytes, { updateMetadata: false })
  if (request.replaceGeneratedStamps) removeGeneratedStamps(pdfDoc)
  if (request.formValues.length > 0) applyFormValues(pdfDoc, request.formValues)
  const pages = pdfDoc.getPages()
  // Apply rotations first so markup appearances draw lines for the page's final orientation
  for (const r of request.rotations ?? []) {
    const page = pages[r.pageIndex]
    if (page) page.setRotation(degrees((page.getRotation().angle + r.delta) % 360))
  }
  for (const m of request.markups) {
    const page = pages[m.pageIndex]
    if (page) addMarkup(pdfDoc, page, m)
  }
  const noteRefs = new Map<string, PDFRef>()
  for (const d of request.drawings ?? []) {
    const page = pages[d.pageIndex]
    if (!page) continue
    if (d.kind === 'image') await addImageStamp(pdfDoc, page, d)
    else addDrawing(pdfDoc, page, d, noteRefs)
  }
  for (const edit of request.noteEdits ?? []) {
    const page = pages[edit.pageIndex]
    if (!page) continue
    const reference = findNoteAnnotationRef(pdfDoc, page, {
      objNum: edit.objNum,
      rect: edit.rect,
      contents: edit.oldContents,
    })
    const annotation = reference ? pdfDoc.context.lookupMaybe(reference, PDFDict) : null
    if (!annotation) continue
    annotation.set(PDFName.of('Contents'), PDFHexString.fromText(edit.contents))
    annotation.set(PDFName.of('M'), PDFString.of(pdfDateString(Date.now())))
  }
  for (const s of request.stamps ?? []) {
    const page = pages[s.pageIndex]
    if (!page) continue
    await addGeneratedStamp(pdfDoc, page, s)
  }
  if (request.metadata) applyMetadata(pdfDoc, request.metadata)
  // Deletions go last, in descending order; earlier ops all address original page indices
  for (const idx of [...(request.deletedPages ?? [])].sort((a, b) => b - a)) {
    if (idx >= 0 && idx < pdfDoc.getPageCount() && pdfDoc.getPageCount() > 1) pdfDoc.removePage(idx)
  }
  // Reorder last: pageOrder gives the new order of remaining-after-delete pages by original index.
  // pdf-lib's removePage never invalidates its page cache, so getPages() here would return the
  // stale pre-deletion list — derive the surviving pages from the pre-deletion snapshot instead.
  const order = request.pageOrder
  if (order && order.length > 0) {
    const deletedSet = new Set(request.deletedPages ?? [])
    const target = order
      .filter((o) => !deletedSet.has(o))
      .map((o) => pages[o])
      .filter((p) => p !== undefined)
    if (target.length === pdfDoc.getPageCount()) {
      while (pdfDoc.getPageCount() > 0) pdfDoc.removePage(0)
      for (const p of target) pdfDoc.addPage(p)
    }
  }
  const staticFormFills = resultingStaticFormFills(request, pages.length)
  if (staticFormFills !== undefined) {
    if (staticFormFills.length === 0) pdfDoc.catalog.delete(STATIC_FORM_FILLS_KEY)
    else
      pdfDoc.catalog.set(
        STATIC_FORM_FILLS_KEY,
        PDFHexString.fromText(JSON.stringify(staticFormFills)),
      )
  }
  try {
    return {
      bytes: await pdfDoc.save({ useObjectStreams: false }),
      skippedTextEdits,
      skippedTextInserts,
      skippedImageEdits,
    }
  } catch (err) {
    // Form values beyond WinAnsi (e.g. CJK) make pdf-lib's appearance generation fail:
    // skip it and set NeedAppearances so viewers rebuild them (Acrobat/pdfjs both support this)
    if (request.formValues.length === 0) throw err
    pdfDoc.getForm().acroForm.dict.set(PDFName.of('NeedAppearances'), PDFBool.True)
    return {
      bytes: await pdfDoc.save({ useObjectStreams: false, updateFieldAppearances: false }),
      skippedTextEdits,
      skippedTextInserts,
      skippedImageEdits,
    }
  }
}
