import { foldRadicals } from '../shared/radicals'
import { readLiveEditorBundledFontAsset } from '@tandemfolio/host-bridge'
import type {
  TextEditFailure,
  TextEditInput,
  TextEditValidation,
  TextInsertFailure,
  TextInsertInput,
} from '../shared/ipc'
import {
  chainBrowserPdfium,
  loadBrowserPdfium,
  saveBrowserPdfiumDocument,
  withBrowserPdfiumDocument,
  type BrowserPdfium,
} from './pdfium-browser'

const FPDF_PAGEOBJ_TEXT = 1
const FPDF_FONT_TYPE1 = 1
const FPDF_FONT_TRUETYPE = 2
const LINE_GAP = 1.2

type Rect = readonly [number, number, number, number]

interface PageTextObject {
  readonly object: number
  readonly index: number
  readonly text: string
  readonly bounds: [number, number, number, number]
  readonly color: [number, number, number]
}

export const normalizePdfText = (text: string): string =>
  foldRadicals(text).normalize('NFKC').replace(/\s+/g, '')

function utf16Pointer(pdfium: BrowserPdfium, value: string): number {
  const bytes = new Uint8Array((value.length + 1) * 2)
  const view = new DataView(bytes.buffer)
  for (let index = 0; index < value.length; index += 1) {
    view.setUint16(index * 2, value.charCodeAt(index), true)
  }
  const pointer = pdfium._malloc(bytes.byteLength)
  pdfium.HEAPU8.set(bytes, pointer)
  return pointer
}

function asciiPointer(pdfium: BrowserPdfium, value: string): number {
  const bytes = new TextEncoder().encode(`${value}\0`)
  const pointer = pdfium._malloc(bytes.byteLength)
  pdfium.HEAPU8.set(bytes, pointer)
  return pointer
}

function textObjectText(pdfium: BrowserPdfium, object: number, textPage: number): string {
  const length = pdfium._FPDFTextObj_GetText(object, textPage, 0, 0)
  if (length <= 2) return ''
  const pointer = pdfium._malloc(length)
  try {
    pdfium._FPDFTextObj_GetText(object, textPage, pointer, length)
    return new TextDecoder('utf-16le').decode(pdfium.HEAPU8.subarray(pointer, pointer + length - 2))
  } finally {
    pdfium._free(pointer)
  }
}

function objectColor(pdfium: BrowserPdfium, object: number): [number, number, number] {
  const pointers = [0, 1, 2, 3].map(() => pdfium._malloc(4))
  try {
    if (
      !pdfium._FPDFPageObj_GetFillColor(
        object,
        pointers[0]!,
        pointers[1]!,
        pointers[2]!,
        pointers[3]!,
      )
    )
      return [0, 0, 0]
    return [
      pdfium.HEAPU8[pointers[0]!]!,
      pdfium.HEAPU8[pointers[1]!]!,
      pdfium.HEAPU8[pointers[2]!]!,
    ]
  } finally {
    for (const pointer of pointers) pdfium._free(pointer)
  }
}

function collectTextObjects(
  pdfium: BrowserPdfium,
  page: number,
  textPage: number,
): PageTextObject[] {
  const pointers = [0, 1, 2, 3].map(() => pdfium._malloc(4))
  try {
    const objects: PageTextObject[] = []
    for (let index = 0; index < pdfium._FPDFPage_CountObjects(page); index += 1) {
      const object = pdfium._FPDFPage_GetObject(page, index)
      if (
        pdfium._FPDFPageObj_GetType(object) !== FPDF_PAGEOBJ_TEXT ||
        !pdfium._FPDFPageObj_GetBounds(
          object,
          pointers[0]!,
          pointers[1]!,
          pointers[2]!,
          pointers[3]!,
        )
      ) {
        continue
      }
      objects.push({
        object,
        index,
        text: textObjectText(pdfium, object, textPage),
        bounds: pointers.map((pointer) => pdfium.HEAPF32[pointer >> 2]!) as [
          number,
          number,
          number,
          number,
        ],
        color: objectColor(pdfium, object),
      })
    }
    return objects
  } finally {
    for (const pointer of pointers) pdfium._free(pointer)
  }
}

function overlapArea(left: Rect, right: Rect): number {
  const width = Math.min(left[2], right[2]) - Math.max(left[0], right[0])
  const height = Math.min(left[3], right[3]) - Math.max(left[1], right[1])
  return width > 0 && height > 0 ? width * height : 0
}

function matchedTextObjects(objects: readonly PageTextObject[], edit: TextEditInput): PageTextObject[] {
  const rectArea = Math.max((edit.rect[2] - edit.rect[0]) * (edit.rect[3] - edit.rect[1]), 1e-6)
  const matches = objects.filter((object) => {
    const objectArea = Math.max(
      (object.bounds[2] - object.bounds[0]) * (object.bounds[3] - object.bounds[1]),
      1e-6,
    )
    const overlap = overlapArea(object.bounds, edit.rect)
    return overlap / objectArea >= 0.5 || overlap / rectArea >= 0.8
  })
  matches.sort((left, right) => left.index - right.index)
  return normalizePdfText(matches.map((object) => object.text).join('')) ===
    normalizePdfText(edit.oldText)
    ? matches
    : []
}

function standardFontName(
  family: string | undefined,
  bold: boolean | undefined,
  italic: boolean | undefined,
): string {
  const suffix = bold && italic ? '-BoldOblique' : bold ? '-Bold' : italic ? '-Oblique' : ''
  if (family === 'times') {
    return bold && italic
      ? 'Times-BoldItalic'
      : bold
        ? 'Times-Bold'
        : italic
          ? 'Times-Italic'
          : 'Times-Roman'
  }
  if (family === 'courier') return `Courier${suffix}`
  return `Helvetica${suffix}`
}

function loadStandardFont(
  pdfium: BrowserPdfium,
  document: number,
  family: string | undefined,
  bold: boolean | undefined,
  italic: boolean | undefined,
): number {
  const pointer = asciiPointer(pdfium, standardFontName(family, bold, italic))
  try {
    const font = pdfium._FPDFText_LoadStandardFont(document, pointer)
    if (!font) throw new Error('The requested standard PDF font is unavailable.')
    return font
  } finally {
    pdfium._free(pointer)
  }
}

interface BrowserFontSpec {
  readonly key: string
  readonly cssFamily: string
  readonly fileName?: string
  readonly standardFamily?: string
  readonly bold?: boolean
  readonly italic?: boolean
}

const fontBytes = new Map<string, Promise<Uint8Array>>()
const registeredBrowserFonts = new Map<string, Promise<void>>()

const hasArabic = (character: string): boolean =>
  /[\u0600-\u077f\u08a0-\u08ff\ufb50-\ufdff\ufe70-\ufeff]/u.test(character)
const hasKorean = (character: string): boolean =>
  /[\u1100-\u11ff\u3130-\u318f\uac00-\ud7af]/u.test(character)
const needsCjk = (character: string): boolean =>
  /[\u3000-\u30ff\u3400-\u9fff\uf900-\ufaff\uff00-\uffef]/u.test(character)

function latinAssetName(
  family: string | undefined,
  bold: boolean | undefined,
  italic: boolean | undefined,
): { fileName: string; family: string } {
  const suffix = bold && italic ? 'BoldItalic' : bold ? 'Bold' : italic ? 'Italic' : 'Regular'
  if (family === 'times') {
    return { fileName: `LiberationSerif-${suffix}.ttf`, family: 'PDF Edit Liberation Serif' }
  }
  if (family === 'courier') {
    return { fileName: `LiberationMono-${suffix}.ttf`, family: 'PDF Edit Liberation Mono' }
  }
  return { fileName: `LiberationSans-${suffix}.ttf`, family: 'PDF Edit Liberation Sans' }
}

function fontSpecForCharacter(
  character: string,
  input: { readonly font?: string; readonly bold?: boolean; readonly italic?: boolean },
): BrowserFontSpec {
  if (hasArabic(character)) {
    return {
      key: 'arabic',
      cssFamily: 'PDF Edit Noto Naskh Arabic',
      fileName: 'NotoNaskhArabic-Regular-subset.ttf',
    }
  }
  if (hasKorean(character)) {
    return {
      key: 'korean',
      cssFamily: 'PDF Edit GenOffice Sans KR',
      fileName: 'GenOfficeSansKR-Regular-subset.otf',
    }
  }
  if (needsCjk(character)) {
    return {
      key: 'cjk',
      cssFamily: 'PDF Edit Noto Sans CJK SC',
      fileName: 'NotoSansCJKsc-Regular-subset.otf',
    }
  }
  if (character.codePointAt(0)! > 0x7e || input.font) {
    const asset = latinAssetName(input.font, input.bold, input.italic)
    return { key: asset.fileName, cssFamily: asset.family, fileName: asset.fileName }
  }
  return {
    key: standardFontName(input.font, input.bold, input.italic),
    cssFamily:
      input.font === 'times' ? 'Times New Roman' : input.font === 'courier' ? 'Courier New' : 'Arial',
    standardFamily: input.font,
    bold: input.bold,
    italic: input.italic,
  }
}

async function readFontBytes(fileName: string): Promise<Uint8Array> {
  let pending = fontBytes.get(fileName)
  if (!pending) {
    pending = readLiveEditorBundledFontAsset(fileName).then((data) => new Uint8Array(data))
    fontBytes.set(fileName, pending)
  }
  return pending
}

async function registerBrowserFont(spec: BrowserFontSpec, bytes: Uint8Array): Promise<void> {
  if (typeof FontFace === 'undefined' || !document.fonts) return
  let pending = registeredBrowserFonts.get(spec.key)
  if (!pending) {
    const data = Uint8Array.from(bytes).buffer
    pending = new FontFace(spec.cssFamily, data).load().then((face) => {
      document.fonts.add(face)
    })
    registeredBrowserFonts.set(spec.key, pending)
  }
  await pending
}

function readCffIndex(
  bytes: Uint8Array,
  offset: number,
): { readonly items: readonly [number, number][]; readonly end: number } {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  const count = view.getUint16(offset)
  if (count === 0) return { items: [], end: offset + 2 }
  const offSize = bytes[offset + 2]!
  const readOffset = (index: number): number => {
    let value = 0
    for (let byte = 0; byte < offSize; byte += 1) {
      value = value * 256 + bytes[offset + 3 + index * offSize + byte]!
    }
    return value
  }
  const dataStart = offset + 3 + (count + 1) * offSize - 1
  return {
    items: Array.from({ length: count }, (_, index) => [
      dataStart + readOffset(index),
      dataStart + readOffset(index + 1),
    ]),
    end: dataStart + readOffset(count),
  }
}

function parseCffDict(bytes: Uint8Array, start: number, end: number): Map<number, number[]> {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  const operations = new Map<number, number[]>()
  let operands: number[] = []
  let index = start
  while (index < end) {
    const first = bytes[index]!
    if (first <= 21) {
      let operation = first
      if (first === 12) operation = 1200 + bytes[++index]!
      operations.set(operation, operands)
      operands = []
      index += 1
    } else if (first === 28) {
      operands.push(view.getInt16(index + 1))
      index += 3
    } else if (first === 29) {
      operands.push(view.getInt32(index + 1))
      index += 5
    } else if (first === 30) {
      index += 1
      while (
        index < end &&
        (bytes[index]! & 0x0f) !== 0x0f &&
        bytes[index]! >> 4 !== 0x0f
      ) {
        index += 1
      }
      index += 1
      operands.push(Number.NaN)
    } else if (first >= 32 && first <= 246) {
      operands.push(first - 139)
      index += 1
    } else if (first >= 247 && first <= 250) {
      operands.push((first - 247) * 256 + bytes[index + 1]! + 108)
      index += 2
    } else if (first >= 251 && first <= 254) {
      operands.push(-(first - 251) * 256 - bytes[index + 1]! - 108)
      index += 2
    } else {
      index += 1
    }
  }
  return operations
}

function sfntTableOffset(font: Uint8Array, tag: string): number {
  if (font.byteLength < 12) return -1
  const view = new DataView(font.buffer, font.byteOffset, font.byteLength)
  const tableCount = view.getUint16(4)
  for (let table = 0; table < tableCount; table += 1) {
    const record = 12 + table * 16
    if (record + 16 > font.byteLength) return -1
    if (new TextDecoder('latin1').decode(font.subarray(record, record + 4)) === tag) {
      return view.getUint32(record + 8)
    }
  }
  return -1
}

/** Minimal browser-safe sfnt cmap check; retained edit fonts use format 4 or 12. */
export function browserFontCoversText(font: Uint8Array, text: string): boolean {
  try {
    const view = new DataView(font.buffer, font.byteOffset, font.byteLength)
    const cmap = sfntTableOffset(font, 'cmap')
    if (cmap < 0 || cmap + 4 > font.byteLength) return false
    const subtableCount = view.getUint16(cmap + 2)
    let selected: { offset: number; format: 4 | 12 } | null = null
    for (let index = 0; index < subtableCount; index += 1) {
      const record = cmap + 4 + index * 8
      const platform = view.getUint16(record)
      const encoding = view.getUint16(record + 2)
      if (platform !== 0 && !(platform === 3 && (encoding === 1 || encoding === 10))) continue
      const offset = cmap + view.getUint32(record + 4)
      const format = view.getUint16(offset)
      if (format === 12) {
        selected = { offset, format }
        break
      }
      if (format === 4 && !selected) selected = { offset, format }
    }
    if (!selected) return false
    const glyphId = (codePoint: number): number => {
      const offset = selected!.offset
      if (selected!.format === 12) {
        const groupCount = view.getUint32(offset + 12)
        for (let group = 0; group < groupCount; group += 1) {
          const entry = offset + 16 + group * 12
          const start = view.getUint32(entry)
          if (codePoint < start) return 0
          const end = view.getUint32(entry + 4)
          if (codePoint <= end) return view.getUint32(entry + 8) + codePoint - start
        }
        return 0
      }
      if (codePoint > 0xffff) return 0
      const segmentBytes = view.getUint16(offset + 6)
      const ends = offset + 14
      const starts = ends + segmentBytes + 2
      const deltas = starts + segmentBytes
      const rangeOffsets = deltas + segmentBytes
      for (let segment = 0; segment < segmentBytes; segment += 2) {
        if (codePoint > view.getUint16(ends + segment)) continue
        const start = view.getUint16(starts + segment)
        if (codePoint < start) return 0
        const rangeOffset = view.getUint16(rangeOffsets + segment)
        if (rangeOffset === 0) {
          return (codePoint + view.getInt16(deltas + segment)) & 0xffff
        }
        const glyph = view.getUint16(
          rangeOffsets + segment + rangeOffset + (codePoint - start) * 2,
        )
        return glyph === 0 ? 0 : (glyph + view.getInt16(deltas + segment)) & 0xffff
      }
      return 0
    }
    return [...text.replace(/[\r\n]/g, '')].every(
      (character) => glyphId(character.codePointAt(0)!) !== 0,
    )
  } catch {
    return false
  }
}

/** PDFium authors CID-keyed CFF text by GID; normalize the decompressed subset charset. */
function identityCffCharset(font: Uint8Array): Uint8Array {
  if (font.byteLength < 12) return font
  const view = new DataView(font.buffer, font.byteOffset, font.byteLength)
  if (view.getUint32(0) !== 0x4f54544f) return font
  const tableCount = view.getUint16(4)
  let cffOffset = -1
  for (let table = 0; table < tableCount; table += 1) {
    const offset = 12 + table * 16
    if (new TextDecoder('latin1').decode(font.subarray(offset, offset + 4)) === 'CFF ') {
      cffOffset = view.getUint32(offset + 8)
      break
    }
  }
  if (cffOffset < 0) return font
  const names = readCffIndex(font, cffOffset + font[cffOffset + 2]!)
  const top = readCffIndex(font, names.end).items[0]
  if (!top) return font
  const operations = parseCffDict(font, top[0], top[1])
  if (!operations.has(1230)) return font
  const charsetOffset = operations.get(15)?.[0]
  const charStringsOffset = operations.get(17)?.[0]
  if (charsetOffset === undefined || charsetOffset <= 2 || charStringsOffset === undefined) {
    return font
  }
  const glyphCount = view.getUint16(cffOffset + charStringsOffset)
  const charset = cffOffset + charsetOffset
  const format = font[charset]!
  let oldSize = 0
  if (format === 0) oldSize = 1 + 2 * (glyphCount - 1)
  else if (format === 1 || format === 2) {
    const step = format === 1 ? 3 : 4
    let covered = 1
    let cursor = charset + 1
    while (covered < glyphCount) {
      covered += (format === 1 ? font[cursor + 2]! : view.getUint16(cursor + 2)) + 1
      cursor += step
    }
    oldSize = cursor - charset
  } else return font
  const output = Uint8Array.from(font)
  const outputView = new DataView(output.buffer)
  if (glyphCount >= 2 && oldSize >= 5) {
    output[charset] = 2
    outputView.setUint16(charset + 1, 1)
    outputView.setUint16(charset + 3, glyphCount - 2)
  } else if (glyphCount >= 2 && oldSize === 4 && glyphCount - 2 <= 0xff) {
    output[charset] = 1
    outputView.setUint16(charset + 1, 1)
    output[charset + 3] = glyphCount - 2
  }
  return output
}

async function loadBrowserFont(
  pdfium: BrowserPdfium,
  document: number,
  spec: BrowserFontSpec,
): Promise<number> {
  if (!spec.fileName) {
    return loadStandardFont(pdfium, document, spec.standardFamily, spec.bold, spec.italic)
  }
  const raw = await readFontBytes(spec.fileName)
  await registerBrowserFont(spec, raw)
  const bytes = identityCffCharset(raw)
  const pointer = pdfium._malloc(bytes.byteLength)
  pdfium.HEAPU8.set(bytes, pointer)
  try {
    const isCff =
      bytes.byteLength >= 4 &&
      new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength).getUint32(0) === 0x4f54544f
    const font = pdfium._FPDFText_LoadFont(
      document,
      pointer,
      bytes.byteLength,
      isCff ? FPDF_FONT_TYPE1 : FPDF_FONT_TRUETYPE,
      1,
    )
    if (!font) throw new Error(`PDFium could not load the bundled font ${spec.fileName}.`)
    return font
  } finally {
    pdfium._free(pointer)
  }
}

async function assertBrowserFontCoverage(spec: BrowserFontSpec, text: string): Promise<void> {
  if (!spec.fileName) return
  const bytes = await readFontBytes(spec.fileName)
  if (!browserFontCoversText(bytes, text)) {
    throw new Error('The replacement contains characters no bundled PDF edit font can draw.')
  }
}

function measureSegment(text: string, fontSize: number, family: string): number {
  const canvas = document.createElement('canvas')
  const context = canvas.getContext('2d')
  if (!context) return text.length * fontSize * 0.6
  context.font = `${fontSize}px "${family}"`
  return context.measureText(text).width
}

interface TextSegment {
  readonly text: string
  readonly color: readonly [number, number, number]
  readonly font: BrowserFontSpec
}

function textSegments(
  text: string,
  offset: number,
  input: {
    readonly color: [number, number, number]
    readonly colorRuns?: readonly { start: number; end: number; color: [number, number, number] }[]
    readonly font?: string
    readonly bold?: boolean
    readonly italic?: boolean
  },
): TextSegment[] {
  const output: TextSegment[] = []
  let cursor = offset
  for (const character of text) {
    const run = input.colorRuns?.find((candidate) => cursor >= candidate.start && cursor < candidate.end)
    const color = run?.color ?? input.color
    const prior = output.at(-1)
    const font = /^\s$/u.test(character) && prior ? prior.font : fontSpecForCharacter(character, input)
    if (prior && prior.font.key === font.key && prior.color.join(',') === color.join(',')) {
      output[output.length - 1] = { ...prior, text: prior.text + character }
    } else {
      output.push({ text: character, color, font })
    }
    cursor += character.length
  }
  return output
}

export function textInsertAxes(rotation = 0): readonly [number, number, number, number] {
  switch (((rotation % 360) + 360) % 360) {
    case 90:
      return [0, 1, -1, 0]
    case 180:
      return [-1, 0, 0, -1]
    case 270:
      return [0, -1, 1, 0]
    default:
      return [1, 0, 0, 1]
  }
}

async function createTextObjects(
  pdfium: BrowserPdfium,
  document: number,
  input: {
    readonly text: string
    readonly fontSize: number
    readonly color: [number, number, number]
    readonly font?: string
    readonly bold?: boolean
    readonly italic?: boolean
    readonly origin: [number, number]
    readonly lineLeading?: number
    readonly lineXOffsets?: number[]
    readonly rotate?: number
    readonly colorRuns?: readonly {
      start: number
      end: number
      color: [number, number, number]
    }[]
  },
): Promise<number[]> {
  const matrixPointer = pdfium._malloc(24)
  const created: number[] = []
  const loadedFonts = new Map<string, number>()
  try {
    const [a, b, c, d] = textInsertAxes(input.rotate)
    const leading = input.lineLeading ?? input.fontSize * LINE_GAP
    let textOffset = 0
    for (const [lineIndex, line] of input.text.split('\n').entries()) {
      const offset = input.lineXOffsets?.[lineIndex] ?? 0
      const drop = lineIndex * leading
      let advance = 0
      for (const segment of textSegments(line, textOffset, input)) {
        await assertBrowserFontCoverage(segment.font, segment.text)
        let font = loadedFonts.get(segment.font.key)
        if (!font) {
          font = await loadBrowserFont(pdfium, document, segment.font)
          loadedFonts.set(segment.font.key, font)
        }
        const object = pdfium._FPDFPageObj_CreateTextObj(document, font, input.fontSize)
        const textPointer = utf16Pointer(pdfium, segment.text)
        const set = pdfium._FPDFText_SetText(object, textPointer)
        pdfium._free(textPointer)
        if (!set) {
          pdfium._FPDFPageObj_Destroy(object)
          throw new Error('The selected PDF font cannot encode the replacement text.')
        }
        pdfium.HEAPF32.set(
          [
            a,
            b,
            c,
            d,
            input.origin[0] + a * (offset + advance) - c * drop,
            input.origin[1] + b * (offset + advance) - d * drop,
          ],
          matrixPointer >> 2,
        )
        pdfium._FPDFPageObj_SetMatrix(object, matrixPointer)
        pdfium._FPDFPageObj_SetFillColor(object, ...segment.color, 255)
        created.push(object)
        advance += measureSegment(segment.text, input.fontSize, segment.font.cssFamily)
      }
      textOffset += line.length + 1
    }
    if (created.length === 0) throw new Error('Inserted PDF text must not be empty.')
    return created
  } catch (error) {
    for (const object of created) pdfium._FPDFPageObj_Destroy(object)
    throw error
  } finally {
    pdfium._free(matrixPointer)
  }
}

export interface BrowserTextEditResult {
  readonly bytes: Uint8Array
  readonly skipped: readonly TextEditFailure[]
}

export function applyBrowserTextEdits(
  bytes: Uint8Array,
  edits: readonly TextEditInput[],
): Promise<BrowserTextEditResult> {
  return chainBrowserPdfium(async () => {
    const pdfium = await loadBrowserPdfium()
    const skipped: TextEditFailure[] = []
    return withBrowserPdfiumDocument(pdfium, bytes, async (document) => {
      let applied = 0
      for (const edit of edits) {
        if (edit.pageIndex < 0 || edit.pageIndex >= pdfium._FPDF_GetPageCount(document)) {
          skipped.push({
            pageIndex: edit.pageIndex,
            oldText: edit.oldText,
            reason: 'page does not exist',
          })
          continue
        }
        const page = pdfium._FPDF_LoadPage(document, edit.pageIndex)
        if (!page) throw new Error(`PDFium could not load page ${edit.pageIndex + 1}.`)
        const textPage = pdfium._FPDFText_LoadPage(page)
        try {
          const matches = matchedTextObjects(collectTextObjects(pdfium, page, textPage), edit)
          if (matches.length === 0) {
            skipped.push({
              pageIndex: edit.pageIndex,
              oldText: edit.oldText,
              reason: 'the addressed text could not be matched',
            })
            continue
          }
          const canReuse =
            matches.length === 1 &&
            !edit.newFont &&
            !edit.newBold &&
            !edit.newItalic &&
            edit.newFontSize === undefined &&
            edit.lineXOffsets === undefined &&
            !edit.newText.includes('\n') &&
            /^[\x20-\x7e]*$/.test(edit.newText) &&
            (!edit.colorRuns || edit.colorRuns.length === 0)
          if (canReuse) {
            const pointer = utf16Pointer(pdfium, edit.newText)
            const set = pdfium._FPDFText_SetText(matches[0]!.object, pointer)
            pdfium._free(pointer)
            if (!set) {
              skipped.push({
                pageIndex: edit.pageIndex,
                oldText: edit.oldText,
                reason: 'the original PDF font cannot encode the replacement text',
              })
              continue
            }
            if (edit.newColor) {
              pdfium._FPDFPageObj_SetFillColor(matches[0]!.object, ...edit.newColor, 255)
            }
          } else {
            let created: number[]
            try {
              created = await createTextObjects(pdfium, document, {
                text: edit.newText,
                fontSize: edit.newFontSize ?? edit.fontSize,
                color: edit.newColor ?? matches[0]!.color,
                font: edit.newFont,
                bold: edit.newBold,
                italic: edit.newItalic,
                origin: edit.origin ?? [edit.rect[0], edit.rect[1]],
                lineLeading: edit.lineLeading,
                lineXOffsets: edit.lineXOffsets,
                colorRuns: edit.colorRuns,
              })
            } catch (error) {
              skipped.push({
                pageIndex: edit.pageIndex,
                oldText: edit.oldText,
                reason: error instanceof Error ? error.message : String(error),
              })
              continue
            }
            for (const match of matches) {
              if (pdfium._FPDFPage_RemoveObject(page, match.object)) {
                pdfium._FPDFPageObj_Destroy(match.object)
              }
            }
            for (const object of created) pdfium._FPDFPage_InsertObject(page, object)
          }
          if (!pdfium._FPDFPage_GenerateContent(page)) {
            throw new Error(`PDFium could not regenerate page ${edit.pageIndex + 1}.`)
          }
          applied += 1
        } finally {
          if (textPage) pdfium._FPDFText_ClosePage(textPage)
          pdfium._FPDF_ClosePage(page)
        }
      }
      return {
        bytes: applied > 0 ? saveBrowserPdfiumDocument(pdfium, document) : bytes,
        skipped,
      }
    })
  })
}

export interface BrowserTextInsertResult {
  readonly bytes: Uint8Array
  readonly skipped: readonly TextInsertFailure[]
}

export function applyBrowserTextInserts(
  bytes: Uint8Array,
  inserts: readonly TextInsertInput[],
): Promise<BrowserTextInsertResult> {
  return chainBrowserPdfium(async () => {
    const pdfium = await loadBrowserPdfium()
    const skipped: TextInsertFailure[] = []
    return withBrowserPdfiumDocument(pdfium, bytes, async (document) => {
      let applied = 0
      for (const [editIndex, input] of inserts.entries()) {
        if (input.pageIndex < 0 || input.pageIndex >= pdfium._FPDF_GetPageCount(document)) {
          skipped.push({ editIndex, pageIndex: input.pageIndex, reason: 'page does not exist' })
          continue
        }
        const page = pdfium._FPDF_LoadPage(document, input.pageIndex)
        if (!page) throw new Error(`PDFium could not load page ${input.pageIndex + 1}.`)
        try {
          try {
            const created = await createTextObjects(pdfium, document, input)
            for (const object of created) pdfium._FPDFPage_InsertObject(page, object)
            if (!pdfium._FPDFPage_GenerateContent(page)) {
              throw new Error(`PDFium could not regenerate page ${input.pageIndex + 1}.`)
            }
            applied += 1
          } catch (error) {
            skipped.push({
              editIndex,
              pageIndex: input.pageIndex,
              reason: error instanceof Error ? error.message : String(error),
            })
          }
        } finally {
          pdfium._FPDF_ClosePage(page)
        }
      }
      return {
        bytes: applied > 0 ? saveBrowserPdfiumDocument(pdfium, document) : bytes,
        skipped,
      }
    })
  })
}

export function validateBrowserTextEdits(
  bytes: Uint8Array,
  edits: readonly TextEditInput[],
): Promise<TextEditValidation[]> {
  return chainBrowserPdfium(async () => {
    const pdfium = await loadBrowserPdfium()
    return withBrowserPdfiumDocument(pdfium, bytes, async (document) => {
      const results: TextEditValidation[] = []
      for (const edit of edits) {
        if (edit.pageIndex < 0 || edit.pageIndex >= pdfium._FPDF_GetPageCount(document)) {
          results.push({ reason: 'page does not exist' })
          continue
        }
        const page = pdfium._FPDF_LoadPage(document, edit.pageIndex)
        if (!page) {
          results.push({ reason: 'page could not be loaded' })
          continue
        }
        const textPage = pdfium._FPDFText_LoadPage(page)
        try {
          const matches = matchedTextObjects(collectTextObjects(pdfium, page, textPage), edit)
          if (matches.length === 0) {
            results.push({ reason: 'the addressed text could not be matched' })
            continue
          }
          results.push({
            reason: null,
            bounds: [
              Math.min(...matches.map((object) => object.bounds[0])),
              Math.min(...matches.map((object) => object.bounds[1])),
              Math.max(...matches.map((object) => object.bounds[2])),
              Math.max(...matches.map((object) => object.bounds[3])),
            ],
            baseColor: matches[0]!.color,
          })
        } finally {
          if (textPage) pdfium._FPDFText_ClosePage(textPage)
          pdfium._FPDF_ClosePage(page)
        }
      }
      return results
    })
  })
}
