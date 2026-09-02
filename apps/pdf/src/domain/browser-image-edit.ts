import type {
  ImageEditFailure,
  ImageEditInput,
  PageImageRef,
  PagePreviewRequest,
} from '../shared/ipc'
import {
  chainBrowserPdfium,
  loadBrowserPdfium,
  saveBrowserPdfiumDocument,
  withBrowserPdfiumDocument,
  type BrowserPdfium,
} from './pdfium-browser'

const FPDF_PAGEOBJ_TEXT = 1
const FPDF_PAGEOBJ_IMAGE = 3
const FPDF_BITMAP_BGRA = 4
const MATCH_TOLERANCE = 2
const MAX_RENDER_SIDE = 2400

type Rect = readonly [number, number, number, number]

interface PageObject {
  readonly object: number
  readonly index: number
  readonly type: number
  readonly bounds: [number, number, number, number]
}

function collectObjects(pdfium: BrowserPdfium, page: number): PageObject[] {
  const pointers = [0, 1, 2, 3].map(() => pdfium._malloc(4))
  try {
    return Array.from({ length: pdfium._FPDFPage_CountObjects(page) }, (_, index) => {
      const object = pdfium._FPDFPage_GetObject(page, index)
      if (
        !pdfium._FPDFPageObj_GetBounds(
          object,
          pointers[0]!,
          pointers[1]!,
          pointers[2]!,
          pointers[3]!,
        )
      )
        return null
      return {
        object,
        index,
        type: pdfium._FPDFPageObj_GetType(object),
        bounds: pointers.map((pointer) => pdfium.HEAPF32[pointer >> 2]!) as [
          number,
          number,
          number,
          number,
        ],
      }
    }).filter((object): object is PageObject => object !== null)
  } finally {
    for (const pointer of pointers) pdfium._free(pointer)
  }
}

function matchImage(objects: readonly PageObject[], rect: Rect): PageObject | null {
  let match: PageObject | null = null
  let distance = Infinity
  for (const object of objects) {
    if (object.type !== FPDF_PAGEOBJ_IMAGE) continue
    const next = Math.max(...object.bounds.map((value, index) => Math.abs(value - rect[index]!)))
    if (next <= MATCH_TOLERANCE && next < distance) {
      match = object
      distance = next
    }
  }
  return match
}

const firstTextIndex = (objects: readonly PageObject[]): number | undefined =>
  objects.find((object) => object.type === FPDF_PAGEOBJ_TEXT)?.index

function base64Bytes(base64: string): Uint8Array {
  const binary = atob(base64)
  return Uint8Array.from(binary, (character) => character.charCodeAt(0))
}

async function decodePng(base64: string): Promise<{
  readonly width: number
  readonly height: number
  readonly bgra: Uint8Array
}> {
  const bytes = base64Bytes(base64)
  const data = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer
  const bitmap = await createImageBitmap(new Blob([data], { type: 'image/png' }))
  try {
    const canvas = document.createElement('canvas')
    canvas.width = bitmap.width
    canvas.height = bitmap.height
    const context = canvas.getContext('2d', { willReadFrequently: true })
    if (!context || bitmap.width < 1 || bitmap.height < 1) {
      throw new Error('The image could not be decoded.')
    }
    context.drawImage(bitmap, 0, 0)
    const rgba = context.getImageData(0, 0, bitmap.width, bitmap.height).data
    const bgra = new Uint8Array(rgba.length)
    for (let index = 0; index < rgba.length; index += 4) {
      bgra[index] = rgba[index + 2]!
      bgra[index + 1] = rgba[index + 1]!
      bgra[index + 2] = rgba[index]!
      bgra[index + 3] = rgba[index + 3]!
    }
    return { width: bitmap.width, height: bitmap.height, bgra }
  } finally {
    bitmap.close()
  }
}

function bgraToPng(
  bgra: Uint8Array,
  width: number,
  height: number,
  stride = width * 4,
): string | null {
  const canvas = document.createElement('canvas')
  canvas.width = width
  canvas.height = height
  const context = canvas.getContext('2d')
  if (!context) return null
  const rgba = new Uint8ClampedArray(width * height * 4)
  for (let row = 0; row < height; row += 1) {
    for (let column = 0; column < width; column += 1) {
      const source = row * stride + column * 4
      const target = (row * width + column) * 4
      rgba[target] = bgra[source + 2]!
      rgba[target + 1] = bgra[source + 1]!
      rgba[target + 2] = bgra[source]!
      rgba[target + 3] = bgra[source + 3]!
    }
  }
  context.putImageData(new ImageData(rgba, width, height), 0, 0)
  return canvas.toDataURL('image/png').split(',')[1] ?? null
}

function imageMatrix(rect: Rect, rotation: number): number[] {
  const [x1, y1, x2, y2] = rect
  const width = x2 - x1
  const height = y2 - y1
  switch (((rotation % 360) + 360) % 360) {
    case 90:
      return [0, height, -width, 0, x2, y1]
    case 180:
      return [-width, 0, 0, -height, x2, y2]
    case 270:
      return [0, -height, width, 0, x1, y2]
    default:
      return [width, 0, 0, height, x1, y1]
  }
}

async function setImageBitmap(
  pdfium: BrowserPdfium,
  page: number,
  object: number,
  image: string,
): Promise<void> {
  const decoded = await decodePng(image)
  const pointer = pdfium._malloc(decoded.bgra.byteLength)
  pdfium.HEAPU8.set(decoded.bgra, pointer)
  const bitmap = pdfium._FPDFBitmap_CreateEx(
    decoded.width,
    decoded.height,
    FPDF_BITMAP_BGRA,
    pointer,
    decoded.width * 4,
  )
  if (!bitmap) {
    pdfium._free(pointer)
    throw new Error('PDFium could not create an image bitmap.')
  }
  const pagesPointer = pdfium._malloc(4)
  pdfium.HEAP32[pagesPointer >> 2] = page
  try {
    if (!pdfium._FPDFImageObj_SetBitmap(pagesPointer, 1, object, bitmap)) {
      throw new Error('PDFium could not set the image bitmap.')
    }
  } finally {
    pdfium._free(pagesPointer)
    pdfium._FPDFBitmap_Destroy(bitmap)
    pdfium._free(pointer)
  }
}

function moveToLayer(
  pdfium: BrowserPdfium,
  page: number,
  object: number,
  layer: 'belowText' | 'aboveText',
): void {
  if (!pdfium._FPDFPage_RemoveObject(page, object)) {
    throw new Error('PDFium could not detach the image object.')
  }
  const textIndex = layer === 'belowText' ? firstTextIndex(collectObjects(pdfium, page)) : undefined
  if (textIndex !== undefined) {
    if (!pdfium._FPDFPage_InsertObjectAtIndex(page, object, textIndex)) {
      pdfium._FPDFPage_InsertObject(page, object)
      throw new Error('PDFium could not move the image below text.')
    }
  } else {
    pdfium._FPDFPage_InsertObject(page, object)
  }
}

const quarterTurnMatrices: ReadonlyArray<readonly [number, number, number, number]> = [
  [1, 0, 0, 1],
  [0, -1, 1, 0],
  [-1, 0, 0, -1],
  [0, 1, -1, 0],
]

function applyGeometry(
  pdfium: BrowserPdfium,
  object: number,
  bounds: Rect,
  rect: Rect,
  quarterTurns = 0,
): void {
  let [oldLeft, oldBottom, oldRight, oldTop] = bounds
  const turns = ((quarterTurns % 4) + 4) % 4
  if (turns) {
    const centerX = (oldLeft + oldRight) / 2
    const centerY = (oldBottom + oldTop) / 2
    const [a, b, c, d] = quarterTurnMatrices[turns]!
    pdfium._FPDFPageObj_Transform(
      object,
      a,
      b,
      c,
      d,
      centerX - a * centerX - c * centerY,
      centerY - b * centerX - d * centerY,
    )
    if (turns % 2 === 1) {
      const width = oldRight - oldLeft
      const height = oldTop - oldBottom
      ;[oldLeft, oldBottom, oldRight, oldTop] = [
        centerX - height / 2,
        centerY - width / 2,
        centerX + height / 2,
        centerY + width / 2,
      ]
    }
  }
  const [left, bottom, right, top] = rect
  const oldWidth = oldRight - oldLeft
  const oldHeight = oldTop - oldBottom
  if (oldWidth <= 1e-6 || oldHeight <= 1e-6) return
  const scaleX = (right - left) / oldWidth
  const scaleY = (top - bottom) / oldHeight
  pdfium._FPDFPageObj_Transform(
    object,
    scaleX,
    0,
    0,
    scaleY,
    left - scaleX * oldLeft,
    bottom - scaleY * oldBottom,
  )
}

async function applyOneEdit(
  pdfium: BrowserPdfium,
  document: number,
  page: number,
  edit: ImageEditInput,
): Promise<void> {
  if (edit.kind === 'insertImage') {
    const object = pdfium._FPDFPageObj_NewImageObj(document)
    if (!object) throw new Error('PDFium could not create an image object.')
    try {
      await setImageBitmap(pdfium, page, object, edit.image)
      const matrixPointer = pdfium._malloc(24)
      pdfium.HEAPF32.set(imageMatrix(edit.rect, edit.rotate ?? 0), matrixPointer >> 2)
      pdfium._FPDFPageObj_SetMatrix(object, matrixPointer)
      pdfium._free(matrixPointer)
      const textIndex =
        edit.layer === 'belowText' ? firstTextIndex(collectObjects(pdfium, page)) : undefined
      if (textIndex !== undefined) {
        if (!pdfium._FPDFPage_InsertObjectAtIndex(page, object, textIndex)) {
          throw new Error('PDFium could not insert the image below text.')
        }
      } else {
        pdfium._FPDFPage_InsertObject(page, object)
      }
    } catch (error) {
      pdfium._FPDFPageObj_Destroy(object)
      throw error
    }
    return
  }

  const target = matchImage(collectObjects(pdfium, page), edit.oldRect)
  if (!target) throw new Error('The image could not be located on the page.')
  if (edit.kind === 'deleteImage') {
    if (!pdfium._FPDFPage_RemoveObject(page, target.object)) {
      throw new Error('PDFium could not remove the image object.')
    }
    pdfium._FPDFPageObj_Destroy(target.object)
    return
  }
  if (edit.layer) moveToLayer(pdfium, page, target.object, edit.layer)
  if (edit.kind === 'replaceImage') {
    await setImageBitmap(pdfium, page, target.object, edit.image)
  }
  applyGeometry(pdfium, target.object, target.bounds, edit.rect, edit.quarterTurns)
}

export interface BrowserImageEditsResult {
  readonly bytes: Uint8Array
  readonly skipped: readonly ImageEditFailure[]
}

export function applyBrowserImageEdits(
  bytes: Uint8Array,
  edits: readonly ImageEditInput[],
): Promise<BrowserImageEditsResult> {
  return chainBrowserPdfium(async () => {
    const pdfium = await loadBrowserPdfium()
    const skipped: ImageEditFailure[] = []
    return withBrowserPdfiumDocument(pdfium, bytes, async (document) => {
      const pageCount = pdfium._FPDF_GetPageCount(document)
      const byPage = new Map<number, { edit: ImageEditInput; editIndex: number }[]>()
      for (const [editIndex, edit] of edits.entries()) {
        if (edit.pageIndex < 0 || edit.pageIndex >= pageCount) {
          skipped.push({ editIndex, pageIndex: edit.pageIndex, reason: 'page does not exist' })
          continue
        }
        byPage.set(edit.pageIndex, [...(byPage.get(edit.pageIndex) ?? []), { edit, editIndex }])
      }
      let applied = 0
      for (const [pageIndex, pageEdits] of byPage) {
        const page = pdfium._FPDF_LoadPage(document, pageIndex)
        if (!page) throw new Error(`PDFium could not load page ${pageIndex + 1}.`)
        try {
          let pageApplied = 0
          for (const { edit, editIndex } of pageEdits) {
            try {
              await applyOneEdit(pdfium, document, page, edit)
              pageApplied += 1
            } catch (error) {
              skipped.push({
                editIndex,
                pageIndex,
                reason: error instanceof Error ? error.message : String(error),
              })
            }
          }
          if (pageApplied > 0 && !pdfium._FPDFPage_GenerateContent(page)) {
            throw new Error(`PDFium could not regenerate page ${pageIndex + 1}.`)
          }
          applied += pageApplied
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

export function listBrowserPageImages(bytes: Uint8Array): Promise<PageImageRef[]> {
  return chainBrowserPdfium(async () => {
    const pdfium = await loadBrowserPdfium()
    return withBrowserPdfiumDocument(pdfium, bytes, async (document) => {
      const images: PageImageRef[] = []
      for (let pageIndex = 0; pageIndex < pdfium._FPDF_GetPageCount(document); pageIndex += 1) {
        const page = pdfium._FPDF_LoadPage(document, pageIndex)
        if (!page) continue
        try {
          const objects = collectObjects(pdfium, page)
          const textIndex = firstTextIndex(objects)
          for (const object of objects) {
            if (
              object.type !== FPDF_PAGEOBJ_IMAGE ||
              object.bounds[2] - object.bounds[0] < 3 ||
              object.bounds[3] - object.bounds[1] < 3
            ) {
              continue
            }
            images.push({
              pageIndex,
              rect: object.bounds,
              aboveText: textIndex !== undefined && object.index > textIndex,
            })
          }
        } finally {
          pdfium._FPDF_ClosePage(page)
        }
      }
      return images
    })
  })
}

export function renderBrowserImagePng(
  bytes: Uint8Array,
  pageIndex: number,
  rect: Rect,
  scale = 1,
): Promise<string | null> {
  return chainBrowserPdfium(async () => {
    const pdfium = await loadBrowserPdfium()
    return withBrowserPdfiumDocument(pdfium, bytes, async (document) => {
      const page = pdfium._FPDF_LoadPage(document, pageIndex)
      if (!page) return null
      try {
        const target = matchImage(collectObjects(pdfium, page), rect)
        if (!target) return null
        const side = Math.max(
          target.bounds[2] - target.bounds[0],
          target.bounds[3] - target.bounds[1],
        )
        const factor = Math.min(Math.max(scale, 1), side > 0 ? MAX_RENDER_SIDE / side : 1)
        if (factor > 1) pdfium._FPDFPageObj_Transform(target.object, factor, 0, 0, factor, 0, 0)
        const bitmap = pdfium._FPDFImageObj_GetRenderedBitmap(document, page, target.object)
        if (!bitmap) return null
        try {
          const width = pdfium._FPDFBitmap_GetWidth(bitmap)
          const height = pdfium._FPDFBitmap_GetHeight(bitmap)
          const stride = pdfium._FPDFBitmap_GetStride(bitmap)
          const pointer = pdfium._FPDFBitmap_GetBuffer(bitmap)
          return width && height && pointer
            ? bgraToPng(
                pdfium.HEAPU8.subarray(pointer, pointer + stride * height),
                width,
                height,
                stride,
              )
            : null
        } finally {
          pdfium._FPDFBitmap_Destroy(bitmap)
        }
      } finally {
        pdfium._FPDF_ClosePage(page)
      }
    })
  })
}

export function renderBrowserPagePreviewPng(
  bytes: Uint8Array,
  request: Omit<PagePreviewRequest, 'path'>,
): Promise<string | null> {
  return chainBrowserPdfium(async () => {
    const pdfium = await loadBrowserPdfium()
    return withBrowserPdfiumDocument(pdfium, bytes, async (document) => {
      const page = pdfium._FPDF_LoadPage(document, request.pageIndex)
      if (!page || request.clip.width <= 0 || request.clip.height <= 0 || request.pxWidth <= 0) {
        return null
      }
      try {
        for (const rect of request.excludeRects) {
          const target = matchImage(collectObjects(pdfium, page), rect)
          if (target && pdfium._FPDFPage_RemoveObject(page, target.object)) {
            pdfium._FPDFPageObj_Destroy(target.object)
          }
        }
        const turns = ((request.rotate % 4) + 4) % 4
        const baseWidth = pdfium._FPDF_GetPageWidthF(page)
        const baseHeight = pdfium._FPDF_GetPageHeightF(page)
        const displayWidth = turns % 2 === 1 ? baseHeight : baseWidth
        const displayHeight = turns % 2 === 1 ? baseWidth : baseHeight
        const factor = request.pxWidth / request.clip.width
        const width = Math.max(1, Math.round(request.pxWidth))
        const height = Math.max(1, Math.round(request.clip.height * factor))
        const pointer = pdfium._malloc(width * height * 4)
        const bitmap = pdfium._FPDFBitmap_CreateEx(
          width,
          height,
          FPDF_BITMAP_BGRA,
          pointer,
          width * 4,
        )
        if (!bitmap) {
          pdfium._free(pointer)
          return null
        }
        try {
          pdfium._FPDFBitmap_FillRect(bitmap, 0, 0, width, height, 0xffffffff)
          pdfium._FPDF_RenderPageBitmap(
            bitmap,
            page,
            Math.round(-request.clip.x * factor),
            Math.round(-request.clip.y * factor),
            Math.round(displayWidth * factor),
            Math.round(displayHeight * factor),
            turns,
            request.excludeAnnots?.length ? 1 : 0,
          )
          return bgraToPng(
            pdfium.HEAPU8.subarray(pointer, pointer + width * height * 4),
            width,
            height,
          )
        } finally {
          pdfium._FPDFBitmap_Destroy(bitmap)
          pdfium._free(pointer)
        }
      } finally {
        pdfium._FPDF_ClosePage(page)
      }
    })
  })
}
