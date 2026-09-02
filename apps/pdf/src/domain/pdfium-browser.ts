import { init } from '@embedpdf/pdfium'
import pdfiumWasmGzip from 'virtual:pdfium-wasm-gzip'

/** Browser-safe subset of the pinned community PDFium boundary. */
export interface BrowserPdfium {
  HEAPU8: Uint8Array
  HEAP32: Int32Array
  HEAPF32: Float32Array
  HEAPF64: Float64Array
  _malloc(size: number): number
  _free(ptr: number): void
  _PDFiumExt_Init(): void
  _PDFiumExt_OpenFileWriter(): number
  _PDFiumExt_SaveAsCopy(doc: number, writer: number): number
  _PDFiumExt_GetFileWriterSize(writer: number): number
  _PDFiumExt_GetFileWriterData(writer: number, buffer: number, size: number): number
  _PDFiumExt_CloseFileWriter(writer: number): void
  _FPDF_LoadMemDocument(pointer: number, size: number, password: number): number
  _FPDF_CloseDocument(document: number): void
  _FPDF_GetPageCount(document: number): number
  _FPDF_LoadPage(document: number, pageIndex: number): number
  _FPDF_ClosePage(page: number): void
  _FPDFPage_CountObjects(page: number): number
  _FPDFPage_GetObject(page: number, index: number): number
  _FPDFPage_RemoveObject(page: number, object: number): number
  _FPDFPage_InsertObject(page: number, object: number): void
  _FPDFPage_InsertObjectAtIndex(page: number, object: number, index: number): number
  _FPDFPage_GenerateContent(page: number): number
  _FPDFPageObj_GetType(object: number): number
  _FPDFPageObj_GetBounds(
    object: number,
    left: number,
    bottom: number,
    right: number,
    top: number,
  ): number
  _FPDFPageObj_GetFillColor(
    object: number,
    red: number,
    green: number,
    blue: number,
    alpha: number,
  ): number
  _FPDFPageObj_SetFillColor(
    object: number,
    red: number,
    green: number,
    blue: number,
    alpha: number,
  ): number
  _FPDFPageObj_SetMatrix(object: number, matrix: number): number
  _FPDFPageObj_Transform(
    object: number,
    a: number,
    b: number,
    c: number,
    d: number,
    e: number,
    f: number,
  ): void
  _FPDFPageObj_Destroy(object: number): void
  _FPDFPageObj_NewImageObj(document: number): number
  _FPDFImageObj_SetBitmap(
    pages: number,
    count: number,
    object: number,
    bitmap: number,
  ): number
  _FPDFImageObj_GetRenderedBitmap(document: number, page: number, object: number): number
  _FPDFBitmap_CreateEx(
    width: number,
    height: number,
    format: number,
    buffer: number,
    stride: number,
  ): number
  _FPDFBitmap_Destroy(bitmap: number): void
  _FPDFBitmap_FillRect(
    bitmap: number,
    left: number,
    top: number,
    width: number,
    height: number,
    color: number,
  ): void
  _FPDFBitmap_GetBuffer(bitmap: number): number
  _FPDFBitmap_GetWidth(bitmap: number): number
  _FPDFBitmap_GetHeight(bitmap: number): number
  _FPDFBitmap_GetStride(bitmap: number): number
  _FPDF_RenderPageBitmap(
    bitmap: number,
    page: number,
    startX: number,
    startY: number,
    sizeX: number,
    sizeY: number,
    rotate: number,
    flags: number,
  ): void
  _FPDF_GetPageWidthF(page: number): number
  _FPDF_GetPageHeightF(page: number): number
  _FPDFText_LoadPage(page: number): number
  _FPDFText_ClosePage(textPage: number): void
  _FPDFText_CountChars(textPage: number): number
  _FPDFText_GetTextObject(textPage: number, index: number): number
  _FPDFTextObj_GetText(object: number, textPage: number, buffer: number, length: number): number
  _FPDFTextObj_GetFontSize(object: number, size: number): number
  _FPDFText_LoadFont(
    document: number,
    data: number,
    size: number,
    fontType: number,
    cid: number,
  ): number
  _FPDFText_LoadStandardFont(document: number, font: number): number
  _FPDFPageObj_CreateTextObj(document: number, font: number, size: number): number
  _FPDFText_SetText(object: number, text: number): number
}

let instancePromise: Promise<BrowserPdfium> | null = null
let queue: Promise<void> = Promise.resolve()

export function loadBrowserPdfium(): Promise<BrowserPdfium> {
  instancePromise ??= (async () => {
    const binary = atob(pdfiumWasmGzip)
    const compressed = Uint8Array.from(binary, (character) => character.charCodeAt(0))
    const stream = new Blob([compressed.buffer as ArrayBuffer])
      .stream()
      .pipeThrough(new DecompressionStream('gzip'))
    const wasmBinary = await new Response(stream).arrayBuffer()
    const initialized = (await init({ wasmBinary })) as unknown as
      | BrowserPdfium
      | { pdfium: BrowserPdfium }
    const pdfium = 'pdfium' in initialized ? initialized.pdfium : initialized
    pdfium._PDFiumExt_Init()
    return pdfium
  })()
  return instancePromise
}

/** PDFium owns one global WASM heap; serialize document work across autosave/previews. */
export function chainBrowserPdfium<T>(work: () => Promise<T>): Promise<T> {
  const run = queue.then(work, work)
  queue = run.then(
    () => undefined,
    () => undefined,
  )
  return run
}

export async function withBrowserPdfiumDocument<T>(
  pdfium: BrowserPdfium,
  bytes: Uint8Array,
  work: (document: number) => Promise<T>,
): Promise<T> {
  const pointer = pdfium._malloc(bytes.byteLength)
  pdfium.HEAPU8.set(bytes, pointer)
  const document = pdfium._FPDF_LoadMemDocument(pointer, bytes.byteLength, 0)
  if (!document) {
    pdfium._free(pointer)
    throw new Error('PDFium could not open the PDF document.')
  }
  try {
    return await work(document)
  } finally {
    pdfium._FPDF_CloseDocument(document)
    pdfium._free(pointer)
  }
}

export function saveBrowserPdfiumDocument(
  pdfium: BrowserPdfium,
  document: number,
): Uint8Array {
  const writer = pdfium._PDFiumExt_OpenFileWriter()
  if (!writer) throw new Error('PDFium could not create a document writer.')
  try {
    if (!pdfium._PDFiumExt_SaveAsCopy(document, writer)) {
      throw new Error('PDFium could not save the edited document.')
    }
    const size = pdfium._PDFiumExt_GetFileWriterSize(writer)
    const pointer = pdfium._malloc(size)
    try {
      if (!pdfium._PDFiumExt_GetFileWriterData(writer, pointer, size)) {
        throw new Error('PDFium could not read the edited document bytes.')
      }
      return Uint8Array.from(pdfium.HEAPU8.subarray(pointer, pointer + size))
    } finally {
      pdfium._free(pointer)
    }
  } finally {
    pdfium._PDFiumExt_CloseFileWriter(writer)
  }
}
