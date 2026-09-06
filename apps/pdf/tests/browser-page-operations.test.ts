import { PDFDocument } from 'pdf-lib'
import { describe, expect, it, vi } from 'vitest'

import { createBrowserPdfHost } from '../src/renderer/host/browser-pdf-api'

async function onePagePdf(): Promise<ArrayBuffer> {
  const document = await PDFDocument.create()
  document.addPage([612, 792])
  const bytes = await document.save({ useObjectStreams: false })
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer
}

async function pdfWithSizes(sizes: [number, number][]): Promise<ArrayBuffer> {
  const document = await PDFDocument.create()
  for (const size of sizes) document.addPage(size)
  const bytes = await document.save({ useObjectStreams: false })
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer
}

describe('PDF browser native page operations', () => {
  it('inserts a blank page into the mounted document and persists the new bytes', async () => {
    const persist = vi.fn(async (_name: string, _data: ArrayBuffer) => true)
    const host = createBrowserPdfHost({ download: persist })
    const staged = host.stageFile('document.pdf', await onePagePdf())
    const path = await host.api.consumePending()
    expect(path).toBeTruthy()
    host.completeOpen(path!)
    await staged

    await expect(host.api.insertBlankPage({ path: path!, afterPageIndex: 0 })).resolves.toEqual({
      ok: true,
    })
    const reopened = await PDFDocument.load(await host.api.readFile(path!))

    expect(reopened.getPageCount()).toBe(2)
    expect(persist).toHaveBeenCalledWith('document.pdf', expect.any(ArrayBuffer))
  })

  it('undoes and redoes immediate page mutations through the mounted byte history', async () => {
    const persist = vi.fn(async (_name: string, _data: ArrayBuffer) => true)
    const host = createBrowserPdfHost({ download: persist })
    const staged = host.stageFile('document.pdf', await onePagePdf())
    const path = await host.api.consumePending()
    host.completeOpen(path!)
    await staged

    await host.api.insertBlankPage({ path: path!, afterPageIndex: 0 })
    await expect(host.api.documentHistoryState(path!)).resolves.toEqual({
      canUndo: true,
      canRedo: false,
    })
    await expect(host.api.undoDocumentMutation(path!)).resolves.toEqual({ ok: true })
    expect((await PDFDocument.load(await host.api.readFile(path!))).getPageCount()).toBe(1)
    await expect(host.api.documentHistoryState(path!)).resolves.toEqual({
      canUndo: false,
      canRedo: true,
    })

    await expect(host.api.redoDocumentMutation(path!)).resolves.toEqual({ ok: true })
    expect((await PDFDocument.load(await host.api.readFile(path!))).getPageCount()).toBe(2)
    expect(persist).toHaveBeenCalledTimes(3)
  })

  it('keeps mounted bytes and history unchanged when a page mutation cannot persist', async () => {
    const host = createBrowserPdfHost({ download: async () => false })
    const staged = host.stageFile('document.pdf', await onePagePdf())
    const path = await host.api.consumePending()
    host.completeOpen(path!)
    await staged

    await expect(host.api.insertBlankPage({ path: path!, afterPageIndex: 0 })).resolves.toEqual({
      ok: false,
      error: 'The local PDF save failed.',
    })
    expect((await PDFDocument.load(await host.api.readFile(path!))).getPageCount()).toBe(1)
    await expect(host.api.documentHistoryState(path!)).resolves.toEqual({
      canUndo: false,
      canRedo: false,
    })
  })

  it('exports consecutive page chunks with deterministic names', async () => {
    const persist = vi.fn(async (_name: string, _data: ArrayBuffer) => true)
    const document = await PDFDocument.create()
    for (let index = 0; index < 3; index += 1) document.addPage([612, 792])
    const source = await document.save({ useObjectStreams: false })
    const host = createBrowserPdfHost({ download: persist })
    const staged = host.stageFile(
      'source.pdf',
      source.buffer.slice(source.byteOffset, source.byteOffset + source.byteLength) as ArrayBuffer,
    )
    const path = await host.api.consumePending()
    host.completeOpen(path!)
    await staged

    await expect(host.api.splitPdf({ path: path!, chunkSize: 2, baseName: 'part' })).resolves.toEqual({
      ok: true,
      savedDir: 'browser-downloads',
      count: 2,
    })
    expect(persist.mock.calls.map(([name]) => name)).toEqual(['part-1.pdf', 'part-2.pdf'])
  })

  it('exports N-up imposed pages without replacing the mounted document', async () => {
    const persist = vi.fn(async (_name: string, _data: ArrayBuffer) => true)
    const document = await PDFDocument.create()
    document.addPage([600, 800])
    document.addPage([600, 800])
    const source = await document.save({ useObjectStreams: false })
    const host = createBrowserPdfHost({ download: persist })
    const staged = host.stageFile(
      'source.pdf',
      source.buffer.slice(source.byteOffset, source.byteOffset + source.byteLength) as ArrayBuffer,
    )
    const path = await host.api.consumePending()
    host.completeOpen(path!)
    await staged

    await expect(
      host.api.mergePages({
        path: path!,
        perSheet: 2,
        direction: 'horizontal',
        separator: true,
        suggestedName: 'two-up.pdf',
      }),
    ).resolves.toEqual({ ok: true, savedPath: 'two-up.pdf' })
    const exported = persist.mock.calls[0]?.[1]
    expect(exported).toBeInstanceOf(ArrayBuffer)
    expect((await PDFDocument.load(exported as ArrayBuffer)).getPageCount()).toBe(1)
    expect((await PDFDocument.load(await host.api.readFile(path!))).getPageCount()).toBe(2)
  })

  it('exports inverse page splitting as a separate PDF', async () => {
    const persist = vi.fn(async (_name: string, _data: ArrayBuffer) => true)
    const host = createBrowserPdfHost({ download: persist })
    const staged = host.stageFile('source.pdf', await onePagePdf())
    const path = await host.api.consumePending()
    host.completeOpen(path!)
    await staged

    await expect(
      host.api.splitPages({ path: path!, perPage: 4, suggestedName: 'quarters.pdf' }),
    ).resolves.toEqual({ ok: true, savedPath: 'quarters.pdf' })
    const exported = persist.mock.calls[0]?.[1]
    expect((await PDFDocument.load(exported as ArrayBuffer)).getPageCount()).toBe(4)
  })

  it('changes the mounted document paper size and persists it', async () => {
    const persist = vi.fn(async (_name: string, _data: ArrayBuffer) => true)
    const host = createBrowserPdfHost({ download: persist })
    const staged = host.stageFile('source.pdf', await onePagePdf())
    const path = await host.api.consumePending()
    host.completeOpen(path!)
    await staged

    await expect(host.api.setPageSize({ path: path!, width: 595, height: 842 })).resolves.toEqual({
      ok: true,
    })
    const reopened = await PDFDocument.load(await host.api.readFile(path!))
    expect(reopened.getPage(0).getSize()).toEqual({ width: 595, height: 842 })
  })

  it('crops selected mounted pages and persists the crop box', async () => {
    const persist = vi.fn(async (_name: string, _data: ArrayBuffer) => true)
    const host = createBrowserPdfHost({ download: persist })
    const staged = host.stageFile('source.pdf', await onePagePdf())
    const path = await host.api.consumePending()
    host.completeOpen(path!)
    await staged

    await expect(
      host.api.cropPages({ path: path!, pages: [0], rect: { l: 0.1, t: 0.1, r: 0.9, b: 0.9 } }),
    ).resolves.toEqual({ ok: true })
    const crop = (await PDFDocument.load(await host.api.readFile(path!))).getPage(0).getCropBox()
    expect(crop.width).toBeCloseTo(489.6)
    expect(crop.height).toBeCloseTo(633.6)
  })

  it('replaces selected mounted pages with a picked PDF', async () => {
    const persist = vi.fn(async (_name: string, _data: ArrayBuffer) => true)
    const replacement = await pdfWithSizes([
      [300, 400],
      [320, 420],
    ])
    const host = createBrowserPdfHost({
      download: persist,
      pickPdfFiles: async () => [{ name: 'replacement.pdf', data: replacement }],
    })
    const staged = host.stageFile(
      'source.pdf',
      await pdfWithSizes([
        [600, 800],
        [600, 800],
        [600, 800],
      ]),
    )
    const path = await host.api.consumePending()
    host.completeOpen(path!)
    await staged

    await expect(host.api.replacePages({ path: path!, pages: [1, 2] })).resolves.toEqual({
      ok: true,
      removed: 2,
      inserted: 2,
    })
    const reopened = await PDFDocument.load(await host.api.readFile(path!))
    expect(reopened.getPages().map((page) => page.getSize())).toEqual([
      { width: 600, height: 800 },
      { width: 300, height: 400 },
      { width: 320, height: 420 },
    ])
  })

  it('exports a merged PDF from every picked source in order', async () => {
    const persist = vi.fn(async (_name: string, _data: ArrayBuffer) => true)
    const host = createBrowserPdfHost({
      download: persist,
      pickPdfFiles: async () => [
        { name: 'second.pdf', data: await pdfWithSizes([[300, 400]]) },
        { name: 'third.pdf', data: await pdfWithSizes([[500, 700]]) },
      ],
    })
    const staged = host.stageFile('source.pdf', await onePagePdf())
    const path = await host.api.consumePending()
    host.completeOpen(path!)
    await staged

    await expect(host.api.mergePdf({ path: path!, suggestedName: 'merged.pdf' })).resolves.toEqual({
      ok: true,
      savedPath: 'merged.pdf',
      appendedCount: 2,
    })
    const exported = persist.mock.calls[0]?.[1]
    expect((await PDFDocument.load(exported as ArrayBuffer)).getPageCount()).toBe(3)
  })

  it('executes blank-page insertion through the canonical typed operation', async () => {
    const host = createBrowserPdfHost({
      download: async (_name: string, _data: ArrayBuffer) => true,
    })
    const staged = host.stageFile('source.pdf', await onePagePdf())
    const path = await host.api.consumePending()
    host.completeOpen(path!)
    await staged

    await expect(
      host.adapter.execute({
        commandId: 'insert-blank',
        baseRevision: 0,
        operation: 'pdf.page.insert_blank',
        arguments: { afterPageIndex: 0 },
      }),
    ).resolves.toEqual({ ok: true, output: { insertedAfterPage: 0 } })
    expect((await PDFDocument.load(await host.api.readFile(path!))).getPageCount()).toBe(2)
  })

  it('executes page-size changes through the canonical typed operation', async () => {
    const host = createBrowserPdfHost({
      download: async (_name: string, _data: ArrayBuffer) => true,
    })
    const staged = host.stageFile('source.pdf', await onePagePdf())
    const path = await host.api.consumePending()
    host.completeOpen(path!)
    await staged

    await expect(
      host.adapter.execute({
        commandId: 'set-size',
        baseRevision: 0,
        operation: 'pdf.page.set_size',
        arguments: { width: 595, height: 842 },
      }),
    ).resolves.toEqual({ ok: true, output: { width: 595, height: 842 } })
    expect((await PDFDocument.load(await host.api.readFile(path!))).getPage(0).getSize()).toEqual({
      width: 595,
      height: 842,
    })
  })

  it('executes page cropping through the canonical typed operation', async () => {
    const host = createBrowserPdfHost({
      download: async (_name: string, _data: ArrayBuffer) => true,
    })
    const staged = host.stageFile('source.pdf', await onePagePdf())
    const path = await host.api.consumePending()
    host.completeOpen(path!)
    await staged
    const rect = { l: 0.1, t: 0.1, r: 0.9, b: 0.9 }

    await expect(
      host.adapter.execute({
        commandId: 'crop',
        baseRevision: 0,
        operation: 'pdf.page.crop',
        arguments: { pages: [0], rect },
      }),
    ).resolves.toEqual({ ok: true, output: { croppedPages: 1 } })
    const crop = (await PDFDocument.load(await host.api.readFile(path!))).getPage(0).getCropBox()
    expect(crop.width).toBeCloseTo(489.6)
  })

  it('executes staged page replacement through the canonical typed operation', async () => {
    const host = createBrowserPdfHost({
      download: async (_name: string, _data: ArrayBuffer) => true,
    })
    const staged = host.stageFile(
      'source.pdf',
      await pdfWithSizes([
        [600, 800],
        [600, 800],
      ]),
    )
    const path = await host.api.consumePending()
    host.completeOpen(path!)
    await staged
    const data = await pdfWithSizes([[300, 400]])

    await expect(
      host.adapter.execute({
        commandId: 'replace-staged',
        baseRevision: 0,
        operation: 'pdf.page.replace_staged',
        arguments: {
          blobId: 'replacement',
          name: 'replacement.pdf',
          size: data.byteLength,
          data,
          pages: [1],
        },
      }),
    ).resolves.toEqual({ ok: true, output: { removed: 1, inserted: 1 } })
    expect((await PDFDocument.load(await host.api.readFile(path!))).getPage(1).getSize()).toEqual({
      width: 300,
      height: 400,
    })
  })

  it('creates a blank A4 PDF through the canonical typed operation', async () => {
    const host = createBrowserPdfHost()
    await expect(host.api.consumePending()).resolves.toBeNull()
    const opened = new Promise<string>((resolve) => {
      window.addEventListener(
        'tandemfolio:pdf-open',
        (event) => resolve((event as CustomEvent<{ path: string }>).detail.path),
        { once: true },
      )
    })
    const execution = host.adapter.execute({
      commandId: 'create-blank',
      baseRevision: 0,
      operation: 'pdf.document.create_blank',
      arguments: {},
    })
    const path = await opened
    host.completeOpen(path)

    await expect(execution).resolves.toEqual({
      ok: true,
      output: { opened: true, fileName: 'Untitled.pdf', pageCount: 1 },
    })
    expect((await PDFDocument.load(await host.api.readFile(path))).getPageCount()).toBe(1)
  })

  it('requires explicit confirmation before a blank PDF replaces the mounted document', async () => {
    const host = createBrowserPdfHost()
    const staged = host.stageFile('existing.pdf', await onePagePdf())
    const path = await host.api.consumePending()
    host.completeOpen(path!)
    await staged

    await expect(
      host.adapter.execute({
        commandId: 'replace-with-blank',
        baseRevision: 0,
        operation: 'pdf.document.create_blank',
        arguments: {},
      }),
    ).resolves.toMatchObject({
      ok: false,
      error: 'execution_failed',
      message: expect.stringContaining('confirmReplace'),
    })
    expect((await PDFDocument.load(await host.api.readFile(path!))).getPageCount()).toBe(1)
  })
})
