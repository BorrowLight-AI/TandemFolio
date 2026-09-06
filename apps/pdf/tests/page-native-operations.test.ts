import { PDFDocument } from 'pdf-lib'
import { describe, expect, it } from 'vitest'
import {
  cropPagesBytes,
  insertBlankPageBytes,
  mergePdfBytes,
  mergePagesBytes,
  replacePagesBytes,
  setPageSizeBytes,
  splitPdfBytes,
  splitPagesBytes,
} from '../src/domain/save-pdf'

async function documentWithPages(sizes: [number, number][]): Promise<Uint8Array> {
  const document = await PDFDocument.create()
  for (const size of sizes) document.addPage(size)
  return document.save({ useObjectStreams: false })
}

describe('native PDF page operations', () => {
  it('inserts a blank page with the neighboring page size', async () => {
    const bytes = await insertBlankPageBytes(await documentWithPages([[612, 792]]), 0)
    const reopened = await PDFDocument.load(bytes)

    expect(reopened.getPageCount()).toBe(2)
    expect(reopened.getPage(1).getSize()).toEqual({ width: 612, height: 792 })
  })

  it('resizes every page to the requested paper while keeping content centered', async () => {
    const bytes = await setPageSizeBytes(await documentWithPages([[612, 792]]), 595, 842)
    const reopened = await PDFDocument.load(bytes)

    expect(reopened.getPage(0).getSize()).toEqual({ width: 595, height: 842 })
  })

  it('crops the requested visual region without removing the page content', async () => {
    const bytes = await cropPagesBytes(
      await documentWithPages([[600, 800]]),
      [0],
      { l: 0.1, t: 0.2, r: 0.9, b: 0.8 },
    )
    const crop = (await PDFDocument.load(bytes)).getPage(0).getCropBox()

    expect(crop.x).toBeCloseTo(60)
    expect(crop.y).toBeCloseTo(160)
    expect(crop.width).toBeCloseTo(480)
    expect(crop.height).toBeCloseTo(480)
  })

  it('replaces selected pages with every page from another PDF at the first selection', async () => {
    const result = await replacePagesBytes(
      await documentWithPages([
        [600, 800],
        [600, 800],
        [600, 800],
      ]),
      await documentWithPages([
        [300, 400],
        [320, 420],
      ]),
      [1, 2],
    )
    const reopened = await PDFDocument.load(result.merged)

    expect(result).toMatchObject({ removed: 2, inserted: 2 })
    expect(reopened.getPages().map((page) => page.getSize())).toEqual([
      { width: 600, height: 800 },
      { width: 300, height: 400 },
      { width: 320, height: 420 },
    ])
  })

  it('imposes two portrait pages side by side on one landscape sheet', async () => {
    const bytes = await mergePagesBytes(
      await documentWithPages([
        [600, 800],
        [600, 800],
      ]),
      { perSheet: 2, direction: 'horizontal', separator: true },
    )
    const reopened = await PDFDocument.load(bytes)

    expect(reopened.getPageCount()).toBe(1)
    expect(reopened.getPage(0).getSize()).toEqual({ width: 800, height: 600 })
  })

  it('splits one page into four visual quadrants in reading order', async () => {
    const bytes = await splitPagesBytes(await documentWithPages([[600, 800]]), 4)
    const reopened = await PDFDocument.load(bytes)

    expect(reopened.getPageCount()).toBe(4)
    expect(reopened.getPages().map((page) => page.getSize())).toEqual([
      { width: 300, height: 400 },
      { width: 300, height: 400 },
      { width: 300, height: 400 },
      { width: 300, height: 400 },
    ])
  })

  it('splits a document into consecutive page chunks', async () => {
    const parts = await splitPdfBytes(
      await documentWithPages(Array.from({ length: 5 }, () => [600, 800])),
      2,
    )
    const pageCounts = await Promise.all(
      parts.map(async (bytes) => (await PDFDocument.load(bytes)).getPageCount()),
    )

    expect(pageCounts).toEqual([2, 2, 1])
  })

  it('appends every page from several PDFs in source order', async () => {
    const result = await mergePdfBytes(
      await documentWithPages([[600, 800]]),
      [
        await documentWithPages([
          [300, 400],
          [320, 420],
        ]),
        await documentWithPages([[500, 700]]),
      ],
    )
    const reopened = await PDFDocument.load(result.merged)

    expect(result.appended).toBe(3)
    expect(reopened.getPages().map((page) => page.getSize())).toEqual([
      { width: 600, height: 800 },
      { width: 300, height: 400 },
      { width: 320, height: 420 },
      { width: 500, height: 700 },
    ])
  })
})
