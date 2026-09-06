import { Blob as NodeBlob } from 'node:buffer'
import { PDFDocument, StandardFonts } from 'pdf-lib'
import { beforeAll, describe, expect, it, vi } from 'vitest'

import {
  applyBrowserTextEdits,
  validateBrowserTextEdits,
} from '../src/domain/browser-text-edit'

async function textPdf(text: string): Promise<Uint8Array> {
  const document = await PDFDocument.create()
  const page = document.addPage([600, 800])
  const font = await document.embedFont(StandardFonts.Helvetica)
  page.drawText(text, { x: 50, y: 700, size: 12, font })
  return document.save({ useObjectStreams: false })
}

async function twoRunPdf(): Promise<Uint8Array> {
  const document = await PDFDocument.create()
  const page = document.addPage([600, 800])
  const font = await document.embedFont(StandardFonts.Helvetica)
  page.drawText('Delete ', { x: 50, y: 700, size: 12, font })
  page.drawText('both', { x: 90, y: 700, size: 12, font })
  return document.save({ useObjectStreams: false })
}

describe('native browser PDF text operations', () => {
  beforeAll(() => {
    vi.stubGlobal('Blob', NodeBlob)
    vi.spyOn(HTMLCanvasElement.prototype, 'getContext').mockReturnValue({
      font: '',
      measureText: (text: string) => ({ width: text.length * 7 }),
    } as unknown as CanvasRenderingContext2D)
  })

  it('moves an exact text run without rebuilding its glyphs', async () => {
    const result = await applyBrowserTextEdits(await textPdf('Move me'), [
      {
        pageIndex: 0,
        rect: [40, 690, 120, 720],
        oldText: 'Move me',
        newText: 'Move me',
        fontSize: 12,
        translate: [120, 0],
      },
    ])

    expect(result.skipped).toEqual([])
    await expect(
      validateBrowserTextEdits(result.bytes, [
        {
          pageIndex: 0,
          rect: [160, 690, 240, 720],
          oldText: 'Move me',
          newText: 'Move me',
          fontSize: 12,
        },
      ]),
    ).resolves.toMatchObject([{ reason: null }])
  })

  it('persists selection-level text styles as separate native text runs', async () => {
    const result = await applyBrowserTextEdits(await textPdf('Color'), [
      {
        pageIndex: 0,
        rect: [40, 690, 120, 720],
        oldText: 'Color',
        newText: 'Color',
        fontSize: 12,
        origin: [50, 700],
        newColor: [255, 0, 0],
        styleRuns: [{ start: 2, end: 5, color: [0, 0, 255], bold: true }],
      },
    ])

    expect(result.skipped).toEqual([])
    await expect(
      validateBrowserTextEdits(result.bytes, [
        {
          pageIndex: 0,
          rect: [40, 690, 140, 720],
          oldText: 'Color',
          newText: 'Color',
          fontSize: 12,
        },
      ]),
    ).resolves.toMatchObject([
      {
        reason: null,
        baseColor: [255, 0, 0],
        colorRuns: [
          { start: 0, end: 2, color: [255, 0, 0] },
          { start: 2, end: 5, color: [0, 0, 255] },
        ],
      },
    ])
  })

  it('replaces a selected fragment inside one native text object without dropping its neighbors', async () => {
    const result = await applyBrowserTextEdits(await textPdf('Hello world'), [
      {
        pageIndex: 0,
        rect: [77, 690, 120, 720],
        oldText: 'world',
        newText: 'PDF',
        fontSize: 12,
      },
    ])

    expect(result.skipped).toEqual([])
    await expect(
      validateBrowserTextEdits(result.bytes, [
        {
          pageIndex: 0,
          rect: [40, 690, 130, 720],
          oldText: 'Hello PDF',
          newText: 'Hello PDF',
          fontSize: 12,
        },
      ]),
    ).resolves.toMatchObject([{ reason: null }])
  })

  it('deletes every matched native text object when the final text is empty', async () => {
    const result = await applyBrowserTextEdits(await twoRunPdf(), [
      {
        pageIndex: 0,
        rect: [40, 690, 140, 720],
        oldText: 'Delete both',
        newText: '',
        fontSize: 12,
      },
    ])

    expect(result.skipped).toEqual([])
    await expect(
      validateBrowserTextEdits(result.bytes, [
        {
          pageIndex: 0,
          rect: [40, 690, 140, 720],
          oldText: 'Delete both',
          newText: 'Delete both',
          fontSize: 12,
        },
      ]),
    ).resolves.toEqual([{ reason: 'the addressed text could not be matched' }])
  })
})
