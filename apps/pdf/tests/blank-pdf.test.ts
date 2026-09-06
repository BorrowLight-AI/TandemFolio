import { PDFDocument } from 'pdf-lib'
import { describe, expect, it } from 'vitest'

import { blankPdfBytes } from '../src/domain/blank-pdf'

describe('blank PDF', () => {
  it('creates one A4 portrait page', async () => {
    const document = await PDFDocument.load(await blankPdfBytes())

    expect(document.getPageCount()).toBe(1)
    expect(document.getPage(0).getSize()).toEqual({ width: 595.28, height: 841.89 })
  })
})
