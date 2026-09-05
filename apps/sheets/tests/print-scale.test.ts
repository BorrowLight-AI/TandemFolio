import { describe, expect, it } from 'vitest'

import { countPages, fitToPageScale } from '../src/renderer/print-scale'

const page = { printableWidthPt: 494.6, printableHeightPt: 733.5 }
const rows = (count: number, heightPt = 15) => Array.from({ length: count }, () => heightPt)

describe('print scale', () => {
  it('paginates whole rows and repeats titles', () => {
    expect(countPages([{ repeatedHeightPt: 50, rowHeightsPt: rows(10, 100) }], 250)).toBe(5)
    expect(countPages([{ repeatedHeightPt: 0, rowHeightsPt: [500, 10, 500] }], 100)).toBe(3)
  })

  it('fits width without enlarging content', () => {
    expect(
      fitToPageScale({ ...page, fitToWidth: 1, fitToHeight: 0, contentWidthPt: 989.2, areas: [] }),
    ).toBeCloseTo(0.5, 5)
    expect(
      fitToPageScale({ ...page, fitToWidth: 3, fitToHeight: 3, contentWidthPt: 100, areas: [] }),
    ).toBe(1)
  })

  it('fits whole-row height and accounts for repeated titles', () => {
    const area = { repeatedHeightPt: 100, rowHeightsPt: rows(5, 100) }
    const scale = fitToPageScale({
      printableWidthPt: 1000,
      printableHeightPt: 300,
      fitToWidth: 0,
      fitToHeight: 2,
      contentWidthPt: 100,
      areas: [area],
    })
    expect(scale).toBeLessThanOrEqual(0.7500001)
    expect(countPages([area], 300 / scale)).toBe(2)
  })

  it('uses a 10 percent floor for extreme content', () => {
    expect(
      fitToPageScale({
        ...page,
        fitToWidth: 0,
        fitToHeight: 1,
        contentWidthPt: 100,
        areas: [{ repeatedHeightPt: 0, rowHeightsPt: rows(100_000) }],
      }),
    ).toBe(0.1)
  })
})
