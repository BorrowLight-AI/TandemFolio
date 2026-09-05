import { describe, expect, it } from 'vitest'

import { createEditJournal, recordPageSetup, recordStructuralOp } from '../src/renderer/edit-journal'
import {
  computePageBoundaries,
  effectivePageBreaks,
  printablePagePt,
} from '../src/renderer/page-break-preview'
import type { LazyWorkbookState } from '../src/renderer/univer-state'

describe('printablePagePt', () => {
  it('defaults to A4 portrait with normal margins', () => {
    expect(printablePagePt({})).toEqual({
      width: (8.27 - 1.4) * 72,
      height: (11.69 - 1.5) * 72,
    })
  })

  it('swaps axes in landscape and honors paper and margin choices', () => {
    const page = printablePagePt({ paperSize: 1, orientation: 'landscape', margins: 'narrow' })
    expect(page.width).toBeCloseTo((11 - 0.5) * 72, 5)
    expect(page.height).toBeCloseTo((8.5 - 1.5) * 72, 5)
  })

  it('uses exact file margins from effective print settings', () => {
    const page = printablePagePt({
        orientation: 'portrait',
        paperSize: 1,
        scale: 100,
        fitToWidth: 0,
        fitToHeight: 0,
        fitToPage: false,
        margins: { left: 0.9, right: 0.8, top: 0.7, bottom: 0.6, header: 0.2, footer: 0.2 },
        printGridlines: false,
        printHeadings: false,
        printAreas: [],
        printTitles: null,
        header: null,
        footer: null,
        firstPage: null,
        evenPages: null,
        headerFooterScaleWithDoc: true,
        headerFooterPictures: [],
      })
    expect(page.width).toBeCloseTo((8.5 - 1.7) * 72, 8)
    expect(page.height).toBeCloseTo((11 - 1.3) * 72, 8)
  })
})

describe('computePageBoundaries', () => {
  const rows100 = () => 100

  it('places automatic boundaries where the next row would overflow', () => {
    expect(computePageBoundaries(rows100, 7, 200, 1, [])).toEqual([
      { index: 2, manual: false },
      { index: 4, manual: false },
      { index: 6, manual: false },
    ])
  })

  it('lets a manual break reset the accumulator', () => {
    expect(computePageBoundaries(rows100, 6, 200, 1, [1])).toEqual([
      { index: 1, manual: true },
      { index: 3, manual: false },
      { index: 5, manual: false },
    ])
  })
})

function stateWith(setup: (state: LazyWorkbookState) => void): LazyWorkbookState {
  const state = {
    editJournal: createEditJournal(),
    file: {
      sheets: [{ id: 'sheet-1', name: 'Budget', pageSetup: { rowBreaks: [5], colBreaks: [2] } }],
    },
  } as unknown as LazyWorkbookState
  setup(state)
  return state
}

describe('effectivePageBreaks', () => {
  it('prefers the journal set over the file set per axis', () => {
    const state = stateWith((current) => {
      recordPageSetup(current.editJournal, 'sheet-1', { rowBreaks: [9] })
    })
    expect(effectivePageBreaks(state, 'sheet-1')).toEqual({ rowBreaks: [9], colBreaks: [2] })
  })

  it('shifts file breaks through structural operations', () => {
    const state = stateWith((current) => {
      recordStructuralOp(
        current.editJournal,
        'sheet-1',
        { kind: 'remove-rows', index: 4, count: 3 },
        'Budget',
      )
      current.file.sheets[0]!.pageSetup = { rowBreaks: [5, 10], colBreaks: [] }
    })
    expect(effectivePageBreaks(state, 'sheet-1')).toEqual({ rowBreaks: [7], colBreaks: [] })
  })
})
