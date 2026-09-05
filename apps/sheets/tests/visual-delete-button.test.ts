import { BooleanNumber } from '@univerjs/core'
import { describe, expect, it } from 'vitest'

import {
  type GridInsetSource,
  gridTopInset,
  shouldShowVisualDeleteButton,
  visualDeleteButtonPosition,
} from '../src/renderer/visual-delete-button'

const bounds = { left: 50, top: 120, right: 850, bottom: 620 }
const size = 20
const gap = 4

describe('visual delete button visibility', () => {
  it('only appears on the selected visual outside text editing', () => {
    expect(shouldShowVisualDeleteButton({ selected: false, textEditing: false })).toBe(false)
    expect(shouldShowVisualDeleteButton({ selected: true, textEditing: false })).toBe(true)
    expect(shouldShowVisualDeleteButton({ selected: true, textEditing: true })).toBe(false)
  })
})

describe('visual delete button placement', () => {
  it('sits above the frame without overlapping even a small visual', () => {
    const frame = { left: 300, top: 300, right: 324, bottom: 324 }
    const position = visualDeleteButtonPosition(frame, bounds, size, gap)
    expect(position).toEqual({ left: 304, top: 276, placement: 'above' })
    expect(position.top + size).toBeLessThanOrEqual(frame.top)
  })

  it('flips below at the grid header and falls inside when neither side fits', () => {
    expect(
      visualDeleteButtonPosition(
        { left: 300, top: 120, right: 400, bottom: 180 },
        bounds,
        size,
        gap,
      ),
    ).toEqual({ left: 380, top: 184, placement: 'below' })
    expect(
      visualDeleteButtonPosition(
        { left: 300, top: 60, right: 400, bottom: 700 },
        bounds,
        size,
        gap,
      ),
    ).toEqual({ left: 380, top: 124, placement: 'inside' })
  })

  it('keeps the button inside the horizontal grid bounds', () => {
    expect(
      visualDeleteButtonPosition(
        { left: 800, top: 300, right: 900, bottom: 360 },
        bounds,
        size,
        gap,
      ).left,
    ).toBe(bounds.right - size)
    expect(
      visualDeleteButtonPosition(
        { left: 20, top: 300, right: 40, bottom: 360 },
        bounds,
        size,
        gap,
      ).left,
    ).toBe(bounds.left)
  })
})

function fakeWorksheet(options: {
  zoom?: number
  headerHidden?: boolean
  freeze?: { ySplit: number; startRow: number }
  hiddenRows?: readonly number[]
}): GridInsetSource {
  const hidden = new Set(options.hiddenRows ?? [])
  return {
    getZoom: () => options.zoom ?? 1,
    getFreeze: () => options.freeze ?? { ySplit: 0, startRow: 0 },
    getRowHeight: (row) => 20 + row,
    getSheet: () => ({
      getConfig: () => ({
        columnHeader: {
          height: 20,
          hidden: options.headerHidden ? BooleanNumber.TRUE : BooleanNumber.FALSE,
        },
      }),
      getRowVisible: (row) => !hidden.has(row),
    }),
  }
}

describe('grid top inset', () => {
  it('includes visible headings and frozen rows at worksheet zoom', () => {
    expect(gridTopInset(fakeWorksheet({}))).toBe(20)
    expect(gridTopInset(fakeWorksheet({ headerHidden: true }))).toBe(0)
    expect(gridTopInset(fakeWorksheet({ zoom: 1.5 }))).toBe(30)
    expect(gridTopInset(fakeWorksheet({ freeze: { ySplit: 3, startRow: 3 } }))).toBe(83)
    expect(
      gridTopInset(fakeWorksheet({ freeze: { ySplit: 3, startRow: 3 }, hiddenRows: [1] })),
    ).toBe(62)
  })
})
