import { describe, expect, it } from 'vitest'

import { applyRowProperties } from '../src/renderer/univer-sync'
import { loadAutoHeightSuppression } from '../src/renderer/univer-state'

function makeWorksheet() {
  const calls: Array<{ method: string; row: number; count: number; px?: number; suppressed?: boolean }> = []
  const worksheet = {
    setRowHeightsForced: (row: number, count: number, px: number) =>
      calls.push({ method: 'forced', row, count, px }),
    setRowAutoHeight: (row: number, count: number) =>
      calls.push({ method: 'auto', row, count, suppressed: loadAutoHeightSuppression.active }),
    hideRows: (row: number, count: number) => calls.push({ method: 'hide', row, count }),
    getSheet: () => ({ setRowStyle: () => undefined }),
  }
  return { worksheet, calls }
}

function makeState(defaultRowHeight: number | null = null) {
  return {
    file: { styles: [], sheets: [{ id: 'sheet-1', defaultRowHeight }] },
    appliedRowKeys: new Map<string, Set<string>>(),
    outline: new Map(),
  }
}

describe('native row height import', () => {
  it('keeps fixed and spacer rows locked but restores auto mode on ordinary rows', () => {
    const { worksheet, calls } = makeWorksheet()
    applyRowProperties(worksheet as never, makeState(15) as never, 'sheet-1', [
      { row: 0, height: 56, customHeight: true, hidden: false },
      { row: 1, height: 38.25, hidden: false },
      { row: 2, height: 2.25, hidden: false },
      { row: 3, hidden: true },
    ] as never)

    expect(calls).toEqual([
      { method: 'forced', row: 0, count: 1, px: 75 },
      { method: 'forced', row: 1, count: 1, px: 51 },
      { method: 'forced', row: 2, count: 1, px: 3 },
      { method: 'auto', row: 1, count: 1, suppressed: true },
      { method: 'hide', row: 3, count: 1 },
    ])
    expect(loadAutoHeightSuppression.active).toBe(false)
  })

  it('includes customHeight in deduplication keys', () => {
    const { worksheet, calls } = makeWorksheet()
    const state = makeState()
    applyRowProperties(worksheet as never, state as never, 'sheet-1', [
      { row: 0, height: 30, hidden: false },
    ] as never)
    applyRowProperties(worksheet as never, state as never, 'sheet-1', [
      { row: 0, height: 30, customHeight: true, hidden: false },
    ] as never)
    expect(calls.filter((call) => call.method === 'forced')).toHaveLength(2)
  })
})
