import type { IRange } from '@univerjs/core'
import { describe, expect, it } from 'vitest'

import { revealCellBelowFreeze } from '../src/renderer/univer-sync'

function range(startRow: number, startColumn: number): IRange {
  return { startRow, startColumn, endRow: startRow + 20, endColumn: startColumn + 10 } as IRange
}

describe('revealCellBelowFreeze', () => {
  it('corrects a frozen-pane offset until the target row is visible', async () => {
    const calls: Array<[number, number]> = []
    let scrolledRow = 0
    const sheet = {
      scrollToCell(row: number, column: number) {
        calls.push([row, column])
        scrolledRow = row
      },
      getVisibleRange: () => range(scrolledRow + 2, 0),
    }
    await revealCellBelowFreeze(sheet, 10, 0)
    expect(calls).toEqual([
      [9, 0],
      [7, 0],
    ])
  })

  it('stops when the viewport clamps at the frozen pane', async () => {
    const calls: Array<[number, number]> = []
    const sheet = {
      scrollToCell(row: number, column: number) {
        calls.push([row, column])
      },
      getVisibleRange: () => range(2, 0),
    }
    await revealCellBelowFreeze(sheet, 2, 0)
    expect(calls).toEqual([
      [1, 0],
      [0, 0],
    ])
  })

  it('lets a newer reveal supersede an older correction loop', async () => {
    const calls: Array<[number, number]> = []
    let scrolledRow = 0
    const sheet = {
      scrollToCell(row: number, column: number) {
        calls.push([row, column])
        scrolledRow = row
      },
      getVisibleRange: () => range(scrolledRow + 2, 0),
    }
    await Promise.all([
      revealCellBelowFreeze(sheet, 50, 0),
      revealCellBelowFreeze(sheet, 10, 0),
    ])
    expect(calls).toEqual([
      [49, 0],
      [9, 0],
      [7, 0],
    ])
  })
})
