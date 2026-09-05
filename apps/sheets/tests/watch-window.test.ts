import { describe, expect, it } from 'vitest'

import {
  appendWatchSelection,
  removeWatch,
  watchKey,
  type WatchCell,
} from '../src/renderer/WatchWindowPanel'

describe('Watch Window cell set', () => {
  it('uses stable sheet ids and coordinates for cell identity', () => {
    expect(watchKey({ sheetId: 'sheet-1', row: 4, column: 2 })).toBe('sheet-1:4:2')
  })

  it('adds a rectangular selection in row-major order without duplicates', () => {
    const existing: WatchCell[] = [{ sheetId: 'sheet-1', row: 1, column: 1 }]
    expect(
      appendWatchSelection(existing, 'sheet-1', {
        startRow: 1,
        endRow: 2,
        startColumn: 1,
        endColumn: 2,
      }),
    ).toEqual([
      { sheetId: 'sheet-1', row: 1, column: 1 },
      { sheetId: 'sheet-1', row: 1, column: 2 },
      { sheetId: 'sheet-1', row: 2, column: 1 },
      { sheetId: 'sheet-1', row: 2, column: 2 },
    ])
  })

  it('treats the same coordinate on another worksheet as a distinct watch', () => {
    const existing: WatchCell[] = [{ sheetId: 'sheet-1', row: 0, column: 0 }]
    expect(
      appendWatchSelection(existing, 'sheet-2', {
        startRow: 0,
        endRow: 0,
        startColumn: 0,
        endColumn: 0,
      }),
    ).toEqual([
      { sheetId: 'sheet-1', row: 0, column: 0 },
      { sheetId: 'sheet-2', row: 0, column: 0 },
    ])
  })

  it('caps the complete watch list at twenty cells for whole-axis selections', () => {
    const watches = appendWatchSelection([], 'sheet-1', {
      startRow: 0,
      endRow: 1_048_575,
      startColumn: 0,
      endColumn: 0,
    })
    expect(watches).toHaveLength(20)
    expect(watches[19]).toEqual({ sheetId: 'sheet-1', row: 19, column: 0 })
  })

  it('honors the remaining capacity when watches already exist', () => {
    const existing = Array.from({ length: 19 }, (_, row) => ({
      sheetId: 'sheet-1',
      row,
      column: 0,
    }))
    const watches = appendWatchSelection(existing, 'sheet-2', {
      startRow: 0,
      endRow: 2,
      startColumn: 0,
      endColumn: 0,
    })
    expect(watches).toHaveLength(20)
    expect(watches[19]).toEqual({ sheetId: 'sheet-2', row: 0, column: 0 })
  })

  it('removes only the requested stable watch key', () => {
    const watches: WatchCell[] = [
      { sheetId: 'sheet-1', row: 0, column: 0 },
      { sheetId: 'sheet-2', row: 0, column: 0 },
    ]
    expect(removeWatch(watches, 'sheet-1:0:0')).toEqual([
      { sheetId: 'sheet-2', row: 0, column: 0 },
    ])
  })
})
