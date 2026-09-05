import { HorizontalAlign, WrapStrategy, type ICellData, type IStyleData } from '@univerjs/core'
import { describe, expect, it } from 'vitest'

import { CENTER_ACROSS_END_KEY } from '../src/renderer/center-continuous'
import { patchWorksheetRangeInner } from '../src/renderer/univer-sync'
import type { WorkbookCellStyle, WorkbookRangeResult } from '../src/shared/desktop-api'

const center: WorkbookCellStyle = {
  bold: false,
  italic: false,
  underline: false,
  strikethrough: false,
  wrapText: true,
  diagonalUp: false,
  diagonalDown: false,
  horizontalAlignment: 'centerContinuous',
}

function patched(cells: WorkbookRangeResult['cells']): ICellData[][] {
  let matrix: ICellData[][] = []
  patchWorksheetRangeInner(
    { getRange: () => ({ setValues: (value: ICellData[][]) => { matrix = value } }) } as never,
    undefined,
    { startRow: 0, endRow: 0, startColumn: 0, endColumn: 4 },
    cells,
    [center],
    [],
    [],
    [],
    null,
    false,
  )
  return matrix
}

describe('centerContinuous loading', () => {
  it('marks the anchor through trailing blank cells', () => {
    const matrix = patched([
      { row: 0, column: 0, value: 'Title', styleIndex: 0 },
      { row: 0, column: 1, value: null, styleIndex: 0 },
      { row: 0, column: 2, value: null, styleIndex: 0 },
      { row: 0, column: 3, value: 'Stop', styleIndex: 0 },
    ])
    expect(matrix[0]?.[0]?.custom).toEqual({ [CENTER_ACROSS_END_KEY]: 2 })
  })

  it('uses centered overflow even when the source xf says wrap', () => {
    const cell = patched([{ row: 0, column: 0, value: 'Title', styleIndex: 0 }])[0]?.[0]
    const style = cell?.s as IStyleData
    expect(style.ht).toBe(HorizontalAlign.CENTER)
    expect(style.tb).toBe(WrapStrategy.OVERFLOW)
  })
})
