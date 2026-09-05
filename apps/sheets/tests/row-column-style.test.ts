import { describe, expect, it } from 'vitest'

import { createColumnData } from '../src/renderer/univer-sync'
import type { WorkbookCellStyle } from '../src/shared/desktop-api'

const style: WorkbookCellStyle = {
  fontFamily: 'Aptos',
  fontSize: 11,
  bold: true,
  italic: false,
  underline: false,
  strikethrough: false,
  wrapText: false,
  fillColor: '#DDEBF7',
  diagonalUp: false,
  diagonalDown: false,
}

describe('native row and column styles', () => {
  it('installs a column default style together with its width', () => {
    const data = createColumnData(
      {
        columnCount: 3,
        columnWidths: [
          { startColumn: 0, endColumn: 1, width: 10, hidden: false, styleIndex: 1 },
        ],
      } as never,
      [{ ...style, bold: false }, style],
    )
    expect(data[0]?.s).toMatchObject({ bl: 1, bg: { rgb: '#DDEBF7' } })
    expect(data[1]?.s).toEqual(data[0]?.s)
    expect(data[2]).toBeUndefined()
  })
})
