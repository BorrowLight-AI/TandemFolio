import { describe, expect, it } from 'vitest'

import { csvField, csvFromDisplayRows, serializeActiveSheetCsv } from '../src/renderer/csv-export'

describe('CSV export', () => {
  it('quotes delimiters, quotes, and normalized newlines', () => {
    expect(csvField('a,b')).toBe('"a,b"')
    expect(csvField('say "hi"')).toBe('"say ""hi"""')
    expect(csvField('a\r\nb')).toBe('"a\nb"')
  })

  it('writes rectangular display rows with CRLF', () => {
    expect(csvFromDisplayRows([['a', 'b'], ['c', '']])).toBe('a,b\r\nc,\r\n')
  })

  it('serializes display values rather than formulas', () => {
    const rows = [['name', 'total'], ['a,b', '3']]
    const sheet = {
      getLastRow: () => 1,
      getLastColumn: () => 1,
      getSheetId: () => 's1',
      getRange: (row: number, column: number, rowCount: number, columnCount: number) => ({
        getDisplayValues: () => rows.slice(row, row + rowCount).map((values) => values.slice(column, column + columnCount)),
      }),
    }
    expect(serializeActiveSheetCsv(sheet as never, null)).toBe('name,total\r\n"a,b",3\r\n')
  })
})
