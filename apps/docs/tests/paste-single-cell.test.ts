import { describe, expect, it } from 'vitest'
import { unwrapSingleCellTable } from '../src/renderer/App'

describe('spreadsheet single-cell paste', () => {
  it('unwraps a lone cell while keeping real tables and mixed content intact', () => {
    expect(unwrapSingleCellTable('<table><tr><td><b>Value</b></td></tr></table>')).toBe(
      '<b>Value</b>',
    )
    const matrix = '<table><tr><td>A</td><td>B</td></tr></table>'
    expect(unwrapSingleCellTable(matrix)).toBe(matrix)
    const mixed = '<p>Lead</p><table><tr><td>Value</td></tr></table>'
    expect(unwrapSingleCellTable(mixed)).toBe(mixed)
  })
})
