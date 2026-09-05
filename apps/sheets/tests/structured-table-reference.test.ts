import { describe, expect, it } from 'vitest'

import { resolveStructuredTableColumn } from '../src/renderer/structured-table-reference'

describe('structured table references', () => {
  const sheets = [
    {
      name: 'Data',
      tables: [
        {
          name: 'Sales',
          columns: ['Region', 'Amount'],
          range: { startRow: 2, startColumn: 4, endRow: 8, endColumn: 5 },
          headerRowCount: 1,
          totalsRowCount: 1,
        },
      ],
    },
  ]

  it('maps a table column to its data-body range', () => {
    expect(resolveStructuredTableColumn(sheets, 'Sales', 'Amount')).toEqual({
      sheetName: 'Data',
      range: { startRow: 3, startColumn: 5, endRow: 7, endColumn: 5 },
    })
  })

  it('matches Excel table and column names without case sensitivity', () => {
    expect(resolveStructuredTableColumn(sheets, 'sales', 'amount')).toEqual({
      sheetName: 'Data',
      range: { startRow: 3, startColumn: 5, endRow: 7, endColumn: 5 },
    })
  })

  it('rejects unknown, malformed, and empty data-body references', () => {
    expect(resolveStructuredTableColumn(sheets, 'Sales', 'Missing')).toBeNull()
    expect(
      resolveStructuredTableColumn(
        [
          {
            name: 'Data',
            tables: [
              {
                name: 'Empty',
                columns: ['Amount'],
                range: { startRow: 2, startColumn: 4, endRow: 2, endColumn: 4 },
                headerRowCount: 1,
                totalsRowCount: 0,
              },
            ],
          },
        ],
        'Empty',
        'Amount',
      ),
    ).toBeNull()
  })
})
