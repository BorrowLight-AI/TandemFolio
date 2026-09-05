import type { RangeBounds } from '../domain/cell-address'

interface StructuredTable {
  readonly name?: string | undefined
  readonly columns?: readonly string[] | undefined
  readonly range: RangeBounds
  readonly headerRowCount: number
  readonly totalsRowCount?: number | undefined
}

interface StructuredTableSheet {
  readonly name: string
  readonly tables: readonly StructuredTable[]
}

/** Resolve `Table[Column]` to the native table data body, excluding headers and totals. */
export function resolveStructuredTableColumn(
  sheets: readonly StructuredTableSheet[],
  tableName: string,
  columnName: string,
): { sheetName: string; range: RangeBounds } | null {
  const wantedTable = tableName.toLocaleLowerCase()
  const wantedColumn = columnName.toLocaleLowerCase()
  for (const sheet of sheets) {
    const table = sheet.tables.find((entry) => entry.name?.toLocaleLowerCase() === wantedTable)
    if (!table) continue
    const columnIndex = table.columns?.findIndex(
      (entry) => entry.toLocaleLowerCase() === wantedColumn,
    )
    if (columnIndex === undefined || columnIndex < 0) return null
    const column = table.range.startColumn + columnIndex
    if (column > table.range.endColumn) return null
    const startRow = table.range.startRow + table.headerRowCount
    const endRow = table.range.endRow - (table.totalsRowCount ?? 0)
    if (endRow < startRow) return null
    return {
      sheetName: sheet.name,
      range: { startRow, endRow, startColumn: column, endColumn: column },
    }
  }
  return null
}
