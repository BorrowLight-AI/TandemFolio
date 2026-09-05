import { formulaViewSheets } from './formula-view'
import type { LazyWorkbookState } from './univer-state'

const EXPORT_ROW_BLOCK = 4096
const MAX_CSV_EXPORT_CHARS = 50_000_000

export interface CsvWorksheet {
  getLastRow(): number
  getLastColumn(): number
  getSheetId(): string
  getRange(
    row: number,
    column: number,
    rowCount: number,
    columnCount: number,
  ): { getDisplayValues(): string[][] }
}

export function csvField(text: string): string {
  const normalized = text.replace(/\r\n|\r/g, '\n')
  return /[",\n]/.test(normalized) ? `"${normalized.replace(/"/g, '""')}"` : normalized
}

export function csvFromDisplayRows(rows: readonly (readonly string[])[]): string {
  return rows.map((row) => `${row.map(csvField).join(',')}\r\n`).join('')
}

export function withoutFormulaView<T>(sheets: Set<string>, sheetId: string, read: () => T): T {
  const enabled = sheets.delete(sheetId)
  try {
    return read()
  } finally {
    if (enabled) sheets.add(sheetId)
  }
}

export function serializeActiveSheetCsv(
  sheet: CsvWorksheet,
  state: LazyWorkbookState | null,
): string | 'too-large' {
  const rowCount = Math.max(sheet.getLastRow(), 0) + 1
  const columnCount = Math.max(sheet.getLastColumn(), 0) + 1
  const parts: string[] = []
  let length = 0
  const tooLarge = withoutFormulaView(formulaViewSheets(state), sheet.getSheetId(), () => {
    for (let start = 0; start < rowCount; start += EXPORT_ROW_BLOCK) {
      const count = Math.min(EXPORT_ROW_BLOCK, rowCount - start)
      const part = csvFromDisplayRows(sheet.getRange(start, 0, count, columnCount).getDisplayValues())
      length += part.length
      if (length > MAX_CSV_EXPORT_CHARS) return true
      parts.push(part)
    }
    return false
  })
  return tooLarge ? 'too-large' : parts.join('')
}
