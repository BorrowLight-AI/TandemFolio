import { columnIndex } from '../domain/cell-address'
import { shiftFormulaText, StructuralShiftError } from '../gateway/xlsx-structure'
import { isSheetRemoved } from './edit-journal'
import type { LazyWorkbookState } from './univer-state'

export type DeleteSpanOperation =
  | { op: 'delete_rows'; sheetId: string; row: number; count: number }
  | { op: 'delete_cols'; sheetId: string; column: string; count: number }

interface DeleteSpanSpec {
  axis: 'row' | 'column'
  shift: { boundary: number; delta: number; deleted: { start: number; end: number } }
  deletedSheetName: string
}

function deleteSpanSpec(
  state: Pick<LazyWorkbookState, 'file'>,
  sheetNameOf: (sheetId: string) => string | undefined,
  operation: DeleteSpanOperation,
): DeleteSpanSpec {
  const axis = operation.op === 'delete_cols' ? ('column' as const) : ('row' as const)
  const index = operation.op === 'delete_cols' ? columnIndex(operation.column) : operation.row - 1
  return {
    axis,
    shift: {
      boundary: index,
      delta: -operation.count,
      deleted: { start: index, end: index + operation.count - 1 },
    },
    deletedSheetName:
      state.file.sheets.find((sheet) => sheet.id === operation.sheetId)?.name ??
      sheetNameOf(operation.sheetId) ??
      '',
  }
}

function deletedSpanTextError(
  texts: readonly string[],
  spec: DeleteSpanSpec,
  qualifiedOnly: boolean,
): string | null {
  for (const text of texts) {
    try {
      shiftFormulaText(text, spec.deletedSheetName, spec.shift, spec.axis, qualifiedOnly)
    } catch (error) {
      if (error instanceof StructuralShiftError) {
        const span = spec.axis === 'column' ? 'columns' : 'rows'
        return (
          `A formula (${text.length > 80 ? `${text.slice(0, 80)}…` : text}) references only the ` +
          `deleted ${span} — the save cannot rewrite it to #REF! yet, so the deletion would ` +
          'fail there. Update or remove such formulas first, then retry.'
        )
      }
      throw error
    }
  }
  return null
}

export async function structuralDeleteFormulaError(
  state: Pick<LazyWorkbookState, 'file' | 'editJournal'>,
  workbook: { getSheets(): { getSheetId(): string; getSheetName(): string }[] },
  operation: DeleteSpanOperation,
): Promise<string | null> {
  if ([...state.editJournal.structuralOps.values()].some((ops) => ops.length > 0)) return null
  const sheets = workbook.getSheets()
  const spec = deleteSpanSpec(
    state,
    (id) => sheets.find((sheet) => sheet.getSheetId() === id)?.getSheetName(),
    operation,
  )
  const fileSheetIds = new Set(state.file.sheets.map((sheet) => sheet.id))
  for (const sheet of sheets) {
    const sheetId = sheet.getSheetId()
    if (isSheetRemoved(state.editJournal, sheetId)) continue
    const journalCells = state.editJournal.cells.get(sheetId)
    const texts: string[] = []
    if (journalCells) {
      for (const entry of journalCells.values()) if (entry.formula) texts.push(entry.formula)
    }
    if (fileSheetIds.has(sheetId)) {
      let result
      try {
        result = await window.desktopApi.readWorkbookFormulas({
          sessionId: state.file.sessionId,
          sheetId,
        })
      } catch {
        return null
      }
      if (result.truncated || !result.indexingComplete) return null
      for (const cell of result.cells) {
        if (!cell.formula) continue
        const entry = journalCells?.get(`${cell.row}:${cell.column}`)
        if (entry && (entry.hasValue || entry.formula)) continue
        texts.push(cell.formula)
      }
    }
    const error = deletedSpanTextError(texts, spec, sheetId !== operation.sheetId)
    if (error) return error
  }
  return null
}

export function structuralDeleteFormulaErrorSync(
  state: LazyWorkbookState,
  workbook: {
    getSheets(): {
      getSheetId(): string
      getSheetName(): string
      getMaxRows(): number
      getMaxColumns(): number
      getRange(
        row: number,
        column: number,
        rows: number,
        columns: number,
      ): { getFormulas(): string[][] }
    }[]
  },
  operation: DeleteSpanOperation,
): string | null {
  const structuralShifted = [...state.editJournal.structuralOps.values()].some(
    (ops) => ops.length > 0,
  )
  if (!state.formulaMode && structuralShifted) return null
  const sheets = workbook.getSheets()
  const spec = deleteSpanSpec(
    state,
    (id) => sheets.find((sheet) => sheet.getSheetId() === id)?.getSheetName(),
    operation,
  )
  for (const sheet of sheets) {
    const sheetId = sheet.getSheetId()
    if (isSheetRemoved(state.editJournal, sheetId)) continue
    const texts: string[] = []
    if (state.formulaMode) {
      const formulas = sheet.getRange(0, 0, sheet.getMaxRows(), sheet.getMaxColumns()).getFormulas()
      for (const row of formulas) for (const formula of row) if (formula) texts.push(formula)
    } else {
      const journalCells = state.editJournal.cells.get(sheetId)
      if (journalCells) {
        for (const entry of journalCells.values()) if (entry.formula) texts.push(entry.formula)
      }
      const harvested = state.formulaText.get(sheetId)
      if (harvested) {
        for (const [key, text] of harvested) {
          const entry = journalCells?.get(key)
          if (entry && (entry.hasValue || entry.formula)) continue
          texts.push(text)
        }
      }
    }
    const error = deletedSpanTextError(texts, spec, sheetId !== operation.sheetId)
    if (error) return error
  }
  return null
}
