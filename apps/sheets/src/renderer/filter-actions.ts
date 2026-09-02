import { BooleanNumber, IUniverInstanceService, type IRange } from '@univerjs/core'
import {
  expandToContinuousRange,
  isSingleCellSelection,
  SheetsSelectionsService,
} from '@univerjs/sheets'

import type { UniverRuntime, UniverWorksheet } from './univer-state'

export type WorkbookFilterApplyResult = 'changed' | 'unchanged' | 'range_conflict' | 'failed'
export type WorkbookFilterCriteriaResult =
  'changed' | 'missing' | 'range_conflict' | 'column_outside' | 'failed'

export type WorkbookCustomFilterOperator =
  'equal' | 'notEqual' | 'greaterThan' | 'greaterThanOrEqual' | 'lessThan' | 'lessThanOrEqual'

export interface WorkbookCustomFilterCondition {
  readonly operator: WorkbookCustomFilterOperator
  readonly value: string
}

function sameRange(left: IRange, right: IRange): boolean {
  return (
    left.startRow === right.startRow &&
    left.endRow === right.endRow &&
    left.startColumn === right.startColumn &&
    left.endColumn === right.endColumn
  )
}

function workbookFilterCommandTarget(
  runtime: UniverRuntime,
  worksheet: UniverWorksheet,
  range: IRange,
):
  | { readonly unitId: string; readonly subUnitId: string }
  | Exclude<WorkbookFilterCriteriaResult, 'changed' | 'column_outside'> {
  const workbook = runtime.univerAPI.getActiveWorkbook()
  if (!workbook) return 'failed'
  const currentRange = worksheet.getFilter()?.getRange().getRange()
  if (!currentRange) return 'missing'
  if (!sameRange(currentRange, range)) return 'range_conflict'
  return { unitId: workbook.getId(), subUnitId: worksheet.getSheetId() }
}

async function setWorkbookFilterColumnCriteria(
  runtime: UniverRuntime,
  worksheet: UniverWorksheet,
  range: IRange,
  column: number,
  criteria: Readonly<Record<string, unknown>>,
): Promise<WorkbookFilterCriteriaResult> {
  const target = workbookFilterCommandTarget(runtime, worksheet, range)
  if (typeof target === 'string') return target
  if (column < range.startColumn || column > range.endColumn) return 'column_outside'
  const completed = await runtime.univerAPI.executeCommand('sheet.command.set-filter-criteria', {
    ...target,
    col: column,
    criteria: { colId: column - range.startColumn, ...criteria },
  })
  return completed ? 'changed' : 'failed'
}

/** Applies an explicit final AutoFilter state through Univer's native undoable commands. */
export async function applyWorkbookFilter(
  runtime: UniverRuntime,
  worksheet: UniverWorksheet,
  range: IRange,
  enabled: boolean,
): Promise<WorkbookFilterApplyResult> {
  const workbook = runtime.univerAPI.getActiveWorkbook()
  if (!workbook) return 'failed'
  const currentRange = worksheet.getFilter()?.getRange().getRange()
  if (currentRange && !sameRange(currentRange, range)) return 'range_conflict'
  if ((enabled && currentRange) || (!enabled && !currentRange)) return 'unchanged'

  const completed = await runtime.univerAPI.executeCommand(
    enabled ? 'sheet.command.set-filter-range' : 'sheet.command.remove-sheet-filter',
    {
      unitId: workbook.getId(),
      subUnitId: worksheet.getSheetId(),
      ...(enabled ? { range } : {}),
    },
  )
  return completed ? 'changed' : 'failed'
}

/** Clears every criterion while retaining the matching AutoFilter range. */
export async function clearWorkbookFilterCriteria(
  runtime: UniverRuntime,
  worksheet: UniverWorksheet,
  range: IRange,
): Promise<WorkbookFilterCriteriaResult> {
  const target = workbookFilterCommandTarget(runtime, worksheet, range)
  if (typeof target === 'string') return target
  const completed = await runtime.univerAPI.executeCommand(
    'sheet.command.clear-filter-criteria',
    target,
  )
  return completed ? 'changed' : 'failed'
}

/** Replaces one column's value-list criterion inside the matching AutoFilter. */
export async function setWorkbookFilterValues(
  runtime: UniverRuntime,
  worksheet: UniverWorksheet,
  range: IRange,
  column: number,
  values: readonly string[],
  includeBlank: boolean,
): Promise<WorkbookFilterCriteriaResult> {
  return setWorkbookFilterColumnCriteria(runtime, worksheet, range, column, {
    filters: {
      filters: [...values],
      ...(includeBlank ? { blank: BooleanNumber.TRUE } : {}),
    },
  })
}

/** Replaces one column's custom criterion inside the matching AutoFilter. */
export async function setWorkbookCustomFilter(
  runtime: UniverRuntime,
  worksheet: UniverWorksheet,
  range: IRange,
  column: number,
  conjunction: 'and' | 'or',
  conditions: readonly WorkbookCustomFilterCondition[],
): Promise<WorkbookFilterCriteriaResult> {
  return setWorkbookFilterColumnCriteria(runtime, worksheet, range, column, {
    customFilters: {
      ...(conjunction === 'and' && conditions.length === 2 ? { and: BooleanNumber.TRUE } : {}),
      customFilters: conditions.map(({ operator, value }) => ({
        ...(operator === 'equal' ? {} : { operator }),
        val: value,
      })),
    },
  })
}

/** Resolves the retained Ribbon toggle into an explicit range and final state. */
export function resolveWorkbookFilterToggle(
  runtime: UniverRuntime,
): { worksheet: UniverWorksheet; range: IRange; enabled: boolean } | null {
  const workbook = runtime.univerAPI.getActiveWorkbook()
  const worksheet = workbook?.getActiveSheet()
  if (!workbook || !worksheet) return null

  const currentRange = worksheet.getFilter()?.getRange().getRange()
  if (currentRange) return { worksheet, range: currentRange, enabled: false }

  const activeRange = workbook.getActiveRange()?.getRange()
  if (!activeRange) return null
  if (activeRange.startRow !== activeRange.endRow) {
    return { worksheet, range: activeRange, enabled: true }
  }

  const injector = runtime.univer.__getInjector()
  const selection = injector.get(SheetsSelectionsService).getCurrentLastSelection()
  const rawWorksheet = injector
    .get(IUniverInstanceService)
    .getUniverSheetInstance(workbook.getId())
    ?.getSheetBySheetId(worksheet.getSheetId())
  if (!selection || !rawWorksheet) return { worksheet, range: activeRange, enabled: true }
  const range = isSingleCellSelection(selection)
    ? expandToContinuousRange(
        selection.range,
        { left: true, right: true, up: true, down: true },
        rawWorksheet,
      )
    : expandToContinuousRange(selection.range, { down: true }, rawWorksheet)
  return { worksheet, range, enabled: true }
}
