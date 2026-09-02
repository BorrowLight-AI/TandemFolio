/**
 * Cell navigation (Name Box / Go To / formula / symbol) and data tools
 * (advanced filter, subtotal, consolidate, outline, format-as-table).
 * Extracted from App.tsx; the App component passes a DataToolsContext built
 * fresh per call so refs and state never go stale.
 */
import { ILayoutService } from '@univerjs/preset-sheets-core'

import { columnLabel } from '../domain/cell-address'
import { decodeCsvBuffer, isNumericCell, parseCsv } from '../gateway/csv-import'
import type { AdvancedFilterColumn, AdvancedFilterCriteria } from './AdvancedFilterDialog'
import { setWorkbookCustomFilter } from './filter-actions'
import {
  buildLabelMatrix,
  buildPositionMatrix,
  parseConsolidateReference,
  targetOverlapsSource,
  type ConsolidateArea,
  type ConsolidateConfig,
  type OutputCell,
} from './consolidate'
import { isSheetRemoved, journalSize } from './edit-journal'
import { resolveGoToRef, type GoToNameEntry } from './goto'
import { getLang, t } from './i18n/locale'
import {
  applyWorkbookOutlineDetailVisibility,
  applyWorkbookOutlineLevels,
  type WorkbookOutlineAxis,
} from './outline-actions'
import { appendSymbol } from './SymbolDialog'
import {
  advancedFilterColumnOptions,
  applyWorkbookCellValue,
  applyWorkbookCellMatrix,
  columnLetter,
  loadVisibleRange,
  sheetOutline,
  univerDefinedNames,
} from './univer-sync'
import type { LazyWorkbookState, UniverRuntime, UniverWorksheet } from './univer-state'
import { applyWorkbookTableAdd } from './workbook-ops'

/** The App refs/state the data-tool actions need; built fresh per call. */
export interface DataToolsContext {
  univerRef: { readonly current: UniverRuntime | null }
  lazyWorkbookRef: { readonly current: LazyWorkbookState | null }
  setMessage: (message: string) => void
  setPendingEdits: (count: number) => void
  setAdvancedFilterColumns: (columns: readonly AdvancedFilterColumn[] | null) => void
}

/// Matches the paste ceiling in spirit: one setValues command, journaled and
/// undoable; larger files should open as their own workbook instead.
const CSV_IMPORT_MAX_CELLS = 50_000

/// Excel writes CSV in the system's legacy charset; the UI language is the
/// best tie-breaker we have (same map as the main-process open path).
const CSV_CHARSET_BY_LANG: Record<string, string> = {
  zh: 'gb18030',
  'zh-TW': 'big5',
  ja: 'shift_jis',
  ko: 'euc-kr',
}

/// Data → From Text/CSV: reads a delimited file (encoding and delimiter
/// sniffed by the shared CSV pipeline) into the active sheet at the selection.
export function handleImportCsv(ctx: DataToolsContext): void {
  const input = document.createElement('input')
  input.type = 'file'
  input.accept = '.csv,.tsv,.txt,text/csv,text/tab-separated-values,text/plain'
  input.onchange = () => {
    const file = input.files?.[0]
    if (!file) return
    if (file.size > 32 * 1024 * 1024) {
      ctx.setMessage(t('appCsvTooLarge'))
      return
    }
    void file.arrayBuffer().then((buffer) => {
      importCsvText(ctx, decodeCsvBuffer(new Uint8Array(buffer), CSV_CHARSET_BY_LANG[getLang()]))
    })
  }
  input.click()
}

function importCsvText(ctx: DataToolsContext, text: string): void {
  const runtime = ctx.univerRef.current
  const workbook = runtime?.univerAPI.getActiveWorkbook()
  const worksheet = workbook?.getActiveSheet()
  const range = workbook?.getActiveRange()
  if (!worksheet || !range) {
    ctx.setMessage(t('appSelectCellFirst'))
    return
  }
  const rows = parseCsv(text)
  const columns = rows.reduce((width, row) => Math.max(width, row.length), 0)
  if (rows.length === 0 || columns === 0) {
    ctx.setMessage(t('appCsvEmpty'))
    return
  }
  if (rows.length * columns > CSV_IMPORT_MAX_CELLS) {
    ctx.setMessage(t('appCsvTooLarge'))
    return
  }
  const values = rows.map((row) =>
    Array.from({ length: columns }, (_, index) => {
      const cell = row[index] ?? ''
      return isNumericCell(cell) ? { v: Number(cell) } : { v: cell }
    }),
  )
  const row = range.getRow()
  const column = range.getColumn()
  try {
    applyWorkbookCellMatrix(worksheet.getRange(row, column, rows.length, columns), values)
  } catch (error: unknown) {
    ctx.setMessage(error instanceof Error ? error.message : t('appCsvImportFailed'))
    return
  }
  ctx.setMessage(
    t('appCsvImported', {
      rows: rows.length,
      columns,
      cell: `${columnLetter(column)}${row + 1}`,
    }),
  )
}

export function activeCellLabel(ctx: DataToolsContext): string {
  const range = ctx.univerRef.current?.univerAPI.getActiveWorkbook()?.getActiveRange()
  if (!range) return t('appActiveCellFallback')
  return `${columnLetter(range.getColumn())}${range.getRow() + 1}`
}

/// Name Box / Go To jump: resolves the typed reference (an A1 address or a
/// defined name) and selects it, switching sheets when the target lives on
/// another one. Returns null on success or a user-facing error message.
/// Univer's parser yields NaN rows instead of throwing on garbage, so
/// validity is decided by resolveGoToRef — the try/catch covers what
/// getRange itself rejects: unknown sheet names and out-of-bounds ranges.
export function goToReference(ctx: DataToolsContext, ref: string): string | null {
  const workbook = ctx.univerRef.current?.univerAPI.getActiveWorkbook()
  const worksheet = workbook?.getActiveSheet()
  if (!workbook || !worksheet) return t('appGoToNotReady')
  const trimmed = ref.trim()
  if (trimmed === '') return t('appGoToEmpty')
  const resolved = resolveGoToRef(trimmed, listDefinedNames(ctx))
  if (resolved === null) {
    return t('appGoToUnresolved', { ref: trimmed })
  }
  try {
    // A jump must not leave an editor open on the previous cell — later
    // keystrokes would land there. Commit it before moving.
    const editing = workbook as unknown as {
      isCellEditing?(): boolean
      endEditingAsync?(save: boolean): Promise<boolean>
    }
    if (editing.isCellEditing?.()) void editing.endEditingAsync?.(true)
    const range = worksheet.getRange(resolved)
    const target = workbook.getSheetBySheetId(range.getSheetId()) ?? worksheet
    workbook.setActiveRange(range)
    target.scrollToCell(range.getRow(), range.getColumn())
    // Hand keyboard focus back to the grid (Univer's hidden editor host, the
    // same handoff its own name box does); typing right after a jump then
    // lands in the target cell instead of being dropped on <body>.
    ctx.univerRef.current?.univer.__getInjector().get(ILayoutService).focus()
    // The single Scroll event a jump produces computes its viewport from a
    // stale getVisibleRange; load around the target explicitly.
    if (ctx.univerRef.current) {
      void loadVisibleRange(
        ctx.univerRef.current,
        ctx.lazyWorkbookRef as { current: LazyWorkbookState | null },
        target,
        ctx.setMessage,
        { row: range.getRow(), column: range.getColumn() },
      )
    }
  } catch (error: unknown) {
    const detail = error instanceof Error ? error.message : ''
    if (detail.includes('out of bounds')) return t('appGoToOutOfBounds', { ref: trimmed })
    if (detail.includes('Range not found')) return t('appGoToSheetNotFound', { ref: trimmed })
    return detail === ''
      ? t('appGoToInvalidAddress', { ref: trimmed })
      : t('appGoToFailed', { ref: trimmed, detail })
  }
  return null
}

/// All defined names of the active workbook, feeding both the Go To
/// dialog's list and Name Box resolution.
export function listDefinedNames(ctx: DataToolsContext): GoToNameEntry[] {
  return univerDefinedNames(ctx.univerRef.current).map((defined) => ({
    name: defined.getName(),
    ref: defined.getFormulaOrRefString(),
  }))
}

export function handleApplyFormula(ctx: DataToolsContext, formula: string): string | null {
  const runtime = ctx.univerRef.current
  if (!runtime) return t('appWorkbookNotReady')
  const workbook = runtime.univerAPI.getActiveWorkbook()
  const worksheet = workbook?.getActiveSheet()
  const range = workbook?.getActiveRange()
  if (!workbook || !worksheet || !range) return t('appSelectCellFirst')
  const trimmed = formula.trim()
  if (!trimmed.startsWith('=')) return t('appFormulaStartsEquals')
  const opens = (trimmed.match(/\(/g) ?? []).length
  const closes = (trimmed.match(/\)/g) ?? []).length
  if (opens !== closes) return t('appUnbalancedParens')
  try {
    worksheet.getRange(range.getRow(), range.getColumn(), 1, 1).setValue({ f: trimmed })
  } catch (error: unknown) {
    return error instanceof Error ? error.message : t('appSetFormulaFailed')
  }
  ctx.setMessage(t('appFormulaSet', { cell: activeCellLabel(ctx) }))
  return null
}

/// The Symbol dialog's click: appends the picked character to the active
/// range's top-left cell as text. setValue lands as a normal
/// set-range-values command, so it journals and undoes like typing.
export function handleInsertSymbol(ctx: DataToolsContext, char: string): void {
  const workbook = ctx.univerRef.current?.univerAPI.getActiveWorkbook()
  const worksheet = workbook?.getActiveSheet()
  const range = workbook?.getActiveRange()
  if (!workbook || !worksheet || !range) {
    ctx.setMessage(t('appSymbolNeedsCell'))
    return
  }
  const cell = worksheet.getRange(range.getRow(), range.getColumn(), 1, 1)
  try {
    applyWorkbookCellValue(cell, appendSymbol(cell.getValue(), char))
  } catch (error: unknown) {
    ctx.setMessage(error instanceof Error ? error.message : t('appSymbolInsertFailed'))
    return
  }
  ctx.setMessage(t('appSymbolInserted', { char, cell: activeCellLabel(ctx) }))
}

/// Samples the active filter range's header row and opens the Advanced
/// Filter dialog; without an auto-filter it explains instead of opening.
export function openAdvancedFilterDialog(ctx: DataToolsContext): void {
  const worksheet = ctx.univerRef.current?.univerAPI.getActiveWorkbook()?.getActiveSheet()
  const filter = worksheet?.getFilter()
  if (!worksheet || !filter) {
    ctx.setMessage(t('appAdvFilterNeedsFilter'))
    return
  }
  ctx.setAdvancedFilterColumns(advancedFilterColumnOptions(worksheet, filter))
}

/// The Advanced Filter dialog's OK lands through the same explicit native
/// command seam as `xlsx.range.set_custom_filter`.
export async function handleApplyAdvancedFilter(
  ctx: DataToolsContext,
  criteria: AdvancedFilterCriteria,
): Promise<string | null> {
  const runtime = ctx.univerRef.current
  const worksheet = runtime?.univerAPI.getActiveWorkbook()?.getActiveSheet()
  const filter = worksheet?.getFilter()
  if (!runtime || !worksheet || !filter) return t('appAdvFilterGone')
  const range = filter.getRange().getRange()
  const result = await setWorkbookCustomFilter(
    runtime,
    worksheet,
    range,
    range.startColumn + criteria.colId,
    criteria.and ? 'and' : 'or',
    criteria.filters.map(({ operator, val }) => ({ operator, value: val })),
  )
  if (result !== 'changed') return t('appAdvFilterFailed')
  ctx.setMessage(t('appAdvFilterApplied'))
  return null
}

export interface WorkbookSubtotalConfig {
  readonly groupCol: number
  readonly valueCol: number
  readonly agg: 'sum' | 'count' | 'average'
}

/** Shared retained action for the Subtotal dialog and the explicit registry tracer. */
export function applyWorkbookSubtotals(
  worksheet: UniverWorksheet,
  range: ReturnType<UniverWorksheet['getRange']>,
  config: WorkbookSubtotalConfig,
): number {
  const values = range.getValues()
  if (values.length < 2) throw new Error('subtotal_requires_header')
  const startRow = range.getRow()
  const startColumn = range.getColumn()
  const groupOffset = config.groupCol - startColumn
  const valueOffset = config.valueCol - startColumn
  if (
    groupOffset < 0 ||
    groupOffset >= range.getWidth() ||
    valueOffset < 0 ||
    valueOffset >= range.getWidth()
  ) {
    throw new Error('subtotal_columns_outside_range')
  }
  if (groupOffset === valueOffset) throw new Error('subtotal_same_columns')
  const body = values.slice(1)
  if (body.length === 0) throw new Error('subtotal_no_data_rows')
  // SUBTOTAL function numbers: 9 = SUM, 3 = COUNTA, 1 = AVERAGE.
  const fn = config.agg === 'sum' ? 9 : config.agg === 'count' ? 3 : 1
  const valueColumn = columnLetter(config.valueCol)
  const groups: { label: string; startAbs: number; endAbs: number }[] = []
  let currentLabel = String(body[0]?.[groupOffset] ?? '')
  let groupStart = startRow + 1
  for (let index = 1; index <= body.length; index++) {
    const label = index < body.length ? String(body[index]?.[groupOffset] ?? '') : null
    if (label !== currentLabel) {
      groups.push({ label: currentLabel, startAbs: groupStart, endAbs: startRow + index })
      if (label !== null) {
        currentLabel = label
        groupStart = startRow + 1 + index
      }
    }
  }
  if (groups.length > 200) throw new Error('subtotal_too_many_groups')
  const writeTotalRow = (row: number, label: string, from: number, to: number): void => {
    const labelCell = worksheet.getRange(row, config.groupCol, 1, 1)
    labelCell.setValue({ v: label })
    labelCell.setFontWeight('bold')
    const valueCell = worksheet.getRange(row, config.valueCol, 1, 1)
    valueCell.setValue({
      f: `=SUBTOTAL(${fn},$${valueColumn}$${from + 1}:$${valueColumn}$${to + 1})`,
    })
    valueCell.setFontWeight('bold')
  }
  // Bottom-up so earlier group indices stay valid across the inserts.
  for (const group of [...groups].reverse()) {
    worksheet.insertRowsBefore(group.endAbs + 1, 1)
    writeTotalRow(group.endAbs + 1, `${group.label} Total`, group.startAbs, group.endAbs)
  }
  const lastRowAbs = startRow + body.length + groups.length
  worksheet.insertRowsBefore(lastRowAbs + 1, 1)
  // SUBTOTAL skips the nested per-group subtotal rows inside the span.
  writeTotalRow(lastRowAbs + 1, 'Grand Total', startRow + 1, lastRowAbs)
  return groups.length
}

export function handleCreateSubtotal(ctx: DataToolsContext, config: WorkbookSubtotalConfig): string | null {
  const runtime = ctx.univerRef.current
  if (!runtime) return t('appWorkbookNotReady')
  const workbook = runtime.univerAPI.getActiveWorkbook()
  const worksheet = workbook?.getActiveSheet()
  const range = workbook?.getActiveRange()
  if (!workbook || !worksheet || !range) return t('appSelectSourceRangeFirst')
  try {
    const groups = applyWorkbookSubtotals(worksheet, range, config)
    ctx.setMessage(t('appSubtotalsInserted', { count: groups }))
  } catch (error: unknown) {
    const code = error instanceof Error ? error.message : ''
    if (code === 'subtotal_requires_header') return t('appSelectRangeWithHeader')
    if (code === 'subtotal_same_columns') return t('appSubtotalSameColumns')
    if (code === 'subtotal_no_data_rows') return t('appRangeNoDataRows')
    if (code === 'subtotal_too_many_groups') return t('appSubtotalTooManyGroups')
    if (code === 'subtotal_columns_outside_range') return t('appSelectSourceRangeFirst')
    return code || t('appSubtotalInsertFailed')
  }
  return null
}

export function consolidateDefaultReference(ctx: DataToolsContext): string {
  const workbook = ctx.univerRef.current?.univerAPI.getActiveWorkbook()
  const worksheet = workbook?.getActiveSheet()
  const range = workbook?.getActiveRange()
  if (!worksheet || !range || (range.getHeight() === 1 && range.getWidth() === 1)) return ''
  const name = worksheet.getSheetName()
  const prefix = /^[A-Za-z_][A-Za-z0-9_.]*$/.test(name) ? name : `'${name.replaceAll("'", "''")}'`
  return `${prefix}!${range.getA1Notation()}`
}

export interface WorkbookConsolidationSource {
  readonly worksheet: UniverWorksheet
  readonly area: ConsolidateArea
}

/** Shared formula-matrix action for the Consolidate dialog and registry tracer. */
export function applyWorkbookConsolidation(
  targetWorksheet: UniverWorksheet,
  target: { readonly row: number; readonly column: number },
  sources: readonly WorkbookConsolidationSource[],
  fn: ConsolidateConfig['fn'],
  leftLabels: boolean,
): { readonly rows: number; readonly columns: number } {
  const areas = sources.map(({ area }) => area)
  let matrix: OutputCell[][]
  if (leftLabels) {
    const areaLabels = sources.map(({ worksheet, area }) =>
      worksheet
        .getRange(area.startRow, area.startColumn, area.rows, 1)
        .getValues()
        .map((row) => String(row?.[0] ?? '')),
    )
    const built = buildLabelMatrix(fn, areas, areaLabels, target)
    if (typeof built === 'string') throw new Error(built)
    matrix = built
  } else {
    matrix = buildPositionMatrix(fn, areas)
  }
  const columns = matrix[0]?.length ?? 0
  if (matrix.length === 0 || columns === 0) throw new Error('consolidate_empty_sources')
  if (
    targetOverlapsSource(areas, targetWorksheet.getSheetName(), {
      ...target,
      rows: matrix.length,
      columns,
    })
  ) {
    throw new Error('consolidate_target_overlap')
  }
  const range = targetWorksheet.getRange(target.row, target.column, matrix.length, columns)
  applyWorkbookCellMatrix(
    range,
    matrix.map((line) =>
      line.map((cell) => (cell.f !== undefined ? { f: cell.f } : { v: cell.v ?? null })),
    ),
  )
  range.activate()
  return { rows: matrix.length, columns }
}

export function handleCreateConsolidate(
  ctx: DataToolsContext,
  config: ConsolidateConfig,
): string | null {
  const runtime = ctx.univerRef.current
  if (!runtime) return t('appWorkbookNotReady')
  const workbook = runtime.univerAPI.getActiveWorkbook()
  const worksheet = workbook?.getActiveSheet()
  const active = workbook?.getActiveRange()
  if (!workbook || !worksheet || !active) return t('appSelectTargetCellFirst')
  const activeSheetName = worksheet.getSheetName()
  const sheets = workbook.getSheets()
  const areas: ConsolidateArea[] = []
  const sources: UniverWorksheet[] = []
  for (const reference of config.references) {
    const parsed = parseConsolidateReference(reference)
    if (!parsed) return t('appConsolidateBadRef', { ref: reference })
    const name = parsed.sheetName ?? activeSheetName
    const source = sheets.find((sheet) => sheet.getSheetName().toLowerCase() === name.toLowerCase())
    if (!source) return t('appNoSheetNamed', { name })
    sources.push(source)
    areas.push({
      sheetName: source.getSheetName(),
      startRow: parsed.range.startRow,
      startColumn: parsed.range.startColumn,
      rows: parsed.range.endRow - parsed.range.startRow + 1,
      columns: parsed.range.endColumn - parsed.range.startColumn + 1,
    })
  }
  const target = { row: active.getRow(), column: active.getColumn() }
  try {
    const result = applyWorkbookConsolidation(
      worksheet,
      target,
      areas.map((area, index) => ({ worksheet: sources[index] as UniverWorksheet, area })),
      config.fn,
      config.leftLabels,
    )
    ctx.setMessage(
      t('appConsolidateDone', {
        count: areas.length,
        cell: `${columnLetter(target.column)}${target.row + 1}`,
        rows: result.rows,
        columns: result.columns,
      }),
    )
  } catch (error: unknown) {
    const message = error instanceof Error ? error.message : ''
    if (message === 'consolidate_empty_sources') return t('appConsolidateEmptySources')
    if (message === 'consolidate_target_overlap') return t('appConsolidateOverlap')
    return message || t('appConsolidateWriteFailed')
  }
  return null
}

/// Outline groups: level edits journal directly (Univer has no outline
/// model, so there is no command to undo); Hide/Show Detail rides the
/// normal hidden-rows pipeline plus a collapsed flag on the summary line.
export function handleOutline(
  ctx: DataToolsContext,
  action: 'group' | 'ungroup' | 'hide-detail' | 'show-detail',
  axis: 'rows' | 'cols',
): void {
  const runtime = ctx.univerRef.current
  if (!runtime) return
  const state = ctx.lazyWorkbookRef.current
  if (!state) {
    ctx.setMessage(t('appOutlineNeedsFile'))
    return
  }
  const workbook = runtime.univerAPI.getActiveWorkbook()
  const worksheet = workbook?.getActiveSheet()
  const range = workbook?.getActiveRange()
  if (!workbook || !worksheet || !range) {
    ctx.setMessage(t('appOutlineSelectFirst'))
    return
  }
  const sheetId = worksheet.getSheetId()
  if (isSheetRemoved(state.editJournal, sheetId)) return
  const start = axis === 'rows' ? range.getRow() : range.getColumn()
  const end = axis === 'rows' ? start + range.getHeight() - 1 : start + range.getWidth() - 1
  const outline = sheetOutline(state, sheetId)
  const entries = axis === 'rows' ? outline.rows : outline.cols
  const operationAxis: WorkbookOutlineAxis = axis === 'rows' ? 'rows' : 'columns'

  if (action === 'hide-detail' || action === 'show-detail') {
    // The selection is the group's detail span; the summary row/column is
    // the next one after it (Excel's summaryBelow/summaryRight default).
    const count = end - start + 1
    const maximum = axis === 'rows' ? worksheet.getMaxRows() : worksheet.getMaxColumns()
    if (end + 1 >= maximum) {
      ctx.setMessage(t('appOutlineSelectFirst'))
      return
    }
    const hidden = action === 'hide-detail'
    applyWorkbookOutlineDetailVisibility(
      runtime,
      state,
      worksheet,
      operationAxis,
      start,
      count,
      hidden,
      ctx.setPendingEdits,
    )
    ctx.setMessage(hidden ? t('appDetailHidden') : t('appDetailShown'))
    return
  }

  // Group/Ungroup shifts each contiguous run of equal levels by ±1
  // (levels clamp to 0-7). Runs already at the boundary are skipped.
  const delta = action === 'group' ? 1 : -1
  const ops: { start: number; end: number; level: number }[] = []
  let runStart = start
  let runLevel = entries.get(start)?.level ?? 0
  const closeRun = (runEnd: number): void => {
    const level = Math.min(7, Math.max(0, runLevel + delta))
    if (level !== runLevel) ops.push({ start: runStart, end: runEnd, level })
  }
  for (let index = start + 1; index <= end; index += 1) {
    const level = entries.get(index)?.level ?? 0
    if (level !== runLevel) {
      closeRun(index - 1)
      runStart = index
      runLevel = level
    }
  }
  closeRun(end)
  if (ops.length === 0) {
    ctx.setMessage(action === 'group' ? t('appOutlineMaxLevel') : t('appNothingToUngroup'))
    return
  }
  applyWorkbookOutlineLevels(
    runtime,
    state,
    worksheet,
    operationAxis,
    ops.map((op) => ({ start: op.start, count: op.end - op.start + 1, level: op.level })),
    ctx.setPendingEdits,
  )
  ctx.setMessage(
    action === 'group'
      ? axis === 'rows'
        ? t('appRowsGrouped')
        : t('appColsGrouped')
      : axis === 'rows'
        ? t('appRowsUngrouped')
        : t('appColsUngrouped'),
  )
}

/// Home → Format as Table: the manual entry over the same engine as the
/// AI add_table op (Univer table rendering + journal + native table part).
export function handleFormatAsTable(ctx: DataToolsContext, style: string): void {
  const runtime = ctx.univerRef.current
  if (!runtime) return
  const state = ctx.lazyWorkbookRef.current
  if (!state) {
    ctx.setMessage(t('appTablesNeedFile'))
    return
  }
  const workbook = runtime.univerAPI.getActiveWorkbook()
  const worksheet = workbook?.getActiveSheet()
  const range = workbook?.getActiveRange()
  if (!workbook || !worksheet || !range) {
    ctx.setMessage(t('appTableSelectRange'))
    return
  }
  const sheetId = worksheet.getSheetId()
  if (isSheetRemoved(state.editJournal, sheetId)) return
  const startRow = range.getRow()
  const startColumn = range.getColumn()
  const endRow = startRow + range.getHeight() - 1
  const endColumn = startColumn + range.getWidth() - 1
  try {
    applyWorkbookTableAdd(runtime, state, {
      op: 'add_table',
      sheetId,
      range: `${columnLabel(startColumn)}${startRow + 1}:${columnLabel(endColumn)}${endRow + 1}`,
      style,
      bandedRows: true,
    })
    ctx.setPendingEdits(journalSize(state.editJournal))
    ctx.setMessage(t('appTableCreated'))
  } catch (error: unknown) {
    ctx.setMessage(error instanceof Error ? error.message : t('appTableCreateFailed'))
  }
}
