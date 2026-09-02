/**
 * Formula view — Excel's Show Formulas (⌘` / Ctrl+`), persisted per sheet as
 * sheetView/@showFormulas.
 *
 * Univer's RENDER_RAW_FORMULA_KEY only reaches the rich-text/rotated cell
 * path in this version, so plain cells never change: the actual value swap
 * happens in a CELL_CONTENT interceptor below. The context key is still
 * toggled for its side effect — every sheet skeleton resets its cache and
 * repaints when it flips.
 */
import { CellValueType, IContextService } from '@univerjs/core'
import { RENDER_RAW_FORMULA_KEY } from '@univerjs/preset-sheets-core'
import { INTERCEPTOR_POINT, SheetInterceptorService } from '@univerjs/sheets'

import { isSheetRemoved, journalSize, recordPageSetup } from './edit-journal'
import type { LazyWorkbookState, UniverRuntime } from './univer-state'

type FormulaViewState = Pick<LazyWorkbookState, 'editJournal' | 'showFormulaSheets'>
export interface FormulaViewUndoStep {
  undo(): void
  redo(): void
}

/// Formula-view sheets for blank/demo workbooks (no LazyWorkbookState).
const demoFormulaViewSheets = new Set<string>()

export function formulaViewSheets(
  state: Pick<LazyWorkbookState, 'showFormulaSheets'> | null,
): Set<string> {
  return state ? state.showFormulaSheets : demoFormulaViewSheets
}

/// Sync the (global) raw-formula render key to the sheet's per-sheet flag;
/// flipping it makes every skeleton drop its text cache and repaint.
export function applyShowFormulasView(
  runtime: UniverRuntime,
  state: Pick<LazyWorkbookState, 'showFormulaSheets'> | null,
  sheetId: string,
): void {
  const contextService = runtime.univer.__getInjector().get(IContextService)
  const next = formulaViewSheets(state).has(sheetId)
  if (Boolean(contextService.getContextValue(RENDER_RAW_FORMULA_KEY)) !== next) {
    contextService.setContextValue(RENDER_RAW_FORMULA_KEY, next)
  }
}

/** Applies an explicit persisted formula-view state to one worksheet. */
export function applyWorkbookFormulaView(
  runtime: UniverRuntime,
  state: FormulaViewState | null,
  sheetId: string,
  enabled: boolean,
  setPendingEdits?: (count: number) => void,
  pushUndo?: (step: FormulaViewUndoStep) => void,
): void {
  const sheets = formulaViewSheets(state)
  const previous = sheets.has(sheetId)
  if (previous === enabled) return
  const apply = (next: boolean): void => {
    if (next) sheets.add(sheetId)
    else sheets.delete(sheetId)
    applyShowFormulasView(runtime, state, sheetId)
    if (state && !isSheetRemoved(state.editJournal, sheetId)) {
      recordPageSetup(state.editJournal, sheetId, { showFormulas: next })
      setPendingEdits?.(journalSize(state.editJournal))
    }
  }
  apply(enabled)
  pushUndo?.({ undo: () => apply(previous), redo: () => apply(enabled) })
}

/// Formula text for a cell whose model carries no `f`: the harvested
/// per-sheet index, unless the coordinates are unreliable (structural
/// edits shifted them) or the user overwrote the cell this session.
function indexedFormulaText(
  state: LazyWorkbookState | null,
  sheetId: string,
  row: number,
  column: number,
): string | undefined {
  if (!state) return undefined
  if ((state.editJournal.structuralOps.get(sheetId)?.length ?? 0) > 0) return undefined
  if (state.editJournal.cells.get(sheetId)?.has(`${row}:${column}`)) return undefined
  return state.formulaText.get(sheetId)?.get(`${row}:${column}`)
}

export function installFormulaViewInterceptor(
  runtime: UniverRuntime,
  lazyWorkbookRef: { readonly current: LazyWorkbookState | null },
): { dispose(): void } {
  const injector = runtime.univer.__getInjector()
  const interceptorService = injector.get(SheetInterceptorService)
  return interceptorService.intercept(INTERCEPTOR_POINT.CELL_CONTENT, {
    // Above NUMFMT (10): a formula cell in formula view shows its formula
    // text, not the formatted value, so the chain stops here.
    priority: 9999,
    handler: (cell, location, next) => {
      const state = lazyWorkbookRef.current
      if (!formulaViewSheets(state).has(location.subUnitId)) {
        return next(cell)
      }
      const formula =
        location.rawData?.f ??
        indexedFormulaText(state, location.subUnitId, location.row, location.col)
      if (typeof formula !== 'string' || formula === '') return next(cell)
      return { ...(cell ?? {}), v: formula, t: CellValueType.STRING, p: null }
    },
  })
}

/// Streamed workbooks whose closure gave up have no `f` in the model, so the
/// formula bar shows only values. Inject the harvested formula text
/// into the view model: the grid ignores `f` (displayRawFormula=false) and
/// the engine reads the raw cell matrix, so this is display-only.
export function installFormulaTextInterceptor(
  runtime: UniverRuntime,
  lazyWorkbookRef: { readonly current: LazyWorkbookState | null },
): { dispose(): void } {
  const injector = runtime.univer.__getInjector()
  const interceptorService = injector.get(SheetInterceptorService)
  return interceptorService.intercept(INTERCEPTOR_POINT.CELL_CONTENT, {
    priority: 9998,
    handler: (cell, location, next) => {
      if (location.rawData?.f) return next(cell)
      const formula = indexedFormulaText(
        lazyWorkbookRef.current,
        location.subUnitId,
        location.row,
        location.col,
      )
      if (!formula) return next(cell)
      return next({ ...(cell ?? {}), f: formula })
    },
  })
}
