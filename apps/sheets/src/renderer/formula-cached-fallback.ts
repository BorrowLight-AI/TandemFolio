import { CellValueType } from '@univerjs/core'
import { ERROR_TYPE_SET, ErrorType } from '@univerjs/engine-formula'
import { INTERCEPTOR_POINT, SheetInterceptorService } from '@univerjs/sheets'

import type { LazyWorkbookState, UniverRuntime } from './univer-state'

const COORDINATE_SHIFTING_OPS = new Set([
  'insert-rows',
  'remove-rows',
  'insert-cols',
  'remove-cols',
  'move-rows',
])

export function isPlainArithmeticFormula(formula: string): boolean {
  let body = formula.startsWith('=') ? formula.slice(1) : formula
  body = body.replace(/"(?:[^"]|"")*"/g, '""').replace(/\$/g, '')
  if (/[[{#]/.test(body)) return false
  body = body.replace(/'(?:[^']|'')*'!/g, '').replace(/[A-Za-z_][\w.]*!/g, '')
  body = body.replace(/(?<![\w.])\d+(?:\.\d+)?E[+-]?\d+/gi, '0')
  for (const match of body.matchAll(/[A-Za-z_][\w.]*/g)) {
    const token = match[0]
    if (/^\s*\(/.test(body.slice((match.index ?? 0) + token.length))) return false
    if (!/^[A-Za-z]{1,3}\d+$/.test(token) && !/^(?:TRUE|FALSE)$/i.test(token)) return false
  }
  return true
}

export function installCachedValueFallbackInterceptor(
  runtime: UniverRuntime,
  lazyWorkbookRef: { readonly current: LazyWorkbookState | null },
): { dispose(): void } {
  const interceptorService = runtime.univer.__getInjector().get(SheetInterceptorService)
  return interceptorService.intercept(INTERCEPTOR_POINT.CELL_CONTENT, {
    priority: 9997,
    handler: (cell, location, next) => {
      const value = cell?.v
      if (typeof value !== 'string' || !ERROR_TYPE_SET.has(value as ErrorType)) return next(cell)
      if (value === ErrorType.VALUE) {
        const formula = location.rawData?.f
        if (typeof formula === 'string' && isPlainArithmeticFormula(formula)) return next(cell)
      }
      const state = lazyWorkbookRef.current
      if (!state) return next(cell)
      const sheetId = location.subUnitId
      const ops = state.editJournal.structuralOps.get(sheetId)
      if (ops?.some((operation) => COORDINATE_SHIFTING_OPS.has(operation.kind))) return next(cell)
      const key = `${location.row}:${location.col}`
      const edits = state.editJournal.cells.get(sheetId)
      if (edits?.get(key)?.hasValue) return next(cell)
      if (value !== ErrorType.NAME && edits) {
        for (const entry of edits.values()) if (entry.hasValue) return next(cell)
      }
      const cached = state.cachedFormulaValues.get(sheetId)?.get(key)
      if (cached === undefined || cached === value) return next(cell)
      return next({
        ...cell,
        v: cached,
        t:
          typeof cached === 'number'
            ? CellValueType.NUMBER
            : typeof cached === 'boolean'
              ? CellValueType.BOOLEAN
              : CellValueType.STRING,
      })
    },
  })
}
