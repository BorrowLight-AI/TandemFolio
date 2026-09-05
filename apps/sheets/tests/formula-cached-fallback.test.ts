import { CellValueType } from '@univerjs/core'
import { describe, expect, it } from 'vitest'

import {
  installCachedValueFallbackInterceptor,
  isPlainArithmeticFormula,
} from '../src/renderer/formula-cached-fallback'

interface Interceptor {
  priority: number
  handler: (
    cell: Record<string, unknown> | undefined,
    location: { subUnitId: string; row: number; col: number; rawData?: { f?: string } },
    next: (cell: Record<string, unknown> | undefined) => unknown,
  ) => unknown
}

function captureInterceptor(lazyWorkbookRef: { current: unknown }): Interceptor {
  let captured: Interceptor | undefined
  const runtime = {
    univer: {
      __getInjector: () => ({
        get: () => ({
          intercept: (_point: unknown, interceptor: Interceptor) => {
            captured = interceptor
            return { dispose: () => undefined }
          },
        }),
      }),
    },
  }
  installCachedValueFallbackInterceptor(runtime as never, lazyWorkbookRef as { current: null })
  if (!captured) throw new Error('interceptor not registered')
  return captured
}

function makeState(cached: Record<string, string | number | boolean>) {
  return {
    editJournal: { structuralOps: new Map(), cells: new Map() },
    cachedFormulaValues: new Map([['sheet-1', new Map(Object.entries(cached))]]),
  }
}

const at = (row: number, col: number) => ({ subUnitId: 'sheet-1', row, col })
const passthrough = (cell: Record<string, unknown> | undefined) => cell

describe('isPlainArithmeticFormula', () => {
  it('accepts refs/operators and rejects function or structured refs', () => {
    expect(isPlainArithmeticFormula('=(D23+D24+D25)/3')).toBe(true)
    expect(isPlainArithmeticFormula('=1E-3*A1')).toBe(true)
    expect(isPlainArithmeticFormula('=SUM(A1:A3)')).toBe(false)
    expect(isPlainArithmeticFormula('=Table1[Total]+1')).toBe(false)
  })
})

describe('installCachedValueFallbackInterceptor', () => {
  it('uses a typed file-cached value for an engine error', () => {
    const { handler } = captureInterceptor({ current: makeState({ '0:0': 51.6373 }) })
    expect(handler({ v: '#NAME?', f: '=X[Y]' }, at(0, 0), passthrough)).toEqual({
      v: 51.6373,
      f: '=X[Y]',
      t: CellValueType.NUMBER,
    })
  })

  it('does not mask user-owned cells or genuine arithmetic errors', () => {
    const state = makeState({ '0:0': 42 })
    state.editJournal.cells.set('sheet-1', new Map([['0:0', { hasValue: true }]]))
    const { handler } = captureInterceptor({ current: state })
    const error = { v: '#NAME?' }
    expect(handler(error, at(0, 0), passthrough)).toBe(error)

    const clean = makeState({ '0:0': 42 })
    const cleanHandler = captureInterceptor({ current: clean }).handler
    const arithmetic = { v: '#VALUE!' }
    expect(
      cleanHandler(arithmetic, { ...at(0, 0), rawData: { f: '=D1+D2' } }, passthrough),
    ).toBe(arithmetic)
  })

  it('keeps cache coordinates for layout-only operations', () => {
    const state = makeState({ '0:0': true })
    state.editJournal.structuralOps.set('sheet-1', [{ kind: 'set-row-size' }])
    const { handler } = captureInterceptor({ current: state })
    expect(handler({ v: '#N/A' }, at(0, 0), passthrough)).toMatchObject({
      v: true,
      t: CellValueType.BOOLEAN,
    })
  })
})
