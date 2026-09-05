import { describe, expect, it } from 'vitest'

import {
  calculateSheet,
  applyCalculationMode,
  isManualCalculation,
  resetCalculationMode,
  setManualCalculation,
} from '../src/renderer/calc-options'
import type { UniverRuntime } from '../src/renderer/univer-state'

const START = 'formula.mutation.set-formula-calculation-start'

function makeRuntime() {
  const listeners: ((command: { id: string; params?: Record<string, unknown> }) => void)[] = []
  const executed: { id: string; params?: Record<string, unknown> }[] = []
  const stack: object[] = []
  const commandService = {
    _commandExecutionStack: stack,
    beforeCommandExecuted: (listener: (command: { id: string }) => void) => listeners.push(listener),
    syncExecuteCommand: (id: string, params?: Record<string, unknown>) => {
      const info = { id, params }
      stack.push(info)
      try {
        for (const listener of listeners) listener(info)
      } catch {
        return false
      }
      executed.push(info)
      stack.splice(stack.indexOf(info), 1)
      return true
    },
  }
  const worksheet = {
    getSheet: () => ({ getRowCount: () => 100, getColumnCount: () => 10 }),
    getSheetId: () => 's1',
    getSheetName: () => 'Sheet1',
  }
  const runtime = {
    univer: { __getInjector: () => ({ get: () => commandService }) },
    univerAPI: { getActiveWorkbook: () => ({ getActiveSheet: () => worksheet, getId: () => 'wb' }) },
  } as unknown as UniverRuntime
  return {
    runtime,
    executed,
    stack,
    recalc: (params: Record<string, unknown>) => commandService.syncExecuteCommand(START, params),
    listenerCount: () => listeners.length,
  }
}

describe('calculation options', () => {
  it('manual mode vetoes automatic recalc without stranding its stack entry', () => {
    const { runtime, recalc, executed, stack } = makeRuntime()
    setManualCalculation(runtime, true)
    expect(recalc({ dirtyRanges: [{ unitId: 'wb' }] })).toBe(false)
    expect(executed).toHaveLength(0)
    expect(stack).toHaveLength(0)
  })

  it('allows forced calculation and resets per workbook', () => {
    const { runtime, recalc, executed } = makeRuntime()
    setManualCalculation(runtime, true)
    recalc({ forceCalculation: true })
    expect(executed).toHaveLength(1)
    expect(isManualCalculation(runtime)).toBe(true)
    resetCalculationMode(runtime)
    expect(isManualCalculation(runtime)).toBe(false)
  })

  it('calculate-sheet dirties the full active sheet', () => {
    const { runtime, executed } = makeRuntime()
    setManualCalculation(runtime, true)
    calculateSheet(runtime)
    expect(executed[0]?.params).toMatchObject({
      dirtyNameMap: { wb: { s1: 'Sheet1' } },
      dirtyRanges: [{ unitId: 'wb', sheetId: 's1', range: { endRow: 99, endColumn: 9 } }],
    })
  })

  it('installs one hook for each runtime', () => {
    const first = makeRuntime()
    setManualCalculation(first.runtime, true)
    setManualCalculation(first.runtime, false)
    expect(first.listenerCount()).toBe(1)
  })

  it('records a reversible calculation-mode change', () => {
    const { runtime } = makeRuntime()
    let step: { undo(): void; redo(): void } | undefined
    applyCalculationMode(runtime, true, (value) => {
      step = value
    })
    expect(isManualCalculation(runtime)).toBe(true)
    step?.undo()
    expect(isManualCalculation(runtime)).toBe(false)
    step?.redo()
    expect(isManualCalculation(runtime)).toBe(true)
  })
})
