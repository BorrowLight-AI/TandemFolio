import { ICommandService, IUndoRedoService } from '@univerjs/core'
import { describe, expect, it } from 'vitest'

import { executeXlsxOperation } from '../src/renderer/operations/registry'
import { isManualCalculation } from '../src/renderer/calc-options'

function makeServices() {
  const listeners: ((command: { id: string; params?: unknown }) => void)[] = []
  const registered: unknown[] = []
  const undo: unknown[] = []
  const executed: { id: string; params?: unknown }[] = []
  let workbookCalculations = 0
  const commandService = {
    _commandExecutionStack: [],
    beforeCommandExecuted: (listener: (command: { id: string }) => void) => listeners.push(listener),
    registerCommand: (command: unknown) => registered.push(command),
    syncExecuteCommand: (id: string, params?: unknown) => {
      executed.push({ id, params })
      return true
    },
  }
  const undoService = { pushUndoRedo: (item: unknown) => undo.push(item) }
  const worksheet = {
    getSheetId: () => 's1',
    getSheetName: () => 'Sheet1',
    getSheet: () => ({ getRowCount: () => 20, getColumnCount: () => 5 }),
  }
  const runtime = {
    univer: {
      __getInjector: () => ({
        get: (token: unknown) => (token === ICommandService ? commandService : token === IUndoRedoService ? undoService : undefined),
      }),
    },
    univerAPI: {
      getActiveWorkbook: () => ({ getId: () => 'wb', getActiveSheet: () => worksheet }),
      getFormula: () => ({ executeCalculation: () => { workbookCalculations += 1 } }),
    },
  }
  return {
    runtime,
    services: { runtime: () => runtime as never },
    undo,
    executed,
    workbookCalculations: () => workbookCalculations,
  }
}

describe('typed calculation operations', () => {
  it('sets manual mode through the renderer state and records native undo', async () => {
    const fixture = makeServices()
    const result = await executeXlsxOperation(
      { operation: 'xlsx.calculation.set_mode', arguments: { mode: 'manual' } },
      fixture.services,
    )
    expect(result).toMatchObject({ handled: true, ok: true, output: { mode: 'manual' } })
    expect(isManualCalculation(fixture.runtime as never)).toBe(true)
    expect(fixture.undo).toHaveLength(1)
  })

  it('recalculates the workbook or active sheet through the same engine', async () => {
    const fixture = makeServices()
    await executeXlsxOperation(
      { operation: 'xlsx.calculation.recalculate', arguments: { scope: 'workbook' } },
      fixture.services,
    )
    await executeXlsxOperation(
      { operation: 'xlsx.calculation.recalculate', arguments: { scope: 'sheet' } },
      fixture.services,
    )
    expect(fixture.workbookCalculations()).toBe(1)
    expect(fixture.executed.some((entry) => entry.id.includes('calculation-start'))).toBe(true)
  })
})
