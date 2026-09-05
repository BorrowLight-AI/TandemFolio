import { CustomCommandExecutionError, ICommandService } from '@univerjs/core'

import type { UniverRuntime } from './univer-state'

const START_MUTATION = 'formula.mutation.set-formula-calculation-start'

interface CalcState {
  manual: boolean
  allowNext: boolean
}

const states = new WeakMap<UniverRuntime, CalcState>()

function stateFor(runtime: UniverRuntime): CalcState {
  const existing = states.get(runtime)
  if (existing) return existing
  const state: CalcState = { manual: false, allowNext: false }
  states.set(runtime, state)
  const commandService = runtime.univer.__getInjector().get(ICommandService)
  commandService.beforeCommandExecuted((command) => {
    if (command.id !== START_MUTATION || !state.manual || state.allowNext) return
    const params = command.params as { forceCalculation?: boolean } | undefined
    if (params?.forceCalculation) return
    const stack = (commandService as unknown as { _commandExecutionStack?: unknown[] })
      ._commandExecutionStack
    const index = stack?.indexOf(command) ?? -1
    if (index >= 0) stack?.splice(index, 1)
    throw new CustomCommandExecutionError('manual calculation mode')
  })
  return state
}

export function isManualCalculation(runtime: UniverRuntime | null): boolean {
  return runtime !== null && stateFor(runtime).manual
}

export function setManualCalculation(runtime: UniverRuntime, manual: boolean): void {
  stateFor(runtime).manual = manual
}

export function applyCalculationMode(
  runtime: UniverRuntime,
  manual: boolean,
  pushUndo?: (step: { undo(): void; redo(): void }) => void,
): void {
  const previous = isManualCalculation(runtime)
  if (previous === manual) return
  setManualCalculation(runtime, manual)
  pushUndo?.({
    undo: () => setManualCalculation(runtime, previous),
    redo: () => setManualCalculation(runtime, manual),
  })
}

export function resetCalculationMode(runtime: UniverRuntime | null): void {
  if (runtime) stateFor(runtime).manual = false
}

export function calculateNow(runtime: UniverRuntime): void {
  runtime.univerAPI.getFormula().executeCalculation()
}

export function calculateSheet(runtime: UniverRuntime): void {
  const workbook = runtime.univerAPI.getActiveWorkbook()
  const worksheet = workbook?.getActiveSheet()
  if (!workbook || !worksheet) return
  const sheet = worksheet.getSheet()
  const commandService = runtime.univer.__getInjector().get(ICommandService)
  const unitId = workbook.getId()
  const sheetId = worksheet.getSheetId()
  const state = stateFor(runtime)
  state.allowNext = true
  try {
    commandService.syncExecuteCommand(START_MUTATION, {
      commands: [],
      dirtyRanges: [
        {
          unitId,
          sheetId,
          range: {
            startRow: 0,
            endRow: sheet.getRowCount() - 1,
            startColumn: 0,
            endColumn: sheet.getColumnCount() - 1,
          },
        },
      ],
      dirtyNameMap: { [unitId]: { [sheetId]: worksheet.getSheetName() } },
    })
  } finally {
    state.allowNext = false
  }
}
