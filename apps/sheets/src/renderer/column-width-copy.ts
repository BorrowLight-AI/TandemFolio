import { ICommandService, IUndoRedoService, IUniverInstanceService } from '@univerjs/core'
import {
  SetWorksheetColWidthMutation,
  SetWorksheetColWidthMutationFactory,
  type ISetWorksheetColWidthMutationParams,
} from '@univerjs/sheets'

import type { UniverRuntime, UniverWorksheet } from './univer-state'

/** Copies one explicit column-width span as a single native workbook history entry. */
export function applyWorkbookColumnWidthCopy(
  runtime: UniverRuntime,
  sourceWorksheet: UniverWorksheet,
  sourceColumn: number,
  destinationWorksheet: UniverWorksheet,
  destinationColumn: number,
  count: number,
): boolean {
  const workbook = runtime.univerAPI.getActiveWorkbook()
  if (!workbook) return false
  const injector = runtime.univer.__getInjector()
  const destinationModel = injector
    .get(IUniverInstanceService)
    .getUniverSheetInstance(workbook.getId())
    ?.getSheetBySheetId(destinationWorksheet.getSheetId())
  if (!destinationModel) return false

  const colWidth = Object.fromEntries(
    Array.from({ length: count }, (_, offset) => [
      destinationColumn + offset,
      sourceWorksheet.getColumnWidth(sourceColumn + offset),
    ]),
  )
  const params: ISetWorksheetColWidthMutationParams = {
    unitId: workbook.getId(),
    subUnitId: destinationWorksheet.getSheetId(),
    ranges: [
      {
        startRow: 0,
        endRow: destinationModel.getMaxRows() - 1,
        startColumn: destinationColumn,
        endColumn: destinationColumn + count - 1,
      },
    ],
    colWidth,
  }
  const undoParams = SetWorksheetColWidthMutationFactory(params, destinationModel)
  if (!injector.get(ICommandService).syncExecuteCommand(SetWorksheetColWidthMutation.id, params)) {
    return false
  }
  injector.get(IUndoRedoService).pushUndoRedo({
    unitID: params.unitId,
    undoMutations: [{ id: SetWorksheetColWidthMutation.id, params: undoParams }],
    redoMutations: [{ id: SetWorksheetColWidthMutation.id, params }],
  })
  return true
}
