import {
  ICommandService,
  IUndoRedoService,
  ObjectMatrix,
  STYLE_KEYS,
  type ICellData,
  type IStyleData,
} from '@univerjs/core'
import { LexerTreeBuilder } from '@univerjs/engine-formula'
import { SetRangeValuesMutation, SetRangeValuesUndoMutationFactory } from '@univerjs/sheets'

import type { UniverRuntime, UniverWorksheet } from './univer-state'

type UniverRange = ReturnType<UniverWorksheet['getRange']>

function replacementCellStyle(style: IStyleData | null): IStyleData | null {
  if (!style) return null
  return Object.fromEntries(
    STYLE_KEYS.map((key) => [key, style[key] ?? null]),
  ) as unknown as IStyleData
}

function replacementCellStyleWithoutBorders(style: IStyleData | null): IStyleData {
  return Object.fromEntries(
    STYLE_KEYS.filter((key) => key !== 'bd').map((key) => [key, style?.[key] ?? null]),
  ) as unknown as IStyleData
}

export function commitWorkbookRangeCells(
  runtime: UniverRuntime,
  destination: UniverRange,
  cellValue: ObjectMatrix<ICellData>,
): boolean {
  const injector = runtime.univer.__getInjector()
  const params = {
    unitId: runtime.univerAPI.getActiveWorkbook()!.getId(),
    subUnitId: destination.getSheetId(),
    cellValue: cellValue.getMatrix(),
  }
  const undoParams = SetRangeValuesUndoMutationFactory(injector, params)
  const commandService = injector.get(ICommandService)
  if (!commandService.syncExecuteCommand(SetRangeValuesMutation.id, params)) return false
  injector.get(IUndoRedoService).pushUndoRedo({
    unitID: params.unitId,
    undoMutations: [{ id: SetRangeValuesMutation.id, params: undoParams }],
    redoMutations: [{ id: SetRangeValuesMutation.id, params }],
  })
  return true
}

/**
 * Replaces one destination style matrix without routing through the value-edit command.
 *
 * Missing source style fields become explicit nulls so this behaves like Paste Formats
 * rather than merging onto the destination. Existing destination content and formulas
 * ride in the mutation payload unchanged, which keeps the save journal from mistaking
 * formula-engine result events for value replacements. The mutation and its inverse are
 * registered as one native workbook history entry.
 */
export function applyWorkbookRangeFormatCopy(
  runtime: UniverRuntime,
  source: UniverRange,
  destination: UniverRange,
): boolean {
  const cellValue = new ObjectMatrix<ICellData>()
  const destinationRow = destination.getRow()
  const destinationColumn = destination.getColumn()
  const destinationCells = destination.getCellDataGrid()
  const destinationFormulas = destination.getFormulas()
  for (const [rowOffset, row] of source.getCellStyles().entries()) {
    for (const [columnOffset, style] of row.entries()) {
      const formula = destinationFormulas[rowOffset]?.[columnOffset]
      cellValue.setValue(destinationRow + rowOffset, destinationColumn + columnOffset, {
        ...(destinationCells[rowOffset]?.[columnOffset] ?? {}),
        ...(formula ? { f: formula } : {}),
        s: replacementCellStyle((style?.getValue() as IStyleData | undefined) ?? null),
      })
    }
  }
  return commitWorkbookRangeCells(runtime, destination, cellValue)
}

/**
 * Copies one explicit cell matrix while retaining every destination border.
 *
 * Source cells, formulas, and non-border styles are snapshotted before the mutation.
 * Missing source style fields are explicit nulls so they replace destination formatting,
 * while omitting `bd` lets Univer's native style merge retain the destination border.
 */
export function applyWorkbookRangeCopyWithoutBorders(
  runtime: UniverRuntime,
  source: UniverRange,
  destination: UniverRange,
  sourceFormulaText?: ReadonlyMap<string, string>,
): boolean {
  const sourceCells = source.getCellDataGrid()
  const sourceFormulas = source.getFormulas()
  const sourceStyles = source.getCellStyles()
  const destinationRow = destination.getRow()
  const destinationColumn = destination.getColumn()
  const rowOffset = destinationRow - source.getRow()
  const columnOffset = destinationColumn - source.getColumn()
  const lexer = runtime.univer.__getInjector().get(LexerTreeBuilder)
  const copiedCells = Array.from({ length: source.getHeight() }, (_, row) =>
    Array.from({ length: source.getWidth() }, (_, column) => {
      const sourceCell = sourceCells[row]?.[column] ?? null
      const rawFormula =
        sourceFormulas[row]?.[column] ||
        sourceCell?.f ||
        sourceFormulaText?.get(`${source.getRow() + row}:${source.getColumn() + column}`)
      const formula = rawFormula
        ? rawFormula.startsWith('=')
          ? rawFormula
          : `=${rawFormula}`
        : null
      const style = sourceStyles[row]?.[column]?.getValue() as IStyleData | undefined
      return {
        ...(formula
          ? { f: lexer.moveFormulaRefOffset(formula, columnOffset, rowOffset) }
          : { v: sourceCell?.v ?? null, f: null }),
        p: sourceCell?.p ?? null,
        ref: null,
        xf: sourceCell?.xf ?? null,
        si: null,
        custom: sourceCell?.custom ?? null,
        s: replacementCellStyleWithoutBorders(style ?? null),
      } satisfies ICellData
    }),
  )

  destination.setValues(copiedCells)
  return true
}
