import { formatAddress } from '../domain/cell-address'
import { recordSparklineAdd, removeSparklineAdd } from './edit-journal'
import { absRangeRef, pushWorkbookUndo } from './univer-sync'
import type { LazyWorkbookState, UniverRuntime, UniverWorksheet } from './univer-state'

export type WorkbookSparklineType = 'line' | 'column' | 'stacked'

export type WorkbookSparklineAddResult =
  | { readonly ok: true; readonly count: number }
  | {
      readonly ok: false
      readonly reason: 'source_width' | 'shape' | 'limit' | 'overlap' | 'occupied'
    }

export function applyWorkbookSparklineAdd(
  runtime: UniverRuntime,
  state: Pick<LazyWorkbookState, 'editJournal' | 'file'>,
  worksheet: UniverWorksheet,
  source: ReturnType<UniverWorksheet['getRange']>,
  target: ReturnType<UniverWorksheet['getRange']>,
  type: WorkbookSparklineType,
  refresh: (sheetId: string) => void,
): WorkbookSparklineAddResult {
  if (source.getWidth() < 2) return { ok: false, reason: 'source_width' }
  if (target.getWidth() !== 1 || target.getHeight() !== source.getHeight()) {
    return { ok: false, reason: 'shape' }
  }
  if (source.getHeight() > 200) return { ok: false, reason: 'limit' }

  const sourceStartRow = source.getRow()
  const sourceEndRow = sourceStartRow + source.getHeight() - 1
  const sourceStartColumn = source.getColumn()
  const sourceEndColumn = sourceStartColumn + source.getWidth() - 1
  const targetStartRow = target.getRow()
  const targetEndRow = targetStartRow + target.getHeight() - 1
  const targetColumn = target.getColumn()
  if (
    targetColumn >= sourceStartColumn &&
    targetColumn <= sourceEndColumn &&
    targetEndRow >= sourceStartRow &&
    targetStartRow <= sourceEndRow
  ) {
    return { ok: false, reason: 'overlap' }
  }

  const sheetId = worksheet.getSheetId()
  const sheetName = worksheet.getSheetName()
  const cells = Array.from({ length: source.getHeight() }, (_, offset) => ({
    cell: formatAddress(targetStartRow + offset, targetColumn),
    sourceRef: absRangeRef(
      sheetName,
      `${formatAddress(sourceStartRow + offset, sourceStartColumn)}` +
        `:${formatAddress(sourceStartRow + offset, sourceEndColumn)}`,
    ),
  }))
  const occupied = new Set([
    ...(state.file.sheets.find((sheet) => sheet.id === sheetId)?.sparklines ?? []).flatMap(
      (group) => group.cells.map((cell) => cell.cell),
    ),
    ...state.editJournal.sparklineAdds
      .filter((entry) => entry.sheetId === sheetId)
      .flatMap((entry) => entry.cells.map((cell) => cell.cell)),
  ])
  if (cells.some((cell) => occupied.has(cell.cell))) return { ok: false, reason: 'occupied' }

  const entry = {
    id: `sparkline-${Date.now().toString(36)}-${state.editJournal.sparklineAdds.length + 1}`,
    sheetId,
    type,
    cells,
  }
  const redraw = (): void => refresh(sheetId)
  recordSparklineAdd(state.editJournal, entry)
  pushWorkbookUndo(runtime, {
    undo: () => {
      removeSparklineAdd(state.editJournal, entry.id)
      redraw()
    },
    redo: () => {
      recordSparklineAdd(state.editJournal, entry)
      redraw()
    },
  })
  redraw()
  target.activate()
  return { ok: true, count: cells.length }
}
