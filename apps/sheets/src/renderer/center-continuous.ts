import { Font, SpreadsheetSkeleton } from '@univerjs/engine-render'

export const CENTER_ACROSS_END_KEY = 'centerAcrossEnd'

interface WorksheetLike {
  getCell?(row: number, column: number): { custom?: Record<string, unknown> | null } | null
}

function centerAcrossEnd(
  worksheet: WorksheetLike | null | undefined,
  row: number,
  column: number,
): number | undefined {
  const value = worksheet?.getCell?.(row, column)?.custom?.[CENTER_ACROSS_END_KEY]
  return typeof value === 'number' && Number.isInteger(value) && value > column ? value : undefined
}

let installed = false

export function installCenterContinuousRender(): void {
  if (installed) return
  installed = true
  const skeletonPrototype = SpreadsheetSkeleton.prototype as unknown as Record<string, Function>
  const originalOverflow = skeletonPrototype.getOverflowPosition
  skeletonPrototype.getOverflowPosition = function (
    this: { worksheet?: WorksheetLike },
    contentSize: unknown,
    horizontalAlign: unknown,
    row: number,
    column: number,
    columnCount: number,
  ) {
    const end = centerAcrossEnd(this.worksheet, row, column)
    if (end !== undefined) return { startColumn: column, endColumn: Math.min(end, columnCount - 1) }
    return originalOverflow?.call(this, contentSize, horizontalAlign, row, column, columnCount)
  }

  const fontPrototype = Font.prototype as unknown as Record<string, Function>
  for (const method of ['_renderText', '_renderDocuments']) {
    const original = fontPrototype[method]
    fontPrototype[method] = function (
      this: unknown,
      context: unknown,
      row: number,
      column: number,
      renderContext: {
        startX: number
        endX: number
        spreadsheetSkeleton?: {
          worksheet?: WorksheetLike
          getColumnCount?(): number
          getCellWithCoordByIndex?(row: number, column: number, header: boolean): { endX: number }
        }
      },
      overflowCache: unknown,
    ) {
      const skeleton = renderContext.spreadsheetSkeleton
      const end = skeleton ? centerAcrossEnd(skeleton.worksheet, row, column) : undefined
      if (end === undefined || !skeleton) {
        return original?.call(this, context, row, column, renderContext, overflowCache)
      }
      const savedEndX = renderContext.endX
      const lastColumn = Math.min(end, (skeleton.getColumnCount?.() ?? end + 1) - 1)
      const coordinate = skeleton.getCellWithCoordByIndex?.(row, lastColumn, false)
      if (coordinate && coordinate.endX > renderContext.startX) renderContext.endX = coordinate.endX
      try {
        return original?.call(this, context, row, column, renderContext, overflowCache)
      } finally {
        renderContext.endX = savedEndX
      }
    }
  }
}
