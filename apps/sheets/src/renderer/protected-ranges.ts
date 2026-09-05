import { columnLabel, parseRange } from '../domain/cell-address'
import type { StructuralJournalOp } from './edit-journal'
import { fileRangeToScreenRange, fileToScreen } from './view-transform'

export interface ProtectedRangeEntry {
  readonly name: string
  readonly sqref: string
  readonly hasPassword: boolean
}

interface Area {
  startRow: number
  endRow: number
  startColumn: number
  endColumn: number
}

function formatArea(area: Area): string {
  const start = `${columnLabel(area.startColumn)}${area.startRow + 1}`
  if (area.startRow === area.endRow && area.startColumn === area.endColumn) return start
  return `${start}:${columnLabel(area.endColumn)}${area.endRow + 1}`
}

const EXACT_MOVE_ROW_CAP = 50_000

function mapArea(area: Area, operations: readonly StructuralJournalOp[]): Area[] {
  const hasMove = operations.some((operation) => operation.kind === 'move-rows')
  if (!hasMove || area.endRow - area.startRow > EXACT_MOVE_ROW_CAP) {
    const moved = fileRangeToScreenRange(operations, area)
    return moved === null ? [] : [moved]
  }
  let probeRow: number | null = null
  const rows: number[] = []
  for (let row = area.startRow; row <= area.endRow; row += 1) {
    const screen = fileToScreen(operations, 'row', row)
    if (screen === null) continue
    if (probeRow === null) probeRow = row
    rows.push(screen)
  }
  if (probeRow === null) return []
  const columns = fileRangeToScreenRange(operations, { ...area, startRow: probeRow, endRow: probeRow })
  if (columns === null) return []
  rows.sort((left, right) => left - right)
  const areas: Area[] = []
  for (const row of rows) {
    const last = areas.at(-1)
    if (last && row === last.endRow + 1) last.endRow = row
    else {
      areas.push({
        startRow: row,
        endRow: row,
        startColumn: columns.startColumn,
        endColumn: columns.endColumn,
      })
    }
  }
  return areas
}

function mapSqref(sqref: string, operations: readonly StructuralJournalOp[]): string | null {
  const parts = sqref
    .split(/\s+/)
    .filter(Boolean)
    .flatMap((part) => {
      try {
        return mapArea(parseRange(part.replaceAll('$', '')), operations).map(formatArea)
      } catch {
        return [part]
      }
    })
  return parts.length === 0 ? null : parts.join(' ')
}

export function mapProtectedRanges(
  ranges: readonly ProtectedRangeEntry[],
  operations: readonly StructuralJournalOp[],
): ProtectedRangeEntry[] {
  if (operations.length === 0) return [...ranges]
  return ranges.flatMap((range) => {
    const sqref = mapSqref(range.sqref, operations)
    return sqref === null ? [] : [{ ...range, sqref }]
  })
}
