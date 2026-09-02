import type { Editor } from '@tiptap/core'
import type { Node as PmNode } from '@tiptap/pm/model'
import type { Command, EditorState, Transaction } from '@tiptap/pm/state'
import { TableMap, isInTable, selectedRect } from '@tiptap/pm/tables'

export const MIN_TABLE_COLUMN_PX = 40

export interface SetTableColumnWidthsInput {
  tableBlockIndex: number
  widthsPx: number[]
}

export type SetTableColumnWidthsResult =
  | { ok: true; matchedColumns: number; changedColumns: number }
  | { ok: false; error: 'invalid_arguments'; message: string }

function finiteWidth(value: unknown, fallback: number): number {
  const width = Number(value)
  return Number.isFinite(width) && width > 0 ? width : fallback
}

/**
 * Fit a table grid into the section content box. Requested columns keep their
 * requested size where possible; unselected columns give up space first.
 */
export function fitColumnWidths(
  current: number[],
  requested: ReadonlyMap<number, number>,
  maxWidth: number,
  minimum = MIN_TABLE_COLUMN_PX,
): number[] {
  if (current.length === 0) return []
  const limit = Math.max(1, finiteWidth(maxWidth, 1))
  const floor = Math.min(Math.max(1, minimum), limit / current.length)
  const widths = current.map((width, index) =>
    Math.max(floor, finiteWidth(requested.get(index), finiteWidth(width, floor))),
  )

  let excess = widths.reduce((sum, width) => sum + width, 0) - limit
  const shrink = (indexes: number[]) => {
    if (excess <= 0.01 || indexes.length === 0) return
    const capacity = indexes.reduce((sum, index) => sum + Math.max(0, widths[index] - floor), 0)
    if (capacity <= 0.01) return
    const take = Math.min(excess, capacity)
    for (const index of indexes) {
      const available = Math.max(0, widths[index] - floor)
      widths[index] -= take * (available / capacity)
    }
    excess -= take
  }

  const untouched = widths.map((_, index) => index).filter((index) => !requested.has(index))
  const changed = widths.map((_, index) => index).filter((index) => requested.has(index))
  shrink(untouched)
  shrink(changed)

  // Floating point drift must never write an oversized OOXML grid.
  const total = widths.reduce((sum, width) => sum + width, 0)
  if (total > limit)
    widths[widths.length - 1] = Math.max(floor, widths[widths.length - 1] - (total - limit))
  const rounded = widths.map((width) => Math.max(1, Math.round(width * 100) / 100))
  const roundedTotal = rounded.reduce((sum, width) => sum + width, 0)
  if (roundedTotal > limit) {
    rounded[rounded.length - 1] = Math.max(
      1,
      Math.round((rounded[rounded.length - 1] - (roundedTotal - limit)) * 100) / 100,
    )
  }
  return rounded
}

interface GridTarget {
  table: PmNode
  tableStart: number
  map: TableMap
  /** requested column range [left, right); empty for whole-grid constraints */
  left: number
  right: number
}

function targetFromSelection(state: EditorState): GridTarget | null {
  if (!isInTable(state)) return null
  const rect = selectedRect(state)
  return {
    table: rect.table,
    tableStart: rect.tableStart,
    map: rect.map,
    left: rect.left,
    right: rect.right,
  }
}

/** Resize handles never set a selection; locate the table from the dragged cell's pos instead. */
function targetFromCellPos(state: EditorState, cellPos: number): GridTarget | null {
  if (cellPos < 0 || cellPos > state.doc.content.size) return null
  const $cell = state.doc.resolve(cellPos)
  if ($cell.parent.type.spec.tableRole !== 'row' || !$cell.nodeAfter) return null
  const table = $cell.node(-1)
  if (table.type.spec.tableRole !== 'table') return null
  return { table, tableStart: $cell.start(-1), map: TableMap.get(table), left: 0, right: 0 }
}

function tableGridWidths(
  state: EditorState,
  target: GridTarget,
  maxWidth: number,
): number[] | null {
  const { map, table, tableStart } = target
  const columns = map.width
  if (columns <= 0) return null
  const pct = (table.attrs.colWidthsPct as number[] | null) ?? []
  const fallback = maxWidth / columns
  const widths: number[] = []
  for (let col = 0; col < columns; col++) {
    const cellPos = map.map[col]
    const cell = state.doc.nodeAt(tableStart + cellPos)
    const cellRect = map.findCell(cellPos)
    const slot = col - cellRect.left
    const declared = (cell?.attrs.colwidth as number[] | null)?.[slot]
    widths.push(finiteWidth(declared, finiteWidth((pct[col] / 100) * maxWidth, fallback)))
  }
  return widths
}

function writeGridWidths(
  state: EditorState,
  tr: Transaction,
  target: GridTarget,
  widths: number[],
): Transaction {
  const { map, table, tableStart } = target
  const seen = new Set<number>()
  for (let row = 0; row < map.height; row++) {
    for (let col = 0; col < map.width; col++) {
      const cellPos = map.map[row * map.width + col]
      if (seen.has(cellPos)) continue
      seen.add(cellPos)
      const pos = tableStart + cellPos
      const cell = state.doc.nodeAt(pos)
      if (!cell) continue
      const cellRect = map.findCell(cellPos)
      tr = tr.setNodeMarkup(pos, undefined, {
        ...cell.attrs,
        colwidth: widths.slice(cellRect.left, cellRect.right),
      })
    }
  }
  const total = widths.reduce((sum, width) => sum + width, 0)
  tr = tr.setNodeMarkup(tableStart - 1, undefined, {
    ...table.attrs,
    widthPx: total,
    widthPct: null,
    colWidthsPct: widths.map((width) => (width / total) * 100),
  })
  return tr
}

function resizeColumns(
  state: EditorState,
  dispatch: ((tr: Transaction) => void) | undefined,
  desiredPx: number | null,
  maxWidthPx: number,
  target: GridTarget | null,
): boolean {
  if (!target) return false
  const current = tableGridWidths(state, target, maxWidthPx)
  if (!current) return false
  const requested = new Map<number, number>()
  if (desiredPx !== null) {
    for (let col = target.left; col < target.right; col++) requested.set(col, desiredPx)
  }
  const widths = fitColumnWidths(current, requested, maxWidthPx)
  const changed = widths.some((width, index) => Math.abs(width - current[index]) > 0.01)
  if (!changed && desiredPx !== null) return true
  dispatch?.(writeGridWidths(state, state.tr, target, widths))
  return true
}

/** Ribbon command: resize selected columns and redistribute the remaining grid. */
export function setSelectedColumnWidth(desiredPx: number, maxWidthPx: number): Command {
  return (state, dispatch) =>
    resizeColumns(state, dispatch, desiredPx, maxWidthPx, targetFromSelection(state))
}

/** Drag-resize cleanup: proportionally clamp an overflowing grid to the content box. */
export function constrainSelectedTableWidth(maxWidthPx: number): Command {
  return (state, dispatch) =>
    resizeColumns(state, dispatch, null, maxWidthPx, targetFromSelection(state))
}

/** Same clamp keyed off the drag handle's cell pos, for drags that start without a selection. */
export function constrainTableWidthAtCell(cellPos: number, maxWidthPx: number): Command {
  return (state, dispatch) =>
    resizeColumns(state, dispatch, null, maxWidthPx, targetFromCellPos(state, cellPos))
}

/** Write one explicit complete grid-width vector through the same kernel used by Ribbon resizing. */
export function setTableColumnWidths(
  editor: Editor,
  input: SetTableColumnWidthsInput,
): SetTableColumnWidthsResult {
  if (
    !Number.isInteger(input.tableBlockIndex) ||
    input.tableBlockIndex < 0 ||
    !Array.isArray(input.widthsPx) ||
    input.widthsPx.length < 1 ||
    input.widthsPx.length > 63 ||
    input.widthsPx.some(
      (width) => !Number.isFinite(width) || width < MIN_TABLE_COLUMN_PX || width > 4096,
    ) ||
    input.widthsPx.reduce((sum, width) => sum + width, 0) > 4096
  ) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: 'docx.table.set_column_widths requires a bounded complete width vector.',
    }
  }
  const table = editor.state.doc.maybeChild(input.tableBlockIndex)
  if (!table || table.type.name !== 'docTable') {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: `DOCX block ${input.tableBlockIndex} is not a table.`,
    }
  }
  const map = TableMap.get(table)
  if (input.widthsPx.length !== map.width) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: `DOCX table block ${input.tableBlockIndex} requires exactly ${map.width} column width(s).`,
    }
  }
  let tablePosition = 0
  editor.state.doc.forEach((_node, offset, index) => {
    if (index === input.tableBlockIndex) tablePosition = offset
  })
  const target: GridTarget = {
    table,
    tableStart: tablePosition + 1,
    map,
    left: 0,
    right: map.width,
  }
  const total = input.widthsPx.reduce((sum, width) => sum + width, 0)
  const current = tableGridWidths(editor.state, target, total)
  const hasExplicitGrid =
    (Array.isArray(table.attrs.colWidthsPct) && table.attrs.colWidthsPct.length === map.width) ||
    (Number.isFinite(Number(table.attrs.widthPx)) && Number(table.attrs.widthPx) > 0) ||
    [...new Set(map.map)].some((position) => {
      const cell = table.nodeAt(position)
      return Array.isArray(cell?.attrs.colwidth) && cell.attrs.colwidth.length > 0
    })
  const changedColumns = hasExplicitGrid
    ? input.widthsPx.filter((width, index) => Math.abs(width - (current?.[index] ?? 0)) > 0.01)
        .length
    : input.widthsPx.length
  if (changedColumns > 0) {
    editor.view.dispatch(writeGridWidths(editor.state, editor.state.tr, target, input.widthsPx))
  }
  return { ok: true, matchedColumns: map.width, changedColumns }
}
