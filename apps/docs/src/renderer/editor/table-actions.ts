import type { Editor } from '@tiptap/core'
import type { TableModel } from '@genoffice/docx-engine'
import type { Node as PmNode } from '@tiptap/pm/model'
import type { Command, EditorState, Transaction } from '@tiptap/pm/state'
import {
  CellSelection,
  TableMap,
  addColumn,
  addRow,
  isInTable,
  mergeCells,
  removeColumn,
  removeRow,
  selectedRect,
  splitCell,
} from '@tiptap/pm/tables'

import { tableModelToPmNode } from './convert'

export interface InsertTableInput {
  afterBlockIndex: number
  rows: number
  columns: number
}

export type InsertTableResult =
  | { ok: true; tableBlockIndex: number }
  | { ok: false; error: 'invalid_arguments' | 'execution_failed'; message: string }

export type DeleteTableResult =
  { ok: true } | { ok: false; error: 'invalid_arguments' | 'execution_failed'; message: string }

export interface InsertTableRowsInput {
  tableBlockIndex: number
  rowIndex: number
  count: number
}

export interface DeleteTableRowsInput {
  tableBlockIndex: number
  rowIndex: number
  count: number
}

export interface InsertTableColumnsInput {
  tableBlockIndex: number
  columnIndex: number
  count: number
}

export interface DeleteTableColumnsInput {
  tableBlockIndex: number
  columnIndex: number
  count: number
}

export interface MergeTableCellsInput {
  tableBlockIndex: number
  topRow: number
  leftColumn: number
  bottomRow: number
  rightColumn: number
}

export interface SplitTableCellInput {
  tableBlockIndex: number
  rowIndex: number
  columnIndex: number
}

export type SplitTableCellResult =
  | { ok: true; splitCells: number }
  | { ok: false; error: 'invalid_arguments' | 'execution_failed'; message: string }

export type TableCellFormatField = 'fill' | 'verticalAlignment'

export interface SetTableCellFormatInput {
  tableBlockIndex: number
  topRow: number
  leftColumn: number
  bottomRow: number
  rightColumn: number
  format: {
    fill?: string | null
    verticalAlignment?: 'top' | 'center' | 'bottom'
  }
  fields: TableCellFormatField[]
}

export type SetTableCellFormatResult =
  | { ok: true; matchedCells: number; changedCells: number }
  | { ok: false; error: 'invalid_arguments' | 'execution_failed'; message: string }

export type TableCellBorderMode = 'all' | 'outer' | 'inner' | 'none'

export interface TableCellBorder {
  color: string
  sizeEighths: number
}

export interface SetTableCellBordersInput {
  tableBlockIndex: number
  topRow: number
  leftColumn: number
  bottomRow: number
  rightColumn: number
  mode: TableCellBorderMode
  border: TableCellBorder | null
}

export type SetTableCellBordersResult =
  | { ok: true; matchedCells: number; changedCells: number }
  | { ok: false; error: 'invalid_arguments' | 'execution_failed'; message: string }

export interface SetTableStyleInput {
  tableBlockIndex: number
  styleId: string | null
}

export type SetTableStyleResult =
  { ok: true; changed: boolean } | { ok: false; error: 'invalid_arguments'; message: string }

export interface SetTableRowHeightInput {
  tableBlockIndex: number
  rowIndex: number
  count: number
  heightTwips: number | null
}

export type SetTableRowHeightResult =
  | { ok: true; matchedRows: number; changedRows: number }
  | { ok: false; error: 'invalid_arguments'; message: string }

interface CellRectangleTarget {
  map: TableMap
  tableStart: number
  top: number
  left: number
  bottom: number
  right: number
}

function applyCellFormatToPositions(
  state: EditorState,
  dispatch: ((tr: Transaction) => void) | undefined,
  positions: ReadonlySet<number>,
  format: SetTableCellFormatInput['format'],
  fields: readonly TableCellFormatField[],
): { matchedCells: number; changedCells: number } | null {
  let tr = state.tr
  let changedCells = 0
  for (const position of positions) {
    const cell = state.doc.nodeAt(position)
    if (!cell) return null
    const attrs = { ...cell.attrs }
    if (fields.includes('fill')) attrs.fill = format.fill
    if (fields.includes('verticalAlignment')) {
      attrs.vAlign = format.verticalAlignment === 'top' ? null : format.verticalAlignment
    }
    const changed =
      (fields.includes('fill') && attrs.fill !== cell.attrs.fill) ||
      (fields.includes('verticalAlignment') && attrs.vAlign !== cell.attrs.vAlign)
    if (!changed) continue
    tr = tr.setNodeMarkup(position, undefined, attrs)
    changedCells++
  }
  if (changedCells > 0) dispatch?.(tr)
  return { matchedCells: positions.size, changedCells }
}

/** Retained Ribbon adapter: resolve its current native selection, then use the shared format kernel. */
export function setSelectedTableCellFormat(
  format: SetTableCellFormatInput['format'],
  fields: readonly TableCellFormatField[],
): Command {
  return (state, dispatch) => {
    if (!isInTable(state)) return false
    const rect = selectedRect(state)
    const positions = new Set<number>()
    for (let row = rect.top; row < rect.bottom; row++) {
      for (let column = rect.left; column < rect.right; column++) {
        positions.add(rect.tableStart + rect.map.map[row * rect.map.width + column])
      }
    }
    return applyCellFormatToPositions(state, dispatch, positions, format, fields) !== null
  }
}

function applyCellBordersToRectangle(
  state: EditorState,
  dispatch: ((tr: Transaction) => void) | undefined,
  target: CellRectangleTarget,
  mode: TableCellBorderMode,
  border: TableCellBorder | null,
): { matchedCells: number; changedCells: number } | null {
  const positions = new Set<number>()
  let tr = state.tr
  let changedCells = 0
  for (let row = target.top; row < target.bottom; row++) {
    for (let column = target.left; column < target.right; column++) {
      const cellPosition = target.map.map[row * target.map.width + column]
      if (positions.has(cellPosition)) continue
      positions.add(cellPosition)
      const absolutePosition = target.tableStart + cellPosition
      const cell = state.doc.nodeAt(absolutePosition)
      if (!cell) return null
      const cellRect = target.map.findCell(cellPosition)
      const edge = {
        top: cellRect.top <= target.top,
        bottom: cellRect.bottom >= target.bottom,
        left: cellRect.left <= target.left,
        right: cellRect.right >= target.right,
      }
      const current = (cell.attrs.borders as Record<string, unknown> | null) ?? {}
      const next: Record<string, unknown> = { ...current }
      const solid = border
        ? { style: 'single', color: border.color, szEighths: border.sizeEighths }
        : null
      for (const side of ['top', 'bottom', 'left', 'right'] as const) {
        if (mode === 'all' && solid) next[side] = solid
        else if (mode === 'none') next[side] = { style: 'none' }
        else if (mode === 'outer' && edge[side] && solid) next[side] = solid
        else if (mode === 'inner' && !edge[side] && solid) next[side] = solid
      }
      if (JSON.stringify(next) === JSON.stringify(current)) continue
      tr = tr.setNodeMarkup(absolutePosition, undefined, { ...cell.attrs, borders: next })
      changedCells++
    }
  }
  if (changedCells > 0) dispatch?.(tr)
  return { matchedCells: positions.size, changedCells }
}

/** Retained Ribbon adapter for the shared cell-border geometry and write kernel. */
export function setSelectedTableCellBorders(
  mode: TableCellBorderMode,
  border: TableCellBorder | null,
): Command {
  return (state, dispatch) => {
    if (!isInTable(state)) return false
    const rect = selectedRect(state)
    return (
      applyCellBordersToRectangle(
        state,
        dispatch,
        {
          map: rect.map,
          tableStart: rect.tableStart,
          top: rect.top,
          left: rect.left,
          bottom: rect.bottom,
          right: rect.right,
        },
        mode,
        border,
      ) !== null
    )
  }
}

function applyTableStyleAtPosition(
  state: EditorState,
  dispatch: ((tr: Transaction) => void) | undefined,
  tablePosition: number,
  styleId: string | null,
): boolean | null {
  const table = state.doc.nodeAt(tablePosition)
  if (!table || table.type.name !== 'docTable') return null
  if (table.attrs.tblStyleId === styleId) return false
  dispatch?.(
    state.tr.setNodeMarkup(tablePosition, undefined, {
      ...table.attrs,
      tblStyleId: styleId,
    }),
  )
  return true
}

/** Retained Ribbon adapter for the shared native table-style attribute kernel. */
export function setSelectedTableStyle(styleId: string | null): Command {
  return (state, dispatch) => {
    if (!isInTable(state)) return false
    const rect = selectedRect(state)
    return applyTableStyleAtPosition(state, dispatch, rect.tableStart - 1, styleId) !== null
  }
}

function applyTableRowHeight(
  state: EditorState,
  dispatch: ((tr: Transaction) => void) | undefined,
  table: PmNode,
  tableStart: number,
  rowIndex: number,
  count: number,
  heightTwips: number | null,
): { matchedRows: number; changedRows: number } {
  let tr = state.tr
  let changedRows = 0
  table.forEach((row, offset, index) => {
    if (index < rowIndex || index >= rowIndex + count) return
    if (row.attrs.heightTwips === heightTwips) return
    tr = tr.setNodeMarkup(tableStart + offset, undefined, {
      ...row.attrs,
      heightTwips,
    })
    changedRows++
  })
  if (changedRows > 0) dispatch?.(tr)
  return { matchedRows: count, changedRows }
}

/** Retained Ribbon adapter for shared row-height final-state writes. */
export function setSelectedTableRowHeight(heightTwips: number | null): Command {
  return (state, dispatch) => {
    if (!isInTable(state)) return false
    const rect = selectedRect(state)
    applyTableRowHeight(
      state,
      dispatch,
      rect.table,
      rect.tableStart,
      rect.top,
      rect.bottom - rect.top,
      heightTwips,
    )
    return true
  }
}

export function buildEmptyDocxTable(rows: number, columns: number): TableModel {
  return {
    rows: Array.from({ length: rows }, () =>
      Array.from({ length: columns }, () => ({ paras: [''] })),
    ),
    colWidthsPct: Array.from({ length: columns }, () => 100 / columns),
  }
}

/** Insert a top-level table at one stable revision-scoped block boundary. */
export function insertTopLevelTableAfterBlock(
  editor: Editor,
  input: InsertTableInput,
): InsertTableResult {
  if (
    !Number.isInteger(input.afterBlockIndex) ||
    input.afterBlockIndex < -1 ||
    !Number.isInteger(input.rows) ||
    input.rows < 1 ||
    input.rows > 100 ||
    !Number.isInteger(input.columns) ||
    input.columns < 1 ||
    input.columns > 63
  ) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: 'docx.table.insert requires bounded block, row, and column values.',
    }
  }
  if (input.rows * input.columns > 4096) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: 'docx.table.insert supports at most 4096 cells.',
    }
  }
  if (input.afterBlockIndex >= editor.state.doc.childCount) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: `DOCX block ${input.afterBlockIndex} does not exist.`,
    }
  }

  let insertionPos = 0
  if (input.afterBlockIndex >= 0) {
    editor.state.doc.forEach((node, offset, index) => {
      if (index === input.afterBlockIndex) insertionPos = offset + node.nodeSize
    })
  }
  const inserted = editor
    .chain()
    .focus()
    .insertContentAt(
      insertionPos,
      tableModelToPmNode(buildEmptyDocxTable(input.rows, input.columns)),
    )
    .run()
  if (!inserted) {
    return {
      ok: false,
      error: 'execution_failed',
      message: 'The mounted DOCX editor rejected table insertion.',
    }
  }
  return { ok: true, tableBlockIndex: input.afterBlockIndex + 1 }
}

/** Delete one top-level native table by stable revision-scoped block index. */
export function deleteTopLevelTable(editor: Editor, tableBlockIndex: number): DeleteTableResult {
  if (!Number.isInteger(tableBlockIndex) || tableBlockIndex < 0) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: 'docx.table.delete requires a non-negative tableBlockIndex.',
    }
  }
  const table = editor.state.doc.maybeChild(tableBlockIndex)
  if (!table || table.type.name !== 'docTable') {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: `DOCX block ${tableBlockIndex} is not a table.`,
    }
  }

  let position = 0
  editor.state.doc.forEach((_node, offset, index) => {
    if (index === tableBlockIndex) position = offset
  })
  const tr = editor.state.tr
  if (editor.state.doc.childCount === 1) {
    const paragraph = editor.schema.nodes.docParagraph?.createAndFill()
    if (!paragraph) {
      return {
        ok: false,
        error: 'execution_failed',
        message: 'The DOCX schema could not create a replacement paragraph.',
      }
    }
    tr.replaceWith(position, position + table.nodeSize, paragraph)
  } else {
    tr.delete(position, position + table.nodeSize)
  }
  editor.view.dispatch(tr)
  return { ok: true }
}

/** Insert rows at an explicit table boundary using ProseMirror's rowspan-aware table map. */
export function insertTableRows(editor: Editor, input: InsertTableRowsInput): DeleteTableResult {
  if (
    !Number.isInteger(input.tableBlockIndex) ||
    input.tableBlockIndex < 0 ||
    !Number.isInteger(input.rowIndex) ||
    input.rowIndex < 0 ||
    !Number.isInteger(input.count) ||
    input.count < 1 ||
    input.count > 100
  ) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: 'docx.table.insert_rows requires bounded table, boundary, and count values.',
    }
  }
  const sourceTable = editor.state.doc.maybeChild(input.tableBlockIndex)
  if (!sourceTable || sourceTable.type.name !== 'docTable') {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: `DOCX block ${input.tableBlockIndex} is not a table.`,
    }
  }
  if (input.rowIndex > sourceTable.childCount) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: `DOCX table block ${input.tableBlockIndex} has ${sourceTable.childCount} row(s); boundary ${input.rowIndex} is invalid.`,
    }
  }

  const sourceMap = TableMap.get(sourceTable)
  if (sourceMap.width * (sourceMap.height + input.count) > 4096) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: 'docx.table.insert_rows supports at most 4096 resulting cells.',
    }
  }
  let tablePosition = 0
  editor.state.doc.forEach((_node, offset, index) => {
    if (index === input.tableBlockIndex) tablePosition = offset
  })
  let tr = editor.state.tr
  for (let index = 0; index < input.count; index++) {
    const table = tr.doc.nodeAt(tablePosition)
    if (!table || table.type.name !== 'docTable') {
      return {
        ok: false,
        error: 'execution_failed',
        message: 'The DOCX table disappeared during row insertion.',
      }
    }
    const map = TableMap.get(table)
    tr = addRow(
      tr,
      {
        map,
        tableStart: tablePosition + 1,
        table,
        left: 0,
        top: 0,
        right: map.width,
        bottom: map.height,
      },
      input.rowIndex,
    )
  }
  editor.view.dispatch(tr)
  return { ok: true }
}

/** Delete rows from an explicit index using ProseMirror's rowspan-aware table map. */
export function deleteTableRows(editor: Editor, input: DeleteTableRowsInput): DeleteTableResult {
  if (
    !Number.isInteger(input.tableBlockIndex) ||
    input.tableBlockIndex < 0 ||
    !Number.isInteger(input.rowIndex) ||
    input.rowIndex < 0 ||
    !Number.isInteger(input.count) ||
    input.count < 1 ||
    input.count > 100
  ) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: 'docx.table.delete_rows requires bounded table, row, and count values.',
    }
  }
  const sourceTable = editor.state.doc.maybeChild(input.tableBlockIndex)
  if (!sourceTable || sourceTable.type.name !== 'docTable') {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: `DOCX block ${input.tableBlockIndex} is not a table.`,
    }
  }
  if (input.rowIndex + input.count > sourceTable.childCount) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: `DOCX table block ${input.tableBlockIndex} row range [${input.rowIndex}, ${input.rowIndex + input.count}) is invalid.`,
    }
  }
  if (input.count === sourceTable.childCount) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: `DOCX table block ${input.tableBlockIndex} cannot delete every row; use docx.table.delete instead.`,
    }
  }

  let tablePosition = 0
  editor.state.doc.forEach((_node, offset, index) => {
    if (index === input.tableBlockIndex) tablePosition = offset
  })
  const tr = editor.state.tr
  for (let index = 0; index < input.count; index++) {
    const table = tr.doc.nodeAt(tablePosition)
    if (!table || table.type.name !== 'docTable') {
      return {
        ok: false,
        error: 'execution_failed',
        message: 'The DOCX table disappeared during row deletion.',
      }
    }
    const map = TableMap.get(table)
    removeRow(
      tr,
      {
        map,
        tableStart: tablePosition + 1,
        table,
        left: 0,
        top: 0,
        right: map.width,
        bottom: map.height,
      },
      input.rowIndex,
    )
  }
  editor.view.dispatch(tr)
  return { ok: true }
}

/** Insert columns at an explicit table boundary using ProseMirror's span-aware table map. */
export function insertTableColumns(
  editor: Editor,
  input: InsertTableColumnsInput,
): DeleteTableResult {
  if (
    !Number.isInteger(input.tableBlockIndex) ||
    input.tableBlockIndex < 0 ||
    !Number.isInteger(input.columnIndex) ||
    input.columnIndex < 0 ||
    !Number.isInteger(input.count) ||
    input.count < 1 ||
    input.count > 63
  ) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: 'docx.table.insert_columns requires bounded table, boundary, and count values.',
    }
  }
  const sourceTable = editor.state.doc.maybeChild(input.tableBlockIndex)
  if (!sourceTable || sourceTable.type.name !== 'docTable') {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: `DOCX block ${input.tableBlockIndex} is not a table.`,
    }
  }
  const sourceMap = TableMap.get(sourceTable)
  if (input.columnIndex > sourceMap.width) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: `DOCX table block ${input.tableBlockIndex} has ${sourceMap.width} column(s); boundary ${input.columnIndex} is invalid.`,
    }
  }
  if (sourceMap.width + input.count > 63) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: 'docx.table.insert_columns supports at most 63 resulting columns.',
    }
  }
  if ((sourceMap.width + input.count) * sourceMap.height > 4096) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: 'docx.table.insert_columns supports at most 4096 resulting cells.',
    }
  }

  let tablePosition = 0
  editor.state.doc.forEach((_node, offset, index) => {
    if (index === input.tableBlockIndex) tablePosition = offset
  })
  let tr = editor.state.tr
  for (let index = 0; index < input.count; index++) {
    const table = tr.doc.nodeAt(tablePosition)
    if (!table || table.type.name !== 'docTable') {
      return {
        ok: false,
        error: 'execution_failed',
        message: 'The DOCX table disappeared during column insertion.',
      }
    }
    const map = TableMap.get(table)
    tr = addColumn(
      tr,
      {
        map,
        tableStart: tablePosition + 1,
        table,
        left: 0,
        top: 0,
        right: map.width,
        bottom: map.height,
      },
      input.columnIndex,
    )
  }
  editor.view.dispatch(tr)
  return { ok: true }
}

/** Delete columns from an explicit index using ProseMirror's span-aware table map. */
export function deleteTableColumns(
  editor: Editor,
  input: DeleteTableColumnsInput,
): DeleteTableResult {
  if (
    !Number.isInteger(input.tableBlockIndex) ||
    input.tableBlockIndex < 0 ||
    !Number.isInteger(input.columnIndex) ||
    input.columnIndex < 0 ||
    !Number.isInteger(input.count) ||
    input.count < 1 ||
    input.count > 63
  ) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: 'docx.table.delete_columns requires bounded table, column, and count values.',
    }
  }
  const sourceTable = editor.state.doc.maybeChild(input.tableBlockIndex)
  if (!sourceTable || sourceTable.type.name !== 'docTable') {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: `DOCX block ${input.tableBlockIndex} is not a table.`,
    }
  }
  const sourceMap = TableMap.get(sourceTable)
  if (input.columnIndex + input.count > sourceMap.width) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: `DOCX table block ${input.tableBlockIndex} column range [${input.columnIndex}, ${input.columnIndex + input.count}) is invalid.`,
    }
  }
  if (input.count === sourceMap.width) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: `DOCX table block ${input.tableBlockIndex} cannot delete every column; use docx.table.delete instead.`,
    }
  }

  let tablePosition = 0
  editor.state.doc.forEach((_node, offset, index) => {
    if (index === input.tableBlockIndex) tablePosition = offset
  })
  const tr = editor.state.tr
  for (let index = 0; index < input.count; index++) {
    const table = tr.doc.nodeAt(tablePosition)
    if (!table || table.type.name !== 'docTable') {
      return {
        ok: false,
        error: 'execution_failed',
        message: 'The DOCX table disappeared during column deletion.',
      }
    }
    const map = TableMap.get(table)
    removeColumn(
      tr,
      {
        map,
        tableStart: tablePosition + 1,
        table,
        left: 0,
        top: 0,
        right: map.width,
        bottom: map.height,
      },
      input.columnIndex,
    )
  }
  editor.view.dispatch(tr)
  return { ok: true }
}

/** Merge one explicit half-open logical-cell rectangle through the native table command. */
export function mergeTableCells(editor: Editor, input: MergeTableCellsInput): DeleteTableResult {
  if (
    !Number.isInteger(input.tableBlockIndex) ||
    input.tableBlockIndex < 0 ||
    !Number.isInteger(input.topRow) ||
    input.topRow < 0 ||
    input.topRow > 99 ||
    !Number.isInteger(input.leftColumn) ||
    input.leftColumn < 0 ||
    input.leftColumn > 62 ||
    !Number.isInteger(input.bottomRow) ||
    input.bottomRow < 1 ||
    input.bottomRow > 100 ||
    !Number.isInteger(input.rightColumn) ||
    input.rightColumn < 1 ||
    input.rightColumn > 63 ||
    input.topRow >= input.bottomRow ||
    input.leftColumn >= input.rightColumn
  ) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: 'docx.table.merge_cells requires one bounded non-empty half-open rectangle.',
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
  if (input.bottomRow > map.height || input.rightColumn > map.width) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: `DOCX table block ${input.tableBlockIndex} rectangle [${input.topRow}, ${input.bottomRow}) × [${input.leftColumn}, ${input.rightColumn}) is outside its ${map.height}×${map.width} logical grid.`,
    }
  }
  const logicalCellCount = (input.bottomRow - input.topRow) * (input.rightColumn - input.leftColumn)
  if (logicalCellCount < 2) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: `DOCX table block ${input.tableBlockIndex} must cover at least two logical cells.`,
    }
  }

  const anchorCell = map.map[input.topRow * map.width + input.leftColumn]
  const headCell = map.map[(input.bottomRow - 1) * map.width + input.rightColumn - 1]
  const exactRect = map.rectBetween(anchorCell, headCell)
  if (
    exactRect.top !== input.topRow ||
    exactRect.left !== input.leftColumn ||
    exactRect.bottom !== input.bottomRow ||
    exactRect.right !== input.rightColumn
  ) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: `DOCX table block ${input.tableBlockIndex} rectangle crosses an existing merged-cell boundary.`,
    }
  }

  let tablePosition = 0
  editor.state.doc.forEach((_node, offset, index) => {
    if (index === input.tableBlockIndex) tablePosition = offset
  })
  const tableStart = tablePosition + 1
  const selection = CellSelection.create(
    editor.state.doc,
    tableStart + anchorCell,
    tableStart + headCell,
  )
  const selectionState = editor.state.apply(
    editor.state.tr.setSelection(selection).setMeta('addToHistory', false),
  )
  let dispatched = false
  const merged = mergeCells(selectionState, (tr) => {
    editor.view.dispatch(tr)
    dispatched = true
  })
  if (!merged || !dispatched) {
    return {
      ok: false,
      error: 'execution_failed',
      message: 'The mounted DOCX editor rejected cell merging.',
    }
  }
  return { ok: true }
}

/** Split the merged cell covering one explicit logical coordinate through the native command. */
export function splitTableCell(editor: Editor, input: SplitTableCellInput): SplitTableCellResult {
  if (
    !Number.isInteger(input.tableBlockIndex) ||
    input.tableBlockIndex < 0 ||
    !Number.isInteger(input.rowIndex) ||
    input.rowIndex < 0 ||
    input.rowIndex > 99 ||
    !Number.isInteger(input.columnIndex) ||
    input.columnIndex < 0 ||
    input.columnIndex > 62
  ) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: 'docx.table.split_cell requires bounded table, row, and column values.',
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
  if (input.rowIndex >= map.height || input.columnIndex >= map.width) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: `DOCX table block ${input.tableBlockIndex} logical cell (${input.rowIndex}, ${input.columnIndex}) is outside its ${map.height}×${map.width} grid.`,
    }
  }
  const cellPosition = map.map[input.rowIndex * map.width + input.columnIndex]
  const cell = table.nodeAt(cellPosition)
  if (!cell) {
    return {
      ok: false,
      error: 'execution_failed',
      message: 'The DOCX table cell could not be resolved.',
    }
  }
  const splitCells = Number(cell.attrs.rowspan) * Number(cell.attrs.colspan)
  if (splitCells <= 1) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: `DOCX table block ${input.tableBlockIndex} logical cell (${input.rowIndex}, ${input.columnIndex}) is not merged.`,
    }
  }

  let tablePosition = 0
  editor.state.doc.forEach((_node, offset, index) => {
    if (index === input.tableBlockIndex) tablePosition = offset
  })
  const absoluteCellPosition = tablePosition + 1 + cellPosition
  const selection = CellSelection.create(editor.state.doc, absoluteCellPosition)
  const selectionState = editor.state.apply(
    editor.state.tr.setSelection(selection).setMeta('addToHistory', false),
  )
  let dispatched = false
  const split = splitCell(selectionState, (tr) => {
    editor.view.dispatch(tr)
    dispatched = true
  })
  if (!split || !dispatched) {
    return {
      ok: false,
      error: 'execution_failed',
      message: 'The mounted DOCX editor rejected cell splitting.',
    }
  }
  return { ok: true, splitCells }
}

/** Apply an exact masked final state to every physical cell in one logical rectangle. */
export function setTableCellFormat(
  editor: Editor,
  input: SetTableCellFormatInput,
): SetTableCellFormatResult {
  if (
    !Number.isInteger(input.tableBlockIndex) ||
    input.tableBlockIndex < 0 ||
    !Number.isInteger(input.topRow) ||
    input.topRow < 0 ||
    input.topRow > 99 ||
    !Number.isInteger(input.leftColumn) ||
    input.leftColumn < 0 ||
    input.leftColumn > 62 ||
    !Number.isInteger(input.bottomRow) ||
    input.bottomRow < 1 ||
    input.bottomRow > 100 ||
    !Number.isInteger(input.rightColumn) ||
    input.rightColumn < 1 ||
    input.rightColumn > 63 ||
    input.topRow >= input.bottomRow ||
    input.leftColumn >= input.rightColumn
  ) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: 'docx.table.set_cell_format requires one bounded non-empty half-open rectangle.',
    }
  }
  if (
    !Array.isArray(input.fields) ||
    input.fields.length < 1 ||
    input.fields.length > 2 ||
    new Set(input.fields).size !== input.fields.length ||
    input.fields.some((field) => field !== 'fill' && field !== 'verticalAlignment')
  ) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: 'docx.table.set_cell_format fields must be a unique non-empty field mask.',
    }
  }
  for (const field of input.fields) {
    if (!Object.prototype.hasOwnProperty.call(input.format, field)) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: `docx.table.set_cell_format format.${field} is required by fields.`,
      }
    }
  }
  if (
    input.fields.includes('fill') &&
    input.format.fill !== null &&
    (typeof input.format.fill !== 'string' || !/^[0-9A-F]{6}$/.test(input.format.fill))
  ) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: 'docx.table.set_cell_format fill must be null or six uppercase hexadecimal digits.',
    }
  }
  if (
    input.fields.includes('verticalAlignment') &&
    !['top', 'center', 'bottom'].includes(input.format.verticalAlignment ?? '')
  ) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: 'docx.table.set_cell_format verticalAlignment is invalid.',
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
  if (input.bottomRow > map.height || input.rightColumn > map.width) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: `DOCX table block ${input.tableBlockIndex} rectangle [${input.topRow}, ${input.bottomRow}) × [${input.leftColumn}, ${input.rightColumn}) is outside its ${map.height}×${map.width} logical grid.`,
    }
  }
  const anchorCell = map.map[input.topRow * map.width + input.leftColumn]
  const headCell = map.map[(input.bottomRow - 1) * map.width + input.rightColumn - 1]
  const exactRect = map.rectBetween(anchorCell, headCell)
  if (
    exactRect.top !== input.topRow ||
    exactRect.left !== input.leftColumn ||
    exactRect.bottom !== input.bottomRow ||
    exactRect.right !== input.rightColumn
  ) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: `DOCX table block ${input.tableBlockIndex} rectangle crosses an existing merged-cell boundary.`,
    }
  }

  let tablePosition = 0
  editor.state.doc.forEach((_node, offset, index) => {
    if (index === input.tableBlockIndex) tablePosition = offset
  })
  const positions = new Set<number>()
  for (let row = input.topRow; row < input.bottomRow; row++) {
    for (let column = input.leftColumn; column < input.rightColumn; column++) {
      positions.add(tablePosition + 1 + map.map[row * map.width + column])
    }
  }
  const applied = applyCellFormatToPositions(
    editor.state,
    editor.view.dispatch,
    positions,
    input.format,
    input.fields,
  )
  if (!applied) {
    return {
      ok: false,
      error: 'execution_failed',
      message: 'A DOCX table cell disappeared during formatting.',
    }
  }
  return { ok: true, ...applied }
}

/** Apply one explicit border policy to an exact half-open logical-cell rectangle. */
export function setTableCellBorders(
  editor: Editor,
  input: SetTableCellBordersInput,
): SetTableCellBordersResult {
  if (
    !Number.isInteger(input.tableBlockIndex) ||
    input.tableBlockIndex < 0 ||
    !Number.isInteger(input.topRow) ||
    input.topRow < 0 ||
    input.topRow > 99 ||
    !Number.isInteger(input.leftColumn) ||
    input.leftColumn < 0 ||
    input.leftColumn > 62 ||
    !Number.isInteger(input.bottomRow) ||
    input.bottomRow < 1 ||
    input.bottomRow > 100 ||
    !Number.isInteger(input.rightColumn) ||
    input.rightColumn < 1 ||
    input.rightColumn > 63 ||
    input.topRow >= input.bottomRow ||
    input.leftColumn >= input.rightColumn
  ) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: 'docx.table.set_cell_borders requires one bounded non-empty half-open rectangle.',
    }
  }
  if (!['all', 'outer', 'inner', 'none'].includes(input.mode)) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: 'docx.table.set_cell_borders mode is invalid.',
    }
  }
  if (input.mode === 'none' && input.border !== null) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: 'docx.table.set_cell_borders mode none requires border: null.',
    }
  }
  if (
    input.mode !== 'none' &&
    (!input.border ||
      !/^[0-9A-F]{6}$/.test(input.border.color) ||
      !Number.isInteger(input.border.sizeEighths) ||
      input.border.sizeEighths < 1 ||
      input.border.sizeEighths > 96)
  ) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: 'docx.table.set_cell_borders requires a bounded uppercase hexadecimal border.',
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
  if (input.bottomRow > map.height || input.rightColumn > map.width) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: `DOCX table block ${input.tableBlockIndex} rectangle [${input.topRow}, ${input.bottomRow}) × [${input.leftColumn}, ${input.rightColumn}) is outside its ${map.height}×${map.width} logical grid.`,
    }
  }
  const anchorCell = map.map[input.topRow * map.width + input.leftColumn]
  const headCell = map.map[(input.bottomRow - 1) * map.width + input.rightColumn - 1]
  const exactRect = map.rectBetween(anchorCell, headCell)
  if (
    exactRect.top !== input.topRow ||
    exactRect.left !== input.leftColumn ||
    exactRect.bottom !== input.bottomRow ||
    exactRect.right !== input.rightColumn
  ) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: `DOCX table block ${input.tableBlockIndex} rectangle crosses an existing merged-cell boundary.`,
    }
  }

  let tablePosition = 0
  editor.state.doc.forEach((_node, offset, index) => {
    if (index === input.tableBlockIndex) tablePosition = offset
  })
  const applied = applyCellBordersToRectangle(
    editor.state,
    editor.view.dispatch,
    {
      map,
      tableStart: tablePosition + 1,
      top: input.topRow,
      left: input.leftColumn,
      bottom: input.bottomRow,
      right: input.rightColumn,
    },
    input.mode,
    input.border,
  )
  if (!applied) {
    return {
      ok: false,
      error: 'execution_failed',
      message: 'A DOCX table cell disappeared during border formatting.',
    }
  }
  return { ok: true, ...applied }
}

/** Set or clear one explicit top-level table's current-document style identity. */
export function setTableStyle(editor: Editor, input: SetTableStyleInput): SetTableStyleResult {
  if (!Number.isInteger(input.tableBlockIndex) || input.tableBlockIndex < 0) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: 'docx.table.set_style requires a non-negative tableBlockIndex.',
    }
  }
  let tablePosition = 0
  let tableFound = false
  editor.state.doc.forEach((node, offset, index) => {
    if (index !== input.tableBlockIndex) return
    tableFound = node.type.name === 'docTable'
    tablePosition = offset
  })
  if (!tableFound) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: `DOCX block ${input.tableBlockIndex} is not a table.`,
    }
  }
  const changed = applyTableStyleAtPosition(
    editor.state,
    editor.view.dispatch,
    tablePosition,
    input.styleId,
  )
  if (changed === null) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: `DOCX block ${input.tableBlockIndex} is not a table.`,
    }
  }
  return { ok: true, changed }
}

/** Set or clear row height over one explicit physical-row interval. */
export function setTableRowHeight(
  editor: Editor,
  input: SetTableRowHeightInput,
): SetTableRowHeightResult {
  if (
    !Number.isInteger(input.tableBlockIndex) ||
    input.tableBlockIndex < 0 ||
    !Number.isInteger(input.rowIndex) ||
    input.rowIndex < 0 ||
    input.rowIndex > 99 ||
    !Number.isInteger(input.count) ||
    input.count < 1 ||
    input.count > 100 ||
    (input.heightTwips !== null &&
      (!Number.isInteger(input.heightTwips) || input.heightTwips < 1 || input.heightTwips > 31680))
  ) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: 'docx.table.set_row_height requires bounded table, row, count, and height values.',
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
  if (input.rowIndex + input.count > table.childCount) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: `DOCX table block ${input.tableBlockIndex} row range [${input.rowIndex}, ${input.rowIndex + input.count}) is invalid.`,
    }
  }
  let tablePosition = 0
  editor.state.doc.forEach((_node, offset, index) => {
    if (index === input.tableBlockIndex) tablePosition = offset
  })
  const result = applyTableRowHeight(
    editor.state,
    editor.view.dispatch,
    table,
    tablePosition + 1,
    input.rowIndex,
    input.count,
    input.heightTwips,
  )
  return { ok: true, ...result }
}
