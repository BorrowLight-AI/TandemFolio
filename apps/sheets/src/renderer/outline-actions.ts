import { ICommandService, type IRange } from '@univerjs/core'
import {
  SetColHiddenMutation,
  SetColVisibleMutation,
  SetRowHiddenMutation,
  SetRowVisibleMutation,
} from '@univerjs/sheets'

import {
  journalSize,
  recordStructuralOp,
  type EditJournal,
  type StructuralJournalOp,
} from './edit-journal'
import { pushWorkbookUndo, sheetOutline } from './univer-sync'
import {
  journalSuppression,
  type LazyWorkbookState,
  type UniverRuntime,
  type UniverWorksheet,
} from './univer-state'

export type WorkbookOutlineAxis = 'rows' | 'columns'

type OutlineState = Pick<LazyWorkbookState, 'editJournal' | 'outline'>
type OutlineEntry = { level: number; collapsed: boolean; hidden?: boolean }
type OutlineSegment = { readonly start: number; readonly count: number; readonly level: number }

function outlineEntries(
  state: OutlineState,
  sheetId: string,
  axis: WorkbookOutlineAxis,
): Map<number, OutlineEntry> {
  const outline = sheetOutline(state as LazyWorkbookState, sheetId)
  return axis === 'rows' ? outline.rows : outline.cols
}

function snapshotJournal(journal: EditJournal, sheetId: string): StructuralJournalOp[] | null {
  const operations = journal.structuralOps.get(sheetId)
  return operations ? [...operations] : null
}

function restoreJournal(
  journal: EditJournal,
  sheetId: string,
  snapshot: readonly StructuralJournalOp[] | null,
): void {
  if (snapshot && snapshot.length > 0) journal.structuralOps.set(sheetId, [...snapshot])
  else journal.structuralOps.delete(sheetId)
}

function snapshotEntries(
  entries: ReadonlyMap<number, OutlineEntry>,
  indices: readonly number[],
): Map<number, OutlineEntry | null> {
  return new Map(
    indices.map((index) => {
      const entry = entries.get(index)
      return [index, entry ? { ...entry } : null]
    }),
  )
}

function restoreEntries(
  entries: Map<number, OutlineEntry>,
  snapshot: ReadonlyMap<number, OutlineEntry | null>,
): void {
  for (const [index, entry] of snapshot) {
    if (entry) entries.set(index, { ...entry })
    else entries.delete(index)
  }
}

function notifyPending(state: OutlineState, setPendingEdits?: (count: number) => void): void {
  setPendingEdits?.(journalSize(state.editJournal))
}

function visibilityRange(
  worksheet: UniverWorksheet,
  axis: WorkbookOutlineAxis,
  start: number,
  count: number,
): IRange {
  return axis === 'rows'
    ? {
        startRow: start,
        endRow: start + count - 1,
        startColumn: 0,
        endColumn: worksheet.getMaxColumns() - 1,
      }
    : {
        startRow: 0,
        endRow: worksheet.getMaxRows() - 1,
        startColumn: start,
        endColumn: start + count - 1,
      }
}

/**
 * Applies hidden state through Univer's mutation layer without asking the
 * built-in command to create a second undo item. The paired renderer-owned
 * undo step owns both live visibility and the OOXML journal atomically.
 */
function mutateOutlineVisibility(
  runtime: UniverRuntime,
  worksheet: UniverWorksheet,
  axis: WorkbookOutlineAxis,
  start: number,
  count: number,
  hidden: boolean,
): void {
  const workbook = runtime.univerAPI.getActiveWorkbook()
  const unitId = workbook?.getId()
  const injector = (
    runtime.univer as unknown as {
      __getInjector?: () => { get<T>(token: unknown): T }
    }
  )?.__getInjector?.()
  if (!unitId || !injector) {
    if (axis === 'rows') {
      if (hidden) worksheet.hideRows(start, count)
      else worksheet.showRows(start, count)
    } else if (hidden) worksheet.hideColumns(start, count)
    else worksheet.showColumns(start, count)
    return
  }

  const mutation =
    axis === 'rows'
      ? hidden
        ? SetRowHiddenMutation
        : SetRowVisibleMutation
      : hidden
        ? SetColHiddenMutation
        : SetColVisibleMutation
  const previousSuppression = journalSuppression.active
  journalSuppression.active = true
  try {
    const ok = injector
      .get<{ syncExecuteCommand(id: string, params: unknown): boolean }>(ICommandService)
      .syncExecuteCommand(mutation.id, {
        unitId,
        subUnitId: worksheet.getSheetId(),
        ranges: [visibilityRange(worksheet, axis, start, count)],
      })
    if (!ok) throw new Error('Univer rejected the outline visibility mutation.')
  } finally {
    journalSuppression.active = previousSuppression
  }
}

function applyVisibilitySnapshot(
  runtime: UniverRuntime,
  worksheet: UniverWorksheet,
  axis: WorkbookOutlineAxis,
  snapshot: ReadonlyMap<number, OutlineEntry | null>,
): void {
  const ordered = [...snapshot]
    .filter(([index]) => index < Math.max(...snapshot.keys()))
    .sort(([left], [right]) => left - right)
  let runStart: number | null = null
  let runHidden = false
  let previous = -2
  const close = (end: number): void => {
    if (runStart === null) return
    mutateOutlineVisibility(runtime, worksheet, axis, runStart, end - runStart + 1, runHidden)
  }
  for (const [index, entry] of ordered) {
    const hidden = entry?.hidden ?? false
    if (runStart === null || index !== previous + 1 || hidden !== runHidden) {
      close(previous)
      runStart = index
      runHidden = hidden
    }
    previous = index
  }
  close(previous)
}

/** Applies one or more absolute outline-level spans as a single undo unit. */
export function applyWorkbookOutlineLevels(
  runtime: UniverRuntime,
  state: OutlineState,
  worksheet: UniverWorksheet,
  axis: WorkbookOutlineAxis,
  segments: readonly OutlineSegment[],
  setPendingEdits?: (count: number) => void,
): boolean {
  const sheetId = worksheet.getSheetId()
  const entries = outlineEntries(state, sheetId, axis)
  const indices = [
    ...new Set(
      segments.flatMap((segment) =>
        Array.from({ length: segment.count }, (_, offset) => segment.start + offset),
      ),
    ),
  ]
  if (
    indices.length === 0 ||
    segments.every((segment) =>
      Array.from({ length: segment.count }, (_, offset) => segment.start + offset).every(
        (index) => (entries.get(index)?.level ?? 0) === segment.level,
      ),
    )
  ) {
    return false
  }

  const beforeEntries = snapshotEntries(entries, indices)
  const beforeJournal = snapshotJournal(state.editJournal, sheetId)
  const apply = (): void => {
    for (const segment of segments) {
      for (let index = segment.start; index < segment.start + segment.count; index += 1) {
        const previous = entries.get(index)
        entries.set(index, {
          level: segment.level,
          collapsed: previous?.collapsed ?? false,
          hidden: previous?.hidden ?? false,
        })
      }
      recordStructuralOp(state.editJournal, sheetId, {
        kind: axis === 'rows' ? 'set-rows-outline' : 'set-cols-outline',
        start: segment.start,
        end: segment.start + segment.count - 1,
        level: segment.level,
      })
    }
    notifyPending(state, setPendingEdits)
  }
  apply()
  const afterEntries = snapshotEntries(entries, indices)
  const afterJournal = snapshotJournal(state.editJournal, sheetId)
  pushWorkbookUndo(runtime, {
    undo: () => {
      restoreEntries(entries, beforeEntries)
      restoreJournal(state.editJournal, sheetId, beforeJournal)
      notifyPending(state, setPendingEdits)
    },
    redo: () => {
      restoreEntries(entries, afterEntries)
      restoreJournal(state.editJournal, sheetId, afterJournal)
      notifyPending(state, setPendingEdits)
    },
  })
  return true
}

/** Applies one absolute outline level to a bounded row/column span. */
export function applyWorkbookOutlineLevel(
  runtime: UniverRuntime,
  state: OutlineState,
  worksheet: UniverWorksheet,
  axis: WorkbookOutlineAxis,
  start: number,
  count: number,
  level: number,
  setPendingEdits?: (count: number) => void,
): boolean {
  return applyWorkbookOutlineLevels(
    runtime,
    state,
    worksheet,
    axis,
    [{ start, count, level }],
    setPendingEdits,
  )
}

/**
 * Sets detail rows/columns to one final hidden state and sets the immediately
 * following summary item to the matching collapsed state.
 */
export function applyWorkbookOutlineDetailVisibility(
  runtime: UniverRuntime,
  state: OutlineState,
  worksheet: UniverWorksheet,
  axis: WorkbookOutlineAxis,
  start: number,
  count: number,
  hidden: boolean,
  setPendingEdits?: (count: number) => void,
): boolean {
  const sheetId = worksheet.getSheetId()
  const entries = outlineEntries(state, sheetId, axis)
  const summary = start + count
  const indices = Array.from({ length: count + 1 }, (_, offset) => start + offset)
  if (
    indices.slice(0, -1).every((index) => (entries.get(index)?.hidden ?? false) === hidden) &&
    (entries.get(summary)?.collapsed ?? false) === hidden
  ) {
    return false
  }

  const beforeEntries = snapshotEntries(entries, indices)
  const beforeJournal = snapshotJournal(state.editJournal, sheetId)
  const apply = (): void => {
    mutateOutlineVisibility(runtime, worksheet, axis, start, count, hidden)
    for (let index = start; index < summary; index += 1) {
      const previous = entries.get(index)
      entries.set(index, {
        level: previous?.level ?? 0,
        collapsed: previous?.collapsed ?? false,
        hidden,
      })
    }
    const summaryEntry = entries.get(summary)
    entries.set(summary, {
      level: summaryEntry?.level ?? 0,
      collapsed: hidden,
      hidden: summaryEntry?.hidden ?? false,
    })
    recordStructuralOp(state.editJournal, sheetId, {
      kind: axis === 'rows' ? 'set-rows-hidden' : 'set-cols-hidden',
      start,
      end: summary - 1,
      hidden,
    })
    recordStructuralOp(state.editJournal, sheetId, {
      kind: axis === 'rows' ? 'set-rows-outline' : 'set-cols-outline',
      start: summary,
      end: summary,
      level: summaryEntry?.level ?? 0,
      collapsed: hidden,
    })
    notifyPending(state, setPendingEdits)
  }
  apply()
  const afterEntries = snapshotEntries(entries, indices)
  const afterJournal = snapshotJournal(state.editJournal, sheetId)
  pushWorkbookUndo(runtime, {
    undo: () => {
      applyVisibilitySnapshot(runtime, worksheet, axis, beforeEntries)
      restoreEntries(entries, beforeEntries)
      restoreJournal(state.editJournal, sheetId, beforeJournal)
      notifyPending(state, setPendingEdits)
    },
    redo: () => {
      mutateOutlineVisibility(runtime, worksheet, axis, start, count, hidden)
      restoreEntries(entries, afterEntries)
      restoreJournal(state.editJournal, sheetId, afterJournal)
      notifyPending(state, setPendingEdits)
    },
  })
  return true
}
