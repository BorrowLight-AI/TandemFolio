import { parseRange } from '../domain/cell-address'
import { isSheetRemoved, journalSize, recordProtectedRangesChange } from './edit-journal'
import { pushWorkbookUndo } from './univer-sync'
import type { LazyWorkbookState, UniverRuntime } from './univer-state'

export interface EditableProtectedRange {
  readonly name: string
  readonly sqref: string
  readonly hasPassword: boolean
}

type ProtectedRangeState = Pick<
  LazyWorkbookState,
  'editJournal' | 'sheetProtectedRanges'
>

function normalizeRanges(
  ranges: readonly { readonly name: string; readonly sqref: string }[],
): EditableProtectedRange[] | string {
  const names = new Set<string>()
  const normalized: EditableProtectedRange[] = []
  for (const range of ranges) {
    const name = range.name.trim()
    const key = name.toLowerCase()
    if (!name || names.has(key)) return 'Allow-edit range names must be unique and non-empty.'
    const parts = range.sqref.trim().split(/\s+/).filter(Boolean)
    if (parts.length === 0) return `Allow-edit range ${name} needs a cell reference.`
    try {
      for (const part of parts) parseRange(part.replaceAll('$', ''))
    } catch {
      return `Allow-edit range ${name} has an invalid cell reference.`
    }
    names.add(key)
    normalized.push({ name, sqref: parts.join(' ').toUpperCase().replaceAll('$', ''), hasPassword: false })
  }
  return normalized
}

function sameRanges(left: readonly EditableProtectedRange[], right: readonly EditableProtectedRange[]): boolean {
  return JSON.stringify(left) === JSON.stringify(right)
}

/** Replaces one sheet's complete allow-edit-range set as one native Undo item. */
export function applyWorkbookProtectedRanges(
  runtime: UniverRuntime,
  state: ProtectedRangeState,
  sheetId: string,
  ranges: readonly { readonly name: string; readonly sqref: string }[],
  setPendingEdits?: (count: number) => void,
): string | null {
  if (isSheetRemoved(state.editJournal, sheetId)) return `Unknown sheet: ${sheetId}`
  const current = state.sheetProtectedRanges.get(sheetId)
  if (!current) return 'Allow-edit ranges are unavailable until the worksheet finishes loading.'
  if (current.some((range) => range.hasPassword)) {
    return 'Password- or permission-protected allow-edit ranges cannot be changed here.'
  }
  const next = normalizeRanges(ranges)
  if (typeof next === 'string') return next
  if (sameRanges(current, next)) return null
  const before = current.map((range) => ({ ...range }))
  const wasDirty = state.editJournal.protectedRangesDirty.has(sheetId)
  const install = (value: readonly EditableProtectedRange[], dirty: boolean): void => {
    state.sheetProtectedRanges.set(sheetId, value.map((range) => ({ ...range })))
    if (dirty) recordProtectedRangesChange(state.editJournal, sheetId)
    else state.editJournal.protectedRangesDirty.delete(sheetId)
    setPendingEdits?.(journalSize(state.editJournal))
  }
  install(next, true)
  pushWorkbookUndo(runtime, {
    undo: () => install(before, wasDirty),
    redo: () => install(next, true),
  })
  return null
}
