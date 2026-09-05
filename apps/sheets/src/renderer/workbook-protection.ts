import { journalSize, recordWorkbookProtection } from './edit-journal'
import { pushWorkbookUndo } from './univer-sync'
import type { LazyWorkbookState, UniverRuntime } from './univer-state'

type WorkbookProtectionState = Pick<LazyWorkbookState, 'editJournal' | 'file'>

export function workbookStructureLocked(state: WorkbookProtectionState): boolean {
  return (
    state.editJournal.workbookProtection.desired ??
    state.file.workbookProtection?.lockStructure ??
    false
  )
}

/** Records one passwordless workbook structure-lock change with native Undo. */
export function applyWorkbookStructureProtection(
  runtime: UniverRuntime,
  state: WorkbookProtectionState,
  lockStructure: boolean,
  setPendingEdits?: (count: number) => void,
): string | null {
  const original = state.file.workbookProtection?.lockStructure ?? false
  if (!lockStructure && state.file.workbookProtection?.hasPassword) {
    return 'This workbook structure is password-protected and cannot be unlocked here.'
  }
  const before = state.editJournal.workbookProtection.desired
  recordWorkbookProtection(state.editJournal, lockStructure, original)
  const after = state.editJournal.workbookProtection.desired
  setPendingEdits?.(journalSize(state.editJournal))
  if (before === after) return null
  pushWorkbookUndo(runtime, {
    undo: () => {
      state.editJournal.workbookProtection.desired = before
      setPendingEdits?.(journalSize(state.editJournal))
    },
    redo: () => {
      state.editJournal.workbookProtection.desired = after
      setPendingEdits?.(journalSize(state.editJournal))
    },
  })
  return null
}
