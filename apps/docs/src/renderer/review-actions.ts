/**
 * Review-tab actions: footnotes/endnotes, comments, revisions, ink
 * annotations, document protection and compare. Extracted from App.tsx; the
 * App component passes a ReviewContext built fresh per call so state never
 * goes stale.
 */
import type { Editor } from '@tiptap/core'
import {
  nextNoteId,
  type CommentInfo,
  type DocProtection,
  type NoteInfo,
} from '@genoffice/docx-engine'
import type { Dispatch, SetStateAction } from 'react'
import type { DocState } from './doc-state'
import {
  addDocxComment,
  deleteDocxComment,
  nextCommentId,
  replyDocxComment,
  setDocxCommentResolved,
} from './editor/comments'
import { compareDocxBytes, type CompareEntry } from './editor/compare'
import { pendingCommentPluginKey } from './editor/extensions'
import {
  collectDocxNotes,
  deleteDocxNote,
  insertDocxNote,
  updateDocxNote,
  type ManagedDocxNoteKey,
  type ManagedDocxNoteState,
} from './editor/note-actions'
import { applyDocxInk, type InkAnnotation } from './editor/ink'
import { setDocxProtection } from './editor/protection-actions'
import { applyDocxRevisionDecision } from './editor/revisions'
import { t } from './i18n/locale'

/** open the note text dialog; id set = editing an existing note */
export interface NotePrompt {
  kind: 'footnote' | 'endnote'
  id?: string
}

/** Protection toggle dialog: set = enable (password may be blank), unlock = removing requires password verification */
export interface ProtectModalState {
  mode: 'set' | 'unlock'
  value: string
  error?: string
}

/** The App state the review actions need; built fresh per call. */
export interface ReviewContext {
  editor: Editor | null
  doc: DocState | null
  dirtyRef: { current: boolean }
  setStatus: (status: string) => void
  notePrompt: NotePrompt | null
  setNotePrompt: (value: NotePrompt | null) => void
  footnotes: NoteInfo[]
  endnotes: NoteInfo[]
  setFootnotes: Dispatch<SetStateAction<NoteInfo[]>>
  setEndnotes: Dispatch<SetStateAction<NoteInfo[]>>
  setNotesDirty: (dirty: boolean) => void
  managedNotes: Map<ManagedDocxNoteKey, ManagedDocxNoteState>
  comments: CommentInfo[]
  setComments: Dispatch<SetStateAction<CommentInfo[]>>
  setCommentsDirty: (dirty: boolean) => void
  setCommentComposing: (composing: boolean) => void
  setShowComments: (show: boolean) => void
  inkAnnotations: InkAnnotation[]
  setInkAnnotations: Dispatch<SetStateAction<InkAnnotation[]>>
  setInksDirty: (dirty: boolean) => void
  protection: DocProtection | null
  setProtection: (value: DocProtection | null) => void
  setProtectionDirty: (dirty: boolean) => void
  protectModal: ProtectModalState | null
  setProtectModal: (value: ProtectModalState | null) => void
  setCompareResult: (value: { otherName: string; entries: CompareEntry[] } | null) => void
}

// ---- References: footnotes / endnotes ----

/** dialog submit: create a new note (+ caret marker) or update an existing one */
export function submitNote(ctx: ReviewContext, text: string): void {
  if (!ctx.notePrompt || !ctx.editor) return
  const { kind, id } = ctx.notePrompt
  const list = kind === 'footnote' ? ctx.footnotes : ctx.endnotes
  if (id !== undefined) {
    const result = updateDocxNote(ctx.editor, list, {
      kind,
      noteId: Number(id),
      text,
    })
    if (!result.ok) {
      ctx.setStatus(result.message)
      return
    }
    if (!result.changed) return
    const key: ManagedDocxNoteKey = `${kind}:${id}`
    if (!ctx.managedNotes.has(key)) {
      ctx.managedNotes.set(key, { original: result.original, deleteWhenUnreferenced: false })
    }
  } else {
    const newId = nextNoteId(list)
    const { from, to } = ctx.editor.state.selection
    const result = insertDocxNote(ctx.editor, list, {
      range: { from, to },
      kind,
      noteId: Number(newId),
      text,
    })
    if (!result.ok) {
      ctx.setStatus(result.message)
      return
    }
    ctx.managedNotes.set(`${kind}:${newId}`, {
      original: null,
      deleteWhenUnreferenced: true,
    })
  }
  const next = collectDocxNotes(ctx.editor, ctx.footnotes, ctx.endnotes, ctx.managedNotes)
  ctx.setFootnotes(next.footnotes)
  ctx.setEndnotes(next.endnotes)
  ctx.setNotesDirty(true)
}

/** delete a note, remove its in-text marker, renumber the remaining markers */
export function deleteNote(ctx: ReviewContext, kind: 'footnote' | 'endnote', id: string): void {
  const list = kind === 'footnote' ? ctx.footnotes : ctx.endnotes
  if (!ctx.editor) return
  const result = deleteDocxNote(ctx.editor, list, { kind, noteId: Number(id) })
  if (!result.ok) {
    ctx.setStatus(result.message)
    return
  }
  const key: ManagedDocxNoteKey = `${kind}:${id}`
  const baseline = ctx.managedNotes.get(key)?.original ?? result.original
  ctx.managedNotes.set(key, { original: baseline, deleteWhenUnreferenced: true })
  const next = collectDocxNotes(ctx.editor, ctx.footnotes, ctx.endnotes, ctx.managedNotes)
  ctx.setFootnotes(next.footnotes)
  ctx.setEndnotes(next.endnotes)
  ctx.setNotesDirty(true)
}

// ---- Review: comments / revisions / compare / protection ----

/** keeps the picked range highlighted while the composer holds focus */
export function setPendingCommentRange(
  ctx: ReviewContext,
  range: { from: number; to: number } | null,
): void {
  if (!ctx.editor) return
  ctx.editor.view.dispatch(ctx.editor.state.tr.setMeta(pendingCommentPluginKey, range))
}

export function cancelNewComment(ctx: ReviewContext): void {
  ctx.setCommentComposing(false)
  setPendingCommentRange(ctx, null)
}

/** New comment: open the pane with the composer; the mark is applied on submit */
export function startNewComment(ctx: ReviewContext): void {
  if (!ctx.editor || ctx.editor.state.selection.empty) {
    ctx.setStatus(t('appSelectTextToComment'))
    return
  }
  const { from, to } = ctx.editor.state.selection
  setPendingCommentRange(ctx, { from, to })
  ctx.setShowComments(true)
  ctx.setCommentComposing(true)
}

export function submitNewComment(ctx: ReviewContext, text: string): void {
  if (!ctx.editor) return
  const pending = pendingCommentPluginKey.getState(ctx.editor.state)?.find()[0]
  if (!pending) {
    ctx.setStatus(t('appCommentSelectionLost'))
    ctx.setCommentComposing(false)
    return
  }
  const id = nextCommentId(ctx.comments)
  const now = new Date().toISOString().replace(/\.\d{3}Z$/, 'Z')
  const result = addDocxComment(ctx.editor, ctx.comments, {
    range: { from: pending.from, to: pending.to },
    comment: { id, author: 'User', initials: null, date: now, text },
  })
  setPendingCommentRange(ctx, null)
  if (!result.ok) {
    ctx.setStatus(t('appCommentSelectionLost'))
    ctx.setCommentComposing(false)
    return
  }
  ctx.setComments([...result.comments])
  ctx.setCommentsDirty(true)
  ctx.setCommentComposing(false)
  ctx.dirtyRef.current = true
  ctx.setStatus(t('appCommentAdded'))
}

/** Reply to a comment: the new entry carries parentId; the anchor shares the parent comment's range */
export function replyToComment(ctx: ReviewContext, parentId: string, text: string): void {
  if (!ctx.editor) return
  const id = nextCommentId(ctx.comments)
  const now = new Date().toISOString().replace(/\.\d{3}Z$/, 'Z')
  const result = replyDocxComment(ctx.editor, ctx.comments, {
    parentId,
    comment: { id, author: 'User', initials: null, date: now, text },
  })
  if (!result.ok) {
    ctx.setStatus(t('appCommentAnchorGone'))
    return
  }
  ctx.setComments([...result.comments])
  ctx.setCommentsDirty(true)
  ctx.dirtyRef.current = true
  ctx.setStatus(t('appCommentReplied'))
}

/** Resolve/reopen: the whole thread (parent + replies) gets done set together */
export function resolveComment(ctx: ReviewContext, id: string, done: boolean): void {
  if (!ctx.editor) return
  const result = setDocxCommentResolved(ctx.editor, ctx.comments, { id, resolved: done })
  if (!result.ok) {
    ctx.setStatus(result.message)
    return
  }
  ctx.setComments([...result.comments])
  ctx.setCommentsDirty(true)
  ctx.dirtyRef.current = true
  ctx.setStatus(done ? t('appCommentResolvedMsg') : t('appCommentReopenedMsg'))
}

export function deleteComment(ctx: ReviewContext, id: string): void {
  if (!ctx.editor) return
  const result = deleteDocxComment(ctx.editor, ctx.comments, { id })
  if (!result.ok) {
    ctx.setStatus(result.message)
    return
  }
  ctx.setComments([...result.comments])
  ctx.setCommentsDirty(true)
  ctx.dirtyRef.current = true
}

export function handleRevision(
  ctx: ReviewContext,
  action: 'accept' | 'reject',
  all: boolean,
): void {
  if (!ctx.editor) return
  const result = applyDocxRevisionDecision(ctx.editor, {
    decision: action,
    scope: all ? 'all' : 'current',
  })
  if (!result.ok) {
    ctx.setStatus(t('appNoRevisionsToHandle'))
    return
  }
  if (all) {
    ctx.setStatus(action === 'accept' ? t('appAllRevisionsAccepted') : t('appAllRevisionsRejected'))
  }
  ctx.dirtyRef.current = true
}

// ---- Draw: overlay annotations ----

export function addInk(ctx: ReviewContext, annotation: InkAnnotation): void {
  if (!annotation.stroke) return
  const result = applyDocxInk(ctx.inkAnnotations, {
    action: 'add',
    annotation: {
      id: annotation.id,
      anchorIndex: annotation.anchorIndex,
      ...annotation.stroke,
    },
  })
  if (!result.ok) return
  ctx.setInkAnnotations([...result.annotations])
  ctx.setInksDirty(true)
  ctx.dirtyRef.current = true
}

export function removeInks(ctx: ReviewContext, ids: string[]): void {
  const result = applyDocxInk(ctx.inkAnnotations, { action: 'delete', ids })
  if (!result.ok) return
  ctx.setInkAnnotations([...result.annotations])
  ctx.setInksDirty(true)
  ctx.dirtyRef.current = true
}

export function clearInks(ctx: ReviewContext): void {
  const result = applyDocxInk(ctx.inkAnnotations, { action: 'clear' })
  if (!result.ok || !result.changed) return
  ctx.setInkAnnotations([...result.annotations])
  ctx.setInksDirty(true)
  ctx.dirtyRef.current = true
  ctx.setStatus(t('appInksCleared'))
}

export function toggleProtection(ctx: ReviewContext): void {
  if (ctx.protection?.enforced && ctx.protection.edit === 'readOnly') {
    if (ctx.protection.hash) {
      ctx.setProtectModal({ mode: 'unlock', value: '' })
    } else {
      void setDocxProtection(ctx.protection, { enabled: false, password: null }).then((result) => {
        if (!result.ok || !result.changed) return
        ctx.setProtection(result.protection)
        ctx.setProtectionDirty(true)
        ctx.dirtyRef.current = true
        ctx.setStatus(t('appProtectionRemoved'))
      })
    }
  } else {
    ctx.setProtectModal({ mode: 'set', value: '' })
  }
}

export async function submitProtectModal(ctx: ReviewContext): Promise<void> {
  if (!ctx.protectModal) return
  if (ctx.protectModal.mode === 'set') {
    const pwd = ctx.protectModal.value
    const result = await setDocxProtection(ctx.protection, {
      enabled: true,
      password: pwd || null,
    })
    if (!result.ok) {
      ctx.setStatus(result.message)
      return
    }
    if (result.changed) {
      ctx.setProtection(result.protection)
      ctx.setProtectionDirty(true)
      ctx.dirtyRef.current = true
    }
    ctx.setProtectModal(null)
    ctx.setStatus(pwd ? t('appProtectionEnabledPwd') : t('appProtectionEnabled'))
  } else {
    const result = await setDocxProtection(ctx.protection, {
      enabled: false,
      password: ctx.protectModal.value || null,
    })
    if (!result.ok) {
      ctx.setProtectModal({ ...ctx.protectModal, error: t('appWrongPassword') })
      return
    }
    if (result.changed) {
      ctx.setProtection(result.protection)
      ctx.setProtectionDirty(true)
      ctx.dirtyRef.current = true
    }
    ctx.setProtectModal(null)
    ctx.setStatus(t('appProtectionRemoved'))
  }
}

/** Compare: pick a second .docx and diff it against the open document */
export async function compareWithFile(ctx: ReviewContext): Promise<void> {
  if (!ctx.doc) return
  const other = await window.desktop.openDocx()
  if (!other) return
  try {
    const result = await compareDocxBytes(ctx.doc.parsed.blocks, {
      name: other.name,
      data: other.data,
    })
    if (!result.ok) throw new Error(result.message)
    ctx.setCompareResult({ otherName: result.otherName, entries: result.entries })
  } catch (err) {
    ctx.setStatus(t('appCompareFailed', { error: String(err) }))
  }
}
