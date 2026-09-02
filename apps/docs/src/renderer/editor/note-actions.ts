import type { Editor } from '@tiptap/core'
import type { NoteInfo } from '@genoffice/docx-engine'
import { closeHistory } from '@tiptap/pm/history'

export type DocxNoteKind = 'footnote' | 'endnote'
export type ManagedDocxNoteKey = `${DocxNoteKind}:${string}`
export interface ManagedDocxNoteState {
  readonly original: NoteInfo | null
  readonly deleteWhenUnreferenced: boolean
}

export interface InsertDocxNoteInput {
  readonly range: { readonly from: number; readonly to: number }
  readonly kind: DocxNoteKind
  readonly noteId: number
  readonly text: string
}

export type InsertDocxNoteResult =
  | {
      readonly ok: true
      readonly from: number
      readonly to: number
      readonly kind: DocxNoteKind
      readonly noteId: string
      readonly number: number
      readonly changed: true
    }
  | {
      readonly ok: false
      readonly error: 'invalid_arguments' | 'execution_failed'
      readonly message: string
    }

export interface UpdateDocxNoteInput {
  readonly kind: DocxNoteKind
  readonly noteId: number
  readonly text: string
}

export interface DeleteDocxNoteInput {
  readonly kind: DocxNoteKind
  readonly noteId: number
}

export type DeleteDocxNoteResult =
  | {
      readonly ok: true
      readonly kind: DocxNoteKind
      readonly noteId: string
      readonly references: number
      readonly renumbered: number
      readonly changed: true
      readonly original: NoteInfo
    }
  | {
      readonly ok: false
      readonly error: 'invalid_arguments' | 'execution_failed'
      readonly message: string
    }

export type UpdateDocxNoteResult =
  | {
      readonly ok: true
      readonly kind: DocxNoteKind
      readonly noteId: string
      readonly references: number
      readonly changed: boolean
      readonly original: NoteInfo
    }
  | {
      readonly ok: false
      readonly error: 'invalid_arguments' | 'execution_failed'
      readonly message: string
    }

function managedKey(kind: DocxNoteKind, id: string): ManagedDocxNoteKey {
  return `${kind}:${id}`
}

/**
 * Insert one explicit note reference and its body metadata in the same native
 * editor transaction. The body metadata makes removal/restoration observable
 * when ProseMirror Undo/Redo removes or restores the reference atom.
 */
export function insertDocxNote(
  editor: Editor,
  existingNotes: readonly NoteInfo[],
  input: InsertDocxNoteInput,
): InsertDocxNoteResult {
  const { from, to } = input?.range ?? {}
  const textLength = typeof input?.text === 'string' ? Array.from(input.text).length : 0
  if (
    !Number.isInteger(from) ||
    !Number.isInteger(to) ||
    from < 1 ||
    to < from ||
    to > editor.state.doc.content.size ||
    (input.kind !== 'footnote' && input.kind !== 'endnote') ||
    !Number.isInteger(input.noteId) ||
    input.noteId < 1 ||
    input.noteId > 2_147_483_647 ||
    textLength < 1 ||
    textLength > 65_536
  ) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message:
        'DOCX note insertion requires one inline range, footnote/endnote kind, a positive bounded noteId, and 1 through 65536 Unicode characters.',
    }
  }

  const $from = editor.state.doc.resolve(from)
  const $to = editor.state.doc.resolve(to)
  if (!$from.sameParent($to) || !$from.parent.inlineContent) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: 'DOCX note insertion requires an exact range inside one inline text block.',
    }
  }

  const noteId = String(input.noteId)
  let duplicate = existingNotes.some((note) => note.id === noteId)
  editor.state.doc.descendants((node) => {
    if (
      node.type.name === 'docNoteRef' &&
      node.attrs.kind === input.kind &&
      String(node.attrs.id) === noteId
    ) {
      duplicate = true
    }
  })
  if (duplicate) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: `DOCX ${input.kind} noteId ${noteId} already exists.`,
    }
  }

  try {
    const ref = editor.schema.nodes.docNoteRef.create({
      kind: input.kind,
      id: noteId,
      num: 1,
      noteText: input.text,
      managed: true,
    })
    let transaction = editor.state.tr.replaceWith(from, to, ref)
    let number = 0
    transaction.doc.descendants((node, pos) => {
      if (node.type.name !== 'docNoteRef' || node.attrs.kind !== input.kind) return
      number++
      if (node.attrs.num !== number) {
        transaction = transaction.setNodeMarkup(pos, undefined, { ...node.attrs, num: number })
      }
    })
    editor.view.dispatch(closeHistory(transaction.scrollIntoView()))
    return { ok: true, from, to, kind: input.kind, noteId, number, changed: true }
  } catch (error) {
    return {
      ok: false,
      error: 'execution_failed',
      message: `The mounted DOCX editor rejected note insertion: ${
        error instanceof Error ? error.message : String(error)
      }`,
    }
  }
}

/** Update one stable note body while keeping its original body in native history. */
export function updateDocxNote(
  editor: Editor,
  existingNotes: readonly NoteInfo[],
  input: UpdateDocxNoteInput,
): UpdateDocxNoteResult {
  const textLength = typeof input?.text === 'string' ? Array.from(input.text).length : 0
  if (
    (input.kind !== 'footnote' && input.kind !== 'endnote') ||
    !Number.isInteger(input.noteId) ||
    input.noteId < 1 ||
    input.noteId > 2_147_483_647 ||
    textLength < 1 ||
    textLength > 65_536
  ) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message:
        'DOCX note update requires footnote/endnote kind, a positive bounded noteId, and 1 through 65536 Unicode characters.',
    }
  }

  const noteId = String(input.noteId)
  const original = existingNotes.find((note) => note.id === noteId)
  if (!original) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: `DOCX ${input.kind} noteId ${noteId} does not exist.`,
    }
  }

  const positions: number[] = []
  editor.state.doc.descendants((node, pos) => {
    if (
      node.type.name === 'docNoteRef' &&
      node.attrs.kind === input.kind &&
      String(node.attrs.id) === noteId
    ) {
      positions.push(pos)
    }
  })
  if (positions.length === 0) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: `DOCX ${input.kind} noteId ${noteId} has no document reference.`,
    }
  }
  if (original.text === input.text) {
    return {
      ok: true,
      kind: input.kind,
      noteId,
      references: positions.length,
      changed: false,
      original,
    }
  }

  try {
    let transaction = editor.state.tr
    for (const pos of positions) {
      const node = transaction.doc.nodeAt(pos)
      if (!node) continue
      transaction = transaction.setNodeMarkup(pos, undefined, {
        ...node.attrs,
        noteText: input.text,
        managed: true,
      })
    }
    editor.view.dispatch(closeHistory(transaction))
    return {
      ok: true,
      kind: input.kind,
      noteId,
      references: positions.length,
      changed: true,
      original,
    }
  } catch (error) {
    return {
      ok: false,
      error: 'execution_failed',
      message: `The mounted DOCX editor rejected note update: ${
        error instanceof Error ? error.message : String(error)
      }`,
    }
  }
}

/** Delete every reference to one stable note and renumber peers in one Undo unit. */
export function deleteDocxNote(
  editor: Editor,
  existingNotes: readonly NoteInfo[],
  input: DeleteDocxNoteInput,
): DeleteDocxNoteResult {
  if (
    (input.kind !== 'footnote' && input.kind !== 'endnote') ||
    !Number.isInteger(input.noteId) ||
    input.noteId < 1 ||
    input.noteId > 2_147_483_647
  ) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: 'DOCX note deletion requires footnote/endnote kind and a positive bounded noteId.',
    }
  }

  const noteId = String(input.noteId)
  const original = existingNotes.find((note) => note.id === noteId)
  if (!original) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: `DOCX ${input.kind} noteId ${noteId} does not exist.`,
    }
  }
  const positions: Array<{ pos: number; size: number }> = []
  editor.state.doc.descendants((node, pos) => {
    if (
      node.type.name === 'docNoteRef' &&
      node.attrs.kind === input.kind &&
      String(node.attrs.id) === noteId
    ) {
      positions.push({ pos, size: node.nodeSize })
    }
  })
  if (positions.length === 0) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: `DOCX ${input.kind} noteId ${noteId} has no document reference.`,
    }
  }

  try {
    let transaction = editor.state.tr
    for (const position of positions.reverse()) {
      transaction = transaction.delete(position.pos, position.pos + position.size)
    }
    let number = 0
    let renumbered = 0
    transaction.doc.descendants((node, pos) => {
      if (node.type.name !== 'docNoteRef' || node.attrs.kind !== input.kind) return
      number++
      if (node.attrs.num !== number) {
        transaction = transaction.setNodeMarkup(pos, undefined, { ...node.attrs, num: number })
        renumbered++
      }
    })
    editor.view.dispatch(closeHistory(transaction.scrollIntoView()))
    return {
      ok: true,
      kind: input.kind,
      noteId,
      references: positions.length,
      renumbered,
      changed: true,
      original,
    }
  } catch (error) {
    return {
      ok: false,
      error: 'execution_failed',
      message: `The mounted DOCX editor rejected note deletion: ${
        error instanceof Error ? error.message : String(error)
      }`,
    }
  }
}

export interface CollectedDocxNotes {
  readonly footnotes: NoteInfo[]
  readonly endnotes: NoteInfo[]
}

/**
 * Reconcile editor-managed notes with parsed note-part state. Existing orphan
 * notes stay intact; a managed note is present exactly while its reference atom
 * is present, so native Undo/Redo drives both marker and body final state.
 */
export function collectDocxNotes(
  editor: Editor,
  footnotes: readonly NoteInfo[],
  endnotes: readonly NoteInfo[],
  managedNotes: ReadonlyMap<ManagedDocxNoteKey, ManagedDocxNoteState>,
): CollectedDocxNotes {
  const fallback = {
    footnote: new Map(footnotes.map((note) => [note.id, note])),
    endnote: new Map(endnotes.map((note) => [note.id, note])),
  }
  const collected: Record<DocxNoteKind, NoteInfo[]> = { footnote: [], endnote: [] }
  const referenced: Record<DocxNoteKind, Set<string>> = {
    footnote: new Set(),
    endnote: new Set(),
  }

  editor.state.doc.descendants((node) => {
    if (node.type.name !== 'docNoteRef') return
    const kind: DocxNoteKind = node.attrs.kind === 'endnote' ? 'endnote' : 'footnote'
    const id = String(node.attrs.id ?? '')
    if (!id || referenced[kind].has(id)) return
    const existing = fallback[kind].get(id)
    const key = managedKey(kind, id)
    const noteText = typeof node.attrs.noteText === 'string' ? node.attrs.noteText : undefined
    if (noteText !== undefined) collected[kind].push({ id, text: noteText })
    else if (managedNotes.has(key)) {
      const original = managedNotes.get(key)?.original
      if (original) collected[kind].push(original)
    } else if (existing) collected[kind].push(existing)
    referenced[kind].add(id)
  })

  for (const [kind, notes] of [
    ['footnote', footnotes],
    ['endnote', endnotes],
  ] as const) {
    for (const note of notes) {
      if (referenced[kind].has(note.id)) continue
      const key = managedKey(kind, note.id)
      if (managedNotes.has(key)) {
        const state = managedNotes.get(key)!
        if (!state.deleteWhenUnreferenced && state.original) collected[kind].push(state.original)
        continue
      }
      collected[kind].push(note)
    }
  }

  return { footnotes: collected.footnote, endnotes: collected.endnote }
}
