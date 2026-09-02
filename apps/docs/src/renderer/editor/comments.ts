/**
 * Comment mark operations for Review → New/Delete Comment.
 *
 * The comment mark stores space-separated ids so overlapping comments share
 * one mark instance; adding/removing an id therefore rewrites the ids attr
 * per text node instead of blindly stacking marks.
 */
import type { Editor } from '@tiptap/core'
import type { CommentInfo } from '@genoffice/docx-engine'
import { closeHistory } from '@tiptap/pm/history'
import { TRACK_IGNORE } from './revisions'

export interface AddDocxCommentInput {
  readonly range: { readonly from: number; readonly to: number }
  readonly comment: {
    readonly id: string
    readonly author: string
    readonly initials: string | null
    readonly date: string
    readonly text: string
  }
}

export type AddDocxCommentResult =
  | {
      readonly ok: true
      readonly id: string
      readonly from: number
      readonly to: number
      readonly changed: true
      readonly comments: readonly CommentInfo[]
    }
  | {
      readonly ok: false
      readonly error: 'invalid_arguments' | 'execution_failed'
      readonly message: string
    }

export interface ReplyDocxCommentInput {
  readonly parentId: string
  readonly comment: AddDocxCommentInput['comment']
}

export type ReplyDocxCommentResult =
  | {
      readonly ok: true
      readonly id: string
      readonly parentId: string
      readonly references: number
      readonly changed: true
      readonly comments: readonly CommentInfo[]
    }
  | {
      readonly ok: false
      readonly error: 'invalid_arguments' | 'execution_failed'
      readonly message: string
    }

export interface SetDocxCommentResolvedInput {
  readonly id: string
  readonly resolved: boolean
}

export type SetDocxCommentResolvedResult =
  | {
      readonly ok: true
      readonly id: string
      readonly resolved: boolean
      readonly affected: number
      readonly changed: boolean
      readonly comments: readonly CommentInfo[]
    }
  | {
      readonly ok: false
      readonly error: 'invalid_arguments' | 'execution_failed'
      readonly message: string
    }

export interface DeleteDocxCommentInput {
  readonly id: string
}

export type DeleteDocxCommentResult =
  | {
      readonly ok: true
      readonly id: string
      readonly deleted: number
      readonly anchors: number
      readonly changed: true
      readonly comments: readonly CommentInfo[]
    }
  | {
      readonly ok: false
      readonly error: 'invalid_arguments' | 'execution_failed'
      readonly message: string
    }

function unicodeLength(value: unknown): number {
  return typeof value === 'string' ? Array.from(value).length : -1
}

/** Read the Undo-owned final comment snapshot, falling back to the parsed comments part. */
export function collectDocxComments(
  editor: Editor,
  parsedComments: readonly CommentInfo[],
): CommentInfo[] {
  const raw = editor.state.doc.firstChild?.attrs.commentStateOverride
  if (typeof raw !== 'string') return [...parsedComments]
  try {
    const value = JSON.parse(raw) as unknown
    return Array.isArray(value) ? (value as CommentInfo[]) : [...parsedComments]
  } catch {
    return [...parsedComments]
  }
}

/** Add one explicit comment body and exact range anchor in the same native-history transaction. */
export function addDocxComment(
  editor: Editor,
  currentComments: readonly CommentInfo[],
  input: AddDocxCommentInput,
): AddDocxCommentResult {
  const { from, to } = input?.range ?? {}
  const comment = input?.comment
  if (
    !Number.isInteger(from) ||
    !Number.isInteger(to) ||
    from < 1 ||
    to <= from ||
    to > editor.state.doc.content.size ||
    !comment ||
    unicodeLength(comment.id) < 1 ||
    unicodeLength(comment.id) > 10 ||
    !/^[0-9]+$/.test(comment.id) ||
    currentComments.some((entry) => entry.id === comment.id) ||
    currentComments.length >= 1024 ||
    unicodeLength(comment.author) < 1 ||
    unicodeLength(comment.author) > 255 ||
    !(comment.initials === null || unicodeLength(comment.initials) <= 16) ||
    unicodeLength(comment.date) < 20 ||
    unicodeLength(comment.date) > 32 ||
    !/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(?:\.\d{1,3})?Z$/.test(comment.date) ||
    Number.isNaN(Date.parse(comment.date)) ||
    unicodeLength(comment.text) < 1 ||
    unicodeLength(comment.text) > 65_536
  ) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message:
        'DOCX comment addition requires one non-empty exact range and unique bounded explicit comment metadata.',
    }
  }

  const normalized: CommentInfo = {
    id: comment.id,
    author: comment.author,
    ...(comment.initials ? { initials: comment.initials } : {}),
    date: comment.date,
    text: comment.text,
  }
  const comments = [...currentComments, normalized]
  if (unicodeLength(JSON.stringify(comments)) > 1_048_576) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: 'DOCX comment snapshot exceeds 1048576 Unicode characters.',
    }
  }

  try {
    const anchor = editor.state.doc.firstChild
    if (!anchor) throw new Error('document has no comment-state anchor')
    const markType = editor.state.schema.marks.comment
    const transaction = editor.state.tr
    let matched = false
    editor.state.doc.nodesBetween(from, to, (node, position) => {
      if (!node.isText) return
      const start = Math.max(position, from)
      const end = Math.min(position + node.nodeSize, to)
      if (start >= end) return
      matched = true
      const existing = node.marks.find((mark) => mark.type === markType)
      const ids = new Set(
        String(existing?.attrs.ids ?? '')
          .split(' ')
          .filter(Boolean),
      )
      ids.add(comment.id)
      transaction.addMark(start, end, markType.create({ ids: [...ids].sort().join(' ') }))
    })
    if (!matched) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'DOCX comment range must cover editable text.',
      }
    }
    transaction.setNodeMarkup(0, undefined, {
      ...anchor.attrs,
      commentStateOverride: JSON.stringify(comments),
    })
    transaction.setMeta(TRACK_IGNORE, true)
    editor.view.dispatch(closeHistory(transaction))
    return { ok: true, id: comment.id, from, to, changed: true, comments }
  } catch (error) {
    return {
      ok: false,
      error: 'execution_failed',
      message: `The mounted DOCX editor rejected comment addition: ${
        error instanceof Error ? error.message : String(error)
      }`,
    }
  }
}

/** Add one explicit reply to every text range anchored by a stable parent comment. */
export function replyDocxComment(
  editor: Editor,
  currentComments: readonly CommentInfo[],
  input: ReplyDocxCommentInput,
): ReplyDocxCommentResult {
  const comment = input?.comment
  const parent = currentComments.find((entry) => entry.id === input?.parentId && !entry.parentId)
  if (
    !parent ||
    unicodeLength(input.parentId) < 1 ||
    unicodeLength(input.parentId) > 10 ||
    !/^[0-9]+$/.test(input.parentId) ||
    !comment ||
    unicodeLength(comment.id) < 1 ||
    unicodeLength(comment.id) > 10 ||
    !/^[0-9]+$/.test(comment.id) ||
    currentComments.some((entry) => entry.id === comment.id) ||
    currentComments.length >= 1024 ||
    unicodeLength(comment.author) < 1 ||
    unicodeLength(comment.author) > 255 ||
    !(comment.initials === null || unicodeLength(comment.initials) <= 16) ||
    unicodeLength(comment.date) < 20 ||
    unicodeLength(comment.date) > 32 ||
    !/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(?:\.\d{1,3})?Z$/.test(comment.date) ||
    Number.isNaN(Date.parse(comment.date)) ||
    unicodeLength(comment.text) < 1 ||
    unicodeLength(comment.text) > 65_536
  ) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message:
        'DOCX comment reply requires one anchored top-level parent and unique bounded explicit reply metadata.',
    }
  }

  const normalized: CommentInfo = {
    id: comment.id,
    author: comment.author,
    ...(comment.initials ? { initials: comment.initials } : {}),
    date: comment.date,
    text: comment.text,
    parentId: input.parentId,
  }
  const comments = [...currentComments, normalized]
  if (unicodeLength(JSON.stringify(comments)) > 1_048_576) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: 'DOCX comment snapshot exceeds 1048576 Unicode characters.',
    }
  }

  try {
    const anchor = editor.state.doc.firstChild
    if (!anchor) throw new Error('document has no comment-state anchor')
    const markType = editor.state.schema.marks.comment
    const transaction = editor.state.tr
    let references = 0
    editor.state.doc.descendants((node, position) => {
      if (!node.isText) return
      const existing = node.marks.find((mark) => mark.type === markType)
      if (!existing) return
      const ids = String(existing.attrs.ids ?? '')
        .split(' ')
        .filter(Boolean)
      if (!ids.includes(input.parentId)) return
      references += 1
      transaction.addMark(
        position,
        position + node.nodeSize,
        markType.create({ ids: [...new Set([...ids, comment.id])].sort().join(' ') }),
      )
    })
    if (references === 0) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: `DOCX parent comment ${input.parentId} has no editable anchor.`,
      }
    }
    transaction.setNodeMarkup(0, undefined, {
      ...anchor.attrs,
      commentStateOverride: JSON.stringify(comments),
    })
    transaction.setMeta(TRACK_IGNORE, true)
    editor.view.dispatch(closeHistory(transaction))
    return {
      ok: true,
      id: comment.id,
      parentId: input.parentId,
      references,
      changed: true,
      comments,
    }
  } catch (error) {
    return {
      ok: false,
      error: 'execution_failed',
      message: `The mounted DOCX editor rejected comment reply: ${
        error instanceof Error ? error.message : String(error)
      }`,
    }
  }
}

/** Set one top-level comment thread to an explicit final resolved state. */
export function setDocxCommentResolved(
  editor: Editor,
  currentComments: readonly CommentInfo[],
  input: SetDocxCommentResolvedInput,
): SetDocxCommentResolvedResult {
  const parent = currentComments.find((entry) => entry.id === input?.id && !entry.parentId)
  if (
    !parent ||
    unicodeLength(input.id) < 1 ||
    unicodeLength(input.id) > 10 ||
    !/^[0-9]+$/.test(input.id) ||
    typeof input.resolved !== 'boolean'
  ) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: 'DOCX comment resolution requires one stable top-level thread and final boolean.',
    }
  }
  const affected = currentComments.filter(
    (entry) => entry.id === input.id || entry.parentId === input.id,
  ).length
  const changed = currentComments.some(
    (entry) =>
      (entry.id === input.id || entry.parentId === input.id) &&
      Boolean(entry.done) !== input.resolved,
  )
  const comments = currentComments.map((entry) =>
    entry.id === input.id || entry.parentId === input.id
      ? { ...entry, done: input.resolved }
      : entry,
  )
  if (!changed) {
    return { ok: true, id: input.id, resolved: input.resolved, affected, changed, comments }
  }

  try {
    const anchor = editor.state.doc.firstChild
    if (!anchor) throw new Error('document has no comment-state anchor')
    const transaction = editor.state.tr.setNodeMarkup(0, undefined, {
      ...anchor.attrs,
      commentStateOverride: JSON.stringify(comments),
    })
    transaction.setMeta(TRACK_IGNORE, true)
    editor.view.dispatch(closeHistory(transaction))
    return { ok: true, id: input.id, resolved: input.resolved, affected, changed, comments }
  } catch (error) {
    return {
      ok: false,
      error: 'execution_failed',
      message: `The mounted DOCX editor rejected comment resolution: ${
        error instanceof Error ? error.message : String(error)
      }`,
    }
  }
}

/** Delete one stable comment, cascading direct replies for a top-level thread. */
export function deleteDocxComment(
  editor: Editor,
  currentComments: readonly CommentInfo[],
  input: DeleteDocxCommentInput,
): DeleteDocxCommentResult {
  const target = currentComments.find((entry) => entry.id === input?.id)
  if (
    !target ||
    unicodeLength(input.id) < 1 ||
    unicodeLength(input.id) > 10 ||
    !/^[0-9]+$/.test(input.id)
  ) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: 'DOCX comment deletion requires one existing stable numeric comment ID.',
    }
  }
  const victims = new Set([
    input.id,
    ...(target.parentId
      ? []
      : currentComments.filter((entry) => entry.parentId === input.id).map((entry) => entry.id)),
  ])
  const comments = currentComments.filter((entry) => !victims.has(entry.id))

  try {
    const anchor = editor.state.doc.firstChild
    if (!anchor) throw new Error('document has no comment-state anchor')
    const markType = editor.state.schema.marks.comment
    const transaction = editor.state.tr
    let anchors = 0
    editor.state.doc.descendants((node, position) => {
      if (!node.isText) return
      const existing = node.marks.find((mark) => mark.type === markType)
      if (!existing) return
      const ids = String(existing.attrs.ids ?? '')
        .split(' ')
        .filter(Boolean)
      if (!ids.some((id) => victims.has(id))) return
      anchors += 1
      const remaining = ids.filter((id) => !victims.has(id))
      transaction.removeMark(position, position + node.nodeSize, markType)
      if (remaining.length > 0) {
        transaction.addMark(
          position,
          position + node.nodeSize,
          markType.create({ ids: remaining.join(' ') }),
        )
      }
    })
    transaction.setNodeMarkup(0, undefined, {
      ...anchor.attrs,
      commentStateOverride: JSON.stringify(comments),
    })
    transaction.setMeta(TRACK_IGNORE, true)
    editor.view.dispatch(closeHistory(transaction))
    return {
      ok: true,
      id: input.id,
      deleted: victims.size,
      anchors,
      changed: true,
      comments,
    }
  } catch (error) {
    return {
      ok: false,
      error: 'execution_failed',
      message: `The mounted DOCX editor rejected comment deletion: ${
        error instanceof Error ? error.message : String(error)
      }`,
    }
  }
}

/** smallest unused numeric comment id */
export function nextCommentId(comments: CommentInfo[]): string {
  const max = comments.reduce((acc, c) => Math.max(acc, parseInt(c.id, 10) || 0), 0)
  return String(max + 1)
}

/** attach `id` to every text node in the current selection; false when selection is empty */
export function addCommentToSelection(editor: Editor, id: string): boolean {
  const { state } = editor
  const { from, to } = state.selection
  if (from === to) return false
  const markType = state.schema.marks.comment
  const tr = state.tr
  tr.setMeta(TRACK_IGNORE, true)
  state.doc.nodesBetween(from, to, (node, pos) => {
    if (!node.isText) return
    const start = Math.max(pos, from)
    const end = Math.min(pos + node.nodeSize, to)
    if (start >= end) return
    const existing = node.marks.find((m) => m.type === markType)
    const ids = new Set(
      String(existing?.attrs.ids ?? '')
        .split(' ')
        .filter(Boolean),
    )
    ids.add(id)
    tr.addMark(start, end, markType.create({ ids: [...ids].sort().join(' ') }))
  })
  editor.view.dispatch(tr)
  return true
}

/** strip `id` from every comment mark in the document (mark removed when it was the last id) */
export function removeCommentFromDoc(editor: Editor, id: string): void {
  const { state } = editor
  const markType = state.schema.marks.comment
  const tr = state.tr
  tr.setMeta(TRACK_IGNORE, true)
  state.doc.descendants((node, pos) => {
    if (!node.isText) return
    const existing = node.marks.find((m) => m.type === markType)
    if (!existing) return
    const ids = String(existing.attrs.ids ?? '')
      .split(' ')
      .filter(Boolean)
    if (!ids.includes(id)) return
    const remaining = ids.filter((x) => x !== id)
    const from = pos
    const to = pos + node.nodeSize
    if (remaining.length === 0) tr.removeMark(from, to, markType)
    else tr.addMark(from, to, markType.create({ ids: remaining.join(' ') }))
  })
  if (tr.steps.length > 0) editor.view.dispatch(tr)
}

/** Append `newId` to every text range anchored by the `parentId` comment (Word: replies share the parent anchor) */
export function addReplyToCommentRange(editor: Editor, parentId: string, newId: string): boolean {
  const { state } = editor
  const markType = state.schema.marks.comment
  const tr = state.tr
  tr.setMeta(TRACK_IGNORE, true)
  let found = false
  state.doc.descendants((node, pos) => {
    if (!node.isText) return
    const existing = node.marks.find((m) => m.type === markType)
    if (!existing) return
    const ids = String(existing.attrs.ids ?? '')
      .split(' ')
      .filter(Boolean)
    if (!ids.includes(parentId)) return
    found = true
    const merged = [...new Set([...ids, newId])].sort()
    tr.addMark(pos, pos + node.nodeSize, markType.create({ ids: merged.join(' ') }))
  })
  if (found && tr.steps.length > 0) editor.view.dispatch(tr)
  return found
}
