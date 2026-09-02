import type { Editor } from '@tiptap/core'
import type { Node as PmNode } from '@tiptap/pm/model'
import { closeHistory } from '@tiptap/pm/history'
import { TextSelection } from '@tiptap/pm/state'

import { isValidDocxBookmarkName } from './bookmark-actions'

export interface InsertDocxCrossReferenceInput {
  readonly range: { readonly from: number; readonly to: number }
  readonly bookmarkName: string
  readonly displayText: string
}

export const docxFieldInstructionValues = ['DATE', 'TIME', 'PAGE', 'NUMPAGES', 'FILENAME'] as const
export type DocxFieldInstruction = (typeof docxFieldInstructionValues)[number]

export interface InsertDocxFieldInput {
  readonly range: { readonly from: number; readonly to: number }
  readonly instruction: DocxFieldInstruction
  readonly displayText: string
}

export const docxUpdatableFieldKeywordValues = [
  'DATE',
  'TIME',
  'PAGE',
  'NUMPAGES',
  'FILENAME',
  'CREATEDATE',
  'SAVEDATE',
] as const

export interface UpdateDocxFieldCacheInput {
  readonly range: { readonly from: number; readonly to: number }
  /** Exact retained instruction, including any Word formatting switches. */
  readonly instruction: string
  readonly displayText: string
}

export interface UpdateDocxFieldsInput {
  readonly updates: readonly UpdateDocxFieldCacheInput[]
}

type DocxFieldActionFailure = {
  readonly ok: false
  readonly error: 'invalid_arguments' | 'execution_failed'
  readonly message: string
}

export type UpdateDocxFieldsResult =
  | {
      readonly ok: true
      readonly matched: number
      readonly changed: number
    }
  | DocxFieldActionFailure

export type InsertDocxFieldResult =
  | {
      readonly ok: true
      readonly changed: true
      readonly from: number
      readonly to: number
      readonly instruction: DocxFieldInstruction
    }
  | DocxFieldActionFailure

export type InsertDocxCrossReferenceResult =
  | {
      readonly ok: true
      readonly changed: true
      readonly from: number
      readonly to: number
      readonly bookmarkName: string
    }
  | DocxFieldActionFailure

function invalid(message: string): DocxFieldActionFailure {
  return { ok: false, error: 'invalid_arguments', message }
}

function unicodeLength(value: string): number {
  return Array.from(value).length
}

/** Insert one exact REF field through the mounted editor's native history. */
export function insertDocxCrossReference(
  editor: Editor,
  input: InsertDocxCrossReferenceInput,
): InsertDocxCrossReferenceResult {
  if (!input || !input.range) return invalid('DOCX cross-references require an explicit range.')
  const { from, to } = input.range
  const { doc } = editor.state
  if (
    !Number.isInteger(from) ||
    !Number.isInteger(to) ||
    from < 1 ||
    from > to ||
    to > doc.content.size
  ) {
    return invalid(
      `DOCX cross-reference range ${from}..${to} is outside document size ${doc.content.size}.`,
    )
  }
  const $from = doc.resolve(from)
  const $to = doc.resolve(to)
  if ($from.parent !== $to.parent || !$from.parent.inlineContent) {
    return invalid('DOCX cross-reference ranges must stay inside one text-bearing block.')
  }
  if (typeof input.bookmarkName !== 'string' || !isValidDocxBookmarkName(input.bookmarkName)) {
    return invalid('DOCX cross-references require a valid 1 through 40 character bookmark name.')
  }
  if (
    typeof input.displayText !== 'string' ||
    unicodeLength(input.displayText) < 1 ||
    unicodeLength(input.displayText) > 65_536
  ) {
    return invalid('DOCX cross-reference display text must contain 1 through 65536 characters.')
  }

  let targetExists = false
  doc.forEach((node) => {
    if (
      Array.isArray(node.attrs.bookmarks) &&
      (node.attrs.bookmarks as readonly unknown[]).includes(input.bookmarkName)
    ) {
      targetExists = true
    }
  })
  if (!targetExists) {
    return invalid(`DOCX bookmark ${input.bookmarkName} does not exist in the current document.`)
  }

  try {
    const marked = editor.state.schema.text(input.displayText, [
      editor.state.schema.marks.refField.create({ name: input.bookmarkName }),
    ])
    const resultTo = from + input.displayText.length
    let transaction = editor.state.tr.replaceWith(from, to, marked)
    transaction = transaction.setSelection(TextSelection.create(transaction.doc, resultTo))
    editor.view.dispatch(closeHistory(transaction))
    return {
      ok: true,
      changed: true,
      from,
      to: resultTo,
      bookmarkName: input.bookmarkName,
    }
  } catch (error) {
    return {
      ok: false,
      error: 'execution_failed',
      message: `The mounted DOCX editor rejected the cross-reference insertion: ${
        error instanceof Error ? error.message : String(error)
      }`,
    }
  }
}

/** Insert one finite generic field through the mounted editor's native history. */
export function insertDocxField(
  editor: Editor,
  input: InsertDocxFieldInput,
): InsertDocxFieldResult {
  if (!input || !input.range) return invalid('DOCX fields require an explicit range.')
  const { from, to } = input.range
  const { doc } = editor.state
  if (
    !Number.isInteger(from) ||
    !Number.isInteger(to) ||
    from < 1 ||
    from > to ||
    to > doc.content.size
  ) {
    return invalid(`DOCX field range ${from}..${to} is outside document size ${doc.content.size}.`)
  }
  const $from = doc.resolve(from)
  const $to = doc.resolve(to)
  if ($from.parent !== $to.parent || !$from.parent.inlineContent) {
    return invalid('DOCX field ranges must stay inside one text-bearing block.')
  }
  if (
    typeof input.instruction !== 'string' ||
    !docxFieldInstructionValues.includes(input.instruction as DocxFieldInstruction)
  ) {
    return invalid(`Unsupported DOCX field instruction ${String(input.instruction)}.`)
  }
  if (
    typeof input.displayText !== 'string' ||
    unicodeLength(input.displayText) < 1 ||
    unicodeLength(input.displayText) > 65_536
  ) {
    return invalid('DOCX field display text must contain 1 through 65536 characters.')
  }

  try {
    const marked = editor.state.schema.text(input.displayText, [
      editor.state.schema.marks.instrField.create({ instr: input.instruction }),
    ])
    const resultTo = from + input.displayText.length
    let transaction = editor.state.tr.replaceWith(from, to, marked)
    transaction = transaction.setSelection(TextSelection.create(transaction.doc, resultTo))
    editor.view.dispatch(closeHistory(transaction))
    return {
      ok: true,
      changed: true,
      from,
      to: resultTo,
      instruction: input.instruction,
    }
  } catch (error) {
    return {
      ok: false,
      error: 'execution_failed',
      message: `The mounted DOCX editor rejected the field insertion: ${
        error instanceof Error ? error.message : String(error)
      }`,
    }
  }
}

/** Replace exact generic-field caches in one native Undo transaction. */
export function updateDocxFields(
  editor: Editor,
  input: UpdateDocxFieldsInput,
): UpdateDocxFieldsResult {
  if (!input || !Array.isArray(input.updates) || input.updates.length < 1) {
    return invalid('DOCX field updates require at least one explicit field cache.')
  }
  if (input.updates.length > 1024) {
    return invalid('DOCX field updates may contain at most 1024 entries.')
  }

  const { state } = editor
  const exactFields = new Map<string, { readonly node: PmNode; readonly instruction: string }>()
  state.doc.descendants((node, position) => {
    if (!node.isText) return
    const mark = node.marks.find((candidate) => candidate.type.name === 'instrField')
    if (!mark) return
    exactFields.set(`${position}:${position + node.nodeSize}`, {
      node,
      instruction: String(mark.attrs.instr),
    })
  })

  let totalCharacters = 0
  const seen = new Set<string>()
  const validated: Array<{
    readonly from: number
    readonly to: number
    readonly displayText: string
    readonly node: PmNode
  }> = []
  for (const update of input.updates) {
    const from = update?.range?.from
    const to = update?.range?.to
    if (!Number.isInteger(from) || !Number.isInteger(to) || from < 1 || from >= to) {
      return invalid(`DOCX field update range ${String(from)}..${String(to)} is invalid.`)
    }
    const key = `${from}:${to}`
    if (seen.has(key)) return invalid(`DOCX field update range ${from}..${to} is duplicated.`)
    seen.add(key)
    if (
      typeof update.instruction !== 'string' ||
      unicodeLength(update.instruction) < 1 ||
      unicodeLength(update.instruction) > 512 ||
      !/^[A-Za-z]+(?:\s.*)?$/.test(update.instruction)
    ) {
      return invalid('DOCX field update instructions must be bounded Word field codes.')
    }
    const keyword = update.instruction.trim().split(/\s+/)[0]?.toUpperCase()
    if (!(docxUpdatableFieldKeywordValues as readonly string[]).includes(keyword ?? '')) {
      return invalid(`Unsupported DOCX field update instruction ${update.instruction}.`)
    }
    const displayLength =
      typeof update.displayText === 'string' ? unicodeLength(update.displayText) : 0
    if (displayLength < 1 || displayLength > 65_536) {
      return invalid('DOCX field display text must contain 1 through 65536 characters.')
    }
    totalCharacters += displayLength
    if (totalCharacters > 65_536) {
      return invalid('DOCX field updates exceed the 65536 character aggregate limit.')
    }
    const candidate = exactFields.get(key)
    if (!candidate || candidate.instruction !== update.instruction) {
      return invalid(
        `DOCX field update range ${from}..${to} does not exactly match ${update.instruction}.`,
      )
    }
    validated.push({ from, to, displayText: update.displayText, node: candidate.node })
  }

  try {
    let transaction = state.tr
    let changed = 0
    for (const update of validated.sort((left, right) => right.from - left.from)) {
      if (update.node.text === update.displayText) continue
      transaction = transaction.replaceWith(
        update.from,
        update.to,
        state.schema.text(update.displayText, [...update.node.marks]),
      )
      changed += 1
    }
    if (changed > 0) editor.view.dispatch(closeHistory(transaction))
    return { ok: true, matched: validated.length, changed }
  } catch (error) {
    return {
      ok: false,
      error: 'execution_failed',
      message: `The mounted DOCX editor rejected the field update: ${
        error instanceof Error ? error.message : String(error)
      }`,
    }
  }
}
