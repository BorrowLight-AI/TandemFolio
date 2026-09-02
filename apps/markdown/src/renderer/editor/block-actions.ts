import type { Editor } from '@tiptap/core'
import { TextSelection, type Transaction } from '@tiptap/pm/state'
import type { Node as ProseMirrorNode } from '@tiptap/pm/model'

export const markdownBlockUpdateActions = ['duplicate', 'delete', 'add_below', 'move'] as const
export type MarkdownBlockUpdateAction = (typeof markdownBlockUpdateActions)[number]

export type MarkdownBlockUpdateResult =
  | { readonly ok: true; readonly changed: boolean }
  | { readonly ok: false; readonly message: string }

function blockPosition(nodes: readonly ProseMirrorNode[], index: number): number {
  let position = 0
  for (let current = 0; current < index; current += 1) position += nodes[current]!.nodeSize
  return position
}

export function getMarkdownTopLevelBlockIndexAtSelection(editor: Editor): number | null {
  const selection = editor.state.selection
  if (selection.$from.depth === 0) return selection.$from.index(0)
  return selection.$from.index(0)
}

/** Apply one explicit top-level block mutation in one native transaction. */
export function updateMarkdownBlock(
  editor: Editor,
  input: {
    readonly blockIndex: number
    readonly action: MarkdownBlockUpdateAction
    readonly afterBlockIndex: number | null
    readonly content: string | null
  },
  dispatch: (transaction: Transaction) => void = (transaction) => editor.view.dispatch(transaction),
  transactionBase: Transaction = editor.state.tr,
): MarkdownBlockUpdateResult {
  const moveFields = input.action === 'move' && input.afterBlockIndex !== null && input.content === null
  const addFields =
    input.action === 'add_below' && input.afterBlockIndex === null && input.content !== null
  const simpleFields =
    (input.action === 'duplicate' || input.action === 'delete') &&
    input.afterBlockIndex === null &&
    input.content === null
  if (!moveFields && !addFields && !simpleFields) {
    return { ok: false, message: 'Markdown block fields must match the selected action.' }
  }

  const doc = editor.state.doc
  if (
    !Number.isInteger(input.blockIndex) ||
    input.blockIndex < 0 ||
    input.blockIndex >= doc.childCount
  ) {
    return {
      ok: false,
      message: `Markdown block ${input.blockIndex} is invalid for ${doc.childCount} top-level block(s).`,
    }
  }
  const nodes = Array.from({ length: doc.childCount }, (_, index) => doc.child(index))
  const target = nodes[input.blockIndex]!

  if (input.action === 'duplicate') {
    const position = blockPosition(nodes, input.blockIndex + 1)
    let transaction = transactionBase.insert(position, target)
    transaction = transaction.setSelection(TextSelection.near(transaction.doc.resolve(position + 1)))
    dispatch(transaction.scrollIntoView())
    return { ok: true, changed: true }
  }

  if (input.action === 'delete') {
    const from = blockPosition(nodes, input.blockIndex)
    const to = from + target.nodeSize
    let transaction = transactionBase.delete(from, to)
    if (transaction.doc.childCount === 0) {
      transaction = transaction.insert(0, editor.schema.nodes.paragraph!.create())
    }
    transaction = transaction.setSelection(
      TextSelection.near(transaction.doc.resolve(Math.min(from + 1, transaction.doc.content.size))),
    )
    dispatch(transaction.scrollIntoView())
    return { ok: true, changed: true }
  }

  if (input.action === 'add_below') {
    const content = input.content ? editor.schema.text(input.content) : undefined
    const paragraph = editor.schema.nodes.paragraph!.create(null, content)
    const position = blockPosition(nodes, input.blockIndex + 1)
    let transaction = transactionBase.insert(position, paragraph)
    transaction = transaction.setSelection(TextSelection.near(transaction.doc.resolve(position + 1)))
    dispatch(transaction.scrollIntoView())
    return { ok: true, changed: true }
  }

  const afterBlockIndex = input.afterBlockIndex!
  if (afterBlockIndex < -1 || afterBlockIndex >= doc.childCount) {
    return {
      ok: false,
      message: `Markdown move boundary ${afterBlockIndex} is invalid for ${doc.childCount} top-level block(s).`,
    }
  }
  if (afterBlockIndex === input.blockIndex) return { ok: true, changed: false }
  const boundary = afterBlockIndex === -1 ? null : nodes[afterBlockIndex]!
  const reordered = nodes.filter((_node, index) => index !== input.blockIndex)
  const insertionIndex = boundary === null ? 0 : reordered.indexOf(boundary) + 1
  reordered.splice(insertionIndex, 0, target)
  if (reordered.every((node, index) => node === nodes[index])) return { ok: true, changed: false }

  const movedPosition = blockPosition(reordered, insertionIndex)
  let transaction = transactionBase.replaceWith(0, doc.content.size, reordered)
  transaction = transaction.setSelection(
    TextSelection.near(transaction.doc.resolve(movedPosition + 1)),
  )
  dispatch(transaction.scrollIntoView())
  return { ok: true, changed: true }
}
