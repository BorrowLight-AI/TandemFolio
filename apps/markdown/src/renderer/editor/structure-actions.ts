import type { Editor } from '@tiptap/core'

export type MarkdownStructureResult =
  | { readonly ok: true; readonly changed?: boolean }
  | { readonly ok: false; readonly message: string }

function validateInsertionPosition(editor: Editor, position: number): MarkdownStructureResult {
  const documentSize = editor.state.doc.content.size
  if (!Number.isInteger(position) || position < 1 || position > documentSize) {
    return {
      ok: false,
      message: `Markdown insertion position ${position} is invalid for document size ${documentSize}.`,
    }
  }
  return { ok: true }
}

export function insertMarkdownTable(
  editor: Editor,
  input: {
    readonly position: number
    readonly rows: number
    readonly columns: number
    readonly headerRow: boolean
  },
): MarkdownStructureResult {
  const valid = validateInsertionPosition(editor, input.position)
  if (!valid.ok) return valid
  const inserted = editor
    .chain()
    .focus()
    .setTextSelection(input.position)
    .insertTable({ rows: input.rows, cols: input.columns, withHeaderRow: input.headerRow })
    .run()
  return inserted
    ? { ok: true }
    : { ok: false, message: `Markdown table cannot be inserted at position ${input.position}.` }
}

export function insertMarkdownDivider(
  editor: Editor,
  input: { readonly position: number },
): MarkdownStructureResult {
  const valid = validateInsertionPosition(editor, input.position)
  if (!valid.ok) return valid
  const inserted = editor
    .chain()
    .focus()
    .setTextSelection(input.position)
    .setHorizontalRule()
    .run()
  return inserted
    ? { ok: true }
    : { ok: false, message: `Markdown divider cannot be inserted at position ${input.position}.` }
}

export const markdownTableUpdateActions = [
  'add_row_before',
  'add_row_after',
  'delete_row',
  'add_column_before',
  'add_column_after',
  'delete_column',
  'set_header_row',
  'delete_table',
] as const

export type MarkdownTableUpdateAction = (typeof markdownTableUpdateActions)[number]

function tableAtPosition(editor: Editor, position: number) {
  if (position < 1 || position > editor.state.doc.content.size) return null
  const resolved = editor.state.doc.resolve(position)
  for (let depth = resolved.depth; depth > 0; depth -= 1) {
    const node = resolved.node(depth)
    if (node.type.name === 'table') return node
  }
  return null
}

export function markdownTableHasHeaderRow(editor: Editor, position: number): boolean | null {
  const table = tableAtPosition(editor, position)
  if (!table) return null
  return table.firstChild?.firstChild?.type.name === 'tableHeader'
}

/** Apply one bounded table-relative mutation at an explicit cell position. */
export function updateMarkdownTable(
  editor: Editor,
  input: {
    readonly position: number
    readonly action: MarkdownTableUpdateAction
    readonly headerRow: boolean | null
  },
): MarkdownStructureResult {
  const table = tableAtPosition(editor, input.position)
  if (!table) {
    return {
      ok: false,
      message: `Markdown table update position ${input.position} is not inside a table.`,
    }
  }
  if (
    (input.action === 'set_header_row' && input.headerRow === null) ||
    (input.action !== 'set_header_row' && input.headerRow !== null)
  ) {
    return {
      ok: false,
      message:
        'Markdown table headerRow must be boolean only for set_header_row and null otherwise.',
    }
  }

  editor.commands.setTextSelection(input.position)
  let chain = editor.chain().focus()
  if (input.action === 'add_row_before') chain = chain.addRowBefore()
  else if (input.action === 'add_row_after') chain = chain.addRowAfter()
  else if (input.action === 'delete_row') chain = chain.deleteRow()
  else if (input.action === 'add_column_before') chain = chain.addColumnBefore()
  else if (input.action === 'add_column_after') chain = chain.addColumnAfter()
  else if (input.action === 'delete_column') chain = chain.deleteColumn()
  else if (input.action === 'delete_table') chain = chain.deleteTable()
  else {
    const enabled = markdownTableHasHeaderRow(editor, input.position)
    if (enabled === input.headerRow) return { ok: true, changed: false }
    chain = chain.toggleHeaderRow()
  }
  if (!chain.run()) {
    return {
      ok: false,
      message: `Markdown table action ${input.action} cannot run at position ${input.position}.`,
    }
  }
  return { ok: true, changed: true }
}
