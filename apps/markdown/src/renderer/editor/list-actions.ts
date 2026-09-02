import type { Editor } from '@tiptap/core'

import { resolveMarkdownTextBlockPosition } from './block-type-actions'

export const markdownListTypes = ['none', 'bullet', 'ordered', 'task'] as const
export type MarkdownListType = (typeof markdownListTypes)[number]

export type MarkdownListResult =
  | { readonly ok: true; readonly changed: boolean }
  | { readonly ok: false; readonly message: string }

export function getActiveMarkdownListType(editor: Editor): MarkdownListType {
  if (editor.isActive('taskList')) return 'task'
  if (editor.isActive('bulletList')) return 'bullet'
  if (editor.isActive('orderedList')) return 'ordered'
  return 'none'
}

/** Set the list context containing one revision-scoped text block to an explicit final type. */
export function setMarkdownListType(
  editor: Editor,
  input: { readonly textBlockIndex: number; readonly type: MarkdownListType },
): MarkdownListResult {
  const address = resolveMarkdownTextBlockPosition(editor, input.textBlockIndex)
  if (!address.ok) return address
  editor.commands.setTextSelection(Math.min(address.position + 1, editor.state.doc.content.size))

  const currentType = getActiveMarkdownListType(editor)
  if (currentType === input.type) return { ok: true, changed: false }

  let chain = editor.chain().focus()
  if (input.type === 'none') {
    const itemNames: Array<'listItem' | 'taskItem'> = []
    const resolved = editor.state.selection.$from
    for (let depth = resolved.depth; depth > 0; depth -= 1) {
      const name = resolved.node(depth).type.name
      if (name === 'listItem' || name === 'taskItem') itemNames.push(name)
    }
    if (itemNames.length === 0) return { ok: true, changed: false }
    for (const itemName of itemNames) chain = chain.liftListItem(itemName)
  } else if (input.type === 'bullet') {
    chain = chain.toggleBulletList()
  } else if (input.type === 'ordered') {
    chain = chain.toggleOrderedList()
  } else {
    chain = chain.toggleTaskList()
  }
  if (!chain.run()) {
    return {
      ok: false,
      message: `Markdown text block ${input.textBlockIndex} cannot be converted to list type ${input.type}.`,
    }
  }
  return { ok: true, changed: true }
}
