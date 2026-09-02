import type { Editor } from '@tiptap/core'

export type MarkdownHistoryAction = 'undo' | 'redo'

export function runMarkdownHistoryAction(
  editor: Editor,
  action: MarkdownHistoryAction,
): boolean {
  return editor.chain().focus()[action]().run()
}
