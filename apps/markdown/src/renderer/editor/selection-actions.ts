import type { Editor } from '@tiptap/core'

export type MarkdownSelectionResult =
  | { readonly ok: true; readonly from: number; readonly to: number }
  | { readonly ok: false; readonly message: string }

export function setMarkdownSelection(
  editor: Editor,
  input: { readonly from: number; readonly to: number },
): MarkdownSelectionResult {
  if (input.from > input.to) {
    return { ok: false, message: 'Markdown selection requires from <= to.' }
  }
  const maximum = editor.state.doc.content.size
  if (input.from < 1 || input.to > maximum) {
    return {
      ok: false,
      message: `Markdown selection must stay within document positions 1..${maximum}.`,
    }
  }
  if (!editor.commands.setTextSelection({ from: input.from, to: input.to })) {
    return { ok: false, message: 'The Markdown text selection could not be applied.' }
  }
  return { ok: true, from: input.from, to: input.to }
}
