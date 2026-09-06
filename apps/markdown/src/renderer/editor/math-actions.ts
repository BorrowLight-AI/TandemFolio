import type { Editor } from '@tiptap/core'

export type MarkdownMathDisplay = 'inline' | 'block'

export type MarkdownMathResult =
  { readonly ok: true } | { readonly ok: false; readonly message: string }

export function insertMarkdownMath(
  editor: Editor,
  input: {
    readonly position: number
    readonly display: MarkdownMathDisplay
    readonly latex: string
  },
): MarkdownMathResult {
  const latex = input.latex.trim()
  if (!latex) return { ok: false, message: 'Markdown math LaTeX must not be empty.' }
  if (input.position < 1 || input.position > editor.state.doc.content.size) {
    return {
      ok: false,
      message: `Markdown math position ${input.position} is invalid for document size ${editor.state.doc.content.size}.`,
    }
  }
  const inserted =
    input.display === 'block'
      ? editor.commands.insertBlockMath({ latex, pos: input.position })
      : editor.commands.insertInlineMath({ latex, pos: input.position })
  return inserted
    ? { ok: true }
    : { ok: false, message: `Markdown ${input.display} math could not be inserted.` }
}

export function setMarkdownMath(
  editor: Editor,
  input: { readonly position: number; readonly latex: string | null },
): MarkdownMathResult {
  const node = editor.state.doc.nodeAt(input.position)
  if (!node || (node.type.name !== 'blockMath' && node.type.name !== 'inlineMath')) {
    return {
      ok: false,
      message: `Markdown math position ${input.position} does not address a formula.`,
    }
  }
  if (input.latex === null) {
    return editor.commands.deleteRange({ from: input.position, to: input.position + node.nodeSize })
      ? { ok: true }
      : { ok: false, message: `Markdown math at position ${input.position} could not be deleted.` }
  }
  const latex = input.latex.trim()
  if (!latex) return { ok: false, message: 'Markdown math LaTeX must not be empty.' }
  const updated =
    node.type.name === 'blockMath'
      ? editor.commands.updateBlockMath({ latex, pos: input.position })
      : editor.commands.updateInlineMath({ latex, pos: input.position })
  return updated
    ? { ok: true }
    : { ok: false, message: `Markdown math at position ${input.position} could not be updated.` }
}
