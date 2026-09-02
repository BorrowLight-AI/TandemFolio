import type { Editor } from '@tiptap/core'

export interface MarkdownTextMarks {
  readonly bold: boolean
  readonly italic: boolean
  readonly strike: boolean
  readonly code: boolean
  readonly link: string | null
}

export type MarkdownTextMarkResult =
  { readonly ok: true } | { readonly ok: false; readonly message: string }

export function getMarkdownTextMarks(editor: Editor): MarkdownTextMarks {
  const href = editor.getAttributes('link').href
  return {
    bold: editor.isActive('bold'),
    italic: editor.isActive('italic'),
    strike: editor.isActive('strike'),
    code: editor.isActive('code'),
    link: editor.isActive('link') && typeof href === 'string' ? href : null,
  }
}

/** Apply one complete, replayable inline-mark state to a revision-scoped PM range. */
export function setMarkdownTextMarks(
  editor: Editor,
  input: { readonly from: number; readonly to: number; readonly marks: MarkdownTextMarks },
): MarkdownTextMarkResult {
  const documentSize = editor.state.doc.content.size
  const isAllSelection = input.from === 0 && input.to === documentSize
  const from = isAllSelection ? 1 : input.from
  const to = isAllSelection ? Math.max(1, documentSize - 1) : input.to
  if (
    !Number.isInteger(from) ||
    !Number.isInteger(to) ||
    from < 1 ||
    to < from ||
    to > documentSize
  ) {
    return {
      ok: false,
      message: `Markdown mark range ${input.from}..${input.to} is invalid for document size ${documentSize}.`,
    }
  }
  if (
    input.marks.code &&
    (input.marks.bold || input.marks.italic || input.marks.strike || input.marks.link !== null)
  ) {
    return {
      ok: false,
      message: 'Markdown inline code cannot be combined with bold, italic, strike, or link.',
    }
  }

  let chain = editor.chain().focus().setTextSelection({ from, to })
  chain = input.marks.bold ? chain.setMark('bold') : chain.unsetMark('bold')
  chain = input.marks.italic ? chain.setMark('italic') : chain.unsetMark('italic')
  chain = input.marks.strike ? chain.setMark('strike') : chain.unsetMark('strike')
  chain = input.marks.code ? chain.setMark('code') : chain.unsetMark('code')
  chain = input.marks.link ? chain.setLink({ href: input.marks.link }) : chain.unsetMark('link')
  chain.run()
  return { ok: true }
}
