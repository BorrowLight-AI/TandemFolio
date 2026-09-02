import type { Editor, JSONContent } from '@tiptap/core'

const PAGE_BREAK_PARAGRAPH: JSONContent = {
  type: 'docParagraph',
  attrs: { pageBreakBefore: true },
}

export interface InsertPageBreakInput {
  readonly afterBlockIndex: number
}

export type InsertPageBreakResult =
  | { readonly ok: true; readonly insertedBlockIndex: number }
  | { readonly ok: false; readonly error: 'invalid_arguments'; readonly message: string }

function insertNativePageBreak(
  editor: Editor,
  target: { readonly kind: 'selection' } | { readonly kind: 'position'; readonly pos: number },
): boolean {
  if (target.kind === 'selection') {
    return editor.chain().focus().insertContent(PAGE_BREAK_PARAGRAPH).run()
  }

  const paragraph = editor.state.schema.nodeFromJSON(PAGE_BREAK_PARAGRAPH)
  editor.view.dispatch(editor.state.tr.insert(target.pos, paragraph))
  return true
}

/** Preserve the retained Ribbon command while sharing the native write kernel. */
export function insertPageBreakAt(editor: Editor): void {
  insertNativePageBreak(editor, { kind: 'selection' })
}

/** Insert a page-break paragraph after one explicit top-level block boundary. */
export function insertPageBreakAfterBlock(
  editor: Editor,
  input: InsertPageBreakInput,
): InsertPageBreakResult {
  const { doc } = editor.state
  if (input.afterBlockIndex < -1 || input.afterBlockIndex >= doc.childCount) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: `DOCX page-break boundary after block ${input.afterBlockIndex} is invalid for ${doc.childCount} block(s).`,
    }
  }

  let pos = 0
  for (let index = 0; index <= input.afterBlockIndex; index += 1) {
    pos += doc.child(index).nodeSize
  }
  insertNativePageBreak(editor, { kind: 'position', pos })
  return { ok: true, insertedBlockIndex: input.afterBlockIndex + 1 }
}
