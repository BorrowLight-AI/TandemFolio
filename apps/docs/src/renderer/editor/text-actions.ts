import type { Editor } from '@tiptap/core'
import { closeHistory } from '@tiptap/pm/history'

export interface InsertDocxTextInput {
  readonly text: string
}

export type InsertDocxTextResult =
  | { readonly ok: true; readonly changed: true }
  | {
      readonly ok: false
      readonly error: 'invalid_arguments' | 'execution_failed'
      readonly message: string
    }

/** Insert bounded text at the mounted editor's active selection through native history. */
export function insertDocxText(editor: Editor, input: InsertDocxTextInput): InsertDocxTextResult {
  const length = typeof input?.text === 'string' ? Array.from(input.text).length : 0
  if (length < 1 || length > 65_536) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: 'DOCX text insertion requires 1 through 65536 Unicode characters.',
    }
  }

  try {
    const { from, to } = editor.state.selection
    const transaction = closeHistory(editor.state.tr.insertText(input.text, from, to))
    editor.view.dispatch(transaction)
    return { ok: true, changed: true }
  } catch (error) {
    return {
      ok: false,
      error: 'execution_failed',
      message: `The mounted DOCX editor rejected text insertion: ${
        error instanceof Error ? error.message : String(error)
      }`,
    }
  }
}
