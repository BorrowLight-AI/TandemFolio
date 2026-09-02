import type { Editor } from '@tiptap/core'
import { closeHistory } from '@tiptap/pm/history'

export interface SetDocxDropCapInput {
  readonly blockIndex: number
  readonly mode: 'none' | 'drop' | 'margin'
  readonly lines: number | null
}

export type SetDocxDropCapResult =
  | {
      readonly ok: true
      readonly blockIndex: number
      readonly mode: SetDocxDropCapInput['mode']
      readonly lines: number | null
      readonly changed: boolean
    }
  | {
      readonly ok: false
      readonly error: 'invalid_arguments' | 'execution_failed'
      readonly message: string
    }

const DROP_CAP_BLOCKS = new Set(['docParagraph', 'docHeading', 'docListItem'])

export function activeTopLevelBlockIndex(editor: Editor): number {
  return Math.min(editor.state.selection.$from.index(0), editor.state.doc.childCount - 1)
}

/** Set one top-level paragraph-like block to an explicit drop-cap final state. */
export function setDocxDropCap(
  editor: Editor,
  input: SetDocxDropCapInput,
): SetDocxDropCapResult {
  const { doc } = editor.state
  if (
    !input ||
    !Number.isInteger(input.blockIndex) ||
    input.blockIndex < 0 ||
    input.blockIndex >= doc.childCount ||
    !['none', 'drop', 'margin'].includes(input.mode) ||
    (input.mode === 'none'
      ? input.lines !== null
      : !Number.isInteger(input.lines) || input.lines! < 2 || input.lines! > 10)
  ) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message:
        'DOCX drop caps require a valid block, none with null lines, or drop/margin with 2–10 lines.',
    }
  }
  const target = doc.child(input.blockIndex)
  if (!DROP_CAP_BLOCKS.has(target.type.name)) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: `DOCX drop-cap block ${input.blockIndex} must be a paragraph, heading, or list item.`,
    }
  }
  const next =
    input.mode === 'none' ? null : JSON.stringify({ type: input.mode, lines: input.lines })
  if ((target.attrs.dropCap ?? null) === next) {
    return { ok: true, blockIndex: input.blockIndex, mode: input.mode, lines: input.lines, changed: false }
  }
  let position = 0
  for (let index = 0; index < input.blockIndex; index += 1) position += doc.child(index).nodeSize
  try {
    editor.view.dispatch(
      closeHistory(
        editor.state.tr.setNodeMarkup(position, undefined, { ...target.attrs, dropCap: next }),
      ),
    )
    return { ok: true, blockIndex: input.blockIndex, mode: input.mode, lines: input.lines, changed: true }
  } catch (error) {
    return {
      ok: false,
      error: 'execution_failed',
      message: `The mounted DOCX editor rejected the drop-cap update: ${
        error instanceof Error ? error.message : String(error)
      }`,
    }
  }
}
