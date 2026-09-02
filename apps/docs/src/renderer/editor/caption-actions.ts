import type { Editor } from '@tiptap/core'
import { generateCaptionXml } from '@genoffice/docx-engine'
import { closeHistory } from '@tiptap/pm/history'

export interface InsertDocxCaptionInput {
  readonly afterBlockIndex: number
  readonly label: string
  readonly number: number
  readonly text: string
}

export type InsertDocxCaptionResult =
  | {
      readonly ok: true
      readonly afterBlockIndex: number
      readonly label: string
      readonly number: number
      readonly changed: true
    }
  | {
      readonly ok: false
      readonly error: 'invalid_arguments' | 'execution_failed'
      readonly message: string
    }

function unicodeLength(value: unknown): number {
  return typeof value === 'string' ? Array.from(value).length : -1
}

/** Insert one explicit dirty SEQ caption after a stable top-level block boundary. */
export function insertDocxCaption(
  editor: Editor,
  input: InsertDocxCaptionInput,
): InsertDocxCaptionResult {
  if (
    !Number.isInteger(input?.afterBlockIndex) ||
    input.afterBlockIndex < -1 ||
    input.afterBlockIndex >= editor.state.doc.childCount ||
    unicodeLength(input.label) < 1 ||
    unicodeLength(input.label) > 255 ||
    !Number.isInteger(input.number) ||
    input.number < 1 ||
    input.number > 2_147_483_647 ||
    unicodeLength(input.text) < 0 ||
    unicodeLength(input.text) > 4096
  ) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message:
        'DOCX caption insertion requires one stable block boundary, bounded label/text, and an explicit positive number.',
    }
  }

  try {
    let position = 0
    for (let index = 0; index <= input.afterBlockIndex; index += 1) {
      position += editor.state.doc.child(index).nodeSize
    }
    const display = `${input.label} ${input.number}${input.text ? ` ${input.text}` : ''}`
    const caption = editor.schema.nodes.docProtected.create({
      docxIndex: null,
      blockType: 'passthrough',
      label: input.label,
      genXml: generateCaptionXml(input.label, input.number, input.text),
      fieldDisplay: { kind: 'text', left: display },
    })
    editor.view.dispatch(closeHistory(editor.state.tr.insert(position, caption)))
    return {
      ok: true,
      afterBlockIndex: input.afterBlockIndex,
      label: input.label,
      number: input.number,
      changed: true,
    }
  } catch (error) {
    return {
      ok: false,
      error: 'execution_failed',
      message: `The mounted DOCX editor rejected caption insertion: ${
        error instanceof Error ? error.message : String(error)
      }`,
    }
  }
}
