import type { Editor } from '@tiptap/core'
import { generateIndexFieldXml } from '@genoffice/docx-engine'
import { closeHistory } from '@tiptap/pm/history'
import { Fragment } from '@tiptap/pm/model'

export interface MarkDocxIndexEntryInput {
  readonly range: { readonly from: number; readonly to: number }
  readonly term: string
}

export type MarkDocxIndexEntryResult =
  | {
      readonly ok: true
      readonly from: number
      readonly to: number
      readonly term: string
      readonly changed: true
    }
  | {
      readonly ok: false
      readonly error: 'invalid_arguments' | 'execution_failed'
      readonly message: string
    }

export interface InsertDocxIndexInput {
  readonly afterBlockIndex: number
  readonly label: string
  readonly terms: readonly string[]
}

export type InsertDocxIndexResult =
  | {
      readonly ok: true
      readonly afterBlockIndex: number
      readonly entries: number
      readonly insertedBlocks: number
    }
  | {
      readonly ok: false
      readonly error: 'invalid_arguments' | 'execution_failed'
      readonly message: string
    }

function unicodeLength(value: unknown): number {
  return typeof value === 'string' ? Array.from(value).length : -1
}

/** Attach one hidden XE marker to the end of an exact inline source range. */
export function markDocxIndexEntry(
  editor: Editor,
  input: MarkDocxIndexEntryInput,
): MarkDocxIndexEntryResult {
  const { from, to } = input?.range ?? {}
  if (
    !Number.isInteger(from) ||
    !Number.isInteger(to) ||
    from < 1 ||
    to < from ||
    to > editor.state.doc.content.size ||
    unicodeLength(input?.term) < 1 ||
    unicodeLength(input.term) > 4096 ||
    !/^[^"\u0000-\u001F]+$/u.test(input.term)
  ) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message:
        'DOCX index marking requires one exact inline range and a bounded round-trippable term.',
    }
  }
  const $from = editor.state.doc.resolve(from)
  const $to = editor.state.doc.resolve(to)
  if (!$from.sameParent($to) || !$from.parent.inlineContent) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: 'DOCX index marking requires one inline text block.',
    }
  }

  try {
    const marker = editor.schema.nodes.docXeMark.create({ term: input.term })
    editor.view.dispatch(closeHistory(editor.state.tr.insert(to, marker)))
    return { ok: true, from, to, term: input.term, changed: true }
  } catch (error) {
    return {
      ok: false,
      error: 'execution_failed',
      message: `The mounted DOCX editor rejected index marking: ${
        error instanceof Error ? error.message : String(error)
      }`,
    }
  }
}

/** Insert one explicit cached INDEX field after a stable top-level block boundary. */
export function insertDocxIndex(
  editor: Editor,
  input: InsertDocxIndexInput,
): InsertDocxIndexResult {
  const terms = input?.terms
  if (
    !Number.isInteger(input?.afterBlockIndex) ||
    input.afterBlockIndex < -1 ||
    input.afterBlockIndex >= editor.state.doc.childCount ||
    unicodeLength(input.label) < 1 ||
    unicodeLength(input.label) > 255 ||
    !Array.isArray(terms) ||
    terms.length < 1 ||
    terms.length > 1024 ||
    terms.some((term) => unicodeLength(term) < 1 || unicodeLength(term) > 4096) ||
    unicodeLength(JSON.stringify(terms)) > 65_536
  ) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message:
        'DOCX index insertion requires one stable block boundary, a bounded label, and bounded terms.',
    }
  }
  const normalizedTerms = [...new Set(terms.map((term) => term.trim()).filter(Boolean))].sort(
    (left, right) => left.localeCompare(right, 'zh-CN'),
  )
  if (normalizedTerms.length === 0) {
    return { ok: false, error: 'invalid_arguments', message: 'DOCX index terms must not be blank.' }
  }

  try {
    let position = 0
    for (let index = 0; index <= input.afterBlockIndex; index += 1) {
      position += editor.state.doc.child(index).nodeSize
    }
    const nodes = generateIndexFieldXml(normalizedTerms).map((xml, index) =>
      editor.schema.nodes.docProtected.create({
        docxIndex: null,
        blockType: 'passthrough',
        label: input.label,
        genXml: xml,
        fieldDisplay: {
          kind: 'tocLine',
          left: normalizedTerms[index],
          right: '',
          level: 1,
        },
      }),
    )
    editor.view.dispatch(
      closeHistory(editor.state.tr.insert(position, Fragment.fromArray(nodes))),
    )
    return {
      ok: true,
      afterBlockIndex: input.afterBlockIndex,
      entries: nodes.length,
      insertedBlocks: nodes.length,
    }
  } catch (error) {
    return {
      ok: false,
      error: 'execution_failed',
      message: `The mounted DOCX editor rejected index insertion: ${
        error instanceof Error ? error.message : String(error)
      }`,
    }
  }
}
