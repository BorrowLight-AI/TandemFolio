import type { Editor } from '@tiptap/core'
import { patchMathTokens, ommlToMathML, type FormulaDisplay } from '@genoffice/docx-engine'
import { closeHistory } from '@tiptap/pm/history'
import type { EditorState, Transaction } from '@tiptap/pm/state'
import { equationBlockJson, inlineEquationNodeJson } from './equation'

export interface InsertDocxEquationInput {
  readonly placement: 'block' | 'inline'
  readonly latex: string
  readonly afterBlockIndex: number | null
  readonly from: number | null
  readonly to: number | null
}

export type InsertDocxEquationResult =
  | {
      readonly ok: true
      readonly equationBlockIndex: number | null
      readonly from: number | null
      readonly to: number | null
    }
  | {
      readonly ok: false
      readonly error: 'invalid_arguments' | 'execution_failed'
      readonly message: string
    }

export interface UpdateDocxEquationInput {
  readonly placement: 'block' | 'inline'
  readonly mode: 'latex' | 'tokens'
  readonly latex: string | null
  readonly tokens: readonly string[] | null
  readonly equationBlockIndex: number | null
  readonly from: number | null
  readonly to: number | null
}

export type UpdateDocxEquationResult =
  | { readonly ok: true; readonly changed: boolean }
  | {
      readonly ok: false
      readonly error: 'invalid_arguments' | 'execution_failed'
      readonly message: string
    }

type DispatchTransaction = (transaction: Transaction) => void

type EquationUpdate = Pick<UpdateDocxEquationInput, 'placement' | 'mode' | 'latex' | 'tokens'>

function invalid(message: string): InsertDocxEquationResult {
  return { ok: false, error: 'invalid_arguments', message }
}

function invalidUpdate(message: string): UpdateDocxEquationResult {
  return { ok: false, error: 'invalid_arguments', message }
}

function validateUpdate(input: EquationUpdate): string | null {
  if (input.placement !== 'block' && input.placement !== 'inline') {
    return `Unsupported DOCX equation placement: ${String(input.placement)}.`
  }
  if (input.mode !== 'latex' && input.mode !== 'tokens') {
    return `Unsupported DOCX equation update mode: ${String(input.mode)}.`
  }
  if (input.mode === 'latex') {
    if (typeof input.latex !== 'string' || input.tokens !== null) {
      return 'LaTeX DOCX equation updates require latex and null tokens.'
    }
    const latex = input.latex.trim()
    if (Array.from(latex).length < 1 || Array.from(latex).length > 4096) {
      return 'DOCX equation LaTeX must contain 1 through 4096 characters after trimming.'
    }
    return null
  }
  if (input.placement !== 'block') {
    return 'DOCX equation token updates are supported only for retained block equations.'
  }
  if (input.latex !== null || !Array.isArray(input.tokens)) {
    return 'Token DOCX equation updates require null latex and an explicit tokens array.'
  }
  if (input.tokens.length < 1 || input.tokens.length > 1024) {
    return 'DOCX equation tokens must contain 1 through 1024 items.'
  }
  let totalCharacters = 0
  for (const token of input.tokens) {
    if (typeof token !== 'string' || Array.from(token).length > 4096) {
      return 'Each DOCX equation token must be a string of at most 4096 characters.'
    }
    totalCharacters += Array.from(token).length
  }
  return totalCharacters <= 4096
    ? null
    : 'DOCX equation token content is limited to 4096 total characters.'
}

/** Shared exact-node equation writer for the retained modal/token editor and Registry. */
export function updateDocxEquationAtPosition(
  state: EditorState,
  dispatch: DispatchTransaction,
  position: number,
  input: EquationUpdate,
): UpdateDocxEquationResult {
  const invalid = validateUpdate(input)
  if (invalid) return invalidUpdate(invalid)
  const node = state.doc.nodeAt(position)
  if (
    !node ||
    (input.placement === 'block'
      ? node.type.name !== 'docProtected' || !node.attrs.formulaDisplay
      : node.type.name !== 'docInlineMath')
  ) {
    return invalidUpdate(`The DOCX ${input.placement} equation target is not editable.`)
  }

  try {
    if (input.mode === 'latex') {
      const latex = input.latex!.trim()
      if (input.placement === 'block') {
        const next = state.schema.nodeFromJSON(equationBlockJson(latex))
        const changed = JSON.stringify(node.toJSON()) !== JSON.stringify(next.toJSON())
        if (changed) {
          dispatch(closeHistory(state.tr.replaceWith(position, position + node.nodeSize, next)))
        }
        return { ok: true, changed }
      }
      const next = inlineEquationNodeJson(latex)
      const changed = JSON.stringify(node.attrs) !== JSON.stringify(next.attrs)
      if (changed) {
        dispatch(closeHistory(state.tr.setNodeMarkup(position, undefined, next.attrs)))
      }
      return { ok: true, changed }
    }

    const current = node.attrs.formulaDisplay as FormulaDisplay
    if (current.tokens.length !== input.tokens!.length) {
      return invalidUpdate(
        `DOCX equation token updates must preserve the current ${current.tokens.length}-token shape.`,
      )
    }
    const tokens = [...input.tokens!]
    const omml = current.omml ? patchMathTokens(current.omml, tokens) : undefined
    const nextFormula: FormulaDisplay = omml
      ? { tokens, omml, mathml: ommlToMathML(omml) }
      : { tokens }
    const attrs: Record<string, unknown> = { ...node.attrs, formulaDisplay: nextFormula }
    if (node.attrs.genXml) attrs.genXml = patchMathTokens(String(node.attrs.genXml), tokens)
    const changed = JSON.stringify(node.attrs) !== JSON.stringify(attrs)
    if (changed) dispatch(closeHistory(state.tr.setNodeMarkup(position, undefined, attrs)))
    return { ok: true, changed }
  } catch (error) {
    return invalidUpdate(
      `DOCX equation update is invalid: ${error instanceof Error ? error.message : String(error)}`,
    )
  }
}

/** Resolve one stable block identity or exact inline node range and update it. */
export function updateDocxEquation(
  editor: Editor,
  input: UpdateDocxEquationInput,
): UpdateDocxEquationResult {
  const { doc } = editor.state
  let position: number
  if (input.placement === 'block') {
    if (
      input.from !== null ||
      input.to !== null ||
      input.equationBlockIndex === null ||
      !Number.isInteger(input.equationBlockIndex) ||
      input.equationBlockIndex < 0 ||
      input.equationBlockIndex >= doc.childCount
    ) {
      return invalidUpdate(
        'Block DOCX equation updates require a valid equationBlockIndex and null from/to coordinates.',
      )
    }
    position = 0
    for (let index = 0; index < input.equationBlockIndex; index += 1) {
      position += doc.child(index).nodeSize
    }
  } else {
    if (
      input.equationBlockIndex !== null ||
      input.from === null ||
      input.to === null ||
      !Number.isInteger(input.from) ||
      !Number.isInteger(input.to) ||
      input.from < 1 ||
      input.to > doc.content.size
    ) {
      return invalidUpdate(
        'Inline DOCX equation updates require null equationBlockIndex and an explicit valid from/to range.',
      )
    }
    const node = doc.nodeAt(input.from)
    if (!node || input.to !== input.from + node.nodeSize) {
      return invalidUpdate('Inline DOCX equation from/to must identify one exact node.')
    }
    position = input.from
  }
  return updateDocxEquationAtPosition(
    editor.state,
    (transaction) => editor.view.dispatch(transaction),
    position,
    input,
  )
}

/** Insert one block or inline equation through the mounted editor's native history. */
export function insertDocxEquation(
  editor: Editor,
  input: InsertDocxEquationInput,
): InsertDocxEquationResult {
  const latex = input.latex.trim()
  if (Array.from(latex).length < 1 || Array.from(latex).length > 4096) {
    return invalid('DOCX equation LaTeX must contain 1 through 4096 characters after trimming.')
  }

  if (input.placement === 'block') {
    if (input.afterBlockIndex === null || input.from !== null || input.to !== null) {
      return invalid('Block DOCX equations require afterBlockIndex and null from/to coordinates.')
    }
    const { doc } = editor.state
    if (input.afterBlockIndex < -1 || input.afterBlockIndex >= doc.childCount) {
      return invalid(
        `DOCX equation boundary after block ${input.afterBlockIndex} is invalid for ${doc.childCount} block(s).`,
      )
    }
    let position = 0
    for (let index = 0; index <= input.afterBlockIndex; index += 1) {
      position += doc.child(index).nodeSize
    }
    try {
      const inserted = editor
        .chain()
        .focus()
        .insertContentAt(position, equationBlockJson(latex))
        .run()
      return inserted
        ? {
            ok: true,
            equationBlockIndex: input.afterBlockIndex + 1,
            from: null,
            to: null,
          }
        : {
            ok: false,
            error: 'execution_failed',
            message: 'The mounted DOCX editor rejected block equation insertion.',
          }
    } catch (error) {
      return invalid(
        `DOCX equation LaTeX is invalid: ${error instanceof Error ? error.message : String(error)}`,
      )
    }
  }

  if (input.afterBlockIndex !== null || input.from === null || input.to === null) {
    return invalid(
      'Inline DOCX equations require null afterBlockIndex and explicit from/to coordinates.',
    )
  }
  if (input.from > input.to) {
    return invalid('Inline DOCX equation from must be less than or equal to to.')
  }
  const { doc } = editor.state
  if (input.from < 1 || input.to > doc.content.size) {
    return invalid(
      `Inline DOCX equation range ${input.from}..${input.to} is outside document size ${doc.content.size}.`,
    )
  }
  const $from = doc.resolve(input.from)
  const $to = doc.resolve(input.to)
  if ($from.parent !== $to.parent || !$from.parent.inlineContent) {
    return invalid('Inline DOCX equation ranges must stay inside one text-bearing block.')
  }
  try {
    const inserted = editor
      .chain()
      .focus()
      .insertContentAt({ from: input.from, to: input.to }, inlineEquationNodeJson(latex))
      .run()
    return inserted
      ? {
          ok: true,
          equationBlockIndex: null,
          from: input.from,
          to: input.to,
        }
      : {
          ok: false,
          error: 'execution_failed',
          message: 'The mounted DOCX editor rejected inline equation insertion.',
        }
  } catch (error) {
    return invalid(
      `DOCX equation LaTeX is invalid: ${error instanceof Error ? error.message : String(error)}`,
    )
  }
}

/** Retained Gallery/modal block producer converging on the registry action. */
export function insertEquationFromLatex(editor: Editor, latex: string): void {
  const afterBlockIndex = editor.state.selection.$from.index(0)
  insertDocxEquation(editor, {
    placement: 'block',
    latex,
    afterBlockIndex,
    from: null,
    to: null,
  })
}

/** Retained modal inline producer converging on the registry action. */
export function insertInlineEquationFromLatex(editor: Editor, latex: string): void {
  const { from, to } = editor.state.selection
  insertDocxEquation(editor, {
    placement: 'inline',
    latex,
    afterBlockIndex: null,
    from,
    to,
  })
}
