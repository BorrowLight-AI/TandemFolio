import type { Editor, JSONContent } from '@tiptap/core'
import { buildTextboxParagraphXml, type TextboxDisplay } from '@genoffice/docx-engine'
import { closeHistory } from '@tiptap/pm/history'
import type { EditorState, Transaction } from '@tiptap/pm/state'

const EMU_PER_PX = 9_525

export const DEFAULT_DOCX_TEXTBOX_WIDTH_EMU = 1_800_000
export const DEFAULT_DOCX_TEXTBOX_HEIGHT_EMU = 1_080_000

export interface InsertDocxTextboxInput {
  readonly afterBlockIndex: number
  readonly widthEmu: number
  readonly heightEmu: number
}

export type InsertDocxTextboxResult =
  | { readonly ok: true; readonly textboxBlockIndex: number }
  | {
      readonly ok: false
      readonly error: 'invalid_arguments' | 'execution_failed'
      readonly message: string
    }

export interface DocxTextboxRunInput {
  readonly text: string
  readonly bold?: boolean
  readonly italic?: boolean
  readonly underline?: boolean
  readonly strike?: boolean
  readonly color?: string
  readonly sizeHalfPoints?: number
  readonly font?: string
  readonly fontAscii?: string
  readonly charSpacingTwips?: number
  readonly highlight?: string
  readonly shading?: string
  readonly vertAlign?: 'superscript' | 'subscript'
  readonly styleId?: string
}

export interface DocxTextboxParagraphInput {
  readonly runs: readonly DocxTextboxRunInput[]
  readonly align?: 'left' | 'center' | 'right' | 'justify' | 'distribute'
  readonly lineSpacing?: number
  readonly indentLeft?: number
  readonly indentRight?: number
  readonly indentFirstLine?: number
  readonly spaceBefore?: number
  readonly spaceAfter?: number
  readonly shadingFill?: string
  readonly borders?: string
}

export interface SetDocxTextboxContentInput {
  readonly objectBlockIndex: number
  readonly textboxIndex: number
  readonly paragraphs: readonly DocxTextboxParagraphInput[]
  /** Explicit final fixed height; null preserves the current fixed/autofit state. */
  readonly heightPx: number | null
}

export type SetDocxTextboxContentResult =
  | { readonly ok: true; readonly changed: boolean; readonly heightPx: number | null }
  | {
      readonly ok: false
      readonly error: 'invalid_arguments' | 'execution_failed'
      readonly message: string
    }

type SetDocxTextboxContentFailure = Exclude<SetDocxTextboxContentResult, { readonly ok: true }>

type DispatchTransaction = (transaction: Transaction) => void

function textboxContent(widthEmu: number, heightEmu: number, label: string): JSONContent {
  const xml = buildTextboxParagraphXml({
    widthEmu,
    heightEmu,
    id: Math.floor(Math.random() * 900_000) + 100_000,
  })
  const textbox: TextboxDisplay = {
    fill: 'FFFFFF',
    borderColor: '000000',
    widthPx: Math.round(widthEmu / EMU_PER_PX),
    heightPx: Math.round(heightEmu / EMU_PER_PX),
    paras: [{ runs: [{ text: '' }] }],
  }
  return {
    type: 'docProtected',
    attrs: {
      docxIndex: null,
      blockType: 'passthrough',
      label,
      genXml: xml,
      textboxes: [textbox],
    },
  }
}

function invalidContent(message: string): SetDocxTextboxContentFailure {
  return { ok: false, error: 'invalid_arguments', message }
}

const HIGHLIGHTS = new Set([
  'yellow',
  'green',
  'cyan',
  'magenta',
  'blue',
  'red',
  'darkBlue',
  'darkCyan',
  'darkGreen',
  'darkMagenta',
  'darkRed',
  'darkYellow',
  'darkGray',
  'lightGray',
  'black',
  'white',
])

function boundedText(value: unknown, maximum: number): value is string {
  return typeof value === 'string' && Array.from(value).length <= maximum
}

function optionalHex(value: unknown): boolean {
  return value === undefined || (typeof value === 'string' && /^[0-9A-F]{6}$/.test(value))
}

function validateParagraphs(paragraphs: readonly DocxTextboxParagraphInput[]): string | null {
  if (!Array.isArray(paragraphs) || paragraphs.length < 1 || paragraphs.length > 256) {
    return 'DOCX textbox content must contain 1 through 256 paragraphs.'
  }
  let totalRuns = 0
  let totalCharacters = 0
  for (const paragraph of paragraphs) {
    if (!paragraph || !Array.isArray(paragraph.runs)) {
      return 'Every DOCX textbox paragraph requires a runs array.'
    }
    if (paragraph.runs.length < 1 || paragraph.runs.length > 1024) {
      return 'Every DOCX textbox paragraph must contain 1 through 1024 runs.'
    }
    totalRuns += paragraph.runs.length
    if (
      paragraph.align !== undefined &&
      !['left', 'center', 'right', 'justify', 'distribute'].includes(paragraph.align)
    ) {
      return `Unsupported DOCX textbox paragraph alignment: ${String(paragraph.align)}.`
    }
    if (
      paragraph.lineSpacing !== undefined &&
      (!Number.isFinite(paragraph.lineSpacing) ||
        paragraph.lineSpacing < 0.06 ||
        paragraph.lineSpacing > 132)
    ) {
      return 'DOCX textbox paragraph lineSpacing must be from 0.06 through 132.'
    }
    for (const [name, value, minimum, maximum] of [
      ['indentLeft', paragraph.indentLeft, -31_680, 31_680],
      ['indentRight', paragraph.indentRight, -31_680, 31_680],
      ['indentFirstLine', paragraph.indentFirstLine, -31_680, 31_680],
      ['spaceBefore', paragraph.spaceBefore, 0, 31_680],
      ['spaceAfter', paragraph.spaceAfter, 0, 31_680],
    ] as const) {
      if (value !== undefined && (!Number.isFinite(value) || value < minimum || value > maximum)) {
        return `DOCX textbox paragraph ${name} is out of range.`
      }
    }
    if (!optionalHex(paragraph.shadingFill)) {
      return 'DOCX textbox paragraph shadingFill must be uppercase six-digit hex.'
    }
    if (
      paragraph.borders !== undefined &&
      (!/^[tblr]{1,4}$/.test(paragraph.borders) ||
        new Set(paragraph.borders).size !== paragraph.borders.length)
    ) {
      return 'DOCX textbox paragraph borders must contain unique t/b/l/r edge letters.'
    }
    for (const run of paragraph.runs) {
      if (!run || !boundedText(run.text, 4096)) {
        return 'Every DOCX textbox run requires text of at most 4096 Unicode characters.'
      }
      totalCharacters += Array.from(run.text).length
      for (const key of ['bold', 'italic', 'underline', 'strike'] as const) {
        if (run[key] !== undefined && typeof run[key] !== 'boolean') {
          return `DOCX textbox run ${key} must be boolean.`
        }
      }
      if (!optionalHex(run.color) || !optionalHex(run.shading)) {
        return 'DOCX textbox run color and shading must be uppercase six-digit hex.'
      }
      if (
        run.sizeHalfPoints !== undefined &&
        (!Number.isInteger(run.sizeHalfPoints) ||
          run.sizeHalfPoints < 2 ||
          run.sizeHalfPoints > 3276)
      ) {
        return 'DOCX textbox run sizeHalfPoints must be an integer from 2 through 3276.'
      }
      if (
        (run.font !== undefined && !boundedText(run.font, 256)) ||
        (run.fontAscii !== undefined && !boundedText(run.fontAscii, 256))
      ) {
        return 'DOCX textbox run font names must contain at most 256 Unicode characters.'
      }
      if (
        run.charSpacingTwips !== undefined &&
        (!Number.isInteger(run.charSpacingTwips) ||
          run.charSpacingTwips < -31_680 ||
          run.charSpacingTwips > 31_680)
      ) {
        return 'DOCX textbox run charSpacingTwips is out of range.'
      }
      if (run.highlight !== undefined && !HIGHLIGHTS.has(run.highlight)) {
        return `Unsupported DOCX textbox run highlight: ${run.highlight}.`
      }
      if (
        run.vertAlign !== undefined &&
        run.vertAlign !== 'superscript' &&
        run.vertAlign !== 'subscript'
      ) {
        return `Unsupported DOCX textbox run vertAlign: ${String(run.vertAlign)}.`
      }
      if (run.styleId !== undefined && !boundedText(run.styleId, 256)) {
        return 'DOCX textbox run styleId must contain at most 256 Unicode characters.'
      }
    }
  }
  if (totalRuns > 4096) return 'DOCX textbox content is limited to 4096 total runs.'
  if (totalCharacters > 65_536) {
    return 'DOCX textbox content is limited to 65536 total Unicode characters.'
  }
  return null
}

type PositionedTextboxContentInput = Pick<
  SetDocxTextboxContentInput,
  'textboxIndex' | 'paragraphs' | 'heightPx'
>

function prepareTextboxContent(
  textboxes: readonly TextboxDisplay[],
  input: PositionedTextboxContentInput,
):
  | { readonly ok: true; readonly box: TextboxDisplay; readonly heightPx: number | null }
  | SetDocxTextboxContentFailure {
  if (
    !Number.isInteger(input.textboxIndex) ||
    input.textboxIndex < 0 ||
    input.textboxIndex >= textboxes.length
  ) {
    return invalidContent(
      `DOCX textbox index ${input.textboxIndex} is invalid for ${textboxes.length} textbox(es).`,
    )
  }
  const box = textboxes[input.textboxIndex]
  if (box.readOnly) {
    return invalidContent(
      'The DOCX textbox target is read-only because its structure is flattened.',
    )
  }
  const invalid = validateParagraphs(input.paragraphs)
  if (invalid) return invalidContent(invalid)
  if (
    input.heightPx !== null &&
    (!Number.isInteger(input.heightPx) || input.heightPx < 8 || input.heightPx > 4096)
  ) {
    return invalidContent('DOCX textbox heightPx must be null or an integer from 8 through 4096.')
  }

  const paragraphs = input.paragraphs.map((paragraph) => ({
    ...paragraph,
    runs: paragraph.runs.map((run) => ({ ...run })),
  })) as TextboxDisplay['paras']
  const heightPx = input.heightPx ?? box.heightPx ?? null
  const nextBox: TextboxDisplay = {
    ...box,
    paras: paragraphs,
    ...(input.heightPx === null ? {} : { heightPx: input.heightPx }),
  }
  return { ok: true, box: nextBox, heightPx }
}

/** Batch-capable shared writer so one UI commit remains one native Undo unit. */
export function setDocxTextboxContentsAtPosition(
  state: EditorState,
  dispatch: DispatchTransaction,
  position: number,
  inputs: readonly PositionedTextboxContentInput[],
): SetDocxTextboxContentResult {
  const node = state.doc.nodeAt(position)
  const textboxes = node?.attrs.textboxes as TextboxDisplay[] | null | undefined
  if (!node || node.type.name !== 'docProtected' || !Array.isArray(textboxes)) {
    return invalidContent('The DOCX object target has no editable textbox content.')
  }
  if (inputs.length < 1 || inputs.length > textboxes.length) {
    return invalidContent('DOCX textbox updates require one unique entry per changed textbox.')
  }
  const seen = new Set<number>()
  const next = [...textboxes]
  let heightPx: number | null = null
  for (const input of inputs) {
    if (seen.has(input.textboxIndex)) {
      return invalidContent(`DOCX textbox index ${input.textboxIndex} must not be repeated.`)
    }
    seen.add(input.textboxIndex)
    const prepared = prepareTextboxContent(next, input)
    if (!prepared.ok) return prepared
    next[input.textboxIndex] = prepared.box
    heightPx = prepared.heightPx
  }
  const changed = JSON.stringify(next) !== JSON.stringify(textboxes)
  if (changed) {
    dispatch(
      closeHistory(state.tr.setNodeMarkup(position, undefined, { ...node.attrs, textboxes: next })),
    )
  }
  return { ok: true, changed, heightPx }
}

/** Shared single nested-textbox writer for Registry. */
export function setDocxTextboxContentAtPosition(
  state: EditorState,
  dispatch: DispatchTransaction,
  position: number,
  input: PositionedTextboxContentInput,
): SetDocxTextboxContentResult {
  return setDocxTextboxContentsAtPosition(state, dispatch, position, [input])
}

/** Resolve one stable top-level object plus nested textbox identity and update its content. */
export function setDocxTextboxContentAtBlock(
  editor: Editor,
  input: SetDocxTextboxContentInput,
): SetDocxTextboxContentResult {
  const { doc } = editor.state
  if (
    !Number.isInteger(input.objectBlockIndex) ||
    input.objectBlockIndex < 0 ||
    input.objectBlockIndex >= doc.childCount
  ) {
    return invalidContent(
      `DOCX object block ${input.objectBlockIndex} is invalid for ${doc.childCount} block(s).`,
    )
  }
  let position = 0
  for (let index = 0; index < input.objectBlockIndex; index += 1) {
    position += doc.child(index).nodeSize
  }
  return setDocxTextboxContentAtPosition(
    editor.state,
    (transaction) => editor.view.dispatch(transaction),
    position,
    input,
  )
}

/** Shared insertion primitive for retained UI and registry textbox producers. */
export function insertDocxTextboxAtPosition(
  editor: Editor,
  input: Omit<InsertDocxTextboxInput, 'afterBlockIndex'> & { readonly position: number },
  label: string,
): number | null {
  return editor
    .chain()
    .focus()
    .insertContentAt(input.position, textboxContent(input.widthEmu, input.heightEmu, label))
    .run()
    ? input.position
    : null
}

/** Insert a textbox after one stable revision-scoped top-level block. */
export function insertDocxTextboxAfterBlock(
  editor: Editor,
  input: InsertDocxTextboxInput,
): InsertDocxTextboxResult {
  const { doc } = editor.state
  if (input.afterBlockIndex < -1 || input.afterBlockIndex >= doc.childCount) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: `DOCX textbox boundary after block ${input.afterBlockIndex} is invalid for ${doc.childCount} block(s).`,
    }
  }
  let position = 0
  for (let index = 0; index <= input.afterBlockIndex; index += 1) {
    position += doc.child(index).nodeSize
  }
  const inserted = insertDocxTextboxAtPosition(editor, { ...input, position }, 'Text box')
  if (inserted === null) {
    return {
      ok: false,
      error: 'execution_failed',
      message: 'The mounted DOCX editor rejected textbox insertion.',
    }
  }
  return { ok: true, textboxBlockIndex: input.afterBlockIndex + 1 }
}
