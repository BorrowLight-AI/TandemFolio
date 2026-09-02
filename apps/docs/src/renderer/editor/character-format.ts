import type { Editor } from '@tiptap/core'
import { TextSelection } from '@tiptap/pm/state'
import { isEastAsianFontName } from '../font-list'

export const docxCharacterFormatFields = [
  'bold',
  'italic',
  'underline',
  'strike',
  'verticalAlign',
  'color',
  'highlight',
  'fontFamily',
  'fontSizePoints',
] as const

export type DocxCharacterFormatField = (typeof docxCharacterFormatFields)[number]

export interface DocxCharacterFormat {
  readonly bold?: boolean
  readonly italic?: boolean
  readonly underline?: boolean
  readonly strike?: boolean
  readonly verticalAlign?: 'baseline' | 'superscript' | 'subscript'
  readonly color?: string | null
  readonly highlight?:
    | 'yellow'
    | 'green'
    | 'cyan'
    | 'magenta'
    | 'blue'
    | 'red'
    | 'darkBlue'
    | 'darkCyan'
    | 'darkGreen'
    | 'darkMagenta'
    | 'darkRed'
    | 'darkYellow'
    | 'darkGray'
    | 'lightGray'
    | 'black'
    | 'white'
    | null
  readonly fontFamily?: string | null
  readonly fontSizePoints?: number | null
}

export interface DocxCharacterFormatInput {
  readonly range: { readonly from: number; readonly to: number }
  readonly format: DocxCharacterFormat
  readonly fields: readonly DocxCharacterFormatField[]
}

export type DocxCharacterFormatResult =
  | { readonly ok: true; readonly changed: boolean }
  | { readonly ok: false; readonly message: string }

export interface DocxCharacterRangeInput {
  readonly range: { readonly from: number; readonly to: number }
}

export type DocxTextCaseMode = 'sentence' | 'lower' | 'upper' | 'title'

export interface DocxTextCaseInput extends DocxCharacterRangeInput {
  readonly mode: DocxTextCaseMode
}

export interface DocxCharacterStyleInput extends DocxCharacterRangeInput {
  readonly styleId: string | null
}

const BOOLEAN_MARKS = {
  bold: 'bold',
  italic: 'italic',
  underline: 'underline',
  strike: 'strike',
} as const

function isBooleanMarkField(field: DocxCharacterFormatField): field is keyof typeof BOOLEAN_MARKS {
  return Object.hasOwn(BOOLEAN_MARKS, field)
}

function hasExactFieldMask(input: DocxCharacterFormatInput): boolean {
  const fields = [...input.fields]
  if (new Set(fields).size !== fields.length) return false
  const fieldSet = new Set<string>(fields)
  const keys = Object.keys(input.format)
  return keys.length === fields.length && keys.every((key) => fieldSet.has(key))
}

function validateBoundedFormat(format: DocxCharacterFormat): string | null {
  if (
    format.color !== undefined &&
    format.color !== null &&
    !/^#[0-9A-Fa-f]{6}$/.test(format.color)
  ) {
    return 'docx.text.set_character_format color must be null or #RRGGBB.'
  }
  if (format.fontFamily !== undefined && format.fontFamily !== null) {
    const font = format.fontFamily.trim()
    if (!font || font.length > 128 || /[\u0000-\u001f\u007f]/.test(font)) {
      return 'docx.text.set_character_format fontFamily must contain 1 through 128 visible characters.'
    }
  }
  if (format.fontSizePoints !== undefined && format.fontSizePoints !== null) {
    if (
      !Number.isFinite(format.fontSizePoints) ||
      format.fontSizePoints < 1 ||
      format.fontSizePoints > 1638 ||
      !Number.isInteger(format.fontSizePoints * 2)
    ) {
      return 'docx.text.set_character_format fontSizePoints must be 1 through 1638 in 0.5-point steps.'
    }
  }
  return null
}

function validateTextRange(
  editor: Editor,
  range: DocxCharacterRangeInput['range'],
  operation: string,
): string | null {
  const { from, to } = range
  const { doc } = editor.state
  if (
    from >= to ||
    to > doc.content.size ||
    !doc.resolve(from).parent.inlineContent ||
    !doc.resolve(to).parent.inlineContent
  ) {
    return `${operation} requires a non-empty text range inside the document.`
  }
  let containsText = false
  doc.nodesBetween(from, to, (node, pos) => {
    if (node.isText && Math.max(from, pos) < Math.min(to, pos + node.nodeSize)) containsText = true
  })
  return containsText ? null : `${operation} requires a non-empty text range inside the document.`
}

export function setDocxCharacterFormat(
  editor: Editor,
  input: DocxCharacterFormatInput,
): DocxCharacterFormatResult {
  if (!hasExactFieldMask(input)) {
    return {
      ok: false,
      message: 'docx.text.set_character_format fields must uniquely and exactly match format.',
    }
  }
  const formatError = validateBoundedFormat(input.format)
  if (formatError) return { ok: false, message: formatError }

  const { from, to } = input.range
  const { state } = editor
  const rangeError = validateTextRange(editor, input.range, 'docx.text.set_character_format')
  if (rangeError) return { ok: false, message: rangeError }

  const tr = state.tr.setSelection(TextSelection.create(state.doc, from, to))
  for (const field of input.fields) {
    if (!isBooleanMarkField(field)) continue
    const mark = state.schema.marks[BOOLEAN_MARKS[field]]
    if (input.format[field]) tr.addMark(from, to, mark.create())
    else tr.removeMark(from, to, mark)
  }

  const stylePatch: Record<string, unknown> = {}
  if (input.fields.includes('verticalAlign')) {
    stylePatch.vertAlign =
      input.format.verticalAlign === 'baseline' ? null : (input.format.verticalAlign ?? null)
  }
  if (input.fields.includes('color')) {
    stylePatch.color =
      input.format.color === null ? null : input.format.color?.slice(1).toUpperCase()
  }
  if (input.fields.includes('highlight')) stylePatch.highlight = input.format.highlight ?? null
  if (input.fields.includes('fontSizePoints')) {
    stylePatch.sizeHalfPoints =
      input.format.fontSizePoints === null ? null : input.format.fontSizePoints! * 2
  }
  if (input.fields.includes('fontFamily')) {
    const font = input.format.fontFamily?.trim() || null
    if (font === null) Object.assign(stylePatch, { font: null, fontAscii: null })
    else if (isEastAsianFontName(font)) stylePatch.font = font
    else stylePatch.fontAscii = font
  }

  if (Object.keys(stylePatch).length > 0) {
    const type = state.schema.marks.docTextStyle
    const jobs: Array<{
      readonly from: number
      readonly to: number
      readonly attrs: Readonly<Record<string, unknown>> | null
    }> = []
    tr.doc.nodesBetween(from, to, (node, pos) => {
      if (!node.isText) return
      const start = Math.max(from, pos)
      const end = Math.min(to, pos + node.nodeSize)
      if (start >= end) return
      const existing = node.marks.find((mark) => mark.type === type)?.attrs ?? {}
      const attrs = { ...existing, ...stylePatch }
      jobs.push({
        from: start,
        to: end,
        attrs: Object.values(attrs).every((value) => value === null) ? null : attrs,
      })
    })
    for (const job of jobs) {
      tr.removeMark(job.from, job.to, type)
      if (job.attrs) tr.addMark(job.from, job.to, type.create(job.attrs))
    }
  }

  const changed = !tr.doc.eq(state.doc)
  editor.view.dispatch(tr)
  return { ok: true, changed }
}

export function clearDocxCharacterFormat(
  editor: Editor,
  input: DocxCharacterRangeInput,
): DocxCharacterFormatResult {
  const rangeError = validateTextRange(editor, input.range, 'docx.text.clear_character_format')
  if (rangeError) return { ok: false, message: rangeError }

  const { from, to } = input.range
  const { state } = editor
  const tr = state.tr.setSelection(TextSelection.create(state.doc, from, to)).removeMark(from, to)
  const changed = !tr.doc.eq(state.doc)
  editor.view.dispatch(tr)
  return { ok: true, changed }
}

export function transformDocxCaseText(value: string, mode: DocxTextCaseMode): string {
  switch (mode) {
    case 'upper':
      return value.toUpperCase()
    case 'lower':
      return value.toLowerCase()
    case 'title':
      return value.toLowerCase().replace(/(^|\s)(\p{L})/gu, (match) => match.toUpperCase())
    case 'sentence':
      return value
        .toLowerCase()
        .replace(/(^\s*\p{L})|([.!?。!?]\s*\p{L})/gu, (match) => match.toUpperCase())
  }
}

export function transformDocxTextCase(
  editor: Editor,
  input: DocxTextCaseInput,
): DocxCharacterFormatResult {
  const rangeError = validateTextRange(editor, input.range, 'docx.text.transform_case')
  if (rangeError) return { ok: false, message: rangeError }

  const { from, to } = input.range
  const { state } = editor
  const tr = state.tr.setSelection(TextSelection.create(state.doc, from, to))
  state.doc.nodesBetween(from, to, (node, pos) => {
    if (!node.isText || !node.text) return
    const start = Math.max(from, pos)
    const end = Math.min(to, pos + node.nodeSize)
    const slice = node.text.slice(start - pos, end - pos)
    const next = transformDocxCaseText(slice, input.mode)
    if (next !== slice) {
      tr.replaceWith(
        tr.mapping.map(start),
        tr.mapping.map(end),
        state.schema.text(next, node.marks),
      )
    }
  })
  const changed = !tr.doc.eq(state.doc)
  editor.view.dispatch(tr)
  return { ok: true, changed }
}

export function setDocxCharacterStyle(
  editor: Editor,
  input: DocxCharacterStyleInput,
): DocxCharacterFormatResult {
  const rangeError = validateTextRange(editor, input.range, 'docx.text.set_character_style')
  if (rangeError) return { ok: false, message: rangeError }

  const { from, to } = input.range
  const { state } = editor
  const type = state.schema.marks.docTextStyle
  const tr = state.tr.setSelection(TextSelection.create(state.doc, from, to))
  if (input.styleId === null) {
    tr.removeMark(from, to, type)
  } else {
    const jobs: Array<{
      readonly from: number
      readonly to: number
      readonly attrs: Readonly<Record<string, unknown>>
    }> = []
    tr.doc.nodesBetween(from, to, (node, pos) => {
      if (!node.isText) return
      const start = Math.max(from, pos)
      const end = Math.min(to, pos + node.nodeSize)
      if (start >= end) return
      const existing = node.marks.find((mark) => mark.type === type)?.attrs ?? {}
      jobs.push({ from: start, to: end, attrs: { ...existing, styleId: input.styleId } })
    })
    for (const job of jobs) {
      tr.removeMark(job.from, job.to, type)
      tr.addMark(job.from, job.to, type.create(job.attrs))
    }
  }
  const changed = !tr.doc.eq(state.doc)
  editor.view.dispatch(tr)
  return { ok: true, changed }
}
