import type { Editor, JSONContent } from '@tiptap/core'
import {
  buildLineParagraphXml,
  buildShapeParagraphXml,
  type TextboxDisplay,
} from '@genoffice/docx-engine'
import { isStraightLineKind } from './shape-svg'

const EMU_PER_PX = 9525
const DEFAULT_FILL = '4472C4'
const DEFAULT_BORDER = '2F5496'
export const DOCX_STRAIGHT_LINE_HEIGHT_EMU = 114_300

export interface InsertDocxShapeInput {
  readonly afterBlockIndex: number
  readonly preset: string
  readonly widthEmu: number
  readonly heightEmu: number
}

export type InsertDocxShapeResult =
  | { readonly ok: true; readonly shapeBlockIndex: number }
  | {
      readonly ok: false
      readonly error: 'invalid_arguments' | 'execution_failed'
      readonly message: string
    }

export interface InsertDocxLineInput {
  readonly afterBlockIndex: number
  readonly kind: string
  readonly widthEmu: number
  readonly heightEmu: number
}

export type InsertDocxLineResult =
  | { readonly ok: true; readonly lineBlockIndex: number }
  | {
      readonly ok: false
      readonly error: 'invalid_arguments' | 'execution_failed'
      readonly message: string
    }

function shapeContent(
  preset: string,
  widthEmu: number,
  heightEmu: number,
  label: string,
): JSONContent {
  const xml = buildShapeParagraphXml({
    prst: preset,
    widthEmu,
    heightEmu,
    id: Math.floor(Math.random() * 900_000) + 100_000,
    fillHex: DEFAULT_FILL,
    borderHex: DEFAULT_BORDER,
    withTextbox: true,
  })
  const textbox: TextboxDisplay = {
    fill: DEFAULT_FILL,
    borderColor: DEFAULT_BORDER,
    widthPx: Math.round(widthEmu / EMU_PER_PX),
    heightPx: Math.round(heightEmu / EMU_PER_PX),
    prst: preset,
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

function lineContent(
  kind: string,
  widthEmu: number,
  heightEmu: number,
  label: string,
): JSONContent {
  const canonicalHeight = isStraightLineKind(kind) ? DOCX_STRAIGHT_LINE_HEIGHT_EMU : heightEmu
  const xml = buildLineParagraphXml({
    kind,
    widthEmu,
    heightEmu: canonicalHeight,
    id: Math.floor(Math.random() * 900_000) + 100_000,
    colorHex: '000000',
  })
  const textbox: TextboxDisplay = {
    borderColor: '000000',
    widthPx: Math.round(widthEmu / EMU_PER_PX),
    heightPx: Math.round(canonicalHeight / EMU_PER_PX),
    prst: kind,
    paras: [],
    readOnly: true,
    insetTopPx: 0,
    insetRightPx: 0,
    insetBottomPx: 0,
    insetLeftPx: 0,
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

/** Shared insertion primitive for the retained draw/gallery routes. */
export function insertDocxShapeAtPosition(
  editor: Editor,
  input: Omit<InsertDocxShapeInput, 'afterBlockIndex'> & { readonly position: number },
  label: string,
): number | null {
  const content = shapeContent(input.preset, input.widthEmu, input.heightEmu, label)
  return editor.chain().focus().insertContentAt(input.position, content).run()
    ? input.position
    : null
}

/** Shared stroke-only insertion primitive for line Gallery/draw routes. */
export function insertDocxLineAtPosition(
  editor: Editor,
  input: Omit<InsertDocxLineInput, 'afterBlockIndex'> & { readonly position: number },
  label: string,
): number | null {
  const content = lineContent(input.kind, input.widthEmu, input.heightEmu, label)
  return editor.chain().focus().insertContentAt(input.position, content).run()
    ? input.position
    : null
}

/** Insert a preset shape after one stable revision-scoped top-level block. */
export function insertDocxShapeAfterBlock(
  editor: Editor,
  input: InsertDocxShapeInput,
): InsertDocxShapeResult {
  const { doc } = editor.state
  if (input.afterBlockIndex < -1 || input.afterBlockIndex >= doc.childCount) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: `DOCX shape boundary after block ${input.afterBlockIndex} is invalid for ${doc.childCount} block(s).`,
    }
  }
  let position = 0
  for (let index = 0; index <= input.afterBlockIndex; index += 1) {
    position += doc.child(index).nodeSize
  }
  const inserted = insertDocxShapeAtPosition(
    editor,
    { ...input, position },
    `Shape(${input.preset})`,
  )
  if (inserted === null) {
    return {
      ok: false,
      error: 'execution_failed',
      message: 'The mounted DOCX editor rejected shape insertion.',
    }
  }
  return { ok: true, shapeBlockIndex: input.afterBlockIndex + 1 }
}

/** Insert a retained line/connector after one stable top-level block. */
export function insertDocxLineAfterBlock(
  editor: Editor,
  input: InsertDocxLineInput,
): InsertDocxLineResult {
  const { doc } = editor.state
  if (input.afterBlockIndex < -1 || input.afterBlockIndex >= doc.childCount) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: `DOCX line boundary after block ${input.afterBlockIndex} is invalid for ${doc.childCount} block(s).`,
    }
  }
  if (isStraightLineKind(input.kind) && input.heightEmu !== DOCX_STRAIGHT_LINE_HEIGHT_EMU) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: `Straight DOCX lines require heightEmu ${DOCX_STRAIGHT_LINE_HEIGHT_EMU}.`,
    }
  }
  let position = 0
  for (let index = 0; index <= input.afterBlockIndex; index += 1) {
    position += doc.child(index).nodeSize
  }
  const inserted = insertDocxLineAtPosition(editor, { ...input, position }, `Line(${input.kind})`)
  if (inserted === null) {
    return {
      ok: false,
      error: 'execution_failed',
      message: 'The mounted DOCX editor rejected line insertion.',
    }
  }
  return { ok: true, lineBlockIndex: input.afterBlockIndex + 1 }
}
