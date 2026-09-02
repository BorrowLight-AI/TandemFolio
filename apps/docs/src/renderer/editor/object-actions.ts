import type { Editor } from '@tiptap/core'
import { NodeSelection, type EditorState, type Transaction } from '@tiptap/pm/state'
import { closeHistory } from '@tiptap/pm/history'
import type { ChartDisplay, ImageWrap, TextboxDisplay } from '@genoffice/docx-engine'
import { isStraightLineKind } from './shape-svg'

export type DocxResizableObjectKind = 'shape' | 'line' | 'textbox' | 'chart'
export type DocxPositionableObjectKind = Exclude<DocxResizableObjectKind, 'chart'> | 'diagram'
export type DocxStyleableObjectKind = Exclude<DocxPositionableObjectKind, 'chart'>
export type DocxObjectStyleField = 'fillHex' | 'borderHex'
export type DocxRemovableObjectKind = DocxStyleableObjectKind | 'chart' | 'equation'

export interface SetDocxObjectSizeInput {
  readonly objectBlockIndex: number
  readonly widthPx: number
  readonly heightPx: number
}

export type SetDocxObjectSizeResult =
  | {
      readonly ok: true
      readonly objectKind: DocxResizableObjectKind
      readonly changed: boolean
    }
  | {
      readonly ok: false
      readonly error: 'invalid_arguments' | 'execution_failed'
      readonly message: string
    }

export interface SetDocxObjectOffsetInput {
  readonly objectBlockIndex: number
  readonly wrap: Exclude<ImageWrap, null>
  readonly offsetXEmu: number
  readonly offsetYEmu: number
}

export type SetDocxObjectOffsetResult =
  | {
      readonly ok: true
      readonly objectKind: DocxPositionableObjectKind
      readonly changed: boolean
    }
  | {
      readonly ok: false
      readonly error: 'invalid_arguments' | 'execution_failed'
      readonly message: string
    }

export interface SetDocxObjectStyleInput {
  readonly objectBlockIndex: number
  readonly style: Readonly<{
    fillHex?: string | null
    borderHex?: string | null
  }>
  readonly fields: readonly DocxObjectStyleField[]
}

export type SetDocxObjectStyleResult =
  | {
      readonly ok: true
      readonly objectKind: DocxStyleableObjectKind
      readonly changed: boolean
    }
  | {
      readonly ok: false
      readonly error: 'invalid_arguments' | 'execution_failed'
      readonly message: string
    }

export interface RemoveDocxObjectInput {
  readonly objectBlockIndex: number
}

export type RemoveDocxObjectResult =
  | {
      readonly ok: true
      readonly objectKind: DocxRemovableObjectKind
      readonly replacementParagraphInserted: boolean
    }
  | {
      readonly ok: false
      readonly error: 'invalid_arguments' | 'execution_failed'
      readonly message: string
    }

type DispatchTransaction = (transaction: Transaction) => void

function invalid(message: string): SetDocxObjectSizeResult {
  return { ok: false, error: 'invalid_arguments', message }
}

function invalidOffset(message: string): SetDocxObjectOffsetResult {
  return { ok: false, error: 'invalid_arguments', message }
}

function invalidStyle(message: string): SetDocxObjectStyleResult {
  return { ok: false, error: 'invalid_arguments', message }
}

function invalidRemoval(message: string): RemoveDocxObjectResult {
  return { ok: false, error: 'invalid_arguments', message }
}

function textboxObjectKind(
  attrs: Record<string, unknown>,
  box: TextboxDisplay,
): DocxResizableObjectKind {
  if (box.readOnly && typeof box.prst === 'string' && box.prst.startsWith('line')) return 'line'
  if (
    typeof box.prst === 'string' ||
    (typeof attrs.label === 'string' && attrs.label.startsWith('Shape('))
  ) {
    return 'shape'
  }
  return 'textbox'
}

function removableObjectKind(attrs: Record<string, unknown>): DocxRemovableObjectKind | null {
  if (attrs.blockType === 'image') return null
  if (attrs.chartDisplay) return 'chart'
  if (attrs.formulaDisplay) return 'equation'
  if (attrs.diagramDisplay) return 'diagram'
  const boxes = attrs.textboxes as TextboxDisplay[] | null
  if (!Array.isArray(boxes) || boxes.length === 0) return null
  return boxes.length === 1 ? textboxObjectKind(attrs, boxes[0]) : 'textbox'
}

/** Shared exact-object deletion kernel for keyboard object mode and Registry. */
export function removeDocxObjectAtPosition(
  state: EditorState,
  dispatch: DispatchTransaction,
  position: number,
): RemoveDocxObjectResult {
  const node = state.doc.nodeAt(position)
  const objectKind =
    node?.type.name === 'docProtected'
      ? removableObjectKind(node.attrs as Record<string, unknown>)
      : null
  if (!node || objectKind === null) {
    return invalidRemoval('The DOCX target is not a removable drawing, chart, or block equation.')
  }

  const replacementParagraphInserted = state.doc.childCount === 1
  let transaction = state.tr
  if (replacementParagraphInserted) {
    const paragraph = state.schema.nodes.docParagraph?.createAndFill()
    if (!paragraph) {
      return {
        ok: false,
        error: 'execution_failed',
        message: 'The DOCX schema could not create a replacement paragraph.',
      }
    }
    transaction = transaction.replaceWith(position, position + node.nodeSize, paragraph)
  } else {
    transaction = transaction.delete(position, position + node.nodeSize)
  }
  dispatch(closeHistory(transaction.scrollIntoView()))
  return { ok: true, objectKind, replacementParagraphInserted }
}

/** Resolve one stable top-level block and remove its exact retained object. */
export function removeDocxObjectAtBlock(
  editor: Editor,
  input: RemoveDocxObjectInput,
): RemoveDocxObjectResult {
  const { doc } = editor.state
  if (
    !Number.isInteger(input.objectBlockIndex) ||
    input.objectBlockIndex < 0 ||
    input.objectBlockIndex >= doc.childCount
  ) {
    return invalidRemoval(
      `DOCX object block ${input.objectBlockIndex} is invalid for ${doc.childCount} block(s).`,
    )
  }
  let position = 0
  for (let index = 0; index < input.objectBlockIndex; index += 1) {
    position += doc.child(index).nodeSize
  }
  const node = doc.child(input.objectBlockIndex)
  if (
    node.type.name !== 'docProtected' ||
    removableObjectKind(node.attrs as Record<string, unknown>) === null
  ) {
    return invalidRemoval(
      `DOCX block ${input.objectBlockIndex} is not a removable drawing, chart, or block equation.`,
    )
  }
  return removeDocxObjectAtPosition(
    editor.state,
    (transaction) => editor.view.dispatch(transaction),
    position,
  )
}

/** Delete only an explicitly selected retained object; let other keymaps handle other nodes. */
export function removeSelectedDocxObject(
  state: EditorState,
  dispatch: DispatchTransaction,
): boolean {
  const { selection } = state
  if (!(selection instanceof NodeSelection) || selection.node.type.name !== 'docProtected') {
    return false
  }
  return removeDocxObjectAtPosition(state, dispatch, selection.from).ok
}

/** Shared final-size kernel for the retained corner handle and Registry. */
export function setDocxObjectSizeAtPosition(
  state: EditorState,
  dispatch: DispatchTransaction,
  position: number,
  input: Pick<SetDocxObjectSizeInput, 'widthPx' | 'heightPx'>,
): SetDocxObjectSizeResult {
  const node = state.doc.nodeAt(position)
  if (!node || node.type.name !== 'docProtected') {
    return invalid('The DOCX object target is not a protected drawing block.')
  }
  const attrs = node.attrs as Record<string, unknown>
  const chart = attrs.chartDisplay as ChartDisplay | null
  if (chart) {
    if (!Number.isInteger(input.widthPx) || input.widthPx < 120 || input.widthPx > 660) {
      return invalid('DOCX chart widthPx must be an integer from 120 through 660.')
    }
    if (!Number.isInteger(input.heightPx) || input.heightPx < 80 || input.heightPx > 4096) {
      return invalid('DOCX chart heightPx must be an integer from 80 through 4096.')
    }
    const changed = chart.widthPx !== input.widthPx || chart.heightPx !== input.heightPx
    if (changed) {
      dispatch(
        closeHistory(
          state.tr.setNodeMarkup(position, undefined, {
            ...attrs,
            chartDisplay: { ...chart, widthPx: input.widthPx, heightPx: input.heightPx },
          }),
        ),
      )
    }
    return { ok: true, objectKind: 'chart', changed }
  }

  const boxes = attrs.textboxes as TextboxDisplay[] | null
  if (!Array.isArray(boxes) || boxes.length !== 1) {
    return invalid('The DOCX object target is not one resizable shape, line, or textbox.')
  }
  if (!Number.isInteger(input.widthPx) || input.widthPx < 24 || input.widthPx > 4096) {
    return invalid('DOCX object widthPx must be an integer from 24 through 4096.')
  }
  if (!Number.isInteger(input.heightPx) || input.heightPx < 8 || input.heightPx > 4096) {
    return invalid('DOCX object heightPx must be an integer from 8 through 4096.')
  }

  const box = boxes[0]
  const objectKind = textboxObjectKind(attrs, box)
  const straight = objectKind === 'line' && isStraightLineKind(box.prst)
  if (straight && input.heightPx !== box.heightPx) {
    return invalid(`Straight DOCX lines require their current ${box.heightPx ?? 12}px grab height.`)
  }
  const changed = box.widthPx !== input.widthPx || (!straight && box.heightPx !== input.heightPx)
  if (changed) {
    const next = straight
      ? { ...box, widthPx: input.widthPx }
      : {
          ...box,
          widthPx: input.widthPx,
          heightPx: input.heightPx,
          minHeightPx: input.heightPx,
        }
    dispatch(
      closeHistory(state.tr.setNodeMarkup(position, undefined, { ...attrs, textboxes: [next] })),
    )
  }
  return { ok: true, objectKind, changed }
}

/** Resolve one stable top-level block and set its final drawing size. */
export function setDocxObjectSizeAtBlock(
  editor: Editor,
  input: SetDocxObjectSizeInput,
): SetDocxObjectSizeResult {
  const { doc } = editor.state
  if (input.objectBlockIndex < 0 || input.objectBlockIndex >= doc.childCount) {
    return invalid(
      `DOCX object block ${input.objectBlockIndex} is invalid for ${doc.childCount} block(s).`,
    )
  }
  let position = 0
  for (let index = 0; index < input.objectBlockIndex; index += 1) {
    position += doc.child(index).nodeSize
  }
  return setDocxObjectSizeAtPosition(
    editor.state,
    (transaction) => editor.view.dispatch(transaction),
    position,
    input,
  )
}

/** Shared final-position kernel for shape draw, move handle, and Registry. */
export function setDocxObjectOffsetAtPosition(
  state: EditorState,
  dispatch: DispatchTransaction,
  position: number,
  input: Pick<SetDocxObjectOffsetInput, 'wrap' | 'offsetXEmu' | 'offsetYEmu'>,
  options: { readonly closeHistory?: boolean } = {},
): SetDocxObjectOffsetResult {
  const node = state.doc.nodeAt(position)
  if (!node || node.type.name !== 'docProtected' || node.attrs.blockType === 'image') {
    return invalidOffset('The DOCX object target is not a positionable drawing block.')
  }
  const boxes = node.attrs.textboxes as TextboxDisplay[] | null
  if (!Array.isArray(boxes) || boxes.length === 0) {
    return invalidOffset('The DOCX object target has no positionable shape or textbox content.')
  }
  if (
    !Number.isInteger(input.offsetXEmu) ||
    input.offsetXEmu < -2_147_483_648 ||
    input.offsetXEmu > 2_147_483_647 ||
    !Number.isInteger(input.offsetYEmu) ||
    input.offsetYEmu < -2_147_483_648 ||
    input.offsetYEmu > 2_147_483_647
  ) {
    return invalidOffset('DOCX object offsets must be signed 32-bit integer EMU values.')
  }
  const objectKind: DocxPositionableObjectKind = node.attrs.diagramDisplay
    ? 'diagram'
    : boxes.length === 1
      ? (textboxObjectKind(node.attrs, boxes[0]) as DocxPositionableObjectKind)
      : 'textbox'
  const changed =
    node.attrs.imageWrap !== input.wrap ||
    node.attrs.imageOffsetXEmu !== input.offsetXEmu ||
    node.attrs.imageOffsetYEmu !== input.offsetYEmu ||
    node.attrs.imagePosH !== null ||
    node.attrs.imagePosV !== null
  if (changed) {
    const transaction = state.tr.setNodeMarkup(position, undefined, {
      ...node.attrs,
      imageWrap: input.wrap,
      imageOffsetXEmu: input.offsetXEmu,
      imageOffsetYEmu: input.offsetYEmu,
      imagePosH: null,
      imagePosV: null,
    })
    dispatch(options.closeHistory === false ? transaction : closeHistory(transaction))
  }
  return { ok: true, objectKind, changed }
}

/** Resolve one stable top-level object block and set its final floating position. */
export function setDocxObjectOffsetAtBlock(
  editor: Editor,
  input: SetDocxObjectOffsetInput,
): SetDocxObjectOffsetResult {
  const { doc } = editor.state
  if (input.objectBlockIndex < 0 || input.objectBlockIndex >= doc.childCount) {
    return invalidOffset(
      `DOCX object block ${input.objectBlockIndex} is invalid for ${doc.childCount} block(s).`,
    )
  }
  let position = 0
  for (let index = 0; index < input.objectBlockIndex; index += 1) {
    position += doc.child(index).nodeSize
  }
  return setDocxObjectOffsetAtPosition(
    editor.state,
    (transaction) => editor.view.dispatch(transaction),
    position,
    input,
  )
}

/** Shared masked final-style kernel for the retained Shape Format palette and Registry. */
export function setDocxObjectStyleAtPosition(
  state: EditorState,
  dispatch: DispatchTransaction,
  position: number,
  input: Pick<SetDocxObjectStyleInput, 'style' | 'fields'>,
): SetDocxObjectStyleResult {
  const node = state.doc.nodeAt(position)
  if (!node || node.type.name !== 'docProtected' || node.attrs.blockType === 'image') {
    return invalidStyle('The DOCX object target is not a styleable drawing block.')
  }
  const boxes = node.attrs.textboxes as TextboxDisplay[] | null
  if (!Array.isArray(boxes) || boxes.length === 0) {
    return invalidStyle('The DOCX object target has no styleable shape or textbox content.')
  }
  if (!Array.isArray(input.fields) || input.fields.length < 1 || input.fields.length > 2) {
    return invalidStyle('DOCX object style fields must contain one or two masked fields.')
  }
  const fields = new Set<DocxObjectStyleField>()
  for (const field of input.fields) {
    if (field !== 'fillHex' && field !== 'borderHex') {
      return invalidStyle(`Unsupported DOCX object style field: ${String(field)}.`)
    }
    if (fields.has(field)) {
      return invalidStyle(`DOCX object style field ${field} must not be repeated.`)
    }
    fields.add(field)
  }
  const styleKeys = Object.keys(input.style)
  if (
    styleKeys.length !== fields.size ||
    styleKeys.some((key) => !fields.has(key as DocxObjectStyleField))
  ) {
    return invalidStyle('DOCX object style must contain exactly the fields named by fields.')
  }
  for (const field of fields) {
    const value = input.style[field]
    if (value !== null && (typeof value !== 'string' || !/^[0-9A-F]{6}$/.test(value))) {
      return invalidStyle(`DOCX object ${field} must be null or a six-digit uppercase hex color.`)
    }
  }

  const objectKind: DocxStyleableObjectKind = node.attrs.diagramDisplay
    ? 'diagram'
    : boxes.length === 1
      ? (textboxObjectKind(node.attrs, boxes[0]) as DocxStyleableObjectKind)
      : 'textbox'
  if (objectKind === 'line' && fields.has('fillHex')) {
    return invalidStyle('Stroke-only DOCX lines do not support fillHex.')
  }

  const box = boxes[0]
  const next: TextboxDisplay = { ...box }
  let changed = false
  if (fields.has('fillHex')) {
    const fill = input.style.fillHex
    changed ||= fill === null ? box.fill !== undefined : box.fill !== fill
    if (fill === null) delete next.fill
    else next.fill = fill
  }
  if (fields.has('borderHex')) {
    const border = input.style.borderHex
    changed ||= border === null ? box.borderColor !== undefined : box.borderColor !== border
    if (border === null) delete next.borderColor
    else next.borderColor = border
  }
  if (changed) {
    dispatch(
      closeHistory(
        state.tr.setNodeMarkup(position, undefined, {
          ...node.attrs,
          textboxes: [next, ...boxes.slice(1)],
        }),
      ),
    )
  }
  return { ok: true, objectKind, changed }
}

/** Resolve one stable top-level object block and set its masked final style. */
export function setDocxObjectStyleAtBlock(
  editor: Editor,
  input: SetDocxObjectStyleInput,
): SetDocxObjectStyleResult {
  const { doc } = editor.state
  if (input.objectBlockIndex < 0 || input.objectBlockIndex >= doc.childCount) {
    return invalidStyle(
      `DOCX object block ${input.objectBlockIndex} is invalid for ${doc.childCount} block(s).`,
    )
  }
  let position = 0
  for (let index = 0; index < input.objectBlockIndex; index += 1) {
    position += doc.child(index).nodeSize
  }
  return setDocxObjectStyleAtPosition(
    editor.state,
    (transaction) => editor.view.dispatch(transaction),
    position,
    input,
  )
}

/** Resolve the selected protected drawing so UI and Registry share the same style mutation. */
export function setSelectedDocxObjectStyle(
  editor: Editor,
  input: Pick<SetDocxObjectStyleInput, 'style' | 'fields'>,
): SetDocxObjectStyleResult {
  editor.view.focus()
  const { selection } = editor.state
  let position: number | null = null
  if (editor.state.doc.nodeAt(selection.from)?.type.name === 'docProtected') {
    position = selection.from
  } else {
    for (let depth = selection.$from.depth; depth > 0; depth -= 1) {
      if (selection.$from.node(depth).type.name === 'docProtected') {
        position = selection.$from.before(depth)
        break
      }
    }
  }
  if (position === null) {
    return invalidStyle('Select one DOCX shape, line, or textbox before changing its style.')
  }
  return setDocxObjectStyleAtPosition(
    editor.state,
    (transaction) => editor.view.dispatch(transaction),
    position,
    input,
  )
}
