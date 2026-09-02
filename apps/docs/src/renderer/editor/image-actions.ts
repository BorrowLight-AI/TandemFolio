import type { Editor, JSONContent } from '@tiptap/core'
import type { ImageWrap, NewImage } from '@genoffice/docx-engine'

export type DocxImageAlignment = NonNullable<NewImage['align']>
export type DocxImageMediaType = NewImage['mime']
export type RetainedDocxImageWrap =
  'square-left' | 'square-right' | 'topBottom' | 'behind' | 'front' | null
export type DocxImageHorizontalPosition = 'left' | 'center' | 'right'
export type DocxImageVerticalPosition = 'top' | 'center' | 'bottom'
export interface DocxImageCrop {
  readonly l: number
  readonly t: number
  readonly r: number
  readonly b: number
}

export interface DocxImagePayload {
  readonly base64: string
  readonly mime: DocxImageMediaType
  readonly widthPx: number
  readonly heightPx: number
  readonly align?: DocxImageAlignment
  readonly label: string
}

export interface InsertDocxImageInput {
  readonly afterBlockIndex: number
  readonly widthPx: number
  readonly heightPx: number
  readonly alignment: DocxImageAlignment
  readonly name: string
  readonly data: ArrayBuffer
  readonly mime: DocxImageMediaType
}

export interface ReplaceDocxImageInput {
  readonly imageBlockIndex: number
  readonly widthPx: number
  readonly heightPx: number
  readonly data: ArrayBuffer
  readonly mime: DocxImageMediaType
}

export interface SetDocxImageWrapInput {
  readonly imageBlockIndex: number
  readonly wrap: RetainedDocxImageWrap
}

export interface SetDocxImageMarginPositionInput {
  readonly imageBlockIndex: number
  readonly horizontal: DocxImageHorizontalPosition
  readonly vertical: DocxImageVerticalPosition
}

export interface SetDocxImageOffsetPositionInput {
  readonly imageBlockIndex: number
  readonly wrap: ImageWrap
  readonly offsetXEmu: number
  readonly offsetYEmu: number
}

export interface SetDocxImageTransformInput {
  readonly imageBlockIndex: number
  readonly rotationDegrees: number
  readonly flipHorizontal: boolean
  readonly flipVertical: boolean
}

export interface SetDocxImageCropInput {
  readonly imageBlockIndex: number
  readonly left: number
  readonly top: number
  readonly right: number
  readonly bottom: number
}

export interface DocxImageReplacementPayload {
  readonly base64: string
  readonly mime: DocxImageMediaType
  readonly widthPx: number
  readonly heightPx: number
}

export type InsertDocxImageResult =
  | { readonly ok: true; readonly insertedBlockIndex: number }
  | { readonly ok: false; readonly error: 'invalid_arguments'; readonly message: string }

export type ReplaceDocxImageResult =
  | { readonly ok: true; readonly imageBlockIndex: number }
  | { readonly ok: false; readonly error: 'invalid_arguments'; readonly message: string }

export type SetDocxImageWrapResult =
  | { readonly ok: true; readonly imageBlockIndex: number; readonly changed: boolean }
  | { readonly ok: false; readonly error: 'invalid_arguments'; readonly message: string }

export type SetDocxImageMarginPositionResult =
  | { readonly ok: true; readonly imageBlockIndex: number; readonly changed: boolean }
  | { readonly ok: false; readonly error: 'invalid_arguments'; readonly message: string }

export type SetDocxImageOffsetPositionResult =
  | { readonly ok: true; readonly imageBlockIndex: number; readonly changed: boolean }
  | { readonly ok: false; readonly error: 'invalid_arguments'; readonly message: string }

export type SetDocxImageTransformResult =
  | { readonly ok: true; readonly imageBlockIndex: number; readonly changed: boolean }
  | { readonly ok: false; readonly error: 'invalid_arguments'; readonly message: string }

export type SetDocxImageCropResult =
  | { readonly ok: true; readonly imageBlockIndex: number; readonly changed: boolean }
  | { readonly ok: false; readonly error: 'invalid_arguments'; readonly message: string }

export type RemoveDocxImageResult =
  | { readonly ok: true; readonly imageBlockIndex: number }
  | {
      readonly ok: false
      readonly error: 'invalid_arguments' | 'execution_failed'
      readonly message: string
    }

function encodeBase64(data: ArrayBuffer): string {
  const bytes = new Uint8Array(data)
  let binary = ''
  for (let offset = 0; offset < bytes.length; offset += 32_768) {
    binary += String.fromCharCode(...bytes.subarray(offset, offset + 32_768))
  }
  return btoa(binary)
}

function imageContent(payload: DocxImagePayload): JSONContent {
  const image: NewImage = {
    base64: payload.base64,
    mime: payload.mime,
    widthPx: payload.widthPx,
    heightPx: payload.heightPx,
    ...(payload.align ? { align: payload.align } : {}),
  }
  return {
    type: 'docProtected',
    attrs: {
      docxIndex: null,
      blockType: 'image',
      label: payload.label,
      imageDataUrl: `data:${payload.mime};base64,${payload.base64}`,
      imageWidthPx: payload.widthPx,
      imageHeightPx: payload.heightPx,
      imageAlign: payload.align ?? null,
      genImage: image,
    },
  }
}

function positionAfterBlock(editor: Editor, afterBlockIndex: number): number {
  let pos = 0
  for (let index = 0; index <= afterBlockIndex; index += 1) {
    pos += editor.state.doc.child(index).nodeSize
  }
  return pos
}

function positionAtBlock(editor: Editor, blockIndex: number): number {
  let pos = 0
  for (let index = 0; index < blockIndex; index += 1) {
    pos += editor.state.doc.child(index).nodeSize
  }
  return pos
}

function replacementAttrs(
  attrs: Readonly<Record<string, unknown>>,
  payload: DocxImageReplacementPayload,
): Readonly<Record<string, unknown>> {
  const original = attrs.docxIndex !== null && attrs.docxIndex !== undefined
  const align =
    attrs.imageAlign === 'left' || attrs.imageAlign === 'center' || attrs.imageAlign === 'right'
      ? attrs.imageAlign
      : undefined
  return {
    imageDataUrl: `data:${payload.mime};base64,${payload.base64}`,
    imageWidthPx: payload.widthPx,
    imageHeightPx: payload.heightPx,
    imageCrop: null,
    imageFillRect: null,
    ...(original
      ? { imageReplace: { base64: payload.base64, mime: payload.mime } }
      : {
          imageReplace: null,
          genImage: {
            base64: payload.base64,
            mime: payload.mime,
            widthPx: payload.widthPx,
            heightPx: payload.heightPx,
            ...(align ? { align } : {}),
          } satisfies NewImage,
        }),
  }
}

function wrapAttrs(wrap: RetainedDocxImageWrap): Readonly<Record<string, unknown>> {
  return wrap === null
    ? {
        imageWrap: null,
        imagePosH: null,
        imagePosV: null,
        imageOffsetXEmu: null,
        imageOffsetYEmu: null,
      }
    : { imageWrap: wrap }
}

function marginPositionAttrs(
  horizontal: DocxImageHorizontalPosition,
  vertical: DocxImageVerticalPosition,
): Readonly<Record<string, unknown>> {
  return {
    imageWrap: horizontal === 'right' ? 'square-right' : 'square-left',
    imagePosH: horizontal,
    imagePosV: vertical,
    imageOffsetXEmu: null,
    imageOffsetYEmu: null,
  }
}

/** Project an explicit floating-image position into the retained DOCX node model. */
export function docxImageOffsetPositionAttrs(
  wrap: ImageWrap,
  offsetXEmu: number,
  offsetYEmu: number,
): Readonly<Record<string, unknown>> {
  return {
    imageWrap: wrap,
    imageOffsetXEmu: offsetXEmu,
    imageOffsetYEmu: offsetYEmu,
    imagePosH: null,
    imagePosV: null,
  }
}

/** Project the complete replayable picture transform into canonical node attrs. */
export function docxImageTransformAttrs(
  rotationDegrees: number,
  flipHorizontal: boolean,
  flipVertical: boolean,
): Readonly<Record<string, unknown>> {
  return {
    imageRotDeg: rotationDegrees === 0 ? null : rotationDegrees,
    imageFlipH: flipHorizontal,
    imageFlipV: flipVertical,
  }
}

/** Project non-destructive source cropping and discard any stale fill window. */
export function docxImageCropAttrs(
  left: number,
  top: number,
  right: number,
  bottom: number,
): Readonly<Record<string, unknown>> {
  const clear = left === 0 && top === 0 && right === 0 && bottom === 0
  return {
    imageCrop: clear ? null : ({ l: left, t: top, r: right, b: bottom } satisfies DocxImageCrop),
    imageFillRect: null,
  }
}

function attrsDiffer(
  attrs: Readonly<Record<string, unknown>>,
  patch: Readonly<Record<string, unknown>>,
): boolean {
  return Object.entries(patch).some(([key, value]) => attrs[key] !== value)
}

/** Preserve selection-oriented Ribbon/paste behavior while sharing the image node model. */
export function insertDocxImageAtSelection(editor: Editor, payload: DocxImagePayload): boolean {
  return editor.chain().focus().insertContent(imageContent(payload)).run()
}

/** Insert hydrated bytes at a stable top-level boundary in one undo-owned transaction. */
export function insertDocxImageAfterBlock(
  editor: Editor,
  input: InsertDocxImageInput,
): InsertDocxImageResult {
  const { doc } = editor.state
  if (input.afterBlockIndex < -1 || input.afterBlockIndex >= doc.childCount) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: `DOCX image boundary after block ${input.afterBlockIndex} is invalid for ${doc.childCount} block(s).`,
    }
  }

  const payload: DocxImagePayload = {
    base64: encodeBase64(input.data),
    mime: input.mime,
    widthPx: input.widthPx,
    heightPx: input.heightPx,
    align: input.alignment,
    label: `Image (${input.name})`,
  }
  const node = editor.state.schema.nodeFromJSON(imageContent(payload))
  editor.view.dispatch(
    editor.state.tr.insert(positionAfterBlock(editor, input.afterBlockIndex), node),
  )
  return { ok: true, insertedBlockIndex: input.afterBlockIndex + 1 }
}

/** Preserve the retained selected-picture UI while sharing the exact replacement attrs. */
export function replaceSelectedDocxImage(
  editor: Editor,
  payload: DocxImageReplacementPayload,
): boolean {
  const attrs = editor.getAttributes('docProtected')
  if (attrs.blockType !== 'image') return false
  return editor
    .chain()
    .focus()
    .updateAttributes('docProtected', replacementAttrs(attrs, payload))
    .run()
}

/** Replace one exact top-level image in one native Undo transaction. */
export function replaceDocxImageAtBlock(
  editor: Editor,
  input: ReplaceDocxImageInput,
): ReplaceDocxImageResult {
  const { doc } = editor.state
  if (input.imageBlockIndex < 0 || input.imageBlockIndex >= doc.childCount) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: `DOCX image block ${input.imageBlockIndex} is invalid for ${doc.childCount} block(s).`,
    }
  }
  const node = doc.child(input.imageBlockIndex)
  if (node.type.name !== 'docProtected' || node.attrs.blockType !== 'image') {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: `DOCX block ${input.imageBlockIndex} is not an image.`,
    }
  }
  const payload: DocxImageReplacementPayload = {
    base64: encodeBase64(input.data),
    mime: input.mime,
    widthPx: input.widthPx,
    heightPx: input.heightPx,
  }
  editor.view.dispatch(
    editor.state.tr.setNodeMarkup(positionAtBlock(editor, input.imageBlockIndex), undefined, {
      ...node.attrs,
      ...replacementAttrs(node.attrs, payload),
    }),
  )
  return { ok: true, imageBlockIndex: input.imageBlockIndex }
}

/** Remove one exact top-level image while preserving the document block+ invariant. */
export function removeDocxImageAtBlock(
  editor: Editor,
  imageBlockIndex: number,
): RemoveDocxImageResult {
  const { doc } = editor.state
  if (imageBlockIndex < 0 || imageBlockIndex >= doc.childCount) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: `DOCX image block ${imageBlockIndex} is invalid for ${doc.childCount} block(s).`,
    }
  }
  const node = doc.child(imageBlockIndex)
  if (node.type.name !== 'docProtected' || node.attrs.blockType !== 'image') {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: `DOCX block ${imageBlockIndex} is not an image.`,
    }
  }
  const pos = positionAtBlock(editor, imageBlockIndex)
  const tr = editor.state.tr
  if (doc.childCount === 1) {
    const paragraph = editor.schema.nodes.docParagraph?.createAndFill()
    if (!paragraph) {
      return {
        ok: false,
        error: 'execution_failed',
        message: 'The DOCX schema could not create a replacement paragraph.',
      }
    }
    tr.replaceWith(pos, pos + node.nodeSize, paragraph)
  } else {
    tr.delete(pos, pos + node.nodeSize)
  }
  editor.view.dispatch(tr)
  return { ok: true, imageBlockIndex }
}

export function setSelectedDocxImageWrap(editor: Editor, wrap: RetainedDocxImageWrap): boolean {
  const attrs = editor.getAttributes('docProtected')
  if (attrs.blockType !== 'image') return false
  const patch = wrapAttrs(wrap)
  if (!attrsDiffer(attrs, patch)) return false
  return editor.chain().focus().updateAttributes('docProtected', patch).run()
}

export function setSelectedDocxImageMarginPosition(
  editor: Editor,
  horizontal: DocxImageHorizontalPosition,
  vertical: DocxImageVerticalPosition,
): boolean {
  const attrs = editor.getAttributes('docProtected')
  if (attrs.blockType !== 'image' || attrs.docxIndex === null || attrs.docxIndex === undefined) {
    return false
  }
  const patch = marginPositionAttrs(horizontal, vertical)
  if (!attrsDiffer(attrs, patch)) return false
  return editor.chain().focus().updateAttributes('docProtected', patch).run()
}

export function setSelectedDocxImageTransform(
  editor: Editor,
  rotationDegrees: number,
  flipHorizontal: boolean,
  flipVertical: boolean,
): boolean {
  const attrs = editor.getAttributes('docProtected')
  if (attrs.blockType !== 'image') return false
  const patch = docxImageTransformAttrs(rotationDegrees, flipHorizontal, flipVertical)
  if (!attrsDiffer(attrs, patch)) return false
  return editor.chain().focus().updateAttributes('docProtected', patch).run()
}

export function setSelectedDocxImageCrop(
  editor: Editor,
  left: number,
  top: number,
  right: number,
  bottom: number,
): boolean {
  const attrs = editor.getAttributes('docProtected')
  if (attrs.blockType !== 'image' || left + right >= 1 || top + bottom >= 1) return false
  const patch = docxImageCropAttrs(left, top, right, bottom)
  if (!attrsDiffer(attrs, patch)) return false
  return editor.chain().focus().updateAttributes('docProtected', patch).run()
}

export function setDocxImageMarginPositionAtBlock(
  editor: Editor,
  input: SetDocxImageMarginPositionInput,
): SetDocxImageMarginPositionResult {
  const { doc } = editor.state
  if (input.imageBlockIndex < 0 || input.imageBlockIndex >= doc.childCount) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: `DOCX image block ${input.imageBlockIndex} is invalid for ${doc.childCount} block(s).`,
    }
  }
  const node = doc.child(input.imageBlockIndex)
  if (node.type.name !== 'docProtected' || node.attrs.blockType !== 'image') {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: `DOCX block ${input.imageBlockIndex} is not an image.`,
    }
  }
  if (node.attrs.docxIndex === null || node.attrs.docxIndex === undefined) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: `DOCX image block ${input.imageBlockIndex} must be saved before setting a margin position.`,
    }
  }
  const patch = marginPositionAttrs(input.horizontal, input.vertical)
  const changed = attrsDiffer(node.attrs, patch)
  if (changed) {
    editor.view.dispatch(
      editor.state.tr.setNodeMarkup(positionAtBlock(editor, input.imageBlockIndex), undefined, {
        ...node.attrs,
        ...patch,
      }),
    )
  }
  return { ok: true, imageBlockIndex: input.imageBlockIndex, changed }
}

/** Set one exact image to an explicit offset in one native Undo transaction. */
export function setDocxImageOffsetPositionAtBlock(
  editor: Editor,
  input: SetDocxImageOffsetPositionInput,
): SetDocxImageOffsetPositionResult {
  const { doc } = editor.state
  if (input.imageBlockIndex < 0 || input.imageBlockIndex >= doc.childCount) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: `DOCX image block ${input.imageBlockIndex} is invalid for ${doc.childCount} block(s).`,
    }
  }
  const node = doc.child(input.imageBlockIndex)
  if (node.type.name !== 'docProtected' || node.attrs.blockType !== 'image') {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: `DOCX block ${input.imageBlockIndex} is not an image.`,
    }
  }
  const patch = docxImageOffsetPositionAttrs(input.wrap, input.offsetXEmu, input.offsetYEmu)
  const changed = attrsDiffer(node.attrs, patch)
  if (changed) {
    editor.view.dispatch(
      editor.state.tr.setNodeMarkup(positionAtBlock(editor, input.imageBlockIndex), undefined, {
        ...node.attrs,
        ...patch,
      }),
    )
  }
  return { ok: true, imageBlockIndex: input.imageBlockIndex, changed }
}

/** Set one exact image transform in one native Undo transaction. */
export function setDocxImageTransformAtBlock(
  editor: Editor,
  input: SetDocxImageTransformInput,
): SetDocxImageTransformResult {
  const { doc } = editor.state
  if (input.imageBlockIndex < 0 || input.imageBlockIndex >= doc.childCount) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: `DOCX image block ${input.imageBlockIndex} is invalid for ${doc.childCount} block(s).`,
    }
  }
  const node = doc.child(input.imageBlockIndex)
  if (node.type.name !== 'docProtected' || node.attrs.blockType !== 'image') {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: `DOCX block ${input.imageBlockIndex} is not an image.`,
    }
  }
  const patch = docxImageTransformAttrs(
    input.rotationDegrees,
    input.flipHorizontal,
    input.flipVertical,
  )
  const changed = attrsDiffer(node.attrs, patch)
  if (changed) {
    editor.view.dispatch(
      editor.state.tr.setNodeMarkup(positionAtBlock(editor, input.imageBlockIndex), undefined, {
        ...node.attrs,
        ...patch,
      }),
    )
  }
  return { ok: true, imageBlockIndex: input.imageBlockIndex, changed }
}

/** Set one exact image crop in one native Undo transaction. */
export function setDocxImageCropAtBlock(
  editor: Editor,
  input: SetDocxImageCropInput,
): SetDocxImageCropResult {
  const { doc } = editor.state
  if (input.imageBlockIndex < 0 || input.imageBlockIndex >= doc.childCount) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: `DOCX image block ${input.imageBlockIndex} is invalid for ${doc.childCount} block(s).`,
    }
  }
  const node = doc.child(input.imageBlockIndex)
  if (node.type.name !== 'docProtected' || node.attrs.blockType !== 'image') {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: `DOCX block ${input.imageBlockIndex} is not an image.`,
    }
  }
  if (input.left + input.right >= 1 || input.top + input.bottom >= 1) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: 'DOCX image crop must retain positive horizontal and vertical source area.',
    }
  }
  const patch = docxImageCropAttrs(input.left, input.top, input.right, input.bottom)
  const changed = attrsDiffer(node.attrs, patch)
  if (changed) {
    editor.view.dispatch(
      editor.state.tr.setNodeMarkup(positionAtBlock(editor, input.imageBlockIndex), undefined, {
        ...node.attrs,
        ...patch,
      }),
    )
  }
  return { ok: true, imageBlockIndex: input.imageBlockIndex, changed }
}

export function setDocxImageWrapAtBlock(
  editor: Editor,
  input: SetDocxImageWrapInput,
): SetDocxImageWrapResult {
  const { doc } = editor.state
  if (input.imageBlockIndex < 0 || input.imageBlockIndex >= doc.childCount) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: `DOCX image block ${input.imageBlockIndex} is invalid for ${doc.childCount} block(s).`,
    }
  }
  const node = doc.child(input.imageBlockIndex)
  if (node.type.name !== 'docProtected' || node.attrs.blockType !== 'image') {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: `DOCX block ${input.imageBlockIndex} is not an image.`,
    }
  }
  const patch = wrapAttrs(input.wrap)
  const changed = attrsDiffer(node.attrs, patch)
  if (changed) {
    editor.view.dispatch(
      editor.state.tr.setNodeMarkup(positionAtBlock(editor, input.imageBlockIndex), undefined, {
        ...node.attrs,
        ...patch,
      }),
    )
  }
  return { ok: true, imageBlockIndex: input.imageBlockIndex, changed }
}

export function stagedDocxImageMediaType(
  name: string,
  data: ArrayBuffer,
): DocxImageMediaType | null {
  const bytes = new Uint8Array(data)
  if (
    name.toLowerCase().endsWith('.png') &&
    bytes.length >= 8 &&
    [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a].every((value, index) => bytes[index] === value)
  ) {
    return 'image/png'
  }
  if (
    /\.jpe?g$/i.test(name) &&
    bytes.length >= 3 &&
    bytes[0] === 0xff &&
    bytes[1] === 0xd8 &&
    bytes[2] === 0xff
  ) {
    return 'image/jpeg'
  }
  if (
    name.toLowerCase().endsWith('.gif') &&
    bytes.length >= 6 &&
    /^GIF8[79]a$/.test(String.fromCharCode(...bytes.subarray(0, 6)))
  ) {
    return 'image/gif'
  }
  return null
}
