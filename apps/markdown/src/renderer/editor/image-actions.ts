import type { Editor } from '@tiptap/core'

export interface MarkdownImageInput {
  readonly position: number
  readonly src: string
  readonly alt: string
  readonly title: string | null
}

export type MarkdownImageResult =
  | { readonly ok: true }
  | { readonly ok: false; readonly message: string }

export function stagedMarkdownImageMediaType(
  name: string,
  data: ArrayBuffer,
): 'image/png' | 'image/jpeg' | 'image/gif' | null {
  const bytes = new Uint8Array(data)
  if (
    /\.png$/i.test(name) &&
    bytes.length >= 8 &&
    [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a].every(
      (value, index) => bytes[index] === value,
    )
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
  if (/\.gif$/i.test(name) && bytes.length >= 6) {
    const signature = String.fromCharCode(...bytes.subarray(0, 6))
    if (signature === 'GIF87a' || signature === 'GIF89a') return 'image/gif'
  }
  return null
}

export function markdownImageDataUrl(mediaType: string, data: ArrayBuffer): string {
  const bytes = new Uint8Array(data)
  let binary = ''
  for (let index = 0; index < bytes.length; index += 0x8000) {
    binary += String.fromCharCode(...bytes.subarray(index, index + 0x8000))
  }
  return `data:${mediaType};base64,${btoa(binary)}`
}

/** Insert one fully resolved browser-safe image through the mounted TipTap history route. */
export function insertMarkdownImage(
  editor: Editor,
  input: MarkdownImageInput,
): MarkdownImageResult {
  const documentSize = editor.state.doc.content.size
  if (
    !Number.isInteger(input.position) ||
    input.position < 1 ||
    input.position > documentSize
  ) {
    return {
      ok: false,
      message: `Markdown insertion position ${input.position} is invalid for document size ${documentSize}.`,
    }
  }
  const inserted = editor
    .chain()
    .focus()
    .insertContentAt(input.position, {
      type: 'image',
      attrs: { src: input.src, alt: input.alt, title: input.title },
    })
    .run()
  return inserted
    ? { ok: true }
    : { ok: false, message: `Markdown image cannot be inserted at position ${input.position}.` }
}
