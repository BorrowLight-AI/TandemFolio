import type { Editor } from '@tiptap/core'
import { closeHistory } from '@tiptap/pm/history'

export interface SetDocxBookmarkInput {
  readonly blockIndex: number
  readonly name: string
  readonly enabled: boolean
}

export type SetDocxBookmarkResult =
  | {
      readonly ok: true
      readonly changed: boolean
      readonly name: string
      readonly enabled: boolean
      readonly blockIndex: number
    }
  | {
      readonly ok: false
      readonly error: 'invalid_arguments' | 'execution_failed'
      readonly message: string
    }

const BOOKMARK_NAME_PATTERN = /^[A-Za-z_一-鿿][A-Za-z0-9_一-鿿]*$/
const BOOKMARK_BLOCK_TYPES = new Set(['docParagraph', 'docHeading', 'docListItem'])

function invalid(message: string): SetDocxBookmarkResult {
  return { ok: false, error: 'invalid_arguments', message }
}

export function isValidDocxBookmarkName(name: string): boolean {
  const length = Array.from(name).length
  return length >= 1 && length <= 40 && BOOKMARK_NAME_PATTERN.test(name)
}

/** Set one top-level paragraph bookmark to an explicit final state through native history. */
export function setDocxBookmark(
  editor: Editor,
  input: SetDocxBookmarkInput,
): SetDocxBookmarkResult {
  const { doc } = editor.state
  if (
    !input ||
    !Number.isInteger(input.blockIndex) ||
    input.blockIndex < 0 ||
    input.blockIndex >= doc.childCount
  ) {
    return invalid(
      `DOCX bookmark block ${input?.blockIndex ?? 'unknown'} is invalid for ${doc.childCount} block(s).`,
    )
  }
  if (typeof input.name !== 'string' || !isValidDocxBookmarkName(input.name)) {
    return invalid(
      'DOCX bookmark names must contain 1 through 40 letters, digits, underscores, or CJK characters and may not start with a digit.',
    )
  }
  if (typeof input.enabled !== 'boolean') {
    return invalid('DOCX bookmark enabled must be a boolean final state.')
  }

  const target = doc.child(input.blockIndex)
  if (!BOOKMARK_BLOCK_TYPES.has(target.type.name)) {
    return invalid(
      `DOCX bookmark block ${input.blockIndex} must be a paragraph, heading, or list item.`,
    )
  }

  let duplicateBlockIndex = -1
  doc.forEach((node, _position, blockIndex) => {
    if (
      blockIndex !== input.blockIndex &&
      Array.isArray(node.attrs.bookmarks) &&
      (node.attrs.bookmarks as readonly unknown[]).includes(input.name)
    ) {
      duplicateBlockIndex = blockIndex
    }
  })
  if (input.enabled && duplicateBlockIndex >= 0) {
    return invalid(`DOCX bookmark ${input.name} already exists on block ${duplicateBlockIndex}.`)
  }

  const current = Array.isArray(target.attrs.bookmarks)
    ? (target.attrs.bookmarks as readonly string[])
    : []
  const alreadyEnabled = current.includes(input.name)
  if (alreadyEnabled === input.enabled) {
    return {
      ok: true,
      changed: false,
      name: input.name,
      enabled: input.enabled,
      blockIndex: input.blockIndex,
    }
  }
  const next = input.enabled
    ? [...current, input.name]
    : current.filter((name) => name !== input.name)
  let position = 0
  for (let index = 0; index < input.blockIndex; index += 1) position += doc.child(index).nodeSize

  try {
    editor.view.dispatch(
      closeHistory(
        editor.state.tr.setNodeMarkup(position, undefined, {
          ...target.attrs,
          bookmarks: next.length > 0 ? next : null,
        }),
      ),
    )
    return {
      ok: true,
      changed: true,
      name: input.name,
      enabled: input.enabled,
      blockIndex: input.blockIndex,
    }
  } catch (error) {
    return {
      ok: false,
      error: 'execution_failed',
      message: `The mounted DOCX editor rejected the bookmark update: ${
        error instanceof Error ? error.message : String(error)
      }`,
    }
  }
}
