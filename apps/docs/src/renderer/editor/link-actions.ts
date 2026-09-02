import type { Editor } from '@tiptap/core'
import { closeHistory } from '@tiptap/pm/history'
import { TextSelection } from '@tiptap/pm/state'

export interface SetDocxTextLinkInput {
  readonly range: { readonly from: number; readonly to: number }
  /** Non-null sets/inserts a link; null removes link marks. */
  readonly href: string | null
  /** Non-null replaces the exact range before setting href; null preserves text. */
  readonly text: string | null
}

export type SetDocxTextLinkResult =
  | {
      readonly ok: true
      readonly changed: boolean
      readonly from: number
      readonly to: number
    }
  | {
      readonly ok: false
      readonly error: 'invalid_arguments' | 'execution_failed'
      readonly message: string
    }

function invalid(message: string): SetDocxTextLinkResult {
  return { ok: false, error: 'invalid_arguments', message }
}

function bounded(value: string, minimum: number, maximum: number): boolean {
  const length = Array.from(value).length
  return length >= minimum && length <= maximum
}

/** Set, replace, or remove one exact inline link through native history. */
export function setDocxTextLink(
  editor: Editor,
  input: SetDocxTextLinkInput,
): SetDocxTextLinkResult {
  if (!input || !input.range) return invalid('DOCX links require an explicit range.')
  const { from, to } = input.range
  const { state } = editor
  const { doc } = state
  if (
    !Number.isInteger(from) ||
    !Number.isInteger(to) ||
    from < 1 ||
    from > to ||
    to > doc.content.size
  ) {
    return invalid(`DOCX link range ${from}..${to} is outside document size ${doc.content.size}.`)
  }
  const $from = doc.resolve(from)
  const $to = doc.resolve(to)
  if ($from.parent !== $to.parent || !$from.parent.inlineContent) {
    return invalid('DOCX link ranges must stay inside one text-bearing block.')
  }
  if (input.href === null && input.text !== null) {
    return invalid('DOCX link removal requires null text.')
  }
  const href = input.href?.trim() ?? null
  if (href !== null && !bounded(href, 1, 4096)) {
    return invalid('DOCX link href must contain 1 through 4096 Unicode characters after trimming.')
  }
  if (input.text !== null && !bounded(input.text, 1, 65_536)) {
    return invalid('DOCX link replacement text must contain 1 through 65536 Unicode characters.')
  }
  if (input.text === null && from === to) {
    return invalid('DOCX link mark updates require a non-empty range when text is null.')
  }
  if (input.text === null) {
    let containsText = false
    let containsUnsupportedAtom = false
    doc.nodesBetween(from, to, (node, position) => {
      if (node.isText && Math.max(from, position) < Math.min(to, position + node.nodeSize)) {
        containsText = true
      } else if (node.isLeaf && !node.isText) {
        containsUnsupportedAtom = true
      }
    })
    if (!containsText || containsUnsupportedAtom) {
      return invalid('DOCX link mark updates require an exact text-only range.')
    }
  }

  try {
    const link = state.schema.marks.link
    let tr = state.tr
    let resultTo = to
    if (input.text !== null) {
      const marked = state.schema.text(input.text, [link.create({ href, rId: null })])
      tr = tr.replaceWith(from, to, marked)
      resultTo = from + input.text.length
      tr = tr.setSelection(TextSelection.create(tr.doc, resultTo))
    } else if (href === null) {
      tr = tr.removeMark(from, to, link).setSelection(TextSelection.create(tr.doc, from, to))
    } else {
      tr = tr
        .addMark(from, to, link.create({ href, rId: null }))
        .setSelection(TextSelection.create(tr.doc, from, to))
    }
    const changed = tr.docChanged
    if (changed) editor.view.dispatch(closeHistory(tr))
    return { ok: true, changed, from, to: resultTo }
  } catch (error) {
    return {
      ok: false,
      error: 'execution_failed',
      message: `The mounted DOCX editor rejected the link update: ${
        error instanceof Error ? error.message : String(error)
      }`,
    }
  }
}
