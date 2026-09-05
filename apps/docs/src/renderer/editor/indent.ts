import type { Editor } from '@tiptap/core'

/** Word's ribbon indent step: half an inch, in twips */
const STEP = 720

export function normalizeListLevel(level: number): number {
  return Math.min(Math.max(Math.trunc(level), 0), 8)
}

/** paragraph-like blocks that carry indent attributes */
const PARA_TYPES = new Set(['docParagraph', 'docHeading', 'docListItem'])

/** Word Ctrl+T / Ctrl+Shift+T hanging-indent steps. */
export function stepHangingIndent(editor: Editor, delta: 1 | -1): boolean {
  return editor.chain().focus().command(({ state, tr, dispatch }) => {
    const { from, to } = state.selection
    let changed = false
    state.doc.nodesBetween(from, to, (node, pos) => {
      if (!PARA_TYPES.has(node.type.name)) return
      const left = Number(node.attrs.indentLeft) || 0
      const nextLeft = Math.max(delta > 0 ? Math.floor(left / STEP) * STEP + STEP : Math.ceil(left / STEP) * STEP - STEP, 0)
      const nextHanging = nextLeft > 0 ? -Math.min(nextLeft, STEP) : 0
      if (nextLeft === left && nextHanging === (Number(node.attrs.indentFirstLine) || 0)) return
      tr.setNodeMarkup(pos, undefined, { ...node.attrs, indentLeft: nextLeft || null, indentFirstLine: nextHanging || null })
      changed = true
    })
    if (changed && dispatch) dispatch(tr)
    return changed
  }).run()
}

/** Ribbon increase/decrease indent: list items change level, plain paragraphs and headings snap to the next half-inch stop */
export function stepParagraphIndent(editor: Editor, delta: 1 | -1): boolean {
  if (editor.isActive('docListItem')) {
    const ilvl = Number(editor.getAttributes('docListItem').ilvl) || 0
    const next = normalizeListLevel(ilvl + delta)
    return editor.chain().focus().updateAttributes('docListItem', { ilvl: next }).run()
  }
  return editor
    .chain()
    .focus()
    .command(({ state, tr, dispatch }) => {
      const { from, to } = state.selection
      let changed = false
      state.doc.nodesBetween(from, to, (node, pos) => {
        if (node.type.name !== 'docParagraph' && node.type.name !== 'docHeading') return
        const cur = Number(node.attrs.indentLeft) || 0
        const next =
          delta > 0
            ? Math.floor(cur / STEP) * STEP + STEP
            : Math.max(Math.ceil(cur / STEP) * STEP - STEP, 0)
        if (next === cur) return
        tr.setNodeMarkup(pos, undefined, { ...node.attrs, indentLeft: next || null })
        changed = true
      })
      if (changed && dispatch) dispatch(tr)
      return changed
    })
    .run()
}
