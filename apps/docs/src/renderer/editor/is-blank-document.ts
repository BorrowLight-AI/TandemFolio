import type { Editor } from '@tiptap/core'

export function isBlankDocument(editor: Editor): boolean {
  const doc = editor.state.doc
  if (doc.childCount !== 1) return false
  const first = doc.child(0)
  return first.type.name === 'docParagraph' && first.content.size === 0
}
