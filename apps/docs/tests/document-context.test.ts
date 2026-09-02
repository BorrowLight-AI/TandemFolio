import { Editor } from '@tiptap/core'
import { describe, expect, it } from 'vitest'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { buildDocxSelectionContext } from '../src/renderer/host/document-context'

describe('DOCX selection context', () => {
  it('identifies the active block and returns bounded neighboring summaries', () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          { type: 'docHeading', attrs: { level: 1 }, content: [{ type: 'text', text: 'Plan' }] },
          { type: 'docParagraph', content: [{ type: 'text', text: 'Selected paragraph' }] },
        ],
      },
    })
    let secondBlockPosition = 0
    editor.state.doc.forEach((_node, offset, index) => {
      if (index === 1) secondBlockPosition = offset + 1
    })
    editor.commands.setTextSelection(secondBlockPosition)

    expect(buildDocxSelectionContext(editor)).toMatchObject({
      activeBlockIndex: 1,
      totalBlocks: 2,
      blocks: [
        { index: 0, type: 'docHeading', text: 'Plan' },
        { index: 1, type: 'docParagraph', text: 'Selected paragraph' },
      ],
    })
    editor.destroy()
  })
})
