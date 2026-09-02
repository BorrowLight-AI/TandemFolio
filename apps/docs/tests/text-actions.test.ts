import { Editor } from '@tiptap/core'
import { parseDocx, saveDocx } from '@genoffice/docx-engine'
import { describe, expect, it } from 'vitest'

import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { insertDocxText } from '../src/renderer/editor/text-actions'

function createEditor(): Editor {
  return new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: {
      type: 'doc',
      content: [
        {
          type: 'docParagraph',
          attrs: { docxIndex: 0 },
          content: [{ type: 'text', text: 'Alpha' }],
        },
      ],
    },
  })
}

describe('DOCX text actions', () => {
  it('inserts a Unicode symbol at the active selection as one Undo unit', () => {
    const editor = createEditor()
    editor.commands.setTextSelection(1)
    expect(insertDocxText(editor, { text: '∞' })).toEqual({ ok: true, changed: true })
    expect(editor.state.doc.textContent).toBe('∞Alpha')
    expect(editor.commands.undo()).toBe(true)
    expect(editor.state.doc.textContent).toBe('Alpha')
    editor.destroy()
  })

  it('rejects empty insertion without mutation', () => {
    const editor = createEditor()
    expect(insertDocxText(editor, { text: '' })).toMatchObject({
      ok: false,
      error: 'invalid_arguments',
    })
    expect(editor.state.doc.textContent).toBe('Alpha')
    editor.destroy()
  })

  it('persists an inserted symbol through save/reopen', async () => {
    const source = await buildDocx({
      bodyXml: '<w:p><w:r><w:t>Alpha</w:t></w:r></w:p>',
    })
    const parsed = await parseDocx(source)
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks) as never,
    })
    editor.commands.setTextSelection(1)
    expect(insertDocxText(editor, { text: '∞' })).toMatchObject({ ok: true })
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    const reopened = await parseDocx(await saveDocx(parsed, plan.saveBlocks))
    expect(reopened.blocks[0].runs?.map((run) => run.text).join('')).toBe('∞Alpha')
    editor.destroy()
  })
})
