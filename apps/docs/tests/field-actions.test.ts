import { Editor } from '@tiptap/core'
import { parseDocx, saveDocx } from '@genoffice/docx-engine'
import { describe, expect, it } from 'vitest'

import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { insertDocxField, updateDocxFields } from '../src/renderer/editor/field-actions'

describe('DOCX generic field actions', () => {
  it('inserts through native Undo and survives save/reopen', async () => {
    const source = await buildDocx({
      bodyXml: '<w:p><w:r><w:t>Alpha</w:t></w:r></w:p>',
    })
    const parsed = await parseDocx(source)
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks) as never,
    })

    expect(
      insertDocxField(editor, {
        range: { from: 1, to: 1 },
        instruction: 'PAGE',
        displayText: '1',
      }),
    ).toMatchObject({ ok: true, changed: true, from: 1, to: 2 })
    expect(editor.commands.undo()).toBe(true)
    expect(editor.state.doc.textContent).toBe('Alpha')
    expect(editor.commands.redo()).toBe(true)
    expect(
      updateDocxFields(editor, {
        updates: [{ range: { from: 1, to: 2 }, instruction: 'PAGE', displayText: '12' }],
      }),
    ).toEqual({ ok: true, matched: 1, changed: 1 })
    expect(editor.commands.undo()).toBe(true)
    expect(editor.state.doc.textContent).toBe('1Alpha')
    expect(editor.commands.redo()).toBe(true)

    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    const reopened = await parseDocx(await saveDocx(parsed, plan.saveBlocks))
    expect(reopened.blocks[0].runs?.[0]).toMatchObject({ text: '12', instrField: 'PAGE' })
    editor.destroy()
  })

  it('updates multiple exact caches as one Undo unit', () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docParagraph',
            attrs: { docxIndex: 0 },
            content: [
              {
                type: 'text',
                text: '1',
                marks: [{ type: 'instrField', attrs: { instr: 'PAGE' } }],
              },
              {
                type: 'text',
                text: '2',
                marks: [{ type: 'instrField', attrs: { instr: 'NUMPAGES' } }],
              },
            ],
          },
        ],
      },
    })
    expect(
      updateDocxFields(editor, {
        updates: [
          { range: { from: 1, to: 2 }, instruction: 'PAGE', displayText: '3' },
          { range: { from: 2, to: 3 }, instruction: 'NUMPAGES', displayText: '10' },
        ],
      }),
    ).toEqual({ ok: true, matched: 2, changed: 2 })
    expect(editor.state.doc.textContent).toBe('310')
    expect(editor.commands.undo()).toBe(true)
    expect(editor.state.doc.textContent).toBe('12')
    editor.destroy()
  })
})
