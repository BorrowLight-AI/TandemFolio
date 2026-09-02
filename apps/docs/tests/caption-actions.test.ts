import { Editor } from '@tiptap/core'
import { parseDocx, saveDocx } from '@genoffice/docx-engine'
import { describe, expect, it } from 'vitest'

import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { insertDocxCaption } from '../src/renderer/editor/caption-actions'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'

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

describe('DOCX caption actions', () => {
  it('inserts one explicit caption after a stable boundary as one Undo unit', () => {
    const editor = createEditor()
    expect(
      insertDocxCaption(editor, {
        afterBlockIndex: 0,
        label: 'Figure',
        number: 2,
        text: 'Registry architecture',
      }),
    ).toEqual({
      ok: true,
      afterBlockIndex: 0,
      label: 'Figure',
      number: 2,
      changed: true,
    })
    expect(editor.state.doc.childCount).toBe(2)
    expect(editor.state.doc.child(1).type.name).toBe('docProtected')
    expect(editor.state.doc.child(1).attrs.fieldDisplay).toEqual({
      kind: 'text',
      left: 'Figure 2 Registry architecture',
    })
    expect(editor.commands.undo()).toBe(true)
    expect(editor.state.doc.childCount).toBe(1)
    editor.destroy()
  })

  it('rejects an invalid caption number without mutation', () => {
    const editor = createEditor()
    expect(
      insertDocxCaption(editor, {
        afterBlockIndex: 0,
        label: 'Figure',
        number: 0,
        text: '',
      }),
    ).toMatchObject({ ok: false, error: 'invalid_arguments' })
    expect(editor.state.doc.childCount).toBe(1)
    editor.destroy()
  })

  it('persists the explicit caption field through save and reopen', async () => {
    const parsed = await parseDocx(
      await buildDocx({ bodyXml: '<w:p><w:r><w:t>Alpha</w:t></w:r></w:p>' }),
    )
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks) as never,
    })
    expect(
      insertDocxCaption(editor, {
        afterBlockIndex: 0,
        label: 'Figure',
        number: 2,
        text: 'Registry architecture',
      }),
    ).toMatchObject({ ok: true })
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    const reopened = await parseDocx(await saveDocx(parsed, plan.saveBlocks))
    expect(reopened.blocks.find((block) => block.fieldDisplay)?.fieldDisplay).toEqual({
      kind: 'text',
      left: 'Figure 2 Registry architecture',
    })
    editor.destroy()
  })
})
