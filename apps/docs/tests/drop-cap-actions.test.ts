import { Editor } from '@tiptap/core'
import { parseDocx, saveDocx } from '@genoffice/docx-engine'
import { describe, expect, it } from 'vitest'

import { setDocxDropCap } from '../src/renderer/editor/drop-cap-actions'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'

describe('explicit DOCX drop-cap state', () => {
  it('sets one block, reports no-op, supports native undo, and saves/reopens', async () => {
    const parsed = await parseDocx(
      await buildDocx({ bodyXml: '<w:p><w:r><w:t>Alpha paragraph</w:t></w:r></w:p>' }),
    )
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks) as never,
    })
    expect(setDocxDropCap(editor, { blockIndex: 0, mode: 'margin', lines: 3 })).toMatchObject({
      ok: true,
      blockIndex: 0,
      mode: 'margin',
      lines: 3,
      changed: true,
    })
    expect(setDocxDropCap(editor, { blockIndex: 0, mode: 'margin', lines: 3 })).toMatchObject({
      ok: true,
      changed: false,
    })

    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    const reopened = await parseDocx(await saveDocx(parsed, plan.saveBlocks))
    expect(reopened.blocks[0].format?.dropCap).toEqual({ type: 'margin', lines: 3 })

    expect(editor.commands.undo()).toBe(true)
    expect(editor.state.doc.child(0).attrs.dropCap).toBeNull()
    editor.destroy()
  })

  it('requires null lines for none and bounded lines for visible modes', async () => {
    const parsed = await parseDocx(
      await buildDocx({ bodyXml: '<w:p><w:r><w:t>Alpha paragraph</w:t></w:r></w:p>' }),
    )
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks) as never,
    })
    expect(setDocxDropCap(editor, { blockIndex: 0, mode: 'none', lines: 3 })).toMatchObject({
      ok: false,
      error: 'invalid_arguments',
    })
    expect(setDocxDropCap(editor, { blockIndex: 0, mode: 'drop', lines: 1 })).toMatchObject({
      ok: false,
      error: 'invalid_arguments',
    })
    editor.destroy()
  })
})
