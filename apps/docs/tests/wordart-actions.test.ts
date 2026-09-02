import { Editor } from '@tiptap/core'
import { parseDocx, saveDocx } from '@genoffice/docx-engine'
import { describe, expect, it } from 'vitest'

import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { insertDocxWordArtAfterBlock } from '../src/renderer/editor/wordart-actions'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'

describe('deterministic DOCX WordArt insertion', () => {
  it('inserts explicit text/geometry/identity, supports native undo, and saves/reopens', async () => {
    const parsed = await parseDocx(
      await buildDocx({ bodyXml: '<w:p><w:r><w:t>Original body</w:t></w:r></w:p>' }),
    )
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks) as never,
    })
    const result = insertDocxWordArtAfterBlock(editor, {
      afterBlockIndex: 0,
      preset: 'white-orange',
      text: 'Registry WordArt',
      widthEmu: 2_700_000,
      heightEmu: 720_000,
      drawingId: 42,
    })
    expect(result).toMatchObject({
      ok: true,
      blockIndex: 1,
      preset: 'white-orange',
      drawingId: 42,
      changed: true,
    })
    expect(editor.state.doc.child(1).attrs.genXml).toContain('wp:docPr id="42"')

    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    const reopened = await parseDocx(await saveDocx(parsed, plan.saveBlocks))
    expect(reopened.blocks[1].textboxes?.[0].paras[0].runs[0].text).toBe('Registry WordArt')
    expect(reopened.blocks[1].textboxes?.[0].paras[0].runs[0].color).toBe('ED7D31')
    expect(reopened.blocks[1].textboxes?.[0]).toMatchObject({ widthPx: 283, heightPx: 76 })

    expect(editor.commands.undo()).toBe(true)
    expect(editor.state.doc.childCount).toBe(1)
    editor.destroy()
  })

  it('rejects unknown presets and duplicate drawing identities before mutation', async () => {
    const parsed = await parseDocx(
      await buildDocx({ bodyXml: '<w:p><w:r><w:t>Original body</w:t></w:r></w:p>' }),
    )
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks) as never,
    })
    const input = {
      afterBlockIndex: 0,
      preset: 'blue' as const,
      text: 'Registry WordArt',
      widthEmu: 2_700_000,
      heightEmu: 720_000,
      drawingId: 42,
    }
    expect(insertDocxWordArtAfterBlock(editor, input)).toMatchObject({ ok: true })
    expect(insertDocxWordArtAfterBlock(editor, { ...input, afterBlockIndex: 1 })).toMatchObject({
      ok: false,
      error: 'invalid_arguments',
    })
    expect(
      insertDocxWordArtAfterBlock(editor, { ...input, preset: 'unknown' as 'blue', drawingId: 43 }),
    ).toMatchObject({ ok: false, error: 'invalid_arguments' })
    expect(editor.state.doc.childCount).toBe(2)
    editor.destroy()
  })
})
