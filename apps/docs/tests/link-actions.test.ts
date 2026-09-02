import { Editor } from '@tiptap/core'
import { parseDocx, saveDocx } from '@genoffice/docx-engine'
import { describe, expect, it } from 'vitest'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { setDocxTextLink } from '../src/renderer/editor/link-actions'

describe('DOCX exact-range links', () => {
  it('persists a new external relationship and reopens the linked run', async () => {
    const parsed = await parseDocx(
      await buildDocx({ bodyXml: '<w:p><w:r><w:t>Link me</w:t></w:r></w:p>' }),
    )
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks) as never,
    })
    expect(
      setDocxTextLink(editor, {
        range: { from: 1, to: 8 },
        href: 'https://example.com/docs',
        text: null,
      }),
    ).toEqual({ ok: true, changed: true, from: 1, to: 8 })

    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    const reopened = await parseDocx(await saveDocx(parsed, plan.saveBlocks))
    expect(reopened.blocks[0].runs?.[0]).toMatchObject({
      text: 'Link me',
      link: { href: 'https://example.com/docs' },
    })
    editor.destroy()
  })
})
