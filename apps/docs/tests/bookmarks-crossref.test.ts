import { Editor } from '@tiptap/core'
import JSZip from 'jszip'
import { parseDocx, saveDocx } from '@genoffice/docx-engine'
import { describe, expect, it } from 'vitest'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { setDocxBookmark } from '../src/renderer/editor/bookmark-actions'
import { insertDocxCrossReference } from '../src/renderer/editor/field-actions'

const BODY =
  '<w:p><w:bookmarkStart w:id="1" w:name="Conclusion"/><w:bookmarkEnd w:id="1"/>' +
  '<w:r><w:t>Conclusion paragraph.</w:t></w:r></w:p>' +
  '<w:p><w:r><w:t>Plain paragraph.</w:t></w:r></w:p>'

async function open() {
  const source = await buildDocx({ bodyXml: BODY })
  const parsed = await parseDocx(source)
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
  return { editor, parsed, source }
}

async function docXmlOf(bytes: Uint8Array): Promise<string> {
  const zip = await JSZip.loadAsync(bytes)
  return zip.file('word/document.xml')!.async('string')
}

describe('bookmarks & cross-references in the editor', () => {
  it('renders data-bookmarks and keeps an untouched doc byte-identical', async () => {
    const { editor, parsed, source } = await open()
    expect(editor.view.dom.querySelector('[data-bookmarks~="Conclusion"]')?.textContent).toBe(
      'Conclusion paragraph.',
    )
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    expect(plan.changedCount).toBe(0)
    expect(await saveDocx(parsed, plan.saveBlocks)).toEqual(source)
    editor.destroy()
  })

  it('the shared bookmark action supports Undo and save/reopen', async () => {
    const { editor, parsed } = await open()
    expect(
      setDocxBookmark(editor, { blockIndex: 1, name: 'NewBookmark', enabled: true }),
    ).toMatchObject({ ok: true, changed: true })
    expect(editor.state.doc.child(1).attrs.bookmarks).toEqual(['NewBookmark'])
    expect(editor.commands.undo()).toBe(true)
    expect(editor.state.doc.child(1).attrs.bookmarks).toBeNull()
    expect(editor.commands.redo()).toBe(true)
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    expect(plan.changedCount).toBe(1)
    const saved = await saveDocx(parsed, plan.saveBlocks)
    const reparsed = await parseDocx(saved)
    expect(reparsed.blocks[1].bookmarks).toEqual(['NewBookmark'])
    expect(reparsed.blocks[0].bookmarks).toEqual(['Conclusion'])
    editor.destroy()
  })

  it('inserting a cross-reference saves a REF field pointing at the bookmark', async () => {
    const { editor, parsed } = await open()
    const position = editor.state.doc.content.size - 1
    expect(
      insertDocxCrossReference(editor, {
        range: { from: position, to: position },
        bookmarkName: 'Conclusion',
        displayText: 'Conclusion paragraph.',
      }),
    ).toMatchObject({ ok: true, changed: true })
    expect(editor.view.dom.querySelector('span[data-ref-field="Conclusion"]')).toBeTruthy()
    expect(editor.commands.undo()).toBe(true)
    expect(editor.view.dom.querySelector('span[data-ref-field="Conclusion"]')).toBeNull()
    expect(editor.commands.redo()).toBe(true)
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    const saved = await saveDocx(parsed, plan.saveBlocks)
    expect(await docXmlOf(saved)).toContain('REF Conclusion \\h')
    const reparsed = await parseDocx(saved)
    const refRun = reparsed.blocks[1].runs!.find((r) => r.refField === 'Conclusion')
    expect(refRun?.text).toBe('Conclusion paragraph.')
    editor.destroy()
  })
})
