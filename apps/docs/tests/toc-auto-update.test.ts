/**
 * TOC page-number persistence: an editor-generated TOC line whose fieldDisplay
 * page was auto-refreshed after repagination must save the refreshed number
 * (the genXml still carries the number from generation time).
 */
import { describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { generateTocFieldXml, parseDocx, saveDocx } from '@genoffice/docx-engine'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { applyTocPageDisplays } from '../src/renderer/editor/toc-refresh'
import { createDocxTocFieldNodes, refreshDocxToc } from '../src/renderer/editor/toc-actions'

async function openBlankDoc() {
  const source = await buildDocx({ bodyXml: '<w:p><w:r><w:t>Body text</w:t></w:r></w:p>' })
  const parsed = await parseDocx(source)
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
  return { editor, parsed }
}

describe('TOC auto page refresh persistence', () => {
  it('saves the refreshed page number of a generated TOC line', async () => {
    const { editor, parsed } = await openBlankDoc()
    const [xml] = generateTocFieldXml([{ level: 1, text: 'Chapter One', pageNo: 1 }])
    editor.commands.insertContentAt(0, {
      type: 'docProtected',
      attrs: {
        docxIndex: null,
        blockType: 'passthrough',
        label: 'TOC line',
        genXml: xml,
        // repagination auto-refresh moved the heading to page 3
        fieldDisplay: { kind: 'tocLine', left: 'Chapter One', right: '3', level: 1 },
      },
    })
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    const xmlBlock = plan.saveBlocks.find(
      (b) => b.kind === 'xml' && b.xml.includes('Chapter One'),
    ) as { xml: string } | undefined
    expect(xmlBlock).toBeDefined()
    expect(xmlBlock!.xml).toMatch(/<w:t[^>]*>3<\/w:t>/)
    expect(xmlBlock!.xml).not.toMatch(/<w:t[^>]*>1<\/w:t>/)
  })

  it('refreshes a reopened TOC through native Undo and saves the regenerated field', async () => {
    const { editor, parsed } = await openBlankDoc()
    editor.commands.insertContentAt(
      0,
      createDocxTocFieldNodes([{ level: 1, text: 'Old', pageNo: 1 }]) as never,
    )
    const initialPlan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    const withToc = await saveDocx(parsed, initialPlan.saveBlocks)
    editor.destroy()

    const reopenedSource = await parseDocx(withToc)
    const reopenedEditor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(reopenedSource.blocks) as never,
    })
    expect(
      refreshDocxToc(reopenedEditor, reopenedSource.blocks, {
        tocBlockIndex: 0,
        entries: [{ level: 1, text: 'New', pageNumber: 3 }],
      }),
    ).toEqual({ ok: true, entries: 1, replacedBlocks: 1 })
    expect(reopenedEditor.commands.undo()).toBe(true)
    expect(reopenedEditor.state.doc.child(0).attrs.fieldDisplay).toMatchObject({ left: 'Old' })
    expect(reopenedEditor.commands.redo()).toBe(true)

    const refreshedPlan = pmDocToSavePlan(reopenedEditor.getJSON() as PmNode, reopenedSource.blocks)
    const finalDoc = await parseDocx(await saveDocx(reopenedSource, refreshedPlan.saveBlocks))
    expect(finalDoc.blocks[0].fieldDisplay).toMatchObject({
      kind: 'tocLine',
      left: 'New',
      right: '3',
      level: 1,
    })
    reopenedEditor.destroy()
  })
})

function tocLineNode(title: string, right = '') {
  return {
    type: 'docProtected',
    attrs: {
      docxIndex: null,
      blockType: 'passthrough',
      label: 'TOC line',
      genXml: generateTocFieldXml([{ level: 1, text: title, pageNo: 1 }])[0],
      fieldDisplay: { kind: 'tocLine', left: title, right, level: 1 },
    },
  }
}

function tocRights(editor: Editor): string[] {
  const rights: string[] = []
  editor.state.doc.forEach((node) => {
    const field = node.attrs.fieldDisplay as { kind?: string; right?: string } | null
    if (field?.kind === 'tocLine') rights.push(field.right ?? '')
  })
  return rights
}

describe('applyTocPageDisplays', () => {
  it('keeps the formatted page string (no Arabic coercion)', async () => {
    const { editor } = await openBlankDoc()
    editor.commands.insertContentAt(0, tocLineNode('Preface', '2'))
    const tr = editor.state.tr
    const changed = applyTocPageDisplays(
      editor.state.doc,
      tr,
      [{ text: 'Preface', level: 1, pos: 0 }],
      ['iv'],
    )
    expect(changed).toBe(true)
    editor.view.dispatch(tr)
    expect(tocRights(editor)).toEqual(['iv'])
  })

  it('assigns duplicate titles their own pages in document order', async () => {
    const { editor } = await openBlankDoc()
    editor.commands.insertContentAt(0, [
      tocLineNode('Summary', '1'),
      tocLineNode('Summary', '1'),
    ] as never)
    const headings = [
      { text: 'Summary', level: 1, pos: 0 },
      { text: 'Summary', level: 1, pos: 10 },
    ]
    const tr = editor.state.tr
    expect(applyTocPageDisplays(editor.state.doc, tr, headings, ['2', '7'])).toBe(true)
    editor.view.dispatch(tr)
    expect(tocRights(editor)).toEqual(['2', '7'])
  })

  it('skips unmeasured headings and reports no change when values match', async () => {
    const { editor } = await openBlankDoc()
    editor.commands.insertContentAt(0, tocLineNode('Chapter', '5'))
    const tr = editor.state.tr
    const changed = applyTocPageDisplays(
      editor.state.doc,
      tr,
      [
        { text: 'Chapter', level: 1, pos: 0 },
        { text: 'Chapter', level: 1, pos: 10 },
      ],
      [undefined, '5'],
    )
    expect(changed).toBe(false)
  })
})
