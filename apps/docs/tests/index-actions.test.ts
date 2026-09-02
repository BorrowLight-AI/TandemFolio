import { Editor } from '@tiptap/core'
import { parseDocx, saveDocx } from '@genoffice/docx-engine'
import { describe, expect, it } from 'vitest'

import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { insertDocxIndex, markDocxIndexEntry } from '../src/renderer/editor/index-actions'

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

describe('DOCX index actions', () => {
  it('marks one exact inline range as one native Undo unit', () => {
    const editor = createEditor()
    expect(markDocxIndexEntry(editor, { range: { from: 1, to: 6 }, term: 'Alpha' })).toEqual({
      ok: true,
      from: 1,
      to: 6,
      term: 'Alpha',
      changed: true,
    })
    expect(editor.state.doc.child(0).lastChild?.type.name).toBe('docXeMark')
    expect(editor.state.doc.child(0).lastChild?.attrs.term).toBe('Alpha')
    expect(editor.commands.undo()).toBe(true)
    expect(editor.state.doc.child(0).childCount).toBe(1)
    editor.destroy()
  })

  it('rejects a term that cannot round-trip without mutation', () => {
    const editor = createEditor()
    expect(
      markDocxIndexEntry(editor, { range: { from: 1, to: 6 }, term: '"Alpha"' }),
    ).toMatchObject({ ok: false, error: 'invalid_arguments' })
    expect(editor.state.doc.child(0).childCount).toBe(1)
    editor.destroy()
  })

  it('persists the marked term through save and reopen', async () => {
    const parsed = await parseDocx(
      await buildDocx({ bodyXml: '<w:p><w:r><w:t>Alpha</w:t></w:r></w:p>' }),
    )
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks) as never,
    })
    expect(markDocxIndexEntry(editor, { range: { from: 1, to: 6 }, term: 'Alpha' })).toMatchObject({
      ok: true,
    })
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    const reopened = await parseDocx(await saveDocx(parsed, plan.saveBlocks))
    expect(reopened.blocks[0].runs?.find((run) => run.xeTerm)?.xeTerm).toBe('Alpha')
    editor.destroy()
  })

  it('inserts one explicit deduplicated index after a stable boundary as one Undo unit', () => {
    const editor = createEditor()
    expect(
      insertDocxIndex(editor, {
        afterBlockIndex: 0,
        label: 'Index',
        terms: ['Beta', 'Alpha', 'Alpha'],
      }),
    ).toEqual({ ok: true, afterBlockIndex: 0, entries: 2, insertedBlocks: 2 })
    expect(editor.state.doc.childCount).toBe(3)
    expect(editor.state.doc.child(1).attrs.fieldDisplay).toMatchObject({ left: 'Alpha' })
    expect(editor.state.doc.child(2).attrs.fieldDisplay).toMatchObject({ left: 'Beta' })
    expect(editor.commands.undo()).toBe(true)
    expect(editor.state.doc.childCount).toBe(1)
    editor.destroy()
  })

  it('rejects an empty final index without mutation', () => {
    const editor = createEditor()
    expect(
      insertDocxIndex(editor, { afterBlockIndex: 0, label: 'Index', terms: [] }),
    ).toMatchObject({ ok: false, error: 'invalid_arguments' })
    expect(editor.state.doc.childCount).toBe(1)
    editor.destroy()
  })

  it('persists the explicit index field through save and reopen', async () => {
    const parsed = await parseDocx(
      await buildDocx({ bodyXml: '<w:p><w:r><w:t>Alpha</w:t></w:r></w:p>' }),
    )
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks) as never,
    })
    expect(
      insertDocxIndex(editor, {
        afterBlockIndex: 0,
        label: 'Index',
        terms: ['Beta', 'Alpha'],
      }),
    ).toMatchObject({ ok: true })
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    const reopened = await parseDocx(await saveDocx(parsed, plan.saveBlocks))
    expect(
      reopened.blocks
        .map((block) => block.fieldDisplay?.left)
        .filter((text): text is string => typeof text === 'string'),
    ).toEqual(['Alpha', 'Beta'])
    editor.destroy()
  })
})
