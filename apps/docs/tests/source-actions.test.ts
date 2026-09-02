import { Editor } from '@tiptap/core'
import { parseDocx, saveDocx, type SourceInfo } from '@genoffice/docx-engine'
import { describe, expect, it } from 'vitest'

import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'
import {
  collectDocxSources,
  insertDocxBibliography,
  insertDocxCitation,
  upsertDocxSource,
} from '../src/renderer/editor/source-actions'

const source: SourceInfo = {
  tag: 'Wang2026',
  type: 'Book',
  author: 'Wang, Wei',
  title: 'Registry Architecture',
  year: '2026',
  publisher: 'Lite Press',
}

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

describe('DOCX source actions', () => {
  it('upserts a stable source through native Undo/Redo', () => {
    const editor = createEditor()
    expect(upsertDocxSource(editor, [], { source })).toEqual({
      ok: true,
      tag: 'Wang2026',
      created: true,
      changed: true,
      sources: [source],
    })
    expect(collectDocxSources(editor, [])).toEqual([source])
    expect(editor.commands.undo()).toBe(true)
    expect(collectDocxSources(editor, [])).toEqual([])
    expect(editor.commands.redo()).toBe(true)
    expect(collectDocxSources(editor, [])).toEqual([source])
    editor.destroy()
  })

  it('updates the same tag and restores the parsed baseline through Undo', () => {
    const editor = createEditor()
    const updated = { ...source, title: 'Updated Architecture' }
    expect(upsertDocxSource(editor, [source], { source: updated })).toMatchObject({
      ok: true,
      created: false,
      changed: true,
      sources: [updated],
    })
    expect(collectDocxSources(editor, [source])).toEqual([updated])
    expect(editor.commands.undo()).toBe(true)
    expect(collectDocxSources(editor, [source])).toEqual([source])
    editor.destroy()
  })

  it('rejects an empty title without mutation', () => {
    const editor = createEditor()
    expect(upsertDocxSource(editor, [], { source: { ...source, title: '' } })).toMatchObject({
      ok: false,
      error: 'invalid_arguments',
    })
    expect(collectDocxSources(editor, [])).toEqual([])
    editor.destroy()
  })

  it('persists an upserted source through save and reopen', async () => {
    const parsed = await parseDocx(
      await buildDocx({ bodyXml: '<w:p><w:r><w:t>Alpha</w:t></w:r></w:p>' }),
    )
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks) as never,
    })
    expect(upsertDocxSource(editor, parsed.sources, { source })).toMatchObject({ ok: true })
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    const reopened = await parseDocx(
      await saveDocx(parsed, plan.saveBlocks, {
        sources: collectDocxSources(editor, parsed.sources),
      }),
    )
    expect(reopened.sources).toEqual([source])
    editor.destroy()
  })

  it('inserts one source-backed citation at an exact range and undoes it', () => {
    const editor = createEditor()
    expect(
      insertDocxCitation(editor, [source], {
        range: { from: 6, to: 6 },
        sourceTag: 'Wang2026',
        displayText: '(Wang, Wei, 2026)',
      }),
    ).toEqual({
      ok: true,
      from: 6,
      to: 6,
      sourceTag: 'Wang2026',
      changed: true,
    })
    expect(editor.state.doc.textContent).toBe('Alpha(Wang, Wei, 2026)')
    expect(editor.commands.undo()).toBe(true)
    expect(editor.state.doc.textContent).toBe('Alpha')
    editor.destroy()
  })

  it('rejects a missing citation source without mutation', () => {
    const editor = createEditor()
    expect(
      insertDocxCitation(editor, [], {
        range: { from: 1, to: 1 },
        sourceTag: 'Missing',
        displayText: '(Missing)',
      }),
    ).toMatchObject({ ok: false, error: 'invalid_arguments' })
    expect(editor.state.doc.textContent).toBe('Alpha')
    editor.destroy()
  })

  it('persists an inserted citation through save and reopen', async () => {
    const parsed = await parseDocx(
      await buildDocx({ bodyXml: '<w:p><w:r><w:t>Alpha</w:t></w:r></w:p>' }),
    )
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks) as never,
    })
    expect(
      insertDocxCitation(editor, [source], {
        range: { from: 6, to: 6 },
        sourceTag: source.tag,
        displayText: '(Wang, Wei, 2026)',
      }),
    ).toMatchObject({ ok: true })
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    const reopened = await parseDocx(await saveDocx(parsed, plan.saveBlocks))
    expect(reopened.blocks[0].runs?.map((run) => run.text).join('')).toBe('Alpha(Wang, Wei, 2026)')
    editor.destroy()
  })

  it('inserts explicit bibliography blocks at a stable boundary as one Undo unit', () => {
    const editor = createEditor()
    expect(
      insertDocxBibliography(editor, [source], {
        afterBlockIndex: 0,
        heading: 'Bibliography',
        entries: [{ sourceTag: source.tag, text: 'Wang, Wei. (2026). Registry Architecture.' }],
      }),
    ).toEqual({ ok: true, afterBlockIndex: 0, entries: 1, insertedBlocks: 2 })
    expect(editor.state.doc.childCount).toBe(3)
    expect(editor.state.doc.child(1).type.name).toBe('docHeading')
    expect(editor.state.doc.child(1).textContent).toBe('Bibliography')
    expect(editor.state.doc.child(2).textContent).toBe(
      'Wang, Wei. (2026). Registry Architecture.',
    )
    expect(editor.commands.undo()).toBe(true)
    expect(editor.state.doc.childCount).toBe(1)
    editor.destroy()
  })

  it('rejects bibliography entries with unknown source tags without mutation', () => {
    const editor = createEditor()
    expect(
      insertDocxBibliography(editor, [source], {
        afterBlockIndex: 0,
        heading: 'Bibliography',
        entries: [{ sourceTag: 'Missing', text: 'Missing.' }],
      }),
    ).toMatchObject({ ok: false, error: 'invalid_arguments' })
    expect(editor.state.doc.childCount).toBe(1)
    editor.destroy()
  })

  it('persists explicit bibliography blocks through save and reopen', async () => {
    const parsed = await parseDocx(
      await buildDocx({ bodyXml: '<w:p><w:r><w:t>Alpha</w:t></w:r></w:p>' }),
    )
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks) as never,
    })
    expect(
      insertDocxBibliography(editor, [source], {
        afterBlockIndex: 0,
        heading: 'Bibliography',
        entries: [{ sourceTag: source.tag, text: 'Wang, Wei. (2026). Registry Architecture.' }],
      }),
    ).toMatchObject({ ok: true })
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    const reopened = await parseDocx(await saveDocx(parsed, plan.saveBlocks))
    expect(
      reopened.blocks
        .map((block) => block.runs?.map((run) => run.text).join(''))
        .filter((text): text is string => typeof text === 'string'),
    ).toEqual(['Alpha', 'Bibliography', 'Wang, Wei. (2026). Registry Architecture.'])
    editor.destroy()
  })
})
