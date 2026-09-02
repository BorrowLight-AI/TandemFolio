import { Editor } from '@tiptap/core'
import { parseDocx, saveDocx } from '@genoffice/docx-engine'
import { describe, expect, it } from 'vitest'

import { editorExtensions } from '../src/renderer/editor/extensions'
import { insertDocxCoverPage } from '../src/renderer/editor/cover-pages'
import { pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'

async function openEditor() {
  const parsed = await parseDocx(
    await buildDocx({ bodyXml: '<w:p><w:r><w:t>Original body</w:t></w:r></w:p>' }),
  )
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: {
      type: 'doc',
      content: [
        {
          type: 'docParagraph',
          attrs: { docxIndex: 0 },
          content: [{ type: 'text', text: 'Original body' }],
        },
      ],
    },
  })
  return { editor, parsed }
}

describe('deterministic DOCX cover-page insertion', () => {
  it('uses explicit replayable content, one native undo, and save/reopen projection', async () => {
    const { editor, parsed } = await openEditor()
    const result = insertDocxCoverPage(editor, {
      preset: 'classic',
      title: 'Registry Architecture',
      subtitle: 'Typed operations',
      author: 'Agent',
      date: '30 August 2026',
      year: 2026,
    })
    expect(result).toMatchObject({
      ok: true,
      preset: 'classic',
      insertedBlocks: 6,
      changed: true,
    })
    expect(editor.getText()).toContain('Registry Architecture')
    expect(editor.getText()).toContain('30 August 2026')
    expect(editor.getText()).toContain('Original body')

    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    const reopened = await parseDocx(await saveDocx(parsed, plan.saveBlocks))
    expect(
      reopened.blocks.map((block) => block.runs?.map((run) => run.text).join('') ?? '').join('\n'),
    ).toContain('Registry Architecture')

    expect(editor.commands.undo()).toBe(true)
    expect(editor.getText()).toBe('Original body')
    editor.destroy()
  })

  it('rejects unknown presets and malformed explicit content before mutation', async () => {
    const { editor } = await openEditor()
    expect(
      insertDocxCoverPage(editor, {
        preset: 'unknown',
        title: '',
        subtitle: 'Typed operations',
        author: 'Agent',
        date: '30 August 2026',
        year: 2026,
      }),
    ).toMatchObject({ ok: false, error: 'invalid_arguments' })
    expect(editor.getText()).toBe('Original body')
    editor.destroy()
  })
})
