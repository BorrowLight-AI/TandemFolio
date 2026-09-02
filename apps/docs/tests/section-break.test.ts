import { Editor } from '@tiptap/core'
import { parseDocx, readSections, saveDocx } from '@genoffice/docx-engine'
import { describe, expect, it } from 'vitest'

import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'
import {
  insertSectionBreakAfterBlock,
  materializeSectionBreakStarts,
  resolveSectionBreakSource,
} from '../src/renderer/editor/section-actions'

const P = (text: string) =>
  `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`

describe('DOCX section-break persistence', () => {
  it('projects an undo-owned final section-break type through save and reopen', async () => {
    const parsed = await parseDocx(await buildDocx({ bodyXml: P('Before') + P('After') }))
    const sections = readSections(parsed)
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks, sections) as never,
    })

    expect(
      insertSectionBreakAfterBlock(
        editor,
        { afterBlockIndex: 0, startType: 'continuous' },
        (afterBlockIndex) =>
          resolveSectionBreakSource(editor, afterBlockIndex, sections, sections[0].settings),
      ),
    ).toMatchObject({ ok: true, insertedBlockIndex: 1, startType: 'continuous' })

    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    expect(plan.sectionBreakStarts).toEqual([
      { saveBlockIndex: 1, startType: 'continuous' },
    ])
    const materialized = materializeSectionBreakStarts(
      plan.saveBlocks,
      plan.sectionBreakStarts,
      parsed.blocks,
    )
    expect(materialized.trailingStartType).toBe('continuous')

    const reparsed = await parseDocx(
      await saveDocx(parsed, materialized.saveBlocks, {
        sectionStartType: materialized.trailingStartType ?? undefined,
      }),
    )
    expect(readSections(reparsed).map((section) => section.startType)).toEqual([
      'nextPage',
      'continuous',
    ])

    expect(editor.commands.undo()).toBe(true)
    const undone = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    expect(undone.sectionBreakStarts).toEqual([])
    const undoneSaved = await saveDocx(parsed, undone.saveBlocks)
    expect(readSections(await parseDocx(undoneSaved))).toHaveLength(1)
    editor.destroy()
  })

  it('projects consecutive inserted break types onto the following inserted and trailing sectPr', async () => {
    const parsed = await parseDocx(await buildDocx({ bodyXml: P('Only section') }))
    const sections = readSections(parsed)
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks, sections) as never,
    })
    const resolve = (afterBlockIndex: number) =>
      resolveSectionBreakSource(editor, afterBlockIndex, sections, sections[0].settings)

    expect(
      insertSectionBreakAfterBlock(
        editor,
        { afterBlockIndex: 0, startType: 'continuous' },
        resolve,
      ),
    ).toMatchObject({ ok: true })
    expect(
      insertSectionBreakAfterBlock(
        editor,
        { afterBlockIndex: 1, startType: 'oddPage' },
        resolve,
      ),
    ).toMatchObject({ ok: true })

    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    const materialized = materializeSectionBreakStarts(
      plan.saveBlocks,
      plan.sectionBreakStarts,
      parsed.blocks,
    )
    const reparsed = await parseDocx(
      await saveDocx(parsed, materialized.saveBlocks, {
        sectionStartType: materialized.trailingStartType ?? undefined,
      }),
    )
    expect(readSections(reparsed).map((section) => section.startType)).toEqual([
      'nextPage',
      'continuous',
      'oddPage',
    ])
    editor.destroy()
  })
})
