import { Editor } from '@tiptap/core'
import { applyPageNumType, parseDocx, readSections, saveDocx } from '@genoffice/docx-engine'
import { describe, expect, it } from 'vitest'

import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'
import {
  applySectionSettingsOverrides,
  setDocumentDifferentOddEvenPages,
  setHeaderFooterText,
  setHeaderFooterPageNumber,
  setHeaderFooterParagraphs,
  setSectionColumns,
  setSectionDifferentFirstPage,
  setSectionMargins,
  setSectionOrientation,
  setSectionPageNumbering,
} from '../src/renderer/editor/section-layout'

const P = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`

describe('DOCX section-layout persistence', () => {
  it('persists an Undo-owned final-section orientation and reopens it exactly', async () => {
    const parsed = await parseDocx(await buildDocx({ bodyXml: P('Portrait') }))
    const sections = readSections(parsed)
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks, sections) as never,
    })

    expect(
      setSectionOrientation(editor, { sectionIndex: 0, orientation: 'landscape' }, sections),
    ).toEqual({ ok: true, changed: true })
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    expect(plan.sectionSettingsOverrides).toHaveLength(1)
    const effective = applySectionSettingsOverrides(sections, plan.sectionSettingsOverrides)
    expect(effective[0].settings).toMatchObject({
      orientation: 'landscape',
      pageWidth: sections[0].settings.pageHeight,
      pageHeight: sections[0].settings.pageWidth,
    })

    const reparsed = await parseDocx(
      await saveDocx(parsed, plan.saveBlocks, { section: effective[0].settings }),
    )
    expect(readSections(reparsed)[0].settings).toMatchObject({
      orientation: 'landscape',
      pageWidth: sections[0].settings.pageHeight,
      pageHeight: sections[0].settings.pageWidth,
    })

    expect(editor.commands.undo()).toBe(true)
    const undone = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    expect(undone.sectionSettingsOverrides).toEqual([])
    editor.destroy()
  })

  it('persists exact margins through the same section override projection', async () => {
    const parsed = await parseDocx(await buildDocx({ bodyXml: P('Margins') }))
    const sections = readSections(parsed)
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks, sections) as never,
    })
    expect(
      setSectionMargins(
        editor,
        {
          sectionIndex: 0,
          margins: {
            topTwips: 720,
            rightTwips: 1080,
            bottomTwips: 720,
            leftTwips: 1080,
          },
        },
        sections,
      ),
    ).toEqual({ ok: true, changed: true })
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    const effective = applySectionSettingsOverrides(sections, plan.sectionSettingsOverrides)
    const reopened = readSections(
      await parseDocx(await saveDocx(parsed, plan.saveBlocks, { section: effective[0].settings })),
    )[0]
    expect(reopened.settings).toMatchObject({
      marginTop: 720,
      marginRight: 1080,
      marginBottom: 720,
      marginLeft: 1080,
    })
    editor.destroy()
  })

  it('persists exact column count and spacing through save and reopen', async () => {
    const parsed = await parseDocx(await buildDocx({ bodyXml: P('Columns') }))
    const sections = readSections(parsed)
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks, sections) as never,
    })
    expect(
      setSectionColumns(
        editor,
        { sectionIndex: 0, count: 3, spacingTwips: 480 },
        sections,
      ),
    ).toEqual({ ok: true, changed: true })
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    const effective = applySectionSettingsOverrides(sections, plan.sectionSettingsOverrides)
    const reopened = readSections(
      await parseDocx(await saveDocx(parsed, plan.saveBlocks, { section: effective[0].settings })),
    )[0]
    expect(reopened.settings.columns).toBe(3)
    expect(reopened.settings.colSpace).toBe(480)
    editor.destroy()
  })

  it('persists different-first-page state through the section journal', async () => {
    const parsed = await parseDocx(await buildDocx({ bodyXml: P('First-page header') }))
    const sections = readSections(parsed)
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks, sections) as never,
    })
    expect(
      setSectionDifferentFirstPage(editor, { sectionIndex: 0, enabled: true }, sections),
    ).toEqual({ ok: true, changed: true })
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    const effective = applySectionSettingsOverrides(sections, plan.sectionSettingsOverrides)
    expect(effective[0].titlePg).toBe(true)
    const reopened = readSections(
      await parseDocx(await saveDocx(parsed, plan.saveBlocks, { titlePg: effective[0].titlePg })),
    )[0]
    expect(reopened.titlePg).toBe(true)
    editor.destroy()
  })

  it('persists different odd/even pages through the document journal and reopens it exactly', async () => {
    const parsed = await parseDocx(await buildDocx({ bodyXml: P('Odd/even headers') }))
    const sections = readSections(parsed)
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks, sections) as never,
    })

    expect(
      setDocumentDifferentOddEvenPages(editor, { enabled: true }, sections, false),
    ).toEqual({ ok: true, changed: true })
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    expect(plan.sectionSettingsOverrides[0]?.evenOddHeaders).toBe(true)

    const reopened = await parseDocx(
      await saveDocx(parsed, plan.saveBlocks, { evenAndOddHeaders: true }),
    )
    expect(reopened.evenAndOddHeaders).toBe(true)

    expect(editor.commands.undo()).toBe(true)
    expect(
      pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks).sectionSettingsOverrides,
    ).toEqual([])
    editor.destroy()
  })

  it('persists page numbering for non-final and final sections and undoes the native history group', async () => {
    const breakParagraph =
      '<w:p><w:pPr><w:sectPr><w:pgSz w:w="11906" w:h="16838"/>' +
      '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>' +
      '</w:sectPr></w:pPr><w:r><w:t>Section one</w:t></w:r></w:p>'
    const parsed = await parseDocx(
      await buildDocx({ bodyXml: breakParagraph + P('Section two') }),
    )
    const sections = readSections(parsed)
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks, sections) as never,
    })

    expect(
      setSectionPageNumbering(
        editor,
        { sectionIndex: 0, format: 'upperRoman', start: 3 },
        sections,
      ),
    ).toEqual({ ok: true, changed: true })
    expect(
      setSectionPageNumbering(
        editor,
        { sectionIndex: 1, format: 'lowerLetter', start: null },
        sections,
      ),
    ).toEqual({ ok: true, changed: true })

    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    const effective = applySectionSettingsOverrides(sections, plan.sectionSettingsOverrides)
    expect(effective[0]).toMatchObject({ pageNumberFmt: 'upperRoman', pageNumberStart: 3 })
    expect(effective[1]).toMatchObject({ pageNumberFmt: 'lowerLetter' })
    expect(effective[1].pageNumberStart).toBeUndefined()

    const first = sections[0]
    const firstBlock = parsed.blocks.find((block) => block.docxIndex === first.lastBlockIndex)!
    const saveBlocks = plan.saveBlocks.map((block) =>
      block.kind === 'original' && block.docxIndex === first.lastBlockIndex
        ? {
            kind: 'xml' as const,
            docxIndex: block.docxIndex,
            xml: firstBlock.originalXml!.replace(
              first.sectPrXml,
              applyPageNumType(first.sectPrXml, 'upperRoman', 3),
            ),
          }
        : block,
    )
    const reopened = readSections(
      await parseDocx(
        await saveDocx(parsed, saveBlocks, {
          pgNumType: { fmt: 'lowerLetter', start: undefined },
        }),
      ),
    )
    expect(reopened[0]).toMatchObject({ pageNumberFmt: 'upperRoman', pageNumberStart: 3 })
    expect(reopened[1].pageNumberFmt).toBe('lowerLetter')
    expect(reopened[1].pageNumberStart).toBeUndefined()

    expect(editor.commands.undo()).toBe(true)
    expect(
      pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks).sectionSettingsOverrides,
    ).toEqual([])
    editor.destroy()
  })

  it('persists an explicit header/footer text journal entry and reopens it exactly', async () => {
    const parsed = await parseDocx(await buildDocx({ bodyXml: P('Header body') }))
    const sections = readSections(parsed)
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks, sections) as never,
    })

    expect(
      setHeaderFooterText(
        editor,
        {
          sectionIndex: 0,
          kind: 'header',
          variant: 'first',
          text: 'Confidential',
        },
        sections,
      ),
    ).toEqual({ ok: true, changed: true })
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    const edit = plan.sectionSettingsOverrides[0]?.headerFooterEdits?.[0]
    expect(edit).toEqual({
      kind: 'header',
      variant: 'first',
      value: { text: 'Confidential' },
    })
    const reopened = await parseDocx(
      await saveDocx(parsed, plan.saveBlocks, { headerFirst: edit?.value }),
    )
    expect(reopened.headerFirst?.text).toContain('Confidential')

    expect(editor.commands.undo()).toBe(true)
    expect(
      pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks).sectionSettingsOverrides,
    ).toEqual([])
    editor.destroy()
  })

  it('persists canonical PAGE field placement and reopens it as a page number', async () => {
    const parsed = await parseDocx(await buildDocx({ bodyXml: P('Page field body') }))
    const sections = readSections(parsed)
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks, sections) as never,
    })

    expect(
      setHeaderFooterPageNumber(
        editor,
        {
          sectionIndex: 0,
          kind: 'footer',
          variant: 'default',
          enabled: true,
          alignment: 'right',
        },
        sections,
      ),
    ).toEqual({ ok: true, changed: true })
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    const edit = plan.sectionSettingsOverrides[0]?.headerFooterEdits?.[0]
    const reopened = await parseDocx(
      await saveDocx(parsed, plan.saveBlocks, { footer: edit?.value }),
    )
    expect(reopened.footerHasPageNumber).toBe(true)
    expect(reopened.footerParas?.[0]?.align).toBe('right')

    expect(editor.commands.undo()).toBe(true)
    expect(
      pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks).sectionSettingsOverrides,
    ).toEqual([])
    editor.destroy()
  })

  it('persists bounded rich header/footer paragraphs and field tokens', async () => {
    const parsed = await parseDocx(await buildDocx({ bodyXml: P('Rich footer body') }))
    const sections = readSections(parsed)
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks, sections) as never,
    })

    expect(
      setHeaderFooterParagraphs(
        editor,
        {
          sectionIndex: 0,
          kind: 'footer',
          variant: 'default',
          paragraphs: [
            {
              alignment: 'right',
              segments: [
                { type: 'text', text: 'Page ', bold: true, color: 'aabbcc' },
                { type: 'page', text: '' },
                { type: 'text', text: ' / ' },
                { type: 'total_pages', text: '' },
              ],
            },
          ],
        },
        sections,
      ),
    ).toEqual({ ok: true, changed: true })
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    const value = plan.sectionSettingsOverrides[0]?.headerFooterEdits?.[0]?.value
    const reopened = await parseDocx(await saveDocx(parsed, plan.saveBlocks, { footer: value }))
    expect(reopened.footerHasPageNumber).toBe(true)
    expect(reopened.footerParas?.[0]?.align).toBe('right')
    expect(reopened.footerParas?.[0]?.runs[0]).toMatchObject({
      text: 'Page ',
      bold: true,
      color: 'AABBCC',
    })

    expect(
      setHeaderFooterParagraphs(
        editor,
        {
          sectionIndex: 0,
          kind: 'header',
          variant: 'default',
          paragraphs: [
            { alignment: 'center', segments: [{ type: 'page', text: 'not-empty' }] },
          ],
        },
        sections,
      ),
    ).toMatchObject({ ok: false, error: 'invalid_arguments' })
    editor.destroy()
  })
})
