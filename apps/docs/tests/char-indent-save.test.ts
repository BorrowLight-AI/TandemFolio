import { afterEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { parseDocx, saveDocx } from '@genoffice/docx-engine'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { setParaAttrs } from '../src/renderer/components/ribbon-tabs'

const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
const STYLES =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ${NS}>` +
  '<w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="24"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr/></w:pPrDefault></w:docDefaults>' +
  '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/>' +
  '<w:pPr><w:ind w:firstLineChars="200"/></w:pPr></w:style>' +
  '</w:styles>'
const BODY =
  '<w:p><w:r><w:t>plain body paragraph</w:t></w:r></w:p>' +
  '<w:p><w:pPr><w:jc w:val="both"/></w:pPr><w:r><w:t>justified body paragraph</w:t></w:r></w:p>'

const editors: Editor[] = []
afterEach(() => {
  for (const editor of editors.splice(0)) editor.destroy()
})

async function openEditor() {
  const { editorExtensions } = await import('../src/renderer/editor/extensions')
  const parsed = await parseDocx(await buildDocx({ bodyXml: BODY, stylesXml: STYLES }))
  const editor = new Editor({ element: document.createElement('div'), extensions: editorExtensions })
  editors.push(editor)
  editor.commands.setContent(blocksToPmDoc(parsed.blocks) as never)
  return { editor, parsed }
}

function setFirstLine(editor: Editor, indentFirstLine: number | null) {
  setParaAttrs(editor, { indentFirstLine }, { from: 0, to: editor.state.doc.content.size })
}

describe('character-unit indents: saving a native indent edit', () => {
  it('parses the style indent into both paragraphs', async () => {
    const { parsed } = await openEditor()
    expect(parsed.blocks[0].rawPPr).toBeUndefined()
    expect(parsed.blocks[0].format).toEqual({
      indentFirstLine: 480,
      charIndents: { firstLine: 200 },
    })
    expect(parsed.blocks[1].format).toMatchObject({ align: 'justify', indentFirstLine: 480 })
  })

  it('saves a point indent with a character-unit cancel and survives reload', async () => {
    const { editor, parsed } = await openEditor()
    setFirstLine(editor, 720)
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    expect(plan.changedCount).toBe(2)
    const reparsed = await parseDocx(await saveDocx(parsed, plan.saveBlocks))

    expect(reparsed.internal.documentXml).toContain(
      '<w:pPr><w:ind w:firstLine="720" w:firstLineChars="0"/></w:pPr>',
    )
    expect(reparsed.internal.documentXml).toContain(
      '<w:pPr><w:ind w:firstLine="720" w:firstLineChars="0"/><w:jc w:val="both"/></w:pPr>',
    )
    expect(reparsed.blocks[0].format).toEqual({ indentFirstLine: 720 })
    expect(reparsed.blocks[1].format).toEqual({ align: 'justify', indentFirstLine: 720 })
  })

  it('writes a bare character-unit cancel when removing the indent', async () => {
    const { editor, parsed } = await openEditor()
    setFirstLine(editor, null)
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    const reparsed = await parseDocx(await saveDocx(parsed, plan.saveBlocks))

    expect(reparsed.internal.documentXml).toContain('<w:pPr><w:ind w:firstLineChars="0"/></w:pPr>')
    expect(reparsed.internal.documentXml).toContain(
      '<w:pPr><w:ind w:firstLineChars="0"/><w:jc w:val="both"/></w:pPr>',
    )
    expect(reparsed.blocks[0].format).toBeUndefined()
    expect(reparsed.blocks[1].format).toEqual({ align: 'justify' })
  })

  it('keeps character-unit indents through an unrelated paragraph edit', async () => {
    const { editor, parsed } = await openEditor()
    const second = editor.state.doc.child(0).nodeSize
    setParaAttrs(editor, { align: 'center' }, { from: second, to: editor.state.doc.content.size })
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    const reparsed = await parseDocx(await saveDocx(parsed, plan.saveBlocks))

    expect(reparsed.internal.documentXml).not.toContain('firstLineChars="0"')
    expect(reparsed.blocks[1].format).toEqual({
      align: 'center',
      indentFirstLine: 480,
      charIndents: { firstLine: 200 },
    })
  })
})
