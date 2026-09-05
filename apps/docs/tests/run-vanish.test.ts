// Modified by TandemFolio contributors: adapt upstream hidden-run renderer coverage (2026-09-05).
import { Editor } from '@tiptap/core'
import { describe, expect, it } from 'vitest'
import type { Block } from '@genoffice/docx-engine'
import { blocksToPmDoc } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'

describe('hidden DOCX runs in the mounted editor model', () => {
  it('renders w:vanish text as hidden content while retaining the text node', () => {
    const block: Block = {
      id: 'b0',
      type: 'paragraph',
      docxIndex: 0,
      originalXml: '<w:p/>',
      runs: [
        { text: 'shown' },
        { text: 'hidden', vanish: true, rawRPr: '<w:rPr><w:vanish/></w:rPr>' },
      ],
    }
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc([block]) as never,
    })

    const html = editor.getHTML()
    expect(editor.getText()).toContain('hidden')
    const hiddenSpan = editor.view.dom.querySelector('span[data-doc-style]') as HTMLElement | null
    expect(hiddenSpan?.style.display).toBe('none')
    expect(html).toContain('hidden')
    editor.destroy()
  })
})
