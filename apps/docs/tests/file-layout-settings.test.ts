import { Editor } from '@tiptap/core'
import { describe, expect, it } from 'vitest'
import type { ParsedDocFull } from '@genoffice/docx-engine'
import { applyDocLayoutSettings } from '../src/renderer/file-actions'
import { editorExtensions } from '../src/renderer/editor/extensions'

describe('DOCX document layout settings', () => {
  it('hydrates tab grid, compatibility spacing, CJK punctuation compression, and hyphenation language', () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: { type: 'doc', content: [{ type: 'docParagraph' }] },
    })
    applyDocLayoutSettings(
      editor,
      {
        defaultTabStopTwips: 480,
        compatibilityMode: 15,
        compressPunctuation: true,
        autoHyphenation: true,
        docDefaults: { lang: 'en-US' },
      } as ParsedDocFull,
    )
    expect(editor.storage.tabStops.defaultTabStopTwips).toBe(480)
    expect(editor.storage.justifyShrink.enabled).toBe(true)
    expect(editor.storage.cjkPunctShrink.enabled).toBe(true)
    expect(editor.view.dom.getAttribute('lang')).toBe('en-US')

    applyDocLayoutSettings(editor, {} as ParsedDocFull)
    expect(editor.storage.justifyShrink.enabled).toBe(false)
    expect(editor.storage.cjkPunctShrink.enabled).toBe(false)
    expect(editor.view.dom.hasAttribute('lang')).toBe(false)
    editor.destroy()
  })
})
