import { createElement } from 'react'
import { flushSync } from 'react-dom'
import { createRoot, type Root } from 'react-dom/client'
import { LANGS, setUiLanguagePreference } from '@genoffice/i18n'
import { afterEach, beforeEach, describe, expect, it } from 'vitest'

import { Ribbon } from '../src/renderer/components/Ribbon'
import { LocaleProvider } from '../src/renderer/i18n/locale'

let root: Root | null = null

beforeEach(() => {
  setUiLanguagePreference('zh')
})

afterEach(() => {
  root?.unmount()
  root = null
  document.body.replaceChildren()
  setUiLanguagePreference('system')
})

describe('Markdown Ribbon', () => {
  it('renders the top-level file and display actions as localized icon buttons', () => {
    const container = document.createElement('div')
    document.body.append(container)
    root = createRoot(container)
    flushSync(() => {
      root!.render(
        createElement(LocaleProvider, {
          initial: 'zh',
          children: createElement(Ribbon, {
            editor: null,
            disabled: false,
            dirty: false,
            onOpen: () => undefined,
            onSave: () => undefined,
            onExportDocx: () => undefined,
            onPrintPdf: () => undefined,
            autoSave: false,
            onToggleAutoSave: () => undefined,
            onInsertImage: () => undefined,
            frontmatterOpen: false,
            onToggleFrontmatter: () => undefined,
            fullscreen: false,
            onToggleFullscreen: () => undefined,
          }),
        }),
      )
    })

    for (const label of ['打开 Markdown', '导出 DOCX', '打印 / PDF', '全屏']) {
      const button = container.querySelector<HTMLButtonElement>(`button[aria-label="${label}"]`)
      expect(button, label).not.toBeNull()
      expect(button?.querySelector('svg'), label).not.toBeNull()
      expect(button?.textContent?.trim(), label).toBe('')
    }
  })

  it('offers system mode and every supported UI language from the top toolbar', () => {
    const container = document.createElement('div')
    document.body.append(container)
    root = createRoot(container)
    flushSync(() => {
      root!.render(
        createElement(LocaleProvider, {
          initial: 'zh',
          children: createElement(Ribbon, {
            editor: null,
            disabled: false,
            dirty: false,
            onOpen: () => undefined,
            onSave: () => undefined,
            onExportDocx: () => undefined,
            onPrintPdf: () => undefined,
            autoSave: false,
            onToggleAutoSave: () => undefined,
            onInsertImage: () => undefined,
            frontmatterOpen: false,
            onToggleFrontmatter: () => undefined,
            fullscreen: false,
            onToggleFullscreen: () => undefined,
          }),
        }),
      )
    })

    const picker = container.querySelector<HTMLSelectElement>('select[aria-label="界面语言"]')
    expect(picker).not.toBeNull()
    expect([...picker!.options].map(({ value }) => value)).toEqual(['system', ...LANGS])
  })

  it('switches the mounted editor UI language without reloading it', () => {
    const container = document.createElement('div')
    document.body.append(container)
    root = createRoot(container)
    flushSync(() => {
      root!.render(
        createElement(LocaleProvider, {
          initial: 'zh',
          children: createElement(Ribbon, {
            editor: null,
            disabled: false,
            dirty: false,
            onOpen: () => undefined,
            onSave: () => undefined,
            onExportDocx: () => undefined,
            onPrintPdf: () => undefined,
            autoSave: false,
            onToggleAutoSave: () => undefined,
            onInsertImage: () => undefined,
            frontmatterOpen: false,
            onToggleFrontmatter: () => undefined,
            fullscreen: false,
            onToggleFullscreen: () => undefined,
          }),
        }),
      )
    })
    const picker = container.querySelector<HTMLSelectElement>('select[aria-label="界面语言"]')!

    flushSync(() => {
      picker.value = 'en'
      picker.dispatchEvent(new Event('change', { bubbles: true }))
    })

    expect(container.querySelector('button[aria-label="Open Markdown"]')).not.toBeNull()
    expect(document.documentElement.lang).toBe('en-US')
  })
})
