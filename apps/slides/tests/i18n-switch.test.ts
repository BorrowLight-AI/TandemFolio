import { createElement } from 'react'
import { flushSync } from 'react-dom'
import { createRoot } from 'react-dom/client'
import { setUiLanguagePreference } from '@genoffice/i18n'
import { afterEach, beforeEach, describe, expect, it } from 'vitest'

import { LocaleProvider, useI18n } from '../src/renderer/i18n/locale'

beforeEach(() => setUiLanguagePreference('zh'))
afterEach(() => {
  setUiLanguagePreference('system')
  document.body.replaceChildren()
})

describe('PPTX UI language', () => {
  it('rerenders the mounted format UI when the user changes language', () => {
    const container = document.createElement('div')
    document.body.append(container)
    const root = createRoot(container)
    function Probe() {
      const { t } = useI18n()
      return createElement('span', null, t('ribbonFileOpen'))
    }
    flushSync(() => {
      root.render(createElement(LocaleProvider, { initial: 'zh', children: createElement(Probe) }))
    })

    flushSync(() => setUiLanguagePreference('en'))

    expect(container.textContent).toBe('Open…')
    expect(document.documentElement.lang).toBe('en-US')
    root.unmount()
  })
})
