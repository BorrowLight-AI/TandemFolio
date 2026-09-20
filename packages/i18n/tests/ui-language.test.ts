// @vitest-environment jsdom

import { afterEach, describe, expect, it, vi } from 'vitest'

import {
  getUiLanguageSnapshot,
  setUiLanguagePreference,
  subscribeUiLanguage,
  UI_LANGUAGE_STORAGE_KEY,
} from '../src/index'

afterEach(() => {
  setUiLanguagePreference('system')
  localStorage.clear()
  vi.restoreAllMocks()
})

describe('UI language preference', () => {
  it('lets the user select a language that is immediately observable and persisted', () => {
    const observed: string[] = []
    const unsubscribe = subscribeUiLanguage(() => {
      observed.push(getUiLanguageSnapshot().lang)
    })

    setUiLanguagePreference('ja')

    expect(getUiLanguageSnapshot()).toEqual({ preference: 'ja', lang: 'ja' })
    expect(localStorage.getItem(UI_LANGUAGE_STORAGE_KEY)).toBe('ja')
    expect(observed).toEqual(['ja'])
    unsubscribe()
  })

  it('restores the persisted preference when a renderer starts', async () => {
    localStorage.setItem(UI_LANGUAGE_STORAGE_KEY, 'fr')
    vi.resetModules()

    const freshI18n = await import('../src/index')

    expect(freshI18n.getUiLanguageSnapshot()).toEqual({ preference: 'fr', lang: 'fr' })
  })

  it('updates a system preference when the browser language changes', () => {
    setUiLanguagePreference('system')
    vi.spyOn(window.navigator, 'language', 'get').mockReturnValue('de-DE')

    window.dispatchEvent(new Event('languagechange'))

    expect(getUiLanguageSnapshot()).toEqual({ preference: 'system', lang: 'de' })
  })

  it('follows a language preference selected in another same-origin renderer', () => {
    window.dispatchEvent(
      new StorageEvent('storage', {
        key: UI_LANGUAGE_STORAGE_KEY,
        newValue: 'ko',
      }),
    )

    expect(getUiLanguageSnapshot()).toEqual({ preference: 'ko', lang: 'ko' })
  })
})
