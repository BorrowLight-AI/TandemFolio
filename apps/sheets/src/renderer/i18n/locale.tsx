import { createContext, useContext, useEffect, useState } from 'react'
import type { ReactNode } from 'react'
import { createI18n, type Lang, type Params } from '@genoffice/i18n'

import type { strings as StringCatalog } from './strings'

export type StringKey = keyof typeof StringCatalog.zh
export type TFunc = (key: StringKey, params?: Params) => string

let translate: ReturnType<typeof createI18n<typeof StringCatalog.zh>> | null = null
let stringsReady: Promise<void> | null = null

function ensureLocaleStrings(): Promise<void> {
  stringsReady ??= import('./strings').then(({ strings }) => {
    translate = createI18n(strings)
  })
  return stringsReady
}

export async function loadLocaleStrings(): Promise<void> {
  await ensureLocaleStrings()
}

let moduleLang: Lang = 'zh'
export const getLang = (): Lang => moduleLang
export const setModuleLang = (lang: Lang): void => {
  moduleLang = lang
}
export const t: TFunc = (key, params) => translate?.(moduleLang, key, params) ?? String(key)

export const DATE_LOCALES: Record<Lang, string> = {
  zh: 'zh-CN',
  en: 'en-US',
  ja: 'ja-JP',
  ko: 'ko-KR',
  fr: 'fr-FR',
  de: 'de-DE',
  es: 'es-ES',
  th: 'th-TH',
  id: 'id-ID',
  ru: 'ru-RU',
  ar: 'ar-SA',
  pt: 'pt-BR',
  it: 'it-IT',
  pl: 'pl-PL',
  nl: 'nl-NL',
  ms: 'ms-MY',
  he: 'he-IL',
  hi: 'hi-IN',
  'zh-TW': 'zh-TW',
}

const LocaleContext = createContext({ lang: 'zh' as Lang, catalogRevision: 0 })

export function LocaleProvider({ initial, children }: { initial: Lang; children: ReactNode }) {
  const [lang] = useState<Lang>(initial)
  const [catalogRevision, setCatalogRevision] = useState(translate ? 1 : 0)
  useEffect(() => setModuleLang(lang), [lang])
  useEffect(() => {
    if (translate) return
    let active = true
    void ensureLocaleStrings().then(() => {
      if (active) setCatalogRevision(1)
    })
    return () => {
      active = false
    }
  }, [])
  return (
    <LocaleContext.Provider value={{ lang, catalogRevision }}>{children}</LocaleContext.Provider>
  )
}

export interface I18n {
  lang: Lang
  t: TFunc
  dateLocale: string
}

export function useI18n(): I18n {
  const { lang } = useContext(LocaleContext)
  return {
    lang,
    t: (key, params) => translate?.(lang, key, params) ?? String(key),
    dateLocale: DATE_LOCALES[lang],
  }
}
