import { useSyncExternalStore } from 'react'
import {
  getUiLanguageSnapshot,
  LANGS,
  setUiLanguagePreference,
  subscribeUiLanguage,
  type Lang,
} from '@genoffice/i18n'

const LANGUAGE_NAMES: Record<Lang, string> = {
  zh: '简体中文',
  en: 'English',
  ja: '日本語',
  ko: '한국어',
  fr: 'Français',
  de: 'Deutsch',
  es: 'Español',
  th: 'ไทย',
  id: 'Bahasa Indonesia',
  ru: 'Русский',
  ar: 'العربية',
  pt: 'Português',
  it: 'Italiano',
  pl: 'Polski',
  nl: 'Nederlands',
  ms: 'Bahasa Melayu',
  he: 'עברית',
  hi: 'हिन्दी',
  'zh-TW': '繁體中文',
}

const CONTROL_LABELS: Record<Lang, string> = {
  zh: '界面语言',
  en: 'Interface language',
  ja: '表示言語',
  ko: '인터페이스 언어',
  fr: 'Langue de l’interface',
  de: 'Oberflächensprache',
  es: 'Idioma de la interfaz',
  th: 'ภาษาของอินเทอร์เฟซ',
  id: 'Bahasa antarmuka',
  ru: 'Язык интерфейса',
  ar: 'لغة الواجهة',
  pt: 'Idioma da interface',
  it: 'Lingua dell’interfaccia',
  pl: 'Język interfejsu',
  nl: 'Interfacetaal',
  ms: 'Bahasa antara muka',
  he: 'שפת הממשק',
  hi: 'इंटरफ़ेस भाषा',
  'zh-TW': '介面語言',
}

const SYSTEM_LABELS: Record<Lang, string> = {
  zh: '跟随系统',
  en: 'Follow system',
  ja: 'システムに従う',
  ko: '시스템 설정 사용',
  fr: 'Suivre le système',
  de: 'Systemeinstellung',
  es: 'Seguir el sistema',
  th: 'ตามระบบ',
  id: 'Ikuti sistem',
  ru: 'Как в системе',
  ar: 'اتباع النظام',
  pt: 'Seguir o sistema',
  it: 'Segui il sistema',
  pl: 'Użyj ustawień systemu',
  nl: 'Systeem volgen',
  ms: 'Ikut sistem',
  he: 'לפי המערכת',
  hi: 'सिस्टम के अनुसार',
  'zh-TW': '跟隨系統',
}

export function EditorLanguageMenu(): React.JSX.Element {
  const { preference, lang } = useSyncExternalStore(
    subscribeUiLanguage,
    getUiLanguageSnapshot,
    getUiLanguageSnapshot,
  )
  const label = CONTROL_LABELS[lang]

  return (
    <label className="editor-language-menu" title={label}>
      <span className="editor-language-menu-icon" aria-hidden="true">
        🌐
      </span>
      <select
        className="editor-language-menu-select"
        aria-label={label}
        value={preference}
        onChange={(event) =>
          setUiLanguagePreference(event.currentTarget.value as typeof preference)
        }
      >
        <option value="system">{SYSTEM_LABELS[lang]}</option>
        {LANGS.map((option) => (
          <option key={option} value={option}>
            {LANGUAGE_NAMES[option]}
          </option>
        ))}
      </select>
    </label>
  )
}
