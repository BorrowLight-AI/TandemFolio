// Modified by TandemFolio contributors: initialize the shared UI-language preference.
import { createRoot } from 'react-dom/client'
import { getUiLanguageSnapshot, htmlLang } from '@genoffice/i18n'
import { App } from './App'
import { installBrowserDesktop } from './host/browser-desktop'
import { LocaleProvider, setModuleLang } from './i18n/locale'
import type { UiTheme } from '../shared/lite-api'
import '@genoffice/ui/tokens.css'
import '@genoffice/ui/dropdown.css'
import '@genoffice/ui/language-menu.css'
import './styles.css'

function applyTheme(theme: UiTheme): void {
  if (theme === 'system') document.documentElement.removeAttribute('data-theme')
  else document.documentElement.setAttribute('data-theme', theme)
}

async function bootstrap(): Promise<void> {
  installBrowserDesktop()
  const lang = getUiLanguageSnapshot().lang
  let theme: UiTheme = 'system'
  try {
    theme = await window.desktop.getTheme().catch(() => 'system' as const)
  } catch {
    /* dev renderer without the preload bridge */
  }
  setModuleLang(lang)
  document.documentElement.lang = htmlLang(lang)
  applyTheme(theme)
  window.desktop?.onThemeChanged(applyTheme)
  createRoot(document.getElementById('root')!).render(
    <LocaleProvider initial={lang}>
      <App />
    </LocaleProvider>,
  )
}

void bootstrap()
