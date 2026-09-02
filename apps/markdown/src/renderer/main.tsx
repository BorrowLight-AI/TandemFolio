import { createRoot } from 'react-dom/client'
import { htmlLang, normalizeLang } from '@genoffice/i18n'
import { installScreenTips } from '@genoffice/ui'
import App from './App'
import { LocaleProvider } from './i18n/locale'
import '@genoffice/ui/tokens.css'
import '@genoffice/ui/screentip.css'
import './styles.css'

installScreenTips()

const lang = normalizeLang(navigator.language)
document.documentElement.lang = htmlLang(lang)

const preferredTheme = localStorage.getItem('tandemfolio.theme')
if (preferredTheme === 'light' || preferredTheme === 'dark') {
  document.documentElement.dataset.theme = preferredTheme
}

createRoot(document.getElementById('root')!).render(
  <LocaleProvider initial={lang}>
    <App />
  </LocaleProvider>,
)
