// Modified by TandemFolio contributors: initialize the shared UI-language preference.
import ReactDOM from 'react-dom/client'
import { getUiLanguageSnapshot, htmlLang } from '@genoffice/i18n'
import '@genoffice/ui/tokens.css'
import '@genoffice/ui/language-menu.css'
import { attachMcpLiveSession } from '@tandemfolio/host-bridge'

import { App } from './App'
import { AudienceView } from './components/AudienceView'
import { createBrowserSlidesHost } from './host/browser-slides-api'
import { LocaleProvider } from './i18n/locale'
import './styles.css'

const root = document.getElementById('root')
if (!root) throw new Error('Missing application root.')

const mode = new URLSearchParams(window.location.search).get('mode')
const host = createBrowserSlidesHost()
window.slidesApi = host.api
if (mode !== 'audience') attachMcpLiveSession(host.adapter)

const lang = getUiLanguageSnapshot().lang
document.documentElement.lang = htmlLang(lang)
ReactDOM.createRoot(root).render(
  <LocaleProvider initial={lang}>
    {mode === 'audience' ? <AudienceView /> : <App />}
  </LocaleProvider>,
)
