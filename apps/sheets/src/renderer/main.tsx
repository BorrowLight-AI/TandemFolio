import ReactDOM from 'react-dom/client'
import { htmlLang, normalizeLang } from '@genoffice/i18n'
import { installScreenTips } from '@genoffice/ui'

import '@genoffice/ui/tokens.css'
import '@genoffice/ui/screentip.css'
import '@univerjs/preset-sheets-core/lib/index.css'

import { App } from './App'
import { LocaleProvider, setModuleLang } from './i18n/locale'
import './styles.css'

// Static dependencies have finished evaluating when the entry module body runs.
// App.tsx combines this mark with Navigation Timing and its first effect boundary.
window.__genofficeXlsxEntryModuleReadyAt = performance.now()

if (import.meta.hot) {
  import.meta.hot.on('vite:beforeUpdate', ({ updates }) => {
    const replacesUniverRuntime = updates.some(
      ({ path }) => path.endsWith('/App.tsx') || path.endsWith('/univer-sync.ts'),
    )
    if (replacesUniverRuntime) window.location.reload()
  })
}

const root = document.getElementById('root')
if (!root) throw new Error('Missing application root.')

installScreenTips()

async function bootstrap(): Promise<void> {
  const lang = normalizeLang(navigator.language)
  setModuleLang(lang)
  document.documentElement.lang = htmlLang(lang)
  ReactDOM.createRoot(root!).render(
    <LocaleProvider initial={lang}>
      <App />
    </LocaleProvider>,
  )
}

void bootstrap()
