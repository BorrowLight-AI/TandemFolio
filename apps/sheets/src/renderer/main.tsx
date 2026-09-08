import '@genoffice/ui/tokens.css'
import '@genoffice/ui/screentip.css'
import '@genoffice/ui/color-picker.css'
import '@univerjs/preset-sheets-core/lib/index.css'
import '@univerjs/preset-sheets-conditional-formatting/lib/index.css'
import '@univerjs/preset-sheets-data-validation/lib/index.css'
import '@univerjs/preset-sheets-drawing/lib/index.css'
import '@univerjs/preset-sheets-find-replace/lib/index.css'
import '@univerjs/preset-sheets-filter/lib/index.css'
import '@univerjs/preset-sheets-note/lib/index.css'
import '@univerjs/preset-sheets-sort/lib/index.css'
import '@univerjs/preset-sheets-table/lib/index.css'

import ReactDOM from 'react-dom/client'
import { htmlLang, normalizeLang } from '@genoffice/i18n'
import { installScreenTips } from '@genoffice/ui'

import { App } from './App'
import { cellFontAliasesReady } from './cell-font-bootstrap'
import { LocaleProvider, setModuleLang } from './i18n/locale'
import './styles.css'

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

function prewarmCommonCellFonts(): void {
  for (const variant of ['', 'bold ', 'italic ', 'italic bold ']) {
    for (const family of ['Calibri', 'Aptos', "'Aptos Narrow'", 'Carlito']) {
      void document.fonts?.load?.(`${variant}16px ${family}`)?.catch(() => {})
    }
  }
}

async function loadCellFontAliases(): Promise<void> {
  await Promise.race([
    cellFontAliasesReady,
    new Promise((resolve) => window.setTimeout(resolve, 2_000)),
  ])
}

async function bootstrap(): Promise<void> {
  await loadCellFontAliases()
  const lang = normalizeLang(navigator.language)
  setModuleLang(lang)
  document.documentElement.lang = htmlLang(lang)
  installScreenTips()
  // The statically linked application graph and required font aliases are
  // ready. App.tsx combines this mark with Navigation Timing and its first
  // layout-effect boundary.
  window.__genofficeXlsxEntryModuleReadyAt = performance.now()
  ReactDOM.createRoot(root!).render(
    <LocaleProvider initial={lang}>
      <App />
    </LocaleProvider>,
  )
  // Alias faces must precede Univer's first measurement; speculative native
  // font variants can warm after startup without delaying the first canvas.
  window.setTimeout(prewarmCommonCellFonts, 500)
}

void bootstrap()
