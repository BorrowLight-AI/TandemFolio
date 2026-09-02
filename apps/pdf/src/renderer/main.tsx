import { useEffect, useRef, useSyncExternalStore } from 'react'
import { createRoot } from 'react-dom/client'
import {
  attachMcpLiveSession,
  getLiveEditorActivity,
  getLiveEditorDisplayMode,
  subscribeLiveEditorActivity,
  subscribeLiveEditorDisplayMode,
  toggleLiveEditorFullscreen,
} from '@tandemfolio/host-bridge'
import { EditorFileIcon, EditorFullscreenIcon, installScreenTips } from '@genoffice/ui'
import '@genoffice/ui/tokens.css'
import '@genoffice/ui/screentip.css'
import App from './App'
import { LocaleProvider } from './i18n/locale'
import { createBrowserPdfHost } from './host/browser-pdf-api'
import { createPdfCommunityCommandBridge } from './host/community-command-bridge'
import './styles.css'

installScreenTips()
const commandBridge = createPdfCommunityCommandBridge()
const host = createBrowserPdfHost({ commandBridge })
window.pdfApi = host.api

function CommunityPdfEditor(): React.ReactElement {
  const inputRef = useRef<HTMLInputElement>(null)
  const display = useSyncExternalStore(
    subscribeLiveEditorDisplayMode,
    getLiveEditorDisplayMode,
    getLiveEditorDisplayMode,
  )
  const editorActive = useSyncExternalStore(
    subscribeLiveEditorActivity,
    getLiveEditorActivity,
    () => true,
  )
  useEffect(() => attachMcpLiveSession(host.adapter), [])

  return (
    <div className="community-pdf-host">
      <div className="community-host-bar">
        <button
          type="button"
          className="qa-btn host-open"
          aria-label="打开 PDF"
          data-tip="打开 PDF"
          onClick={() => inputRef.current?.click()}
        >
          <EditorFileIcon />
        </button>
        <button
          type="button"
          className="qa-btn host-fullscreen"
          aria-label={display.mode === 'fullscreen' ? '退出全屏' : '全屏'}
          data-tip={display.mode === 'fullscreen' ? '退出全屏' : '全屏'}
          aria-pressed={display.mode === 'fullscreen'}
          onClick={() => void toggleLiveEditorFullscreen()}
        >
          <EditorFullscreenIcon exit={display.mode === 'fullscreen'} />
        </button>
        <input
          ref={inputRef}
          className="community-file-input"
          type="file"
          accept=".pdf,application/pdf"
          onChange={(event) => {
            const file = event.currentTarget.files?.[0]
            event.currentTarget.value = ''
            if (!file) return
            void file
              .arrayBuffer()
              .then((data) => host.stageFile(file.name, data))
              .catch(() => undefined)
          }}
        />
      </div>
      <div className="community-pdf-renderer">
        <LocaleProvider initial="zh">
          <App
            active={editorActive}
            commandBridge={commandBridge}
            onDocumentOpened={host.completeOpen}
          />
        </LocaleProvider>
      </div>
    </div>
  )
}

const root = document.getElementById('root')
if (!root) throw new Error('Missing application root.')
createRoot(root).render(<CommunityPdfEditor />)
