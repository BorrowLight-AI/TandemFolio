import { useSyncExternalStore } from 'react'
import { EditorFullscreenIcon } from '@genoffice/ui'
import {
  getLiveEditorDisplayMode,
  subscribeLiveEditorDisplayMode,
  toggleLiveEditorFullscreen,
} from '@tandemfolio/host-bridge'

export function HostFullscreenButton() {
  const display = useSyncExternalStore(
    subscribeLiveEditorDisplayMode,
    getLiveEditorDisplayMode,
    getLiveEditorDisplayMode,
  )
  const fullscreen = display.mode === 'fullscreen'
  const label = fullscreen ? '退出全屏' : '全屏'

  return (
    <button
      type="button"
      className="qa-btn host-fullscreen"
      aria-label={label}
      aria-pressed={fullscreen}
      data-tip={label}
      disabled={!display.connected || !display.fullscreenAvailable || display.pending}
      onClick={() => void toggleLiveEditorFullscreen()}
    >
      <EditorFullscreenIcon exit={fullscreen} />
    </button>
  )
}
