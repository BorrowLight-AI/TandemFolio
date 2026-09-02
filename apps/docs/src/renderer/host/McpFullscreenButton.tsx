import { useSyncExternalStore } from 'react'
import { EditorFullscreenIcon } from '@genoffice/ui'
import { useI18n } from '../i18n/locale'
import {
  getLiveEditorDisplayMode,
  subscribeLiveEditorDisplayMode,
  toggleLiveEditorFullscreen,
} from '@tandemfolio/host-bridge'

const labels = {
  en: { enter: 'Fullscreen', exit: 'Exit fullscreen', unavailable: 'Fullscreen is unavailable' },
  zh: { enter: '全屏', exit: '退出全屏', unavailable: '当前宿主不支持全屏' },
  'zh-TW': { enter: '全螢幕', exit: '退出全螢幕', unavailable: '目前宿主不支援全螢幕' },
} as const

export function McpFullscreenButton() {
  const { lang } = useI18n()
  const state = useSyncExternalStore(
    subscribeLiveEditorDisplayMode,
    getLiveEditorDisplayMode,
    getLiveEditorDisplayMode,
  )

  if (!state.embedded) return null

  const copy = labels[lang as keyof typeof labels] ?? labels.en
  const isFullscreen = state.mode === 'fullscreen'
  const label = isFullscreen ? copy.exit : copy.enter
  const unavailable = !state.connected || !state.fullscreenAvailable
  const title = unavailable
    ? copy.unavailable
    : state.error === 'request_failed'
      ? `${label} — request failed`
      : state.error === 'request_declined'
        ? `${label} — host kept the current mode`
        : label

  return (
    <button
      type="button"
      className="mcp-fullscreen-button"
      aria-label={label}
      aria-pressed={isFullscreen}
      aria-disabled={unavailable || state.pending}
      aria-busy={state.pending}
      title={title}
      onClick={() => {
        if (!unavailable && !state.pending) void toggleLiveEditorFullscreen()
      }}
    >
      <EditorFullscreenIcon exit={isFullscreen} />
    </button>
  )
}
