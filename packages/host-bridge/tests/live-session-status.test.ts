// @vitest-environment jsdom
import { expect, it } from 'vitest'
import { createLiveSessionStatus } from '../src/live-session-status'

it('shows an explicit unsaved file location instead of hiding the location panel', () => {
  const status = createLiveSessionStatus(() => undefined)
  try {
    status.path(null)
    const panel = document.querySelector<HTMLElement>('[data-live-session-status]')!
    expect(panel.hidden).toBe(false)
    expect(panel.querySelector('details')!.hidden).toBe(false)
    expect(panel.textContent).toContain('尚未保存')
  } finally {
    status.dispose()
  }
})
