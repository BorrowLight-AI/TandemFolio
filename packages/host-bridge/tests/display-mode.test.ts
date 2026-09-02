import { describe, expect, it } from 'vitest'
import { createMcpDisplayModeController } from '../src/index'

describe('MCP display mode controller', () => {
  it('requests fullscreen once when a format editor first connects', async () => {
    const requests: string[] = []
    const controller = createMcpDisplayModeController()
    const host = {
      requestDisplayMode: async (mode: 'inline' | 'fullscreen') => {
        requests.push(mode)
        return { mode }
      },
    }

    await controller.connect(host, {
      availableDisplayModes: ['inline', 'fullscreen'],
      displayMode: 'inline',
    })
    expect(controller.getState()).toMatchObject({ connected: true, mode: 'fullscreen' })

    await controller.connect(host, {
      availableDisplayModes: ['inline', 'fullscreen'],
      displayMode: 'inline',
    })

    expect(requests).toEqual(['fullscreen'])
  })

  it('keeps newer host state when an older fullscreen request resolves', async () => {
    let resolveRequest!: (result: { mode: 'fullscreen' }) => void
    const requestResult = new Promise<{ mode: 'fullscreen' }>((resolve) => {
      resolveRequest = resolve
    })
    const controller = createMcpDisplayModeController()
    const connecting = controller.connect(
      { requestDisplayMode: async () => requestResult },
      { availableDisplayModes: ['inline', 'fullscreen'], displayMode: 'inline' },
    )

    await Promise.resolve()
    expect(controller.getState()).toMatchObject({ mode: 'inline', pending: true })

    controller.sync({ availableDisplayModes: ['inline'], displayMode: 'inline' })
    resolveRequest({ mode: 'fullscreen' })
    await connecting

    expect(controller.getState()).toMatchObject({
      mode: 'inline',
      fullscreenAvailable: false,
      pending: false,
      error: null,
    })
  })
})
