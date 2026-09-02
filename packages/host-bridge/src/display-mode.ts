export type McpDisplayMode = 'inline' | 'fullscreen' | 'pip'

export interface McpDisplayModeState {
  connected: boolean
  fullscreenAvailable: boolean
  mode: McpDisplayMode
  pending: boolean
  error: 'request_failed' | 'request_declined' | null
}

export interface McpDisplayModeHost {
  requestDisplayMode(mode: 'inline' | 'fullscreen'): Promise<{ mode: McpDisplayMode }>
}

export interface McpDisplayModeContext {
  availableDisplayModes?: readonly McpDisplayMode[]
  displayMode?: McpDisplayMode
}

const initialState: McpDisplayModeState = {
  connected: false,
  fullscreenAvailable: false,
  mode: 'inline',
  pending: false,
  error: null,
}

export function createMcpDisplayModeController() {
  let host: McpDisplayModeHost | null = null
  let state = initialState
  let initialFullscreenRequested = false
  let hostContextRevision = 0
  const listeners = new Set<() => void>()

  const update = (next: Partial<McpDisplayModeState>) => {
    state = { ...state, ...next }
    for (const listener of listeners) listener()
  }

  const request = async (mode: 'inline' | 'fullscreen') => {
    if (!host || !state.connected || (mode === 'fullscreen' && !state.fullscreenAvailable)) return
    const activeHost = host
    const requestContextRevision = hostContextRevision
    update({ pending: true, error: null })
    try {
      const result = await activeHost.requestDisplayMode(mode)
      if (host !== activeHost || hostContextRevision !== requestContextRevision) return
      update({
        mode: result.mode,
        pending: false,
        error: result.mode === mode ? null : 'request_declined',
      })
    } catch {
      if (host === activeHost && hostContextRevision === requestContextRevision) {
        update({ pending: false, error: 'request_failed' })
      }
    }
  }

  const sync = (context: McpDisplayModeContext) => {
    hostContextRevision += 1
    const fullscreenAvailable = context.availableDisplayModes?.includes('fullscreen') ?? true
    update({
      connected: true,
      fullscreenAvailable,
      mode: context.displayMode ?? state.mode,
      pending: false,
      error: null,
    })
  }

  return {
    subscribe(listener: () => void) {
      listeners.add(listener)
      return () => listeners.delete(listener)
    },
    getState() {
      return state
    },
    async connect(nextHost: McpDisplayModeHost, context: McpDisplayModeContext) {
      host = nextHost
      sync(context)
      if (!initialFullscreenRequested && state.fullscreenAvailable && state.mode !== 'fullscreen') {
        initialFullscreenRequested = true
        await request('fullscreen')
      }
    },
    sync,
    async toggle() {
      await request(state.mode === 'fullscreen' ? 'inline' : 'fullscreen')
    },
    disconnect() {
      host = null
      state = initialState
      initialFullscreenRequested = false
      for (const listener of listeners) listener()
    },
  }
}

export type McpDisplayModeController = ReturnType<typeof createMcpDisplayModeController>
