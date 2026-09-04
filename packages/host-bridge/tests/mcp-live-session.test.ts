// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from 'vitest'

const mcp = vi.hoisted(() => ({
  commands: [] as Array<Record<string, unknown>>,
  calls: [] as Array<{ name: string; arguments: Record<string, unknown> }>,
  events: [] as string[],
  stagedBytes: null as Uint8Array | null,
  assetBytes: null as Uint8Array | null,
  fontBytes: null as Uint8Array | null,
  requestedModes: [] as string[],
  sizeNotifications: [] as Array<{ width: number; height: number }>,
  pollWaiters: [] as Array<() => void>,
  pollFailures: 0,
  pollError: null as string | null,
  pollRetryAfterMs: undefined as number | undefined,
  pollFilePath: undefined as string | null | undefined,
  restoreCommandId: undefined as string | undefined,
  connectFailures: 0,
  downloads: [] as Array<Record<string, unknown>>,
  instance: null as {
    ontoolinput?: (input: { arguments?: Record<string, unknown> }) => void
    ontoolresult?: (result: { structuredContent?: Record<string, unknown> }) => void
  } | null,
}))

vi.mock('@modelcontextprotocol/ext-apps', () => ({
  App: class {
    ontoolinput?: (input: { arguments?: Record<string, unknown> }) => void
    ontoolresult?: (result: { structuredContent?: Record<string, unknown> }) => void
    onhostcontextchanged?: () => void
    private readonly autoResize: boolean
    constructor(
      _appInfo: unknown,
      _capabilities: unknown,
      options: { autoResize?: boolean } = { autoResize: true },
    ) {
      this.autoResize = options.autoResize !== false
      mcp.instance = this
    }
    async connect() {
      if (mcp.connectFailures > 0) {
        mcp.connectFailures -= 1
        throw new Error('host iframe source is not ready')
      }
      if (this.autoResize) {
        window.requestAnimationFrame(() => {
          mcp.events.push('size:initial')
          mcp.sizeNotifications.push({ width: window.innerWidth, height: window.innerHeight })
        })
      }
    }
    async close() {}
    getHostContext() {
      return { availableDisplayModes: ['inline', 'fullscreen'], displayMode: 'inline' }
    }
    async requestDisplayMode({ mode }: { mode: string }) {
      mcp.events.push(`display:${mode}`)
      mcp.requestedModes.push(mode)
      return { mode }
    }
    async downloadFile(request: Record<string, unknown>) {
      mcp.downloads.push(request)
      return {}
    }
    async callServerTool(call: { name: string; arguments: Record<string, unknown> }) {
      mcp.events.push(`tool:${call.name}`)
      mcp.calls.push(call)
      if (call.name === 'office_editor_poll') {
        if (mcp.pollError) {
          return {
            isError: true,
            structuredContent: {
              ok: false,
              error: mcp.pollError,
              retryAfterMs: mcp.pollRetryAfterMs,
            },
          }
        }
        if (mcp.pollFailures > 0) {
          mcp.pollFailures -= 1
          throw new Error('temporary transport failure')
        }
        const commands = mcp.commands.splice(0)
        if (commands.length > 0 || !(Number(call.arguments.waitMs) > 0)) {
          return {
            structuredContent: {
              ok: true,
              commands,
              restoreCommandId: mcp.restoreCommandId,
              filePath: mcp.pollFilePath,
            },
          }
        }
        return await new Promise<{
          structuredContent: {
            ok: true
            commands: Array<Record<string, unknown>>
            filePath?: string | null
          }
        }>((resolve) => {
          mcp.pollWaiters.push(() =>
            resolve({
              structuredContent: {
                ok: true,
                commands: mcp.commands.splice(0),
                filePath: mcp.pollFilePath,
              },
            }),
          )
        })
      }
      if (call.name === 'office_editor_read_file_chunk' && mcp.stagedBytes) {
        const offset = call.arguments.offset as number
        const length = call.arguments.length as number
        const chunk = mcp.stagedBytes.subarray(offset, offset + length)
        let binary = ''
        for (const byte of chunk) binary += String.fromCharCode(byte)
        return {
          structuredContent: {
            ok: true,
            data: btoa(binary),
            nextOffset: offset + chunk.length,
          },
        }
      }
      if (call.name === 'office_editor_read_local_asset_chunk' && mcp.assetBytes) {
        const offset = call.arguments.offset as number
        const length = call.arguments.length as number
        const chunk = mcp.assetBytes.subarray(offset, offset + length)
        let binary = ''
        for (const byte of chunk) binary += String.fromCharCode(byte)
        return {
          structuredContent: {
            ok: true,
            size: mcp.assetBytes.length,
            mime: 'image/png',
            data: btoa(binary),
            nextOffset: offset + chunk.length,
            eof: offset + chunk.length >= mcp.assetBytes.length,
          },
        }
      }
      if (call.name === 'office_editor_read_font_chunk' && mcp.fontBytes) {
        const offset = call.arguments.offset as number
        const length = call.arguments.length as number
        const chunk = mcp.fontBytes.subarray(offset, offset + length)
        let binary = ''
        for (const byte of chunk) binary += String.fromCharCode(byte)
        return {
          structuredContent: {
            ok: true,
            data: btoa(binary),
            nextOffset: offset + chunk.length,
            eof: offset + chunk.length >= mcp.fontBytes.length,
          },
        }
      }
      if (call.name === 'office_editor_begin_document_save') {
        return {
          structuredContent: {
            ok: true,
            uploadId: `save-${call.arguments.sessionId}`,
            path: `/tmp/tandemfolio/${call.arguments.fileName}`,
          },
        }
      }
      if (call.name === 'office_editor_write_document_save_chunk') {
        const decoded = atob(call.arguments.data as string)
        return {
          structuredContent: {
            ok: true,
            nextOffset: (call.arguments.offset as number) + decoded.length,
          },
        }
      }
      if (call.name === 'office_editor_commit_document_save') {
        return {
          structuredContent: {
            ok: true,
            path: '/tmp/tandemfolio/Quarterly Review.pptx',
          },
        }
      }
      if (call.name === 'office_editor_abort_document_save') {
        return { structuredContent: { ok: true } }
      }
      if (call.name === 'office_editor_begin_recovery') {
        return { structuredContent: { ok: true, uploadId: `upload-${call.arguments.sessionId}` } }
      }
      if (call.name === 'office_editor_write_recovery_chunk') {
        const decoded = atob(call.arguments.data as string)
        return {
          structuredContent: {
            ok: true,
            nextOffset: (call.arguments.offset as number) + decoded.length,
          },
        }
      }
      if (call.name === 'office_editor_commit_recovery') {
        return { structuredContent: { ok: true } }
      }
      return { structuredContent: { ok: true } }
    }
  },
}))

import {
  attachMcpLiveSession,
  getLiveEditorActivity,
  getLiveEditorDisplayMode,
  readLiveEditorBundledFontAsset,
  readLiveEditorLocalAsset,
  saveLiveEditorFile,
  replaceLiveEditorDocument,
  subscribeLiveEditorActivity,
} from '../src/mcp-live-session'

afterEach(() => {
  vi.useRealTimers()
  vi.unstubAllGlobals()
  mcp.commands = []
  mcp.calls = []
  mcp.events = []
  mcp.stagedBytes = null
  mcp.assetBytes = null
  mcp.fontBytes = null
  mcp.requestedModes = []
  mcp.sizeNotifications = []
  mcp.pollFailures = 0
  mcp.pollError = null
  mcp.pollRetryAfterMs = undefined
  mcp.pollFilePath = undefined
  mcp.restoreCommandId = undefined
  mcp.connectFailures = 0
  mcp.downloads = []
  for (const release of mcp.pollWaiters.splice(0)) release()
  mcp.instance = null
  Object.defineProperty(window, 'parent', { configurable: true, value: window })
  Object.defineProperty(document, 'visibilityState', { configurable: true, value: 'visible' })
})

function stubEditorIntersection(initiallyIntersecting: boolean): (isIntersecting: boolean) => void {
  let observer: FakeIntersectionObserver | null = null

  class FakeIntersectionObserver implements IntersectionObserver {
    readonly root = null
    readonly rootMargin = '0px'
    readonly thresholds = [0]

    constructor(private readonly callback: IntersectionObserverCallback) {
      observer = this
    }

    observe(target: Element): void {
      this.emit(target, initiallyIntersecting)
    }

    emit(target: Element, isIntersecting: boolean): void {
      this.callback(
        [
          {
            isIntersecting,
            intersectionRatio: isIntersecting ? 1 : 0,
            target,
          } as IntersectionObserverEntry,
        ],
        this,
      )
    }

    disconnect(): void {}
    unobserve(): void {}
    takeRecords(): IntersectionObserverEntry[] {
      return []
    }
  }

  vi.stubGlobal('IntersectionObserver', FakeIntersectionObserver)
  return (isIntersecting) => observer?.emit(document.documentElement, isIntersecting)
}

function bindEditor(sessionId: string, viewId = `view-${sessionId}`): void {
  mcp.instance?.ontoolinput?.({ arguments: { sessionId } })
  mcp.instance?.ontoolresult?.({ structuredContent: { ok: true, sessionId, viewId } })
}

describe('format-neutral MCP live session', () => {
  it('honors explicit continuation despite a stale inactive viewport hint', async () => {
    vi.useFakeTimers()
    Object.defineProperty(window, 'parent', { configurable: true, value: {} })
    const intersect = stubEditorIntersection(true)
    mcp.pollError = 'editor_view_conflict'
    mcp.pollRetryAfterMs = 1000
    const teardown = attachMcpLiveSession({
      execute: async () => ({ ok: true }),
      snapshot: (revision) => ({ revision, fileName: null, dirty: false, selection: null }),
    })
    try {
      bindEditor('explicit-inactive')
      await vi.advanceTimersByTimeAsync(100)
      intersect(false)
      const prior = mcp.calls.filter((call) => call.name === 'office_editor_poll').length
      const button = [...document.querySelectorAll('button')].find(
        (el) => el.textContent === '在此继续编辑',
      )!
      button.click()
      await vi.advanceTimersByTimeAsync(100)
      const polls = mcp.calls.filter((call) => call.name === 'office_editor_poll')
      expect(polls).toHaveLength(prior + 1)
      expect(polls.at(-1)?.arguments).toMatchObject({ activateView: true, active: true })
      // The outgoing owner needs time to checkpoint: retain the explicit intent for retries.
      mcp.pollError = null
      await vi.advanceTimersByTimeAsync(1100)
      expect(document.documentElement.dataset.liveEditorConnection).toBe('connected')
      expect(document.documentElement.dataset.liveEditorActive).toBe('false')
    } finally {
      teardown()
    }
  })

  it('stops waiting after 30 seconds and requires an explicit action to try again', async () => {
    vi.useFakeTimers()
    Object.defineProperty(window, 'parent', { configurable: true, value: {} })
    const intersect = stubEditorIntersection(true)
    mcp.pollError = 'editor_view_conflict'
    mcp.pollRetryAfterMs = 1000
    const teardown = attachMcpLiveSession({
      execute: async () => ({ ok: true }),
      snapshot: (revision) => ({ revision, fileName: null, dirty: false, selection: null }),
    })
    try {
      bindEditor('bounded-wait')
      await vi.advanceTimersByTimeAsync(100)
      expect(document.documentElement.dataset.liveEditorConnection).toBe('editor_view_conflict')
      await vi.advanceTimersByTimeAsync(30_000)
      expect(document.querySelector('[data-live-session-status]')?.textContent).toContain(
        '等待已超时',
      )
      const polls = mcp.calls.filter((call) => call.name === 'office_editor_poll').length
      intersect(false)
      intersect(true)
      await vi.advanceTimersByTimeAsync(20_000)
      expect(mcp.calls.filter((call) => call.name === 'office_editor_poll')).toHaveLength(polls)
      mcp.pollError = null
      const button = [...document.querySelectorAll('button')].find(
        (el) => el.textContent === '在此继续编辑',
      )!
      button.click()
      await vi.advanceTimersByTimeAsync(100)
      expect(document.documentElement.dataset.liveEditorConnection).toBe('connected')
    } finally {
      teardown()
    }
  })

  it.each(['hydrate', 'execute'])(
    'stops a failed %s restore instead of accepting a blank session',
    async (phase) => {
      Object.defineProperty(window, 'parent', { configurable: true, value: {} })
      mcp.restoreCommandId = 'restore'
      mcp.commands.push({
        commandId: 'restore',
        operation: 'markdown.document.load_staged',
        baseRevision: 0,
        arguments: { blobId: 'recovery', fileName: 'restored.md', size: 4 },
      })
      if (phase === 'execute') mcp.stagedBytes = new TextEncoder().encode('test')
      const teardown = attachMcpLiveSession({
        execute: async () => {
          throw new Error('Cannot load recovered bytes')
        },
        snapshot: (revision) => ({ revision, fileName: null, dirty: false, selection: null }),
      })
      try {
        bindEditor('failed-restore')
        await vi.waitFor(() =>
          expect(document.documentElement.dataset.liveEditorConnection).toBe(
            'document_restore_failed',
          ),
        )
        expect(mcp.calls.filter((call) => call.name === 'office_editor_poll')).toHaveLength(1)
        expect(
          mcp.calls.find((call) => call.name === 'office_editor_acknowledge')?.arguments,
        ).toMatchObject({ commandId: 'restore', ok: false })
      } finally {
        teardown()
      }
    },
  )

  it('shows the actual committed path and lets the user copy it', async () => {
    Object.defineProperty(window, 'parent', { configurable: true, value: {} })
    let copied = ''
    Object.defineProperty(navigator, 'clipboard', {
      configurable: true,
      value: {
        writeText: async (text: string) => {
          copied = text
        },
      },
    })
    const teardown = attachMcpLiveSession({
      execute: async () => ({ ok: true }),
      snapshot: (revision) => ({
        revision,
        fileName: 'Quarterly Review.pptx',
        dirty: false,
        selection: null,
      }),
    })
    try {
      bindEditor('path-visible')
      await vi.waitFor(() =>
        expect(document.documentElement.dataset.liveEditorConnection).toBe('connected'),
      )
      await saveLiveEditorFile({ fileName: 'Quarterly Review.pptx', data: new ArrayBuffer(0) })
      const path = document.querySelector<HTMLInputElement>('[aria-label="文件绝对路径"]')
      expect(path?.value).toBe('/tmp/tandemfolio/Quarterly Review.pptx')
      document.querySelector<HTMLButtonElement>('[aria-label="复制文件路径"]')?.click()
      await vi.waitFor(() => expect(copied).toBe(path?.value))
    } finally {
      teardown()
    }
  })

  it('detaches the previous target before a browser replacement and checkpoints the new document', async () => {
    Object.defineProperty(window, 'parent', { configurable: true, value: {} })
    const teardown = attachMcpLiveSession({
      execute: async () => ({ ok: true }),
      snapshot: (revision) => ({ revision, fileName: 'new.md', dirty: false, selection: null }),
      recoverySnapshot: async () => ({
        fileName: 'new.md',
        data: new TextEncoder().encode('new document').buffer,
      }),
    })
    try {
      bindEditor('replace')
      await vi.waitFor(() =>
        expect(document.documentElement.dataset.liveEditorConnection).toBe('connected'),
      )
      await replaceLiveEditorDocument(async () => {
        expect(mcp.calls.some((call) => call.name === 'office_editor_reset_document')).toBe(true)
      })
      expect(mcp.calls.some((call) => call.name === 'office_editor_commit_recovery')).toBe(true)
    } finally {
      teardown()
    }
  })

  it('keeps the Export Copy destination visible when ordinary polls repeat the bound source path', async () => {
    Object.defineProperty(window, 'parent', { configurable: true, value: {} })
    mcp.pollFilePath = '/tmp/source.pptx'
    const teardown = attachMcpLiveSession({
      execute: async () => ({ ok: true }),
      snapshot: (revision) => ({
        revision,
        fileName: 'source.pptx',
        dirty: false,
        selection: null,
      }),
    })
    try {
      bindEditor('export-path')
      await vi.waitFor(() => expect(mcp.pollWaiters.length).toBeGreaterThan(0))
      await saveLiveEditorFile({
        fileName: 'Quarterly Review.pptx',
        data: new ArrayBuffer(0),
        mode: 'export-copy',
      })
      for (let index = 0; index < 2; index++) {
        mcp.pollWaiters.shift()!()
        await vi.waitFor(() => expect(mcp.pollWaiters.length).toBeGreaterThan(0))
      }
      expect(document.querySelector<HTMLInputElement>('[aria-label="文件绝对路径"]')?.value).toBe(
        '/tmp/tandemfolio/Quarterly Review.pptx',
      )
      expect(document.querySelector('[data-live-session-status] summary')?.textContent).toBe(
        '导出副本位置',
      )
    } finally {
      teardown()
    }
  })

  it('waits for browser replacement before executing a command delivered by an already pending poll', async () => {
    Object.defineProperty(window, 'parent', { configurable: true, value: {} })
    const execute = vi.fn(async () => ({ ok: true as const }))
    const teardown = attachMcpLiveSession({
      execute,
      snapshot: (revision) => ({ revision, fileName: 'new.md', dirty: false, selection: null }),
    })
    let release!: () => void
    const replacing = new Promise<void>((resolve) => {
      release = resolve
    })
    let replacement: Promise<void> | undefined
    try {
      bindEditor('replace-concurrent')
      await vi.waitFor(() => expect(mcp.pollWaiters.length).toBeGreaterThan(0))
      replacement = replaceLiveEditorDocument(() => replacing)
      await vi.waitFor(() =>
        expect(mcp.calls.some((call) => call.name === 'office_editor_reset_document')).toBe(true),
      )
      mcp.commands.push({
        commandId: 'after-open',
        operation: 'markdown.text.insert',
        arguments: {},
        baseRevision: 0,
      })
      mcp.pollWaiters.shift()!()
      await new Promise((resolve) => window.setTimeout(resolve, 50))
      expect(execute).not.toHaveBeenCalled()
      release()
      await replacement
      await vi.waitFor(() => expect(execute).toHaveBeenCalledTimes(1))
    } finally {
      release()
      await replacement
      teardown()
    }
  })

  it('gives each cold mount a distinct transport identity even when the host replays the same show result', async () => {
    Object.defineProperty(window, 'parent', { configurable: true, value: {} })
    const adapter = {
      execute: async () => ({ ok: true as const }),
      snapshot: (revision: number) => ({ revision, fileName: null, dirty: false, selection: null }),
    }
    const mounts: unknown[] = []
    for (let index = 0; index < 2; index++) {
      const teardown = attachMcpLiveSession(adapter)
      try {
        bindEditor('same-session', 'same-view')
        await vi.waitFor(() =>
          expect(
            mcp.calls.filter((call) => call.name === 'office_editor_poll').length,
          ).toBeGreaterThan(0),
        )
        mounts.push(mcp.calls.find((call) => call.name === 'office_editor_poll')!.arguments.mountId)
      } finally {
        teardown()
        mcp.calls = []
      }
    }
    expect(mounts[0]).toEqual(expect.any(String))
    expect(mounts[1]).not.toBe(mounts[0])
  })

  it('backs off a structured transient poll error and then resumes command delivery', async () => {
    Object.defineProperty(window, 'parent', { configurable: true, value: {} })
    mcp.pollError = 'internal_error'
    const teardown = attachMcpLiveSession({
      execute: async () => ({ ok: true }),
      snapshot: (revision) => ({ revision, fileName: null, dirty: false, selection: null }),
    })
    try {
      bindEditor('transient-error')
      await vi.waitFor(() =>
        expect(mcp.calls.some((call) => call.name === 'office_editor_poll')).toBe(true),
      )
      await new Promise((resolve) => window.setTimeout(resolve, 250))
      expect(mcp.calls.filter((call) => call.name === 'office_editor_poll')).toHaveLength(1)
      mcp.pollError = null
      await vi.waitFor(() =>
        expect(document.documentElement.dataset.liveEditorConnection).toBe('connected'),
      )
    } finally {
      teardown()
    }
  })

  it('rechecks an initially hidden conflicting view when it first becomes visible', async () => {
    Object.defineProperty(window, 'parent', { configurable: true, value: {} })
    const intersect = stubEditorIntersection(false)
    mcp.pollError = 'editor_view_conflict'
    const teardown = attachMcpLiveSession({
      execute: async () => ({ ok: true }),
      snapshot: (revision) => ({ revision, fileName: null, dirty: false, selection: null }),
    })
    try {
      bindEditor('initially-hidden')
      await vi.waitFor(() =>
        expect(document.documentElement.dataset.liveEditorConnection).toBe('editor_view_conflict'),
      )
      expect(mcp.calls.filter((call) => call.name === 'office_editor_poll')).toHaveLength(1)
      mcp.pollError = null
      intersect(true)
      await vi.waitFor(() =>
        expect(document.documentElement.dataset.liveEditorConnection).toBe('connected'),
      )
    } finally {
      teardown()
    }
  })

  it.each(['editor_view_conflict', 'session_not_found'])(
    'stops a rejected view instead of spinning on %s, including while hidden',
    async (error) => {
      Object.defineProperty(window, 'parent', { configurable: true, value: {} })
      mcp.pollError = error
      const teardown = attachMcpLiveSession({
        execute: async () => ({ ok: true }),
        snapshot: (revision) => ({ revision, fileName: null, dirty: false, selection: null }),
      })
      try {
        bindEditor('rejected-view')
        await vi.waitFor(() =>
          expect(mcp.calls.some((call) => call.name === 'office_editor_poll')).toBe(true),
        )
        Object.defineProperty(document, 'visibilityState', { configurable: true, value: 'hidden' })
        document.dispatchEvent(new Event('visibilitychange'))
        await new Promise((resolve) => window.setTimeout(resolve, 600))
        expect(mcp.calls.filter((call) => call.name === 'office_editor_poll')).toHaveLength(1)
        expect(document.documentElement.dataset.liveEditorConnection).toBe(error)
      } finally {
        teardown()
      }
    },
  )

  it('persists a generated binary file through the local server-tool save protocol', async () => {
    Object.defineProperty(window, 'parent', { configurable: true, value: {} })
    const teardown = attachMcpLiveSession({
      execute: async () => ({ ok: true }),
      snapshot: (revision) => ({
        revision,
        fileName: 'Quarterly Review.pptx',
        dirty: false,
        selection: null,
      }),
    })

    try {
      await vi.waitFor(() => expect(mcp.instance).not.toBeNull())
      bindEditor('session-save', 'view-save')
      await vi.waitFor(() =>
        expect(document.documentElement.dataset.liveEditorConnection).toBe('connected'),
      )
      await expect(
        saveLiveEditorFile({
          fileName: 'Quarterly Review.pptx',
          data: Uint8Array.from([1, 2, 3]).buffer,
          mode: 'save-as',
        }),
      ).resolves.toEqual({
        ok: true,
        path: '/tmp/tandemfolio/Quarterly Review.pptx',
      })
      expect(mcp.calls.filter((call) => call.name.includes('document_save'))).toEqual([
        {
          name: 'office_editor_begin_document_save',
          arguments: {
            sessionId: 'session-save',
            viewId: 'view-save',
            mountId: expect.any(String),
            fileName: 'Quarterly Review.pptx',
            size: 3,
            mode: 'save-as',
          },
        },
        {
          name: 'office_editor_write_document_save_chunk',
          arguments: {
            sessionId: 'session-save',
            viewId: 'view-save',
            mountId: expect.any(String),
            uploadId: 'save-session-save',
            offset: 0,
            data: 'AQID',
          },
        },
        {
          name: 'office_editor_commit_document_save',
          arguments: {
            sessionId: 'session-save',
            viewId: 'view-save',
            mountId: expect.any(String),
            uploadId: 'save-session-save',
          },
        },
      ])
      expect(mcp.downloads).toEqual([])
    } finally {
      teardown()
    }
  })

  it('reports whether the live editor is embedded through the shared display snapshot', () => {
    expect(getLiveEditorDisplayMode().embedded).toBe(false)

    Object.defineProperty(window, 'parent', { configurable: true, value: {} })

    expect(getLiveEditorDisplayMode().embedded).toBe(true)
  })

  it('publishes browser rendering skip and resume activity for heavy editor work', () => {
    Object.defineProperty(window, 'parent', { configurable: true, value: {} })
    stubEditorIntersection(true)
    const root = document.createElement('div')
    root.id = 'root'
    document.body.append(root)
    const activity: boolean[] = []
    const unsubscribe = subscribeLiveEditorActivity(() => activity.push(getLiveEditorActivity()))
    const teardown = attachMcpLiveSession({
      execute: async () => ({ ok: true }),
      snapshot: (revision) => ({
        revision,
        fileName: 'activity.pptx',
        dirty: false,
        selection: null,
      }),
    })

    try {
      expect(getLiveEditorActivity()).toBe(true)
      document
        .getElementById('root')
        ?.dispatchEvent(
          Object.assign(new Event('contentvisibilityautostatechange'), { skipped: true }),
        )
      expect(getLiveEditorActivity()).toBe(false)
      document
        .getElementById('root')
        ?.dispatchEvent(
          Object.assign(new Event('contentvisibilityautostatechange'), { skipped: false }),
        )
      expect(getLiveEditorActivity()).toBe(true)
      expect(activity).toEqual([false, true])
    } finally {
      teardown()
      unsubscribe()
      root.remove()
    }
  })

  it('keeps the shared display snapshot referentially stable until its state changes', () => {
    const standalone = getLiveEditorDisplayMode()

    expect(getLiveEditorDisplayMode()).toBe(standalone)

    Object.defineProperty(window, 'parent', { configurable: true, value: {} })
    const embedded = getLiveEditorDisplayMode()

    expect(embedded).not.toBe(standalone)
    expect(getLiveEditorDisplayMode()).toBe(embedded)
  })

  it('waits for the XLSX first commit and delivers its startup trace once after transport recovery', async () => {
    Object.defineProperty(window, 'parent', { configurable: true, value: {} })
    let resolveStartup!: (trace: {
      operation: 'xlsx.editor.cold_start'
      phases: {
        bootstrapMs: number
        univerCreateMs: number
        worksheetInstallMs: number
        firstCommitMs: number
      }
      bootstrapPhases: {
        resourceReceiveMs: number
        moduleGraphReadyMs: number
        reactMountMs: number
      }
    }) => void
    const startup = new Promise<Parameters<typeof resolveStartup>[0]>((resolve) => {
      resolveStartup = resolve
    })
    mcp.pollFailures = 1
    const teardown = attachMcpLiveSession({
      startupTrace: () => startup,
      execute: async () => ({ ok: true }),
      snapshot: (revision) => ({
        revision,
        fileName: null,
        dirty: false,
        selection: { sheet: 'Sheet1', range: 'A1' },
      }),
    })

    try {
      bindEditor('session-1')
      await new Promise((resolve) => window.setTimeout(resolve, 20))
      expect(mcp.calls).not.toContainEqual(expect.objectContaining({ name: 'office_editor_poll' }))

      const trace = {
        operation: 'xlsx.editor.cold_start' as const,
        phases: {
          bootstrapMs: 600,
          univerCreateMs: 8,
          worksheetInstallMs: 14,
          firstCommitMs: 60,
        },
        bootstrapPhases: {
          resourceReceiveMs: 25,
          moduleGraphReadyMs: 550,
          reactMountMs: 25,
        },
      }
      resolveStartup(trace)

      await vi.waitFor(
        () => {
          expect(mcp.calls.filter((call) => call.name === 'office_editor_poll')).toHaveLength(3)
        },
        { timeout: 1_500 },
      )
      const polls = mcp.calls.filter((call) => call.name === 'office_editor_poll')
      expect(polls[0].arguments.startupTrace).toEqual(trace)
      expect(polls[1].arguments.startupTrace).toEqual(trace)
      expect(polls[2].arguments).not.toHaveProperty('startupTrace')
    } finally {
      teardown()
    }
  })

  it('reports the embedded renderer size when its MCP connection mounts', async () => {
    Object.defineProperty(window, 'parent', { configurable: true, value: {} })

    const teardown = attachMcpLiveSession({
      execute: async () => ({ ok: true }),
      snapshot: (revision) => ({
        revision,
        fileName: null,
        dirty: false,
        selection: null,
      }),
    })

    await vi.waitFor(() => {
      expect(mcp.sizeNotifications).toEqual([
        { width: window.innerWidth, height: window.innerHeight },
      ])
    })
    teardown()
  })

  it('completes the first session poll before requesting fullscreen', async () => {
    Object.defineProperty(window, 'parent', { configurable: true, value: {} })
    const teardown = attachMcpLiveSession({
      execute: async () => ({ ok: true }),
      snapshot: (revision) => ({
        revision,
        fileName: null,
        dirty: false,
        selection: null,
      }),
    })

    try {
      bindEditor('session-1')
      await vi.waitFor(() => expect(mcp.requestedModes).toEqual(['fullscreen']))

      expect(mcp.events.indexOf('tool:office_editor_poll')).toBeLessThan(
        mcp.events.indexOf('display:fullscreen'),
      )
    } finally {
      teardown()
    }
  })

  it('uses an immediate bootstrap poll followed by a bounded waiting poll', async () => {
    Object.defineProperty(window, 'parent', { configurable: true, value: {} })
    const teardown = attachMcpLiveSession({
      execute: async () => ({ ok: true }),
      snapshot: (revision) => ({
        revision,
        fileName: null,
        dirty: false,
        selection: null,
      }),
    })

    try {
      bindEditor('session-1')
      await vi.waitFor(() => {
        expect(mcp.calls.filter((call) => call.name === 'office_editor_poll')).toHaveLength(2)
      })

      const polls = mcp.calls.filter((call) => call.name === 'office_editor_poll')
      expect(polls.map((call) => call.arguments.waitMs)).toEqual([0, 10_000])
    } finally {
      teardown()
    }
  })

  it('keeps a transiently hidden host view active while it still intersects the viewport', () => {
    Object.defineProperty(window, 'parent', { configurable: true, value: {} })
    Object.defineProperty(document, 'visibilityState', { configurable: true, value: 'visible' })
    stubEditorIntersection(true)
    const teardown = attachMcpLiveSession({
      execute: async () => ({ ok: true }),
      snapshot: (revision) => ({
        revision,
        fileName: null,
        dirty: false,
        selection: null,
      }),
    })

    Object.defineProperty(document, 'visibilityState', {
      configurable: true,
      value: 'hidden',
    })
    document.dispatchEvent(new Event('visibilitychange'))

    expect(document.documentElement.dataset.liveEditorActive).toBe('true')
    teardown()
  })

  it('suspends a visible host view after it leaves the nested viewport', () => {
    Object.defineProperty(window, 'parent', { configurable: true, value: {} })
    Object.defineProperty(document, 'visibilityState', { configurable: true, value: 'visible' })
    stubEditorIntersection(false)
    const teardown = attachMcpLiveSession({
      execute: async () => ({ ok: true }),
      snapshot: (revision) => ({
        revision,
        fileName: 'offscreen.pptx',
        dirty: false,
        selection: null,
      }),
    })

    expect(document.documentElement.dataset.liveEditorActive).toBe('false')
    teardown()
  })

  it('checkpoints once while leaving the viewport and resumes periodic recovery on return', async () => {
    vi.useFakeTimers({ toFake: ['setInterval', 'clearInterval'] })
    Object.defineProperty(window, 'parent', { configurable: true, value: {} })
    Object.defineProperty(document, 'visibilityState', { configurable: true, value: 'visible' })
    const setIntersecting = stubEditorIntersection(true)
    const recoverySnapshot = vi.fn(async () => ({
      fileName: 'offscreen.pptx',
      data: new Uint8Array([1, 2, 3]).buffer,
    }))
    const teardown = attachMcpLiveSession({
      execute: async () => ({ ok: true }),
      snapshot: (revision) => ({
        revision,
        fileName: 'offscreen.pptx',
        dirty: true,
        selection: null,
      }),
      recoverySnapshot,
    })

    try {
      bindEditor('session-1')
      await vi.waitFor(() => expect(mcp.requestedModes).toEqual(['fullscreen']))

      setIntersecting(false)
      await vi.advanceTimersByTimeAsync(4_000)
      expect(recoverySnapshot).toHaveBeenCalledTimes(1)

      setIntersecting(true)
      await vi.advanceTimersByTimeAsync(2_000)
      expect(recoverySnapshot).toHaveBeenCalledTimes(2)
    } finally {
      teardown()
    }
  })

  it('serializes each recovery version only once until the document changes', async () => {
    vi.useFakeTimers({ toFake: ['setInterval', 'clearInterval'] })
    Object.defineProperty(window, 'parent', { configurable: true, value: {} })
    stubEditorIntersection(true)
    let recoveryVersion = 1
    const recoverySnapshot = vi.fn(async () => ({
      fileName: 'versioned.pptx',
      data: new Uint8Array([1, 2, 3]).buffer,
    }))
    const teardown = attachMcpLiveSession({
      execute: async () => ({ ok: true }),
      snapshot: (revision) => ({
        revision,
        fileName: 'versioned.pptx',
        dirty: true,
        selection: null,
      }),
      recoverySnapshot,
      recoveryVersion: () => recoveryVersion,
    })

    try {
      bindEditor('session-1')
      await vi.waitFor(() => expect(mcp.requestedModes).toEqual(['fullscreen']))

      await vi.advanceTimersByTimeAsync(4_000)
      expect(recoverySnapshot).toHaveBeenCalledTimes(1)

      recoveryVersion += 1
      await vi.advanceTimersByTimeAsync(2_000)
      expect(recoverySnapshot).toHaveBeenCalledTimes(2)
    } finally {
      teardown()
    }
  })

  it('does not serialize a recovery version again after command acknowledgement stored it', async () => {
    vi.useFakeTimers({ toFake: ['setInterval', 'clearInterval'] })
    Object.defineProperty(window, 'parent', { configurable: true, value: {} })
    stubEditorIntersection(true)
    mcp.commands = [
      {
        commandId: 'command-1',
        baseRevision: 0,
        operation: 'pptx.slide.add_blank',
        arguments: { afterSlideIndex: 0 },
      },
    ]
    const recoverySnapshot = vi.fn(async () => ({
      fileName: 'acknowledged.pptx',
      data: new Uint8Array([4, 5, 6]).buffer,
    }))
    const teardown = attachMcpLiveSession({
      execute: async () => ({
        ok: true,
        recovery: {
          fileName: 'acknowledged.pptx',
          data: new Uint8Array([1, 2, 3]).buffer,
        },
      }),
      snapshot: (revision) => ({
        revision,
        fileName: 'acknowledged.pptx',
        dirty: true,
        selection: null,
      }),
      recoverySnapshot,
      recoveryVersion: () => 1,
    })

    try {
      bindEditor('session-1')
      await vi.waitFor(() => {
        expect(mcp.calls).toContainEqual({
          name: 'office_editor_acknowledge',
          arguments: expect.objectContaining({ sessionId: 'session-1', commandId: 'command-1' }),
        })
      })

      await vi.advanceTimersByTimeAsync(2_000)
      expect(recoverySnapshot).not.toHaveBeenCalled()
    } finally {
      teardown()
    }
  })

  it('keeps one bounded command poll while suppressing inactive-editor rendering', async () => {
    Object.defineProperty(window, 'parent', { configurable: true, value: {} })
    Object.defineProperty(document, 'visibilityState', { configurable: true, value: 'visible' })
    const setIntersecting = stubEditorIntersection(false)
    const teardown = attachMcpLiveSession({
      execute: async () => ({ ok: true }),
      snapshot: (revision) => ({
        revision,
        fileName: 'inactive.xlsx',
        dirty: false,
        selection: { sheet: 'Sheet1', range: 'A1' },
      }),
    })

    try {
      bindEditor('session-1')
      await vi.waitFor(() => {
        expect(mcp.calls.filter((call) => call.name === 'office_editor_poll')).toHaveLength(2)
      })

      Object.defineProperty(document, 'visibilityState', {
        configurable: true,
        value: 'hidden',
      })
      document.dispatchEvent(new Event('visibilitychange'))
      expect(document.documentElement.dataset.liveEditorActive).toBe('false')
      for (const release of mcp.pollWaiters.splice(0)) release()
      await vi.waitFor(() => {
        expect(mcp.calls.filter((call) => call.name === 'office_editor_poll')).toHaveLength(3)
      })
      expect(
        mcp.calls.filter((call) => call.name === 'office_editor_poll').at(-1)?.arguments.waitMs,
      ).toBe(10_000)

      Object.defineProperty(document, 'visibilityState', {
        configurable: true,
        value: 'visible',
      })
      document.dispatchEvent(new Event('visibilitychange'))
      expect(document.documentElement.dataset.liveEditorActive).toBe('false')
      setIntersecting(true)
      expect(document.documentElement.dataset.liveEditorActive).toBe('true')
    } finally {
      teardown()
    }
  })

  it('retains a 500 ms fallback only after a transport failure', async () => {
    Object.defineProperty(window, 'parent', { configurable: true, value: {} })
    mcp.pollFailures = 1
    const teardown = attachMcpLiveSession({
      execute: async () => ({ ok: true }),
      snapshot: (revision) => ({
        revision,
        fileName: null,
        dirty: false,
        selection: null,
      }),
    })

    try {
      bindEditor('session-1')
      await vi.waitFor(() => {
        expect(mcp.calls.filter((call) => call.name === 'office_editor_poll')).toHaveLength(1)
      })
      await new Promise((resolve) => window.setTimeout(resolve, 250))
      expect(mcp.calls.filter((call) => call.name === 'office_editor_poll')).toHaveLength(1)

      await vi.waitFor(
        () => {
          expect(mcp.calls.filter((call) => call.name === 'office_editor_poll')).toHaveLength(3)
        },
        { timeout: 1_000 },
      )
      const polls = mcp.calls.filter((call) => call.name === 'office_editor_poll')
      expect(polls.map((call) => call.arguments.waitMs)).toEqual([0, 0, 10_000])
    } finally {
      teardown()
    }
  })

  it('waits for a lease-bearing show result before polling and requesting fullscreen', async () => {
    Object.defineProperty(window, 'parent', { configurable: true, value: {} })
    const teardown = attachMcpLiveSession({
      execute: async () => ({ ok: true }),
      snapshot: (revision) => ({
        revision,
        fileName: null,
        dirty: false,
        selection: null,
      }),
    })

    try {
      await vi.waitFor(() => expect(mcp.sizeNotifications).toHaveLength(1))
      await new Promise((resolve) => window.setTimeout(resolve, 0))
      expect(mcp.calls).not.toContainEqual(expect.objectContaining({ name: 'office_editor_poll' }))
      expect(mcp.requestedModes).toEqual([])

      mcp.instance?.ontoolinput?.({ arguments: { sessionId: 'session-1' } })
      await new Promise((resolve) => window.setTimeout(resolve, 0))
      expect(mcp.calls).not.toContainEqual(expect.objectContaining({ name: 'office_editor_poll' }))
      expect(mcp.requestedModes).toEqual([])

      mcp.instance?.ontoolresult?.({
        structuredContent: { ok: true, sessionId: 'session-1', viewId: 'view-1' },
      })
      await vi.waitFor(() => expect(mcp.requestedModes).toEqual(['fullscreen']))
      expect(mcp.calls).toContainEqual({
        name: 'office_editor_poll',
        arguments: expect.objectContaining({
          sessionId: 'session-1',
          viewId: 'view-1',
          waitMs: 0,
        }),
      })
      expect(mcp.events.indexOf('tool:office_editor_poll')).toBeLessThan(
        mcp.events.indexOf('display:fullscreen'),
      )
    } finally {
      teardown()
    }
  })

  it('reconnects a restored host view when its show result replays the session identity', async () => {
    Object.defineProperty(window, 'parent', { configurable: true, value: {} })
    const teardown = attachMcpLiveSession({
      execute: async () => ({ ok: true }),
      snapshot: (revision) => ({
        revision,
        fileName: null,
        dirty: false,
        selection: null,
      }),
    })

    try {
      mcp.instance?.ontoolresult?.({
        structuredContent: { ok: true, sessionId: 'restored-session', viewId: 'restored-view' },
      })

      await vi.waitFor(() => {
        expect(mcp.calls).toContainEqual({
          name: 'office_editor_poll',
          arguments: expect.objectContaining({
            sessionId: 'restored-session',
            viewId: 'restored-view',
            waitMs: 0,
          }),
        })
      })
    } finally {
      teardown()
    }
  })

  it('retries the host handshake after a responsive layout switch rejects initialization', async () => {
    Object.defineProperty(window, 'parent', { configurable: true, value: {} })
    mcp.connectFailures = 1
    const teardown = attachMcpLiveSession({
      execute: async () => ({ ok: true }),
      snapshot: (revision) => ({
        revision,
        fileName: null,
        dirty: false,
        selection: null,
      }),
    })
    bindEditor('responsive-session')

    try {
      await vi.waitFor(
        () => {
          expect(mcp.calls).toContainEqual({
            name: 'office_editor_poll',
            arguments: expect.objectContaining({ sessionId: 'responsive-session', waitMs: 0 }),
          })
        },
        { timeout: 1_500 },
      )
    } finally {
      teardown()
    }
  })

  it('does not disconnect a session when a rejected duplicate view is torn down before polling', async () => {
    Object.defineProperty(window, 'parent', { configurable: true, value: {} })
    mcp.connectFailures = 1
    const teardown = attachMcpLiveSession({
      execute: async () => ({ ok: true }),
      snapshot: (revision) => ({
        revision,
        fileName: null,
        dirty: false,
        selection: null,
      }),
    })
    bindEditor('healthy-session')
    await Promise.resolve()

    teardown()

    expect(mcp.calls).not.toContainEqual({
      name: 'office_editor_disconnect',
      arguments: { sessionId: 'healthy-session' },
    })
  })

  it('keeps one mounted renderer bound to its original session when another show input arrives', async () => {
    Object.defineProperty(window, 'parent', { configurable: true, value: {} })
    const teardown = attachMcpLiveSession({
      execute: async () => ({ ok: true }),
      snapshot: (revision) => ({
        revision,
        fileName: 'previous.md',
        dirty: false,
        selection: null,
      }),
    })

    try {
      bindEditor('session-1')
      await vi.waitFor(() => {
        expect(mcp.calls).toContainEqual({
          name: 'office_editor_poll',
          arguments: expect.objectContaining({ sessionId: 'session-1' }),
        })
      })
      mcp.calls = []
      mcp.events = []

      mcp.instance?.ontoolinput?.({ arguments: { sessionId: 'session-2' } })
      mcp.instance?.ontoolresult?.({
        structuredContent: { ok: true, sessionId: 'session-2', viewId: 'view-session-2' },
      })
      await new Promise((resolve) => window.setTimeout(resolve, 0))

      expect(mcp.calls).not.toContainEqual({
        name: 'office_editor_disconnect',
        arguments: { sessionId: 'session-1' },
      })
      expect(mcp.calls).not.toContainEqual({
        name: 'office_editor_poll',
        arguments: expect.objectContaining({ sessionId: 'session-2' }),
      })
    } finally {
      teardown()
    }
  })

  it('reports the initial shared-renderer size before polling and requesting fullscreen', async () => {
    Object.defineProperty(window, 'parent', { configurable: true, value: {} })
    const teardown = attachMcpLiveSession({
      execute: async () => ({ ok: true }),
      snapshot: (revision) => ({
        revision,
        fileName: null,
        dirty: false,
        selection: null,
      }),
    })

    try {
      bindEditor('session-1')
      await vi.waitFor(() => {
        expect(mcp.sizeNotifications).toHaveLength(1)
        expect(mcp.requestedModes).toEqual(['fullscreen'])
      })

      expect(mcp.events.indexOf('size:initial')).toBeLessThan(
        mcp.events.indexOf('tool:office_editor_poll'),
      )
      expect(mcp.events.indexOf('tool:office_editor_poll')).toBeLessThan(
        mcp.events.indexOf('display:fullscreen'),
      )
    } finally {
      teardown()
    }
  })

  it('requests fullscreen once and acknowledges a renderer mutation', async () => {
    Object.defineProperty(window, 'parent', { configurable: true, value: {} })
    mcp.commands.push({
      commandId: 'command-1',
      baseRevision: 0,
      operation: 'set_cell_value',
      arguments: { address: 'A1', value: 'Ready' },
    })
    const teardown = attachMcpLiveSession({
      execute: async () => ({ ok: true, output: { changed: 1 } }),
      snapshot: (revision) => ({
        revision,
        fileName: 'book.xlsx',
        dirty: true,
        selection: { sheet: 'Sheet1', range: 'A1' },
      }),
    })
    bindEditor('session-1')

    await vi.waitFor(() => {
      expect(mcp.requestedModes).toEqual(['fullscreen'])
      expect(mcp.calls).toContainEqual({
        name: 'office_editor_acknowledge',
        arguments: {
          sessionId: 'session-1',
          viewId: 'view-session-1',
          mountId: expect.any(String),
          commandId: 'command-1',
          ok: true,
          output: { changed: 1 },
          timing: {
            hydrateMs: expect.any(Number),
            executeMs: expect.any(Number),
          },
          revision: 1,
          fileName: 'book.xlsx',
          dirty: true,
          selection: { sheet: 'Sheet1', range: 'A1' },
        },
      })
    })
    const acknowledgement = mcp.calls.find((call) => call.name === 'office_editor_acknowledge')!
    expect(
      (acknowledgement.arguments.timing as { hydrateMs: number; executeMs: number }).hydrateMs,
    ).toBeGreaterThanOrEqual(0)
    expect(
      (acknowledgement.arguments.timing as { hydrateMs: number; executeMs: number }).executeMs,
    ).toBeGreaterThanOrEqual(0)
    teardown()
  })

  it('hydrates a canonical staged-load operation before renderer execution', async () => {
    Object.defineProperty(window, 'parent', { configurable: true, value: {} })
    mcp.stagedBytes = new TextEncoder().encode('# Hydrated Markdown')
    mcp.commands.push({
      commandId: 'command-staged',
      baseRevision: 0,
      operation: 'markdown.document.load_staged',
      arguments: {
        blobId: 'markdown-blob',
        name: 'hydrated.md',
        size: mcp.stagedBytes.byteLength,
      },
    })
    let executed: {
      readonly operation: string
      readonly arguments: Record<string, unknown>
    } | null = null
    const teardown = attachMcpLiveSession({
      execute: async (command) => {
        executed = command
        return {
          ok: true,
          trace: {
            operation: 'markdown.document.load_staged',
            phases: {
              decodeMs: 0.1,
              parseMs: 1.2,
              tiptapStateInstallMs: 0.3,
              reactCommitMs: 0.4,
            },
          },
        }
      },
      snapshot: (revision) => ({
        revision,
        fileName: 'hydrated.md',
        dirty: false,
        selection: null,
      }),
    })
    bindEditor('session-1')

    await vi.waitFor(() => {
      expect(executed).toMatchObject({
        operation: 'markdown.document.load_staged',
        arguments: {
          blobId: 'markdown-blob',
          name: 'hydrated.md',
          size: mcp.stagedBytes!.byteLength,
          data: expect.any(ArrayBuffer),
        },
      })
    })
    const data = (executed as { arguments: { data: ArrayBuffer } } | null)!.arguments.data
    expect(new TextDecoder().decode(data)).toBe('# Hydrated Markdown')
    await vi.waitFor(() => {
      expect(mcp.calls).toContainEqual({
        name: 'office_editor_acknowledge',
        arguments: expect.objectContaining({
          commandId: 'command-staged',
          timing: {
            hydrateMs: expect.any(Number),
            executeMs: expect.any(Number),
            trace: {
              operation: 'markdown.document.load_staged',
              phases: {
                decodeMs: 0.1,
                parseMs: 1.2,
                tiptapStateInstallMs: 0.3,
                reactCommitMs: 0.4,
              },
            },
          },
        }),
      })
    })
    teardown()
  })

  it('hydrates internal staged PDF page insertion before renderer execution', async () => {
    Object.defineProperty(window, 'parent', { configurable: true, value: {} })
    mcp.stagedBytes = new TextEncoder().encode('%PDF-1.7\ninsert')
    mcp.commands.push({
      commandId: 'command-pdf-insert',
      baseRevision: 0,
      operation: 'pdf.page.insert_staged',
      arguments: {
        blobId: 'pdf-insert-blob',
        name: 'insert.pdf',
        size: mcp.stagedBytes.byteLength,
        afterPageIndex: 2,
      },
    })
    let executed: {
      readonly operation: string
      readonly arguments: Record<string, unknown>
    } | null = null
    const teardown = attachMcpLiveSession({
      execute: async (command) => {
        executed = command
        return { ok: true }
      },
      snapshot: (revision) => ({ revision, fileName: 'base.pdf', dirty: false, selection: null }),
    })
    bindEditor('session-1')

    await vi.waitFor(() => {
      expect(executed).toMatchObject({
        operation: 'pdf.page.insert_staged',
        arguments: {
          blobId: 'pdf-insert-blob',
          name: 'insert.pdf',
          size: mcp.stagedBytes!.byteLength,
          afterPageIndex: 2,
          data: expect.any(ArrayBuffer),
        },
      })
    })
    const data = (executed as { arguments: { data: ArrayBuffer } } | null)!.arguments.data
    expect(new TextDecoder().decode(data)).toBe('%PDF-1.7\ninsert')
    teardown()
  })

  it('reads a session-bound local asset through the connected app bridge', async () => {
    Object.defineProperty(window, 'parent', { configurable: true, value: {} })
    mcp.assetBytes = Uint8Array.from([0x89, 0x50, 0x4e, 0x47])
    const teardown = attachMcpLiveSession({
      execute: async () => ({ ok: true }),
      snapshot: (revision) => ({ revision, fileName: 'notes.md', dirty: false, selection: null }),
    })
    bindEditor('session-1')

    await vi.waitFor(() => expect(mcp.requestedModes).toEqual(['fullscreen']))
    await expect(readLiveEditorLocalAsset('root-1', 'assets/pixel.png')).resolves.toEqual({
      mime: 'image/png',
      data: mcp.assetBytes.buffer,
    })
    expect(mcp.calls).toContainEqual({
      name: 'office_editor_read_local_asset_chunk',
      arguments: {
        mountId: expect.any(String),
        sessionId: 'session-1',
        viewId: 'view-session-1',
        rootId: 'root-1',
        path: 'assets/pixel.png',
        offset: 0,
        length: 262144,
      },
    })
    teardown()
  })

  it('reads a bundled font through the connected format-neutral bridge', async () => {
    Object.defineProperty(window, 'parent', { configurable: true, value: {} })
    mcp.fontBytes = Uint8Array.from([0x4f, 0x54, 0x54, 0x4f, 0x01, 0x02])
    const teardown = attachMcpLiveSession({
      execute: async () => ({ ok: true }),
      snapshot: (revision) => ({ revision, fileName: 'font.pdf', dirty: false, selection: null }),
    })
    bindEditor('session-1')

    await vi.waitFor(() => expect(mcp.requestedModes).toEqual(['fullscreen']))
    await expect(
      readLiveEditorBundledFontAsset('NotoSansCJKsc-Regular-subset.otf'),
    ).resolves.toEqual(mcp.fontBytes.buffer)
    expect(mcp.calls).toContainEqual({
      name: 'office_editor_read_font_chunk',
      arguments: {
        fileName: 'NotoSansCJKsc-Regular-subset.otf',
        offset: 0,
        length: 262144,
      },
    })
    teardown()
  })
})
