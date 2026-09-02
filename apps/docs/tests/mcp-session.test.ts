import { afterEach, describe, expect, it, vi } from 'vitest'

const mcp = vi.hoisted(() => ({
  commands: [] as Array<{
    commandId: string
    baseRevision: number
    operation: string
    arguments: Record<string, unknown>
  }>,
  fileBytes: new Uint8Array(),
  calls: [] as Array<{ name: string; arguments: Record<string, unknown> }>,
  events: [] as string[],
  hostContext: { availableDisplayModes: ['inline'] as string[], displayMode: 'inline' },
  sizeNotifications: [] as Array<{ width: number; height: number }>,
  pollWaiters: [] as Array<() => void>,
  instances: [] as Array<{
    ontoolinput?: (input: { arguments?: Record<string, unknown> }) => void
    ontoolresult?: (result: { structuredContent?: Record<string, unknown> }) => void
  }>,
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
      mcp.instances.push(this)
    }

    async connect() {
      if (this.autoResize) {
        mcp.sizeNotifications.push({ width: window.innerWidth, height: window.innerHeight })
      }
    }
    async close() {}
    async requestDisplayMode({ mode }: { mode: string }) {
      mcp.events.push(`display:${mode}`)
      return { mode }
    }
    getHostContext() {
      return mcp.hostContext
    }
    async callServerTool(call: { name: string; arguments: Record<string, unknown> }) {
      mcp.events.push(`tool:${call.name}`)
      mcp.calls.push(call)
      if (call.name === 'office_editor_poll') {
        const commands = mcp.commands.splice(0)
        if (commands.length > 0 || !(Number(call.arguments.waitMs) > 0)) {
          return { structuredContent: { ok: true, commands } }
        }
        return await new Promise<{ structuredContent: { ok: true; commands: never[] } }>(
          (resolve) => {
            mcp.pollWaiters.push(() => resolve({ structuredContent: { ok: true, commands: [] } }))
          },
        )
      }
      if (call.name === 'office_editor_disconnect') {
        mcp.pollWaiters.shift()?.()
        return { structuredContent: { ok: true } }
      }
      if (
        call.name === 'office_editor_read_file_chunk' ||
        call.name === 'office_editor_read_font_chunk'
      ) {
        const offset = call.arguments.offset as number
        const length = call.arguments.length as number
        const chunk = mcp.fileBytes.subarray(offset, offset + length)
        const nextOffset = offset + chunk.length
        return {
          structuredContent: {
            ok: true,
            offset,
            data: Buffer.from(chunk).toString('base64'),
            nextOffset,
            eof: nextOffset >= mcp.fileBytes.length,
          },
        }
      }
      if (call.name === 'office_editor_begin_recovery') {
        return { structuredContent: { ok: true, uploadId: 'upload-1' } }
      }
      if (call.name === 'office_editor_write_recovery_chunk') {
        const data = Buffer.from(call.arguments.data as string, 'base64')
        return {
          structuredContent: {
            ok: true,
            nextOffset: (call.arguments.offset as number) + data.length,
          },
        }
      }
      return { structuredContent: { ok: true } }
    }
  },
}))

import {
  attachMcpLiveSession as attachMcpSession,
  readLiveEditorBundledFontAsset as readMcpBundledFontAsset,
} from '@tandemfolio/host-bridge'

function bindEditor(sessionId = 'session-1', viewId = 'view-1'): void {
  mcp.instances[0].ontoolinput?.({ arguments: { sessionId } })
  mcp.instances[0].ontoolresult?.({
    structuredContent: { ok: true, sessionId, viewId },
  })
}

afterEach(() => {
  mcp.commands = []
  mcp.fileBytes = new Uint8Array()
  mcp.calls = []
  mcp.events = []
  mcp.hostContext = { availableDisplayModes: ['inline'], displayMode: 'inline' }
  mcp.sizeNotifications = []
  for (const release of mcp.pollWaiters.splice(0)) release()
  mcp.instances = []
  Object.defineProperty(window, 'parent', { configurable: true, value: window })
})

describe('MCP editor session', () => {
  it('reports the DOCX renderer size when its MCP connection mounts', async () => {
    Object.defineProperty(window, 'parent', { configurable: true, value: {} })

    const teardown = attachMcpSession({
      execute: async () => ({ ok: true }),
      snapshot: (revision) => ({ revision, fileName: null, dirty: false, selection: null }),
    })

    await vi.waitFor(() => {
      expect(mcp.sizeNotifications).toEqual([
        { width: window.innerWidth, height: window.innerHeight },
      ])
    })
    teardown()
  })

  it('completes the first DOCX session poll before requesting fullscreen', async () => {
    Object.defineProperty(window, 'parent', { configurable: true, value: {} })
    mcp.hostContext = {
      availableDisplayModes: ['inline', 'fullscreen'],
      displayMode: 'inline',
    }
    const teardown = attachMcpSession({
      execute: async () => ({ ok: true }),
      snapshot: (revision) => ({ revision, fileName: null, dirty: false, selection: null }),
    })

    try {
      bindEditor()
      await vi.waitFor(() => expect(mcp.events).toContain('display:fullscreen'))

      expect(mcp.events.indexOf('tool:office_editor_poll')).toBeLessThan(
        mcp.events.indexOf('display:fullscreen'),
      )
    } finally {
      teardown()
    }
  })

  it('uses the shared immediate bootstrap and bounded waiting poll schedule for DOCX', async () => {
    Object.defineProperty(window, 'parent', { configurable: true, value: {} })
    const teardown = attachMcpSession({
      execute: async () => ({ ok: true }),
      snapshot: (revision) => ({ revision, fileName: null, dirty: false, selection: null }),
    })

    try {
      bindEditor()
      await vi.waitFor(() => {
        expect(mcp.calls.filter((call) => call.name === 'office_editor_poll')).toHaveLength(2)
      })

      const polls = mcp.calls.filter((call) => call.name === 'office_editor_poll')
      expect(polls.map((call) => call.arguments.waitMs)).toEqual([0, 10_000])
    } finally {
      teardown()
    }
  })

  it('waits for the lease-bearing DOCX show result before polling and requesting fullscreen', async () => {
    Object.defineProperty(window, 'parent', { configurable: true, value: {} })
    mcp.hostContext = {
      availableDisplayModes: ['inline', 'fullscreen'],
      displayMode: 'inline',
    }
    const teardown = attachMcpSession({
      execute: async () => ({ ok: true }),
      snapshot: (revision) => ({ revision, fileName: null, dirty: false, selection: null }),
    })

    try {
      await vi.waitFor(() => expect(mcp.sizeNotifications).toHaveLength(1))
      await new Promise((resolve) => window.setTimeout(resolve, 0))
      expect(mcp.calls).not.toContainEqual(expect.objectContaining({ name: 'office_editor_poll' }))
      expect(mcp.events).not.toContain('display:fullscreen')

      mcp.instances[0].ontoolinput?.({ arguments: { sessionId: 'session-1' } })
      await new Promise((resolve) => window.setTimeout(resolve, 0))
      expect(mcp.calls).not.toContainEqual(expect.objectContaining({ name: 'office_editor_poll' }))
      expect(mcp.events).not.toContain('display:fullscreen')

      mcp.instances[0].ontoolresult?.({
        structuredContent: { ok: true, sessionId: 'session-1', viewId: 'view-1' },
      })
      await vi.waitFor(() => expect(mcp.events).toContain('display:fullscreen'))
      expect(mcp.events.indexOf('tool:office_editor_poll')).toBeLessThan(
        mcp.events.indexOf('display:fullscreen'),
      )
    } finally {
      teardown()
    }
  })

  it('reads an external bundled font without embedding it in the editor HTML', async () => {
    Object.defineProperty(window, 'parent', { configurable: true, value: {} })
    mcp.fileBytes = new Uint8Array([0, 1, 0, 0, 8, 9, 10])
    const teardown = attachMcpSession({
      execute: async () => ({ ok: true }),
      snapshot: (revision) => ({ revision, fileName: null, dirty: false, selection: null }),
    })

    await expect(readMcpBundledFontAsset('Carlito-Regular.ttf')).resolves.toEqual(
      mcp.fileBytes.buffer,
    )
    expect(mcp.calls).toContainEqual({
      name: 'office_editor_read_font_chunk',
      arguments: { fileName: 'Carlito-Regular.ttf', offset: 0, length: 262_144 },
    })
    teardown()
  })

  it('negatively acknowledges an unsupported renderer operation', async () => {
    Object.defineProperty(window, 'parent', { configurable: true, value: {} })
    mcp.commands.push({
      commandId: 'command-1',
      baseRevision: 0,
      operation: 'unknown_operation',
      arguments: {},
    })

    const teardown = attachMcpSession({
      execute: async () => ({
        ok: false,
        error: 'unsupported_operation',
        message: 'Unsupported operation: unknown_operation',
      }),
      snapshot: (revision) => ({
        revision,
        fileName: null,
        dirty: false,
        selection: null,
      }),
    })
    bindEditor()

    await vi.waitFor(() => {
      expect(mcp.calls).toContainEqual({
        name: 'office_editor_acknowledge',
        arguments: {
          sessionId: 'session-1',
          viewId: 'view-1',
          commandId: 'command-1',
          ok: false,
          error: 'unsupported_operation',
          message: 'Unsupported operation: unknown_operation',
        },
      })
    })
    teardown()
  })

  it('keeps the mounted DOCX context on its original Session/view lease', async () => {
    Object.defineProperty(window, 'parent', { configurable: true, value: {} })
    const teardown = attachMcpSession({
      execute: async () => ({ ok: true }),
      snapshot: (revision) => ({
        revision,
        fileName: 'kept.docx',
        dirty: true,
        selection: { from: 4, to: 4, empty: true },
      }),
    })
    bindEditor('session-1', 'view-1')

    await vi.waitFor(() => {
      expect(mcp.calls).toContainEqual({
        name: 'office_editor_poll',
        arguments: {
          sessionId: 'session-1',
          viewId: 'view-1',
          fileName: 'kept.docx',
          dirty: true,
          selection: { from: 4, to: 4, empty: true },
          waitMs: 0,
        },
      })
    })
    mcp.instances[0].ontoolinput?.({ arguments: { sessionId: 'session-2' } })
    mcp.instances[0].ontoolresult?.({
      structuredContent: { ok: true, sessionId: 'session-2', viewId: 'view-2' },
    })
    await new Promise((resolve) => window.setTimeout(resolve, 0))

    expect(mcp.calls).not.toContainEqual({
      name: 'office_editor_disconnect',
      arguments: { sessionId: 'session-1', viewId: 'view-1' },
    })
    expect(mcp.calls).not.toContainEqual({
      name: 'office_editor_poll',
      arguments: expect.objectContaining({ sessionId: 'session-2' }),
    })
    teardown()
  })

  it('hydrates a canonical staged DOCX load before passing it to the format adapter', async () => {
    Object.defineProperty(window, 'parent', { configurable: true, value: {} })
    mcp.fileBytes = new Uint8Array([80, 75, 3, 4, 1, 2, 3, 4])
    mcp.commands.push({
      commandId: 'open-1',
      baseRevision: 0,
      operation: 'docx.document.load_staged',
      arguments: { blobId: 'blob-1', name: 'local.docx', size: mcp.fileBytes.length },
    })
    let received: Record<string, unknown> | undefined

    const teardown = attachMcpSession({
      execute: async (command) => {
        received = command.arguments
        return { ok: true }
      },
      snapshot: (revision) => ({
        revision,
        fileName: 'local.docx',
        dirty: false,
        selection: null,
      }),
    })
    bindEditor()

    await vi.waitFor(() => {
      expect(received).toMatchObject({
        blobId: 'blob-1',
        name: 'local.docx',
        size: mcp.fileBytes.length,
        data: expect.any(ArrayBuffer),
      })
      expect(Array.from(new Uint8Array(received!.data as ArrayBuffer))).toEqual(
        Array.from(mcp.fileBytes),
      )
    })
    teardown()
  })

  it('hydrates a staged DOCX image before passing it to the format adapter', async () => {
    Object.defineProperty(window, 'parent', { configurable: true, value: {} })
    mcp.fileBytes = new TextEncoder().encode('GIF87a')
    mcp.commands.push({
      commandId: 'image-1',
      baseRevision: 0,
      operation: 'docx.image.insert_staged',
      arguments: {
        blobId: 'image-blob',
        name: 'logo.gif',
        size: mcp.fileBytes.length,
        afterBlockIndex: 0,
        widthPx: 320,
        heightPx: 180,
        alignment: 'center',
      },
    })
    let received: Record<string, unknown> | undefined

    const teardown = attachMcpSession({
      execute: async (command) => {
        received = command.arguments
        return { ok: true }
      },
      snapshot: (revision) => ({
        revision,
        fileName: 'local.docx',
        dirty: true,
        selection: null,
      }),
    })
    bindEditor()

    await vi.waitFor(() => {
      expect(received).toMatchObject({
        blobId: 'image-blob',
        name: 'logo.gif',
        size: mcp.fileBytes.length,
        afterBlockIndex: 0,
        widthPx: 320,
        heightPx: 180,
        alignment: 'center',
        data: expect.any(ArrayBuffer),
      })
      expect(Array.from(new Uint8Array(received!.data as ArrayBuffer))).toEqual(
        Array.from(mcp.fileBytes),
      )
    })
    teardown()
  })

  it('hydrates staged DOCX replacement bytes before passing them to the format adapter', async () => {
    Object.defineProperty(window, 'parent', { configurable: true, value: {} })
    mcp.fileBytes = new TextEncoder().encode('GIF87a')
    mcp.commands.push({
      commandId: 'replace-image-1',
      baseRevision: 0,
      operation: 'docx.image.replace_staged',
      arguments: {
        blobId: 'replacement-blob',
        name: 'replacement.gif',
        size: mcp.fileBytes.length,
        imageBlockIndex: 3,
        widthPx: 400,
        heightPx: 225,
      },
    })
    let received: Record<string, unknown> | undefined

    const teardown = attachMcpSession({
      execute: async (command) => {
        received = command.arguments
        return { ok: true }
      },
      snapshot: (revision) => ({
        revision,
        fileName: 'local.docx',
        dirty: true,
        selection: null,
      }),
    })
    bindEditor()

    await vi.waitFor(() => {
      expect(received).toMatchObject({
        blobId: 'replacement-blob',
        name: 'replacement.gif',
        size: mcp.fileBytes.length,
        imageBlockIndex: 3,
        widthPx: 400,
        heightPx: 225,
        data: expect.any(ArrayBuffer),
      })
    })
    teardown()
  })

  it('commits the renderer recovery snapshot before acknowledging a mutation', async () => {
    Object.defineProperty(window, 'parent', { configurable: true, value: {} })
    mcp.commands.push({
      commandId: 'command-1',
      baseRevision: 0,
      operation: 'insert_text',
      arguments: { text: 'hello' },
    })
    const recovery = new Uint8Array([80, 75, 3, 4, 9, 8, 7, 6]).buffer

    const teardown = attachMcpSession({
      execute: async () => ({
        ok: true,
        recovery: { fileName: 'draft.docx', data: recovery },
      }),
      snapshot: (revision) => ({
        revision,
        fileName: 'draft.docx',
        dirty: true,
        selection: null,
      }),
    })
    bindEditor()

    await vi.waitFor(() => {
      const names = mcp.calls.map((call) => call.name)
      expect(names).toContain('office_editor_commit_recovery')
      expect(names.indexOf('office_editor_commit_recovery')).toBeLessThan(
        names.indexOf('office_editor_acknowledge'),
      )
      expect(mcp.calls).toContainEqual({
        name: 'office_editor_write_recovery_chunk',
        arguments: {
          sessionId: 'session-1',
          viewId: 'view-1',
          uploadId: 'upload-1',
          offset: 0,
          data: Buffer.from(recovery).toString('base64'),
        },
      })
    })
    teardown()
  })

  it('periodically checkpoints dirty user edits without an agent command', async () => {
    vi.useFakeTimers()
    try {
      Object.defineProperty(window, 'parent', { configurable: true, value: {} })
      const recovery = new Uint8Array([80, 75, 3, 4, 5, 6]).buffer
      const teardown = attachMcpSession({
        execute: async () => ({ ok: true }),
        snapshot: (revision) => ({
          revision,
          fileName: 'manual.docx',
          dirty: true,
          selection: null,
        }),
        recoverySnapshot: async () => ({ fileName: 'manual.docx', data: recovery }),
      })
      bindEditor()
      await Promise.resolve()

      await vi.advanceTimersByTimeAsync(2_050)

      expect(mcp.calls.map((call) => call.name)).toContain('office_editor_commit_recovery')
      teardown()
    } finally {
      vi.useRealTimers()
    }
  })
})
