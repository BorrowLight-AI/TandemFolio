import { describe, expect, it } from 'vitest'

import { createBrowserPdfHost, PDF_OPEN_EVENT } from '../src/renderer/host/browser-pdf-api'

describe('PDF browser staged-load boundary', () => {
  it('shows an empty session immediately and still opens a PDF staged later', async () => {
    const host = createBrowserPdfHost()

    await expect(host.api.consumePending()).resolves.toBeNull()

    const opened = new Promise<string>((resolve) => {
      window.addEventListener(
        PDF_OPEN_EVENT,
        (event) => resolve((event as CustomEvent<{ path: string }>).detail.path),
        { once: true },
      )
    })
    const data = Uint8Array.from([0x25, 0x50, 0x44, 0x46]).buffer
    const staged = host.stageFile('later.pdf', data)
    const path = await opened

    host.completeOpen(path)
    await expect(staged).resolves.toBeUndefined()
    expect(new Uint8Array(await host.api.readFile(path))).toEqual(new Uint8Array(data))
  })

  it('loads hydrated bytes through the canonical internal PDF operation', async () => {
    const host = createBrowserPdfHost()
    const data = Uint8Array.from([0x25, 0x50, 0x44, 0x46]).buffer

    const execution = host.adapter.execute({
      commandId: 'load-pdf',
      baseRevision: 0,
      operation: 'pdf.document.load_staged',
      arguments: {
        blobId: 'pdf-blob',
        name: 'loaded.pdf',
        size: data.byteLength,
        data,
      },
    })

    const path = await host.api.consumePending()
    expect(path).toBeTruthy()
    if (!path) throw new Error('The staged PDF did not enter the mounted-App load queue.')
    host.completeOpen(path)
    await expect(execution).resolves.toEqual({ ok: true, output: { opened: 'loaded.pdf' } })
    expect(new Uint8Array(await host.api.readFile(path))).toEqual(new Uint8Array(data))
    expect(host.adapter.snapshot(1)).toMatchObject({ revision: 1, fileName: 'loaded.pdf' })
  })

  it('keeps the retired open_local_file alias out of the mounted PDF queue', async () => {
    const host = createBrowserPdfHost()
    const data = Uint8Array.from([0x25, 0x50, 0x44, 0x46, 0x2d]).buffer

    await expect(
      host.adapter.execute({
        commandId: 'legacy-load-pdf',
        baseRevision: 0,
        operation: 'open_local_file',
        arguments: {
          blobId: 'legacy-pdf-blob',
          name: 'legacy.pdf',
          size: data.byteLength,
          data,
        },
      }),
    ).resolves.toMatchObject({ ok: false, error: 'unsupported_operation' })

    await expect(host.api.consumePending()).resolves.toBeNull()
    expect(host.adapter.snapshot(0)).toMatchObject({ revision: 0, fileName: null })
  })
})
