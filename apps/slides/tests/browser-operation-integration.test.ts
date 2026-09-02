import { readFile } from 'node:fs/promises'
import { describe, expect, it, vi } from 'vitest'

import { createBrowserSlidesHost } from '../src/renderer/host/browser-slides-api'
import type { OpenResult } from '../src/shared/ipc'
import type { ShapeRenderNode } from '@genoffice/pptx-render'
import { BrowserPresentation } from '../src/renderer/host/browser-presentation'

describe('PPTX browser operation integration', () => {
  it('writes the first PPTX save to the selected local file before reporting success', async () => {
    let persisted = new ArrayBuffer(0)
    const writable = {
      write: vi.fn(async (data: FileSystemWriteChunkType) => {
        if (data instanceof ArrayBuffer) persisted = data.slice(0)
        else if (ArrayBuffer.isView(data)) {
          persisted = data.buffer.slice(
            data.byteOffset,
            data.byteOffset + data.byteLength,
          ) as ArrayBuffer
        } else if (data instanceof Blob) persisted = await data.arrayBuffer()
        else throw new Error('Unexpected PPTX write payload.')
      }),
      close: vi.fn(async () => undefined),
    }
    const handle = {
      kind: 'file' as const,
      name: 'Quarterly Review.pptx',
      createWritable: vi.fn(async () => writable as unknown as FileSystemWritableFileStream),
      getFile: vi.fn(async () => new File([persisted], 'Quarterly Review.pptx')),
    } as unknown as FileSystemFileHandle
    const fakeWindow = {
      location: { href: 'https://example.test/slides/index.html' },
      showSaveFilePicker: vi.fn(async () => handle),
    } as unknown as Window
    Object.assign(fakeWindow, { parent: fakeWindow })
    vi.stubGlobal('window', fakeWindow)

    try {
      const host = createBrowserSlidesHost()
      await host.adapter.execute({
        commandId: 'blank',
        baseRevision: 0,
        operation: 'pptx.document.create_blank',
        arguments: {},
      })

      await expect(host.api.save()).resolves.toEqual({
        ok: true,
        path: 'Quarterly Review.pptx',
      })
      expect(writable.close).toHaveBeenCalledOnce()
      expect(persisted.byteLength).toBeGreaterThan(0)
      const reopened = await BrowserPresentation.open('Quarterly Review.pptx', persisted)
      expect(reopened.slideCount).toBe(1)
    } finally {
      vi.unstubAllGlobals()
    }
  })

  it('overwrites the selected PPTX file on later saves without reopening the picker', async () => {
    let persisted = new ArrayBuffer(0)
    const createWritable = vi.fn(async () => ({
      write: async (data: FileSystemWriteChunkType) => {
        if (!(data instanceof ArrayBuffer)) throw new Error('Expected PPTX ArrayBuffer bytes.')
        persisted = data.slice(0)
      },
      close: async () => undefined,
    }) as unknown as FileSystemWritableFileStream)
    const handle = {
      kind: 'file' as const,
      name: 'Quarterly Review.pptx',
      createWritable,
    } as unknown as FileSystemFileHandle
    const showSaveFilePicker = vi.fn(async () => handle)
    const fakeWindow = {
      location: { href: 'https://example.test/slides/index.html' },
      showSaveFilePicker,
    } as unknown as Window
    Object.assign(fakeWindow, { parent: fakeWindow })
    vi.stubGlobal('window', fakeWindow)

    try {
      const host = createBrowserSlidesHost()
      await host.adapter.execute({
        commandId: 'blank',
        baseRevision: 0,
        operation: 'pptx.document.create_blank',
        arguments: {},
      })

      await host.api.save()
      await host.adapter.execute({
        commandId: 'slide',
        baseRevision: 1,
        operation: 'pptx.slide.add_blank',
        arguments: { afterSlideIndex: 0 },
      })
      await host.api.save()

      expect(showSaveFilePicker).toHaveBeenCalledOnce()
      expect(createWritable).toHaveBeenCalledTimes(2)
      const reopened = await BrowserPresentation.open('Quarterly Review.pptx', persisted)
      expect(reopened.slideCount).toBe(2)
    } finally {
      vi.unstubAllGlobals()
    }
  })

  it('saves a PPTX copy to a newly selected file and adopts its name', async () => {
    const bytesByName = new Map<string, ArrayBuffer>()
    const fileHandle = (name: string): FileSystemFileHandle =>
      ({
        kind: 'file' as const,
        name,
        createWritable: async () =>
          ({
            write: async (data: FileSystemWriteChunkType) => {
              if (!(data instanceof ArrayBuffer)) throw new Error('Expected PPTX ArrayBuffer bytes.')
              bytesByName.set(name, data.slice(0))
            },
            close: async () => undefined,
          }) as unknown as FileSystemWritableFileStream,
      }) as unknown as FileSystemFileHandle
    const firstHandle = fileHandle('Quarterly Review.pptx')
    const renamedHandle = fileHandle('Renamed Deck.pptx')
    const showSaveFilePicker = vi
      .fn<() => Promise<FileSystemFileHandle>>()
      .mockResolvedValueOnce(firstHandle)
      .mockResolvedValueOnce(renamedHandle)
    const fakeWindow = {
      location: { href: 'https://example.test/slides/index.html' },
      showSaveFilePicker,
    } as unknown as Window
    Object.assign(fakeWindow, { parent: fakeWindow })
    vi.stubGlobal('window', fakeWindow)

    try {
      const host = createBrowserSlidesHost()
      await host.adapter.execute({
        commandId: 'blank',
        baseRevision: 0,
        operation: 'pptx.document.create_blank',
        arguments: {},
      })
      await host.api.save()

      await expect(host.api.saveAs('Quarterly Review.pptx')).resolves.toEqual({
        ok: true,
        path: 'Renamed Deck.pptx',
      })
      expect(host.adapter.snapshot(2).fileName).toBe('Renamed Deck.pptx')
      expect(showSaveFilePicker).toHaveBeenCalledTimes(2)
      expect(bytesByName.get('Quarterly Review.pptx')?.byteLength).toBeGreaterThan(0)
      expect(bytesByName.get('Renamed Deck.pptx')?.byteLength).toBeGreaterThan(0)
    } finally {
      vi.unstubAllGlobals()
    }
  })

  it('does not report a PPTX save when the destination picker is cancelled', async () => {
    const fakeWindow = {
      location: { href: 'https://example.test/slides/index.html' },
      showSaveFilePicker: vi.fn(async () => {
        throw new DOMException('Cancelled', 'AbortError')
      }),
    } as unknown as Window
    Object.assign(fakeWindow, { parent: fakeWindow })
    vi.stubGlobal('window', fakeWindow)

    try {
      const host = createBrowserSlidesHost()
      await host.adapter.execute({
        commandId: 'blank',
        baseRevision: 0,
        operation: 'pptx.document.create_blank',
        arguments: {},
      })

      await expect(host.api.save()).resolves.toEqual({ ok: false })
      expect(host.adapter.snapshot(1).fileName).toBeNull()
    } finally {
      vi.unstubAllGlobals()
    }
  })

  it('advances its recovery version only after the dirty presentation changes', async () => {
    const host = createBrowserSlidesHost()
    const recoveryVersion = (): number => {
      const version = host.adapter.recoveryVersion?.()
      if (typeof version !== 'number') throw new Error('PPTX recovery version is unavailable.')
      return version
    }

    await host.adapter.execute({
      commandId: 'blank',
      baseRevision: 0,
      operation: 'pptx.document.create_blank',
      arguments: {},
    })
    expect(recoveryVersion()).toBe(0)

    await host.adapter.execute({
      commandId: 'slide-1',
      baseRevision: 1,
      operation: 'pptx.slide.add_blank',
      arguments: { afterSlideIndex: 0 },
    })
    const changedVersion = recoveryVersion()
    expect(changedVersion).toBeGreaterThan(0)

    await host.adapter.recoverySnapshot?.()
    await host.adapter.recoverySnapshot?.()
    expect(recoveryVersion()).toBe(changedVersion)

    await host.adapter.execute({
      commandId: 'slide-2',
      baseRevision: 2,
      operation: 'pptx.slide.add_blank',
      arguments: { afterSlideIndex: 1 },
    })
    expect(recoveryVersion()).toBeGreaterThan(changedVersion)
  })

  it('exports retained image and PDF output through browser download primitives', async () => {
    const host = createBrowserSlidesHost()
    const downloads: string[] = []
    vi.stubGlobal('document', {
      createElement: vi.fn((tag: string) => {
        if (tag !== 'a') throw new Error(`Unexpected element: ${tag}`)
        return {
          href: '',
          download: '',
          click() {
            downloads.push(this.download)
          },
        }
      }),
    })
    const png =
      'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9WlRjGQAAAAASUVORK5CYII='

    try {
      await expect(host.api.pickExportDir()).resolves.toBe('Downloads')
      await expect(
        host.api.exportImages({ dir: 'Downloads', baseName: 'Deck', pngsBase64: [png, png] }),
      ).resolves.toEqual({
        ok: true,
        paths: ['Downloads/Deck-01.png', 'Downloads/Deck-02.png'],
      })
      await expect(host.api.pickExportPdfPath('Deck.pdf')).resolves.toBe('Deck.pdf')
      await expect(
        host.api.exportPdf({ filePath: 'Deck.pdf', pngsBase64: [png], widthPx: 16, heightPx: 9 }),
      ).resolves.toEqual({ ok: true, path: 'Deck.pdf' })
      expect(downloads).toEqual(['Deck-01.png', 'Deck-02.png', 'Deck.pdf'])
    } finally {
      vi.unstubAllGlobals()
    }
  })

  it('prints retained full-page, handout, and notes layouts through an isolated browser frame', async () => {
    const host = createBrowserSlidesHost()
    const printed: string[] = []
    const frames: Array<{ srcdoc: string; onload: (() => void) | null; remove: () => void }> = []
    vi.stubGlobal('document', {
      createElement: vi.fn((tag: string) => {
        if (tag !== 'iframe') throw new Error(`Unexpected element: ${tag}`)
        const frame = {
          srcdoc: '',
          onload: null as (() => void) | null,
          style: { cssText: '' },
          contentWindow: {
            focus: vi.fn(),
            print: () => printed.push(frame.srcdoc),
          },
          remove: vi.fn(),
        }
        frames.push(frame)
        return frame
      }),
      body: {
        append: (frame: { onload: (() => void) | null }) => queueMicrotask(() => frame.onload?.()),
      },
    })

    try {
      await expect(
        host.api.printSlides({
          pngsBase64: ['AAAA', 'BBBB'],
          widthPx: 16,
          heightPx: 9,
          layout: 'notes',
          notes: ['First <note>', 'Second'],
        }),
      ).resolves.toEqual({ ok: true })
      expect(printed).toHaveLength(1)
      expect(printed[0]).toContain('First &lt;note&gt;')
      expect(printed[0]).toContain('data:image/png;base64,AAAA')
      expect(frames[0]?.remove).toHaveBeenCalledOnce()
    } finally {
      vi.unstubAllGlobals()
    }
  })

  it('coordinates retained presenter and audience events through the browser host', async () => {
    const audience = { close: vi.fn(), closed: false }
    const fakeWindow = {
      location: { href: 'https://example.test/slides/index.html' },
      open: vi.fn(() => audience),
    } as unknown as Window
    Object.assign(fakeWindow, { parent: fakeWindow })
    class FakeBroadcastChannel {
      onmessage: ((event: MessageEvent) => void) | null = null
      readonly postMessage = vi.fn()
      readonly close = vi.fn()
    }
    vi.stubGlobal('window', fakeWindow)
    vi.stubGlobal('BroadcastChannel', FakeBroadcastChannel)

    try {
      const host = createBrowserSlidesHost()
      await host.adapter.execute({
        commandId: 'blank',
        baseRevision: 0,
        operation: 'pptx.document.create_blank',
        arguments: {},
      })
      const sync = {
        idx: 0,
        played: 0,
        playing: false,
        fresh: true,
        ended: false,
        black: false,
      }
      const onSync = vi.fn()
      const onInk = vi.fn()
      const onNav = vi.fn()
      host.api.onShowSync(onSync)
      host.api.onShowInk(onInk)
      host.api.onAudienceNav(onNav)

      await expect(host.api.presenterStart()).resolves.toEqual({ audience: true })
      expect(fakeWindow.open).toHaveBeenCalledWith(
        expect.stringContaining('mode=audience'),
        'genoffice-pptx-audience',
        'popup',
      )
      host.api.presenterSync(sync)
      host.api.presenterInk({ type: 'clear' })
      host.api.audienceNav('next')
      expect(onSync).toHaveBeenCalledWith(sync)
      expect(onInk).toHaveBeenCalledWith({ type: 'clear' })
      expect(onNav).toHaveBeenCalledWith('next')
      await expect(host.api.audienceReady()).resolves.toEqual(sync)
      await expect(host.api.presenterSwap()).resolves.toBe(false)
      await host.api.presenterEnd()
      expect(audience.close).toHaveBeenCalledOnce()
    } finally {
      vi.unstubAllGlobals()
    }
  })

  it('sets and reads PPTX animations through the registry-owned browser primitive', async () => {
    const host = createBrowserSlidesHost()
    let opened: OpenResult | null = null
    host.api.onOpened((result) => {
      opened = result
    })
    await host.adapter.execute({
      commandId: 'blank',
      baseRevision: 0,
      operation: 'pptx.document.create_blank',
      arguments: {},
    })
    const added = await host.adapter.execute({
      commandId: 'shape',
      baseRevision: 1,
      operation: 'pptx.object.add',
      arguments: {
        slideIndex: 0,
        kind: 'rect',
        xEmu: 100,
        yEmu: 100,
        widthEmu: 1_000,
        heightEmu: 1_000,
      },
    })
    const objectId = (added as unknown as { output: { objectId: string } }).output.objectId
    await expect(host.api.getShapeKeys(0)).resolves.toEqual([
      { sourceId: objectId, spid: expect.any(Number), name: expect.any(String) },
    ])
    await expect(
      host.adapter.execute({
        commandId: 'animation',
        baseRevision: 2,
        operation: 'pptx.animation.set',
        arguments: {
          slideIndex: 0,
          animations: [
            { objectId, effect: 'fade', trigger: 'onClick', durationMs: 500, delayMs: 0 },
          ],
        },
      }),
    ).resolves.toMatchObject({
      ok: true,
      output: { animationCount: 1 },
      recovery: { fileName: 'Untitled.pptx' },
    })
    await expect(host.api.getAnimations(0)).resolves.toEqual([
      expect.objectContaining({
        sourceId: objectId,
        effect: 'fade',
        trigger: 'onClick',
        durationMs: 500,
        delayMs: 0,
      }),
    ])
    expect((opened as OpenResult | null)?.slides[0]).toBeTruthy()
    await host.api.undo()
    await expect(host.api.getAnimations(0)).resolves.toEqual([])
  })

  it('sets hyperlinks, header/footer, and adds a slide with a retained layout', async () => {
    const host = createBrowserSlidesHost()
    let opened: OpenResult | null = null
    host.api.onOpened((result) => {
      opened = result
    })
    await host.adapter.execute({
      commandId: 'blank',
      baseRevision: 0,
      operation: 'pptx.document.create_blank',
      arguments: {},
    })
    const added = await host.adapter.execute({
      commandId: 'shape',
      baseRevision: 1,
      operation: 'pptx.object.add',
      arguments: {
        slideIndex: 0,
        kind: 'rect',
        xEmu: 100,
        yEmu: 100,
        widthEmu: 1_000,
        heightEmu: 1_000,
      },
    })
    const objectId = (added as unknown as { output: { objectId: string } }).output.objectId
    const linked = await host.adapter.execute({
      commandId: 'link',
      baseRevision: 2,
      operation: 'pptx.hyperlink.set',
      arguments: { slideIndex: 0, objectId, target: { kind: 'url', url: 'https://example.com' } },
    })
    expect(linked).toMatchObject({ ok: true, output: { objectId: expect.any(String) } })
    await expect(
      host.api.getLink(0, (linked as unknown as { output: { objectId: string } }).output.objectId),
    ).resolves.toEqual({ kind: 'url', url: 'https://example.com' })
    await expect(
      host.adapter.execute({
        commandId: 'footer',
        baseRevision: 3,
        operation: 'pptx.slide.apply_header_footer',
        arguments: { footer: 'Confidential', slideNumber: true, date: null, automaticDate: false },
      }),
    ).resolves.toMatchObject({ ok: true, output: { changed: 1 } })
    const layouts = await host.api.getLayouts()
    const layoutPath = layouts?.layouts[0]?.path
    expect(layoutPath).toBeTruthy()
    await expect(
      host.adapter.execute({
        commandId: 'layout-slide',
        baseRevision: 4,
        operation: 'pptx.slide.add_with_layout',
        arguments: { afterSlideIndex: 0, layoutPath },
      }),
    ).resolves.toMatchObject({
      ok: true,
      output: { slideIndex: 1, slideCount: 2 },
      recovery: { fileName: 'Untitled.pptx' },
    })
    expect((opened as OpenResult | null)?.slides).toHaveLength(2)
    await host.api.undo()
    await host.api.undo()
    await host.api.undo()
    expect(host.adapter.snapshot(8)).toMatchObject({ dirty: true })
  })

  it('applies a PPTX table preset style through the registry and UI primitive', async () => {
    const host = createBrowserSlidesHost()
    await host.adapter.execute({
      commandId: 'blank',
      baseRevision: 0,
      operation: 'pptx.document.create_blank',
      arguments: {},
    })
    const added = await host.adapter.execute({
      commandId: 'table',
      baseRevision: 1,
      operation: 'pptx.table.add',
      arguments: {
        slideIndex: 0,
        rows: 2,
        columns: 2,
        xEmu: 100,
        yEmu: 100,
        widthEmu: 4_000,
        heightEmu: 2_000,
      },
    })
    const tableId = (added as unknown as { output: { tableId: string } }).output.tableId
    await expect(
      host.adapter.execute({
        commandId: 'style',
        baseRevision: 2,
        operation: 'pptx.table.set_style',
        arguments: { slideIndex: 0, tableId, styleName: 'zebraBlue' },
      }),
    ).resolves.toMatchObject({
      ok: true,
      output: { tableId: expect.any(String) },
      recovery: { fileName: 'Untitled.pptx' },
    })
    await host.api.undo()
    expect(host.adapter.snapshot(4)).toMatchObject({ dirty: true })
  })

  it('undoes a table preset together with its injected table-style package part', async () => {
    const presentation = await BrowserPresentation.blank()
    const added = await presentation.addTableEmu({
      slideIndex: 0,
      rows: 2,
      columns: 2,
      xEmu: 100,
      yEmu: 100,
      widthEmu: 4_000,
      heightEmu: 2_000,
    })
    expect(added).toBeTruthy()
    const before = await presentation.checkpoint()
    const { openPptx } = await import('@genoffice/pptx-engine')
    expect((await openPptx(before)).archive.has('ppt/tableStyles.xml')).toBe(false)

    await presentation.setTableStyle({
      slideIndex: 0,
      tableId: added!.tableId,
      styleName: 'zebraBlue',
    })
    expect(
      (await openPptx(await presentation.checkpoint())).archive.has('ppt/tableStyles.xml'),
    ).toBe(true)
    await presentation.undo()
    expect(
      (await openPptx(await presentation.checkpoint())).archive.has('ppt/tableStyles.xml'),
    ).toBe(false)
  })

  it('adds, updates, and reads back a PPTX chart through registry primitives', async () => {
    const host = createBrowserSlidesHost()
    await host.adapter.execute({
      commandId: 'blank',
      baseRevision: 0,
      operation: 'pptx.document.create_blank',
      arguments: {},
    })
    const added = await host.adapter.execute({
      commandId: 'chart-add',
      baseRevision: 1,
      operation: 'pptx.chart.add',
      arguments: {
        slideIndex: 0,
        kind: 'bar',
        title: 'Sales',
        categories: ['A', 'B'],
        series: [{ name: 'FY26', values: [10, 20] }],
        xEmu: 100,
        yEmu: 100,
        widthEmu: 5_000,
        heightEmu: 3_000,
      },
    })
    expect(added).toMatchObject({ ok: true, output: { chartId: expect.any(String) } })
    const chartId = (added as unknown as { output: { chartId: string } }).output.chartId
    const updated = await host.adapter.execute({
      commandId: 'chart-update',
      baseRevision: 2,
      operation: 'pptx.chart.update',
      arguments: {
        slideIndex: 0,
        chartId,
        title: 'Updated sales',
        colors: ['#4472C4'],
        legendPosition: 'bottom',
        dataLabels: true,
        allowImportedSimplification: false,
      },
    })
    expect(updated).toMatchObject({
      ok: true,
      output: { chartId: expect.any(String) },
      recovery: { fileName: 'Untitled.pptx' },
    })
    const newChartId = (updated as unknown as { output: { chartId: string } }).output.chartId
    await expect(host.api.getChartData(0, newChartId)).resolves.toMatchObject({
      title: 'Updated sales',
      categories: ['A', 'B'],
      series: [{ name: 'FY26', values: [10, 20] }],
    })
    await host.api.undo()
    expect(host.adapter.snapshot(4)).toMatchObject({ dirty: true })
  })

  it('adds SmartArt as one editable shape group through the registry', async () => {
    const host = createBrowserSlidesHost()
    await host.adapter.execute({
      commandId: 'blank',
      baseRevision: 0,
      operation: 'pptx.document.create_blank',
      arguments: {},
    })
    const added = await host.adapter.execute({
      commandId: 'smartart-add',
      baseRevision: 1,
      operation: 'pptx.smartart.add',
      arguments: {
        slideIndex: 0,
        layout: 'process',
        items: ['Plan', 'Build', 'Ship'],
        xEmu: 100,
        yEmu: 100,
        widthEmu: 5_000,
        heightEmu: 3_000,
      },
    })
    expect(added).toMatchObject({
      ok: true,
      output: { objectId: expect.any(String) },
      recovery: { fileName: 'Untitled.pptx' },
    })
    const objectId = (added as unknown as { output: { objectId: string } }).output.objectId
    await expect(host.api.getRenderSlides()).resolves.toMatchObject([
      { nodes: [expect.objectContaining({ sourceId: objectId, type: 'group' })] },
    ])
    await host.api.undo()
    await expect(host.api.getRenderSlides()).resolves.toMatchObject([{ nodes: [] }])
  })

  it('adds image bytes through the registry and keeps the mutation undoable', async () => {
    const host = createBrowserSlidesHost()
    await host.adapter.execute({
      commandId: 'blank',
      baseRevision: 0,
      operation: 'pptx.document.create_blank',
      arguments: {},
    })
    const added = await host.adapter.execute({
      commandId: 'image-add',
      baseRevision: 1,
      operation: 'pptx.image.add_bytes',
      arguments: {
        slideIndex: 0,
        data: 'iVBORw0KGgo=',
        extension: 'png',
        xEmu: 100,
        yEmu: 100,
        widthEmu: 3_000,
        heightEmu: 2_000,
        name: 'Icon',
      },
    })
    expect(added).toMatchObject({
      ok: true,
      output: { objectId: expect.any(String) },
      recovery: { fileName: 'Untitled.pptx' },
    })
    const objectId = (added as unknown as { output: { objectId: string } }).output.objectId
    await expect(host.api.getRenderSlides()).resolves.toMatchObject([
      { nodes: [expect.objectContaining({ sourceId: objectId, type: 'picture' })] },
    ])
    await host.api.undo()
    await expect(host.api.getRenderSlides()).resolves.toMatchObject([{ nodes: [] }])
  })

  it('routes the retained image picker through the same image-bytes mutation', async () => {
    const host = createBrowserSlidesHost()
    await host.adapter.execute({
      commandId: 'blank',
      baseRevision: 0,
      operation: 'pptx.document.create_blank',
      arguments: {},
    })

    let change: (() => void) | null = null
    const input = {
      type: '',
      accept: '',
      hidden: false,
      files: [
        {
          name: 'picked.png',
          arrayBuffer: async () => new Uint8Array([137, 80, 78, 71]).buffer,
        },
      ],
      addEventListener: (_type: string, listener: () => void) => {
        change = listener
      },
      remove: vi.fn(),
      click: () => change?.(),
    }
    vi.stubGlobal('document', {
      createElement: vi.fn(() => input),
      body: { append: vi.fn() },
    })

    try {
      const added = await host.api.insertImage(0, 1200)
      expect(added).toMatchObject({
        slide: { nodes: [expect.objectContaining({ type: 'picture' })] },
        sourceId: expect.any(String),
      })
      expect(input.accept).toBe('.png,.jpg,.jpeg,.gif,.bmp,.webp,.tif,.tiff')
      await host.api.undo()
      await expect(host.api.getRenderSlides()).resolves.toMatchObject([{ nodes: [] }])
    } finally {
      vi.unstubAllGlobals()
    }
  })

  it('replaces picture bytes through the registry and preserves the object frame', async () => {
    const host = createBrowserSlidesHost()
    await host.adapter.execute({
      commandId: 'blank',
      baseRevision: 0,
      operation: 'pptx.document.create_blank',
      arguments: {},
    })
    const added = await host.adapter.execute({
      commandId: 'image-add',
      baseRevision: 1,
      operation: 'pptx.image.add_bytes',
      arguments: {
        slideIndex: 0,
        data: 'iVBORw0KGgo=',
        extension: 'png',
        xEmu: 100,
        yEmu: 200,
        widthEmu: 3_000,
        heightEmu: 2_000,
      },
    })
    const pictureId = (added as unknown as { output: { objectId: string } }).output.objectId
    await expect(
      host.adapter.execute({
        commandId: 'image-replace',
        baseRevision: 2,
        operation: 'pptx.image.replace_bytes',
        arguments: {
          slideIndex: 0,
          pictureId,
          data: 'R0lGODlhAQABAAAAACw=',
          extension: 'gif',
          preserveCrop: false,
        },
      }),
    ).resolves.toMatchObject({ ok: true, output: { updated: true } })
    await expect(host.api.getRenderSlides()).resolves.toMatchObject([
      {
        nodes: [
          expect.objectContaining({
            sourceId: pictureId,
            type: 'picture',
            box: expect.objectContaining({ x: expect.any(Number), y: expect.any(Number) }),
          }),
        ],
      },
    ])
    await host.api.undo()
    await expect(host.api.getRenderSlides()).resolves.toMatchObject([
      { nodes: [expect.objectContaining({ type: 'picture' })] },
    ])
  })

  it('adds and reads embedded media through the registry-backed browser host', async () => {
    const host = createBrowserSlidesHost()
    await host.adapter.execute({
      commandId: 'blank',
      baseRevision: 0,
      operation: 'pptx.document.create_blank',
      arguments: {},
    })
    const added = await host.adapter.execute({
      commandId: 'media-add',
      baseRevision: 1,
      operation: 'pptx.media.add_bytes',
      arguments: {
        slideIndex: 0,
        kind: 'video',
        data: 'AAAAAA==',
        extension: 'mp4',
        xEmu: 100,
        yEmu: 100,
        widthEmu: 5_000,
        heightEmu: 3_000,
        name: 'Clip',
      },
    })
    expect(added).toMatchObject({
      ok: true,
      output: { objectId: expect.any(String) },
      recovery: { fileName: 'Untitled.pptx' },
    })
    const objectId = (added as unknown as { output: { objectId: string } }).output.objectId
    await expect(host.api.getMediaData(0, objectId)).resolves.toEqual({
      kind: 'video',
      dataUrl: 'data:video/mp4;base64,AAAAAA==',
    })
    await host.api.undo()
    await expect(host.api.getRenderSlides()).resolves.toMatchObject([{ nodes: [] }])
  })

  it('adds a 3D model poster through the registry and keeps it undoable', async () => {
    const host = createBrowserSlidesHost()
    await host.adapter.execute({
      commandId: 'blank',
      baseRevision: 0,
      operation: 'pptx.document.create_blank',
      arguments: {},
    })
    const added = await host.adapter.execute({
      commandId: 'model-add',
      baseRevision: 1,
      operation: 'pptx.model3d.add_bytes',
      arguments: {
        slideIndex: 0,
        data: 'Z2xiAA==',
        extension: 'glb',
        xEmu: 100,
        yEmu: 100,
        widthEmu: 5_000,
        heightEmu: 3_000,
        name: 'Model',
      },
    })
    expect(added).toMatchObject({ ok: true, output: { objectId: expect.any(String) } })
    const objectId = (added as unknown as { output: { objectId: string } }).output.objectId
    await expect(host.api.getRenderSlides()).resolves.toMatchObject([
      { nodes: [expect.objectContaining({ sourceId: objectId, type: 'picture' })] },
    ])
    await host.api.undo()
    await expect(host.api.getRenderSlides()).resolves.toMatchObject([{ nodes: [] }])
  })

  it('applies and undoes a built-in theme through the registry', async () => {
    const host = createBrowserSlidesHost()
    await host.adapter.execute({
      commandId: 'blank',
      baseRevision: 0,
      operation: 'pptx.document.create_blank',
      arguments: {},
    })
    await expect(
      host.adapter.execute({
        commandId: 'theme',
        baseRevision: 1,
        operation: 'pptx.theme.apply',
        arguments: { preset: 'midnight' },
      }),
    ).resolves.toMatchObject({
      ok: true,
      output: { changedSlides: 1 },
      recovery: { fileName: 'Untitled.pptx' },
    })
    await expect(host.api.getRenderSlides()).resolves.toMatchObject([
      { background: { kind: 'solid', color: expect.stringMatching(/^#[0-9A-F]{6}$/i) } },
    ])
    await host.api.undo()
    expect(host.adapter.snapshot(3)).toMatchObject({ dirty: false })
  })

  it('adds an ink stroke through the registry and exposes it to the renderer', async () => {
    const host = createBrowserSlidesHost()
    await host.adapter.execute({
      commandId: 'blank',
      baseRevision: 0,
      operation: 'pptx.document.create_blank',
      arguments: {},
    })
    const added = await host.adapter.execute({
      commandId: 'ink',
      baseRevision: 1,
      operation: 'pptx.ink.add',
      arguments: {
        slideIndex: 0,
        data: 'iVBORw0KGgo=',
        payload: '{"points":[[0,0],[1,1]]}',
        xEmu: 100,
        yEmu: 100,
        widthEmu: 3_000,
        heightEmu: 2_000,
      },
    })
    expect(added).toMatchObject({ ok: true, output: { objectId: expect.any(String) } })
    const objectId = (added as unknown as { output: { objectId: string } }).output.objectId
    await expect(host.api.getRenderSlides()).resolves.toMatchObject([
      {
        nodes: [
          expect.objectContaining({
            sourceId: objectId,
            type: 'picture',
            name: expect.stringMatching(/^aislides-ink /),
          }),
        ],
      },
    ])
    await host.api.undo()
    await expect(host.api.getRenderSlides()).resolves.toMatchObject([{ nodes: [] }])
  })

  it('duplicates objects through an explicit registry offset and undoes atomically', async () => {
    const host = createBrowserSlidesHost()
    await host.adapter.execute({
      commandId: 'blank',
      baseRevision: 0,
      operation: 'pptx.document.create_blank',
      arguments: {},
    })
    const original = await host.adapter.execute({
      commandId: 'add',
      baseRevision: 1,
      operation: 'pptx.object.add',
      arguments: {
        slideIndex: 0,
        kind: 'rect',
        xEmu: 100,
        yEmu: 100,
        widthEmu: 3_000,
        heightEmu: 2_000,
      },
    })
    const objectId = (original as unknown as { output: { objectId: string } }).output.objectId
    await expect(
      host.adapter.execute({
        commandId: 'duplicate',
        baseRevision: 2,
        operation: 'pptx.object.duplicate',
        arguments: {
          slideIndex: 0,
          objectIds: [objectId],
          deltaXEmu: 500,
          deltaYEmu: 500,
        },
      }),
    ).resolves.toMatchObject({
      ok: true,
      output: { objectIds: [expect.any(String)] },
    })
    await expect(host.api.getRenderSlides()).resolves.toMatchObject([
      { nodes: [expect.any(Object), expect.any(Object)] },
    ])
    await host.api.undo()
    await expect(host.api.getRenderSlides()).resolves.toMatchObject([
      { nodes: [expect.any(Object)] },
    ])
  })

  it('copies objects across slides explicitly and backs UI copy/paste with the same primitive', async () => {
    const host = createBrowserSlidesHost()
    await host.adapter.execute({
      commandId: 'blank',
      baseRevision: 0,
      operation: 'pptx.document.create_blank',
      arguments: {},
    })
    const original = await host.adapter.execute({
      commandId: 'add',
      baseRevision: 1,
      operation: 'pptx.object.add',
      arguments: {
        slideIndex: 0,
        kind: 'rect',
        xEmu: 100,
        yEmu: 100,
        widthEmu: 3_000,
        heightEmu: 2_000,
      },
    })
    const objectId = (original as unknown as { output: { objectId: string } }).output.objectId
    await host.adapter.execute({
      commandId: 'slide',
      baseRevision: 2,
      operation: 'pptx.slide.add_blank',
      arguments: { afterSlideIndex: 0 },
    })
    await expect(
      host.adapter.execute({
        commandId: 'copy-to',
        baseRevision: 3,
        operation: 'pptx.object.copy_to',
        arguments: {
          sourceSlideIndex: 0,
          objectIds: [objectId],
          destinationSlideIndex: 1,
          deltaXEmu: 500,
          deltaYEmu: 500,
        },
      }),
    ).resolves.toMatchObject({ ok: true, output: { objectIds: [expect.any(String)] } })
    await expect(host.api.getRenderSlides()).resolves.toMatchObject([
      { nodes: [expect.any(Object)] },
      { nodes: [expect.any(Object)] },
    ])
    expect(
      await host.api.copyElements({
        slideIndex: 1,
        sourceIds: [(await host.api.getRenderSlides())![1]!.nodes[0]!.sourceId],
      }),
    ).toBe(1)
    await expect(
      host.api.pasteElements({ slideIndex: 1, fitWidthPx: 1200 }),
    ).resolves.toMatchObject({ sourceIds: [expect.any(String)] })
    await expect(host.api.getRenderSlides()).resolves.toMatchObject([
      expect.any(Object),
      { nodes: [expect.any(Object), expect.any(Object)] },
    ])
  })

  it('copies a slide explicitly and backs UI slide copy/paste with the same primitive', async () => {
    const host = createBrowserSlidesHost()
    await host.adapter.execute({
      commandId: 'blank',
      baseRevision: 0,
      operation: 'pptx.document.create_blank',
      arguments: {},
    })
    await expect(
      host.adapter.execute({
        commandId: 'copy-slide',
        baseRevision: 1,
        operation: 'pptx.slide.copy_to',
        arguments: { sourceSlideIndex: 0, afterSlideIndex: 0, mode: 'theme' },
      }),
    ).resolves.toMatchObject({
      ok: true,
      output: { slideIndex: 1, slideCount: 2 },
    })
    expect(await host.api.copySlide(0)).toBe(true)
    await expect(host.api.clipboardExternal()).resolves.toEqual({ kind: 'slide' })
    await expect(
      host.api.pasteSlide({ afterIndex: 1, fitWidthPx: 1200, mode: 'source' }),
    ).resolves.toMatchObject({ index: 2, slides: [{}, {}, {}] })
    await expect(host.api.repasteSlide({ mode: 'theme', fitWidthPx: 1200 })).resolves.toMatchObject(
      { index: 2, slides: [{}, {}, {}] },
    )
  })

  it('classifies external browser clipboard text without mutating the deck', async () => {
    vi.stubGlobal('navigator', { clipboard: { readText: vi.fn(async () => 'Plain text') } })
    const host = createBrowserSlidesHost()
    await expect(host.api.clipboardExternal()).resolves.toEqual({
      kind: 'text',
      text: 'Plain text',
    })
    vi.unstubAllGlobals()
  })

  it('deletes and restores a master object through the registry and native history', async () => {
    const bytes = await readFile(
      new URL(
        '../../../packages/pptx-engine/tests/fixtures/01_standard_business.pptx',
        import.meta.url,
      ),
    )
    const host = createBrowserSlidesHost()
    await host.adapter.execute({
      commandId: 'load',
      baseRevision: 0,
      operation: 'pptx.document.load_staged',
      arguments: {
        blobId: 'fixture',
        name: 'fixture.pptx',
        size: bytes.byteLength,
        data: bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength),
      },
    })
    const masters = await host.api.masterEnter(1200)
    const target = masters!.items.find((item) => item.slide.nodes.length > 0)!
    const before = target.slide.nodes.length
    const objectId = target.slide.nodes[0]!.sourceId
    await expect(
      host.adapter.execute({
        commandId: 'master-delete',
        baseRevision: 1,
        operation: 'pptx.master.object.delete',
        arguments: { partPath: target.partPath, objectId },
      }),
    ).resolves.toMatchObject({ ok: true, output: { deleted: true } })
    expect((await host.api.masterOpen(target.partPath))?.nodes).toHaveLength(before - 1)
    await host.api.undo()
    expect((await host.api.masterOpen(target.partPath))?.nodes).toHaveLength(before)
  })

  it('creates a blank presentation through the registry-owned document route', async () => {
    const host = createBrowserSlidesHost()

    await expect(
      host.adapter.execute({
        commandId: 'create-blank-pptx',
        baseRevision: 0,
        operation: 'pptx.document.create_blank',
        arguments: {},
      }),
    ).resolves.toEqual({
      ok: true,
      output: { opened: true, fileName: 'Untitled.pptx', slideCount: 1 },
    })
    expect(host.adapter.snapshot(1)).toMatchObject({
      fileName: null,
      dirty: false,
      selection: { activeSlide: 0, slideCount: 1 },
    })
  })

  it('loads staged PPTX bytes through the format registry and mounted presentation', async () => {
    const bytes = await readFile(
      new URL(
        '../../../packages/pptx-engine/tests/fixtures/01_standard_business.pptx',
        import.meta.url,
      ),
    )
    const data = bytes.buffer.slice(
      bytes.byteOffset,
      bytes.byteOffset + bytes.byteLength,
    ) as ArrayBuffer
    const host = createBrowserSlidesHost()

    await expect(
      host.adapter.execute({
        commandId: 'load-pptx',
        baseRevision: 0,
        operation: 'pptx.document.load_staged',
        arguments: {
          blobId: 'pptx-blob',
          name: 'deck.pptx',
          size: data.byteLength,
          data,
        },
      }),
    ).resolves.toEqual({
      ok: true,
      output: { opened: true, fileName: 'deck.pptx' },
    })
    expect(host.adapter.snapshot(1)).toMatchObject({
      revision: 1,
      fileName: 'deck.pptx',
      dirty: false,
      selection: { presentation: 'deck.pptx' },
    })
  })

  it('selects a rendered object through pptx.selection.set without dirtying the deck', async () => {
    const bytes = await readFile(
      new URL(
        '../../../packages/pptx-engine/tests/fixtures/01_standard_business.pptx',
        import.meta.url,
      ),
    )
    const data = bytes.buffer.slice(
      bytes.byteOffset,
      bytes.byteOffset + bytes.byteLength,
    ) as ArrayBuffer
    const host = createBrowserSlidesHost()
    let opened: OpenResult | null = null
    host.api.onOpened((result) => {
      opened = result
    })
    await host.adapter.execute({
      commandId: 'load-pptx',
      baseRevision: 0,
      operation: 'pptx.document.load_staged',
      arguments: { blobId: 'pptx-blob', name: 'deck.pptx', size: data.byteLength, data },
    })
    const objectId = (opened as OpenResult | null)?.slides[0]?.nodes[0]?.sourceId
    expect(objectId).toBeTruthy()
    await expect(
      host.adapter.execute({
        commandId: 'select-shape',
        baseRevision: 1,
        operation: 'pptx.selection.set',
        arguments: { slideIndex: 0, objectIds: [objectId] },
      }),
    ).resolves.toEqual({
      ok: true,
      output: { selected: 1, slideIndex: 0 },
    })
    expect(host.adapter.snapshot(2)).toMatchObject({
      dirty: false,
      selection: { activeSlide: 0, selectedObjects: [{ id: objectId }] },
    })
  })

  it('replaces selected PPTX text through the registry and restores it through native undo', async () => {
    const bytes = await readFile(
      new URL(
        '../../../packages/pptx-engine/tests/fixtures/01_standard_business.pptx',
        import.meta.url,
      ),
    )
    const data = bytes.buffer.slice(
      bytes.byteOffset,
      bytes.byteOffset + bytes.byteLength,
    ) as ArrayBuffer
    const host = createBrowserSlidesHost()
    let opened: OpenResult | null = null
    host.api.onOpened((result) => {
      opened = result
    })
    await host.adapter.execute({
      commandId: 'load-pptx',
      baseRevision: 0,
      operation: 'pptx.document.load_staged',
      arguments: { blobId: 'pptx-blob', name: 'deck.pptx', size: data.byteLength, data },
    })
    const textNode = (opened as OpenResult | null)?.slides[0]?.nodes.find(
      (node) => node.type === 'text' || node.type === 'shape',
    )
    expect(textNode).toBeTruthy()
    await host.adapter.execute({
      commandId: 'select-text',
      baseRevision: 1,
      operation: 'pptx.selection.set',
      arguments: { slideIndex: 0, objectIds: [textNode?.sourceId] },
    })
    const before = (
      host.adapter.snapshot(2).selection as { selectedObjects: Array<{ text?: string }> }
    ).selectedObjects[0]?.text

    const execution = await host.adapter.execute({
      commandId: 'replace-text',
      baseRevision: 2,
      operation: 'pptx.text.replace_selection',
      arguments: { text: 'Registry replacement' },
    })
    expect(execution).toMatchObject({ ok: true, output: { changed: 1 } })
    expect(execution).toHaveProperty('recovery')
    expect(host.adapter.snapshot(3)).toMatchObject({
      dirty: true,
      selection: { selectedObjects: [{ text: 'Registry replacement' }] },
    })

    const restoredSlides = await host.api.undo()
    const restoredTextNode = restoredSlides?.[0]?.nodes.find(
      (node) => node.type === 'text' || node.type === 'shape',
    )
    await host.adapter.execute({
      commandId: 'select-restored-text',
      baseRevision: 4,
      operation: 'pptx.selection.set',
      arguments: { slideIndex: 0, objectIds: [restoredTextNode?.sourceId] },
    })
    expect(host.adapter.snapshot(4)).toMatchObject({
      dirty: false,
      selection: { selectedObjects: [{ text: before }] },
    })
  })

  it('moves selected PPTX objects through the registry and restores geometry through undo', async () => {
    const bytes = await readFile(
      new URL(
        '../../../packages/pptx-engine/tests/fixtures/01_standard_business.pptx',
        import.meta.url,
      ),
    )
    const data = bytes.buffer.slice(
      bytes.byteOffset,
      bytes.byteOffset + bytes.byteLength,
    ) as ArrayBuffer
    const host = createBrowserSlidesHost()
    let opened: OpenResult | null = null
    host.api.onOpened((result) => {
      opened = result
    })
    await host.adapter.execute({
      commandId: 'load-pptx',
      baseRevision: 0,
      operation: 'pptx.document.load_staged',
      arguments: { blobId: 'pptx-blob', name: 'deck.pptx', size: data.byteLength, data },
    })
    const objectId = (opened as OpenResult | null)?.slides[0]?.nodes[0]?.sourceId
    await host.adapter.execute({
      commandId: 'select-object',
      baseRevision: 1,
      operation: 'pptx.selection.set',
      arguments: { slideIndex: 0, objectIds: [objectId] },
    })
    const before = (
      host.adapter.snapshot(2).selection as {
        selectedObjects: Array<{ box: { x: number; y: number } }>
      }
    ).selectedObjects[0]!.box

    const execution = await host.adapter.execute({
      commandId: 'move-object',
      baseRevision: 2,
      operation: 'pptx.object.move_selection',
      arguments: { deltaXEmu: 91_440, deltaYEmu: -45_720 },
    })
    expect(execution).toMatchObject({ ok: true, output: { moved: 1 } })
    expect(execution).toHaveProperty('recovery')
    expect(host.adapter.snapshot(3)).toMatchObject({
      dirty: true,
      selection: {
        selectedObjects: [{ box: { x: before.x + 91_440, y: before.y - 45_720 } }],
      },
    })

    const restoredSlides = await host.api.undo()
    const restoredId = restoredSlides?.[0]?.nodes[0]?.sourceId
    await host.adapter.execute({
      commandId: 'select-restored-object',
      baseRevision: 4,
      operation: 'pptx.selection.set',
      arguments: { slideIndex: 0, objectIds: [restoredId] },
    })
    expect(host.adapter.snapshot(5)).toMatchObject({
      dirty: false,
      selection: { selectedObjects: [{ box: before }] },
    })
  })

  it('adds a blank slide through the same browser presentation history route', async () => {
    const bytes = await readFile(
      new URL(
        '../../../packages/pptx-engine/tests/fixtures/01_standard_business.pptx',
        import.meta.url,
      ),
    )
    const data = bytes.buffer.slice(
      bytes.byteOffset,
      bytes.byteOffset + bytes.byteLength,
    ) as ArrayBuffer
    const host = createBrowserSlidesHost()
    await host.adapter.execute({
      commandId: 'load-pptx',
      baseRevision: 0,
      operation: 'pptx.document.load_staged',
      arguments: { blobId: 'pptx-blob', name: 'deck.pptx', size: data.byteLength, data },
    })
    const beforeCount = (host.adapter.snapshot(1).selection as { slideCount: number }).slideCount

    const execution = await host.adapter.execute({
      commandId: 'add-blank-slide',
      baseRevision: 1,
      operation: 'pptx.slide.add_blank',
      arguments: { afterSlideIndex: 0 },
    })
    expect(execution).toMatchObject({
      ok: true,
      output: { slideIndex: 1, slideCount: beforeCount + 1 },
    })
    expect(execution).toHaveProperty('recovery')
    expect(host.adapter.snapshot(2)).toMatchObject({
      dirty: true,
      selection: { activeSlide: 1, slideCount: beforeCount + 1 },
    })

    await host.api.undo()
    expect(host.adapter.snapshot(3)).toMatchObject({
      dirty: false,
      selection: { activeSlide: 0, slideCount: beforeCount },
    })
  })

  it('duplicates a slide through the registry with one native undo step', async () => {
    const bytes = await readFile(
      new URL(
        '../../../packages/pptx-engine/tests/fixtures/01_standard_business.pptx',
        import.meta.url,
      ),
    )
    const data = bytes.buffer.slice(
      bytes.byteOffset,
      bytes.byteOffset + bytes.byteLength,
    ) as ArrayBuffer
    const host = createBrowserSlidesHost()
    await host.adapter.execute({
      commandId: 'load-pptx',
      baseRevision: 0,
      operation: 'pptx.document.load_staged',
      arguments: { blobId: 'pptx-blob', name: 'deck.pptx', size: data.byteLength, data },
    })
    const beforeCount = (host.adapter.snapshot(1).selection as { slideCount: number }).slideCount

    await expect(
      host.adapter.execute({
        commandId: 'duplicate-slide',
        baseRevision: 1,
        operation: 'pptx.slide.duplicate',
        arguments: { slideIndex: 0, clearText: false },
      }),
    ).resolves.toMatchObject({
      ok: true,
      output: { slideIndex: 1, slideCount: beforeCount + 1 },
      recovery: { fileName: 'deck.pptx' },
    })
    await host.api.undo()
    expect(host.adapter.snapshot(3)).toMatchObject({
      dirty: false,
      selection: { activeSlide: 0, slideCount: beforeCount },
    })
  })

  it('deletes and moves slides through registry-owned native history routes', async () => {
    const bytes = await readFile(
      new URL(
        '../../../packages/pptx-engine/tests/fixtures/01_standard_business.pptx',
        import.meta.url,
      ),
    )
    const data = bytes.buffer.slice(
      bytes.byteOffset,
      bytes.byteOffset + bytes.byteLength,
    ) as ArrayBuffer
    const host = createBrowserSlidesHost()
    await host.adapter.execute({
      commandId: 'load-pptx',
      baseRevision: 0,
      operation: 'pptx.document.load_staged',
      arguments: { blobId: 'pptx-blob', name: 'deck.pptx', size: data.byteLength, data },
    })
    const beforeCount = (host.adapter.snapshot(1).selection as { slideCount: number }).slideCount

    await expect(
      host.adapter.execute({
        commandId: 'delete-slide',
        baseRevision: 1,
        operation: 'pptx.slide.delete',
        arguments: { slideIndex: 1 },
      }),
    ).resolves.toMatchObject({
      ok: true,
      output: { slideIndex: 1, slideCount: beforeCount - 1 },
    })
    await host.api.undo()

    await expect(
      host.adapter.execute({
        commandId: 'move-slide',
        baseRevision: 3,
        operation: 'pptx.slide.move',
        arguments: { fromIndex: 0, toIndex: 1 },
      }),
    ).resolves.toMatchObject({
      ok: true,
      output: { slideIndex: 1, slideCount: beforeCount },
      recovery: { fileName: 'deck.pptx' },
    })
    await host.api.undo()
    expect(host.adapter.snapshot(5)).toMatchObject({
      dirty: false,
      selection: { activeSlide: 0, slideCount: beforeCount },
    })
  })

  it('shares a replayable undo and redo stack between UI and registry operations', async () => {
    const bytes = await readFile(
      new URL(
        '../../../packages/pptx-engine/tests/fixtures/01_standard_business.pptx',
        import.meta.url,
      ),
    )
    const data = bytes.buffer.slice(
      bytes.byteOffset,
      bytes.byteOffset + bytes.byteLength,
    ) as ArrayBuffer
    const host = createBrowserSlidesHost()
    await host.adapter.execute({
      commandId: 'load-pptx',
      baseRevision: 0,
      operation: 'pptx.document.load_staged',
      arguments: { blobId: 'pptx-blob', name: 'deck.pptx', size: data.byteLength, data },
    })
    const beforeCount = (host.adapter.snapshot(1).selection as { slideCount: number }).slideCount
    await host.adapter.execute({
      commandId: 'add-slide',
      baseRevision: 1,
      operation: 'pptx.slide.add_blank',
      arguments: { afterSlideIndex: 0 },
    })

    await expect(
      host.adapter.execute({
        commandId: 'undo-slide',
        baseRevision: 2,
        operation: 'pptx.history.undo',
        arguments: {},
      }),
    ).resolves.toMatchObject({
      ok: true,
      output: { undone: true, slideIndex: 0, slideCount: beforeCount },
    })
    expect(host.adapter.snapshot(3)).toMatchObject({ dirty: false })

    await expect(
      host.adapter.execute({
        commandId: 'redo-slide',
        baseRevision: 3,
        operation: 'pptx.history.redo',
        arguments: {},
      }),
    ).resolves.toMatchObject({
      ok: true,
      output: { redone: true, slideIndex: 1, slideCount: beforeCount + 1 },
      recovery: { fileName: 'deck.pptx' },
    })
    expect(host.adapter.snapshot(4)).toMatchObject({ dirty: true })
  })

  it('deletes an explicit PPTX object through the registry and restores it through undo', async () => {
    const bytes = await readFile(
      new URL(
        '../../../packages/pptx-engine/tests/fixtures/01_standard_business.pptx',
        import.meta.url,
      ),
    )
    const data = bytes.buffer.slice(
      bytes.byteOffset,
      bytes.byteOffset + bytes.byteLength,
    ) as ArrayBuffer
    const host = createBrowserSlidesHost()
    let opened: OpenResult | null = null
    host.api.onOpened((result) => {
      opened = result
    })
    await host.adapter.execute({
      commandId: 'load-pptx',
      baseRevision: 0,
      operation: 'pptx.document.load_staged',
      arguments: { blobId: 'pptx-blob', name: 'deck.pptx', size: data.byteLength, data },
    })
    const objectId = (opened as OpenResult | null)?.slides[0]?.nodes[0]?.sourceId

    await expect(
      host.adapter.execute({
        commandId: 'delete-object',
        baseRevision: 1,
        operation: 'pptx.object.delete',
        arguments: { slideIndex: 0, objectId },
      }),
    ).resolves.toMatchObject({
      ok: true,
      output: { deleted: true, slideIndex: 0, objectId },
      recovery: { fileName: 'deck.pptx' },
    })
    await host.api.undo()
    expect(host.adapter.snapshot(3)).toMatchObject({ dirty: false })
  })

  it('sets explicit EMU object geometry through the registry and restores it through undo', async () => {
    const bytes = await readFile(
      new URL(
        '../../../packages/pptx-engine/tests/fixtures/01_standard_business.pptx',
        import.meta.url,
      ),
    )
    const data = bytes.buffer.slice(
      bytes.byteOffset,
      bytes.byteOffset + bytes.byteLength,
    ) as ArrayBuffer
    const host = createBrowserSlidesHost()
    let opened: OpenResult | null = null
    host.api.onOpened((result) => {
      opened = result
    })
    await host.adapter.execute({
      commandId: 'load-pptx',
      baseRevision: 0,
      operation: 'pptx.document.load_staged',
      arguments: { blobId: 'pptx-blob', name: 'deck.pptx', size: data.byteLength, data },
    })
    const objectId = (opened as OpenResult | null)?.slides[0]?.nodes[0]?.sourceId
    await host.adapter.execute({
      commandId: 'select-object',
      baseRevision: 1,
      operation: 'pptx.selection.set',
      arguments: { slideIndex: 0, objectIds: [objectId] },
    })
    const before = (
      host.adapter.snapshot(2).selection as {
        selectedObjects: Array<{
          box: { x: number; y: number; cx: number; cy: number }
        }>
      }
    ).selectedObjects[0]!.box

    await expect(
      host.adapter.execute({
        commandId: 'transform-object',
        baseRevision: 2,
        operation: 'pptx.object.set_transform',
        arguments: {
          slideIndex: 0,
          objectId,
          xEmu: before.x + 91_440,
          yEmu: before.y + 45_720,
          widthEmu: before.cx,
          heightEmu: before.cy,
          rotationDegrees: 12,
        },
      }),
    ).resolves.toMatchObject({
      ok: true,
      output: { transformed: true, slideIndex: 0, objectId },
      recovery: { fileName: 'deck.pptx' },
    })
    expect(host.adapter.snapshot(3)).toMatchObject({
      dirty: true,
      selection: {
        selectedObjects: [
          { box: { x: before.x + 91_440, y: before.y + 45_720, cx: before.cx, cy: before.cy } },
        ],
      },
    })
    await host.api.undo()
    expect(host.adapter.snapshot(4)).toMatchObject({ dirty: false })
  })

  it('adds and recovery-reopens an explicit PPTX shape through the registry', async () => {
    const bytes = await readFile(
      new URL(
        '../../../packages/pptx-engine/tests/fixtures/01_standard_business.pptx',
        import.meta.url,
      ),
    )
    const data = bytes.buffer.slice(
      bytes.byteOffset,
      bytes.byteOffset + bytes.byteLength,
    ) as ArrayBuffer
    const host = createBrowserSlidesHost()
    let opened: OpenResult | null = null
    host.api.onOpened((result) => {
      opened = result
    })
    await host.adapter.execute({
      commandId: 'load-pptx',
      baseRevision: 0,
      operation: 'pptx.document.load_staged',
      arguments: { blobId: 'pptx-blob', name: 'deck.pptx', size: data.byteLength, data },
    })
    const beforeCount = (opened as unknown as OpenResult).slides[0]!.nodes.length

    const execution = await host.adapter.execute({
      commandId: 'add-shape',
      baseRevision: 1,
      operation: 'pptx.object.add',
      arguments: {
        slideIndex: 0,
        kind: 'rect',
        xEmu: 914_400,
        yEmu: 457_200,
        widthEmu: 1_828_800,
        heightEmu: 914_400,
        fillColor: '#C43E1C',
      },
    })
    expect(execution).toMatchObject({
      ok: true,
      output: { added: true, slideIndex: 0 },
      recovery: { fileName: 'deck.pptx' },
    })
    const recovery = (execution as { recovery: { data: ArrayBuffer } }).recovery
    const reopened = await BrowserPresentation.open('deck.pptx', recovery.data)
    expect(reopened.render(0, 960).nodes).toHaveLength(beforeCount + 1)
    await host.api.undo()
    expect(host.adapter.snapshot(3)).toMatchObject({ dirty: false })
  })

  it('sets an explicit object flip state through the registry and native undo', async () => {
    const host = createBrowserSlidesHost()
    await host.adapter.execute({
      commandId: 'create-blank',
      baseRevision: 0,
      operation: 'pptx.document.create_blank',
      arguments: {},
    })
    const added = await host.adapter.execute({
      commandId: 'add-arrow',
      baseRevision: 1,
      operation: 'pptx.object.add',
      arguments: {
        slideIndex: 0,
        kind: 'rightArrow',
        xEmu: 914_400,
        yEmu: 457_200,
        widthEmu: 1_828_800,
        heightEmu: 914_400,
      },
    })
    const objectId = (added as unknown as { output: { objectId: string } }).output.objectId

    await expect(
      host.adapter.execute({
        commandId: 'flip-arrow',
        baseRevision: 2,
        operation: 'pptx.object.set_flip',
        arguments: { slideIndex: 0, objectIds: [objectId], horizontal: true },
      }),
    ).resolves.toMatchObject({
      ok: true,
      output: { changed: 1 },
      recovery: { fileName: 'Untitled.pptx' },
    })
    await host.api.undo()
    expect(host.adapter.snapshot(4)).toMatchObject({ dirty: true })
  })

  it('sets PPTX object fill and stroke through the registry and shared renderer state', async () => {
    const host = createBrowserSlidesHost()
    await host.adapter.execute({
      commandId: 'create-blank',
      baseRevision: 0,
      operation: 'pptx.document.create_blank',
      arguments: {},
    })
    const added = await host.adapter.execute({
      commandId: 'add-shape',
      baseRevision: 1,
      operation: 'pptx.object.add',
      arguments: {
        slideIndex: 0,
        kind: 'rect',
        xEmu: 914_400,
        yEmu: 457_200,
        widthEmu: 1_828_800,
        heightEmu: 914_400,
      },
    })
    const objectId = (added as unknown as { output: { objectId: string } }).output.objectId

    await expect(
      host.adapter.execute({
        commandId: 'fill-shape',
        baseRevision: 2,
        operation: 'pptx.object.set_fill',
        arguments: {
          slideIndex: 0,
          objectId,
          fill: { kind: 'solid', color: '#336699' },
        },
      }),
    ).resolves.toMatchObject({ ok: true, output: { updated: true } })
    await expect(
      host.adapter.execute({
        commandId: 'stroke-shape',
        baseRevision: 3,
        operation: 'pptx.object.set_stroke',
        arguments: {
          slideIndex: 0,
          objectId,
          stroke: { color: '#112233', widthPt: 2, dash: 'dash' },
        },
      }),
    ).resolves.toMatchObject({
      ok: true,
      output: { updated: true },
      recovery: { fileName: 'Untitled.pptx' },
    })
    await host.api.undo()
    await host.api.undo()
    expect(host.adapter.snapshot(6)).toMatchObject({ dirty: true })
  })

  it('shares the explicit image-fill primitive between MCP and the retained picker', async () => {
    const host = createBrowserSlidesHost()
    await host.adapter.execute({
      commandId: 'create-blank',
      baseRevision: 0,
      operation: 'pptx.document.create_blank',
      arguments: {},
    })
    const added = await host.adapter.execute({
      commandId: 'add-shape',
      baseRevision: 1,
      operation: 'pptx.object.add',
      arguments: {
        slideIndex: 0,
        kind: 'rect',
        xEmu: 914_400,
        yEmu: 457_200,
        widthEmu: 1_828_800,
        heightEmu: 914_400,
      },
    })
    const objectId = (added as unknown as { output: { objectId: string } }).output.objectId

    await expect(
      host.adapter.execute({
        commandId: 'image-fill',
        baseRevision: 2,
        operation: 'pptx.object.set_image_fill',
        arguments: {
          slideIndex: 0,
          objectId,
          data: 'iVBORw0KGgo=',
          extension: 'png',
        },
      }),
    ).resolves.toMatchObject({ ok: true, output: { updated: true } })
    await expect(host.api.getRenderSlides()).resolves.toMatchObject([
      {
        nodes: [expect.objectContaining({ fill: expect.objectContaining({ kind: 'image' }) })],
      },
    ])
    await host.api.undo()

    const restoredId = (await host.api.getRenderSlides())?.[0]?.nodes[0]?.sourceId
    expect(restoredId).toBeTruthy()
    let change: (() => void) | null = null
    const input = {
      type: '',
      accept: '',
      hidden: false,
      files: [
        {
          name: 'fill.png',
          arrayBuffer: async () => new Uint8Array([137, 80, 78, 71]).buffer,
        },
      ],
      addEventListener: (_type: string, listener: () => void) => {
        change = listener
      },
      remove: vi.fn(),
      click: () => change?.(),
    }
    vi.stubGlobal('document', {
      createElement: vi.fn(() => input),
      body: { append: vi.fn() },
    })
    try {
      await expect(
        host.api.editImageFill({ slideIndex: 0, sourceId: restoredId! }),
      ).resolves.toMatchObject({
        nodes: [expect.objectContaining({ fill: expect.objectContaining({ kind: 'image' }) })],
      })
    } finally {
      vi.unstubAllGlobals()
    }
  })

  it('formats explicit PPTX text objects through the registry and native undo', async () => {
    const bytes = await readFile(
      new URL(
        '../../../packages/pptx-engine/tests/fixtures/01_standard_business.pptx',
        import.meta.url,
      ),
    )
    const data = bytes.buffer.slice(
      bytes.byteOffset,
      bytes.byteOffset + bytes.byteLength,
    ) as ArrayBuffer
    const host = createBrowserSlidesHost()
    let opened: OpenResult | null = null
    host.api.onOpened((result) => {
      opened = result
    })
    await host.adapter.execute({
      commandId: 'load-pptx',
      baseRevision: 0,
      operation: 'pptx.document.load_staged',
      arguments: { blobId: 'pptx-blob', name: 'deck.pptx', size: data.byteLength, data },
    })
    const textNode = (opened as OpenResult | null)?.slides[0]?.nodes.find(
      (node) => node.type === 'text' || node.type === 'shape',
    ) as ShapeRenderNode | undefined
    expect(textNode).toBeTruthy()

    await expect(
      host.adapter.execute({
        commandId: 'format-text',
        baseRevision: 1,
        operation: 'pptx.text.set_font',
        arguments: { slideIndex: 0, objectIds: [textNode?.sourceId], bold: true },
      }),
    ).resolves.toMatchObject({
      ok: true,
      output: { changed: 1 },
      recovery: { fileName: 'deck.pptx' },
    })
    const formatted = (opened as OpenResult | null)?.slides[0]?.nodes.find(
      (node) => node.sourceId === textNode?.sourceId,
    ) as ShapeRenderNode | undefined
    expect(
      formatted?.text?.lines.flatMap((line) => line.runs).filter((run) => !run.isBullet),
    ).toEqual(expect.arrayContaining([expect.objectContaining({ bold: true })]))
    await host.api.undo()
    expect(host.adapter.snapshot(3)).toMatchObject({ dirty: false })
  })

  it('formats explicit PPTX paragraphs through the registry and native undo', async () => {
    const bytes = await readFile(
      new URL(
        '../../../packages/pptx-engine/tests/fixtures/01_standard_business.pptx',
        import.meta.url,
      ),
    )
    const data = bytes.buffer.slice(
      bytes.byteOffset,
      bytes.byteOffset + bytes.byteLength,
    ) as ArrayBuffer
    const host = createBrowserSlidesHost()
    let opened: OpenResult | null = null
    host.api.onOpened((result) => {
      opened = result
    })
    await host.adapter.execute({
      commandId: 'load-pptx',
      baseRevision: 0,
      operation: 'pptx.document.load_staged',
      arguments: { blobId: 'pptx-blob', name: 'deck.pptx', size: data.byteLength, data },
    })
    const textNode = (opened as OpenResult | null)?.slides[0]?.nodes.find(
      (node) => node.type === 'text' || node.type === 'shape',
    )

    await expect(
      host.adapter.execute({
        commandId: 'format-paragraph',
        baseRevision: 1,
        operation: 'pptx.paragraph.set_format',
        arguments: { slideIndex: 0, objectIds: [textNode?.sourceId], align: 'center' },
      }),
    ).resolves.toMatchObject({
      ok: true,
      output: { changed: 1 },
      recovery: { fileName: 'deck.pptx' },
    })
    await host.api.undo()
    expect(host.adapter.snapshot(3)).toMatchObject({ dirty: false })
  })

  it('finds and replaces PPTX text through the registry and native undo', async () => {
    const bytes = await readFile(
      new URL(
        '../../../packages/pptx-engine/tests/fixtures/01_standard_business.pptx',
        import.meta.url,
      ),
    )
    const data = bytes.buffer.slice(
      bytes.byteOffset,
      bytes.byteOffset + bytes.byteLength,
    ) as ArrayBuffer
    const host = createBrowserSlidesHost()
    await host.adapter.execute({
      commandId: 'load-pptx',
      baseRevision: 0,
      operation: 'pptx.document.load_staged',
      arguments: { blobId: 'pptx-blob', name: 'deck.pptx', size: data.byteLength, data },
    })

    await expect(
      host.adapter.execute({
        commandId: 'replace-title',
        baseRevision: 1,
        operation: 'pptx.text.replace_all',
        arguments: { find: 'Q3 Business Review', replace: 'Q4 Business Review' },
      }),
    ).resolves.toMatchObject({
      ok: true,
      output: { changed: 1 },
      recovery: { fileName: 'deck.pptx' },
    })
    await host.api.undo()
    expect(host.adapter.snapshot(3)).toMatchObject({ dirty: false })
  })

  it('sets rich PPTX paragraphs through the registry and retained text mapper', async () => {
    const bytes = await readFile(
      new URL(
        '../../../packages/pptx-engine/tests/fixtures/01_standard_business.pptx',
        import.meta.url,
      ),
    )
    const data = bytes.buffer.slice(
      bytes.byteOffset,
      bytes.byteOffset + bytes.byteLength,
    ) as ArrayBuffer
    const host = createBrowserSlidesHost()
    let opened: OpenResult | null = null
    host.api.onOpened((result) => {
      opened = result
    })
    await host.adapter.execute({
      commandId: 'load-pptx',
      baseRevision: 0,
      operation: 'pptx.document.load_staged',
      arguments: { blobId: 'pptx-blob', name: 'deck.pptx', size: data.byteLength, data },
    })
    const textNode = (opened as OpenResult | null)?.slides[0]?.nodes.find(
      (node) => node.type === 'text' || node.type === 'shape',
    )

    await expect(
      host.adapter.execute({
        commandId: 'set-rich-text',
        baseRevision: 1,
        operation: 'pptx.text.set_paragraphs',
        arguments: {
          slideIndex: 0,
          objectId: textNode?.sourceId,
          paragraphs: [
            {
              runs: [
                { text: 'Registry ', bold: true },
                { text: 'rich text', italic: true },
              ],
              align: 'center',
            },
          ],
        },
      }),
    ).resolves.toMatchObject({
      ok: true,
      output: { updated: true, slideIndex: 0, objectId: textNode?.sourceId },
      recovery: { fileName: 'deck.pptx' },
    })
    await host.adapter.execute({
      commandId: 'select-rich-text',
      baseRevision: 2,
      operation: 'pptx.selection.set',
      arguments: { slideIndex: 0, objectIds: [textNode?.sourceId] },
    })
    expect(host.adapter.snapshot(3)).toMatchObject({
      dirty: true,
      selection: { selectedObjects: [{ text: 'Registry rich text' }] },
    })
    await host.api.undo()
    expect(host.adapter.snapshot(4)).toMatchObject({ dirty: false })
  })

  it('sets a PPTX slide background through the registry, persists it, and undoes natively', async () => {
    const host = createBrowserSlidesHost()
    let opened: OpenResult | null = null
    host.api.onOpened((result) => {
      opened = result
    })
    await host.adapter.execute({
      commandId: 'create-blank',
      baseRevision: 0,
      operation: 'pptx.document.create_blank',
      arguments: {},
    })

    const execution = await host.adapter.execute({
      commandId: 'set-background',
      baseRevision: 1,
      operation: 'pptx.slide.set_background',
      arguments: { scope: 'slide', slideIndex: 0, color: '#336699' },
    })
    expect(execution).toMatchObject({
      ok: true,
      output: { changed: 1 },
      recovery: { fileName: 'Untitled.pptx' },
    })
    expect((opened as OpenResult | null)?.slides[0]?.background).toMatchObject({
      kind: 'solid',
      color: '#336699',
    })

    const recovery = (execution as { recovery: { fileName: string; data: ArrayBuffer } }).recovery
    const reopened = await BrowserPresentation.open(recovery.fileName, recovery.data)
    expect(reopened.render(0).background).toMatchObject({ kind: 'solid', color: '#336699' })

    await host.api.undo()
    expect(host.adapter.snapshot(3)).toMatchObject({ dirty: false })
  })

  it('groups, ungroups, reorders, and batch-transforms PPTX objects through registry primitives', async () => {
    const host = createBrowserSlidesHost()
    let opened: OpenResult | null = null
    host.api.onOpened((result) => {
      opened = result
    })
    await host.adapter.execute({
      commandId: 'blank',
      baseRevision: 0,
      operation: 'pptx.document.create_blank',
      arguments: {},
    })
    const add = async (commandId: string, xEmu: number) => {
      const result = await host.adapter.execute({
        commandId,
        baseRevision: 1,
        operation: 'pptx.object.add',
        arguments: {
          slideIndex: 0,
          kind: 'rect',
          xEmu,
          yEmu: 100,
          widthEmu: 1_000,
          heightEmu: 1_000,
        },
      })
      return (result as unknown as { output: { objectId: string } }).output.objectId
    }
    const first = await add('add-first', 100)
    const second = await add('add-second', 2_000)

    await expect(
      host.adapter.execute({
        commandId: 'batch-transform',
        baseRevision: 3,
        operation: 'pptx.object.set_transforms',
        arguments: {
          slideIndex: 0,
          objects: [
            {
              objectId: first,
              xEmu: 500,
              yEmu: 600,
              widthEmu: 1_000,
              heightEmu: 1_000,
              rotationDegrees: 0,
            },
            {
              objectId: second,
              xEmu: 2_500,
              yEmu: 600,
              widthEmu: 1_000,
              heightEmu: 1_000,
              rotationDegrees: 0,
            },
          ],
        },
      }),
    ).resolves.toMatchObject({ ok: true, output: { changed: 2 } })
    await expect(
      host.adapter.execute({
        commandId: 'reorder',
        baseRevision: 4,
        operation: 'pptx.object.reorder',
        arguments: { slideIndex: 0, objectId: first, position: 'front' },
      }),
    ).resolves.toMatchObject({ ok: true, output: { reordered: true } })
    const grouped = await host.adapter.execute({
      commandId: 'group',
      baseRevision: 5,
      operation: 'pptx.object.group',
      arguments: { slideIndex: 0, objectIds: [first, second] },
    })
    expect(grouped).toMatchObject({ ok: true, output: { grouped: true } })
    const groupId = (grouped as unknown as { output: { groupId: string } }).output.groupId
    expect((opened as OpenResult | null)?.slides[0]?.nodes).toEqual([
      expect.objectContaining({ type: 'group', sourceId: groupId }),
    ])
    await expect(
      host.adapter.execute({
        commandId: 'ungroup',
        baseRevision: 6,
        operation: 'pptx.object.ungroup',
        arguments: { slideIndex: 0, groupId },
      }),
    ).resolves.toMatchObject({ ok: true, output: { ungrouped: 2 } })
    expect((opened as OpenResult | null)?.slides[0]?.nodes).toHaveLength(2)
    await host.api.undo()
    expect(host.adapter.snapshot(8)).toMatchObject({ dirty: true })
  })

  it('edits group-child rich text and font through explicit parent addressing', async () => {
    const host = createBrowserSlidesHost()
    await host.adapter.execute({
      commandId: 'blank',
      baseRevision: 0,
      operation: 'pptx.document.create_blank',
      arguments: {},
    })
    const addText = async (commandId: string, xEmu: number, text: string) => {
      const result = await host.adapter.execute({
        commandId,
        baseRevision: 1,
        operation: 'pptx.object.add',
        arguments: {
          slideIndex: 0,
          kind: 'rect',
          xEmu,
          yEmu: 100,
          widthEmu: 2_000,
          heightEmu: 1_000,
          text,
        },
      })
      return (result as unknown as { output: { objectId: string } }).output.objectId
    }
    const first = await addText('first', 100, 'First')
    const second = await addText('second', 2_500, 'Second')
    const grouped = await host.adapter.execute({
      commandId: 'group',
      baseRevision: 3,
      operation: 'pptx.object.group',
      arguments: { slideIndex: 0, objectIds: [first, second] },
    })
    const groupId = (grouped as unknown as { output: { groupId: string } }).output.groupId
    const childId = (
      (await host.api.getRenderSlides())?.[0]?.nodes.find(
        (node) => node.sourceId === groupId && node.type === 'group',
      ) as { children?: Array<{ sourceId: string }> } | undefined
    )?.children?.[0]?.sourceId
    expect(childId).toBeTruthy()

    await expect(
      host.adapter.execute({
        commandId: 'font-child',
        baseRevision: 4,
        operation: 'pptx.text.set_font',
        arguments: { slideIndex: 0, groupId, objectIds: [childId!], bold: true },
      }),
    ).resolves.toMatchObject({ ok: true, output: { changed: 1 } })
    await expect(
      host.adapter.execute({
        commandId: 'text-child',
        baseRevision: 5,
        operation: 'pptx.text.set_paragraphs',
        arguments: {
          slideIndex: 0,
          groupId,
          objectId: childId!,
          paragraphs: [{ runs: [{ text: 'Nested text', bold: true }] }],
        },
      }),
    ).resolves.toMatchObject({
      ok: true,
      output: { updated: true },
      recovery: { fileName: 'Untitled.pptx' },
    })
    await expect(
      host.adapter.execute({
        commandId: 'transform-child',
        baseRevision: 6,
        operation: 'pptx.object.set_transform',
        arguments: {
          slideIndex: 0,
          groupId,
          objectId: childId!,
          xEmu: 250,
          yEmu: 150,
          widthEmu: 1_800,
          heightEmu: 900,
          rotationDegrees: 10,
        },
      }),
    ).resolves.toMatchObject({ ok: true, output: { transformed: true } })
    await host.api.undo()
    await host.api.undo()
    await host.api.undo()
    expect(host.adapter.snapshot(8)).toMatchObject({ dirty: true })
  })

  it('sets PPTX connector endpoints and attachment final states through the registry', async () => {
    const host = createBrowserSlidesHost()
    await host.adapter.execute({
      commandId: 'blank',
      baseRevision: 0,
      operation: 'pptx.document.create_blank',
      arguments: {},
    })
    const add = async (commandId: string, kind: string, xEmu: number) => {
      const result = await host.adapter.execute({
        commandId,
        baseRevision: 1,
        operation: 'pptx.object.add',
        arguments: { slideIndex: 0, kind, xEmu, yEmu: 100, widthEmu: 1_000, heightEmu: 1_000 },
      })
      return (result as unknown as { output: { objectId: string } }).output.objectId
    }
    const target = await add('target', 'rect', 100)
    const connector = await add('connector', 'line', 2_000)
    const execution = await host.adapter.execute({
      commandId: 'set-endpoints',
      baseRevision: 3,
      operation: 'pptx.connector.set_endpoints',
      arguments: {
        slideIndex: 0,
        connectorId: connector,
        start: { xEmu: 1_000, yEmu: 2_000, attachment: { objectId: target, connectionPoint: 3 } },
        end: { xEmu: 4_000, yEmu: 5_000, attachment: null },
      },
    })
    expect(execution).toMatchObject({
      ok: true,
      output: { updated: true },
      recovery: { fileName: 'Untitled.pptx' },
    })
    await host.api.undo()
    expect(host.adapter.snapshot(5)).toMatchObject({ dirty: true })
  })

  it('sets PPTX picture crop/opacity and text vertical anchor through the registry', async () => {
    const bytes = await readFile(
      new URL(
        '../../../packages/pptx-engine/tests/fixtures/01_standard_business.pptx',
        import.meta.url,
      ),
    )
    const data = bytes.buffer.slice(
      bytes.byteOffset,
      bytes.byteOffset + bytes.byteLength,
    ) as ArrayBuffer
    const host = createBrowserSlidesHost()
    let opened: OpenResult | null = null
    host.api.onOpened((result) => {
      opened = result
    })
    await host.adapter.execute({
      commandId: 'load',
      baseRevision: 0,
      operation: 'pptx.document.load_staged',
      arguments: { blobId: 'fixture', name: 'deck.pptx', size: data.byteLength, data },
    })
    const slides = (opened as OpenResult | null)?.slides ?? []
    const pictureSlide = slides.findIndex((slide) =>
      slide.nodes.some((node) => node.type === 'picture'),
    )
    const pictureId = slides[pictureSlide]?.nodes.find((node) => node.type === 'picture')?.sourceId
    const textSlide = slides.findIndex((slide) =>
      slide.nodes.some((node) => node.type === 'text' || node.type === 'shape'),
    )
    const textId = slides[textSlide]?.nodes.find(
      (node) => node.type === 'text' || node.type === 'shape',
    )?.sourceId
    expect(pictureId).toBeTruthy()
    expect(textId).toBeTruthy()

    await expect(
      host.adapter.execute({
        commandId: 'crop',
        baseRevision: 1,
        operation: 'pptx.picture.set_crop',
        arguments: {
          slideIndex: pictureSlide,
          pictureId,
          crop: { left: 0.05, top: 0.05, right: 0.05, bottom: 0.05 },
        },
      }),
    ).resolves.toMatchObject({ ok: true, output: { updated: true } })
    await expect(
      host.adapter.execute({
        commandId: 'opacity',
        baseRevision: 2,
        operation: 'pptx.picture.set_opacity',
        arguments: { slideIndex: pictureSlide, pictureId, opacity: 0.5 },
      }),
    ).resolves.toMatchObject({ ok: true, output: { updated: true } })
    await expect(
      host.adapter.execute({
        commandId: 'anchor',
        baseRevision: 3,
        operation: 'pptx.text.set_vertical_anchor',
        arguments: { slideIndex: textSlide, objectId: textId, anchor: 'middle' },
      }),
    ).resolves.toMatchObject({
      ok: true,
      output: { updated: true },
      recovery: { fileName: 'deck.pptx' },
    })
    await host.api.undo()
    await host.api.undo()
    await host.api.undo()
    expect(host.adapter.snapshot(7)).toMatchObject({ dirty: false })
  })

  it('creates and edits a PPTX table through bounded registry operations', async () => {
    const host = createBrowserSlidesHost()
    await host.adapter.execute({
      commandId: 'blank',
      baseRevision: 0,
      operation: 'pptx.document.create_blank',
      arguments: {},
    })
    const added = await host.adapter.execute({
      commandId: 'add-table',
      baseRevision: 1,
      operation: 'pptx.table.add',
      arguments: {
        slideIndex: 0,
        rows: 2,
        columns: 2,
        xEmu: 100,
        yEmu: 200,
        widthEmu: 4_000,
        heightEmu: 2_000,
      },
    })
    expect(added).toMatchObject({ ok: true, output: { tableId: expect.any(String) } })
    let tableId = (added as unknown as { output: { tableId: string } }).output.tableId
    await expect(
      host.adapter.execute({
        commandId: 'cell-content',
        baseRevision: 2,
        operation: 'pptx.table.set_cell_content',
        arguments: {
          slideIndex: 0,
          tableId,
          row: 0,
          column: 0,
          paragraphs: [{ runs: [{ text: 'Registry table', bold: true }] }],
        },
      }),
    ).resolves.toMatchObject({ ok: true, output: { updated: true } })
    await expect(
      host.adapter.execute({
        commandId: 'cell-anchor',
        baseRevision: 3,
        operation: 'pptx.table.set_cell_anchor',
        arguments: { slideIndex: 0, tableId, row: 0, column: 0, anchor: 'middle' },
      }),
    ).resolves.toMatchObject({ ok: true, output: { updated: true } })
    await expect(
      host.adapter.execute({
        commandId: 'column-width',
        baseRevision: 4,
        operation: 'pptx.table.set_column_width',
        arguments: { slideIndex: 0, tableId, column: 0, widthEmu: 2_500 },
      }),
    ).resolves.toMatchObject({ ok: true, output: { updated: true } })
    await expect(
      host.adapter.execute({
        commandId: 'row-height',
        baseRevision: 5,
        operation: 'pptx.table.set_row_height',
        arguments: { slideIndex: 0, tableId, row: 0, heightEmu: 1_500 },
      }),
    ).resolves.toMatchObject({ ok: true, output: { updated: true } })
    const merged = await host.adapter.execute({
      commandId: 'merge',
      baseRevision: 6,
      operation: 'pptx.table.merge_cells',
      arguments: { slideIndex: 0, tableId, action: 'merge-right', row: 0, column: 0 },
    })
    expect(merged).toMatchObject({ ok: true, output: { tableId: expect.any(String) } })
    tableId = (merged as unknown as { output: { tableId: string } }).output.tableId
    const structured = await host.adapter.execute({
      commandId: 'insert-row',
      baseRevision: 7,
      operation: 'pptx.table.edit_structure',
      arguments: { slideIndex: 0, tableId, action: 'insert-row', index: 1, before: false },
    })
    expect(structured).toMatchObject({ ok: false, error: 'invalid_arguments' })
    await host.adapter.execute({
      commandId: 'split',
      baseRevision: 8,
      operation: 'pptx.table.merge_cells',
      arguments: { slideIndex: 0, tableId, action: 'split', row: 0, column: 0 },
    })
    await host.api.undo()
    expect(host.adapter.snapshot(10)).toMatchObject({ dirty: true })
  })

  it('sets PPTX slide size, layout, transition, visibility, and advance timing through the registry', async () => {
    const host = createBrowserSlidesHost()
    await host.adapter.execute({
      commandId: 'blank',
      baseRevision: 0,
      operation: 'pptx.document.create_blank',
      arguments: {},
    })
    await expect(
      host.adapter.execute({
        commandId: 'size',
        baseRevision: 1,
        operation: 'pptx.slide.set_size',
        arguments: { widthEmu: 9_144_000, heightEmu: 6_858_000 },
      }),
    ).resolves.toMatchObject({ ok: true, output: { changed: 1 } })
    await expect(
      host.adapter.execute({
        commandId: 'layout',
        baseRevision: 2,
        operation: 'pptx.slide.set_layout',
        arguments: { slideIndex: 0, layoutPath: null },
      }),
    ).resolves.toMatchObject({ ok: true, output: { updated: true } })
    await expect(
      host.adapter.execute({
        commandId: 'transition',
        baseRevision: 3,
        operation: 'pptx.slide.set_transition',
        arguments: { scope: 'slide', slideIndex: 0, transition: 'fade' },
      }),
    ).resolves.toMatchObject({ ok: true, output: { changed: 1 } })
    await expect(
      host.adapter.execute({
        commandId: 'hidden',
        baseRevision: 4,
        operation: 'pptx.slide.set_hidden',
        arguments: { slideIndex: 0, hidden: true },
      }),
    ).resolves.toMatchObject({ ok: true, output: { updated: true } })
    await expect(
      host.adapter.execute({
        commandId: 'timing',
        baseRevision: 5,
        operation: 'pptx.slide.set_advance_times',
        arguments: { slides: [{ slideIndex: 0, milliseconds: 1_500 }] },
      }),
    ).resolves.toMatchObject({
      ok: true,
      output: { changed: 1 },
      recovery: { fileName: 'Untitled.pptx' },
    })
    await host.api.undo()
    await host.api.undo()
    await host.api.undo()
    await host.api.undo()
    await host.api.undo()
    expect(host.adapter.snapshot(11)).toMatchObject({ dirty: false })
  })

  it('edits PPTX notes, comments, and sections through the registry and native history', async () => {
    const host = createBrowserSlidesHost()
    await host.adapter.execute({
      commandId: 'blank',
      baseRevision: 0,
      operation: 'pptx.document.create_blank',
      arguments: {},
    })
    await host.adapter.execute({
      commandId: 'second-slide',
      baseRevision: 1,
      operation: 'pptx.slide.add_blank',
      arguments: { afterSlideIndex: 0 },
    })
    await expect(
      host.adapter.execute({
        commandId: 'notes',
        baseRevision: 2,
        operation: 'pptx.notes.set',
        arguments: { slideIndex: 0, text: 'Speaker note' },
      }),
    ).resolves.toMatchObject({ ok: true, output: { updated: true } })
    const comment = await host.adapter.execute({
      commandId: 'comment',
      baseRevision: 3,
      operation: 'pptx.comment.add',
      arguments: { slideIndex: 0, author: 'Codex', text: 'Review this' },
    })
    expect(comment).toMatchObject({
      ok: true,
      output: { authorId: expect.any(Number), index: expect.any(Number) },
    })
    const ref = (comment as unknown as { output: { authorId: number; index: number } }).output
    await expect(
      host.adapter.execute({
        commandId: 'delete-comment',
        baseRevision: 4,
        operation: 'pptx.comment.delete',
        arguments: { slideIndex: 0, authorId: ref.authorId, index: ref.index },
      }),
    ).resolves.toMatchObject({ ok: true, output: { deleted: true } })
    const firstSection = await host.adapter.execute({
      commandId: 'section-a',
      baseRevision: 5,
      operation: 'pptx.section.add',
      arguments: { beforeSlideIndex: 0, name: 'Intro' },
    })
    expect(firstSection).toMatchObject({ ok: true, output: { sectionCount: expect.any(Number) } })
    const sections = await host.api.getSections()
    const sectionId = sections[0]?.id
    expect(sectionId).toBeTruthy()
    await expect(
      host.adapter.execute({
        commandId: 'rename-section',
        baseRevision: 6,
        operation: 'pptx.section.rename',
        arguments: { sectionId, name: 'Opening' },
      }),
    ).resolves.toMatchObject({ ok: true })
    await expect(
      host.adapter.execute({
        commandId: 'section-b',
        baseRevision: 7,
        operation: 'pptx.section.add',
        arguments: { beforeSlideIndex: 1, name: 'Body' },
      }),
    ).resolves.toMatchObject({ ok: true })
    await expect(
      host.adapter.execute({
        commandId: 'move-section',
        baseRevision: 8,
        operation: 'pptx.section.move',
        arguments: { sectionId, direction: 'down' },
      }),
    ).resolves.toMatchObject({ ok: true })
    await expect(
      host.adapter.execute({
        commandId: 'remove-section',
        baseRevision: 9,
        operation: 'pptx.section.remove',
        arguments: { sectionId },
      }),
    ).resolves.toMatchObject({ ok: true, recovery: { fileName: 'Untitled.pptx' } })
    await host.api.undo()
    expect(host.adapter.snapshot(11)).toMatchObject({ dirty: true })
  })
})
