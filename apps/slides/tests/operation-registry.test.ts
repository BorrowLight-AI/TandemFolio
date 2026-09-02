import { describe, expect, it, vi } from 'vitest'
import { createOperationManifest } from '@tandemfolio/operation-contract'

import { pptxOperationCatalog } from '../src/renderer/operations/catalog'
import { pptxRetainedCommandAudit } from '../src/renderer/operations/baseline'
import { executePptxOperation, pptxOperationHandlerIds } from '../src/renderer/operations/registry'

describe('PPTX operation registry', () => {
  it('maps every retained command family and every registry descriptor without gaps', () => {
    expect(pptxRetainedCommandAudit.map((entry) => entry.disposition)).not.toContain('missing')
    const mapped = new Set(pptxRetainedCommandAudit.flatMap((entry) => entry.operationIds))
    expect([...mapped].sort()).toEqual(
      pptxOperationCatalog.operations.map((operation) => operation.id).sort(),
    )
  })

  it('sets a bounded PPTX animation list through stable object ids', async () => {
    const setAnimations = vi.fn(async () => 1)
    const arguments_ = {
      slideIndex: 0,
      animations: [
        { objectId: 'shape', effect: 'fade', trigger: 'onClick', durationMs: 500, delayMs: 0 },
      ],
    }
    await expect(
      executePptxOperation(
        { operation: 'pptx.animation.set', arguments: arguments_ },
        {
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'deck.pptx' }),
          setAnimations,
        },
      ),
    ).resolves.toMatchObject({ ok: true, output: { animationCount: 1 } })
    expect(setAnimations).toHaveBeenCalledWith(arguments_)
  })

  it('sets an explicit PPTX object hyperlink target', async () => {
    const setHyperlink = vi.fn(async () => 'new-id')
    await expect(
      executePptxOperation(
        {
          operation: 'pptx.hyperlink.set',
          arguments: {
            slideIndex: 0,
            objectId: 'shape',
            target: { kind: 'url', url: 'https://example.com' },
          },
        },
        {
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'deck.pptx' }),
          setHyperlink,
        },
      ),
    ).resolves.toMatchObject({ ok: true, output: { objectId: 'new-id' } })
  })

  it('applies PPTX header/footer final state to all slides', async () => {
    const applyHeaderFooter = vi.fn(async () => 3)
    const arguments_ = {
      footer: 'Confidential',
      slideNumber: true,
      date: '2026-08-30',
      automaticDate: false,
    }
    await expect(
      executePptxOperation(
        { operation: 'pptx.slide.apply_header_footer', arguments: arguments_ },
        {
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'deck.pptx' }),
          applyHeaderFooter,
        },
      ),
    ).resolves.toMatchObject({ ok: true, output: { changed: 3 } })
    expect(applyHeaderFooter).toHaveBeenCalledWith(arguments_)
  })

  it('adds a PPTX slide with an explicit layout path', async () => {
    const addSlideWithLayout = vi.fn(async () => ({ slideIndex: 1, slideCount: 2 }))
    await expect(
      executePptxOperation(
        {
          operation: 'pptx.slide.add_with_layout',
          arguments: { afterSlideIndex: 0, layoutPath: 'ppt/slideLayouts/slideLayout1.xml' },
        },
        {
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'deck.pptx' }),
          addSlideWithLayout,
        },
      ),
    ).resolves.toMatchObject({ ok: true, output: { slideIndex: 1, slideCount: 2 } })
  })

  it('sets a PPTX table style through a bounded final-state contract', async () => {
    const setTableStyle = vi.fn(async () => 'table-new')
    const arguments_ = { slideIndex: 0, tableId: 'table', styleName: 'zebraBlue' }
    await expect(
      executePptxOperation(
        { operation: 'pptx.table.set_style', arguments: arguments_ },
        {
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'deck.pptx' }),
          setTableStyle,
        },
      ),
    ).resolves.toMatchObject({ ok: true, output: { tableId: 'table-new' } })
    expect(setTableStyle).toHaveBeenCalledWith(arguments_)
  })

  it('adds and updates PPTX charts through bounded data contracts', async () => {
    const addChart = vi.fn(async () => 'chart-1')
    const updateChart = vi.fn(async () => 'chart-2')
    const addArguments = {
      slideIndex: 0,
      kind: 'bar',
      categories: ['A', 'B'],
      series: [{ name: 'Sales', values: [1, 2] }],
      xEmu: 1,
      yEmu: 2,
      widthEmu: 3_000,
      heightEmu: 2_000,
    }
    await expect(
      executePptxOperation(
        { operation: 'pptx.chart.add', arguments: addArguments },
        {
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'deck.pptx' }),
          addChart,
        },
      ),
    ).resolves.toMatchObject({ ok: true, output: { chartId: 'chart-1' } })
    const updateArguments = {
      slideIndex: 0,
      chartId: 'chart-1',
      title: 'Updated',
      colors: ['#4472C4'],
      allowImportedSimplification: false,
    }
    await expect(
      executePptxOperation(
        { operation: 'pptx.chart.update', arguments: updateArguments },
        {
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'deck.pptx' }),
          updateChart,
        },
      ),
    ).resolves.toMatchObject({ ok: true, output: { chartId: 'chart-2' } })
  })

  it('adds a bounded SmartArt shape group through the registry', async () => {
    const addSmartArt = vi.fn(async () => 'smartart-1')
    const arguments_ = {
      slideIndex: 0,
      layout: 'process',
      items: ['Plan', 'Build', 'Ship'],
      xEmu: 100,
      yEmu: 200,
      widthEmu: 5_000,
      heightEmu: 3_000,
    }
    await expect(
      executePptxOperation(
        { operation: 'pptx.smartart.add', arguments: arguments_ },
        {
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'deck.pptx' }),
          addSmartArt,
        },
      ),
    ).resolves.toMatchObject({ ok: true, output: { objectId: 'smartart-1' } })
    expect(addSmartArt).toHaveBeenCalledWith(arguments_)
  })

  it('adds bounded image bytes through an explicit registry contract', async () => {
    const addImage = vi.fn(async () => 'picture-1')
    const arguments_ = {
      slideIndex: 0,
      data: 'iVBORw0KGgo=',
      extension: 'png',
      xEmu: 100,
      yEmu: 200,
      widthEmu: 3_000,
      heightEmu: 2_000,
      name: 'Icon',
    }
    await expect(
      executePptxOperation(
        { operation: 'pptx.image.add_bytes', arguments: arguments_ },
        {
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'deck.pptx' }),
          addImage,
        },
      ),
    ).resolves.toMatchObject({ ok: true, output: { objectId: 'picture-1' } })
    expect(addImage).toHaveBeenCalledWith(arguments_)
  })

  it('sets a bounded image fill on an explicit shape through the registry', async () => {
    const setObjectImageFill = vi.fn(async () => true)
    const arguments_ = {
      slideIndex: 0,
      objectId: 'shape-1',
      data: 'iVBORw0KGgo=',
      extension: 'png',
    }
    await expect(
      executePptxOperation(
        { operation: 'pptx.object.set_image_fill', arguments: arguments_ },
        {
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'deck.pptx' }),
          setObjectImageFill,
        },
      ),
    ).resolves.toMatchObject({
      handled: true,
      operationId: 'pptx.object.set_image_fill',
      ok: true,
      output: { updated: true, slideIndex: 0, objectId: 'shape-1' },
    })
    expect(setObjectImageFill).toHaveBeenCalledWith(arguments_)
  })

  it('replaces bounded image bytes with an explicit crop-preservation state', async () => {
    const replaceImage = vi.fn(async () => true)
    const arguments_ = {
      slideIndex: 0,
      pictureId: 'picture-1',
      data: 'iVBORw0KGgo=',
      extension: 'png',
      preserveCrop: true,
    }
    await expect(
      executePptxOperation(
        { operation: 'pptx.image.replace_bytes', arguments: arguments_ },
        {
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'deck.pptx' }),
          replaceImage,
        },
      ),
    ).resolves.toMatchObject({ ok: true, output: { updated: true } })
    expect(replaceImage).toHaveBeenCalledWith(arguments_)
  })

  it('adds bounded media bytes with explicit kind and geometry', async () => {
    const addMedia = vi.fn(async () => 'media-1')
    const arguments_ = {
      slideIndex: 0,
      kind: 'video',
      data: 'AAAAAA==',
      extension: 'mp4',
      xEmu: 100,
      yEmu: 200,
      widthEmu: 5_000,
      heightEmu: 3_000,
      name: 'Clip',
    }
    await expect(
      executePptxOperation(
        { operation: 'pptx.media.add_bytes', arguments: arguments_ },
        {
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'deck.pptx' }),
          addMedia,
        },
      ),
    ).resolves.toMatchObject({ ok: true, output: { objectId: 'media-1' } })
    expect(addMedia).toHaveBeenCalledWith(arguments_)
  })

  it('adds bounded 3D model bytes with explicit poster geometry', async () => {
    const addModel3d = vi.fn(async () => 'model-1')
    const arguments_ = {
      slideIndex: 0,
      data: 'Z2xiAA==',
      extension: 'glb',
      xEmu: 100,
      yEmu: 200,
      widthEmu: 5_000,
      heightEmu: 3_000,
      name: 'Model',
    }
    await expect(
      executePptxOperation(
        { operation: 'pptx.model3d.add_bytes', arguments: arguments_ },
        {
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'deck.pptx' }),
          addModel3d,
        },
      ),
    ).resolves.toMatchObject({ ok: true, output: { objectId: 'model-1' } })
    expect(addModel3d).toHaveBeenCalledWith(arguments_)
  })

  it('applies one bounded built-in presentation theme', async () => {
    const applyTheme = vi.fn(async () => 2)
    await expect(
      executePptxOperation(
        { operation: 'pptx.theme.apply', arguments: { preset: 'midnight' } },
        {
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'deck.pptx' }),
          applyTheme,
        },
      ),
    ).resolves.toMatchObject({ ok: true, output: { changedSlides: 2 } })
    expect(applyTheme).toHaveBeenCalledWith({ preset: 'midnight' })
  })

  it('adds one bounded ink stroke through the registry', async () => {
    const addInk = vi.fn(async () => 'ink-1')
    const arguments_ = {
      slideIndex: 0,
      data: 'iVBORw0KGgo=',
      payload: '{"points":[[0,0],[1,1]]}',
      xEmu: 100,
      yEmu: 200,
      widthEmu: 3_000,
      heightEmu: 2_000,
    }
    await expect(
      executePptxOperation(
        { operation: 'pptx.ink.add', arguments: arguments_ },
        {
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'deck.pptx' }),
          addInk,
        },
      ),
    ).resolves.toMatchObject({ ok: true, output: { objectId: 'ink-1' } })
    expect(addInk).toHaveBeenCalledWith(arguments_)
  })

  it('duplicates bounded PPTX object targets by an explicit EMU offset', async () => {
    const duplicateObjects = vi.fn(async () => ['copy-1', 'copy-2'])
    const arguments_ = {
      slideIndex: 0,
      objectIds: ['shape-1', 'shape-2'],
      deltaXEmu: 1_000,
      deltaYEmu: 2_000,
    }
    await expect(
      executePptxOperation(
        { operation: 'pptx.object.duplicate', arguments: arguments_ },
        {
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'deck.pptx' }),
          duplicateObjects,
        },
      ),
    ).resolves.toMatchObject({ ok: true, output: { objectIds: ['copy-1', 'copy-2'] } })
    expect(duplicateObjects).toHaveBeenCalledWith(arguments_)
  })

  it('copies explicit PPTX objects to a destination slide without hidden clipboard state', async () => {
    const copyObjectsTo = vi.fn(async () => ['copy-1'])
    const arguments_ = {
      sourceSlideIndex: 0,
      objectIds: ['shape-1'],
      destinationSlideIndex: 1,
      deltaXEmu: 1_000,
      deltaYEmu: 2_000,
    }
    await expect(
      executePptxOperation(
        { operation: 'pptx.object.copy_to', arguments: arguments_ },
        {
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'deck.pptx' }),
          copyObjectsTo,
        },
      ),
    ).resolves.toMatchObject({ ok: true, output: { objectIds: ['copy-1'] } })
    expect(copyObjectsTo).toHaveBeenCalledWith(arguments_)
  })

  it('copies an explicit slide without hidden clipboard state', async () => {
    const copySlideTo = vi.fn(async () => ({ slideIndex: 2, slideCount: 3 }))
    const arguments_ = {
      sourceSlideIndex: 0,
      afterSlideIndex: 1,
      mode: 'theme',
    }
    await expect(
      executePptxOperation(
        { operation: 'pptx.slide.copy_to', arguments: arguments_ },
        {
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'deck.pptx' }),
          copySlideTo,
        },
      ),
    ).resolves.toMatchObject({ ok: true, output: { slideIndex: 2, slideCount: 3 } })
    expect(copySlideTo).toHaveBeenCalledWith(arguments_)
  })

  it('deletes an explicit master-part object through the registry', async () => {
    const deleteMasterObject = vi.fn(async () => true)
    const arguments_ = {
      partPath: 'ppt/slideMasters/slideMaster1.xml',
      objectId: 'shape-1',
    }
    await expect(
      executePptxOperation(
        { operation: 'pptx.master.object.delete', arguments: arguments_ },
        {
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'deck.pptx' }),
          deleteMasterObject,
        },
      ),
    ).resolves.toMatchObject({ ok: true, output: { deleted: true } })
    expect(deleteMasterObject).toHaveBeenCalledWith(arguments_)
  })

  it('edits master rich text, transform, fill, and stroke through explicit contracts', async () => {
    const services = {
      loadStaged: async () => undefined,
      save: async () => ({ ok: true as const, fileName: 'deck.pptx' }),
      setMasterText: vi.fn(async () => true),
      setMasterTransform: vi.fn(async () => true),
      setMasterFill: vi.fn(async () => true),
      setMasterStroke: vi.fn(async () => true),
    }
    const target = { partPath: 'ppt/slideLayouts/slideLayout1.xml', objectId: 'shape-1' }
    await expect(
      executePptxOperation(
        {
          operation: 'pptx.master.text.set_paragraphs',
          arguments: { ...target, paragraphs: [{ runs: [{ text: 'Title' }] }] },
        },
        services,
      ),
    ).resolves.toMatchObject({ ok: true, output: { updated: true } })
    await expect(
      executePptxOperation(
        {
          operation: 'pptx.master.object.set_transform',
          arguments: {
            ...target,
            xEmu: 1,
            yEmu: 2,
            widthEmu: 3_000,
            heightEmu: 2_000,
            rotationDegrees: 5,
          },
        },
        services,
      ),
    ).resolves.toMatchObject({ ok: true, output: { updated: true } })
    await expect(
      executePptxOperation(
        {
          operation: 'pptx.master.object.set_fill',
          arguments: { ...target, fill: { kind: 'solid', color: '#4472C4' } },
        },
        services,
      ),
    ).resolves.toMatchObject({ ok: true, output: { updated: true } })
    await expect(
      executePptxOperation(
        {
          operation: 'pptx.master.object.set_stroke',
          arguments: {
            ...target,
            stroke: { color: '#000000', widthEmu: 12_700 },
          },
        },
        services,
      ),
    ).resolves.toMatchObject({ ok: true, output: { updated: true } })
  })

  it('executes pptx.document.create_blank through the mounted presentation factory', async () => {
    const createBlank = vi.fn(async () => ({ fileName: 'Untitled.pptx', slideCount: 1 }))

    await expect(
      executePptxOperation(
        { operation: 'pptx.document.create_blank', arguments: {} },
        {
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'deck.pptx' }),
          createBlank,
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'pptx.document.create_blank',
      ok: true,
      output: { opened: true, fileName: 'Untitled.pptx', slideCount: 1 },
      checkpointRecovery: false,
    })
    expect(createBlank).toHaveBeenCalledOnce()
  })

  it('loads staged PPTX bytes through pptx.document.load_staged', async () => {
    const data = new Uint8Array([80, 75, 3, 4]).buffer as ArrayBuffer
    let loaded: { name: string; data: ArrayBuffer } | null = null

    await expect(
      executePptxOperation(
        {
          operation: 'pptx.document.load_staged',
          arguments: {
            blobId: 'pptx-blob',
            name: 'deck.pptx',
            size: data.byteLength,
            data,
          },
        },
        {
          save: async () => ({ ok: true, fileName: 'deck.pptx' }),
          loadStaged: async (input) => {
            loaded = input
          },
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'pptx.document.load_staged',
      ok: true,
      output: { opened: true, fileName: 'deck.pptx' },
    })
    expect(loaded).toEqual({ name: 'deck.pptx', data })
  })

  it('rejects the retired open_local_file transport alias without loading PPTX bytes', async () => {
    const data = new Uint8Array([80, 75, 3, 4]).buffer as ArrayBuffer
    let loaded = false

    await expect(
      executePptxOperation(
        {
          operation: 'open_local_file',
          arguments: {
            blobId: 'pptx-blob',
            name: 'legacy.PPTX',
            size: data.byteLength,
            data,
          },
        },
        {
          save: async () => ({ ok: true, fileName: 'deck.pptx' }),
          loadStaged: async () => {
            loaded = true
          },
        },
      ),
    ).resolves.toEqual({ handled: false })
    expect(loaded).toBe(false)
  })

  it('rejects staged PPTX input that was not hydrated to an ArrayBuffer', async () => {
    const loadStaged = vi.fn(async () => undefined)

    await expect(
      executePptxOperation(
        {
          operation: 'pptx.document.load_staged',
          arguments: {
            blobId: 'pptx-blob',
            name: 'deck.pptx',
            size: 4,
            data: {},
          },
        },
        {
          save: async () => ({ ok: true, fileName: 'deck.pptx' }),
          loadStaged,
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'pptx.document.load_staged',
      ok: false,
      error: 'invalid_arguments',
      message: '$.data must be a hydrated ArrayBuffer.',
    })
    expect(loadStaged).not.toHaveBeenCalled()
  })

  it('rejects staged PPTX bytes whose length does not match the descriptor', async () => {
    const loadStaged = vi.fn(async () => undefined)
    const data = new Uint8Array([80, 75, 3, 4]).buffer as ArrayBuffer

    await expect(
      executePptxOperation(
        {
          operation: 'pptx.document.load_staged',
          arguments: {
            blobId: 'pptx-blob',
            name: 'deck.pptx',
            size: data.byteLength + 1,
            data,
          },
        },
        {
          save: async () => ({ ok: true, fileName: 'deck.pptx' }),
          loadStaged,
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'pptx.document.load_staged',
      ok: false,
      error: 'invalid_arguments',
      message: 'pptx.document.load_staged requires a valid staged PPTX descriptor.',
    })
    expect(loadStaged).not.toHaveBeenCalled()
  })

  it('rejects staged bytes without a PPTX file identity', async () => {
    const loadStaged = vi.fn(async () => undefined)
    const data = new Uint8Array([80, 75, 3, 4]).buffer as ArrayBuffer

    await expect(
      executePptxOperation(
        {
          operation: 'pptx.document.load_staged',
          arguments: {
            blobId: 'pptx-blob',
            name: 'deck.docx',
            size: data.byteLength,
            data,
          },
        },
        {
          save: async () => ({ ok: true, fileName: 'deck.pptx' }),
          loadStaged,
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'pptx.document.load_staged',
      ok: false,
      error: 'invalid_arguments',
      message: 'pptx.document.load_staged requires a valid staged PPTX descriptor.',
    })
    expect(loadStaged).not.toHaveBeenCalled()
  })

  it('maps PPTX loader failures to execution_failed', async () => {
    const data = new Uint8Array([80, 75, 3, 4]).buffer as ArrayBuffer

    await expect(
      executePptxOperation(
        {
          operation: 'pptx.document.load_staged',
          arguments: {
            blobId: 'pptx-blob',
            name: 'broken.pptx',
            size: data.byteLength,
            data,
          },
        },
        {
          save: async () => ({ ok: true, fileName: 'deck.pptx' }),
          loadStaged: async () => {
            throw new Error('PPTX parse failed.')
          },
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'pptx.document.load_staged',
      ok: false,
      error: 'execution_failed',
      message: 'PPTX parse failed.',
    })
  })

  it('provides a serializable PPTX-owned operation catalog', () => {
    const manifest = createOperationManifest([pptxOperationCatalog])

    expect(JSON.parse(JSON.stringify(manifest))).toEqual(manifest)
    expect(manifest.operations).toEqual(pptxOperationCatalog.operations)
  })

  it('keeps one executable handler for every catalog descriptor', () => {
    expect(pptxOperationHandlerIds).toEqual(
      pptxOperationCatalog.operations.map((descriptor) => descriptor.id).sort(),
    )
  })

  it('executes pptx.document.save through the injected presentation save service', async () => {
    const save = vi.fn(async () => ({ ok: true as const, fileName: 'deck.pptx' }))

    await expect(
      executePptxOperation(
        { operation: 'pptx.document.save', arguments: {} },
        { loadStaged: async () => undefined, save },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'pptx.document.save',
      ok: true,
      output: { saved: true, fileName: 'deck.pptx' },
    })
    expect(save).toHaveBeenCalledOnce()
  })

  it('rejects extra save arguments before calling the persistence service', async () => {
    const save = vi.fn(async () => ({ ok: true as const, fileName: 'deck.pptx' }))

    await expect(
      executePptxOperation(
        { operation: 'pptx.document.save', arguments: { overwrite: true } },
        { loadStaged: async () => undefined, save },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'pptx.document.save',
      ok: false,
      error: 'invalid_arguments',
      message: '$.overwrite is not allowed.',
    })
    expect(save).not.toHaveBeenCalled()
  })

  it('maps save cancellation to execution_failed', async () => {
    await expect(
      executePptxOperation(
        { operation: 'pptx.document.save', arguments: {} },
        {
          loadStaged: async () => undefined,
          save: async () => ({ ok: false, message: 'Save canceled.' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'pptx.document.save',
      ok: false,
      error: 'execution_failed',
      message: 'Save canceled.',
    })
  })

  it('maps persistence failures to execution_failed', async () => {
    await expect(
      executePptxOperation(
        { operation: 'pptx.document.save', arguments: {} },
        {
          loadStaged: async () => undefined,
          save: async () => {
            throw new Error('Presentation write failed.')
          },
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'pptx.document.save',
      ok: false,
      error: 'execution_failed',
      message: 'Presentation write failed.',
    })
  })

  it('executes pptx.document.save_as with an explicit bounded file name', async () => {
    const saveAs = vi.fn(async () => ({ ok: true as const, fileName: 'Q3 Review.pptx' }))

    await expect(
      executePptxOperation(
        { operation: 'pptx.document.save_as', arguments: { fileName: 'Q3 Review.pptx' } },
        {
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'deck.pptx' }),
          saveAs,
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'pptx.document.save_as',
      ok: true,
      output: { saved: true, fileName: 'Q3 Review.pptx' },
    })
    expect(saveAs).toHaveBeenCalledWith({ fileName: 'Q3 Review.pptx' })
  })

  it('rejects a save-as target without a PPTX identity', async () => {
    const saveAs = vi.fn()

    await expect(
      executePptxOperation(
        { operation: 'pptx.document.save_as', arguments: { fileName: 'Q3 Review.pdf' } },
        {
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'deck.pptx' }),
          saveAs,
        },
      ),
    ).resolves.toMatchObject({
      handled: true,
      operationId: 'pptx.document.save_as',
      ok: false,
      error: 'invalid_arguments',
    })
    expect(saveAs).not.toHaveBeenCalled()
  })

  it('executes pptx.selection.set through the injected presentation selection service', async () => {
    const setSelection = vi.fn(() => 2)

    await expect(
      executePptxOperation(
        {
          operation: 'pptx.selection.set',
          arguments: { slideIndex: 1, objectIds: ['shape-1', 'shape-2'] },
        },
        {
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'deck.pptx' }),
          setSelection,
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'pptx.selection.set',
      ok: true,
      output: { selected: 2, slideIndex: 1 },
      checkpointRecovery: false,
    })
    expect(setSelection).toHaveBeenCalledWith({
      slideIndex: 1,
      objectIds: ['shape-1', 'shape-2'],
    })
  })

  it('bounds PPTX selection before calling the presentation', async () => {
    const setSelection = vi.fn(() => 0)

    await expect(
      executePptxOperation(
        {
          operation: 'pptx.selection.set',
          arguments: { slideIndex: -1, objectIds: ['shape-1'] },
        },
        {
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'deck.pptx' }),
          setSelection,
        },
      ),
    ).resolves.toMatchObject({
      handled: true,
      operationId: 'pptx.selection.set',
      ok: false,
      error: 'invalid_arguments',
      message: '$.slideIndex must be greater than or equal to 0.',
    })
    expect(setSelection).not.toHaveBeenCalled()
  })

  it('executes pptx.text.replace_selection through the selected-text service', async () => {
    const replaceSelectedText = vi.fn(async () => 2)

    await expect(
      executePptxOperation(
        { operation: 'pptx.text.replace_selection', arguments: { text: 'Quarterly update' } },
        {
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'deck.pptx' }),
          replaceSelectedText,
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'pptx.text.replace_selection',
      ok: true,
      output: { changed: 2 },
    })
    expect(replaceSelectedText).toHaveBeenCalledWith({ text: 'Quarterly update' })
  })

  it('reports an empty editable selection without creating a false success', async () => {
    await expect(
      executePptxOperation(
        { operation: 'pptx.text.replace_selection', arguments: { text: 'No target' } },
        {
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'deck.pptx' }),
          replaceSelectedText: async () => 0,
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'pptx.text.replace_selection',
      ok: false,
      error: 'invalid_arguments',
      message: 'Select at least one editable text object first.',
    })
  })

  it('executes pptx.object.move_selection through the selected-object service', async () => {
    const moveSelectedObjects = vi.fn(async () => 2)

    await expect(
      executePptxOperation(
        {
          operation: 'pptx.object.move_selection',
          arguments: { deltaXEmu: 91_440, deltaYEmu: -45_720 },
        },
        {
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'deck.pptx' }),
          moveSelectedObjects,
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'pptx.object.move_selection',
      ok: true,
      output: { moved: 2 },
    })
    expect(moveSelectedObjects).toHaveBeenCalledWith({
      deltaXEmu: 91_440,
      deltaYEmu: -45_720,
    })
  })

  it('rejects every retired public PPTX alias before entering presentation execution', async () => {
    const services = {
      loadStaged: async () => undefined,
      save: async () => ({ ok: true as const, fileName: 'deck.pptx' }),
    }

    for (const operation of [
      'save',
      'select_objects',
      'replace_selected_text',
      'move_selected_objects',
    ] as const) {
      await expect(executePptxOperation({ operation, arguments: {} }, services)).resolves.toEqual({
        handled: false,
      })
    }
  })

  it('rejects a no-op object movement before entering history', async () => {
    const moveSelectedObjects = vi.fn(async () => 0)

    await expect(
      executePptxOperation(
        {
          operation: 'pptx.object.move_selection',
          arguments: { deltaXEmu: 0, deltaYEmu: 0 },
        },
        {
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'deck.pptx' }),
          moveSelectedObjects,
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'pptx.object.move_selection',
      ok: false,
      error: 'invalid_arguments',
      message: 'At least one movement delta must be non-zero.',
    })
    expect(moveSelectedObjects).not.toHaveBeenCalled()
  })

  it('executes pptx.object.delete through the shared object action', async () => {
    const deleteObject = vi.fn(async () => true)

    await expect(
      executePptxOperation(
        {
          operation: 'pptx.object.delete',
          arguments: { slideIndex: 1, objectId: 'shape-7' },
        },
        {
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'deck.pptx' }),
          deleteObject,
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'pptx.object.delete',
      ok: true,
      output: { deleted: true, slideIndex: 1, objectId: 'shape-7' },
    })
    expect(deleteObject).toHaveBeenCalledWith({ slideIndex: 1, objectId: 'shape-7' })
  })

  it('reports an unknown PPTX object delete target', async () => {
    await expect(
      executePptxOperation(
        {
          operation: 'pptx.object.delete',
          arguments: { slideIndex: 0, objectId: 'missing' },
        },
        {
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'deck.pptx' }),
          deleteObject: async () => false,
        },
      ),
    ).resolves.toMatchObject({
      handled: true,
      operationId: 'pptx.object.delete',
      ok: false,
      error: 'invalid_arguments',
      message: 'The PPTX object does not exist on the requested slide.',
    })
  })

  it('executes pptx.object.set_transform with document-native EMU geometry', async () => {
    const setObjectTransform = vi.fn(async () => true)
    const transform = {
      slideIndex: 1,
      groupId: 'group-2',
      objectId: 'shape-7',
      xEmu: 914_400,
      yEmu: 457_200,
      widthEmu: 1_828_800,
      heightEmu: 914_400,
      rotationDegrees: 15,
    }

    await expect(
      executePptxOperation(
        { operation: 'pptx.object.set_transform', arguments: transform },
        {
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'deck.pptx' }),
          setObjectTransform,
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'pptx.object.set_transform',
      ok: true,
      output: { transformed: true, slideIndex: 1, objectId: 'shape-7' },
    })
    expect(setObjectTransform).toHaveBeenCalledWith(transform)
  })

  it('bounds PPTX object dimensions before entering history', async () => {
    const setObjectTransform = vi.fn(async () => true)

    await expect(
      executePptxOperation(
        {
          operation: 'pptx.object.set_transform',
          arguments: {
            slideIndex: 0,
            objectId: 'shape-1',
            xEmu: 0,
            yEmu: 0,
            widthEmu: 0,
            heightEmu: 10,
            rotationDegrees: 0,
          },
        },
        {
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'deck.pptx' }),
          setObjectTransform,
        },
      ),
    ).resolves.toMatchObject({
      handled: true,
      operationId: 'pptx.object.set_transform',
      ok: false,
      error: 'invalid_arguments',
      message: '$.widthEmu must be greater than or equal to 1.',
    })
    expect(setObjectTransform).not.toHaveBeenCalled()
  })

  it('executes pptx.object.add through the shared element-insertion primitive', async () => {
    const addObject = vi.fn(async () => 'shape-new')
    const input = {
      slideIndex: 0,
      kind: 'rect',
      xEmu: 914_400,
      yEmu: 457_200,
      widthEmu: 1_828_800,
      heightEmu: 914_400,
      fillColor: '#C43E1C',
    }

    await expect(
      executePptxOperation(
        { operation: 'pptx.object.add', arguments: input },
        {
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'deck.pptx' }),
          addObject,
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'pptx.object.add',
      ok: true,
      output: { added: true, slideIndex: 0, objectId: 'shape-new' },
    })
    expect(addObject).toHaveBeenCalledWith(input)
  })

  it('rejects conflicting plain and rich initial PPTX text', async () => {
    const addObject = vi.fn(async () => 'shape-new')

    await expect(
      executePptxOperation(
        {
          operation: 'pptx.object.add',
          arguments: {
            slideIndex: 0,
            kind: 'textbox',
            xEmu: 0,
            yEmu: 0,
            widthEmu: 914_400,
            heightEmu: 457_200,
            text: 'plain',
            paragraphs: [{ runs: [{ text: 'rich' }] }],
          },
        },
        {
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'deck.pptx' }),
          addObject,
        },
      ),
    ).resolves.toMatchObject({
      handled: true,
      operationId: 'pptx.object.add',
      ok: false,
      error: 'invalid_arguments',
      message: 'Provide either text or paragraphs when adding a PPTX object, not both.',
    })
    expect(addObject).not.toHaveBeenCalled()
  })

  it('executes pptx.object.set_flip as an explicit final state', async () => {
    const setObjectFlip = vi.fn(async () => 2)
    const input = {
      slideIndex: 0,
      objectIds: ['arrow-1', 'arrow-2'],
      horizontal: true,
      vertical: false,
    }

    await expect(
      executePptxOperation(
        { operation: 'pptx.object.set_flip', arguments: input },
        {
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'deck.pptx' }),
          setObjectFlip,
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'pptx.object.set_flip',
      ok: true,
      output: { changed: 2 },
    })
    expect(setObjectFlip).toHaveBeenCalledWith(input)
  })

  it('rejects an empty explicit PPTX flip patch', async () => {
    const setObjectFlip = vi.fn(async () => 0)

    await expect(
      executePptxOperation(
        {
          operation: 'pptx.object.set_flip',
          arguments: { slideIndex: 0, objectIds: ['arrow-1'] },
        },
        {
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'deck.pptx' }),
          setObjectFlip,
        },
      ),
    ).resolves.toMatchObject({
      handled: true,
      operationId: 'pptx.object.set_flip',
      ok: false,
      error: 'invalid_arguments',
      message: 'Provide horizontal, vertical, or both for the final PPTX flip state.',
    })
    expect(setObjectFlip).not.toHaveBeenCalled()
  })

  it('executes PPTX fill and stroke final-state operations', async () => {
    const setObjectFill = vi.fn(async () => true)
    const setObjectStroke = vi.fn(async () => true)
    const services = {
      loadStaged: async () => undefined,
      save: async () => ({ ok: true as const, fileName: 'deck.pptx' }),
      setObjectFill,
      setObjectStroke,
    }

    await expect(
      executePptxOperation(
        {
          operation: 'pptx.object.set_fill',
          arguments: {
            slideIndex: 0,
            objectId: 'shape-1',
            fill: { kind: 'gradient', from: '#112233', to: '#AABBCC', angleDegrees: 45 },
          },
        },
        services,
      ),
    ).resolves.toMatchObject({
      handled: true,
      operationId: 'pptx.object.set_fill',
      ok: true,
      output: { updated: true, slideIndex: 0, objectId: 'shape-1' },
    })
    await expect(
      executePptxOperation(
        {
          operation: 'pptx.object.set_stroke',
          arguments: { slideIndex: 0, objectId: 'shape-1', stroke: null },
        },
        services,
      ),
    ).resolves.toMatchObject({
      handled: true,
      operationId: 'pptx.object.set_stroke',
      ok: true,
      output: { updated: true, slideIndex: 0, objectId: 'shape-1' },
    })
    expect(setObjectFill).toHaveBeenCalledOnce()
    expect(setObjectStroke).toHaveBeenCalledOnce()
  })

  it('rejects contradictory PPTX gradient direction fields', async () => {
    const setObjectFill = vi.fn(async () => true)
    await expect(
      executePptxOperation(
        {
          operation: 'pptx.object.set_fill',
          arguments: {
            slideIndex: 0,
            objectId: 'shape-1',
            fill: {
              kind: 'gradient',
              from: '#112233',
              to: '#AABBCC',
              angleDegrees: 45,
              radial: true,
            },
          },
        },
        {
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'deck.pptx' }),
          setObjectFill,
        },
      ),
    ).resolves.toMatchObject({
      handled: true,
      operationId: 'pptx.object.set_fill',
      ok: false,
      error: 'invalid_arguments',
      message: 'A radial PPTX gradient cannot also specify angleDegrees.',
    })
    expect(setObjectFill).not.toHaveBeenCalled()
  })

  it('executes pptx.text.set_font through the shared element-font action', async () => {
    const setElementFont = vi.fn(async () => 2)
    const input = {
      slideIndex: 0,
      objectIds: ['title', 'subtitle'],
      fontFamily: 'Aptos',
      fontSizePt: 24,
      bold: true,
      color: '#336699',
    }

    await expect(
      executePptxOperation(
        { operation: 'pptx.text.set_font', arguments: input },
        {
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'deck.pptx' }),
          setElementFont,
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'pptx.text.set_font',
      ok: true,
      output: { changed: 2 },
    })
    expect(setElementFont).toHaveBeenCalledWith(input)
  })

  it('addresses group-child text and font mutations with an explicit parent group', async () => {
    const setElementFont = vi.fn(async () => 1)
    const setTextParagraphs = vi.fn(async () => true)
    const services = {
      loadStaged: async () => undefined,
      save: async () => ({ ok: true as const, fileName: 'deck.pptx' }),
      setElementFont,
      setTextParagraphs,
    }
    const target = { slideIndex: 0, groupId: 'group-1' }

    await expect(
      executePptxOperation(
        {
          operation: 'pptx.text.set_font',
          arguments: { ...target, objectIds: ['child-1'], bold: true },
        },
        services,
      ),
    ).resolves.toMatchObject({ ok: true, output: { changed: 1 } })
    await expect(
      executePptxOperation(
        {
          operation: 'pptx.text.set_paragraphs',
          arguments: {
            ...target,
            objectId: 'child-1',
            paragraphs: [{ runs: [{ text: 'Nested text' }] }],
          },
        },
        services,
      ),
    ).resolves.toMatchObject({ ok: true, output: { updated: true } })
    expect(setElementFont).toHaveBeenCalledWith({
      ...target,
      objectIds: ['child-1'],
      bold: true,
    })
    expect(setTextParagraphs).toHaveBeenCalledWith({
      ...target,
      objectId: 'child-1',
      paragraphs: [{ runs: [{ text: 'Nested text' }] }],
    })
  })

  it('rejects an empty PPTX font patch before entering history', async () => {
    const setElementFont = vi.fn(async () => 0)

    await expect(
      executePptxOperation(
        {
          operation: 'pptx.text.set_font',
          arguments: { slideIndex: 0, objectIds: ['title'] },
        },
        {
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'deck.pptx' }),
          setElementFont,
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'pptx.text.set_font',
      ok: false,
      error: 'invalid_arguments',
      message: 'Provide at least one PPTX font property to change.',
    })
    expect(setElementFont).not.toHaveBeenCalled()
  })

  it('executes pptx.paragraph.set_format through the shared paragraph action', async () => {
    const setParagraphFormat = vi.fn(async () => 2)
    const input = {
      slideIndex: 0,
      objectIds: ['title', 'subtitle'],
      align: 'center',
      lineSpacingPct: 120,
      spaceAfterPt: 6,
    }

    await expect(
      executePptxOperation(
        { operation: 'pptx.paragraph.set_format', arguments: input },
        {
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'deck.pptx' }),
          setParagraphFormat,
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'pptx.paragraph.set_format',
      ok: true,
      output: { changed: 2 },
    })
    expect(setParagraphFormat).toHaveBeenCalledWith(input)
  })

  it('rejects an empty PPTX paragraph patch before entering history', async () => {
    const setParagraphFormat = vi.fn(async () => 0)

    await expect(
      executePptxOperation(
        {
          operation: 'pptx.paragraph.set_format',
          arguments: { slideIndex: 0, objectIds: ['title'] },
        },
        {
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'deck.pptx' }),
          setParagraphFormat,
        },
      ),
    ).resolves.toMatchObject({
      handled: true,
      operationId: 'pptx.paragraph.set_format',
      ok: false,
      error: 'invalid_arguments',
      message: 'Provide at least one PPTX paragraph property to change.',
    })
    expect(setParagraphFormat).not.toHaveBeenCalled()
  })

  it('executes pptx.text.replace_all through the shared find/replace action', async () => {
    const replaceAllText = vi.fn(async () => 3)
    const input = { find: 'Q2', replace: 'Q3', matchCase: true, slideIndex: 0 }

    await expect(
      executePptxOperation(
        { operation: 'pptx.text.replace_all', arguments: input },
        {
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'deck.pptx' }),
          replaceAllText,
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'pptx.text.replace_all',
      ok: true,
      output: { changed: 3 },
    })
    expect(replaceAllText).toHaveBeenCalledWith(input)
  })

  it('reports a PPTX find/replace with no matches without creating history', async () => {
    await expect(
      executePptxOperation(
        { operation: 'pptx.text.replace_all', arguments: { find: 'missing', replace: 'new' } },
        {
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'deck.pptx' }),
          replaceAllText: async () => 0,
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'pptx.text.replace_all',
      ok: true,
      output: { changed: 0 },
      checkpointRecovery: false,
    })
  })

  it('executes pptx.text.set_paragraphs through the retained rich-text mapper', async () => {
    const setTextParagraphs = vi.fn(async () => true)
    const input = {
      slideIndex: 0,
      objectId: 'title',
      paragraphs: [
        {
          runs: [
            { text: 'Quarterly ', bold: true },
            { text: 'review', italic: true, color: '#336699' },
          ],
          align: 'center',
          level: 0,
        },
      ],
    }

    await expect(
      executePptxOperation(
        { operation: 'pptx.text.set_paragraphs', arguments: input },
        {
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'deck.pptx' }),
          setTextParagraphs,
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'pptx.text.set_paragraphs',
      ok: true,
      output: { updated: true, slideIndex: 0, objectId: 'title' },
    })
    expect(setTextParagraphs).toHaveBeenCalledWith(input)
  })

  it('rejects contradictory PPTX run-link fields before editing', async () => {
    const setTextParagraphs = vi.fn(async () => true)

    await expect(
      executePptxOperation(
        {
          operation: 'pptx.text.set_paragraphs',
          arguments: {
            slideIndex: 0,
            objectId: 'title',
            paragraphs: [
              {
                runs: [
                  {
                    text: 'link',
                    link: { kind: 'url', url: 'https://example.com', slideIndex: 2 },
                  },
                ],
              },
            ],
          },
        },
        {
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'deck.pptx' }),
          setTextParagraphs,
        },
      ),
    ).resolves.toMatchObject({
      handled: true,
      operationId: 'pptx.text.set_paragraphs',
      ok: false,
      error: 'invalid_arguments',
      message: 'A URL run link must provide only url.',
    })
    expect(setTextParagraphs).not.toHaveBeenCalled()
  })

  it('executes pptx.slide.add_blank through the shared slide action', async () => {
    const addBlankSlide = vi.fn(async () => ({ slideIndex: 3, slideCount: 4 }))

    await expect(
      executePptxOperation(
        { operation: 'pptx.slide.add_blank', arguments: { afterSlideIndex: 2 } },
        {
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'deck.pptx' }),
          addBlankSlide,
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'pptx.slide.add_blank',
      ok: true,
      output: { slideIndex: 3, slideCount: 4 },
      refreshDocument: true,
    })
    expect(addBlankSlide).toHaveBeenCalledWith({ afterSlideIndex: 2 })
  })

  it('bounds the blank-slide insertion index before calling the presentation', async () => {
    const addBlankSlide = vi.fn(async () => ({ slideIndex: 0, slideCount: 1 }))

    await expect(
      executePptxOperation(
        { operation: 'pptx.slide.add_blank', arguments: { afterSlideIndex: -1 } },
        {
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'deck.pptx' }),
          addBlankSlide,
        },
      ),
    ).resolves.toMatchObject({
      handled: true,
      operationId: 'pptx.slide.add_blank',
      ok: false,
      error: 'invalid_arguments',
      message: '$.afterSlideIndex must be greater than or equal to 0.',
    })
    expect(addBlankSlide).not.toHaveBeenCalled()
  })

  it('executes pptx.slide.duplicate through the shared slide action', async () => {
    const duplicateSlide = vi.fn(async () => ({ slideIndex: 2, slideCount: 5 }))

    await expect(
      executePptxOperation(
        {
          operation: 'pptx.slide.duplicate',
          arguments: { slideIndex: 1, clearText: true },
        },
        {
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'deck.pptx' }),
          duplicateSlide,
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'pptx.slide.duplicate',
      ok: true,
      output: { slideIndex: 2, slideCount: 5 },
    })
    expect(duplicateSlide).toHaveBeenCalledWith({ slideIndex: 1, clearText: true })
  })

  it('defaults pptx.slide.duplicate to preserving slide text', async () => {
    const duplicateSlide = vi.fn(async () => ({ slideIndex: 1, slideCount: 2 }))

    await executePptxOperation(
      { operation: 'pptx.slide.duplicate', arguments: { slideIndex: 0 } },
      {
        loadStaged: async () => undefined,
        save: async () => ({ ok: true, fileName: 'deck.pptx' }),
        duplicateSlide,
      },
    )
    expect(duplicateSlide).toHaveBeenCalledWith({ slideIndex: 0, clearText: false })
  })

  it('executes pptx.slide.delete through the shared slide action', async () => {
    const deleteSlide = vi.fn(async () => ({ slideIndex: 1, slideCount: 3 }))

    await expect(
      executePptxOperation(
        { operation: 'pptx.slide.delete', arguments: { slideIndex: 2 } },
        {
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'deck.pptx' }),
          deleteSlide,
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'pptx.slide.delete',
      ok: true,
      output: { slideIndex: 1, slideCount: 3 },
    })
    expect(deleteSlide).toHaveBeenCalledWith({ slideIndex: 2 })
  })

  it('executes pptx.slide.move through the shared slide action', async () => {
    const moveSlide = vi.fn(async () => ({ slideIndex: 3, slideCount: 5 }))

    await expect(
      executePptxOperation(
        { operation: 'pptx.slide.move', arguments: { fromIndex: 1, toIndex: 3 } },
        {
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'deck.pptx' }),
          moveSlide,
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'pptx.slide.move',
      ok: true,
      output: { slideIndex: 3, slideCount: 5 },
    })
    expect(moveSlide).toHaveBeenCalledWith({ fromIndex: 1, toIndex: 3 })
  })

  it('rejects a no-op slide move before entering history', async () => {
    const moveSlide = vi.fn(async () => ({ slideIndex: 1, slideCount: 3 }))

    await expect(
      executePptxOperation(
        { operation: 'pptx.slide.move', arguments: { fromIndex: 1, toIndex: 1 } },
        {
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'deck.pptx' }),
          moveSlide,
        },
      ),
    ).resolves.toMatchObject({
      handled: true,
      operationId: 'pptx.slide.move',
      ok: false,
      error: 'invalid_arguments',
      message: 'fromIndex and toIndex must identify different slides.',
    })
    expect(moveSlide).not.toHaveBeenCalled()
  })

  it.each([
    ['pptx.history.undo', 'undoHistory', 'undone'],
    ['pptx.history.redo', 'redoHistory', 'redone'],
  ] as const)(
    'executes %s through the native presentation history',
    async (operation, service, key) => {
      const history = vi.fn(async () => ({ slideIndex: 2, slideCount: 4 }))

      await expect(
        executePptxOperation(
          { operation, arguments: {} },
          {
            loadStaged: async () => undefined,
            save: async () => ({ ok: true, fileName: 'deck.pptx' }),
            [service]: history,
          },
        ),
      ).resolves.toEqual({
        handled: true,
        operationId: operation,
        ok: true,
        output: { [key]: true, slideIndex: 2, slideCount: 4 },
        refreshDocument: true,
      })
      expect(history).toHaveBeenCalledOnce()
    },
  )

  it('reports an exhausted PPTX history stack explicitly', async () => {
    await expect(
      executePptxOperation(
        { operation: 'pptx.history.undo', arguments: {} },
        {
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'deck.pptx' }),
          undoHistory: async () => null,
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'pptx.history.undo',
      ok: false,
      error: 'invalid_arguments',
      message: 'There is no PPTX history entry to undo.',
    })
  })

  it('sets one PPTX slide background through an explicit target scope', async () => {
    const setSlideBackground = vi.fn(async () => 1)

    await expect(
      executePptxOperation(
        {
          operation: 'pptx.slide.set_background',
          arguments: { scope: 'slide', slideIndex: 2, color: '#336699' },
        },
        {
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'deck.pptx' }),
          setSlideBackground,
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'pptx.slide.set_background',
      ok: true,
      output: { changed: 1 },
      refreshDocument: true,
    })
    expect(setSlideBackground).toHaveBeenCalledWith({
      scope: 'slide',
      slideIndex: 2,
      color: '#336699',
    })
  })

  it('rejects contradictory PPTX background target fields', async () => {
    const setSlideBackground = vi.fn(async () => 1)

    await expect(
      executePptxOperation(
        {
          operation: 'pptx.slide.set_background',
          arguments: { scope: 'all', slideIndex: 2, color: '#336699' },
        },
        {
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'deck.pptx' }),
          setSlideBackground,
        },
      ),
    ).resolves.toMatchObject({
      handled: true,
      operationId: 'pptx.slide.set_background',
      ok: false,
      error: 'invalid_arguments',
    })
    expect(setSlideBackground).not.toHaveBeenCalled()
  })

  it.each([
    [
      'pptx.object.group',
      { slideIndex: 0, objectIds: ['a', 'b'] },
      'groupObjects',
      'group-1',
      { grouped: true, groupId: 'group-1' },
    ],
    [
      'pptx.object.ungroup',
      { slideIndex: 0, groupId: 'group-1' },
      'ungroupObject',
      2,
      { ungrouped: 2 },
    ],
    [
      'pptx.object.reorder',
      { slideIndex: 0, objectId: 'a', position: 'front' },
      'reorderObject',
      true,
      { reordered: true },
    ],
  ] as const)(
    'executes %s through the retained arrangement primitive',
    async (operation, arguments_, service, serviceResult, output) => {
      const action = vi.fn(async () => serviceResult)
      await expect(
        executePptxOperation(
          { operation, arguments: arguments_ },
          {
            loadStaged: async () => undefined,
            save: async () => ({ ok: true, fileName: 'deck.pptx' }),
            [service]: action,
          },
        ),
      ).resolves.toMatchObject({ handled: true, operationId: operation, ok: true, output })
      expect(action).toHaveBeenCalledWith(arguments_)
    },
  )

  it('executes bounded PPTX batch transforms as one operation', async () => {
    const setObjectTransforms = vi.fn(async () => 2)
    const arguments_ = {
      slideIndex: 0,
      objects: [
        { objectId: 'a', xEmu: 1, yEmu: 2, widthEmu: 3, heightEmu: 4, rotationDegrees: 5 },
        { objectId: 'b', xEmu: 6, yEmu: 7, widthEmu: 8, heightEmu: 9, rotationDegrees: 10 },
      ],
    }
    await expect(
      executePptxOperation(
        { operation: 'pptx.object.set_transforms', arguments: arguments_ },
        {
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'deck.pptx' }),
          setObjectTransforms,
        },
      ),
    ).resolves.toMatchObject({ ok: true, output: { changed: 2 } })
    expect(setObjectTransforms).toHaveBeenCalledWith(arguments_)
  })

  it('executes explicit PPTX connector endpoints and attachments', async () => {
    const setConnectorEndpoints = vi.fn(async () => true)
    const arguments_ = {
      slideIndex: 0,
      connectorId: 'connector',
      start: { xEmu: 10, yEmu: 20, attachment: { objectId: 'a', connectionPoint: 3 } },
      end: { xEmu: 30, yEmu: 40, attachment: null },
    }
    await expect(
      executePptxOperation(
        { operation: 'pptx.connector.set_endpoints', arguments: arguments_ },
        {
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'deck.pptx' }),
          setConnectorEndpoints,
        },
      ),
    ).resolves.toMatchObject({ ok: true, output: { updated: true } })
    expect(setConnectorEndpoints).toHaveBeenCalledWith(arguments_)
  })

  it.each([
    [
      'pptx.picture.set_crop',
      {
        slideIndex: 0,
        pictureId: 'picture',
        crop: { left: 0.1, top: 0.2, right: 0.1, bottom: 0.2 },
      },
      'setPictureCrop',
    ],
    [
      'pptx.picture.set_opacity',
      { slideIndex: 0, pictureId: 'picture', opacity: 0.5 },
      'setPictureOpacity',
    ],
    [
      'pptx.text.set_vertical_anchor',
      { slideIndex: 0, objectId: 'text', anchor: 'middle' },
      'setTextVerticalAnchor',
    ],
  ] as const)(
    'executes %s through its retained engine primitive',
    async (operation, arguments_, service) => {
      const action = vi.fn(async () => true)
      await expect(
        executePptxOperation(
          { operation, arguments: arguments_ },
          {
            loadStaged: async () => undefined,
            save: async () => ({ ok: true, fileName: 'deck.pptx' }),
            [service]: action,
          },
        ),
      ).resolves.toMatchObject({ handled: true, operationId: operation, ok: true })
      expect(action).toHaveBeenCalledWith(arguments_)
    },
  )

  it('rejects a PPTX crop whose horizontal edges consume the whole picture', async () => {
    const setPictureCrop = vi.fn(async () => true)
    await expect(
      executePptxOperation(
        {
          operation: 'pptx.picture.set_crop',
          arguments: {
            slideIndex: 0,
            pictureId: 'picture',
            crop: { left: 0.6, top: 0, right: 0.4, bottom: 0 },
          },
        },
        {
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'deck.pptx' }),
          setPictureCrop,
        },
      ),
    ).resolves.toMatchObject({ ok: false, error: 'invalid_arguments' })
    expect(setPictureCrop).not.toHaveBeenCalled()
  })

  it.each([
    [
      'pptx.table.add',
      { slideIndex: 0, rows: 2, columns: 3, xEmu: 1, yEmu: 2, widthEmu: 3_000, heightEmu: 2_000 },
      'addTable',
      'table-1',
      { tableId: 'table-1' },
    ],
    [
      'pptx.table.set_cell_content',
      {
        slideIndex: 0,
        tableId: 'table',
        row: 0,
        column: 0,
        paragraphs: [{ runs: [{ text: 'Hello' }] }],
      },
      'setTableCellContent',
      true,
      { updated: true },
    ],
    [
      'pptx.table.edit_structure',
      { slideIndex: 0, tableId: 'table', action: 'insert-row', index: 1, before: true },
      'editTableStructure',
      'table-2',
      { tableId: 'table-2' },
    ],
    [
      'pptx.table.merge_cells',
      { slideIndex: 0, tableId: 'table', action: 'merge-right', row: 0, column: 0 },
      'mergeTableCells',
      'table-2',
      { tableId: 'table-2' },
    ],
    [
      'pptx.table.set_column_width',
      { slideIndex: 0, tableId: 'table', column: 0, widthEmu: 1_000 },
      'setTableColumnWidth',
      true,
      { updated: true },
    ],
    [
      'pptx.table.set_row_height',
      { slideIndex: 0, tableId: 'table', row: 0, heightEmu: 1_000 },
      'setTableRowHeight',
      true,
      { updated: true },
    ],
    [
      'pptx.table.set_cell_anchor',
      { slideIndex: 0, tableId: 'table', row: 0, column: 0, anchor: 'middle' },
      'setTableCellAnchor',
      true,
      { updated: true },
    ],
  ] as const)(
    'executes %s through the table service',
    async (operation, arguments_, service, result, output) => {
      const action = vi.fn(async () => result)
      await expect(
        executePptxOperation(
          { operation, arguments: arguments_ },
          {
            loadStaged: async () => undefined,
            save: async () => ({ ok: true, fileName: 'deck.pptx' }),
            [service]: action,
          },
        ),
      ).resolves.toMatchObject({ handled: true, operationId: operation, ok: true, output })
      expect(action).toHaveBeenCalledWith(arguments_)
    },
  )

  it.each([
    [
      'pptx.slide.set_layout',
      { slideIndex: 0, layoutPath: null },
      'setSlideLayout',
      true,
      { updated: true },
    ],
    [
      'pptx.slide.set_size',
      { widthEmu: 12_192_000, heightEmu: 6_858_000 },
      'setSlideSize',
      3,
      { changed: 3 },
    ],
    [
      'pptx.slide.set_transition',
      { scope: 'slide', slideIndex: 0, transition: 'fade' },
      'setSlideTransition',
      1,
      { changed: 1 },
    ],
    [
      'pptx.slide.set_hidden',
      { slideIndex: 0, hidden: true },
      'setSlideHidden',
      true,
      { updated: true },
    ],
    [
      'pptx.slide.set_advance_times',
      { slides: [{ slideIndex: 0, milliseconds: 1_500 }] },
      'setSlideAdvanceTimes',
      1,
      { changed: 1 },
    ],
  ] as const)(
    'executes %s through the slide metadata service',
    async (operation, arguments_, service, result, output) => {
      const action = vi.fn(async () => result)
      await expect(
        executePptxOperation(
          { operation, arguments: arguments_ },
          {
            loadStaged: async () => undefined,
            save: async () => ({ ok: true, fileName: 'deck.pptx' }),
            [service]: action,
          },
        ),
      ).resolves.toMatchObject({ handled: true, operationId: operation, ok: true, output })
      expect(action).toHaveBeenCalledWith(arguments_)
    },
  )

  it.each([
    [
      'pptx.notes.set',
      { slideIndex: 0, text: 'Speaker note' },
      'setNotes',
      true,
      { updated: true },
    ],
    [
      'pptx.comment.add',
      { slideIndex: 0, author: 'Codex', text: 'Review this' },
      'addComment',
      { authorId: 1, index: 2 },
      { authorId: 1, index: 2 },
    ],
    [
      'pptx.comment.delete',
      { slideIndex: 0, authorId: 1, index: 2 },
      'deleteComment',
      true,
      { deleted: true },
    ],
    [
      'pptx.section.add',
      { beforeSlideIndex: 0, name: 'Intro' },
      'addSection',
      1,
      { sectionCount: 1 },
    ],
    [
      'pptx.section.rename',
      { sectionId: 's1', name: 'Opening' },
      'renameSection',
      1,
      { sectionCount: 1 },
    ],
    ['pptx.section.remove', { sectionId: 's1' }, 'removeSection', 0, { sectionCount: 0 }],
    [
      'pptx.section.move',
      { sectionId: 's1', direction: 'down' },
      'moveSection',
      2,
      { sectionCount: 2 },
    ],
  ] as const)(
    'executes %s through its metadata primitive',
    async (operation, arguments_, service, result, output) => {
      const action = vi.fn(async () => result)
      await expect(
        executePptxOperation(
          { operation, arguments: arguments_ },
          {
            loadStaged: async () => undefined,
            save: async () => ({ ok: true, fileName: 'deck.pptx' }),
            [service]: action,
          },
        ),
      ).resolves.toMatchObject({ handled: true, operationId: operation, ok: true, output })
      expect(action).toHaveBeenCalledWith(arguments_)
    },
  )
})
