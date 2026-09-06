import { describe, expect, it } from 'vitest'

import { BrowserPresentation } from '../src/renderer/host/browser-presentation'
import { createBrowserSlidesHost } from '../src/renderer/host/browser-slides-api'

describe('source-current native PPTX mutations', () => {
  it('persists geometry, effects, and text-body properties and undoes each native step', async () => {
    const host = createBrowserSlidesHost()
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
        text: 'Native text',
        xEmu: 100_000,
        yEmu: 100_000,
        widthEmu: 2_000_000,
        heightEmu: 1_000_000,
      },
    })
    const objectId = (added as unknown as { output: { objectId: string } }).output.objectId

    await expect(
      host.adapter.execute({
        commandId: 'geometry',
        baseRevision: 2,
        operation: 'pptx.object.set_geometry',
        arguments: { slideIndex: 0, objectId, preset: 'roundRect', adjustments: { adj: 25_000 } },
      }),
    ).resolves.toMatchObject({ ok: true, output: { updated: true } })
    await expect(
      host.adapter.execute({
        commandId: 'effects',
        baseRevision: 3,
        operation: 'pptx.object.set_effects',
        arguments: {
          slideIndex: 0,
          objectId,
          glow: { color: '#4472C4', radiusEmu: 63_500 },
          softEdgeRadiusEmu: 12_700,
        },
      }),
    ).resolves.toMatchObject({ ok: true, output: { updated: true } })
    const final = await host.adapter.execute({
      commandId: 'text-body',
      baseRevision: 4,
      operation: 'pptx.text.set_body_properties',
      arguments: {
        slideIndex: 0,
        objectId,
        vertical: 'vert270',
        autofit: 'shrink',
        wrap: false,
        insetLeftEmu: 91_440,
      },
    })
    expect(final).toMatchObject({ ok: true, output: { updated: true } })

    const recovery = (final as { recovery: { fileName: string; data: ArrayBuffer } }).recovery
    const reopened = await BrowserPresentation.open(recovery.fileName, recovery.data)
    const node = reopened
      .render(0)
      .nodes.find(
        (candidate) => 'presetGeometry' in candidate && candidate.presetGeometry === 'roundRect',
      )
    expect(node).toMatchObject({
      presetGeometry: 'roundRect',
      glow: expect.objectContaining({ color: '#4472C4' }),
      softEdgePx: expect.any(Number),
      text: expect.objectContaining({ vert: 'vert270', autofit: 'shrink', wrap: false }),
    })

    await host.api.undo()
    const afterUndo = (await host.api.getRenderSlides())?.[0]?.nodes.find(
      (candidate) => 'presetGeometry' in candidate && candidate.presetGeometry === 'roundRect',
    )
    expect(afterUndo).toMatchObject({ presetGeometry: 'roundRect' })
    expect(afterUndo && 'text' in afterUndo ? afterUndo.text?.vert : undefined).toBeUndefined()
  })

  it('persists gradient and image-independent background controls', async () => {
    const host = createBrowserSlidesHost()
    await host.adapter.execute({
      commandId: 'blank',
      baseRevision: 0,
      operation: 'pptx.document.create_blank',
      arguments: {},
    })
    await expect(
      host.adapter.execute({
        commandId: 'gradient',
        baseRevision: 1,
        operation: 'pptx.slide.set_background_gradient',
        arguments: {
          scope: 'slide',
          slideIndex: 0,
          from: '#112233',
          to: '#DDEEFF',
          angleDeg: 45,
        },
      }),
    ).resolves.toMatchObject({ ok: true, output: { changed: 1 } })
    await expect(
      host.adapter.execute({
        commandId: 'hide-graphics',
        baseRevision: 2,
        operation: 'pptx.slide.set_background_graphics_hidden',
        arguments: { slideIndex: 0, hidden: true },
      }),
    ).resolves.toMatchObject({ ok: true, output: { changed: 1 } })
    expect((await host.api.getRenderSlides())?.[0]?.background).toMatchObject({ kind: 'gradient' })
  })

  it('persists a native tiled picture fill through the typed operation', async () => {
    const host = createBrowserSlidesHost()
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
        xEmu: 100_000,
        yEmu: 100_000,
        widthEmu: 1_000_000,
        heightEmu: 1_000_000,
      },
    })
    const objectId = (added as unknown as { output: { objectId: string } }).output.objectId
    const result = await host.adapter.execute({
      commandId: 'tile',
      baseRevision: 2,
      operation: 'pptx.object.set_image_fill',
      arguments: {
        slideIndex: 0,
        objectId,
        mode: 'tile',
        extension: 'png',
        data: 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAAEElEQVR42mP4z8DwHwAFgAIACu8B9QAAAABJRU5ErkJggg==',
      },
    })

    expect(result).toMatchObject({ ok: true, output: { updated: true } })
    expect((await host.api.getRenderSlides())?.[0]?.nodes).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          sourceId: objectId,
          fill: expect.objectContaining({ kind: 'image', mode: 'tile' }),
        }),
      ]),
    )
    const recovery = (result as { recovery: { fileName: string; data: ArrayBuffer } }).recovery
    const reopened = await BrowserPresentation.open(recovery.fileName, recovery.data)
    expect(reopened.render(0).nodes).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          fill: expect.objectContaining({ kind: 'image', mode: 'tile' }),
        }),
      ]),
    )
  })
})
