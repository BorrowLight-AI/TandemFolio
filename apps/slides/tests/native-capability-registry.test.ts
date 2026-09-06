import { describe, expect, it, vi } from 'vitest'

import { pptxOperationCatalog } from '../src/renderer/operations/catalog'
import { executePptxOperation } from '../src/renderer/operations/registry'

const baseServices = {
  loadStaged: async () => undefined,
  save: async () => ({ ok: true as const, fileName: 'deck.pptx' }),
}

describe('source-current PPTX native capability registry', () => {
  it('publishes typed routes for the new native mutation families', () => {
    const ids = new Set(pptxOperationCatalog.operations.map((operation) => operation.id))
    expect([...ids]).toEqual(
      expect.arrayContaining([
        'pptx.slide.set_background_gradient',
        'pptx.slide.set_background_image',
        'pptx.slide.reset_background',
        'pptx.slide.set_background_graphics_hidden',
        'pptx.object.set_effects',
        'pptx.object.set_geometry',
        'pptx.text.set_body_properties',
      ]),
    )
  })

  it('dispatches native background variants without bypassing the format service', async () => {
    const setSlideBackgroundGradient = vi.fn(async () => 2)
    const setSlideBackgroundImage = vi.fn(async () => 1)
    const resetSlideBackground = vi.fn(async () => 1)
    const setSlideBackgroundGraphicsHidden = vi.fn(async () => 1)
    const services = {
      ...baseServices,
      setSlideBackgroundGradient,
      setSlideBackgroundImage,
      resetSlideBackground,
      setSlideBackgroundGraphicsHidden,
    }

    await expect(
      executePptxOperation(
        {
          operation: 'pptx.slide.set_background_gradient',
          arguments: {
            scope: 'all',
            from: '#112233',
            to: '#DDEEFF',
            angleDeg: 45,
            radial: false,
          },
        },
        services,
      ),
    ).resolves.toMatchObject({ ok: true, output: { changed: 2 } })
    await expect(
      executePptxOperation(
        {
          operation: 'pptx.slide.set_background_image',
          arguments: {
            scope: 'slide',
            slideIndex: 0,
            data: 'iVBORw0KGgo=',
            extension: 'png',
            mode: 'tile',
          },
        },
        services,
      ),
    ).resolves.toMatchObject({ ok: true, output: { changed: 1 } })
    await expect(
      executePptxOperation(
        { operation: 'pptx.slide.reset_background', arguments: { slideIndex: 0 } },
        services,
      ),
    ).resolves.toMatchObject({ ok: true, output: { changed: 1 } })
    await expect(
      executePptxOperation(
        {
          operation: 'pptx.slide.set_background_graphics_hidden',
          arguments: { slideIndex: 0, hidden: true },
        },
        services,
      ),
    ).resolves.toMatchObject({ ok: true, output: { changed: 1 } })
  })

  it('dispatches effects, geometry, and text-body properties through typed services', async () => {
    const setObjectEffects = vi.fn(async () => true)
    const setObjectGeometry = vi.fn(async () => true)
    const setTextBodyProperties = vi.fn(async () => true)
    const services = {
      ...baseServices,
      setObjectEffects,
      setObjectGeometry,
      setTextBodyProperties,
    }

    await expect(
      executePptxOperation(
        {
          operation: 'pptx.object.set_effects',
          arguments: {
            slideIndex: 0,
            objectId: 'shape-1',
            glow: { color: '#4472C4', radiusEmu: 63500 },
            softEdgeRadiusEmu: 12700,
          },
        },
        services,
      ),
    ).resolves.toMatchObject({ ok: true, output: { updated: true } })
    await expect(
      executePptxOperation(
        {
          operation: 'pptx.object.set_geometry',
          arguments: {
            slideIndex: 0,
            objectId: 'shape-1',
            preset: 'roundRect',
            adjustments: { adj: 25000 },
          },
        },
        services,
      ),
    ).resolves.toMatchObject({ ok: true, output: { updated: true } })
    await expect(
      executePptxOperation(
        {
          operation: 'pptx.text.set_body_properties',
          arguments: {
            slideIndex: 0,
            objectId: 'shape-1',
            wrap: true,
            autofit: 'shrink',
            vertical: 'vert270',
            insetLeftEmu: 91440,
          },
        },
        services,
      ),
    ).resolves.toMatchObject({ ok: true, output: { updated: true } })
  })

  it('retains 3-D chart kinds and RTL flags in existing typed routes', () => {
    const byId = new Map(
      pptxOperationCatalog.operations.map((operation) => [
        operation.id,
        operation as { inputSchema: { properties?: Record<string, { enum?: readonly string[] }> } },
      ]),
    )
    expect(byId.get('pptx.chart.add')?.inputSchema.properties?.kind?.enum).toEqual(
      expect.arrayContaining(['bar3D', 'pie3D']),
    )
    expect(byId.get('pptx.paragraph.set_format')?.inputSchema.properties).toHaveProperty('rtl')
    expect(byId.get('pptx.table.set_style')?.inputSchema.properties).toHaveProperty('rtl')
  })
})
