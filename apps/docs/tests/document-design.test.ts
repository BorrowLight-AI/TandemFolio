import { parseDocx, readPageColor, saveDocx } from '@genoffice/docx-engine'
import { describe, expect, it } from 'vitest'

import { applyDocxDocumentDesign } from '../src/renderer/editor/document-design'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'

describe('explicit DOCX document design state', () => {
  it('applies a bounded masked update and saves/reopens every retained design field', async () => {
    const result = applyDocxDocumentDesign(
      { pageColor: null, watermark: null, themeFonts: null, themeColors: null },
      {
        fields: ['pageColor', 'watermark', 'themeFonts', 'themeColors'],
        pageColor: 'FFF9E6',
        watermark: 'DRAFT',
        themeFonts: { major: 'Trebuchet MS', minor: 'Trebuchet MS', eastAsia: '微软雅黑' },
        themeColors: {
          name: 'Facet',
          dk2: '3E3D2D',
          lt2: 'E1DFDD',
          accent1: '90C226',
          accent2: '54A021',
          accent3: 'E6B91E',
          accent4: 'E76618',
          accent5: 'C42F1A',
          accent6: '918655',
        },
      },
    )
    expect(result).toMatchObject({
      ok: true,
      fields: ['pageColor', 'watermark', 'themeFonts', 'themeColors'],
      changedFields: ['pageColor', 'watermark', 'themeFonts', 'themeColors'],
      changed: true,
    })
    if (!result.ok) throw new Error(result.message)

    const parsed = await parseDocx(
      await buildDocx({ bodyXml: '<w:p><w:r><w:t>Design</w:t></w:r></w:p>' }),
    )
    const blocks = parsed.blocks
      .filter((block) => !block.hidden && block.docxIndex !== null)
      .map((block) => ({ kind: 'original' as const, docxIndex: block.docxIndex! }))
    const reopened = await parseDocx(
      await saveDocx(parsed, blocks, {
        pageColor: result.state.pageColor,
        watermark: result.state.watermark,
        themeFonts: result.state.themeFonts!,
        themeColors: result.state.themeColors!,
      }),
    )
    expect(readPageColor(reopened)).toBe('FFF9E6')
    expect(reopened.watermarkText).toBe('DRAFT')
    expect(reopened.themeFonts).toMatchObject(result.state.themeFonts!)
    expect(reopened.themeColors).toMatchObject(result.state.themeColors!)
  })

  it('rejects mismatched masks and reports repeated final state as a no-op', () => {
    const state = { pageColor: null, watermark: 'DRAFT', themeFonts: null, themeColors: null }
    expect(
      applyDocxDocumentDesign(state, {
        fields: ['pageColor'],
        pageColor: 'fff9e6',
      }),
    ).toMatchObject({ ok: false, error: 'invalid_arguments' })
    expect(
      applyDocxDocumentDesign(state, {
        fields: ['watermark'],
        watermark: 'DRAFT',
      }),
    ).toMatchObject({ ok: true, changedFields: [], changed: false, state })
  })
})
