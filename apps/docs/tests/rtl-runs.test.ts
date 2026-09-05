// Modified by TandemFolio contributors: adapt upstream RTL editor round-trip coverage (2026-09-05).
import { describe, expect, it } from 'vitest'
import type { Run } from '@genoffice/docx-engine'
import { inlineToRuns, runsToInline } from '../src/renderer/editor/convert'

describe('run-level RTL editor conversion', () => {
  it('keeps explicit direction and complex-script selection through the mounted editor model', () => {
    const source: Run[] = [
      {
        text: 'مرحبا',
        rtl: true,
        cs: true,
        csFont: 'Traditional Arabic',
        rawRPr: '<w:rPr><w:rFonts w:cs="Traditional Arabic"/><w:rtl/></w:rPr>',
      },
    ]

    expect(inlineToRuns(runsToInline(source))[0]).toMatchObject({
      text: 'مرحبا',
      rtl: true,
      cs: true,
      csFont: 'Traditional Arabic',
    })
  })
})
