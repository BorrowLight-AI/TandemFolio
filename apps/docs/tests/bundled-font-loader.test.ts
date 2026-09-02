import { describe, expect, it, vi } from 'vitest'
import {
  loadBundledFallbackFonts,
  requiredBundledFontFamilies,
  resetBundledFontLoaderForTests,
} from '../src/renderer/bundled-font-loader'

describe('bundled font loader', () => {
  it('selects only fallback families required by the document scripts', () => {
    expect(requiredBundledFontFamilies(['Calibri'], 'English only')).toEqual(['Carlito GO'])
    expect(requiredBundledFontFamilies(['Calibri'], 'English 中文')).toEqual([
      'Carlito GO',
      'Noto Sans CJK SC',
    ])
    expect(requiredBundledFontFamilies(['Traditional Arabic'], 'العربية')).toEqual([
      'Noto Naskh Arabic TA',
    ])
  })

  it('loads each asset once and registers every face for a required family', async () => {
    resetBundledFontLoaderForTests()
    const read = vi.fn(async () => new Uint8Array([0, 1, 0, 0]).buffer)
    const register = vi.fn(async () => undefined)

    await loadBundledFallbackFonts(['Calibri'], 'English only', read, register)
    await loadBundledFallbackFonts(['Calibri'], 'English only', read, register)

    expect(read).toHaveBeenCalledTimes(4)
    expect(register).toHaveBeenCalledTimes(4)
    expect(register.mock.calls.map(([face]) => [face.family, face.weight, face.style])).toEqual([
      ['Carlito GO', '400', 'normal'],
      ['Carlito GO', '700', 'normal'],
      ['Carlito GO', '400', 'italic'],
      ['Carlito GO', '700', 'italic'],
    ])
  })
})
