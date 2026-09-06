import { describe, expect, it, vi } from 'vitest'

import {
  registerBundledFallbackFontFaces,
  registerEmbeddedFontFaces,
} from '../src/renderer/doc-fonts'

describe('embedded PPTX fonts', () => {
  it('loads every native style into the browser font set', async () => {
    const added: Array<{ family: string; weight?: string; style?: string }> = []
    class FakeFontFace {
      constructor(
        readonly family: string,
        _source: ArrayBuffer,
        readonly descriptors: { weight?: string; style?: string },
      ) {}
      async load() {
        return this
      }
    }
    const fontSet = {
      add: vi.fn((face: FakeFontFace) => added.push({ family: face.family, ...face.descriptors })),
    }

    await expect(
      registerEmbeddedFontFaces(
        [
          { typeface: 'Deck Sans', style: 'regular', sfnt: new Uint8Array([0, 1, 2, 3]) },
          { typeface: 'Deck Sans', style: 'boldItalic', sfnt: new Uint8Array([4, 5, 6, 7]) },
        ],
        { FontFaceCtor: FakeFontFace, fontSet },
      ),
    ).resolves.toBe(2)
    expect(added).toEqual([
      { family: 'Deck Sans', weight: '400', style: 'normal' },
      { family: 'Deck Sans', weight: '700', style: 'italic' },
    ])
  })

  it('keeps opening a deck when one embedded face is malformed', async () => {
    class RejectingFontFace {
      constructor(
        readonly family: string,
        _source: ArrayBuffer,
        _descriptors: { weight?: string; style?: string },
      ) {}
      async load() {
        if (this.family === 'Bad') throw new Error('bad font')
        return this
      }
    }
    const fontSet = { add: vi.fn() }
    const count = await registerEmbeddedFontFaces(
      [
        { typeface: 'Bad', style: 'regular', sfnt: new Uint8Array([0]) },
        { typeface: 'Good', style: 'bold', sfnt: new Uint8Array([1]) },
      ],
      { FontFaceCtor: RejectingFontFace, fontSet },
    )
    expect(count).toBe(1)
    expect(fontSet.add).toHaveBeenCalledTimes(1)
  })

  it('registers the four bundled Calibri-compatible faces under both browser aliases', async () => {
    class FakeFontFace {
      constructor(
        readonly family: string,
        _source: ArrayBuffer,
        readonly descriptors: { weight?: string; style?: string },
      ) {}
      async load() {
        return this
      }
    }
    const read = vi.fn(async () => new ArrayBuffer(4))
    const fontSet = { add: vi.fn() }

    await expect(
      registerBundledFallbackFontFaces(read, { FontFaceCtor: FakeFontFace, fontSet }),
    ).resolves.toBe(8)
    expect(read).toHaveBeenCalledTimes(4)
    expect(fontSet.add).toHaveBeenCalledTimes(8)
  })
})
