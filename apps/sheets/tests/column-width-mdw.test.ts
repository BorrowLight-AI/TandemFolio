import { afterEach, describe, expect, it } from 'vitest'

import {
  getWorkbookMdw,
  pixelsToCharacterWidth,
  setWorkbookMdw,
} from '../src/renderer/app-constants'
import { generalCharBudget } from '../src/renderer/numfmt-fix'
import {
  characterWidthToPixels,
  measureNormalFontMdw,
  paddedBaseColumnWidth,
} from '../src/renderer/univer-sync'

afterEach(() => setWorkbookMdw(7))

describe('workbook max digit width', () => {
  it('uses one workbook-specific unit for both column-width directions', () => {
    expect(getWorkbookMdw()).toBe(7)
    setWorkbookMdw(8)
    expect(characterWidthToPixels(34.83203125)).toBe(284)
    const pixels = characterWidthToPixels(12)
    expect(Math.abs(pixelsToCharacterWidth(pixels) - 12)).toBeLessThan(0.05)
    expect(generalCharBudget(characterWidthToPixels(40))).toBe(40)
  })

  it('clamps invalid units and derives the padded built-in default', () => {
    setWorkbookMdw(Number.NaN)
    expect(getWorkbookMdw()).toBe(7)
    setWorkbookMdw(8)
    expect(paddedBaseColumnWidth(null)).toBe(8.625)
    expect(characterWidthToPixels(paddedBaseColumnWidth(null))).toBe(74)
    expect(paddedBaseColumnWidth(10)).toBe(10.625)
  })

  it('derives known Normal-font widths without browser font substitution', () => {
    expect(
      measureNormalFontMdw({ styles: [{ fontFamily: 'Calibri', fontSize: 11 }] } as never),
    ).toBe(8)
    expect(
      measureNormalFontMdw({ styles: [{ fontFamily: 'Verdana', fontSize: 10 }] } as never),
    ).toBe(8)
    expect(
      measureNormalFontMdw({ styles: [{ fontFamily: 'Aptos Narrow', fontSize: 11 }] } as never),
    ).toBe(8)
  })
})
