import { describe, expect, it } from 'vitest'

import { excelWidthScale, hashFill, overflowHashes } from '../src/renderer/numfmt-fix'

const measure = (text: string): number => text.length * 8

describe('overflowHashes', () => {
  it('returns null when the text fits the column', () => {
    expect(overflowHashes('$472.00', 85, measure)).toBeNull()
  })

  it('fills available width with hashes on overflow', () => {
    expect(overflowHashes('2015-12-15 22:50', 85, measure)).toBe('##########')
  })

  it('keeps at least one hash in a sliver column and ignores empty text', () => {
    expect(overflowHashes('99', 6, measure)).toBe('#')
    expect(overflowHashes('', 85, measure)).toBeNull()
  })

  it('tolerates uncalibrated measurement noise', () => {
    expect(overflowHashes('2026/8/30x', 82, measure)).toBeNull()
  })
})

describe('excelWidthScale', () => {
  it('scales substituted Calibri back to its GDI digit width', () => {
    expect(excelWidthScale('Calibri', 11, () => 8.25, false)).toBeCloseTo(7 / 8.25)
  })

  it('never inflates a genuine or unknown font measurement', () => {
    expect(excelWidthScale('Calibri', 11, () => 6, false)).toBe(1)
    expect(excelWidthScale('Arial', 11, () => 8.25, false)).toBe(1)
    expect(excelWidthScale(undefined, 11, () => 8.25, false)).toBe(1)
  })

  it('can inflate a registered narrower substitute', () => {
    expect(excelWidthScale('Aptos Narrow', 11, () => 7.14, true)).toBeCloseTo(8 / 7.14)
    expect(excelWidthScale('Aptos Narrow', 11, () => 7.14, false)).toBe(1)
  })
})

describe('hashFill', () => {
  it('fills regardless of the source text', () => {
    expect(hashFill(85, measure)).toBe('##########')
  })

  it('returns null when the hash glyph cannot be measured', () => {
    expect(hashFill(85, () => 0)).toBeNull()
  })
})
