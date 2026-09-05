import { describe, expect, it } from 'vitest'

import {
  CELL_INSET_PX,
  expandAsteriskFill,
  fillRepeatCount,
  fillRewriteForPattern,
  sectionFillToken,
} from '../src/renderer/numfmt-fix'

const NBSP = '\u00a0'
const measure = (text: string): number => text.length
const ACCOUNTING = '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)'

describe('sectionFillToken', () => {
  it('finds the last valid fill character and its span', () => {
    expect(sectionFillToken('_("$"* #,##0.00_)')).toEqual({ start: 5, end: 7, fill: ' ' })
    expect(sectionFillToken('0*-*x')).toEqual({ start: 3, end: 5, fill: 'x' })
    expect(sectionFillToken('0*\u{1F600}')).toEqual({ start: 1, end: 4, fill: '\u{1F600}' })
  })

  it('ignores quoted, escaped, skip-width, bracketed, and trailing asterisks', () => {
    expect(sectionFillToken('"*"0')).toBeNull()
    expect(sectionFillToken('0\\*')).toBeNull()
    expect(sectionFillToken('0_*')).toBeNull()
    expect(sectionFillToken('[$*-409]0')).toBeNull()
    expect(sectionFillToken('0*')).toBeNull()
  })
})

describe('fillRewriteForPattern', () => {
  it('marks each section fill and keeps quoted separators intact', () => {
    expect(fillRewriteForPattern(ACCOUNTING)?.fills).toEqual([' ', ' ', ' ', undefined])
    expect(fillRewriteForPattern('"a;b"* 0')?.pattern).toBe('"a;b""\uE000"0')
  })

  it('returns null for patterns without a fill', () => {
    expect(fillRewriteForPattern('#,##0.00;[Red](#,##0.00)')).toBeNull()
    expect(fillRewriteForPattern('General')).toBeNull()
  })
})

describe('fillRepeatCount', () => {
  it('fits whole repetitions into the remaining cell room', () => {
    expect(fillRepeatCount(100, 20, 7)).toBe(10)
    expect(CELL_INSET_PX).toBe(5)
    expect(fillRepeatCount(30, 40, 7)).toBe(0)
    expect(fillRepeatCount(100, 20, 0)).toBe(0)
  })
})

describe('expandAsteriskFill', () => {
  it('pins accounting symbols left and values right', () => {
    expect(expandAsteriskFill(ACCOUNTING, 1234.5, 25, measure)).toBe(
      `${NBSP}$${NBSP.repeat(9)}1,234.50${NBSP}`,
    )
    expect(expandAsteriskFill(ACCOUNTING, -1234.5, 25, measure)).toBe(
      `${NBSP}$${NBSP.repeat(8)}(1,234.50)`,
    )
  })

  it('supports trailing fills, leading fills, and text dot leaders', () => {
    expect(expandAsteriskFill('0*-', 5, 15, measure)).toBe(`5${'-'.repeat(9)}`)
    expect(expandAsteriskFill('*00', 42, 15, measure)).toBe(`${'0'.repeat(8)}42`)
    expect(expandAsteriskFill('@*.', 'Total', 20, measure)).toBe(`Total${'.'.repeat(10)}`)
  })

  it('returns null when no fill applies or no repetition fits', () => {
    expect(expandAsteriskFill('#,##0.00', 1, 25, measure)).toBeNull()
    expect(expandAsteriskFill(ACCOUNTING, 'n/a', 25, measure)).toBeNull()
    expect(expandAsteriskFill(ACCOUNTING, 1234567.89, 10, measure)).toBeNull()
  })

  it('uses measured widths for the repetition count', () => {
    const proportional = (text: string): number =>
      Array.from(text).reduce((sum, character) => sum + (character === NBSP ? 3 : 7), 0)
    expect(expandAsteriskFill('"$"* 0', 5, 50, proportional)).toBe(`$${NBSP.repeat(10)}5`)
  })
})
