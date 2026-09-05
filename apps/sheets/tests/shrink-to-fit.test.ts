import { describe, expect, it } from 'vitest'

import { shrinkToFitFontSize } from '../src/renderer/univer-sync'

const measure = (line: string): number => line.length * 10

describe('shrinkToFitFontSize', () => {
  it('leaves text that already fits unchanged', () => {
    expect(shrinkToFitFontSize('1980', 12, 50, measure)).toBeNull()
  })

  it('scales overflowing text proportionally to an integer font size', () => {
    expect(shrinkToFitFontSize('1980X', 12, 42, measure)).toBe(10)
  })

  it('uses the widest line in a multiline cell', () => {
    expect(shrinkToFitFontSize('ab\r\nabcdef', 12, 30, measure)).toBe(6)
  })

  it('clamps at 1pt and rejects a zero-width column', () => {
    expect(shrinkToFitFontSize('abcdefghij', 12, 1, measure)).toBe(1)
    expect(shrinkToFitFontSize('abc', 12, 0, measure)).toBeNull()
  })
})
