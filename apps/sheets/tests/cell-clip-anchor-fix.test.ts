import { VerticalAlign } from '@univerjs/core'
import { describe, expect, it } from 'vitest'

import {
  effectiveTextDrawProps,
  excelClipOffsetTop,
  shouldAnchorTextTop,
} from '../src/renderer/cell-clip-anchor-fix'

describe('cell text clipping', () => {
  it('keeps the first line visible when multiline text overflows', () => {
    expect(excelClipOffsetTop(-24, 3)).toBe(0)
    expect(shouldAnchorTextTop(3, 54, 20)).toBe(true)
    expect(effectiveTextDrawProps({ text: 'a\\nb', fontStyle: '12px Arial', width: 80, height: 20, vAlign: VerticalAlign.BOTTOM }, 2, 36).vAlign).toBe(VerticalAlign.TOP)
  })

  it('preserves native alignment for single-line and fitting content', () => {
    expect(excelClipOffsetTop(-4, 1)).toBe(-4)
    expect(shouldAnchorTextTop(1, 22, 20)).toBe(false)
    expect(shouldAnchorTextTop(2, 20, 20)).toBe(false)
  })
})
