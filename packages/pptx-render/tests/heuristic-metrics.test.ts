import { describe, expect, it } from 'vitest'
import { HeuristicMetrics, type RunStyle } from '../src/metrics'

const metrics = new HeuristicMetrics()
const style: RunStyle = { fontFamily: '宋体', fontSizePx: 43, bold: false, italic: false }

describe('full-width fallback advances', () => {
  it('keeps CJK titles at one em per glyph in regular and bold faces', () => {
    const title = '物业企业应收账款管理研究'
    expect(metrics.measure(title, style)).toBe(516)
    expect(metrics.measure(title, { ...style, bold: true })).toBe(516)
  })

  it('preserves full-width punctuation while retaining the proportional Latin bold estimate', () => {
    const bold = { ...style, bold: true }
    expect(metrics.measure('中文，全角。', bold)).toBe(6 * 43)
    expect(metrics.measure('中A', bold)).toBeCloseTo(43 + metrics.measure('A', bold))
    expect(metrics.measure('A', bold)).toBeCloseTo(metrics.measure('A', style) * 1.04)
  })
})
