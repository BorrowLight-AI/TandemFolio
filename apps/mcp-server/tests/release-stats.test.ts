import { describe, expect, it } from 'vitest'
import { summarizeSamples, suggestedCeiling } from '../../../tools/release-gate/stats'

describe('release metric statistics', () => {
  it('uses deterministic nearest-rank percentiles and a 20 percent regression ceiling', () => {
    const summary = summarizeSamples([9, 1, 5, 3, 7])

    expect(summary).toEqual({ min: 1, median: 5, p95: 9, max: 9, samples: 5 })
    expect(suggestedCeiling(summary)).toBe(10.8)
  })

})
