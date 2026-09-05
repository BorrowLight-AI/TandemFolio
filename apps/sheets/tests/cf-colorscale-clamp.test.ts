import { describe, expect, it } from 'vitest'

import { clampColorScaleStops } from '../src/renderer/cf-thresholds'

const num = (value: string) => ({ kind: 'num', value })

describe('clampColorScaleStops', () => {
  it('lifts a later stop and epsilon-steps a shared boundary', () => {
    const clamped = clampColorScaleStops([num('0'), num('4'), num('0')])
    expect(Number(clamped[0]!.value)).toBe(0)
    expect(Number(clamped[2]!.value)).toBe(4)
    expect(Number(clamped[1]!.value)).toBeLessThan(4)
    expect(Number(clamped[1]!.value)).toBeGreaterThan(3.999)
  })

  it('returns the original array when stops are already ascending', () => {
    const stops = [num('0'), num('5'), num('10')]
    expect(clampColorScaleStops(stops)).toBe(stops)
  })

  it('does not clamp across a non-numeric stop', () => {
    const stops = [{ kind: 'min' }, { kind: 'percent', value: '50' }, num('2')]
    expect(clampColorScaleStops(stops)).toBe(stops)
  })
})
