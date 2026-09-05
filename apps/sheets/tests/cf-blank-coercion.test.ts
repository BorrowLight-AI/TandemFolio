import { describe, expect, it } from 'vitest'

import { cellIsBlankDiverges } from '../src/renderer/univer-sync'

describe('cellIsBlankDiverges', () => {
  it('flags rules whose Excel blank-as-zero result differs from Univer', () => {
    expect(cellIsBlankDiverges('equal', 0, Number.NaN)).toBe(true)
    expect(cellIsBlankDiverges('lessThan', 5, Number.NaN)).toBe(true)
    expect(cellIsBlankDiverges('greaterThan', -1, Number.NaN)).toBe(true)
    expect(cellIsBlankDiverges('between', -1, 1)).toBe(true)
    expect(cellIsBlankDiverges('notEqual', 0, Number.NaN)).toBe(true)
    expect(cellIsBlankDiverges('notBetween', -1, 1)).toBe(true)
  })

  it('leaves rules on which both engines agree on the native condition', () => {
    expect(cellIsBlankDiverges('equal', 1, Number.NaN)).toBe(false)
    expect(cellIsBlankDiverges('greaterThan', 0, Number.NaN)).toBe(false)
    expect(cellIsBlankDiverges('notEqual', 5, Number.NaN)).toBe(false)
    expect(cellIsBlankDiverges('unknown', 0, Number.NaN)).toBe(false)
  })

  it('normalizes reversed between bounds', () => {
    expect(cellIsBlankDiverges('between', 1, -1)).toBe(true)
    expect(cellIsBlankDiverges('notBetween', 1, -1)).toBe(true)
  })
})
