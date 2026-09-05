import { describe, expect, it } from 'vitest'

import {
  defaultThreshold,
  emulateBarExtents,
  evaluateArithmetic,
  evaluateThresholdFormula,
  isSelfContainedFormula,
  layoutDataBar,
  middleAxisBounds,
  percentileInc,
  resolveBarBound,
  type ThresholdReader,
} from '../src/renderer/cf-thresholds'

const reader: ThresholdReader = {
  readValues: async (sheetName, range) => {
    if (sheetName?.toLowerCase() === 'other') return [7]
    if (range.startColumn === 1) return [10, 20, 30, 40]
    if (range.startColumn === 19) return [42]
    return []
  },
  definedName: (name) => (name.toLowerCase() === 'threshold' ? '$T$5*2' : null),
  tableColumn: (table, column) =>
    table === 'Sales' && column === 'Amount'
      ? { sheetName: 'Data', range: { startRow: 1, endRow: 3, startColumn: 1, endColumn: 1 } }
      : null,
}

describe('conditional-format formula thresholds', () => {
  it('folds absolute references, aggregates, names, tables, and arithmetic', async () => {
    expect(await evaluateThresholdFormula('$T$5', reader)).toBe(42)
    expect(await evaluateThresholdFormula("'Other'!$A$1+1", reader)).toBe(8)
    expect(await evaluateThresholdFormula('AVERAGE($B$2:$B$5)*2', reader)).toBe(50)
    expect(await evaluateThresholdFormula('Threshold/4', reader)).toBe(21)
    expect(await evaluateThresholdFormula('SUM(Sales[Amount])', reader)).toBe(100)
  })

  it('treats relative references as zero and rejects opaque expressions', async () => {
    expect(await evaluateThresholdFormula('3*A1+2', reader)).toBe(2)
    expect(await evaluateThresholdFormula('TODAY()-30', reader)).toBeNull()
    expect(await evaluateThresholdFormula('"12"', reader)).toBeNull()
  })

  it('recognizes formulas that do not depend on workbook cells', () => {
    expect(isSelfContainedFormula('TODAY()-30')).toBe(true)
    expect(isSelfContainedFormula('A1+1')).toBe(false)
    expect(isSelfContainedFormula('Threshold*2')).toBe(false)
  })

  it('evaluates arithmetic and percentile boundaries', () => {
    expect(evaluateArithmetic('3-(-5)')).toBe(8)
    expect(evaluateArithmetic('2*1e3')).toBe(2000)
    expect(evaluateArithmetic('1/0')).toBeNull()
    expect(percentileInc([1, 2, 3, 4], 0.5)).toBe(2.5)
  })

  it('supplies Excel defaults for unresolved threshold slots', () => {
    expect(defaultThreshold('colorScale', 0, 3)).toEqual({ kind: 'min' })
    expect(defaultThreshold('colorScale', 1, 3)).toEqual({ kind: 'percentile', value: '50' })
    expect(defaultThreshold('iconSet', 2, 3)).toEqual({ kind: 'percent', value: '67' })
  })
})

describe('native data-bar layout', () => {
  it('resolves data bounds and auto bounds', () => {
    expect(resolveBarBound({ kind: 'min' }, [1, 2, 4], 'min')).toBe(1)
    expect(resolveBarBound({ kind: 'autoMin' }, [1, 2, 4], 'min')).toBe(0)
    expect(resolveBarBound({ kind: 'autoMax' }, [-4, -1], 'max')).toBe(0)
  })

  it('maps legacy 10/90 extents onto widened numeric bounds', () => {
    const bounds = emulateBarExtents(1000, 5000, 10, 90)!
    expect((1000 - bounds.min) / (bounds.max - bounds.min)).toBeCloseTo(0.1, 12)
    expect((5000 - bounds.min) / (bounds.max - bounds.min)).toBeCloseTo(0.9, 12)
  })

  it('centres two-signed middle-axis bars', () => {
    expect(middleAxisBounds(-2, 4)).toEqual({ min: -4, max: 4 })
    expect(
      layoutDataBar(
        { cfvos: [{ kind: 'autoMin' }, { kind: 'autoMax' }], axisPosition: 'middle' },
        [-2, 1, 4],
      ),
    ).toEqual([
      { kind: 'num', value: '-4' },
      { kind: 'num', value: '4' },
    ])
  })
})
