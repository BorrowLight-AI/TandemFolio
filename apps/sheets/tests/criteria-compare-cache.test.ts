import { ArrayValueObject, NumberValueObject, StringValueObject } from '@univerjs/engine-formula'
import type { BaseValueObject } from '@univerjs/engine-formula'
import { afterEach, describe, expect, it } from 'vitest'

import { installCriteriaCompareCacheFix } from '../src/renderer/criteria-compare-cache'

let unit = 0
const disposables: { dispose(): void }[] = []
afterEach(() => disposables.splice(0).forEach((item) => item.dispose()))

function column(values: readonly (string | number)[], unitId: string) {
  return ArrayValueObject.create({
    calculateValueList: values.map((value) => [
      typeof value === 'number'
        ? (NumberValueObject.create(value) as BaseValueObject)
        : (StringValueObject.create(value) as BaseValueObject),
    ]),
    rowCount: values.length,
    columnCount: 1,
    unitId,
    sheetId: 'sheet1',
    row: 4,
    column: 7,
  } as unknown as Parameters<typeof ArrayValueObject.create>[0]) as ArrayValueObject
}

function repaired(result: ArrayValueObject, range: ArrayValueObject, criteria: BaseValueObject) {
  const mapped = (
    result as unknown as {
      mapValue(map: (value: BaseValueObject, row: number, column: number) => BaseValueObject): ArrayValueObject
    }
  ).mapValue((value, row, columnIndex) => {
    const cell = (range as unknown as { get(row: number, column: number): BaseValueObject }).get(
      row,
      columnIndex,
    )
    if (cell.isString() && criteria.isNumber()) {
      const coerced = cell.convertToNumberObjectValue()
      if (coerced.isNumber()) return coerced.compare(criteria, '=' as never)
    }
    return value
  })
  return mapped.getArrayValue().map((row) => row[0]?.getValue() === true)
}

describe('criteria compare cache', () => {
  it('keeps coerced date criteria matching text cells on repeated evaluations', () => {
    disposables.push(installCriteriaCompareCacheFix())
    const unitId = `wb-${unit++}`
    const criteria = NumberValueObject.create(46174) as BaseValueObject
    for (let call = 0; call < 3; call += 1) {
      const range = column(['2026-06', '2026-05', '2026-06'], unitId)
      expect(repaired(range.compare(criteria, '=' as never) as ArrayValueObject, range, criteria)).toEqual([
        true,
        false,
        true,
      ])
    }
  })

  it('restores the engine implementation on dispose', () => {
    const prototype = ArrayValueObject.prototype as unknown as Record<string, unknown>
    const original = prototype._batchOperatorValue
    const fix = installCriteriaCompareCacheFix()
    expect(prototype._batchOperatorValue).not.toBe(original)
    fix.dispose()
    expect(prototype._batchOperatorValue).toBe(original)
  })
})
