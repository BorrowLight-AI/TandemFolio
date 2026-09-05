import { ArrayValueObject, CELL_INVERTED_INDEX_CACHE, ERROR_TYPE_SET } from '@univerjs/engine-formula'

type Compare = (
  this: SheetArrayLike,
  value: CriteriaLike,
  column: number,
  result: unknown[],
  batchType: unknown,
  operator?: string,
  caseSensitive?: boolean,
) => void

interface CriteriaLike {
  isString?(): boolean
  isError?(): boolean
}

interface SheetArrayLike {
  getUnitId(): string
  getSheetId(): string
  getCurrentColumn(): number
}

interface CacheLike {
  canUseCache(...arguments_: unknown[]): unknown
  getCellValuePositions(unitId: string, sheetId: string, column: number): ReadonlyMap<unknown, unknown> | null
}

const verdicts = new WeakMap<object, { size: number; hasStringKey: boolean }>()

function cacheHoldsStrings(cache: CacheLike, array: SheetArrayLike, column: number): boolean {
  const keyed = cache.getCellValuePositions(
    array.getUnitId(),
    array.getSheetId(),
    column + array.getCurrentColumn(),
  )
  if (!keyed) return false
  const memo = verdicts.get(keyed)
  if (memo?.size === keyed.size) return memo.hasStringKey
  const errors = ERROR_TYPE_SET as ReadonlySet<string>
  const hasStringKey = [...keyed.keys()].some(
    (key) => typeof key === 'string' && !errors.has(key),
  )
  verdicts.set(keyed, { size: keyed.size, hasStringKey })
  return hasStringKey
}

export function installCriteriaCompareCacheFix(): { dispose(): void } {
  const prototype = ArrayValueObject.prototype as unknown as Record<string, Compare | undefined>
  const original = prototype._batchOperatorValue
  if (!original) return { dispose() {} }
  const cache = CELL_INVERTED_INDEX_CACHE as unknown as CacheLike
  const patched: Compare = function (
    value,
    column,
    result,
    batchType,
    operator,
    caseSensitive,
  ) {
    if (
      operator === '=' &&
      value.isString?.() === false &&
      value.isError?.() === false &&
      cacheHoldsStrings(cache, this, column)
    ) {
      const canUseCache = cache.canUseCache
      cache.canUseCache = () => ({ rowsInCache: [], rowsNotInCache: [] })
      try {
        return original.call(this, value, column, result, batchType, operator, caseSensitive)
      } finally {
        cache.canUseCache = canUseCache
      }
    }
    return original.call(this, value, column, result, batchType, operator, caseSensitive)
  }
  prototype._batchOperatorValue = patched
  return {
    dispose() {
      if (prototype._batchOperatorValue === patched) prototype._batchOperatorValue = original
    },
  }
}
