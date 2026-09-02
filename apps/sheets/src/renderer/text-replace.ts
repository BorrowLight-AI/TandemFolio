import { ObjectMatrix, type ICellData } from '@univerjs/core'

import { commitWorkbookRangeCells } from './range-copy'
import type { UniverRuntime, UniverWorksheet } from './univer-state'

type UniverRange = ReturnType<UniverWorksheet['getRange']>

function replaceOccurrences(
  value: string,
  find: string,
  replacement: string,
  matchCase: boolean,
): string {
  if (matchCase) return value.split(find).join(replacement)
  const escaped = find.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
  return value.replace(new RegExp(escaped, 'gi'), replacement)
}

/**
 * Replaces literal text cells in one range while preserving formulas and all
 * untouched cells. The sparse mutation plus inverse is one native undo item.
 */
export function applyWorkbookTextReplacement(
  runtime: UniverRuntime,
  range: UniverRange,
  input: {
    readonly find: string
    readonly replace: string
    readonly matchCase: boolean
    readonly wholeCell: boolean
  },
): number {
  const values = range.getValues()
  const formulas = range.getFormulas()
  const changes = new ObjectMatrix<ICellData>()
  let changed = 0
  for (let row = 0; row < range.getHeight(); row += 1) {
    for (let column = 0; column < range.getWidth(); column += 1) {
      if (formulas[row]?.[column]) continue
      const current = values[row]?.[column]
      if (typeof current !== 'string') continue
      const haystack = input.matchCase ? current : current.toLocaleLowerCase()
      const needle = input.matchCase ? input.find : input.find.toLocaleLowerCase()
      const next = input.wholeCell
        ? haystack === needle
          ? input.replace
          : current
        : haystack.includes(needle)
          ? replaceOccurrences(current, input.find, input.replace, input.matchCase)
          : current
      if (next === current) continue
      changes.setValue(range.getRow() + row, range.getColumn() + column, { v: next, f: null })
      changed += 1
    }
  }
  if (changed === 0) return 0
  if (!commitWorkbookRangeCells(runtime, range, changes)) {
    throw new Error('Univer rejected the XLSX text replacement mutation.')
  }
  range.activate()
  return changed
}
