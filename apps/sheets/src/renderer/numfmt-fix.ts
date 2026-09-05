/**
 * Excel-parity fixes for number-format display that Univer's own CELL_CONTENT
 * interceptors get wrong:
 *  - an empty format section renders '' — Univer treats '' as a formatting
 *    failure and falls back to the raw value, so `#,##0;(#,##0);` shows 0
 *  - string cells never enter the formatter, so the 4th (text) section of a
 *    pattern like `0.0_);(0.0);0.0_);@_)` is ignored
 *  - `_x`/`*x` padding comes out as plain spaces, which layout collapses on
 *    right-aligned cells; re-render with NBSP so accounting columns align
 *  - General is left at stripErrorMargin's 12 significant digits instead of
 *    Excel's "at most 11, shrunk to what the column width can hold"
 *
 * Registered just below Univer's NUMFMT interceptor (priority 10) so its
 * color handling and render cache still run first; this pass only fixes up
 * the resulting value, comparing against the raw cell to see what Univer did.
 */
import { CellValueType, InterceptorEffectEnum, isDefaultFormat, numfmt } from '@univerjs/core'
import { INTERCEPTOR_POINT, SheetInterceptorService } from '@univerjs/sheets'

import type { UniverRuntime } from './univer-state'

/// Same max-digit-width constant as characterWidthToPixels (univer-sync.ts),
/// so a column imported as N chars yields a budget of floor(N) regardless of
/// the render font.
const MDW = 7

export function generalCharBudget(columnWidthPx: number): number {
  return Math.max(1, Math.floor((columnWidthPx - 5) / MDW))
}

function toScientific(value: number, decimals: number): string {
  const [mantissa, exponent = '+0'] = value.toExponential(decimals).split('e')
  const sign = exponent.startsWith('-') ? '-' : '+'
  return `${mantissa}E${sign}${exponent.replace(/[+-]/, '').padStart(2, '0')}`
}

/**
 * Excel's General display: numfmt already applies the 11-significant-digit
 * cap and the fixed/scientific switch; on top of that, shrink decimals (then
 * the scientific mantissa) until the text fits the column's char budget.
 */
export function formatGeneral(value: number, budget: number): string {
  if (!Number.isFinite(value)) return String(value)
  const base = numfmt.format('General', value)
  if (base.length <= budget) return base
  const abs = Math.abs(value)
  const intLen = (abs < 1 ? 1 : Math.floor(Math.log10(abs)) + 1) + (value < 0 ? 1 : 0)
  for (let dec = Math.min(budget - intLen - 1, 10); dec >= 0; dec -= 1) {
    let t = value.toFixed(dec)
    if (t.includes('.')) t = t.replace(/0+$/, '').replace(/\.$/, '')
    if (Number(t) === 0 && value !== 0) break
    if (t.length <= budget) return t
  }
  for (let dec = 5; dec > 0; dec -= 1) {
    const t = toScientific(value, dec)
    if (t.length <= budget) return t
  }
  return toScientific(value, 0)
}

function safeFormat(pattern: string, value: number | string): string | null {
  try {
    return numfmt.format(pattern, value, { nbsp: true, throws: false })
  } catch {
    return null
  }
}

const NBSP = /\u00a0/g

/// Days between the 1900 and 1904 date-system epochs. numfmt is 1900-only,
/// so 1904 workbooks shift date serials before formatting while the model
/// keeps the file's original serial value for lossless saving.
export const DATE_1904_OFFSET = 1462

const datePatternCache = new Map<string, boolean>()

/// Calendar dates only: time-only and elapsed patterns render the serial's
/// magnitude and must not receive the epoch shift.
export function isCalendarDatePattern(pattern: string): boolean {
  let isDate = datePatternCache.get(pattern)
  if (isDate === undefined) {
    try {
      const type = (numfmt.getFormatInfo(pattern) as { type?: string }).type
      isDate = type === 'date' || type === 'datetime'
    } catch {
      isDate = false
    }
    datePatternCache.set(pattern, isDate)
  }
  return isDate
}

/**
 * Fix-ups for a cell that already went through Univer's NUMFMT interceptor.
 * `displayed` is the value Univer left in the cell, `raw` the model value.
 * Returns the corrected display text, or null to leave the cell alone.
 */
export function fixFormattedValue(
  pattern: string,
  raw: string | number,
  displayed: string | number | boolean | undefined,
  date1904 = false,
): string | null {
  if (typeof raw === 'string') {
    // Univer never formats plain-text cells: apply the text (4th) section.
    if (displayed !== raw) return null
    const text = safeFormat(pattern, raw)
    return text !== null && text !== raw ? text : null
  }
  const shiftDate1904 = date1904 && isCalendarDatePattern(pattern)
  const formatValue = shiftDate1904 ? raw + DATE_1904_OFFSET : raw
  const text = safeFormat(pattern, formatValue)
  if (text === null) return null
  if (shiftDate1904) return text === String(displayed) ? null : text
  // Empty section (e.g. `#,##0;(#,##0);` for 0, or `;;;`): Univer fell back
  // to the raw value; Excel renders an empty cell.
  if (text === '') return displayed === '' ? null : ''
  // Padding upgrade: only override when the outputs differ by NBSP alone, so
  // any locale-specific rendering Univer did stays untouched.
  if (
    typeof displayed === 'string' &&
    text !== displayed &&
    text.replace(NBSP, ' ') === displayed
  ) {
    return text
  }
  return null
}

export function installNumberFormatFix(
  runtime: UniverRuntime,
  isDate1904?: () => boolean,
): { dispose(): void } {
  const injector = runtime.univer.__getInjector()
  const interceptorService = injector.get(SheetInterceptorService)
  const cache = new Map<string, string | null>()
  return interceptorService.intercept(INTERCEPTOR_POINT.CELL_CONTENT, {
    // Below NUMFMT (10): Univer formats first (keeping its section colors and
    // render cache); this pass only corrects the value it left behind.
    priority: 9.5,
    effect: InterceptorEffectEnum.Value,
    handler: (cell, location, next) => {
      if (!cell || cell.p != null) return next(cell)
      if (cell.t === CellValueType.BOOLEAN || cell.t === CellValueType.FORCE_STRING) {
        return next(cell)
      }
      const raw = location.rawData?.v
      if (raw === undefined || raw === null || typeof raw === 'boolean') return next(cell)
      const style = location.workbook.getStyles().getStyleByCell(cell)
      const pattern = style?.n?.pattern
      if (isDefaultFormat(pattern)) {
        if (cell.t !== CellValueType.NUMBER || typeof raw !== 'number') return next(cell)
        const budget = generalCharBudget(location.worksheet.getColumnWidth(location.col))
        const key = `G\u0000${raw}\u0000${budget}`
        let text = cache.get(key)
        if (text === undefined) {
          text = formatGeneral(raw, budget)
          if (cache.size > 50_000) cache.clear()
          cache.set(key, text)
        }
        if (text === null || text === String(cell.v)) return next(cell)
        return next({ ...cell, v: text, t: CellValueType.NUMBER })
      }
      // Formula results are produced by Univer's 1900-based calculation
      // engine. Shift only static serials loaded from a 1904 workbook.
      const date1904 =
        isDate1904?.() === true &&
        typeof raw === 'number' &&
        location.rawData?.f == null &&
        location.rawData?.si == null
      const key = `${date1904 ? '1904' : '1900'}\u0000${pattern}\u0000${raw}\u0000${cell.v}`
      let text = cache.get(key)
      if (text === undefined) {
        text = fixFormattedValue(pattern as string, raw, cell.v ?? undefined, date1904)
        if (cache.size > 50_000) cache.clear()
        cache.set(key, text)
      }
      if (text === null) return next(cell)
      return next({
        ...cell,
        v: text,
        t: typeof raw === 'string' ? CellValueType.STRING : CellValueType.NUMBER,
      })
    },
  })
}
