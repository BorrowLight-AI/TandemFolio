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
import {
  CellValueType,
  InterceptorEffectEnum,
  isDefaultFormat,
  numfmt,
  WrapStrategy,
} from '@univerjs/core'
import { FontCache, getFontStyleString } from '@univerjs/engine-render'
import { INTERCEPTOR_POINT, SheetInterceptorService } from '@univerjs/sheets'

import type { UniverRuntime } from './univer-state'
import { getWorkbookMdw } from './app-constants'

export const CELL_INSET_PX = 5

export function generalCharBudget(columnWidthPx: number): number {
  return Math.max(1, Math.floor((columnWidthPx - CELL_INSET_PX) / getWorkbookMdw()))
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

function patternSections(pattern: string): string[] {
  const sections: string[] = []
  let current = ''
  let quoted = false
  for (let index = 0; index < pattern.length; index += 1) {
    const character = pattern[index] ?? ''
    if (character === '"') quoted = !quoted
    if (character === ';' && !quoted) {
      sections.push(current)
      current = ''
      continue
    }
    if (character === '\\' && !quoted) {
      current += character + (pattern[index + 1] ?? '')
      index += 1
      continue
    }
    current += character
  }
  sections.push(current)
  return sections
}

export function sectionFillToken(
  section: string,
): { start: number; end: number; fill: string } | null {
  let token: { start: number; end: number; fill: string } | null = null
  for (let index = 0; index < section.length; index += 1) {
    const character = section[index]
    if (character === '"') {
      const end = section.indexOf('"', index + 1)
      if (end === -1) return null
      index = end
      continue
    }
    if (character === '[') {
      const end = section.indexOf(']', index + 1)
      if (end === -1) return null
      index = end
      continue
    }
    if (character === '\\' || character === '_') {
      index += 1
      continue
    }
    if (character === '*') {
      const codePoint = section.codePointAt(index + 1)
      if (codePoint === undefined) break
      const fill = String.fromCodePoint(codePoint)
      token = { start: index, end: index + 1 + fill.length, fill }
      index += fill.length
    }
  }
  return token
}

function fillSentinel(sectionIndex: number): string {
  return String.fromCharCode(0xe000 + sectionIndex)
}

const FILL_SENTINEL_RANGE = /[\uE000-\uE0FF]/g

interface FillRewrite {
  pattern: string
  fills: (string | undefined)[]
}

const fillRewriteCache = new Map<string, FillRewrite | null>()

export function fillRewriteForPattern(pattern: string): FillRewrite | null {
  let rewrite = fillRewriteCache.get(pattern)
  if (rewrite !== undefined) return rewrite
  rewrite = null
  if (pattern.includes('*')) {
    const sections = patternSections(pattern)
    const fills: (string | undefined)[] = []
    const rewritten = sections.map((section, index) => {
      const token = sectionFillToken(section)
      fills.push(token?.fill)
      if (!token) return section
      return `${section.slice(0, token.start)}"${fillSentinel(index)}"${section.slice(token.end)}`
    })
    if (fills.some((fill) => fill !== undefined)) rewrite = { pattern: rewritten.join(';'), fills }
  }
  if (fillRewriteCache.size > 5_000) fillRewriteCache.clear()
  fillRewriteCache.set(pattern, rewrite)
  return rewrite
}

export function fillRepeatCount(
  columnWidthPx: number,
  textWidthPx: number,
  fillWidthPx: number,
): number {
  if (!(fillWidthPx > 0)) return 0
  return Math.max(0, Math.floor((columnWidthPx - CELL_INSET_PX - textWidthPx) / fillWidthPx))
}

export function expandAsteriskFill(
  pattern: string,
  value: number | string,
  columnWidthPx: number,
  measure: (text: string) => number,
): string | null {
  const rewrite = fillRewriteForPattern(pattern)
  if (!rewrite) return null
  const marked = safeFormat(rewrite.pattern, value)
  if (marked === null) return null
  const sentinels = marked.match(FILL_SENTINEL_RANGE)
  if (sentinels === null || sentinels.length !== 1) return null
  const sentinel = sentinels[0] as string
  const sectionIndex = sentinel.charCodeAt(0) - 0xe000
  const rawFill = rewrite.fills[sectionIndex]
  if (rawFill === undefined) return null
  const at = marked.indexOf(sentinel)
  const head = marked.slice(0, at)
  const tail = marked.slice(at + 1)
  if (head + tail !== safeFormat(pattern, value)) return null
  const fill = rawFill === ' ' ? '\u00a0' : rawFill
  const count = fillRepeatCount(columnWidthPx, measure(head + tail), measure(fill))
  if (count <= 0) return null
  return head + fill.repeat(count) + tail
}

interface MergedSpanSheet {
  getMergedCell(
    row: number,
    column: number,
  ): { startColumn: number; endColumn: number } | null | undefined | void
  getColumnWidth(column: number): number
  getColVisible(column: number): boolean
}

export function mergedSpanWidth(
  sheet: MergedSpanSheet,
  row: number,
  column: number,
): number | null {
  const merged = sheet.getMergedCell(row, column)
  if (!merged) return null
  let total = 0
  for (let current = merged.startColumn; current <= merged.endColumn; current += 1) {
    if (sheet.getColVisible(current)) total += sheet.getColumnWidth(current)
  }
  return total
}

export function hashFill(
  columnWidthPx: number,
  measure: (text: string) => number,
): string | null {
  const hashWidth = measure('#')
  if (!(hashWidth > 0)) return null
  return '#'.repeat(Math.max(1, Math.floor((columnWidthPx - CELL_INSET_PX) / hashWidth)))
}

export function overflowHashes(
  display: string,
  columnWidthPx: number,
  measure: (text: string) => number,
  scale = 1,
  calibrated = scale !== 1,
): string | null {
  if (display === '') return null
  const available = columnWidthPx - CELL_INSET_PX
  const limit = calibrated ? available - measure('0') * scale : available * 1.05
  if (measure(display) * scale <= limit) return null
  return hashFill(columnWidthPx, measure)
}

const EXCEL_DIGIT_PER_PT: Record<string, number> = {
  Calibri: 7 / 11,
  Verdana: 8 / 10,
  'Malgun Gothic': 7 / 11,
  '맑은 고딕': 7 / 11,
  'Aptos Narrow': 8 / 11,
  'ＭＳ Ｐゴシック': 8 / 11,
  'MS PGothic': 8 / 11,
  宋体: 8 / 11,
  SimSun: 8 / 11,
  新細明體: 8 / 11,
  PMingLiU: 8 / 11,
}

export function excelWidthScale(
  family: string | undefined,
  sizePt: number,
  measureDigit: () => number,
  substituteActive?: boolean,
): number {
  if (!family) return 1
  const perPt = EXCEL_DIGIT_PER_PT[family]
  if (perPt === undefined || (substituteActive === undefined && fontAvailable(family))) return 1
  const digit = measureDigit()
  if (!(digit > 0)) return 1
  const scale = (perPt * sizePt) / digit
  return substituteActive === true ? scale : Math.min(1, scale)
}

const fontAvailabilityCache = new Map<string, boolean>()

function fontAvailable(family: string): boolean {
  let known = fontAvailabilityCache.get(family)
  if (known !== undefined) return known
  known = false
  try {
    const context =
      typeof document === 'undefined' ? null : document.createElement('canvas').getContext('2d')
    if (context) {
      const width = (font: string): number => {
        context.font = `16px ${font}`
        return context.measureText('01mWi').width
      }
      known =
        width(`"${family}", monospace`) !== width('monospace') ||
        width(`"${family}", serif`) !== width('serif')
    }
  } catch {
    known = false
  }
  fontAvailabilityCache.set(family, known)
  return known
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

const patternTypeCache = new Map<string, string>()

function patternType(pattern: string): string {
  let type = patternTypeCache.get(pattern)
  if (type !== undefined) return type
  try {
    type = (numfmt.getFormatInfo(pattern) as { type?: string }).type ?? 'unknown'
  } catch {
    type = 'unknown'
  }
  patternTypeCache.set(pattern, type)
  return type
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
      if (cell.t === CellValueType.BOOLEAN) {
        const style = location.workbook.getStyles().getStyleByCell(cell)
        if (
          style?.tb !== WrapStrategy.WRAP &&
          !style?.tr?.a &&
          !style?.tr?.v &&
          !location.worksheet.getMergedCell(location.row, location.col)
        ) {
          const fontString = getFontStyleString(style ?? undefined).fontString
          const measure = (text: string): number => FontCache.getMeasureText(text, fontString).width
          const hashes = overflowHashes(
            cell.v === 0 || cell.v === false ? 'FALSE' : 'TRUE',
            location.worksheet.getColumnWidth(location.col),
            measure,
            excelWidthScale(style?.ff ?? undefined, style?.fs ?? 11, () => measure('0')),
          )
          if (hashes !== null) return next({ ...cell, v: hashes, t: CellValueType.NUMBER })
        }
        return next(cell)
      }
      if (cell.t === CellValueType.FORCE_STRING) return next(cell)
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
      const maybeHash = (outCell: typeof cell, patternUsed: unknown): typeof cell => {
        if (typeof raw !== 'number' || typeof patternUsed !== 'string') return outCell
        if (isDefaultFormat(patternUsed)) return outCell
        const type = patternType(patternUsed)
        if (type === 'text' || type === 'unknown') return outCell
        if (style?.tb === WrapStrategy.WRAP || style?.tr?.a || style?.tr?.v) return outCell
        const fontString = getFontStyleString(style ?? undefined).fontString
        const measure = (text: string): number => FontCache.getMeasureText(text, fontString).width
        const width =
          mergedSpanWidth(location.worksheet, location.row, location.col) ??
          location.worksheet.getColumnWidth(location.col)
        const negativeDate =
          raw < 0 && !date1904 && (type === 'date' || type === 'datetime' || type === 'time')
        const hashes = negativeDate
          ? hashFill(width, measure)
          : overflowHashes(
              String(outCell.v ?? ''),
              width,
              measure,
              excelWidthScale(style?.ff ?? undefined, style?.fs ?? 11, () => measure('0')),
            )
        return hashes === null ? outCell : { ...outCell, v: hashes, t: CellValueType.NUMBER }
      }
      if (
        typeof pattern === 'string' &&
        pattern.includes('*') &&
        (typeof raw === 'number' || (typeof raw === 'string' && !raw.startsWith('#'))) &&
        style?.tb !== WrapStrategy.WRAP &&
        !style?.tr?.a &&
        !style?.tr?.v
      ) {
        const fontString = getFontStyleString(style ?? undefined).fontString
        const measure = (text: string): number => FontCache.getMeasureText(text, fontString).width
        const width =
          mergedSpanWidth(location.worksheet, location.row, location.col) ??
          location.worksheet.getColumnWidth(location.col)
        const fillValue =
          date1904 && typeof raw === 'number' && isCalendarDatePattern(pattern)
            ? raw + DATE_1904_OFFSET
            : raw
        const fillKey = `*\u0000${typeof fillValue}\u0000${pattern}\u0000${fillValue}\u0000${width}\u0000${fontString}`
        let expanded = cache.get(fillKey)
        if (expanded === undefined) {
          expanded = expandAsteriskFill(pattern, fillValue, width, measure)
          if (cache.size > 50_000) cache.clear()
          cache.set(fillKey, expanded)
        }
        if (expanded !== null && expanded !== String(cell.v)) {
          return next({
            ...cell,
            v: expanded,
            t: typeof raw === 'string' ? CellValueType.STRING : CellValueType.NUMBER,
          })
        }
      }
      const key = `${date1904 ? '1904' : '1900'}\u0000${pattern}\u0000${raw}\u0000${cell.v}`
      let text = cache.get(key)
      if (text === undefined) {
        text = fixFormattedValue(pattern as string, raw, cell.v ?? undefined, date1904)
        if (cache.size > 50_000) cache.clear()
        cache.set(key, text)
      }
      if (text === null) return next(maybeHash(cell, pattern))
      return next(
        maybeHash(
          {
            ...cell,
            v: text,
            t: typeof raw === 'string' ? CellValueType.STRING : CellValueType.NUMBER,
          },
          pattern,
        ),
      )
    },
  })
}
