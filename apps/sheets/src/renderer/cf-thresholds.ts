import { parseAddress, type RangeBounds } from '../domain/cell-address'

/** A scale threshold as stored by the XLSX conditional-format model. */
export interface ScaleCfvo {
  kind: string
  value?: string | undefined
  gte?: boolean | undefined
}

export interface ThresholdReader {
  readValues(sheetName: string | null, range: RangeBounds): Promise<number[] | null>
  definedName(name: string): string | null
  tableColumn(table: string, column: string): { sheetName: string; range: RangeBounds } | null
}

export const THRESHOLD_RANGE_CELL_CAP = 512_000
const NAME_DEPTH_LIMIT = 3
const SHEET_PREFIX = String.raw`(?:(?:'([^']+)'|([A-Za-z0-9_.]+))!)?`
const CELL = String.raw`(\$?)([A-Za-z]{1,3})(\$?)([0-9]{1,7})`
const RANGE_ARGUMENT = new RegExp(
  String.raw`^\s*${SHEET_PREFIX}${CELL}(?::${CELL})?\s*(?:,\s*(-?[0-9]+(?:\.[0-9]+)?)\s*)?$`,
)
const AGGREGATE_CALL =
  /(?<![\w.])(AVERAGE|MIN|MAX|SUM|COUNT|MEDIAN|PERCENTILE(?:\.INC)?|QUARTILE(?:\.INC)?)\(([^()]*)\)/gi
const STRUCTURED_CALL =
  /(?<![\w.])(AVERAGE|MIN|MAX|SUM|COUNT|MEDIAN)\(\s*([A-Za-z_][\w.]*)\[([^\]]+)\]\s*\)/gi
const CELL_REFERENCE = new RegExp(String.raw`(?<![\w$.!'])${SHEET_PREFIX}${CELL}(?![\w(])`, 'g')
const NAME_TOKEN = /(?<![\w$.!'])[A-Za-z_\\][\w.]*(?![\w.(![])/g

export async function evaluateThresholdFormula(
  body: string,
  reader: ThresholdReader,
  depth = 0,
): Promise<number | null> {
  let expression = body.replace(/^=/, '').trim()
  if (expression === '' || expression.includes('"') || depth > NAME_DEPTH_LIMIT) return null

  const structured = await replaceAsync(expression, STRUCTURED_CALL, async (match) => {
    const column = reader.tableColumn(match[2]!, match[3]!)
    if (!column) return null
    const values = await reader.readValues(column.sheetName, column.range)
    return values === null
      ? null
      : aggregate(match[1]!.toUpperCase(), values.filter(Number.isFinite), undefined)
  })
  if (structured === null) return null
  expression = structured

  const aggregated = await replaceAsync(expression, AGGREGATE_CALL, async (match) => {
    const argument = RANGE_ARGUMENT.exec(match[2]!)
    if (!argument) return null
    const [, quoted, bare, c1, col1, r1, row1, c2, col2, r2, row2, k] = argument
    if (c1 !== '$' || r1 !== '$' || (col2 !== undefined && (c2 !== '$' || r2 !== '$'))) {
      return 0
    }
    const first = parseAddress(`${col1!.toUpperCase()}${row1}`)
    const second = col2 === undefined ? first : parseAddress(`${col2.toUpperCase()}${row2}`)
    const values = await reader.readValues(quoted ?? bare ?? null, {
      startRow: Math.min(first.row, second.row),
      endRow: Math.max(first.row, second.row),
      startColumn: Math.min(first.column, second.column),
      endColumn: Math.max(first.column, second.column),
    })
    return values === null
      ? null
      : aggregate(
          match[1]!.toUpperCase(),
          values.filter(Number.isFinite),
          k === undefined ? undefined : Number(k),
        )
  })
  if (aggregated === null) return null
  expression = aggregated

  const referenced = await replaceAsync(expression, CELL_REFERENCE, async (match) => {
    const [, quoted, bare, columnAnchor, column, rowAnchor, row] = match
    if (columnAnchor !== '$' || rowAnchor !== '$') return 0
    const cell = parseAddress(`${column!.toUpperCase()}${row}`)
    const values = await reader.readValues(quoted ?? bare ?? null, {
      startRow: cell.row,
      endRow: cell.row,
      startColumn: cell.column,
      endColumn: cell.column,
    })
    if (values === null) return null
    if (values.length === 0) return 0
    return Number.isFinite(values[0]) ? values[0]! : null
  })
  if (referenced === null) return null
  expression = referenced

  const named = await replaceAsync(expression, NAME_TOKEN, async (match) => {
    const formula = reader.definedName(match[0])
    return formula === null ? null : evaluateThresholdFormula(formula, reader, depth + 1)
  })
  return named === null ? null : evaluateArithmetic(named)
}

export function isSelfContainedFormula(body: string): boolean {
  const bare = body.replace(/^=/, '').replace(/"[^"]*"/g, '""')
  if (/[$![']/.test(bare)) return false
  if (/(?<![\w.])[A-Za-z]{1,3}[0-9]{1,7}(?![\w(])/.test(bare)) return false
  return [...bare.matchAll(/(?<![\w.])[A-Za-z_][\w.]*/g)].every((match) =>
    /^\s*\(/.test(bare.slice(match.index + match[0].length)),
  )
}

async function replaceAsync(
  input: string,
  pattern: RegExp,
  resolve: (match: RegExpExecArray) => Promise<number | null>,
): Promise<string | null> {
  const matches = [...input.matchAll(pattern)] as RegExpExecArray[]
  if (matches.length === 0) return input
  let output = ''
  let cursor = 0
  for (const match of matches) {
    const value = await resolve(match)
    if (value === null || !Number.isFinite(value)) return null
    output += input.slice(cursor, match.index) + `(${String(value)})`
    cursor = match.index + match[0].length
  }
  return output + input.slice(cursor)
}

export function percentileInc(sorted: readonly number[], k: number): number | null {
  if (sorted.length === 0 || !Number.isFinite(k) || k < 0 || k > 1) return null
  const rank = (sorted.length - 1) * k
  const lower = Math.floor(rank)
  const upper = Math.min(lower + 1, sorted.length - 1)
  return sorted[lower]! + (sorted[upper]! - sorted[lower]!) * (rank - lower)
}

function aggregate(name: string, values: number[], k: number | undefined): number | null {
  switch (name) {
    case 'SUM':
      return values.reduce((sum, value) => sum + value, 0)
    case 'COUNT':
      return values.length
    case 'MIN': {
      let low = Number.POSITIVE_INFINITY
      for (const value of values) if (value < low) low = value
      return values.length === 0 ? 0 : low
    }
    case 'MAX': {
      let high = Number.NEGATIVE_INFINITY
      for (const value of values) if (value > high) high = value
      return values.length === 0 ? 0 : high
    }
    case 'AVERAGE':
      return values.length === 0
        ? null
        : values.reduce((sum, value) => sum + value, 0) / values.length
    case 'MEDIAN':
      return percentileInc(
        [...values].sort((a, b) => a - b),
        0.5,
      )
    case 'PERCENTILE':
    case 'PERCENTILE.INC':
      return k === undefined
        ? null
        : percentileInc(
            [...values].sort((a, b) => a - b),
            k,
          )
    case 'QUARTILE':
    case 'QUARTILE.INC':
      return k === undefined || !Number.isInteger(k)
        ? null
        : percentileInc(
            [...values].sort((a, b) => a - b),
            k / 4,
          )
    default:
      return null
  }
}

export function evaluateArithmetic(expression: string): number | null {
  const source = expression.replace(/^=/, '').replace(/\s+/g, '')
  if (source === '' || !/^[\d+\-*/().eE]+$/.test(source)) return null
  let position = 0
  const parseExpression = (): number => {
    let value = parseTerm()
    while (source[position] === '+' || source[position] === '-') {
      const operator = source[position++]
      const term = parseTerm()
      value = operator === '+' ? value + term : value - term
    }
    return value
  }
  const parseTerm = (): number => {
    let value = parseFactor()
    while (source[position] === '*' || source[position] === '/') {
      const operator = source[position++]
      const factor = parseFactor()
      value = operator === '*' ? value * factor : value / factor
    }
    return value
  }
  const parseFactor = (): number => {
    if (source[position] === '-') {
      position += 1
      return -parseFactor()
    }
    if (source[position] === '+') {
      position += 1
      return parseFactor()
    }
    if (source[position] === '(') {
      position += 1
      const value = parseExpression()
      if (source[position] !== ')') return Number.NaN
      position += 1
      return value
    }
    const match = /^\d+(?:\.\d+)?(?:[eE][+-]?\d+)?/.exec(source.slice(position))
    if (!match) return Number.NaN
    position += match[0].length
    return Number(match[0])
  }
  const value = parseExpression()
  return position === source.length && Number.isFinite(value) ? value : null
}

export function defaultThreshold(ruleType: string, index: number, count: number): ScaleCfvo {
  if (index <= 0) return { kind: 'min' }
  if (ruleType === 'iconSet') {
    return { kind: 'percent', value: String(Math.round((index * 100) / count)) }
  }
  if (index >= count - 1) return { kind: 'max' }
  return { kind: 'percentile', value: '50' }
}

/**
 * Excel forces color-scale stops to be non-decreasing. Equal neighbours need
 * a tiny step below the later stop so its color owns their shared boundary.
 */
export function clampColorScaleStops<T extends ScaleCfvo>(cfvos: T[]): T[] {
  const stops = cfvos.map((cfvo) =>
    cfvo.kind === 'num' && cfvo.value !== undefined && Number.isFinite(Number(cfvo.value))
      ? Number(cfvo.value)
      : null,
  )
  let previous: number | null = null
  const clamped = stops.map((stop) => {
    if (stop === null) {
      previous = null
      return null
    }
    const lifted = previous !== null && stop < previous ? previous : stop
    previous = lifted
    return lifted
  })
  for (let index = clamped.length - 1; index > 0; index -= 1) {
    const current = clamped[index] ?? null
    const before = clamped[index - 1] ?? null
    if (current !== null && before !== null && before >= current) {
      clamped[index - 1] = current - Math.max(Math.abs(current) * 1e-9, 1e-9)
    }
  }
  if (clamped.every((stop, index) => stop === null || stop === stops[index])) return cfvos
  return cfvos.map((cfvo, index) => {
    const stop = clamped[index] ?? null
    return stop === null || stop === stops[index] ? cfvo : { ...cfvo, value: String(stop) }
  })
}

export interface DataBarLayoutInput {
  readonly cfvos: readonly ScaleCfvo[]
  readonly minLength?: number | undefined
  readonly maxLength?: number | undefined
  readonly axisPosition?: string | undefined
}

const DATA_DRIVEN_KINDS = new Set(['min', 'max', 'autoMin', 'autoMax', 'percent', 'percentile'])

export function dataBarNeedsLayout(input: DataBarLayoutInput): boolean {
  const [low, high] = input.cfvos
  if (!low || !high) return false
  if (low.kind === 'autoMin' || low.kind === 'autoMax') return true
  if (high.kind === 'autoMin' || high.kind === 'autoMax') return true
  return (
    (input.minLength ?? 0) !== 0 ||
    (input.maxLength ?? 100) !== 100 ||
    input.axisPosition === 'middle'
  )
}

export function dataBarNeedsValues(input: DataBarLayoutInput): boolean {
  return dataBarNeedsLayout(input) && input.cfvos.some((cfvo) => DATA_DRIVEN_KINDS.has(cfvo.kind))
}

export function resolveBarBound(
  cfvo: ScaleCfvo,
  values: readonly number[] | null,
  side: 'min' | 'max',
): number | null {
  if (cfvo.kind === 'num') {
    const value = Number(cfvo.value)
    return Number.isFinite(value) ? value : null
  }
  if (values === null) return null
  const finite = values.filter(Number.isFinite)
  if (finite.length === 0) return null
  let low = Number.POSITIVE_INFINITY
  let high = Number.NEGATIVE_INFINITY
  for (const value of finite) {
    if (value < low) low = value
    if (value > high) high = value
  }
  switch (cfvo.kind) {
    case 'min':
      return low
    case 'max':
      return high
    case 'autoMin':
      return Math.min(0, low)
    case 'autoMax':
      return Math.max(0, high)
    case 'percent': {
      const percent = Number(cfvo.value ?? (side === 'min' ? 0 : 100))
      return Number.isFinite(percent)
        ? low + (Math.max(0, Math.min(100, percent)) / 100) * (high - low)
        : null
    }
    case 'percentile': {
      const percent = Number(cfvo.value ?? (side === 'min' ? 0 : 100))
      return Number.isFinite(percent)
        ? percentileInc(
            [...finite].sort((a, b) => a - b),
            Math.max(0, Math.min(100, percent)) / 100,
          )
        : null
    }
    default:
      return null
  }
}

export function emulateBarExtents(
  min: number,
  max: number,
  minLength: number,
  maxLength: number,
): { min: number; max: number } | null {
  if (minLength === 0 && maxLength === 100) return null
  if (!(max > min) || !(maxLength > minLength) || minLength < 0 || maxLength > 100) return null
  const start = minLength / 100
  const end = maxLength / 100
  const span = (max - min) / (end - start)
  const lowered = min - start * span
  return lowered < 0 ? null : { min: lowered, max: lowered + span }
}

export function middleAxisBounds(min: number, max: number): { min: number; max: number } | null {
  if (!(min < 0 && max > 0)) return null
  const extent = Math.max(-min, max)
  return { min: -extent, max: extent }
}

export function layoutDataBar(
  input: DataBarLayoutInput,
  values: readonly number[] | null,
): ScaleCfvo[] | null {
  if (!dataBarNeedsLayout(input)) return null
  const [low, high] = input.cfvos as [ScaleCfvo, ScaleCfvo]
  const lower = resolveBarBound(low, values, 'min')
  const upper = resolveBarBound(high, values, 'max')
  const num = (value: number): ScaleCfvo => ({ kind: 'num', value: String(value) })
  const autoResolved = [
    low.kind === 'autoMin' || low.kind === 'autoMax' ? (lower === null ? low : num(lower)) : low,
    high.kind === 'autoMin' || high.kind === 'autoMax'
      ? upper === null
        ? high
        : num(upper)
      : high,
  ]
  const fallback = autoResolved[0] === low && autoResolved[1] === high ? null : autoResolved
  if (lower === null || upper === null) return fallback
  if (input.axisPosition === 'middle') {
    const bounds = middleAxisBounds(lower, upper)
    return bounds ? [num(bounds.min), num(bounds.max)] : fallback
  }
  const bounds = emulateBarExtents(lower, upper, input.minLength ?? 0, input.maxLength ?? 100)
  return bounds ? [num(bounds.min), num(bounds.max)] : fallback
}
