/** A scale threshold as stored by the XLSX conditional-format model. */
export interface ScaleCfvo {
  kind: string
  value?: string | undefined
  gte?: boolean | undefined
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
