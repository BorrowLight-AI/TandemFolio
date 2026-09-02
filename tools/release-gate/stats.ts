export interface MetricSummary {
  min: number
  median: number
  p95: number
  max: number
  samples: number
}

function rounded(value: number): number {
  return Number(value.toFixed(3))
}

function nearestRank(values: readonly number[], percentile: number): number {
  const rank = Math.max(1, Math.ceil(percentile * values.length))
  return values[Math.min(values.length - 1, rank - 1)]
}

export function summarizeSamples(samples: readonly number[]): MetricSummary {
  if (samples.length === 0) throw new Error('At least one release metric sample is required.')
  const values = [...samples].sort((left, right) => left - right)
  const middle = Math.floor(values.length / 2)
  const median =
    values.length % 2 === 0 ? (values[middle - 1] + values[middle]) / 2 : values[middle]
  return {
    min: rounded(values[0]),
    median: rounded(median),
    p95: rounded(nearestRank(values, 0.95)),
    max: rounded(values.at(-1)!),
    samples: values.length,
  }
}

export function suggestedCeiling(summary: MetricSummary): number {
  return rounded(summary.p95 * 1.2)
}
