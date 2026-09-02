import { existsSync, readFileSync, writeFileSync } from 'node:fs'
import { createHash } from 'node:crypto'
import { dirname, resolve, sep } from 'node:path'
import { fileURLToPath } from 'node:url'
import { parseReleaseEvidence, type ReleaseEvidence } from './release-gate/contract'
import { computeReleaseSourceFingerprint } from './release-gate/fingerprint'
import { releaseMeasurementMatrix } from './release-gate/scenarios'

function argument(name: string): string | undefined {
  const index = process.argv.indexOf(name)
  return index === -1 ? undefined : process.argv[index + 1]
}

const evidencePath = resolve(argument('--evidence') ?? 'release/evidence.json')
const readinessPath = argument('--write-readiness')
  ? resolve(argument('--write-readiness')!)
  : undefined
const root = resolve(dirname(fileURLToPath(import.meta.url)), '..')

function writeReadiness(
  ready: boolean,
  upstreamCommit: string,
  sourceFingerprint: string,
): void {
  if (!readinessPath) return
  writeFileSync(
    readinessPath,
    `${JSON.stringify(
      {
        schemaVersion: 1,
        ready,
        formats: {
          docx: ready,
          markdown: ready,
          xlsx: ready,
          pptx: ready,
          pdf: ready,
        },
        upstreamCommit,
        sourceFingerprint,
      },
      null,
      2,
    )}\n`,
  )
}

function fail(reasons: Array<Record<string, unknown>>): void {
  if (readinessPath) {
    const upstream = JSON.parse(readFileSync(resolve(root, 'upstream.config.json'), 'utf8')) as {
      baseline: string
    }
    writeReadiness(false, upstream.baseline, computeReleaseSourceFingerprint(root))
  }
  process.stdout.write(`${JSON.stringify({ ok: false, ready: false, reasons })}\n`)
  process.exitCode = 1
}

function metricBudgetFailures(value: unknown, path = '$'): Array<Record<string, unknown>> {
  if (typeof value !== 'object' || value === null || Array.isArray(value)) return []
  const record = value as Record<string, unknown>
  if (typeof record.ceiling === 'number') {
    const measured =
      typeof record.measured === 'number'
        ? record.measured
        : typeof record.measured === 'object' &&
            record.measured !== null &&
            typeof (record.measured as { p95?: unknown }).p95 === 'number'
          ? (record.measured as { p95: number }).p95
          : undefined
    if (measured !== undefined && measured > record.ceiling) {
      return [
        {
          code: 'metric_budget_exceeded',
          path,
          measured,
          ceiling: record.ceiling,
        },
      ]
    }
  }
  return Object.entries(record).flatMap(([key, child]) =>
    metricBudgetFailures(child, `${path}.${key}`),
  )
}

function visualFailures(formats: ReleaseEvidence['formats']): Array<Record<string, unknown>> {
  return Object.entries(formats).flatMap(([format, evidence]) => {
    if (typeof evidence !== 'object' || evidence === null) return []
    const visual = (evidence as { visual?: unknown }).visual
    if (typeof visual !== 'object' || visual === null) return []
    const value = visual as Record<string, unknown>
    if (
      value.passed !== true ||
      typeof value.diffPixelRatio !== 'number' ||
      typeof value.maxDiffPixelRatio !== 'number' ||
      value.diffPixelRatio > value.maxDiffPixelRatio
    ) {
      return [
        {
          code: 'visual_gate_failed',
          format,
          diffPixelRatio: value.diffPixelRatio,
          maxDiffPixelRatio: value.maxDiffPixelRatio,
        },
      ]
    }
    return []
  })
}

function sampleFailures(profile: Record<string, unknown>): Array<Record<string, unknown>> {
  const samples = profile.samples
  if (typeof samples !== 'object' || samples === null) {
    return [
      { code: 'evidence_invalid', issues: [{ path: '$.profile.samples', message: 'is required' }] },
    ]
  }
  const record = samples as Record<string, unknown>
  return Object.entries(releaseMeasurementMatrix.defaultSamples).flatMap(([metric, expected]) => {
    const actual = record[metric]
    return typeof actual !== 'number' || actual < expected
      ? [{ code: 'insufficient_samples', metric, expected, actual }]
      : []
  })
}

function visualArtifactFailures(
  formats: ReleaseEvidence['formats'],
): Array<Record<string, unknown>> {
  return Object.entries(formats).flatMap(([format, evidence]) => {
    if (typeof evidence !== 'object' || evidence === null) return []
    const visual = (evidence as { visual?: unknown }).visual
    if (typeof visual !== 'object' || visual === null) return []
    const value = visual as Record<string, unknown>
    for (const kind of ['baseline', 'actual'] as const) {
      const path = value[`${kind}Path`]
      const expected = value[`${kind}Sha256`]
      if (typeof path !== 'string' || typeof expected !== 'string') continue
      const absolute = resolve(root, path)
      if (!absolute.startsWith(`${root}${sep}`) || !existsSync(absolute)) {
        return [{ code: 'visual_artifact_missing', format, kind, path }]
      }
      const actual = createHash('sha256').update(readFileSync(absolute)).digest('hex')
      if (actual !== expected) {
        return [{ code: 'visual_artifact_hash_mismatch', format, kind, path, expected, actual }]
      }
    }
    return []
  })
}

if (!existsSync(evidencePath)) {
  fail([{ code: 'evidence_missing', path: evidencePath }])
} else {
  let value: unknown
  try {
    value = JSON.parse(readFileSync(evidencePath, 'utf8'))
  } catch (error) {
    fail([
      {
        code: 'evidence_invalid',
        message: error instanceof Error ? error.message : String(error),
      },
    ])
  }
  if (value !== undefined) {
    const parsed = parseReleaseEvidence(value)
    if (!parsed.ok) {
      fail([{ code: 'evidence_invalid', issues: parsed.issues }])
    } else {
      const upstream = JSON.parse(readFileSync(resolve(root, 'upstream.config.json'), 'utf8')) as {
        baseline: string
      }
      if (parsed.evidence.upstream.commit !== upstream.baseline) {
        fail([
          {
            code: 'upstream_mismatch',
            expected: upstream.baseline,
            actual: parsed.evidence.upstream.commit,
          },
        ])
      } else {
        const expected = computeReleaseSourceFingerprint(root)
        if (parsed.evidence.sourceFingerprint !== expected) {
          fail([
            {
              code: 'source_mismatch',
              expected,
              actual: parsed.evidence.sourceFingerprint,
            },
          ])
        } else if (!parsed.evidence.approved) {
          fail([{ code: 'evidence_unapproved' }])
        } else {
          const insufficientSamples = sampleFailures(parsed.evidence.profile)
          if (insufficientSamples.length > 0) {
            fail(insufficientSamples)
          } else {
            const failedFormats = Object.entries(parsed.evidence.formats)
              .filter(([, evidence]) => {
                return (
                  typeof evidence !== 'object' ||
                  evidence === null ||
                  (evidence as { passed?: unknown }).passed !== true
                )
              })
              .map(([format]) => format)
            if (failedFormats.length > 0) {
              fail([{ code: 'format_gate_failed', formats: failedFormats }])
            } else if (visualFailures(parsed.evidence.formats).length > 0) {
              fail(visualFailures(parsed.evidence.formats))
            } else if (visualArtifactFailures(parsed.evidence.formats).length > 0) {
              fail(visualArtifactFailures(parsed.evidence.formats))
            } else if (metricBudgetFailures(parsed.evidence.formats, '$.formats').length > 0) {
              fail(metricBudgetFailures(parsed.evidence.formats, '$.formats'))
            } else if (!parsed.evidence.ready) {
              fail([{ code: 'evidence_not_ready' }])
            } else {
              writeReadiness(
                true,
                parsed.evidence.upstream.commit,
                parsed.evidence.sourceFingerprint,
              )
              process.stdout.write(`${JSON.stringify({ ok: true, ready: true, reasons: [] })}\n`)
            }
          }
        }
      }
    }
  }
}
