import { spawnSync } from 'node:child_process'
import { createHash } from 'node:crypto'
import { mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'
import { computeReleaseSourceFingerprint } from '../../../tools/release-gate/fingerprint'

const root = dirname(dirname(dirname(dirname(fileURLToPath(import.meta.url)))))
const tsx = join(root, 'node_modules', '.bin', 'tsx')
const cli = join(root, 'tools', 'release-gate.ts')
const visualBaselinePath =
  'tests/visual/__screenshots__/host-widths.spec.ts/markdown-split-view.png'
const visualBaselineSha256 = createHash('sha256')
  .update(readFileSync(join(root, visualBaselinePath)))
  .digest('hex')

describe('release source fingerprint', () => {
  it('stays stable when a platform rebuild regenerates packaged editor bundles', () => {
    const directory = mkdtempSync(join(tmpdir(), 'genoffice-release-fingerprint-'))
    const generatedEditor = join(
      directory,
      'plugins/tandemfolio/assets/editors/xlsx/assets/index.js',
    )
    mkdirSync(dirname(generatedEditor), { recursive: true })
    writeFileSync(join(directory, 'package.json'), JSON.stringify({ version: '0.1.0-beta.1' }))
    writeFileSync(generatedEditor, 'macOS bundle')

    try {
      const before = computeReleaseSourceFingerprint(directory)
      writeFileSync(generatedEditor, 'Linux bundle')

      expect(computeReleaseSourceFingerprint(directory)).toBe(before)
    } finally {
      rmSync(directory, { recursive: true, force: true })
    }
  })

  it('changes when format-owned renderer source changes', () => {
    const directory = mkdtempSync(join(tmpdir(), 'genoffice-release-fingerprint-'))
    const rendererSource = join(directory, 'apps/docs/src/renderer/App.tsx')
    mkdirSync(dirname(rendererSource), { recursive: true })
    writeFileSync(join(directory, 'package.json'), JSON.stringify({ version: '0.1.0-beta.1' }))
    writeFileSync(rendererSource, 'export const version = 1')

    try {
      const before = computeReleaseSourceFingerprint(directory)
      writeFileSync(rendererSource, 'export const version = 2')

      expect(computeReleaseSourceFingerprint(directory)).not.toBe(before)
    } finally {
      rmSync(directory, { recursive: true, force: true })
    }
  })
})

function completeEvidence(overrides: Record<string, unknown> = {}) {
  const metric = {
    measured: { min: 1, median: 1, p95: 1, max: 1, samples: 7 },
    ceiling: 1.2,
  }
  const format = {
    passed: true,
    fixtures: {
      small: { bytes: 1, sha256: '1'.repeat(64) },
      medium: { bytes: 2, sha256: '2'.repeat(64) },
      large: { bytes: 3, sha256: '3'.repeat(64) },
    },
    visual: {
      passed: true,
      upstreamCommit: 'dc4d7e5927864498913b7ba42d0da06cc7cf628e',
      baselinePath: visualBaselinePath,
      actualPath: visualBaselinePath,
      baselineSha256: visualBaselineSha256,
      actualSha256: visualBaselineSha256,
      dimensionsMatch: true,
      diffPixelRatio: 0,
      maxDiffPixelRatio: 0.03,
    },
    coldStartMs: metric,
    openDocumentMs: { small: metric, medium: metric, large: metric },
    interactionMs: metric,
    acknowledgement: {
      pollWaitMs: metric,
      hydrateMs: metric,
      executeMs: metric,
      ackTransportMs: metric,
      totalMs: metric,
    },
    memory: {
      peakJsHeapBytes: { measured: 1, ceiling: 2 },
      peakRendererMemoryBytes: { measured: 1, ceiling: 2 },
    },
  }
  const markdown = {
    ...format,
    stagedLoadPhases: {
      small: {
        decodeMs: metric,
        parseMs: metric,
        tiptapStateInstallMs: metric,
        reactCommitMs: metric,
      },
      medium: {
        decodeMs: metric,
        parseMs: metric,
        tiptapStateInstallMs: metric,
        reactCommitMs: metric,
      },
      large: {
        decodeMs: metric,
        parseMs: metric,
        tiptapStateInstallMs: metric,
        reactCommitMs: metric,
      },
    },
  }
  const xlsx = {
    ...format,
    coldStartPhases: {
      bootstrapMs: { ...metric, ceiling: 500 },
      univerCreateMs: metric,
      worksheetInstallMs: metric,
      firstCommitMs: metric,
    },
    bootstrapPhases: {
      resourceReceiveMs: metric,
      moduleGraphReadyMs: metric,
      reactMountMs: metric,
    },
  }
  return {
    schemaVersion: 4,
    approved: false,
    ready: false,
    capturedAt: '2026-08-30T00:00:00.000Z',
    sourceFingerprint: computeReleaseSourceFingerprint(root),
    upstream: {
      repository: 'https://github.com/genspark-ai/genoffice.git',
      commit: 'dc4d7e5927864498913b7ba42d0da06cc7cf628e',
    },
    profile: {
      id: 'test',
      platform: 'darwin',
      arch: 'arm64',
      cpu: 'test',
      logicalCpuCount: 1,
      totalMemoryBytes: 1,
      browser: 'chromium',
      browserVersion: '140.0.0.0',
      viewport: { width: 1440, height: 1360, editorWidth: 720, editorHeight: 900 },
      samples: { coldStart: 7, openDocument: 7, interaction: 21 },
    },
    formats: { docx: format, markdown, xlsx, pptx: format, pdf: format },
    ...overrides,
  }
}

describe('release gate CLI', () => {
  it('fails closed with a machine-readable reason when release evidence is missing', () => {
    const missingEvidence = join(root, 'test-results', 'missing-release-evidence.json')
    const result = spawnSync(tsx, [cli, '--evidence', missingEvidence], {
      cwd: root,
      encoding: 'utf8',
    })

    expect(result.status).toBe(1)
    expect(JSON.parse(result.stdout)).toEqual({
      ok: false,
      ready: false,
      reasons: [
        {
          code: 'evidence_missing',
          path: missingEvidence,
        },
      ],
    })
  })

  it('rejects evidence that does not satisfy the release evidence contract', () => {
    const directory = mkdtempSync(join(tmpdir(), 'genoffice-release-gate-'))
    const evidence = join(directory, 'evidence.json')
    writeFileSync(evidence, JSON.stringify({ ready: true }))

    try {
      const result = spawnSync(tsx, [cli, '--evidence', evidence], {
        cwd: root,
        encoding: 'utf8',
      })

      expect(result.status).toBe(1)
      expect(JSON.parse(result.stdout)).toMatchObject({
        ok: false,
        ready: false,
        reasons: [{ code: 'evidence_invalid' }],
      })
    } finally {
      rmSync(directory, { recursive: true, force: true })
    }
  })

  it('rejects evidence captured against a different upstream commit', () => {
    const directory = mkdtempSync(join(tmpdir(), 'genoffice-release-gate-'))
    const evidence = join(directory, 'evidence.json')
    writeFileSync(
      evidence,
      JSON.stringify(
        completeEvidence({
          upstream: {
            repository: 'https://github.com/genspark-ai/genoffice.git',
            commit: '1'.repeat(40),
          },
        }),
      ),
    )

    try {
      const result = spawnSync(tsx, [cli, '--evidence', evidence], {
        cwd: root,
        encoding: 'utf8',
      })

      expect(result.status).toBe(1)
      expect(JSON.parse(result.stdout)).toMatchObject({
        ok: false,
        ready: false,
        reasons: [
          {
            code: 'upstream_mismatch',
            expected: 'dc4d7e5927864498913b7ba42d0da06cc7cf628e',
            actual: '1111111111111111111111111111111111111111',
          },
        ],
      })
    } finally {
      rmSync(directory, { recursive: true, force: true })
    }
  })

  it('rejects evidence when release-relevant source inputs have changed', () => {
    const directory = mkdtempSync(join(tmpdir(), 'genoffice-release-gate-'))
    const evidence = join(directory, 'evidence.json')
    writeFileSync(
      evidence,
      JSON.stringify(
        completeEvidence({
          sourceFingerprint: '0'.repeat(64),
        }),
      ),
    )

    try {
      const result = spawnSync(tsx, [cli, '--evidence', evidence], {
        cwd: root,
        encoding: 'utf8',
      })
      const output = JSON.parse(result.stdout)

      expect(result.status).toBe(1)
      expect(output).toMatchObject({
        ok: false,
        ready: false,
        reasons: [{ code: 'source_mismatch', actual: '0'.repeat(64) }],
      })
      expect(output.reasons[0].expected).toMatch(/^[a-f0-9]{64}$/)
      expect(output.reasons[0].expected).not.toBe('0'.repeat(64))
    } finally {
      rmSync(directory, { recursive: true, force: true })
    }
  })

  it('does not treat an unapproved candidate measurement as release evidence', () => {
    const directory = mkdtempSync(join(tmpdir(), 'genoffice-release-gate-'))
    const evidence = join(directory, 'evidence.json')
    writeFileSync(evidence, JSON.stringify(completeEvidence()))

    try {
      const result = spawnSync(tsx, [cli, '--evidence', evidence], {
        cwd: root,
        encoding: 'utf8',
      })

      expect(result.status).toBe(1)
      expect(JSON.parse(result.stdout)).toEqual({
        ok: false,
        ready: false,
        reasons: [{ code: 'evidence_unapproved' }],
      })
    } finally {
      rmSync(directory, { recursive: true, force: true })
    }
  })

  it('rejects approved evidence when any format gate failed', () => {
    const directory = mkdtempSync(join(tmpdir(), 'genoffice-release-gate-'))
    const evidencePath = join(directory, 'evidence.json')
    const evidence = completeEvidence({ approved: true, ready: false })
    evidence.formats = {
      ...evidence.formats,
      markdown: { ...evidence.formats.markdown, passed: false },
    }
    writeFileSync(evidencePath, JSON.stringify(evidence))

    try {
      const result = spawnSync(tsx, [cli, '--evidence', evidencePath], {
        cwd: root,
        encoding: 'utf8',
      })

      expect(result.status).toBe(1)
      expect(JSON.parse(result.stdout)).toEqual({
        ok: false,
        ready: false,
        reasons: [{ code: 'format_gate_failed', formats: ['markdown'] }],
      })
    } finally {
      rmSync(directory, { recursive: true, force: true })
    }
  })

  it('returns ready only for current, approved, all-format passing evidence', () => {
    const directory = mkdtempSync(join(tmpdir(), 'genoffice-release-gate-'))
    const evidencePath = join(directory, 'evidence.json')
    writeFileSync(evidencePath, JSON.stringify(completeEvidence({ approved: true, ready: true })))

    try {
      const result = spawnSync(tsx, [cli, '--evidence', evidencePath], {
        cwd: root,
        encoding: 'utf8',
      })

      expect(result.status).toBe(0)
      expect(JSON.parse(result.stdout)).toEqual({ ok: true, ready: true, reasons: [] })
    } finally {
      rmSync(directory, { recursive: true, force: true })
    }
  })

  it('generates the build-time readiness projection only from passing evidence', () => {
    const directory = mkdtempSync(join(tmpdir(), 'genoffice-release-gate-'))
    const evidencePath = join(directory, 'evidence.json')
    const outputPath = join(directory, 'release-readiness.json')
    writeFileSync(evidencePath, JSON.stringify(completeEvidence({ approved: true, ready: true })))

    try {
      const result = spawnSync(
        tsx,
        [cli, '--evidence', evidencePath, '--write-readiness', outputPath],
        { cwd: root, encoding: 'utf8' },
      )

      expect(result.status).toBe(0)
      expect(JSON.parse(readFileSync(outputPath, 'utf8'))).toEqual({
        schemaVersion: 1,
        ready: true,
        formats: { docx: true, markdown: true, xlsx: true, pptx: true, pdf: true },
        upstreamCommit: 'dc4d7e5927864498913b7ba42d0da06cc7cf628e',
        sourceFingerprint: computeReleaseSourceFingerprint(root),
      })
    } finally {
      rmSync(directory, { recursive: true, force: true })
    }
  })

  it('writes a not-ready projection when current evidence is rejected', () => {
    const directory = mkdtempSync(join(tmpdir(), 'tandemfolio-release-gate-'))
    const evidencePath = join(directory, 'evidence.json')
    const outputPath = join(directory, 'release-readiness.json')
    writeFileSync(
      evidencePath,
      JSON.stringify(completeEvidence({ sourceFingerprint: '0'.repeat(64) })),
    )

    try {
      const result = spawnSync(
        tsx,
        [cli, '--evidence', evidencePath, '--write-readiness', outputPath],
        { cwd: root, encoding: 'utf8' },
      )

      expect(result.status).toBe(1)
      expect(JSON.parse(readFileSync(outputPath, 'utf8'))).toEqual({
        schemaVersion: 1,
        ready: false,
        formats: { docx: false, markdown: false, xlsx: false, pptx: false, pdf: false },
        upstreamCommit: 'dc4d7e5927864498913b7ba42d0da06cc7cf628e',
        sourceFingerprint: computeReleaseSourceFingerprint(root),
      })
    } finally {
      rmSync(directory, { recursive: true, force: true })
    }
  })

  it('rejects a passing claim that omits a required measurement family', () => {
    const directory = mkdtempSync(join(tmpdir(), 'genoffice-release-gate-'))
    const evidencePath = join(directory, 'evidence.json')
    const evidence = completeEvidence({ approved: true, ready: true })
    const markdown = { ...evidence.formats.markdown } as Record<string, unknown>
    delete markdown.memory
    evidence.formats = {
      ...evidence.formats,
      markdown: markdown as typeof evidence.formats.markdown,
    }
    writeFileSync(evidencePath, JSON.stringify(evidence))

    try {
      const result = spawnSync(tsx, [cli, '--evidence', evidencePath], {
        cwd: root,
        encoding: 'utf8',
      })

      expect(result.status).toBe(1)
      expect(JSON.parse(result.stdout)).toMatchObject({
        ok: false,
        ready: false,
        reasons: [
          {
            code: 'evidence_invalid',
            issues: [{ path: '$.formats.markdown.memory', message: 'is required' }],
          },
        ],
      })
    } finally {
      rmSync(directory, { recursive: true, force: true })
    }
  })

  it('rejects Markdown evidence that omits the four staged-load phase families', () => {
    const directory = mkdtempSync(join(tmpdir(), 'genoffice-release-gate-'))
    const evidencePath = join(directory, 'evidence.json')
    const evidence = completeEvidence({ approved: true, ready: true })
    const markdown = { ...evidence.formats.markdown } as Record<string, unknown>
    delete markdown.stagedLoadPhases
    evidence.formats = {
      ...evidence.formats,
      markdown: markdown as typeof evidence.formats.markdown,
    }
    writeFileSync(evidencePath, JSON.stringify(evidence))

    try {
      const result = spawnSync(tsx, [cli, '--evidence', evidencePath], {
        cwd: root,
        encoding: 'utf8',
      })

      expect(result.status).toBe(1)
      expect(JSON.parse(result.stdout)).toMatchObject({
        ok: false,
        ready: false,
        reasons: [
          {
            code: 'evidence_invalid',
            issues: [{ path: '$.formats.markdown.stagedLoadPhases', message: 'is required' }],
          },
        ],
      })
    } finally {
      rmSync(directory, { recursive: true, force: true })
    }
  })

  it('rejects XLSX evidence that omits the four cold-start phase families', () => {
    const directory = mkdtempSync(join(tmpdir(), 'genoffice-release-gate-'))
    const evidencePath = join(directory, 'evidence.json')
    const evidence = completeEvidence({ approved: true, ready: true })
    const xlsx = { ...evidence.formats.xlsx } as Record<string, unknown>
    delete xlsx.coldStartPhases
    evidence.formats = {
      ...evidence.formats,
      xlsx: xlsx as typeof evidence.formats.xlsx,
    }
    writeFileSync(evidencePath, JSON.stringify(evidence))

    try {
      const result = spawnSync(tsx, [cli, '--evidence', evidencePath], {
        cwd: root,
        encoding: 'utf8',
      })

      expect(result.status).toBe(1)
      expect(JSON.parse(result.stdout)).toMatchObject({
        ok: false,
        ready: false,
        reasons: [
          {
            code: 'evidence_invalid',
            issues: [{ path: '$.formats.xlsx.coldStartPhases', message: 'is required' }],
          },
        ],
      })
    } finally {
      rmSync(directory, { recursive: true, force: true })
    }
  })

  it('rejects XLSX evidence that omits the three bootstrap phase families', () => {
    const directory = mkdtempSync(join(tmpdir(), 'genoffice-release-gate-'))
    const evidencePath = join(directory, 'evidence.json')
    const evidence = completeEvidence({ approved: true, ready: true })
    const xlsx = { ...evidence.formats.xlsx } as Record<string, unknown>
    delete xlsx.bootstrapPhases
    evidence.formats = {
      ...evidence.formats,
      xlsx: xlsx as typeof evidence.formats.xlsx,
    }
    writeFileSync(evidencePath, JSON.stringify(evidence))

    try {
      const result = spawnSync(tsx, [cli, '--evidence', evidencePath], {
        cwd: root,
        encoding: 'utf8',
      })

      expect(result.status).toBe(1)
      expect(JSON.parse(result.stdout)).toMatchObject({
        ok: false,
        ready: false,
        reasons: [
          {
            code: 'evidence_invalid',
            issues: [{ path: '$.formats.xlsx.bootstrapPhases', message: 'is required' }],
          },
        ],
      })
    } finally {
      rmSync(directory, { recursive: true, force: true })
    }
  })

  it('rejects XLSX evidence that changes the fixed bootstrap ceiling', () => {
    const directory = mkdtempSync(join(tmpdir(), 'genoffice-release-gate-'))
    const evidencePath = join(directory, 'evidence.json')
    const evidence = completeEvidence({ approved: true, ready: true })
    evidence.formats = {
      ...evidence.formats,
      xlsx: {
        ...evidence.formats.xlsx,
        coldStartPhases: {
          ...evidence.formats.xlsx.coldStartPhases,
          bootstrapMs: {
            ...evidence.formats.xlsx.coldStartPhases.bootstrapMs,
            ceiling: 501,
          },
        },
      },
    }
    writeFileSync(evidencePath, JSON.stringify(evidence))

    try {
      const result = spawnSync(tsx, [cli, '--evidence', evidencePath], {
        cwd: root,
        encoding: 'utf8',
      })

      expect(result.status).toBe(1)
      expect(JSON.parse(result.stdout)).toMatchObject({
        ok: false,
        ready: false,
        reasons: [
          {
            code: 'evidence_invalid',
            issues: [
              {
                path: '$.formats.xlsx.coldStartPhases.bootstrapMs.ceiling',
                message: 'must equal 500',
              },
            ],
          },
        ],
      })
    } finally {
      rmSync(directory, { recursive: true, force: true })
    }
  })

  it('rejects a passing claim when a measured p95 exceeds its approved ceiling', () => {
    const directory = mkdtempSync(join(tmpdir(), 'genoffice-release-gate-'))
    const evidencePath = join(directory, 'evidence.json')
    const evidence = completeEvidence({ approved: true, ready: true })
    evidence.formats = {
      ...evidence.formats,
      markdown: {
        ...evidence.formats.markdown,
        interactionMs: {
          measured: { min: 1, median: 1, p95: 2, max: 2, samples: 21 },
          ceiling: 1.2,
        },
      },
    }
    writeFileSync(evidencePath, JSON.stringify(evidence))

    try {
      const result = spawnSync(tsx, [cli, '--evidence', evidencePath], {
        cwd: root,
        encoding: 'utf8',
      })

      expect(result.status).toBe(1)
      expect(JSON.parse(result.stdout)).toEqual({
        ok: false,
        ready: false,
        reasons: [
          {
            code: 'metric_budget_exceeded',
            path: '$.formats.markdown.interactionMs',
            measured: 2,
            ceiling: 1.2,
          },
        ],
      })
    } finally {
      rmSync(directory, { recursive: true, force: true })
    }
  })

  it('rejects a passing claim when its pinned-upstream visual diff exceeds the ceiling', () => {
    const directory = mkdtempSync(join(tmpdir(), 'genoffice-release-gate-'))
    const evidencePath = join(directory, 'evidence.json')
    const evidence = completeEvidence({ approved: true, ready: true })
    evidence.formats = {
      ...evidence.formats,
      markdown: {
        ...evidence.formats.markdown,
        visual: {
          ...evidence.formats.markdown.visual,
          passed: false,
          diffPixelRatio: 0.04,
          maxDiffPixelRatio: 0.03,
        },
      },
    }
    writeFileSync(evidencePath, JSON.stringify(evidence))

    try {
      const result = spawnSync(tsx, [cli, '--evidence', evidencePath], {
        cwd: root,
        encoding: 'utf8',
      })

      expect(result.status).toBe(1)
      expect(JSON.parse(result.stdout)).toEqual({
        ok: false,
        ready: false,
        reasons: [
          {
            code: 'visual_gate_failed',
            format: 'markdown',
            diffPixelRatio: 0.04,
            maxDiffPixelRatio: 0.03,
          },
        ],
      })
    } finally {
      rmSync(directory, { recursive: true, force: true })
    }
  })

  it('rejects approved evidence captured with fewer than the reproducibility sample counts', () => {
    const directory = mkdtempSync(join(tmpdir(), 'genoffice-release-gate-'))
    const evidencePath = join(directory, 'evidence.json')
    const evidence = completeEvidence({ approved: true, ready: true })
    evidence.profile = {
      ...evidence.profile,
      samples: { coldStart: 1, openDocument: 7, interaction: 21 },
    }
    writeFileSync(evidencePath, JSON.stringify(evidence))

    try {
      const result = spawnSync(tsx, [cli, '--evidence', evidencePath], {
        cwd: root,
        encoding: 'utf8',
      })

      expect(result.status).toBe(1)
      expect(JSON.parse(result.stdout)).toEqual({
        ok: false,
        ready: false,
        reasons: [
          {
            code: 'insufficient_samples',
            metric: 'coldStart',
            expected: 7,
            actual: 1,
          },
        ],
      })
    } finally {
      rmSync(directory, { recursive: true, force: true })
    }
  })

  it('rejects visual evidence whose recorded artifact is unavailable', () => {
    const directory = mkdtempSync(join(tmpdir(), 'genoffice-release-gate-'))
    const evidencePath = join(directory, 'evidence.json')
    const evidence = completeEvidence({ approved: true, ready: true })
    evidence.formats = {
      ...evidence.formats,
      markdown: {
        ...evidence.formats.markdown,
        visual: {
          ...evidence.formats.markdown.visual,
          actualPath: 'release/artifacts/missing.png',
        },
      },
    }
    writeFileSync(evidencePath, JSON.stringify(evidence))

    try {
      const result = spawnSync(tsx, [cli, '--evidence', evidencePath], {
        cwd: root,
        encoding: 'utf8',
      })

      expect(result.status).toBe(1)
      expect(JSON.parse(result.stdout)).toEqual({
        ok: false,
        ready: false,
        reasons: [
          {
            code: 'visual_artifact_missing',
            format: 'markdown',
            kind: 'actual',
            path: 'release/artifacts/missing.png',
          },
        ],
      })
    } finally {
      rmSync(directory, { recursive: true, force: true })
    }
  })
})
