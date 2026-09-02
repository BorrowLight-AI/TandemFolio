export const RELEASE_EVIDENCE_SCHEMA_VERSION = 4 as const

export interface ReleaseEvidence {
  schemaVersion: typeof RELEASE_EVIDENCE_SCHEMA_VERSION
  approved: boolean
  ready: boolean
  capturedAt: string
  sourceFingerprint: string
  upstream: {
    repository: string
    commit: string
  }
  formats: Record<'docx' | 'markdown' | 'xlsx' | 'pptx' | 'pdf', unknown>
  profile: Record<string, unknown>
}

export interface EvidenceIssue {
  path: string
  message: string
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}

export function parseReleaseEvidence(
  value: unknown,
): { ok: true; evidence: ReleaseEvidence } | { ok: false; issues: EvidenceIssue[] } {
  const issues: EvidenceIssue[] = []
  if (!isRecord(value)) return { ok: false, issues: [{ path: '$', message: 'must be an object' }] }

  if (value.schemaVersion !== RELEASE_EVIDENCE_SCHEMA_VERSION) {
    issues.push({ path: '$.schemaVersion', message: 'must equal 4' })
  }
  if (typeof value.approved !== 'boolean') {
    issues.push({ path: '$.approved', message: 'must be a boolean' })
  }
  if (typeof value.ready !== 'boolean') {
    issues.push({ path: '$.ready', message: 'must be a boolean' })
  }
  if (typeof value.capturedAt !== 'string' || Number.isNaN(Date.parse(value.capturedAt))) {
    issues.push({ path: '$.capturedAt', message: 'must be an ISO-8601 timestamp' })
  }
  if (
    typeof value.sourceFingerprint !== 'string' ||
    !/^[a-f0-9]{64}$/.test(value.sourceFingerprint)
  ) {
    issues.push({ path: '$.sourceFingerprint', message: 'must be a SHA-256 hex digest' })
  }
  if (!isRecord(value.upstream)) {
    issues.push({ path: '$.upstream', message: 'must be an object' })
  } else {
    if (typeof value.upstream.repository !== 'string') {
      issues.push({ path: '$.upstream.repository', message: 'must be a string' })
    }
    if (
      typeof value.upstream.commit !== 'string' ||
      !/^[a-f0-9]{40}$/.test(value.upstream.commit)
    ) {
      issues.push({ path: '$.upstream.commit', message: 'must be a full Git commit SHA' })
    }
  }
  if (!isRecord(value.formats)) {
    issues.push({ path: '$.formats', message: 'must be an object' })
  } else {
    for (const format of ['docx', 'markdown', 'xlsx', 'pptx', 'pdf']) {
      if (!(format in value.formats)) {
        issues.push({ path: `$.formats.${format}`, message: 'is required' })
        continue
      }
      const formatEvidence = value.formats[format]
      if (!isRecord(formatEvidence)) {
        issues.push({ path: `$.formats.${format}`, message: 'must be an object' })
        continue
      }
      for (const field of [
        'passed',
        'fixtures',
        'visual',
        'coldStartMs',
        'openDocumentMs',
        'interactionMs',
        'acknowledgement',
        'memory',
      ]) {
        if (!(field in formatEvidence)) {
          issues.push({ path: `$.formats.${format}.${field}`, message: 'is required' })
        }
      }
      if (format === 'markdown') {
        const phases = formatEvidence.stagedLoadPhases
        if (!isRecord(phases)) {
          issues.push({
            path: '$.formats.markdown.stagedLoadPhases',
            message: 'is required',
          })
        } else {
          for (const size of ['small', 'medium', 'large']) {
            const sizedPhases = phases[size]
            if (!isRecord(sizedPhases)) {
              issues.push({
                path: `$.formats.markdown.stagedLoadPhases.${size}`,
                message: 'must be an object',
              })
              continue
            }
            for (const phase of ['decodeMs', 'parseMs', 'tiptapStateInstallMs', 'reactCommitMs']) {
              if (!(phase in sizedPhases)) {
                issues.push({
                  path: `$.formats.markdown.stagedLoadPhases.${size}.${phase}`,
                  message: 'is required',
                })
              }
            }
          }
        }
      }
      if (format === 'xlsx') {
        const phases = formatEvidence.coldStartPhases
        if (!isRecord(phases)) {
          issues.push({
            path: '$.formats.xlsx.coldStartPhases',
            message: 'is required',
          })
        } else {
          for (const phase of [
            'bootstrapMs',
            'univerCreateMs',
            'worksheetInstallMs',
            'firstCommitMs',
          ]) {
            if (!(phase in phases)) {
              issues.push({
                path: `$.formats.xlsx.coldStartPhases.${phase}`,
                message: 'is required',
              })
            }
          }
          const bootstrapMetric = phases.bootstrapMs
          if (isRecord(bootstrapMetric) && bootstrapMetric.ceiling !== 500) {
            issues.push({
              path: '$.formats.xlsx.coldStartPhases.bootstrapMs.ceiling',
              message: 'must equal 500',
            })
          }
        }
        const bootstrapPhases = formatEvidence.bootstrapPhases
        if (!isRecord(bootstrapPhases)) {
          issues.push({
            path: '$.formats.xlsx.bootstrapPhases',
            message: 'is required',
          })
        } else {
          for (const phase of ['resourceReceiveMs', 'moduleGraphReadyMs', 'reactMountMs']) {
            if (!(phase in bootstrapPhases)) {
              issues.push({
                path: `$.formats.xlsx.bootstrapPhases.${phase}`,
                message: 'is required',
              })
            }
          }
        }
      }
    }
  }
  if (!isRecord(value.profile)) {
    issues.push({ path: '$.profile', message: 'must be an object' })
  }

  return issues.length
    ? { ok: false, issues }
    : { ok: true, evidence: value as unknown as ReleaseEvidence }
}
