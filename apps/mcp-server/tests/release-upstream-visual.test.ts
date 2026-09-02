import { execFileSync } from 'node:child_process'
import { createHash } from 'node:crypto'
import { readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'

const root = dirname(dirname(dirname(dirname(fileURLToPath(import.meta.url)))))

describe('pinned-upstream visual provenance', () => {
  it('grounds every reviewed screenshot in permitted paths at the pinned commit', () => {
    const manifest = JSON.parse(
      readFileSync(join(root, 'release/upstream-visual-manifest.json'), 'utf8'),
    ) as {
      upstreamCommit: string
      formats: Record<
        string,
        { baselinePath: string; baselineSha256: string; upstreamPaths: string[]; masks: unknown[] }
      >
    }

    expect(manifest.upstreamCommit).toBe('dc4d7e5927864498913b7ba42d0da06cc7cf628e')
    expect(Object.keys(manifest.formats)).toEqual(['docx', 'markdown', 'xlsx', 'pptx', 'pdf'])
    for (const format of Object.values(manifest.formats)) {
      expect(format.masks).toEqual([])
      expect(format.upstreamPaths.length).toBeGreaterThan(0)
      for (const path of format.upstreamPaths) {
        expect(path === 'ee' || path.startsWith('ee/')).toBe(false)
        execFileSync('git', ['cat-file', '-e', `${manifest.upstreamCommit}:${path}`], { cwd: root })
      }
      const bytes = readFileSync(join(root, format.baselinePath))
      expect(createHash('sha256').update(bytes).digest('hex')).toBe(format.baselineSha256)
    }
  })
})
