import { spawnSync } from 'node:child_process'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'

const root = dirname(dirname(dirname(dirname(fileURLToPath(import.meta.url)))))

describe('release baseline capture CLI', () => {
  it('publishes the complete deterministic release measurement matrix', () => {
    const result = spawnSync(
      join(root, 'node_modules', '.bin', 'tsx'),
      [join(root, 'tools', 'capture-release-baseline.ts'), '--describe'],
      { cwd: root, encoding: 'utf8' },
    )

    expect(result.status).toBe(0)
    expect(JSON.parse(result.stdout)).toEqual({
      formats: ['docx', 'markdown', 'xlsx', 'pptx', 'pdf'],
      fixtureSizes: ['small', 'medium', 'large'],
      measurements: [
        'pinnedUpstreamVisual',
        'coldStart',
        'openDocument',
        'interaction',
        'ack.pollWait',
        'ack.hydrate',
        'ack.execute',
        'ack.transport',
        'markdown.load.decode',
        'markdown.load.parse',
        'markdown.load.tiptapStateInstall',
        'markdown.load.reactCommit',
        'xlsx.coldStart.bootstrap',
        'xlsx.coldStart.univerCreate',
        'xlsx.coldStart.worksheetInstall',
        'xlsx.coldStart.firstCommit',
        'xlsx.bootstrap.resourceReceive',
        'xlsx.bootstrap.moduleGraphReady',
        'xlsx.bootstrap.reactMount',
        'peakJsHeap',
        'peakRendererMemory',
      ],
      defaultSamples: { coldStart: 7, openDocument: 7, interaction: 21 },
    })
  })
})
