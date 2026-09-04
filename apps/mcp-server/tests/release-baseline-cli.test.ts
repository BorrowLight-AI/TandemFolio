import { spawnSync } from 'node:child_process'
import { cpSync, mkdirSync, mkdtempSync, rmSync, symlinkSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'

const root = dirname(dirname(dirname(dirname(fileURLToPath(import.meta.url)))))

describe('release baseline capture CLI', () => {
  it('publishes the complete deterministic matrix in a checkout without release evidence', () => {
    const checkout = mkdtempSync(join(tmpdir(), 'tandemfolio-capture-cli-'))
    let result: ReturnType<typeof spawnSync>
    try {
      mkdirSync(join(checkout, 'tools'))
      cpSync(
        join(root, 'tools/capture-release-baseline.ts'),
        join(checkout, 'tools/capture-release-baseline.ts'),
      )
      cpSync(join(root, 'tools/release-gate'), join(checkout, 'tools/release-gate'), {
        recursive: true,
      })
      symlinkSync(join(root, 'node_modules'), join(checkout, 'node_modules'), 'junction')
      writeFileSync(join(checkout, 'package.json'), JSON.stringify({ type: 'module' }))
      result = spawnSync(
        process.execPath,
        ['--import', 'tsx', join(checkout, 'tools/capture-release-baseline.ts'), '--describe'],
        { cwd: checkout, encoding: 'utf8' },
      )
    } finally {
      rmSync(checkout, { recursive: true, force: true })
    }

    expect(result.status).toBe(0)
    expect(JSON.parse(String(result.stdout))).toEqual({
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
