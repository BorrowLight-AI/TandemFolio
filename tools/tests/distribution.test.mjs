import assert from 'node:assert/strict'
import { execFileSync, spawnSync } from 'node:child_process'
import { cp, mkdir, mkdtemp, readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { delimiter, dirname, join, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import test from 'node:test'
import {
  archiveDistribution,
  distributionVersion,
  stageDistribution,
} from '../package-distribution.mjs'
import { verifyBundle } from '../../distribution/verify.mjs'

const repo = resolve(dirname(fileURLToPath(import.meta.url)), '../..')

async function fixture(t) {
  const temp = await mkdtemp(join(tmpdir(), 'tandemfolio distribution '))
  t.after(() => rm(temp, { recursive: true, force: true }))
  const source = join(temp, 'source')
  const manifest = {
    name: 'tandemfolio',
    version: '1.2.3+codex.local',
    skills: './skills/',
    mcpServers: './.mcp.json',
  }
  const files = {
    'package.json': JSON.stringify({ version: '1.2.3' }),
    'plugins/tandemfolio/.codex-plugin/plugin.json': JSON.stringify(manifest),
    'plugins/tandemfolio/.mcp.json': JSON.stringify({
      mcpServers: { tandemfolio: { command: 'node', args: ['./dist/server.js'], cwd: '.' } },
    }),
    'plugins/tandemfolio/skills/tandemfolio/SKILL.md': '# Skill',
    'plugins/tandemfolio/assets/editor/index.html': '<html>DOCX</html>',
    'plugins/tandemfolio/dist/server.js':
      'import process from "node:process"; console.log(process.platform)',
    'plugins/tandemfolio/.env': 'must not ship',
    'docs/distribution.md': 'Install instructions',
    'docs/project-facts.md': 'pre-release',
    'docs/migration/provenance.md': 'provenance',
    'docs/migration/ledger.md': 'ledger',
    'apps/docs/src/renderer/fonts/README.md': 'fonts',
    LICENSE: 'Apache-2.0',
    NOTICE: 'upstream notice',
    'LICENSE-UNICODE.txt': 'Unicode notice',
  }
  for (const [path, contents] of Object.entries(files)) {
    await mkdir(dirname(join(source, path)), { recursive: true })
    await writeFile(join(source, path), contents)
  }
  await cp(join(repo, 'distribution'), join(source, 'distribution'), { recursive: true })
  const bundle = join(temp, 'tandemfolio-1.2.3')
  await stageDistribution(source, bundle, await distributionVersion(source, 'v1.2.3'))
  return { temp, source, bundle }
}

test('accepts an exact SemVer prerelease tag and rejects its stable counterpart', async (t) => {
  const temp = await mkdtemp(join(tmpdir(), 'tandemfolio prerelease '))
  t.after(() => rm(temp, { recursive: true, force: true }))
  await mkdir(join(temp, 'plugins/tandemfolio/.codex-plugin'), { recursive: true })
  await writeFile(join(temp, 'package.json'), JSON.stringify({ version: '0.1.0-beta.1' }))
  await writeFile(
    join(temp, 'plugins/tandemfolio/.codex-plugin/plugin.json'),
    JSON.stringify({ version: '0.1.0-beta.1+codex.local' }),
  )

  assert.equal(await distributionVersion(temp, 'v0.1.0-beta.1'), '0.1.0-beta.1')
  await assert.rejects(distributionVersion(temp, 'v0.1.0'), /Expected tag v0\.1\.0-beta\.1/)
})

test('stages an independent, verifiable marketplace without modifying source metadata', async (t) => {
  const { source, bundle } = await fixture(t)
  assert.equal(await verifyBundle(bundle), '1.2.3')
  await writeFile(join(bundle, '.DS_Store'), 'Finder metadata')
  assert.equal(await verifyBundle(bundle), '1.2.3')
  const original = JSON.parse(
    await readFile(join(source, 'plugins/tandemfolio/.codex-plugin/plugin.json')),
  )
  assert.equal(original.version, '1.2.3+codex.local')
  await assert.rejects(readFile(join(bundle, 'plugins/tandemfolio/.env')), { code: 'ENOENT' })
  assert.equal(
    await readFile(join(bundle, 'plugins/tandemfolio/LICENSE-UNICODE.txt'), 'utf8'),
    'Unicode notice',
  )
  assert.match(
    execFileSync(process.execPath, ['dist/server.js'], {
      cwd: join(bundle, 'plugins/tandemfolio'),
      encoding: 'utf8',
    }),
    new RegExp(process.platform),
  )
})

test('rejects tag/version mismatches and missing, modified, or unexpected package files', async (t) => {
  const { source, bundle } = await fixture(t)
  await assert.rejects(distributionVersion(source, 'v9.0.0'), /Expected tag/)
  const server = join(bundle, 'plugins/tandemfolio/dist/server.js')
  const original = await readFile(server)
  await writeFile(server, 'tampered')
  await assert.rejects(verifyBundle(bundle), /Checksum mismatch/)
  await writeFile(server, original)
  await writeFile(join(bundle, 'unexpected.txt'), 'extra')
  await assert.rejects(verifyBundle(bundle), /file list differs/)
  await rm(join(bundle, 'unexpected.txt'))
  await rm(server)
  await assert.rejects(verifyBundle(bundle), /file list differs/)
})

test('native installer check accepts a path containing spaces and rejects corruption', async (t) => {
  const { bundle } = await fixture(t)
  const command = process.platform === 'win32' ? 'powershell.exe' : 'bash'
  const args =
    process.platform === 'win32'
      ? ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', join(bundle, 'install.ps1'), '-Check']
      : [join(bundle, 'install.sh'), '--check']
  assert.match(execFileSync(command, args, { encoding: 'utf8' }), /Verified TandemFolio/)
  await writeFile(join(bundle, 'plugins/tandemfolio/dist/server.js'), 'tampered')
  assert.notEqual(spawnSync(command, args, { encoding: 'utf8' }).status, 0)
})

test(
  'ZIP and tar archives preserve hidden manifests, notices, and all checksums',
  { skip: process.platform === 'win32' },
  async (t) => {
    const { temp } = await fixture(t)
    archiveDistribution(temp, 'tandemfolio-1.2.3')
    for (const extension of ['zip', 'tar.gz']) {
      const extracted = join(temp, `extracted ${extension}`)
      await mkdir(extracted)
      if (extension === 'zip')
        execFileSync('unzip', ['-q', join(temp, `tandemfolio-1.2.3.${extension}`), '-d', extracted])
      else
        execFileSync('tar', ['-xzf', join(temp, `tandemfolio-1.2.3.${extension}`), '-C', extracted])
      assert.equal(await verifyBundle(join(extracted, 'tandemfolio-1.2.3')), '1.2.3')
    }
  },
)

test('native installer registers the exact path, updates without registration, and stops on CLI failure', async (t) => {
  const { temp, bundle } = await fixture(t)
  const bin = join(temp, 'fake cli')
  await mkdir(bin)
  const log = join(temp, 'commands.jsonl')
  const mock = join(bin, 'cli.mjs')
  await writeFile(
    mock,
    `import { appendFileSync } from 'node:fs';
const args = process.argv.slice(2);
appendFileSync(process.env.TEST_CODEX_LOG, JSON.stringify(args) + '\\n');
if (process.env.TEST_CODEX_FAIL === '1' && args.includes('marketplace')) process.exit(42);
`,
  )
  const shellQuote = (value) => `'${value.replaceAll("'", "'\\''")}'`
  if (process.platform === 'win32') {
    await writeFile(
      join(bin, 'codex.cmd'),
      `@echo off\r\n"${process.execPath}" "${mock}" %*\r\nexit /b %errorlevel%\r\n`,
    )
  } else {
    await writeFile(
      join(bin, 'codex'),
      `#!/usr/bin/env bash\nexec ${shellQuote(process.execPath)} ${shellQuote(mock)} "$@"\n`,
      { mode: 0o755 },
    )
  }
  const env = Object.fromEntries(
    Object.entries(process.env).filter(([key]) => key.toLowerCase() !== 'path'),
  )
  env.PATH = `${bin}${delimiter}${process.env.PATH ?? process.env.Path}`
  env.TEST_CODEX_LOG = log
  const command = process.platform === 'win32' ? 'powershell.exe' : 'bash'
  const args =
    process.platform === 'win32'
      ? ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', join(bundle, 'install.ps1')]
      : [join(bundle, 'install.sh')]
  const commands = async () => (await readFile(log, 'utf8')).trim().split('\n').map(JSON.parse)
  execFileSync(command, args, { env })
  const installed = await commands()
  assert.deepEqual(
    installed.map((entry) => entry.slice(0, 3)),
    [
      ['--version'],
      ['plugin', 'marketplace', 'add'],
      ['plugin', 'add', 'tandemfolio@tandemfolio-releases'],
    ],
  )
  // macOS resolves /var to /private/var; compare filesystem identity, not spelling.
  const { realpathSync } = await import('node:fs')
  assert.equal(realpathSync(installed[1][3]), realpathSync(bundle))
  await writeFile(log, '')
  execFileSync(command, [...args, process.platform === 'win32' ? '-Update' : '--update'], { env })
  assert.deepEqual(await commands(), [
    ['--version'],
    ['plugin', 'add', 'tandemfolio@tandemfolio-releases'],
  ])
  await writeFile(log, '')
  assert.notEqual(spawnSync(command, args, { env: { ...env, TEST_CODEX_FAIL: '1' } }).status, 0)
  assert.equal((await commands()).length, 2)
})
