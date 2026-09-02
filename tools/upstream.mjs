import { execFileSync, spawnSync } from 'node:child_process'
import { existsSync, readFileSync } from 'node:fs'
import { dirname, join, relative } from 'node:path'
import { fileURLToPath } from 'node:url'

const root = dirname(dirname(fileURLToPath(import.meta.url)))
const configPath = join(root, 'upstream.config.json')
const config = JSON.parse(readFileSync(configPath, 'utf8'))
const action = process.argv[2] ?? 'status'
const candidateRef = `refs/remotes/${config.remote}/${config.branch}`

function git(args, options = {}) {
  const output = execFileSync('git', args, {
    cwd: root,
    encoding: 'utf8',
    stdio: options.stdio ?? ['ignore', 'pipe', 'pipe'],
  })
  return typeof output === 'string' ? output.trim() : ''
}

function tryGit(args) {
  const result = spawnSync('git', args, { cwd: root, encoding: 'utf8' })
  return {
    ok: result.status === 0,
    stdout: result.stdout.trim(),
    stderr: result.stderr.trim(),
  }
}

function fail(message) {
  console.error(message)
  process.exit(1)
}

function validateConfig() {
  if (config.repository !== 'https://github.com/genspark-ai/genoffice.git') {
    fail(`unexpected upstream repository in ${relative(root, configPath)}`)
  }
  if (config.remote !== 'upstream' || config.branch !== 'main') {
    fail('upstream remote and branch must remain upstream/main')
  }
  if (!/^[0-9a-f]{40}$/.test(config.baseline)) {
    fail('upstream baseline must be a full 40-character commit SHA')
  }
  if (config.pushUrl !== 'DISABLED') {
    fail('upstream push URL must remain disabled')
  }
  if (config.fetchFilter !== 'blob:none') {
    fail('upstream fetches must use blob:none to avoid a second source checkout')
  }
  if (!Array.isArray(config.excludedPaths) || !config.excludedPaths.includes('ee/**')) {
    fail('upstream excluded paths must contain ee/**')
  }
  if (existsSync(join(root, 'upstream', 'genoffice'))) {
    fail('upstream/genoffice must not be checked out in TandemFolio; use the upstream remote')
  }
}

function configuredRemoteUrl() {
  const result = tryGit(['remote', 'get-url', config.remote])
  return result.ok ? result.stdout : undefined
}

function ensureGitConfig(key, value) {
  const current = tryGit(['config', '--get', key])
  if (!current.ok || current.stdout !== value) git(['config', key, value])
}

function ensureRemote({ create = false, announce = false } = {}) {
  const currentUrl = configuredRemoteUrl()
  if (!currentUrl && !create) {
    fail(`missing Git remote ${config.remote}; run npm run upstream:setup`)
  }
  if (!currentUrl) {
    git(['remote', 'add', config.remote, config.repository])
    ensureGitConfig(`remote.${config.remote}.pushurl`, config.pushUrl)
    ensureGitConfig(`remote.${config.remote}.promisor`, 'true')
    ensureGitConfig(`remote.${config.remote}.partialclonefilter`, config.fetchFilter)
    console.log(`Added ${config.remote} -> ${config.repository}`)
    return
  }
  if (currentUrl !== config.repository) {
    fail(
      `remote ${config.remote} points to ${currentUrl}; refusing to replace it with ${config.repository}`,
    )
  }
  ensureGitConfig(`remote.${config.remote}.pushurl`, config.pushUrl)
  ensureGitConfig(`remote.${config.remote}.promisor`, 'true')
  ensureGitConfig(`remote.${config.remote}.partialclonefilter`, config.fetchFilter)
  if (announce) console.log(`Remote ${config.remote} already matches ${config.repository}`)
}

function hasCommit(ref) {
  return tryGit(['cat-file', '-e', `${ref}^{commit}`]).ok
}

function requireFetchedRefs() {
  if (!hasCommit(config.baseline) || !hasCommit(candidateRef)) {
    fail(`upstream refs are not available locally; run npm run upstream:fetch`)
  }
}

function changedPaths() {
  const output = git([
    'diff',
    '--name-only',
    `${config.baseline}..${candidateRef}`,
    '--',
    '.',
    ':(exclude,glob)ee/**',
  ])
  return output ? output.split('\n').filter(Boolean) : []
}

function classify(path) {
  if (
    path.startsWith('packages/agent-core/') ||
    path.startsWith('packages/ai-') ||
    /(^|\/)ai(\/|$)/.test(path)
  ) {
    return 'excluded AI/product'
  }
  if (
    path.startsWith('apps/shell/') ||
    /(^|\/)src\/(main|preload)(\/|$)/.test(path) ||
    path.startsWith('packages/electron-utils/')
  ) {
    return 'excluded Electron/host'
  }
  if (path.includes('/test') || path.startsWith('fixtures/') || path.startsWith('e2e/')) {
    return 'tests/fixtures'
  }
  if (path.startsWith('apps/docs/') || path.startsWith('packages/docx-engine/')) return 'DOCX'
  if (path.startsWith('apps/sheets/')) return 'XLSX'
  if (path.startsWith('apps/slides/') || path.startsWith('packages/pptx-')) return 'PPTX'
  if (path.startsWith('apps/pdf/')) return 'PDF'
  if (
    path.startsWith('.github/') ||
    path === 'package.json' ||
    path === 'package-lock.json' ||
    path.startsWith('scripts/')
  ) {
    return 'build/repository'
  }
  return 'shared/other'
}

function printStatus() {
  ensureRemote()
  requireFetchedRefs()

  const candidate = git(['rev-parse', candidateRef])
  const isAncestor = tryGit(['merge-base', '--is-ancestor', config.baseline, candidateRef]).ok
  const commitCount = isAncestor
    ? git(['rev-list', '--count', `${config.baseline}..${candidateRef}`])
    : 'not comparable as a fast-forward'
  const paths = changedPaths()
  const groups = new Map()
  for (const path of paths) groups.set(classify(path), (groups.get(classify(path)) ?? 0) + 1)

  console.log(`Baseline:  ${config.baseline}`)
  console.log(`Candidate: ${candidate} (${config.remote}/${config.branch})`)
  console.log(`Commits:   ${commitCount}`)
  console.log(`Paths:     ${paths.length} community path(s) changed; ee/** was not inspected`)

  for (const [group, count] of [...groups.entries()].sort((left, right) =>
    left[0].localeCompare(right[0]),
  )) {
    console.log(`  ${group}: ${count}`)
  }

  if (isAncestor && candidate !== config.baseline) {
    const log = git([
      'log',
      '--no-decorate',
      '--format=%h %cs %s',
      '--max-count=20',
      `${config.baseline}..${candidateRef}`,
    ])
    if (log) console.log(`\nRecent upstream commits:\n${log}`)
  }
}

function validateRequestedPaths(paths) {
  const normalizedPaths = []
  for (const path of paths) {
    const normalized = path.replace(/^\.\//, '').replace(/\/$/, '')
    if (normalized === 'ee' || normalized.startsWith('ee/')) {
      fail('refusing to inspect the excluded ee/ tree')
    }
    if (normalized.startsWith('/') || normalized.split('/').includes('..')) {
      fail(`invalid upstream path: ${path}`)
    }
    normalizedPaths.push(normalized)
  }
  return normalizedPaths
}

function printDiff() {
  ensureRemote()
  requireFetchedRefs()
  const requestedPaths = process.argv.slice(3)
  const normalizedPaths = validateRequestedPaths(requestedPaths)
  const pathspecs = normalizedPaths.length
    ? [...normalizedPaths.map((path) => `:(literal)${path}`), ':(exclude,glob)ee/**']
    : ['.', ':(exclude,glob)ee/**']
  const args = normalizedPaths.length
    ? ['diff', `${config.baseline}..${candidateRef}`, '--', ...pathspecs]
    : ['diff', '--name-status', `${config.baseline}..${candidateRef}`, '--', ...pathspecs]
  const result = spawnSync('git', args, { cwd: root, stdio: 'inherit' })
  process.exit(result.status ?? 1)
}

function checkDocumentation() {
  const requiredFiles = [
    'AGENTS.md',
    'docs/adr/0001-tandemfolio-plugin-and-live-editor.md',
    'docs/migration/ledger.md',
    'docs/migration/provenance.md',
    'docs/migration/roadmap.md',
  ]
  for (const file of requiredFiles) {
    const contents = readFileSync(join(root, file), 'utf8')
    if (!contents.includes(config.baseline)) fail(`${file} does not record the configured baseline`)
  }
  const currentUrl = configuredRemoteUrl()
  if (currentUrl && currentUrl !== config.repository) {
    fail(`configured ${config.remote} remote does not match upstream.config.json`)
  }
  const pushUrl = tryGit(['remote', 'get-url', '--push', config.remote])
  if (currentUrl && (!pushUrl.ok || pushUrl.stdout !== config.pushUrl)) {
    fail(`configured ${config.remote} push URL is not disabled`)
  }
  console.log(`Verified upstream tracking configuration at ${config.baseline}.`)
}

validateConfig()

switch (action) {
  case 'setup':
    ensureRemote({ create: true, announce: true })
    break
  case 'fetch':
    ensureRemote({ create: true, announce: true })
    git(
      [
        'fetch',
        '--prune',
        '--no-tags',
        `--filter=${config.fetchFilter}`,
        config.remote,
        `refs/heads/${config.branch}:${candidateRef}`,
      ],
      { stdio: 'inherit' },
    )
    printStatus()
    break
  case 'status':
    printStatus()
    break
  case 'diff':
    printDiff()
    break
  case 'check':
    checkDocumentation()
    break
  default:
    fail('usage: node tools/upstream.mjs <setup|fetch|status|diff|check> [paths...]')
}
