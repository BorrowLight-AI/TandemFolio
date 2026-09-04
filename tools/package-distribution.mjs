import { execFileSync } from 'node:child_process'
import { createHash } from 'node:crypto'
import { realpathSync } from 'node:fs'
import { cp, mkdir, mkdtemp, readFile, writeFile } from 'node:fs/promises'
import { dirname, join, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { filesUnder, verifyBundle } from '../distribution/verify.mjs'

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..')
const json = async (path) => JSON.parse(await readFile(path, 'utf8'))
const writeJson = (path, value) => writeFile(path, `${JSON.stringify(value, null, 2)}\n`)
const semver =
  /^(0|[1-9]\d*)\.(0|[1-9]\d*)\.(0|[1-9]\d*)(?:-((?:0|[1-9]\d*|\d*[A-Za-z-][0-9A-Za-z-]*)(?:\.(?:0|[1-9]\d*|\d*[A-Za-z-][0-9A-Za-z-]*))*))?$/

export async function distributionVersion(sourceRoot, tag) {
  const { version } = await json(join(sourceRoot, 'package.json'))
  const manifest = await json(join(sourceRoot, 'plugins/tandemfolio/.codex-plugin/plugin.json'))
  if (!semver.test(version) || manifest.version.split('+')[0] !== version) {
    throw new Error(
      'Root and plugin versions must have the same SemVer prefix (without build metadata).',
    )
  }
  if (tag !== undefined && tag !== `v${version}`)
    throw new Error(`Expected tag v${version}, got ${tag}.`)
  return version
}

// Staging is separate so tests can exercise real archive layouts without issuing a release.
// Only the CLI below produces release archives, after the source-current evidence gate.
export async function stageDistribution(sourceRoot, bundleRoot, version) {
  await mkdir(bundleRoot, { recursive: false })
  const sourcePlugin = join(sourceRoot, 'plugins/tandemfolio')
  await filesUnder(sourcePlugin) // Reject links before copying from the build tree.
  const pluginRoot = join(bundleRoot, 'plugins/tandemfolio')
  await mkdir(pluginRoot, { recursive: true })
  for (const path of ['.codex-plugin', '.mcp.json', 'skills', 'assets', 'dist']) {
    await cp(join(sourcePlugin, path), join(pluginRoot, path), { recursive: true })
  }
  const manifestPath = join(pluginRoot, '.codex-plugin/plugin.json')
  const manifest = await json(manifestPath)
  // Release metadata is independent of the developer's local Codex cachebuster.
  await writeJson(manifestPath, { ...manifest, version })
  await writeJson(join(pluginRoot, 'package.json'), {
    name: 'tandemfolio',
    version,
    private: true,
    type: 'module',
  })
  for (const path of ['LICENSE', 'NOTICE', 'LICENSE-UNICODE.txt']) {
    await cp(join(sourceRoot, path), join(bundleRoot, path))
    await cp(join(sourceRoot, path), join(pluginRoot, path))
  }
  for (const path of [
    'docs/project-facts.md',
    'docs/migration/provenance.md',
    'docs/migration/ledger.md',
    'apps/docs/src/renderer/fonts/README.md',
  ]) {
    await mkdir(dirname(join(bundleRoot, path)), { recursive: true })
    await cp(join(sourceRoot, path), join(bundleRoot, path))
  }
  for (const path of ['install.sh', 'install.ps1', 'verify.mjs']) {
    await cp(join(sourceRoot, 'distribution', path), join(bundleRoot, path))
  }
  await cp(join(sourceRoot, 'docs/distribution.md'), join(bundleRoot, 'README.md'))
  await mkdir(join(bundleRoot, '.agents/plugins'), { recursive: true })
  await writeJson(join(bundleRoot, '.agents/plugins/marketplace.json'), {
    name: 'tandemfolio-releases',
    interface: { displayName: 'TandemFolio Releases' },
    plugins: [
      {
        name: 'tandemfolio',
        source: { source: 'local', path: './plugins/tandemfolio' },
        policy: { installation: 'AVAILABLE', authentication: 'ON_INSTALL' },
        category: 'Productivity',
      },
    ],
  })
  const checksums = {}
  for (const path of await filesUnder(bundleRoot)) {
    checksums[path] = createHash('sha256')
      .update(await readFile(join(bundleRoot, path)))
      .digest('hex')
  }
  await writeJson(join(bundleRoot, 'checksums.json'), checksums)
  await verifyBundle(bundleRoot)
}

export function archiveDistribution(workDir, bundleName) {
  execFileSync('tar', ['-czf', `${bundleName}.tar.gz`, bundleName], {
    cwd: workDir,
    stdio: 'inherit',
  })
  execFileSync('zip', ['-q', '-r', `${bundleName}.zip`, bundleName], {
    cwd: workDir,
    stdio: 'inherit',
  })
}

async function main() {
  if (process.argv.slice(2).some((arg) => !arg.startsWith('--tag=')))
    throw new Error('Usage: npm run distribution:archive -- [--tag=vX.Y.Z]')
  const tag = process.argv
    .slice(2)
    .find((arg) => arg.startsWith('--tag='))
    ?.slice(6)
  const version = await distributionVersion(root, tag)
  // Read-only validation: no stale ready flag or manual packaging can bypass ADR 0005.
  execFileSync(
    process.execPath,
    ['--import', 'tsx', 'tools/release-gate.ts', '--evidence', 'release/release-evidence.json'],
    { cwd: root, stdio: 'inherit' },
  )
  const outputRoot = join(root, 'out/releases')
  await mkdir(outputRoot, { recursive: true })
  const workDir = await mkdtemp(join(outputRoot, `v${version}-`))
  const bundleName = `tandemfolio-${version}`
  await stageDistribution(root, join(workDir, bundleName), version)
  archiveDistribution(workDir, bundleName)
  const lines = []
  for (const extension of ['zip', 'tar.gz']) {
    const name = `${bundleName}.${extension}`
    lines.push(
      `${createHash('sha256')
        .update(await readFile(join(workDir, name)))
        .digest('hex')}  ${name}`,
    )
  }
  await writeFile(join(workDir, 'SHA256SUMS'), `${lines.join('\n')}\n`)
  if (process.env.GITHUB_OUTPUT) {
    const { appendFile } = await import('node:fs/promises')
    await appendFile(process.env.GITHUB_OUTPUT, `directory=${workDir}\nversion=${version}\n`)
  }
  console.log(`Release archives: ${workDir}`)
}

if (process.argv[1] && realpathSync(resolve(process.argv[1])) === fileURLToPath(import.meta.url)) {
  main().catch((error) => {
    console.error(error.message)
    process.exitCode = 1
  })
}
