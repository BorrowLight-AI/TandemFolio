import { createHash } from 'node:crypto'
import { existsSync, readFileSync, readdirSync, statSync } from 'node:fs'
import { join, relative } from 'node:path'

const SOURCE_ROOTS = [
  'package.json',
  'package-lock.json',
  'upstream.config.json',
  'playwright.host.config.ts',
  'apps',
  'packages',
  'tests/visual',
  'tools/generate-operation-manifest.ts',
  'tools/operation-catalogs.ts',
  'tools/package-plugin.mjs',
  'tools/capture-release-baseline.ts',
  'tools/release-gate.ts',
  'tools/release-gate',
  'release/upstream-visual-manifest.json',
  'plugins/tandemfolio/.codex-plugin/plugin.json',
  'plugins/tandemfolio/assets/editor/index.html',
  'plugins/tandemfolio/assets/fonts',
] as const

function releaseRelevant(path: string): boolean {
  if (path.includes('/node_modules/') || path.includes('/dist/')) return false
  if (/^\/(?:apps|packages)\/[^/]+\/tests\//.test(path)) return false
  if (path.endsWith('/src/generated/release-readiness.json')) return false
  if (/\.(?:ts|tsx|js|mjs|cjs|json|html|css|wasm|ttf|otf|woff2)$/.test(path)) return true
  return !path.includes('.')
}

function filesUnder(root: string, input: string): string[] {
  const absolute = join(root, input)
  if (!existsSync(absolute)) return []
  if (!statSync(absolute).isDirectory()) return [absolute]
  const files: string[] = []
  for (const entry of readdirSync(absolute, { withFileTypes: true })) {
    const child = join(absolute, entry.name)
    const normalized = `/${relative(root, child).replaceAll('\\', '/')}`
    if (!releaseRelevant(normalized)) continue
    if (entry.isDirectory()) files.push(...filesUnder(root, relative(root, child)))
    else if (entry.isFile()) files.push(child)
  }
  return files
}

export function computeReleaseSourceFingerprint(root: string): string {
  const files = SOURCE_ROOTS.flatMap((input) => filesUnder(root, input)).sort()
  const hash = createHash('sha256')
  for (const file of files) {
    hash.update(relative(root, file).replaceAll('\\', '/'))
    hash.update('\0')
    hash.update(readFileSync(file))
    hash.update('\0')
  }
  return hash.digest('hex')
}
