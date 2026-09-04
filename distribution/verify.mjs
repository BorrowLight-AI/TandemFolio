import { createHash } from 'node:crypto'
import { realpathSync } from 'node:fs'
import { readFile, readdir, lstat } from 'node:fs/promises'
import { dirname, join, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'

export async function filesUnder(root, prefix = '') {
  const files = []
  for (const name of (await readdir(join(root, prefix))).sort()) {
    // Finder/Explorer may create these while the user extracts or opens the bundle.
    if (['.DS_Store', 'Thumbs.db', 'desktop.ini'].includes(name)) continue
    const path = prefix ? `${prefix}/${name}` : name
    const stat = await lstat(join(root, path))
    if (stat.isSymbolicLink()) throw new Error(`Symlinks are not allowed: ${path}`)
    if (stat.isDirectory()) files.push(...(await filesUnder(root, path)))
    else if (stat.isFile()) files.push(path)
    else throw new Error(`Unsupported file: ${path}`)
  }
  return files
}

export async function verifyBundle(root) {
  const [major, minor] = process.versions.node.split('.').map(Number)
  if (major < 22 || (major === 22 && minor < 12)) throw new Error('Node.js 22.12+ is required.')
  const expected = JSON.parse(await readFile(join(root, 'checksums.json'), 'utf8'))
  const files = (await filesUnder(root)).filter((path) => path !== 'checksums.json')
  if (JSON.stringify(files.sort()) !== JSON.stringify(Object.keys(expected).sort())) {
    throw new Error('Bundle file list differs from checksums.json. Extract a fresh download.')
  }
  for (const path of files) {
    const actual = createHash('sha256')
      .update(await readFile(join(root, path)))
      .digest('hex')
    if (actual !== expected[path]) throw new Error(`Checksum mismatch: ${path}`)
  }
  const catalog = JSON.parse(await readFile(join(root, '.agents/plugins/marketplace.json'), 'utf8'))
  const manifest = JSON.parse(
    await readFile(join(root, 'plugins/tandemfolio/.codex-plugin/plugin.json'), 'utf8'),
  )
  if (
    catalog.name !== 'tandemfolio-releases' ||
    catalog.plugins.length !== 1 ||
    catalog.plugins[0].name !== 'tandemfolio' ||
    catalog.plugins[0].source.source !== 'local' ||
    catalog.plugins[0].source.path !== './plugins/tandemfolio' ||
    manifest.name !== 'tandemfolio'
  ) {
    throw new Error('Unexpected plugin or marketplace identity.')
  }
  return manifest.version
}

if (process.argv[1] && realpathSync(resolve(process.argv[1])) === fileURLToPath(import.meta.url)) {
  try {
    const version = await verifyBundle(dirname(fileURLToPath(import.meta.url)))
    console.log(`Verified TandemFolio ${version}; Node.js ${process.versions.node}.`)
  } catch (error) {
    console.error(error.message)
    process.exitCode = 1
  }
}
