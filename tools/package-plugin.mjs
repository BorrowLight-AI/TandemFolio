import { copyFile, cp, mkdir, readFile, readdir, rm, writeFile } from 'node:fs/promises'
import { basename, dirname, join, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { gzipSync } from 'node:zlib'
import { init, parse } from 'es-module-lexer'

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..')
const editorDist = join(root, 'apps/docs/dist/renderer')
const markdownEditorDist = join(root, 'apps/markdown/dist/renderer')
const xlsxEditorDist = join(root, 'apps/sheets/dist/renderer')
const pptxEditorDist = join(root, 'apps/slides/dist/renderer')
const pdfEditorDist = join(root, 'apps/pdf/dist/renderer')
const serverBundle = join(root, 'apps/mcp-server/dist/server.js')
const pluginRoot = join(root, 'plugins/tandemfolio')
const pluginDist = join(pluginRoot, 'dist')
const pluginEditor = join(pluginRoot, 'assets/editor')
const pluginEditors = join(pluginRoot, 'assets/editors')
const pluginFonts = join(pluginRoot, 'assets/fonts')

async function inlineEditor(dist) {
  let html = await readFile(join(dist, 'index.html'), 'utf8')

  for (const match of [...html.matchAll(/<link[^>]+href="([^"]+\.css)"[^>]*>/g)]) {
    const cssPath = resolve(dist, match[1])
    const css = await readFile(cssPath, 'utf8')
    html = html.replace(match[0], () => `<style>${css}</style>`)
  }

  for (const match of [...html.matchAll(/<script[^>]+src="([^"]+\.js)"[^>]*><\/script>/g)]) {
    const jsPath = resolve(dist, match[1])
    const js = (await readFile(jsPath, 'utf8')).replaceAll('</script', '<\\/script')
    html = html.replace(match[0], () => `<script type="module">${js}</script>`)
  }

  return html
}

function replaceModuleSpecifiers(source, knownModules) {
  const [imports] = parse(source)
  const replacements = []
  for (const imported of imports) {
    if (!imported.n?.startsWith('./') || !imported.n.endsWith('.js')) continue
    const moduleId = basename(imported.n)
    if (!knownModules.has(moduleId)) {
      throw new Error(`Unknown deferred XLSX module: ${imported.n}`)
    }
    if (imported.d >= 0) {
      replacements.push({
        start: imported.ss,
        end: imported.se,
        value: `globalThis.__genofficeImport(${JSON.stringify(moduleId)})`,
      })
    } else {
      replacements.push({
        start: imported.s,
        end: imported.e,
        value: `genoffice-static:${moduleId}`,
      })
    }
  }
  let transformed = source
  for (const replacement of replacements.sort((left, right) => right.start - left.start)) {
    transformed =
      transformed.slice(0, replacement.start) +
      replacement.value +
      transformed.slice(replacement.end)
  }
  return transformed
}

const moduleVaultBootstrap = `(function () {
  const dependencyNode = document.querySelector('script[data-tandemfolio-module-dependencies]');
  const dependencies = JSON.parse(dependencyNode.textContent);
  dependencyNode.remove();
  const payloadNodes = document.querySelectorAll('script[type="application/x-tandemfolio-module"]');
  const payloads = new Map();
  let entryModule = '';
  for (const node of payloadNodes) {
    const moduleId = node.dataset.module;
    if (!moduleId) continue;
    payloads.set(moduleId, {
      source: node.textContent || '',
      encoding: node.dataset.encoding || 'gzip',
    });
    if (node.dataset.entry === 'true') entryModule = moduleId;
    node.remove();
  }
  const urls = new Map();
  const pending = new Map();
  async function inflate(encoded) {
    const binary = atob(encoded);
    const bytes = new Uint8Array(binary.length);
    for (let index = 0; index < binary.length; index += 1) bytes[index] = binary.charCodeAt(index);
    const stream = new Blob([bytes]).stream().pipeThrough(new DecompressionStream('gzip'));
    return new Response(stream);
  }
  async function ensureModule(moduleId) {
    const existing = urls.get(moduleId);
    if (existing) return existing;
    const inflight = pending.get(moduleId);
    if (inflight) return inflight;
    const task = (async function () {
      const payload = payloads.get(moduleId);
      if (!payload) throw new Error('Missing embedded module: ' + moduleId);
      const dependencyIds = dependencies[moduleId];
      if (!Array.isArray(dependencyIds)) throw new Error('Missing module dependencies: ' + moduleId);
      let source;
      if (payload.encoding === 'identity') {
        source = payload.source;
      } else if (dependencyIds.length === 0) {
        // Keep dependency-free modules as UTF-8 bytes: avoid decoding a large
        // entry to a JS string only to re-encode it when constructing its Blob.
        const response = await inflate(payload.source.trim());
        source = await response.blob();
      } else {
        const response = await inflate(payload.source.trim());
        source = await response.text();
        for (const dependencyId of dependencyIds) {
          const dependencyUrl = await ensureModule(dependencyId);
          source = source.replaceAll('genoffice-static:' + dependencyId, dependencyUrl);
        }
      }
      const url = URL.createObjectURL(new Blob([source], { type: 'text/javascript' }));
      urls.set(moduleId, url);
      return url;
    })();
    pending.set(moduleId, task);
    try {
      return await task;
    } finally {
      pending.delete(moduleId);
    }
  }
  globalThis.__genofficeImport = async function (moduleId) {
    return import(await ensureModule(moduleId));
  };
  globalThis.__genofficeModuleVaultReady = globalThis.__genofficeImport(entryModule);
  globalThis.__genofficeModuleVaultReady.catch(function (error) { console.error(error); });
})();`

async function inlineDeferredEditor(dist) {
  await init
  let html = await readFile(join(dist, 'index.html'), 'utf8')

  for (const match of [...html.matchAll(/<link[^>]+href="([^"]+\.css)"[^>]*>/g)]) {
    const cssPath = resolve(dist, match[1])
    const css = await readFile(cssPath, 'utf8')
    html = html.replace(match[0], () => `<style>${css}</style>`)
  }
  html = html.replace(/<link[^>]+rel="modulepreload"[^>]*>/g, '')

  const entryTag = html.match(/<script[^>]+src="([^"]+\.js)"[^>]*><\/script>/)
  if (!entryTag) throw new Error('Missing XLSX entry module.')
  const entryPath = resolve(dist, entryTag[1])
  const entryModule = basename(entryPath)
  const assetDir = dirname(entryPath)
  const moduleIds = (await readdir(assetDir)).filter((fileName) => fileName.endsWith('.js'))
  const knownModules = new Set(moduleIds)
  const payloads = []
  const dependencies = {}
  let entryPayload = ''
  for (const moduleId of moduleIds.sort()) {
    const source = replaceModuleSpecifiers(
      await readFile(join(assetDir, moduleId), 'utf8'),
      knownModules,
    )
    dependencies[moduleId] = [
      ...new Set(
        parse(source)[0]
          .filter((entry) => entry.d === -1 && entry.n?.startsWith('genoffice-static:'))
          .map((entry) => entry.n.slice('genoffice-static:'.length)),
      ),
    ]
    if (moduleId === entryModule) {
      if (dependencies[moduleId].length > 0) {
        throw new Error('The directly executable XLSX entry cannot have static chunk imports.')
      }
      entryPayload = `<script type="application/x-tandemfolio-module" data-module="${moduleId}" data-entry="true" data-encoding="identity">${source.replaceAll('</script', '<\\/script')}</script>`
      continue
    }
    const encoded = gzipSync(source, { level: 9 }).toString('base64')
    payloads.push(
      `<script type="application/x-tandemfolio-module" data-module="${moduleId}">${encoded}</script>`,
    )
  }
  if (!entryPayload) throw new Error('The XLSX entry module was not packaged.')
  html = html.replace(
    entryTag[0],
    () =>
      `${payloads.join('')}${entryPayload}<script type="application/json" data-tandemfolio-module-dependencies>${JSON.stringify(dependencies)}</script><script data-tandemfolio-module-bootstrap>${moduleVaultBootstrap}</script>`,
  )
  html = html.replace(
    "script-src 'self' 'unsafe-inline'",
    "script-src 'self' 'unsafe-inline' blob:",
  )
  return html
}

await rm(pluginDist, { recursive: true, force: true })
await rm(pluginEditor, { recursive: true, force: true })
await rm(pluginEditors, { recursive: true, force: true })
await rm(pluginFonts, { recursive: true, force: true })
await mkdir(pluginDist, { recursive: true })
await mkdir(pluginEditor, { recursive: true })
await mkdir(join(pluginEditors, 'markdown'), { recursive: true })
await mkdir(join(pluginEditors, 'xlsx'), { recursive: true })
await mkdir(join(pluginEditors, 'pptx'), { recursive: true })
await mkdir(join(pluginEditors, 'pdf'), { recursive: true })
await copyFile(serverBundle, join(pluginDist, 'server.js'))
await writeFile(join(pluginEditor, 'index.html'), await inlineEditor(editorDist))
await writeFile(join(pluginEditors, 'markdown/index.html'), await inlineEditor(markdownEditorDist))
await writeFile(join(pluginEditors, 'xlsx/index.html'), await inlineDeferredEditor(xlsxEditorDist))
await writeFile(join(pluginEditors, 'pptx/index.html'), await inlineEditor(pptxEditorDist))
await writeFile(join(pluginEditors, 'pdf/index.html'), await inlineEditor(pdfEditorDist))
await cp(join(root, 'apps/docs/src/renderer/fonts'), pluginFonts, {
  recursive: true,
  filter: (source) => !source.endsWith('fonts.css'),
})
await copyFile(join(root, 'LICENSE'), join(pluginRoot, 'LICENSE'))
await copyFile(join(root, 'NOTICE'), join(pluginRoot, 'NOTICE'))
