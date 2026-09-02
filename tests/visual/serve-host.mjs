import { build } from 'esbuild'
import { createServer } from 'node:http'
import { readFile } from 'node:fs/promises'
import { basename, join } from 'node:path'

const root = process.cwd()
const pluginRoot = join(root, 'plugins/tandemfolio')
const editorAssets = {
  docx: 'assets/editor/index.html',
  markdown: 'assets/editors/markdown/index.html',
  xlsx: 'assets/editors/xlsx/index.html',
  pptx: 'assets/editors/pptx/index.html',
  pdf: 'assets/editors/pdf/index.html',
}
const hostBundle = await build({
  entryPoints: [join(root, 'tests/visual/host-harness.ts')],
  bundle: true,
  define: { 'process.env.NODE_ENV': '"production"' },
  format: 'iife',
  platform: 'browser',
  target: 'chrome140',
  write: false,
})
const hostScript = hostBundle.outputFiles[0].contents

const hostHtml = `<!doctype html>
<html lang="zh-CN">
  <head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>Codex MCP visual host</title>
    <style>
      :root { color-scheme: light; font-family: Arial, sans-serif; background: #f5f5f5; }
      * { box-sizing: border-box; }
      html, body { margin: 0; min-height: 100%; }
      body { padding: 16px; }
      #editor-frame { display: block; border: 0; background: white; box-shadow: 0 0 0 1px #d8d8d8; }
    </style>
  </head>
  <body>
    <iframe id="editor-frame" title="TandemFolio editor"></iframe>
    <script src="/host.js"></script>
  </body>
</html>`

function send(response, status, contentType, body) {
  response.writeHead(status, { 'content-type': contentType, 'cache-control': 'no-store' })
  response.end(body)
}

const server = createServer(async (request, response) => {
  const url = new URL(request.url ?? '/', 'http://127.0.0.1:4178')
  try {
    if (url.pathname === '/health') return send(response, 200, 'text/plain', 'ok')
    if (url.pathname === '/') return send(response, 200, 'text/html; charset=utf-8', hostHtml)
    if (url.pathname === '/host.js') {
      return send(response, 200, 'text/javascript; charset=utf-8', hostScript)
    }
    if (url.pathname === '/editor.html') {
      const format = url.searchParams.get('format') ?? 'docx'
      const editorAsset = editorAssets[format]
      if (!editorAsset) return send(response, 404, 'text/plain', 'unknown editor format')
      const html = await readFile(join(pluginRoot, editorAsset))
      return send(response, 200, 'text/html; charset=utf-8', html)
    }
    if (url.pathname === '/font') {
      const requestedName = url.searchParams.get('name') ?? ''
      const fileName = basename(requestedName)
      if (!fileName || fileName !== requestedName)
        return send(response, 400, 'text/plain', 'bad font')
      const bytes = await readFile(join(pluginRoot, 'assets/fonts', fileName))
      const offset = Math.max(0, Number(url.searchParams.get('offset') ?? 0))
      const length = Math.max(0, Number(url.searchParams.get('length') ?? 262_144))
      const chunk = bytes.subarray(offset, Math.min(offset + length, bytes.length))
      const nextOffset = offset + chunk.length
      return send(
        response,
        200,
        'application/json',
        JSON.stringify({
          ok: true,
          data: chunk.toString('base64'),
          nextOffset,
          eof: nextOffset >= bytes.length,
        }),
      )
    }
    return send(response, 404, 'text/plain', 'not found')
  } catch (error) {
    return send(response, 500, 'text/plain', error instanceof Error ? error.message : String(error))
  }
})

server.listen(4178, '127.0.0.1')

for (const signal of ['SIGINT', 'SIGTERM']) {
  process.on(signal, () => server.close(() => process.exit(0)))
}
