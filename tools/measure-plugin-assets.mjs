import { readFile } from 'node:fs/promises'
import { gzipSync } from 'node:zlib'

const resources = [
  // Registry completion grows the retained renderer deliberately; this
  // remains a regression ceiling, never permission to delete capabilities.
  // The 2026-09-05 source-current DOCX port adds the upstream pagination,
  // layout, fonts and native dialogs; retain modest headroom over that audited build.
  ['DOCX', 'plugins/tandemfolio/assets/editor/index.html', 3_650_000],
  ['Markdown', 'plugins/tandemfolio/assets/editors/markdown/index.html', 2_750_000],
  // Codex measures the returned MCP App HTML as a UTF-8 Blob and rejects it
  // above 10,000,000 bytes, before mounting the renderer iframe.
  ['XLSX', 'plugins/tandemfolio/assets/editors/xlsx/index.html', 10_000_000],
  // The complete pinned non-AI community renderer replaces the rejected
  // narrow scaffold. Keep a regression ceiling without deleting capabilities.
  ['PPTX', 'plugins/tandemfolio/assets/editors/pptx/index.html', 4_000_000],
  // Browser-safe content-stream text/image parity carries the pinned PDFium
  // implementation and compressed WASM inside the self-contained editor.
  ['PDF', 'plugins/tandemfolio/assets/editors/pdf/index.html', 7_000_000],
]

let failed = false
console.log('| Format | Raw bytes | Gzip bytes | Budget |')
console.log('| --- | ---: | ---: | ---: |')
for (const [format, path, budget] of resources) {
  const bytes = await readFile(path)
  console.log(`| ${format} | ${bytes.length} | ${gzipSync(bytes).length} | ${budget} |`)
  if (bytes.length > budget) failed = true
}
if (failed) process.exitCode = 1
