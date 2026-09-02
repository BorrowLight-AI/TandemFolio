import { readFile } from 'node:fs/promises'
import { gzipSync } from 'node:zlib'

const resources = [
  // Registry completion grows the retained renderer deliberately; this
  // remains a regression ceiling, never permission to delete capabilities.
  ['DOCX', 'plugins/tandemfolio/assets/editor/index.html', 3_500_000],
  ['Markdown', 'plugins/tandemfolio/assets/editors/markdown/index.html', 2_500_000],
  // The restored Univer renderer is intentionally much larger than the
  // rejected five-file scaffold. This is a regression ceiling, not a target
  // that permits deleting community capabilities.
  ['XLSX', 'plugins/tandemfolio/assets/editors/xlsx/index.html', 21_000_000],
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
