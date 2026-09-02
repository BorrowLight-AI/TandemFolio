import react from '@vitejs/plugin-react'
import { defineConfig } from 'vite'
import { readFileSync } from 'node:fs'
import { gzipSync } from 'node:zlib'

const PDFIUM_WASM_GZIP = '\0tandemfolio:pdfium-wasm-gzip'

function compressedPdfiumWasm() {
  return {
    name: 'tandemfolio-compressed-pdfium-wasm',
    enforce: 'pre' as const,
    resolveId(id: string) {
      return id === 'virtual:pdfium-wasm-gzip' ? PDFIUM_WASM_GZIP : null
    },
    load(id: string) {
      if (id !== PDFIUM_WASM_GZIP) return null
      const wasm = readFileSync(
        new URL('../../node_modules/@embedpdf/pdfium/dist/pdfium.wasm', import.meta.url),
      )
      const encoded = gzipSync(wasm, { level: 9 }).toString('base64')
      return `export default ${JSON.stringify(encoded)}`
    },
    transform(code: string, id: string) {
      if (!id.includes('@embedpdf/pdfium/dist/index.browser.js')) return null
      return code.replace("new URL('pdfium.wasm', import.meta.url).href", "'pdfium.wasm'")
    },
  }
}

export default defineConfig({
  root: 'src/renderer',
  plugins: [react(), compressedPdfiumWasm()],
  base: './',
  build: {
    outDir: new URL('./dist/renderer', import.meta.url).pathname,
    emptyOutDir: true,
    assetsInlineLimit: Number.MAX_SAFE_INTEGER,
    rollupOptions: { output: { inlineDynamicImports: true } },
  },
  server: { port: Number(process.env.PDF_DEV_PORT) || 5176, strictPort: true },
})
