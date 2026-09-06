import { existsSync, readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { gzipSync } from 'node:zlib'

const wasmCandidates = [
  resolve(process.cwd(), 'node_modules/@embedpdf/pdfium/dist/pdfium.wasm'),
  resolve(process.cwd(), '../../node_modules/@embedpdf/pdfium/dist/pdfium.wasm'),
]
const wasmPath = wasmCandidates.find(existsSync)
if (!wasmPath) throw new Error('Could not locate the bundled PDFium WASM payload.')

/** Real compressed PDFium payload so native text/image save tests cross the WASM boundary. */
export default gzipSync(
  readFileSync(wasmPath),
  { level: 1 },
).toString('base64')
