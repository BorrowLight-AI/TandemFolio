/** Load the Node-side PDF.js evidence API without relying on its optional native canvas package. */
export async function loadPdfJsForNodeEvidence() {
  if (!globalThis.DOMMatrix) {
    Object.defineProperty(globalThis, 'DOMMatrix', {
      configurable: true,
      writable: true,
      value: class NodeEvidenceDOMMatrix {},
    })
  }
  return import('pdfjs-dist/legacy/build/pdf.mjs')
}
