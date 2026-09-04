import { spawnSync } from 'node:child_process'
import { dirname, join, resolve } from 'node:path'
import { fileURLToPath, pathToFileURL } from 'node:url'
import { expect, it } from 'vitest'

it('loads PDF.js evidence helpers when the optional native canvas dependency is unavailable', () => {
  const repositoryRoot = resolve(dirname(fileURLToPath(import.meta.url)), '../../..')
  const bootstrapUrl = pathToFileURL(
    join(repositoryRoot, 'tests/visual/pdfjs-node-bootstrap.ts'),
  ).href
  const script = `
    import Module from 'node:module'
    const originalLoad = Module._load
    Module._load = function (request, parent, isMain) {
      if (request === '@napi-rs/canvas') {
        const error = new Error('simulated optional dependency omission')
        error.code = 'MODULE_NOT_FOUND'
        throw error
      }
      return originalLoad.apply(this, arguments)
    }
    const { loadPdfJsForNodeEvidence } = await import(${JSON.stringify(bootstrapUrl)})
    const { getDocument, OPS } = await loadPdfJsForNodeEvidence()
    if (typeof getDocument !== 'function' || !OPS) {
      throw new Error('PDF.js evidence API did not load')
    }
    const { PDFDocument, StandardFonts } = await import('pdf-lib')
    const fixture = await PDFDocument.create()
    const font = await fixture.embedFont(StandardFonts.Helvetica)
    const sourcePage = fixture.addPage([420, 300])
    sourcePage.drawText('CI PDF evidence', { x: 30, y: 250, size: 18, font })
    const loadingTask = getDocument({ data: new Uint8Array(await fixture.save()) })
    try {
      const document = await loadingTask.promise
      const page = await document.getPage(1)
      const content = await page.getTextContent()
      const text = content.items.map((item) => 'str' in item ? item.str : '').join(' ')
      const operators = await page.getOperatorList()
      if (text !== 'CI PDF evidence' || operators.fnArray.length === 0) {
        throw new Error('PDF.js evidence parsing did not complete')
      }
    } finally {
      await loadingTask.destroy()
    }
  `
  const result = spawnSync(
    process.execPath,
    ['--import', 'tsx', '--input-type=module', '--eval', script],
    { cwd: repositoryRoot, encoding: 'utf8' },
  )

  expect(result.status, result.stderr || result.stdout).toBe(0)
})
