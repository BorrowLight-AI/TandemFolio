import { readFile, stat } from 'node:fs/promises'
import { gunzipSync } from 'node:zlib'
import { describe, expect, it } from 'vitest'

const packagedEditor = new URL(
  '../../../plugins/tandemfolio/assets/editor/index.html',
  import.meta.url,
)
const packagedXlsxEditor = new URL(
  '../../../plugins/tandemfolio/assets/editors/xlsx/index.html',
  import.meta.url,
)
const packagedPptxEditor = new URL(
  '../../../plugins/tandemfolio/assets/editors/pptx/index.html',
  import.meta.url,
)
const packagedPdfEditor = new URL(
  '../../../plugins/tandemfolio/assets/editors/pdf/index.html',
  import.meta.url,
)
const packagedMarkdownEditor = new URL(
  '../../../plugins/tandemfolio/assets/editors/markdown/index.html',
  import.meta.url,
)
const packagedOperationReference = new URL(
  '../../../plugins/tandemfolio/skills/tandemfolio/references/operations.md',
  import.meta.url,
)
const packagedSkill = new URL(
  '../../../plugins/tandemfolio/skills/tandemfolio/SKILL.md',
  import.meta.url,
)
const generatedOperationManifest = new URL(
  '../src/generated/operation-manifest.json',
  import.meta.url,
)

describe('packaged editor UI', () => {
  it('keeps every initial format resource inside its explicit raw-byte budget', async () => {
    for (const [resource, budget] of [
      // Registry completion grows the retained renderer deliberately; this
      // remains a regression ceiling, never permission to delete capabilities.
      [packagedEditor, 3_500_000],
      [packagedMarkdownEditor, 2_500_000],
      // Full permitted pinned Univer renderer; this is a regression ceiling,
      // not permission to delete community capabilities for bundle size.
      [packagedXlsxEditor, 21_000_000],
      [packagedPptxEditor, 4_000_000],
      // Browser PDFium restores retained searchable text/image mutation. Its
      // WASM is gzip-compressed in the self-contained resource; fonts stay lazy.
      [packagedPdfEditor, 7_000_000],
    ] as const) {
      await expect(stat(resource)).resolves.toMatchObject({ size: expect.any(Number) })
      expect((await stat(resource)).size).toBeLessThanOrEqual(budget)
    }
  })

  it('ships the styled document status bar controls', async () => {
    const html = await readFile(packagedEditor, 'utf8')

    for (const selector of [
      '.status-bar',
      '.status-left',
      '.status-item',
      '.status-wordcount',
      '.status-msg',
      '.status-right',
      '.zoom-btn',
      '.zoom-slider',
      '.zoom-value',
    ]) {
      expect(html, `missing packaged selector ${selector}`).toContain(selector)
    }
  })

  it('lists every Agent-visible generated operation in the installed Skill reference', async () => {
    const manifest = JSON.parse(await readFile(generatedOperationManifest, 'utf8')) as {
      operations: { id: string; visibility: 'agent' | 'internal' }[]
    }
    const reference = await readFile(packagedOperationReference, 'utf8')
    const missing = manifest.operations
      .filter((operation) => operation.visibility === 'agent')
      .filter((operation) => !reference.includes(`\`${operation.id}`))
      .map((operation) => operation.id)

    expect(missing).toEqual([])
  })

  it('reuses the exact mounted session for follow-up edits across every editor format', async () => {
    const skill = await readFile(packagedSkill, 'utf8')

    expect(skill).toMatch(/follow-up edit/i)
    expect(skill).toMatch(/connected:\s*true[\s\S]{0,240}(?:do not|never)[\s\S]{0,120}show/i)
    expect(skill).toMatch(/resume:\s*"exact"/)
    expect(skill).toMatch(/resume:\s*"none"/)
    expect(skill).toMatch(
      /DOCX[\s\S]{0,120}Markdown[\s\S]{0,120}XLSX[\s\S]{0,120}PPTX[\s\S]{0,120}PDF/,
    )
  })

  it('keeps the initial editor under 3.5 MB and packages fonts as lazy external assets', async () => {
    const html = await readFile(packagedEditor, 'utf8')
    const info = await stat(packagedEditor)
    const font = new URL(
      '../../../plugins/tandemfolio/assets/fonts/Carlito-Regular.ttf',
      import.meta.url,
    )

    expect(info.size).toBeLessThan(3_500_000)
    expect(html).not.toMatch(/data:font\/(?:ttf|woff2?);base64/i)
    await expect(stat(font)).resolves.toMatchObject({ size: 631_712 })
  })

  it('packages the XLSX live editor as a separate self-contained resource', async () => {
    const html = await readFile(packagedXlsxEditor, 'utf8')
    expect(html).toContain('TandemFolio · XLSX')
    expect(html).toContain('univer-container')
    expect(html).toContain('ribbon-tabs')
    expect(html).not.toMatch(/src="[^"]+\.js"/)
    expect(html).not.toMatch(/Genspark|ai-chat|ai-composer|ai-login|telemetry\.service/i)
  })

  it('keeps deferred XLSX modules inside one resource without compiling them at bootstrap', async () => {
    const html = await readFile(packagedXlsxEditor, 'utf8')
    const payloads = [
      ...html.matchAll(
        /<script type="application\/x-tandemfolio-module" data-module="([^"]+)"( data-entry="true")?>([^<]+)<\/script>/g,
      ),
    ]
    const entry = payloads.find((match) => match[2] !== undefined)

    expect(entry, 'missing embedded XLSX entry module').toBeDefined()
    expect(payloads.length).toBeGreaterThan(1)
    expect(gunzipSync(Buffer.from(entry![3], 'base64')).byteLength).toBeLessThanOrEqual(11_000_000)
    expect(html).toContain('data-tandemfolio-module-bootstrap')
    expect(html).not.toMatch(/(?:src|href)="[^"]+\.(?:js|css)"/)
  })

  it('packages the Markdown community renderer without product AI or Electron code', async () => {
    const html = await readFile(packagedMarkdownEditor, 'utf8')
    expect(html).toContain('TandemFolio · Markdown')
    expect(html).toContain('doc-editor')
    expect(html).toContain('Export DOCX')
    expect(html).not.toMatch(/src="[^"]+\.js"/)
    expect(html).not.toMatch(/Genspark|electron|markdownApi|ai-panel/i)
  })

  it('packages the retained PPTX community renderer without prohibited product code', async () => {
    const html = await readFile(packagedPptxEditor, 'utf8')
    expect(html).toContain('TandemFolio · PPTX')
    expect(html).toContain('stage-wrap')
    expect(html).toContain('ribbon-body')
    expect(html).toContain('退出全屏')
    expect(html).not.toMatch(/src="[^"]+\.js"/)
    expect(html).not.toMatch(
      /Genspark|ai-(?:chat|clarify|composer|dock|login|panel)|ipcRenderer|electronApi|telemetry\.service/i,
    )
  })

  it('packages the retained PDF community renderer without prohibited product code', async () => {
    const html = await readFile(packagedPdfEditor, 'utf8')
    expect(html).toContain('TandemFolio · PDF')
    expect(html).toContain('ribbon-tabs')
    expect(html).toContain('pdf-body')
    expect(html).toContain('pdf-page-content')
    expect(html).toContain('host-fullscreen')
    expect(html).not.toMatch(/src="[^"]+\.js"/)
    expect(html).not.toMatch(/Genspark|ai-panel|ipcRenderer|electronApi/i)
  })
})
