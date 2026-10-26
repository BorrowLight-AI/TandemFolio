import { readFile } from 'node:fs/promises'
import { expect, test, type Locator } from '@playwright/test'
import type { ShapeRenderNode } from '../../packages/pptx-render/src/render-tree'

const BLUE =
  'iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAIAAACQkWg2AAAAFUlEQVR4nGNgyF1FGhrVMKph+GoAAKoGFxBobgn7AAAAAElFTkSuQmCC'

async function bluePixels(canvas: Locator): Promise<number> {
  return canvas.evaluate((element) => {
    const target = element as HTMLCanvasElement
    const pixels = target.getContext('2d')!.getImageData(0, 0, target.width, target.height).data
    let count = 0
    for (let i = 0; i < pixels.length; i += 4) {
      if (
        pixels[i] === 0 &&
        pixels[i + 1] === 109 &&
        pixels[i + 2] === 170 &&
        pixels[i + 3] === 255
      )
        count++
    }
    return count
  })
}

test('bold CJK layout advances match the glyphs painted by the browser', async ({ page }) => {
  await page.goto('/?format=pptx&width=1332&height=900')
  await page.waitForFunction(
    () => window.__codexVisualHost?.initialized && window.__codexVisualHost.polls > 0,
  )
  await expect(
    page.frameLocator('#editor-frame').locator('.stage-wrap .konvajs-content canvas').first(),
  ).toBeVisible()
  await page.evaluate(() =>
    window.__codexVisualHost.enqueueCommand({
      commandId: 'insert-cjk-title',
      baseRevision: 0,
      operation: 'pptx.object.add',
      arguments: {
        slideIndex: 0,
        kind: 'textbox',
        xEmu: 914400,
        yEmu: 914400,
        widthEmu: 10000000,
        heightEmu: 1828800,
        paragraphs: [
          {
            align: 'center',
            runs: [
              // 30 pt renders at 40 px. Linux system CJK hinting can round a
              // fractional 42⅔ px face to 43 px per glyph; that is unrelated
              // to the regression where layout expanded every bold glyph 4%.
              { text: '物业企业应收账款管理研究', fontFamily: '宋体', fontSize: 30, bold: true },
            ],
          },
        ],
      },
    }),
  )
  await expect
    .poll(() =>
      page.evaluate(() =>
        window.__codexVisualHost.acknowledgements.find(
          (entry) => entry.commandId === 'insert-cjk-title',
        ),
      ),
    )
    .toMatchObject({ ok: true })
  const widths = await page
    .frameLocator('#editor-frame')
    .locator('body')
    .evaluate(async () => {
      const slide = (await window.slidesApi.getRenderSlides())![0]!
      const title = slide.nodes.find(
        (node): node is ShapeRenderNode =>
          (node.type === 'shape' || node.type === 'text') &&
          node.text?.lines.flatMap((line) => line.runs.map((run) => run.text)).join('') ===
            '物业企业应收账款管理研究',
      )!
      const runs = title.text!.lines.flatMap((line) => line.runs)
      const context = document.createElement('canvas').getContext('2d')!
      context.font = `bold ${runs[0]!.fontSizePx}px SimSun, "Songti SC", serif`
      return {
        fontSize: runs[0]!.fontSizePx,
        family: runs[0]!.fontFamily,
        advances: runs.map((run) => ({ width: run.widthPx, size: run.fontSizePx, bold: run.bold })),
        layout: runs.reduce((width, run) => width + run.widthPx, 0),
        painted: context.measureText(runs.map((run) => run.text).join('')).width,
      }
    })
  for (const advance of widths.advances) {
    expect(advance.bold).toBe(true)
    expect(advance.width).toBeCloseTo(advance.size, 8)
  }
  // Keep the strict painted-width check without fractional-size OS hinting.
  expect(Math.abs(widths.layout - widths.painted), JSON.stringify(widths)).toBeLessThan(0.5)
})

test('PPTX package pictures paint on the stage and thumbnail after Agent insert and saved-file reopen', async ({
  page,
}) => {
  await page.goto('/?format=pptx&width=1332&height=900')
  await page.waitForFunction(
    () => window.__codexVisualHost?.initialized && window.__codexVisualHost.polls > 0,
  )
  const editor = page.frameLocator('#editor-frame')
  const stage = editor.locator('.stage-wrap .konvajs-content canvas').first()
  const thumb = editor.locator('.thumb .konvajs-content canvas').first()
  await expect(stage).toBeVisible()

  await page.evaluate(
    (data) =>
      window.__codexVisualHost.enqueueCommand({
        commandId: 'insert-package-picture',
        baseRevision: 0,
        operation: 'pptx.image.add_bytes',
        arguments: {
          slideIndex: 0,
          data,
          extension: 'png',
          xEmu: 914400,
          yEmu: 914400,
          widthEmu: 1828800,
          heightEmu: 914400,
        },
      }),
    BLUE,
  )
  await page.waitForFunction(() =>
    window.__codexVisualHost.acknowledgements.some(
      (entry) => entry.commandId === 'insert-package-picture' && entry.ok === true,
    ),
  )
  await expect.poll(() => bluePixels(stage)).toBeGreaterThan(1000)
  await expect.poll(() => bluePixels(thumb)).toBeGreaterThan(5)

  const downloadPromise = page.waitForEvent('download')
  await page.evaluate(() =>
    window.__codexVisualHost.enqueueCommand({
      commandId: 'save-package-picture',
      baseRevision: 1,
      operation: 'pptx.document.save',
      arguments: {},
    }),
  )
  await page.waitForFunction(() =>
    window.__codexVisualHost.acknowledgements.some(
      (entry) => entry.commandId === 'save-package-picture' && entry.ok === true,
    ),
  )
  const savedPath = await (await downloadPromise).path()
  await editor.getByRole('button', { name: '文件', exact: true }).click()
  await editor.getByRole('button', { name: /打开/ }).click()
  await editor.locator('input[type="file"][accept=".pptx"]').setInputFiles({
    name: 'package-picture.pptx',
    mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    buffer: await readFile(savedPath!),
  })
  await expect(editor.locator('.status-bar')).toContainText('package-picture.pptx')
  await expect.poll(() => bluePixels(stage)).toBeGreaterThan(1000)
  await expect.poll(() => bluePixels(thumb)).toBeGreaterThan(5)
  expect(await page.evaluate(() => window.__codexVisualHost.editorLoads)).toBe(1)
})
