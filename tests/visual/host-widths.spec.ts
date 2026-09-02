import { expect, test, type Page } from '@playwright/test'
import { PDFDocument, StandardFonts } from 'pdf-lib'
import {
  activeHostVisualFormats,
  pendingHostVisualFormats,
  type ActiveHostVisualFormat,
} from './host-format-gates'

const NARROW = { width: 420, height: 900 }
const COMPACT = { width: 280, height: 900 }
const SPLIT = { width: 720, height: 900 }
const FULLSCREEN = { width: 1332, height: 1280 }

function editorUrl(
  format: ActiveHostVisualFormat,
  dimensions: { width: number; height: number },
  acceptFullscreen = false,
): string {
  const params = new URLSearchParams({
    format: format.id,
    width: String(dimensions.width),
    height: String(dimensions.height),
  })
  if (acceptFullscreen) {
    params.set('acceptFullscreen', 'true')
    params.set('fullscreenWidth', String(FULLSCREEN.width))
    params.set('fullscreenHeight', String(FULLSCREEN.height))
  }
  return `/?${params}`
}

async function waitForMountedHost(page: Page): Promise<void> {
  await page.waitForFunction(() => {
    const state = window.__codexVisualHost
    return (
      state?.initialized &&
      state.polls > 0 &&
      state.firstPollArguments?.sessionId === 'visual-session' &&
      state.firstPollArguments?.viewId === 'visual-view' &&
      state.sizeNotifications.length > 0 &&
      state.events.includes('display-request:fullscreen')
    )
  })
}

async function prepareFormatDocument(page: Page, format: ActiveHostVisualFormat): Promise<void> {
  if (format.id !== 'pdf') return
  const document = await PDFDocument.create()
  const font = await document.embedFont(StandardFonts.Helvetica)
  const pdfPage = document.addPage([420, 300])
  pdfPage.drawText('TandemFolio PDF host gate', { x: 30, y: 250, size: 18, font })
  const editor = page.frameLocator('#editor-frame')
  await editor.locator('input[type="file"][accept*=".pdf"]').setInputFiles({
    name: 'host-gate.pdf',
    mimeType: 'application/pdf',
    buffer: Buffer.from(await document.save()),
  })
  await expect(editor.locator(format.viewportSelector).first()).toBeVisible()
}

for (const format of activeHostVisualFormats) {
  test.describe(`${format.label} Codex host width matrix`, () => {
    test(`${format.label} keeps its document surface inside a 280px compact Codex pane`, async ({
      page,
    }) => {
      await page.goto(editorUrl(format, COMPACT))
      await waitForMountedHost(page)
      await prepareFormatDocument(page, format)

      const editor = page.frameLocator('#editor-frame')
      await expect(editor.locator(format.documentSelector)).toBeVisible()
      await expect(editor.locator(format.viewportSelector)).toBeVisible()
      const geometry = await editor
        .locator(format.documentSelector)
        .first()
        .evaluate((element) => {
          const rect = element.getBoundingClientRect()
          return {
            bodyScrollWidth: document.body.scrollWidth,
            viewportWidth: window.innerWidth,
            intersectsViewport: rect.right > 0 && rect.left < window.innerWidth && rect.height > 0,
          }
        })
      expect(geometry.intersectsViewport).toBe(true)
      expect(geometry.bodyScrollWidth).toBeLessThanOrEqual(geometry.viewportWidth)

      const state = await page.evaluate(() => window.__codexVisualHost)
      expect(state.editorLoads).toBe(1)
      expect(state.errors).toEqual([])
    })

    test(`${format.label} remains visible in the 420px Codex narrow sidebar`, async ({ page }) => {
      await page.goto(editorUrl(format, NARROW))
      await waitForMountedHost(page)
      await prepareFormatDocument(page, format)

      const editor = page.frameLocator('#editor-frame')
      await expect(editor.locator(format.documentSelector)).toBeVisible()
      await expect(editor.locator(format.viewportSelector)).toBeVisible()
      await expect(editor.locator(format.statusSelector)).toBeVisible()
      await expect(editor.locator(format.fullscreenButtonSelector)).toBeVisible()
      await expect(page.locator('#editor-frame')).toHaveScreenshot(
        `${format.id}-narrow-sidebar.png`,
      )

      const state = await page.evaluate(() => window.__codexVisualHost)
      expect(state.editorLoads).toBe(1)
      expect(state.errors).toEqual([])
    })

    test(`${format.label} remains fully mounted in the 720px Codex split view`, async ({
      page,
    }) => {
      await page.goto(editorUrl(format, SPLIT))
      await waitForMountedHost(page)
      await prepareFormatDocument(page, format)

      const editor = page.frameLocator('#editor-frame')
      await expect(editor.locator(format.documentSelector)).toBeVisible()
      await expect(editor.locator(format.viewportSelector)).toBeVisible()
      await expect(editor.locator(format.statusSelector)).toBeVisible()
      await expect(editor.locator(format.fullscreenButtonSelector)).toBeInViewport()
      await expect(page.locator('#editor-frame')).toHaveScreenshot(`${format.id}-split-view.png`)

      const state = await page.evaluate(() => window.__codexVisualHost)
      expect(state.editorLoads).toBe(1)
      expect(state.errors).toEqual([])
    })

    test(`${format.label} first fullscreen follows mount readiness at the 1332px Codex width`, async ({
      page,
    }) => {
      await page.goto(editorUrl(format, SPLIT, true))
      await page.waitForFunction(
        ({ fullscreenWidth }) => {
          const state = window.__codexVisualHost
          return (
            state?.displayMode === 'fullscreen' &&
            state.polls > 0 &&
            state.sizeNotifications.some((size) => size.width === fullscreenWidth)
          )
        },
        { fullscreenWidth: FULLSCREEN.width },
      )
      await prepareFormatDocument(page, format)

      const editor = page.frameLocator('#editor-frame')
      await expect(editor.locator(format.documentSelector)).toBeVisible()
      await expect(editor.locator(format.viewportSelector)).toBeVisible()
      await expect(editor.locator(format.fullscreenButtonSelector)).toHaveAttribute(
        'aria-pressed',
        'true',
      )
      await expect(page.locator('#editor-frame')).toHaveCSS('width', `${FULLSCREEN.width}px`)
      await expect(page.locator('#editor-frame')).toHaveScreenshot(
        `${format.id}-first-fullscreen.png`,
      )

      const state = await page.evaluate(() => window.__codexVisualHost)
      const firstSize = state.events.findIndex((event) => event.startsWith('size:'))
      const firstPoll = state.events.indexOf('tool:office_editor_poll')
      const fullscreenRequest = state.events.indexOf('display-request:fullscreen')
      expect(firstSize).toBeGreaterThanOrEqual(0)
      expect(firstSize).toBeLessThan(firstPoll)
      expect(firstPoll).toBeLessThan(fullscreenRequest)
      expect(state.editorLoads).toBe(1)
      expect(state.errors).toEqual([])
    })

    test(`${format.label} exits fullscreen back to split view without remounting the iframe`, async ({
      page,
    }) => {
      await page.goto(editorUrl(format, SPLIT, true))
      await page.waitForFunction(() => window.__codexVisualHost?.displayMode === 'fullscreen')
      await prepareFormatDocument(page, format)

      const editor = page.frameLocator('#editor-frame')
      await editor.locator(format.fullscreenButtonSelector).click()
      await page.waitForFunction(() => window.__codexVisualHost?.displayMode === 'inline')

      await expect(page.locator('#editor-frame')).toHaveCSS('width', `${SPLIT.width}px`)
      await expect(editor.locator(format.documentSelector)).toBeVisible()
      await expect(editor.locator(format.viewportSelector)).toBeVisible()
      await expect(editor.locator(format.fullscreenButtonSelector)).toHaveAttribute(
        'aria-pressed',
        'false',
      )
      await expect(page.locator('#editor-frame')).toHaveScreenshot(
        `${format.id}-exit-fullscreen.png`,
      )

      const state = await page.evaluate(() => window.__codexVisualHost)
      expect(state.events).toContain('display-request:inline')
      expect(state.editorLoads).toBe(1)
      expect(state.polls).toBeGreaterThan(0)
      expect(state.errors).toEqual([])
    })
  })
}

const pendingScenarios = [
  '420px narrow sidebar',
  '720px split view',
  '1332px first fullscreen',
  'fullscreen exit without iframe remount',
]

for (const format of pendingHostVisualFormats) {
  test.describe(`${format.label} Codex host width matrix`, () => {
    test.skip(true, format.reason)
    for (const scenario of pendingScenarios) {
      test(`${format.label} ${scenario}`, async () => {})
    }
  })
}
