import { readFile } from 'node:fs/promises'
import { fileURLToPath } from 'node:url'
import { expect, test, type Page } from '@playwright/test'

const fixturePath = fileURLToPath(
  new URL('../../packages/pptx-engine/tests/fixtures/01_standard_business.pptx', import.meta.url),
)

const NARROW = { width: 420, height: 900 }
const SPLIT = { width: 720, height: 900 }
const FULLSCREEN = { width: 1332, height: 1280 }

function hostUrl(dimensions: { width: number; height: number }, acceptFullscreen = false): string {
  const params = new URLSearchParams({
    format: 'pptx',
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

async function waitForCommunityRenderer(page: Page): Promise<void> {
  await page.waitForFunction(() => {
    const state = window.__codexVisualHost
    return state?.initialized && state.polls > 0 && state.sizeNotifications.length > 0
  })
  await expect(
    page.frameLocator('#editor-frame').locator('.stage-wrap .konvajs-content canvas').first(),
  ).toBeVisible()
}

async function thumbnailPixels(page: Page, slideIndex: number): Promise<string> {
  return page
    .frameLocator('#editor-frame')
    .locator('.thumb')
    .nth(slideIndex)
    .locator('.konvajs-content canvas')
    .first()
    .evaluate((canvas) => (canvas as HTMLCanvasElement).toDataURL())
}

test('the product boots the pinned community App and its original Ribbon', async ({ page }) => {
  await page.goto('/?format=pptx&width=1332&height=900')
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)

  const editor = page.frameLocator('#editor-frame')
  await expect(editor.locator('.app')).toBeVisible()
  await expect(editor.locator('.ribbon')).toBeVisible()
  await expect(editor.getByRole('button', { name: '切换' })).toBeVisible()
  await expect(editor.getByRole('button', { name: '动画' })).toBeVisible()
  await expect(editor.locator('.ribbon-body')).toBeVisible()
})

test('PPTX exposes file, save, and fullscreen as accessible icon-only controls', async ({
  page,
}) => {
  await page.goto('/?format=pptx&width=1332&height=900')
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)

  const editor = page.frameLocator('#editor-frame')
  const controls = [
    editor.getByRole('button', { name: '文件', exact: true }),
    editor.getByRole('button', { name: /^保存/ }),
    editor.getByRole('button', { name: '全屏', exact: true }),
  ]

  for (const control of controls) {
    await expect(control).toBeVisible()
    await expect(control.locator('svg')).toHaveCount(1)
    await expect(control).toHaveText('')
  }
})

test('PPTX file control starts at the leading edge of the embedded ribbon', async ({ page }) => {
  await page.goto('/?format=pptx&width=1332&height=900')
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)

  const editor = page.frameLocator('#editor-frame')
  const ribbonLeft = await editor
    .locator('.ribbon-tabs')
    .evaluate((element) => element.getBoundingClientRect().left)
  const fileLeft = await editor
    .getByRole('button', { name: '文件', exact: true })
    .evaluate((element) => element.getBoundingClientRect().left)

  expect(fileLeft).toBe(ribbonLeft)
})

test('a generated PPTX saves through the session-bound local persistence protocol', async ({
  page,
}) => {
  await page.goto('/?format=pptx&width=1332&height=900')
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)

  const editor = page.frameLocator('#editor-frame')
  await editor.getByRole('button', { name: '新建幻灯片' }).click()
  const downloadPromise = page.waitForEvent('download')
  await editor.getByRole('button', { name: /^保存/ }).click()
  const download = await downloadPromise

  expect(download.suggestedFilename()).toBe('Untitled.pptx')
  await expect
    .poll(() => page.evaluate(() => window.__codexVisualHost.downloads))
    .toEqual([
      {
        fileName: 'Untitled.pptx',
        mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
        size: expect.any(Number),
      },
    ])
  expect(await page.evaluate(() => window.__codexVisualHost.events)).toEqual(
    expect.arrayContaining([
      'tool:office_editor_begin_document_save',
      'tool:office_editor_commit_document_save',
    ]),
  )
})

test('Save As lets the user rename a generated PPTX before local persistence', async ({ page }) => {
  await page.goto('/?format=pptx&width=1332&height=900')
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)

  const editor = page.frameLocator('#editor-frame')
  await editor.getByRole('button', { name: '新建幻灯片' }).click()
  await editor.getByRole('button', { name: '文件' }).click()
  await editor.getByRole('button', { name: /另存为/ }).click()

  const dialog = editor.getByRole('dialog', { name: /另存为/ })
  await expect(dialog).toBeVisible()
  await dialog.getByLabel('名称').fill('Renamed Deck')
  const downloadPromise = page.waitForEvent('download')
  await dialog.getByRole('button', { name: '保存' }).click()
  const download = await downloadPromise

  expect(download.suggestedFilename()).toBe('Renamed Deck.pptx')
  await expect
    .poll(() => page.evaluate(() => window.__codexVisualHost.downloads.at(-1)?.fileName))
    .toBe('Renamed Deck.pptx')
})

test('the visible presentation stays painted while the host toggles pinned summary', async ({
  page,
}) => {
  await page.goto('/?format=pptx&width=1332&height=900')
  await waitForCommunityRenderer(page)

  const editor = page.frameLocator('#editor-frame')
  await editor.locator('body').evaluate(() => {
    Object.defineProperty(document, 'visibilityState', {
      configurable: true,
      value: 'hidden',
    })
    document.dispatchEvent(new Event('visibilitychange'))
  })

  await expect(editor.locator('.app')).toBeVisible()
  await expect(editor.locator('.stage-wrap')).toBeVisible()
})

test('an offscreen PPTX releases its canvas work and resumes without remounting', async ({
  page,
}) => {
  await page.goto('/?format=pptx&width=1332&height=900')
  await waitForCommunityRenderer(page)

  const editor = page.frameLocator('#editor-frame')
  await editor.locator('#root').evaluate((root) => {
    root.dispatchEvent(
      Object.assign(new Event('contentvisibilityautostatechange'), { skipped: true }),
    )
  })
  await expect(editor.locator('.workspace-suspended')).toHaveCount(1)
  await expect(editor.locator('.konvajs-content canvas')).toHaveCount(0)
  await expect
    .poll(() =>
      editor.locator('#root').evaluate((root) => getComputedStyle(root).contentVisibility),
    )
    .toBe('auto')

  await editor.locator('#root').evaluate((root) => {
    root.dispatchEvent(
      Object.assign(new Event('contentvisibilityautostatechange'), { skipped: false }),
    )
  })
  await expect(editor.locator('.stage-wrap .konvajs-content canvas').first()).toBeVisible()

  const state = await page.evaluate(() => window.__codexVisualHost)
  expect(state.editorLoads).toBe(1)
  expect(state.errors).toEqual([])
})

test('a user opens a PPTX and changes slide through the original community renderer', async ({
  page,
}) => {
  await page.goto('/?format=pptx&width=1332&height=900')
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)

  const editor = page.frameLocator('#editor-frame')
  await editor.getByRole('button', { name: '文件' }).click()
  await editor.getByRole('button', { name: /打开/ }).click()
  await editor.locator('input[type="file"][accept=".pptx"]').setInputFiles({
    name: 'community-slides.pptx',
    mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    buffer: await readFile(fixturePath),
  })

  const thumbnails = editor.locator('.thumb')
  await expect.poll(() => thumbnails.count()).toBeGreaterThan(1)
  await expect(editor.locator('.stage-wrap .konvajs-content canvas').first()).toBeVisible()

  await thumbnails.nth(1).click()

  await expect(thumbnails.nth(1)).toHaveClass(/active/)
})

test('the original community find and replace survives save and reopen', async ({ page }) => {
  await page.goto('/?format=pptx&width=1332&height=900')
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)

  const editor = page.frameLocator('#editor-frame')
  await editor.getByRole('button', { name: '文件' }).click()
  await editor.getByRole('button', { name: /打开/ }).click()
  await editor.locator('input[type="file"][accept=".pptx"]').setInputFiles({
    name: 'find-replace-source.pptx',
    mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    buffer: await readFile(fixturePath),
  })
  await expect.poll(() => editor.locator('.thumb').count()).toBeGreaterThan(1)

  await editor.getByRole('button', { name: '查找/替换' }).click()
  await editor.getByPlaceholder('查找内容').fill('Q3 Business Review')
  await editor.getByPlaceholder('替换为').fill('Q4 Business Review')
  await editor.getByRole('button', { name: '全部替换' }).click()
  await expect(editor.locator('.find-panel-status')).toHaveText('已替换 1 处')

  const downloadPromise = page.waitForEvent('download')
  await editor.getByRole('button', { name: /^保存/ }).click()
  const savedPath = await (await downloadPromise).path()
  expect(savedPath).not.toBeNull()

  await editor.getByRole('button', { name: '文件' }).click()
  await editor.getByRole('button', { name: /打开/ }).click()
  await editor.locator('input[type="file"][accept=".pptx"]').setInputFiles({
    name: 'find-replace-saved.pptx',
    mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    buffer: await readFile(savedPath!),
  })
  await editor.getByPlaceholder('查找内容').fill('Q4 Business Review')
  await expect(editor.locator('.find-panel-status')).toHaveText('第 0 项，共 1 项')
})

test('the original community text editor preserves a typed title through save and reopen', async ({
  page,
}) => {
  await page.goto('/?format=pptx&width=1332&height=900')
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)

  const editor = page.frameLocator('#editor-frame')
  await editor.getByRole('button', { name: '文件' }).click()
  await editor.getByRole('button', { name: /打开/ }).click()
  await editor.locator('input[type="file"][accept=".pptx"]').setInputFiles({
    name: 'text-edit-source.pptx',
    mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    buffer: await readFile(fixturePath),
  })
  await expect.poll(() => editor.locator('.thumb').count()).toBeGreaterThan(1)

  await editor.locator('body').press('Tab')
  await editor.locator('body').press('x')
  const textEditor = editor.locator('[contenteditable="true"]')
  await expect(textEditor).toBeVisible()
  await textEditor.fill('Community Renderer Restored')
  await textEditor.press('Escape')
  await expect(textEditor).toHaveCount(0)

  const downloadPromise = page.waitForEvent('download')
  await editor.getByRole('button', { name: /^保存/ }).click()
  const savedPath = await (await downloadPromise).path()
  expect(savedPath).not.toBeNull()

  await editor.getByRole('button', { name: '文件' }).click()
  await editor.getByRole('button', { name: /打开/ }).click()
  await editor.locator('input[type="file"][accept=".pptx"]').setInputFiles({
    name: 'text-edit-saved.pptx',
    mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    buffer: await readFile(savedPath!),
  })
  await editor.getByRole('button', { name: '查找/替换' }).click()
  await editor.getByPlaceholder('查找内容').fill('Community Renderer Restored')
  await expect(editor.locator('.find-panel-status')).toHaveText('第 0 项，共 1 项')
})

test('the original community canvas preserves a keyboard-moved shape through save and reopen', async ({
  page,
}) => {
  await page.goto('/?format=pptx&width=1332&height=900')
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)

  const editor = page.frameLocator('#editor-frame')
  await editor.getByRole('button', { name: '文件' }).click()
  await editor.getByRole('button', { name: /打开/ }).click()
  await editor.locator('input[type="file"][accept=".pptx"]').setInputFiles({
    name: 'transform-source.pptx',
    mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    buffer: await readFile(fixturePath),
  })
  await expect.poll(() => editor.locator('.thumb').count()).toBeGreaterThan(1)

  const before = await thumbnailPixels(page, 0)
  await editor.locator('body').press('Tab')
  await editor.locator('body').press('Shift+ArrowRight')
  await expect.poll(() => thumbnailPixels(page, 0)).not.toBe(before)
  const moved = await thumbnailPixels(page, 0)

  const downloadPromise = page.waitForEvent('download')
  await editor.getByRole('button', { name: /^保存/ }).click()
  const savedPath = await (await downloadPromise).path()
  expect(savedPath).not.toBeNull()

  await editor.getByRole('button', { name: '文件' }).click()
  await editor.getByRole('button', { name: /打开/ }).click()
  await editor.locator('input[type="file"][accept=".pptx"]').setInputFiles({
    name: 'transform-saved.pptx',
    mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    buffer: await readFile(savedPath!),
  })
  await expect.poll(() => thumbnailPixels(page, 0)).toBe(moved)
})

test('the original community ribbon preserves selected-shape bold formatting through save and reopen', async ({
  page,
}) => {
  await page.goto('/?format=pptx&width=1332&height=900')
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)

  const editor = page.frameLocator('#editor-frame')
  await editor.getByRole('button', { name: '文件' }).click()
  await editor.getByRole('button', { name: /打开/ }).click()
  await editor.locator('input[type="file"][accept=".pptx"]').setInputFiles({
    name: 'bold-source.pptx',
    mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    buffer: await readFile(fixturePath),
  })
  await expect.poll(() => editor.locator('.thumb').count()).toBeGreaterThan(1)

  const before = await thumbnailPixels(page, 0)
  await editor.locator('body').press('Tab')
  await editor.getByRole('button', { name: '加粗' }).click()
  await expect.poll(() => thumbnailPixels(page, 0)).not.toBe(before)
  const bold = await thumbnailPixels(page, 0)

  const downloadPromise = page.waitForEvent('download')
  await editor.getByRole('button', { name: /^保存/ }).click()
  const savedPath = await (await downloadPromise).path()
  expect(savedPath).not.toBeNull()

  await editor.getByRole('button', { name: '文件' }).click()
  await editor.getByRole('button', { name: /打开/ }).click()
  await editor.locator('input[type="file"][accept=".pptx"]').setInputFiles({
    name: 'bold-saved.pptx',
    mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    buffer: await readFile(savedPath!),
  })
  await expect.poll(() => thumbnailPixels(page, 0)).toBe(bold)
})

test('the original community ribbon preserves paragraph alignment through save and reopen', async ({
  page,
}) => {
  await page.goto('/?format=pptx&width=1332&height=900')
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)

  const editor = page.frameLocator('#editor-frame')
  await editor.getByRole('button', { name: '文件' }).click()
  await editor.getByRole('button', { name: /打开/ }).click()
  await editor.locator('input[type="file"][accept=".pptx"]').setInputFiles({
    name: 'paragraph-alignment-source.pptx',
    mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    buffer: await readFile(fixturePath),
  })
  await expect.poll(() => editor.locator('.thumb').count()).toBeGreaterThan(1)

  const before = await thumbnailPixels(page, 0)
  await editor.locator('body').press('Tab')
  await editor.getByRole('button', { name: '段落', exact: true }).click()
  await editor.getByRole('button', { name: '右对齐', exact: true }).click()
  await expect.poll(() => thumbnailPixels(page, 0)).not.toBe(before)
  const aligned = await thumbnailPixels(page, 0)

  const downloadPromise = page.waitForEvent('download')
  await editor.getByRole('button', { name: /^保存/ }).click()
  const savedPath = await (await downloadPromise).path()
  expect(savedPath).not.toBeNull()

  await editor.getByRole('button', { name: '文件' }).click()
  await editor.getByRole('button', { name: /打开/ }).click()
  await editor.locator('input[type="file"][accept=".pptx"]').setInputFiles({
    name: 'paragraph-alignment-saved.pptx',
    mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    buffer: await readFile(savedPath!),
  })
  await expect.poll(() => thumbnailPixels(page, 0)).toBe(aligned)
})

test('the original community canvas deletes a selected object through save and reopen', async ({
  page,
}) => {
  await page.goto('/?format=pptx&width=1332&height=900')
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)

  const editor = page.frameLocator('#editor-frame')
  await editor.getByRole('button', { name: '文件' }).click()
  await editor.getByRole('button', { name: /打开/ }).click()
  await editor.locator('input[type="file"][accept=".pptx"]').setInputFiles({
    name: 'delete-object-source.pptx',
    mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    buffer: await readFile(fixturePath),
  })
  await expect.poll(() => editor.locator('.thumb').count()).toBeGreaterThan(1)

  const before = await thumbnailPixels(page, 0)
  await editor.locator('body').press('Tab')
  await editor.locator('body').press('Delete')
  await expect.poll(() => thumbnailPixels(page, 0)).not.toBe(before)
  const deleted = await thumbnailPixels(page, 0)

  const downloadPromise = page.waitForEvent('download')
  await editor.getByRole('button', { name: /^保存/ }).click()
  const savedPath = await (await downloadPromise).path()
  expect(savedPath).not.toBeNull()

  await editor.getByRole('button', { name: '文件' }).click()
  await editor.getByRole('button', { name: /打开/ }).click()
  await editor.locator('input[type="file"][accept=".pptx"]').setInputFiles({
    name: 'delete-object-saved.pptx',
    mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    buffer: await readFile(savedPath!),
  })
  await expect.poll(() => thumbnailPixels(page, 0)).toBe(deleted)
})

test('the original community ribbon adds a slide that survives save and reopen', async ({
  page,
}) => {
  await page.goto('/?format=pptx&width=1332&height=900')
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)

  const editor = page.frameLocator('#editor-frame')
  await editor.getByRole('button', { name: '文件' }).click()
  await editor.getByRole('button', { name: /打开/ }).click()
  await editor.locator('input[type="file"][accept=".pptx"]').setInputFiles({
    name: 'community-slides.pptx',
    mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    buffer: await readFile(fixturePath),
  })

  const thumbnails = editor.locator('.thumb')
  await expect.poll(() => thumbnails.count()).toBeGreaterThan(1)
  const initialSlideCount = await thumbnails.count()

  await editor.getByRole('button', { name: '新建幻灯片' }).click()
  await expect(thumbnails).toHaveCount(initialSlideCount + 1)

  const downloadPromise = page.waitForEvent('download')
  await editor.getByRole('button', { name: /^保存/ }).click()
  const download = await downloadPromise
  const savedPath = await download.path()
  expect(savedPath).not.toBeNull()

  await editor.getByRole('button', { name: '文件' }).click()
  await editor.getByRole('button', { name: /打开/ }).click()
  await editor.locator('input[type="file"][accept=".pptx"]').setInputFiles({
    name: 'saved-community-slides.pptx',
    mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    buffer: await readFile(savedPath!),
  })
  await expect(thumbnails).toHaveCount(initialSlideCount + 1)
})

test('the original community renderer duplicates a slide through save and reopen', async ({
  page,
}) => {
  await page.goto('/?format=pptx&width=1332&height=900')
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)

  const editor = page.frameLocator('#editor-frame')
  await editor.getByRole('button', { name: '文件' }).click()
  await editor.getByRole('button', { name: /打开/ }).click()
  await editor.locator('input[type="file"][accept=".pptx"]').setInputFiles({
    name: 'duplicate-source.pptx',
    mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    buffer: await readFile(fixturePath),
  })

  const thumbnails = editor.locator('.thumb')
  await expect.poll(() => thumbnails.count()).toBeGreaterThan(1)
  const initialSlideCount = await thumbnails.count()
  await thumbnails.nth(1).click({ button: 'right' })
  await editor.getByRole('button', { name: '复制幻灯片' }).click()
  await expect(thumbnails).toHaveCount(initialSlideCount + 1)
  await expect(thumbnails.nth(2)).toHaveClass(/active/)

  const downloadPromise = page.waitForEvent('download')
  await editor.getByRole('button', { name: /^保存/ }).click()
  const savedPath = await (await downloadPromise).path()
  expect(savedPath).not.toBeNull()
  await editor.getByRole('button', { name: '文件' }).click()
  await editor.getByRole('button', { name: /打开/ }).click()
  await editor.locator('input[type="file"][accept=".pptx"]').setInputFiles({
    name: 'duplicated-slides.pptx',
    mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    buffer: await readFile(savedPath!),
  })
  await expect(thumbnails).toHaveCount(initialSlideCount + 1)
})

test('the original community renderer deletes a slide through save and reopen', async ({
  page,
}) => {
  await page.goto('/?format=pptx&width=1332&height=900')
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)

  const editor = page.frameLocator('#editor-frame')
  await editor.getByRole('button', { name: '文件' }).click()
  await editor.getByRole('button', { name: /打开/ }).click()
  await editor.locator('input[type="file"][accept=".pptx"]').setInputFiles({
    name: 'delete-source.pptx',
    mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    buffer: await readFile(fixturePath),
  })

  const thumbnails = editor.locator('.thumb')
  await expect.poll(() => thumbnails.count()).toBeGreaterThan(1)
  const initialSlideCount = await thumbnails.count()
  await thumbnails.nth(1).click({ button: 'right' })
  await editor.getByRole('button', { name: '删除幻灯片' }).click()
  await expect(thumbnails).toHaveCount(initialSlideCount - 1)
  await expect(thumbnails.nth(1)).toHaveClass(/active/)

  const downloadPromise = page.waitForEvent('download')
  await editor.getByRole('button', { name: /^保存/ }).click()
  const savedPath = await (await downloadPromise).path()
  expect(savedPath).not.toBeNull()
  await editor.getByRole('button', { name: '文件' }).click()
  await editor.getByRole('button', { name: /打开/ }).click()
  await editor.locator('input[type="file"][accept=".pptx"]').setInputFiles({
    name: 'deleted-slide.pptx',
    mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    buffer: await readFile(savedPath!),
  })
  await expect(thumbnails).toHaveCount(initialSlideCount - 1)
})

test('the original community renderer reorders slides through save and reopen', async ({
  page,
}) => {
  await page.goto('/?format=pptx&width=1332&height=900')
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)

  const editor = page.frameLocator('#editor-frame')
  await editor.getByRole('button', { name: '文件' }).click()
  await editor.getByRole('button', { name: /打开/ }).click()
  await editor.locator('input[type="file"][accept=".pptx"]').setInputFiles({
    name: 'reorder-source.pptx',
    mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    buffer: await readFile(fixturePath),
  })
  const thumbnails = editor.locator('.thumb')
  await expect.poll(() => thumbnails.count()).toBeGreaterThan(1)

  const firstBefore = await thumbnailPixels(page, 0)
  const secondBefore = await thumbnailPixels(page, 1)
  expect(firstBefore).not.toBe(secondBefore)

  await thumbnails.nth(0).evaluate((source) => {
    const transfer = new DataTransfer()
    ;(window as Window & { __pptxDragTransfer?: DataTransfer }).__pptxDragTransfer = transfer
    source.dispatchEvent(new DragEvent('dragstart', { bubbles: true, dataTransfer: transfer }))
  })
  await page.waitForTimeout(50)
  await thumbnails.nth(1).evaluate((target) => {
    const transfer = (window as Window & { __pptxDragTransfer?: DataTransfer }).__pptxDragTransfer
    const rect = target.getBoundingClientRect()
    target.dispatchEvent(
      new DragEvent('dragover', {
        bubbles: true,
        cancelable: true,
        clientX: rect.left + rect.width / 2,
        clientY: rect.bottom - 1,
        dataTransfer: transfer,
      }),
    )
    target.dispatchEvent(
      new DragEvent('drop', {
        bubbles: true,
        cancelable: true,
        clientX: rect.left + rect.width / 2,
        clientY: rect.bottom - 1,
        dataTransfer: transfer,
      }),
    )
  })
  await expect(thumbnails.nth(1)).toHaveClass(/active/)
  await expect.poll(() => thumbnailPixels(page, 0)).toBe(secondBefore)
  await expect.poll(() => thumbnailPixels(page, 1)).toBe(firstBefore)

  const downloadPromise = page.waitForEvent('download')
  await editor.getByRole('button', { name: /^保存/ }).click()
  const savedPath = await (await downloadPromise).path()
  expect(savedPath).not.toBeNull()
  await editor.getByRole('button', { name: '文件' }).click()
  await editor.getByRole('button', { name: /打开/ }).click()
  await editor.locator('input[type="file"][accept=".pptx"]').setInputFiles({
    name: 'reordered-slides.pptx',
    mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    buffer: await readFile(savedPath!),
  })
  await expect.poll(() => thumbnailPixels(page, 0)).toBe(secondBefore)
  await expect.poll(() => thumbnailPixels(page, 1)).toBe(firstBefore)
})

test('the original community ribbon undoes a slide lifecycle edit before save and reopen', async ({
  page,
}) => {
  await page.goto('/?format=pptx&width=1332&height=900')
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)

  const editor = page.frameLocator('#editor-frame')
  await editor.getByRole('button', { name: '文件' }).click()
  await editor.getByRole('button', { name: /打开/ }).click()
  await editor.locator('input[type="file"][accept=".pptx"]').setInputFiles({
    name: 'undo-source.pptx',
    mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    buffer: await readFile(fixturePath),
  })

  const thumbnails = editor.locator('.thumb')
  await expect.poll(() => thumbnails.count()).toBeGreaterThan(1)
  const initialSlideCount = await thumbnails.count()
  await editor.getByRole('button', { name: '新建幻灯片' }).click()
  await expect(thumbnails).toHaveCount(initialSlideCount + 1)
  await editor.getByRole('button', { name: '撤销' }).click()
  await expect(thumbnails).toHaveCount(initialSlideCount)

  const downloadPromise = page.waitForEvent('download')
  await editor.getByRole('button', { name: '文件' }).click()
  await editor.getByRole('button', { name: /另存为/ }).click()
  const saveAsDialog = editor.getByRole('dialog', { name: /另存为/ })
  await expect(saveAsDialog).toBeVisible()
  await saveAsDialog.getByRole('button', { name: '保存' }).click()
  const savedPath = await (await downloadPromise).path()
  expect(savedPath).not.toBeNull()
  await editor.getByRole('button', { name: '文件' }).click()
  await editor.getByRole('button', { name: /打开/ }).click()
  await editor.locator('input[type="file"][accept=".pptx"]').setInputFiles({
    name: 'undo-saved.pptx',
    mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    buffer: await readFile(savedPath!),
  })
  await expect(thumbnails).toHaveCount(initialSlideCount)
})

test('PPTX community renderer remains visible in the 420px Codex narrow sidebar', async ({
  page,
}) => {
  await page.goto(hostUrl(NARROW))
  await waitForCommunityRenderer(page)

  const editor = page.frameLocator('#editor-frame')
  await expect(editor.locator('.app')).toBeVisible()
  await expect(editor.locator('.stage-wrap')).toBeVisible()
  await expect(editor.locator('.status-bar')).toBeVisible()
  await expect(editor.locator('.host-fullscreen[aria-pressed]')).toBeVisible()
  await expect(page.locator('#editor-frame')).toHaveScreenshot('pptx-narrow-sidebar.png')

  const state = await page.evaluate(() => window.__codexVisualHost)
  expect(state.editorLoads).toBe(1)
  expect(state.errors).toEqual([])
})

test('PPTX community renderer remains mounted in the 720px Codex split view', async ({ page }) => {
  await page.goto(hostUrl(SPLIT))
  await waitForCommunityRenderer(page)

  const editor = page.frameLocator('#editor-frame')
  await expect(editor.locator('.app')).toBeVisible()
  await expect(editor.locator('.stage-wrap')).toBeVisible()
  await expect(editor.locator('.host-fullscreen[aria-pressed]')).toBeInViewport()
  await expect(page.locator('#editor-frame')).toHaveScreenshot('pptx-split-view.png')

  const state = await page.evaluate(() => window.__codexVisualHost)
  expect(state.editorLoads).toBe(1)
  expect(state.errors).toEqual([])
})

test('PPTX community renderer first fullscreen follows host mount readiness', async ({ page }) => {
  await page.goto(hostUrl(SPLIT, true))
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

  const editor = page.frameLocator('#editor-frame')
  await expect(editor.locator('.stage-wrap .konvajs-content canvas').first()).toBeVisible()
  await expect(editor.locator('.host-fullscreen[aria-pressed]')).toHaveAttribute(
    'aria-pressed',
    'true',
  )
  await expect(page.locator('#editor-frame')).toHaveCSS('width', `${FULLSCREEN.width}px`)
  await expect(page.locator('#editor-frame')).toHaveScreenshot('pptx-first-fullscreen.png')

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

test('PPTX community renderer exits fullscreen without remounting its iframe', async ({ page }) => {
  await page.goto(hostUrl(SPLIT, true))
  await page.waitForFunction(() => window.__codexVisualHost?.displayMode === 'fullscreen')

  const editor = page.frameLocator('#editor-frame')
  await editor.locator('.host-fullscreen[aria-pressed]').click()
  await page.waitForFunction(() => window.__codexVisualHost?.displayMode === 'inline')

  await expect(page.locator('#editor-frame')).toHaveCSS('width', `${SPLIT.width}px`)
  await expect(editor.locator('.app')).toBeVisible()
  await expect(editor.locator('.stage-wrap .konvajs-content canvas').first()).toBeVisible()
  await expect(editor.locator('.host-fullscreen[aria-pressed]')).toHaveAttribute(
    'aria-pressed',
    'false',
  )
  await expect(page.locator('#editor-frame')).toHaveScreenshot('pptx-exit-fullscreen.png')

  const state = await page.evaluate(() => window.__codexVisualHost)
  expect(state.events).toContain('display-request:inline')
  expect(state.editorLoads).toBe(1)
  expect(state.polls).toBeGreaterThan(0)
  expect(state.errors).toEqual([])
})
