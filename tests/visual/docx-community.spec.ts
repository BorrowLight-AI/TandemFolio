import { expect, test } from '@playwright/test'

test('DOCX exposes file, save, and fullscreen as accessible icon-only controls', async ({
  page,
}) => {
  await page.goto('/?format=docx&width=1280&height=900')
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

test('DOCX file menu stays above the ribbon and its Open action is clickable', async ({ page }) => {
  await page.goto('/?format=docx&width=1280&height=900')
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)

  const editor = page.frameLocator('#editor-frame')
  await editor.getByRole('button', { name: '文件', exact: true }).click()

  const open = editor.getByRole('button', { name: /^\u6253开/ })
  await expect(open).toBeVisible()
  await expect(
    open.evaluate((button) => {
      const bounds = button.getBoundingClientRect()
      const hit = document.elementFromPoint(
        bounds.left + bounds.width / 2,
        bounds.top + bounds.height / 2,
      )
      return hit === button || button.contains(hit)
    }),
  ).resolves.toBe(true)

  await open.click()
  await expect(open).toHaveCount(0)
})

test('a generated DOCX saves through the session-bound local persistence protocol', async ({
  page,
}) => {
  await page.goto('/?format=docx&width=1280&height=900')
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)

  await page.evaluate(() => {
    window.__codexVisualHost.enqueueCommand({
      commandId: 'docx-save-content',
      baseRevision: 0,
      operation: 'docx.text.insert',
      arguments: { text: 'Quarterly Review' },
    })
  })
  await expect
    .poll(() => page.evaluate(() => window.__codexVisualHost.acknowledgements.length))
    .toBe(1)

  const downloadPromise = page.waitForEvent('download', { timeout: 5_000 })
  await page.evaluate(() => {
    window.__codexVisualHost.enqueueCommand({
      commandId: 'docx-save',
      baseRevision: 1,
      operation: 'docx.document.save',
      arguments: {},
    })
  })
  await expect
    .poll(() => page.evaluate(() => window.__codexVisualHost.acknowledgements.length))
    .toBe(2)
  await expect
    .poll(() => page.evaluate(() => window.__codexVisualHost.acknowledgements.at(-1)))
    .toMatchObject({ commandId: 'docx-save', ok: true })
  const download = await downloadPromise

  expect(download.suggestedFilename()).toMatch(/\.docx$/)
  await expect
    .poll(() => page.evaluate(() => window.__codexVisualHost.downloads.at(-1)))
    .toMatchObject({
      fileName: download.suggestedFilename(),
      mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      size: expect.any(Number),
    })
  expect(await page.evaluate(() => window.__codexVisualHost.events)).toEqual(
    expect.arrayContaining([
      'tool:office_editor_begin_document_save',
      'tool:office_editor_commit_document_save',
    ]),
  )
})

test('an offscreen DOCX releases its document DOM and resumes the same session', async ({
  page,
}) => {
  await page.goto('/?format=docx&width=720&height=900')
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)

  const editor = page.frameLocator('#editor-frame')
  await expect(editor.locator('html')).toHaveAttribute('data-live-editor-connection', 'connected', {
    timeout: 30_000,
  })
  await page.evaluate(() => {
    window.__codexVisualHost.enqueueCommand({
      commandId: 'docx-lifecycle-content',
      baseRevision: 0,
      operation: 'docx.text.insert',
      arguments: { text: 'Retained DOCX content' },
    })
  })
  await expect
    .poll(() =>
      page.evaluate(() =>
        window.__codexVisualHost.acknowledgements.find(
          (entry) => entry.commandId === 'docx-lifecycle-content',
        ),
      ),
    )
    .toMatchObject({ ok: true, revision: 1 })
  await expect(editor.locator('.editor-scroll .doc-page.ProseMirror')).toContainText(
    'Retained DOCX content',
  )

  await page.evaluate(() => {
    window.__codexVisualHost.enqueueCommand({
      commandId: 'docx-lifecycle-textbox',
      baseRevision: 1,
      operation: 'docx.textbox.insert',
      arguments: { afterBlockIndex: 0, widthEmu: 2_400_000, heightEmu: 900_000 },
    })
  })
  const textbox = editor.locator('.doc-textbox .ProseMirror')
  await expect(textbox).toHaveCount(1)
  await textbox.evaluate((element) => {
    element.dispatchEvent(new MouseEvent('dblclick', { bubbles: true }))
  })
  await expect(textbox).toHaveAttribute('contenteditable', 'true')
  await textbox.fill('Pending textbox content', { force: true })

  await editor.locator('#root').evaluate((root) => {
    root.dispatchEvent(
      Object.assign(new Event('contentvisibilityautostatechange'), { skipped: true }),
    )
  })
  await expect(editor.locator('.editor-workspace-suspended')).toHaveCount(1)
  await expect(editor.locator('.editor-scroll .doc-page.ProseMirror')).toHaveCount(0)

  await editor.locator('#root').evaluate((root) => {
    root.dispatchEvent(
      Object.assign(new Event('contentvisibilityautostatechange'), { skipped: false }),
    )
  })
  await expect(editor.locator('.editor-scroll .doc-page.ProseMirror')).toContainText(
    'Retained DOCX content',
  )
  await expect(editor.locator('.doc-textbox .ProseMirror')).toContainText('Pending textbox content')

  const state = await page.evaluate(() => window.__codexVisualHost)
  expect(state.editorLoads).toBe(1)
  expect(state.errors).toEqual([])
})

test('DOCX checkpoints direct UI changes once per persisted state version', async ({ page }) => {
  await page.goto('/?format=docx&width=1280&height=900')
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)
  const editor = page.frameLocator('#editor-frame')
  await expect(editor.locator('html')).toHaveAttribute('data-live-editor-connection', 'connected', {
    timeout: 30_000,
  })

  await page.evaluate(() => {
    window.__codexVisualHost.enqueueCommand({
      commandId: 'docx-recovery-content',
      baseRevision: 0,
      operation: 'docx.text.insert',
      arguments: { text: 'Recovery baseline' },
    })
  })
  await expect
    .poll(() =>
      page.evaluate(() =>
        window.__codexVisualHost.acknowledgements.find(
          (entry) => entry.commandId === 'docx-recovery-content',
        ),
      ),
    )
    .toMatchObject({ ok: true, revision: 1 })
  await expect.poll(() => page.evaluate(() => window.__codexVisualHost.recoveryCommits)).toBe(1)

  await editor.locator('.ribbon-tab:not(.ribbon-tab-file)').nth(3).click()
  await editor.locator('.rb-big').nth(3).click()
  await editor.locator('.color-palette-page .color-swatch').nth(1).click()
  await expect.poll(() => page.evaluate(() => window.__codexVisualHost.recoveryCommits)).toBe(2)

  await page.waitForTimeout(2300)
  expect(await page.evaluate(() => window.__codexVisualHost.recoveryCommits)).toBe(2)

  await editor.locator('.rb-big').nth(3).click()
  await editor.locator('.color-palette-page .color-swatch').nth(2).click()
  await expect.poll(() => page.evaluate(() => window.__codexVisualHost.recoveryCommits)).toBe(3)
})
