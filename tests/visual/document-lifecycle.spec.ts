import { expect, test } from '@playwright/test'
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'
import { randomUUID } from 'node:crypto'
import { fixture } from './document-fixtures'


for (const format of ['docx', 'markdown', 'xlsx', 'pptx', 'pdf']) {
  test(`${format}: real save path and native content survive a broker restart and cold remount`, async ({
    page,
    request,
  }) => {
    test.setTimeout(90_000)
    const broker = randomUUID()
    const root = await mkdtemp(join(tmpdir(), 'tandemfolio-native-reopen-'))
    const fileName = `Exact-${format}.${format === 'markdown' ? 'md' : format}`
    const source = join(root, fileName)
    await writeFile(source, await fixture(format))
    await page.addInitScript(() => {
      // Exercise the standard File input used by sandboxed MCP hosts.
      Object.defineProperty(window, 'showDirectoryPicker', { configurable: true, value: undefined })
      Object.defineProperty(window, 'showOpenFilePicker', { configurable: true, value: undefined })
    })
    const call = async (name: string, args: Record<string, unknown>) => {
      const response = await request.post(`/broker?id=${broker}`, {
        data: { name, arguments: args },
      })
      const result = await response.json()
      expect(result.isError, JSON.stringify(result.structuredContent)).not.toBe(true)
      return result.structuredContent
    }
    try {
      const { session } = await call('office_create_session', { format })
      const { viewId } = await call(
        format === 'docx' ? 'office_show_editor' : `office_show_${format}_editor`,
        { sessionId: session.id },
      )
      const query = new URLSearchParams({
        format,
        broker,
        sessionId: session.id,
        viewId,
        width: '1280',
        height: '900',
      })
      await page.goto(`/?${query}`)
      const editor = page.frameLocator('#editor-frame')
      await expect(editor.locator('html')).toHaveAttribute(
        'data-live-editor-connection',
        'connected',
      )
      await call('office_open_local_file', {
        sessionId: session.id,
        baseRevision: 0,
        path: resolve(source),
      })
      const before = (await call('office_get_context', { sessionId: session.id })).session
      await call('office_execute', {
        sessionId: session.id,
        baseRevision: before.revision,
        requestId: 'native-save',
        operations: [{ id: `${format}.document.save`, arguments: {} }],
      })
      await editor.locator('[data-live-session-status] summary').click()
      await expect(editor.getByRole('textbox', { name: '文件绝对路径' })).toHaveValue(source)
      const savedBytes = await readFile(source)
      expect(savedBytes.byteLength).toBeGreaterThan(20)
      await page.goto('about:blank')
      await request.post(`/broker?id=${broker}&action=restart`, { data: {} })
      await page.goto(`/?${query}`)
      await expect
        .poll(async () => {
          const response = await request.post(`/broker?id=${broker}`, {
            data: { name: 'office_get_context', arguments: { sessionId: session.id } },
          })
          return (await response.json()).structuredContent?.session?.fileName
        })
        .toBe(fileName)
      await expect
        .poll(() => page.evaluate(() => window.__codexVisualHost.acknowledgements.length))
        .toBe(1)
      const after = (await call('office_get_context', { sessionId: session.id })).session
      expect(after.filePath).toBe(source)
      expect(after.selection).toEqual(before.selection)
      expect(await readFile(source)).toEqual(savedBytes)
      if (format === 'markdown')
        await expect(editor.locator('.tiptap')).toContainText('Retained contents after restart')
      if (format === 'docx')
        await expect(editor.locator('.editor-scroll .doc-page.ProseMirror').first()).not.toBeEmpty()
      if (format === 'pptx') await expect(editor.locator('canvas').first()).toBeVisible()
      if (format === 'pdf') await expect(editor.locator('canvas').first()).toBeVisible()
      if (format === 'xlsx') await expect(editor.locator('canvas').first()).toBeVisible()

      // Replacing A with a browser-selected file of the same name must detach A's target,
      // checkpoint even clean imported content, and never overwrite A on the next Save.
      const checkpointCount = await page.evaluate(() => window.__codexVisualHost.recoveryCommits)
      if (format === 'xlsx' || format === 'pdf') {
        await editor.locator(`input[type="file"][accept*=".${format}"]`).setInputFiles(source)
      } else {
        const chooserPromise = page.waitForEvent('filechooser')
        if (format !== 'markdown')
          await editor.getByRole('button', { name: '文件', exact: true }).click()
        await editor
          .getByRole('button', { name: /打开|Open Markdown/ })
          .first()
          .click()
        const chooser = await chooserPromise
        await chooser.setFiles(format === 'markdown' ? root : source)
      }
      await expect
        .poll(() => page.evaluate(() => window.__codexVisualHost.recoveryCommits))
        .toBeGreaterThan(checkpointCount)
      expect(
        (await call('office_get_context', { sessionId: session.id })).session.filePath,
      ).toBeNull()
      await page.goto('about:blank')
      await request.post(`/broker?id=${broker}&action=restart`, { data: {} })
      await page.goto(`/?${query}`)
      await expect
        .poll(() => page.evaluate(() => window.__codexVisualHost?.acknowledgements.length))
        .toBe(1)
      const imported = (await call('office_get_context', { sessionId: session.id })).session
      expect(imported.fileName).toBe(fileName)
      expect(imported.filePath).toBeNull()
      await call('office_execute', {
        sessionId: session.id,
        baseRevision: imported.revision,
        requestId: 'save-import',
        operations: [{ id: `${format}.document.save`, arguments: {} }],
      })
      const copyPath = (await call('office_get_context', { sessionId: session.id })).session
        .filePath
      expect(copyPath).not.toBe(source)
      expect((await readFile(copyPath)).byteLength).toBeGreaterThan(20)
      expect(await readFile(source)).toEqual(savedBytes)
      await editor.locator('[data-live-session-status] summary').click()
      await expect(editor.getByRole('textbox', { name: '文件绝对路径' })).toHaveValue(copyPath)
      await page.screenshot({ path: test.info().outputPath(`${format}-saved-path.png`) })
    } finally {
      await page.goto('about:blank')
      await request.post(`/broker?id=${broker}&action=close`, { data: {} })
      await rm(root, { recursive: true, force: true })
    }
  })
}

test('multiple visible and hidden rejected PPTX editors stop polling', async ({ page }) => {
  await page.goto('/')
  await page.setContent(
    Array.from(
      { length: 5 },
      (_, index) =>
        `<iframe data-test-host src="/?format=pptx&pollError=editor_view_conflict&width=420&height=900" style="width:440px;height:920px;${index ? 'display:none' : ''}"></iframe>`,
    ).join(''),
  )
  await expect
    .poll(async () =>
      page.evaluate(() =>
        [...document.querySelectorAll<HTMLIFrameElement>('[data-test-host]')].map(
          (frame) => frame.contentWindow?.__codexVisualHost?.polls ?? 0,
        ),
      ),
    )
    .toEqual([1, 1, 1, 1, 1])
  await page.waitForTimeout(1200)
  expect(
    await page.evaluate(() =>
      [...document.querySelectorAll<HTMLIFrameElement>('[data-test-host]')].map(
        (frame) => frame.contentWindow?.__codexVisualHost?.polls ?? 0,
      ),
    ),
  ).toEqual([1, 1, 1, 1, 1])
})
