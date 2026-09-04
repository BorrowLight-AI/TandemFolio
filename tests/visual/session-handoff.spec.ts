import { expect, test } from '@playwright/test'
import { randomUUID } from 'node:crypto'
import { resolve } from 'node:path'
import { join } from 'node:path'
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { fixture } from './document-fixtures'

test('Continue here restores unsaved PPTX from an owner that still reports active', async ({
  page: owner,
  context,
  request,
}) => {
  const broker = randomUUID()
  const call = async (name: string, args: Record<string, unknown>) => {
    const result = await (
      await request.post(`/broker?id=${broker}`, {
        data: { name, arguments: args },
      })
    ).json()
    expect(result.isError, JSON.stringify(result.structuredContent)).not.toBe(true)
    return result.structuredContent
  }
  const returning = await context.newPage()
  try {
    const { session } = await call('office_create_session', { format: 'pptx' })
    const { viewId } = await call('office_show_pptx_editor', { sessionId: session.id })
    const url =
      '/?' +
      new URLSearchParams({
        broker,
        format: 'pptx',
        sessionId: session.id,
        viewId,
        width: '1280',
        height: '900',
      })
    await owner.goto(url)
    await expect(owner.frameLocator('#editor-frame').locator('html')).toHaveAttribute(
      'data-live-editor-connection',
      'connected',
    )
    await call('office_open_local_file', {
      sessionId: session.id,
      baseRevision: 0,
      path: resolve('packages/pptx-engine/tests/fixtures/01_standard_business.pptx'),
    })
    const before = (await call('office_get_context', { sessionId: session.id })).session
    await call('office_execute', {
      sessionId: session.id,
      baseRevision: before.revision,
      requestId: 'unsaved-continue-here',
      operations: [{ id: 'pptx.slide.add_blank', arguments: { afterSlideIndex: 0 } }],
    })
    // Native hosts can hide/occlude a retained surface without removing its layout box.
    await owner.locator('#editor-frame').evaluate((frame) => {
      frame.style.visibility = 'hidden'
    })
    await returning.goto(url)
    const editor = returning.frameLocator('#editor-frame')
    await expect(editor.locator('html')).toHaveAttribute(
      'data-live-editor-connection',
      'editor_view_conflict',
    )
    await expect(owner.frameLocator('#editor-frame').locator('html')).toHaveAttribute(
      'data-live-editor-active',
      'true',
    )
    await editor.getByRole('button', { name: '在此继续编辑' }).click({ timeout: 3_000 })
    await expect(editor.locator('html')).toHaveAttribute('data-live-editor-connection', 'connected')
    const after = (await call('office_get_context', { sessionId: session.id })).session
    expect(after.selection.slideCount).toBe(6)
    expect(after.filePath).toBe(before.filePath)
    await expect(owner.frameLocator('#editor-frame').locator('html')).toHaveAttribute(
      'data-live-editor-connection',
      'suspended',
    )
    await expect(owner.frameLocator('#editor-frame').locator('#root')).toHaveAttribute('inert', '')
  } finally {
    await owner.goto('about:blank')
    await returning.close()
    await request.post(`/broker?id=${broker}&action=close`, { data: {} })
  }
})

test('Continue here is not lost while a connection retry is already in flight', async ({
  page: owner,
  context,
  request,
}) => {
  const broker = randomUUID()
  const call = async (name: string, args: Record<string, unknown>) => {
    const result = await (
      await request.post(`/broker?id=${broker}`, {
        data: { name, arguments: args },
      })
    ).json()
    expect(result.isError, JSON.stringify(result.structuredContent)).not.toBe(true)
    return result.structuredContent
  }
  const returning = await context.newPage()
  let releaseRetryPoll = () => undefined
  try {
    const { session } = await call('office_create_session', { format: 'pptx' })
    const { viewId } = await call('office_show_pptx_editor', { sessionId: session.id })
    const url =
      '/?' +
      new URLSearchParams({
        broker,
        format: 'pptx',
        sessionId: session.id,
        viewId,
        width: '1280',
        height: '900',
      })
    await owner.goto(url)
    await expect(owner.frameLocator('#editor-frame').locator('html')).toHaveAttribute(
      'data-live-editor-connection',
      'connected',
    )
    await call('office_open_local_file', {
      sessionId: session.id,
      baseRevision: 0,
      path: resolve('packages/pptx-engine/tests/fixtures/01_standard_business.pptx'),
    })
    const before = (await call('office_get_context', { sessionId: session.id })).session
    await call('office_execute', {
      sessionId: session.id,
      baseRevision: before.revision,
      requestId: 'retry-then-continue-unsaved',
      operations: [{ id: 'pptx.slide.add_blank', arguments: { afterSlideIndex: 0 } }],
    })

    let markRetryPollStarted = () => undefined
    const retryPollStarted = new Promise<void>((resolve) => {
      markRetryPollStarted = resolve
    })
    const retryPollGate = new Promise<void>((resolve) => {
      releaseRetryPoll = resolve
    })
    let heldRetryPoll = false
    await returning.route('**/broker?*', async (route) => {
      const body = route.request().postDataJSON()
      if (
        !heldRetryPoll &&
        body?.name === 'office_editor_poll' &&
        body.arguments?.retryHandoff === true &&
        body.arguments?.activateView !== true
      ) {
        heldRetryPoll = true
        markRetryPollStarted()
        await retryPollGate
      }
      await route.continue()
    })

    await returning.goto(url)
    const editor = returning.frameLocator('#editor-frame')
    await expect(editor.locator('html')).toHaveAttribute(
      'data-live-editor-connection',
      'editor_view_conflict',
    )
    await editor.getByRole('button', { name: '重试连接' }).click()
    await retryPollStarted
    await editor.getByRole('button', { name: '在此继续编辑' }).click()
    await expect(editor.locator('[data-live-session-status]')).toContainText(
      '已请求在此继续编辑，正在等待原编辑器安全交接…',
    )
    await expect(editor.getByRole('button', { name: '在此继续编辑' })).toBeDisabled()
    await expect(editor.getByRole('button', { name: '重试连接' })).toBeDisabled()
    releaseRetryPoll()

    await expect(editor.locator('html')).toHaveAttribute(
      'data-live-editor-connection',
      'connected',
      { timeout: 8_000 },
    )
    const after = (await call('office_get_context', { sessionId: session.id })).session
    expect(after.selection.slideCount).toBe(6)
    expect(after.filePath).toBe(before.filePath)
    await expect(owner.frameLocator('#editor-frame').locator('html')).toHaveAttribute(
      'data-live-editor-connection',
      'suspended',
    )
  } finally {
    releaseRetryPoll()
    await owner.goto('about:blank')
    await returning.close()
    await request.post(`/broker?id=${broker}&action=close`, { data: {} })
  }
})

test('PPTX: a hidden owner hands off automatically and requires explicit continuation to return', async ({
  page: owner,
  context,
  request,
}) => {
  const broker = randomUUID()
  const call = async (name: string, args: Record<string, unknown>) => {
    const response = await request.post(`/broker?id=${broker}`, {
      data: { name, arguments: args },
    })
    const result = await response.json()
    expect(result.isError, JSON.stringify(result.structuredContent)).not.toBe(true)
    return result.structuredContent
  }
  const returning = await context.newPage()
  try {
    const { session } = await call('office_create_session', { format: 'pptx' })
    const { viewId } = await call('office_show_pptx_editor', { sessionId: session.id })
    const url =
      '/?' +
      new URLSearchParams({
        broker,
        format: 'pptx',
        sessionId: session.id,
        viewId,
        width: '1280',
        height: '900',
      })
    await owner.goto(url)
    await expect(owner.frameLocator('#editor-frame').locator('html')).toHaveAttribute(
      'data-live-editor-connection',
      'connected',
    )
    await call('office_open_local_file', {
      sessionId: session.id,
      baseRevision: 0,
      path: resolve('packages/pptx-engine/tests/fixtures/01_standard_business.pptx'),
    })
    const before = (await call('office_get_context', { sessionId: session.id })).session
    expect(before.selection.slideCount).toBe(5)
    await call('office_execute', {
      sessionId: session.id,
      baseRevision: before.revision,
      requestId: 'unsaved-slide',
      operations: [{ id: 'pptx.slide.add_blank', arguments: { afterSlideIndex: 0 } }],
    })
    await owner.locator('#editor-frame').evaluate((frame) => {
      frame.style.display = 'none'
    })
    await returning.goto(url)
    await expect(returning.frameLocator('#editor-frame').locator('html')).toHaveAttribute(
      'data-live-editor-connection',
      'connected',
    )
    await expect
      .poll(
        async () =>
          (await call('office_get_context', { sessionId: session.id })).session.selection
            ?.slideCount,
      )
      .toBe(6)
    await expect
      .poll(() => returning.evaluate(() => window.__codexVisualHost.acknowledgements.length))
      .toBe(1)
    await expect(owner.frameLocator('#editor-frame').locator('html')).toHaveAttribute(
      'data-live-editor-connection',
      'suspended',
    )
    const resumed = (await call('office_get_context', { sessionId: session.id })).session
    await call('office_execute', {
      sessionId: session.id,
      baseRevision: resumed.revision,
      requestId: 'second-unsaved-slide',
      operations: [{ id: 'pptx.slide.add_blank', arguments: { afterSlideIndex: 0 } }],
    })
    await returning.locator('#editor-frame').evaluate((frame) => {
      frame.style.display = 'none'
    })
    await owner.locator('#editor-frame').evaluate((frame) => {
      frame.style.display = 'block'
    })
    await expect(owner.frameLocator('#editor-frame').locator('html')).toHaveAttribute(
      'data-live-editor-connection',
      'suspended',
    )
    await owner.frameLocator('#editor-frame').getByRole('button', { name: '在此继续编辑' }).click()
    await expect(owner.frameLocator('#editor-frame').locator('html')).toHaveAttribute(
      'data-live-editor-connection',
      'connected',
    )
    await expect
      .poll(
        async () =>
          (await call('office_get_context', { sessionId: session.id })).session.selection
            ?.slideCount,
      )
      .toBe(7)
  } finally {
    await owner.goto('about:blank')
    await returning.close()
    await request.post(`/broker?id=${broker}&action=close`, { data: {} })
  }
})

for (const format of ['docx', 'markdown', 'xlsx', 'pdf']) {
  test(`${format}: a replayed view restores the hidden owner's native document`, async ({
    page: owner,
    context,
    request,
  }) => {
    const broker = randomUUID()
    const directory = await mkdtemp(join(tmpdir(), 'tandemfolio-handoff-format-'))
    const source = join(directory, `Retained.${format === 'markdown' ? 'md' : format}`)
    await writeFile(source, await fixture(format))
    const call = async (name: string, args: Record<string, unknown>) => {
      const result = await (
        await request.post(`/broker?id=${broker}`, { data: { name, arguments: args } })
      ).json()
      expect(result.isError, JSON.stringify(result.structuredContent)).not.toBe(true)
      return result.structuredContent
    }
    const returning = await context.newPage()
    try {
      const { session } = await call('office_create_session', { format })
      const { viewId } = await call(
        format === 'docx' ? 'office_show_editor' : `office_show_${format}_editor`,
        { sessionId: session.id },
      )
      const url =
        '/?' +
        new URLSearchParams({
          broker,
          format,
          sessionId: session.id,
          viewId,
          width: '1280',
          height: '900',
        })
      await owner.goto(url)
      await expect(owner.frameLocator('#editor-frame').locator('html')).toHaveAttribute(
        'data-live-editor-connection',
        'connected',
      )
      await call('office_open_local_file', { sessionId: session.id, baseRevision: 0, path: source })
      const before = (await call('office_get_context', { sessionId: session.id })).session
      const textSelector = format === 'docx' ? '.editor-scroll .doc-page.ProseMirror' : '.tiptap'
      const nativeText = ['docx', 'markdown'].includes(format)
        ? await owner.frameLocator('#editor-frame').locator(textSelector).first().innerText()
        : null
      await owner.locator('#editor-frame').evaluate((frame) => {
        frame.style.display = 'none'
      })
      await returning.goto(url)
      const editor = returning.frameLocator('#editor-frame')
      await expect(editor.locator('html')).toHaveAttribute(
        'data-live-editor-connection',
        'connected',
      )
      await expect
        .poll(() => returning.evaluate(() => window.__codexVisualHost.acknowledgements.length))
        .toBe(1)
      const after = (await call('office_get_context', { sessionId: session.id })).session
      expect(after.fileName).toBe(before.fileName)
      expect(after.filePath).toBe(source)
      expect(after.selection).toEqual(before.selection)
      if (nativeText !== null)
        await expect(editor.locator(textSelector).first()).toHaveText(nativeText, {
          useInnerText: true,
        })
      else await expect(editor.locator('canvas').first()).toBeVisible()
      await expect(owner.frameLocator('#editor-frame').locator('html')).toHaveAttribute(
        'data-live-editor-connection',
        'suspended',
      )
      await editor.locator('[data-live-session-status] summary').click()
      await expect(editor.getByRole('textbox', { name: '文件绝对路径' })).toHaveValue(source)
      await returning.screenshot({ path: test.info().outputPath(`${format}-handoff.png`) })
    } finally {
      await owner.goto('about:blank')
      await returning.close()
      await request.post(`/broker?id=${broker}&action=close`, { data: {} })
      await rm(directory, { recursive: true, force: true })
    }
  })
}

test('a waiting view shows its saved path without acquiring the active editor lease', async ({
  page: owner,
  context,
  request,
}) => {
  const broker = randomUUID()
  const call = async (name: string, args: Record<string, unknown>) => {
    const result = await (
      await request.post(`/broker?id=${broker}`, {
        data: { name, arguments: args },
      })
    ).json()
    expect(result.isError, JSON.stringify(result.structuredContent)).not.toBe(true)
    return result.structuredContent
  }
  const waiting = await context.newPage()
  try {
    const { session } = await call('office_create_session', { format: 'pptx' })
    const { viewId } = await call('office_show_pptx_editor', { sessionId: session.id })
    const url =
      '/?' +
      new URLSearchParams({
        broker,
        format: 'pptx',
        sessionId: session.id,
        viewId,
        width: '1280',
        height: '900',
      })
    await owner.goto(url)
    await expect(owner.frameLocator('#editor-frame').locator('html')).toHaveAttribute(
      'data-live-editor-connection',
      'connected',
    )
    const path = resolve('packages/pptx-engine/tests/fixtures/01_standard_business.pptx')
    await call('office_open_local_file', { sessionId: session.id, baseRevision: 0, path })
    await waiting.goto(url)
    const editor = waiting.frameLocator('#editor-frame')
    await expect(editor.locator('html')).toHaveAttribute(
      'data-live-editor-connection',
      'editor_view_conflict',
    )
    await expect(editor.locator('[data-live-session-status] summary')).toBeVisible({
      timeout: 3_000,
    })
    await editor.locator('[data-live-session-status] summary').click()
    await expect(editor.getByRole('textbox', { name: '文件绝对路径' })).toHaveValue(path)
    await expect(editor.locator('#root')).toHaveAttribute('inert', '')
    expect(
      (await call('office_get_context', { sessionId: session.id })).session.selection.slideCount,
    ).toBe(5)
    await waiting.locator('#editor-frame').evaluate((frame) => {
      frame.style.display = 'none'
    })
    await waiting.waitForTimeout(2500)
    await owner.locator('#editor-frame').evaluate((frame) => {
      frame.style.display = 'none'
    })
    await waiting.locator('#editor-frame').evaluate((frame) => {
      frame.style.display = 'block'
    })
    await expect(editor.locator('html')).toHaveAttribute(
      'data-live-editor-connection',
      'connected',
      { timeout: 8_000 },
    )
  } finally {
    await owner.goto('about:blank')
    await waiting.close()
    await request.post(`/broker?id=${broker}&action=close`, { data: {} })
  }
})

test('checkpoint failure retains the original editor and stops automatic handoff until explicit retry', async ({
  page: owner,
  context,
  request,
}) => {
  const broker = randomUUID()
  const call = async (name: string, args: Record<string, unknown>) => {
    const result = await (
      await request.post(`/broker?id=${broker}`, { data: { name, arguments: args } })
    ).json()
    expect(result.isError, JSON.stringify(result.structuredContent)).not.toBe(true)
    return result.structuredContent
  }
  const waiting = await context.newPage()
  let failures = 0
  try {
    const { session } = await call('office_create_session', { format: 'pptx' })
    const { viewId } = await call('office_show_pptx_editor', { sessionId: session.id })
    const url =
      '/?' +
      new URLSearchParams({
        broker,
        format: 'pptx',
        sessionId: session.id,
        viewId,
        width: '1280',
        height: '900',
      })
    const path = resolve('packages/pptx-engine/tests/fixtures/01_standard_business.pptx')
    const original = await readFile(path)
    await owner.goto(url)
    await expect(owner.frameLocator('#editor-frame').locator('html')).toHaveAttribute(
      'data-live-editor-connection',
      'connected',
    )
    await call('office_open_local_file', { sessionId: session.id, baseRevision: 0, path })
    await owner.route('**/broker?*', async (route) => {
      if (route.request().postDataJSON().name !== 'office_editor_commit_recovery')
        return route.continue()
      failures += 1
      await route.fulfill({
        json: {
          isError: true,
          structuredContent: {
            ok: false,
            error: 'execution_failed',
            message: 'Injected disk failure',
          },
        },
      })
    })
    await owner.locator('#editor-frame').evaluate((frame) => {
      frame.style.display = 'none'
    })
    await waiting.goto(url)
    const editor = waiting.frameLocator('#editor-frame')
    await expect(editor.locator('[data-live-session-status]')).toContainText('交接失败', {
      timeout: 8_000,
    })
    await expect(editor.locator('#root')).toHaveAttribute('inert', '')
    expect(
      (await call('office_get_context', { sessionId: session.id })).session.selection.slideCount,
    ).toBe(5)
    await waiting.waitForTimeout(6000)
    expect(failures).toBe(1)
    expect(await readFile(path)).toEqual(original)
    await owner.unroute('**/broker?*')
    await editor.getByRole('button', { name: '重试连接' }).click()
    await expect(editor.locator('html')).toHaveAttribute('data-live-editor-connection', 'connected')
    expect(
      (await call('office_get_context', { sessionId: session.id })).session.selection.slideCount,
    ).toBe(5)
  } finally {
    await owner.goto('about:blank')
    await waiting.close()
    await request.post(`/broker?id=${broker}&action=close`, { data: {} })
  }
})

test('an abandoned owner restores after bounded waiting and explicit retry without creating another Session', async ({
  page: owner,
  context,
  request,
}) => {
  test.setTimeout(65_000)
  const broker = randomUUID()
  const call = async (name: string, args: Record<string, unknown>) => {
    const result = await (
      await request.post(`/broker?id=${broker}`, { data: { name, arguments: args } })
    ).json()
    expect(result.isError, JSON.stringify(result.structuredContent)).not.toBe(true)
    return result.structuredContent
  }
  const waiting = await context.newPage()
  try {
    const { session } = await call('office_create_session', { format: 'pptx' })
    const { viewId } = await call('office_show_pptx_editor', { sessionId: session.id })
    const url =
      '/?' +
      new URLSearchParams({
        broker,
        format: 'pptx',
        sessionId: session.id,
        viewId,
        width: '1280',
        height: '900',
      })
    await owner.goto(url)
    await expect(owner.frameLocator('#editor-frame').locator('html')).toHaveAttribute(
      'data-live-editor-connection',
      'connected',
    )
    await call('office_open_local_file', {
      sessionId: session.id,
      baseRevision: 0,
      path: resolve('packages/pptx-engine/tests/fixtures/01_standard_business.pptx'),
    })
    await waiting.goto(url)
    const editor = waiting.frameLocator('#editor-frame')
    await expect(editor.locator('html')).toHaveAttribute(
      'data-live-editor-connection',
      'editor_view_conflict',
    )
    await owner.goto('about:blank')
    await expect(editor.locator('[data-live-session-status]')).toContainText('等待已超时', {
      timeout: 35_000,
    })
    await editor.getByRole('button', { name: '在此继续编辑' }).click()
    await expect(editor.locator('html')).toHaveAttribute(
      'data-live-editor-connection',
      'connected',
      { timeout: 40_000 },
    )
    expect(
      (await call('office_get_context', { sessionId: session.id })).session.selection.slideCount,
    ).toBe(5)
    expect(await waiting.evaluate(() => window.__codexVisualHost.polls)).toBeLessThan(16)
  } finally {
    await waiting.close()
    await request.post(`/broker?id=${broker}&action=close`, { data: {} })
  }
})
