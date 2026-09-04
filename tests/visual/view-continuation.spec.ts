import {
  expect,
  test,
  type Page,
  type BrowserContext,
  type APIRequestContext,
} from '@playwright/test'
import { randomUUID } from 'node:crypto'
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises'
import { join } from 'node:path'
import { tmpdir } from 'node:os'
import { fixture } from './document-fixtures'
import JSZip from 'jszip'

async function openReplicas(
  owner: Page,
  context: BrowserContext,
  request: APIRequestContext,
  format = 'pptx',
) {
  const broker = randomUUID()
  const directory = await mkdtemp(join(tmpdir(), 'tandemfolio-continuation-'))
  const source = join(directory, `Retained.${format === 'markdown' ? 'md' : format}`)
  const waiting = await context.newPage()
  const call = async (name: string, args: Record<string, unknown>) => {
    const result = await (
      await request.post(`/broker?id=${broker}`, { data: { name, arguments: args } })
    ).json()
    expect(result.isError, JSON.stringify(result.structuredContent)).not.toBe(true)
    return result.structuredContent
  }
  const close = async () => {
    await owner.goto('about:blank')
    await waiting.close()
    await request.post(`/broker?id=${broker}&action=close`, { data: {} })
    await rm(directory, { recursive: true, force: true })
  }
  try {
    await writeFile(source, await fixture(format))
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
    await waiting.goto(url)
    const editor = waiting.frameLocator('#editor-frame')
    await expect(editor.locator('html')).toHaveAttribute(
      'data-live-editor-connection',
      'editor_view_conflict',
    )
    return { call, broker, sessionId: session.id, source, before, waiting, editor, close, url }
  } catch (error) {
    await close()
    throw error
  }
}

test('PPTX continuation rejects an implicit document reset without losing the original file binding', async ({
  page: owner,
  context,
  request,
}) => {
  const replicas = await openReplicas(owner, context, request)
  try {
    const savedBytes = await readFile(replicas.source)
    await replicas.editor.getByRole('button', { name: '在此继续编辑' }).click()
    await expect(replicas.editor.locator('html')).toHaveAttribute(
      'data-live-editor-connection',
      'connected',
    )
    const before = (await replicas.call('office_get_context', { sessionId: replicas.sessionId }))
      .session
    const reset = await (
      await request.post(`/broker?id=${replicas.broker}`, {
        data: {
          name: 'office_execute',
          arguments: {
            sessionId: replicas.sessionId,
            baseRevision: before.revision,
            requestId: 'implicit-reset-after-continuation',
            operations: [{ id: 'pptx.document.create_blank', arguments: {} }],
          },
        },
      })
    ).json()
    expect(reset.isError, JSON.stringify(reset.structuredContent)).toBe(true)
    expect(JSON.stringify(reset.structuredContent)).toContain('confirmReplace')
    const after = (await replicas.call('office_get_context', { sessionId: replicas.sessionId }))
      .session
    expect(after.revision).toBe(before.revision)
    expect(after.fileName).toBe(before.fileName)
    expect(after.filePath).toBe(replicas.source)
    expect(after.selection).toEqual(before.selection)
    expect(await readFile(replicas.source)).toEqual(savedBytes)

    // A rejected reset must not poison the next edit or detach the original save destination.
    await replicas.call('office_execute', {
      sessionId: replicas.sessionId,
      baseRevision: after.revision,
      requestId: 'continue-editing-original',
      operations: [{ id: 'pptx.slide.add_blank', arguments: { afterSlideIndex: 0 } }],
    })
    const edited = (await replicas.call('office_get_context', { sessionId: replicas.sessionId }))
      .session
    expect(edited.selection.slideCount).toBe(before.selection.slideCount + 1)
    await replicas.call('office_execute', {
      sessionId: replicas.sessionId,
      baseRevision: edited.revision,
      requestId: 'save-continued-original',
      operations: [{ id: 'pptx.document.save', arguments: {} }],
    })
    const saved = (await replicas.call('office_get_context', { sessionId: replicas.sessionId }))
      .session
    expect(saved.filePath).toBe(replicas.source)
    expect(saved.fileName).toBe(before.fileName)
    const zip = await JSZip.loadAsync(await readFile(replicas.source))
    expect(zip.file(/^ppt\/slides\/slide\d+\.xml$/)).toHaveLength(before.selection.slideCount + 1)
  } finally {
    await replicas.close()
  }
})

test('PPTX explicit replacement starts a separate save target and leaves the original file intact', async ({
  page: owner,
  context,
  request,
}) => {
  const replicas = await openReplicas(owner, context, request)
  try {
    const originalBytes = await readFile(replicas.source)
    await replicas.editor.getByRole('button', { name: '在此继续编辑' }).click()
    await expect(replicas.editor.locator('html')).toHaveAttribute(
      'data-live-editor-connection',
      'connected',
    )
    const before = (await replicas.call('office_get_context', { sessionId: replicas.sessionId }))
      .session
    await replicas.call('office_execute', {
      sessionId: replicas.sessionId,
      baseRevision: before.revision,
      requestId: 'confirmed-new-document',
      operations: [{ id: 'pptx.document.create_blank', arguments: { confirmReplace: true } }],
    })
    const blank = (await replicas.call('office_get_context', { sessionId: replicas.sessionId }))
      .session
    expect(blank.filePath).toBeNull()
    expect(blank.selection.slideCount).toBe(1)
    await replicas.call('office_execute', {
      sessionId: replicas.sessionId,
      baseRevision: blank.revision,
      requestId: 'save-confirmed-new-document',
      operations: [{ id: 'pptx.document.save', arguments: {} }],
    })
    const saved = (await replicas.call('office_get_context', { sessionId: replicas.sessionId }))
      .session
    expect(saved.fileName).toBe('Untitled.pptx')
    expect(saved.filePath).not.toBe(replicas.source)
    expect((await readFile(saved.filePath)).byteLength).toBeGreaterThan(20)
    expect(await readFile(replicas.source)).toEqual(originalBytes)
  } finally {
    await replicas.close()
  }
})

test('an uncertain commit reply never unlocks the outgoing editor', async ({
  page: owner,
  context,
  request,
}) => {
  const replicas = await openReplicas(owner, context, request)
  try {
    await owner.route('**/broker?*', async (route) => {
      const input = route.request().postDataJSON()
      if (input.name !== 'office_editor_handoff' || input.arguments.action !== 'commit')
        return route.continue()
      // The real Broker commits, but its successful reply does not reach the old renderer.
      await route.fetch()
      await route.fulfill({
        json: { isError: true, structuredContent: { ok: false, message: 'Lost commit reply' } },
      })
    })
    await replicas.editor.getByRole('button', { name: '在此继续编辑' }).click()
    await expect(replicas.editor.locator('html')).toHaveAttribute(
      'data-live-editor-connection',
      'connected',
    )
    const original = owner.frameLocator('#editor-frame')
    await expect(original.locator('[data-live-session-status]')).toContainText('交接结果尚未确认')
    await expect(original.locator('#root')).toHaveAttribute('inert', '')
    expect(
      (await replicas.call('office_get_context', { sessionId: replicas.sessionId })).session
        .selection,
    ).toEqual(replicas.before.selection)
    await owner.unroute('**/broker?*')
    await original.getByRole('button', { name: '重试连接' }).click({ timeout: 3_000 })
    await expect(original.locator('html')).toHaveAttribute(
      'data-live-editor-connection',
      'suspended',
    )
    await expect(original.locator('#root')).toHaveAttribute('inert', '')
    await original.getByRole('button', { name: '在此继续编辑' }).click()
    await expect(original.locator('html')).toHaveAttribute(
      'data-live-editor-connection',
      'connected',
    )
    expect(
      (await replicas.call('office_get_context', { sessionId: replicas.sessionId })).session
        .selection,
    ).toEqual(replicas.before.selection)
  } finally {
    await replicas.close()
  }
})

test('an uncertain prepare can resume the original document only after ownership is confirmed', async ({
  page: owner,
  context,
  request,
}) => {
  const replicas = await openReplicas(owner, context, request)
  try {
    await owner.route('**/broker?*', async (route) => {
      const input = route.request().postDataJSON()
      if (input.name !== 'office_editor_handoff') return route.continue()
      if (input.arguments.action === 'prepare') await route.fetch()
      await route.fulfill({
        json: {
          isError: true,
          structuredContent: { ok: false, message: 'Handoff transport unavailable' },
        },
      })
    })
    await replicas.editor.getByRole('button', { name: '在此继续编辑' }).click()
    const original = owner.frameLocator('#editor-frame')
    await expect(original.locator('html')).toHaveAttribute(
      'data-live-editor-connection',
      'handoff_uncertain',
    )
    await expect(original.locator('#root')).toHaveAttribute('inert', '')
    await expect(replicas.editor.locator('#root')).toHaveAttribute('inert', '')
    await owner.unroute('**/broker?*')
    await original.getByRole('button', { name: '重试连接' }).click()
    await expect(original.locator('html')).toHaveAttribute(
      'data-live-editor-connection',
      'connected',
    )
    await expect(original.locator('#root')).not.toHaveAttribute('inert', '')
    const after = (await replicas.call('office_get_context', { sessionId: replicas.sessionId }))
      .session
    expect(after.selection).toEqual(replicas.before.selection)
    expect(after.filePath).toBe(replicas.source)
    await expect(replicas.editor.locator('#root')).toHaveAttribute('inert', '')
  } finally {
    await replicas.close()
  }
})

for (const format of ['docx', 'markdown', 'xlsx', 'pptx', 'pdf']) {
  test(`${format}: explicit continuation restores the chosen document and keeps the old surface suspended`, async ({
    page: owner,
    context,
    request,
  }) => {
    const replicas = await openReplicas(owner, context, request, format)
    try {
      const original = owner.frameLocator('#editor-frame')
      const textSelector = format === 'docx' ? '.editor-scroll .doc-page.ProseMirror' : '.tiptap'
      const originalText = ['docx', 'markdown'].includes(format)
        ? await original.locator(textSelector).first().innerText()
        : null
      await replicas.editor.getByRole('button', { name: '在此继续编辑' }).click()
      await expect(replicas.editor.locator('html')).toHaveAttribute(
        'data-live-editor-connection',
        'connected',
      )
      const after = (await replicas.call('office_get_context', { sessionId: replicas.sessionId }))
        .session
      expect(after.selection).toEqual(replicas.before.selection)
      expect(after.filePath).toBe(replicas.source)
      if (originalText !== null)
        await expect(replicas.editor.locator(textSelector).first()).toHaveText(originalText, {
          useInnerText: true,
        })
      else await expect(replicas.editor.locator('canvas').first()).toBeVisible()
      await expect(original.locator('html')).toHaveAttribute(
        'data-live-editor-connection',
        'suspended',
      )
      await owner.locator('#editor-frame').evaluate((frame) => {
        frame.style.display = 'none'
      })
      await owner.locator('#editor-frame').evaluate((frame) => {
        frame.style.display = 'block'
      })
      await replicas.waiting.waitForTimeout(2_000)
      await expect(original.locator('html')).toHaveAttribute(
        'data-live-editor-connection',
        'suspended',
      )
      await expect(original.locator('#root')).toHaveAttribute('inert', '')
      await expect(replicas.editor.locator('html')).toHaveAttribute(
        'data-live-editor-connection',
        'connected',
      )
    } finally {
      await replicas.close()
    }
  })
}
