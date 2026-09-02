import { expect, test } from '@playwright/test'

test('Markdown exposes file, save, and fullscreen as accessible icon-only controls', async ({
  page,
}) => {
  await page.goto('/?format=markdown&width=720&height=900')
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)

  const editor = page.frameLocator('#editor-frame')
  const controls = [
    editor.getByRole('button', { name: '打开 Markdown', exact: true }),
    editor.getByRole('button', { name: /^保存/ }),
    editor.getByRole('button', { name: '全屏', exact: true }),
  ]

  for (const control of controls) {
    await expect(control).toBeVisible()
    await expect(control.locator('svg')).toHaveCount(1)
    await expect(control).toHaveText('')
  }
})

test('a generated Markdown file saves through the session-bound local persistence protocol', async ({
  page,
}) => {
  await page.goto('/?format=markdown&width=720&height=900')
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)

  await page.evaluate(() => {
    window.__codexVisualHost.enqueueCommand({
      commandId: 'markdown-save-content',
      baseRevision: 0,
      operation: 'markdown.text.insert',
      arguments: { text: '# Quarterly Review' },
    })
  })
  await expect
    .poll(() => page.evaluate(() => window.__codexVisualHost.acknowledgements.length))
    .toBe(1)

  const downloadPromise = page.waitForEvent('download', { timeout: 5_000 })
  await page.evaluate(() => {
    window.__codexVisualHost.enqueueCommand({
      commandId: 'markdown-save',
      baseRevision: 1,
      operation: 'markdown.document.save',
      arguments: {},
    })
  })
  await expect
    .poll(() => page.evaluate(() => window.__codexVisualHost.acknowledgements.length))
    .toBe(2)
  await expect
    .poll(() => page.evaluate(() => window.__codexVisualHost.acknowledgements.at(-1)))
    .toMatchObject({ commandId: 'markdown-save', ok: true })
  const download = await downloadPromise

  expect(download.suggestedFilename()).toMatch(/\.md$/)
  await expect
    .poll(() => page.evaluate(() => window.__codexVisualHost.downloads.at(-1)))
    .toMatchObject({
      fileName: download.suggestedFilename(),
      mimeType: 'text/markdown;charset=utf-8',
      size: expect.any(Number),
    })
  expect(await page.evaluate(() => window.__codexVisualHost.events)).toEqual(
    expect.arrayContaining([
      'tool:office_editor_begin_document_save',
      'tool:office_editor_commit_document_save',
    ]),
  )
})

test('the packaged host holds one bounded poll and wakes it without a 500 ms interval', async ({
  page,
}) => {
  await page.goto('/?format=markdown&width=720&height=900')
  await page.waitForFunction(() => {
    const state = window.__codexVisualHost
    return state?.initialized && state.firstPollAt != null
  })

  const idlePollDelta = await page.evaluate(async () => {
    const before = window.__codexVisualHost.polls
    await new Promise((resolve) => setTimeout(resolve, 150))
    return window.__codexVisualHost.polls - before
  })
  expect(idlePollDelta).toBeLessThanOrEqual(1)

  await page.evaluate(() => {
    window.__codexVisualHost.enqueueCommand({
      commandId: 'markdown-wakeup',
      baseRevision: 0,
      operation: 'markdown.text.insert',
      arguments: { text: 'wake immediately' },
    })
  })
  await page.waitForFunction(() =>
    window.__codexVisualHost.acknowledgements.some(
      (entry) => entry.commandId === 'markdown-wakeup' && entry.ok === true,
    ),
  )

  const timing = await page.evaluate(
    () => window.__codexVisualHost.commandTimings['markdown-wakeup'],
  )
  expect(timing.pollWaitMs).toBeLessThan(100)
})

test('an offscreen Markdown editor releases its document DOM and resumes the same session', async ({
  page,
}) => {
  await page.goto('/?format=markdown&width=720&height=900')
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)

  const editor = page.frameLocator('#editor-frame')
  await page.evaluate(() => {
    window.__codexVisualHost.enqueueCommand({
      commandId: 'markdown-lifecycle-content',
      baseRevision: 0,
      operation: 'markdown.text.insert',
      arguments: { text: 'Retained while offscreen' },
    })
  })
  await expect(editor.locator('.doc-editor')).toContainText('Retained while offscreen')

  await editor.locator('#root').evaluate((root) => {
    root.dispatchEvent(
      Object.assign(new Event('contentvisibilityautostatechange'), { skipped: true }),
    )
  })
  await expect(editor.locator('.editor-workspace-suspended')).toHaveCount(1)
  await expect(editor.locator('.doc-editor')).toHaveCount(0)

  await editor.locator('#root').evaluate((root) => {
    root.dispatchEvent(
      Object.assign(new Event('contentvisibilityautostatechange'), { skipped: false }),
    )
  })
  await expect(editor.locator('.doc-editor')).toContainText('Retained while offscreen')

  const state = await page.evaluate(() => window.__codexVisualHost)
  expect(state.editorLoads).toBe(1)
  expect(state.errors).toEqual([])
})

test('Markdown stores each dirty document version only once', async ({ page }) => {
  await page.goto('/?format=markdown&width=720&height=900')
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)

  await page.evaluate(() => {
    window.__codexVisualHost.enqueueCommand({
      commandId: 'markdown-recovery-v1',
      baseRevision: 0,
      operation: 'markdown.text.insert',
      arguments: { text: 'first version' },
    })
  })
  await expect.poll(() => page.evaluate(() => window.__codexVisualHost.recoveryCommits)).toBe(1)

  await page.waitForTimeout(2_300)
  expect(await page.evaluate(() => window.__codexVisualHost.recoveryCommits)).toBe(1)

  await page.evaluate(() => {
    window.__codexVisualHost.enqueueCommand({
      commandId: 'markdown-recovery-v2',
      baseRevision: 1,
      operation: 'markdown.text.insert',
      arguments: { text: ' second version' },
    })
  })
  await expect.poll(() => page.evaluate(() => window.__codexVisualHost.recoveryCommits)).toBe(2)
})

test('Markdown exposes reproducible cold-start, staged-open, interaction, ack, and heap metrics', async ({
  page,
}) => {
  await page.goto('/?format=markdown&width=720&height=900')
  await page.waitForFunction(() => {
    const state = window.__codexVisualHost
    return state?.initialized && state.firstPollAt != null
  })

  await page.evaluate(() => {
    const bytes = new TextEncoder().encode('# Release baseline\n\nSmall deterministic fixture.')
    let binary = ''
    for (const byte of bytes) binary += String.fromCharCode(byte)
    window.__codexVisualHost.enqueueCommand(
      {
        commandId: 'markdown-open-small',
        baseRevision: 0,
        operation: 'markdown.document.load_staged',
        arguments: { blobId: 'markdown-small', name: 'small.md', size: bytes.byteLength },
      },
      { blobId: 'markdown-small', base64: btoa(binary) },
    )
  })

  await page.waitForFunction(() =>
    window.__codexVisualHost.acknowledgements.some(
      (entry) => entry.commandId === 'markdown-open-small' && entry.ok === true,
    ),
  )
  await expect(page.frameLocator('#editor-frame').locator('.doc-editor')).toContainText(
    'Small deterministic fixture.',
  )

  await page.evaluate(() => {
    window.__codexVisualHost.enqueueCommand({
      commandId: 'markdown-insert',
      baseRevision: 1,
      operation: 'markdown.text.insert',
      arguments: { text: ' measured interaction' },
    })
  })
  await page.waitForFunction(() =>
    window.__codexVisualHost.acknowledgements.some(
      (entry) => entry.commandId === 'markdown-insert' && entry.ok === true,
    ),
  )

  const metrics = await page.evaluate(() => {
    const state = window.__codexVisualHost
    return {
      coldStartMs: state.firstPollAt! - state.startedAt,
      open: state.commandTimings['markdown-open-small'],
      interaction: state.commandTimings['markdown-insert'],
      usedJsHeapBytes: performance.memory?.usedJSHeapSize,
    }
  })

  expect(metrics.coldStartMs).toBeGreaterThan(0)
  expect(metrics.usedJsHeapBytes).toBeGreaterThan(0)
  for (const timing of [metrics.open, metrics.interaction]) {
    expect(timing.totalMs).toBeGreaterThanOrEqual(0)
    expect(timing.pollWaitMs).toBeGreaterThanOrEqual(0)
    expect(timing.hydrateMs).toBeGreaterThanOrEqual(0)
    expect(timing.executeMs).toBeGreaterThanOrEqual(0)
    expect(timing.ackTransportMs).toBeGreaterThanOrEqual(0)
    expect(
      timing.pollWaitMs + timing.hydrateMs + timing.executeMs + timing.ackTransportMs,
    ).toBeCloseTo(timing.totalMs, 4)
  }
})

test('Markdown staged load acknowledges four phases after the committed document is visible', async ({
  page,
}) => {
  await page.goto('/?format=markdown&width=720&height=900')
  await page.waitForFunction(() => {
    const state = window.__codexVisualHost
    return state?.initialized && state.firstPollAt != null
  })

  await page.evaluate(() => {
    const bytes = new TextEncoder().encode('# Traced load\n\nCommitted content.')
    let binary = ''
    for (const byte of bytes) binary += String.fromCharCode(byte)
    window.__codexVisualHost.enqueueCommand(
      {
        commandId: 'markdown-traced-load',
        baseRevision: 0,
        operation: 'markdown.document.load_staged',
        arguments: {
          blobId: 'markdown-traced-load-blob',
          name: 'traced.md',
          size: bytes.byteLength,
        },
      },
      { blobId: 'markdown-traced-load-blob', base64: btoa(binary) },
    )
  })

  await page.waitForFunction(() =>
    window.__codexVisualHost.acknowledgements.some(
      (entry) => entry.commandId === 'markdown-traced-load' && entry.ok === true,
    ),
  )

  const acknowledgement = await page.evaluate(() => ({
    payload: window.__codexVisualHost.acknowledgements.find(
      (entry) => entry.commandId === 'markdown-traced-load',
    ),
    snapshot: window.__codexVisualHost.acknowledgementSnapshots['markdown-traced-load'],
  }))
  const timing = acknowledgement.payload?.timing as
    | {
        executeMs?: number
        trace?: {
          operation?: string
          phases?: Record<string, number>
        }
      }
    | undefined

  expect(timing?.trace).toEqual({
    operation: 'markdown.document.load_staged',
    phases: {
      decodeMs: expect.any(Number),
      parseMs: expect.any(Number),
      tiptapStateInstallMs: expect.any(Number),
      reactCommitMs: expect.any(Number),
    },
  })
  const phases = timing!.trace!.phases!
  for (const value of Object.values(phases)) expect(value).toBeGreaterThanOrEqual(0)
  expect(Object.values(phases).reduce((sum, value) => sum + value, 0)).toBeLessThanOrEqual(
    timing!.executeMs!,
  )
  expect(acknowledgement.snapshot).toEqual({
    editorText: 'Traced loadCommitted content.',
    fileName: 'traced.md',
    ready: true,
  })
})

test('Markdown opens the canonical large fixture within five seconds', async ({ page }) => {
  await page.goto('/?format=markdown&width=720&height=900')
  await page.waitForFunction(() => {
    const state = window.__codexVisualHost
    return state?.initialized && state.firstPollAt != null
  })

  await page.evaluate(() => {
    const sections = Array.from(
      { length: 12_000 },
      (_, index) =>
        `## Section ${index + 1}\n\nDeterministic release fixture paragraph ${index + 1}.`,
    )
    const bytes = new TextEncoder().encode(
      `# TandemFolio large fixture\n\n${sections.join('\n\n')}`,
    )
    let binary = ''
    for (let offset = 0; offset < bytes.length; offset += 32_768) {
      binary += String.fromCharCode(...bytes.subarray(offset, offset + 32_768))
    }
    window.__codexVisualHost.enqueueCommand(
      {
        commandId: 'markdown-large-five-second-gate',
        baseRevision: 0,
        operation: 'markdown.document.load_staged',
        arguments: {
          blobId: 'markdown-large-five-second-blob',
          name: 'large.md',
          size: bytes.byteLength,
        },
      },
      { blobId: 'markdown-large-five-second-blob', base64: btoa(binary) },
    )
  })

  await page.waitForFunction(
    () =>
      window.__codexVisualHost.acknowledgements.some(
        (entry) => entry.commandId === 'markdown-large-five-second-gate' && entry.ok === true,
      ),
    undefined,
    { timeout: 30_000 },
  )

  const result = await page.evaluate(() => ({
    snapshot: window.__codexVisualHost.acknowledgementSnapshots['markdown-large-five-second-gate'],
    timing: window.__codexVisualHost.commandTimings['markdown-large-five-second-gate'],
  }))
  expect(result.snapshot.fileName).toBe('large.md')
  expect(result.snapshot.ready).toBe(true)
  expect(result.snapshot.editorText).toContain('TandemFolio large fixture')
  expect(result.snapshot.editorText).toContain('Deterministic release fixture paragraph 12000.')
  expect(result.timing.executeMs).toBeLessThanOrEqual(5_000)
})
