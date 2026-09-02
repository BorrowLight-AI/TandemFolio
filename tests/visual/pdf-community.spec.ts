import { expect, test } from '@playwright/test'
import type { Page } from '@playwright/test'
import { readFile } from 'node:fs/promises'
import { PDFArray, PDFDict, PDFDocument, PDFName, StandardFonts } from 'pdf-lib'
import { getDocument, OPS } from 'pdfjs-dist/legacy/build/pdf.mjs'

const NARROW = { width: 420, height: 900 }
const SPLIT = { width: 720, height: 900 }
const FULLSCREEN = { width: 1332, height: 1280 }

test('PDF exposes file, save, and fullscreen as accessible icon-only controls', async ({
  page,
}) => {
  await page.goto('/?format=pdf&width=1332&height=900')
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)

  const editor = page.frameLocator('#editor-frame')
  await editor.locator('input[type="file"][accept*=".pdf"]').setInputFiles({
    name: 'toolbar-contract.pdf',
    mimeType: 'application/pdf',
    buffer: await fixture(),
  })
  await expect(editor.locator('.ribbon')).toBeVisible()
  const controls = [
    editor.getByRole('button', { name: '打开 PDF', exact: true }),
    editor.getByRole('button', { name: /^保存/ }),
    editor.getByRole('button', { name: '全屏', exact: true }),
  ]

  for (const control of controls) {
    await expect(control).toBeVisible()
    await expect(control.locator('svg')).toHaveCount(1)
    await expect(control).toHaveText('')
  }
})

/** 1x1 red pixel PNG. */
const TINY_PNG =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=='

function hostUrl(dimensions: { width: number; height: number }, acceptFullscreen = false): string {
  const params = new URLSearchParams({
    format: 'pdf',
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

async function fixture(): Promise<Buffer> {
  const document = await PDFDocument.create()
  document.setTitle('Community PDF renderer')
  const font = await document.embedFont(StandardFonts.Helvetica)
  const page = document.addPage([420, 300])
  page.drawText('Original community PDF content', { x: 30, y: 250, size: 18, font })
  return Buffer.from(await document.save())
}

async function annotatedFixture(): Promise<{
  buffer: Buffer
  deletion: { pageIndex: number; objNum: number; subtype: string; rect: number[] }
}> {
  const document = await PDFDocument.create()
  const page = document.addPage([420, 300])
  const annotations = document.context.obj([]) as PDFArray
  page.node.set(PDFName.of('Annots'), annotations)
  const add = (subtype: 'Highlight' | 'Underline', rect: [number, number, number, number]) => {
    const reference = document.context.register(
      document.context.obj({
        Type: 'Annot',
        Subtype: subtype,
        Rect: rect,
        QuadPoints: [rect[0], rect[3], rect[2], rect[3], rect[0], rect[1], rect[2], rect[1]],
        C: [1, 0.87, 0.35],
        F: 4,
        P: page.ref,
      }),
    )
    annotations.push(reference)
    return reference
  }
  const highlight = add('Highlight', [40, 220, 180, 242])
  add('Underline', [40, 180, 180, 202])
  return {
    buffer: Buffer.from(await document.save({ useObjectStreams: false })),
    deletion: {
      pageIndex: 0,
      objNum: highlight.objectNumber,
      subtype: 'highlight',
      rect: [40, 220, 180, 242],
    },
  }
}

async function openFixture(page: Page): Promise<void> {
  const editor = page.frameLocator('#editor-frame')
  await editor.locator('input[type="file"][accept*=".pdf"]').setInputFiles({
    name: 'community.pdf',
    mimeType: 'application/pdf',
    buffer: await fixture(),
  })
  await expect(editor.locator('.pdf-page-content canvas').first()).toBeVisible()
}

async function executePdfCommand(
  page: Page,
  command: {
    commandId: string
    baseRevision: number
    operation: string
    arguments: Record<string, unknown>
  },
): Promise<Record<string, unknown>> {
  await page.evaluate((next) => window.__codexVisualHost.commands.push(next), command)
  await expect
    .poll(() =>
      page.evaluate((commandId) => {
        return window.__codexVisualHost.acknowledgements.find(
          (acknowledgement) => acknowledgement.commandId === commandId,
        )
      }, command.commandId),
    )
    .not.toBeUndefined()
  const acknowledgement = await page.evaluate((commandId) => {
    return window.__codexVisualHost.acknowledgements.find(
      (candidate) => candidate.commandId === commandId,
    )!
  }, command.commandId)
  if (acknowledgement.ok !== true) {
    throw new Error(`Unexpected acknowledgement: ${JSON.stringify(acknowledgement)}`)
  }
  return acknowledgement
}

async function savePdfCommand(
  page: Page,
  commandId: string,
  baseRevision: number,
): Promise<Buffer> {
  const downloadPromise = page.waitForEvent('download')
  const acknowledgementPromise = executePdfCommand(page, {
    commandId,
    baseRevision,
    operation: 'pdf.document.save',
    arguments: {},
  })
  const download = await downloadPromise
  const acknowledgement = await acknowledgementPromise
  expect(acknowledgement).toMatchObject({ ok: true, dirty: false })
  const savedPath = await download.path()
  expect(savedPath).not.toBeNull()
  return readFile(savedPath!)
}

async function pdfPageEvidence(bytes: Uint8Array): Promise<{
  text: string
  imageCount: number
  fillColors: string[]
  annotationSubtypes: string[]
}> {
  const loadingTask = getDocument({ data: Uint8Array.from(bytes) })
  try {
    const document = await loadingTask.promise
    const page = await document.getPage(1)
    const content = await page.getTextContent()
    const operators = await page.getOperatorList()
    const annotations = await page.getAnnotations()
    const fillColors = operators.fnArray.flatMap((operator, index) => {
      if (operator !== OPS.setFillRGBColor) return []
      const color = operators.argsArray[index]?.[0] as ArrayLike<number> | string | undefined
      if (typeof color === 'string' && /^#[\da-f]{6}$/i.test(color)) {
        return [
          [color.slice(1, 3), color.slice(3, 5), color.slice(5, 7)]
            .map((component) => Number.parseInt(component, 16))
            .join(','),
        ]
      }
      return color ? [Array.from(color).join(',')] : []
    })
    return {
      text: content.items.map((item) => ('str' in item ? item.str : '')).join(' '),
      imageCount: operators.fnArray.filter((operator) =>
        [OPS.paintImageXObject, OPS.paintInlineImageXObject, OPS.paintImageMaskXObject].includes(
          operator,
        ),
      ).length,
      fillColors: [...new Set(fillColors)],
      annotationSubtypes: annotations.map((annotation) => annotation.subtype),
    }
  } finally {
    await loadingTask.destroy()
  }
}

test('a user opens a PDF in the retained community renderer', async ({ page }) => {
  await page.goto('/?format=pdf&width=1332&height=900')
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)

  const editor = page.frameLocator('#editor-frame')
  await editor.locator('input[type="file"][accept*=".pdf"]').setInputFiles({
    name: 'community.pdf',
    mimeType: 'application/pdf',
    buffer: await fixture(),
  })

  await expect(editor.locator('.ribbon-tabs')).toBeVisible()
  await expect(editor.locator('.pdf-body')).toBeVisible()
  await expect(editor.locator('.pdf-page-content canvas').first()).toBeVisible()
})

test('an edited PDF saves through the session-bound local persistence protocol', async ({
  page,
}) => {
  await page.goto(hostUrl(SPLIT))
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)
  await openFixture(page)

  await executePdfCommand(page, {
    commandId: 'pdf-host-save-rotation',
    baseRevision: 0,
    operation: 'pdf.page.set_rotation',
    arguments: { pageIndex: 0, rotation: 90 },
  })
  const downloadPromise = page.waitForEvent('download', { timeout: 5_000 })
  const acknowledgement = await executePdfCommand(page, {
    commandId: 'pdf-host-save',
    baseRevision: 1,
    operation: 'pdf.document.save',
    arguments: {},
  })
  const download = await downloadPromise

  expect(acknowledgement).toMatchObject({ ok: true, dirty: false })
  expect(download.suggestedFilename()).toBe('community.pdf')
  await expect
    .poll(() => page.evaluate(() => window.__codexVisualHost.downloads.at(-1)))
    .toMatchObject({
      fileName: 'community.pdf',
      mimeType: 'application/pdf',
      size: expect.any(Number),
    })
  expect(await page.evaluate(() => window.__codexVisualHost.events)).toEqual(
    expect.arrayContaining([
      'tool:office_editor_begin_document_save',
      'tool:office_editor_commit_document_save',
    ]),
  )
})

test('an offscreen PDF releases page canvases and resumes the same document', async ({ page }) => {
  await page.goto(hostUrl(SPLIT))
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)
  await openFixture(page)

  const editor = page.frameLocator('#editor-frame')
  await editor.locator('#root').evaluate((root) => {
    root.dispatchEvent(
      Object.assign(new Event('contentvisibilityautostatechange'), { skipped: true }),
    )
  })
  await expect(editor.locator('.pdf-page-content canvas')).toHaveCount(0)

  await editor.locator('#root').evaluate((root) => {
    root.dispatchEvent(
      Object.assign(new Event('contentvisibilityautostatechange'), { skipped: false }),
    )
  })
  await expect(editor.locator('.pdf-page-content canvas').first()).toBeVisible()

  const state = await page.evaluate(() => window.__codexVisualHost)
  expect(state.editorLoads).toBe(1)
  expect(state.errors).toEqual([])
})

test('PDF stores each dirty edit version only once', async ({ page }) => {
  await page.goto(hostUrl(SPLIT))
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)
  await openFixture(page)

  await page.evaluate(() => {
    window.__codexVisualHost.enqueueCommand({
      commandId: 'pdf-recovery-v1',
      baseRevision: 0,
      operation: 'pdf.page.set_rotation',
      arguments: { pageIndex: 0, rotation: 90 },
    })
  })
  await expect.poll(() => page.evaluate(() => window.__codexVisualHost.recoveryCommits)).toBe(1)

  await page.waitForTimeout(2_300)
  expect(await page.evaluate(() => window.__codexVisualHost.recoveryCommits)).toBe(1)

  await page.evaluate(() => {
    window.__codexVisualHost.enqueueCommand({
      commandId: 'pdf-recovery-v2',
      baseRevision: 1,
      operation: 'pdf.page.set_rotation',
      arguments: { pageIndex: 0, rotation: 180 },
    })
  })
  await expect.poll(() => page.evaluate(() => window.__codexVisualHost.recoveryCommits)).toBe(2)
})

test('an MCP staged PDF opens in the same retained community renderer', async ({ page }) => {
  const bytes = await fixture()
  await page.goto('/?format=pdf&width=720&height=900')
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)

  await page.evaluate(
    ({ data, size }) => {
      const state = window.__codexVisualHost
      state.stagedFiles['pdf-fixture'] = data
      state.commands.push({
        commandId: 'open-pdf',
        baseRevision: 0,
        operation: 'pdf.document.load_staged',
        arguments: {
          blobId: 'pdf-fixture',
          name: 'community.pdf',
          size,
        },
      })
    },
    { data: bytes.toString('base64'), size: bytes.length },
  )

  const editor = page.frameLocator('#editor-frame')
  await expect(editor.locator('.ribbon-tabs')).toBeVisible()
  await expect(editor.locator('.pdf-page-content canvas').first()).toBeVisible()
  await expect
    .poll(() =>
      page.evaluate(() =>
        window.__codexVisualHost.acknowledgements.find(
          (acknowledgement) => acknowledgement.commandId === 'open-pdf',
        ),
      ),
    )
    .toMatchObject({ ok: true, fileName: 'community.pdf' })
})

test('a positive staged-open acknowledgement makes the next PDF mutation immediately safe', async ({
  page,
}) => {
  const bytes = await fixture()
  await page.goto('/?format=pdf&width=720&height=900')
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)

  await page.evaluate(
    ({ data, size }) => {
      window.__codexVisualHost.enqueueCommand(
        {
          commandId: 'open-before-immediate-rotation',
          baseRevision: 0,
          operation: 'pdf.document.load_staged',
          arguments: {
            blobId: 'pdf-immediate-rotation-fixture',
            name: 'immediate.pdf',
            size,
          },
        },
        { blobId: 'pdf-immediate-rotation-fixture', base64: data },
      )
    },
    { data: bytes.toString('base64'), size: bytes.length },
  )
  await page.waitForFunction(() =>
    window.__codexVisualHost.acknowledgements.some(
      (entry) => entry.commandId === 'open-before-immediate-rotation' && entry.ok === true,
    ),
  )

  await page.evaluate(() => {
    window.__codexVisualHost.enqueueCommand({
      commandId: 'immediate-rotation',
      baseRevision: 1,
      operation: 'pdf.page.set_rotation',
      arguments: { pageIndex: 0, rotation: 90 },
    })
  })
  await expect
    .poll(() =>
      page.evaluate(() =>
        window.__codexVisualHost.acknowledgements.find(
          (entry) => entry.commandId === 'immediate-rotation',
        ),
      ),
    )
    .toMatchObject({ ok: true, revision: 2 })
})

test('user undo and typed MCP delete share state through save and reopen', async ({ page }) => {
  const fixture = await annotatedFixture()
  await page.goto('/?format=pdf&width=720&height=900')
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)
  const editor = page.frameLocator('#editor-frame')
  const fileInput = editor.locator('input[type="file"][accept*=".pdf"]')
  await fileInput.setInputFiles({
    name: 'review.pdf',
    mimeType: 'application/pdf',
    buffer: fixture.buffer,
  })
  const renderedPage = editor.locator('.pdf-page').first()
  await expect(renderedPage.locator('.pdf-page-content canvas')).toBeVisible()

  const pageBox = await renderedPage.boundingBox()
  expect(pageBox).not.toBeNull()
  const scale = pageBox!.width / 420
  const highlightPosition = { x: 100 * scale, y: 69 * scale }
  await renderedPage.click({ position: highlightPosition })
  await editor.locator('.pdf-del-popup-danger').click()

  const undo = editor.locator('button[aria-label^="撤销"]')
  await expect(undo).toBeEnabled()
  await undo.click()
  await renderedPage.click({ position: highlightPosition })
  await expect(editor.locator('.pdf-del-popup-danger')).toBeVisible()
  await page.keyboard.press('Escape')

  await page.evaluate((deletion) => {
    window.__codexVisualHost.commands.push({
      commandId: 'delete-saved-annotation',
      baseRevision: 0,
      operation: 'pdf.annotation.delete_saved',
      arguments: deletion,
    })
  }, fixture.deletion)
  await expect
    .poll(() =>
      page.evaluate(() =>
        window.__codexVisualHost.acknowledgements.find(
          (acknowledgement) => acknowledgement.commandId === 'delete-saved-annotation',
        ),
      ),
    )
    .toMatchObject({ ok: true, dirty: true })

  const downloadPromise = page.waitForEvent('download')
  await page.evaluate(() => {
    window.__codexVisualHost.commands.push({
      commandId: 'save-pdf',
      baseRevision: 1,
      operation: 'pdf.document.save',
      arguments: {},
    })
  })
  const download = await downloadPromise
  await expect
    .poll(() =>
      page.evaluate(() =>
        window.__codexVisualHost.acknowledgements.find(
          (acknowledgement) => acknowledgement.commandId === 'save-pdf',
        ),
      ),
    )
    .toMatchObject({ ok: true, dirty: false })

  const savedPath = await download.path()
  expect(savedPath).not.toBeNull()
  const savedBytes = await readFile(savedPath!)
  const reopened = await PDFDocument.load(savedBytes)
  const subtypes = reopened
    .getPage(0)
    .node.lookup(PDFName.of('Annots'), PDFArray)
    .asArray()
    .map((reference) =>
      reopened.context
        .lookup(reference, PDFDict)
        .lookup(PDFName.of('Subtype'), PDFName)
        .decodeText(),
    )
  expect(subtypes).toEqual(['Underline'])

  await fileInput.setInputFiles({
    name: 'review-reopened.pdf',
    mimeType: 'application/pdf',
    buffer: savedBytes,
  })
  await expect(editor.locator('.pdf-page-content canvas').first()).toBeVisible()
  await renderedPage.click({ position: highlightPosition })
  await expect(editor.locator('.pdf-del-popup-danger')).toBeHidden()
})

test('typed MCP text and image content-stream edits survive browser save and reopen', async ({
  page,
}) => {
  await page.goto('/?format=pdf&width=720&height=900')
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)
  const editor = page.frameLocator('#editor-frame')
  const fileInput = editor.locator('input[type="file"][accept*=".pdf"]')
  await fileInput.setInputFiles({
    name: 'content-stream.pdf',
    mimeType: 'application/pdf',
    buffer: await fixture(),
  })
  await expect(editor.locator('.pdf-page-content canvas').first()).toBeVisible()

  await page.evaluate((image) => {
    window.__codexVisualHost.commands.push(
      {
        commandId: 'insert-searchable-text',
        baseRevision: 0,
        operation: 'pdf.text.insert',
        arguments: {
          pageIndex: 0,
          origin: [30, 190],
          text: 'Inserted through typed MCP',
          fontSize: 14,
          color: [20, 40, 60],
          font: 'arial',
        },
      },
      {
        commandId: 'insert-content-image',
        baseRevision: 1,
        operation: 'pdf.image.insert',
        arguments: {
          pageIndex: 0,
          rect: [250, 170, 300, 220],
          image,
          layer: 'aboveText',
        },
      },
    )
  }, TINY_PNG)

  for (const commandId of ['insert-searchable-text', 'insert-content-image']) {
    await expect
      .poll(() =>
        page.evaluate((id) => {
          return window.__codexVisualHost.acknowledgements.find(
            (acknowledgement) => acknowledgement.commandId === id,
          )
        }, commandId),
      )
      .not.toBeUndefined()
    const acknowledgement = await page.evaluate((id) => {
      return window.__codexVisualHost.acknowledgements.find(
        (candidate) => candidate.commandId === id,
      )
    }, commandId)
    if (acknowledgement?.ok !== true) {
      throw new Error(`Unexpected acknowledgement: ${JSON.stringify(acknowledgement)}`)
    }
    expect(acknowledgement).toMatchObject({ ok: true, dirty: true })
  }

  const downloadPromise = page.waitForEvent('download')
  await page.evaluate(() => {
    window.__codexVisualHost.commands.push({
      commandId: 'save-content-stream',
      baseRevision: 2,
      operation: 'pdf.document.save',
      arguments: {},
    })
  })
  const download = await downloadPromise
  await expect
    .poll(() =>
      page.evaluate(() =>
        window.__codexVisualHost.acknowledgements.find(
          (acknowledgement) => acknowledgement.commandId === 'save-content-stream',
        ),
      ),
    )
    .toMatchObject({ ok: true, dirty: false })

  const savedPath = await download.path()
  expect(savedPath).not.toBeNull()
  const savedBytes = await readFile(savedPath!)
  const loadingTask = getDocument({ data: new Uint8Array(savedBytes) })
  try {
    const reopened = await loadingTask.promise
    const reopenedPage = await reopened.getPage(1)
    const content = await reopenedPage.getTextContent()
    expect(content.items.map((item) => ('str' in item ? item.str : '')).join(' ')).toContain(
      'Inserted through typed MCP',
    )
    const operators = await reopenedPage.getOperatorList()
    expect(
      operators.fnArray.filter((operator) =>
        [OPS.paintImageXObject, OPS.paintInlineImageXObject, OPS.paintImageMaskXObject].includes(
          operator,
        ),
      ),
    ).not.toEqual([])
  } finally {
    await loadingTask.destroy()
  }

  await fileInput.setInputFiles({
    name: 'content-stream-reopened.pdf',
    mimeType: 'application/pdf',
    buffer: savedBytes,
  })
  await expect(editor.locator('.pdf-page-content canvas').first()).toBeVisible()
})

test('typed MCP replaces text and transforms, replaces, then deletes an existing image', async ({
  page,
}) => {
  await page.goto('/?format=pdf&width=720&height=900')
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)
  const editor = page.frameLocator('#editor-frame')
  const fileInput = editor.locator('input[type="file"][accept*=".pdf"]')
  await fileInput.setInputFiles({
    name: 'existing-content.pdf',
    mimeType: 'application/pdf',
    buffer: await fixture(),
  })
  await expect(editor.locator('.pdf-page-content canvas').first()).toBeVisible()

  await executePdfCommand(page, {
    commandId: 'replace-existing-text',
    baseRevision: 0,
    operation: 'pdf.text.replace',
    arguments: {
      pageIndex: 0,
      rect: [0, 230, 420, 280],
      oldText: 'Original community PDF content',
      newText: 'Replaced through typed MCP',
      fontSize: 18,
    },
  })
  await executePdfCommand(page, {
    commandId: 'seed-existing-image',
    baseRevision: 1,
    operation: 'pdf.image.insert',
    arguments: {
      pageIndex: 0,
      rect: [250, 170, 300, 220],
      image: TINY_PNG,
      layer: 'aboveText',
    },
  })
  const seeded = await savePdfCommand(page, 'save-replaced-content', 2)
  await expect(pdfPageEvidence(seeded)).resolves.toMatchObject({
    text: expect.stringContaining('Replaced through typed MCP'),
    imageCount: 1,
  })

  await executePdfCommand(page, {
    commandId: 'transform-existing-image',
    baseRevision: 3,
    operation: 'pdf.image.transform',
    arguments: {
      pageIndex: 0,
      oldRect: [250, 170, 300, 220],
      rect: [240, 140, 320, 220],
      layer: 'belowText',
      quarterTurns: 1,
    },
  })
  const transformed = await savePdfCommand(page, 'save-transformed-image', 4)
  await expect(pdfPageEvidence(transformed)).resolves.toMatchObject({ imageCount: 1 })

  await executePdfCommand(page, {
    commandId: 'replace-existing-image',
    baseRevision: 5,
    operation: 'pdf.image.replace',
    arguments: {
      pageIndex: 0,
      oldRect: [240, 140, 320, 220],
      rect: [240, 140, 320, 220],
      image: TINY_PNG,
      layer: 'aboveText',
    },
  })
  const replaced = await savePdfCommand(page, 'save-replaced-image', 6)
  await expect(pdfPageEvidence(replaced)).resolves.toMatchObject({ imageCount: 1 })

  await executePdfCommand(page, {
    commandId: 'delete-existing-image',
    baseRevision: 7,
    operation: 'pdf.image.delete',
    arguments: { pageIndex: 0, oldRect: [240, 140, 320, 220] },
  })
  const deleted = await savePdfCommand(page, 'save-deleted-image', 8)
  await expect(pdfPageEvidence(deleted)).resolves.toMatchObject({
    text: expect.stringContaining('Replaced through typed MCP'),
    imageCount: 0,
  })

  await fileInput.setInputFiles({
    name: 'existing-content-reopened.pdf',
    mimeType: 'application/pdf',
    buffer: deleted,
  })
  await expect(editor.locator('.pdf-page-content canvas').first()).toBeVisible()
})

test('typed MCP embeds CJK text and preserves bounded selection-level colors', async ({ page }) => {
  await page.goto('/?format=pdf&width=720&height=900')
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)
  const editor = page.frameLocator('#editor-frame')
  await editor.locator('input[type="file"][accept*=".pdf"]').setInputFiles({
    name: 'font-parity.pdf',
    mimeType: 'application/pdf',
    buffer: await fixture(),
  })
  await expect(editor.locator('.pdf-page-content canvas').first()).toBeVisible()

  await executePdfCommand(page, {
    commandId: 'insert-cjk-text',
    baseRevision: 0,
    operation: 'pdf.text.insert',
    arguments: {
      pageIndex: 0,
      origin: [30, 190],
      text: '中文可搜索\n한국어 검색\nنص عربي',
      fontSize: 14,
      color: [10, 20, 30],
      lineLeading: 20,
    },
  })
  await executePdfCommand(page, {
    commandId: 'replace-with-mixed-colors',
    baseRevision: 1,
    operation: 'pdf.text.replace',
    arguments: {
      pageIndex: 0,
      rect: [0, 230, 420, 280],
      oldText: 'Original community PDF content',
      newText: 'Color split',
      fontSize: 18,
      colorRuns: [
        { start: 0, end: 5, color: [255, 0, 0] },
        { start: 6, end: 11, color: [0, 0, 255] },
      ],
    },
  })
  const saved = await savePdfCommand(page, 'save-font-parity', 2)
  const evidence = await pdfPageEvidence(saved)
  expect(evidence.text).toContain('中文可搜索')
  expect(evidence.text.replaceAll(' ', '')).toContain('한국어검색')
  expect(
    Array.from(evidence.text.normalize('NFKC').replaceAll(' ', '')).reverse().join(''),
  ).toContain('نصعربي')
  expect(evidence.text).toContain('Color split')
  expect(evidence.fillColors).toEqual(expect.arrayContaining(['255,0,0', '0,0,255']))
})

test('typed MCP replaces and explicitly clears generated PDF stamps', async ({ page }) => {
  await page.goto('/?format=pdf&width=720&height=900')
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)
  const editor = page.frameLocator('#editor-frame')
  await editor.locator('input[type="file"][accept*=".pdf"]').setInputFiles({
    name: 'stamp-final-state.pdf',
    mimeType: 'application/pdf',
    buffer: await fixture(),
  })
  await expect(editor.locator('.pdf-page-content canvas').first()).toBeVisible()

  await executePdfCommand(page, {
    commandId: 'set-generated-stamp',
    baseRevision: 0,
    operation: 'pdf.stamp.set',
    arguments: {
      watermark: {
        text: 'DRAFT',
        angle: -30,
        opacity: 0.2,
        color: '#cc0000',
        sizeRatio: 0.2,
      },
      headerFooter: null,
    },
  })
  const stamped = await savePdfCommand(page, 'save-generated-stamp', 1)
  await expect(pdfPageEvidence(stamped)).resolves.toMatchObject({
    annotationSubtypes: ['Stamp'],
  })

  await executePdfCommand(page, {
    commandId: 'clear-generated-stamp',
    baseRevision: 2,
    operation: 'pdf.stamp.set',
    arguments: { watermark: null, headerFooter: null },
  })
  const cleared = await savePdfCommand(page, 'save-cleared-stamp', 3)
  await expect(pdfPageEvidence(cleared)).resolves.toMatchObject({
    annotationSubtypes: [],
  })
})

test('PDF community renderer remains visible in the 420px Codex narrow sidebar', async ({
  page,
}) => {
  await page.goto(hostUrl(NARROW))
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)
  await openFixture(page)

  const editor = page.frameLocator('#editor-frame')
  await expect(editor.locator('.app')).toBeVisible()
  await expect(editor.locator('.pdf-body')).toBeVisible()
  await expect(editor.locator('.status-bar')).toBeVisible()
  await expect(editor.locator('.host-fullscreen[aria-pressed]')).toBeVisible()
  await expect(page.locator('#editor-frame')).toHaveScreenshot('pdf-narrow-sidebar.png')

  const state = await page.evaluate(() => window.__codexVisualHost)
  expect(state.editorLoads).toBe(1)
  expect(state.errors).toEqual([])
})

test('PDF community renderer remains mounted in the 720px Codex split view', async ({ page }) => {
  await page.goto(hostUrl(SPLIT))
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)
  await openFixture(page)

  const editor = page.frameLocator('#editor-frame')
  await expect(editor.locator('.app')).toBeVisible()
  await expect(editor.locator('.pdf-body')).toBeVisible()
  await expect(editor.locator('.host-fullscreen[aria-pressed]')).toBeInViewport()
  await expect(page.locator('#editor-frame')).toHaveScreenshot('pdf-split-view.png')

  const state = await page.evaluate(() => window.__codexVisualHost)
  expect(state.editorLoads).toBe(1)
  expect(state.errors).toEqual([])
})

test('PDF community renderer first fullscreen follows host mount readiness', async ({ page }) => {
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
  await openFixture(page)

  const editor = page.frameLocator('#editor-frame')
  await expect(editor.locator('.pdf-page-content canvas').first()).toBeVisible()
  await expect(editor.locator('.host-fullscreen[aria-pressed]')).toHaveAttribute(
    'aria-pressed',
    'true',
  )
  await expect(page.locator('#editor-frame')).toHaveCSS('width', `${FULLSCREEN.width}px`)
  await expect(page.locator('#editor-frame')).toHaveScreenshot('pdf-first-fullscreen.png')

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

test('PDF community renderer exits fullscreen without remounting its iframe', async ({ page }) => {
  await page.goto(hostUrl(SPLIT, true))
  await page.waitForFunction(() => window.__codexVisualHost?.displayMode === 'fullscreen')
  await openFixture(page)

  const editor = page.frameLocator('#editor-frame')
  await editor.locator('.host-fullscreen[aria-pressed]').click()
  await page.waitForFunction(() => window.__codexVisualHost?.displayMode === 'inline')

  await expect(page.locator('#editor-frame')).toHaveCSS('width', `${SPLIT.width}px`)
  await expect(editor.locator('.app')).toBeVisible()
  await expect(editor.locator('.pdf-page-content canvas').first()).toBeVisible()
  await expect(editor.locator('.host-fullscreen[aria-pressed]')).toHaveAttribute(
    'aria-pressed',
    'false',
  )
  await expect(page.locator('#editor-frame')).toHaveScreenshot('pdf-exit-fullscreen.png')

  const state = await page.evaluate(() => window.__codexVisualHost)
  expect(state.events).toContain('display-request:inline')
  expect(state.editorLoads).toBe(1)
  expect(state.polls).toBeGreaterThan(0)
  expect(state.errors).toEqual([])
})
