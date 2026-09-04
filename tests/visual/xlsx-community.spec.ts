import { expect, test, type FrameLocator } from '@playwright/test'
import JSZip from 'jszip'

test('XLSX exposes file, save, and fullscreen as accessible icon-only controls', async ({
  page,
}) => {
  await page.goto('/?format=xlsx&width=1280&height=900')
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)

  const editor = page.frameLocator('#editor-frame')
  const controls = [
    editor.getByRole('button', { name: '打开 XLSX', exact: true }),
    editor.getByRole('button', { name: /^保存/ }),
    editor.getByRole('button', { name: '全屏', exact: true }),
  ]

  for (const control of controls) {
    await expect(control).toBeVisible()
    await expect(control.locator('svg')).toHaveCount(1)
    await expect(control).toHaveText('')
  }
})

async function installWorkbookSaveHandle(
  editor: FrameLocator,
  fileName: string,
): Promise<() => Promise<Buffer>> {
  await editor.locator('body').evaluate((body, name) => {
    const frameWindow = body.ownerDocument.defaultView as Window & {
      __savedWorkbookBytes?: ArrayBuffer
      showSaveFilePicker?: () => Promise<FileSystemFileHandle>
    }
    frameWindow.showSaveFilePicker = async () =>
      ({
        kind: 'file',
        name,
        createWritable: async () =>
          ({
            write: async (data: FileSystemWriteChunkType) => {
              if (data instanceof ArrayBuffer) {
                frameWindow.__savedWorkbookBytes = data.slice(0)
              } else if (ArrayBuffer.isView(data)) {
                frameWindow.__savedWorkbookBytes = data.buffer.slice(
                  data.byteOffset,
                  data.byteOffset + data.byteLength,
                ) as ArrayBuffer
              } else if (data instanceof Blob) {
                frameWindow.__savedWorkbookBytes = await data.arrayBuffer()
              } else {
                throw new Error('Unexpected XLSX write payload.')
              }
            },
            close: async () => undefined,
          }) as FileSystemWritableFileStream,
        getFile: async () =>
          new File([frameWindow.__savedWorkbookBytes ?? new ArrayBuffer(0)], name, {
            type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
          }),
      }) as FileSystemFileHandle
  }, fileName)
  return async () => {
    const bytes = await editor.locator('body').evaluate((body) => {
      const frameWindow = body.ownerDocument.defaultView as Window & {
        __savedWorkbookBytes?: ArrayBuffer
      }
      return [...new Uint8Array(frameWindow.__savedWorkbookBytes ?? new ArrayBuffer(0))]
    })
    return Buffer.from(bytes)
  }
}

async function xlsxWithA1(value: string): Promise<Buffer> {
  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>',
  )
  zip.file(
    '_rels/.rels',
    '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>',
  )
  zip.file(
    'xl/workbook.xml',
    '<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>',
  )
  zip.file(
    'xl/_rels/workbook.xml.rels',
    '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>',
  )
  zip.file(
    'xl/styles.xml',
    '<?xml version="1.0"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>',
  )
  zip.file(
    'xl/worksheets/sheet1.xml',
    `<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>${value}</t></is></c></row></sheetData></worksheet>`,
  )
  return zip.generateAsync({ type: 'nodebuffer' })
}

async function xlsxWithFitToPages(value: string): Promise<Buffer> {
  const zip = await JSZip.loadAsync(await xlsxWithA1(value))
  const path = 'xl/worksheets/sheet1.xml'
  const worksheetXml = await zip.file(path)!.async('text')
  zip.file(
    path,
    worksheetXml
      .replace('<sheetData>', '<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr><sheetData>')
      .replace('</worksheet>', '<pageSetup fitToWidth="2" fitToHeight="3"/></worksheet>'),
  )
  return zip.generateAsync({ type: 'nodebuffer' })
}

async function xlsxWithCopyValuesSource(): Promise<Buffer> {
  const zip = await JSZip.loadAsync(await xlsxWithA1('placeholder'))
  zip.file(
    'xl/worksheets/sheet1.xml',
    '<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><f>1+2</f><v>3</v></c><c r="B1" t="inlineStr"><is><t>North</t></is></c></row><row r="2"><c r="A2"><v>7</v></c><c r="B2" t="b"><v>1</v></c></row></sheetData></worksheet>',
  )
  return zip.generateAsync({ type: 'nodebuffer' })
}

async function xlsxWithCopyFormulasSource(): Promise<Buffer> {
  const zip = await JSZip.loadAsync(await xlsxWithA1('placeholder'))
  zip.file(
    'xl/worksheets/sheet1.xml',
    '<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><v>10</v></c><c r="B1"><f>A1+$A$1+A$1+$A1</f><v>40</v></c></row><row r="2"><c r="A2"><v>5</v></c><c r="B2"><f>SUM(A1:B1)</f><v>50</v></c></row></sheetData></worksheet>',
  )
  return zip.generateAsync({ type: 'nodebuffer' })
}

async function xlsxWithCopyFormatsSource(): Promise<Buffer> {
  const zip = await JSZip.loadAsync(await xlsxWithA1('placeholder'))
  zip.file(
    'xl/worksheets/sheet1.xml',
    '<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><v>0.25</v></c><c r="B1" t="inlineStr"><is><t>plain</t></is></c></row><row r="3"><c r="D3"><v>99</v></c><c r="E3"><f>D3*2</f><v>198</v></c></row></sheetData></worksheet>',
  )
  return zip.generateAsync({ type: 'nodebuffer' })
}

async function xlsxWithCopyColumnWidthsSource(): Promise<Buffer> {
  const zip = await JSZip.loadAsync(await xlsxWithA1('placeholder'))
  for (const path of ['[Content_Types].xml', 'xl/workbook.xml', 'xl/_rels/workbook.xml.rels']) {
    const xml = await zip.file(path)!.async('text')
    zip.file(
      path,
      path === '[Content_Types].xml'
        ? xml.replace(
            '</Types>',
            '<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>',
          )
        : path === 'xl/workbook.xml'
          ? xml.replace('</sheets>', '<sheet name="Archive" sheetId="2" r:id="rId3"/></sheets>')
          : xml.replace(
              '</Relationships>',
              '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/></Relationships>',
            ),
    )
  }
  zip.file(
    'xl/worksheets/sheet1.xml',
    '<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c></row></sheetData></worksheet>',
  )
  zip.file(
    'xl/worksheets/sheet2.xml',
    '<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="D1"><v>99</v></c><c r="E1"><f>D1*2</f><v>198</v></c></row></sheetData></worksheet>',
  )
  return zip.generateAsync({ type: 'nodebuffer' })
}

async function xlsxWithCopyWithoutBordersSource(): Promise<Buffer> {
  const zip = await JSZip.loadAsync(await xlsxWithCopyColumnWidthsSource())
  zip.file(
    'xl/worksheets/sheet1.xml',
    '<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><v>10</v></c><c r="B1"><f>A1+$A$1+A$1+$A1</f><v>40</v></c></row></sheetData></worksheet>',
  )
  zip.file(
    'xl/worksheets/sheet2.xml',
    '<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="3"><c r="D3"><v>99</v></c><c r="E3"><f>D3*2</f><v>198</v></c></row></sheetData></worksheet>',
  )
  return zip.generateAsync({ type: 'nodebuffer' })
}

async function xlsxWithTableSource(): Promise<Buffer> {
  const zip = await JSZip.loadAsync(await xlsxWithA1('placeholder'))
  zip.file(
    '[Content_Types].xml',
    (await zip.file('[Content_Types].xml')!.async('text')).replace(
      '</Types>',
      '<Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/></Types>',
    ),
  )
  zip.file(
    'xl/worksheets/sheet1.xml',
    '<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Region</t></is></c><c r="B1" t="inlineStr"><is><t>Amount</t></is></c></row><row r="2"><c r="A2" t="inlineStr"><is><t>North</t></is></c><c r="B2"><v>10</v></c></row><row r="3"><c r="A3" t="inlineStr"><is><t>South</t></is></c><c r="B3"><v>20</v></c></row></sheetData><tableParts count="1"><tablePart r:id="rId1"/></tableParts></worksheet>',
  )
  zip.file(
    'xl/worksheets/_rels/sheet1.xml.rels',
    '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/></Relationships>',
  )
  zip.file(
    'xl/tables/table1.xml',
    '<?xml version="1.0"?><table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="SalesTable" displayName="SalesTable" ref="A1:B3" headerRowCount="1"><autoFilter ref="A1:B3"/><tableColumns count="2"><tableColumn id="1" name="Region"/><tableColumn id="2" name="Amount"/></tableColumns><tableStyleInfo name="TableStyleMedium4" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/></table>',
  )
  return zip.generateAsync({ type: 'nodebuffer' })
}

function savedColumnWidth(sheetXml: string, oneBasedColumn: number): string | undefined {
  const tags = sheetXml.match(/<col\b[^>]*\/>/g) ?? []
  return tags
    .map((tag) => ({
      min: Number(/\bmin="([^"]+)"/.exec(tag)?.[1]),
      max: Number(/\bmax="([^"]+)"/.exec(tag)?.[1]),
      width: /\bwidth="([^"]+)"/.exec(tag)?.[1],
    }))
    .find(({ min, max }) => min <= oneBasedColumn && oneBasedColumn <= max)?.width
}

test('packaged XLSX opens the complete community workbook workspace', async ({ page }) => {
  await page.goto('/?format=xlsx&width=1332&height=900')
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)

  const editor = page.frameLocator('#editor-frame')
  await expect(editor.locator('.app-shell')).toBeVisible()
  await expect(editor.locator('.ribbon-tabs[aria-label="Workbook commands"]')).toBeVisible()
  await expect(editor.locator('#univer-container')).toBeVisible()
  await expect(editor.locator('canvas[id^="univer-sheet-main-canvas_"]')).toBeVisible()
})

test('the retained Insert Symbol command opens the community dialog', async ({ page }) => {
  await page.goto('/?format=xlsx&width=1332&height=900')
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)

  const editor = page.frameLocator('#editor-frame')
  await editor.getByRole('button', { name: '插入', exact: true }).click()
  await editor.locator('button[data-tip="符号"]').click()

  await expect(editor.getByRole('dialog', { name: '插入符号' })).toBeVisible()
})

test('opening a local XLSX loads its cells into the mounted workbook session', async ({ page }) => {
  await page.goto('/?format=xlsx&width=1332&height=900')
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)

  const editor = page.frameLocator('#editor-frame')
  await editor.locator('input[type="file"][accept*=".xlsx"]').setInputFiles({
    name: 'local-fixture.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: await xlsxWithA1('Community renderer'),
  })

  await expect(editor.locator('.workbook-status')).toContainText('local-fixture.xlsx')
  await expect
    .poll(() =>
      page.evaluate(() => {
        const selection = window.__codexVisualHost.lastPollArguments?.selection as
          { activeCell?: { value?: unknown } } | undefined
        return selection?.activeCell?.value
      }),
    )
    .toBe('Community renderer')
})

test('the MCP staged-file descriptor opens a local XLSX in the mounted renderer', async ({
  page,
}) => {
  await page.goto('/?format=xlsx&width=720&height=900')
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)
  const bytes = await xlsxWithA1('Opened by agent')

  await page.evaluate(
    ({ base64, size }) => {
      window.__codexVisualHost.stagedFiles['xlsx-open-blob'] = base64
      window.__codexVisualHost.commands.push({
        commandId: 'xlsx-open-local',
        baseRevision: 0,
        operation: 'xlsx.document.load_staged',
        arguments: { blobId: 'xlsx-open-blob', name: 'agent-open.xlsx', size },
      })
    },
    { base64: bytes.toString('base64'), size: bytes.byteLength },
  )

  await expect
    .poll(() =>
      page.evaluate(() =>
        window.__codexVisualHost.acknowledgements.find(
          (entry) => entry.commandId === 'xlsx-open-local',
        ),
      ),
    )
    .toMatchObject({ ok: true, revision: 1 })
  await expect
    .poll(() => page.evaluate(() => window.__codexVisualHost.lastPollArguments))
    .toMatchObject({
      fileName: 'agent-open.xlsx',
      selection: { activeCell: { value: 'Opened by agent' } },
    })
})

test('an MCP cell edit changes the mounted workbook revision and saved XLSX', async ({ page }) => {
  await page.goto('/?format=xlsx&width=1332&height=900')
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)
  const editor = page.frameLocator('#editor-frame')
  await editor.locator('input[type="file"][accept*=".xlsx"]').setInputFiles({
    name: 'roundtrip.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: await xlsxWithA1('Before'),
  })
  await expect(editor.locator('.workbook-status')).toContainText('roundtrip.xlsx')
  const readSavedWorkbook = await installWorkbookSaveHandle(editor, 'roundtrip.xlsx')

  await page.evaluate(() => {
    window.__codexVisualHost.commands.push({
      commandId: 'xlsx-set-a1',
      baseRevision: 0,
      operation: 'xlsx.cell.set_value',
      arguments: { sheet: 'Sheet1', address: 'A1', value: 'After' },
    })
  })
  await expect
    .poll(() =>
      page.evaluate(() =>
        window.__codexVisualHost.acknowledgements.find(
          (entry) => entry.commandId === 'xlsx-set-a1',
        ),
      ),
    )
    .toMatchObject({ ok: true, revision: 1 })
  await expect
    .poll(() =>
      page.evaluate(() => {
        const selection = window.__codexVisualHost.lastPollArguments?.selection as
          { activeCell?: { value?: unknown } } | undefined
        return selection?.activeCell?.value
      }),
    )
    .toBe('After')

  await editor.locator('button.qa-btn').first().click()
  const zip = await JSZip.loadAsync(await readSavedWorkbook())
  const sheetXml = await zip.file('xl/worksheets/sheet1.xml')!.async('text')
  expect(sheetXml).toContain('After')
})

test('XLSX stores each dirty workbook version only once', async ({ page }) => {
  await page.goto('/?format=xlsx&width=1332&height=900')
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)
  const editor = page.frameLocator('#editor-frame')
  await editor.locator('input[type="file"][accept*=".xlsx"]').setInputFiles({
    name: 'recovery-version.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: await xlsxWithA1('Before'),
  })
  await expect(editor.locator('.workbook-status')).toContainText('recovery-version.xlsx')

  // The browser import first checkpoints the clean workbook for exact continuation.
  await expect.poll(() => page.evaluate(() => window.__codexVisualHost.recoveryCommits)).toBe(1)

  await page.evaluate(() => {
    window.__codexVisualHost.enqueueCommand({
      commandId: 'xlsx-recovery-v1',
      baseRevision: 0,
      operation: 'xlsx.cell.set_value',
      arguments: { sheet: 'Sheet1', address: 'A1', value: 'First' },
    })
  })
  await expect.poll(() => page.evaluate(() => window.__codexVisualHost.recoveryCommits)).toBe(2)

  await page.waitForTimeout(2_300)
  expect(await page.evaluate(() => window.__codexVisualHost.recoveryCommits)).toBe(2)

  await page.evaluate(() => {
    window.__codexVisualHost.enqueueCommand({
      commandId: 'xlsx-recovery-v2',
      baseRevision: 1,
      operation: 'xlsx.cell.set_value',
      arguments: { sheet: 'Sheet1', address: 'A1', value: 'Second' },
    })
  })
  await expect.poll(() => page.evaluate(() => window.__codexVisualHost.recoveryCommits)).toBe(3)
})

test('the XLSX history operations undo and redo the latest mounted workbook edit', async ({
  page,
}) => {
  await page.goto('/?format=xlsx&width=720&height=900')
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)
  const editor = page.frameLocator('#editor-frame')
  await editor.locator('input[type="file"][accept*=".xlsx"]').setInputFiles({
    name: 'registry-history.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: await xlsxWithA1('Before'),
  })
  const readSavedWorkbook = await installWorkbookSaveHandle(editor, 'registry-history.xlsx')

  await page.evaluate(() => {
    window.__codexVisualHost.commands.push({
      commandId: 'xlsx-history-edit',
      baseRevision: 0,
      operation: 'xlsx.cell.set_value',
      arguments: { sheet: 'Sheet1', address: 'A1', value: 'After' },
    })
  })
  await expect
    .poll(() =>
      page.evaluate(() =>
        window.__codexVisualHost.acknowledgements.find(
          (entry) => entry.commandId === 'xlsx-history-edit',
        ),
      ),
    )
    .toMatchObject({ ok: true, revision: 1 })

  await page.evaluate(() => {
    window.__codexVisualHost.commands.push({
      commandId: 'xlsx-history-undo',
      baseRevision: 1,
      operation: 'xlsx.history.undo',
      arguments: {},
    })
  })
  await expect
    .poll(() =>
      page.evaluate(() =>
        window.__codexVisualHost.acknowledgements.find(
          (entry) => entry.commandId === 'xlsx-history-undo',
        ),
      ),
    )
    .toMatchObject({ ok: true, revision: 2, output: { undone: true } })
  await expect
    .poll(() =>
      page.evaluate(() => {
        const selection = window.__codexVisualHost.lastPollArguments?.selection as
          { activeCell?: { value?: unknown } } | undefined
        return selection?.activeCell?.value
      }),
    )
    .toBe('Before')

  await page.evaluate(() => {
    window.__codexVisualHost.commands.push({
      commandId: 'xlsx-history-redo',
      baseRevision: 2,
      operation: 'xlsx.history.redo',
      arguments: {},
    })
  })
  await expect
    .poll(() =>
      page.evaluate(() =>
        window.__codexVisualHost.acknowledgements.find(
          (entry) => entry.commandId === 'xlsx-history-redo',
        ),
      ),
    )
    .toMatchObject({ ok: true, revision: 3, output: { redone: true } })
  await expect
    .poll(() =>
      page.evaluate(() => {
        const selection = window.__codexVisualHost.lastPollArguments?.selection as
          { activeCell?: { value?: unknown } } | undefined
        return selection?.activeCell?.value
      }),
    )
    .toBe('After')

  await editor.locator('button.qa-btn').first().click()
  const zip = await JSZip.loadAsync(await readSavedWorkbook())
  const sheetXml = await zip.file('xl/worksheets/sheet1.xml')!.async('text')
  expect(sheetXml).toContain('After')
})

test('the XLSX copy-values operation copies computed values through shared undo and save', async ({
  page,
}) => {
  await page.goto('/?format=xlsx&width=720&height=900')
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)
  const editor = page.frameLocator('#editor-frame')
  await editor.locator('input[type="file"][accept*=".xlsx"]').setInputFiles({
    name: 'registry-copy-values.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: await xlsxWithCopyValuesSource(),
  })
  const readSavedWorkbook = await installWorkbookSaveHandle(editor, 'registry-copy-values.xlsx')

  await page.evaluate(() => {
    window.__codexVisualHost.commands.push({
      commandId: 'xlsx-copy-values',
      baseRevision: 0,
      operation: 'xlsx.range.copy_values',
      arguments: {
        sourceSheet: 'Sheet1',
        sourceRange: 'A1:B2',
        destinationSheet: 'Sheet1',
        destinationRange: 'D1:E2',
      },
    })
  })
  await expect
    .poll(() =>
      page.evaluate(() =>
        window.__codexVisualHost.acknowledgements.find(
          (entry) => entry.commandId === 'xlsx-copy-values',
        ),
      ),
    )
    .toMatchObject({
      ok: true,
      revision: 1,
      output: {
        sourceSheet: 'Sheet1',
        sourceRange: 'A1:B2',
        destinationSheet: 'Sheet1',
        destinationRange: 'D1:E2',
        copied: 4,
      },
    })
  await expect
    .poll(() =>
      page.evaluate(() => {
        const selection = window.__codexVisualHost.lastPollArguments?.selection as
          { activeCell?: { address?: string; value?: unknown } } | undefined
        return selection?.activeCell
      }),
    )
    .toMatchObject({ address: 'D1', value: 3 })

  const nameBox = editor.getByRole('textbox', { name: 'Name Box' })
  await editor.locator('button.qa-btn').nth(1).click()
  await nameBox.fill('D1')
  await nameBox.press('Enter')
  await expect
    .poll(() =>
      page.evaluate(() => {
        const selection = window.__codexVisualHost.lastPollArguments?.selection as
          { activeCell?: { address?: string; value?: unknown } } | undefined
        return selection?.activeCell
      }),
    )
    .toMatchObject({ address: 'D1', value: null })
  await editor.locator('button.qa-btn').nth(2).click()
  await nameBox.fill('D1')
  await nameBox.press('Enter')
  await expect
    .poll(() =>
      page.evaluate(() => {
        const selection = window.__codexVisualHost.lastPollArguments?.selection as
          { activeCell?: { address?: string; value?: unknown } } | undefined
        return selection?.activeCell
      }),
    )
    .toMatchObject({ address: 'D1', value: 3 })

  await editor.locator('button.qa-btn').first().click()
  const savedZip = await JSZip.loadAsync(await readSavedWorkbook())
  const sheetXml = await savedZip.file('xl/worksheets/sheet1.xml')!.async('text')
  const destinationCell = /<c r="D1"[^>]*>([\s\S]*?)<\/c>/.exec(sheetXml)?.[1]
  expect(destinationCell).toContain('<v>3</v>')
  expect(destinationCell).not.toContain('<f>')
})

test('the XLSX copy-formulas operation translates references through shared undo and save', async ({
  page,
}) => {
  await page.goto('/?format=xlsx&width=720&height=900')
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)
  const editor = page.frameLocator('#editor-frame')
  await editor.locator('input[type="file"][accept*=".xlsx"]').setInputFiles({
    name: 'registry-copy-formulas.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: await xlsxWithCopyFormulasSource(),
  })
  const readSavedWorkbook = await installWorkbookSaveHandle(editor, 'registry-copy-formulas.xlsx')

  await page.evaluate(() => {
    window.__codexVisualHost.commands.push({
      commandId: 'xlsx-copy-formulas',
      baseRevision: 0,
      operation: 'xlsx.range.copy_formulas',
      arguments: {
        sourceSheet: 'Sheet1',
        sourceRange: 'B1:B2',
        destinationSheet: 'Sheet1',
        destinationRange: 'D3:D4',
      },
    })
  })
  await expect
    .poll(() =>
      page.evaluate(() =>
        window.__codexVisualHost.acknowledgements.find(
          (entry) => entry.commandId === 'xlsx-copy-formulas',
        ),
      ),
    )
    .toMatchObject({
      ok: true,
      revision: 1,
      output: {
        sourceSheet: 'Sheet1',
        sourceRange: 'B1:B2',
        destinationSheet: 'Sheet1',
        destinationRange: 'D3:D4',
        copied: 2,
      },
    })
  await expect
    .poll(() => page.evaluate(() => window.__codexVisualHost.lastPollArguments))
    .toMatchObject({ selection: { activeCell: { address: 'D3' } } })

  const nameBox = editor.getByRole('textbox', { name: 'Name Box' })
  await editor.locator('button.qa-btn').nth(1).click()
  await nameBox.fill('D3')
  await nameBox.press('Enter')
  await expect
    .poll(() => page.evaluate(() => window.__codexVisualHost.lastPollArguments))
    .toMatchObject({ selection: { activeCell: { address: 'D3', value: null } } })
  await editor.locator('button.qa-btn').nth(2).click()

  await editor.locator('button.qa-btn').first().click()
  const savedZip = await JSZip.loadAsync(await readSavedWorkbook())
  const sheetXml = await savedZip.file('xl/worksheets/sheet1.xml')!.async('text')
  expect(/<c r="D3"[^>]*>[\s\S]*?<f>C3\+\$A\$1\+C\$1\+\$A3<\/f>/.test(sheetXml)).toBe(true)
  expect(/<c r="D4"[^>]*>[\s\S]*?<f>SUM\(C3:D3\)<\/f>/.test(sheetXml)).toBe(true)
})

test('the XLSX copy-formats operation replaces only cell formats through shared undo and save', async ({
  page,
}) => {
  await page.goto('/?format=xlsx&width=720&height=900')
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)
  const editor = page.frameLocator('#editor-frame')
  await editor.locator('input[type="file"][accept*=".xlsx"]').setInputFiles({
    name: 'registry-copy-formats.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: await xlsxWithCopyFormatsSource(),
  })

  const setupCommands = [
    {
      operation: 'xlsx.range.set_text_style',
      arguments: {
        sheet: 'Sheet1',
        range: 'A1',
        style: { bold: true },
        fields: ['bold'],
      },
    },
    {
      operation: 'xlsx.range.set_font',
      arguments: {
        sheet: 'Sheet1',
        range: 'A1',
        font: { family: 'Aptos', size: 14, color: '#44546A' },
        fields: ['family', 'size', 'color'],
      },
    },
    {
      operation: 'xlsx.range.set_fill',
      arguments: { sheet: 'Sheet1', range: 'A1', color: '#DDEBF7' },
    },
    {
      operation: 'xlsx.range.set_alignment',
      arguments: {
        sheet: 'Sheet1',
        range: 'A1',
        alignment: { horizontal: 'center', wrap: true },
        fields: ['horizontal', 'wrap'],
      },
    },
    {
      operation: 'xlsx.range.set_number_format',
      arguments: { sheet: 'Sheet1', range: 'A1', pattern: '0.00%' },
    },
    {
      operation: 'xlsx.range.set_fill',
      arguments: { sheet: 'Sheet1', range: 'D3', color: '#FFCC99' },
    },
  ] as const

  for (const [index, command] of setupCommands.entries()) {
    await page.evaluate(
      ({ index, command }) => {
        window.__codexVisualHost.commands.push({
          commandId: `xlsx-copy-formats-setup-${index}`,
          baseRevision: index,
          operation: command.operation,
          arguments: command.arguments,
        })
      },
      { index, command },
    )
    await expect
      .poll(() =>
        page.evaluate(
          (commandId) =>
            window.__codexVisualHost.acknowledgements.find(
              (entry) => entry.commandId === commandId,
            ),
          `xlsx-copy-formats-setup-${index}`,
        ),
      )
      .toMatchObject({ ok: true, revision: index + 1 })
  }

  await page.evaluate(() => {
    window.__codexVisualHost.commands.push({
      commandId: 'xlsx-copy-formats',
      baseRevision: 6,
      operation: 'xlsx.range.copy_formats',
      arguments: {
        sourceSheet: 'Sheet1',
        sourceRange: 'A1:B1',
        destinationSheet: 'Sheet1',
        destinationRange: 'D3:E3',
      },
    })
  })
  await expect
    .poll(() =>
      page.evaluate(() =>
        window.__codexVisualHost.acknowledgements.find(
          (entry) => entry.commandId === 'xlsx-copy-formats',
        ),
      ),
    )
    .toMatchObject({
      ok: true,
      revision: 7,
      output: {
        sourceSheet: 'Sheet1',
        sourceRange: 'A1:B1',
        destinationSheet: 'Sheet1',
        destinationRange: 'D3:E3',
        copied: 2,
      },
    })
  await expect
    .poll(() => page.evaluate(() => window.__codexVisualHost.lastPollArguments))
    .toMatchObject({
      selection: {
        activeCell: {
          address: 'D3',
          value: '9900.00%',
          style: {
            bold: true,
            fillColor: '#DDEBF7',
            fontColor: '#44546A',
            fontFamily: 'Aptos',
            fontSize: 14,
            horizontalAlignment: 'center',
            numberFormat: '0.00%',
            wrap: true,
          },
        },
      },
    })

  await editor.locator('button.qa-btn').nth(1).click()
  await expect
    .poll(() => page.evaluate(() => window.__codexVisualHost.lastPollArguments))
    .toMatchObject({ selection: { activeCell: { value: 99, style: { fillColor: '#FFCC99' } } } })
  await editor.locator('button.qa-btn').nth(2).click()
  await expect
    .poll(() => page.evaluate(() => window.__codexVisualHost.lastPollArguments))
    .toMatchObject({
      selection: { activeCell: { value: '9900.00%', style: { fillColor: '#DDEBF7' } } },
    })

  const downloadPromise = page.waitForEvent('download')
  await editor.locator('button.qa-btn').first().click()
  const savedPath = await (await downloadPromise).path()
  expect(savedPath).not.toBeNull()
  const savedZip = await JSZip.loadAsync(
    await import('node:fs/promises').then((fs) => fs.readFile(savedPath!)),
  )
  const sheetXml = await savedZip.file('xl/worksheets/sheet1.xml')!.async('text')
  const cells = Object.fromEntries(
    ['A1', 'B1', 'D3', 'E3'].map((address) => [
      address,
      new RegExp(`<c\\b[^>]*\\br="${address}"[^>]*>(?:[\\s\\S]*?)<\\/c>`).exec(sheetXml)?.[0] ?? '',
    ]),
  )
  const styleIndex = (cell: string) => /\bs="([0-9]+)"/.exec(cell)?.[1] ?? '0'
  expect(cells.D3).toContain('<v>99</v>')
  expect(cells.E3).toContain('<f>D3*2</f>')
  expect(cells.E3).toContain('<v>198</v>')
  expect(styleIndex(cells.D3)).toBe(styleIndex(cells.A1))
  expect(styleIndex(cells.E3)).toBe(styleIndex(cells.B1))
})

test('the XLSX copy-without-borders operation copies cells while preserving destination borders', async ({
  page,
}) => {
  await page.goto('/?format=xlsx&width=720&height=900')
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)
  const editor = page.frameLocator('#editor-frame')
  await editor.locator('input[type="file"][accept*=".xlsx"]').setInputFiles({
    name: 'registry-copy-without-borders.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: await xlsxWithCopyWithoutBordersSource(),
  })

  const setupCommands = [
    {
      operation: 'xlsx.range.set_text_style',
      arguments: {
        sheet: 'Sheet1',
        range: 'A1:B1',
        style: { bold: true },
        fields: ['bold'],
      },
    },
    {
      operation: 'xlsx.range.set_fill',
      arguments: { sheet: 'Sheet1', range: 'A1:B1', color: '#DDEBF7' },
    },
    {
      operation: 'xlsx.range.set_number_format',
      arguments: { sheet: 'Sheet1', range: 'A1:B1', pattern: '0.00' },
    },
    {
      operation: 'xlsx.range.set_border',
      arguments: {
        sheet: 'Sheet1',
        range: 'A1:B1',
        border: { preset: 'all', lineStyle: 'thick', color: '#C00000' },
      },
    },
    {
      operation: 'xlsx.range.set_fill',
      arguments: { sheet: 'Archive', range: 'D3:E3', color: '#FFCC99' },
    },
    {
      operation: 'xlsx.range.set_border',
      arguments: {
        sheet: 'Archive',
        range: 'D3:E3',
        border: { preset: 'all', lineStyle: 'medium', color: '#4472C4' },
      },
    },
  ] as const

  for (const [index, command] of setupCommands.entries()) {
    await page.evaluate(
      ({ index, command }) => {
        window.__codexVisualHost.commands.push({
          commandId: `xlsx-copy-without-borders-setup-${index}`,
          baseRevision: index,
          operation: command.operation,
          arguments: command.arguments,
        })
      },
      { index, command },
    )
    await expect
      .poll(() =>
        page.evaluate(
          (commandId) =>
            window.__codexVisualHost.acknowledgements.find(
              (entry) => entry.commandId === commandId,
            ),
          `xlsx-copy-without-borders-setup-${index}`,
        ),
      )
      .toMatchObject({ ok: true, revision: index + 1 })
  }

  await page.evaluate(() => {
    window.__codexVisualHost.commands.push({
      commandId: 'xlsx-copy-without-borders',
      baseRevision: 6,
      operation: 'xlsx.range.copy_without_borders',
      arguments: {
        sourceSheet: 'Sheet1',
        sourceRange: 'A1:B1',
        destinationSheet: 'Archive',
        destinationRange: 'D3:E3',
      },
    })
  })
  await expect
    .poll(() =>
      page.evaluate(() =>
        window.__codexVisualHost.acknowledgements.find(
          (entry) => entry.commandId === 'xlsx-copy-without-borders',
        ),
      ),
    )
    .toMatchObject({
      ok: true,
      revision: 7,
      output: {
        sourceSheet: 'Sheet1',
        sourceRange: 'A1:B1',
        destinationSheet: 'Archive',
        destinationRange: 'D3:E3',
        copied: 2,
      },
    })
  await expect
    .poll(() => page.evaluate(() => window.__codexVisualHost.lastPollArguments))
    .toMatchObject({
      selection: {
        activeCell: {
          address: 'D3',
          value: '10.00',
          style: { bold: true, fillColor: '#DDEBF7', numberFormat: '0.00' },
        },
      },
    })

  await editor.locator('button.qa-btn').nth(1).click()
  await expect
    .poll(() => page.evaluate(() => window.__codexVisualHost.lastPollArguments))
    .toMatchObject({
      selection: { activeCell: { value: 99, style: { fillColor: '#FFCC99' } } },
    })
  await editor.locator('button.qa-btn').nth(2).click()
  await expect
    .poll(() => page.evaluate(() => window.__codexVisualHost.lastPollArguments))
    .toMatchObject({
      selection: { activeCell: { value: '10.00', style: { fillColor: '#DDEBF7' } } },
    })

  const downloadPromise = page.waitForEvent('download')
  await editor.locator('button.qa-btn').first().click()
  const savedPath = await (await downloadPromise).path()
  expect(savedPath).not.toBeNull()
  const savedZip = await JSZip.loadAsync(
    await import('node:fs/promises').then((fs) => fs.readFile(savedPath!)),
  )
  const sourceXml = await savedZip.file('xl/worksheets/sheet1.xml')!.async('text')
  const destinationXml = await savedZip.file('xl/worksheets/sheet2.xml')!.async('text')
  const stylesXml = await savedZip.file('xl/styles.xml')!.async('text')
  const savedCell = (xml: string, address: string) =>
    new RegExp(`<c\\b[^>]*\\br="${address}"[^>]*>(?:[\\s\\S]*?)<\\/c>`).exec(xml)?.[0] ?? ''
  const styleIndex = (cell: string) => Number(/\bs="([0-9]+)"/.exec(cell)?.[1] ?? '0')
  const cellXfs = /<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/.exec(stylesXml)?.[1] ?? ''
  const styleRecords = cellXfs.match(/<xf\b[^>]*(?:\/>|>[\s\S]*?<\/xf>)/g) ?? []
  const borders = /<borders\b[^>]*>([\s\S]*?)<\/borders>/.exec(stylesXml)?.[1] ?? ''
  const borderRecords =
    borders.match(/<border(?!s)(?:\s[^>]*)?\/>|<border(?!s)(?:\s[^>]*)?>[\s\S]*?<\/border>/g) ?? []
  const styleRecord = (cell: string) => styleRecords[styleIndex(cell)] ?? ''
  const styleAttribute = (record: string, name: string) =>
    new RegExp(`\\b${name}="([^"]+)"`).exec(record)?.[1] ?? '0'
  const sourceA1 = savedCell(sourceXml, 'A1')
  const sourceB1 = savedCell(sourceXml, 'B1')
  const destinationD3 = savedCell(destinationXml, 'D3')
  const destinationE3 = savedCell(destinationXml, 'E3')
  expect(destinationD3).toContain('<v>10</v>')
  expect(sourceB1).toContain('<f>A1+$A$1+A$1+$A1</f>')
  expect(destinationE3).toContain('<f>D3+$A$1+D$1+$A3</f>')
  for (const attribute of ['fontId', 'fillId', 'numFmtId']) {
    expect(styleAttribute(styleRecord(destinationD3), attribute)).toBe(
      styleAttribute(styleRecord(sourceA1), attribute),
    )
  }
  const sourceBorderId = Number(styleAttribute(styleRecord(sourceA1), 'borderId'))
  const destinationBorderId = Number(styleAttribute(styleRecord(destinationD3), 'borderId'))
  expect(destinationBorderId).not.toBe(sourceBorderId)
  expect(borderRecords[destinationBorderId]).toContain('style="medium"')
  expect(borderRecords[destinationBorderId]).toMatch(/rgb="(?:FF)?4472C4"/)
})

test('the XLSX copy-column-widths operation copies explicit widths through shared undo and save', async ({
  page,
}) => {
  await page.goto('/?format=xlsx&width=720&height=900')
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)
  const editor = page.frameLocator('#editor-frame')
  await editor.locator('input[type="file"][accept*=".xlsx"]').setInputFiles({
    name: 'registry-copy-column-widths.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: await xlsxWithCopyColumnWidthsSource(),
  })

  const setupCommands = [
    { sheet: 'Sheet1', column: 'A', widthCharacters: 18.5 },
    { sheet: 'Sheet1', column: 'B', widthCharacters: 7.25 },
    { sheet: 'Archive', column: 'D', widthCharacters: 4 },
    { sheet: 'Archive', column: 'E', widthCharacters: 30 },
  ] as const
  for (const [index, { sheet, column, widthCharacters }] of setupCommands.entries()) {
    await page.evaluate(
      ({ index, sheet, column, widthCharacters }) => {
        window.__codexVisualHost.commands.push({
          commandId: `xlsx-copy-column-widths-setup-${index}`,
          baseRevision: index,
          operation: 'xlsx.column.set_width',
          arguments: { sheet, column, count: 1, widthCharacters },
        })
      },
      { index, sheet, column, widthCharacters },
    )
    await expect
      .poll(() =>
        page.evaluate(
          (commandId) =>
            window.__codexVisualHost.acknowledgements.find(
              (entry) => entry.commandId === commandId,
            ),
          `xlsx-copy-column-widths-setup-${index}`,
        ),
      )
      .toMatchObject({ ok: true, revision: index + 1 })
  }

  await page.evaluate(() => {
    window.__codexVisualHost.commands.push({
      commandId: 'xlsx-copy-column-widths',
      baseRevision: 4,
      operation: 'xlsx.column.copy_widths',
      arguments: {
        sourceSheet: 'Sheet1',
        sourceColumn: 'A',
        destinationSheet: 'Archive',
        destinationColumn: 'D',
        count: 2,
      },
    })
  })
  await expect
    .poll(() =>
      page.evaluate(() =>
        window.__codexVisualHost.acknowledgements.find(
          (entry) => entry.commandId === 'xlsx-copy-column-widths',
        ),
      ),
    )
    .toMatchObject({
      ok: true,
      revision: 5,
      output: {
        sourceSheet: 'Sheet1',
        sourceColumn: 'A',
        destinationSheet: 'Archive',
        destinationColumn: 'D',
        count: 2,
      },
    })

  await editor.locator('button.qa-btn').nth(1).click()
  await expect(editor.locator('button.qa-btn').nth(2)).toBeEnabled()
  await editor.locator('button.qa-btn').nth(2).click()
  await expect(editor.locator('button.qa-btn').nth(1)).toBeEnabled()
  const redoDownloadPromise = page.waitForEvent('download')
  await editor.locator('button.qa-btn').first().click()
  const redoPath = await (await redoDownloadPromise).path()
  expect(redoPath).not.toBeNull()
  const redoZip = await JSZip.loadAsync(
    await import('node:fs/promises').then((fs) => fs.readFile(redoPath!)),
  )
  const sourceXml = await redoZip.file('xl/worksheets/sheet1.xml')!.async('text')
  const destinationXml = await redoZip.file('xl/worksheets/sheet2.xml')!.async('text')
  expect(savedColumnWidth(destinationXml, 4)).toBe(savedColumnWidth(sourceXml, 1))
  expect(savedColumnWidth(destinationXml, 5)).toBe(savedColumnWidth(sourceXml, 2))
  expect(/<c r="D1"[^>]*>[\s\S]*?<v>99<\/v>/.test(destinationXml)).toBe(true)
  expect(/<c r="E1"[^>]*>[\s\S]*?<f>D1\*2<\/f>[\s\S]*?<v>198<\/v>/.test(destinationXml)).toBe(true)
})

test('a user grid edit is preserved by the browser XLSX save path', async ({ page }) => {
  await page.goto('/?format=xlsx&width=720&height=900')
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)
  const editor = page.frameLocator('#editor-frame')
  await editor.locator('input[type="file"][accept*=".xlsx"]').setInputFiles({
    name: 'user-edit.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: await xlsxWithA1('Before'),
  })
  await expect(editor.locator('.workbook-status')).toContainText('user-edit.xlsx')

  const canvas = editor.locator('canvas[id^="univer-sheet-main-canvas_"]')
  await expect(canvas).toBeVisible()
  await canvas.click({ position: { x: 70, y: 35 } })
  await page.keyboard.type('User edit', { delay: 30 })
  await page.keyboard.press('Enter')

  const downloadPromise = page.waitForEvent('download')
  await editor.locator('button.qa-btn').first().click()
  const download = await downloadPromise
  const savedPath = await download.path()
  expect(savedPath).not.toBeNull()
  const zip = await JSZip.loadAsync(
    await import('node:fs/promises').then((fs) => fs.readFile(savedPath!)),
  )
  const sheetXml = await zip.file('xl/worksheets/sheet1.xml')!.async('text')
  expect(sheetXml).toContain('User edit')
  expect(sheetXml).not.toContain('Before')
})

test('a user formula remains a formula after browser XLSX save', async ({ page }) => {
  await page.goto('/?format=xlsx&width=720&height=900')
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)
  const editor = page.frameLocator('#editor-frame')
  await editor.locator('input[type="file"][accept*=".xlsx"]').setInputFiles({
    name: 'formula.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: await xlsxWithA1('Before'),
  })
  await expect(editor.locator('.workbook-status')).toContainText('formula.xlsx')

  const canvas = editor.locator('canvas[id^="univer-sheet-main-canvas_"]')
  await expect(canvas).toBeVisible()
  await canvas.click({ position: { x: 70, y: 35 } })
  await page.keyboard.type('=1+1', { delay: 30 })
  await page.keyboard.press('Enter')

  const downloadPromise = page.waitForEvent('download')
  await editor.locator('button.qa-btn').first().click()
  const savedPath = await (await downloadPromise).path()
  expect(savedPath).not.toBeNull()
  const zip = await JSZip.loadAsync(
    await import('node:fs/promises').then((fs) => fs.readFile(savedPath!)),
  )
  const sheetXml = await zip.file('xl/worksheets/sheet1.xml')!.async('text')
  expect(sheetXml).toContain('<f>1+1</f>')
  expect(sheetXml).toContain('<v>2</v>')
})

test('an MCP range edit updates the shared grid and saved XLSX', async ({ page }) => {
  await page.goto('/?format=xlsx&width=720&height=900')
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)
  const editor = page.frameLocator('#editor-frame')
  await editor.locator('input[type="file"][accept*=".xlsx"]').setInputFiles({
    name: 'range-edit.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: await xlsxWithA1('Before'),
  })
  await expect(editor.locator('.workbook-status')).toContainText('range-edit.xlsx')

  await page.evaluate(() => {
    window.__codexVisualHost.commands.push({
      commandId: 'xlsx-set-range',
      baseRevision: 0,
      operation: 'xlsx.range.set_values',
      arguments: {
        sheet: 'Sheet1',
        range: 'A1:B2',
        values: [
          ['North', 10],
          ['South', 20],
        ],
      },
    })
  })
  await expect
    .poll(() =>
      page.evaluate(() =>
        window.__codexVisualHost.acknowledgements.find(
          (entry) => entry.commandId === 'xlsx-set-range',
        ),
      ),
    )
    .toMatchObject({ ok: true, revision: 1 })

  const downloadPromise = page.waitForEvent('download')
  await editor.locator('button.qa-btn').first().click()
  const savedPath = await (await downloadPromise).path()
  expect(savedPath).not.toBeNull()
  const zip = await JSZip.loadAsync(
    await import('node:fs/promises').then((fs) => fs.readFile(savedPath!)),
  )
  const sheetXml = await zip.file('xl/worksheets/sheet1.xml')!.async('text')
  expect(sheetXml).toContain('North')
  expect(sheetXml).toContain('South')
  expect(sheetXml).toContain('<v>10</v>')
  expect(sheetXml).toContain('<v>20</v>')
})

test('the MCP save operation uses the community browser save path', async ({ page }) => {
  await page.goto('/?format=xlsx&width=720&height=900')
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)
  const editor = page.frameLocator('#editor-frame')
  await editor.locator('input[type="file"][accept*=".xlsx"]').setInputFiles({
    name: 'mcp-save.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: await xlsxWithA1('Saved through MCP'),
  })
  await expect(editor.locator('.workbook-status')).toContainText('mcp-save.xlsx')

  const downloadPromise = page.waitForEvent('download')
  await page.evaluate(() => {
    window.__codexVisualHost.commands.push({
      commandId: 'xlsx-save',
      baseRevision: 0,
      operation: 'xlsx.document.save',
      arguments: {},
    })
  })
  const download = await downloadPromise
  expect(download.suggestedFilename()).toBe('mcp-save.xlsx')
  await expect
    .poll(() =>
      page.evaluate(() =>
        window.__codexVisualHost.acknowledgements.find((entry) => entry.commandId === 'xlsx-save'),
      ),
    )
    .toMatchObject({ ok: true, revision: 1, output: { saved: true } })
})

test('an edited XLSX saves through local persistence when an iframe picker is unavailable', async ({
  page,
}) => {
  await page.goto('/?format=xlsx&width=720&height=900')
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)
  const editor = page.frameLocator('#editor-frame')
  await editor.locator('input[type="file"][accept*=".xlsx"]').setInputFiles({
    name: 'host-save.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: await xlsxWithA1('Before'),
  })
  await expect(editor.locator('.workbook-status')).toContainText('host-save.xlsx')
  await editor.locator('body').evaluate((body) => {
    delete (
      body.ownerDocument.defaultView as Window & {
        showSaveFilePicker?: unknown
      }
    ).showSaveFilePicker
  })

  await page.evaluate(() => {
    window.__codexVisualHost.enqueueCommand({
      commandId: 'xlsx-host-save-content',
      baseRevision: 0,
      operation: 'xlsx.cell.set_value',
      arguments: { sheet: 'Sheet1', address: 'A1', value: 'Saved by host' },
    })
  })
  await expect
    .poll(() =>
      page.evaluate(() =>
        window.__codexVisualHost.acknowledgements.find(
          (entry) => entry.commandId === 'xlsx-host-save-content',
        ),
      ),
    )
    .toMatchObject({ ok: true, revision: 1 })

  const downloadPromise = page.waitForEvent('download', { timeout: 5_000 })
  await page.evaluate(() => {
    window.__codexVisualHost.enqueueCommand({
      commandId: 'xlsx-host-save',
      baseRevision: 1,
      operation: 'xlsx.document.save',
      arguments: {},
    })
  })
  await expect
    .poll(() =>
      page.evaluate(() =>
        window.__codexVisualHost.acknowledgements.find(
          (entry) => entry.commandId === 'xlsx-host-save',
        ),
      ),
    )
    .toMatchObject({ ok: true, revision: 2, output: { saved: true, fileName: 'host-save.xlsx' } })
  const download = await downloadPromise

  expect(download.suggestedFilename()).toBe('host-save.xlsx')
  await expect
    .poll(() => page.evaluate(() => window.__codexVisualHost.downloads.at(-1)))
    .toMatchObject({
      fileName: 'host-save.xlsx',
      mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      size: expect.any(Number),
    })
  const savedPath = await download.path()
  expect(savedPath).not.toBeNull()
  const zip = await JSZip.loadAsync(
    await import('node:fs/promises').then((fs) => fs.readFile(savedPath!)),
  )
  expect(await zip.file('xl/worksheets/sheet1.xml')!.async('text')).toContain('Saved by host')
  expect(await page.evaluate(() => window.__codexVisualHost.events)).toEqual(
    expect.arrayContaining([
      'tool:office_editor_begin_document_save',
      'tool:office_editor_commit_document_save',
    ]),
  )
})

test('an XLSX hyperlink registry edit survives typed save and browser reopen', async ({ page }) => {
  await page.goto('/?format=xlsx&width=720&height=900')
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)
  const editor = page.frameLocator('#editor-frame')
  await editor.locator('input[type="file"][accept*=".xlsx"]').setInputFiles({
    name: 'registry-hyperlink.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: await xlsxWithA1('Open report'),
  })

  await page.evaluate(() => {
    window.__codexVisualHost.commands.push({
      commandId: 'xlsx-hyperlink-set',
      baseRevision: 0,
      operation: 'xlsx.hyperlink.set',
      arguments: {
        sheet: 'Sheet1',
        address: 'A1',
        target: 'https://example.com/report',
      },
    })
  })
  await expect
    .poll(() =>
      page.evaluate(() =>
        window.__codexVisualHost.acknowledgements.find(
          (entry) => entry.commandId === 'xlsx-hyperlink-set',
        ),
      ),
    )
    .toMatchObject({ ok: true, revision: 1 })

  const downloadPromise = page.waitForEvent('download')
  await page.evaluate(() => {
    window.__codexVisualHost.commands.push({
      commandId: 'xlsx-hyperlink-save',
      baseRevision: 1,
      operation: 'xlsx.document.save',
      arguments: {},
    })
  })
  const savedPath = await (await downloadPromise).path()
  expect(savedPath).not.toBeNull()
  await expect
    .poll(() =>
      page.evaluate(() =>
        window.__codexVisualHost.acknowledgements.find(
          (entry) => entry.commandId === 'xlsx-hyperlink-save',
        ),
      ),
    )
    .toMatchObject({ ok: true, revision: 2, output: { saved: true } })

  const savedBytes = await import('node:fs/promises').then((fs) => fs.readFile(savedPath!))
  const zip = await JSZip.loadAsync(savedBytes)
  const sheetXml = await zip.file('xl/worksheets/sheet1.xml')!.async('text')
  const relsXml = await zip.file('xl/worksheets/_rels/sheet1.xml.rels')!.async('text')
  expect(sheetXml).toMatch(/<hyperlink\b[^>]*\bref="A1"[^>]*\br:id="rId1"/)
  expect(relsXml).toContain('Target="https://example.com/report"')

  await editor.locator('input[type="file"][accept*=".xlsx"]').setInputFiles({
    name: 'registry-hyperlink-reopened.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: savedBytes,
  })
  await expect(editor.locator('.workbook-status')).toContainText('registry-hyperlink-reopened.xlsx')
  const canvas = editor.locator('canvas[id^="univer-sheet-main-canvas_"]')
  await canvas.click({ position: { x: 135, y: 35 } })
  const popupPromise = page.waitForEvent('popup')
  await canvas.click({ position: { x: 70, y: 35 } })
  const popup = await popupPromise
  await expect(popup).toHaveURL('https://example.com/report')
  await popup.close()
})

test('an XLSX table registry edit survives typed save and browser reopen', async ({ page }) => {
  await page.goto('/?format=xlsx&width=720&height=900')
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)
  const editor = page.frameLocator('#editor-frame')
  await editor.locator('input[type="file"][accept*=".xlsx"]').setInputFiles({
    name: 'registry-table.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: await xlsxWithA1('placeholder'),
  })

  const commands = [
    {
      commandId: 'xlsx-table-values',
      operation: 'xlsx.range.set_values',
      arguments: {
        sheet: 'Sheet1',
        range: 'A1:B3',
        values: [
          ['Region', 'Amount'],
          ['North', 10],
          ['South', 20],
        ],
      },
    },
    {
      commandId: 'xlsx-table-add',
      operation: 'xlsx.table.add',
      arguments: { sheet: 'Sheet1', range: 'A1:B3', style: 'TableStyleMedium4' },
    },
  ] as const
  for (const [index, command] of commands.entries()) {
    await page.evaluate(
      ({ index, command }) => {
        window.__codexVisualHost.commands.push({
          commandId: command.commandId,
          baseRevision: index,
          operation: command.operation,
          arguments: command.arguments,
        })
      },
      { index, command },
    )
    await expect
      .poll(() =>
        page.evaluate(
          (commandId) =>
            window.__codexVisualHost.acknowledgements.find(
              (entry) => entry.commandId === commandId,
            ),
          command.commandId,
        ),
      )
      .toMatchObject({ ok: true, revision: index + 1 })
  }

  const downloadPromise = page.waitForEvent('download')
  await page.evaluate(() => {
    window.__codexVisualHost.commands.push({
      commandId: 'xlsx-table-save',
      baseRevision: 2,
      operation: 'xlsx.document.save',
      arguments: {},
    })
  })
  const savedPath = await (await downloadPromise).path()
  expect(savedPath).not.toBeNull()
  await expect
    .poll(() =>
      page.evaluate(() =>
        window.__codexVisualHost.acknowledgements.find(
          (entry) => entry.commandId === 'xlsx-table-save',
        ),
      ),
    )
    .toMatchObject({ ok: true, revision: 3, output: { saved: true } })

  const savedBytes = await import('node:fs/promises').then((fs) => fs.readFile(savedPath!))
  const zip = await JSZip.loadAsync(savedBytes)
  const tableXml = await zip.file('xl/tables/table1.xml')!.async('text')
  const sheetXml = await zip.file('xl/worksheets/sheet1.xml')!.async('text')
  const relsXml = await zip.file('xl/worksheets/_rels/sheet1.xml.rels')!.async('text')
  expect(tableXml).toMatch(/<table\b[^>]*\bdisplayName="Table1"[^>]*\bref="A1:B3"/)
  expect(tableXml).toContain('name="TableStyleMedium4"')
  expect(sheetXml).toMatch(/<tablePart\b[^>]*\br:id="rId1"/)
  expect(relsXml).toContain('Target="../tables/table1.xml"')

  await editor.locator('input[type="file"][accept*=".xlsx"]').setInputFiles({
    name: 'registry-table-reopened.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: savedBytes,
  })
  await expect(editor.locator('.workbook-status')).toContainText('registry-table-reopened.xlsx')
  await expect
    .poll(() => {
      return page.evaluate(() => {
        const selection = window.__codexVisualHost.lastPollArguments?.selection as
          { activeCell?: { value?: unknown } } | undefined
        return selection?.activeCell?.value
      })
    })
    .toBe('Region')
})

test('an XLSX chart registry edit renders and survives typed browser save', async ({ page }) => {
  await page.goto('/?format=xlsx&width=720&height=900')
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)
  const editor = page.frameLocator('#editor-frame')
  await editor.locator('input[type="file"][accept*=".xlsx"]').setInputFiles({
    name: 'registry-chart.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: await xlsxWithA1('placeholder'),
  })

  const commands = [
    {
      commandId: 'xlsx-chart-values',
      operation: 'xlsx.range.set_values',
      arguments: {
        sheet: 'Sheet1',
        range: 'A1:B3',
        values: [
          ['Quarter', 'Revenue'],
          ['Q1', 10],
          ['Q2', 20],
        ],
      },
    },
    {
      commandId: 'xlsx-chart-add',
      operation: 'xlsx.chart.add',
      arguments: { sheet: 'Sheet1', dataRange: 'A1:B3', type: 'line', anchorCell: 'D2' },
    },
  ] as const
  for (const [index, command] of commands.entries()) {
    await page.evaluate(
      ({ index, command }) => {
        window.__codexVisualHost.commands.push({
          commandId: command.commandId,
          baseRevision: index,
          operation: command.operation,
          arguments: command.arguments,
        })
      },
      { index, command },
    )
    await expect
      .poll(() =>
        page.evaluate(
          (commandId) =>
            window.__codexVisualHost.acknowledgements.find(
              (entry) => entry.commandId === commandId,
            ),
          command.commandId,
        ),
      )
      .toMatchObject({ ok: true, revision: index + 1 })
  }
  await expect(editor.locator('.xlsx-chart')).toHaveCount(1)

  const chartId = await page.evaluate(() => {
    const acknowledgement = window.__codexVisualHost.acknowledgements.find(
      (entry) => entry.commandId === 'xlsx-chart-add',
    )
    return (acknowledgement?.output as { visualId?: string } | undefined)?.visualId
  })
  expect(chartId).toMatch(/^added-chart-/)
  await page.evaluate((chartId) => {
    window.__codexVisualHost.commands.push({
      commandId: 'xlsx-chart-update',
      baseRevision: 2,
      operation: 'xlsx.chart.update',
      arguments: {
        chartId,
        title: 'Updated Revenue',
        type: 'bar',
        legend: 'bottom',
      },
    })
  }, chartId!)
  await expect
    .poll(() =>
      page.evaluate(() =>
        window.__codexVisualHost.acknowledgements.find(
          (entry) => entry.commandId === 'xlsx-chart-update',
        ),
      ),
    )
    .toMatchObject({ ok: true, revision: 3 })
  await page.evaluate((chartId) => {
    window.__codexVisualHost.commands.push({
      commandId: 'xlsx-chart-series',
      baseRevision: 3,
      operation: 'xlsx.chart.set_series',
      arguments: {
        chartId,
        series: [{ name: 'Net Revenue', values: [30, 40], categories: ['H1', 'H2'] }],
      },
    })
  }, chartId!)
  await expect
    .poll(() =>
      page.evaluate(() =>
        window.__codexVisualHost.acknowledgements.find(
          (entry) => entry.commandId === 'xlsx-chart-series',
        ),
      ),
    )
    .toMatchObject({ ok: true, revision: 4 })
  await page.evaluate((chartId) => {
    window.__codexVisualHost.commands.push({
      commandId: 'xlsx-chart-colors',
      baseRevision: 4,
      operation: 'xlsx.chart.set_colors',
      arguments: { chartId, seriesColors: ['#112233'] },
    })
  }, chartId!)
  await expect
    .poll(() =>
      page.evaluate(() =>
        window.__codexVisualHost.acknowledgements.find(
          (entry) => entry.commandId === 'xlsx-chart-colors',
        ),
      ),
    )
    .toMatchObject({ ok: true, revision: 5 })

  const downloadPromise = page.waitForEvent('download')
  await page.evaluate(() => {
    window.__codexVisualHost.commands.push({
      commandId: 'xlsx-chart-save',
      baseRevision: 5,
      operation: 'xlsx.document.save',
      arguments: {},
    })
  })
  const savedPath = await (await downloadPromise).path()
  expect(savedPath).not.toBeNull()
  await expect
    .poll(() =>
      page.evaluate(() =>
        window.__codexVisualHost.acknowledgements.find(
          (entry) => entry.commandId === 'xlsx-chart-save',
        ),
      ),
    )
    .toMatchObject({ ok: true, revision: 6, output: { saved: true } })

  const savedBytes = await import('node:fs/promises').then((fs) => fs.readFile(savedPath!))
  const zip = await JSZip.loadAsync(savedBytes)
  const chartXml = await zip.file('xl/charts/chart1.xml')!.async('text')
  const drawingXml = await zip.file('xl/drawings/drawing1.xml')!.async('text')
  expect(chartXml).toContain('<c:barChart>')
  expect(chartXml).toContain('<c:barDir val="bar"/>')
  expect(chartXml).toContain('<a:t>Updated Revenue</a:t>')
  expect(chartXml).toContain('<c:legendPos val="b"/>')
  expect(chartXml).toContain('<c:v>Net Revenue</c:v>')
  expect(chartXml).toContain('<a:srgbClr val="112233"/>')
  expect(chartXml).toContain('<c:v>40</c:v>')
  expect(drawingXml).toContain('<xdr:from><xdr:col>3</xdr:col>')

  await editor.locator('input[type="file"][accept*=".xlsx"]').setInputFiles({
    name: 'registry-chart-reopened.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: savedBytes,
  })
  await expect
    .poll(() => page.evaluate(() => window.__codexVisualHost.lastPollArguments?.fileName))
    .toBe('registry-chart-reopened.xlsx')

  await expect
    .poll(() =>
      page.evaluate(() => {
        const selection = window.__codexVisualHost.lastPollArguments?.selection as
          { visuals?: { id: string; title?: string }[] } | undefined
        return selection?.visuals?.[0]
      }),
    )
    .toMatchObject({ id: 'file-chart-xl-charts-chart1-xml', title: 'Updated Revenue' })

  await page.evaluate(() => {
    window.__codexVisualHost.commands.push({
      commandId: 'xlsx-chart-reopened-update',
      baseRevision: 6,
      operation: 'xlsx.chart.update',
      arguments: {
        chartId: 'file-chart-xl-charts-chart1-xml',
        title: 'Reopened Revenue',
        gridlines: true,
        valueAxisMin: 0,
        valueAxisMax: 100,
        gapWidthPct: 80,
        dataLabels: 'value',
        dataLabelPosition: 'outside-end',
        dataLabelFormat: '0.0%',
      },
    })
  })
  await expect
    .poll(() =>
      page.evaluate(() =>
        window.__codexVisualHost.acknowledgements.find(
          (entry) => entry.commandId === 'xlsx-chart-reopened-update',
        ),
      ),
    )
    .toMatchObject({ ok: true, revision: 7 })

  const reopenedDownloadPromise = page.waitForEvent('download')
  await page.evaluate(() => {
    window.__codexVisualHost.commands.push({
      commandId: 'xlsx-chart-reopened-save',
      baseRevision: 7,
      operation: 'xlsx.document.save',
      arguments: {},
    })
  })
  const reopenedSavedPath = await (await reopenedDownloadPromise).path()
  expect(reopenedSavedPath).not.toBeNull()
  await expect
    .poll(() =>
      page.evaluate(() =>
        window.__codexVisualHost.acknowledgements.find(
          (entry) => entry.commandId === 'xlsx-chart-reopened-save',
        ),
      ),
    )
    .toMatchObject({ ok: true, revision: 8, output: { saved: true } })
  const reopenedZip = await JSZip.loadAsync(
    await import('node:fs/promises').then((fs) => fs.readFile(reopenedSavedPath!)),
  )
  const reopenedChartXml = await reopenedZip.file('xl/charts/chart1.xml')!.async('text')
  expect(reopenedChartXml).toContain('<a:t>Reopened Revenue</a:t>')
  expect(reopenedChartXml).toContain('<c:majorGridlines/>')
  expect(reopenedChartXml).toContain('<c:max val="100"/>')
  expect(reopenedChartXml).toContain('<c:min val="0"/>')
  expect(reopenedChartXml).toContain('<c:gapWidth val="80"/>')
  expect(reopenedChartXml).toContain('<c:dLblPos val="outEnd"/>')
  expect(reopenedChartXml).toContain('<c:numFmt formatCode="0.0%" sourceLinked="0"/>')
})

test('an XLSX shape registry add shares undo and survives typed browser save/reopen', async ({
  page,
}) => {
  await page.goto('/?format=xlsx&width=720&height=900')
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)
  const editor = page.frameLocator('#editor-frame')
  await editor.locator('input[type="file"][accept*=".xlsx"]').setInputFiles({
    name: 'registry-shape.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: await xlsxWithA1('Shape host'),
  })

  await page.evaluate(() => {
    window.__codexVisualHost.commands.push({
      commandId: 'xlsx-shape-add',
      baseRevision: 0,
      operation: 'xlsx.shape.add',
      arguments: {
        sheet: 'Sheet1',
        type: 'roundRect',
        anchorCell: 'D2',
        fillColor: '#DDEBF7',
        text: 'Forecast',
      },
    })
  })
  await expect
    .poll(() =>
      page.evaluate(() =>
        window.__codexVisualHost.acknowledgements.find(
          (entry) => entry.commandId === 'xlsx-shape-add',
        ),
      ),
    )
    .toMatchObject({
      ok: true,
      revision: 1,
      output: { visualId: expect.stringMatching(/^added-shape-/) },
    })
  await expect(editor.locator('.xlsx-shape-drawn')).toHaveCount(1)

  await editor.locator('button.qa-btn').nth(1).click()
  await expect(editor.locator('.xlsx-shape-drawn')).toHaveCount(0)
  await editor.locator('button.qa-btn').nth(2).click()
  await expect(editor.locator('.xlsx-shape-drawn')).toHaveCount(1)

  const downloadPromise = page.waitForEvent('download')
  await page.evaluate(() => {
    window.__codexVisualHost.commands.push({
      commandId: 'xlsx-shape-save',
      baseRevision: 1,
      operation: 'xlsx.document.save',
      arguments: {},
    })
  })
  const savedPath = await (await downloadPromise).path()
  expect(savedPath).not.toBeNull()
  await expect
    .poll(() =>
      page.evaluate(() =>
        window.__codexVisualHost.acknowledgements.find(
          (entry) => entry.commandId === 'xlsx-shape-save',
        ),
      ),
    )
    .toMatchObject({ ok: true, revision: 2, output: { saved: true } })
  const savedBytes = await import('node:fs/promises').then((fs) => fs.readFile(savedPath!))
  const zip = await JSZip.loadAsync(savedBytes)
  const drawingXml = await zip.file('xl/drawings/drawing1.xml')!.async('text')
  expect(drawingXml).toContain('<a:prstGeom prst="roundRect">')
  expect(drawingXml).toContain('<a:srgbClr val="DDEBF7"/>')
  expect(drawingXml).toContain('<a:t>Forecast</a:t>')

  await editor.locator('input[type="file"][accept*=".xlsx"]').setInputFiles({
    name: 'registry-shape-reopened.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: savedBytes,
  })
  await expect(editor.locator('.xlsx-shape-drawn')).toHaveCount(1)
  await expect
    .poll(() =>
      page.evaluate(() => {
        const selection = window.__codexVisualHost.lastPollArguments?.selection as
          { visuals?: { id: string; title?: string }[] } | undefined
        return selection?.visuals?.[0]
      }),
    )
    .toMatchObject({ id: 'file-shape-xl-drawings-drawing1-xml-0', kind: 'shape' })

  await page.evaluate(() => {
    window.__codexVisualHost.commands.push({
      commandId: 'xlsx-shape-reopened-update',
      baseRevision: 2,
      operation: 'xlsx.shape.update',
      arguments: {
        shapeId: 'file-shape-xl-drawings-drawing1-xml-0',
        anchorCell: 'F4',
        fillColor: '#4472C4',
        text: 'Updated forecast',
      },
    })
  })
  await expect
    .poll(() =>
      page.evaluate(() =>
        window.__codexVisualHost.acknowledgements.find(
          (entry) => entry.commandId === 'xlsx-shape-reopened-update',
        ),
      ),
    )
    .toMatchObject({ ok: true, revision: 3 })
  await expect(editor.locator('.xlsx-shape-drawn .shape-text')).toHaveText('Updated forecast')

  const updatedDownloadPromise = page.waitForEvent('download')
  await page.evaluate(() => {
    window.__codexVisualHost.commands.push({
      commandId: 'xlsx-shape-updated-save',
      baseRevision: 3,
      operation: 'xlsx.document.save',
      arguments: {},
    })
  })
  const updatedPath = await (await updatedDownloadPromise).path()
  expect(updatedPath).not.toBeNull()
  await expect
    .poll(() =>
      page.evaluate(() =>
        window.__codexVisualHost.acknowledgements.find(
          (entry) => entry.commandId === 'xlsx-shape-updated-save',
        ),
      ),
    )
    .toMatchObject({ ok: true, revision: 4 })
  const updatedBytes = await import('node:fs/promises').then((fs) => fs.readFile(updatedPath!))
  const updatedZip = await JSZip.loadAsync(updatedBytes)
  const updatedDrawingXml = await updatedZip.file('xl/drawings/drawing1.xml')!.async('text')
  expect(updatedDrawingXml).toContain('<xdr:col>5</xdr:col>')
  expect(updatedDrawingXml).toContain('<a:srgbClr val="4472C4"/>')
  expect(updatedDrawingXml).toContain('<a:t>Updated forecast</a:t>')

  await editor.locator('input[type="file"][accept*=".xlsx"]').setInputFiles({
    name: 'registry-shape-updated.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: updatedBytes,
  })
  await page.evaluate(() => {
    window.__codexVisualHost.commands.push({
      commandId: 'xlsx-shape-reopened-remove',
      baseRevision: 4,
      operation: 'xlsx.shape.remove',
      arguments: { shapeId: 'file-shape-xl-drawings-drawing1-xml-0' },
    })
  })
  await expect
    .poll(() =>
      page.evaluate(() =>
        window.__codexVisualHost.acknowledgements.find(
          (entry) => entry.commandId === 'xlsx-shape-reopened-remove',
        ),
      ),
    )
    .toMatchObject({ ok: true, revision: 5 })
  await expect(editor.locator('.xlsx-shape-drawn')).toHaveCount(0)

  const removedDownloadPromise = page.waitForEvent('download')
  await page.evaluate(() => {
    window.__codexVisualHost.commands.push({
      commandId: 'xlsx-shape-removed-save',
      baseRevision: 5,
      operation: 'xlsx.document.save',
      arguments: {},
    })
  })
  const removedPath = await (await removedDownloadPromise).path()
  expect(removedPath).not.toBeNull()
  const removedZip = await JSZip.loadAsync(
    await import('node:fs/promises').then((fs) => fs.readFile(removedPath!)),
  )
  expect(await removedZip.file('xl/drawings/drawing1.xml')!.async('text')).not.toContain('<xdr:sp ')
})

test('an XLSX chart registry removal cancels a session visual before typed save', async ({
  page,
}) => {
  await page.goto('/?format=xlsx&width=720&height=900')
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)
  const editor = page.frameLocator('#editor-frame')
  await editor.locator('input[type="file"][accept*=".xlsx"]').setInputFiles({
    name: 'registry-chart-remove.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: await xlsxWithA1('placeholder'),
  })
  const commands = [
    {
      commandId: 'xlsx-chart-remove-values',
      operation: 'xlsx.range.set_values',
      arguments: {
        sheet: 'Sheet1',
        range: 'A1:B2',
        values: [
          ['Quarter', 'Revenue'],
          ['Q1', 10],
        ],
      },
    },
    {
      commandId: 'xlsx-chart-remove-add',
      operation: 'xlsx.chart.add',
      arguments: { sheet: 'Sheet1', dataRange: 'A1:B2', type: 'column' },
    },
  ] as const
  for (const [index, command] of commands.entries()) {
    await page.evaluate(
      ({ index, command }) => {
        window.__codexVisualHost.commands.push({
          commandId: command.commandId,
          baseRevision: index,
          operation: command.operation,
          arguments: command.arguments,
        })
      },
      { index, command },
    )
    await expect
      .poll(() =>
        page.evaluate(
          (commandId) =>
            window.__codexVisualHost.acknowledgements.find(
              (entry) => entry.commandId === commandId,
            ),
          command.commandId,
        ),
      )
      .toMatchObject({ ok: true, revision: index + 1 })
  }
  const chartId = await page.evaluate(() => {
    const acknowledgement = window.__codexVisualHost.acknowledgements.find(
      (entry) => entry.commandId === 'xlsx-chart-remove-add',
    )
    return (acknowledgement?.output as { visualId?: string } | undefined)?.visualId
  })
  await page.evaluate((chartId) => {
    window.__codexVisualHost.commands.push({
      commandId: 'xlsx-chart-remove',
      baseRevision: 2,
      operation: 'xlsx.chart.remove',
      arguments: { chartId },
    })
  }, chartId!)
  await expect
    .poll(() =>
      page.evaluate(() =>
        window.__codexVisualHost.acknowledgements.find(
          (entry) => entry.commandId === 'xlsx-chart-remove',
        ),
      ),
    )
    .toMatchObject({ ok: true, revision: 3 })
  await expect(editor.locator('.xlsx-chart')).toHaveCount(0)

  const downloadPromise = page.waitForEvent('download')
  await page.evaluate(() => {
    window.__codexVisualHost.commands.push({
      commandId: 'xlsx-chart-remove-save',
      baseRevision: 3,
      operation: 'xlsx.document.save',
      arguments: {},
    })
  })
  const savedPath = await (await downloadPromise).path()
  expect(savedPath).not.toBeNull()
  const savedBytes = await import('node:fs/promises').then((fs) => fs.readFile(savedPath!))
  const zip = await JSZip.loadAsync(savedBytes)
  expect(zip.file('xl/charts/chart1.xml')).toBeNull()
  expect(zip.file('xl/drawings/drawing1.xml')).toBeNull()
})

test('a staged XLSX image registry edit renders and survives typed browser save', async ({
  page,
}) => {
  await page.goto('/?format=xlsx&width=720&height=900')
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)
  const editor = page.frameLocator('#editor-frame')
  await editor.locator('input[type="file"][accept*=".xlsx"]').setInputFiles({
    name: 'registry-image.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: await xlsxWithA1('image'),
  })

  await page.evaluate(() => {
    const base64 =
      'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Y9ZlS8AAAAASUVORK5CYII='
    const binary = atob(base64)
    window.__codexVisualHost.stagedFiles['playwright-image'] = base64
    window.__codexVisualHost.commands.push({
      commandId: 'xlsx-image-add',
      baseRevision: 0,
      operation: 'xlsx.image.add_staged',
      arguments: {
        blobId: 'playwright-image',
        name: 'pixel.png',
        size: binary.length,
        sheet: 'Sheet1',
        anchorCell: 'C3',
      },
    })
  })
  await page.waitForFunction(() =>
    window.__codexVisualHost.acknowledgements.some((entry) => entry.commandId === 'xlsx-image-add'),
  )
  const imageAck = await page.evaluate(() =>
    window.__codexVisualHost.acknowledgements.find((entry) => entry.commandId === 'xlsx-image-add'),
  )
  expect(imageAck?.ok, JSON.stringify(imageAck)).toBe(true)
  expect(imageAck).toMatchObject({ revision: 1 })
  await expect(editor.locator('.xlsx-image')).toHaveCount(1)

  const downloadPromise = page.waitForEvent('download')
  await page.evaluate(() => {
    window.__codexVisualHost.commands.push({
      commandId: 'xlsx-image-save',
      baseRevision: 1,
      operation: 'xlsx.document.save',
      arguments: {},
    })
  })
  await page.waitForFunction(() =>
    window.__codexVisualHost.acknowledgements.some(
      (entry) => entry.commandId === 'xlsx-image-save',
    ),
  )
  const saveAck = await page.evaluate(() =>
    window.__codexVisualHost.acknowledgements.find(
      (entry) => entry.commandId === 'xlsx-image-save',
    ),
  )
  expect(saveAck?.ok, JSON.stringify(saveAck)).toBe(true)
  const savedPath = await (await downloadPromise).path()
  expect(savedPath).not.toBeNull()
  const savedBytes = await import('node:fs/promises').then((fs) => fs.readFile(savedPath!))
  const zip = await JSZip.loadAsync(savedBytes)
  const media = await zip.file('xl/media/image1.png')!.async('nodebuffer')
  const drawingXml = await zip.file('xl/drawings/drawing1.xml')!.async('text')
  expect(media.subarray(0, 8)).toEqual(
    Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
  )
  expect(drawingXml).toContain('<xdr:pic>')

  await editor.locator('input[type="file"][accept*=".xlsx"]').setInputFiles({
    name: 'registry-image-reopened.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: savedBytes,
  })
  await expect(editor.locator('.workbook-status')).toContainText('registry-image-reopened.xlsx')
  await expect(editor.locator('.xlsx-image')).toHaveCount(1)
  await expect
    .poll(() =>
      page.evaluate(() => {
        const selection = window.__codexVisualHost.lastPollArguments?.selection as
          { visuals?: { id: string; kind: string }[] } | undefined
        return selection?.visuals?.[0]
      }),
    )
    .toMatchObject({ id: 'file-image-xl-media-image1-png-0', kind: 'image' })

  await page.evaluate(() => {
    window.__codexVisualHost.commands.push({
      commandId: 'xlsx-image-reopened-move',
      baseRevision: 2,
      operation: 'xlsx.image.move',
      arguments: { imageId: 'file-image-xl-media-image1-png-0', anchorCell: 'F4' },
    })
  })
  await expect
    .poll(() =>
      page.evaluate(() =>
        window.__codexVisualHost.acknowledgements.find(
          (entry) => entry.commandId === 'xlsx-image-reopened-move',
        ),
      ),
    )
    .toMatchObject({ ok: true, revision: 3 })

  const movedDownloadPromise = page.waitForEvent('download')
  await page.evaluate(() => {
    window.__codexVisualHost.commands.push({
      commandId: 'xlsx-image-moved-save',
      baseRevision: 3,
      operation: 'xlsx.document.save',
      arguments: {},
    })
  })
  const movedPath = await (await movedDownloadPromise).path()
  expect(movedPath).not.toBeNull()
  const movedBytes = await import('node:fs/promises').then((fs) => fs.readFile(movedPath!))
  const movedZip = await JSZip.loadAsync(movedBytes)
  expect(await movedZip.file('xl/drawings/drawing1.xml')!.async('text')).toContain(
    '<xdr:from><xdr:col>5</xdr:col>',
  )

  await editor.locator('input[type="file"][accept*=".xlsx"]').setInputFiles({
    name: 'registry-image-moved.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: movedBytes,
  })
  await page.evaluate(() => {
    window.__codexVisualHost.commands.push({
      commandId: 'xlsx-image-reopened-remove',
      baseRevision: 4,
      operation: 'xlsx.image.remove',
      arguments: { imageId: 'file-image-xl-media-image1-png-0' },
    })
  })
  await expect
    .poll(() =>
      page.evaluate(() =>
        window.__codexVisualHost.acknowledgements.find(
          (entry) => entry.commandId === 'xlsx-image-reopened-remove',
        ),
      ),
    )
    .toMatchObject({ ok: true, revision: 5 })
  await expect(editor.locator('.xlsx-image')).toHaveCount(0)

  const removedDownloadPromise = page.waitForEvent('download')
  await page.evaluate(() => {
    window.__codexVisualHost.commands.push({
      commandId: 'xlsx-image-removed-save',
      baseRevision: 5,
      operation: 'xlsx.document.save',
      arguments: {},
    })
  })
  const removedPath = await (await removedDownloadPromise).path()
  expect(removedPath).not.toBeNull()
  const removedZip = await JSZip.loadAsync(
    await import('node:fs/promises').then((fs) => fs.readFile(removedPath!)),
  )
  expect(await removedZip.file('xl/drawings/drawing1.xml')!.async('text')).not.toContain(
    '<xdr:pic>',
  )
})

test('an XLSX note registry edit survives typed save, reopen, and native removal', async ({
  page,
}) => {
  await page.goto('/?format=xlsx&width=720&height=900')
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)
  const editor = page.frameLocator('#editor-frame')
  await editor.locator('input[type="file"][accept*=".xlsx"]').setInputFiles({
    name: 'registry-note.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: await xlsxWithA1('note'),
  })

  await page.evaluate(() => {
    window.__codexVisualHost.commands.push({
      commandId: 'xlsx-note-set',
      baseRevision: 0,
      operation: 'xlsx.note.set',
      arguments: { sheet: 'Sheet1', address: 'C4', text: 'Check this total.' },
    })
  })
  await page.waitForFunction(() =>
    window.__codexVisualHost.acknowledgements.some((entry) => entry.commandId === 'xlsx-note-set'),
  )
  const noteSetAck = await page.evaluate(() =>
    window.__codexVisualHost.acknowledgements.find((entry) => entry.commandId === 'xlsx-note-set'),
  )
  expect(noteSetAck?.ok, JSON.stringify(noteSetAck)).toBe(true)
  expect(noteSetAck).toMatchObject({ revision: 1 })

  const notedDownload = page.waitForEvent('download')
  await page.evaluate(() => {
    window.__codexVisualHost.commands.push({
      commandId: 'xlsx-note-save',
      baseRevision: 1,
      operation: 'xlsx.document.save',
      arguments: {},
    })
  })
  const notedPath = await (await notedDownload).path()
  expect(notedPath).not.toBeNull()
  const notedBytes = await import('node:fs/promises').then((fs) => fs.readFile(notedPath!))
  const notedZip = await JSZip.loadAsync(notedBytes)
  expect(await notedZip.file('xl/comments1.xml')!.async('text')).toContain(
    '<comment ref="C4" authorId="0"><text><t xml:space="preserve">Check this total.</t>',
  )
  expect(await notedZip.file('xl/drawings/vmlDrawing1.vml')!.async('text')).toContain(
    '<x:Row>3</x:Row><x:Column>2</x:Column>',
  )

  await editor.locator('input[type="file"][accept*=".xlsx"]').setInputFiles({
    name: 'registry-note-reopened.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: notedBytes,
  })
  await expect(editor.locator('.workbook-status')).toContainText('registry-note-reopened.xlsx')

  await page.evaluate(() => {
    window.__codexVisualHost.commands.push({
      commandId: 'xlsx-note-remove',
      baseRevision: 0,
      operation: 'xlsx.note.remove',
      arguments: { sheet: 'Sheet1', address: 'C4' },
    })
  })
  await expect
    .poll(() =>
      page.evaluate(() =>
        window.__codexVisualHost.acknowledgements.find(
          (entry) => entry.commandId === 'xlsx-note-remove',
        ),
      ),
    )
    .toMatchObject({ ok: true, revision: 1 })

  const clearedDownload = page.waitForEvent('download')
  await page.evaluate(() => {
    window.__codexVisualHost.commands.push({
      commandId: 'xlsx-note-clear-save',
      baseRevision: 1,
      operation: 'xlsx.document.save',
      arguments: {},
    })
  })
  const clearedPath = await (await clearedDownload).path()
  expect(clearedPath).not.toBeNull()
  const clearedBytes = await import('node:fs/promises').then((fs) => fs.readFile(clearedPath!))
  const clearedZip = await JSZip.loadAsync(clearedBytes)
  expect(clearedZip.file('xl/comments1.xml')).toBeNull()
  expect(clearedZip.file('xl/drawings/vmlDrawing1.vml')).toBeNull()
})

test('browser reopen hydrates native XLSX table metadata into the mounted view', async ({
  page,
}) => {
  await page.goto('/?format=xlsx&width=720&height=900')
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)
  const editor = page.frameLocator('#editor-frame')
  await editor.locator('input[type="file"][accept*=".xlsx"]').setInputFiles({
    name: 'native-table.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: await xlsxWithTableSource(),
  })

  const canvas = editor.locator('canvas[id^="univer-sheet-main-canvas_"]')
  await canvas.click({ position: { x: 135, y: 75 } })
  await canvas.click({ position: { x: 70, y: 35 } })
  await expect(editor.locator('button[data-tip="加粗"]')).toHaveClass(/is-active/)
})

test('XLSX cell and sheet protection registry edits survive typed save and reopen', async ({
  page,
}) => {
  await page.goto('/?format=xlsx&width=720&height=900')
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)
  const editor = page.frameLocator('#editor-frame')
  await editor.locator('input[type="file"][accept*=".xlsx"]').setInputFiles({
    name: 'registry-protection.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: await xlsxWithA1('protected'),
  })

  const commands = [
    {
      commandId: 'xlsx-cell-protection',
      operation: 'xlsx.range.set_protection',
      arguments: {
        sheet: 'Sheet1',
        range: 'A1',
        protection: { locked: false, hidden: true },
        fields: ['locked', 'hidden'],
      },
    },
    {
      commandId: 'xlsx-sheet-protection',
      operation: 'xlsx.sheet.set_protection',
      arguments: { sheet: 'Sheet1', protected: true },
    },
  ] as const
  for (const [index, command] of commands.entries()) {
    await page.evaluate(
      ({ index, command }) => {
        window.__codexVisualHost.commands.push({
          commandId: command.commandId,
          baseRevision: index,
          operation: command.operation,
          arguments: command.arguments,
        })
      },
      { index, command },
    )
    await expect
      .poll(() =>
        page.evaluate(
          (commandId) =>
            window.__codexVisualHost.acknowledgements.find(
              (entry) => entry.commandId === commandId,
            ),
          command.commandId,
        ),
      )
      .toMatchObject({ ok: true, revision: index + 1 })
  }

  const downloadPromise = page.waitForEvent('download')
  await page.evaluate(() => {
    window.__codexVisualHost.commands.push({
      commandId: 'xlsx-protection-save',
      baseRevision: 2,
      operation: 'xlsx.document.save',
      arguments: {},
    })
  })
  const savedPath = await (await downloadPromise).path()
  expect(savedPath).not.toBeNull()
  const savedBytes = await import('node:fs/promises').then((fs) => fs.readFile(savedPath!))
  const zip = await JSZip.loadAsync(savedBytes)
  const sheetXml = await zip.file('xl/worksheets/sheet1.xml')!.async('text')
  const stylesXml = await zip.file('xl/styles.xml')!.async('text')
  expect(sheetXml).toContain('<sheetProtection sheet="1"')
  expect(stylesXml).toContain('<protection locked="0" hidden="1"/>')

  await editor.locator('input[type="file"][accept*=".xlsx"]').setInputFiles({
    name: 'registry-protection-reopened.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: savedBytes,
  })
  await expect(editor.locator('.workbook-status')).toContainText(
    'registry-protection-reopened.xlsx',
  )
  await page.evaluate(() => {
    window.__codexVisualHost.commands.push({
      commandId: 'xlsx-sheet-unprotect-reopened',
      baseRevision: 3,
      operation: 'xlsx.sheet.set_protection',
      arguments: { sheet: 'Sheet1', protected: false },
    })
  })
  await expect
    .poll(() =>
      page.evaluate(() =>
        window.__codexVisualHost.acknowledgements.find(
          (entry) => entry.commandId === 'xlsx-sheet-unprotect-reopened',
        ),
      ),
    )
    .toMatchObject({ ok: true, revision: 4 })
})

test('an XLSX sparkline registry edit survives typed save and browser reopen', async ({ page }) => {
  await page.goto('/?format=xlsx&width=720&height=900')
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)
  const editor = page.frameLocator('#editor-frame')
  await editor.locator('input[type="file"][accept*=".xlsx"]').setInputFiles({
    name: 'registry-sparkline.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: await xlsxWithA1('placeholder'),
  })

  const commands = [
    {
      commandId: 'xlsx-sparkline-values',
      operation: 'xlsx.range.set_values',
      arguments: {
        sheet: 'Sheet1',
        range: 'A1:C2',
        values: [
          [1, 2, 3],
          [3, -1, 2],
        ],
      },
    },
    {
      commandId: 'xlsx-sparkline-add',
      operation: 'xlsx.sparkline.add',
      arguments: {
        sheet: 'Sheet1',
        sourceRange: 'A1:C2',
        targetRange: 'D1:D2',
        type: 'column',
      },
    },
  ] as const
  for (const [index, command] of commands.entries()) {
    await page.evaluate(
      ({ index, command }) => {
        window.__codexVisualHost.commands.push({
          commandId: command.commandId,
          baseRevision: index,
          operation: command.operation,
          arguments: command.arguments,
        })
      },
      { index, command },
    )
    await expect
      .poll(() =>
        page.evaluate(
          (commandId) =>
            window.__codexVisualHost.acknowledgements.find(
              (entry) => entry.commandId === commandId,
            ),
          command.commandId,
        ),
      )
      .toMatchObject({ ok: true, revision: index + 1 })
  }
  await expect(editor.locator('svg.sparkline-svg')).toHaveCount(2)

  const downloadPromise = page.waitForEvent('download')
  await page.evaluate(() => {
    window.__codexVisualHost.commands.push({
      commandId: 'xlsx-sparkline-save',
      baseRevision: 2,
      operation: 'xlsx.document.save',
      arguments: {},
    })
  })
  const savedPath = await (await downloadPromise).path()
  expect(savedPath).not.toBeNull()
  const savedBytes = await import('node:fs/promises').then((fs) => fs.readFile(savedPath!))
  const zip = await JSZip.loadAsync(savedBytes)
  const sheetXml = await zip.file('xl/worksheets/sheet1.xml')!.async('text')
  expect(sheetXml).toContain('<x14:sparklineGroup displayEmptyCellsAs="gap" type="column">')
  expect(sheetXml).toContain("<xm:f>'Sheet1'!$A$1:$C$1</xm:f><xm:sqref>D1</xm:sqref>")

  await editor.locator('input[type="file"][accept*=".xlsx"]').setInputFiles({
    name: 'registry-sparkline-reopened.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: savedBytes,
  })
  await expect(editor.locator('.workbook-status')).toContainText('registry-sparkline-reopened.xlsx')
  await expect(editor.locator('svg.sparkline-svg')).toHaveCount(2)
})

test('XLSX outline registry state survives typed save, reopen, and explicit expansion', async ({
  page,
}) => {
  await page.goto('/?format=xlsx&width=720&height=900')
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)
  const editor = page.frameLocator('#editor-frame')
  await editor.locator('input[type="file"][accept*=".xlsx"]').setInputFiles({
    name: 'registry-outline.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: await xlsxWithA1('outline'),
  })

  const commands = [
    {
      commandId: 'xlsx-outline-level',
      operation: 'xlsx.outline.set_level',
      arguments: { sheet: 'Sheet1', axis: 'rows', start: 2, count: 2, level: 1 },
    },
    {
      commandId: 'xlsx-outline-collapse',
      operation: 'xlsx.outline.set_detail_visibility',
      arguments: { sheet: 'Sheet1', axis: 'rows', start: 2, count: 2, hidden: true },
    },
  ] as const
  for (const [index, command] of commands.entries()) {
    await page.evaluate(
      ({ index, command }) => {
        window.__codexVisualHost.commands.push({
          commandId: command.commandId,
          baseRevision: index,
          operation: command.operation,
          arguments: command.arguments,
        })
      },
      { index, command },
    )
    await expect
      .poll(() =>
        page.evaluate(
          (commandId) =>
            window.__codexVisualHost.acknowledgements.find(
              (entry) => entry.commandId === commandId,
            ),
          command.commandId,
        ),
      )
      .toMatchObject({ ok: true, revision: index + 1 })
  }

  const firstDownload = page.waitForEvent('download')
  await page.evaluate(() => {
    window.__codexVisualHost.commands.push({
      commandId: 'xlsx-outline-save',
      baseRevision: 2,
      operation: 'xlsx.document.save',
      arguments: {},
    })
  })
  const firstPath = await (await firstDownload).path()
  expect(firstPath).not.toBeNull()
  const collapsedBytes = await import('node:fs/promises').then((fs) => fs.readFile(firstPath!))
  const collapsedZip = await JSZip.loadAsync(collapsedBytes)
  const collapsedXml = await collapsedZip.file('xl/worksheets/sheet1.xml')!.async('text')
  expect(collapsedXml).toMatch(
    /<row\b(?=[^>]*\br="2")(?=[^>]*\bhidden="1")(?=[^>]*\boutlineLevel="1")[^>]*>/,
  )
  expect(collapsedXml).toMatch(
    /<row\b(?=[^>]*\br="3")(?=[^>]*\bhidden="1")(?=[^>]*\boutlineLevel="1")[^>]*>/,
  )
  expect(collapsedXml).toMatch(/<row\b[^>]*\br="4"[^>]*\bcollapsed="1"/)

  await editor.locator('input[type="file"][accept*=".xlsx"]').setInputFiles({
    name: 'registry-outline-reopened.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: collapsedBytes,
  })
  await expect(editor.locator('.workbook-status')).toContainText('registry-outline-reopened.xlsx')
  await page.evaluate(() => {
    window.__codexVisualHost.commands.push({
      commandId: 'xlsx-outline-expand-reopened',
      baseRevision: 3,
      operation: 'xlsx.outline.set_detail_visibility',
      arguments: { sheet: 'Sheet1', axis: 'rows', start: 2, count: 2, hidden: false },
    })
  })
  await expect
    .poll(() =>
      page.evaluate(() =>
        window.__codexVisualHost.acknowledgements.find(
          (entry) => entry.commandId === 'xlsx-outline-expand-reopened',
        ),
      ),
    )
    .toMatchObject({ ok: true, revision: 4 })

  const secondDownload = page.waitForEvent('download')
  await page.evaluate(() => {
    window.__codexVisualHost.commands.push({
      commandId: 'xlsx-outline-expanded-save',
      baseRevision: 4,
      operation: 'xlsx.document.save',
      arguments: {},
    })
  })
  const secondPath = await (await secondDownload).path()
  expect(secondPath).not.toBeNull()
  const expandedBytes = await import('node:fs/promises').then((fs) => fs.readFile(secondPath!))
  const expandedZip = await JSZip.loadAsync(expandedBytes)
  const expandedXml = await expandedZip.file('xl/worksheets/sheet1.xml')!.async('text')
  expect(expandedXml).not.toMatch(/<row\b[^>]*\br="(?:2|3)"[^>]*\bhidden="1"/)
  expect(expandedXml).not.toMatch(/<row\b[^>]*\br="4"[^>]*\bcollapsed="1"/)
})

test('XLSX checkbox validation survives typed save, reopen, and explicit removal', async ({
  page,
}) => {
  await page.goto('/?format=xlsx&width=720&height=900')
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)
  const editor = page.frameLocator('#editor-frame')
  await editor.locator('input[type="file"][accept*=".xlsx"]').setInputFiles({
    name: 'registry-checkbox.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: await xlsxWithA1('checkbox'),
  })

  await page.evaluate(() => {
    window.__codexVisualHost.commands.push({
      commandId: 'xlsx-checkbox-enable',
      baseRevision: 0,
      operation: 'xlsx.range.set_checkbox',
      arguments: { sheet: 'Sheet1', range: 'B2:C3', enabled: true },
    })
  })
  await expect
    .poll(() =>
      page.evaluate(() =>
        window.__codexVisualHost.acknowledgements.find(
          (entry) => entry.commandId === 'xlsx-checkbox-enable',
        ),
      ),
    )
    .toMatchObject({ ok: true, revision: 1 })

  const firstDownload = page.waitForEvent('download')
  await page.evaluate(() => {
    window.__codexVisualHost.commands.push({
      commandId: 'xlsx-checkbox-save',
      baseRevision: 1,
      operation: 'xlsx.document.save',
      arguments: {},
    })
  })
  const firstPath = await (await firstDownload).path()
  expect(firstPath).not.toBeNull()
  const checkedBytes = await import('node:fs/promises').then((fs) => fs.readFile(firstPath!))
  const checkedZip = await JSZip.loadAsync(checkedBytes)
  const checkedXml = await checkedZip.file('xl/worksheets/sheet1.xml')!.async('text')
  expect(checkedXml).toContain('<dataValidations count="1">')
  expect(checkedXml).toMatch(/<dataValidation\b[^>]*\btype="list"[^>]*\bsqref="B2:C3"/)
  expect(checkedXml).toContain('<formula1>"1,0"</formula1>')

  await editor.locator('input[type="file"][accept*=".xlsx"]').setInputFiles({
    name: 'registry-checkbox-reopened.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: checkedBytes,
  })
  await expect(editor.locator('.workbook-status')).toContainText('registry-checkbox-reopened.xlsx')
  await page.evaluate(() => {
    window.__codexVisualHost.commands.push({
      commandId: 'xlsx-checkbox-disable',
      baseRevision: 2,
      operation: 'xlsx.range.set_checkbox',
      arguments: { sheet: 'Sheet1', range: 'B2:C3', enabled: false },
    })
  })
  await expect
    .poll(() =>
      page.evaluate(() =>
        window.__codexVisualHost.acknowledgements.find(
          (entry) => entry.commandId === 'xlsx-checkbox-disable',
        ),
      ),
    )
    .toMatchObject({ ok: true, revision: 3 })

  const secondDownload = page.waitForEvent('download')
  await page.evaluate(() => {
    window.__codexVisualHost.commands.push({
      commandId: 'xlsx-checkbox-removed-save',
      baseRevision: 3,
      operation: 'xlsx.document.save',
      arguments: {},
    })
  })
  const secondPath = await (await secondDownload).path()
  expect(secondPath).not.toBeNull()
  const removedBytes = await import('node:fs/promises').then((fs) => fs.readFile(secondPath!))
  const removedZip = await JSZip.loadAsync(removedBytes)
  const removedXml = await removedZip.file('xl/worksheets/sheet1.xml')!.async('text')
  expect(removedXml).not.toContain('<dataValidations')
})

test('the restored Home ribbon formats the active Univer selection and shares undo', async ({
  page,
}) => {
  await page.goto('/?format=xlsx&width=720&height=900')
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)
  const editor = page.frameLocator('#editor-frame')

  await editor.locator('button[data-tip="加粗"]').click()
  await expect
    .poll(() =>
      page.evaluate(() => {
        const selection = window.__codexVisualHost.lastPollArguments?.selection as
          { activeCell?: { style?: { bold?: boolean } } } | undefined
        return selection?.activeCell?.style?.bold
      }),
    )
    .toBe(true)
  await expect(editor.locator('button[data-tip="加粗"]')).toHaveClass(/is-active/)

  await editor.locator('button.qa-btn').nth(1).click()
  await expect
    .poll(() =>
      page.evaluate(() => {
        const selection = window.__codexVisualHost.lastPollArguments?.selection as
          { activeCell?: { style?: { bold?: boolean } } } | undefined
        return selection?.activeCell?.style?.bold
      }),
    )
    .toBe(false)
})

test('a user Ribbon format is preserved in XLSX styles on save', async ({ page }) => {
  await page.goto('/?format=xlsx&width=720&height=900')
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)
  const editor = page.frameLocator('#editor-frame')
  await editor.locator('input[type="file"][accept*=".xlsx"]').setInputFiles({
    name: 'formatted.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: await xlsxWithA1('Format me'),
  })
  await expect(editor.locator('.workbook-status')).toContainText('formatted.xlsx')

  await editor.locator('button[data-tip="加粗"]').click()
  await expect(editor.locator('button[data-tip="加粗"]')).toHaveClass(/is-active/)

  const downloadPromise = page.waitForEvent('download')
  await editor.locator('button.qa-btn').first().click()
  const savedPath = await (await downloadPromise).path()
  expect(savedPath).not.toBeNull()
  const zip = await JSZip.loadAsync(
    await import('node:fs/promises').then((fs) => fs.readFile(savedPath!)),
  )
  const sheetXml = await zip.file('xl/worksheets/sheet1.xml')!.async('text')
  const stylesXml = await zip.file('xl/styles.xml')!.async('text')
  expect(sheetXml).toMatch(/<c\b[^>]*\br="A1"[^>]*\bs="[1-9][0-9]*"/)
  expect(stylesXml).toContain('<b/>')
})

test('undo removes a Ribbon format from the XLSX that is subsequently saved', async ({ page }) => {
  await page.goto('/?format=xlsx&width=720&height=900')
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)
  const editor = page.frameLocator('#editor-frame')
  await editor.locator('input[type="file"][accept*=".xlsx"]').setInputFiles({
    name: 'undo-format.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: await xlsxWithA1('Undo format'),
  })
  await expect(editor.locator('.workbook-status')).toContainText('undo-format.xlsx')

  await editor.locator('button[data-tip="加粗"]').click()
  await expect(editor.locator('button[data-tip="加粗"]')).toHaveClass(/is-active/)
  await editor.locator('button.qa-btn').nth(1).click()
  await expect(editor.locator('button[data-tip="加粗"]')).not.toHaveClass(/is-active/)

  const downloadPromise = page.waitForEvent('download')
  await editor.locator('button.qa-btn').first().click()
  const savedPath = await (await downloadPromise).path()
  expect(savedPath).not.toBeNull()
  const zip = await JSZip.loadAsync(
    await import('node:fs/promises').then((fs) => fs.readFile(savedPath!)),
  )
  const sheetXml = await zip.file('xl/worksheets/sheet1.xml')!.async('text')
  const stylesXml = await zip.file('xl/styles.xml')!.async('text')
  const styleIndex = Number(/<c\b[^>]*\br="A1"[^>]*\bs="([0-9]+)"/.exec(sheetXml)?.[1] ?? 0)
  const cellXfs = /<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/.exec(stylesXml)?.[1] ?? ''
  const xf = [...cellXfs.matchAll(/<xf\b[^>]*(?:\/>|>[\s\S]*?<\/xf>)/g)][styleIndex]?.[0]
  const fontId = Number(/\bfontId="([0-9]+)"/.exec(xf ?? '')?.[1] ?? 0)
  const fonts = /<fonts\b[^>]*>([\s\S]*?)<\/fonts>/.exec(stylesXml)?.[1] ?? ''
  const font = [...fonts.matchAll(/<font\b[^>]*(?:\/>|>[\s\S]*?<\/font>)/g)][fontId]?.[0]
  expect(font).not.toMatch(/<b\b/)
})

test('the community Ribbon dispatcher applies italic through the active Univer range', async ({
  page,
}) => {
  await page.goto('/?format=xlsx&width=720&height=900')
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)
  const editor = page.frameLocator('#editor-frame')

  await editor.locator('button[data-tip="倾斜"]').click()
  await expect(editor.locator('button[data-tip="倾斜"]')).toHaveClass(/is-active/)
})

test('the MCP text-style registry operation shares the active Univer state and user undo stack', async ({
  page,
}) => {
  await page.goto('/?format=xlsx&width=720&height=900')
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)
  const editor = page.frameLocator('#editor-frame')

  await page.evaluate(() => {
    window.__codexVisualHost.commands.push({
      commandId: 'xlsx-italic',
      baseRevision: 0,
      operation: 'xlsx.range.set_text_style',
      arguments: {
        sheet: 'Sheet1',
        range: 'A1',
        style: { italic: true },
        fields: ['italic'],
      },
    })
  })

  await expect
    .poll(() =>
      page.evaluate(() =>
        window.__codexVisualHost.acknowledgements.find(
          (entry) => entry.commandId === 'xlsx-italic',
        ),
      ),
    )
    .toMatchObject({ ok: true, revision: 1 })
  await expect(editor.locator('button[data-tip="倾斜"]')).toHaveClass(/is-active/)

  await editor.locator('button.qa-btn').nth(1).click()
  await expect(editor.locator('button[data-tip="倾斜"]')).not.toHaveClass(/is-active/)
})

test('registry-owned Ribbon commands are rejected in favor of explicit range operations', async ({
  page,
}) => {
  await page.goto('/?format=xlsx&width=720&height=900')
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)

  const cases = [
    ['italic', 'xlsx.range.set_text_style'],
    ['align:left', 'xlsx.range.set_alignment'],
    ['font-family:Aptos', 'xlsx.range.set_font'],
    ['fill:#DDEBF7', 'xlsx.range.set_fill'],
    ['border:all:#4472C4', 'xlsx.range.set_border'],
    ['cell-style:good', 'xlsx.range.apply_cell_style'],
    ['format-painter', 'explicit xlsx.range style operations'],
    ['format:0.00%', 'xlsx.range.set_number_format'],
    ['decimal-inc', 'xlsx.range.set_number_format'],
    ['merge:center', 'xlsx.range.merge'],
    ['clear-all', 'xlsx.range.clear'],
    ['fill-down', 'xlsx.range.fill'],
    ['sort:asc', 'xlsx.range.sort'],
    ['sort-custom:1:0a,1d', 'xlsx.range.sort_custom'],
    ['remove-duplicates:1', 'xlsx.range.remove_duplicates'],
    ['autofn:SUM', 'xlsx.formula.insert_aggregate'],
    ['flash-fill', 'xlsx.range.flash_fill'],
    ['text-to-columns:4', 'xlsx.range.text_to_columns'],
    ['row-height:24.75', 'xlsx.row.set_height'],
    ['col-width:12.5', 'xlsx.column.set_width'],
    ['freeze-here', 'xlsx.sheet.set_freeze'],
    ['freeze-top-row', 'xlsx.sheet.set_freeze'],
    ['freeze-first-col', 'xlsx.sheet.set_freeze'],
    ['unfreeze', 'xlsx.sheet.set_freeze'],
    ['toggle-gridlines', 'xlsx.sheet.set_gridlines'],
    ['toggle-show-formulas', 'xlsx.sheet.set_formula_view'],
    ['page-layout:margins:wide', 'xlsx.sheet.set_page_margins'],
    ['page-layout:orientation:landscape', 'xlsx.sheet.set_page_orientation'],
    ['page-layout:paper:9', 'xlsx.sheet.set_paper_size'],
    ['page-layout:fit-width:2', 'xlsx.sheet.set_fit_to_pages'],
    ['page-layout:fit-height:3', 'xlsx.sheet.set_fit_to_pages'],
    ['page-layout:print-gridlines:1', 'xlsx.sheet.set_print_gridlines'],
    ['page-layout:print-headings:1', 'xlsx.sheet.set_print_headings'],
    ['page-layout:print-area:set', 'xlsx.sheet.set_print_area'],
    ['page-layout:print-area:clear', 'xlsx.sheet.set_print_area'],
    ['page-layout:print-titles:first-row', 'xlsx.sheet.set_print_titles'],
    ['page-layout:print-titles:set', 'xlsx.sheet.set_print_titles'],
    ['page-layout:print-titles:clear', 'xlsx.sheet.set_print_titles'],
    ['page-layout:scale:80', 'xlsx.sheet.set_print_scale'],
    ['link-set:https%3A%2F%2Fexample.com', 'xlsx.hyperlink.set'],
    ['link-remove', 'xlsx.hyperlink.remove'],
    ['format-as-table:TableStyleMedium2', 'xlsx.table.add'],
    ['sparkline:column', 'xlsx.sparkline.add'],
    ['outline-group:rows', 'xlsx.outline.set_level'],
    ['outline-hide-detail:cols', 'xlsx.outline.set_detail_visibility'],
    ['insert-checkbox', 'xlsx.range.set_checkbox'],
    ['insert-symbol', 'xlsx.cell.set_value'],
    ['import-csv', 'xlsx.range.set_values'],
    ['cellprot:locked-on:hidden-off', 'xlsx.range.set_protection'],
    ['sheet-protect', 'xlsx.sheet.set_protection'],
    ['filter-toggle', 'xlsx.range.set_filter'],
    ['filter-clear', 'xlsx.range.clear_filter_criteria'],
    ['filter-advanced', 'xlsx.range.set_custom_filter'],
    ['insert-sheet', 'xlsx.sheet.add'],
    ['insert-row-here', 'xlsx.row.insert'],
    ['delete-row-here', 'xlsx.row.delete'],
    ['insert-col-here', 'xlsx.column.insert'],
    ['delete-col-here', 'xlsx.column.delete'],
    ['undo', 'xlsx.history.undo'],
    ['redo', 'xlsx.history.redo'],
    ['paste-special:value', 'xlsx.range.copy_values'],
    ['paste-special:formula', 'xlsx.range.copy_formulas'],
    ['paste-special:format', 'xlsx.range.copy_formats'],
    ['paste-special:col-width', 'xlsx.column.copy_widths'],
    ['paste-special:besides-border', 'xlsx.range.copy_without_borders'],
    ['copy', 'explicit source/destination XLSX operations'],
    ['cut', 'explicit source/destination XLSX operations'],
    ['paste', 'explicit source/destination XLSX operations'],
  ] as const
  await page.evaluate((commands) => {
    for (const [index, [command]] of commands.entries()) {
      window.__codexVisualHost.commands.push({
        commandId: `xlsx-legacy-style-${index}`,
        baseRevision: 0,
        operation: 'ribbon_command',
        arguments: { command },
      })
    }
  }, cases)

  for (const [index] of cases.entries()) {
    await expect
      .poll(() =>
        page.evaluate(
          (commandId) =>
            window.__codexVisualHost.acknowledgements.find(
              (entry) => entry.commandId === commandId,
            ),
          `xlsx-legacy-style-${index}`,
        ),
      )
      .toMatchObject({
        ok: false,
        error: 'unsupported_operation',
        message: 'Unsupported XLSX operation: ribbon_command',
      })
  }
})

test('the XLSX range-style registry operations share state, undo, and saved output', async ({
  page,
}) => {
  await page.goto('/?format=xlsx&width=720&height=900')
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)
  const editor = page.frameLocator('#editor-frame')
  await editor.locator('input[type="file"][accept*=".xlsx"]').setInputFiles({
    name: 'registry-styles.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: await xlsxWithA1('Style me'),
  })

  const commands = [
    {
      operation: 'xlsx.range.set_alignment',
      arguments: {
        sheet: 'Sheet1',
        range: 'A1',
        alignment: {
          horizontal: 'center',
          wrap: true,
          indent: 2,
          rotation: { mode: 'angle', degrees: 45 },
        },
        fields: ['horizontal', 'wrap', 'indent', 'rotation'],
      },
    },
    {
      operation: 'xlsx.range.set_font',
      arguments: {
        sheet: 'Sheet1',
        range: 'A1',
        font: { family: 'Aptos', size: 14, color: '#44546A' },
        fields: ['family', 'size', 'color'],
      },
    },
    {
      operation: 'xlsx.range.set_fill',
      arguments: { sheet: 'Sheet1', range: 'A1', color: '#DDEBF7' },
    },
    {
      operation: 'xlsx.range.set_border',
      arguments: {
        sheet: 'Sheet1',
        range: 'A1',
        border: { preset: 'outer', lineStyle: 'medium', color: '#4472C4' },
      },
    },
    {
      operation: 'xlsx.range.apply_cell_style',
      arguments: { sheet: 'Sheet1', range: 'A1', preset: 'input' },
    },
  ] as const

  for (const [index, command] of commands.entries()) {
    await page.evaluate(
      ({ index, command }) => {
        window.__codexVisualHost.commands.push({
          commandId: `xlsx-registry-style-${index}`,
          baseRevision: index,
          operation: command.operation,
          arguments: command.arguments,
        })
      },
      { index, command },
    )
    await expect
      .poll(() =>
        page.evaluate(
          (commandId) =>
            window.__codexVisualHost.acknowledgements.find(
              (entry) => entry.commandId === commandId,
            ),
          `xlsx-registry-style-${index}`,
        ),
      )
      .toMatchObject({ ok: true, revision: index + 1 })
  }

  await expect
    .poll(() => page.evaluate(() => window.__codexVisualHost.lastPollArguments))
    .toMatchObject({
      selection: {
        activeCell: {
          style: {
            fontFamily: 'Aptos',
            fontSize: 14,
            horizontalAlignment: 'center',
            wrap: true,
            indent: 2,
            textRotation: 45,
            fontColor: '#3F3F76',
            fillColor: '#FFCC99',
          },
        },
      },
    })

  await editor.locator('button.qa-btn').nth(1).click()
  await expect
    .poll(() => page.evaluate(() => window.__codexVisualHost.lastPollArguments))
    .toMatchObject({ selection: { activeCell: { style: { fillColor: '#DDEBF7' } } } })
  await editor.locator('button.qa-btn').nth(2).click()
  await expect
    .poll(() => page.evaluate(() => window.__codexVisualHost.lastPollArguments))
    .toMatchObject({ selection: { activeCell: { style: { fillColor: '#FFCC99' } } } })

  const downloadPromise = page.waitForEvent('download')
  await editor.locator('button.qa-btn').first().click()
  const savedPath = await (await downloadPromise).path()
  expect(savedPath).not.toBeNull()
  const zip = await JSZip.loadAsync(
    await import('node:fs/promises').then((fs) => fs.readFile(savedPath!)),
  )
  const stylesXml = await zip.file('xl/styles.xml')!.async('text')
  expect(stylesXml).toContain('Aptos')
  expect(stylesXml).toContain('FFCC99')
  expect(stylesXml).toContain('3F3F76')
  expect(stylesXml).toContain('7F7F7F')
  expect(stylesXml).toMatch(/<alignment\b[^>]*horizontal="center"[^>]*wrapText="1"/)
})

test('the XLSX number-format, merge, clear, and fill registry operations share state, undo, and saved output', async ({
  page,
}) => {
  await page.goto('/?format=xlsx&width=720&height=900')
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)
  const editor = page.frameLocator('#editor-frame')
  await editor.locator('input[type="file"][accept*=".xlsx"]').setInputFiles({
    name: 'registry-range-actions.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: await xlsxWithA1('Before'),
  })

  const commands = [
    {
      operation: 'xlsx.range.set_values',
      arguments: {
        sheet: 'Sheet1',
        range: 'A1:D2',
        values: [
          [12, 'merge', 'source', 'clear'],
          [34, 'keep', 'destination', 'keep'],
        ],
      },
    },
    {
      operation: 'xlsx.range.set_number_format',
      arguments: { sheet: 'Sheet1', range: 'A1:A2', pattern: '0.0000' },
    },
    {
      operation: 'xlsx.range.fill',
      arguments: { sheet: 'Sheet1', range: 'C1:C2', direction: 'down' },
    },
    {
      operation: 'xlsx.range.merge',
      arguments: { sheet: 'Sheet1', range: 'A1:B1', mode: 'cells' },
    },
    {
      operation: 'xlsx.range.clear',
      arguments: { sheet: 'Sheet1', range: 'D1', scope: 'contents' },
    },
  ] as const

  for (const [index, command] of commands.entries()) {
    await page.evaluate(
      ({ index, command }) => {
        window.__codexVisualHost.commands.push({
          commandId: `xlsx-registry-range-${index}`,
          baseRevision: index,
          operation: command.operation,
          arguments: command.arguments,
        })
      },
      { index, command },
    )
    await expect
      .poll(() =>
        page.evaluate(
          (commandId) =>
            window.__codexVisualHost.acknowledgements.find(
              (entry) => entry.commandId === commandId,
            ),
          `xlsx-registry-range-${index}`,
        ),
      )
      .toMatchObject({ ok: true, revision: index + 1 })

    if (command.operation === 'xlsx.range.set_number_format') {
      await expect
        .poll(() =>
          page.evaluate(() => {
            const selection = window.__codexVisualHost.lastPollArguments?.selection as
              { activeCell?: { style?: { numberFormat?: string } } } | undefined
            return selection?.activeCell?.style?.numberFormat
          }),
        )
        .toBe('0.0000')
    }
  }

  await expect
    .poll(() =>
      page.evaluate(() => {
        const selection = window.__codexVisualHost.lastPollArguments?.selection as
          { activeCell?: { value?: unknown } } | undefined
        return selection?.activeCell?.value
      }),
    )
    .toBeNull()
  await editor.locator('button.qa-btn').nth(1).click()
  await expect
    .poll(() =>
      page.evaluate(() => {
        const selection = window.__codexVisualHost.lastPollArguments?.selection as
          { activeCell?: { value?: unknown } } | undefined
        return selection?.activeCell?.value
      }),
    )
    .toBe('clear')
  await editor.locator('button.qa-btn').nth(2).click()
  await expect
    .poll(() =>
      page.evaluate(() => {
        const selection = window.__codexVisualHost.lastPollArguments?.selection as
          { activeCell?: { value?: unknown } } | undefined
        return selection?.activeCell?.value
      }),
    )
    .toBeNull()

  const downloadPromise = page.waitForEvent('download')
  await editor.locator('button.qa-btn').first().click()
  const savedPath = await (await downloadPromise).path()
  expect(savedPath).not.toBeNull()
  const zip = await JSZip.loadAsync(
    await import('node:fs/promises').then((fs) => fs.readFile(savedPath!)),
  )
  const sheetXml = await zip.file('xl/worksheets/sheet1.xml')!.async('text')
  const stylesXml = await zip.file('xl/styles.xml')!.async('text')
  const c2Xml = sheetXml.match(/<c r="C2"[^>]*>[\s\S]*?<\/c>/)?.[0] ?? ''
  expect(c2Xml).toContain('source')
  expect(sheetXml).not.toContain('destination')
  expect(sheetXml).not.toContain('clear')
  expect(sheetXml).toMatch(/<mergeCell\b[^>]*ref="A1:B1"/)
  expect(stylesXml).toContain('formatCode="0.0000"')
})

test('the XLSX sort and remove-duplicates registry operations share state, undo, and saved output', async ({
  page,
}) => {
  await page.goto('/?format=xlsx&width=720&height=900')
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)
  const editor = page.frameLocator('#editor-frame')
  await editor.locator('input[type="file"][accept*=".xlsx"]').setInputFiles({
    name: 'registry-sort-dedupe.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: await xlsxWithA1('Before'),
  })

  const commands = [
    {
      operation: 'xlsx.range.set_values',
      arguments: {
        sheet: 'Sheet1',
        range: 'A1:C5',
        values: [
          ['Name', 'Score', 'Team'],
          ['Bravo', 2, 'B'],
          ['Alpha', 3, 'A'],
          ['Charlie', 1, 'C'],
          ['Alpha', 3, 'A'],
        ],
      },
    },
    {
      operation: 'xlsx.range.sort',
      arguments: { sheet: 'Sheet1', range: 'A2:C5', direction: 'asc' },
    },
    {
      operation: 'xlsx.range.sort_custom',
      arguments: {
        sheet: 'Sheet1',
        range: 'A1:C5',
        keys: [
          { column: 'B', direction: 'asc' },
          { column: 'A', direction: 'asc' },
        ],
        hasHeader: true,
      },
    },
    {
      operation: 'xlsx.range.remove_duplicates',
      arguments: { sheet: 'Sheet1', range: 'A1:C5', hasHeader: true },
    },
  ] as const

  for (const [index, command] of commands.entries()) {
    await page.evaluate(
      ({ index, command }) => {
        window.__codexVisualHost.commands.push({
          commandId: `xlsx-registry-sort-${index}`,
          baseRevision: index,
          operation: command.operation,
          arguments: command.arguments,
        })
      },
      { index, command },
    )
    await expect
      .poll(() =>
        page.evaluate(
          (commandId) =>
            window.__codexVisualHost.acknowledgements.find(
              (entry) => entry.commandId === commandId,
            ),
          `xlsx-registry-sort-${index}`,
        ),
      )
      .toMatchObject({
        ok: true,
        revision: index + 1,
        ...(command.operation === 'xlsx.range.remove_duplicates' ? { output: { removed: 1 } } : {}),
      })
  }

  const canvas = editor.locator('canvas[id^="univer-sheet-main-canvas_"]')
  await canvas.click({ position: { x: 70, y: 135 } })
  await expect
    .poll(() =>
      page.evaluate(() => {
        const selection = window.__codexVisualHost.lastPollArguments?.selection as
          { activeCell?: { address?: string; value?: unknown } } | undefined
        return selection?.activeCell
      }),
    )
    .toMatchObject({ address: 'A5', value: null })

  await editor.locator('button.qa-btn').nth(1).click()
  await canvas.click({ position: { x: 70, y: 135 } })
  await expect
    .poll(() =>
      page.evaluate(() => {
        const selection = window.__codexVisualHost.lastPollArguments?.selection as
          { activeCell?: { value?: unknown } } | undefined
        return selection?.activeCell?.value
      }),
    )
    .toBe('Alpha')
  await editor.locator('button.qa-btn').nth(2).click()
  await canvas.click({ position: { x: 70, y: 135 } })
  await expect
    .poll(() =>
      page.evaluate(() => {
        const selection = window.__codexVisualHost.lastPollArguments?.selection as
          { activeCell?: { value?: unknown } } | undefined
        return selection?.activeCell?.value
      }),
    )
    .toBeNull()

  const downloadPromise = page.waitForEvent('download')
  await editor.locator('button.qa-btn').first().click()
  const savedPath = await (await downloadPromise).path()
  expect(savedPath).not.toBeNull()
  const zip = await JSZip.loadAsync(
    await import('node:fs/promises').then((fs) => fs.readFile(savedPath!)),
  )
  const sheetXml = await zip.file('xl/worksheets/sheet1.xml')!.async('text')
  expect(sheetXml).toMatch(/<c\b[^>]*\br="A2"[^>]*>[\s\S]*?Charlie[\s\S]*?<\/c>/)
  expect(sheetXml).toMatch(/<c\b[^>]*\br="A3"[^>]*>[\s\S]*?Bravo[\s\S]*?<\/c>/)
  expect(sheetXml).toMatch(/<c\b[^>]*\br="A4"[^>]*>[\s\S]*?Alpha[\s\S]*?<\/c>/)
  expect(sheetXml).toMatch(/<c\b[^>]*\br="A5"[^>]*\/>/)
  expect(sheetXml.match(/Alpha/g)).toHaveLength(1)
})

test('the XLSX filter registry operation and Ribbon share undo and saved output', async ({
  page,
}) => {
  await page.goto('/?format=xlsx&width=720&height=900')
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)
  const editor = page.frameLocator('#editor-frame')
  await editor.locator('input[type="file"][accept*=".xlsx"]').setInputFiles({
    name: 'registry-filter.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: await xlsxWithA1('Before'),
  })

  const commands = [
    {
      operation: 'xlsx.range.set_values',
      arguments: {
        sheet: 'Sheet1',
        range: 'A1:C4',
        values: [
          ['Name', 'Score', 'Team'],
          ['Alpha', 3, 'A'],
          ['Bravo', 2, 'B'],
          ['Charlie', 1, 'C'],
        ],
      },
    },
    {
      operation: 'xlsx.range.set_filter',
      arguments: { sheet: 'Sheet1', range: 'A1:C4', enabled: true },
    },
  ] as const
  for (const [index, command] of commands.entries()) {
    await page.evaluate(
      ({ index, command }) => {
        window.__codexVisualHost.commands.push({
          commandId: `xlsx-registry-filter-${index}`,
          baseRevision: index,
          operation: command.operation,
          arguments: command.arguments,
        })
      },
      { index, command },
    )
    await expect
      .poll(() =>
        page.evaluate(
          (commandId) =>
            window.__codexVisualHost.acknowledgements.find(
              (entry) => entry.commandId === commandId,
            ),
          `xlsx-registry-filter-${index}`,
        ),
      )
      .toMatchObject({ ok: true, revision: index + 1 })
  }

  await editor.locator('button.qa-btn').nth(1).click()
  await expect(editor.locator('button.qa-btn').nth(2)).toBeEnabled()
  const undoDownloadPromise = page.waitForEvent('download')
  await editor.locator('button.qa-btn').first().click()
  const undoPath = await (await undoDownloadPromise).path()
  expect(undoPath).not.toBeNull()
  const undoZip = await JSZip.loadAsync(
    await import('node:fs/promises').then((fs) => fs.readFile(undoPath!)),
  )
  const undoSheetXml = await undoZip.file('xl/worksheets/sheet1.xml')!.async('text')
  expect(undoSheetXml).not.toContain('<autoFilter')

  await editor.getByRole('button', { name: '数据', exact: true }).click()
  await editor.locator('.ribbon-tool.large').filter({ hasText: '筛选' }).click()
  await expect(editor.locator('button.qa-btn').nth(2)).toBeDisabled()
  await editor.locator('button.qa-btn').nth(1).click()
  await expect(editor.locator('button.qa-btn').nth(2)).toBeEnabled()
  await editor.locator('button.qa-btn').nth(2).click()

  const redoDownloadPromise = page.waitForEvent('download')
  await editor.locator('button.qa-btn').first().click()
  const redoPath = await (await redoDownloadPromise).path()
  expect(redoPath).not.toBeNull()
  const redoBytes = await import('node:fs/promises').then((fs) => fs.readFile(redoPath!))
  const redoZip = await JSZip.loadAsync(redoBytes)
  const redoSheetXml = await redoZip.file('xl/worksheets/sheet1.xml')!.async('text')
  expect(redoSheetXml).toMatch(/<autoFilter\b[^>]*\bref="A1:C4"/)

  await page.reload()
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)
  const reopened = page.frameLocator('#editor-frame')
  await reopened.locator('input[type="file"][accept*=".xlsx"]').setInputFiles({
    name: 'registry-filter-reopened.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: redoBytes,
  })
  await page.evaluate(() => {
    window.__codexVisualHost.commands.push({
      commandId: 'xlsx-registry-filter-reopened-touch',
      baseRevision: 0,
      operation: 'xlsx.cell.set_value',
      arguments: { sheet: 'Sheet1', address: 'A1', value: 'Name' },
    })
  })
  await expect
    .poll(() =>
      page.evaluate(() =>
        window.__codexVisualHost.acknowledgements.find(
          (entry) => entry.commandId === 'xlsx-registry-filter-reopened-touch',
        ),
      ),
    )
    .toMatchObject({ ok: true, revision: 1 })
  const reopenedDownloadPromise = page.waitForEvent('download')
  await reopened.locator('button.qa-btn').first().click()
  const reopenedPath = await (await reopenedDownloadPromise).path()
  expect(reopenedPath).not.toBeNull()
  const reopenedZip = await JSZip.loadAsync(
    await import('node:fs/promises').then((fs) => fs.readFile(reopenedPath!)),
  )
  const reopenedSheetXml = await reopenedZip.file('xl/worksheets/sheet1.xml')!.async('text')
  expect(reopenedSheetXml).toMatch(/<autoFilter\b[^>]*\bref="A1:C4"/)
})

test('the XLSX filter criteria operations and retained Data controls share undo and saved output', async ({
  page,
}) => {
  await page.goto('/?format=xlsx&width=720&height=900')
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)
  const editor = page.frameLocator('#editor-frame')
  await editor.locator('input[type="file"][accept*=".xlsx"]').setInputFiles({
    name: 'registry-filter-criteria.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: await xlsxWithA1('Before'),
  })

  const commands = [
    {
      operation: 'xlsx.range.set_values',
      arguments: {
        sheet: 'Sheet1',
        range: 'A1:C5',
        values: [
          ['Name', 'Score', 'Team'],
          ['Alpha', 3, 'A'],
          ['Bravo', 2, 'B'],
          ['Charlie', 1, 'C'],
          ['Delta', 4, 'A'],
        ],
      },
    },
    {
      operation: 'xlsx.range.set_filter',
      arguments: { sheet: 'Sheet1', range: 'A1:C5', enabled: true },
    },
    {
      operation: 'xlsx.range.set_filter_values',
      arguments: {
        sheet: 'Sheet1',
        range: 'A1:C5',
        column: 'C',
        values: ['A', 'C'],
        includeBlank: false,
      },
    },
    {
      operation: 'xlsx.range.set_custom_filter',
      arguments: {
        sheet: 'Sheet1',
        range: 'A1:C5',
        column: 'B',
        conjunction: 'and',
        conditions: [
          { operator: 'greaterThanOrEqual', value: '2' },
          { operator: 'lessThanOrEqual', value: '3' },
        ],
      },
    },
  ] as const
  for (const [index, command] of commands.entries()) {
    await page.evaluate(
      ({ index, command }) => {
        window.__codexVisualHost.commands.push({
          commandId: `xlsx-registry-filter-criteria-${index}`,
          baseRevision: index,
          operation: command.operation,
          arguments: command.arguments,
        })
      },
      { index, command },
    )
    await expect
      .poll(() =>
        page.evaluate(
          (commandId) =>
            window.__codexVisualHost.acknowledgements.find(
              (entry) => entry.commandId === commandId,
            ),
          `xlsx-registry-filter-criteria-${index}`,
        ),
      )
      .toMatchObject({ ok: true, revision: index + 1 })
  }

  await editor.locator('button.qa-btn').nth(1).click()
  await editor.locator('button.qa-btn').nth(2).click()
  await editor.getByRole('button', { name: '数据', exact: true }).click()
  await editor.locator('button[data-tip="清除筛选条件"]').click()
  await editor.locator('button.qa-btn').nth(1).click()

  await editor.getByRole('button', { name: '高级', exact: true }).click()
  const dialog = editor.getByRole('dialog', { name: '高级筛选' })
  await expect(dialog).toBeVisible()
  await dialog.locator('select').nth(0).selectOption('1')
  await dialog.locator('select').nth(1).selectOption('greaterThanOrEqual')
  await dialog.locator('input').nth(0).fill('1')
  await dialog.locator('select').nth(2).selectOption('lessThanOrEqual')
  await dialog.locator('input').nth(3).fill('4')
  await dialog.getByRole('button', { name: '确定', exact: true }).click()
  await expect(dialog).toBeHidden()

  const downloadPromise = page.waitForEvent('download')
  await editor.locator('button.qa-btn').first().click()
  const savedPath = await (await downloadPromise).path()
  expect(savedPath).not.toBeNull()
  const savedBytes = await import('node:fs/promises').then((fs) => fs.readFile(savedPath!))
  const zip = await JSZip.loadAsync(savedBytes)
  const sheetXml = await zip.file('xl/worksheets/sheet1.xml')!.async('text')
  expect(sheetXml).toContain('<autoFilter ref="A1:C5">')
  expect(sheetXml).toContain(
    '<filterColumn colId="1"><customFilters and="1">' +
      '<customFilter operator="greaterThanOrEqual" val="1"/>' +
      '<customFilter operator="lessThanOrEqual" val="4"/>' +
      '</customFilters></filterColumn>',
  )
  expect(sheetXml).toContain(
    '<filterColumn colId="2"><filters><filter val="A"/><filter val="C"/></filters></filterColumn>',
  )
  expect(sheetXml).toMatch(/<row\b[^>]*\br="3"[^>]*\bhidden="1"/)

  await page.reload()
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)
  const reopened = page.frameLocator('#editor-frame')
  await reopened.locator('input[type="file"][accept*=".xlsx"]').setInputFiles({
    name: 'registry-filter-criteria-reopened.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: savedBytes,
  })
  await page.evaluate(() => {
    window.__codexVisualHost.commands.push({
      commandId: 'xlsx-registry-filter-criteria-reopened-touch',
      baseRevision: 0,
      operation: 'xlsx.cell.set_value',
      arguments: { sheet: 'Sheet1', address: 'A1', value: 'Name' },
    })
  })
  await expect
    .poll(() =>
      page.evaluate(() =>
        window.__codexVisualHost.acknowledgements.find(
          (entry) => entry.commandId === 'xlsx-registry-filter-criteria-reopened-touch',
        ),
      ),
    )
    .toMatchObject({ ok: true, revision: 1 })
  const reopenedDownloadPromise = page.waitForEvent('download')
  await reopened.locator('button.qa-btn').first().click()
  const reopenedPath = await (await reopenedDownloadPromise).path()
  expect(reopenedPath).not.toBeNull()
  const reopenedZip = await JSZip.loadAsync(
    await import('node:fs/promises').then((fs) => fs.readFile(reopenedPath!)),
  )
  const reopenedSheetXml = await reopenedZip.file('xl/worksheets/sheet1.xml')!.async('text')
  expect(reopenedSheetXml).toContain('<filterColumn colId="1"><customFilters and="1">')
  expect(reopenedSheetXml).toContain(
    '<filterColumn colId="2"><filters><filter val="A"/><filter val="C"/></filters></filterColumn>',
  )
})

test('the XLSX row-height registry operation shares user undo and saved output', async ({
  page,
}) => {
  await page.goto('/?format=xlsx&width=720&height=900')
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)
  const editor = page.frameLocator('#editor-frame')
  await editor.locator('input[type="file"][accept*=".xlsx"]').setInputFiles({
    name: 'registry-row-height.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: await xlsxWithA1('Sized row'),
  })

  await page.evaluate(() => {
    window.__codexVisualHost.commands.push({
      commandId: 'xlsx-row-height',
      baseRevision: 0,
      operation: 'xlsx.row.set_height',
      arguments: { sheet: 'Sheet1', row: 1, count: 2, heightPoints: 24.75 },
    })
  })
  await expect
    .poll(() =>
      page.evaluate(() =>
        window.__codexVisualHost.acknowledgements.find(
          (entry) => entry.commandId === 'xlsx-row-height',
        ),
      ),
    )
    .toMatchObject({
      ok: true,
      revision: 1,
      output: { sheet: 'Sheet1', row: 1, count: 2, heightPoints: 24.75 },
    })

  await editor.locator('button.qa-btn').nth(1).click()
  await expect(editor.locator('button.qa-btn').nth(2)).toBeEnabled()
  await editor.locator('button.qa-btn').nth(2).click()
  await expect(editor.locator('button.qa-btn').nth(1)).toBeEnabled()
  const redoDownloadPromise = page.waitForEvent('download')
  await editor.locator('button.qa-btn').first().click()
  const redoPath = await (await redoDownloadPromise).path()
  expect(redoPath).not.toBeNull()
  const redoBytes = await import('node:fs/promises').then((fs) => fs.readFile(redoPath!))
  const redoZip = await JSZip.loadAsync(redoBytes)
  const redoSheetXml = await redoZip.file('xl/worksheets/sheet1.xml')!.async('text')
  expect(redoSheetXml).toMatch(/<row\b(?=[^>]*\br="1")(?=[^>]*\bht="24\.75")/)
  expect(redoSheetXml).toMatch(/<row\b(?=[^>]*\br="2")(?=[^>]*\bht="24\.75")/)

  await page.reload()
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)
  const reopened = page.frameLocator('#editor-frame')
  await reopened.locator('input[type="file"][accept*=".xlsx"]').setInputFiles({
    name: 'registry-row-height-reopened.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: redoBytes,
  })
  await expect(reopened.locator('.workbook-status')).toContainText(
    'registry-row-height-reopened.xlsx',
  )
  await page.evaluate(() => {
    window.__codexVisualHost.commands.push({
      commandId: 'xlsx-row-height-reopened-touch',
      baseRevision: 0,
      operation: 'xlsx.cell.set_value',
      arguments: { sheet: 'Sheet1', address: 'A1', value: 'Sized row' },
    })
  })
  await expect
    .poll(() =>
      page.evaluate(() =>
        window.__codexVisualHost.acknowledgements.find(
          (entry) => entry.commandId === 'xlsx-row-height-reopened-touch',
        ),
      ),
    )
    .toMatchObject({ ok: true, revision: 1 })
  const reopenedDownloadPromise = page.waitForEvent('download')
  await reopened.locator('button.qa-btn').first().click()
  const reopenedPath = await (await reopenedDownloadPromise).path()
  expect(reopenedPath).not.toBeNull()
  const reopenedZip = await JSZip.loadAsync(
    await import('node:fs/promises').then((fs) => fs.readFile(reopenedPath!)),
  )
  const reopenedSheetXml = await reopenedZip.file('xl/worksheets/sheet1.xml')!.async('text')
  expect(reopenedSheetXml).toContain('ht="24.75"')
})

test('the XLSX column-width registry operation shares user undo and saved output', async ({
  page,
}) => {
  await page.goto('/?format=xlsx&width=720&height=900')
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)
  const editor = page.frameLocator('#editor-frame')
  await editor.locator('input[type="file"][accept*=".xlsx"]').setInputFiles({
    name: 'registry-column-width.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: await xlsxWithA1('Sized column'),
  })

  await page.evaluate(() => {
    window.__codexVisualHost.commands.push({
      commandId: 'xlsx-column-width',
      baseRevision: 0,
      operation: 'xlsx.column.set_width',
      arguments: { sheet: 'Sheet1', column: 'C', count: 2, widthCharacters: 12.5 },
    })
  })
  await expect
    .poll(() =>
      page.evaluate(() =>
        window.__codexVisualHost.acknowledgements.find(
          (entry) => entry.commandId === 'xlsx-column-width',
        ),
      ),
    )
    .toMatchObject({
      ok: true,
      revision: 1,
      output: { sheet: 'Sheet1', column: 'C', count: 2, widthCharacters: 12.4296875 },
    })

  await editor.locator('button.qa-btn').nth(1).click()
  await expect(editor.locator('button.qa-btn').nth(2)).toBeEnabled()
  await editor.locator('button.qa-btn').nth(2).click()
  await expect(editor.locator('button.qa-btn').nth(1)).toBeEnabled()
  const redoDownloadPromise = page.waitForEvent('download')
  await editor.locator('button.qa-btn').first().click()
  const redoPath = await (await redoDownloadPromise).path()
  expect(redoPath).not.toBeNull()
  const redoBytes = await import('node:fs/promises').then((fs) => fs.readFile(redoPath!))
  const redoZip = await JSZip.loadAsync(redoBytes)
  const redoSheetXml = await redoZip.file('xl/worksheets/sheet1.xml')!.async('text')
  expect(redoSheetXml).toMatch(
    /<col\b(?=[^>]*\bmin="3")(?=[^>]*\bmax="3")(?=[^>]*\bwidth="12\.4296875")(?=[^>]*\bcustomWidth="1")/,
  )
  expect(redoSheetXml).toMatch(
    /<col\b(?=[^>]*\bmin="4")(?=[^>]*\bmax="4")(?=[^>]*\bwidth="12\.4296875")(?=[^>]*\bcustomWidth="1")/,
  )

  await page.reload()
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)
  const reopened = page.frameLocator('#editor-frame')
  await reopened.locator('input[type="file"][accept*=".xlsx"]').setInputFiles({
    name: 'registry-column-width-reopened.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: redoBytes,
  })
  await page.evaluate(() => {
    window.__codexVisualHost.commands.push({
      commandId: 'xlsx-column-width-reopened-touch',
      baseRevision: 0,
      operation: 'xlsx.cell.set_value',
      arguments: { sheet: 'Sheet1', address: 'A1', value: 'Sized column' },
    })
  })
  await expect
    .poll(() =>
      page.evaluate(() =>
        window.__codexVisualHost.acknowledgements.find(
          (entry) => entry.commandId === 'xlsx-column-width-reopened-touch',
        ),
      ),
    )
    .toMatchObject({ ok: true, revision: 1 })
  const reopenedDownloadPromise = page.waitForEvent('download')
  await reopened.locator('button.qa-btn').first().click()
  const reopenedPath = await (await reopenedDownloadPromise).path()
  expect(reopenedPath).not.toBeNull()
  const reopenedZip = await JSZip.loadAsync(
    await import('node:fs/promises').then((fs) => fs.readFile(reopenedPath!)),
  )
  const reopenedSheetXml = await reopenedZip.file('xl/worksheets/sheet1.xml')!.async('text')
  expect(reopenedSheetXml).toContain('width="12.4296875"')
})

test('the XLSX freeze registry operation shares user undo and saved output', async ({ page }) => {
  await page.goto('/?format=xlsx&width=720&height=900')
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)
  const editor = page.frameLocator('#editor-frame')
  await editor.locator('input[type="file"][accept*=".xlsx"]').setInputFiles({
    name: 'registry-freeze.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: await xlsxWithA1('Frozen sheet'),
  })

  await page.evaluate(() => {
    window.__codexVisualHost.commands.push({
      commandId: 'xlsx-sheet-freeze',
      baseRevision: 0,
      operation: 'xlsx.sheet.set_freeze',
      arguments: { sheet: 'Sheet1', frozenRows: 2, frozenColumns: 1 },
    })
  })
  await expect
    .poll(() =>
      page.evaluate(() =>
        window.__codexVisualHost.acknowledgements.find(
          (entry) => entry.commandId === 'xlsx-sheet-freeze',
        ),
      ),
    )
    .toMatchObject({
      ok: true,
      revision: 1,
      output: { sheet: 'Sheet1', frozenRows: 2, frozenColumns: 1 },
    })

  await editor.locator('button.qa-btn').nth(1).click()
  await expect(editor.locator('button.qa-btn').nth(2)).toBeEnabled()
  await editor.locator('button.qa-btn').nth(2).click()
  await expect(editor.locator('button.qa-btn').nth(1)).toBeEnabled()
  const redoDownloadPromise = page.waitForEvent('download')
  await editor.locator('button.qa-btn').first().click()
  const redoPath = await (await redoDownloadPromise).path()
  expect(redoPath).not.toBeNull()
  const redoBytes = await import('node:fs/promises').then((fs) => fs.readFile(redoPath!))
  const redoZip = await JSZip.loadAsync(redoBytes)
  const redoSheetXml = await redoZip.file('xl/worksheets/sheet1.xml')!.async('text')
  expect(redoSheetXml).toMatch(
    /<pane\b(?=[^>]*\bxSplit="1")(?=[^>]*\bySplit="2")(?=[^>]*\btopLeftCell="B3")(?=[^>]*\bstate="frozen")/,
  )

  await page.reload()
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)
  const reopened = page.frameLocator('#editor-frame')
  await reopened.locator('input[type="file"][accept*=".xlsx"]').setInputFiles({
    name: 'registry-freeze-reopened.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: redoBytes,
  })
  await page.evaluate(() => {
    window.__codexVisualHost.commands.push({
      commandId: 'xlsx-sheet-freeze-reopened-touch',
      baseRevision: 0,
      operation: 'xlsx.cell.set_value',
      arguments: { sheet: 'Sheet1', address: 'A1', value: 'Frozen sheet' },
    })
  })
  await expect
    .poll(() =>
      page.evaluate(() =>
        window.__codexVisualHost.acknowledgements.find(
          (entry) => entry.commandId === 'xlsx-sheet-freeze-reopened-touch',
        ),
      ),
    )
    .toMatchObject({ ok: true, revision: 1 })
  const reopenedDownloadPromise = page.waitForEvent('download')
  await reopened.locator('button.qa-btn').first().click()
  const reopenedPath = await (await reopenedDownloadPromise).path()
  expect(reopenedPath).not.toBeNull()
  const reopenedZip = await JSZip.loadAsync(
    await import('node:fs/promises').then((fs) => fs.readFile(reopenedPath!)),
  )
  const reopenedSheetXml = await reopenedZip.file('xl/worksheets/sheet1.xml')!.async('text')
  expect(reopenedSheetXml).toContain('topLeftCell="B3"')
})

test('the XLSX gridline registry operation shares user undo and saved output', async ({ page }) => {
  await page.goto('/?format=xlsx&width=720&height=900')
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)
  const editor = page.frameLocator('#editor-frame')
  await editor.locator('input[type="file"][accept*=".xlsx"]').setInputFiles({
    name: 'registry-gridlines.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: await xlsxWithA1('Gridlines'),
  })

  await page.evaluate(() => {
    window.__codexVisualHost.commands.push({
      commandId: 'xlsx-sheet-gridlines',
      baseRevision: 0,
      operation: 'xlsx.sheet.set_gridlines',
      arguments: { sheet: 'Sheet1', visible: false },
    })
  })
  await expect
    .poll(() =>
      page.evaluate(() =>
        window.__codexVisualHost.acknowledgements.find(
          (entry) => entry.commandId === 'xlsx-sheet-gridlines',
        ),
      ),
    )
    .toMatchObject({
      ok: true,
      revision: 1,
      output: { sheet: 'Sheet1', visible: false },
    })

  await editor.locator('button.qa-btn').nth(1).click()
  await expect(editor.locator('button.qa-btn').nth(2)).toBeEnabled()
  await editor.locator('button.qa-btn').nth(2).click()
  await expect(editor.locator('button.qa-btn').nth(1)).toBeEnabled()
  const redoDownloadPromise = page.waitForEvent('download')
  await editor.locator('button.qa-btn').first().click()
  const redoPath = await (await redoDownloadPromise).path()
  expect(redoPath).not.toBeNull()
  const redoBytes = await import('node:fs/promises').then((fs) => fs.readFile(redoPath!))
  const redoZip = await JSZip.loadAsync(redoBytes)
  const redoSheetXml = await redoZip.file('xl/worksheets/sheet1.xml')!.async('text')
  expect(redoSheetXml).toMatch(/<sheetView\b[^>]*\bshowGridLines="0"/)

  await page.reload()
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)
  const reopened = page.frameLocator('#editor-frame')
  await reopened.locator('input[type="file"][accept*=".xlsx"]').setInputFiles({
    name: 'registry-gridlines-reopened.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: redoBytes,
  })
  await page.evaluate(() => {
    window.__codexVisualHost.commands.push({
      commandId: 'xlsx-sheet-gridlines-reopened-touch',
      baseRevision: 0,
      operation: 'xlsx.cell.set_value',
      arguments: { sheet: 'Sheet1', address: 'A1', value: 'Gridlines' },
    })
  })
  await expect
    .poll(() =>
      page.evaluate(() =>
        window.__codexVisualHost.acknowledgements.find(
          (entry) => entry.commandId === 'xlsx-sheet-gridlines-reopened-touch',
        ),
      ),
    )
    .toMatchObject({ ok: true, revision: 1 })
  const reopenedDownloadPromise = page.waitForEvent('download')
  await reopened.locator('button.qa-btn').first().click()
  const reopenedPath = await (await reopenedDownloadPromise).path()
  expect(reopenedPath).not.toBeNull()
  const reopenedZip = await JSZip.loadAsync(
    await import('node:fs/promises').then((fs) => fs.readFile(reopenedPath!)),
  )
  const reopenedSheetXml = await reopenedZip.file('xl/worksheets/sheet1.xml')!.async('text')
  expect(reopenedSheetXml).toContain('showGridLines="0"')
})

test('the XLSX formula-view registry operation shares user undo and saved output', async ({
  page,
}) => {
  test.slow()
  await page.goto('/?format=xlsx&width=720&height=900')
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)
  const editor = page.frameLocator('#editor-frame')
  await editor.locator('input[type="file"][accept*=".xlsx"]').setInputFiles({
    name: 'registry-formula-view.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: await xlsxWithA1('Formula view'),
  })

  await page.evaluate(() => {
    window.__codexVisualHost.commands.push({
      commandId: 'xlsx-sheet-formula-view',
      baseRevision: 0,
      operation: 'xlsx.sheet.set_formula_view',
      arguments: { sheet: 'Sheet1', enabled: true },
    })
  })
  await expect
    .poll(() =>
      page.evaluate(() =>
        window.__codexVisualHost.acknowledgements.find(
          (entry) => entry.commandId === 'xlsx-sheet-formula-view',
        ),
      ),
    )
    .toMatchObject({
      ok: true,
      revision: 1,
      output: { sheet: 'Sheet1', enabled: true },
    })

  await expect(editor.locator('button.qa-btn').nth(1)).toBeEnabled()
  await editor.locator('button.qa-btn').nth(1).click()
  await expect(editor.locator('button.qa-btn').nth(2)).toBeEnabled()
  const undoDownloadPromise = page.waitForEvent('download')
  await editor.locator('button.qa-btn').first().click()
  const undoPath = await (await undoDownloadPromise).path()
  expect(undoPath).not.toBeNull()
  const undoZip = await JSZip.loadAsync(
    await import('node:fs/promises').then((fs) => fs.readFile(undoPath!)),
  )
  const undoSheetXml = await undoZip.file('xl/worksheets/sheet1.xml')!.async('text')
  expect(undoSheetXml).not.toContain('showFormulas')

  await page.evaluate(() => {
    window.__codexVisualHost.commands.push({
      commandId: 'xlsx-sheet-formula-view-redo-cycle',
      baseRevision: 1,
      operation: 'xlsx.sheet.set_formula_view',
      arguments: { sheet: 'Sheet1', enabled: true },
    })
  })
  await expect
    .poll(() =>
      page.evaluate(() =>
        window.__codexVisualHost.acknowledgements.find(
          (entry) => entry.commandId === 'xlsx-sheet-formula-view-redo-cycle',
        ),
      ),
    )
    .toMatchObject({ ok: true, revision: 2 })
  await editor.locator('button.qa-btn').nth(1).click()
  await expect(editor.locator('button.qa-btn').nth(2)).toBeEnabled()
  await editor.locator('button.qa-btn').nth(2).click()
  await expect(editor.locator('button.qa-btn').nth(1)).toBeEnabled()
  const redoDownloadPromise = page.waitForEvent('download')
  await editor.locator('button.qa-btn').first().click()
  const redoPath = await (await redoDownloadPromise).path()
  expect(redoPath).not.toBeNull()
  const redoBytes = await import('node:fs/promises').then((fs) => fs.readFile(redoPath!))
  const redoZip = await JSZip.loadAsync(redoBytes)
  const redoSheetXml = await redoZip.file('xl/worksheets/sheet1.xml')!.async('text')
  expect(redoSheetXml).toMatch(/<sheetView\b[^>]*\bshowFormulas="1"/)

  await page.reload()
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)
  const reopened = page.frameLocator('#editor-frame')
  await reopened.locator('input[type="file"][accept*=".xlsx"]').setInputFiles({
    name: 'registry-formula-view-reopened.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: redoBytes,
  })
  await page.evaluate(() => {
    window.__codexVisualHost.commands.push({
      commandId: 'xlsx-sheet-formula-view-reopened-touch',
      baseRevision: 0,
      operation: 'xlsx.cell.set_value',
      arguments: { sheet: 'Sheet1', address: 'A1', value: 'Formula view' },
    })
  })
  await expect
    .poll(() =>
      page.evaluate(() =>
        window.__codexVisualHost.acknowledgements.find(
          (entry) => entry.commandId === 'xlsx-sheet-formula-view-reopened-touch',
        ),
      ),
    )
    .toMatchObject({ ok: true, revision: 1 })
  const reopenedDownloadPromise = page.waitForEvent('download')
  await reopened.locator('button.qa-btn').first().click()
  const reopenedPath = await (await reopenedDownloadPromise).path()
  expect(reopenedPath).not.toBeNull()
  const reopenedZip = await JSZip.loadAsync(
    await import('node:fs/promises').then((fs) => fs.readFile(reopenedPath!)),
  )
  const reopenedSheetXml = await reopenedZip.file('xl/worksheets/sheet1.xml')!.async('text')
  expect(reopenedSheetXml).toContain('showFormulas="1"')
})

test('the XLSX page-orientation registry operation shares user undo and saved output', async ({
  page,
}) => {
  await page.goto('/?format=xlsx&width=720&height=900')
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)
  const editor = page.frameLocator('#editor-frame')
  await editor.locator('input[type="file"][accept*=".xlsx"]').setInputFiles({
    name: 'registry-page-orientation.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: await xlsxWithA1('Page orientation'),
  })

  await page.evaluate(() => {
    window.__codexVisualHost.commands.push({
      commandId: 'xlsx-sheet-page-orientation',
      baseRevision: 0,
      operation: 'xlsx.sheet.set_page_orientation',
      arguments: { sheet: 'Sheet1', orientation: 'landscape' },
    })
  })
  await expect
    .poll(() =>
      page.evaluate(() =>
        window.__codexVisualHost.acknowledgements.find(
          (entry) => entry.commandId === 'xlsx-sheet-page-orientation',
        ),
      ),
    )
    .toMatchObject({
      ok: true,
      revision: 1,
      output: { sheet: 'Sheet1', orientation: 'landscape' },
    })

  await expect(editor.locator('button.qa-btn').nth(1)).toBeEnabled()
  await editor.locator('button.qa-btn').nth(1).click()
  await expect(editor.locator('button.qa-btn').nth(2)).toBeEnabled()
  await page.evaluate(() => {
    window.__codexVisualHost.commands.push({
      commandId: 'xlsx-sheet-page-orientation-undo-persist',
      baseRevision: 1,
      operation: 'xlsx.cell.set_value',
      arguments: { sheet: 'Sheet1', address: 'B1', value: 'Undo persisted' },
    })
  })
  await expect
    .poll(() =>
      page.evaluate(() =>
        window.__codexVisualHost.acknowledgements.find(
          (entry) => entry.commandId === 'xlsx-sheet-page-orientation-undo-persist',
        ),
      ),
    )
    .toMatchObject({ ok: true, revision: 2 })
  const undoDownloadPromise = page.waitForEvent('download')
  await editor.locator('button.qa-btn').first().click()
  const undoPath = await (await undoDownloadPromise).path()
  expect(undoPath).not.toBeNull()
  const undoZip = await JSZip.loadAsync(
    await import('node:fs/promises').then((fs) => fs.readFile(undoPath!)),
  )
  const undoSheetXml = await undoZip.file('xl/worksheets/sheet1.xml')!.async('text')
  expect(undoSheetXml).not.toMatch(/<pageSetup\b[^>]*\borientation=/)

  await page.evaluate(() => {
    window.__codexVisualHost.commands.push({
      commandId: 'xlsx-sheet-page-orientation-redo-cycle',
      baseRevision: 2,
      operation: 'xlsx.sheet.set_page_orientation',
      arguments: { sheet: 'Sheet1', orientation: 'landscape' },
    })
  })
  await expect
    .poll(() =>
      page.evaluate(() =>
        window.__codexVisualHost.acknowledgements.find(
          (entry) => entry.commandId === 'xlsx-sheet-page-orientation-redo-cycle',
        ),
      ),
    )
    .toMatchObject({ ok: true, revision: 3 })
  await editor.locator('button.qa-btn').nth(1).click()
  await expect(editor.locator('button.qa-btn').nth(2)).toBeEnabled()
  await editor.locator('button.qa-btn').nth(2).click()
  await expect(editor.locator('button.qa-btn').nth(1)).toBeEnabled()
  const redoDownloadPromise = page.waitForEvent('download')
  await editor.locator('button.qa-btn').first().click()
  const redoPath = await (await redoDownloadPromise).path()
  expect(redoPath).not.toBeNull()
  const redoBytes = await import('node:fs/promises').then((fs) => fs.readFile(redoPath!))
  const redoZip = await JSZip.loadAsync(redoBytes)
  const redoSheetXml = await redoZip.file('xl/worksheets/sheet1.xml')!.async('text')
  expect(redoSheetXml).toMatch(/<pageSetup\b[^>]*\borientation="landscape"/)

  await page.reload()
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)
  const reopened = page.frameLocator('#editor-frame')
  await reopened.locator('input[type="file"][accept*=".xlsx"]').setInputFiles({
    name: 'registry-page-orientation-reopened.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: redoBytes,
  })
  await page.evaluate(() => {
    window.__codexVisualHost.commands.push({
      commandId: 'xlsx-sheet-page-orientation-reopened-touch',
      baseRevision: 0,
      operation: 'xlsx.cell.set_value',
      arguments: { sheet: 'Sheet1', address: 'A1', value: 'Page orientation' },
    })
  })
  await expect
    .poll(() =>
      page.evaluate(() =>
        window.__codexVisualHost.acknowledgements.find(
          (entry) => entry.commandId === 'xlsx-sheet-page-orientation-reopened-touch',
        ),
      ),
    )
    .toMatchObject({ ok: true, revision: 1 })
  const reopenedDownloadPromise = page.waitForEvent('download')
  await reopened.locator('button.qa-btn').first().click()
  const reopenedPath = await (await reopenedDownloadPromise).path()
  expect(reopenedPath).not.toBeNull()
  const reopenedZip = await JSZip.loadAsync(
    await import('node:fs/promises').then((fs) => fs.readFile(reopenedPath!)),
  )
  const reopenedSheetXml = await reopenedZip.file('xl/worksheets/sheet1.xml')!.async('text')
  expect(reopenedSheetXml).toContain('orientation="landscape"')
})

test('the XLSX page-margins registry operation shares user undo and saved output', async ({
  page,
}) => {
  await page.goto('/?format=xlsx&width=720&height=900')
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)
  const editor = page.frameLocator('#editor-frame')
  await editor.locator('input[type="file"][accept*=".xlsx"]').setInputFiles({
    name: 'registry-page-margins.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: await xlsxWithA1('Page margins'),
  })

  await page.evaluate(() => {
    window.__codexVisualHost.commands.push({
      commandId: 'xlsx-sheet-page-margins',
      baseRevision: 0,
      operation: 'xlsx.sheet.set_page_margins',
      arguments: { sheet: 'Sheet1', margins: 'wide' },
    })
  })
  await expect
    .poll(() =>
      page.evaluate(() =>
        window.__codexVisualHost.acknowledgements.find(
          (entry) => entry.commandId === 'xlsx-sheet-page-margins',
        ),
      ),
    )
    .toMatchObject({
      ok: true,
      revision: 1,
      output: { sheet: 'Sheet1', margins: 'wide' },
    })

  await expect(editor.locator('button.qa-btn').nth(1)).toBeEnabled()
  await editor.locator('button.qa-btn').nth(1).click()
  await expect(editor.locator('button.qa-btn').nth(2)).toBeEnabled()
  await page.evaluate(() => {
    window.__codexVisualHost.commands.push({
      commandId: 'xlsx-sheet-page-margins-undo-persist',
      baseRevision: 1,
      operation: 'xlsx.cell.set_value',
      arguments: { sheet: 'Sheet1', address: 'B1', value: 'Undo persisted' },
    })
  })
  await expect
    .poll(() =>
      page.evaluate(() =>
        window.__codexVisualHost.acknowledgements.find(
          (entry) => entry.commandId === 'xlsx-sheet-page-margins-undo-persist',
        ),
      ),
    )
    .toMatchObject({ ok: true, revision: 2 })
  const undoDownloadPromise = page.waitForEvent('download')
  await editor.locator('button.qa-btn').first().click()
  const undoPath = await (await undoDownloadPromise).path()
  expect(undoPath).not.toBeNull()
  const undoZip = await JSZip.loadAsync(
    await import('node:fs/promises').then((fs) => fs.readFile(undoPath!)),
  )
  const undoSheetXml = await undoZip.file('xl/worksheets/sheet1.xml')!.async('text')
  expect(undoSheetXml).not.toContain('<pageMargins')

  await page.evaluate(() => {
    window.__codexVisualHost.commands.push({
      commandId: 'xlsx-sheet-page-margins-redo-cycle',
      baseRevision: 2,
      operation: 'xlsx.sheet.set_page_margins',
      arguments: { sheet: 'Sheet1', margins: 'wide' },
    })
  })
  await expect
    .poll(() =>
      page.evaluate(() =>
        window.__codexVisualHost.acknowledgements.find(
          (entry) => entry.commandId === 'xlsx-sheet-page-margins-redo-cycle',
        ),
      ),
    )
    .toMatchObject({ ok: true, revision: 3 })
  await editor.locator('button.qa-btn').nth(1).click()
  await expect(editor.locator('button.qa-btn').nth(2)).toBeEnabled()
  await editor.locator('button.qa-btn').nth(2).click()
  await expect(editor.locator('button.qa-btn').nth(1)).toBeEnabled()
  const redoDownloadPromise = page.waitForEvent('download')
  await editor.locator('button.qa-btn').first().click()
  const redoPath = await (await redoDownloadPromise).path()
  expect(redoPath).not.toBeNull()
  const redoBytes = await import('node:fs/promises').then((fs) => fs.readFile(redoPath!))
  const redoZip = await JSZip.loadAsync(redoBytes)
  const redoSheetXml = await redoZip.file('xl/worksheets/sheet1.xml')!.async('text')
  expect(redoSheetXml).toMatch(
    /<pageMargins\b(?=[^>]*\bleft="1")(?=[^>]*\bright="1")(?=[^>]*\btop="1")(?=[^>]*\bbottom="1")(?=[^>]*\bheader="0\.5")(?=[^>]*\bfooter="0\.5")/,
  )

  await page.reload()
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)
  const reopened = page.frameLocator('#editor-frame')
  await reopened.locator('input[type="file"][accept*=".xlsx"]').setInputFiles({
    name: 'registry-page-margins-reopened.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: redoBytes,
  })
  await page.evaluate(() => {
    window.__codexVisualHost.commands.push({
      commandId: 'xlsx-sheet-page-margins-reopened-touch',
      baseRevision: 0,
      operation: 'xlsx.cell.set_value',
      arguments: { sheet: 'Sheet1', address: 'A1', value: 'Page margins' },
    })
  })
  await expect
    .poll(() =>
      page.evaluate(() =>
        window.__codexVisualHost.acknowledgements.find(
          (entry) => entry.commandId === 'xlsx-sheet-page-margins-reopened-touch',
        ),
      ),
    )
    .toMatchObject({ ok: true, revision: 1 })
  const reopenedDownloadPromise = page.waitForEvent('download')
  await reopened.locator('button.qa-btn').first().click()
  const reopenedPath = await (await reopenedDownloadPromise).path()
  expect(reopenedPath).not.toBeNull()
  const reopenedZip = await JSZip.loadAsync(
    await import('node:fs/promises').then((fs) => fs.readFile(reopenedPath!)),
  )
  const reopenedSheetXml = await reopenedZip.file('xl/worksheets/sheet1.xml')!.async('text')
  expect(reopenedSheetXml).toContain('left="1"')
})

test('the XLSX paper-size registry operation shares user undo and saved output', async ({
  page,
}) => {
  await page.goto('/?format=xlsx&width=720&height=900')
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)
  const editor = page.frameLocator('#editor-frame')
  await editor.locator('input[type="file"][accept*=".xlsx"]').setInputFiles({
    name: 'registry-paper-size.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: await xlsxWithA1('Paper size'),
  })

  await page.evaluate(() => {
    window.__codexVisualHost.commands.push({
      commandId: 'xlsx-sheet-paper-size',
      baseRevision: 0,
      operation: 'xlsx.sheet.set_paper_size',
      arguments: { sheet: 'Sheet1', paperSize: 9 },
    })
  })
  await expect
    .poll(() =>
      page.evaluate(() =>
        window.__codexVisualHost.acknowledgements.find(
          (entry) => entry.commandId === 'xlsx-sheet-paper-size',
        ),
      ),
    )
    .toMatchObject({
      ok: true,
      revision: 1,
      output: { sheet: 'Sheet1', paperSize: 9 },
    })

  await expect(editor.locator('button.qa-btn').nth(1)).toBeEnabled()
  await editor.locator('button.qa-btn').nth(1).click()
  await expect(editor.locator('button.qa-btn').nth(2)).toBeEnabled()
  await page.evaluate(() => {
    window.__codexVisualHost.commands.push({
      commandId: 'xlsx-sheet-paper-size-undo-persist',
      baseRevision: 1,
      operation: 'xlsx.cell.set_value',
      arguments: { sheet: 'Sheet1', address: 'B1', value: 'Undo persisted' },
    })
  })
  await expect
    .poll(() =>
      page.evaluate(() =>
        window.__codexVisualHost.acknowledgements.find(
          (entry) => entry.commandId === 'xlsx-sheet-paper-size-undo-persist',
        ),
      ),
    )
    .toMatchObject({ ok: true, revision: 2 })
  const undoDownloadPromise = page.waitForEvent('download')
  await editor.locator('button.qa-btn').first().click()
  const undoPath = await (await undoDownloadPromise).path()
  expect(undoPath).not.toBeNull()
  const undoZip = await JSZip.loadAsync(
    await import('node:fs/promises').then((fs) => fs.readFile(undoPath!)),
  )
  const undoSheetXml = await undoZip.file('xl/worksheets/sheet1.xml')!.async('text')
  expect(undoSheetXml).not.toMatch(/<pageSetup\b[^>]*\bpaperSize=/)

  await page.evaluate(() => {
    window.__codexVisualHost.commands.push({
      commandId: 'xlsx-sheet-paper-size-redo-cycle',
      baseRevision: 2,
      operation: 'xlsx.sheet.set_paper_size',
      arguments: { sheet: 'Sheet1', paperSize: 9 },
    })
  })
  await expect
    .poll(() =>
      page.evaluate(() =>
        window.__codexVisualHost.acknowledgements.find(
          (entry) => entry.commandId === 'xlsx-sheet-paper-size-redo-cycle',
        ),
      ),
    )
    .toMatchObject({ ok: true, revision: 3 })
  await editor.locator('button.qa-btn').nth(1).click()
  await expect(editor.locator('button.qa-btn').nth(2)).toBeEnabled()
  await editor.locator('button.qa-btn').nth(2).click()
  await expect(editor.locator('button.qa-btn').nth(1)).toBeEnabled()
  const redoDownloadPromise = page.waitForEvent('download')
  await editor.locator('button.qa-btn').first().click()
  const redoPath = await (await redoDownloadPromise).path()
  expect(redoPath).not.toBeNull()
  const redoBytes = await import('node:fs/promises').then((fs) => fs.readFile(redoPath!))
  const redoZip = await JSZip.loadAsync(redoBytes)
  const redoSheetXml = await redoZip.file('xl/worksheets/sheet1.xml')!.async('text')
  expect(redoSheetXml).toMatch(/<pageSetup\b[^>]*\bpaperSize="9"/)

  await page.reload()
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)
  const reopened = page.frameLocator('#editor-frame')
  await reopened.locator('input[type="file"][accept*=".xlsx"]').setInputFiles({
    name: 'registry-paper-size-reopened.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: redoBytes,
  })
  await page.evaluate(() => {
    window.__codexVisualHost.commands.push({
      commandId: 'xlsx-sheet-paper-size-reopened-touch',
      baseRevision: 0,
      operation: 'xlsx.cell.set_value',
      arguments: { sheet: 'Sheet1', address: 'A1', value: 'Paper size' },
    })
  })
  await expect
    .poll(() =>
      page.evaluate(() =>
        window.__codexVisualHost.acknowledgements.find(
          (entry) => entry.commandId === 'xlsx-sheet-paper-size-reopened-touch',
        ),
      ),
    )
    .toMatchObject({ ok: true, revision: 1 })
  const reopenedDownloadPromise = page.waitForEvent('download')
  await reopened.locator('button.qa-btn').first().click()
  const reopenedPath = await (await reopenedDownloadPromise).path()
  expect(reopenedPath).not.toBeNull()
  const reopenedZip = await JSZip.loadAsync(
    await import('node:fs/promises').then((fs) => fs.readFile(reopenedPath!)),
  )
  const reopenedSheetXml = await reopenedZip.file('xl/worksheets/sheet1.xml')!.async('text')
  expect(reopenedSheetXml).toContain('paperSize="9"')
})

test('the XLSX fit-to-pages registry operation shares user undo and saved output', async ({
  page,
}) => {
  await page.goto('/?format=xlsx&width=720&height=900')
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)
  const editor = page.frameLocator('#editor-frame')
  await editor.locator('input[type="file"][accept*=".xlsx"]').setInputFiles({
    name: 'registry-fit-to-pages.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: await xlsxWithA1('Fit to pages'),
  })

  await page.evaluate(() => {
    window.__codexVisualHost.commands.push({
      commandId: 'xlsx-sheet-fit-to-pages',
      baseRevision: 0,
      operation: 'xlsx.sheet.set_fit_to_pages',
      arguments: { sheet: 'Sheet1', widthPages: 2, heightPages: 3 },
    })
  })
  await expect
    .poll(() =>
      page.evaluate(() =>
        window.__codexVisualHost.acknowledgements.find(
          (entry) => entry.commandId === 'xlsx-sheet-fit-to-pages',
        ),
      ),
    )
    .toMatchObject({
      ok: true,
      revision: 1,
      output: { sheet: 'Sheet1', widthPages: 2, heightPages: 3 },
    })

  await expect(editor.locator('button.qa-btn').nth(1)).toBeEnabled()
  await editor.locator('button.qa-btn').nth(1).click()
  await expect(editor.locator('button.qa-btn').nth(2)).toBeEnabled()
  await page.evaluate(() => {
    window.__codexVisualHost.commands.push({
      commandId: 'xlsx-sheet-fit-to-pages-undo-persist',
      baseRevision: 1,
      operation: 'xlsx.cell.set_value',
      arguments: { sheet: 'Sheet1', address: 'B1', value: 'Undo persisted' },
    })
  })
  await expect
    .poll(() =>
      page.evaluate(() =>
        window.__codexVisualHost.acknowledgements.find(
          (entry) => entry.commandId === 'xlsx-sheet-fit-to-pages-undo-persist',
        ),
      ),
    )
    .toMatchObject({ ok: true, revision: 2 })
  const undoDownloadPromise = page.waitForEvent('download')
  await editor.locator('button.qa-btn').first().click()
  const undoPath = await (await undoDownloadPromise).path()
  expect(undoPath).not.toBeNull()
  const undoZip = await JSZip.loadAsync(
    await import('node:fs/promises').then((fs) => fs.readFile(undoPath!)),
  )
  const undoSheetXml = await undoZip.file('xl/worksheets/sheet1.xml')!.async('text')
  expect(undoSheetXml).not.toMatch(/<pageSetup\b[^>]*\bfitTo(?:Width|Height)=/)
  expect(undoSheetXml).not.toMatch(/<pageSetUpPr\b[^>]*\bfitToPage=/)

  await page.evaluate(() => {
    window.__codexVisualHost.commands.push({
      commandId: 'xlsx-sheet-fit-to-pages-redo-cycle',
      baseRevision: 2,
      operation: 'xlsx.sheet.set_fit_to_pages',
      arguments: { sheet: 'Sheet1', widthPages: 2, heightPages: 3 },
    })
  })
  await expect
    .poll(() =>
      page.evaluate(() =>
        window.__codexVisualHost.acknowledgements.find(
          (entry) => entry.commandId === 'xlsx-sheet-fit-to-pages-redo-cycle',
        ),
      ),
    )
    .toMatchObject({ ok: true, revision: 3 })
  await editor.locator('button.qa-btn').nth(1).click()
  await expect(editor.locator('button.qa-btn').nth(2)).toBeEnabled()
  await editor.locator('button.qa-btn').nth(2).click()
  await expect(editor.locator('button.qa-btn').nth(1)).toBeEnabled()
  const redoDownloadPromise = page.waitForEvent('download')
  await editor.locator('button.qa-btn').first().click()
  const redoPath = await (await redoDownloadPromise).path()
  expect(redoPath).not.toBeNull()
  const redoBytes = await import('node:fs/promises').then((fs) => fs.readFile(redoPath!))
  const redoZip = await JSZip.loadAsync(redoBytes)
  const redoSheetXml = await redoZip.file('xl/worksheets/sheet1.xml')!.async('text')
  expect(redoSheetXml).toMatch(/<pageSetup\b(?=[^>]*\bfitToWidth="2")(?=[^>]*\bfitToHeight="3")/)
  expect(redoSheetXml).toMatch(/<pageSetUpPr\b[^>]*\bfitToPage="1"/)

  await page.reload()
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)
  const reopened = page.frameLocator('#editor-frame')
  await reopened.locator('input[type="file"][accept*=".xlsx"]').setInputFiles({
    name: 'registry-fit-to-pages-reopened.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: redoBytes,
  })
  await page.evaluate(() => {
    window.__codexVisualHost.commands.push({
      commandId: 'xlsx-sheet-fit-to-pages-reopened-touch',
      baseRevision: 0,
      operation: 'xlsx.cell.set_value',
      arguments: { sheet: 'Sheet1', address: 'A1', value: 'Fit to pages' },
    })
  })
  await expect
    .poll(() =>
      page.evaluate(() =>
        window.__codexVisualHost.acknowledgements.find(
          (entry) => entry.commandId === 'xlsx-sheet-fit-to-pages-reopened-touch',
        ),
      ),
    )
    .toMatchObject({ ok: true, revision: 1 })
  const reopenedDownloadPromise = page.waitForEvent('download')
  await reopened.locator('button.qa-btn').first().click()
  const reopenedPath = await (await reopenedDownloadPromise).path()
  expect(reopenedPath).not.toBeNull()
  const reopenedZip = await JSZip.loadAsync(
    await import('node:fs/promises').then((fs) => fs.readFile(reopenedPath!)),
  )
  const reopenedSheetXml = await reopenedZip.file('xl/worksheets/sheet1.xml')!.async('text')
  expect(reopenedSheetXml).toMatch(
    /<pageSetup\b(?=[^>]*\bfitToWidth="2")(?=[^>]*\bfitToHeight="3")/,
  )
  expect(reopenedSheetXml).toMatch(/<pageSetUpPr\b[^>]*\bfitToPage="1"/)
})

test('the XLSX print-scale registry operation shares user undo and saved output', async ({
  page,
}) => {
  await page.goto('/?format=xlsx&width=720&height=900')
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)
  const editor = page.frameLocator('#editor-frame')
  await editor.locator('input[type="file"][accept*=".xlsx"]').setInputFiles({
    name: 'registry-print-scale.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: await xlsxWithFitToPages('Print scale'),
  })

  await page.evaluate(() => {
    window.__codexVisualHost.commands.push({
      commandId: 'xlsx-sheet-print-scale',
      baseRevision: 0,
      operation: 'xlsx.sheet.set_print_scale',
      arguments: { sheet: 'Sheet1', scalePercent: 80 },
    })
  })
  await expect
    .poll(() =>
      page.evaluate(() =>
        window.__codexVisualHost.acknowledgements.find(
          (entry) => entry.commandId === 'xlsx-sheet-print-scale',
        ),
      ),
    )
    .toMatchObject({
      ok: true,
      revision: 1,
      output: { sheet: 'Sheet1', scalePercent: 80 },
    })

  await expect(editor.locator('button.qa-btn').nth(1)).toBeEnabled()
  await editor.locator('button.qa-btn').nth(1).click()
  await expect(editor.locator('button.qa-btn').nth(2)).toBeEnabled()
  await page.evaluate(() => {
    window.__codexVisualHost.commands.push({
      commandId: 'xlsx-sheet-print-scale-undo-persist',
      baseRevision: 1,
      operation: 'xlsx.cell.set_value',
      arguments: { sheet: 'Sheet1', address: 'B1', value: 'Undo persisted' },
    })
  })
  await expect
    .poll(() =>
      page.evaluate(() =>
        window.__codexVisualHost.acknowledgements.find(
          (entry) => entry.commandId === 'xlsx-sheet-print-scale-undo-persist',
        ),
      ),
    )
    .toMatchObject({ ok: true, revision: 2 })
  const undoDownloadPromise = page.waitForEvent('download')
  await editor.locator('button.qa-btn').first().click()
  const undoPath = await (await undoDownloadPromise).path()
  expect(undoPath).not.toBeNull()
  const undoZip = await JSZip.loadAsync(
    await import('node:fs/promises').then((fs) => fs.readFile(undoPath!)),
  )
  const undoSheetXml = await undoZip.file('xl/worksheets/sheet1.xml')!.async('text')
  expect(undoSheetXml).toMatch(/<pageSetUpPr\b[^>]*\bfitToPage="1"/)
  expect(undoSheetXml).not.toMatch(/<pageSetup\b[^>]*\bscale=/)

  await page.evaluate(() => {
    window.__codexVisualHost.commands.push({
      commandId: 'xlsx-sheet-print-scale-redo-cycle',
      baseRevision: 2,
      operation: 'xlsx.sheet.set_print_scale',
      arguments: { sheet: 'Sheet1', scalePercent: 80 },
    })
  })
  await expect
    .poll(() =>
      page.evaluate(() =>
        window.__codexVisualHost.acknowledgements.find(
          (entry) => entry.commandId === 'xlsx-sheet-print-scale-redo-cycle',
        ),
      ),
    )
    .toMatchObject({ ok: true, revision: 3 })
  await editor.locator('button.qa-btn').nth(1).click()
  await expect(editor.locator('button.qa-btn').nth(2)).toBeEnabled()
  await editor.locator('button.qa-btn').nth(2).click()
  const redoDownloadPromise = page.waitForEvent('download')
  await editor.locator('button.qa-btn').first().click()
  const redoPath = await (await redoDownloadPromise).path()
  expect(redoPath).not.toBeNull()
  const redoBytes = await import('node:fs/promises').then((fs) => fs.readFile(redoPath!))
  const redoZip = await JSZip.loadAsync(redoBytes)
  const redoSheetXml = await redoZip.file('xl/worksheets/sheet1.xml')!.async('text')
  expect(redoSheetXml).toMatch(/<pageSetup\b[^>]*\bscale="80"/)
  expect(redoSheetXml).not.toMatch(/<pageSetUpPr\b[^>]*\bfitToPage=/)

  await page.reload()
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)
  const reopened = page.frameLocator('#editor-frame')
  await reopened.locator('input[type="file"][accept*=".xlsx"]').setInputFiles({
    name: 'registry-print-scale-reopened.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: redoBytes,
  })
  await page.evaluate(() => {
    window.__codexVisualHost.commands.push({
      commandId: 'xlsx-sheet-print-scale-reopened-touch',
      baseRevision: 0,
      operation: 'xlsx.cell.set_value',
      arguments: { sheet: 'Sheet1', address: 'A1', value: 'Print scale' },
    })
  })
  await expect
    .poll(() =>
      page.evaluate(() =>
        window.__codexVisualHost.acknowledgements.find(
          (entry) => entry.commandId === 'xlsx-sheet-print-scale-reopened-touch',
        ),
      ),
    )
    .toMatchObject({ ok: true, revision: 1 })
  const reopenedDownloadPromise = page.waitForEvent('download')
  await reopened.locator('button.qa-btn').first().click()
  const reopenedPath = await (await reopenedDownloadPromise).path()
  expect(reopenedPath).not.toBeNull()
  const reopenedZip = await JSZip.loadAsync(
    await import('node:fs/promises').then((fs) => fs.readFile(reopenedPath!)),
  )
  const reopenedSheetXml = await reopenedZip.file('xl/worksheets/sheet1.xml')!.async('text')
  expect(reopenedSheetXml).toMatch(/<pageSetup\b[^>]*\bscale="80"/)
  expect(reopenedSheetXml).not.toMatch(/<pageSetUpPr\b[^>]*\bfitToPage=/)
})

test('the XLSX print-gridline registry operation shares user undo and saved output', async ({
  page,
}) => {
  await page.goto('/?format=xlsx&width=720&height=900')
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)
  const editor = page.frameLocator('#editor-frame')
  await editor.locator('input[type="file"][accept*=".xlsx"]').setInputFiles({
    name: 'registry-print-gridlines.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: await xlsxWithA1('Print gridlines'),
  })

  await page.evaluate(() => {
    window.__codexVisualHost.commands.push({
      commandId: 'xlsx-sheet-print-gridlines',
      baseRevision: 0,
      operation: 'xlsx.sheet.set_print_gridlines',
      arguments: { sheet: 'Sheet1', enabled: true },
    })
  })
  await expect
    .poll(() =>
      page.evaluate(() =>
        window.__codexVisualHost.acknowledgements.find(
          (entry) => entry.commandId === 'xlsx-sheet-print-gridlines',
        ),
      ),
    )
    .toMatchObject({
      ok: true,
      revision: 1,
      output: { sheet: 'Sheet1', enabled: true },
    })

  await expect(editor.locator('button.qa-btn').nth(1)).toBeEnabled()
  await editor.locator('button.qa-btn').nth(1).click()
  await expect(editor.locator('button.qa-btn').nth(2)).toBeEnabled()
  await page.evaluate(() => {
    window.__codexVisualHost.commands.push({
      commandId: 'xlsx-sheet-print-gridlines-undo-persist',
      baseRevision: 1,
      operation: 'xlsx.cell.set_value',
      arguments: { sheet: 'Sheet1', address: 'B1', value: 'Undo persisted' },
    })
  })
  await expect
    .poll(() =>
      page.evaluate(() =>
        window.__codexVisualHost.acknowledgements.find(
          (entry) => entry.commandId === 'xlsx-sheet-print-gridlines-undo-persist',
        ),
      ),
    )
    .toMatchObject({ ok: true, revision: 2 })
  const undoDownloadPromise = page.waitForEvent('download')
  await editor.locator('button.qa-btn').first().click()
  const undoPath = await (await undoDownloadPromise).path()
  expect(undoPath).not.toBeNull()
  const undoZip = await JSZip.loadAsync(
    await import('node:fs/promises').then((fs) => fs.readFile(undoPath!)),
  )
  const undoSheetXml = await undoZip.file('xl/worksheets/sheet1.xml')!.async('text')
  expect(undoSheetXml).not.toMatch(/<printOptions\b[^>]*\bgridLines=/)

  await page.evaluate(() => {
    window.__codexVisualHost.commands.push({
      commandId: 'xlsx-sheet-print-gridlines-redo-cycle',
      baseRevision: 2,
      operation: 'xlsx.sheet.set_print_gridlines',
      arguments: { sheet: 'Sheet1', enabled: true },
    })
  })
  await expect
    .poll(() =>
      page.evaluate(() =>
        window.__codexVisualHost.acknowledgements.find(
          (entry) => entry.commandId === 'xlsx-sheet-print-gridlines-redo-cycle',
        ),
      ),
    )
    .toMatchObject({ ok: true, revision: 3 })
  await editor.locator('button.qa-btn').nth(1).click()
  await expect(editor.locator('button.qa-btn').nth(2)).toBeEnabled()
  await editor.locator('button.qa-btn').nth(2).click()
  const redoDownloadPromise = page.waitForEvent('download')
  await editor.locator('button.qa-btn').first().click()
  const redoPath = await (await redoDownloadPromise).path()
  expect(redoPath).not.toBeNull()
  const redoBytes = await import('node:fs/promises').then((fs) => fs.readFile(redoPath!))
  const redoZip = await JSZip.loadAsync(redoBytes)
  const redoSheetXml = await redoZip.file('xl/worksheets/sheet1.xml')!.async('text')
  expect(redoSheetXml).toMatch(/<printOptions\b[^>]*\bgridLines="1"/)

  await page.reload()
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)
  const reopened = page.frameLocator('#editor-frame')
  await reopened.locator('input[type="file"][accept*=".xlsx"]').setInputFiles({
    name: 'registry-print-gridlines-reopened.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: redoBytes,
  })
  await page.evaluate(() => {
    window.__codexVisualHost.commands.push({
      commandId: 'xlsx-sheet-print-gridlines-reopened-touch',
      baseRevision: 0,
      operation: 'xlsx.cell.set_value',
      arguments: { sheet: 'Sheet1', address: 'A1', value: 'Print gridlines' },
    })
  })
  await expect
    .poll(() =>
      page.evaluate(() =>
        window.__codexVisualHost.acknowledgements.find(
          (entry) => entry.commandId === 'xlsx-sheet-print-gridlines-reopened-touch',
        ),
      ),
    )
    .toMatchObject({ ok: true, revision: 1 })
  const reopenedDownloadPromise = page.waitForEvent('download')
  await reopened.locator('button.qa-btn').first().click()
  const reopenedPath = await (await reopenedDownloadPromise).path()
  expect(reopenedPath).not.toBeNull()
  const reopenedZip = await JSZip.loadAsync(
    await import('node:fs/promises').then((fs) => fs.readFile(reopenedPath!)),
  )
  const reopenedSheetXml = await reopenedZip.file('xl/worksheets/sheet1.xml')!.async('text')
  expect(reopenedSheetXml).toMatch(/<printOptions\b[^>]*\bgridLines="1"/)
})

test('the retained print-gridline Ribbon control shares user undo and saved output', async ({
  page,
}) => {
  await page.goto('/?format=xlsx&width=1332&height=900')
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)
  const editor = page.frameLocator('#editor-frame')
  await editor.locator('input[type="file"][accept*=".xlsx"]').setInputFiles({
    name: 'ribbon-print-gridlines.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: await xlsxWithA1('Print gridlines'),
  })

  await editor.getByRole('button', { name: '页面布局', exact: true }).click()
  const printGridlines = editor.locator('button[data-tip="打印网格线"]')
  await printGridlines.click()
  await expect(editor.locator('.workbook-status')).toContainText('将打印网格线')
  await expect(editor.locator('button.qa-btn').nth(1)).toBeEnabled()
  await editor.locator('button.qa-btn').nth(1).click()

  await page.evaluate(() => {
    window.__codexVisualHost.commands.push({
      commandId: 'xlsx-ribbon-print-gridlines-undo-persist',
      baseRevision: 0,
      operation: 'xlsx.cell.set_value',
      arguments: { sheet: 'Sheet1', address: 'B1', value: 'Undo persisted' },
    })
  })
  await expect
    .poll(() =>
      page.evaluate(() =>
        window.__codexVisualHost.acknowledgements.find(
          (entry) => entry.commandId === 'xlsx-ribbon-print-gridlines-undo-persist',
        ),
      ),
    )
    .toMatchObject({ ok: true, revision: 1 })

  const undoDownloadPromise = page.waitForEvent('download')
  await editor.locator('button.qa-btn').first().click()
  const undoPath = await (await undoDownloadPromise).path()
  expect(undoPath).not.toBeNull()
  const undoZip = await JSZip.loadAsync(
    await import('node:fs/promises').then((fs) => fs.readFile(undoPath!)),
  )
  const undoSheetXml = await undoZip.file('xl/worksheets/sheet1.xml')!.async('text')
  expect(undoSheetXml).not.toMatch(/<printOptions\b[^>]*\bgridLines=/)

  await printGridlines.click()
  await editor.locator('button.qa-btn').nth(1).click()
  await expect(editor.locator('button.qa-btn').nth(2)).toBeEnabled()
  await editor.locator('button.qa-btn').nth(2).click()
  const redoDownloadPromise = page.waitForEvent('download')
  await editor.locator('button.qa-btn').first().click()
  const redoPath = await (await redoDownloadPromise).path()
  expect(redoPath).not.toBeNull()
  const redoBytes = await import('node:fs/promises').then((fs) => fs.readFile(redoPath!))
  const redoZip = await JSZip.loadAsync(redoBytes)
  const redoSheetXml = await redoZip.file('xl/worksheets/sheet1.xml')!.async('text')
  expect(redoSheetXml).toMatch(/<printOptions\b[^>]*\bgridLines="1"/)

  await page.reload()
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)
  const reopened = page.frameLocator('#editor-frame')
  await reopened.locator('input[type="file"][accept*=".xlsx"]').setInputFiles({
    name: 'ribbon-print-gridlines-reopened.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: redoBytes,
  })
  await reopened.getByRole('button', { name: '页面布局', exact: true }).click()
  await expect(reopened.locator('button[data-tip="打印网格线"] i.check-box')).toHaveText('✓')
  await reopened.locator('button[data-tip="打印网格线"]').click()
  const disabledDownloadPromise = page.waitForEvent('download')
  await reopened.locator('button.qa-btn').first().click()
  const disabledPath = await (await disabledDownloadPromise).path()
  expect(disabledPath).not.toBeNull()
  const disabledZip = await JSZip.loadAsync(
    await import('node:fs/promises').then((fs) => fs.readFile(disabledPath!)),
  )
  const disabledSheetXml = await disabledZip.file('xl/worksheets/sheet1.xml')!.async('text')
  expect(disabledSheetXml).not.toMatch(/<printOptions\b[^>]*\bgridLines=/)
})

test('the XLSX print-heading registry operation shares user undo and saved output', async ({
  page,
}) => {
  await page.goto('/?format=xlsx&width=720&height=900')
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)
  const editor = page.frameLocator('#editor-frame')
  await editor.locator('input[type="file"][accept*=".xlsx"]').setInputFiles({
    name: 'registry-print-headings.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: await xlsxWithA1('Print headings'),
  })

  await page.evaluate(() => {
    window.__codexVisualHost.commands.push({
      commandId: 'xlsx-sheet-print-headings',
      baseRevision: 0,
      operation: 'xlsx.sheet.set_print_headings',
      arguments: { sheet: 'Sheet1', enabled: true },
    })
  })
  await expect
    .poll(() =>
      page.evaluate(() =>
        window.__codexVisualHost.acknowledgements.find(
          (entry) => entry.commandId === 'xlsx-sheet-print-headings',
        ),
      ),
    )
    .toMatchObject({
      ok: true,
      revision: 1,
      output: { sheet: 'Sheet1', enabled: true },
    })

  await expect(editor.locator('button.qa-btn').nth(1)).toBeEnabled()
  await editor.locator('button.qa-btn').nth(1).click()
  await expect(editor.locator('button.qa-btn').nth(2)).toBeEnabled()
  await page.evaluate(() => {
    window.__codexVisualHost.commands.push({
      commandId: 'xlsx-sheet-print-headings-undo-persist',
      baseRevision: 1,
      operation: 'xlsx.cell.set_value',
      arguments: { sheet: 'Sheet1', address: 'B1', value: 'Undo persisted' },
    })
  })
  await expect
    .poll(() =>
      page.evaluate(() =>
        window.__codexVisualHost.acknowledgements.find(
          (entry) => entry.commandId === 'xlsx-sheet-print-headings-undo-persist',
        ),
      ),
    )
    .toMatchObject({ ok: true, revision: 2 })
  const undoDownloadPromise = page.waitForEvent('download')
  await editor.locator('button.qa-btn').first().click()
  const undoPath = await (await undoDownloadPromise).path()
  expect(undoPath).not.toBeNull()
  const undoZip = await JSZip.loadAsync(
    await import('node:fs/promises').then((fs) => fs.readFile(undoPath!)),
  )
  const undoSheetXml = await undoZip.file('xl/worksheets/sheet1.xml')!.async('text')
  expect(undoSheetXml).not.toMatch(/<printOptions\b[^>]*\bheadings=/)

  await page.evaluate(() => {
    window.__codexVisualHost.commands.push({
      commandId: 'xlsx-sheet-print-headings-redo-cycle',
      baseRevision: 2,
      operation: 'xlsx.sheet.set_print_headings',
      arguments: { sheet: 'Sheet1', enabled: true },
    })
  })
  await expect
    .poll(() =>
      page.evaluate(() =>
        window.__codexVisualHost.acknowledgements.find(
          (entry) => entry.commandId === 'xlsx-sheet-print-headings-redo-cycle',
        ),
      ),
    )
    .toMatchObject({ ok: true, revision: 3 })
  await editor.locator('button.qa-btn').nth(1).click()
  await expect(editor.locator('button.qa-btn').nth(2)).toBeEnabled()
  await editor.locator('button.qa-btn').nth(2).click()
  const redoDownloadPromise = page.waitForEvent('download')
  await editor.locator('button.qa-btn').first().click()
  const redoPath = await (await redoDownloadPromise).path()
  expect(redoPath).not.toBeNull()
  const redoBytes = await import('node:fs/promises').then((fs) => fs.readFile(redoPath!))
  const redoZip = await JSZip.loadAsync(redoBytes)
  const redoSheetXml = await redoZip.file('xl/worksheets/sheet1.xml')!.async('text')
  expect(redoSheetXml).toMatch(/<printOptions\b[^>]*\bheadings="1"/)

  await page.reload()
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)
  const reopened = page.frameLocator('#editor-frame')
  await reopened.locator('input[type="file"][accept*=".xlsx"]').setInputFiles({
    name: 'registry-print-headings-reopened.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: redoBytes,
  })
  await page.evaluate(() => {
    window.__codexVisualHost.commands.push({
      commandId: 'xlsx-sheet-print-headings-reopened-touch',
      baseRevision: 0,
      operation: 'xlsx.cell.set_value',
      arguments: { sheet: 'Sheet1', address: 'A1', value: 'Print headings' },
    })
  })
  await expect
    .poll(() =>
      page.evaluate(() =>
        window.__codexVisualHost.acknowledgements.find(
          (entry) => entry.commandId === 'xlsx-sheet-print-headings-reopened-touch',
        ),
      ),
    )
    .toMatchObject({ ok: true, revision: 1 })
  const reopenedDownloadPromise = page.waitForEvent('download')
  await reopened.locator('button.qa-btn').first().click()
  const reopenedPath = await (await reopenedDownloadPromise).path()
  expect(reopenedPath).not.toBeNull()
  const reopenedZip = await JSZip.loadAsync(
    await import('node:fs/promises').then((fs) => fs.readFile(reopenedPath!)),
  )
  const reopenedSheetXml = await reopenedZip.file('xl/worksheets/sheet1.xml')!.async('text')
  expect(reopenedSheetXml).toMatch(/<printOptions\b[^>]*\bheadings="1"/)
})

test('the retained print-heading Ribbon control shares user undo and saved output', async ({
  page,
}) => {
  await page.goto('/?format=xlsx&width=1332&height=900')
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)
  const editor = page.frameLocator('#editor-frame')
  await editor.locator('input[type="file"][accept*=".xlsx"]').setInputFiles({
    name: 'ribbon-print-headings.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: await xlsxWithA1('Print headings'),
  })

  await editor.getByRole('button', { name: '页面布局', exact: true }).click()
  const printHeadings = editor.locator('button[data-tip="打印行号列标"]')
  await printHeadings.click()
  await expect(editor.locator('.workbook-status')).toContainText('将打印行号列标')
  await expect(editor.locator('button.qa-btn').nth(1)).toBeEnabled()
  await editor.locator('button.qa-btn').nth(1).click()

  await page.evaluate(() => {
    window.__codexVisualHost.commands.push({
      commandId: 'xlsx-ribbon-print-headings-undo-persist',
      baseRevision: 0,
      operation: 'xlsx.cell.set_value',
      arguments: { sheet: 'Sheet1', address: 'B1', value: 'Undo persisted' },
    })
  })
  await expect
    .poll(() =>
      page.evaluate(() =>
        window.__codexVisualHost.acknowledgements.find(
          (entry) => entry.commandId === 'xlsx-ribbon-print-headings-undo-persist',
        ),
      ),
    )
    .toMatchObject({ ok: true, revision: 1 })

  const undoDownloadPromise = page.waitForEvent('download')
  await editor.locator('button.qa-btn').first().click()
  const undoPath = await (await undoDownloadPromise).path()
  expect(undoPath).not.toBeNull()
  const undoZip = await JSZip.loadAsync(
    await import('node:fs/promises').then((fs) => fs.readFile(undoPath!)),
  )
  const undoSheetXml = await undoZip.file('xl/worksheets/sheet1.xml')!.async('text')
  expect(undoSheetXml).not.toMatch(/<printOptions\b[^>]*\bheadings=/)

  await printHeadings.click()
  await editor.locator('button.qa-btn').nth(1).click()
  await expect(editor.locator('button.qa-btn').nth(2)).toBeEnabled()
  await editor.locator('button.qa-btn').nth(2).click()
  const redoDownloadPromise = page.waitForEvent('download')
  await editor.locator('button.qa-btn').first().click()
  const redoPath = await (await redoDownloadPromise).path()
  expect(redoPath).not.toBeNull()
  const redoBytes = await import('node:fs/promises').then((fs) => fs.readFile(redoPath!))
  const redoZip = await JSZip.loadAsync(redoBytes)
  const redoSheetXml = await redoZip.file('xl/worksheets/sheet1.xml')!.async('text')
  expect(redoSheetXml).toMatch(/<printOptions\b[^>]*\bheadings="1"/)

  await page.reload()
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)
  const reopened = page.frameLocator('#editor-frame')
  await reopened.locator('input[type="file"][accept*=".xlsx"]').setInputFiles({
    name: 'ribbon-print-headings-reopened.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: redoBytes,
  })
  await reopened.getByRole('button', { name: '页面布局', exact: true }).click()
  await expect(reopened.locator('button[data-tip="打印行号列标"] i.check-box')).toHaveText('✓')
  await reopened.locator('button[data-tip="打印行号列标"]').click()
  const disabledDownloadPromise = page.waitForEvent('download')
  await reopened.locator('button.qa-btn').first().click()
  const disabledPath = await (await disabledDownloadPromise).path()
  expect(disabledPath).not.toBeNull()
  const disabledZip = await JSZip.loadAsync(
    await import('node:fs/promises').then((fs) => fs.readFile(disabledPath!)),
  )
  const disabledSheetXml = await disabledZip.file('xl/worksheets/sheet1.xml')!.async('text')
  expect(disabledSheetXml).not.toMatch(/<printOptions\b[^>]*\bheadings=/)
})

test('the XLSX print-area registry operation shares user undo and saved output', async ({
  page,
}) => {
  await page.goto('/?format=xlsx&width=720&height=900')
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)
  const editor = page.frameLocator('#editor-frame')
  await editor.locator('input[type="file"][accept*=".xlsx"]').setInputFiles({
    name: 'registry-print-area.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: await xlsxWithA1('Print area'),
  })

  await page.evaluate(() => {
    window.__codexVisualHost.commands.push({
      commandId: 'xlsx-sheet-print-area',
      baseRevision: 0,
      operation: 'xlsx.sheet.set_print_area',
      arguments: { sheet: 'Sheet1', range: 'B2:D8' },
    })
  })
  await expect
    .poll(() =>
      page.evaluate(() =>
        window.__codexVisualHost.acknowledgements.find(
          (entry) => entry.commandId === 'xlsx-sheet-print-area',
        ),
      ),
    )
    .toMatchObject({
      ok: true,
      revision: 1,
      output: { sheet: 'Sheet1', range: 'B2:D8' },
    })

  await expect(editor.locator('button.qa-btn').nth(1)).toBeEnabled()
  await editor.locator('button.qa-btn').nth(1).click()
  await expect(editor.locator('button.qa-btn').nth(2)).toBeEnabled()
  await page.evaluate(() => {
    window.__codexVisualHost.commands.push({
      commandId: 'xlsx-sheet-print-area-undo-persist',
      baseRevision: 1,
      operation: 'xlsx.cell.set_value',
      arguments: { sheet: 'Sheet1', address: 'B1', value: 'Undo persisted' },
    })
  })
  await expect
    .poll(() =>
      page.evaluate(() =>
        window.__codexVisualHost.acknowledgements.find(
          (entry) => entry.commandId === 'xlsx-sheet-print-area-undo-persist',
        ),
      ),
    )
    .toMatchObject({ ok: true, revision: 2 })
  const undoDownloadPromise = page.waitForEvent('download')
  await editor.locator('button.qa-btn').first().click()
  const undoPath = await (await undoDownloadPromise).path()
  expect(undoPath).not.toBeNull()
  const undoZip = await JSZip.loadAsync(
    await import('node:fs/promises').then((fs) => fs.readFile(undoPath!)),
  )
  const undoWorkbookXml = await undoZip.file('xl/workbook.xml')!.async('text')
  expect(undoWorkbookXml).not.toContain('_xlnm.Print_Area')

  await page.evaluate(() => {
    window.__codexVisualHost.commands.push({
      commandId: 'xlsx-sheet-print-area-redo-cycle',
      baseRevision: 2,
      operation: 'xlsx.sheet.set_print_area',
      arguments: { sheet: 'Sheet1', range: 'B2:D8' },
    })
  })
  await expect
    .poll(() =>
      page.evaluate(() =>
        window.__codexVisualHost.acknowledgements.find(
          (entry) => entry.commandId === 'xlsx-sheet-print-area-redo-cycle',
        ),
      ),
    )
    .toMatchObject({ ok: true, revision: 3 })
  await editor.locator('button.qa-btn').nth(1).click()
  await expect(editor.locator('button.qa-btn').nth(2)).toBeEnabled()
  await editor.locator('button.qa-btn').nth(2).click()
  const redoDownloadPromise = page.waitForEvent('download')
  await editor.locator('button.qa-btn').first().click()
  const redoPath = await (await redoDownloadPromise).path()
  expect(redoPath).not.toBeNull()
  const redoBytes = await import('node:fs/promises').then((fs) => fs.readFile(redoPath!))
  const redoZip = await JSZip.loadAsync(redoBytes)
  const redoWorkbookXml = await redoZip.file('xl/workbook.xml')!.async('text')
  expect(redoWorkbookXml).toContain('_xlnm.Print_Area')
  expect(redoWorkbookXml).toContain("'Sheet1'!$B$2:$D$8")

  await page.reload()
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)
  const reopened = page.frameLocator('#editor-frame')
  await reopened.locator('input[type="file"][accept*=".xlsx"]').setInputFiles({
    name: 'registry-print-area-reopened.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: redoBytes,
  })
  await reopened.getByRole('button', { name: '页面布局', exact: true }).click()
  await expect(
    reopened.locator('.ribbon-tool.large').filter({ hasText: '打印区域' }),
  ).toHaveAttribute('data-tip', /B2:D8/)

  await page.evaluate(() => {
    window.__codexVisualHost.commands.push({
      commandId: 'xlsx-sheet-print-area-clear',
      baseRevision: 0,
      operation: 'xlsx.sheet.set_print_area',
      arguments: { sheet: 'Sheet1', range: null },
    })
  })
  await expect
    .poll(() =>
      page.evaluate(() =>
        window.__codexVisualHost.acknowledgements.find(
          (entry) => entry.commandId === 'xlsx-sheet-print-area-clear',
        ),
      ),
    )
    .toMatchObject({
      ok: true,
      revision: 1,
      output: { sheet: 'Sheet1', range: null },
    })
  const clearedDownloadPromise = page.waitForEvent('download')
  await reopened.locator('button.qa-btn').first().click()
  const clearedPath = await (await clearedDownloadPromise).path()
  expect(clearedPath).not.toBeNull()
  const clearedZip = await JSZip.loadAsync(
    await import('node:fs/promises').then((fs) => fs.readFile(clearedPath!)),
  )
  const clearedWorkbookXml = await clearedZip.file('xl/workbook.xml')!.async('text')
  expect(clearedWorkbookXml).not.toContain('_xlnm.Print_Area')
})

test('the XLSX print-title registry operation shares user undo and saved output', async ({
  page,
}) => {
  await page.goto('/?format=xlsx&width=720&height=900')
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)
  const editor = page.frameLocator('#editor-frame')
  await editor.locator('input[type="file"][accept*=".xlsx"]').setInputFiles({
    name: 'registry-print-titles.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: await xlsxWithA1('Print titles'),
  })

  await page.evaluate(() => {
    window.__codexVisualHost.commands.push({
      commandId: 'xlsx-sheet-print-titles',
      baseRevision: 0,
      operation: 'xlsx.sheet.set_print_titles',
      arguments: { sheet: 'Sheet1', rows: '2:8' },
    })
  })
  await expect
    .poll(() =>
      page.evaluate(() =>
        window.__codexVisualHost.acknowledgements.find(
          (entry) => entry.commandId === 'xlsx-sheet-print-titles',
        ),
      ),
    )
    .toMatchObject({
      ok: true,
      revision: 1,
      output: { sheet: 'Sheet1', rows: '2:8' },
    })

  await expect(editor.locator('button.qa-btn').nth(1)).toBeEnabled()
  await editor.locator('button.qa-btn').nth(1).click()
  await expect(editor.locator('button.qa-btn').nth(2)).toBeEnabled()
  await page.evaluate(() => {
    window.__codexVisualHost.commands.push({
      commandId: 'xlsx-sheet-print-titles-undo-persist',
      baseRevision: 1,
      operation: 'xlsx.cell.set_value',
      arguments: { sheet: 'Sheet1', address: 'B1', value: 'Undo persisted' },
    })
  })
  await expect
    .poll(() =>
      page.evaluate(() =>
        window.__codexVisualHost.acknowledgements.find(
          (entry) => entry.commandId === 'xlsx-sheet-print-titles-undo-persist',
        ),
      ),
    )
    .toMatchObject({ ok: true, revision: 2 })
  const undoDownloadPromise = page.waitForEvent('download')
  await editor.locator('button.qa-btn').first().click()
  const undoPath = await (await undoDownloadPromise).path()
  expect(undoPath).not.toBeNull()
  const undoZip = await JSZip.loadAsync(
    await import('node:fs/promises').then((fs) => fs.readFile(undoPath!)),
  )
  const undoWorkbookXml = await undoZip.file('xl/workbook.xml')!.async('text')
  expect(undoWorkbookXml).not.toContain('_xlnm.Print_Titles')

  await page.evaluate(() => {
    window.__codexVisualHost.commands.push({
      commandId: 'xlsx-sheet-print-titles-redo-cycle',
      baseRevision: 2,
      operation: 'xlsx.sheet.set_print_titles',
      arguments: { sheet: 'Sheet1', rows: '2:8' },
    })
  })
  await expect
    .poll(() =>
      page.evaluate(() =>
        window.__codexVisualHost.acknowledgements.find(
          (entry) => entry.commandId === 'xlsx-sheet-print-titles-redo-cycle',
        ),
      ),
    )
    .toMatchObject({ ok: true, revision: 3 })
  await editor.locator('button.qa-btn').nth(1).click()
  await expect(editor.locator('button.qa-btn').nth(2)).toBeEnabled()
  await editor.locator('button.qa-btn').nth(2).click()
  const redoDownloadPromise = page.waitForEvent('download')
  await editor.locator('button.qa-btn').first().click()
  const redoPath = await (await redoDownloadPromise).path()
  expect(redoPath).not.toBeNull()
  const redoBytes = await import('node:fs/promises').then((fs) => fs.readFile(redoPath!))
  const redoZip = await JSZip.loadAsync(redoBytes)
  const redoWorkbookXml = await redoZip.file('xl/workbook.xml')!.async('text')
  expect(redoWorkbookXml).toContain('_xlnm.Print_Titles')
  expect(redoWorkbookXml).toContain("'Sheet1'!$2:$8")

  await page.reload()
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)
  const reopened = page.frameLocator('#editor-frame')
  await reopened.locator('input[type="file"][accept*=".xlsx"]').setInputFiles({
    name: 'registry-print-titles-reopened.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: redoBytes,
  })
  await reopened.getByRole('button', { name: '页面布局', exact: true }).click()
  await expect(
    reopened.locator('.ribbon-tool.large').filter({ hasText: '打印标题' }),
  ).toHaveAttribute('data-tip', /2:8/)

  await page.evaluate(() => {
    window.__codexVisualHost.commands.push({
      commandId: 'xlsx-sheet-print-titles-clear',
      baseRevision: 0,
      operation: 'xlsx.sheet.set_print_titles',
      arguments: { sheet: 'Sheet1', rows: null },
    })
  })
  await expect
    .poll(() =>
      page.evaluate(() =>
        window.__codexVisualHost.acknowledgements.find(
          (entry) => entry.commandId === 'xlsx-sheet-print-titles-clear',
        ),
      ),
    )
    .toMatchObject({
      ok: true,
      revision: 1,
      output: { sheet: 'Sheet1', rows: null },
    })
  const clearedDownloadPromise = page.waitForEvent('download')
  await reopened.locator('button.qa-btn').first().click()
  const clearedPath = await (await clearedDownloadPromise).path()
  expect(clearedPath).not.toBeNull()
  const clearedZip = await JSZip.loadAsync(
    await import('node:fs/promises').then((fs) => fs.readFile(clearedPath!)),
  )
  const clearedWorkbookXml = await clearedZip.file('xl/workbook.xml')!.async('text')
  expect(clearedWorkbookXml).not.toContain('_xlnm.Print_Titles')
})

test('the retained print-title Ribbon menu shares user undo and saved output', async ({ page }) => {
  await page.goto('/?format=xlsx&width=1332&height=900')
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)
  const editor = page.frameLocator('#editor-frame')
  await editor.locator('input[type="file"][accept*=".xlsx"]').setInputFiles({
    name: 'ribbon-print-titles.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: await xlsxWithA1('Print titles'),
  })

  await editor.getByRole('button', { name: '页面布局', exact: true }).click()
  await editor.getByRole('button', { name: '打印标题' }).click()
  await editor.getByRole('option', { name: '重复第 1 行' }).click()
  await expect(editor.locator('.workbook-status')).toContainText('第 1 行将在每页顶端重复')
  await expect(editor.locator('button.qa-btn').nth(1)).toBeEnabled()
  await editor.locator('button.qa-btn').nth(1).click()

  await page.evaluate(() => {
    window.__codexVisualHost.commands.push({
      commandId: 'xlsx-ribbon-print-titles-undo-persist',
      baseRevision: 0,
      operation: 'xlsx.cell.set_value',
      arguments: { sheet: 'Sheet1', address: 'B1', value: 'Undo persisted' },
    })
  })
  await expect
    .poll(() =>
      page.evaluate(() =>
        window.__codexVisualHost.acknowledgements.find(
          (entry) => entry.commandId === 'xlsx-ribbon-print-titles-undo-persist',
        ),
      ),
    )
    .toMatchObject({ ok: true, revision: 1 })
  const undoDownloadPromise = page.waitForEvent('download')
  await editor.locator('button.qa-btn').first().click()
  const undoPath = await (await undoDownloadPromise).path()
  expect(undoPath).not.toBeNull()
  const undoZip = await JSZip.loadAsync(
    await import('node:fs/promises').then((fs) => fs.readFile(undoPath!)),
  )
  const undoWorkbookXml = await undoZip.file('xl/workbook.xml')!.async('text')
  expect(undoWorkbookXml).not.toContain('_xlnm.Print_Titles')

  const nameBox = editor.getByRole('textbox', { name: 'Name Box' })
  await nameBox.fill('B2:D8')
  await nameBox.press('Enter')
  await expect
    .poll(() =>
      page.evaluate(() => {
        const selection = window.__codexVisualHost.lastPollArguments?.selection as
          { range?: string } | undefined
        return selection?.range
      }),
    )
    .toBe('B2:D8')
  await editor.getByRole('button', { name: '打印标题' }).click()
  await editor.getByRole('option', { name: '重复所选行' }).click()
  await expect(editor.locator('.workbook-status')).toContainText('第 2:8 行将在每页顶端重复')
  await editor.locator('button.qa-btn').nth(1).click()
  await expect(editor.locator('button.qa-btn').nth(2)).toBeEnabled()
  await editor.locator('button.qa-btn').nth(2).click()
  const redoDownloadPromise = page.waitForEvent('download')
  await editor.locator('button.qa-btn').first().click()
  const redoPath = await (await redoDownloadPromise).path()
  expect(redoPath).not.toBeNull()
  const redoBytes = await import('node:fs/promises').then((fs) => fs.readFile(redoPath!))
  const redoZip = await JSZip.loadAsync(redoBytes)
  const redoWorkbookXml = await redoZip.file('xl/workbook.xml')!.async('text')
  expect(redoWorkbookXml).toContain('_xlnm.Print_Titles')
  expect(redoWorkbookXml).toContain("'Sheet1'!$2:$8")

  await page.reload()
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)
  const reopened = page.frameLocator('#editor-frame')
  await reopened.locator('input[type="file"][accept*=".xlsx"]').setInputFiles({
    name: 'ribbon-print-titles-reopened.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: redoBytes,
  })
  await reopened.getByRole('button', { name: '页面布局', exact: true }).click()
  await expect(
    reopened.locator('.ribbon-tool.large').filter({ hasText: '打印标题' }),
  ).toHaveAttribute('data-tip', /2:8/)
  await reopened.getByRole('button', { name: '打印标题' }).click()
  await reopened.getByRole('option', { name: '清除打印标题' }).click()
  await expect(reopened.locator('.workbook-status')).toContainText('打印标题已清除')
  const clearedDownloadPromise = page.waitForEvent('download')
  await reopened.locator('button.qa-btn').first().click()
  const clearedPath = await (await clearedDownloadPromise).path()
  expect(clearedPath).not.toBeNull()
  const clearedZip = await JSZip.loadAsync(
    await import('node:fs/promises').then((fs) => fs.readFile(clearedPath!)),
  )
  const clearedWorkbookXml = await clearedZip.file('xl/workbook.xml')!.async('text')
  expect(clearedWorkbookXml).not.toContain('_xlnm.Print_Titles')
})

test('the retained print-area Ribbon menu shares user undo and saved output', async ({ page }) => {
  await page.goto('/?format=xlsx&width=1332&height=900')
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)
  const editor = page.frameLocator('#editor-frame')
  await editor.locator('input[type="file"][accept*=".xlsx"]').setInputFiles({
    name: 'ribbon-print-area.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: await xlsxWithA1('Print area'),
  })

  const nameBox = editor.getByRole('textbox', { name: 'Name Box' })
  await nameBox.fill('B2:D8')
  await nameBox.press('Enter')
  await expect
    .poll(() =>
      page.evaluate(() => {
        const selection = window.__codexVisualHost.lastPollArguments?.selection as
          { range?: string } | undefined
        return selection?.range
      }),
    )
    .toBe('B2:D8')
  await editor.getByRole('button', { name: '页面布局', exact: true }).click()
  await editor.getByRole('button', { name: '打印区域' }).click()
  await editor.getByRole('option', { name: '设置打印区域' }).click()
  await expect(editor.locator('.workbook-status')).toContainText('打印区域：B2:D8')
  await expect(editor.locator('button.qa-btn').nth(1)).toBeEnabled()
  await editor.locator('button.qa-btn').nth(1).click()

  await page.evaluate(() => {
    window.__codexVisualHost.commands.push({
      commandId: 'xlsx-ribbon-print-area-undo-persist',
      baseRevision: 0,
      operation: 'xlsx.cell.set_value',
      arguments: { sheet: 'Sheet1', address: 'B1', value: 'Undo persisted' },
    })
  })
  await expect
    .poll(() =>
      page.evaluate(() =>
        window.__codexVisualHost.acknowledgements.find(
          (entry) => entry.commandId === 'xlsx-ribbon-print-area-undo-persist',
        ),
      ),
    )
    .toMatchObject({ ok: true, revision: 1 })
  const undoDownloadPromise = page.waitForEvent('download')
  await editor.locator('button.qa-btn').first().click()
  const undoPath = await (await undoDownloadPromise).path()
  expect(undoPath).not.toBeNull()
  const undoZip = await JSZip.loadAsync(
    await import('node:fs/promises').then((fs) => fs.readFile(undoPath!)),
  )
  const undoWorkbookXml = await undoZip.file('xl/workbook.xml')!.async('text')
  expect(undoWorkbookXml).not.toContain('_xlnm.Print_Area')

  await nameBox.fill('B2:D8')
  await nameBox.press('Enter')
  await editor.getByRole('button', { name: '打印区域' }).click()
  await editor.getByRole('option', { name: '设置打印区域' }).click()
  await editor.locator('button.qa-btn').nth(1).click()
  await expect(editor.locator('button.qa-btn').nth(2)).toBeEnabled()
  await editor.locator('button.qa-btn').nth(2).click()
  const redoDownloadPromise = page.waitForEvent('download')
  await editor.locator('button.qa-btn').first().click()
  const redoPath = await (await redoDownloadPromise).path()
  expect(redoPath).not.toBeNull()
  const redoBytes = await import('node:fs/promises').then((fs) => fs.readFile(redoPath!))
  const redoZip = await JSZip.loadAsync(redoBytes)
  const redoWorkbookXml = await redoZip.file('xl/workbook.xml')!.async('text')
  expect(redoWorkbookXml).toContain('_xlnm.Print_Area')
  expect(redoWorkbookXml).toContain("'Sheet1'!$B$2:$D$8")

  await page.reload()
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)
  const reopened = page.frameLocator('#editor-frame')
  await reopened.locator('input[type="file"][accept*=".xlsx"]').setInputFiles({
    name: 'ribbon-print-area-reopened.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: redoBytes,
  })
  await reopened.getByRole('button', { name: '页面布局', exact: true }).click()
  await expect(
    reopened.locator('.ribbon-tool.large').filter({ hasText: '打印区域' }),
  ).toHaveAttribute('data-tip', /B2:D8/)
  await reopened.getByRole('button', { name: '打印区域' }).click()
  await reopened.getByRole('option', { name: '取消打印区域' }).click()
  await expect(reopened.locator('.workbook-status')).toContainText('打印区域已清除')
  const clearedDownloadPromise = page.waitForEvent('download')
  await reopened.locator('button.qa-btn').first().click()
  const clearedPath = await (await clearedDownloadPromise).path()
  expect(clearedPath).not.toBeNull()
  const clearedZip = await JSZip.loadAsync(
    await import('node:fs/promises').then((fs) => fs.readFile(clearedPath!)),
  )
  const clearedWorkbookXml = await clearedZip.file('xl/workbook.xml')!.async('text')
  expect(clearedWorkbookXml).not.toContain('_xlnm.Print_Area')
})

test('a user can insert a row and the shifted cells survive save and reopen', async ({ page }) => {
  await page.goto('/?format=xlsx&width=720&height=900')
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)
  const editor = page.frameLocator('#editor-frame')
  await editor.locator('input[type="file"][accept*=".xlsx"]').setInputFiles({
    name: 'insert-row.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: await xlsxWithA1('Shift me'),
  })
  await expect(editor.locator('.workbook-status')).toContainText('insert-row.xlsx')

  await editor.getByRole('button', { name: '插入行' }).click()

  const downloadPromise = page.waitForEvent('download')
  await editor.locator('button.qa-btn').first().click()
  const savedPath = await (await downloadPromise).path()
  expect(savedPath).not.toBeNull()
  const saved = await import('node:fs/promises').then((fs) => fs.readFile(savedPath!))
  const zip = await JSZip.loadAsync(saved)
  const sheetXml = await zip.file('xl/worksheets/sheet1.xml')!.async('text')
  expect(sheetXml).toContain('r="A2"')
  expect(sheetXml).toContain('Shift me')

  await page.reload()
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)
  const reopened = page.frameLocator('#editor-frame')
  await reopened.locator('input[type="file"][accept*=".xlsx"]').setInputFiles({
    name: 'insert-row-reopened.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: saved,
  })
  await reopened
    .locator('canvas[id^="univer-sheet-main-canvas_"]')
    .click({ position: { x: 70, y: 55 } })
  await expect
    .poll(() =>
      page.evaluate(() => {
        const selection = window.__codexVisualHost.lastPollArguments?.selection as
          { activeCell?: { address?: string; value?: unknown } } | undefined
        return selection?.activeCell
      }),
    )
    .toMatchObject({ address: 'A2', value: 'Shift me' })
})

test('undo reverses a user row insertion in the subsequently saved XLSX', async ({ page }) => {
  await page.goto('/?format=xlsx&width=720&height=900')
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)
  const editor = page.frameLocator('#editor-frame')
  await editor.locator('input[type="file"][accept*=".xlsx"]').setInputFiles({
    name: 'undo-insert-row.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: await xlsxWithA1('Stay in A1'),
  })
  await expect(editor.locator('.workbook-status')).toContainText('undo-insert-row.xlsx')

  await editor.getByRole('button', { name: '插入行' }).click()
  await editor.locator('button.qa-btn').nth(1).click()
  await expect
    .poll(() =>
      page.evaluate(() => {
        const selection = window.__codexVisualHost.lastPollArguments?.selection as
          { activeCell?: { value?: unknown } } | undefined
        return selection?.activeCell?.value
      }),
    )
    .toBe('Stay in A1')

  const downloadPromise = page.waitForEvent('download')
  await page.evaluate(() => {
    window.__codexVisualHost.commands.push({
      commandId: 'xlsx-save-after-undo',
      baseRevision: 0,
      operation: 'xlsx.document.save',
      arguments: {},
    })
  })
  const savedPath = await (await downloadPromise).path()
  expect(savedPath).not.toBeNull()
  const zip = await JSZip.loadAsync(
    await import('node:fs/promises').then((fs) => fs.readFile(savedPath!)),
  )
  const sheetXml = await zip.file('xl/worksheets/sheet1.xml')!.async('text')
  expect(sheetXml).toContain('r="A1"')
  expect(sheetXml).not.toContain('r="A2"')
  expect(sheetXml).toContain('Stay in A1')
})

test('an MCP row insertion updates the shared grid and survives save and reopen', async ({
  page,
}) => {
  await page.goto('/?format=xlsx&width=720&height=900')
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)
  const editor = page.frameLocator('#editor-frame')
  await editor.locator('input[type="file"][accept*=".xlsx"]').setInputFiles({
    name: 'mcp-insert-rows.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: await xlsxWithA1('Agent shifted'),
  })
  await expect(editor.locator('.workbook-status')).toContainText('mcp-insert-rows.xlsx')

  await page.evaluate(() => {
    window.__codexVisualHost.commands.push({
      commandId: 'xlsx-insert-rows',
      baseRevision: 0,
      operation: 'xlsx.row.insert',
      arguments: { sheet: 'Sheet1', row: 1, count: 2 },
    })
  })
  await expect
    .poll(() =>
      page.evaluate(() =>
        window.__codexVisualHost.acknowledgements.find(
          (entry) => entry.commandId === 'xlsx-insert-rows',
        ),
      ),
    )
    .toMatchObject({ ok: true, revision: 1, output: { sheet: 'Sheet1', row: 1, count: 2 } })

  const downloadPromise = page.waitForEvent('download')
  await editor.locator('button.qa-btn').first().click()
  const savedPath = await (await downloadPromise).path()
  expect(savedPath).not.toBeNull()
  const saved = await import('node:fs/promises').then((fs) => fs.readFile(savedPath!))
  const zip = await JSZip.loadAsync(saved)
  const sheetXml = await zip.file('xl/worksheets/sheet1.xml')!.async('text')
  expect(sheetXml).toContain('r="A3"')

  await page.reload()
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)
  const reopened = page.frameLocator('#editor-frame')
  await reopened.locator('input[type="file"][accept*=".xlsx"]').setInputFiles({
    name: 'mcp-insert-rows-reopened.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: saved,
  })
  await reopened
    .locator('canvas[id^="univer-sheet-main-canvas_"]')
    .click({ position: { x: 70, y: 75 } })
  await expect
    .poll(() =>
      page.evaluate(() => {
        const selection = window.__codexVisualHost.lastPollArguments?.selection as
          { activeCell?: { address?: string; value?: unknown } } | undefined
        return selection?.activeCell
      }),
    )
    .toMatchObject({ address: 'A3', value: 'Agent shifted' })
})

test('an MCP column insertion updates the shared grid and survives save and reopen', async ({
  page,
}) => {
  await page.goto('/?format=xlsx&width=720&height=900')
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)
  const editor = page.frameLocator('#editor-frame')
  await editor.locator('input[type="file"][accept*=".xlsx"]').setInputFiles({
    name: 'mcp-insert-columns.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: await xlsxWithA1('Move to B1'),
  })
  await expect(editor.locator('.workbook-status')).toContainText('mcp-insert-columns.xlsx')

  await page.evaluate(() => {
    window.__codexVisualHost.commands.push({
      commandId: 'xlsx-insert-columns',
      baseRevision: 0,
      operation: 'xlsx.column.insert',
      arguments: { sheet: 'Sheet1', column: 'A', count: 1 },
    })
  })
  await expect
    .poll(() =>
      page.evaluate(() =>
        window.__codexVisualHost.acknowledgements.find(
          (entry) => entry.commandId === 'xlsx-insert-columns',
        ),
      ),
    )
    .toMatchObject({
      ok: true,
      revision: 1,
      output: { sheet: 'Sheet1', column: 'A', count: 1 },
    })

  const downloadPromise = page.waitForEvent('download')
  await editor.locator('button.qa-btn').first().click()
  const savedPath = await (await downloadPromise).path()
  expect(savedPath).not.toBeNull()
  const saved = await import('node:fs/promises').then((fs) => fs.readFile(savedPath!))
  const zip = await JSZip.loadAsync(saved)
  const sheetXml = await zip.file('xl/worksheets/sheet1.xml')!.async('text')
  expect(sheetXml).toContain('r="B1"')

  await page.reload()
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)
  const reopened = page.frameLocator('#editor-frame')
  await reopened.locator('input[type="file"][accept*=".xlsx"]').setInputFiles({
    name: 'mcp-insert-columns-reopened.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: saved,
  })
  await reopened
    .locator('canvas[id^="univer-sheet-main-canvas_"]')
    .click({ position: { x: 70, y: 35 } })
  await page.keyboard.press('ArrowRight')
  await expect
    .poll(() =>
      page.evaluate(() => {
        const selection = window.__codexVisualHost.lastPollArguments?.selection as
          { activeCell?: { address?: string; value?: unknown } } | undefined
        return selection?.activeCell
      }),
    )
    .toMatchObject({ address: 'B1', value: 'Move to B1' })
})

test('MCP row and column deletion share the community journal and saved package', async ({
  page,
}) => {
  await page.goto('/?format=xlsx&width=720&height=900')
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)
  const editor = page.frameLocator('#editor-frame')
  await editor.locator('input[type="file"][accept*=".xlsx"]').setInputFiles({
    name: 'mcp-delete-axes.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: await xlsxWithA1('Before'),
  })

  await page.evaluate(() => {
    window.__codexVisualHost.commands.push(
      {
        commandId: 'xlsx-seed-delete-axes',
        baseRevision: 0,
        operation: 'xlsx.range.set_values',
        arguments: {
          sheet: 'Sheet1',
          range: 'A1:B2',
          values: [
            ['Drop row', 'Drop row'],
            ['Drop column', 'Keep'],
          ],
        },
      },
      {
        commandId: 'xlsx-delete-row',
        baseRevision: 1,
        operation: 'xlsx.row.delete',
        arguments: { sheet: 'Sheet1', row: 1, count: 1 },
      },
      {
        commandId: 'xlsx-delete-column',
        baseRevision: 2,
        operation: 'xlsx.column.delete',
        arguments: { sheet: 'Sheet1', column: 'A', count: 1 },
      },
    )
  })
  await expect
    .poll(() =>
      page.evaluate(() =>
        window.__codexVisualHost.acknowledgements.find(
          (entry) => entry.commandId === 'xlsx-delete-column',
        ),
      ),
    )
    .toMatchObject({ ok: true, revision: 3 })

  const downloadPromise = page.waitForEvent('download')
  await editor.locator('button.qa-btn').first().click()
  const savedPath = await (await downloadPromise).path()
  expect(savedPath).not.toBeNull()
  const zip = await JSZip.loadAsync(
    await import('node:fs/promises').then((fs) => fs.readFile(savedPath!)),
  )
  const sheetXml = await zip.file('xl/worksheets/sheet1.xml')!.async('text')
  expect(sheetXml).toContain('r="A1"')
  expect(sheetXml).toContain('Keep')
  expect(sheetXml).not.toContain('Drop row')
  expect(sheetXml).not.toContain('Drop column')
})

test('an MCP sheet insertion creates a valid worksheet package that reopens', async ({ page }) => {
  await page.goto('/?format=xlsx&width=720&height=900')
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)
  const editor = page.frameLocator('#editor-frame')
  await editor.locator('input[type="file"][accept*=".xlsx"]').setInputFiles({
    name: 'add-sheet.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: await xlsxWithA1('Existing'),
  })

  await page.evaluate(() => {
    window.__codexVisualHost.commands.push({
      commandId: 'xlsx-add-sheet',
      baseRevision: 0,
      operation: 'xlsx.sheet.add',
      arguments: { name: 'Budget' },
    })
  })
  await expect
    .poll(() =>
      page.evaluate(() =>
        window.__codexVisualHost.acknowledgements.find(
          (entry) => entry.commandId === 'xlsx-add-sheet',
        ),
      ),
    )
    .toMatchObject({ ok: true, revision: 1, output: { name: 'Budget' } })

  const downloadPromise = page.waitForEvent('download')
  await editor.locator('button.qa-btn').first().click()
  const savedPath = await (await downloadPromise).path()
  expect(savedPath).not.toBeNull()
  const saved = await import('node:fs/promises').then((fs) => fs.readFile(savedPath!))
  const zip = await JSZip.loadAsync(saved)
  expect(await zip.file('xl/workbook.xml')!.async('text')).toContain('name="Budget"')
  expect(await zip.file('xl/_rels/workbook.xml.rels')!.async('text')).toContain(
    'worksheets/sheet2.xml',
  )
  expect(await zip.file('[Content_Types].xml')!.async('text')).toContain(
    '/xl/worksheets/sheet2.xml',
  )
  expect(zip.file('xl/worksheets/sheet2.xml')).not.toBeNull()

  await page.reload()
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)
  const reopened = page.frameLocator('#editor-frame')
  await reopened.locator('input[type="file"][accept*=".xlsx"]').setInputFiles({
    name: 'add-sheet-reopened.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: saved,
  })
  await expect
    .poll(() =>
      page.evaluate(() => {
        const selection = window.__codexVisualHost.lastPollArguments?.selection as
          { sheets?: string[] } | undefined
        return selection?.sheets
      }),
    )
    .toEqual(['Sheet1', 'Budget'])
})

test('the community sheet-tab add button shares the saved workbook and undo stack', async ({
  page,
}) => {
  await page.goto('/?format=xlsx&width=720&height=900')
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)
  const editor = page.frameLocator('#editor-frame')
  await editor.locator('input[type="file"][accept*=".xlsx"]').setInputFiles({
    name: 'user-add-sheet.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: await xlsxWithA1('Existing'),
  })

  await editor.locator('[data-u-comp="sheet-bar-append-button"]').first().click()
  await expect(editor.locator('[aria-label="Sheet tabs"] [role="tab"]')).toHaveCount(2)
  await editor.locator('button.qa-btn').nth(1).click()
  await expect(editor.locator('[aria-label="Sheet tabs"] [role="tab"]')).toHaveCount(1)
  await editor.locator('button.qa-btn').nth(2).click()
  await expect(editor.locator('[aria-label="Sheet tabs"] [role="tab"]')).toHaveCount(2)

  const downloadPromise = page.waitForEvent('download')
  await editor.locator('button.qa-btn').first().click()
  const savedPath = await (await downloadPromise).path()
  expect(savedPath).not.toBeNull()
  const zip = await JSZip.loadAsync(
    await import('node:fs/promises').then((fs) => fs.readFile(savedPath!)),
  )
  expect(await zip.file('xl/workbook.xml')!.async('text')).toContain('name="工作表1"')
})

test('an MCP sheet rename updates the tab and survives save and reopen', async ({ page }) => {
  await page.goto('/?format=xlsx&width=720&height=900')
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)
  const editor = page.frameLocator('#editor-frame')
  await editor.locator('input[type="file"][accept*=".xlsx"]').setInputFiles({
    name: 'rename-sheet.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: await xlsxWithA1('Existing'),
  })

  await page.evaluate(() => {
    window.__codexVisualHost.commands.push({
      commandId: 'xlsx-rename-sheet',
      baseRevision: 0,
      operation: 'xlsx.sheet.rename',
      arguments: { sheet: 'Sheet1', name: 'Summary' },
    })
  })
  await expect
    .poll(() =>
      page.evaluate(() =>
        window.__codexVisualHost.acknowledgements.find(
          (entry) => entry.commandId === 'xlsx-rename-sheet',
        ),
      ),
    )
    .toMatchObject({ ok: true, revision: 1, output: { sheet: 'Sheet1', name: 'Summary' } })
  await expect(editor.getByRole('tab', { name: 'Summary' })).toBeVisible()

  const downloadPromise = page.waitForEvent('download')
  await editor.locator('button.qa-btn').first().click()
  const savedPath = await (await downloadPromise).path()
  expect(savedPath).not.toBeNull()
  const saved = await import('node:fs/promises').then((fs) => fs.readFile(savedPath!))
  const zip = await JSZip.loadAsync(saved)
  const workbookXml = await zip.file('xl/workbook.xml')!.async('text')
  expect(workbookXml).toContain('name="Summary"')
  expect(workbookXml).not.toContain('name="Sheet1"')

  await page.reload()
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)
  const reopened = page.frameLocator('#editor-frame')
  await reopened.locator('input[type="file"][accept*=".xlsx"]').setInputFiles({
    name: 'rename-sheet-reopened.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: saved,
  })
  await expect(reopened.getByRole('tab', { name: 'Summary' })).toBeVisible()
})

test('an MCP sheet deletion removes its package parts and survives reopen', async ({ page }) => {
  await page.goto('/?format=xlsx&width=720&height=900')
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)
  const editor = page.frameLocator('#editor-frame')
  await editor.locator('input[type="file"][accept*=".xlsx"]').setInputFiles({
    name: 'delete-sheet.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: await xlsxWithA1('Remove me'),
  })
  await page.evaluate(() => {
    window.__codexVisualHost.commands.push({
      commandId: 'xlsx-add-survivor',
      baseRevision: 0,
      operation: 'xlsx.sheet.add',
      arguments: { name: 'Survivor' },
    })
  })
  await expect
    .poll(() =>
      page.evaluate(() =>
        window.__codexVisualHost.acknowledgements.find(
          (entry) => entry.commandId === 'xlsx-add-survivor',
        ),
      ),
    )
    .toMatchObject({ ok: true, revision: 1 })

  await page.evaluate(() => {
    window.__codexVisualHost.commands.push({
      commandId: 'xlsx-delete-sheet',
      baseRevision: 1,
      operation: 'xlsx.sheet.delete',
      arguments: { sheet: 'Sheet1' },
    })
  })
  await expect
    .poll(() =>
      page.evaluate(() =>
        window.__codexVisualHost.acknowledgements.find(
          (entry) => entry.commandId === 'xlsx-delete-sheet',
        ),
      ),
    )
    .toMatchObject({ ok: true, revision: 2, output: { sheet: 'Sheet1' } })
  await expect(editor.locator('[aria-label="Sheet tabs"] [role="tab"]')).toHaveCount(1)
  await expect(editor.getByRole('tab', { name: 'Survivor' })).toBeVisible()

  const downloadPromise = page.waitForEvent('download')
  await editor.locator('button.qa-btn').first().click()
  const savedPath = await (await downloadPromise).path()
  expect(savedPath).not.toBeNull()
  const saved = await import('node:fs/promises').then((fs) => fs.readFile(savedPath!))
  const zip = await JSZip.loadAsync(saved)
  const workbookXml = await zip.file('xl/workbook.xml')!.async('text')
  expect(workbookXml).toContain('name="Survivor"')
  expect(workbookXml).not.toContain('name="Sheet1"')
  expect(zip.file('xl/worksheets/sheet1.xml')).toBeNull()

  await page.reload()
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)
  const reopened = page.frameLocator('#editor-frame')
  await reopened.locator('input[type="file"][accept*=".xlsx"]').setInputFiles({
    name: 'delete-sheet-reopened.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: saved,
  })
  await expect(reopened.getByRole('tab', { name: 'Survivor' })).toBeVisible()
})

test('an MCP sheet move persists the tab order and reopens in that order', async ({ page }) => {
  await page.goto('/?format=xlsx&width=720&height=900')
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)
  const editor = page.frameLocator('#editor-frame')
  await editor.locator('input[type="file"][accept*=".xlsx"]').setInputFiles({
    name: 'move-sheet.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: await xlsxWithA1('First'),
  })
  await page.evaluate(() => {
    window.__codexVisualHost.commands.push({
      commandId: 'xlsx-add-second',
      baseRevision: 0,
      operation: 'xlsx.sheet.add',
      arguments: { name: 'Second' },
    })
  })
  await expect
    .poll(() =>
      page.evaluate(() =>
        window.__codexVisualHost.acknowledgements.find(
          (entry) => entry.commandId === 'xlsx-add-second',
        ),
      ),
    )
    .toMatchObject({ ok: true, revision: 1 })

  await page.evaluate(() => {
    window.__codexVisualHost.commands.push({
      commandId: 'xlsx-move-sheet',
      baseRevision: 1,
      operation: 'xlsx.sheet.move',
      arguments: { sheet: 'Second', position: 1 },
    })
  })
  await expect
    .poll(() =>
      page.evaluate(() =>
        window.__codexVisualHost.acknowledgements.find(
          (entry) => entry.commandId === 'xlsx-move-sheet',
        ),
      ),
    )
    .toMatchObject({
      ok: true,
      revision: 2,
      output: { sheet: 'Second', position: 1 },
    })

  const downloadPromise = page.waitForEvent('download')
  await editor.locator('button.qa-btn').first().click()
  const savedPath = await (await downloadPromise).path()
  expect(savedPath).not.toBeNull()
  const saved = await import('node:fs/promises').then((fs) => fs.readFile(savedPath!))
  const zip = await JSZip.loadAsync(saved)
  const workbookXml = await zip.file('xl/workbook.xml')!.async('text')
  expect(workbookXml.indexOf('name="Second"')).toBeLessThan(workbookXml.indexOf('name="Sheet1"'))

  await page.reload()
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)
  const reopened = page.frameLocator('#editor-frame')
  await reopened.locator('input[type="file"][accept*=".xlsx"]').setInputFiles({
    name: 'move-sheet-reopened.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: saved,
  })
  await expect
    .poll(() => reopened.locator('[aria-label="Sheet tabs"] [role="tab"]').allTextContents())
    .toEqual(['Second', 'Sheet1'])
})

test('a user-defined name survives XLSX save and reopen through the community Name Manager', async ({
  page,
}) => {
  await page.goto('/?format=xlsx&width=1332&height=900')
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)
  const editor = page.frameLocator('#editor-frame')
  await editor.locator('input[type="file"][accept*=".xlsx"]').setInputFiles({
    name: 'defined-name.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: await xlsxWithA1('Named value'),
  })

  await editor.getByRole('button', { name: '公式', exact: true }).click()
  await editor.getByRole('button', { name: '名称管理器' }).click()
  const dialog = editor.getByRole('dialog', { name: '名称管理器' })
  await dialog.getByLabel('名称').fill('NamedValue')
  await dialog.getByLabel('引用位置').fill('=Sheet1!$A$1')
  await dialog.getByRole('button', { name: '添加' }).click()
  await expect(dialog.getByRole('cell', { name: 'NamedValue' })).toBeVisible()
  await dialog.getByRole('button', { name: '关闭' }).click()

  const downloadPromise = page.waitForEvent('download')
  await editor.locator('button.qa-btn').first().click()
  const savedPath = await (await downloadPromise).path()
  expect(savedPath).not.toBeNull()
  const saved = await import('node:fs/promises').then((fs) => fs.readFile(savedPath!))
  const zip = await JSZip.loadAsync(saved)
  await expect(zip.file('xl/workbook.xml')!.async('text')).resolves.toContain(
    '<definedName name="NamedValue">Sheet1!$A$1</definedName>',
  )

  await page.reload()
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)
  const reopened = page.frameLocator('#editor-frame')
  await reopened.locator('input[type="file"][accept*=".xlsx"]').setInputFiles({
    name: 'defined-name-reopened.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: saved,
  })
  await reopened.getByRole('button', { name: '公式', exact: true }).click()
  await reopened.getByRole('button', { name: '名称管理器' }).click()
  await expect(
    reopened.getByRole('dialog', { name: '名称管理器' }).getByRole('cell', { name: 'NamedValue' }),
  ).toBeVisible()
})

test('the community Insert Function dialog writes a formula that survives XLSX save', async ({
  page,
}) => {
  await page.goto('/?format=xlsx&width=1332&height=900')
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)
  const editor = page.frameLocator('#editor-frame')
  await editor.locator('input[type="file"][accept*=".xlsx"]').setInputFiles({
    name: 'insert-function.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: await xlsxWithA1('Before'),
  })

  await editor.getByRole('button', { name: '公式', exact: true }).click()
  await editor.getByRole('button', { name: '插入函数' }).click()
  const dialog = editor.getByRole('dialog', { name: 'Insert Function' })
  await dialog.getByLabel('公式').fill('=SUM(1,2)')
  await dialog.getByRole('button', { name: '插入' }).click()
  await expect(dialog).toBeHidden()

  const downloadPromise = page.waitForEvent('download')
  await editor.locator('button.qa-btn').first().click()
  const savedPath = await (await downloadPromise).path()
  expect(savedPath).not.toBeNull()
  const zip = await JSZip.loadAsync(
    await import('node:fs/promises').then((fs) => fs.readFile(savedPath!)),
  )
  const sheetXml = await zip.file('xl/worksheets/sheet1.xml')!.async('text')
  expect(sheetXml).toContain('<f>SUM(1,2)</f>')
})

test('community Page Layout and Header & Footer edits survive XLSX save and reopen', async ({
  page,
}) => {
  await page.goto('/?format=xlsx&width=1332&height=900')
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)
  const editor = page.frameLocator('#editor-frame')
  await editor.locator('input[type="file"][accept*=".xlsx"]').setInputFiles({
    name: 'page-layout.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: await xlsxWithA1('Printable'),
  })

  await editor.getByRole('button', { name: '页面布局', exact: true }).click()
  await editor.getByRole('button', { name: '页边距' }).click()
  await editor.getByRole('option', { name: '宽', exact: true }).click()
  await expect(editor.locator('.workbook-status')).toContainText('页边距：宽')
  await editor.getByRole('button', { name: '纸张方向' }).click()
  await editor.getByRole('option', { name: '横向' }).click()
  await expect(editor.locator('.workbook-status')).toContainText('横向')
  await editor.getByRole('button', { name: '纸张大小' }).click()
  await editor.getByRole('option', { name: 'A4', exact: true }).click()
  await expect(editor.locator('.workbook-status')).toContainText('纸张大小：A4')
  await editor.getByRole('button', { name: 'Fit to width' }).click()
  await editor
    .getByRole('listbox', { name: 'Fit to width' })
    .getByRole('option', { name: '2 页', exact: true })
    .click()
  await expect(editor.locator('.workbook-status')).toContainText('调整宽度：2 页')
  await page.keyboard.press('Escape')
  await editor.getByRole('button', { name: 'Fit to height' }).click()
  await editor
    .getByRole('listbox', { name: 'Fit to height' })
    .getByRole('option', { name: '3 页', exact: true })
    .click()
  await expect(editor.locator('.workbook-status')).toContainText('调整高度：3 页')

  await editor.getByRole('button', { name: '插入', exact: true }).click()
  await editor.getByRole('button', { name: '页眉和页脚' }).click()
  const dialog = editor.getByRole('dialog', { name: '页眉页脚' })
  await dialog.getByLabel('页眉 · 中').fill('Quarterly Report')
  await dialog.getByLabel('页脚 · 右').fill('&P / &N')
  await dialog.getByRole('button', { name: '确定' }).click()
  await expect(editor.locator('.workbook-status')).toContainText('页眉页脚已更新')

  const downloadPromise = page.waitForEvent('download')
  await editor.locator('button.qa-btn').first().click()
  const savedPath = await (await downloadPromise).path()
  expect(savedPath).not.toBeNull()
  const saved = await import('node:fs/promises').then((fs) => fs.readFile(savedPath!))
  const zip = await JSZip.loadAsync(saved)
  const sheetXml = await zip.file('xl/worksheets/sheet1.xml')!.async('text')
  expect(sheetXml).toMatch(
    /<pageMargins\b(?=[^>]*\bleft="1")(?=[^>]*\bright="1")(?=[^>]*\btop="1")(?=[^>]*\bbottom="1")(?=[^>]*\bheader="0\.5")(?=[^>]*\bfooter="0\.5")/,
  )
  expect(sheetXml).toContain('<pageSetup orientation="landscape"')
  expect(sheetXml).toMatch(/<pageSetup\b[^>]*\bpaperSize="9"/)
  expect(sheetXml).toMatch(/<pageSetup\b(?=[^>]*\bfitToWidth="2")(?=[^>]*\bfitToHeight="3")/)
  expect(sheetXml).toMatch(/<pageSetUpPr\b[^>]*\bfitToPage="1"/)
  expect(sheetXml).toContain('<oddHeader>&amp;CQuarterly Report</oddHeader>')
  expect(sheetXml).toContain('<oddFooter>&amp;R&amp;P / &amp;N</oddFooter>')

  await page.reload()
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)
  const reopened = page.frameLocator('#editor-frame')
  await reopened.locator('input[type="file"][accept*=".xlsx"]').setInputFiles({
    name: 'page-layout-reopened.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: saved,
  })
  await reopened.getByRole('button', { name: '页面布局', exact: true }).click()
  await expect(
    reopened.locator('.ribbon-tool.large').filter({ hasText: '页边距' }),
  ).toHaveAttribute('data-tip', /宽/)
  await expect(
    reopened.locator('.ribbon-tool.large').filter({ hasText: '纸张方向' }),
  ).toHaveAttribute('data-tip', /横向/)
  await expect(reopened.getByRole('button', { name: 'Fit to width' })).toContainText('2 页')
  await expect(reopened.getByRole('button', { name: 'Fit to height' })).toContainText('3 页')
  await reopened.getByRole('button', { name: '插入', exact: true }).click()
  await reopened.getByRole('button', { name: '页眉和页脚' }).click()
  const reopenedDialog = reopened.getByRole('dialog', { name: '页眉页脚' })
  await expect(reopenedDialog.getByLabel('页眉 · 中')).toHaveValue('Quarterly Report')
  await expect(reopenedDialog.getByLabel('页脚 · 右')).toHaveValue('&P / &N')
})
