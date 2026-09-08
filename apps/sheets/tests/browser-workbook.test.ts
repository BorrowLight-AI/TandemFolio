import JSZip from 'jszip'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { blankXlsxBuffer } from '../src/gateway/csv-import'
import { buildChartXml } from '../src/gateway/xlsx-drawing-add'
import { openBrowserWorkbook } from '../src/host/browser-workbook'
import { BrowserWorkbookDesktopApi } from '../src/renderer/browser-desktop-api'
import type { WorkbookSaveRequest } from '../src/shared/desktop-api'
import { buildEditFixture } from './fixture-builder'

async function fixture(): Promise<Uint8Array> {
  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    '<Types><Override PartName="/xl/workbook.xml"/><Override PartName="/custom/unknown.xml"/></Types>',
  )
  zip.file(
    'xl/workbook.xml',
    '<workbook><sheets><sheet name="Budget" sheetId="1" r:id="rId1"/></sheets></workbook>',
  )
  zip.file(
    'xl/_rels/workbook.xml.rels',
    '<Relationships><Relationship Target="worksheets/sheet1.xml" Id="rId1"/></Relationships>',
  )
  zip.file(
    'xl/worksheets/sheet1.xml',
    '<worksheet><sheetData><row r="1"><c r="A1" s="3" t="inlineStr"><is><t>Revenue</t></is></c>' +
      '<c r="B1"><f>SUM(B2:B3)</f><v>30</v></c></row><row r="2"><c r="B2"><v>10</v></c></row>' +
      '</sheetData><mergeCells count="1"><mergeCell ref="A1:A2"/></mergeCells></worksheet>',
  )
  zip.file(
    'xl/styles.xml',
    '<styleSheet><fonts count="1"><font/></fonts><fills count="1"><fill/></fills>' +
      '<borders count="1"><border/></borders><cellXfs count="4">' +
      '<xf/><xf/><xf/><xf/></cellXfs></styleSheet>',
  )
  zip.file('xl/charts/chart1.xml', '<chart><title>Keep me</title></chart>')
  zip.file('xl/pivotTables/pivotTable1.xml', '<pivotTableDefinition name="KeepPivot"/>')
  zip.file('custom/unknown.xml', '<futureFeature value="opaque"/>')
  return zip.generateAsync({ type: 'uint8array' })
}

async function themeFixture(): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(await fixture())
  const rels = await zip.file('xl/_rels/workbook.xml.rels')!.async('text')
  zip.file(
    'xl/_rels/workbook.xml.rels',
    rels.replace(
      '</Relationships>',
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/></Relationships>',
    ),
  )
  zip.file(
    'xl/theme/theme1.xml',
    '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office"><a:themeElements>' +
      '<a:clrScheme name="Office"><a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>' +
      '<a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>' +
      '<a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2>' +
      '<a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4>' +
      '<a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6>' +
      '<a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme>' +
      '<a:fontScheme name="Office"><a:majorFont><a:latin typeface="Calibri Light"/></a:majorFont>' +
      '<a:minorFont><a:latin typeface="Calibri"/></a:minorFont></a:fontScheme></a:themeElements></a:theme>',
  )
  return zip.generateAsync({ type: 'uint8array' })
}

function inMemoryFileHandle(name: string) {
  let persisted = new Uint8Array()
  const writable = {
    write: vi.fn(async (data: ArrayBuffer | ArrayBufferView) => {
      const bytes =
        data instanceof ArrayBuffer
          ? new Uint8Array(data)
          : new Uint8Array(data.buffer, data.byteOffset, data.byteLength)
      persisted = bytes.slice()
    }),
    close: vi.fn(async () => undefined),
  }
  const handle = {
    kind: 'file' as const,
    name,
    createWritable: vi.fn(async () => writable as unknown as FileSystemWritableFileStream),
    getFile: vi.fn(
      async () =>
        ({
          name,
          lastModified: 1,
          arrayBuffer: async () => persisted.buffer.slice(0),
        }) as File,
    ),
  } as FileSystemFileHandle
  return { handle, writable, bytes: () => persisted }
}

function cellSaveRequest(
  file: Awaited<ReturnType<BrowserWorkbookDesktopApi['openBuffer']>>,
  value: number,
): WorkbookSaveRequest {
  return {
    sessionId: file.sessionId,
    mode: 'save',
    edits: [
      {
        sheetId: file.sheets[0]!.id,
        row: 1,
        column: 1,
        value,
        writeValue: true,
      },
    ],
    structuralOps: [],
    chartEdits: [],
    visualEdits: [],
    visualAdditions: [],
    tableAdditions: [],
    pivotAdditions: [],
    sheetOps: [],
    sheetOrder: [],
    filterStates: [],
    hyperlinkEdits: [],
    cfStates: [],
    dvStates: [],
    pageSetupStates: [],
    noteStates: [],
    formulaValues: [],
    pivotCacheRefreshPaths: [],
    pivotRefreshUpdates: [],
    sheetProtections: [],
    sparklineAdditions: [],
    definedNamesState: null,
  }
}

describe('browser XLSX workbook', () => {
  afterEach(() => vi.unstubAllGlobals())

  it('repairs legacy blank workbooks before saving native styles and themes', async () => {
    const legacy = await JSZip.loadAsync(await blankXlsxBuffer())
    legacy.remove('xl/styles.xml')
    legacy.remove('xl/theme/theme1.xml')
    legacy.file(
      'xl/_rels/workbook.xml.rels',
      (await legacy.file('xl/_rels/workbook.xml.rels')!.async('text'))
        .replace(/<Relationship\b[^>]*Type="[^"]*\/styles"[^>]*\/>/, '')
        .replace(/<Relationship\b[^>]*Type="[^"]*\/theme"[^>]*\/>/, ''),
    )
    legacy.file(
      '[Content_Types].xml',
      (await legacy.file('[Content_Types].xml')!.async('text'))
        .replace(/<Override\b[^>]*PartName="\/xl\/styles\.xml"[^>]*\/>/, '')
        .replace(/<Override\b[^>]*PartName="\/xl\/theme\/theme1\.xml"[^>]*\/>/, ''),
    )

    const workbook = await openBrowserWorkbook(
      await legacy.generateAsync({ type: 'uint8array' }),
      'legacy-blank.xlsx',
    )
    workbook.setRangeStyle('Sheet1', 'A1:B1', { bold: true })
    workbook.applyTheme({ fonts: { name: 'Office', major: 'Aptos Display', minor: 'Aptos' } })

    const saved = await JSZip.loadAsync(await workbook.save())
    expect(await saved.file('xl/styles.xml')!.async('text')).toContain('<b/>')
    expect(await saved.file('xl/theme/theme1.xml')!.async('text')).toContain('Aptos Display')
    expect(await saved.file('xl/_rels/workbook.xml.rels')!.async('text')).toContain(
      '/relationships/styles',
    )
    expect(await saved.file('xl/_rels/workbook.xml.rels')!.async('text')).toContain(
      '/relationships/theme',
    )
    expect(await saved.file('[Content_Types].xml')!.async('text')).toContain(
      'PartName="/xl/styles.xml"',
    )
    expect(await saved.file('[Content_Types].xml')!.async('text')).toContain(
      'PartName="/xl/theme/theme1.xml"',
    )
  })

  it('hydrates rich shared strings including subscript and superscript', async () => {
    vi.stubGlobal('window', {})
    const zip = await JSZip.loadAsync(await buildEditFixture())
    zip.file(
      'xl/sharedStrings.xml',
      '<sst><si><r><rPr><rFont val="Aptos"/><sz val="12"/><b/></rPr><t>H</t></r>' +
        '<r><rPr><vertAlign val="subscript"/></rPr><t>2</t></r>' +
        '<r><rPr><color rgb="FF123456"/><vertAlign val="superscript"/></rPr><t>O</t></r>' +
        '</si></sst>',
    )
    const api = new BrowserWorkbookDesktopApi(async () => null)
    const file = await api.openBuffer(await zip.generateAsync({ type: 'arraybuffer' }), 'rich.xlsx')
    const range = await api.readWorkbookRange({
      sessionId: file.sessionId,
      sheetId: file.sheets[0]!.id,
      range: { startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 },
    })

    expect(range.cells[0]).toMatchObject({
      value: 'H2O',
      rich: [
        expect.objectContaining({ text: 'H', family: 'Aptos', size: 12, bold: true }),
        expect.objectContaining({ text: '2', vertAlign: 'subscript' }),
        expect.objectContaining({ text: 'O', color: '#123456', vertAlign: 'superscript' }),
      ],
    })
  })

  it('exposes native shrink-to-fit styles from the workbook stylesheet', async () => {
    vi.stubGlobal('window', {})
    const zip = await JSZip.loadAsync(await fixture())
    zip.file(
      'xl/styles.xml',
      '<styleSheet><fonts count="2"><font/><font><name val="Aptos"/><sz val="12"/><b/><color rgb="FF123456"/></font></fonts>' +
        '<fills count="1"><fill/></fills><borders count="1"><border/></borders><cellXfs count="4">' +
        '<xf/><xf/><xf/><xf fontId="1" applyFont="1" applyAlignment="1"><alignment shrinkToFit="1" horizontal="right" textRotation="135"/></xf>' +
        '</cellXfs></styleSheet>',
    )
    const source = await zip.generateAsync({ type: 'arraybuffer' })
    const api = new BrowserWorkbookDesktopApi(async () => null)

    const file = await api.openBuffer(source, 'shrink.xlsx')

    expect(file.styles[3]).toMatchObject({
      fontFamily: 'Aptos',
      fontSize: 12,
      bold: true,
      fontColor: '#123456',
      horizontalAlignment: 'right',
      shrinkToFit: true,
      textRotation: 135,
    })
  })

  it('reads row and column default style indices from worksheet XML', async () => {
    vi.stubGlobal('window', {})
    const zip = await JSZip.loadAsync(await fixture())
    const sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('text')
    zip.file(
      'xl/worksheets/sheet1.xml',
      sheet
        .replace('<sheetData>', '<cols><col min="1" max="2" width="12" style="2"/></cols><sheetData>')
        .replace('<row r="1">', '<row r="1" s="1" customFormat="1">'),
    )
    const api = new BrowserWorkbookDesktopApi(async () => null)
    const file = await api.openBuffer(await zip.generateAsync({ type: 'arraybuffer' }), 'styles.xlsx')
    expect(file.sheets[0]!.columnWidths[0]).toMatchObject({ styleIndex: 2 })
    const range = await api.readWorkbookRange({
      sessionId: file.sessionId,
      sheetId: file.sheets[0]!.id,
      range: { startRow: 0, endRow: 2, startColumn: 0, endColumn: 2 },
    })
    expect(range.rows[0]).toMatchObject({ row: 0, styleIndex: 1 })
  })

  it('preserves sheetFormatPr default and base column widths', async () => {
    vi.stubGlobal('window', {})
    const zip = await JSZip.loadAsync(await fixture())
    const sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('text')
    zip.file(
      'xl/worksheets/sheet1.xml',
      sheet.replace('<sheetData>', '<sheetFormatPr defaultRowHeight="18" baseColWidth="10" defaultColWidth="12.5"/><sheetData>'),
    )
    const api = new BrowserWorkbookDesktopApi(async () => null)
    const file = await api.openBuffer(await zip.generateAsync({ type: 'arraybuffer' }), 'widths.xlsx')

    expect(file.sheets[0]).toMatchObject({
      defaultRowHeight: 18,
      defaultColumnWidth: 12.5,
      baseColumnWidth: 10,
    })
  })

  it('preserves row customHeight separately from cached auto heights', async () => {
    vi.stubGlobal('window', {})
    const zip = await JSZip.loadAsync(await fixture())
    const sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('text')
    zip.file(
      'xl/worksheets/sheet1.xml',
      sheet
        .replace('<row r="1">', '<row r="1" ht="30">')
        .replace('<row r="2">', '<row r="2" ht="30" customHeight="1">'),
    )
    const api = new BrowserWorkbookDesktopApi(async () => null)
    const file = await api.openBuffer(await zip.generateAsync({ type: 'arraybuffer' }), 'rows.xlsx')
    const range = await api.readWorkbookRange({
      sessionId: file.sessionId,
      sheetId: file.sheets[0]!.id,
      range: { startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 },
    })

    expect(range.rows).toEqual([
      expect.objectContaining({ row: 0, height: 30 }),
      expect.objectContaining({ row: 1, height: 30, customHeight: true }),
    ])
    expect(range.rows[0]).not.toHaveProperty('customHeight')
  })

  it('hydrates native conditional-format rules and differential styles', async () => {
    vi.stubGlobal('window', {})
    const zip = await JSZip.loadAsync(await fixture())
    const sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('text')
    zip.file(
      'xl/worksheets/sheet1.xml',
      sheet.replace(
        '</worksheet>',
        '<conditionalFormatting sqref="B2:B10"><cfRule type="colorScale" priority="1">' +
          '<colorScale><cfvo type="num" val="10"/><cfvo type="num" val="0"/>' +
          '<color rgb="FFF8696B"/><color rgb="FF63BE7B"/></colorScale></cfRule></conditionalFormatting>' +
          '<conditionalFormatting sqref="A1:A3"><cfRule type="cellIs" dxfId="0" priority="2" operator="lessThan">' +
          '<formula>5</formula></cfRule></conditionalFormatting></worksheet>',
      ),
    )
    const styles = await zip.file('xl/styles.xml')!.async('text')
    zip.file(
      'xl/styles.xml',
      styles.replace(
        '</styleSheet>',
        '<dxfs count="1"><dxf><font><b/><color rgb="FFFF0000"/></font></dxf></dxfs></styleSheet>',
      ),
    )
    const api = new BrowserWorkbookDesktopApi(async () => null)
    const file = await api.openBuffer(await zip.generateAsync({ type: 'arraybuffer' }), 'cf.xlsx')

    expect(file.dxfStyles[0]).toMatchObject({ bold: true, fontColor: '#FF0000' })
    const range = await api.readWorkbookRange({
      sessionId: file.sessionId,
      sheetId: file.sheets[0]!.id,
      range: { startRow: 0, endRow: 9, startColumn: 0, endColumn: 1 },
    })
    expect(range.conditionalRules).toEqual([
      expect.objectContaining({
        ruleType: 'colorScale',
        priority: 1,
        cfvos: [
          { kind: 'num', value: '10' },
          { kind: 'num', value: '0' },
        ],
        colors: ['#F8696B', '#63BE7B'],
      }),
      expect.objectContaining({
        ruleType: 'cellIs',
        operator: 'lessThan',
        dxfIndex: 0,
        formulas: ['5'],
      }),
    ])
  })

  it('reads, rewrites, saves, and reopens the native workbook theme', async () => {
    const workbook = await openBrowserWorkbook(await themeFixture(), 'theme.xlsx')
    expect(workbook.theme()).toMatchObject({
      colors: {
        name: 'Office',
        values: expect.arrayContaining(['#FFFFFF', '#000000', '#E7E6E6', '#44546A']),
      },
      fonts: { name: 'Office', major: 'Calibri Light', minor: 'Calibri' },
    })

    workbook.applyTheme({
      colors: {
        name: 'Forest',
        values: [
          '#FFFFFF', '#1E2B20', '#E9F2EB', '#375E43', '#217346', '#4EA72E',
          '#92D050', '#FFC000', '#3E8E8B', '#70AD47', '#217346', '#954F72',
        ],
      },
      fonts: { name: 'Candara', major: 'Candara', minor: 'Candara' },
    })
    const reopened = await openBrowserWorkbook(await workbook.save(), 'theme.xlsx')
    expect(reopened.theme()).toMatchObject({
      colors: { name: 'Forest', values: expect.arrayContaining(['#217346', '#4EA72E']) },
      fonts: { name: 'Candara', major: 'Candara', minor: 'Candara' },
    })
  })

  it('persists a theme-only Desktop API save and returns the reopened theme state', async () => {
    vi.stubGlobal('window', {})
    const memoryFile = inMemoryFileHandle('theme.xlsx')
    const api = new BrowserWorkbookDesktopApi(async () => null)
    const file = await api.openBuffer((await themeFixture()).buffer as ArrayBuffer, 'theme.xlsx', memoryFile.handle)
    expect(file.themeFonts).toEqual({ major: 'Calibri Light', minor: 'Calibri' })

    const result = await api.saveWorkbookEdits({
      ...cellSaveRequest(file, 10),
      edits: [],
      themeState: { fonts: { name: 'Georgia', major: 'Georgia', minor: 'Georgia' } },
    })
    expect(result.canceled).toBe(false)
    if (result.canceled) return
    expect(result.file.themeFonts).toEqual({ major: 'Georgia', minor: 'Georgia' })
    const reopened = await openBrowserWorkbook(memoryFile.bytes(), 'theme.xlsx')
    expect(reopened.theme()?.fonts).toEqual({
      name: 'Georgia',
      major: 'Georgia',
      minor: 'Georgia',
    })
  })

  it('saves and reopens an explicit worksheet tab color', async () => {
    const workbook = await openBrowserWorkbook(await fixture(), 'budget.xlsx')
    workbook.applyPageSetup('Budget', { tabColor: '#112233' })
    const saved = await workbook.save()
    const archive = await JSZip.loadAsync(saved)
    expect(await archive.file('xl/worksheets/sheet1.xml')!.async('text')).toContain(
      '<sheetPr><tabColor rgb="FF112233"/></sheetPr>',
    )
    const reopened = await openBrowserWorkbook(saved, 'budget.xlsx')
    expect(reopened.pageSetup('Budget').tabColor).toBe('#112233')
  })

  it('writes, discovers, reads, and refresh-flags native pivot definitions', async () => {
    const workbook = await openBrowserWorkbook(await buildEditFixture(), 'pivot.xlsx')
    await workbook.applyPivots(
      [
        {
          sheetName: 'Data',
          sourceSheetName: 'Data',
          sourceArea: { startRow: 0, startColumn: 0, endRow: 2, endColumn: 2 },
          location: { startRow: 0, startColumn: 5, endRow: 4, endColumn: 6 },
          name: 'Pivot1',
          fieldNames: ['Region', 'Product', 'Amount'],
          rowFieldIndices: [0],
          rowItems: ['East', 'South', 'North'],
          values: [{ fieldIndex: 2, agg: 'sum' }],
        },
      ],
      [],
      [],
    )

    const reopened = await openBrowserWorkbook(await workbook.save(), 'pivot.xlsx')
    expect(reopened.sheets[0]?.pivotRanges).toEqual([
      { startRow: 0, startColumn: 5, endRow: 4, endColumn: 6 },
    ])
    expect(reopened.sheets[0]?.pivotTables).toEqual([
      {
        path: 'xl/pivotTables/pivotTable1.xml',
        cachePath: 'xl/pivotCache/pivotCacheDefinition1.xml',
        outputRef: 'F1:G5',
        styled: true,
        firstDataRow: 1,
        firstDataCol: 1,
        rowGrandTotals: true,
        rowKinds: 'dddg',
        headerFill: '#DAE3F3',
        headerBold: true,
        subheadingBold: true,
        subheading2Bold: true,
        subtotalBold: true,
        totalRowFill: '#DAE3F3',
        totalRowBold: true,
      },
    ])
    const pivot = reopened.sheets[0]!.pivotTables[0]!
    await expect(reopened.readPivotDefinition(pivot.path, pivot.cachePath!)).resolves.toMatchObject(
      {
        outputRef: 'F1:G5',
        sourceSheet: 'Data',
        sourceRef: 'A1:C3',
        rowFields: [0],
        dataFields: [{ field: 2, subtotal: 'sum' }],
        unsupported: [],
      },
    )

    await reopened.applyPivots(
      [],
      [pivot.cachePath!],
      [
        {
          cachePath: pivot.cachePath!,
          pivotPath: pivot.path,
          sheetName: 'Data',
          newOutputRef: pivot.outputRef,
          memberHiddenItems: [{ field: 0, hiddenItems: [1] }],
        },
      ],
    )
    const refreshedZip = await JSZip.loadAsync(await reopened.save())
    expect(
      await refreshedZip.file('xl/pivotCache/pivotCacheDefinition1.xml')!.async('text'),
    ).toContain('refreshOnLoad="1"')
    const filtered = await openBrowserWorkbook(
      await refreshedZip.generateAsync({ type: 'uint8array' }),
      'pivot.xlsx',
    )
    const filteredDefinition = await filtered.readPivotDefinition(pivot.path, pivot.cachePath!)
    expect(filteredDefinition.fieldItems[0]?.map((item) => item.hidden)).toEqual([
      false,
      true,
      false,
    ])
  })

  it('reads context and patches cells while preserving formulas, styles, merges, charts, pivots, and unknown parts', async () => {
    const source = await fixture()
    const workbook = await openBrowserWorkbook(source, 'budget.xlsx')

    expect(workbook.sheets[0]).toMatchObject({
      name: 'Budget',
      merges: ['A1:A2'],
    })
    expect(workbook.cell('Budget', 'A1')).toMatchObject({ value: 'Revenue', styleIndex: 3 })
    expect(workbook.cell('Budget', 'B1')).toMatchObject({ value: 30, formula: 'SUM(B2:B3)' })

    workbook.setCellValue('Budget', 'B2', 25)
    const saved = await workbook.save()
    const before = await JSZip.loadAsync(source)
    const after = await JSZip.loadAsync(saved)

    expect(await after.file('xl/worksheets/sheet1.xml')!.async('text')).toContain(
      '<c r="B2"><v>25</v></c>',
    )
    expect(await after.file('xl/worksheets/sheet1.xml')!.async('text')).toContain(
      '<f>SUM(B2:B3)</f>',
    )
    expect(await after.file('xl/worksheets/sheet1.xml')!.async('text')).toContain(
      '<mergeCell ref="A1:A2"/>',
    )
    for (const path of [
      'xl/charts/chart1.xml',
      'xl/pivotTables/pivotTable1.xml',
      'custom/unknown.xml',
    ]) {
      expect(await after.file(path)!.async('uint8array')).toEqual(
        await before.file(path)!.async('uint8array'),
      )
    }
  })

  it('applies a rectangular value matrix as one workbook mutation', async () => {
    const workbook = await openBrowserWorkbook(await fixture(), 'budget.xlsx')
    expect(
      workbook.setRangeValues('Budget', 'C2:D3', [
        [1, 2],
        [3, 4],
      ]),
    ).toBe(4)
    expect(workbook.cell('Budget', 'D3')?.value).toBe(4)
  })

  it('writes and reopens workbook- and sheet-scoped defined names', async () => {
    const workbook = await openBrowserWorkbook(await fixture(), 'budget.xlsx')
    workbook.replaceDefinedNames([
      { name: 'Revenue', formula: 'Budget!$B$2:$B$9' },
      { name: 'LocalRate', formula: 'Budget!$D$2', sheetIndex: 0 },
    ])

    const reopened = await openBrowserWorkbook(await workbook.save(), 'budget.xlsx')
    expect(reopened.definedNames()).toEqual([
      { name: 'Revenue', formula: 'Budget!$B$2:$B$9' },
      { name: 'LocalRate', formula: 'Budget!$D$2', sheetIndex: 0 },
    ])

    reopened.replaceDefinedNames([{ name: 'ForecastRevenue', formula: 'Budget!$C$2:$C$9' }])
    const reduced = await openBrowserWorkbook(await reopened.save(), 'budget.xlsx')
    expect(reduced.definedNames()).toEqual([
      { name: 'ForecastRevenue', formula: 'Budget!$C$2:$C$9' },
    ])
  })

  it('reopens an external hyperlink through the browser workbook boundary', async () => {
    const workbook = await openBrowserWorkbook(await fixture(), 'budget.xlsx')
    workbook.applyHyperlinks('Budget', [
      { row: 0, column: 0, target: 'https://example.com/report' },
    ])

    const reopened = await openBrowserWorkbook(await workbook.save(), 'budget.xlsx')

    expect(reopened.sheets[0]?.hyperlinks.get('0:0')).toBe('https://example.com/report')
  })

  it('removes a reopened hyperlink through the browser workbook boundary', async () => {
    const workbook = await openBrowserWorkbook(await fixture(), 'budget.xlsx')
    workbook.applyHyperlinks('Budget', [
      { row: 0, column: 0, target: 'https://example.com/report' },
    ])
    const linked = await openBrowserWorkbook(await workbook.save(), 'budget.xlsx')

    linked.applyHyperlinks('Budget', [{ row: 0, column: 0, target: null }])
    const reopened = await openBrowserWorkbook(await linked.save(), 'budget.xlsx')

    expect(reopened.sheets[0]?.hyperlinks.has('0:0')).toBe(false)
  })

  it('writes, reopens, and removes legacy notes through the browser workbook boundary', async () => {
    const workbook = await openBrowserWorkbook(await fixture(), 'budget.xlsx')

    await workbook.applyNotes('Budget', [
      { row: 3, column: 2, author: 'Agent', text: 'Check this total.' },
    ])
    const notedBytes = await workbook.save()
    const notedZip = await JSZip.loadAsync(notedBytes)
    expect(await notedZip.file('xl/comments1.xml')!.async('text')).toContain(
      '<comment ref="C4" authorId="0"><text><t xml:space="preserve">Check this total.</t>',
    )
    expect(await notedZip.file('xl/drawings/vmlDrawing1.vml')!.async('text')).toContain(
      '<x:Row>3</x:Row><x:Column>2</x:Column>',
    )

    const reopened = await openBrowserWorkbook(notedBytes, 'budget.xlsx')
    expect(reopened.sheets[0]?.comments).toEqual([
      { row: 3, column: 2, author: 'Agent', text: 'Check this total.' },
    ])

    await reopened.applyNotes('Budget', [])
    const clearedBytes = await reopened.save()
    const clearedZip = await JSZip.loadAsync(clearedBytes)
    expect(clearedZip.file('xl/comments1.xml')).toBeNull()
    expect(clearedZip.file('xl/drawings/vmlDrawing1.vml')).toBeNull()
    expect((await openBrowserWorkbook(clearedBytes, 'budget.xlsx')).sheets[0]?.comments).toEqual([])
  })

  it('sets and reopens passwordless sheet protection through the browser boundary', async () => {
    const workbook = await openBrowserWorkbook(await fixture(), 'budget.xlsx')

    workbook.setSheetProtection('Budget', true)
    const protectedWorkbook = await openBrowserWorkbook(await workbook.save(), 'budget.xlsx')
    expect(protectedWorkbook.sheets[0]?.sheetProtection).toEqual({
      protected: true,
      hasPassword: false,
    })

    protectedWorkbook.setSheetProtection('Budget', false)
    const reopened = await openBrowserWorkbook(await protectedWorkbook.save(), 'budget.xlsx')
    expect(reopened.sheets[0]?.sheetProtection).toBeNull()
  })

  it('sets and reopens passwordless workbook structure protection', async () => {
    const workbook = await openBrowserWorkbook(await fixture(), 'budget.xlsx')
    expect(workbook.workbookProtection()).toBeNull()
    workbook.setWorkbookProtection(true)
    const protectedWorkbook = await openBrowserWorkbook(await workbook.save(), 'budget.xlsx')
    expect(protectedWorkbook.workbookProtection()).toEqual({
      lockStructure: true,
      hasPassword: false,
    })
    protectedWorkbook.setWorkbookProtection(false)
    const reopened = await openBrowserWorkbook(await protectedWorkbook.save(), 'budget.xlsx')
    expect(reopened.workbookProtection()).toBeNull()
  })

  it('reports and preserves the workbook 1904 date system', async () => {
    const zip = await JSZip.loadAsync(await fixture())
    const workbookXml = await zip.file('xl/workbook.xml')!.async('text')
    zip.file(
      'xl/workbook.xml',
      workbookXml.replace('<workbook>', '<workbook><workbookPr date1904="1"/>'),
    )
    const workbook = await openBrowserWorkbook(
      await zip.generateAsync({ type: 'uint8array' }),
      'date-1904.xlsx',
    )

    expect(workbook.date1904()).toBe(true)
    expect((await openBrowserWorkbook(await workbook.save(), 'date-1904.xlsx')).date1904()).toBe(
      true,
    )

    const api = new BrowserWorkbookDesktopApi(async () => null)
    const file = await api.openBuffer(await workbook.save(), 'date-1904.xlsx')
    expect(file.date1904).toBe(true)
  })

  it('reads, replaces, and reopens manual page breaks', async () => {
    const zip = await JSZip.loadAsync(await fixture())
    const sheetXml = await zip.file('xl/worksheets/sheet1.xml')!.async('text')
    zip.file(
      'xl/worksheets/sheet1.xml',
      sheetXml.replace(
        '</worksheet>',
        '<rowBreaks count="1" manualBreakCount="1"><brk id="5" max="16383" man="1"/></rowBreaks>' +
          '<colBreaks count="1" manualBreakCount="1"><brk id="2" max="1048575" man="1"/></colBreaks></worksheet>',
      ),
    )
    const workbook = await openBrowserWorkbook(
      await zip.generateAsync({ type: 'uint8array' }),
      'breaks.xlsx',
    )
    expect(workbook.pageSetup('Budget')).toMatchObject({ rowBreaks: [5], colBreaks: [2] })

    workbook.applyPageSetup('Budget', { rowBreaks: [7], colBreaks: [] })
    expect((await openBrowserWorkbook(await workbook.save(), 'breaks.xlsx')).pageSetup('Budget')).toMatchObject(
      { rowBreaks: [7], colBreaks: [] },
    )
  })

  it('reads exact saved print settings and page-specific header/footer variants', async () => {
    const zip = await JSZip.loadAsync(await fixture())
    const sheetXml = await zip.file('xl/worksheets/sheet1.xml')!.async('text')
    zip.file(
      'xl/worksheets/sheet1.xml',
      sheetXml.replace(
        '</worksheet>',
        '<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>' +
          '<printOptions gridLines="1" headings="1"/>' +
          '<pageMargins left="0.98" right="0.98" top="5" bottom="0.79" header="0" footer="0.51"/>' +
          '<pageSetup orientation="landscape" paperSize="1" scale="65"/>' +
          '<headerFooter differentOddEven="1" differentFirst="1" scaleWithDoc="0">' +
          '<oddFooter>&amp;CPage &amp;P of &amp;N</oddFooter><evenHeader>&amp;CEven</evenHeader>' +
          '<firstFooter>&amp;RFirst &amp;D</firstFooter></headerFooter></worksheet>',
      ),
    )

    const workbook = await openBrowserWorkbook(
      await zip.generateAsync({ type: 'uint8array' }),
      'print-settings.xlsx',
    )

    expect(workbook.pagePrintSettings('Budget')).toEqual({
      orientation: 'landscape',
      paperSize: 1,
      scale: 65,
      fitToPage: true,
      margins: { left: 0.98, right: 0.98, top: 5, bottom: 0.79, header: 0, footer: 0.51 },
      printGridlines: true,
      printHeadings: true,
      oddFooter: '&CPage &P of &N',
      differentOddEven: true,
      differentFirst: true,
      headerFooterFixedSize: true,
      evenHeader: '&CEven',
      firstFooter: '&RFirst &D',
    })

    const api = new BrowserWorkbookDesktopApi(async () => null)
    const file = await api.openBuffer(await workbook.save(), 'print-settings.xlsx')
    expect(file.sheets[0]?.printSettings).toMatchObject({
      orientation: 'landscape',
      margins: { left: 0.98, top: 5 },
      differentFirst: true,
      differentOddEven: true,
    })
  })

  it('sets and reopens native allow-edit ranges on a worksheet', async () => {
    const workbook = await openBrowserWorkbook(await fixture(), 'budget.xlsx')
    expect(workbook.protectedRanges('Budget')).toEqual([])
    workbook.setProtectedRanges('Budget', [
      { name: 'Inputs', sqref: 'B2:B10' },
      { name: 'Rates', sqref: 'D2 D4:D8' },
    ])
    const reopened = await openBrowserWorkbook(await workbook.save(), 'budget.xlsx')
    expect(reopened.protectedRanges('Budget')).toEqual([
      { name: 'Inputs', sqref: 'B2:B10', hasPassword: false },
      { name: 'Rates', sqref: 'D2 D4:D8', hasPassword: false },
    ])
  })

  it('writes and reopens native sparklines through the browser boundary', async () => {
    const workbook = await openBrowserWorkbook(await fixture(), 'budget.xlsx')

    workbook.applySparklines('Budget', [
      {
        type: 'column',
        cells: [{ cell: 'C2', sourceRef: "'Budget'!$A$2:$B$2" }],
      },
    ])
    const reopened = await openBrowserWorkbook(await workbook.save(), 'budget.xlsx')

    expect(reopened.sheets[0]?.sparklines).toEqual([
      {
        type: 'column',
        color: '#376092',
        negativeColor: '#D00000',
        cells: [{ cell: 'C2', sourceRef: "'Budget'!$A$2:$B$2" }],
      },
    ])
  })

  it('writes, reopens, and removes comparison conditional formatting through the browser boundary', async () => {
    const workbook = await openBrowserWorkbook(await fixture(), 'budget.xlsx')
    workbook.applyConditionalFormats('Budget', [
      {
        ranges: [{ startRow: 0, endRow: 4, startColumn: 0, endColumn: 0 }],
        stopIfTrue: true,
        rule: {
          type: 'highlightCell',
          subType: 'number',
          operator: 'greaterThan',
          value: 10,
          style: { bg: { rgb: '#FFF2CC' } },
        },
      },
    ])
    const formattedBytes = await workbook.save()
    const formattedZip = await JSZip.loadAsync(formattedBytes)
    expect(await formattedZip.file('xl/worksheets/sheet1.xml')!.async('text')).toContain(
      '<cfRule type="cellIs" dxfId="0" priority="1" stopIfTrue="1" operator="greaterThan"><formula>10</formula></cfRule>',
    )
    expect(await formattedZip.file('xl/styles.xml')!.async('text')).toContain(
      '<dxf><fill><patternFill><bgColor rgb="FFFFF2CC"/></patternFill></fill></dxf>',
    )

    const reopened = await openBrowserWorkbook(formattedBytes, 'budget.xlsx')
    reopened.applyConditionalFormats('Budget', [])
    const clearedZip = await JSZip.loadAsync(await reopened.save())
    expect(await clearedZip.file('xl/worksheets/sheet1.xml')!.async('text')).not.toContain(
      'conditionalFormatting',
    )
  })

  it('reopens row and column outline metadata after structural edits', async () => {
    const workbook = await openBrowserWorkbook(await fixture(), 'budget.xlsx')
    workbook.applyStructuralOperation('Budget', {
      kind: 'set-rows-outline',
      start: 1,
      end: 2,
      level: 2,
    })
    workbook.applyStructuralOperation('Budget', {
      kind: 'set-rows-hidden',
      start: 1,
      end: 2,
      hidden: true,
    })
    workbook.applyStructuralOperation('Budget', {
      kind: 'set-rows-outline',
      start: 3,
      end: 3,
      level: 0,
      collapsed: true,
    })
    workbook.applyStructuralOperation('Budget', {
      kind: 'set-cols-outline',
      start: 1,
      end: 2,
      level: 1,
    })
    workbook.applyStructuralOperation('Budget', {
      kind: 'set-cols-hidden',
      start: 1,
      end: 2,
      hidden: true,
    })
    workbook.applyStructuralOperation('Budget', {
      kind: 'set-cols-outline',
      start: 3,
      end: 3,
      level: 0,
      collapsed: true,
    })

    const reopened = await openBrowserWorkbook(await workbook.save(), 'budget.xlsx')

    expect(reopened.sheets[0]).toMatchObject({
      rows: expect.arrayContaining([
        { row: 1, hidden: true, outlineLevel: 2 },
        { row: 2, hidden: true, outlineLevel: 2 },
        { row: 3, hidden: false, collapsed: true },
      ]),
      columnWidths: expect.arrayContaining([
        { startColumn: 1, endColumn: 2, hidden: true, outlineLevel: 1 },
        { startColumn: 3, endColumn: 3, hidden: false, collapsed: true },
      ]),
    })
  })

  it('writes and reopens checkbox data validation through the browser boundary', async () => {
    const workbook = await openBrowserWorkbook(await fixture(), 'budget.xlsx')
    ;(
      workbook as unknown as {
        applyDataValidations(
          sheetName: string,
          rules: readonly {
            ranges: readonly {
              startRow: number
              endRow: number
              startColumn: number
              endColumn: number
            }[]
            rule: Record<string, unknown>
          }[],
        ): void
      }
    ).applyDataValidations('Budget', [
      {
        ranges: [{ startRow: 1, endRow: 2, startColumn: 1, endColumn: 2 }],
        rule: { type: 'checkbox', allowBlank: true },
      },
    ])

    const reopened = await openBrowserWorkbook(await workbook.save(), 'budget.xlsx')

    expect(reopened.sheets[0]).toMatchObject({
      dataValidations: [
        {
          ranges: [{ startRow: 1, endRow: 2, startColumn: 1, endColumn: 2 }],
          ruleType: 'list',
          formulas: ['"1,0"'],
          allowBlank: true,
        },
      ],
    })
  })

  it('writes, reopens, and removes inline list validation through the browser boundary', async () => {
    const workbook = await openBrowserWorkbook(await fixture(), 'budget.xlsx')
    const apply = (
      target: typeof workbook,
      rules: readonly {
        ranges: readonly {
          startRow: number
          endRow: number
          startColumn: number
          endColumn: number
        }[]
        rule: Record<string, unknown>
      }[],
    ) =>
      (
        target as unknown as {
          applyDataValidations(sheetName: string, rules: typeof rules): void
        }
      ).applyDataValidations('Budget', rules)

    apply(workbook, [
      {
        ranges: [{ startRow: 1, endRow: 4, startColumn: 1, endColumn: 1 }],
        rule: {
          type: 'list',
          formula1: 'Open,Closed',
          allowBlank: true,
          showDropDown: true,
        },
      },
    ])
    const reopened = await openBrowserWorkbook(await workbook.save(), 'budget.xlsx')
    expect(reopened.sheets[0]?.dataValidations).toEqual([
      expect.objectContaining({
        ranges: [{ startRow: 1, endRow: 4, startColumn: 1, endColumn: 1 }],
        ruleType: 'list',
        formulas: ['"Open,Closed"'],
        allowBlank: true,
        suppressDropdown: false,
      }),
    ])

    apply(reopened, [])
    const cleared = await openBrowserWorkbook(await reopened.save(), 'budget.xlsx')
    expect(cleared.sheets[0]?.dataValidations).toEqual([])
  })

  it('writes and reopens range-backed list validation through the browser boundary', async () => {
    const workbook = await openBrowserWorkbook(await fixture(), 'budget.xlsx')
    ;(
      workbook as unknown as {
        applyDataValidations(
          sheetName: string,
          rules: readonly {
            ranges: readonly {
              startRow: number
              endRow: number
              startColumn: number
              endColumn: number
            }[]
            rule: Record<string, unknown>
          }[],
        ): void
      }
    ).applyDataValidations('Budget', [
      {
        ranges: [{ startRow: 0, endRow: 9, startColumn: 0, endColumn: 0 }],
        rule: {
          type: 'list',
          formula1: '=Budget!$D$1:$D$3',
          allowBlank: false,
          showDropDown: true,
        },
      },
    ])

    const reopened = await openBrowserWorkbook(await workbook.save(), 'budget.xlsx')
    expect(reopened.sheets[0]?.dataValidations).toEqual([
      expect.objectContaining({
        ranges: [{ startRow: 0, endRow: 9, startColumn: 0, endColumn: 0 }],
        ruleType: 'list',
        formulas: ['Budget!$D$1:$D$3'],
        suppressDropdown: false,
      }),
    ])
  })

  it('writes a table through the transactional browser package boundary', async () => {
    const workbook = await openBrowserWorkbook(await fixture(), 'budget.xlsx')

    await workbook.applyTables([
      {
        sheetName: 'Budget',
        area: { startRow: 0, startColumn: 2, endRow: 2, endColumn: 3 },
        name: 'Table1',
        columnNames: ['Region', 'Amount'],
        style: 'TableStyleMedium4',
        bandedRows: true,
      },
    ])

    const saved = await workbook.save()
    const zip = await JSZip.loadAsync(saved)
    expect(await zip.file('xl/tables/table1.xml')!.async('text')).toMatch(
      /<table\b[^>]*\bdisplayName="Table1"[^>]*\bref="C1:D3"/,
    )
    expect(await zip.file('xl/worksheets/sheet1.xml')!.async('text')).toContain('<tableParts')
    expect(await zip.file('xl/worksheets/_rels/sheet1.xml.rels')!.async('text')).toContain(
      'Target="../tables/table1.xml"',
    )

    const reopened = await openBrowserWorkbook(saved, 'budget.xlsx')
    expect(reopened.sheets[0]?.tables).toEqual([
      {
        range: { startRow: 0, startColumn: 2, endRow: 2, endColumn: 3 },
        headerRowCount: 1,
        showRowStripes: true,
        showColumnStripes: false,
        name: 'Table1',
        columns: ['Region', 'Amount'],
        styleName: 'TableStyleMedium4',
        headerFill: '#A5A5A5',
        headerFontColor: '#FFFFFF',
        stripeFill: '#EDEDED',
        totalRowBorderColor: '#A5A5A5',
        totalRowBorderStyle: 'double',
      },
    ])
  })

  it('writes a chart through the transactional browser package boundary', async () => {
    const workbook = await openBrowserWorkbook(await fixture(), 'budget.xlsx')

    await workbook.applyVisuals([
      {
        sheetName: 'Budget',
        anchor: {
          fromRow: 0,
          fromColumn: 3,
          fromRowOffset: 0,
          fromColumnOffset: 0,
          toRow: 15,
          toColumn: 10,
          toRowOffset: 0,
          toColumnOffset: 0,
        },
        chart: {
          chartType: 'line',
          title: 'Revenue',
          series: [
            {
              name: 'Revenue',
              categories: ['Q1', 'Q2'],
              values: [10, 20],
              categoriesRef: "'Budget'!$A$2:$A$3",
              valuesRef: "'Budget'!$B$2:$B$3",
            },
          ],
        },
      },
    ])

    const saved = await workbook.save()
    const zip = await JSZip.loadAsync(saved)
    expect(await zip.file('xl/charts/chart2.xml')!.async('text')).toContain('<c:lineChart>')
    expect(await zip.file('xl/charts/chart2.xml')!.async('text')).toContain("'Budget'!$B$2:$B$3")
    expect(await zip.file('xl/worksheets/sheet1.xml')!.async('text')).toContain('<drawing')
    expect(await zip.file('xl/drawings/drawing1.xml')!.async('text')).toContain('<xdr:graphicFrame')

    const reopened = await openBrowserWorkbook(saved, 'budget.xlsx')
    expect(reopened.visuals).toEqual([
      expect.objectContaining({
        id: 'file-chart-xl-charts-chart2-xml',
        sheetName: 'Budget',
        kind: 'chart',
        chartPath: 'xl/charts/chart2.xml',
        drawingPath: 'xl/drawings/drawing1.xml',
        drawingIndex: 0,
        chart: expect.objectContaining({
          chartTypes: ['lineChart'],
          title: 'Revenue',
          series: [
            expect.objectContaining({
              name: 'Revenue',
              categories: ['Q1', 'Q2'],
              values: [10, 20],
              lineColor: '#4472C4',
            }),
          ],
        }),
      }),
    ])

    const noLineZip = await JSZip.loadAsync(saved)
    const noLineChart = (await noLineZip.file('xl/charts/chart2.xml')!.async('text')).replace(
      /<a:ln\b[^>]*>[\s\S]*?<\/a:ln>/,
      '<a:ln><a:noFill/></a:ln>',
    )
    noLineZip.file('xl/charts/chart2.xml', noLineChart)
    const noLine = await openBrowserWorkbook(
      await noLineZip.generateAsync({ type: 'uint8array' }),
      'no-line.xlsx',
    )
    expect(noLine.visuals[0]?.chart?.series[0]?.lineColor).toBe('none')
  })

  it('reopens sparse chart caches with their blank-point display mode', async () => {
    const workbook = await openBrowserWorkbook(await fixture(), 'budget.xlsx')
    await workbook.applyVisuals([
      {
        sheetName: 'Budget',
        anchor: {
          fromRow: 0,
          fromColumn: 3,
          fromRowOffset: 0,
          fromColumnOffset: 0,
          toRow: 15,
          toColumn: 10,
          toRowOffset: 0,
          toColumnOffset: 0,
        },
        chart: {
          chartType: 'line',
          title: 'Sparse revenue',
          series: [{ name: 'Revenue', categories: ['Q1', 'Q2', 'Q3'], values: [10, 0, 30] }],
        },
      },
    ])

    const zip = await JSZip.loadAsync(await workbook.save())
    const chartXml = await zip.file('xl/charts/chart2.xml')!.async('text')
    zip.file(
      'xl/charts/chart2.xml',
      chartXml
        .replace(
          /<c:num(?:Cache|Lit)>[\s\S]*?<\/c:num(?:Cache|Lit)>/,
          '<c:numLit><c:ptCount val="3"/><c:pt idx="0"><c:v>10</c:v></c:pt><c:pt idx="2"><c:v>30</c:v></c:pt></c:numLit>',
        )
        .replace(
          '<c:valAx>',
          '<c:valAx><c:majorUnit val="5"/><c:numFmt formatCode="$#,##0" sourceLinked="0"/>',
        )
        .replace('</c:chart>', '<c:dispBlanksAs val="gap"/></c:chart>'),
    )

    const reopened = await openBrowserWorkbook(
      await zip.generateAsync({ type: 'uint8array' }),
      'sparse.xlsx',
    )
    expect(reopened.visuals[0]?.chart).toEqual(
      expect.objectContaining({
        dispBlanksAs: 'gap',
        valueAxis: { majorUnit: 5, numFmt: '$#,##0' },
        series: [
          expect.objectContaining({
            values: [10, 0, 30],
            blanks: [1],
          }),
        ],
      }),
    )
  })

  it('writes and reopens a shape through the transactional browser package boundary', async () => {
    const workbook = await openBrowserWorkbook(await fixture(), 'budget.xlsx')

    await workbook.applyVisuals([
      {
        sheetName: 'Budget',
        anchor: {
          fromRow: 1,
          fromColumn: 3,
          fromRowOffset: 0,
          fromColumnOffset: 0,
          toRow: 7,
          toColumn: 6,
          toRowOffset: 0,
          toColumnOffset: 0,
        },
        shape: { shapeType: 'roundRect', fillColor: '#DDEBF7', text: 'Forecast' },
      },
    ])

    const saved = await workbook.save()
    const zip = await JSZip.loadAsync(saved)
    const drawingXml = await zip.file('xl/drawings/drawing1.xml')!.async('text')
    expect(drawingXml).toContain('<a:prstGeom prst="roundRect">')
    expect(drawingXml).toContain('<a:srgbClr val="DDEBF7"/>')
    expect(drawingXml).toContain('<a:t>Forecast</a:t>')

    const reopened = await openBrowserWorkbook(saved, 'budget.xlsx')
    expect(reopened.visuals).toEqual([
      expect.objectContaining({
        id: 'file-shape-xl-drawings-drawing1-xml-0',
        sheetName: 'Budget',
        kind: 'shape',
        shapeType: 'roundRect',
        fillColor: '#DDEBF7',
        text: 'Forecast',
        drawingPath: 'xl/drawings/drawing1.xml',
        drawingIndex: 0,
      }),
    ])

    await reopened.applyVisualEdits([
      {
        drawingPath: 'xl/drawings/drawing1.xml',
        drawingIndex: 0,
        anchor: {
          fromRow: 3,
          fromColumn: 5,
          fromRowOffset: 0,
          fromColumnOffset: 0,
          toRow: 9,
          toColumn: 8,
          toRowOffset: 0,
          toColumnOffset: 0,
        },
        fillColor: '#4472C4',
        text: 'Updated forecast',
      },
    ])
    const edited = await openBrowserWorkbook(await reopened.save(), 'budget.xlsx')
    expect(edited.visuals[0]).toMatchObject({
      id: 'file-shape-xl-drawings-drawing1-xml-0',
      fillColor: '#4472C4',
      text: 'Updated forecast',
      anchor: { fromRow: 3, fromColumn: 5, toRow: 9, toColumn: 8 },
    })

    await edited.applyVisualEdits([
      { drawingPath: 'xl/drawings/drawing1.xml', drawingIndex: 0, remove: true },
    ])
    expect((await openBrowserWorkbook(await edited.save(), 'budget.xlsx')).visuals).toEqual([])
  })

  it('writes, reads, moves, and removes an image through the browser package boundary', async () => {
    const workbook = await openBrowserWorkbook(await fixture(), 'budget.xlsx')
    const bytes = new TextEncoder().encode('GIF87a')

    await workbook.applyVisuals([
      {
        sheetName: 'Budget',
        anchor: {
          fromRow: 1,
          fromColumn: 3,
          fromRowOffset: 0,
          fromColumnOffset: 0,
          toRow: 7,
          toColumn: 6,
          toRowOffset: 0,
          toColumnOffset: 0,
        },
        image: { mediaType: 'image/gif', base64: 'R0lGODdh' },
      },
    ])

    const reopened = await openBrowserWorkbook(await workbook.save(), 'budget.xlsx')
    expect(reopened.visuals).toEqual([
      expect.objectContaining({
        id: 'file-image-xl-media-image1-gif-0',
        sheetName: 'Budget',
        kind: 'image',
        mediaPath: 'xl/media/image1.gif',
        mediaType: 'image/gif',
        drawingPath: 'xl/drawings/drawing1.xml',
        drawingIndex: 0,
      }),
    ])
    await expect(reopened.readBinary('xl/media/image1.gif')).resolves.toEqual(bytes)

    await reopened.applyVisualEdits([
      {
        drawingPath: 'xl/drawings/drawing1.xml',
        drawingIndex: 0,
        anchor: {
          fromRow: 3,
          fromColumn: 5,
          fromRowOffset: 0,
          fromColumnOffset: 0,
          toRow: 9,
          toColumn: 8,
          toRowOffset: 0,
          toColumnOffset: 0,
        },
      },
    ])
    const moved = await openBrowserWorkbook(await reopened.save(), 'budget.xlsx')
    expect(moved.visuals[0]?.anchor).toMatchObject({
      fromRow: 3,
      fromColumn: 5,
      toRow: 9,
      toColumn: 8,
    })
    await moved.applyVisualEdits([
      { drawingPath: 'xl/drawings/drawing1.xml', drawingIndex: 0, remove: true },
    ])
    expect((await openBrowserWorkbook(await moved.save(), 'budget.xlsx')).visuals).toEqual([])
  })

  it('patches an existing chart through the browser package boundary', async () => {
    const zip = await JSZip.loadAsync(await fixture())
    zip.file(
      'xl/charts/chart1.xml',
      buildChartXml({
        chartType: 'line',
        title: 'Old title',
        series: [{ name: 'Revenue', categories: ['Q1'], values: [10] }],
      }),
    )
    const workbook = await openBrowserWorkbook(
      await zip.generateAsync({ type: 'uint8array' }),
      'budget.xlsx',
    )

    await workbook.applyChartEdits([
      { chartPath: 'xl/charts/chart1.xml', title: 'Revenue', legend: 'bottom' },
    ])

    const saved = await JSZip.loadAsync(await workbook.save())
    const chartXml = await saved.file('xl/charts/chart1.xml')!.async('text')
    expect(chartXml).toContain('<a:t>Revenue</a:t>')
    expect(chartXml).toContain('<c:legendPos val="b"/>')
  })

  it('shifts formulas on other sheets when a row is inserted and the workbook reopens', async () => {
    const zip = await JSZip.loadAsync(await fixture())
    zip.file(
      'xl/workbook.xml',
      '<workbook><sheets><sheet name="Budget" sheetId="1" r:id="rId1"/>' +
        '<sheet name="Summary" sheetId="2" r:id="rId2"/></sheets></workbook>',
    )
    zip.file(
      'xl/_rels/workbook.xml.rels',
      '<Relationships><Relationship Target="worksheets/sheet1.xml" Id="rId1"/>' +
        '<Relationship Target="worksheets/sheet2.xml" Id="rId2"/></Relationships>',
    )
    zip.file(
      'xl/worksheets/sheet2.xml',
      '<worksheet><sheetData><row r="1"><c r="A1"><f>Budget!A1</f><v>0</v></c></row>' +
        '</sheetData></worksheet>',
    )
    const workbook = await openBrowserWorkbook(
      await zip.generateAsync({ type: 'uint8array' }),
      'cross-sheet.xlsx',
    )

    workbook.applyStructuralOperation('Budget', { kind: 'insert-rows', index: 0, count: 1 })
    const saved = await workbook.save()
    const reopened = await openBrowserWorkbook(saved, 'cross-sheet.xlsx')

    expect(reopened.cell('Summary', 'A1')?.formula).toBe('Budget!A2')
  })

  it('duplicates and hides sheets through the browser package adapter and reopens both states', async () => {
    const workbook = await openBrowserWorkbook(await fixture(), 'budget.xlsx')

    workbook.duplicateSheet('Budget', 'Budget Copy')
    workbook.setSheetHidden('Budget', true)

    const saved = await workbook.save()
    const archive = await JSZip.loadAsync(saved)
    const workbookXml = await archive.file('xl/workbook.xml')!.async('text')
    expect(workbookXml).toContain('<sheet state="hidden" name="Budget" sheetId="1" r:id="rId1"/>')
    expect(workbookXml).toContain('<sheet name="Budget Copy" sheetId="2" r:id="rId2"/>')
    expect(await archive.file('xl/worksheets/sheet2.xml')!.async('text')).toContain(
      '<f>SUM(B2:B3)</f>',
    )

    const reopened = await openBrowserWorkbook(saved, 'budget.xlsx')
    expect(reopened.sheets.map(({ name, hidden }) => ({ name, hidden }))).toEqual([
      { name: 'Budget', hidden: true },
      { name: 'Budget Copy', hidden: false },
    ])
    expect(reopened.cell('Budget Copy', 'A1')?.value).toBe('Revenue')
    expect(reopened.cell('Budget Copy', 'B1')?.formula).toBe('SUM(B2:B3)')
  })

  it('writes the first browser save to the selected local XLSX file before reporting success', async () => {
    const memoryFile = inMemoryFileHandle('saved-budget.xlsx')
    const showSaveFilePicker = vi.fn(async () => memoryFile.handle)
    vi.stubGlobal('document', {
      createElement: () => ({ href: '', download: '', click: vi.fn() }),
    })
    vi.stubGlobal('window', {
      showSaveFilePicker,
      setTimeout: (callback: () => void) => callback(),
    })
    vi.stubGlobal('URL', { createObjectURL: () => 'blob:test', revokeObjectURL: vi.fn() })
    const api = new BrowserWorkbookDesktopApi(async () => null)
    const bytes = await fixture()
    const opened = await api.openBuffer(
      bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer,
      'budget.xlsx',
    )
    const result = await api.saveWorkbookEdits(cellSaveRequest(opened, 42))

    expect(showSaveFilePicker).toHaveBeenCalledTimes(1)
    expect(memoryFile.writable.write).toHaveBeenCalledTimes(1)
    expect(memoryFile.writable.close).toHaveBeenCalledTimes(1)
    expect(result.canceled).toBe(false)
    const reopened = await openBrowserWorkbook(memoryFile.bytes(), 'saved-budget.xlsx')
    expect(reopened.cell('Budget', 'B2')?.value).toBe(42)
  })

  it('overwrites the same local XLSX handle on later saves without reopening the picker', async () => {
    const memoryFile = inMemoryFileHandle('saved-budget.xlsx')
    const showSaveFilePicker = vi.fn(async () => memoryFile.handle)
    vi.stubGlobal('window', { showSaveFilePicker })
    const api = new BrowserWorkbookDesktopApi(async () => null)
    const bytes = await fixture()
    const opened = await api.openBuffer(
      bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer,
      'budget.xlsx',
    )

    const first = await api.saveWorkbookEdits(cellSaveRequest(opened, 42))
    if (first.canceled) throw new Error('The first browser save was unexpectedly canceled.')
    const second = await api.saveWorkbookEdits(cellSaveRequest(first.file, 84))

    expect(second.canceled).toBe(false)
    expect(showSaveFilePicker).toHaveBeenCalledTimes(1)
    expect(memoryFile.handle.createWritable).toHaveBeenCalledTimes(2)
    const reopened = await openBrowserWorkbook(memoryFile.bytes(), 'saved-budget.xlsx')
    expect(reopened.cell('Budget', 'B2')?.value).toBe(84)
  })

  it.each([false, true])(
    'serializes pending XLSX edits into recovery bytes without a picker or source mutation (embedded: %s)',
    async (embedded) => {
      vi.stubGlobal('window', embedded ? { parent: {} } : {})
      const api = new BrowserWorkbookDesktopApi(async () => null)
      const bytes = await fixture()
      const opened = await api.openBuffer(
        bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer,
        'budget.xlsx',
      )

      const recovery = await api.writeWorkbookRecovery(cellSaveRequest(opened, 42))

      expect(recovery.ok).toBe(true)
      if (!recovery.ok) throw new Error('The browser workbook recovery was not created.')
      const reopened = await openBrowserWorkbook(new Uint8Array(recovery.data), recovery.fileName)
      expect(reopened.cell('Budget', 'B2')?.value).toBe(42)
      const source = await api.readWorkbookRange({
        sessionId: opened.sessionId,
        sheetId: opened.sheets[0]!.id,
        range: { startRow: 1, endRow: 1, startColumn: 1, endColumn: 1 },
      })
      expect(source.cells[0]?.value).toBe(10)
    },
  )

  it('routes duplicate and visibility journal operations through the browser desktop save', async () => {
    const memoryFile = inMemoryFileHandle('saved-budget.xlsx')
    vi.stubGlobal('document', {
      createElement: () => ({ href: '', download: '', click: vi.fn() }),
    })
    vi.stubGlobal('window', {
      showSaveFilePicker: vi.fn(async () => memoryFile.handle),
      setTimeout: (callback: () => void) => callback(),
    })
    vi.stubGlobal('URL', { createObjectURL: () => 'blob:test', revokeObjectURL: vi.fn() })
    const api = new BrowserWorkbookDesktopApi(async () => null)
    const bytes = await fixture()
    const source = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength)
    const opened = await api.openBuffer(source as ArrayBuffer, 'budget.xlsx')
    const sourceSheetId = opened.sheets[0]!.id
    const copySheetId = 'sheet-copy'
    const request: WorkbookSaveRequest = {
      sessionId: opened.sessionId,
      mode: 'save',
      edits: [],
      structuralOps: [],
      chartEdits: [],
      visualEdits: [],
      visualAdditions: [],
      tableAdditions: [],
      pivotAdditions: [],
      sheetOps: [
        {
          kind: 'duplicate-sheet',
          sheetId: copySheetId,
          name: 'Budget Copy',
          sourceSheetId,
        },
        { kind: 'set-sheet-hidden', sheetId: sourceSheetId, hidden: true },
      ],
      sheetOrder: [sourceSheetId, copySheetId],
      filterStates: [],
      hyperlinkEdits: [],
      cfStates: [],
      dvStates: [],
      pageSetupStates: [],
      noteStates: [],
      formulaValues: [],
      pivotCacheRefreshPaths: [],
      pivotRefreshUpdates: [],
      sheetProtections: [],
      sparklineAdditions: [],
      definedNamesState: null,
    }

    const result = await api.saveWorkbookEdits(request)

    expect(result.canceled).toBe(false)
    if (result.canceled) throw new Error('The browser save was unexpectedly canceled.')
    expect(result.file.sheets.map(({ name, hidden }) => ({ name, hidden }))).toEqual([
      { name: 'Budget', hidden: true },
      { name: 'Budget Copy', hidden: false },
    ])
  })
})
