import JSZip from 'jszip'
import { afterEach, describe, expect, it, vi } from 'vitest'
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
        styleName: 'TableStyleMedium4',
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
            }),
          ],
        }),
      }),
    ])
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

  it('serializes pending XLSX edits into recovery bytes without a picker or source mutation', async () => {
    vi.stubGlobal('window', {})
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
  })

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
