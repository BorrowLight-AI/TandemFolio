/// Lays the active sheet out as print HTML from the live Univer model —
/// display strings (number formats applied), cell styles, merges, and the
/// sheet's Page Layout settings (print area, repeated title rows, gridlines,
/// headings). The main process turns the HTML into a PDF.

import { htmlLang, type Lang } from '@genoffice/i18n'
import { BorderStyleTypes } from '@univerjs/core'
import { columnIndex, columnLabel } from '../domain/cell-address'

import type { WorkbookExportPdfRequest } from '../shared/desktop-api'
import type { HeaderFooterParts, PageSetupJournalState } from './edit-journal'
import { getLang, t } from './i18n/locale'
import { fitToPageScale } from './print-scale'
import {
  resolveEffectivePageSetup,
  type EffectivePageSetup,
  type HeaderFooterPair,
  type PrintMargins,
} from './print-settings'

export class PrintError extends Error {}

/** UI-language CJK fallback for the print stack (mirrors the :lang() variables in styles.css) */
function printCjkFonts(lang: Lang): string {
  switch (lang) {
    case 'ja':
      return "'Hiragino Sans', 'Hiragino Kaku Gothic ProN', 'Yu Gothic UI', 'Yu Gothic', 'Meiryo'"
    case 'ko':
      return "'Apple SD Gothic Neo', 'Malgun Gothic'"
    case 'zh-TW':
      return "'PingFang TC', 'Microsoft JhengHei'"
    default:
      return "'PingFang SC'"
  }
}

const MAX_PRINT_CELLS = 50_000

/// The slice of the Univer facade the layout needs (structural, so the
/// caller passes the FWorksheet through a cast).
export interface PrintWorksheet {
  getLastRow(): number
  getLastColumn(): number
  getRowHeight(row: number): number
  getColumnWidth(column: number): number
  getMergedRanges(): {
    getRow(): number
    getColumn(): number
    getWidth(): number
    getHeight(): number
  }[]
  getRange(
    row: number,
    column: number,
    numRows: number,
    numColumns: number,
  ): {
    getDisplayValues(): string[][]
    getValues(): unknown[][]
  }
  getRange(row: number, column: number): { getCellStyleData(): PrintCellStyle | null }
}

/// The IStyleData fields the layout reads (all optional in Univer).
interface PrintCellStyle {
  readonly bl?: number
  readonly it?: number
  readonly ul?: { s?: number } | null
  readonly st?: { s?: number } | null
  readonly fs?: number
  readonly ff?: string | null
  readonly cl?: { rgb?: string | null } | null
  readonly bg?: { rgb?: string | null } | null
  readonly ht?: number
  readonly vt?: number
  readonly tb?: number
  readonly bd?: Partial<
    Record<'t' | 'b' | 'l' | 'r', { s?: number; cl?: { rgb?: string | null } | null } | null>
  > | null
}

export function printBorderWidthPt(style: number | undefined): number {
  switch (style) {
    case BorderStyleTypes.MEDIUM:
    case BorderStyleTypes.MEDIUM_DASHED:
    case BorderStyleTypes.MEDIUM_DASH_DOT:
    case BorderStyleTypes.MEDIUM_DASH_DOT_DOT:
      return 1.5
    case BorderStyleTypes.THICK:
      return 2.25
    default:
      return 0.75
  }
}

/// Inches, mirroring the gateway's Margins presets.
const MARGIN_PRESETS = {
  normal: { left: 0.7, right: 0.7, top: 0.75, bottom: 0.75 },
  wide: { left: 1, right: 1, top: 1, bottom: 1 },
  narrow: { left: 0.25, right: 0.25, top: 0.75, bottom: 0.75 },
} as const

/// OOXML paper-size code → Electron pageSize (custom sizes in inches).
const PAPER_SIZES: Record<number, WorkbookExportPdfRequest['pageSize']> = {
  1: 'Letter',
  3: 'Tabloid',
  5: 'Legal',
  7: { width: 7.25, height: 10.5 },
  8: 'A3',
  9: 'A4',
  11: 'A5',
}

const PAPER_WIDTH_INCHES: Record<string, number> = {
  Letter: 8.5,
  Tabloid: 11,
  Legal: 8.5,
  A3: 11.69,
  A4: 8.27,
  A5: 5.83,
}

export function buildSheetPrintPayload(
  worksheet: PrintWorksheet,
  pageSetup: PageSetupJournalState | EffectivePageSetup,
  fileName: string,
  sheetName: string,
): WorkbookExportPdfRequest {
  const setup =
    'printAreas' in pageSetup ? pageSetup : resolveEffectivePageSetup(pageSetup, null, null)
  if (setup.printAreas.length > 1) {
    const parsed = setup.printAreas.map(parseArea)
    const totalCells = parsed.reduce(
      (total, area) =>
        total +
        (area.endRow - area.startRow + 1) * (area.endColumn - area.startColumn + 1),
      0,
    )
    if (totalCells > MAX_PRINT_CELLS) throw new PrintError(t('appPrintTooLarge'))
    const payloads = setup.printAreas.map((area) =>
      buildSheetPrintPayload(worksheet, { ...setup, printAreas: [area] }, fileName, sheetName),
    )
    const first = payloads[0]!
    const tables = payloads
      .map((payload) => /<body>([\s\S]*)<\/body>/.exec(payload.html)?.[1] ?? '')
      .join('')
    return {
      ...first,
      html: first.html
        .replace('</style>', 'table + table { break-before: page; }</style>')
        .replace(/<body>[\s\S]*<\/body>/, `<body>${tables}</body>`),
    }
  }
  const area = setup.printAreas[0] ? parseArea(setup.printAreas[0]) : usedArea(worksheet)
  const rows = area.endRow - area.startRow + 1
  const columns = area.endColumn - area.startColumn + 1
  if (rows < 1 || columns < 1) throw new PrintError(t('appPrintNothing'))
  if (rows * columns > MAX_PRINT_CELLS) {
    throw new PrintError(t('appPrintTooLarge'))
  }

  const titles = setup.printTitles ? parseTitleRows(setup.printTitles) : null
  const grid = worksheet.getRange(area.startRow, area.startColumn, rows, columns)
  const display = grid.getDisplayValues()
  const raw = grid.getValues()
  const merges = mergeMaps(worksheet, area)

  const headings = setup.printHeadings
  const gridlines = setup.printGridlines
  const columnWidthsPt = Array.from(
    { length: columns },
    (_, offset) => worksheet.getColumnWidth(area.startColumn + offset) * 0.75,
  )
  const rowHeaderPt = headings ? 24 : 0

  const bodyRow = (row: number): string => {
    const cells: string[] = []
    if (headings) {
      cells.push(`<th class="hd">${row + 1}</th>`)
    }
    for (let column = area.startColumn; column <= area.endColumn; column += 1) {
      const key = `${row}:${column}`
      if (merges.covered.has(key)) continue
      const anchor = merges.anchors.get(key)
      const span = anchor
        ? ` rowspan="${Math.min(anchor.rows, area.endRow - row + 1)}"` +
          ` colspan="${Math.min(anchor.columns, area.endColumn - column + 1)}"`
        : ''
      const inArea = row >= area.startRow && row <= area.endRow
      const text = inArea
        ? (display[row - area.startRow]?.[column - area.startColumn] ?? '')
        : cellDisplay(worksheet, row, column)
      const rawValue = inArea ? raw[row - area.startRow]?.[column - area.startColumn] : undefined
      const style = worksheet.getRange(row, column).getCellStyleData()
      cells.push(
        `<td${span} style="${cellCss(style, rawValue, gridlines)}">${escapeHtml(text)}</td>`,
      )
    }
    const heightPt = Math.max(worksheet.getRowHeight(row) * 0.75, 10)
    return `<tr style="height:${round(heightPt)}pt">${cells.join('')}</tr>`
  }

  const headParts: string[] = []
  if (headings) {
    const letters = Array.from(
      { length: columns },
      (_, offset) => `<th class="hd">${columnLabel(area.startColumn + offset)}</th>`,
    )
    headParts.push(`<tr>${headings ? '<th class="hd"></th>' : ''}${letters.join('')}</tr>`)
  }
  if (titles) {
    for (let row = titles.start; row <= titles.end; row += 1) headParts.push(bodyRow(row))
  }

  const bodyParts: string[] = []
  for (let row = area.startRow; row <= area.endRow; row += 1) {
    // Title rows already repeat via the table header.
    if (titles && row >= titles.start && row <= titles.end) continue
    bodyParts.push(bodyRow(row))
  }

  const colgroup = `<colgroup>${headings ? `<col style="width:${rowHeaderPt}pt">` : ''}${columnWidthsPt
    .map((width) => `<col style="width:${round(width)}pt">`)
    .join('')}</colgroup>`
  const html =
    `<!doctype html><html lang="${htmlLang(getLang())}"><head><meta charset="utf-8"><style>
* { box-sizing: border-box; }
body { margin: 0; font-family: Calibri, 'Helvetica Neue', Arial, ${printCjkFonts(getLang())}, sans-serif; }
table { border-collapse: collapse; table-layout: fixed; }
thead { display: table-header-group; }
tfoot { display: table-footer-group; }
td, th { overflow: hidden; padding: 1pt 3pt; font-size: 11pt; vertical-align: bottom; }
th.hd { background: #f1f1f1; border: 0.5pt solid #b7b7b7; color: #444;
  font-size: 8.5pt; font-weight: 400; text-align: center; vertical-align: middle; }
th.hf, td.hf { padding: 2pt 0 6pt; font-weight: 400; }
td.hf { padding: 6pt 0 0; }
.hf > div { display: flex; font-size: 9pt; color: #333; }
.hf span { flex: 1; white-space: pre; }
.hf span:nth-child(2) { text-align: center; }
.hf span:last-child { text-align: right; }
</style></head><body><table>${colgroup}<thead>${headParts.join('')}</thead>` +
    `<tbody>${bodyParts.join('')}</tbody></table></body></html>`

  const margins = setup.margins
  const pageSize = PAPER_SIZES[setup.paperSize] ?? 'A4'
  const landscape = setup.orientation === 'landscape'
  const now = new Date()
  const baseName = fileName.replace(/\.pdf$/, '')
  const scale = computeScale(
    setup,
    pageSize,
    landscape,
    margins,
    rowHeaderPt + columnWidthsPt.reduce((total, width) => total + width, 0),
    {
      repeatedHeightPt:
        (headings ? 15 : 0) +
        (titles
          ? Array.from(
              { length: titles.end - titles.start + 1 },
              (_, offset) => Math.max(worksheet.getRowHeight(titles.start + offset) * 0.75, 10),
            ).reduce((total, height) => total + height, 0)
          : 0),
      rowHeightsPt: Array.from({ length: rows }, (_, offset) => area.startRow + offset)
        .filter((row) => !titles || row < titles.start || row > titles.end)
        .map((row) => Math.max(worksheet.getRowHeight(row) * 0.75, 10)),
    },
  )
  const templateScale = setup.headerFooterScaleWithDoc ? scale : 1
  const templates = (pair: HeaderFooterPair) => ({
    ...(pair.header
      ? {
          headerTemplate: buildHeaderFooterTemplate(
            pair.header,
            'header',
            margins,
            baseName,
            sheetName,
            now,
            templateScale,
          ),
        }
      : {}),
    ...(pair.footer
      ? {
          footerTemplate: buildHeaderFooterTemplate(
            pair.footer,
            'footer',
            margins,
            baseName,
            sheetName,
            now,
            templateScale,
          ),
        }
      : {}),
  })
  return {
    fileName,
    html,
    landscape,
    pageSize,
    margins: { top: margins.top, bottom: margins.bottom, left: margins.left, right: margins.right },
    scale,
    ...templates({ header: setup.header, footer: setup.footer }),
    ...(setup.firstPage === null ? {} : { firstPage: templates(setup.firstPage) }),
    ...(setup.evenPages === null ? {} : { evenPages: templates(setup.evenPages) }),
  }
}

export function renderHeaderFooterHtml(
  text: string,
  fileName: string,
  sheetName: string,
  now: Date,
): string {
  let output = ''
  let literal = ''
  const flush = (): void => {
    output += escapeHtml(literal)
    literal = ''
  }
  let cursor = 0
  for (const match of text.matchAll(/&(&|[A-Za-z])/g)) {
    literal += text.slice(cursor, match.index)
    const code = match[1]
    if (code === 'P' || code === 'N') {
      flush()
      output += `<span class="${code === 'P' ? 'pageNumber' : 'totalPages'}"></span>`
    } else if (code === '&') literal += '&'
    else if (code === 'D') literal += now.toLocaleDateString()
    else if (code === 'T') literal += now.toLocaleTimeString()
    else if (code === 'F') literal += fileName
    else if (code === 'A') literal += sheetName
    else literal += match[0]
    cursor = match.index + match[0].length
  }
  literal += text.slice(cursor)
  flush()
  return output
}

export function buildHeaderFooterTemplate(
  parts: HeaderFooterParts,
  kind: 'header' | 'footer',
  margins: PrintMargins,
  fileName: string,
  sheetName: string,
  now: Date,
  scale = 1,
): string {
  const rendered = [parts.left ?? '', parts.center ?? '', parts.right ?? ''].map((text) =>
    renderHeaderFooterHtml(text, fileName, sheetName, now),
  )
  const spanStyle = 'flex:1;min-width:0;white-space:pre-wrap'
  const offset =
    kind === 'header'
      ? `padding-top:${round(margins.header)}in`
      : `padding-bottom:${round(margins.footer)}in`
  return (
    `<div style="box-sizing:border-box;display:flex;width:100%;font-size:${round(9 * scale)}pt;` +
    `color:#000;font-family:Calibri,'Helvetica Neue',Arial,sans-serif;` +
    `padding-left:${round(margins.left)}in;padding-right:${round(margins.right)}in;${offset}">` +
    `<span style="${spanStyle}">${rendered[0]}</span>` +
    `<span style="${spanStyle};text-align:center">${rendered[1]}</span>` +
    `<span style="${spanStyle};text-align:right">${rendered[2]}</span></div>`
  )
}

/// Excel's fit-to-width only shrinks; an explicit scale applies as-is.
function computeScale(
  pageSetup: Pick<EffectivePageSetup, 'fitToPage' | 'fitToWidth' | 'fitToHeight' | 'scale'>,
  pageSize: WorkbookExportPdfRequest['pageSize'],
  landscape: boolean,
  margins: { left: number; right: number; top: number; bottom: number },
  contentWidthPt: number,
  area: { repeatedHeightPt: number; rowHeightsPt: readonly number[] },
): number {
  if (pageSetup.fitToPage === true) {
    const naturalWidth =
      typeof pageSize === 'string' ? (PAPER_WIDTH_INCHES[pageSize] ?? 8.27) : pageSize.width
    const naturalHeight =
      typeof pageSize === 'string' ? paperHeightInches(pageSize) : pageSize.height
    const paperWidthIn = landscape ? naturalHeight : naturalWidth
    const paperHeightIn = landscape ? naturalWidth : naturalHeight
    return fitToPageScale({
      printableWidthPt: Math.max(0, paperWidthIn - margins.left - margins.right) * 72,
      printableHeightPt: Math.max(0, paperHeightIn - margins.top - margins.bottom) * 72,
      fitToWidth: pageSetup.fitToWidth ?? 0,
      fitToHeight: pageSetup.fitToHeight ?? 0,
      contentWidthPt,
      areas: [area],
    })
  }
  if (pageSetup.scale !== undefined) {
    return clamp(pageSetup.scale / 100, 0.1, 2)
  }
  return 1
}

function paperHeightInches(name: string): number {
  const heights: Record<string, number> = {
    Letter: 11,
    Tabloid: 17,
    Legal: 14,
    A3: 16.54,
    A4: 11.69,
    A5: 8.27,
  }
  return heights[name] ?? 11.69
}

function usedArea(worksheet: PrintWorksheet) {
  return {
    startRow: 0,
    startColumn: 0,
    endRow: Math.max(worksheet.getLastRow(), 0),
    endColumn: Math.max(worksheet.getLastColumn(), 0),
  }
}

function parseArea(area: string) {
  const match = /^\$?([A-Za-z]{1,3})\$?(\d{1,7}):\$?([A-Za-z]{1,3})\$?(\d{1,7})$/.exec(area)
  if (!match) throw new PrintError(t('appPrintBadArea', { area }))
  return {
    startRow: Number(match[2]) - 1,
    startColumn: columnIndex(match[1] ?? 'A'),
    endRow: Number(match[4]) - 1,
    endColumn: columnIndex(match[3] ?? 'A'),
  }
}

function parseTitleRows(titles: string): { start: number; end: number } {
  const match = /^(\d{1,7}):(\d{1,7})$/.exec(titles)
  if (!match) throw new PrintError(t('appPrintBadTitles', { titles }))
  const start = Number(match[1]) - 1
  const end = Number(match[2]) - 1
  if (end - start > 20) throw new PrintError(t('appPrintTitlesLimit'))
  return { start, end }
}

function mergeMaps(
  worksheet: PrintWorksheet,
  area: { startRow: number; endRow: number; startColumn: number; endColumn: number },
) {
  const anchors = new Map<string, { rows: number; columns: number }>()
  const covered = new Set<string>()
  for (const merge of worksheet.getMergedRanges()) {
    const row = merge.getRow()
    const column = merge.getColumn()
    if (row > area.endRow || column > area.endColumn) continue
    if (row + merge.getHeight() - 1 < area.startRow) continue
    if (column + merge.getWidth() - 1 < area.startColumn) continue
    anchors.set(`${row}:${column}`, { rows: merge.getHeight(), columns: merge.getWidth() })
    for (let r = row; r < row + merge.getHeight(); r += 1) {
      for (let c = column; c < column + merge.getWidth(); c += 1) {
        if (r !== row || c !== column) covered.add(`${r}:${c}`)
      }
    }
  }
  return { anchors, covered }
}

function cellDisplay(worksheet: PrintWorksheet, row: number, column: number): string {
  return worksheet.getRange(row, column, 1, 1).getDisplayValues()[0]?.[0] ?? ''
}

function cellCss(style: PrintCellStyle | null, rawValue: unknown, gridlines: boolean): string {
  const rules: string[] = []
  if (style?.bl === 1) rules.push('font-weight:700')
  if (style?.it === 1) rules.push('font-style:italic')
  const decorations = [
    style?.ul?.s === 1 ? 'underline' : '',
    style?.st?.s === 1 ? 'line-through' : '',
  ].filter(Boolean)
  if (decorations.length > 0) rules.push(`text-decoration:${decorations.join(' ')}`)
  if (style?.fs) rules.push(`font-size:${round(style.fs)}pt`)
  // a font name comes straight from styles.xml; anything outside the whitelist could
  // close the style attribute and inject markup into the exported page
  const family = style?.ff?.replace(/[^\p{L}\p{N} \-_.]/gu, '')
  // fallbacks mirror the body stack: an uninstalled family (e.g. Aptos)
  // must not drop to the browser's serif default in the exported page
  if (family)
    rules.push(
      `font-family:'${family}',Calibri,'Helvetica Neue',Arial,${printCjkFonts(getLang())},sans-serif`,
    )
  if (style?.cl?.rgb) rules.push(`color:${cssColor(style.cl.rgb)}`)
  if (style?.bg?.rgb) rules.push(`background:${cssColor(style.bg.rgb)}`)
  const align =
    style?.ht === 1
      ? 'left'
      : style?.ht === 2
        ? 'center'
        : style?.ht === 3
          ? 'right'
          : typeof rawValue === 'number'
            ? 'right'
            : typeof rawValue === 'boolean'
              ? 'center'
              : 'left'
  rules.push(`text-align:${align}`)
  if (style?.vt === 1) rules.push('vertical-align:top')
  else if (style?.vt === 2) rules.push('vertical-align:middle')
  rules.push(style?.tb === 3 ? 'white-space:pre-wrap;word-break:break-word' : 'white-space:pre')
  const defaultBorder = gridlines ? '0.5pt solid #c0c0c0' : 'none'
  for (const [edge, css] of [
    ['t', 'top'],
    ['b', 'bottom'],
    ['l', 'left'],
    ['r', 'right'],
  ]) {
    const border = style?.bd?.[edge as 't' | 'b' | 'l' | 'r']
    rules.push(
      `border-${css}:${
        border
          ? `${printBorderWidthPt(border.s)}pt solid ${cssColor(border.cl?.rgb ?? '#000000')}`
          : defaultBorder
      }`,
    )
  }
  return rules.join(';')
}

function cssColor(rgb: string): string {
  return /^(#[0-9a-fA-F]{3,8}|rgba?\([\d ,.%]+\))$/.test(rgb) ? rgb : '#000'
}

function escapeHtml(value: string): string {
  return value.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
}

function clamp(value: number, min: number, max: number): number {
  return Math.min(max, Math.max(min, value))
}

function round(value: number): number {
  return Math.round(value * 100) / 100
}
