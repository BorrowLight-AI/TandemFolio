// Modified by TandemFolio contributors: browser-owned workbook state and bounded overlays.
/** Page Break Preview boundaries anchored to the live Univer grid. */
import type { PageSetupJournalState } from './edit-journal'
import { t } from './i18n/locale'
import type { LazyWorkbookState, UniverRuntime, UniverWorksheet } from './univer-state'
import { fileToScreen } from './view-transform'

interface Disposable {
  dispose(): void
}

export interface PageBoundary {
  readonly index: number
  readonly manual: boolean
}

const PAPER_INCHES: Record<number, { width: number; height: number }> = {
  1: { width: 8.5, height: 11 },
  3: { width: 11, height: 17 },
  5: { width: 8.5, height: 14 },
  7: { width: 7.25, height: 10.5 },
  8: { width: 11.69, height: 16.54 },
  9: { width: 8.27, height: 11.69 },
  11: { width: 5.83, height: 8.27 },
}

const MARGIN_PRESETS = {
  normal: { left: 0.7, right: 0.7, top: 0.75, bottom: 0.75 },
  wide: { left: 1, right: 1, top: 1, bottom: 1 },
  narrow: { left: 0.25, right: 0.25, top: 0.75, bottom: 0.75 },
} as const

export function printablePagePt(pageSetup: PageSetupJournalState): {
  width: number
  height: number
} {
  const paper = PAPER_INCHES[pageSetup.paperSize ?? 9] ?? PAPER_INCHES[9]!
  const margins = MARGIN_PRESETS[pageSetup.margins ?? 'normal']
  const landscape = pageSetup.orientation === 'landscape'
  const width = (landscape ? paper.height : paper.width) - margins.left - margins.right
  const height = (landscape ? paper.width : paper.height) - margins.top - margins.bottom
  return { width: Math.max(width, 1) * 72, height: Math.max(height, 1) * 72 }
}

export function computePageBoundaries(
  sizeOf: (index: number) => number,
  count: number,
  printableSizePt: number,
  scale: number,
  manual: readonly number[],
): PageBoundary[] {
  const manualSet = new Set(manual.filter((index) => index > 0 && index < count))
  const boundaries: PageBoundary[] = []
  let used = 0
  for (let index = 0; index < count; index += 1) {
    if (manualSet.has(index)) {
      boundaries.push({ index, manual: true })
      used = 0
    }
    const size = sizeOf(index) * 0.75 * scale
    if (used > 0 && used + size > printableSizePt) {
      if (!manualSet.has(index)) boundaries.push({ index, manual: false })
      used = 0
    }
    used += size
  }
  return boundaries
}

type FilePageSetup = PageSetupJournalState & {
  readonly rowBreaks?: readonly number[]
  readonly colBreaks?: readonly number[]
}

export function effectivePageBreaks(
  state: LazyWorkbookState,
  sheetId: string,
): { rowBreaks: number[]; colBreaks: number[] } {
  const journal = state.editJournal.pageSetup.get(sheetId)
  const file = (
    state.file.sheets.find((sheet) => sheet.id === sheetId) as
      | { readonly pageSetup?: FilePageSetup }
      | undefined
  )?.pageSetup
  const ops = state.editJournal.structuralOps.get(sheetId) ?? []
  const mapped = (axis: 'row' | 'column', values: readonly number[]): number[] =>
    values
      .map((value) => fileToScreen(ops, axis, value))
      .filter((value): value is number => value !== null && value > 0)
  return {
    rowBreaks: journal?.rowBreaks ?? mapped('row', file?.rowBreaks ?? []),
    colBreaks: journal?.colBreaks ?? mapped('column', file?.colBreaks ?? []),
  }
}

function effectiveScale(
  pageSetup: PageSetupJournalState,
  contentWidthPt: number,
  printableWidthPt: number,
): number {
  const fitPages = pageSetup.fitToPage === true ? (pageSetup.fitToWidth ?? 0) : 0
  if (fitPages > 0 && contentWidthPt > 0) {
    return Math.min(Math.max((printableWidthPt * fitPages) / contentWidthPt, 0.1), 1)
  }
  if (pageSetup.fitToPage !== true && pageSetup.scale !== undefined) {
    return Math.min(Math.max(pageSetup.scale / 100, 0.1), 2)
  }
  return 1
}

const MAX_PAGE_LABELS = 60
const MAX_BOUNDARIES = 200
const MAX_ROWS = 20_000
const MAX_COLUMNS = 2_000

export function installPageBreakPreview(
  runtime: UniverRuntime,
  worksheet: UniverWorksheet,
  pageSetup: PageSetupJournalState,
  breaks: { rowBreaks: readonly number[]; colBreaks: readonly number[] },
  extent: { rows: number; columns: number },
  idPrefix: string,
): Disposable[] {
  const rows = Math.min(Math.max(worksheet.getLastRow() + 1, extent.rows, 1), MAX_ROWS)
  const columns = Math.min(Math.max(worksheet.getLastColumn() + 1, extent.columns, 1), MAX_COLUMNS)
  const page = printablePagePt(pageSetup)
  let contentWidthPt = 0
  for (let column = 0; column < columns; column += 1) {
    contentWidthPt += worksheet.getColumnWidth(column) * 0.75
  }
  const scale = effectiveScale(pageSetup, contentWidthPt, page.width)
  const rowBoundaries = computePageBoundaries(
    (index) => worksheet.getRowHeight(index),
    rows,
    page.height,
    scale,
    breaks.rowBreaks,
  ).slice(0, MAX_BOUNDARIES)
  const colBoundaries = computePageBoundaries(
    (index) => worksheet.getColumnWidth(index),
    columns,
    page.width,
    scale,
    breaks.colBreaks,
  ).slice(0, MAX_BOUNDARIES)
  const disposables: Disposable[] = []
  const layer = (
    key: string,
    startRow: number,
    startColumn: number,
    rowCount: number,
    columnCount: number,
    render: () => React.JSX.Element,
  ): void => {
    disposables.push(runtime.univerAPI.registerComponent(key, render))
    const floating = worksheet.addFloatDomToRange(
      worksheet.getRange(startRow, startColumn, rowCount, columnCount),
      { componentKey: key, allowTransform: false, eventPassThrough: true },
      {},
      key,
    )
    if (floating) disposables.push(floating)
  }
  layer(`${idPrefix}-outline`, 0, 0, rows, columns, () => (
    <div className="page-break-line page-break-outline" />
  ))
  rowBoundaries.forEach((boundary, index) =>
    layer(
      `${idPrefix}-row-${index}`,
      boundary.index,
      0,
      1,
      columns,
      () => (
        <div className={`page-break-line page-break-row${boundary.manual ? ' manual' : ''}`} />
      ),
    ),
  )
  colBoundaries.forEach((boundary, index) =>
    layer(
      `${idPrefix}-column-${index}`,
      0,
      boundary.index,
      rows,
      1,
      () => (
        <div className={`page-break-line page-break-col${boundary.manual ? ' manual' : ''}`} />
      ),
    ),
  )
  const rowEdges = [0, ...rowBoundaries.map((boundary) => boundary.index), rows]
  const colEdges = [0, ...colBoundaries.map((boundary) => boundary.index), columns]
  let pageNumber = 0
  for (
    let columnPage = 0;
    columnPage < colEdges.length - 1 && pageNumber < MAX_PAGE_LABELS;
    columnPage += 1
  ) {
    for (
      let rowPage = 0;
      rowPage < rowEdges.length - 1 && pageNumber < MAX_PAGE_LABELS;
      rowPage += 1
    ) {
      pageNumber += 1
      const label = t('appPageWatermark', { page: pageNumber })
      const startRow = rowEdges[rowPage] ?? 0
      const startColumn = colEdges[columnPage] ?? 0
      layer(
        `${idPrefix}-page-${pageNumber}`,
        startRow,
        startColumn,
        (rowEdges[rowPage + 1] ?? rows) - startRow,
        (colEdges[columnPage + 1] ?? columns) - startColumn,
        () => <div className="page-break-watermark">{label}</div>,
      )
    }
  }
  return disposables
}
