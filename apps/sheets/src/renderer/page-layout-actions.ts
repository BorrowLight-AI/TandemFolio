/**
 * Page Layout commands, header/footer, freeze journaling and PDF export.
 * Extracted from App.tsx; the App component passes a PageLayoutContext built
 * fresh per call so refs and state never go stale. Page-setup edits journal
 * per-sheet print settings; nothing renders in the grid (Univer has no
 * page-layout view), everything lands in the saved file.
 */
import { columnLabel } from '../domain/cell-address'
import {
  isSheetRemoved,
  journalSize,
  type HeaderFooterParts,
  type PageSetupJournalState,
} from './edit-journal'
import type { HeaderFooterResult } from './HeaderFooterDialog'
import { t } from './i18n/locale'
import { buildSheetPrintPayload, type PrintWorksheet } from './print-html'
import { pushWorkbookUndo } from './univer-sync'
import type { LazyWorkbookState, UniverRuntime } from './univer-state'

const PAPER_NAMES: Record<string, string> = {
  1: 'Letter',
  3: 'Tabloid',
  5: 'Legal',
  7: 'Executive',
  8: 'A3',
  9: 'A4',
  11: 'A5',
}

/** The App refs/state the page-layout actions need; built fresh per call. */
export interface PageLayoutContext {
  univerRef: { readonly current: UniverRuntime | null }
  lazyWorkbookRef: { readonly current: LazyWorkbookState | null }
  setMessage: (message: string) => void
  setPendingEdits: (count: number) => void
}

export type WorkbookPageOrientation = 'portrait' | 'landscape'
export type WorkbookPageMargins = 'normal' | 'wide' | 'narrow'
const WORKBOOK_PAPER_SIZES = [1, 3, 5, 7, 8, 9, 11] as const
export type WorkbookPaperSize = (typeof WORKBOOK_PAPER_SIZES)[number]

type PageSetupPresetValues = {
  readonly orientation: WorkbookPageOrientation
  readonly margins: WorkbookPageMargins
  readonly paperSize: WorkbookPaperSize
  readonly scale: number
  readonly fitToWidth: number
  readonly fitToHeight: number
  readonly fitToPage: boolean
  readonly printGridlines: boolean
  readonly printHeadings: boolean
  readonly printArea: string | null
  readonly printTitles: string | null
  readonly header: HeaderFooterParts | null
  readonly footer: HeaderFooterParts | null
}

type PageSetupPresetState = Pick<LazyWorkbookState, 'editJournal'> &
  Partial<Pick<LazyWorkbookState, 'file'>>

function filePageSetupPreset<Field extends keyof PageSetupPresetValues>(
  state: PageSetupPresetState,
  sheetId: string,
  field: Field,
): PageSetupPresetValues[Field] | undefined {
  const sheet = state.file?.sheets.find((candidate) => candidate.id === sheetId)
  const pageSetup = (
    sheet as
      | (typeof sheet & {
          pageSetup?: PageSetupJournalState
        })
      | undefined
  )?.pageSetup
  return pageSetup?.[field] as PageSetupPresetValues[Field] | undefined
}

type PageSetupPatch = Partial<PageSetupPresetValues>
type PageSetupField = keyof PageSetupPresetValues
type PageSetupFieldValue = PageSetupPresetValues[PageSetupField]
type PageSetupSnapshot = ReadonlyArray<{
  readonly field: PageSetupField
  readonly present: boolean
  readonly value: PageSetupFieldValue | undefined
}>

function effectivePageSetupValue<Field extends PageSetupField>(
  state: PageSetupPresetState,
  sheetId: string,
  field: Field,
): PageSetupPresetValues[Field] | undefined {
  const pageSetup = state.editJournal.pageSetup.get(sheetId)
  return Object.prototype.hasOwnProperty.call(pageSetup ?? {}, field)
    ? (pageSetup?.[field] as PageSetupPresetValues[Field] | undefined)
    : filePageSetupPreset(state, sheetId, field)
}

function snapshotPageSetupFields(
  state: PageSetupPresetState,
  sheetId: string,
  fields: readonly PageSetupField[],
): PageSetupSnapshot {
  const pageSetup = state.editJournal.pageSetup.get(sheetId)
  return fields.map((field) => ({
    field,
    present: Object.prototype.hasOwnProperty.call(pageSetup ?? {}, field),
    value: pageSetup?.[field] as PageSetupFieldValue | undefined,
  }))
}

function restorePageSetupFields(
  state: PageSetupPresetState,
  sheetId: string,
  snapshot: PageSetupSnapshot,
  setPendingEdits?: (count: number) => void,
): void {
  const pageSetup = { ...(state.editJournal.pageSetup.get(sheetId) ?? {}) }
  const mutablePageSetup = pageSetup as Record<PageSetupField, PageSetupFieldValue | undefined>
  for (const { field, present, value } of snapshot) {
    if (present && value !== undefined) mutablePageSetup[field] = value
    else delete mutablePageSetup[field]
  }
  if (Object.keys(pageSetup).length > 0) state.editJournal.pageSetup.set(sheetId, pageSetup)
  else state.editJournal.pageSetup.delete(sheetId)
  setPendingEdits?.(journalSize(state.editJournal))
}

function applyWorkbookPageSetupPatch(
  runtime: UniverRuntime,
  state: PageSetupPresetState,
  sheetId: string,
  patch: PageSetupPatch,
  setPendingEdits?: (count: number) => void,
): void {
  const next = Object.entries(patch).map(([field, value]) => ({
    field: field as PageSetupField,
    present: true,
    value: value as PageSetupFieldValue,
  })) satisfies PageSetupSnapshot
  const before = snapshotPageSetupFields(
    state,
    sheetId,
    next.map(({ field }) => field),
  )
  if (
    next.every(({ field, value }) => {
      const current = effectivePageSetupValue(state, sheetId, field)
      return current === value || JSON.stringify(current) === JSON.stringify(value)
    })
  ) {
    return
  }

  restorePageSetupFields(state, sheetId, next, setPendingEdits)
  pushWorkbookUndo(runtime, {
    undo: () => restorePageSetupFields(state, sheetId, before, setPendingEdits),
    redo: () => restorePageSetupFields(state, sheetId, next, setPendingEdits),
  })
}

function applyWorkbookPageSetupPreset<Field extends keyof PageSetupPresetValues>(
  runtime: UniverRuntime,
  state: PageSetupPresetState,
  sheetId: string,
  field: Field,
  value: PageSetupPresetValues[Field],
  setPendingEdits?: (count: number) => void,
): void {
  applyWorkbookPageSetupPatch(
    runtime,
    state,
    sheetId,
    { [field]: value } as PageSetupPatch,
    setPendingEdits,
  )
}

/** Applies an explicit persisted print orientation to one worksheet. */
export function applyWorkbookPageOrientation(
  runtime: UniverRuntime,
  state: PageSetupPresetState,
  sheetId: string,
  orientation: WorkbookPageOrientation,
  setPendingEdits?: (count: number) => void,
): void {
  applyWorkbookPageSetupPreset(runtime, state, sheetId, 'orientation', orientation, setPendingEdits)
}

/** Applies an explicit persisted print-margin preset to one worksheet. */
export function applyWorkbookPageMargins(
  runtime: UniverRuntime,
  state: PageSetupPresetState,
  sheetId: string,
  margins: WorkbookPageMargins,
  setPendingEdits?: (count: number) => void,
): void {
  applyWorkbookPageSetupPreset(runtime, state, sheetId, 'margins', margins, setPendingEdits)
}

/** Applies an explicit persisted paper-size preset to one worksheet. */
export function applyWorkbookPaperSize(
  runtime: UniverRuntime,
  state: PageSetupPresetState,
  sheetId: string,
  paperSize: WorkbookPaperSize,
  setPendingEdits?: (count: number) => void,
): void {
  applyWorkbookPageSetupPreset(runtime, state, sheetId, 'paperSize', paperSize, setPendingEdits)
}

/** Applies explicit persisted fit-to-page width and height as one undo unit. */
export function applyWorkbookFitToPages(
  runtime: UniverRuntime,
  state: PageSetupPresetState,
  sheetId: string,
  widthPages: number,
  heightPages: number,
  setPendingEdits?: (count: number) => void,
): void {
  applyWorkbookPageSetupPatch(
    runtime,
    state,
    sheetId,
    {
      fitToWidth: widthPages,
      fitToHeight: heightPages,
      fitToPage: widthPages > 0 || heightPages > 0,
    },
    setPendingEdits,
  )
}

/** Applies a fixed print scale and disables fit-to-page as one undo unit. */
export function applyWorkbookPrintScale(
  runtime: UniverRuntime,
  state: PageSetupPresetState,
  sheetId: string,
  scalePercent: number,
  setPendingEdits?: (count: number) => void,
): void {
  applyWorkbookPageSetupPatch(
    runtime,
    state,
    sheetId,
    { scale: scalePercent, fitToPage: false },
    setPendingEdits,
  )
}

/** Applies an explicit persisted print-gridline state to one worksheet. */
export function applyWorkbookPrintGridlines(
  runtime: UniverRuntime,
  state: PageSetupPresetState,
  sheetId: string,
  enabled: boolean,
  setPendingEdits?: (count: number) => void,
): void {
  applyWorkbookPageSetupPreset(runtime, state, sheetId, 'printGridlines', enabled, setPendingEdits)
}

/** Applies an explicit persisted print-heading state to one worksheet. */
export function applyWorkbookPrintHeadings(
  runtime: UniverRuntime,
  state: PageSetupPresetState,
  sheetId: string,
  enabled: boolean,
  setPendingEdits?: (count: number) => void,
): void {
  applyWorkbookPageSetupPreset(runtime, state, sheetId, 'printHeadings', enabled, setPendingEdits)
}

/** Applies an explicit persisted print area, or clears it, for one worksheet. */
export function applyWorkbookPrintArea(
  runtime: UniverRuntime,
  state: PageSetupPresetState,
  sheetId: string,
  range: string | null,
  setPendingEdits?: (count: number) => void,
): void {
  applyWorkbookPageSetupPreset(runtime, state, sheetId, 'printArea', range, setPendingEdits)
}

/** Applies explicit persisted print-title rows, or clears them, for one worksheet. */
export function applyWorkbookPrintTitles(
  runtime: UniverRuntime,
  state: PageSetupPresetState,
  sheetId: string,
  rows: string | null,
  setPendingEdits?: (count: number) => void,
): void {
  applyWorkbookPageSetupPreset(runtime, state, sheetId, 'printTitles', rows, setPendingEdits)
}

/** Applies printed header/footer sections as one renderer-owned undo unit. */
export function applyWorkbookHeaderFooter(
  runtime: UniverRuntime,
  state: PageSetupPresetState,
  sheetId: string,
  header: HeaderFooterParts | null,
  footer: HeaderFooterParts | null,
  setPendingEdits?: (count: number) => void,
): void {
  applyWorkbookPageSetupPatch(
    runtime,
    state,
    sheetId,
    { header, footer },
    setPendingEdits,
  )
}

function isWorkbookPaperSize(value: number): value is WorkbookPaperSize {
  return (WORKBOOK_PAPER_SIZES as readonly number[]).includes(value)
}

export function handlePageLayoutCommand(ctx: PageLayoutContext, rest: string): void {
  const runtime = ctx.univerRef.current
  const state = ctx.lazyWorkbookRef.current
  if (!runtime) return
  if (!state) {
    ctx.setMessage(t('appPageSetupNeedsFile'))
    return
  }
  const worksheet = runtime.univerAPI.getActiveWorkbook()?.getActiveSheet()
  const sheetId = worksheet?.getSheetId()
  if (!sheetId || isSheetRemoved(state.editJournal, sheetId)) return
  const separator = rest.indexOf(':')
  const key = separator === -1 ? rest : rest.slice(0, separator)
  const value = separator === -1 ? '' : rest.slice(separator + 1)
  switch (key) {
    case 'orientation':
      if (value !== 'portrait' && value !== 'landscape') return
      applyWorkbookPageOrientation(runtime, state, sheetId, value, ctx.setPendingEdits)
      ctx.setMessage(
        t('appPageSetupRecorded', {
          note: value === 'portrait' ? t('appOrientationPortrait') : t('appOrientationLandscape'),
        }),
      )
      return
    case 'margins':
      if (value !== 'normal' && value !== 'wide' && value !== 'narrow') return
      applyWorkbookPageMargins(runtime, state, sheetId, value, ctx.setPendingEdits)
      ctx.setMessage(
        t('appPageSetupRecorded', {
          note: t(
            value === 'normal'
              ? 'appMarginsNormal'
              : value === 'wide'
                ? 'appMarginsWide'
                : 'appMarginsNarrow',
          ),
        }),
      )
      return
    case 'paper': {
      const code = Number(value)
      if (!isWorkbookPaperSize(code)) return
      applyWorkbookPaperSize(runtime, state, sheetId, code, ctx.setPendingEdits)
      ctx.setMessage(
        t('appPageSetupRecorded', {
          note: t('appPaperSizeNote', { name: PAPER_NAMES[value] ?? value }),
        }),
      )
      return
    }
    case 'scale': {
      const scale = Number(value)
      if (!Number.isInteger(scale) || scale < 10 || scale > 400) return
      applyWorkbookPrintScale(runtime, state, sheetId, scale, ctx.setPendingEdits)
      ctx.setMessage(t('appPageSetupRecorded', { note: t('appPrintScaleNote', { scale }) }))
      return
    }
    case 'fit-width':
    case 'fit-height': {
      const pages = Number(value)
      if (!Number.isInteger(pages) || pages < 0 || pages > 1_000) return
      const effective = (field: 'fitToWidth' | 'fitToHeight'): number =>
        effectivePageSetupValue(state, sheetId, field) ?? 0
      const fitToWidth = key === 'fit-width' ? pages : effective('fitToWidth')
      const fitToHeight = key === 'fit-height' ? pages : effective('fitToHeight')
      const fitValue = pages === 0 ? t('appFitAutomatic') : t('appFitPages', { count: pages })
      applyWorkbookFitToPages(runtime, state, sheetId, fitToWidth, fitToHeight, ctx.setPendingEdits)
      ctx.setMessage(
        t('appPageSetupRecorded', {
          note:
            key === 'fit-width'
              ? t('appFitWidthNote', { value: fitValue })
              : t('appFitHeightNote', { value: fitValue }),
        }),
      )
      return
    }
    case 'print-gridlines': {
      if (value !== '0' && value !== '1') return
      const enabled = value === '1'
      applyWorkbookPrintGridlines(runtime, state, sheetId, enabled, ctx.setPendingEdits)
      ctx.setMessage(
        t('appPageSetupRecorded', {
          note: enabled ? t('appGridlinesWillPrint') : t('appGridlinesWontPrint'),
        }),
      )
      return
    }
    case 'print-headings': {
      if (value !== '0' && value !== '1') return
      const enabled = value === '1'
      applyWorkbookPrintHeadings(runtime, state, sheetId, enabled, ctx.setPendingEdits)
      ctx.setMessage(
        t('appPageSetupRecorded', {
          note: enabled ? t('appHeadingsWillPrint') : t('appHeadingsWontPrint'),
        }),
      )
      return
    }
    case 'print-area': {
      if (value === 'clear') {
        applyWorkbookPrintArea(runtime, state, sheetId, null, ctx.setPendingEdits)
        ctx.setMessage(t('appPageSetupRecorded', { note: t('appPrintAreaCleared') }))
        return
      }
      if (value !== 'set') return
      const range = runtime.univerAPI.getActiveWorkbook()?.getActiveRange()
      if (!range) {
        ctx.setMessage(t('appSelectPrintRange'))
        return
      }
      const startColumn = range.getColumn()
      const startRow = range.getRow()
      const area =
        `${columnLabel(startColumn)}${startRow + 1}` +
        `:${columnLabel(startColumn + range.getWidth() - 1)}${startRow + range.getHeight()}`
      applyWorkbookPrintArea(runtime, state, sheetId, area, ctx.setPendingEdits)
      ctx.setMessage(t('appPageSetupRecorded', { note: t('appPrintAreaNote', { area }) }))
      return
    }
    case 'print-titles': {
      if (value === 'clear') {
        applyWorkbookPrintTitles(runtime, state, sheetId, null, ctx.setPendingEdits)
        ctx.setMessage(t('appPageSetupRecorded', { note: t('appPrintTitlesCleared') }))
        return
      }
      if (value === 'first-row') {
        applyWorkbookPrintTitles(runtime, state, sheetId, '1:1', ctx.setPendingEdits)
        ctx.setMessage(t('appPageSetupRecorded', { note: t('appRow1Repeats') }))
        return
      }
      if (value !== 'set') return
      const range = runtime.univerAPI.getActiveWorkbook()?.getActiveRange()
      if (!range) {
        ctx.setMessage(t('appSelectRepeatRows'))
        return
      }
      if (range.getHeight() > 21) {
        ctx.setMessage(t('appPrintTitlesLimit'))
        return
      }
      const rows = `${range.getRow() + 1}:${range.getRow() + range.getHeight()}`
      applyWorkbookPrintTitles(runtime, state, sheetId, rows, ctx.setPendingEdits)
      ctx.setMessage(t('appPageSetupRecorded', { note: t('appRowsRepeat', { rows }) }))
      return
    }
    default:
      return
  }
}

/// Insert → Header & Footer OK: journals the printed header/footer of the
/// active sheet; the save writes the worksheet's <headerFooter> element.
export function handleApplyHeaderFooter(
  ctx: PageLayoutContext,
  result: HeaderFooterResult,
): string | null {
  const state = ctx.lazyWorkbookRef.current
  const runtime = ctx.univerRef.current
  if (!state || !runtime) return t('appHfNeedsFile')
  const sheetId = runtime.univerAPI.getActiveWorkbook()?.getActiveSheet()?.getSheetId()
  if (!sheetId || isSheetRemoved(state.editJournal, sheetId)) {
    return t('appActiveSheetUnavailable')
  }
  applyWorkbookHeaderFooter(
    runtime,
    state,
    sheetId,
    result.header,
    result.footer,
    ctx.setPendingEdits,
  )
  ctx.setMessage(t('appHfUpdated'))
  return null
}

/// Lays the active sheet out as HTML with its Page Layout settings and asks
/// the main process to render the PDF (hidden window + save dialog).
export async function handleExportPdf(ctx: PageLayoutContext): Promise<void> {
  const runtime = ctx.univerRef.current
  const worksheet = runtime?.univerAPI.getActiveWorkbook()?.getActiveSheet()
  if (!runtime || !worksheet) return
  const state = ctx.lazyWorkbookRef.current
  if (state && !state.flags.preloadComplete) {
    ctx.setMessage(t('appPdfNeedsFullLoad'))
    return
  }
  try {
    const pageSetup = state?.editJournal.pageSetup.get(worksheet.getSheetId()) ?? {}
    const baseName = (state?.file.name ?? 'Book1').replace(/\.[^.]+$/, '')
    const payload = buildSheetPrintPayload(
      worksheet as unknown as PrintWorksheet,
      pageSetup,
      `${baseName}.pdf`,
      worksheet.getSheetName(),
    )
    ctx.setMessage(t('appPdfRendering'))
    const result = await window.desktopApi.exportPdf(payload)
    ctx.setMessage(
      result.canceled ? t('appPdfCanceled') : t('appPdfExported', { path: result.path }),
    )
  } catch (error: unknown) {
    ctx.setMessage(error instanceof Error ? error.message : t('appPdfExportFailed'))
  }
}
