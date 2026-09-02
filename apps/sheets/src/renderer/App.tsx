import {
  absRangeRef,
  activateFormulaClosure,
  applyDefinedNames,
  applyWorkbookNotes,
  cellValueBounds,
  clearLazyState,
  columnLetter,
  disposeVisuals,
  ensureLazyRangeLoaded,
  journalRangeSnapshot,
  loadSnapshotIntoUniver,
  loadVisibleRange,
  loadWorkbookSkeleton,
  matrixBounds,
  navigateToAnchor,
  preloadEntireWorkbook,
  queueFormulaRecalc,
  queueSparklineInstall,
  RECALC_MAX_FAILURES,
  queueVisualInstall,
  sheetOutline,
  univerDefinedNames,
} from './univer-sync'
import {
  journalSuppression,
  type ActiveWorkbook,
  type LazyWorkbookState,
  type UniverRuntime,
} from './univer-state'
import {
  applyAiPivotAdd,
  applyAiTableColumnAdd,
  applyAiTableColumnDelete,
  applyAiTableRowAdd,
  applyAiTableRowDelete,
  renameChartRefsForSheet,
} from './workbook-ops'
import { isNumericIdentifierText } from './cell-warning'
import { useCallback, useEffect, useRef, useState, useSyncExternalStore } from 'react'

import {
  attachMcpLiveSession,
  getLiveEditorActivity,
  getLiveEditorDisplayMode,
  subscribeLiveEditorActivity,
  subscribeLiveEditorDisplayMode,
  toggleLiveEditorFullscreen,
  type LiveEditorAdapter,
  type LiveEditorStartupTrace,
} from '@tandemfolio/host-bridge'

import {
  CellValueType,
  getNumfmtParseValueFilter,
  InterceptorEffectEnum,
  isRealNum,
  IUndoRedoService,
  LocaleType,
  mergeLocales,
  ThemeService,
  type IRange,
  type IStyleData,
} from '@univerjs/core'
import { UniverSheetsConditionalFormattingPreset } from '@univerjs/preset-sheets-conditional-formatting'
import UniverPresetSheetsConditionalFormattingEnUS from '@univerjs/preset-sheets-conditional-formatting/locales/en-US'
import '@univerjs/preset-sheets-conditional-formatting/lib/index.css'
import {
  INTERCEPTOR_POINT,
  SheetInterceptorService,
  UniverSheetsCorePreset,
} from '@univerjs/preset-sheets-core'
import UniverPresetSheetsCoreEnUS from '@univerjs/preset-sheets-core/locales/en-US'
import { UniverSheetsDataValidationPreset } from '@univerjs/preset-sheets-data-validation'
import UniverPresetSheetsDataValidationEnUS from '@univerjs/preset-sheets-data-validation/locales/en-US'
import '@univerjs/preset-sheets-data-validation/lib/index.css'
import { UniverSheetsDrawingPreset } from '@univerjs/preset-sheets-drawing'
import '@univerjs/preset-sheets-drawing/lib/index.css'
import { UniverSheetsFindReplacePreset } from '@univerjs/preset-sheets-find-replace'
import UniverPresetSheetsFindReplaceEnUS from '@univerjs/preset-sheets-find-replace/locales/en-US'
import '@univerjs/preset-sheets-find-replace/lib/index.css'
import { UniverSheetsFilterPreset } from '@univerjs/preset-sheets-filter'
import UniverPresetSheetsFilterEnUS from '@univerjs/preset-sheets-filter/locales/en-US'
import '@univerjs/preset-sheets-filter/lib/index.css'
import { UniverSheetsNotePreset } from '@univerjs/preset-sheets-note'
import UniverPresetSheetsNoteEnUS from '@univerjs/preset-sheets-note/locales/en-US'
import '@univerjs/preset-sheets-note/lib/index.css'
import { UniverSheetsSortPreset } from '@univerjs/preset-sheets-sort'
import UniverPresetSheetsSortEnUS from '@univerjs/preset-sheets-sort/locales/en-US'
import '@univerjs/preset-sheets-sort/lib/index.css'
import { UniverSheetsTablePreset } from '@univerjs/preset-sheets-table'
import UniverPresetSheetsTableEnUS from '@univerjs/preset-sheets-table/locales/en-US'
import '@univerjs/preset-sheets-table/lib/index.css'
import { greenTheme } from '@univerjs/themes'
import { createUniver } from './create-univer'
import { BrowserWorkbookDesktopApi } from './browser-desktop-api'

import { columnIndex, columnLabel, parseRange } from '../domain/cell-address'
import {
  applyChartStateEdit,
  chartSupportsDataLabels,
  chartSupportsSeriesReplace,
  withDefaultBarLabels,
  type CellBounds,
} from '../domain/chart-visual'
import { InMemoryWorkbookAdapter } from '../domain/in-memory-workbook'
import { cfRuleUnsaveableReason, iconSetSaveable } from '../gateway/xlsx-cf'
import type { MenuAction, WorkbookFile, WorkbookVisualObject } from '../shared/desktop-api'
import type { PageSetupJournalState, StructuralJournalOp } from './edit-journal'
import {
  AUTO_FILL_COMMAND,
  AXIS_ATTR_MUTATIONS,
  BLOCKED_COMMAND_PATTERN,
  CF_MUTATIONS,
  CF_RULE_COMMAND_PATTERN,
  COPY_SHEET_COMMAND,
  DEFINED_NAME_MUTATIONS,
  DV_EDIT_COMMAND_PATTERN,
  DV_MUTATIONS,
  EMPTY_CHART_EDITS,
  FILTER_COMMAND_PATTERN,
  FILTER_MUTATIONS,
  FORMULA_MODE_MAX_CELLS,
  initialSnapshot,
  MERGE_MUTATIONS,
  MOVE_RANGE_COMMAND,
  MOVE_ROWS_COMMAND,
  MOVE_ROWS_MUTATION,
  MOVE_RANGE_MUTATION,
  NOTE_MUTATIONS,
  pixelsToCharacterWidth,
  REMOVE_NUMFMT_MUTATION,
  REORDER_RANGE_MUTATION,
  ROW_COLUMN_MUTATIONS,
  SET_FROZEN_MUTATION,
  SET_NUMFMT_MUTATION,
  SET_TAB_COLOR_MUTATION,
  TOGGLE_GRIDLINES_MUTATION,
  SET_RANGE_VALUES_MUTATION,
  SHEET_LIFECYCLE_MUTATIONS,
  SORT_COMMAND_PATTERN,
  STRUCTURAL_EDIT_COMMAND_PATTERN,
} from './app-constants'
import {
  getSourceRange as getSourceRangeImpl,
  handleCreatePivot as handleCreatePivotImpl,
  handleCreateSlicer as handleCreateSlicerImpl,
  handleEditPivotApply as handleEditPivotApplyImpl,
  handleRefreshPivot as handleRefreshPivotImpl,
  handleCreateTimeline as handleCreateTimelineImpl,
  handleRemoveSlicer as handleRemoveSlicerImpl,
  handleRemoveTimeline as handleRemoveTimelineImpl,
  handleSlicerSelectAll as handleSlicerSelectAllImpl,
  handleSlicerToggle as handleSlicerToggleImpl,
  handleTimelineRange as handleTimelineRangeImpl,
  isSelectionInPivot as isSelectionInPivotImpl,
  pivotEditInitial as pivotEditInitialImpl,
  pivotFieldOptions as pivotFieldOptionsImpl,
  refreshPivotTables as refreshPivotTablesImpl,
  setPivotMemberFilter as setPivotMemberFilterImpl,
  updatePivotLayout as updatePivotLayoutImpl,
  type PivotActionContext,
  type PivotEditContext,
  type SlicerPickerState,
  type TimelinePickerState,
} from './pivot-actions'
import type { ChartRecommendations } from '../domain/chart-recommend'
import {
  handleInsertChart as handleInsertChartImpl,
  handleInsertEquation as handleInsertEquationImpl,
  handleInsertIcon as handleInsertIconImpl,
  handleInsertScreenshot,
  handleRecommendedCharts as handleRecommendedChartsImpl,
  insertAiChartVisual,
  insertAiImageVisual,
  insertAiPivotChartVisual,
  insertAiShapeVisual,
  type VisualActionContext,
} from './visual-actions'
import {
  activeCellLabel as activeCellLabelImpl,
  applyWorkbookConsolidation,
  applyWorkbookSubtotals,
  consolidateDefaultReference as consolidateDefaultReferenceImpl,
  goToReference as goToReferenceImpl,
  handleApplyAdvancedFilter as handleApplyAdvancedFilterImpl,
  handleApplyFormula as handleApplyFormulaImpl,
  handleCreateConsolidate as handleCreateConsolidateImpl,
  handleCreateSubtotal as handleCreateSubtotalImpl,
  handleInsertSymbol as handleInsertSymbolImpl,
  listDefinedNames as listDefinedNamesImpl,
  type DataToolsContext,
} from './data-tools-actions'
import { installTsvClipboardFix } from './clipboard-tsv'
import { installFilteredCopyHook } from './filtered-copy'
import {
  applyShowFormulasView,
  installFormulaTextInterceptor,
  installFormulaViewInterceptor,
} from './formula-view'
import { installCellFilenameFunction } from './cell-function'
import { installFormulaLexerFix } from './formula-lexer-fix'
import { installSheetRenameFix } from './sheet-rename-fix'
import { installSelectionWrapGuard } from './selection-wrap-fix'
import { installMultiRowAutofit } from './autofit-multi-row'
import { installCopyMaterialize } from './copy-materialize'
import { applyUniverLocale } from './univer-locales'
import { installRuleDetail } from './univer-rule-detail'
import { installPopulatedDataValidationArrow } from './data-validation-arrow'
import { installFormulaNullResultFix } from './formula-null-result'
import { installNumberFormatFix } from './numfmt-fix'
import { installRateFallback } from './rate-function'
import {
  handleRibbonCommand as handleRibbonCommandImpl,
  type RibbonCommandContext,
} from './ribbon-actions'
import {
  handleApplyHeaderFooter as handleApplyHeaderFooterImpl,
  handleExportPdf as handleExportPdfImpl,
  handlePageLayoutCommand as handlePageLayoutCommandImpl,
  type PageLayoutContext,
} from './page-layout-actions'
import {
  handleSave as handleSaveImpl,
  type SaveActionResult,
  type SaveContext,
} from './save-actions'
import { executeXlsxOperation } from './operations/registry'
import { listWorkbookConditionalFormats } from './conditional-format-actions'
import {
  applyChartEdit as applyChartEditImpl,
  applyShapeEdit as applyShapeEditImpl,
  queueChartDataSync as queueChartDataSyncImpl,
  readChartVector as readChartVectorImpl,
  type VisualSyncContext,
} from './visual-edit-sync'

import {
  createEditJournal,
  isSheetRemoved,
  hyperlinkEditAt,
  journalSize,
  recordCfChange,
  recordPageSetup,
  recordDefinedNamesChange,
  recordDvChange,
  recordNoteChange,
  recordSheetProtection,
  recordFilterChange,
  recordSetNumfmt,
  recordSetRangeValues,
  recordSheetHidden,
  recordSheetDuplicate,
  recordSheetInsert,
  recordSheetOrderChange,
  recordSheetRemove,
  recordSheetRename,
  recordStructuralOp,
  removeTableAdd,
  shiftVisualForStructuralOp,
} from './edit-journal'
import { shiftPinnedCells } from './formula-closure'
import { getLang, t } from './i18n/locale'
import { netAxisDelta, screenToFile } from './view-transform'
import { selectionFormatEquals, toSelectionFormat, type SelectionFormat } from './selection-format'
import { applyWorkbookTextReplacement } from './text-replace'
import { ExcelShell } from './ExcelShell'
import { ToastHost } from './toast'
import { AdvancedFilterDialog, type AdvancedFilterColumn } from './AdvancedFilterDialog'
import { EquationDialog } from './EquationDialog'
import { IconsDialog } from './IconsDialog'
import { RecommendedChartsDialog } from './RecommendedChartsDialog'
import { ScreenshotDialog } from './ScreenshotDialog'
import { SymbolDialog } from './SymbolDialog'
import { SlicerFieldPicker, SlicerPanels, type SlicerUiState } from './SlicerPanel'
import { TimelineFieldPicker, TimelinePanels, type TimelineUiState } from './TimelinePanel'
import type { DefinedNameAction, DefinedNameRow } from './NameManagerDialog'
import { applyWorkbookDefinedNameAction } from './defined-name-actions'
import {
  clearVisualSelection,
  convertibleType,
  getChartElementSelection,
  installWorkbookVisuals,
  isVisualDragActive,
  setChartDialogListener,
  setVisualSelectionListener,
  subscribeChartElementSelection,
  type ChartDialogKind,
  type ChartEditData,
  type ChartVectorRead,
  type ShapeEditChanges,
} from './WorkbookVisuals'
import { ChartFormatPane, SelectDataDialog } from './ChartPanels'

// Source sheet id of an in-flight copy-sheet command; the next insert-sheet
// mutation is that copy and must journal as a duplicate, not a blank add.
let pendingCopySource: string | undefined

async function stagedImagePayload(
  data: ArrayBuffer,
  mediaType: 'image/png' | 'image/jpeg' | 'image/gif',
): Promise<{ dataUrl: string; width: number; height: number }> {
  const blob = new Blob([data], { type: mediaType })
  const dataUrl = await new Promise<string>((resolve, reject) => {
    const reader = new FileReader()
    reader.onerror = () => reject(new Error('The staged image could not be decoded.'))
    reader.onload = () =>
      typeof reader.result === 'string'
        ? resolve(reader.result)
        : reject(new Error('The staged image could not be decoded.'))
    reader.readAsDataURL(blob)
  })
  const dimensions = await new Promise<{ width: number; height: number }>((resolve) => {
    const image = new Image()
    image.onload = () => resolve({ width: image.naturalWidth, height: image.naturalHeight })
    image.onerror = () => resolve({ width: 480, height: 320 })
    image.src = dataUrl
  })
  return { dataUrl, ...dimensions }
}

function decodeImageBase64(base64: string): ArrayBuffer {
  const binary = atob(base64)
  const bytes = new Uint8Array(binary.length)
  for (let index = 0; index < binary.length; index += 1) bytes[index] = binary.charCodeAt(index)
  return bytes.buffer
}

export function App(): React.JSX.Element {
  const adapterRef = useRef(new InMemoryWorkbookAdapter(initialSnapshot))
  const univerRef = useRef<UniverRuntime | null>(null)
  const lazyWorkbookRef = useRef<LazyWorkbookState | null>(null)
  const initialWorkbookLoadRef = useRef<Promise<void>>(Promise.resolve())
  const workbookReadyRef = useRef<Promise<void>>(Promise.resolve())
  const startupTraceRef = useRef<{
    promise: Promise<LiveEditorStartupTrace>
    resolve: (trace: LiveEditorStartupTrace) => void
  } | null>(null)
  if (!startupTraceRef.current) {
    let resolve!: (trace: LiveEditorStartupTrace) => void
    const promise = new Promise<LiveEditorStartupTrace>((complete) => {
      resolve = complete
    })
    startupTraceRef.current = { promise, resolve }
  }
  /// Univer undo/redo stack occupancy (subscribed at mount): drives the QAT button gray states
  const [univerHist, setUniverHist] = useState({ canUndo: false, canRedo: false })
  /// True while Univer's in-cell editor is open (AutoSave must not save-reload then).
  const editingCellRef = useRef(false)
  const visualDisposablesRef = useRef<{ dispose(): void }[]>([])
  const traceArrowsRef = useRef<{ disposables: { dispose(): void }[]; nextId: number }>({
    disposables: [],
    nextId: 0,
  })
  const visualInstallTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null)
  const sparklineDisposablesRef = useRef<{ dispose(): void }[]>([])
  const sparklineTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null)
  const visualViewportKeyRef = useRef('')
  const canvasBackingStoresRef = useRef(
    new Map<HTMLCanvasElement, { width: number; height: number }>(),
  )
  const demoVisualDisposablesRef = useRef<{ dispose(): void }[]>([])
  const demoVisualInstallTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null)
  const [_revision, setRevision] = useState(0)
  const [pendingEdits, setPendingEditsState] = useState(0)
  const recoveryVersionRef = useRef(0)
  const setPendingEdits = useCallback((count: number) => {
    if (count > 0) recoveryVersionRef.current += 1
    setPendingEditsState(count)
  }, [])
  const fileInputRef = useRef<HTMLInputElement>(null)
  const pendingFilePickerRef = useRef<((file: File | null) => void) | null>(null)
  const browserHostRef = useRef<BrowserWorkbookDesktopApi | null>(null)
  if (!browserHostRef.current) {
    browserHostRef.current = new BrowserWorkbookDesktopApi(
      () =>
        new Promise<File | null>((resolve) => {
          pendingFilePickerRef.current?.(null)
          pendingFilePickerRef.current = resolve
          fileInputRef.current?.click()
        }),
    )
    Object.defineProperty(window, 'desktopApi', {
      configurable: true,
      value: browserHostRef.current,
    })
  }
  const display = useSyncExternalStore(
    subscribeLiveEditorDisplayMode,
    getLiveEditorDisplayMode,
    getLiveEditorDisplayMode,
  )
  const editorActive = useSyncExternalStore(
    subscribeLiveEditorActivity,
    getLiveEditorActivity,
    () => true,
  )

  useEffect(() => {
    const container = document.getElementById('univer-container')
    if (!container) return
    const backingStores = canvasBackingStoresRef.current
    if (editorActive) {
      for (const [canvas, size] of backingStores) {
        if (!canvas.isConnected) continue
        canvas.width = size.width
        canvas.height = size.height
      }
      backingStores.clear()
      window.requestAnimationFrame(() => window.dispatchEvent(new Event('resize')))
      return
    }
    const releaseBackingStores = () => {
      for (const canvas of container.querySelectorAll('canvas')) {
        if (canvas.width === 0 && canvas.height === 0) continue
        if (!backingStores.has(canvas)) {
          backingStores.set(canvas, { width: canvas.width, height: canvas.height })
        }
        canvas.width = 0
        canvas.height = 0
      }
    }
    releaseBackingStores()
    const observer = new MutationObserver(releaseBackingStores)
    observer.observe(container, {
      attributes: true,
      attributeFilter: ['width', 'height'],
      childList: true,
      subtree: true,
    })
    return () => observer.disconnect()
  }, [editorActive])

  useEffect(() => {
    const adapter: LiveEditorAdapter = {
      startupTrace: () => startupTraceRef.current!.promise,
      snapshot: (revision) => {
        const workbook = univerRef.current?.univerAPI.getActiveWorkbook()
        const activeRange = workbook?.getActiveRange()
        let style: ReturnType<typeof toSelectionFormat> | undefined
        if (activeRange) {
          try {
            style = toSelectionFormat(
              activeRange.getCellStyleData() ?? {},
              activeRange.getNumberFormat(),
            )
          } catch {
            // The workbook can briefly dispose between polling and reading.
          }
        }
        const state = lazyWorkbookRef.current
        const visibleVisuals = state
          ? [...state.file.visuals, ...state.editJournal.visualAdds]
              .filter((visual) => !state.editJournal.visualEdits.get(visual.id)?.remove)
              .slice(0, 100)
              .map((visual) => ({
                id: visual.id,
                sheet: state.file.sheets.find((sheet) => sheet.id === visual.sheetId)?.name ?? null,
                kind: visual.kind,
                ...(visual.chart === undefined
                  ? {}
                  : { title: visual.chart.title, chartTypes: visual.chart.chartTypes }),
              }))
          : []
        let conditionalFormats: ReturnType<typeof listWorkbookConditionalFormats> = []
        try {
          conditionalFormats = listWorkbookConditionalFormats(workbook ?? null)
        } catch {
          // The workbook can briefly dispose between polling and reading.
        }
        return {
          revision,
          fileName: state?.file.name ?? null,
          dirty: state ? journalSize(state.editJournal) > 0 : false,
          conditionalFormats,
          selection: activeRange
            ? {
                sheet: activeRange.getSheetName(),
                range: activeRange.getA1Notation(),
                sheets: workbook?.getSheets().map((sheet) => sheet.getSheetName()) ?? [],
                visuals: visibleVisuals,
                activeCell: {
                  address: activeRange.getA1Notation().split(':')[0],
                  value: activeRange.getValue(),
                  ...(style ? { style } : {}),
                },
              }
            : null,
        }
      },
      recoverySnapshot: async () => {
        if (!lazyWorkbookRef.current || journalSize(lazyWorkbookRef.current.editJournal) === 0) {
          return null
        }
        const recovery = await handleSaveRef.current('recovery')
        return recovery?.ok && 'data' in recovery
          ? { fileName: recovery.fileName, data: recovery.data }
          : null
      },
      recoveryVersion: () => recoveryVersionRef.current,
      execute: async (command) => {
        try {
          if (
            command.operation !== 'open_local_file' &&
            command.operation !== 'xlsx.document.load_staged'
          ) {
            await workbookReadyRef.current
          }
          const registered = await executeXlsxOperation(command, {
            runtime: () => univerRef.current,
            state: () => lazyWorkbookRef.current,
            setPendingEdits,
            isSheetDataComplete: (sheetId) => {
              const state = lazyWorkbookRef.current
              return (
                !state ||
                state.editJournal.sheets.added.has(sheetId) ||
                (state.formulaMode && state.flags.preloadComplete)
              )
            },
            duplicateSheet: (worksheet, name) => {
              const workbook = univerRef.current?.univerAPI.getActiveWorkbook()
              if (!workbook) throw new Error('Open an XLSX workbook first.')
              const duplicated = workbook.duplicateSheet(worksheet)
              duplicated.setName(name)
              return duplicated
            },
            mutateTable: async (operation) => {
              const runtime = univerRef.current
              const state = lazyWorkbookRef.current
              const workbook = runtime?.univerAPI.getActiveWorkbook()
              if (!runtime || !state || !workbook) {
                throw new Error('Open a file-backed XLSX workbook first.')
              }
              if (operation.op === 'add_table_row') {
                applyAiTableRowAdd(runtime, state, operation)
              } else if (operation.op === 'delete_table_row') {
                applyAiTableRowDelete(runtime, state, operation)
              } else if (operation.op === 'add_table_column') {
                applyAiTableColumnAdd(runtime, state, operation)
              } else if (operation.op === 'delete_table_column') {
                applyAiTableColumnDelete(runtime, state, operation)
              } else {
                const table = workbook.getTableInfoByName(operation.tableName)
                if (table && table.subUnitId === operation.sheetId) {
                  await workbook.removeTable(table.id)
                }
                if (!removeTableAdd(state.editJournal, operation.sheetId, operation.tableName)) {
                  throw new Error(`Unknown session-created table: ${operation.tableName}`)
                }
              }
              setPendingEdits(journalSize(state.editJournal))
            },
            replaceText: ({ sheetId, range, find, replace, matchCase, wholeCell }) => {
              const runtime = univerRef.current
              const worksheet = runtime?.univerAPI.getActiveWorkbook()?.getSheetBySheetId(sheetId)
              if (!runtime || !worksheet) throw new Error('Open an XLSX workbook first.')
              return applyWorkbookTextReplacement(runtime, worksheet.getRange(range), {
                find,
                replace,
                matchCase,
                wholeCell,
              })
            },
            createSubtotals: ({ sheetId, range, groupColumn, valueColumn, aggregation }) => {
              const worksheet = univerRef.current?.univerAPI
                .getActiveWorkbook()
                ?.getSheetBySheetId(sheetId)
              if (!worksheet) throw new Error('Open an XLSX workbook first.')
              return applyWorkbookSubtotals(worksheet, worksheet.getRange(range), {
                groupCol: groupColumn,
                valueCol: valueColumn,
                agg: aggregation,
              })
            },
            consolidate: ({ targetSheetId, targetCell, sources, aggregation, leftLabels }) => {
              const workbook = univerRef.current?.univerAPI.getActiveWorkbook()
              const targetWorksheet = workbook?.getSheetBySheetId(targetSheetId)
              if (!workbook || !targetWorksheet) throw new Error('Open an XLSX workbook first.')
              const target = parseRange(targetCell)
              return applyWorkbookConsolidation(
                targetWorksheet,
                { row: target.startRow, column: target.startColumn },
                sources.map((source) => {
                  const worksheet = workbook.getSheetBySheetId(source.sheetId)
                  if (!worksheet) throw new Error(`Unknown XLSX worksheet: ${source.sheetId}`)
                  const range = parseRange(source.range)
                  return {
                    worksheet,
                    area: {
                      sheetName: worksheet.getSheetName(),
                      startRow: range.startRow,
                      startColumn: range.startColumn,
                      rows: range.endRow - range.startRow + 1,
                      columns: range.endColumn - range.startColumn + 1,
                    },
                  }
                }),
                aggregation,
                leftLabels,
              )
            },
            readSourceFormulaText: async ({ sheetId, range }) => {
              const state = lazyWorkbookRef.current
              if (!state) return new Map()
              const result = await window.desktopApi.readWorkbookRange({
                sessionId: state.file.sessionId,
                sheetId,
                range,
              })
              return new Map(
                result.cells.flatMap((cell) =>
                  cell.formula ? [[`${cell.row}:${cell.column}`, cell.formula] as const] : [],
                ),
              )
            },
            refreshSparklines: (sheetId) => {
              const runtime = univerRef.current
              if (!runtime) return
              queueSparklineInstall(
                runtime,
                lazyWorkbookRef,
                sparklineDisposablesRef,
                sparklineTimerRef,
                sheetId,
              )
            },
            addChart: async (operation) => {
              const runtime = univerRef.current
              const state = lazyWorkbookRef.current
              if (!runtime || !state) throw new Error('Open a file-backed XLSX workbook first.')
              const visualId = await insertAiChartVisual(visualContext(), runtime, state, operation)
              setPendingEdits(journalSize(state.editJournal))
              return visualId
            },
            addPivot: (operation) => {
              const runtime = univerRef.current
              const state = lazyWorkbookRef.current
              if (!runtime || !state) throw new Error('Open a file-backed XLSX workbook first.')
              const name = applyAiPivotAdd(runtime, state, operation)
              if (!name) throw new Error('The XLSX PivotTable was not created.')
              setPendingEdits(journalSize(state.editJournal))
              return name
            },
            addPivotChart: ({ pivotId, type }) =>
              insertAiPivotChartVisual(visualContext(), pivotId, type),
            refreshPivots: (sheetId) => refreshPivotTablesImpl(pivotContext(), sheetId),
            updatePivot: (input) => updatePivotLayoutImpl(pivotContext(), input),
            setPivotMemberFilter: (input) => setPivotMemberFilterImpl(pivotContext(), input),
            addShape: async (operation) => {
              const runtime = univerRef.current
              const state = lazyWorkbookRef.current
              if (!runtime || !state) throw new Error('Open a file-backed XLSX workbook first.')
              const visualId = insertAiShapeVisual(visualContext(), runtime, state, operation)
              setPendingEdits(journalSize(state.editJournal))
              return visualId
            },
            updateShape: (shapeId, changes) => {
              shapeEditRef.current(shapeId, changes)
            },
            moveVisual: (visualId, anchor) => {
              shapeEditRef.current(visualId, { anchor })
            },
            editChart: (editKey, edit) => {
              chartEditRef.current(editKey, edit)
            },
            removeVisual: (visualId) => {
              shapeEditRef.current(visualId, { remove: true })
            },
            addImage: async ({ sheetId, anchorCell, name, mediaType, data }) => {
              const runtime = univerRef.current
              const state = lazyWorkbookRef.current
              if (!runtime || !state) throw new Error('Open a file-backed XLSX workbook first.')
              const image = await stagedImagePayload(data, mediaType)
              const visualId = insertAiImageVisual(
                visualContext(),
                runtime,
                state,
                { op: 'add_image', sheetId, anchorCell, path: name },
                { ...image, mediaType },
              )
              setPendingEdits(journalSize(state.editJournal))
              return visualId
            },
            addImagePath: async ({ sheetId, anchorCell, path }) => {
              const runtime = univerRef.current
              const state = lazyWorkbookRef.current
              if (!runtime || !state) throw new Error('Open a file-backed XLSX workbook first.')
              const loaded = await window.desktopApi.readLocalImage({ path })
              const data = decodeImageBase64(loaded.base64)
              const image = await stagedImagePayload(data, loaded.mediaType)
              const visualId = insertAiImageVisual(
                visualContext(),
                runtime,
                state,
                { op: 'add_image', sheetId, anchorCell, path },
                { ...image, mediaType: loaded.mediaType },
              )
              setPendingEdits(journalSize(state.editJournal))
              return visualId
            },
            loadStaged: async ({ name, data }) => {
              const opened = await browserHostRef.current!.openBuffer(data, name)
              await trackWorkbookOpen(opened)
            },
            save: async () => {
              const saved = await handleSaveRef.current('save-as')
              return saved ?? { ok: false, message: t('appSaveFailed') }
            },
          })
          if (registered.handled) {
            if (!registered.ok) {
              return {
                ok: false,
                error: registered.error,
                message: registered.message,
              }
            }
            await new Promise<void>((resolve) => requestAnimationFrame(() => resolve()))
            return registered.output ? { ok: true, output: { ...registered.output } } : { ok: true }
          }

          return {
            ok: false,
            error: 'unsupported_operation',
            message: `Unsupported XLSX operation: ${command.operation}`,
          }
        } catch (error) {
          return {
            ok: false,
            error: 'execution_failed',
            message: error instanceof Error ? error.message : String(error),
          }
        }
      },
    }
    return attachMcpLiveSession(adapter)
  }, [])
  // The close guard lives in the main process; keep it fed with the badge count.
  useEffect(() => {
    window.desktopApi?.notifyPendingEdits?.(pendingEdits)
  }, [pendingEdits])
  const [autoSave, setAutoSave] = useState(() => localStorage.getItem('sheets-auto-save') === '1')
  useEffect(() => {
    localStorage.setItem('sheets-auto-save', autoSave ? '1' : '0')
  }, [autoSave])
  // AutoSave tick (docs/slides parity): every 30 s and on window blur, flush
  // pending edits of the open workbook. The journal is read at tick time so
  // the interval stays stable; demo mode has no backing file and is skipped.
  useEffect(() => {
    if (!autoSave) return
    let saving = false
    const tick = () => {
      const state = lazyWorkbookRef.current
      if (saving || !state || journalSize(state.editJournal) === 0) return
      // Never while the in-cell editor is open (saving reloads the workbook
      // and would wipe the edit), and never for converted .xls/.csv imports
      // whose first save opens a Save As dialog.
      if (editingCellRef.current || state.file.needsSaveAs) return
      saving = true
      void handleSaveRef.current('save', true).finally(() => {
        saving = false
      })
    }
    const id = window.setInterval(tick, 30_000)
    window.addEventListener('blur', tick)
    return () => {
      window.clearInterval(id)
      window.removeEventListener('blur', tick)
    }
  }, [autoSave])

  // Crash-recovery copy: independent of the AutoSave pill — a dirty
  // workbook gets a real .xlsx copy under userData every 30 s, so a force-quit or a
  // renderer crash no longer costs everything since the last manual save. A normal
  // save removes the copy; reopening a file whose copy is newer offers Restore.
  useEffect(() => {
    let writing = false
    const tick = () => {
      const state = lazyWorkbookRef.current
      if (writing || !state || journalSize(state.editJournal) === 0) return
      // The in-cell editor's pending text is not in the journal yet, and a
      // converted import has no original file to recover into.
      if (editingCellRef.current || state.file.needsSaveAs) return
      writing = true
      void handleSaveRef.current('recovery').finally(() => {
        writing = false
      })
    }
    const id = window.setInterval(tick, 30_000)
    return () => window.clearInterval(id)
  }, [])
  const [message, setMessage] = useState(t('appReadyInitial'))
  /// Zoom of the active sheet in percent, echoed by the status-bar slider.
  const [zoomPercent, setZoomPercent] = useState(100)
  const [selectionFormat, setSelectionFormat] = useState<SelectionFormat | null>(null)
  /// A1 label of the active cell, echoed live by the Name Box. Updated from
  /// the same SelectionChanged refresh that keeps selectionFormat current.
  const [activeCellA1, setActiveCellA1] = useState('')
  /// Non-null while the Advanced Filter dialog is open: the column choices
  /// sampled from the active filter range's header row.
  const [advancedFilterColumns, setAdvancedFilterColumns] = useState<
    readonly AdvancedFilterColumn[] | null
  >(null)
  /// True while the Insert → Symbol dialog is open.
  const [symbolDialogOpen, setSymbolDialogOpen] = useState(false)
  const [screenshotDialogOpen, setScreenshotDialogOpen] = useState(false)
  const [iconsDialogOpen, setIconsDialogOpen] = useState(false)
  const [equationDialogOpen, setEquationDialogOpen] = useState(false)
  const [recommendedCharts, setRecommendedCharts] = useState<ChartRecommendations | null>(null)
  /// The focused floating visual (chart/shape/image); charts surface a
  /// contextual Chart Design ribbon tab while selected.
  const [selectedVisual, setSelectedVisual] = useState<WorkbookVisualObject | null>(null)
  /// Chart panels (Select Data dialog / format task pane), opened from the
  /// ribbon or the chart context menu, keyed like chart edits.
  const [chartDialog, setChartDialog] = useState<{ kind: ChartDialogKind; editKey: string } | null>(
    null,
  )
  const chartElement = useSyncExternalStore(
    subscribeChartElementSelection,
    getChartElementSelection,
  )
  /// Bumped on every visual/chart edit: journal edits merge in place, so the
  /// journal size (pendingEdits) alone misses same-target re-edits and the
  /// ribbon echo would go stale.
  const [, setVisualEditTick] = useState(0)
  /// In-session slicers (OOXML slicer part persistence: see the TODO in
  /// SlicerPanel).
  const [slicers, setSlicers] = useState<readonly SlicerUiState[]>([])
  /// Non-null while the "Insert Slicer" field picker is open.
  const [slicerPicker, setSlicerPicker] = useState<SlicerPickerState | null>(null)
  /// In-session timelines (same session-only model as slicers).
  const [timelines, setTimelines] = useState<readonly TimelineUiState[]>([])
  /// Non-null while the "Insert Timeline" field picker is open.
  const [timelinePicker, setTimelinePicker] = useState<TimelinePickerState | null>(null)
  const menuActionRef = useRef<(action: MenuAction) => void>(() => {})
  /// Fresh handleSave for the AutoSave tick (assigned each render, like
  /// menuActionRef, so the interval closure never goes stale).
  const handleSaveRef = useRef<
    (
      mode: 'save' | 'save-as' | 'recovery',
      quiet?: boolean,
    ) => Promise<SaveActionResult | undefined>
  >(() => Promise.resolve(undefined))
  const closeSaveRef = useRef<() => Promise<void>>(() => Promise.resolve())
  const refreshSelectionFormatRef = useRef<() => void>(() => {})
  const chartEditRef = useRef<(chartPath: string, edit: ChartEditData) => void>(() => {})
  const chartVectorRef = useRef<(chartPath: string, range: string) => Promise<ChartVectorRead>>(
    () => Promise.reject(new Error('Workbook not ready.')),
  )
  const shapeEditRef = useRef<(visualId: string, changes: ShapeEditChanges) => void>(() => {})
  /// A3 editing of an existing pivot: context locked when the dialog opens, used
  /// on Apply.
  const pivotEditContextRef = useRef<PivotEditContext | null>(null)
  /** App-scope refs/state bundle for the extracted pivot actions (pivot-actions.ts). */
  function pivotContext(): PivotActionContext {
    return {
      univerRef,
      lazyWorkbookRef,
      pivotEditContextRef,
      slicers,
      slicerPicker,
      setSlicers,
      setSlicerPicker,
      timelines,
      timelinePicker,
      setTimelines,
      setTimelinePicker,
      setMessage,
      setPendingEdits,
    }
  }

  /** App-scope refs/state bundle for the extracted visual-insert actions (visual-actions.ts). */
  function visualContext(): VisualActionContext {
    return {
      adapterRef,
      univerRef,
      lazyWorkbookRef,
      visualDisposablesRef,
      visualInstallTimerRef,
      chartEditRef,
      chartVectorRef,
      shapeEditRef,
      setMessage,
      setRevision,
      setPendingEdits,
      pivotContext,
      queueDemoVisualInstall,
      refreshLazyVisuals,
    }
  }

  /** App-scope refs/state bundle for the extracted data-tool actions (data-tools-actions.ts). */
  function dataToolsContext(): DataToolsContext {
    return { univerRef, lazyWorkbookRef, setMessage, setPendingEdits, setAdvancedFilterColumns }
  }

  function pageLayoutContext(): PageLayoutContext {
    return { univerRef, lazyWorkbookRef, setMessage, setPendingEdits }
  }

  function saveContext(): SaveContext {
    return { univerRef, lazyWorkbookRef, setMessage, openLazyWorkbook }
  }

  function visualSyncContext(): VisualSyncContext {
    return {
      adapterRef,
      univerRef,
      lazyWorkbookRef,
      chartSyncRef,
      setMessage,
      setPendingEdits,
      refreshLazyVisuals,
      refreshDemoVisuals,
    }
  }

  /** Guards repeated picker/open requests while the browser host resolves a file. */
  const workbookOpeningRef = useRef(false)
  useEffect(() => {
    setVisualSelectionListener({
      select: (visual) =>
        setSelectedVisual((current) => (current?.id === visual.id ? current : visual)),
      deselect: () => setSelectedVisual(null),
    })
    setChartDialogListener((editKey, dialog) => setChartDialog({ kind: dialog, editKey }))
    return () => {
      setVisualSelectionListener(null)
      setChartDialogListener(null)
    }
  }, [])

  // The format pane follows the chart selection; deselecting closes it.
  useEffect(() => {
    if (!selectedVisual) setChartDialog(null)
  }, [selectedVisual])

  useEffect(() => {
    // Univer paints the grid on canvas, so it can't follow the CSS tokens —
    // mirror the <html data-theme> state into its official darkMode flag
    const prefersDark = window.matchMedia('(prefers-color-scheme: dark)')
    const isDarkTheme = () =>
      document.documentElement.getAttribute('data-theme') === 'dark' ||
      (!document.documentElement.hasAttribute('data-theme') && prefersDark.matches)
    const univerCreateStartedAt = performance.now()
    const bootstrapMs = univerCreateStartedAt
    const entryModuleReadyAt = Math.min(
      univerCreateStartedAt,
      Math.max(0, window.__genofficeXlsxEntryModuleReadyAt ?? 0),
    )
    const navigation = performance.getEntriesByType('navigation')[0] as
      PerformanceNavigationTiming | undefined
    const resourceReceivedAt = Math.min(
      entryModuleReadyAt,
      Math.max(0, navigation?.responseEnd ?? 0),
    )
    const bootstrapPhases = {
      resourceReceiveMs: resourceReceivedAt,
      moduleGraphReadyMs: entryModuleReadyAt - resourceReceivedAt,
      reactMountMs: univerCreateStartedAt - entryModuleReadyAt,
    }
    const runtime = createUniver({
      // green selection/highlight instead of Univer's default blue
      theme: greenTheme,
      darkMode: isDarkTheme(),
      locale: LocaleType.EN_US,
      locales: {
        [LocaleType.EN_US]: mergeLocales(
          UniverPresetSheetsCoreEnUS,
          UniverPresetSheetsConditionalFormattingEnUS,
          UniverPresetSheetsFilterEnUS,
          UniverPresetSheetsDataValidationEnUS,
          UniverPresetSheetsNoteEnUS,
          UniverPresetSheetsFindReplaceEnUS,
          UniverPresetSheetsSortEnUS,
          UniverPresetSheetsTableEnUS,
          // sheets-ui code in 0.25.1 references these two keys, but the language
          // pack shipped without the entries — unless patched, the raw
          // "sheets-ui.info.forceStringInfo" pops up for users.
          // mergeLocales shallow-merges namespaces, so the existing entries must
          // be spread; otherwise the whole sheets-ui namespace gets overwritten
          // (the sheet-tab context menu turns into bare keys).
          {
            'sheets-ui': {
              ...(UniverPresetSheetsCoreEnUS as Record<string, Record<string, unknown>>)[
                'sheets-ui'
              ],
              info: {
                ...(
                  UniverPresetSheetsCoreEnUS as Record<
                    string,
                    Record<string, Record<string, string>>
                  >
                )['sheets-ui']?.info,
                error: 'Number stored as text',
                forceStringInfo:
                  'The value in this cell is stored as text — it will not be treated as a ' +
                  'number in formulas.',
              },
            },
          },
        ),
      },
      presets: [
        UniverSheetsCorePreset({
          container: 'univer-container',
          // header: true + toolbar: false renders only the name box + formula
          // bar (the Univer ribbon needs both flags).
          header: true,
          toolbar: false,
          contextMenu: true,
          formulaBar: true,
          footer: {
            sheetBar: true,
            statisticBar: true,
            menus: true,
            // zoom lives in the custom full-width status bar (unified with docs/slides)
            zoomSlider: false,
          },
          statusBarStatistic: true,
        }),
        UniverSheetsDrawingPreset(),
        UniverSheetsConditionalFormattingPreset(),
        UniverSheetsFilterPreset(),
        UniverSheetsDataValidationPreset(),
        UniverSheetsNotePreset(),
        UniverSheetsFindReplacePreset(),
        UniverSheetsSortPreset(),
        UniverSheetsTablePreset(),
      ],
    })
    const univerCreateMs = performance.now() - univerCreateStartedAt
    const worksheetInstallStartedAt = performance.now()
    loadSnapshotIntoUniver(runtime, initialSnapshot, 'new-workbook', 'Untitled')
    const worksheetInstallMs = performance.now() - worksheetInstallStartedAt
    const firstCommitStartedAt = performance.now()
    univerRef.current = runtime
    // live theme switching: main.tsx updates data-theme first (its listener
    // registered at bootstrap), so reading the attribute here is safe; the
    // matchMedia listener covers OS appearance flips while in system mode
    const themeService = runtime.univer.__getInjector().get(ThemeService)
    const applyUniverDark = () => themeService.setDarkMode(isDarkTheme())
    const offThemeChanged = window.desktopApi?.onThemeChanged?.(applyUniverDark)
    prefersDark.addEventListener('change', applyUniverDark)
    // Undo/redo stack occupancy: the QAT buttons grey out when there is nothing to apply
    const undoRedoService = runtime.univer.__getInjector().get(IUndoRedoService)
    const undoRedoSub = undoRedoService.undoRedoStatus$.subscribe(
      ({ undos, redos }: { undos: number; redos: number }) =>
        setUniverHist({ canUndo: undos > 0, canRedo: redos > 0 }),
    )
    // Programmatic installs (viewport streaming, file loads, merges, row
    // heights, notes, CF/filter rules) run through the same undoable commands
    // as user edits; they all raise journalSuppression, so drop their undo
    // entries there too — a freshly opened workbook starts with an empty
    // stack and undo can never strip loaded file content or layout.
    const originalPushUndoRedo = undoRedoService.pushUndoRedo.bind(undoRedoService)
    undoRedoService.pushUndoRedo = (item) => {
      if (!journalSuppression.active) originalPushUndoRedo(item)
    }
    // IUndoRedoService is registered lazily, so the injector hands out a redi
    // proxy that caches a bound copy of each method on first read — and the
    // initial snapshot load above reads pushUndoRedo before this wrapper is
    // assigned. The assignment lands on the real instance, but every caller
    // resolves the service through the same proxy and keeps getting the stale
    // cached original. Deleting the key clears that per-proxy cache so the
    // next read re-binds to the wrapper.
    Reflect.deleteProperty(undoRedoService, 'pushUndoRedo')
    // The window always starts blank now; still consume the one-shot
    // new-blank flag so it doesn't leak into the next workbook open.
    void window.desktopApi?.consumeNewBlankWorkbook?.()
    // Pull any shell-queued workbook ourselves: the shell's 'open' nudge loop
    // gives up after 30s, and on slow dev cold starts Univer mounts later than
    // that — the tab would strand as a blank in-memory workbook (no save, no
    // shapes) with the queued file silently never opened.
    void window.desktopApi?.hasQueuedWorkbook?.().then((queued) => {
      if (queued) void handleInspectWorkbook()
    })
    // Univer 0.25.1 also badges text parseable as date/time, phone numbers, and
    // other long numeric identifiers with "Number stored as text". Those values
    // should remain text, so clear the view type before the built-in marker
    // interceptor (priority 10). Short numeric text ("007", "20%") keeps its
    // warning.
    const dateTextDisposable = runtime.univer
      .__getInjector()
      .get(SheetInterceptorService)
      .intercept(INTERCEPTOR_POINT.CELL_CONTENT, {
        priority: 11,
        effect: InterceptorEffectEnum.Style,
        handler: (cell, _position, next) => {
          if (cell?.t === CellValueType.STRING && typeof cell.v === 'string') {
            if (isNumericIdentifierText(cell.v)) return next({ ...cell, t: undefined })
            if (isRealNum(cell.v)) return next(cell)
            const parsed = getNumfmtParseValueFilter(cell.v)
            if (parsed?.z && /[ymdhs]/i.test(parsed.z)) return next({ ...cell, t: undefined })
          }
          return next(cell)
        },
      })
    // Copying in a filtered sheet must skip hidden rows;
    // see filtered-copy.ts for why the built-in hook is not enough.
    const filteredCopyDisposable = installFilteredCopyHook(runtime)
    // Excel-compatible TSV plain text (TRUE/FALSE, quoted newlines).
    const tsvClipboardDisposable = installTsvClipboardFix(runtime)
    // Formula view: swap formula cells to their formula text per sheet.
    const formulaViewDisposable = installFormulaViewInterceptor(runtime, lazyWorkbookRef)
    // Formula bar shows harvested formula text on streamed workbooks whose
    // closure gave up; display-only, the engine never sees it.
    const formulaTextDisposable = installFormulaTextInterceptor(runtime, lazyWorkbookRef)
    // Excel-parity number-format display: empty sections, text section,
    // _/* padding, General digit fitting.
    const numberFormatFixDisposable = installNumberFormatFix(runtime)
    // CELL("filename") resolves the session's on-disk path; converted
    // imports (needsSaveAs) count as never-saved, like Excel.
    const cellFilenameDisposable = installCellFilenameFunction(runtime, () => {
      const file = lazyWorkbookRef.current?.file
      return file && !file.needsSaveAs ? (file.path ?? null) : null
    })
    // RATE converges near -100% via bisection instead of erroring.
    const rateFallbackDisposable = installRateFallback(runtime)
    // Escaped quotes ("") no longer shift lexer indices and silently
    // rewrite committed formulas.
    const formulaLexerFixDisposable = installFormulaLexerFix(runtime)
    // Renaming a sheet to a case variant of itself is not a duplicate.
    const sheetRenameFixDisposable = installSheetRenameFix()
    // Arrow keys stop at the sheet edge instead of wrapping to the far side.
    const selectionWrapGuardDisposable = installSelectionWrapGuard(runtime)
    // Row-header double-click autofits every selected row, like Excel.
    const multiRowAutofitDisposable = installMultiRowAutofit(runtime)
    // Empty-value formula results (IFERROR/IF/CHOOSE over blank refs)
    // display as 0 like Excel.
    const nullResultDisposable = installFormulaNullResultFix(runtime)
    // Copy/cut load their selection into the lazy window first so streamed
    // workbooks don't serialize blanks for never-viewed rows.
    const copyMaterializeDisposable = installCopyMaterialize(runtime, lazyWorkbookRef, setMessage)
    // List-validation arrows stay discoverable on values without cluttering empty template rows.
    const dataValidationArrowDisposable = installPopulatedDataValidationArrow(runtime)
    // Univer's own UI (rule-management panels, dialogs) follows the app
    // language instead of hard-coded English.
    void applyUniverLocale(runtime, getLang())
    // Rule-management panels show what each rule actually does: list options /
    // source range, CF formula text, ⚠ on #REF! dead rules.
    const ruleDetailDisposable = installRuleDetail(runtime)
    const scrollDisposable = runtime.univerAPI.addEvent(
      runtime.univerAPI.Event.Scroll,
      (params) => {
        const { worksheet } = params
        // The event carries the true post-scroll position; getVisibleRange
        // inside loadVisibleRange lags a frame.
        const eventStart = params as { sheetViewStartRow?: number; sheetViewStartColumn?: number }
        void loadVisibleRange(
          runtime,
          lazyWorkbookRef,
          worksheet,
          setMessage,
          typeof eventStart.sheetViewStartRow === 'number' &&
            typeof eventStart.sheetViewStartColumn === 'number'
            ? { row: eventStart.sheetViewStartRow, column: eventStart.sheetViewStartColumn }
            : undefined,
        )
        let visible: ReturnType<typeof worksheet.getVisibleRange>
        try {
          visible = worksheet.getVisibleRange()
        } catch {
          // The lazy loader falls back to the top-left viewport while Univer
          // replaces its scroll controller; visual installation can wait.
          return
        }
        const viewportKey = visible
          ? `${worksheet.getSheetId()}:${visible.startRow}:${visible.endRow}:${visible.startColumn}:${visible.endColumn}`
          : worksheet.getSheetId()
        if (visualViewportKeyRef.current === viewportKey) return
        visualViewportKeyRef.current = viewportKey
        queueVisualInstall(
          runtime,
          lazyWorkbookRef,
          visualDisposablesRef,
          visualInstallTimerRef,
          worksheet.getSheetId(),
          chartEditRef,
          chartVectorRef,
          shapeEditRef,
        )
        queueSparklineInstall(
          runtime,
          lazyWorkbookRef,
          sparklineDisposablesRef,
          sparklineTimerRef,
          worksheet.getSheetId(),
        )
      },
    )
    const zoomDisposable = runtime.univerAPI.addEvent(
      runtime.univerAPI.Event.SheetZoomChanged,
      ({ worksheet }) => {
        setZoomPercent(Math.round(worksheet.getZoom() * 100))
      },
    )
    // In-cell editor open/closed, read by the AutoSave tick: saving reloads
    // the workbook and would wipe an in-progress edit.
    const editStartDisposable = runtime.univerAPI.addEvent(
      runtime.univerAPI.Event.SheetEditStarted,
      () => {
        editingCellRef.current = true
      },
    )
    const editEndDisposable = runtime.univerAPI.addEvent(
      runtime.univerAPI.Event.SheetEditEnded,
      () => {
        editingCellRef.current = false
      },
    )
    const sheetDisposable = runtime.univerAPI.addEvent(
      runtime.univerAPI.Event.ActiveSheetChanged,
      ({ activeSheet }) => {
        void loadVisibleRange(runtime, lazyWorkbookRef, activeSheet, setMessage)
        // formula view is per-sheet (sheetView/@showFormulas)
        applyShowFormulasView(runtime, lazyWorkbookRef.current, activeSheet.getSheetId())
        // zoom is per-sheet state; echo the new sheet's level
        setZoomPercent(Math.round(activeSheet.getZoom() * 100))
        refreshSelectionFormatRef.current()
        visualViewportKeyRef.current = ''
        if (!lazyWorkbookRef.current) {
          queueDemoVisualInstall(runtime, activeSheet.getSheetId())
        }
        queueVisualInstall(
          runtime,
          lazyWorkbookRef,
          visualDisposablesRef,
          visualInstallTimerRef,
          activeSheet.getSheetId(),
          chartEditRef,
          chartVectorRef,
          shapeEditRef,
        )
        queueSparklineInstall(
          runtime,
          lazyWorkbookRef,
          sparklineDisposablesRef,
          sparklineTimerRef,
          activeSheet.getSheetId(),
        )
      },
    )
    const editDisposable = runtime.univerAPI.addEvent(
      runtime.univerAPI.Event.BeforeSheetEditStart,
      (event) => {
        const state = lazyWorkbookRef.current
        if (!state) return
        const sheetId = event.worksheet.getSheetId()
        const sheet = state.file.sheets.find((candidate) => candidate.id === sheetId)
        if (!sheet) return
        // Pivot output is baked into the worksheet; editing it would corrupt
        // the file's pivot semantics. Protected in every load mode.
        if (
          sheet.pivotRanges.some(
            (range) =>
              event.row >= range.startRow &&
              event.row <= range.endRow &&
              event.column >= range.startColumn &&
              event.column <= range.endColumn,
          )
        ) {
          event.cancel = true
          setMessage(t('appPivotCellNoEdit'))
          return
        }
        // Fully-loaded workbooks have nothing left to stream in.
        if (state.flags.preloadComplete) return
        // Editing a cell whose original content hasn't streamed in yet would
        // silently overwrite data the user never saw. Beyond the file's used
        // range every cell is genuinely empty, so those edits are safe — and
        // so are rows/columns inserted this session (journal-owned, nothing
        // streams into them). Bounds are screen-space: structural ops shift
        // the data extent.
        const ops = state.editJournal.structuralOps.get(sheetId) ?? []
        const beyondData =
          event.row >= sheet.rowCount + netAxisDelta(ops, 'row') ||
          event.column >= sheet.columnCount + netAxisDelta(ops, 'column')
        const journalOwned =
          ops.length > 0 &&
          (screenToFile(ops, 'row', event.row) === null ||
            screenToFile(ops, 'column', event.column) === null)
        const loaded = state.loadedRanges.get(sheetId)
        const inLoaded =
          loaded !== undefined &&
          event.row >= loaded.startRow &&
          event.row <= loaded.endRow &&
          event.column >= loaded.startColumn &&
          event.column <= loaded.endColumn
        const inFrozen =
          loaded !== undefined &&
          (event.row < (sheet.freeze?.frozenRows ?? 0) ||
            event.column < (sheet.freeze?.frozenColumns ?? 0))
        if (!beyondData && !journalOwned && !inLoaded && !inFrozen) {
          event.cancel = true
          setMessage(t('appAreaStreaming'))
        }
      },
    )
    const journalDisposable = runtime.univerAPI.addEvent(
      runtime.univerAPI.Event.CommandExecuted,
      (event) => {
        if (journalSuppression.active) return
        // The formula engine re-applies cached results with these execution
        // options; they are derived state, never user edits.
        const options = event.options as { fromFormula?: boolean } | undefined
        if (options?.fromFormula) return
        // The copy finished (or failed); a stale source must not claim a
        // later, unrelated insert-sheet.
        if (event.id === COPY_SHEET_COMMAND) {
          pendingCopySource = undefined
          return
        }
        const rowColumn = ROW_COLUMN_MUTATIONS[event.id]
        const merge = MERGE_MUTATIONS[event.id]
        const axisAttr = AXIS_ATTR_MUTATIONS[event.id]
        if (
          event.id !== SET_RANGE_VALUES_MUTATION &&
          event.id !== SET_NUMFMT_MUTATION &&
          !rowColumn &&
          !merge &&
          !axisAttr &&
          !SHEET_LIFECYCLE_MUTATIONS.has(event.id) &&
          event.id !== REORDER_RANGE_MUTATION &&
          !FILTER_MUTATIONS.has(event.id) &&
          !CF_MUTATIONS.has(event.id) &&
          !DV_MUTATIONS.has(event.id) &&
          !DEFINED_NAME_MUTATIONS.has(event.id) &&
          !NOTE_MUTATIONS.has(event.id) &&
          event.id !== MOVE_RANGE_MUTATION &&
          event.id !== MOVE_ROWS_MUTATION &&
          event.id !== SET_FROZEN_MUTATION &&
          event.id !== SET_TAB_COLOR_MUTATION &&
          event.id !== TOGGLE_GRIDLINES_MUTATION
        ) {
          return
        }
        const state = lazyWorkbookRef.current
        if (!state) {
          // Demo mode journals nothing, but chart↔data sync still applies.
          if (event.id === SET_RANGE_VALUES_MUTATION) {
            const demoParams = event.params as
              { subUnitId?: string; cellValue?: unknown } | undefined
            const bounds = cellValueBounds(demoParams?.cellValue)
            if (demoParams?.subUnitId && bounds) queueChartDataSync(demoParams.subUnitId, bounds)
          }
          return
        }
        const params = event.params as
          | {
              unitId?: string
              subUnitId?: string
              sheetId?: string
              cellValue?: unknown
              range?: IRange
              ranges?: IRange[]
              name?: string
              sheet?: { id?: string; name?: string }
            }
          | undefined
        if (params?.unitId !== `file-${state.file.sha256}`) return
        if (SHEET_LIFECYCLE_MUTATIONS.has(event.id)) {
          if (event.id === 'sheet.mutation.insert-sheet') {
            const { id, name } = params.sheet ?? {}
            if (typeof id === 'string' && typeof name === 'string') {
              if (pendingCopySource !== undefined) {
                recordSheetDuplicate(state.editJournal, id, name, pendingCopySource)
                pendingCopySource = undefined
              } else {
                recordSheetInsert(state.editJournal, id, name)
              }
            }
          } else if (event.id === 'sheet.mutation.remove-sheet') {
            if (typeof params.subUnitId === 'string') {
              recordSheetRemove(state.editJournal, params.subUnitId)
            }
          } else if (event.id === 'sheet.mutation.set-worksheet-order') {
            recordSheetOrderChange(state.editJournal)
          } else if (event.id === 'sheet.mutation.set-worksheet-hidden') {
            const hidden = (params as { hidden?: number | boolean }).hidden
            if (typeof params.subUnitId === 'string' && hidden !== undefined) {
              const originallyHidden =
                state.file.sheets.find((sheet) => sheet.id === params.subUnitId)?.hidden ?? false
              recordSheetHidden(
                state.editJournal,
                params.subUnitId,
                hidden === true || hidden === 1,
                originallyHidden,
              )
            }
          } else if (typeof params.subUnitId === 'string' && typeof params.name === 'string') {
            const originalName = state.file.sheets.find(
              (sheet) => sheet.id === params.subUnitId,
            )?.name
            // The live sync matches series refs against live sheet names, so
            // in-memory refs must follow the rename (the file's own c:f refs
            // are rewritten independently at save time).
            const previousName =
              state.editJournal.sheets.renamed.get(params.subUnitId) ??
              state.editJournal.sheets.added.get(params.subUnitId)?.name ??
              originalName
            recordSheetRename(state.editJournal, params.subUnitId, params.name, originalName)
            if (previousName !== undefined && previousName !== params.name) {
              renameChartRefsForSheet(state, previousName, params.name)
            }
          }
          setPendingEdits(journalSize(state.editJournal))
          return
        }
        // Move-range carries its sheet ids inside from/to, not at top level.
        if (event.id === MOVE_RANGE_MUTATION) {
          const move = event.params as
            | {
                from?: { subUnitId?: string; value?: unknown }
                to?: { subUnitId?: string; value?: unknown }
                fromRange?: IRange
                toRange?: IRange
              }
            | undefined
          const fromSheet = move?.from?.subUnitId ?? params.subUnitId
          const toSheet = move?.to?.subUnitId ?? params.subUnitId
          const fromRange = move?.fromRange ?? matrixBounds(move?.from?.value)
          const toRange = move?.toRange ?? matrixBounds(move?.to?.value)
          if (fromSheet && fromRange) journalRangeSnapshot(runtime, state, fromSheet, fromRange)
          if (toSheet && toRange) journalRangeSnapshot(runtime, state, toSheet, toRange)
          // Moved cells feed charts too, same as value mutations.
          if (fromSheet && fromRange) queueChartDataSync(fromSheet, fromRange)
          if (toSheet && toRange) queueChartDataSync(toSheet, toRange)
          setPendingEdits(journalSize(state.editJournal))
          return
        }
        // Workbook-level: defined-name mutations carry no subUnitId.
        if (DEFINED_NAME_MUTATIONS.has(event.id)) {
          recordDefinedNamesChange(state.editJournal)
          setPendingEdits(journalSize(state.editJournal))
          return
        }
        if (!params.subUnitId) return
        if (axisAttr) {
          const attrParams = event.params as {
            ranges?: IRange[]
            rowHeight?: number | Record<number, number>
            colWidth?: number | Record<number, number>
            autoHeightInfo?: number | Record<number, number>
          }
          const uniform = axisAttr.axis === 'row' ? attrParams.rowHeight : attrParams.colWidth
          const toFileSize = (pixels: number): number =>
            axisAttr.axis === 'row'
              ? Math.round(pixels * 0.75 * 100) / 100
              : pixelsToCharacterWidth(pixels)
          for (const range of attrParams.ranges ?? []) {
            const start = axisAttr.axis === 'row' ? range.startRow : range.startColumn
            const end = axisAttr.axis === 'row' ? range.endRow : range.endColumn
            if (!Number.isInteger(start) || !Number.isInteger(end) || end < start) continue
            if (end - start >= 100_000) continue
            const sizeKind = axisAttr.axis === 'row' ? 'set-row-size' : 'set-col-size'
            if (axisAttr.kind === 'hidden') {
              recordStructuralOp(state.editJournal, params.subUnitId, {
                kind: axisAttr.axis === 'row' ? 'set-rows-hidden' : 'set-cols-hidden',
                start,
                end,
                hidden: axisAttr.hidden === true,
              })
            } else if (axisAttr.kind === 'auto-size') {
              // Setting an explicit height ALSO emits this mutation with
              // autoHeightInfo=0 (auto off) — only auto ON resets the height.
              const info = attrParams.autoHeightInfo
              if (typeof info === 'number') {
                if (info === 1) {
                  recordStructuralOp(state.editJournal, params.subUnitId, {
                    kind: sizeKind,
                    start,
                    end,
                    size: null,
                  })
                }
              } else if (info && typeof info === 'object') {
                for (let line = start; line <= end; line += 1) {
                  if (info[line] !== 1) continue
                  recordStructuralOp(state.editJournal, params.subUnitId, {
                    kind: sizeKind,
                    start: line,
                    end: line,
                    size: null,
                  })
                }
              }
            } else if (typeof uniform === 'number') {
              recordStructuralOp(state.editJournal, params.subUnitId, {
                kind: sizeKind,
                start,
                end,
                size: toFileSize(uniform),
              })
            } else if (uniform && typeof uniform === 'object') {
              // Per-line sizes (undo restores): one op per line in the range.
              for (let line = start; line <= end; line += 1) {
                const pixels = uniform[line]
                if (typeof pixels !== 'number') continue
                recordStructuralOp(state.editJournal, params.subUnitId, {
                  kind: sizeKind,
                  start: line,
                  end: line,
                  size: toFileSize(pixels),
                })
              }
            }
          }
          setPendingEdits(journalSize(state.editJournal))
          return
        }
        if (FILTER_MUTATIONS.has(event.id)) {
          recordFilterChange(state.editJournal, params.subUnitId)
          setPendingEdits(journalSize(state.editJournal))
          return
        }
        if (CF_MUTATIONS.has(event.id)) {
          recordCfChange(state.editJournal, params.subUnitId)
          setPendingEdits(journalSize(state.editJournal))
          return
        }
        if (DV_MUTATIONS.has(event.id)) {
          if (params.subUnitId) {
            recordDvChange(state.editJournal, params.subUnitId)
            setPendingEdits(journalSize(state.editJournal))
          }
          return
        }
        if (event.id === SET_FROZEN_MUTATION) {
          // Recording from the mutation (not the ribbon handler) keeps the
          // journal in step with Univer's undo/redo of the freeze.
          const freeze = event.params as
            { subUnitId?: string; ySplit?: number; xSplit?: number } | undefined
          if (freeze?.subUnitId && !isSheetRemoved(state.editJournal, freeze.subUnitId)) {
            recordPageSetup(state.editJournal, freeze.subUnitId, {
              frozenRows: Math.max(0, freeze.ySplit ?? 0),
              frozenColumns: Math.max(0, freeze.xSplit ?? 0),
            })
            setPendingEdits(journalSize(state.editJournal))
          }
          return
        }
        if (event.id === TOGGLE_GRIDLINES_MUTATION) {
          const gridlines = event.params as
            { subUnitId?: string; showGridlines?: number } | undefined
          if (
            gridlines?.subUnitId &&
            gridlines.showGridlines !== undefined &&
            !isSheetRemoved(state.editJournal, gridlines.subUnitId)
          ) {
            recordPageSetup(state.editJournal, gridlines.subUnitId, {
              showGridlines: gridlines.showGridlines === 1,
            })
            setPendingEdits(journalSize(state.editJournal))
          }
          return
        }
        if (event.id === SET_TAB_COLOR_MUTATION) {
          const tabColor = event.params as { subUnitId?: string; color?: string | null } | undefined
          if (tabColor?.subUnitId && !isSheetRemoved(state.editJournal, tabColor.subUnitId)) {
            const color = tabColor.color
            recordPageSetup(state.editJournal, tabColor.subUnitId, {
              tabColor:
                typeof color === 'string' && /^#[0-9A-Fa-f]{6}$/.test(color)
                  ? color.toUpperCase()
                  : null,
            })
            setPendingEdits(journalSize(state.editJournal))
          }
          return
        }
        if (NOTE_MUTATIONS.has(event.id)) {
          const noteSheetId = params.sheetId ?? params.subUnitId
          if (noteSheetId) {
            recordNoteChange(state.editJournal, noteSheetId)
            setPendingEdits(journalSize(state.editJournal))
          }
          return
        }
        if (event.id === REORDER_RANGE_MUTATION) {
          if (params.range) {
            journalRangeSnapshot(runtime, state, params.subUnitId, params.range)
            // Sorted cells feed charts too, same as value mutations.
            queueChartDataSync(params.subUnitId, params.range)
            setPendingEdits(journalSize(state.editJournal))
          }
          return
        }
        if (rowColumn || event.id === MOVE_ROWS_MUTATION) {
          let structuralOp: StructuralJournalOp
          if (rowColumn) {
            const range = params.range
            if (!range) return
            const index = rowColumn.axis === 'row' ? range.startRow : range.startColumn
            const count =
              rowColumn.axis === 'row'
                ? range.endRow - range.startRow + 1
                : range.endColumn - range.startColumn + 1
            if (count <= 0) return
            structuralOp = { kind: rowColumn.kind, index, count }
          } else {
            const move = event.params as { sourceRange?: IRange; targetRange?: IRange } | undefined
            if (!move?.sourceRange || !move.targetRange) return
            const index = move.sourceRange.startRow
            const count = move.sourceRange.endRow - move.sourceRange.startRow + 1
            const before = move.targetRange.startRow
            if (count <= 0 || (before >= index && before <= index + count)) return
            structuralOp = { kind: 'move-rows', index, count, before }
          }
          const structuralSheetId = params.subUnitId
          // Refs are matched by live sheet name (they follow renames).
          const structuralSheetName =
            runtime.univerAPI
              .getActiveWorkbook()
              ?.getSheetBySheetId(structuralSheetId)
              ?.getSheetName() ??
            state.file.sheets.find((sheet) => sheet.id === structuralSheetId)?.name
          recordStructuralOp(
            state.editJournal,
            structuralSheetId,
            structuralOp,
            structuralSheetName,
          )
          // File visuals shift on-screen too (the save shifts the file's own
          // anchors and c:f refs independently); keeping the in-memory copy in
          // the new space keeps the preview and the live data sync honest.
          state.file.visuals.forEach((visual, at) => {
            state.file.visuals[at] = shiftVisualForStructuralOp(
              visual,
              structuralSheetId,
              structuralSheetName,
              structuralOp,
            )
          })
          refreshLazyVisuals(state)
          // Univer shifted its installed cells itself, but the loaded-range
          // bookkeeping and frozen strip are now stale — refetch the viewport
          // through the updated coordinate mapping. Moves are exempt: they
          // are gated to fully loaded sheets, and the refetch would re-install
          // cells through undoable commands, burying the move's undo entry.
          if (structuralOp.kind !== 'move-rows') {
            state.loadedRanges.delete(params.subUnitId)
            state.frozenStripKeys.delete(params.subUnitId)
          }
          // Pinned closure values shift with the model; pinned formulas are
          // dropped — Univer rewrote their references in the model, so a
          // stale snapshot must not be re-applied after eviction.
          const pinnedClosure = state.closure.pinned.get(params.subUnitId)
          if (pinnedClosure && 'index' in structuralOp) {
            const shifted = shiftPinnedCells(pinnedClosure, structuralOp)
            for (const [key, cell] of [...shifted]) {
              if (cell.f !== undefined) shifted.delete(key)
            }
            state.closure.pinned.set(params.subUnitId, shifted)
          }
          // The recalc fallback reads the on-disk file; structural edits
          // desync every coordinate, so its overlays must not re-apply.
          state.recalc.overlay.clear()
          state.recalc.formulaCells.clear()
          const activeSheet = runtime.univerAPI.getActiveWorkbook()?.getActiveSheet()
          if (activeSheet?.getSheetId() === params.subUnitId) {
            void loadVisibleRange(runtime, lazyWorkbookRef, activeSheet, setMessage)
          }
          setPendingEdits(journalSize(state.editJournal))
          return
        }
        if (merge) {
          for (const range of params.ranges ?? []) {
            if (range.endRow < range.startRow || range.endColumn < range.startColumn) continue
            recordStructuralOp(state.editJournal, params.subUnitId, {
              kind: merge,
              range: {
                startRow: range.startRow,
                endRow: range.endRow,
                startColumn: range.startColumn,
                endColumn: range.endColumn,
              },
            })
          }
          setPendingEdits(journalSize(state.editJournal))
          return
        }
        // Copy-sheet batches a large source's cellData into follow-up chunk
        // mutations; the save clones the worksheet part, so journaling them
        // as edits would duplicate (and re-encode) content the clone covers.
        if ((event.params as { __splitChunk__?: boolean } | undefined)?.__splitChunk__) return
        const recorded =
          event.id === SET_NUMFMT_MUTATION
            ? recordSetNumfmt(state.editJournal, params.subUnitId, params)
            : recordSetRangeValues(state.editJournal, params.subUnitId, params.cellValue)
        if (recorded.length === 0) return
        setPendingEdits(journalSize(state.editJournal))
        const contentEdited = recorded.some(
          (entry) => entry.hasValue || entry.formula !== undefined,
        )
        if (contentEdited) {
          const valueEntries = recorded.filter(
            (entry) => entry.hasValue || entry.formula !== undefined,
          )
          queueChartDataSync(params.subUnitId, {
            startRow: Math.min(...valueEntries.map((entry) => entry.row)),
            endRow: Math.max(...valueEntries.map((entry) => entry.row)),
            startColumn: Math.min(...valueEntries.map((entry) => entry.column)),
            endColumn: Math.max(...valueEntries.map((entry) => entry.column)),
          })
          // Sparkline values read live from the grid — re-render them too.
          if (
            state.editJournal.sparklineAdds.length > 0 ||
            state.file.sheets.some((sheet) => sheet.sparklines.length > 0)
          ) {
            queueSparklineInstall(
              runtime,
              lazyWorkbookRef,
              sparklineDisposablesRef,
              sparklineTimerRef,
              params.subUnitId,
            )
          }
        }
        if (
          !state.formulaMode &&
          contentEdited &&
          state.closure.status === 'unavailable' &&
          state.recalc.failures < RECALC_MAX_FAILURES
        ) {
          queueFormulaRecalc(runtime, lazyWorkbookRef, setMessage)
        } else if (
          !state.formulaMode &&
          state.closure.status !== 'active' &&
          recorded.some((entry) => entry.formula)
        ) {
          setMessage(t('appFormulaRecordedPartial'))
        }
      },
    )
    const structuralDisposable = runtime.univerAPI.addEvent(
      runtime.univerAPI.Event.BeforeCommandExecute,
      (event) => {
        const state = lazyWorkbookRef.current
        if (journalSuppression.active || !state) return
        if (CF_RULE_COMMAND_PATTERN.test(event.id)) {
          // The Univer panel offers rules base OOXML cannot hold (x14-only
          // icon sets, date-occurring, equal/notEqual average, …); block them
          // here instead of failing the whole save later.
          const rule = (
            event.params as
              | {
                  rule?: { rule?: Record<string, unknown> & { type?: string; config?: unknown } }
                }
              | undefined
          )?.rule?.rule
          if (rule?.type === 'iconSet' && !iconSetSaveable(rule.config)) {
            event.cancel = true
            setMessage(t('appIconSetUnsupported'))
            return
          }
          if (rule && cfRuleUnsaveableReason(rule) !== null) {
            event.cancel = true
            setMessage(t('appCfRuleUnsaveable'))
          }
          return
        }
        if (STRUCTURAL_EDIT_COMMAND_PATTERN.test(event.id)) {
          // Row/column inserts/removals and merges are allowed in every load
          // mode: viewport reads translate screen ↔ file coordinates through
          // the journaled operation stream (view-transform.ts), and the save
          // replays the same stream against the file. Sheets carrying pivot
          // tables are the exception — a shift would desync the baked pivot
          // output from its definition.
          const subUnitId =
            (event.params as { subUnitId?: string } | undefined)?.subUnitId ??
            runtime.univerAPI.getActiveWorkbook()?.getActiveSheet()?.getSheetId()
          const sheet = state.file.sheets.find((candidate) => candidate.id === subUnitId)
          if (sheet && sheet.pivotRanges.length > 0) {
            event.cancel = true
            setMessage(t('appPivotSheetNoStructural'))
          }
          return
        }
        if (
          SORT_COMMAND_PATTERN.test(event.id) ||
          FILTER_COMMAND_PATTERN.test(event.id) ||
          event.id === MOVE_RANGE_COMMAND ||
          event.id === MOVE_ROWS_COMMAND
        ) {
          const subUnitId =
            (event.params as { subUnitId?: string } | undefined)?.subUnitId ??
            runtime.univerAPI.getActiveWorkbook()?.getActiveSheet()?.getSheetId()
          const isAddedSheet =
            subUnitId !== undefined && state.editJournal.sheets.added.has(subUnitId)
          // Sorting, filtering, and range moves read/rewrite model content,
          // so partially streamed data would silently produce wrong results.
          if (!isAddedSheet && (!state.formulaMode || !state.flags.preloadComplete)) {
            event.cancel = true
            setMessage(t('appNeedFullLoadSort'))
            return
          }
          if (
            (event.id === MOVE_RANGE_COMMAND || event.id === MOVE_ROWS_COMMAND) &&
            state.file.sheets.find((candidate) => candidate.id === subUnitId)?.pivotRanges.length
          ) {
            event.cancel = true
            setMessage(t('appPivotSheetNoMove'))
            return
          }
          if (
            FILTER_COMMAND_PATTERN.test(event.id) &&
            subUnitId !== undefined &&
            state.filterOrigins.get(subUnitId)?.origin === 'table'
          ) {
            event.cancel = true
            setMessage(t('appTableFilterNoEdit'))
          }
          return
        }
        if (event.id === AUTO_FILL_COMMAND && !state.flags.preloadComplete) {
          const target = (event.params as { targetRange?: IRange } | undefined)?.targetRange
          const subUnitId =
            (event.params as { subUnitId?: string } | undefined)?.subUnitId ??
            runtime.univerAPI.getActiveWorkbook()?.getActiveSheet()?.getSheetId()
          const sheet = state.file.sheets.find((candidate) => candidate.id === subUnitId)
          const loaded = subUnitId === undefined ? undefined : state.loadedRanges.get(subUnitId)
          const ops =
            subUnitId === undefined ? [] : (state.editJournal.structuralOps.get(subUnitId) ?? [])
          const beyondRow = sheet === undefined ? 0 : sheet.rowCount + netAxisDelta(ops, 'row')
          const beyondColumn =
            sheet === undefined ? 0 : sheet.columnCount + netAxisDelta(ops, 'column')
          const covered =
            target !== undefined &&
            (target.startRow >= beyondRow ||
              target.startColumn >= beyondColumn ||
              (loaded !== undefined &&
                target.startRow >= loaded.startRow &&
                target.endRow <= loaded.endRow &&
                target.startColumn >= loaded.startColumn &&
                target.endColumn <= loaded.endColumn))
          if (!covered) {
            event.cancel = true
            setMessage(t('appAutofillStreaming'))
          }
          return
        }
        if (DV_EDIT_COMMAND_PATTERN.test(event.id)) {
          const subUnitId =
            (event.params as { subUnitId?: string } | undefined)?.subUnitId ??
            runtime.univerAPI.getActiveWorkbook()?.getActiveSheet()?.getSheetId()
          const isAddedSheet =
            subUnitId !== undefined && state.editJournal.sheets.added.has(subUnitId)
          // The save rewrites the whole section from Univer's model, so the
          // file's own rules must be in the model before any edit.
          if (!isAddedSheet && (subUnitId === undefined || !state.appliedDvSheets.has(subUnitId))) {
            event.cancel = true
            setMessage(t('appDvNeedsIndexed'))
          }
          return
        }
        if (event.id === COPY_SHEET_COMMAND) {
          const subUnitId =
            (event.params as { subUnitId?: string } | undefined)?.subUnitId ??
            runtime.univerAPI.getActiveWorkbook()?.getActiveSheet()?.getSheetId()
          const isAddedSheet =
            subUnitId !== undefined && state.editJournal.sheets.added.has(subUnitId)
          // The Univer-side copy clones the model, so a partially streamed
          // source would produce a copy with silently missing data.
          if (!isAddedSheet && (!state.formulaMode || !state.flags.preloadComplete)) {
            event.cancel = true
            setMessage(t('appDuplicateNeedsFullLoad'))
            return
          }
          const sheet = state.file.sheets.find((candidate) => candidate.id === subUnitId)
          if (sheet && sheet.pivotRanges.length > 0) {
            event.cancel = true
            setMessage(t('appPivotSheetNoDuplicate'))
            return
          }
          if (subUnitId !== undefined) pendingCopySource = subUnitId
          return
        }
        if (BLOCKED_COMMAND_PATTERN.test(event.id)) {
          event.cancel = true
          setMessage(t('appMoveRowsColsUnsaved'))
        }
      },
    )
    // File-menu accelerators (⌘O/⌘S/⇧⌘S) arrive from the main process.
    const unsubscribeMenu =
      window.desktopApi?.onMenuAction((action) => menuActionRef.current(action)) ??
      (() => undefined)
    // Close guard chose Save: run the journal save and report the outcome.
    const unsubscribeCloseSave =
      window.desktopApi?.onCloseSaveRequest?.(() => void closeSaveRef.current()) ??
      (() => undefined)
    const selectionDisposable = runtime.univerAPI.addEvent(
      runtime.univerAPI.Event.SelectionChanged,
      () => {
        refreshSelectionFormatRef.current()
        // A grid click ends any floating-visual selection.
        clearVisualSelection()
      },
    )
    // Style edits from the ribbon, dialogs, and undo/redo all land as these
    // mutations; re-reading the selection keeps the ribbon echo current.
    const formatEchoDisposable = runtime.univerAPI.addEvent(
      runtime.univerAPI.Event.CommandExecuted,
      ({ id }) => {
        if (
          id === SET_RANGE_VALUES_MUTATION ||
          id === SET_NUMFMT_MUTATION ||
          id === REMOVE_NUMFMT_MUTATION
        ) {
          refreshSelectionFormatRef.current()
        }
      },
    )
    const clickDisposable = runtime.univerAPI.addEvent(
      runtime.univerAPI.Event.CellClicked,
      ({ worksheet, row, column }) => {
        const state = lazyWorkbookRef.current
        if (!state) return
        // A journaled link edit (set, changed, or removed) wins over the
        // file's streamed target.
        const journaled = hyperlinkEditAt(state.editJournal, worksheet.getSheetId(), row, column)
        const target =
          journaled !== undefined
            ? journaled
            : state.hyperlinkTargets.get(worksheet.getSheetId())?.get(`${row}:${column}`)
        if (target?.startsWith('http')) {
          void window.desktopApi.openExternal(target)
        } else if (target?.startsWith('#')) {
          navigateToAnchor(runtime, target.slice(1), setMessage)
        }
      },
    )
    let firstCommitFrame = 0
    const waitForFirstCommit = () => {
      firstCommitFrame = window.requestAnimationFrame(() => {
        const workbook = runtime.univerAPI.getActiveWorkbook()
        const worksheet = workbook?.getActiveSheet()
        const canvas = document.querySelector('#univer-container canvas')
        if (!workbook || !worksheet || !canvas) {
          waitForFirstCommit()
          return
        }
        startupTraceRef.current!.resolve({
          operation: 'xlsx.editor.cold_start',
          phases: {
            bootstrapMs,
            univerCreateMs,
            worksheetInstallMs,
            firstCommitMs: performance.now() - firstCommitStartedAt,
          },
          bootstrapPhases,
        })
      })
    }
    waitForFirstCommit()
    return () => {
      window.cancelAnimationFrame(firstCommitFrame)
      unsubscribeMenu()
      unsubscribeCloseSave()
      offThemeChanged?.()
      undoRedoSub.unsubscribe()
      undoRedoService.pushUndoRedo = originalPushUndoRedo
      // clear the proxy's cached bound wrapper (see the install site)
      Reflect.deleteProperty(undoRedoService, 'pushUndoRedo')
      prefersDark.removeEventListener('change', applyUniverDark)
      dateTextDisposable.dispose()
      filteredCopyDisposable.dispose()
      tsvClipboardDisposable.dispose()
      formulaViewDisposable.dispose()
      formulaTextDisposable.dispose()
      numberFormatFixDisposable.dispose()
      cellFilenameDisposable.dispose()
      rateFallbackDisposable.dispose()
      formulaLexerFixDisposable.dispose()
      sheetRenameFixDisposable.dispose()
      selectionWrapGuardDisposable.dispose()
      multiRowAutofitDisposable.dispose()
      nullResultDisposable.dispose()
      copyMaterializeDisposable.dispose()
      dataValidationArrowDisposable.dispose()
      ruleDetailDisposable()
      scrollDisposable.dispose()
      zoomDisposable.dispose()
      editStartDisposable.dispose()
      editEndDisposable.dispose()
      sheetDisposable.dispose()
      editDisposable.dispose()
      journalDisposable.dispose()
      structuralDisposable.dispose()
      selectionDisposable.dispose()
      formatEchoDisposable.dispose()
      clickDisposable.dispose()
      if (visualInstallTimerRef.current) clearTimeout(visualInstallTimerRef.current)
      disposeVisuals(visualDisposablesRef.current)
      visualViewportKeyRef.current = ''
      const lazyState = lazyWorkbookRef.current
      lazyWorkbookRef.current = null
      clearLazyState(lazyState)
      if (lazyState) {
        void window.desktopApi.closeWorkbook(lazyState.file.sessionId)
      }
      runtime.univer.dispose()
      univerRef.current = null
    }
  }, [])

  /// Demo-mode counterpart of queueVisualInstall: charts live in the adapter
  /// snapshot, so every grid rebuild (Apply/undo) and sheet switch re-installs
  /// them from there.
  function queueDemoVisualInstall(runtime: UniverRuntime, sheetId: string): void {
    if (demoVisualInstallTimerRef.current) clearTimeout(demoVisualInstallTimerRef.current)
    demoVisualInstallTimerRef.current = setTimeout(function install() {
      demoVisualInstallTimerRef.current = null
      if (lazyWorkbookRef.current) return
      if (runtime.univerAPI.getActiveWorkbook()?.getActiveSheet()?.getSheetId() !== sheetId) return
      if (isVisualDragActive()) {
        demoVisualInstallTimerRef.current = setTimeout(install, 100)
        return
      }
      disposeVisuals(demoVisualDisposablesRef.current)
      const visuals = adapterRef.current
        .getSnapshot()
        .sheets.flatMap((sheet) => sheet.visuals ?? [])
      demoVisualDisposablesRef.current =
        visuals.length === 0
          ? []
          : installWorkbookVisuals(
              runtime,
              { sessionId: 'demo-workbook', visuals },
              sheetId,
              {
                edits: EMPTY_CHART_EDITS,
                onEdit: (editKey, edit) => chartEditRef.current(editKey, edit),
                readVector: (editKey, range) => chartVectorRef.current(editKey, range),
              },
              { onEdit: (visualId, changes) => shapeEditRef.current(visualId, changes) },
            )
    }, 100)
  }

  function queueDemoVisualInstallForActiveSheet(): void {
    const runtime = univerRef.current
    const sheetId = runtime?.univerAPI.getActiveWorkbook()?.getActiveSheet()?.getSheetId()
    if (runtime && sheetId) queueDemoVisualInstall(runtime, sheetId)
  }

  function handleUndo(): void {
    // Ribbon/MCP commands always mutate Univer. In the blank community
    // workbook they therefore use Univer history too; the legacy in-memory
    // adapter remains only for its own plan-operation demo transactions.
    if (lazyWorkbookRef.current || univerHist.canUndo) {
      void univerRef.current?.univerAPI.undo()
      return
    }
    try {
      const receipt = adapterRef.current.undo()
      // Rebuild instead of patching: undo can remove cells and reverse
      // structural changes, neither of which syncUniver can express.
      loadSnapshotIntoUniver(
        univerRef.current,
        adapterRef.current.getSnapshot(),
        'new-workbook',
        'Untitled',
      )
      queueDemoVisualInstallForActiveSheet()
      setRevision(receipt.revision)
      setMessage(t('appUndoCommitted', { revision: receipt.revision }))
    } catch (error: unknown) {
      setMessage(error instanceof Error ? error.message : t('appUndoFailed'))
    }
  }

  /// QAT Redo: workbook history via Univer, same path as the app menu's ⇧⌘Z
  /// (the demo adapter has no redo, matching the menu's behavior).
  function handleRedo(): void {
    void univerRef.current?.univerAPI.redo()
  }

  /** App-scope refs/state bundle for the extracted ribbon dispatcher (ribbon-actions.ts). */
  function ribbonContext(): RibbonCommandContext {
    return {
      univerRef,
      lazyWorkbookRef,
      traceArrowsRef,
      sparklineDisposablesRef,
      sparklineTimerRef,
      chartEditRef,
      shapeEditRef,
      refreshSelectionFormatRef,
      selectedVisual,
      selectedChart,
      setMessage,
      setChartDialog,
      setSymbolDialogOpen,
      setScreenshotDialogOpen,
      setIconsDialogOpen,
      setEquationDialogOpen,
      openRecommendedCharts: () => {
        void handleRecommendedChartsImpl(visualContext()).then((result) => {
          if (result) setRecommendedCharts(result)
        })
      },
      setPendingEdits,
      visualContext,
      dataToolsContext,
      pivotContext,
      handlePageLayoutCommand: (rest) => handlePageLayoutCommandImpl(pageLayoutContext(), rest),
      handleExportPdf: () => handleExportPdfImpl(pageLayoutContext()),
    }
  }

  function handleRibbonCommand(command: string): void {
    handleRibbonCommandImpl(ribbonContext(), command)
  }

  function selectionStyle(
    range: NonNullable<ReturnType<ActiveWorkbook['getActiveRange']>>,
  ): IStyleData {
    // Resolves interned style references and merges row/col/sheet styles —
    // raw getCellData().s can be a style-id string with no fields on it.
    return range.getCellStyleData() ?? {}
  }

  function anchorCellValue(): number | string | null {
    try {
      const value = univerRef.current?.univerAPI.getActiveWorkbook()?.getActiveRange()?.getValue()
      return typeof value === 'number' || typeof value === 'string' ? value : null
    } catch {
      return null
    }
  }

  refreshSelectionFormatRef.current = () => {
    let range: ReturnType<ActiveWorkbook['getActiveRange']> | undefined
    try {
      range = univerRef.current?.univerAPI.getActiveWorkbook()?.getActiveRange()
    } catch {
      // Sheet changes briefly retain the previous sheet's selection. If that
      // row or column is outside the new sheet, Univer rejects the stale
      // range; the next selection event will refresh the ribbon normally.
      return
    }
    if (!range) {
      setSelectionFormat(null)
      setActiveCellA1('')
      return
    }
    setActiveCellA1(`${columnLetter(range.getColumn())}${range.getRow() + 1}`)
    let pattern: string
    try {
      pattern = range.getNumberFormat()
    } catch {
      // A disposing workbook can race the read; keep the last echo.
      return
    }
    const next = toSelectionFormat(selectionStyle(range), pattern, selectionLinkTarget(range))
    setSelectionFormat((previous) => (selectionFormatEquals(previous, next) ? previous : next))
  }

  function selectionLinkTarget(
    range: NonNullable<ReturnType<ActiveWorkbook['getActiveRange']>>,
  ): string | null {
    const state = lazyWorkbookRef.current
    const sheetId = univerRef.current?.univerAPI.getActiveWorkbook()?.getActiveSheet()?.getSheetId()
    if (!state || !sheetId) return null
    const row = range.getRow()
    const column = range.getColumn()
    const journaled = hyperlinkEditAt(state.editJournal, sheetId, row, column)
    if (journaled !== undefined) return journaled
    return state.hyperlinkTargets.get(sheetId)?.get(`${row}:${column}`) ?? null
  }

  function openLazyWorkbook(opened: WorkbookFile): void {
    const selected: WorkbookFile = {
      ...opened,
      visuals: opened.visuals.map((visual) =>
        visual.kind === 'chart' && visual.chart !== undefined
          ? { ...visual, chart: withDefaultBarLabels(visual.chart) }
          : visual,
      ),
    }
    const previous = lazyWorkbookRef.current
    if (previous) {
      clearLazyState(previous)
      void window.desktopApi.closeWorkbook(previous.file.sessionId).catch(() => undefined)
    }
    if (demoVisualInstallTimerRef.current) {
      clearTimeout(demoVisualInstallTimerRef.current)
      demoVisualInstallTimerRef.current = null
    }
    disposeVisuals(demoVisualDisposablesRef.current)
    demoVisualDisposablesRef.current = []
    const state: LazyWorkbookState = {
      file: selected,
      generation: Date.now(),
      loadedRanges: new Map(),
      loadingKeys: new Map(),
      retryTimers: new Map(),
      appliedMerges: new Map(),
      appliedRowKeys: new Map(),
      sheetProtections: new Map(),
      uninstalledDefinedNames: new Set(),
      appliedCfSheets: new Set(),
      appliedFilterSheets: new Set(),
      appliedDvSheets: new Set(),
      hyperlinkTargets: new Map(),
      frozenStripKeys: new Map(),
      filterOrigins: new Map(),
      showFormulaSheets: new Set(
        selected.sheets.filter((sheet) => sheet.showFormulas).map((sheet) => sheet.id),
      ),
      formulaMode:
        selected.sheets.reduce((sum, sheet) => sum + sheet.rowCount * sheet.columnCount, 0) <=
        FORMULA_MODE_MAX_CELLS,
      editJournal: createEditJournal(),
      flags: { preloadComplete: false },
      closure: { status: 'idle', pinned: new Map() },
      formulaText: new Map(),
      pivotDefinitions: new Map(),
      outline: new Map(),
      recalc: {
        timer: null,
        generation: 0,
        failures: 0,
        formulaCells: new Map(),
        overlay: new Map(),
      },
    }
    // Column outline levels arrive with the sheet metadata; seed them now.
    for (const sheet of selected.sheets) {
      for (const columnWidth of sheet.columnWidths) {
        if (columnWidth.outlineLevel === undefined && !columnWidth.collapsed) continue
        const cols = sheetOutline(state, sheet.id).cols
        const endColumn = Math.min(columnWidth.endColumn, sheet.columnCount - 1)
        for (let column = columnWidth.startColumn; column <= endColumn; column += 1) {
          cols.set(column, {
            level: columnWidth.outlineLevel ?? 0,
            collapsed: columnWidth.collapsed ?? false,
            hidden: columnWidth.hidden,
          })
        }
      }
    }
    lazyWorkbookRef.current = state
    let resolveInitialWorkbookLoad: () => void = () => undefined
    initialWorkbookLoadRef.current = new Promise<void>((resolve) => {
      resolveInitialWorkbookLoad = resolve
    })
    // Pivot definitions load eagerly so refresh (a synchronous apply step)
    // never waits on IPC. Best effort: a failed parse just disables refresh.
    for (const sheet of selected.sheets) {
      for (const pivot of sheet.pivotTables) {
        if (pivot.cachePath === null) continue
        void window.desktopApi
          .readPivotDefinition({
            sessionId: selected.sessionId,
            path: pivot.path,
            cachePath: pivot.cachePath,
          })
          .then((definition) => {
            if (lazyWorkbookRef.current === state) {
              state.pivotDefinitions.set(pivot.path, definition)
            }
          })
          .catch(() => undefined)
      }
    }
    // Dev-only diagnosis hooks: e2e drivers dump journal state and dispatch
    // Univer commands (drag interactions are hard to synthesize over CDP).
    if (import.meta.env.DEV) {
      ;(window as unknown as Record<string, unknown>).__journal = state.editJournal
      ;(window as unknown as Record<string, unknown>).__univerAPI = univerRef.current?.univerAPI
    }
    setRevision(0)
    setPendingEdits(0)
    // Slicers/timelines belong to the previous workbook's session only;
    // switching files invalidates them.
    setSlicers([])
    setSlicerPicker(null)
    setTimelines([])
    setTimelinePicker(null)
    disposeVisuals(visualDisposablesRef.current)
    loadWorkbookSkeleton(univerRef.current, selected)
    applyWorkbookNotes(univerRef.current, selected)
    applyDefinedNames(univerRef.current, selected, state)
    const runtime = univerRef.current
    if (runtime) {
      requestAnimationFrame(() => {
        if (lazyWorkbookRef.current !== state) {
          resolveInitialWorkbookLoad()
          return
        }
        const workbook = runtime.univerAPI.getActiveWorkbook()
        if (!workbook) {
          resolveInitialWorkbookLoad()
          return
        }
        // Register existing file tables so Univer renders filter dropdowns
        // and banding. This is visual-only (the journal is empty for file
        // tables), so failures are swallowed — the data is still usable.
        for (const sheet of selected.sheets) {
          if (sheet.tables.length === 0) continue
          const ws = workbook.getSheetBySheetId(sheet.id)
          if (!ws) continue
          for (let index = 0; index < sheet.tables.length; index += 1) {
            const table = sheet.tables[index]!
            const tableId = `file-table-${sheet.id}-${index}`
            const tableName = `Table${index + 1}_${sheet.id.slice(0, 6)}`
            try {
              void ws.addTable(tableName, table.range, tableId)
            } catch {
              // Best-effort: skip if Univer rejects (e.g. overlapping ranges)
            }
          }
        }
        const worksheet = workbook.getActiveSheet()
        if (!worksheet) {
          resolveInitialWorkbookLoad()
          return
        }
        // apply the opening sheet's formula view (sheetView/@showFormulas)
        applyShowFormulasView(runtime, state, worksheet.getSheetId())
        queueVisualInstall(
          runtime,
          lazyWorkbookRef,
          visualDisposablesRef,
          visualInstallTimerRef,
          worksheet.getSheetId(),
          chartEditRef,
          chartVectorRef,
          shapeEditRef,
        )
        try {
          worksheet.scrollToCell(0, 0)
        } catch {
          // A workbook opened during startup (the shell's queued-open nudge)
          // can land before Univer's Rendered lifecycle registers the scroll
          // render controller, and the facade then throws a redi
          // QuantityCheckError. The fresh view is already at the origin, so
          // skipping the reset is harmless.
        }
        const initialLoads = [loadVisibleRange(runtime, lazyWorkbookRef, worksheet, setMessage)]
        if (state.formulaMode) {
          initialLoads.push(preloadEntireWorkbook(runtime, lazyWorkbookRef, setMessage))
        } else {
          // Deferred so first paint and initial streaming win the sidecar.
          setTimeout(() => {
            void activateFormulaClosure(runtime, lazyWorkbookRef, setMessage)
          }, 1500)
        }
        void Promise.all(initialLoads).finally(resolveInitialWorkbookLoad)
      })
    }
  }

  function trackWorkbookOpen(opened: WorkbookFile): Promise<void> {
    openLazyWorkbook(opened)
    const ready = initialWorkbookLoadRef.current
    workbookReadyRef.current = ready
    return ready
  }

  async function handleInspectWorkbook(): Promise<void> {
    if (workbookOpeningRef.current) return
    workbookOpeningRef.current = true
    try {
      if (!window.desktopApi) {
        throw new Error(t('appBridgeUnavailable'))
      }
      const opening = window.desktopApi.selectWorkbook().then(async (selected) => {
        if (!selected) {
          setMessage(t('appOpenCanceled'))
          return
        }
        await trackWorkbookOpen(selected)
        setMessage(t('appOpened', { name: selected.name }))
      })
      workbookReadyRef.current = opening
      await opening
    } catch (error: unknown) {
      setMessage(error instanceof Error ? error.message : t('appOpenFailed'))
    } finally {
      workbookOpeningRef.current = false
    }
  }

  async function handleSave(
    mode: 'save' | 'save-as' | 'recovery',
    quiet = false,
  ): Promise<SaveActionResult | undefined> {
    return handleSaveImpl(saveContext(), mode, quiet)
  }
  closeSaveRef.current = async () => {
    const state = lazyWorkbookRef.current
    if (!state || journalSize(state.editJournal) === 0) {
      window.desktopApi?.reportCloseSaveResult?.(true)
      return
    }
    await handleSave('save')
    // handleSave swallows errors into the status bar; a drained journal
    // (fresh state after openLazyWorkbook) is the success signal.
    const after = lazyWorkbookRef.current
    window.desktopApi?.reportCloseSaveResult?.(
      after === null || journalSize(after.editJournal) === 0,
    )
  }
  function sortColumnOptions(): { label: string; colIndex: number }[] {
    const range = univerRef.current?.univerAPI.getActiveWorkbook()?.getActiveRange()
    if (!range) return []
    const start = range.getColumn()
    const width = Math.min(range.getWidth(), 26)
    return Array.from({ length: width }, (_, offset) => ({
      label: t('appColumnLabel', { col: columnLabel(start + offset) }),
      colIndex: start + offset,
    }))
  }

  menuActionRef.current = (action) => {
    if (action === 'open') {
      void handleInspectWorkbook()
    } else if (action === 'export-pdf') {
      void handleExportPdfImpl(pageLayoutContext())
    } else if (action === 'undo' || action === 'redo') {
      // Dialog text fields keep native text undo; everywhere else ⌘Z means
      // workbook history.
      const active = document.activeElement
      if (active instanceof HTMLTextAreaElement || active instanceof HTMLInputElement) {
        document.execCommand(action)
      } else if (action === 'undo') {
        void univerRef.current?.univerAPI.undo()
      } else {
        void univerRef.current?.univerAPI.redo()
      }
    } else {
      void handleSave(action)
    }
  }
  handleSaveRef.current = handleSave
  /// Re-renders the floating visuals after a journal mutation (edits and
  /// their undo/redo closures share it).
  function refreshLazyVisuals(state: LazyWorkbookState): void {
    const runtime = univerRef.current
    if (!runtime || lazyWorkbookRef.current !== state) return
    setPendingEdits(journalSize(state.editJournal))
    setVisualEditTick((tick) => tick + 1)
    const sheetId = runtime.univerAPI.getActiveWorkbook()?.getActiveSheet()?.getSheetId()
    if (sheetId) {
      queueVisualInstall(
        runtime,
        lazyWorkbookRef,
        visualDisposablesRef,
        visualInstallTimerRef,
        sheetId,
        chartEditRef,
        chartVectorRef,
        shapeEditRef,
      )
    }
  }

  function refreshDemoVisuals(): void {
    if (lazyWorkbookRef.current) return
    setVisualEditTick((tick) => tick + 1)
    queueDemoVisualInstallForActiveSheet()
  }

  const chartSyncRef = useRef<{
    timer: ReturnType<typeof setTimeout> | null
    dirty: Map<string, CellBounds>
  }>({ timer: null, dirty: new Map() })

  function queueChartDataSync(sheetId: string, bounds: CellBounds): void {
    queueChartDataSyncImpl(visualSyncContext(), sheetId, bounds)
  }

  chartEditRef.current = (editKey, edit) => applyChartEditImpl(visualSyncContext(), editKey, edit)
  chartVectorRef.current = (editKey, rangeText) =>
    readChartVectorImpl(visualSyncContext(), editKey, rangeText)
  shapeEditRef.current = (visualId, changes) =>
    applyShapeEditImpl(visualSyncContext(), visualId, changes)

  // Ribbon echo of the selected chart; a live lookup so deletion, file
  // switches, and pending type edits reflect without extra bookkeeping.
  const selectedChart = (() => {
    if (!selectedVisual || selectedVisual.kind !== 'chart') return null
    const state = lazyWorkbookRef.current
    const live: WorkbookVisualObject | undefined = state
      ? [...state.file.visuals, ...state.editJournal.visualAdds].find(
          (candidate) => candidate.id === selectedVisual.id,
        )
      : (adapterRef.current.findVisual(selectedVisual.id) ?? undefined)
    if (!live?.chart || (state && state.editJournal.visualEdits.get(live.id)?.remove)) return null
    const pending =
      state && live.chartPath ? state.editJournal.chartEdits.get(live.chartPath) : undefined
    const currentChart = pending ? applyChartStateEdit(live.chart, pending) : live.chart
    const convertible = convertibleType(live.chart)
    const currentType = pending?.chartType ?? convertible
    const isPie =
      currentType !== null
        ? currentType === 'pie' || currentType === 'doughnut'
        : live.chart.chartTypes.some((type) => type.includes('pie') || type.includes('doughnut'))
    return {
      title: pending?.title ?? live.chart.title,
      convertible,
      currentType,
      canEdit: !state || live.chartPath !== undefined || live.id.startsWith('added-'),
      isPie,
      hasAxes: currentType !== null && !isPie,
      // A pending conversion always lands on a labelable family.
      canLabel: pending?.chartType !== undefined || chartSupportsDataLabels(live.chart.chartTypes),
      seriesCount: currentChart.series.length,
      categoryCount: currentChart.series[0]?.categories.length ?? 0,
      series: currentChart.series,
      legend: pending?.legend ?? live.chart.legend,
      axisTitles: { ...live.chart.axisTitles, ...pending?.axisTitles },
      dataLabels: pending?.dataLabels ?? live.chart.dataLabels,
      grouping: pending?.grouping ?? live.chart.grouping,
    }
  })()

  const activePageLayout = (() => {
    const worksheet = univerRef.current?.univerAPI.getActiveWorkbook()?.getActiveSheet()
    const fileSheet = worksheet
      ? lazyWorkbookRef.current?.file.sheets.find(
          (candidate) => candidate.id === worksheet.getSheetId(),
        )
      : undefined
    const baseState =
      (
        fileSheet as
          | (typeof fileSheet & {
              pageSetup?: PageSetupJournalState
            })
          | undefined
      )?.pageSetup ?? {}
    const journalState = worksheet
      ? (lazyWorkbookRef.current?.editJournal.pageSetup.get(worksheet.getSheetId()) ?? {})
      : {}
    return {
      ...baseState,
      ...journalState,
      showGridlines:
        journalState.showGridlines ??
        baseState.showGridlines ??
        (worksheet ? !worksheet.hasHiddenGridLines() : true),
    }
  })()

  // Chart panels resolve their chart live (pending edits applied), so every
  // control reflects the state the next save would write.
  const chartDialogTarget = (() => {
    if (!chartDialog) return null
    const state = lazyWorkbookRef.current
    const live: WorkbookVisualObject | undefined = state
      ? [...state.file.visuals, ...state.editJournal.visualAdds].find(
          (candidate) =>
            candidate.chartPath === chartDialog.editKey || candidate.id === chartDialog.editKey,
        )
      : (adapterRef.current.findVisual(chartDialog.editKey) ?? undefined)
    // A visual pending removal must not keep a live panel producing edits.
    if (!live?.chart || (state && state.editJournal.visualEdits.get(live.id)?.remove)) return null
    const pending =
      state && live.chartPath ? state.editJournal.chartEdits.get(live.chartPath) : undefined
    return {
      visualId: live.id,
      chart: applyChartStateEdit(live.chart, pending),
      supported: chartSupportsSeriesReplace(live.chart.chartTypes),
    }
  })()

  return (
    <>
      <ToastHost />
      {chartDialog && chartDialogTarget && chartDialog.kind === 'format' && (
        <ChartFormatPane
          chart={chartDialogTarget.chart}
          element={
            chartElement?.visualId === chartDialogTarget.visualId ? chartElement.element : null
          }
          onEdit={(edit) => chartEditRef.current(chartDialog.editKey, edit)}
          onClose={() => setChartDialog(null)}
        />
      )}
      {chartDialog && chartDialogTarget && chartDialog.kind === 'select-data' && (
        <SelectDataDialog
          chart={chartDialogTarget.chart}
          supported={chartDialogTarget.supported}
          readVector={(range) => chartVectorRef.current(chartDialog.editKey, range)}
          onApply={(edit) => chartEditRef.current(chartDialog.editKey, edit)}
          onClose={() => setChartDialog(null)}
        />
      )}
      <ExcelShell
        onOpen={() => void handleInspectWorkbook()}
        fullscreen={display.mode === 'fullscreen'}
        onToggleFullscreen={() => void toggleLiveEditorFullscreen()}
        pageLayout={activePageLayout}
        selectionFormat={selectionFormat}
        statusMessage={message}
        onUndo={handleUndo}
        canUndo={univerHist.canUndo || (!lazyWorkbookRef.current && adapterRef.current.canUndo)}
        canRedo={univerHist.canRedo}
        onCommand={handleRibbonCommand}
        zoomPercent={zoomPercent}
        canSave={pendingEdits > 0}
        onSave={() => void handleSave('save')}
        onRedo={handleRedo}
        autoSave={autoSave}
        onAutoSaveChange={setAutoSave}
        selectedChart={selectedChart}
        onGetSortColumns={sortColumnOptions}
        onGetSheetProtection={sheetProtectionEcho}
        onGetDefinedNames={definedNameRows}
        onDefinedNameAction={handleDefinedNameAction}
        onGetPivotFields={() => pivotFieldOptionsImpl(pivotContext())}
        onGetSourceRange={() => getSourceRangeImpl(pivotContext())}
        onCreatePivot={(config) => handleCreatePivotImpl(pivotContext(), config)}
        onGetPivotEditSeed={() => pivotEditInitialImpl(pivotContext())}
        onEditPivot={(config) => handleEditPivotApplyImpl(pivotContext(), config)}
        onRefreshPivot={() => handleRefreshPivotImpl(pivotContext())}
        onIsSelectionInPivot={() => isSelectionInPivotImpl(pivotContext())}
        onGetActiveCell={() => activeCellLabelImpl(dataToolsContext())}
        onGetAnchorValue={anchorCellValue}
        activeCellA1={activeCellA1}
        onGoToReference={(ref) => goToReferenceImpl(dataToolsContext(), ref)}
        onListDefinedNames={() => listDefinedNamesImpl(dataToolsContext())}
        onApplyFormula={(formula) => handleApplyFormulaImpl(dataToolsContext(), formula)}
        onCreateSubtotal={(config) => handleCreateSubtotalImpl(dataToolsContext(), config)}
        onCreateConsolidate={(config) => handleCreateConsolidateImpl(dataToolsContext(), config)}
        onGetConsolidateDefault={() => consolidateDefaultReferenceImpl(dataToolsContext())}
        onApplyHeaderFooter={(result) => handleApplyHeaderFooterImpl(pageLayoutContext(), result)}
      />
      <input
        ref={fileInputRef}
        className="visually-hidden-file"
        type="file"
        accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        onChange={(event) => {
          const file = event.currentTarget.files?.[0] ?? null
          const pending = pendingFilePickerRef.current
          pendingFilePickerRef.current = null
          if (pending) {
            pending(file)
          } else if (file) {
            const opening = file
              .arrayBuffer()
              .then((data) => browserHostRef.current!.openBuffer(data, file.name))
              .then(trackWorkbookOpen)
              .then(() => setMessage(t('appOpened', { name: file.name })))
            workbookReadyRef.current = opening
            void opening.catch((error) =>
              setMessage(error instanceof Error ? error.message : t('appOpenFailed')),
            )
          }
          event.currentTarget.value = ''
        }}
      />
      {advancedFilterColumns !== null && (
        <AdvancedFilterDialog
          columns={advancedFilterColumns}
          onApply={(criteria) => handleApplyAdvancedFilterImpl(dataToolsContext(), criteria)}
          onClose={() => setAdvancedFilterColumns(null)}
        />
      )}
      {symbolDialogOpen && (
        <SymbolDialog
          onInsert={(char) => handleInsertSymbolImpl(dataToolsContext(), char)}
          onClose={() => setSymbolDialogOpen(false)}
        />
      )}
      {screenshotDialogOpen && (
        <ScreenshotDialog
          onInsert={(dataUrl, width, height) =>
            handleInsertScreenshot(visualContext(), dataUrl, width, height)
          }
          onClose={() => setScreenshotDialogOpen(false)}
        />
      )}
      {iconsDialogOpen && (
        <IconsDialog
          onInsert={(dataUrl, size, name) =>
            handleInsertIconImpl(visualContext(), dataUrl, size, name)
          }
          onClose={() => setIconsDialogOpen(false)}
        />
      )}
      {equationDialogOpen && (
        <EquationDialog
          onInsert={(dataUrl, width, height) =>
            handleInsertEquationImpl(visualContext(), dataUrl, width, height)
          }
          onClose={() => setEquationDialogOpen(false)}
        />
      )}
      {recommendedCharts !== null && (
        <RecommendedChartsDialog
          recommendations={recommendedCharts}
          onPick={(kind) => void handleInsertChartImpl(visualContext(), kind)}
          onClose={() => setRecommendedCharts(null)}
        />
      )}
      {slicerPicker !== null && (
        <SlicerFieldPicker
          fields={slicerPicker.fields}
          onPick={(field) => handleCreateSlicerImpl(pivotContext(), field)}
          onClose={() => setSlicerPicker(null)}
        />
      )}
      <SlicerPanels
        slicers={slicers}
        onToggle={(slicerId, member) => handleSlicerToggleImpl(pivotContext(), slicerId, member)}
        onSelectAll={(slicerId) => handleSlicerSelectAllImpl(pivotContext(), slicerId)}
        onRemove={(slicerId) => handleRemoveSlicerImpl(pivotContext(), slicerId)}
      />
      {timelinePicker !== null && (
        <TimelineFieldPicker
          fields={timelinePicker.fields}
          onPick={(field) => handleCreateTimelineImpl(pivotContext(), field)}
          onClose={() => setTimelinePicker(null)}
        />
      )}
      <TimelinePanels
        timelines={timelines}
        onRange={(timelineId, start, end) =>
          handleTimelineRangeImpl(pivotContext(), timelineId, { start, end })
        }
        onClear={(timelineId) => handleTimelineRangeImpl(pivotContext(), timelineId, null)}
        onRemove={(timelineId) => handleRemoveTimelineImpl(pivotContext(), timelineId)}
      />
    </>
  )

  function definedNameRows(): {
    names: DefinedNameRow[]
    sheets: { id: string; name: string }[]
  } {
    const workbook = univerRef.current?.univerAPI.getActiveWorkbook()
    const sheets =
      workbook?.getSheets().map((sheet) => ({
        id: sheet.getSheetId(),
        name: sheet.getSheetName(),
      })) ?? []
    const sheetNames = new Map(sheets.map((sheet) => [sheet.id, sheet.name]))
    const names = univerDefinedNames(univerRef.current).map((defined) => {
      const localSheetId = defined.getLocalSheetId()
      const scoped = localSheetId !== undefined && localSheetId !== 'AllDefaultWorkbook'
      return {
        name: defined.getName(),
        ref: defined.getFormulaOrRefString(),
        scopeSheetId: scoped ? localSheetId : null,
        scopeLabel: scoped ? (sheetNames.get(localSheetId) ?? localSheetId) : t('appScopeWorkbook'),
      }
    })
    return { names, sheets }
  }

  function handleDefinedNameAction(action: DefinedNameAction): string | null {
    const runtime = univerRef.current
    const workbook = runtime?.univerAPI.getActiveWorkbook()
    if (!runtime || !workbook || !lazyWorkbookRef.current) {
      return t('appNamesNeedFile')
    }
    try {
      applyWorkbookDefinedNameAction(runtime, action)
    } catch (error: unknown) {
      return error instanceof Error ? error.message : t('appNameApplyFailed')
    }
    setMessage(t('appNamesUpdated'))
    return null
  }

  /// Effective protection of the active sheet: journal override, else file
  /// state; null while unknown (still indexing) or in the demo workbook.
  function sheetProtectionEcho(): boolean | null {
    const state = lazyWorkbookRef.current
    const sheetId = univerRef.current?.univerAPI.getActiveWorkbook()?.getActiveSheet()?.getSheetId()
    if (!state || !sheetId) return null
    const journaled = state.editJournal.sheetProtection.get(sheetId)
    if (journaled !== undefined) return journaled
    const file = state.sheetProtections.get(sheetId)
    if (file) return file.protected
    return state.editJournal.sheets.added.has(sheetId) ? false : null
  }
}
