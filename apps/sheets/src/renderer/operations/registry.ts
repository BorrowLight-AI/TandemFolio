import { validateJsonSchemaValue } from '@tandemfolio/operation-contract'
import {
  BooleanNumber,
  BorderStyleTypes,
  HorizontalAlign,
  VerticalAlign,
  WrapStrategy,
  type ICellData,
} from '@univerjs/core'
import { LexerTreeBuilder } from '@univerjs/engine-formula'

import { formatAddress, parseAddress, parseRange, rangeCellCount } from '../../domain/cell-address'
import type {
  AddChartOperation,
  AddPivotOperation,
  AddShapeOperation,
  AddTableColumnOperation,
  AddTableRowOperation,
  DeleteTableColumnOperation,
  DeleteTableRowOperation,
} from '../../domain/workbook-dsl'
import type { WorkbookChartEdit, WorkbookVisualObject } from '../../shared/desktop-api'
import { CELL_STYLE_PRESETS } from '../app-constants'
import { applyWorkbookColumnWidthCopy } from '../column-width-copy'
import {
  clearWorkbookConditionalFormats,
  isWorkbookVisualConditionalFormatValid,
  removeWorkbookConditionalFormat,
  setWorkbookConditionalFormatPriority,
  setWorkbookComparisonConditionalFormat,
  setWorkbookFormulaConditionalFormat,
  setWorkbookHighlightConditionalFormat,
  setWorkbookStatisticalConditionalFormat,
  setWorkbookVisualConditionalFormat,
  type WorkbookConditionalFormatComparisonOperator,
  type WorkbookConditionalFormatHighlightPredicate,
  type WorkbookConditionalFormatStatisticalDirection,
  type WorkbookConditionalFormatStatisticalKind,
  type WorkbookConditionalFormatStyle,
  type WorkbookConditionalFormatThreshold,
  type WorkbookConditionalFormatVisualInput,
  type WorkbookConditionalFormatVisualKind,
} from '../conditional-format-actions'
import { dedupeRows } from '../dedupe'
import {
  applyWorkbookCheckbox,
  applyWorkbookComparisonValidation,
  applyWorkbookCustomFormulaValidation,
  applyWorkbookListValidation,
  applyWorkbookListReferenceValidation,
  applyWorkbookValidationMessages,
  removeWorkbookDataValidation,
  type WorkbookComparisonValidationKind,
  type WorkbookComparisonValidationOperator,
  type WorkbookValidationMessages,
} from '../data-validation-actions'
import {
  isValidDefinedName,
  removeWorkbookDefinedName,
  setWorkbookDefinedName,
} from '../defined-name-actions'
import {
  applyWorkbookFilter,
  clearWorkbookFilterCriteria,
  setWorkbookCustomFilter,
  setWorkbookFilterValues,
  type WorkbookCustomFilterCondition,
} from '../filter-actions'
import { journalSize, recordNoteChange, type HeaderFooterParts } from '../edit-journal'
import { applyWorkbookFormulaView } from '../formula-view'
import {
  applyWorkbookFitToPages,
  applyWorkbookHeaderFooter,
  applyWorkbookPageMargins,
  applyWorkbookPageOrientation,
  applyWorkbookPaperSize,
  applyWorkbookPrintArea,
  applyWorkbookPrintGridlines,
  applyWorkbookPrintHeadings,
  applyWorkbookPrintScale,
  applyWorkbookPrintTitles,
  type WorkbookPageMargins,
  type WorkbookPageOrientation,
  type WorkbookPaperSize,
} from '../page-layout-actions'
import { INDENT_STEP_PX } from '../selection-format'
import {
  applyWorkbookOutlineDetailVisibility,
  applyWorkbookOutlineLevel,
  type WorkbookOutlineAxis,
} from '../outline-actions'
import { applyWorkbookSparklineAdd, type WorkbookSparklineType } from '../sparkline-actions'
import { applyWorkbookRangeCopyWithoutBorders, applyWorkbookRangeFormatCopy } from '../range-copy'
import {
  applyWorkbookAggregateFormula,
  applyWorkbookCellProtection,
  applyWorkbookCellFormula,
  applyWorkbookCellMatrix,
  applyWorkbookCellValue,
  applyWorkbookColumnWidth,
  applyWorkbookFlashFill,
  applyWorkbookFreeze,
  applyWorkbookGridlineVisibility,
  applyWorkbookHyperlink,
  applyWorkbookRowHeight,
  applyWorkbookSheetProtection,
  applyWorkbookTextToColumns,
  normalizeLinkTarget,
  pushWorkbookUndo,
  streamedWorkbookRectangleAvailable,
  type WorkbookAggregateFunction,
  type WorkbookTextToColumnsDelimiter,
} from '../univer-sync'
import { applyWorkbookTableAdd } from '../workbook-ops'
import {
  BORDER_COMMAND_TYPES,
  type ActiveWorkbook,
  type LazyWorkbookState,
  type UniverRuntime,
  type UniverWorksheet,
} from '../univer-state'
import { xlsxOperationCatalog } from './catalog'

type XlsxOperationDescriptor = (typeof xlsxOperationCatalog.operations)[number]
type XlsxOperationId = XlsxOperationDescriptor['id']

export interface XlsxOperationCommand {
  readonly operation: string
  readonly arguments: Record<string, unknown>
}

type XlsxTableMutation =
  | AddTableRowOperation
  | DeleteTableRowOperation
  | AddTableColumnOperation
  | DeleteTableColumnOperation
  | {
      readonly op: 'convert_table_to_range'
      readonly sheetId: string
      readonly tableName: string
    }

export interface XlsxOperationServices {
  readonly runtime: () => UniverRuntime | null
  readonly state?: () =>
    | (Pick<LazyWorkbookState, 'editJournal' | 'hyperlinkTargets' | 'sheetProtections'> &
        Partial<
          Pick<
            LazyWorkbookState,
            | 'file'
            | 'loadedRanges'
            | 'flags'
            | 'showFormulaSheets'
            | 'formulaText'
            | 'outline'
            | 'appliedCfSheets'
            | 'appliedDvSheets'
          >
        >)
    | null
  readonly setPendingEdits?: (count: number) => void
  readonly isSheetDataComplete?: (sheetId: string) => boolean
  readonly duplicateSheet?: (worksheet: UniverWorksheet, name: string) => UniverWorksheet
  readonly mutateTable?: (operation: XlsxTableMutation) => void | Promise<void>
  readonly replaceText?: (input: {
    readonly sheetId: string
    readonly range: string
    readonly find: string
    readonly replace: string
    readonly matchCase: boolean
    readonly wholeCell: boolean
  }) => number | Promise<number>
  readonly createSubtotals?: (input: {
    readonly sheetId: string
    readonly range: string
    readonly groupColumn: number
    readonly valueColumn: number
    readonly aggregation: 'sum' | 'count' | 'average'
  }) => number | Promise<number>
  readonly consolidate?: (input: {
    readonly targetSheetId: string
    readonly targetCell: string
    readonly sources: readonly { readonly sheetId: string; readonly range: string }[]
    readonly aggregation: 'sum' | 'count' | 'average' | 'max' | 'min'
    readonly leftLabels: boolean
  }) => { readonly rows: number; readonly columns: number } | Promise<{
    readonly rows: number
    readonly columns: number
  }>
  readonly readSourceFormulaText?: (input: {
    readonly sheetId: string
    readonly range: {
      readonly startRow: number
      readonly endRow: number
      readonly startColumn: number
      readonly endColumn: number
    }
  }) => Promise<ReadonlyMap<string, string>>
  readonly refreshSparklines?: (sheetId: string) => void
  readonly addChart?: (operation: AddChartOperation) => Promise<string>
  readonly addPivot?: (operation: AddPivotOperation) => string | Promise<string>
  readonly addPivotChart?: (input: {
    readonly pivotId: string
    readonly type: 'column' | 'bar' | 'line' | 'pie' | 'doughnut' | 'radar'
  }) =>
    | { readonly chartId: string; readonly truncated: boolean }
    | Promise<{ readonly chartId: string; readonly truncated: boolean }>
  readonly refreshPivots?: (sheetId: string) => number | Promise<number>
  readonly updatePivot?: (input: {
    readonly pivotId: string
    readonly targetCell: string
    readonly rowFields: readonly string[]
    readonly columnField?: readonly string[]
    readonly pageFields?: readonly string[]
    readonly values: AddPivotOperation['values']
  }) => void | Promise<void>
  readonly setPivotMemberFilter?: (input: {
    readonly pivotId: string
    readonly field: string
    readonly selectedValues: readonly XlsxCellScalar[] | null
  }) => number | Promise<number>
  readonly addShape?: (operation: AddShapeOperation) => Promise<string>
  readonly updateShape?: (
    shapeId: string,
    changes: {
      readonly anchor?: WorkbookVisualObject['anchor']
      readonly fillColor?: string
      readonly text?: string
    },
  ) => void | Promise<void>
  readonly moveVisual?: (
    visualId: string,
    anchor: WorkbookVisualObject['anchor'],
  ) => void | Promise<void>
  readonly editChart?: (
    editKey: string,
    edit: Omit<WorkbookChartEdit, 'chartPath'>,
  ) => void | Promise<void>
  readonly removeVisual?: (visualId: string) => void | Promise<void>
  readonly addImage?: (input: {
    readonly sheetId: string
    readonly anchorCell: string
    readonly name: string
    readonly mediaType: 'image/png' | 'image/jpeg' | 'image/gif'
    readonly data: ArrayBuffer
  }) => Promise<string>
  readonly addImagePath?: (input: {
    readonly sheetId: string
    readonly anchorCell: string
    readonly path: string
  }) => Promise<string>
  readonly loadStaged: (input: {
    readonly name: string
    readonly data: ArrayBuffer
  }) => void | Promise<void>
  readonly save: () => Promise<
    | { readonly ok: true; readonly fileName: string }
    | { readonly ok: false; readonly message: string }
  >
}

type XlsxOperationHandlerResult =
  | {
      readonly ok: true
      readonly output?: Readonly<Record<string, unknown>>
      readonly checkpointRecovery?: false
    }
  | {
      readonly ok: false
      readonly error: 'invalid_arguments' | 'execution_failed'
      readonly message: string
    }

type XlsxOperationHandler = (
  arguments_: Record<string, unknown>,
  services: XlsxOperationServices,
) => XlsxOperationHandlerResult | Promise<XlsxOperationHandlerResult>

export type XlsxOperationExecution =
  | { readonly handled: false }
  | {
      readonly handled: true
      readonly operationId: XlsxOperationId
      readonly ok: true
      readonly output?: Readonly<Record<string, unknown>>
      readonly checkpointRecovery?: false
    }
  | {
      readonly handled: true
      readonly operationId: XlsxOperationId
      readonly ok: false
      readonly error: 'invalid_arguments' | 'execution_failed'
      readonly message: string
    }

const descriptorsByName = new Map<string, XlsxOperationDescriptor>()
for (const descriptor of xlsxOperationCatalog.operations) {
  descriptorsByName.set(descriptor.id, descriptor)
  for (const alias of descriptor.compatibilityAliases) {
    descriptorsByName.set(alias, descriptor)
  }
}

function isArrayBuffer(value: unknown): value is ArrayBuffer {
  return Object.prototype.toString.call(value) === '[object ArrayBuffer]'
}

function stagedImageMediaType(
  name: string,
  data: ArrayBuffer,
): 'image/png' | 'image/jpeg' | 'image/gif' | null {
  const bytes = new Uint8Array(data)
  if (
    name.toLowerCase().endsWith('.png') &&
    bytes.length >= 8 &&
    [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a].every((value, index) => bytes[index] === value)
  ) {
    return 'image/png'
  }
  if (
    /\.jpe?g$/i.test(name) &&
    bytes.length >= 3 &&
    bytes[0] === 0xff &&
    bytes[1] === 0xd8 &&
    bytes[2] === 0xff
  ) {
    return 'image/jpeg'
  }
  if (
    name.toLowerCase().endsWith('.gif') &&
    bytes.length >= 6 &&
    /^GIF8[79]a$/.test(String.fromCharCode(...bytes.subarray(0, 6)))
  ) {
    return 'image/gif'
  }
  return null
}

type XlsxCellScalar = string | number | boolean | null
type XlsxAlignmentField = 'horizontal' | 'vertical' | 'wrap' | 'indent' | 'rotation'
type XlsxAlignment = {
  readonly horizontal?: 'left' | 'center' | 'right' | 'justify' | 'distributed' | null
  readonly vertical?: 'top' | 'middle' | 'bottom' | null
  readonly wrap?: boolean
  readonly indent?: number
  readonly rotation?: {
    readonly mode: 'angle' | 'none' | 'stacked'
    readonly degrees?: number
  }
}
type XlsxTextStyleField = 'bold' | 'italic' | 'strike' | 'underline'
type XlsxTextStyle = {
  readonly bold?: boolean
  readonly italic?: boolean
  readonly strike?: boolean
  readonly underline?: 'none' | 'single' | 'double'
}
type XlsxFontField = 'family' | 'size' | 'color'
type XlsxFont = {
  readonly family?: string | null
  readonly size?: number | null
  readonly color?: string | null
}
type XlsxProtectionField = 'locked' | 'hidden'
type XlsxProtection = {
  readonly locked?: boolean
  readonly hidden?: boolean
}
type XlsxStreamingState = Pick<LazyWorkbookState, 'file' | 'loadedRanges' | 'flags'>

const XLSX_RANGE_COPY_CELL_LIMIT = 20_000

const xlsxHorizontalAlignment: Record<string, HorizontalAlign | null> = {
  left: HorizontalAlign.LEFT,
  center: HorizontalAlign.CENTER,
  right: HorizontalAlign.RIGHT,
  justify: HorizontalAlign.JUSTIFIED,
  distributed: HorizontalAlign.DISTRIBUTED,
}
const xlsxVerticalAlignment: Record<string, VerticalAlign | null> = {
  top: VerticalAlign.TOP,
  middle: VerticalAlign.MIDDLE,
  bottom: VerticalAlign.BOTTOM,
}
const xlsxBorderLineStyles: Record<string, BorderStyleTypes> = {
  thin: BorderStyleTypes.THIN,
  medium: BorderStyleTypes.MEDIUM,
  thick: BorderStyleTypes.THICK,
  double: BorderStyleTypes.DOUBLE,
  hair: BorderStyleTypes.HAIR,
  dashed: BorderStyleTypes.DASHED,
  dotted: BorderStyleTypes.DOTTED,
}

function isXlsxAlignmentField(value: string): value is XlsxAlignmentField {
  return (
    value === 'horizontal' ||
    value === 'vertical' ||
    value === 'wrap' ||
    value === 'indent' ||
    value === 'rotation'
  )
}

function hasValidXlsxAlignmentValue(alignment: XlsxAlignment, field: XlsxAlignmentField): boolean {
  if (field === 'horizontal') {
    return (
      alignment.horizontal === null ||
      (typeof alignment.horizontal === 'string' && alignment.horizontal in xlsxHorizontalAlignment)
    )
  }
  if (field === 'vertical') {
    return (
      alignment.vertical === null ||
      (typeof alignment.vertical === 'string' && alignment.vertical in xlsxVerticalAlignment)
    )
  }
  if (field === 'wrap') return typeof alignment.wrap === 'boolean'
  if (field === 'indent') {
    return (
      Number.isInteger(alignment.indent) &&
      (alignment.indent as number) >= 0 &&
      (alignment.indent as number) <= 250
    )
  }

  const rotation = alignment.rotation
  if (!rotation) return false
  const rotationFields = Object.keys(rotation)
  if (rotation.mode === 'angle') {
    return (
      rotationFields.length === 2 &&
      rotationFields.includes('degrees') &&
      Number.isInteger(rotation.degrees) &&
      (rotation.degrees as number) >= -90 &&
      (rotation.degrees as number) <= 90
    )
  }
  return (rotation.mode === 'none' || rotation.mode === 'stacked') && rotationFields.length === 1
}

function xlsxAlignmentPatch(
  alignment: XlsxAlignment,
  fields: readonly XlsxAlignmentField[],
): Record<string, unknown> {
  const patch: Record<string, unknown> = {}
  for (const field of fields) {
    if (field === 'horizontal') {
      patch.ht =
        alignment.horizontal === null
          ? null
          : xlsxHorizontalAlignment[alignment.horizontal as string]
    }
    if (field === 'vertical') {
      patch.vt =
        alignment.vertical === null ? null : xlsxVerticalAlignment[alignment.vertical as string]
    }
    if (field === 'wrap') patch.tb = alignment.wrap ? WrapStrategy.WRAP : null
    if (field === 'indent') {
      patch.pd =
        alignment.indent === 0 ? null : { l: (alignment.indent as number) * INDENT_STEP_PX }
    }
    if (field === 'rotation') {
      patch.tr =
        alignment.rotation?.mode === 'stacked'
          ? { a: 0, v: BooleanNumber.TRUE }
          : alignment.rotation?.mode === 'none'
            ? null
            : { a: alignment.rotation?.degrees as number }
    }
  }
  return patch
}

function isXlsxCellScalar(value: unknown): value is XlsxCellScalar {
  return (
    value === null ||
    typeof value === 'string' ||
    typeof value === 'boolean' ||
    (typeof value === 'number' && Number.isFinite(value))
  )
}

function isXlsxCellMatrix(value: unknown): value is XlsxCellScalar[][] {
  return (
    Array.isArray(value) &&
    value.length > 0 &&
    value.length <= 20_000 &&
    value.every((row) => Array.isArray(row) && row.length > 0 && row.every(isXlsxCellScalar))
  )
}

function isXlsxTextStyleField(value: string): value is XlsxTextStyleField {
  return value === 'bold' || value === 'italic' || value === 'strike' || value === 'underline'
}

function hasValidXlsxTextStyleValue(style: XlsxTextStyle, field: XlsxTextStyleField): boolean {
  if (field === 'underline') {
    return (
      style.underline === 'none' || style.underline === 'single' || style.underline === 'double'
    )
  }
  return typeof style[field] === 'boolean'
}

function xlsxTextStylePatch(
  style: XlsxTextStyle,
  fields: readonly XlsxTextStyleField[],
): Record<string, unknown> {
  const patch: Record<string, unknown> = {}
  for (const field of fields) {
    if (field === 'bold') patch.bl = style.bold ? BooleanNumber.TRUE : null
    if (field === 'italic') patch.it = style.italic ? BooleanNumber.TRUE : null
    if (field === 'strike') {
      patch.st = style.strike ? { s: BooleanNumber.TRUE } : null
    }
    if (field === 'underline') {
      patch.ul =
        style.underline === 'none'
          ? null
          : {
              s: BooleanNumber.TRUE,
              ...(style.underline === 'double' ? { t: 10 } : {}),
            }
    }
  }
  return patch
}

function isXlsxFontField(value: string): value is XlsxFontField {
  return value === 'family' || value === 'size' || value === 'color'
}

function isXlsxProtectionField(value: string): value is XlsxProtectionField {
  return value === 'locked' || value === 'hidden'
}

function isXlsxColor(value: unknown): value is string {
  return typeof value === 'string' && /^#[0-9a-fA-F]{6}$/.test(value)
}

function isXlsxConditionalFormatStyleValid(format: WorkbookConditionalFormatStyle): boolean {
  const colors = [format.fillColor, format.fontColor].filter(
    (value): value is string => value !== undefined,
  )
  return (
    colors.every(isXlsxColor) &&
    (colors.length > 0 || format.bold === true || format.italic === true)
  )
}

function hasValidXlsxFontValue(font: XlsxFont, field: XlsxFontField): boolean {
  if (field === 'family') {
    return (
      font.family === null ||
      (typeof font.family === 'string' &&
        font.family.trim().length > 0 &&
        font.family.length <= 128)
    )
  }
  if (field === 'size') {
    return (
      font.size === null ||
      (typeof font.size === 'number' &&
        Number.isFinite(font.size) &&
        font.size >= 1 &&
        font.size <= 409)
    )
  }
  return font.color === null || isXlsxColor(font.color)
}

function xlsxFontPatch(font: XlsxFont, fields: readonly XlsxFontField[]): Record<string, unknown> {
  const patch: Record<string, unknown> = {}
  for (const field of fields) {
    if (field === 'family') patch.ff = font.family
    if (field === 'size') patch.fs = font.size
    if (field === 'color') patch.cl = font.color === null ? null : { rgb: font.color }
  }
  return patch
}

function xlsxCellStylePatch(preset: string): Record<string, unknown> | null {
  const patches = CELL_STYLE_PRESETS[preset]
  if (!patches) return null
  const style: Record<string, unknown> = {}
  for (const patch of patches) {
    if (patch.bold !== undefined) style.bl = patch.bold ? BooleanNumber.TRUE : null
    if (patch.italic !== undefined) style.it = patch.italic ? BooleanNumber.TRUE : null
    if (patch.underline !== undefined) {
      style.ul = patch.underline ? { s: BooleanNumber.TRUE } : null
    }
    if (patch.strikethrough !== undefined) {
      style.st = patch.strikethrough ? { s: BooleanNumber.TRUE } : null
    }
    if (patch.fontFamily !== undefined) style.ff = patch.fontFamily
    if (patch.fontSize !== undefined) style.fs = patch.fontSize
    if (patch.fontColor !== undefined) {
      style.cl = patch.fontColor === null ? null : { rgb: patch.fontColor }
    }
    if (patch.fillColor !== undefined) {
      style.bg = patch.fillColor === null ? null : { rgb: patch.fillColor }
    }
    if (patch.border) {
      const edge = {
        s: BorderStyleTypes.THIN,
        cl: { rgb: patch.border.color ?? '#000000' },
      }
      if (patch.border.type === 'all') style.bd = { t: edge, r: edge, b: edge, l: edge }
      if (patch.border.type === 'top') style.bd = { t: edge }
      if (patch.border.type === 'bottom') style.bd = { b: edge }
      if (patch.border.type === 'left') style.bd = { l: edge }
      if (patch.border.type === 'right') style.bd = { r: edge }
      if (patch.border.type === 'none') style.bd = { t: null, r: null, b: null, l: null }
    }
  }
  return style
}

function xlsxWorkbook(runtime: UniverRuntime | null): ActiveWorkbook {
  const workbook = runtime?.univerAPI.getActiveWorkbook()
  if (!workbook) throw new Error('Open an XLSX workbook first.')
  return workbook
}

function xlsxWorksheet(runtime: UniverRuntime | null, name: string): UniverWorksheet {
  const workbook = xlsxWorkbook(runtime)
  const worksheet = workbook.getSheets().find((candidate) => candidate.getSheetName() === name)
  if (!worksheet) throw new Error(`Unknown worksheet: ${name}`)
  workbook.setActiveSheet(worksheet)
  return worksheet
}

function xlsxOperationState(
  services: XlsxOperationServices,
): Pick<LazyWorkbookState, 'editJournal' | 'hyperlinkTargets' | 'sheetProtections'> &
  Partial<
    Pick<
      LazyWorkbookState,
      | 'file'
      | 'loadedRanges'
      | 'flags'
      | 'showFormulaSheets'
      | 'formulaText'
      | 'outline'
      | 'appliedCfSheets'
      | 'appliedDvSheets'
    >
  > {
  const state = services.state?.()
  if (!state) throw new Error('Open a file-backed XLSX workbook first.')
  return state
}

function xlsxStreamingState(services: XlsxOperationServices): XlsxStreamingState | null {
  const state = services.state?.()
  return state?.file && state.loadedRanges && state.flags
    ? { file: state.file, loadedRanges: state.loadedRanges, flags: state.flags }
    : null
}

function xlsxRangeDataAvailable(
  state: XlsxStreamingState | null,
  worksheet: UniverWorksheet,
  range: ReturnType<UniverWorksheet['getRange']>,
): boolean {
  return streamedWorkbookRectangleAvailable(
    state,
    worksheet.getSheetId(),
    range.getRow(),
    range.getRow() + range.getHeight() - 1,
    range.getColumn(),
    range.getColumn() + range.getWidth() - 1,
  )
}

function xlsxFormulaViewState(
  services: XlsxOperationServices,
): Pick<LazyWorkbookState, 'editJournal' | 'showFormulaSheets'> {
  const state = services.state?.()
  if (!state?.showFormulaSheets) throw new Error('Open a file-backed XLSX workbook first.')
  return { editJournal: state.editJournal, showFormulaSheets: state.showFormulaSheets }
}

function xlsxColumnLocation(column: string): { normalized: string; index: number } | null {
  if (!/^[A-Za-z]{1,3}$/.test(column)) return null
  const normalized = column.toUpperCase()
  return { normalized, index: parseAddress(`${normalized}1`).column }
}

function isXlsxSheetName(name: string): boolean {
  return (
    name.trim().length > 0 &&
    name.length <= 31 &&
    !/[:\\/?*[\]]/.test(name) &&
    !name.startsWith("'") &&
    !name.endsWith("'")
  )
}

function normalizeXlsxCellRange(range: string): string | undefined {
  const normalized = range.toUpperCase()
  if (!/^[A-Z]{1,3}[1-9][0-9]{0,6}(:[A-Z]{1,3}[1-9][0-9]{0,6})?$/.test(normalized)) {
    return undefined
  }
  const bounds = parseRange(normalized)
  if (bounds.endColumn >= 16_384 || bounds.endRow >= 1_048_576) return undefined
  const start = formatAddress(bounds.startRow, bounds.startColumn)
  if (!normalized.includes(':')) return start
  return `${start}:${formatAddress(bounds.endRow, bounds.endColumn)}`
}

type XlsxPivotValueArgument = {
  readonly field: string
  readonly aggregation: AddPivotOperation['values'][number]['agg']
  readonly numberFormat: string | null
  readonly showAs: AddPivotOperation['values'][number]['showDataAs'] | null
}

type XlsxPivotLayout = {
  readonly targetCell: string
  readonly rowFields: string[]
  readonly columnFields: string[]
  readonly pageFields: string[]
  readonly values: XlsxPivotValueArgument[]
}

function normalizeXlsxPivotLayout(
  operationId: 'xlsx.pivot.add' | 'xlsx.pivot.update',
  arguments_: Record<string, unknown>,
):
  | { readonly ok: true; readonly layout: XlsxPivotLayout }
  | Extract<XlsxOperationHandlerResult, { readonly ok: false }> {
  const targetCell = normalizeXlsxCellRange(arguments_.targetCell as string)
  if (!targetCell || targetCell.includes(':')) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: `${operationId} requires a valid single-cell targetCell.`,
    }
  }
  const rowFields = arguments_.rowFields as string[]
  const columnFields = arguments_.columnFields as string[]
  const pageFields = arguments_.pageFields as string[]
  const values = arguments_.values as XlsxPivotValueArgument[]
  const fieldNames = [
    ...rowFields,
    ...columnFields,
    ...pageFields,
    ...values.map((value) => value.field),
  ]
  if (fieldNames.some((field) => field.trim().length === 0 || field.length > 255)) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: `${operationId} requires non-empty field names of at most 255 characters.`,
    }
  }
  const normalized = (field: string): string => field.toLocaleLowerCase()
  const hasDuplicates = (fields: readonly string[]): boolean =>
    new Set(fields.map(normalized)).size !== fields.length
  if (hasDuplicates(rowFields) || hasDuplicates(columnFields) || hasDuplicates(pageFields)) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: `${operationId} dimension lists must not repeat fields.`,
    }
  }
  const dimensions = [...rowFields, ...columnFields, ...pageFields].map(normalized)
  if (new Set(dimensions).size !== dimensions.length) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: `${operationId} dimension fields must be disjoint.`,
    }
  }
  if (values.some((value) => dimensions.includes(normalized(value.field)))) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: `${operationId} values fields cannot also be dimensions.`,
    }
  }
  if (columnFields.length > 0 && values.length !== 1) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: `${operationId} requires exactly one values entry when columnFields are present.`,
    }
  }
  if (
    values.some(
      (value) =>
        value.numberFormat !== null &&
        (value.numberFormat.length === 0 || value.numberFormat.length > 255),
    )
  ) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: `${operationId} number formats must contain 1..255 characters.`,
    }
  }
  return { ok: true, layout: { targetCell, rowFields, columnFields, pageFields, values } }
}

function xlsxPivotValues(values: readonly XlsxPivotValueArgument[]): AddPivotOperation['values'] {
  return values.map((value) => ({
    field: value.field,
    agg: value.aggregation,
    ...(value.numberFormat === null ? {} : { numFmt: value.numberFormat }),
    ...(value.showAs === null ? {} : { showDataAs: value.showAs }),
  }))
}

function xlsxSessionTable(
  services: XlsxOperationServices,
  sheet: string,
  tableName: string,
) {
  if (tableName.trim().length === 0 || tableName.length > 255) {
    throw new Error('XLSX table names must contain 1..255 characters.')
  }
  const state = xlsxOperationState(services)
  const worksheet = xlsxWorksheet(services.runtime(), sheet)
  const table = state.editJournal.tableAdds.find(
    (candidate) =>
      candidate.sheetId === worksheet.getSheetId() &&
      candidate.name.toLowerCase() === tableName.toLowerCase(),
  )
  if (!table) {
    throw new Error(
      `XLSX table "${tableName}" is not a session-created table and cannot use this operation.`,
    )
  }
  return { state, worksheet, table }
}

function xlsxChartVisual(services: XlsxOperationServices, chartId: string) {
  const state = xlsxOperationState(services)
  const visual = [...(state.file?.visuals ?? []), ...state.editJournal.visualAdds].find(
    (candidate) => candidate.id === chartId,
  )
  if (
    !visual ||
    visual.kind !== 'chart' ||
    !visual.chart ||
    state.editJournal.visualEdits.get(visual.id)?.remove
  ) {
    throw new Error(`Unknown XLSX chart: ${chartId}`)
  }
  return visual as typeof visual & { chart: NonNullable<typeof visual.chart> }
}

function xlsxShapeVisual(services: XlsxOperationServices, shapeId: string) {
  const state = xlsxOperationState(services)
  const visual = [...(state.file?.visuals ?? []), ...state.editJournal.visualAdds].find(
    (candidate) => candidate.id === shapeId,
  )
  if (!visual || visual.kind !== 'shape' || state.editJournal.visualEdits.get(visual.id)?.remove) {
    throw new Error(`Unknown XLSX shape: ${shapeId}`)
  }
  return visual
}

function xlsxImageVisual(services: XlsxOperationServices, imageId: string) {
  const state = xlsxOperationState(services)
  const visual = [...(state.file?.visuals ?? []), ...state.editJournal.visualAdds].find(
    (candidate) => candidate.id === imageId,
  )
  if (!visual || visual.kind !== 'image' || state.editJournal.visualEdits.get(visual.id)?.remove) {
    throw new Error(`Unknown XLSX image: ${imageId}`)
  }
  return visual
}

type XlsxRangeCopyOperation =
  | 'xlsx.range.copy_values'
  | 'xlsx.range.copy_formulas'
  | 'xlsx.range.copy_formats'
  | 'xlsx.range.copy_without_borders'
type XlsxRangeCopyPair =
  | {
      readonly ok: true
      readonly runtime: UniverRuntime
      readonly sourceSheet: string
      readonly sourceRange: string
      readonly source: ReturnType<UniverWorksheet['getRange']>
      readonly destinationSheet: string
      readonly destinationRange: string
      readonly destination: ReturnType<UniverWorksheet['getRange']>
      readonly cellCount: number
    }
  | Extract<XlsxOperationHandlerResult, { readonly ok: false }>

function resolveXlsxRangeCopyPair(
  operation: XlsxRangeCopyOperation,
  arguments_: Record<string, unknown>,
  services: XlsxOperationServices,
): XlsxRangeCopyPair {
  const sourceSheet = arguments_.sourceSheet as string
  const sourceRange = normalizeXlsxCellRange(arguments_.sourceRange as string)
  const destinationSheet = arguments_.destinationSheet as string
  const destinationRange = normalizeXlsxCellRange(arguments_.destinationRange as string)
  if (!sourceRange || !destinationRange) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: `${operation} requires valid in-sheet A1 source and destination ranges.`,
    }
  }
  const runtime = services.runtime()
  if (!runtime) throw new Error('Open an XLSX workbook first.')
  const sourceWorksheet = xlsxWorksheet(runtime, sourceSheet)
  const source = sourceWorksheet.getRange(sourceRange)
  const destinationWorksheet = xlsxWorksheet(runtime, destinationSheet)
  const destination = destinationWorksheet.getRange(destinationRange)
  if (
    source.getHeight() !== destination.getHeight() ||
    source.getWidth() !== destination.getWidth()
  ) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: `${operation} requires equally shaped source and destination ranges.`,
    }
  }
  const cellCount = source.getHeight() * source.getWidth()
  if (cellCount > XLSX_RANGE_COPY_CELL_LIMIT) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: `${operation} supports at most 20,000 cells per operation.`,
    }
  }
  const streamingState = xlsxStreamingState(services)
  if (!xlsxRangeDataAvailable(streamingState, sourceWorksheet, source)) {
    return {
      ok: false,
      error: 'execution_failed',
      message: `${operation} requires its source range to finish streaming.`,
    }
  }
  if (!xlsxRangeDataAvailable(streamingState, destinationWorksheet, destination)) {
    return {
      ok: false,
      error: 'execution_failed',
      message: `${operation} cannot overwrite a destination range that is still streaming.`,
    }
  }
  return {
    ok: true,
    runtime,
    sourceSheet,
    sourceRange,
    source,
    destinationSheet,
    destinationRange,
    destination,
    cellCount,
  }
}

function normalizeXlsxPrintArea(range: string | null): string | null | undefined {
  if (range === null) return null
  return normalizeXlsxCellRange(range)
}

function normalizeXlsxPrintTitles(rows: string | null): string | null | undefined {
  if (rows === null) return null
  const match = /^([1-9][0-9]{0,6}):([1-9][0-9]{0,6})$/.exec(rows)
  if (!match) return undefined
  const start = Number(match[1])
  const end = Number(match[2])
  if (start > end || end > 1_048_576 || end - start > 20) return undefined
  return `${start}:${end}`
}

function isBoundedXlsxHeaderFooterParts(
  value: HeaderFooterParts | null,
): value is HeaderFooterParts | null {
  if (value === null) return true
  const sections = [value.left, value.center, value.right].filter(
    (section): section is string => section !== undefined,
  )
  return sections.length > 0 && sections.every((section) => section.length <= 1_024)
}

type XlsxDataValidationTarget =
  | {
      readonly ok: true
      readonly runtime: UniverRuntime
      readonly address: string
      readonly range: ReturnType<UniverWorksheet['getRange']>
    }
  | Extract<XlsxOperationHandlerResult, { readonly ok: false }>

function resolveXlsxDataValidationTarget(
  operation: string,
  sheet: string,
  rangeText: string,
  services: XlsxOperationServices,
): XlsxDataValidationTarget {
  const address = normalizeXlsxCellRange(rangeText)
  if (!address) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: `${operation} requires a valid in-sheet A1 range.`,
    }
  }
  const runtime = services.runtime()
  if (!runtime) throw new Error('Open an XLSX workbook first.')
  const state = xlsxOperationState(services)
  const worksheet = xlsxWorksheet(runtime, sheet)
  const range = worksheet.getRange(address)
  if (range.getHeight() * range.getWidth() > 10_000) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: `${operation} supports at most 10,000 cells per operation.`,
    }
  }
  if (!xlsxRangeDataAvailable(xlsxStreamingState(services), worksheet, range)) {
    return {
      ok: false,
      error: 'execution_failed',
      message: `${operation} requires its target range to finish streaming.`,
    }
  }
  if (
    !state.editJournal.sheets.added.has(worksheet.getSheetId()) &&
    !state.appliedDvSheets?.has(worksheet.getSheetId())
  ) {
    return {
      ok: false,
      error: 'execution_failed',
      message: `${operation} requires worksheet validation metadata to finish loading.`,
    }
  }
  return { ok: true, runtime, address, range }
}

type XlsxConditionalFormatTarget =
  | {
      readonly ok: true
      readonly runtime: UniverRuntime
      readonly worksheet: UniverWorksheet
      readonly address: string
      readonly range: ReturnType<UniverWorksheet['getRange']>
    }
  | Extract<XlsxOperationHandlerResult, { readonly ok: false }>

function resolveXlsxConditionalFormatTarget(
  operation: string,
  sheet: string,
  rangeText: string,
  services: XlsxOperationServices,
): XlsxConditionalFormatTarget {
  const address = normalizeXlsxCellRange(rangeText)
  if (!address) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: `${operation} requires a valid in-sheet A1 range.`,
    }
  }
  const runtime = services.runtime()
  if (!runtime) throw new Error('Open an XLSX workbook first.')
  const state = xlsxOperationState(services)
  const worksheet = xlsxWorksheet(runtime, sheet)
  const range = worksheet.getRange(address)
  if (range.getHeight() * range.getWidth() > 10_000) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: `${operation} supports at most 10,000 cells per operation.`,
    }
  }
  if (!xlsxRangeDataAvailable(xlsxStreamingState(services), worksheet, range)) {
    return {
      ok: false,
      error: 'execution_failed',
      message: `${operation} requires its target range to finish streaming.`,
    }
  }
  if (
    !state.editJournal.sheets.added.has(worksheet.getSheetId()) &&
    !state.appliedCfSheets?.has(worksheet.getSheetId())
  ) {
    return {
      ok: false,
      error: 'execution_failed',
      message: `${operation} requires worksheet conditional-format metadata to finish loading.`,
    }
  }
  return { ok: true, runtime, worksheet, address, range }
}

function resolveXlsxConditionalFormatSheet(
  operation: string,
  sheet: string,
  services: XlsxOperationServices,
):
  | { readonly ok: true; readonly worksheet: UniverWorksheet }
  | Extract<XlsxOperationHandlerResult, { readonly ok: false }> {
  const state = xlsxOperationState(services)
  const worksheet = xlsxWorksheet(services.runtime(), sheet)
  if (
    !state.editJournal.sheets.added.has(worksheet.getSheetId()) &&
    !state.appliedCfSheets?.has(worksheet.getSheetId())
  ) {
    return {
      ok: false,
      error: 'execution_failed',
      message: `${operation} requires worksheet conditional-format metadata to finish loading.`,
    }
  }
  return { ok: true, worksheet }
}

function parseXlsxIsoDate(value: string): Date | null {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(value)) return null
  const date = new Date(`${value}T00:00:00.000Z`)
  return Number.isFinite(date.getTime()) && date.toISOString().slice(0, 10) === value ? date : null
}

function parseXlsxTime(value: string): number | null {
  const match = /^(\d{2}):(\d{2})(?::(\d{2}))?$/.exec(value)
  if (!match) return null
  const hour = Number(match[1])
  const minute = Number(match[2])
  const second = Number(match[3] ?? 0)
  return hour <= 23 && minute <= 59 && second <= 59 ? hour * 3_600 + minute * 60 + second : null
}

function normalizeXlsxComparisonOperand(
  kind: WorkbookComparisonValidationKind,
  operand: string | number,
): { readonly formula: string; readonly rank: number } | null {
  if (kind === 'whole') {
    return typeof operand === 'number' && Number.isSafeInteger(operand)
      ? { formula: String(operand), rank: operand }
      : null
  }
  if (kind === 'decimal') {
    return typeof operand === 'number' && Number.isFinite(operand)
      ? { formula: String(operand), rank: operand }
      : null
  }
  if (kind === 'date') {
    if (typeof operand !== 'string') return null
    const date = parseXlsxIsoDate(operand)
    return date ? { formula: operand, rank: date.getTime() } : null
  }
  if (kind === 'time') {
    if (typeof operand !== 'string') return null
    const seconds = parseXlsxTime(operand)
    return seconds === null ? null : { formula: operand, rank: seconds }
  }
  return typeof operand === 'number' &&
    Number.isInteger(operand) &&
    operand >= 0 &&
    operand <= 32_767
    ? { formula: String(operand), rank: operand }
    : null
}

const handlers = {
  'xlsx.cell.set_value': (arguments_, services) => {
    const sheet = arguments_.sheet as string
    const address = arguments_.address as string
    const value = arguments_.value as XlsxCellScalar
    const range = xlsxWorksheet(services.runtime(), sheet).getRange(address)
    applyWorkbookCellValue(range, value)
    return {
      ok: true,
      output: { changed: 1, sheet, range: address.toUpperCase() },
    }
  },
  'xlsx.cell.set_formula': (arguments_, services) => {
    const sheet = arguments_.sheet as string
    const address = normalizeXlsxCellRange(arguments_.address as string)
    const formula = arguments_.formula as string
    if (!address || address.includes(':') || !formula.startsWith('=') || formula.length > 8_192) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'xlsx.cell.set_formula requires one valid cell and a 1..8,192 character formula beginning with "=".',
      }
    }
    applyWorkbookCellFormula(xlsxWorksheet(services.runtime(), sheet).getRange(address), formula)
    return {
      ok: true,
      output: { changed: 1, sheet, range: address, formula },
    }
  },
  'xlsx.chart.add': async (arguments_, services) => {
    const sheet = arguments_.sheet as string
    const dataRange = normalizeXlsxCellRange(arguments_.dataRange as string)
    const anchorCell =
      arguments_.anchorCell === undefined
        ? undefined
        : normalizeXlsxCellRange(arguments_.anchorCell as string)
    if (!dataRange || (anchorCell !== undefined && anchorCell.includes(':'))) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'xlsx.chart.add requires a valid dataRange and optional single-cell anchorCell.',
      }
    }
    if (rangeCellCount(parseRange(dataRange)) > 2_000) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'xlsx.chart.add supports at most 2,000 source cells.',
      }
    }
    if (!services.addChart) throw new Error('xlsx.chart.add is unavailable in this renderer.')
    const worksheet = xlsxWorksheet(services.runtime(), sheet)
    const type = arguments_.type as AddChartOperation['chartType']
    const visualId = await services.addChart({
      op: 'add_chart',
      sheetId: worksheet.getSheetId(),
      chartType: type,
      dataRange,
      ...(anchorCell === undefined ? {} : { anchorCell }),
    })
    return {
      ok: true,
      output: {
        sheet,
        dataRange,
        type,
        ...(anchorCell === undefined ? {} : { anchorCell }),
        visualId,
      },
    }
  },
  'xlsx.pivot.add': async (arguments_, services) => {
    const sourceSheet = arguments_.sourceSheet as string
    const targetSheet = arguments_.targetSheet as string
    const sourceRange = normalizeXlsxCellRange(arguments_.sourceRange as string)
    const normalizedLayout = normalizeXlsxPivotLayout('xlsx.pivot.add', arguments_)
    if (!normalizedLayout.ok) return normalizedLayout
    const { targetCell, rowFields, columnFields, pageFields, values } = normalizedLayout.layout
    if (!sourceRange) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'xlsx.pivot.add requires a valid sourceRange.',
      }
    }
    const sourceBounds = parseRange(sourceRange)
    const sourceRows = sourceBounds.endRow - sourceBounds.startRow + 1
    const sourceColumns = sourceBounds.endColumn - sourceBounds.startColumn + 1
    if (sourceRows < 2 || sourceRows > 10_001 || sourceColumns > 200) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'xlsx.pivot.add supports 2..10,001 source rows and at most 200 columns.',
      }
    }

    const name = arguments_.name as string | null
    if (name !== null && !/^[\p{L}_\\][\p{L}\p{N}_.\\ ]{0,254}$/u.test(name)) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'xlsx.pivot.add requires a valid optional PivotTable name.',
      }
    }
    if (!services.addPivot) throw new Error('xlsx.pivot.add is unavailable in this renderer.')
    const sourceWorksheet = xlsxWorksheet(services.runtime(), sourceSheet)
    const targetWorksheet = xlsxWorksheet(services.runtime(), targetSheet)
    if (
      !services.isSheetDataComplete?.(sourceWorksheet.getSheetId()) ||
      !services.isSheetDataComplete(targetWorksheet.getSheetId())
    ) {
      throw new Error('xlsx.pivot.add requires fully loaded source and target worksheets.')
    }
    const createdName = await services.addPivot({
      op: 'add_pivot',
      sheetId: sourceWorksheet.getSheetId(),
      sourceRange,
      targetSheetId: targetWorksheet.getSheetId(),
      targetCell,
      ...(name === null ? {} : { name }),
      rowFields,
      ...(columnFields.length === 0 ? {} : { columnField: columnFields }),
      ...(pageFields.length === 0 ? {} : { pageFields }),
      values: xlsxPivotValues(values),
    })
    return {
      ok: true,
      output: {
        sourceSheet,
        sourceRange,
        targetSheet,
        targetCell,
        name: createdName,
      },
    }
  },
  'xlsx.pivot.add_chart': async (arguments_, services) => {
    const pivotId = arguments_.pivotId as string
    const type = arguments_.type as 'column' | 'bar' | 'line' | 'pie' | 'doughnut' | 'radar'
    if (!/^xl\/pivotTables\/[A-Za-z0-9._-]+\.xml$/.test(pivotId)) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'xlsx.pivot.add_chart requires a valid published pivotId.',
      }
    }
    if (!services.addPivotChart) {
      throw new Error('xlsx.pivot.add_chart is unavailable in this renderer.')
    }
    const result = await services.addPivotChart({ pivotId, type })
    return {
      ok: true,
      output: { pivotId, type, chartId: result.chartId, truncated: result.truncated },
    }
  },
  'xlsx.pivot.refresh': async (arguments_, services) => {
    const sheet = arguments_.sheet as string
    if (!services.refreshPivots) {
      throw new Error('xlsx.pivot.refresh is unavailable in this renderer.')
    }
    const worksheet = xlsxWorksheet(services.runtime(), sheet)
    const refreshed = await services.refreshPivots(worksheet.getSheetId())
    return { ok: true, output: { sheet, refreshed } }
  },
  'xlsx.pivot.update': async (arguments_, services) => {
    const pivotId = arguments_.pivotId as string
    if (!/^xl\/pivotTables\/[A-Za-z0-9._-]+\.xml$/.test(pivotId)) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'xlsx.pivot.update requires a valid published pivotId.',
      }
    }
    const normalizedLayout = normalizeXlsxPivotLayout('xlsx.pivot.update', arguments_)
    if (!normalizedLayout.ok) return normalizedLayout
    if (!services.updatePivot) {
      throw new Error('xlsx.pivot.update is unavailable in this renderer.')
    }
    const { targetCell, rowFields, columnFields, pageFields, values } = normalizedLayout.layout
    await services.updatePivot({
      pivotId,
      targetCell,
      rowFields,
      ...(columnFields.length === 0 ? {} : { columnField: columnFields }),
      ...(pageFields.length === 0 ? {} : { pageFields }),
      values: xlsxPivotValues(values),
    })
    return { ok: true, output: { pivotId, targetCell } }
  },
  'xlsx.pivot.set_member_filter': async (arguments_, services) => {
    const pivotId = arguments_.pivotId as string
    const field = arguments_.field as string
    const selectedValues = arguments_.selectedValues as XlsxCellScalar[] | null
    if (!/^xl\/pivotTables\/[A-Za-z0-9._-]+\.xml$/.test(pivotId)) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'xlsx.pivot.set_member_filter requires a valid published pivotId.',
      }
    }
    if (field.trim().length === 0 || field.length > 255) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'xlsx.pivot.set_member_filter requires a bounded field name.',
      }
    }
    const valueKey = (value: XlsxCellScalar): string => `${typeof value}:${String(value)}`
    if (
      selectedValues !== null &&
      new Set(selectedValues.map(valueKey)).size !== selectedValues.length
    ) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'xlsx.pivot.set_member_filter selectedValues must be unique.',
      }
    }
    if (!services.setPivotMemberFilter) {
      throw new Error('xlsx.pivot.set_member_filter is unavailable in this renderer.')
    }
    const selectedCount = await services.setPivotMemberFilter({ pivotId, field, selectedValues })
    return {
      ok: true,
      output: { pivotId, field, selectedCount: selectedValues === null ? null : selectedCount },
    }
  },
  'xlsx.chart.update': async (arguments_, services) => {
    const chartId = arguments_.chartId as string
    if (chartId.length === 0 || chartId.length > 255) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'xlsx.chart.update requires a bounded chartId.',
      }
    }
    const textFields = [
      arguments_.title,
      arguments_.categoryAxisTitle,
      arguments_.valueAxisTitle,
    ].filter((value): value is string => typeof value === 'string')
    if (textFields.some((value) => value.length > 255)) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'XLSX chart titles are limited to 255 characters.',
      }
    }
    if (typeof arguments_.dataLabelFormat === 'string' && arguments_.dataLabelFormat.length > 64) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'XLSX chart data-label formats are limited to 64 characters.',
      }
    }
    const valueAxisMin = arguments_.valueAxisMin as number | null | undefined
    const valueAxisMax = arguments_.valueAxisMax as number | null | undefined
    if (
      typeof valueAxisMin === 'number' &&
      typeof valueAxisMax === 'number' &&
      valueAxisMin > valueAxisMax
    ) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'xlsx.chart.update requires valueAxisMin to be at most valueAxisMax.',
      }
    }
    const keys = [
      'title',
      'type',
      'legend',
      'dataLabels',
      'grouping',
      'categoryAxisTitle',
      'valueAxisTitle',
      'gridlines',
      'valueAxisMin',
      'valueAxisMax',
      'gapWidthPct',
      'holeSizePct',
      'explosionPct',
      'dataLabelPosition',
      'dataLabelFormat',
    ] as const
    if (!keys.some((key) => arguments_[key] !== undefined)) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'xlsx.chart.update requires at least one final-state property.',
      }
    }
    const visual = xlsxChartVisual(services, chartId)
    if (!services.editChart) throw new Error('xlsx.chart.update is unavailable in this renderer.')
    const edit: Omit<WorkbookChartEdit, 'chartPath'> = {
      ...(arguments_.title === undefined ? {} : { title: arguments_.title as string }),
      ...(arguments_.type === undefined
        ? {}
        : { chartType: arguments_.type as NonNullable<WorkbookChartEdit['chartType']> }),
      ...(arguments_.legend === undefined
        ? {}
        : { legend: arguments_.legend as NonNullable<WorkbookChartEdit['legend']> }),
      ...(arguments_.dataLabels === undefined
        ? {}
        : { dataLabels: arguments_.dataLabels as NonNullable<WorkbookChartEdit['dataLabels']> }),
      ...(arguments_.dataLabelPosition === undefined
        ? {}
        : {
            dataLabelPosition: arguments_.dataLabelPosition as NonNullable<
              WorkbookChartEdit['dataLabelPosition']
            >,
          }),
      ...(arguments_.dataLabelFormat === undefined
        ? {}
        : { dataLabelFormat: arguments_.dataLabelFormat as string }),
      ...(arguments_.grouping === undefined
        ? {}
        : { grouping: arguments_.grouping as NonNullable<WorkbookChartEdit['grouping']> }),
      ...(arguments_.categoryAxisTitle === undefined && arguments_.valueAxisTitle === undefined
        ? {}
        : {
            axisTitles: {
              ...(arguments_.categoryAxisTitle === undefined
                ? {}
                : { category: arguments_.categoryAxisTitle as string | null }),
              ...(arguments_.valueAxisTitle === undefined
                ? {}
                : { value: arguments_.valueAxisTitle as string | null }),
            },
          }),
      ...(arguments_.gridlines === undefined ? {} : { gridlines: arguments_.gridlines as boolean }),
      ...(valueAxisMin === undefined && valueAxisMax === undefined
        ? {}
        : {
            valueAxis: {
              ...(valueAxisMin === undefined ? {} : { min: valueAxisMin }),
              ...(valueAxisMax === undefined ? {} : { max: valueAxisMax }),
            },
          }),
      ...(arguments_.gapWidthPct === undefined
        ? {}
        : { gapWidthPct: arguments_.gapWidthPct as number }),
      ...(arguments_.holeSizePct === undefined
        ? {}
        : { holeSizePct: arguments_.holeSizePct as number }),
      ...(arguments_.explosionPct === undefined
        ? {}
        : { explosionPct: arguments_.explosionPct as number }),
    }
    await services.editChart(visual.chartPath ?? visual.id, edit)
    return { ok: true, output: { changed: 1, chartId } }
  },
  'xlsx.chart.remove': async (arguments_, services) => {
    const chartId = arguments_.chartId as string
    if (chartId.length === 0 || chartId.length > 255) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'xlsx.chart.remove requires a bounded chartId.',
      }
    }
    xlsxChartVisual(services, chartId)
    if (!services.removeVisual)
      throw new Error('xlsx.chart.remove is unavailable in this renderer.')
    await services.removeVisual(chartId)
    return { ok: true, output: { changed: 1, chartId } }
  },
  'xlsx.chart.set_colors': async (arguments_, services) => {
    const chartId = arguments_.chartId as string
    const seriesColors = arguments_.seriesColors as string[] | undefined
    const pointColors = arguments_.pointColors as string[] | undefined
    const validColor = (color: string) => /^#[0-9A-Fa-f]{6}$/.test(color)
    if (
      chartId.length === 0 ||
      chartId.length > 255 ||
      (seriesColors === undefined && pointColors === undefined) ||
      (seriesColors !== undefined &&
        (seriesColors.length === 0 ||
          seriesColors.length > 24 ||
          seriesColors.some((color) => !validColor(color)))) ||
      (pointColors !== undefined &&
        (pointColors.length === 0 ||
          pointColors.length > 64 ||
          pointColors.some((color) => !validColor(color))))
    ) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'xlsx.chart.set_colors requires 1..24 series or 1..64 point hex colors.',
      }
    }
    const visual = xlsxChartVisual(services, chartId)
    if (
      pointColors !== undefined &&
      !visual.chart.chartTypes.some((type) => type === 'pieChart' || type === 'doughnutChart')
    ) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'Point colors are supported only for pie and doughnut charts.',
      }
    }
    if (!services.editChart)
      throw new Error('xlsx.chart.set_colors is unavailable in this renderer.')
    await services.editChart(visual.chartPath ?? visual.id, {
      ...(seriesColors === undefined
        ? {}
        : { seriesColors: Object.fromEntries(seriesColors.map((color, index) => [index, color])) }),
      ...(pointColors === undefined
        ? {}
        : {
            pointColors: {
              0: Object.fromEntries(pointColors.map((color, index) => [index, color])),
            },
          }),
    })
    return { ok: true, output: { changed: 1, chartId } }
  },
  'xlsx.chart.set_series': async (arguments_, services) => {
    const chartId = arguments_.chartId as string
    const series = arguments_.series as NonNullable<WorkbookChartEdit['seriesSet']>
    const validColor = (color: string | undefined) =>
      color === undefined || /^#[0-9A-Fa-f]{6}$/.test(color)
    const validSeries =
      chartId.length > 0 &&
      chartId.length <= 255 &&
      series.length >= 1 &&
      series.length <= 24 &&
      series.every(
        (entry) =>
          entry.name.length <= 255 &&
          entry.values.length >= 1 &&
          entry.values.length <= 1_000 &&
          entry.values.every(Number.isFinite) &&
          (entry.categories === undefined ||
            (entry.categories.length === entry.values.length &&
              entry.categories.every((category) => category.length <= 255))) &&
          (entry.valuesRef === undefined || entry.valuesRef.length <= 512) &&
          (entry.categoriesRef === undefined || entry.categoriesRef.length <= 512) &&
          validColor(entry.color),
      )
    if (!validSeries) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'xlsx.chart.set_series requires 1..24 bounded series of 1..1,000 values.',
      }
    }
    const visual = xlsxChartVisual(services, chartId)
    if (!services.editChart)
      throw new Error('xlsx.chart.set_series is unavailable in this renderer.')
    await services.editChart(visual.chartPath ?? visual.id, { seriesSet: series })
    return { ok: true, output: { changed: 1, chartId, seriesCount: series.length } }
  },
  'xlsx.shape.add': async (arguments_, services) => {
    const sheet = arguments_.sheet as string
    const anchorCell = normalizeXlsxCellRange(arguments_.anchorCell as string)
    const fillColor = arguments_.fillColor as string | undefined
    const text = arguments_.text as string | undefined
    if (
      !isXlsxSheetName(sheet) ||
      !anchorCell ||
      anchorCell.includes(':') ||
      (fillColor !== undefined && !/^#[0-9A-Fa-f]{6}$/.test(fillColor)) ||
      (text !== undefined && text.length > 1_000)
    ) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'xlsx.shape.add requires a valid sheet, cell anchor, hex fill, and bounded text.',
      }
    }
    if (!services.addShape) throw new Error('xlsx.shape.add is unavailable in this renderer.')
    const worksheet = xlsxWorksheet(services.runtime(), sheet)
    const type = arguments_.type as AddShapeOperation['shapeType']
    const visualId = await services.addShape({
      op: 'add_shape',
      sheetId: worksheet.getSheetId(),
      shapeType: type,
      anchorCell,
      ...(fillColor === undefined ? {} : { fillColor }),
      ...(text === undefined ? {} : { text }),
    })
    return { ok: true, output: { sheet, type, anchorCell, visualId } }
  },
  'xlsx.shape.update': async (arguments_, services) => {
    const shapeId = arguments_.shapeId as string
    const anchorCell =
      arguments_.anchorCell === undefined
        ? undefined
        : normalizeXlsxCellRange(arguments_.anchorCell as string)
    const fillColor = arguments_.fillColor as string | undefined
    const text = arguments_.text as string | undefined
    if (
      shapeId.length === 0 ||
      shapeId.length > 255 ||
      (arguments_.anchorCell !== undefined && (!anchorCell || anchorCell.includes(':'))) ||
      (fillColor !== undefined && !/^#[0-9A-Fa-f]{6}$/.test(fillColor)) ||
      (text !== undefined && text.length > 1_000) ||
      (anchorCell === undefined && fillColor === undefined && text === undefined)
    ) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'xlsx.shape.update requires one bounded final-state property.',
      }
    }
    const visual = xlsxShapeVisual(services, shapeId)
    if (!services.updateShape) throw new Error('xlsx.shape.update is unavailable in this renderer.')
    const anchor = (() => {
      if (anchorCell === undefined) return undefined
      const base = parseAddress(anchorCell)
      return {
        ...visual.anchor,
        fromRow: base.row,
        fromColumn: base.column,
        toRow: base.row + (visual.anchor.toRow - visual.anchor.fromRow),
        toColumn: base.column + (visual.anchor.toColumn - visual.anchor.fromColumn),
      }
    })()
    await services.updateShape(shapeId, {
      ...(anchor === undefined ? {} : { anchor }),
      ...(fillColor === undefined ? {} : { fillColor }),
      ...(text === undefined ? {} : { text }),
    })
    return { ok: true, output: { changed: 1, shapeId } }
  },
  'xlsx.shape.remove': async (arguments_, services) => {
    const shapeId = arguments_.shapeId as string
    if (shapeId.length === 0 || shapeId.length > 255) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'xlsx.shape.remove requires a bounded shapeId.',
      }
    }
    xlsxShapeVisual(services, shapeId)
    if (!services.removeVisual)
      throw new Error('xlsx.shape.remove is unavailable in this renderer.')
    await services.removeVisual(shapeId)
    return { ok: true, output: { changed: 1, shapeId } }
  },
  'xlsx.image.add': async (arguments_, services) => {
    const sheet = arguments_.sheet as string
    const path = arguments_.path as string
    const anchorCell = normalizeXlsxCellRange(arguments_.anchorCell as string)
    if (
      !anchorCell ||
      anchorCell.includes(':') ||
      !path.startsWith('/') ||
      path.length > 1_024 ||
      !/\.(png|jpe?g|gif)$/i.test(path)
    ) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'xlsx.image.add requires an absolute PNG, JPEG, or GIF path and one cell anchor.',
      }
    }
    if (!services.addImagePath) throw new Error('xlsx.image.add requires Broker staging.')
    const worksheet = xlsxWorksheet(services.runtime(), sheet)
    const visualId = await services.addImagePath({
      sheetId: worksheet.getSheetId(),
      anchorCell,
      path,
    })
    return { ok: true, output: { sheet, anchorCell, visualId } }
  },
  'xlsx.image.add_staged': async (arguments_, services) => {
    const sheet = arguments_.sheet as string
    const anchorCell = normalizeXlsxCellRange(arguments_.anchorCell as string)
    const data = arguments_.data
    const name = arguments_.name as string
    if (
      !anchorCell ||
      anchorCell.includes(':') ||
      !isArrayBuffer(data) ||
      data.byteLength !== arguments_.size ||
      data.byteLength > 20 * 1024 * 1024
    ) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'xlsx.image.add_staged requires a valid bounded staged image descriptor.',
      }
    }
    const mediaType = stagedImageMediaType(name, data)
    if (!mediaType) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'xlsx.image.add_staged requires matching PNG, JPEG, or GIF bytes.',
      }
    }
    if (!services.addImage)
      throw new Error('xlsx.image.add_staged is unavailable in this renderer.')
    const worksheet = xlsxWorksheet(services.runtime(), sheet)
    const visualId = await services.addImage({
      sheetId: worksheet.getSheetId(),
      anchorCell,
      name,
      mediaType,
      data,
    })
    return { ok: true, output: { sheet, anchorCell, visualId } }
  },
  'xlsx.image.move': async (arguments_, services) => {
    const imageId = arguments_.imageId as string
    const anchorCell = normalizeXlsxCellRange(arguments_.anchorCell as string)
    if (imageId.length === 0 || imageId.length > 255 || !anchorCell || anchorCell.includes(':')) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'xlsx.image.move requires a bounded imageId and one final cell anchor.',
      }
    }
    const visual = xlsxImageVisual(services, imageId)
    if (!services.moveVisual) throw new Error('xlsx.image.move is unavailable in this renderer.')
    const base = parseAddress(anchorCell)
    await services.moveVisual(imageId, {
      ...visual.anchor,
      fromRow: base.row,
      fromColumn: base.column,
      toRow: base.row + (visual.anchor.toRow - visual.anchor.fromRow),
      toColumn: base.column + (visual.anchor.toColumn - visual.anchor.fromColumn),
    })
    return { ok: true, output: { changed: 1, imageId } }
  },
  'xlsx.image.remove': async (arguments_, services) => {
    const imageId = arguments_.imageId as string
    if (imageId.length === 0 || imageId.length > 255) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'xlsx.image.remove requires a bounded imageId.',
      }
    }
    xlsxImageVisual(services, imageId)
    if (!services.removeVisual)
      throw new Error('xlsx.image.remove is unavailable in this renderer.')
    await services.removeVisual(imageId)
    return { ok: true, output: { changed: 1, imageId } }
  },
  'xlsx.defined_name.set': (arguments_, services) => {
    const name = arguments_.name as string
    const previousName = arguments_.previousName as string | undefined
    const formula = (arguments_.formula as string).replace(/^=/, '')
    const scopeSheet = arguments_.scopeSheet as string | undefined
    if (
      !isValidDefinedName(name) ||
      (previousName !== undefined && !isValidDefinedName(previousName)) ||
      formula.length === 0 ||
      formula.length > 8_192 ||
      (scopeSheet !== undefined && !isXlsxSheetName(scopeSheet))
    ) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'xlsx.defined_name.set requires a valid name, scope, and bounded formula.',
      }
    }
    const runtime = services.runtime()
    if (!runtime) throw new Error('Open an XLSX workbook first.')
    const scopeSheetId =
      scopeSheet === undefined ? null : xlsxWorksheet(runtime, scopeSheet).getSheetId()
    setWorkbookDefinedName(runtime, {
      name,
      formula,
      scopeSheetId,
      ...(previousName === undefined ? {} : { previousName }),
    })
    return {
      ok: true,
      output: { changed: 1, name, formula, scope: scopeSheet ?? 'workbook' },
    }
  },
  'xlsx.defined_name.remove': (arguments_, services) => {
    const name = arguments_.name as string
    const scopeSheet = arguments_.scopeSheet as string | undefined
    if (!isValidDefinedName(name) || (scopeSheet !== undefined && !isXlsxSheetName(scopeSheet))) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'xlsx.defined_name.remove requires a valid name and scope.',
      }
    }
    const runtime = services.runtime()
    if (!runtime) throw new Error('Open an XLSX workbook first.')
    const scopeSheetId =
      scopeSheet === undefined ? null : xlsxWorksheet(runtime, scopeSheet).getSheetId()
    removeWorkbookDefinedName(runtime, name, scopeSheetId)
    return {
      ok: true,
      output: { changed: 1, name, scope: scopeSheet ?? 'workbook' },
    }
  },
  'xlsx.note.set': async (arguments_, services) => {
    const sheet = arguments_.sheet as string
    const address = normalizeXlsxCellRange(arguments_.address as string)
    const text = arguments_.text as string
    if (!address || address.includes(':') || text.length > 32_767) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'xlsx.note.set requires one valid cell address and at most 32,767 characters.',
      }
    }
    const runtime = services.runtime()
    const workbook = xlsxWorkbook(runtime)
    const worksheet = xlsxWorksheet(runtime, sheet)
    const position = parseAddress(address)
    worksheet.getRange(address).activate()
    const changed = await runtime!.univerAPI.executeCommand('sheet.command.update-note', {
      unitId: workbook.getId(),
      sheetId: worksheet.getSheetId(),
      row: position.row,
      col: position.column,
      note: {
        row: position.row,
        col: position.column,
        width: 220,
        height: 90,
        note: text,
      },
    })
    if (!changed) throw new Error('Univer rejected the note update.')
    const state = services.state?.()
    if (state) {
      recordNoteChange(state.editJournal, worksheet.getSheetId())
      services.setPendingEdits?.(journalSize(state.editJournal))
    }
    return { ok: true, output: { changed: 1, sheet, address } }
  },
  'xlsx.note.remove': async (arguments_, services) => {
    const sheet = arguments_.sheet as string
    const address = normalizeXlsxCellRange(arguments_.address as string)
    if (!address || address.includes(':')) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'xlsx.note.remove requires one valid cell address.',
      }
    }
    const runtime = services.runtime()
    const workbook = xlsxWorkbook(runtime)
    const worksheet = xlsxWorksheet(runtime, sheet)
    const position = parseAddress(address)
    worksheet.getRange(address).activate()
    const changed = await runtime!.univerAPI.executeCommand('sheet.command.delete-note', {
      unitId: workbook.getId(),
      sheetId: worksheet.getSheetId(),
      row: position.row,
      col: position.column,
    })
    if (!changed) throw new Error('The addressed cell has no note to remove.')
    const state = services.state?.()
    if (state) {
      recordNoteChange(state.editJournal, worksheet.getSheetId())
      services.setPendingEdits?.(journalSize(state.editJournal))
    }
    return { ok: true, output: { changed: 1, sheet, address } }
  },
  'xlsx.column.delete': (arguments_, services) => {
    const sheet = arguments_.sheet as string
    const column = arguments_.column as string
    const count = arguments_.count as number
    const location = xlsxColumnLocation(column)
    if (!location) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'xlsx.column.delete requires an A1 column label.',
      }
    }
    xlsxWorksheet(services.runtime(), sheet).deleteColumns(location.index, count)
    return { ok: true, output: { sheet, column: location.normalized, count } }
  },
  'xlsx.column.insert': (arguments_, services) => {
    const sheet = arguments_.sheet as string
    const column = arguments_.column as string
    const count = arguments_.count as number
    const location = xlsxColumnLocation(column)
    if (!location) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'xlsx.column.insert requires an A1 column label.',
      }
    }
    xlsxWorksheet(services.runtime(), sheet).insertColumnsBefore(location.index, count)
    return { ok: true, output: { sheet, column: location.normalized, count } }
  },
  'xlsx.column.set_width': (arguments_, services) => {
    const sheet = arguments_.sheet as string
    const column = arguments_.column as string
    const count = arguments_.count as number
    const widthCharacters = arguments_.widthCharacters as number
    const location = xlsxColumnLocation(column)
    if (!location) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'xlsx.column.set_width requires an A1 column label.',
      }
    }
    const worksheet = xlsxWorksheet(services.runtime(), sheet)
    if (location.index + count > worksheet.getMaxColumns()) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'xlsx.column.set_width must stay inside the worksheet column bounds.',
      }
    }
    const appliedWidthCharacters = applyWorkbookColumnWidth(
      worksheet,
      location.index,
      count,
      widthCharacters,
    )
    return {
      ok: true,
      output: {
        sheet,
        column: location.normalized,
        count,
        widthCharacters: appliedWidthCharacters,
      },
    }
  },
  'xlsx.column.set_visibility': (arguments_, services) => {
    const sheet = arguments_.sheet as string
    const column = arguments_.column as string
    const count = arguments_.count as number
    const visible = arguments_.visible as boolean
    const location = xlsxColumnLocation(column)
    if (!location) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'xlsx.column.set_visibility requires a valid start column.',
      }
    }
    const worksheet = xlsxWorksheet(services.runtime(), sheet)
    if (location.index + count > worksheet.getMaxColumns()) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'xlsx.column.set_visibility exceeds the worksheet column boundary.',
      }
    }
    if (visible) worksheet.showColumns(location.index, count)
    else worksheet.hideColumns(location.index, count)
    return { ok: true, output: { sheet, column: location.normalized, count, visible } }
  },
  'xlsx.column.copy_widths': (arguments_, services) => {
    const sourceSheet = arguments_.sourceSheet as string
    const sourceColumn = xlsxColumnLocation(arguments_.sourceColumn as string)
    const destinationSheet = arguments_.destinationSheet as string
    const destinationColumn = xlsxColumnLocation(arguments_.destinationColumn as string)
    const count = arguments_.count as number
    if (!sourceColumn || !destinationColumn) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'xlsx.column.copy_widths requires A1 source and destination column labels.',
      }
    }
    const runtime = services.runtime()
    if (!runtime) throw new Error('Open an XLSX workbook first.')
    const sourceWorksheet = xlsxWorksheet(runtime, sourceSheet)
    const destinationWorksheet = xlsxWorksheet(runtime, destinationSheet)
    if (
      sourceColumn.index + count > sourceWorksheet.getMaxColumns() ||
      destinationColumn.index + count > destinationWorksheet.getMaxColumns()
    ) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'xlsx.column.copy_widths must stay inside both worksheet column bounds.',
      }
    }
    if (
      !applyWorkbookColumnWidthCopy(
        runtime,
        sourceWorksheet,
        sourceColumn.index,
        destinationWorksheet,
        destinationColumn.index,
        count,
      )
    ) {
      return {
        ok: false,
        error: 'execution_failed',
        message: 'xlsx.column.copy_widths could not update the destination widths.',
      }
    }
    return {
      ok: true,
      output: {
        sourceSheet,
        sourceColumn: sourceColumn.normalized,
        destinationSheet,
        destinationColumn: destinationColumn.normalized,
        count,
      },
    }
  },
  'xlsx.formula.insert_aggregate': (arguments_, services) => {
    const sheet = arguments_.sheet as string
    const address = arguments_.range as string
    const aggregate = arguments_.function as WorkbookAggregateFunction
    const worksheet = xlsxWorksheet(services.runtime(), sheet)
    const targetRange = worksheet.getRange(address)
    if (targetRange.getHeight() < 2) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'xlsx.formula.insert_aggregate requires at least two rows.',
      }
    }
    const streamingState = xlsxStreamingState(services)
    const result = applyWorkbookAggregateFormula(streamingState, worksheet, targetRange, aggregate)
    if ('blocked' in result) {
      return {
        ok: false,
        error: 'execution_failed',
        message: 'xlsx.formula.insert_aggregate cannot overwrite a row that is still streaming.',
      }
    }
    return {
      ok: true,
      output: {
        sheet,
        range: address.toUpperCase(),
        targetRange: result.targetRange,
        function: aggregate,
        inserted: result.inserted,
      },
    }
  },
  'xlsx.range.flash_fill': (arguments_, services) => {
    const sheet = arguments_.sheet as string
    const address = arguments_.range as string
    const worksheet = xlsxWorksheet(services.runtime(), sheet)
    const range = worksheet.getRange(address)
    const streamingState = xlsxStreamingState(services)
    const result = applyWorkbookFlashFill(streamingState, worksheet, range)
    if ('blocked' in result) {
      if (result.blocked === 'needs_left') {
        return {
          ok: false,
          error: 'invalid_arguments',
          message: 'xlsx.range.flash_fill requires a source column to the left of the target.',
        }
      }
      if (result.blocked === 'needs_rows') {
        return {
          ok: false,
          error: 'invalid_arguments',
          message: 'xlsx.range.flash_fill requires at least two rows.',
        }
      }
      if (result.blocked === 'streaming') {
        return {
          ok: false,
          error: 'execution_failed',
          message: 'xlsx.range.flash_fill requires its source rows to finish streaming.',
        }
      }
      if (result.blocked === 'needs_examples') {
        return {
          ok: false,
          error: 'execution_failed',
          message: 'xlsx.range.flash_fill requires at least one non-empty target example.',
        }
      }
      if (result.blocked === 'no_pattern') {
        return {
          ok: false,
          error: 'execution_failed',
          message: 'xlsx.range.flash_fill could not infer a pattern from the target examples.',
        }
      }
      return {
        ok: false,
        error: 'execution_failed',
        message: 'xlsx.range.flash_fill has no empty target cells to fill.',
      }
    }
    return {
      ok: true,
      output: {
        sheet,
        range: address.toUpperCase(),
        targetRange: result.targetRange,
        filled: result.filled,
      },
    }
  },
  'xlsx.range.text_to_columns': async (arguments_, services) => {
    const sheet = arguments_.sheet as string
    const address = arguments_.range as string
    const delimiter = arguments_.delimiter as WorkbookTextToColumnsDelimiter
    const runtime = services.runtime()
    if (!runtime) throw new Error('Open an XLSX workbook first.')
    const worksheet = xlsxWorksheet(runtime, sheet)
    if (services.isSheetDataComplete && !services.isSheetDataComplete(worksheet.getSheetId())) {
      return {
        ok: false,
        error: 'execution_failed',
        message: 'xlsx.range.text_to_columns requires fully loaded sheet data.',
      }
    }
    const targetRange = worksheet.getRange(address)
    if (targetRange.getWidth() !== 1) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'xlsx.range.text_to_columns requires a single-column range.',
      }
    }
    const completed = await applyWorkbookTextToColumns(runtime, worksheet, targetRange, delimiter)
    if (!completed) {
      return {
        ok: false,
        error: 'execution_failed',
        message: 'Univer did not complete xlsx.range.text_to_columns.',
      }
    }
    return {
      ok: true,
      output: { sheet, range: address.toUpperCase(), delimiter },
    }
  },
  'xlsx.range.set_values': (arguments_, services) => {
    const sheet = arguments_.sheet as string
    const range = arguments_.range as string
    const values = arguments_.values
    if (!isXlsxCellMatrix(values)) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'xlsx.range.set_values requires a bounded scalar matrix.',
      }
    }
    const targetRange = xlsxWorksheet(services.runtime(), sheet).getRange(range)
    applyWorkbookCellMatrix(
      targetRange,
      values.map((row) => row.map((value) => ({ v: value })) satisfies ICellData[]),
    )
    return {
      ok: true,
      output: {
        changed: values.reduce((count, row) => count + row.length, 0),
        sheet,
        range: range.toUpperCase(),
      },
    }
  },
  'xlsx.range.replace_text': async (arguments_, services) => {
    const sheet = arguments_.sheet as string
    const address = normalizeXlsxCellRange(arguments_.range as string)
    const find = arguments_.find as string
    const replace = arguments_.replace as string
    const matchCase = arguments_.matchCase as boolean
    const wholeCell = arguments_.wholeCell as boolean
    if (!address || find.length === 0 || find.length > 255 || replace.length > 255) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'xlsx.range.replace_text requires a valid range and bounded find/replace text.',
      }
    }
    if (rangeCellCount(parseRange(address)) > 20_000) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'xlsx.range.replace_text supports at most 20,000 cells.',
      }
    }
    const worksheet = xlsxWorksheet(services.runtime(), sheet)
    const range = worksheet.getRange(address)
    if (!xlsxRangeDataAvailable(xlsxStreamingState(services), worksheet, range)) {
      throw new Error('xlsx.range.replace_text requires the complete target range to be loaded.')
    }
    if (!services.replaceText) {
      throw new Error('xlsx.range.replace_text is unavailable in this renderer.')
    }
    const changed = await services.replaceText({
      sheetId: worksheet.getSheetId(),
      range: address,
      find,
      replace,
      matchCase,
      wholeCell,
    })
    return { ok: true, output: { sheet, range: address, changed } }
  },
  'xlsx.range.insert_subtotals': async (arguments_, services) => {
    const sheet = arguments_.sheet as string
    const address = normalizeXlsxCellRange(arguments_.range as string)
    const groupColumn = xlsxColumnLocation(arguments_.groupColumn as string)
    const valueColumn = xlsxColumnLocation(arguments_.valueColumn as string)
    const aggregation = arguments_.aggregation as 'sum' | 'count' | 'average'
    if (!address || !groupColumn || !valueColumn || groupColumn.index === valueColumn.index) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message:
          'xlsx.range.insert_subtotals requires a valid range and two distinct worksheet columns.',
      }
    }
    const bounds = parseRange(address)
    if (
      bounds.endRow === bounds.startRow ||
      rangeCellCount(bounds) > 50_000 ||
      groupColumn.index < bounds.startColumn ||
      groupColumn.index > bounds.endColumn ||
      valueColumn.index < bounds.startColumn ||
      valueColumn.index > bounds.endColumn
    ) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message:
          'xlsx.range.insert_subtotals requires a header plus data rows, at most 50,000 cells, and both columns inside the range.',
      }
    }
    const worksheet = xlsxWorksheet(services.runtime(), sheet)
    const range = worksheet.getRange(address)
    if (!xlsxRangeDataAvailable(xlsxStreamingState(services), worksheet, range)) {
      throw new Error('xlsx.range.insert_subtotals requires the complete source range to be loaded.')
    }
    if (!services.createSubtotals) {
      throw new Error('xlsx.range.insert_subtotals is unavailable in this renderer.')
    }
    const groups = await services.createSubtotals({
      sheetId: worksheet.getSheetId(),
      range: address,
      groupColumn: groupColumn.index,
      valueColumn: valueColumn.index,
      aggregation,
    })
    return {
      ok: true,
      output: {
        sheet,
        range: address,
        groupColumn: groupColumn.normalized,
        valueColumn: valueColumn.normalized,
        aggregation,
        groups,
        insertedRows: groups + 1,
      },
    }
  },
  'xlsx.range.consolidate': async (arguments_, services) => {
    const targetSheet = arguments_.targetSheet as string
    const targetCell = normalizeXlsxCellRange(arguments_.targetCell as string)
    const sourceInputs = arguments_.sources as readonly {
      readonly sheet: string
      readonly range: string
    }[]
    const aggregation = arguments_.aggregation as 'sum' | 'count' | 'average' | 'max' | 'min'
    const leftLabels = arguments_.leftLabels as boolean
    if (!targetCell || targetCell.includes(':')) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'xlsx.range.consolidate requires a valid single-cell targetCell.',
      }
    }
    if (leftLabels && (aggregation === 'max' || aggregation === 'min')) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'xlsx.range.consolidate does not support max or min with leftLabels.',
      }
    }
    let totalCells = 0
    const runtime = services.runtime()
    const targetWorksheet = xlsxWorksheet(runtime, targetSheet)
    const sources: { sheetId: string; range: string }[] = []
    for (const source of sourceInputs) {
      const range = normalizeXlsxCellRange(source.range)
      if (!range) {
        return {
          ok: false,
          error: 'invalid_arguments',
          message: 'xlsx.range.consolidate requires valid in-sheet A1 source ranges.',
        }
      }
      const bounds = parseRange(range)
      totalCells += rangeCellCount(bounds)
      if (totalCells > 50_000 || (leftLabels && bounds.startColumn === bounds.endColumn)) {
        return {
          ok: false,
          error: 'invalid_arguments',
          message:
            'xlsx.range.consolidate supports at most 50,000 source cells and left-label sources require at least two columns.',
        }
      }
      const worksheet = xlsxWorksheet(runtime, source.sheet)
      const sourceRange = worksheet.getRange(range)
      if (
        leftLabels &&
        !xlsxRangeDataAvailable(xlsxStreamingState(services), worksheet, sourceRange)
      ) {
        throw new Error(
          'xlsx.range.consolidate requires every left-label source range to finish streaming.',
        )
      }
      sources.push({ sheetId: worksheet.getSheetId(), range })
    }
    if (!services.consolidate) {
      throw new Error('xlsx.range.consolidate is unavailable in this renderer.')
    }
    const result = await services.consolidate({
      targetSheetId: targetWorksheet.getSheetId(),
      targetCell,
      sources,
      aggregation,
      leftLabels,
    })
    if (result.rows * result.columns > 50_000) {
      throw new Error('xlsx.range.consolidate produced more than 50,000 output cells.')
    }
    return {
      ok: true,
      output: {
        targetSheet,
        targetCell,
        sources: sources.length,
        aggregation,
        leftLabels,
        rows: result.rows,
        columns: result.columns,
      },
    }
  },
  'xlsx.range.move': async (arguments_, services) => {
    const sourceSheet = arguments_.sourceSheet as string
    const sourceRange = normalizeXlsxCellRange(arguments_.sourceRange as string)
    const destinationSheet = arguments_.destinationSheet as string
    const destinationRange = normalizeXlsxCellRange(arguments_.destinationRange as string)
    if (!sourceRange || !destinationRange) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'xlsx.range.move requires valid in-sheet A1 source and destination ranges.',
      }
    }
    const sourceBounds = parseRange(sourceRange)
    const destinationBounds = parseRange(destinationRange)
    const sourceRows = sourceBounds.endRow - sourceBounds.startRow + 1
    const sourceColumns = sourceBounds.endColumn - sourceBounds.startColumn + 1
    if (
      sourceRows !== destinationBounds.endRow - destinationBounds.startRow + 1 ||
      sourceColumns !== destinationBounds.endColumn - destinationBounds.startColumn + 1 ||
      sourceRows * sourceColumns > 20_000 ||
      (sourceSheet === destinationSheet && sourceRange === destinationRange)
    ) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message:
          'xlsx.range.move requires distinct equally shaped ranges containing at most 20,000 cells.',
      }
    }
    const runtime = services.runtime()
    const workbook = xlsxWorkbook(runtime)
    const sourceWorksheet = xlsxWorksheet(runtime, sourceSheet)
    const destinationWorksheet = xlsxWorksheet(runtime, destinationSheet)
    if (
      !services.isSheetDataComplete?.(sourceWorksheet.getSheetId()) ||
      !services.isSheetDataComplete(destinationWorksheet.getSheetId())
    ) {
      throw new Error('xlsx.range.move requires fully loaded source and destination worksheets.')
    }
    const state = services.state?.()
    if (
      state?.file?.sheets.some(
        (sheet) =>
          (sheet.id === sourceWorksheet.getSheetId() ||
            sheet.id === destinationWorksheet.getSheetId()) &&
          sheet.pivotRanges.length > 0,
      )
    ) {
      throw new Error('xlsx.range.move cannot mutate a worksheet containing PivotTables.')
    }
    const moved = await runtime!.univerAPI.executeCommand('sheet.command.move-range', {
      fromUnitId: workbook.getId(),
      fromSubUnitId: sourceWorksheet.getSheetId(),
      fromRange: sourceBounds,
      toUnitId: workbook.getId(),
      toSubUnitId: destinationWorksheet.getSheetId(),
      toRange: destinationBounds,
    })
    if (!moved) throw new Error('Univer did not complete xlsx.range.move.')
    return {
      ok: true,
      output: {
        sourceSheet,
        sourceRange,
        destinationSheet,
        destinationRange,
        moved: sourceRows * sourceColumns,
      },
    }
  },
  'xlsx.range.copy_values': (arguments_, services) => {
    const pair = resolveXlsxRangeCopyPair('xlsx.range.copy_values', arguments_, services)
    if (!pair.ok) return pair
    const values = pair.source.getValues().map((row) => row.map((value) => ({ v: value ?? null })))
    pair.destination.setValues(values as ICellData[][])
    pair.destination.activate()
    return {
      ok: true,
      output: {
        sourceSheet: pair.sourceSheet,
        sourceRange: pair.sourceRange,
        destinationSheet: pair.destinationSheet,
        destinationRange: pair.destinationRange,
        copied: pair.cellCount,
      },
    }
  },
  'xlsx.range.copy_formulas': (arguments_, services) => {
    const pair = resolveXlsxRangeCopyPair('xlsx.range.copy_formulas', arguments_, services)
    if (!pair.ok) return pair
    const rowOffset = pair.destination.getRow() - pair.source.getRow()
    const columnOffset = pair.destination.getColumn() - pair.source.getColumn()
    const lexer = pair.runtime.univer.__getInjector().get(LexerTreeBuilder)
    const formulas = pair.source.getFormulas()
    const values = pair.source.getValues()
    const copied = formulas.map((row, rowIndex) =>
      row.map((formula, columnIndex) =>
        formula
          ? { f: lexer.moveFormulaRefOffset(formula, columnOffset, rowOffset) }
          : { v: values[rowIndex]?.[columnIndex] ?? null },
      ),
    )
    pair.destination.setValues(copied as ICellData[][])
    pair.destination.activate()
    return {
      ok: true,
      output: {
        sourceSheet: pair.sourceSheet,
        sourceRange: pair.sourceRange,
        destinationSheet: pair.destinationSheet,
        destinationRange: pair.destinationRange,
        copied: pair.cellCount,
      },
    }
  },
  'xlsx.range.copy_formats': (arguments_, services) => {
    const pair = resolveXlsxRangeCopyPair('xlsx.range.copy_formats', arguments_, services)
    if (!pair.ok) return pair
    if (!applyWorkbookRangeFormatCopy(pair.runtime, pair.source, pair.destination)) {
      return {
        ok: false,
        error: 'execution_failed',
        message: 'xlsx.range.copy_formats could not update the destination formats.',
      }
    }
    pair.destination.activate()
    return {
      ok: true,
      output: {
        sourceSheet: pair.sourceSheet,
        sourceRange: pair.sourceRange,
        destinationSheet: pair.destinationSheet,
        destinationRange: pair.destinationRange,
        copied: pair.cellCount,
      },
    }
  },
  'xlsx.range.copy_without_borders': async (arguments_, services) => {
    const pair = resolveXlsxRangeCopyPair('xlsx.range.copy_without_borders', arguments_, services)
    if (!pair.ok) return pair
    const sourceSheetId = pair.source.getSheetId()
    const state = services.state?.()
    const sourceFormulaText = new Map(state?.formulaText?.get(sourceSheetId) ?? [])
    const sourceHasStructuralEdits =
      (state?.editJournal.structuralOps.get(sourceSheetId)?.length ?? 0) > 0
    if (
      services.readSourceFormulaText &&
      !sourceHasStructuralEdits &&
      !state?.editJournal.sheets.added.has(sourceSheetId)
    ) {
      const indexed = await services.readSourceFormulaText({
        sheetId: sourceSheetId,
        range: {
          startRow: pair.source.getRow(),
          endRow: pair.source.getRow() + pair.source.getHeight() - 1,
          startColumn: pair.source.getColumn(),
          endColumn: pair.source.getColumn() + pair.source.getWidth() - 1,
        },
      })
      for (const [key, formula] of indexed) sourceFormulaText.set(key, formula)
    }
    if (sourceHasStructuralEdits) sourceFormulaText.clear()
    for (const [key, entry] of state?.editJournal.cells.get(sourceSheetId) ?? []) {
      if (!entry.hasValue) continue
      if (entry.formula) sourceFormulaText.set(key, entry.formula)
      else sourceFormulaText.delete(key)
    }
    if (
      !applyWorkbookRangeCopyWithoutBorders(
        pair.runtime,
        pair.source,
        pair.destination,
        sourceFormulaText,
      )
    ) {
      return {
        ok: false,
        error: 'execution_failed',
        message: 'xlsx.range.copy_without_borders could not update the destination cells.',
      }
    }
    pair.destination.activate()
    return {
      ok: true,
      output: {
        sourceSheet: pair.sourceSheet,
        sourceRange: pair.sourceRange,
        destinationSheet: pair.destinationSheet,
        destinationRange: pair.destinationRange,
        copied: pair.cellCount,
      },
    }
  },
  'xlsx.history.undo': async (_arguments, services) => {
    const runtime = services.runtime()
    if (!runtime) throw new Error('Open an XLSX workbook first.')
    if (!(await runtime.univerAPI.undo())) {
      return {
        ok: false,
        error: 'execution_failed',
        message: 'xlsx.history.undo requires an available undo entry.',
      }
    }
    return { ok: true, output: { undone: true } }
  },
  'xlsx.history.redo': async (_arguments, services) => {
    const runtime = services.runtime()
    if (!runtime) throw new Error('Open an XLSX workbook first.')
    if (!(await runtime.univerAPI.redo())) {
      return {
        ok: false,
        error: 'execution_failed',
        message: 'xlsx.history.redo requires an available redo entry.',
      }
    }
    return { ok: true, output: { redone: true } }
  },
  'xlsx.range.set_checkbox': (arguments_, services) => {
    const sheet = arguments_.sheet as string
    const enabled = arguments_.enabled as boolean
    const target = resolveXlsxDataValidationTarget(
      'xlsx.range.set_checkbox',
      sheet,
      arguments_.range as string,
      services,
    )
    if (!target.ok) return target
    applyWorkbookCheckbox(target.runtime, target.range, enabled)
    return { ok: true, output: { sheet, range: target.address, enabled } }
  },
  'xlsx.range.set_list_validation': (arguments_, services) => {
    const sheet = arguments_.sheet as string
    const values = arguments_.values as string[]
    const allowBlank = arguments_.allowBlank as boolean
    const showDropdown = arguments_.showDropdown as boolean
    if (
      values.some(
        (value) =>
          value.length === 0 || value.length > 255 || value.includes(',') || /[\r\n]/.test(value),
      ) ||
      values.join(',').length > 255
    ) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message:
          'xlsx.range.set_list_validation requires a valid range and an inline XLSX list source of at most 255 characters.',
      }
    }
    const target = resolveXlsxDataValidationTarget(
      'xlsx.range.set_list_validation',
      sheet,
      arguments_.range as string,
      services,
    )
    if (!target.ok) return target
    applyWorkbookListValidation(target.runtime, target.range, { values, allowBlank, showDropdown })
    return {
      ok: true,
      output: {
        sheet,
        range: target.address,
        valueCount: values.length,
        allowBlank,
        showDropdown,
      },
    }
  },
  'xlsx.range.set_list_reference_validation': (arguments_, services) => {
    const sheet = arguments_.sheet as string
    const sourceSheet = arguments_.sourceSheet as string
    const sourceAddress = normalizeXlsxCellRange(arguments_.sourceRange as string)
    const allowBlank = arguments_.allowBlank as boolean
    const showDropdown = arguments_.showDropdown as boolean
    if (!sourceAddress) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'xlsx.range.set_list_reference_validation requires a valid source range.',
      }
    }
    const target = resolveXlsxDataValidationTarget(
      'xlsx.range.set_list_reference_validation',
      sheet,
      arguments_.range as string,
      services,
    )
    if (!target.ok) return target
    const sourceWorksheet = xlsxWorksheet(target.runtime, sourceSheet)
    const source = sourceWorksheet.getRange(sourceAddress)
    const sourceCells = source.getHeight() * source.getWidth()
    if ((source.getHeight() > 1 && source.getWidth() > 1) || sourceCells > 1_000) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message:
          'xlsx.range.set_list_reference_validation requires one row or column of at most 1,000 cells.',
      }
    }
    if (!xlsxRangeDataAvailable(xlsxStreamingState(services), sourceWorksheet, source)) {
      return {
        ok: false,
        error: 'execution_failed',
        message:
          'xlsx.range.set_list_reference_validation requires its source range to finish streaming.',
      }
    }
    applyWorkbookListReferenceValidation(target.runtime, target.range, source, {
      allowBlank,
      showDropdown,
    })
    return {
      ok: true,
      output: {
        sheet,
        range: target.address,
        sourceSheet,
        sourceRange: sourceAddress,
        allowBlank,
        showDropdown,
      },
    }
  },
  'xlsx.range.remove_data_validation': (arguments_, services) => {
    const sheet = arguments_.sheet as string
    const target = resolveXlsxDataValidationTarget(
      'xlsx.range.remove_data_validation',
      sheet,
      arguments_.range as string,
      services,
    )
    if (!target.ok) return target
    removeWorkbookDataValidation(target.range)
    return { ok: true, output: { sheet, range: target.address, removed: true } }
  },
  'xlsx.range.set_comparison_validation': (arguments_, services) => {
    const sheet = arguments_.sheet as string
    const kind = arguments_.kind as WorkbookComparisonValidationKind
    const operator = arguments_.operator as WorkbookComparisonValidationOperator
    const operand1 = arguments_.operand1 as string | number
    const operand2 = arguments_.operand2 as string | number | undefined
    const allowBlank = arguments_.allowBlank as boolean
    const binary = operator === 'between' || operator === 'notBetween'
    const first = normalizeXlsxComparisonOperand(kind, operand1)
    const second = operand2 === undefined ? null : normalizeXlsxComparisonOperand(kind, operand2)
    if (
      !first ||
      (binary && !second) ||
      (!binary && operand2 !== undefined) ||
      (binary && second !== null && first.rank > second.rank)
    ) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message:
          'xlsx.range.set_comparison_validation operands must match the kind, operator arity, and ascending bounds.',
      }
    }
    const target = resolveXlsxDataValidationTarget(
      'xlsx.range.set_comparison_validation',
      sheet,
      arguments_.range as string,
      services,
    )
    if (!target.ok) return target
    applyWorkbookComparisonValidation(target.range, {
      kind,
      operator,
      formula1: first.formula,
      ...(second === null ? {} : { formula2: second.formula }),
      allowBlank,
    })
    return {
      ok: true,
      output: {
        sheet,
        range: target.address,
        kind,
        operator,
        operand1,
        ...(operand2 === undefined ? {} : { operand2 }),
        allowBlank,
      },
    }
  },
  'xlsx.range.set_custom_formula_validation': (arguments_, services) => {
    const sheet = arguments_.sheet as string
    const formula = arguments_.formula as string
    const allowBlank = arguments_.allowBlank as boolean
    if (!formula.startsWith('=') || formula.length < 2 || formula.length > 8_192) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message:
          'xlsx.range.set_custom_formula_validation requires a bounded equals-prefixed formula.',
      }
    }
    const target = resolveXlsxDataValidationTarget(
      'xlsx.range.set_custom_formula_validation',
      sheet,
      arguments_.range as string,
      services,
    )
    if (!target.ok) return target
    applyWorkbookCustomFormulaValidation(target.runtime, target.range, formula, allowBlank)
    return {
      ok: true,
      output: { sheet, range: target.address, formula, allowBlank },
    }
  },
  'xlsx.range.set_validation_messages': (arguments_, services) => {
    const sheet = arguments_.sheet as string
    const messages: WorkbookValidationMessages = {
      inputTitle: arguments_.inputTitle as string | null,
      inputMessage: arguments_.inputMessage as string | null,
      errorStyle: arguments_.errorStyle as WorkbookValidationMessages['errorStyle'],
      errorTitle: arguments_.errorTitle as string | null,
      errorMessage: arguments_.errorMessage as string | null,
    }
    const bounded = (value: string | null, maximum: number) =>
      value === null || (value.length > 0 && value.length <= maximum)
    if (
      !bounded(messages.inputTitle, 32) ||
      !bounded(messages.inputMessage, 255) ||
      !bounded(messages.errorTitle, 32) ||
      !bounded(messages.errorMessage, 255) ||
      (messages.errorStyle === 'none' &&
        (messages.errorTitle !== null || messages.errorMessage !== null))
    ) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message:
          'xlsx.range.set_validation_messages requires bounded explicit messages consistent with errorStyle.',
      }
    }
    const target = resolveXlsxDataValidationTarget(
      'xlsx.range.set_validation_messages',
      sheet,
      arguments_.range as string,
      services,
    )
    if (!target.ok) return target
    if (!applyWorkbookValidationMessages(target.range, messages)) {
      return {
        ok: false,
        error: 'execution_failed',
        message: 'xlsx.range.set_validation_messages requires one existing validation rule.',
      }
    }
    return { ok: true, output: { sheet, range: target.address, ...messages } }
  },
  'xlsx.conditional_format.set_comparison': (arguments_, services) => {
    const sheet = arguments_.sheet as string
    const ruleId = arguments_.ruleId as string | null
    const operator = arguments_.operator as WorkbookConditionalFormatComparisonOperator
    const operand1 = arguments_.operand1 as number
    const operand2 = arguments_.operand2 as number | undefined
    const format = arguments_.format as WorkbookConditionalFormatStyle
    const stopIfTrue = arguments_.stopIfTrue as boolean
    const binary = operator === 'between' || operator === 'notBetween'
    if (
      !Number.isFinite(operand1) ||
      (binary && (!Number.isFinite(operand2) || operand1 > (operand2 as number))) ||
      (!binary && operand2 !== undefined) ||
      (ruleId !== null && (ruleId.length === 0 || ruleId.length > 128)) ||
      !isXlsxConditionalFormatStyleValid(format)
    ) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message:
          'xlsx.conditional_format.set_comparison requires bounded operands, an explicit visible format, and valid rule identity.',
      }
    }
    const target = resolveXlsxConditionalFormatTarget(
      'xlsx.conditional_format.set_comparison',
      sheet,
      arguments_.range as string,
      services,
    )
    if (!target.ok) return target
    const result = setWorkbookComparisonConditionalFormat(target.worksheet, target.range, {
      ruleId,
      operator,
      operand1,
      ...(operand2 === undefined ? {} : { operand2 }),
      format,
      stopIfTrue,
    })
    if (!result) {
      return {
        ok: false,
        error: 'execution_failed',
        message:
          'xlsx.conditional_format.set_comparison can update only an existing comparison rule.',
      }
    }
    return {
      ok: true,
      output: {
        sheet,
        range: target.address,
        ...result,
        operator,
        operand1,
        ...(operand2 === undefined ? {} : { operand2 }),
        stopIfTrue,
      },
    }
  },
  'xlsx.conditional_format.remove': (arguments_, services) => {
    const sheet = arguments_.sheet as string
    const ruleId = arguments_.ruleId as string
    if (ruleId.length === 0 || ruleId.length > 128) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'xlsx.conditional_format.remove requires one bounded rule ID.',
      }
    }
    const target = resolveXlsxConditionalFormatSheet(
      'xlsx.conditional_format.remove',
      sheet,
      services,
    )
    if (!target.ok) return target
    if (!removeWorkbookConditionalFormat(target.worksheet, ruleId)) {
      return {
        ok: false,
        error: 'execution_failed',
        message: `Conditional-format rule not found: ${ruleId}`,
      }
    }
    return { ok: true, output: { sheet, ruleId, removed: true } }
  },
  'xlsx.conditional_format.clear': (arguments_, services) => {
    const sheet = arguments_.sheet as string
    const scope = arguments_.scope as 'range' | 'sheet'
    const rangeText = arguments_.range as string | null
    if ((scope === 'range' && rangeText === null) || (scope === 'sheet' && rangeText !== null)) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message:
          'xlsx.conditional_format.clear requires a range only when scope is range.',
      }
    }
    if (scope === 'range') {
      const target = resolveXlsxConditionalFormatTarget(
        'xlsx.conditional_format.clear',
        sheet,
        rangeText as string,
        services,
      )
      if (!target.ok) return target
      const cleared = clearWorkbookConditionalFormats(target.range)
      target.range.activate()
      return {
        ok: true,
        output: { sheet, scope, range: target.address, cleared },
      }
    }
    const target = resolveXlsxConditionalFormatSheet(
      'xlsx.conditional_format.clear',
      sheet,
      services,
    )
    if (!target.ok) return target
    return {
      ok: true,
      output: {
        sheet,
        scope,
        range: null,
        cleared: clearWorkbookConditionalFormats(target.worksheet),
      },
    }
  },
  'xlsx.conditional_format.set_priority': (arguments_, services) => {
    const sheet = arguments_.sheet as string
    const ruleId = arguments_.ruleId as string
    const position = arguments_.position as number
    if (ruleId.length === 0 || ruleId.length > 128) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'xlsx.conditional_format.set_priority requires one bounded rule ID.',
      }
    }
    const target = resolveXlsxConditionalFormatSheet(
      'xlsx.conditional_format.set_priority',
      sheet,
      services,
    )
    if (!target.ok) return target
    const result = setWorkbookConditionalFormatPriority(target.worksheet, ruleId, position)
    if (!result) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: `Conditional-format rule or priority position is unavailable: ${ruleId} at ${position}.`,
      }
    }
    return { ok: true, output: { sheet, ruleId, position, moved: result.moved } }
  },
  'xlsx.conditional_format.set_highlight': (arguments_, services) => {
    const sheet = arguments_.sheet as string
    const ruleId = arguments_.ruleId as string | null
    const predicate = arguments_.predicate as WorkbookConditionalFormatHighlightPredicate
    const text = arguments_.text as string | null
    const format = arguments_.format as WorkbookConditionalFormatStyle
    const stopIfTrue = arguments_.stopIfTrue as boolean
    const requiresText = predicate.startsWith('text')
    if (
      (requiresText && (text === null || text.length === 0 || text.length > 255)) ||
      (!requiresText && text !== null) ||
      (ruleId !== null && (ruleId.length === 0 || ruleId.length > 128)) ||
      !isXlsxConditionalFormatStyleValid(format)
    ) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message:
          'xlsx.conditional_format.set_highlight requires predicate-matched text, an explicit visible format, and valid rule identity.',
      }
    }
    const target = resolveXlsxConditionalFormatTarget(
      'xlsx.conditional_format.set_highlight',
      sheet,
      arguments_.range as string,
      services,
    )
    if (!target.ok) return target
    const result = setWorkbookHighlightConditionalFormat(target.worksheet, target.range, {
      ruleId,
      predicate,
      text,
      format,
      stopIfTrue,
    })
    if (!result) {
      return {
        ok: false,
        error: 'execution_failed',
        message:
          'xlsx.conditional_format.set_highlight can update only an existing text, blank-state, or duplicate-state rule.',
      }
    }
    return {
      ok: true,
      output: {
        sheet,
        range: target.address,
        ...result,
        predicate,
        text,
        stopIfTrue,
      },
    }
  },
  'xlsx.conditional_format.set_statistical': (arguments_, services) => {
    const sheet = arguments_.sheet as string
    const ruleId = arguments_.ruleId as string | null
    const kind = arguments_.kind as WorkbookConditionalFormatStatisticalKind
    const direction = arguments_.direction as WorkbookConditionalFormatStatisticalDirection
    const rank = arguments_.rank as number | null
    const percent = arguments_.percent as boolean | null
    const inclusive = arguments_.inclusive as boolean | null
    const format = arguments_.format as WorkbookConditionalFormatStyle
    const stopIfTrue = arguments_.stopIfTrue as boolean
    const validRank =
      kind === 'rank' &&
      (direction === 'top' || direction === 'bottom') &&
      Number.isInteger(rank) &&
      (rank as number) >= 1 &&
      (rank as number) <= (percent === true ? 100 : 1_000) &&
      typeof percent === 'boolean' &&
      inclusive === null
    const validAverage =
      kind === 'average' &&
      (direction === 'above' || direction === 'below') &&
      rank === null &&
      percent === null &&
      typeof inclusive === 'boolean'
    if (
      (!validRank && !validAverage) ||
      (ruleId !== null && (ruleId.length === 0 || ruleId.length > 128)) ||
      !isXlsxConditionalFormatStyleValid(format)
    ) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message:
          'xlsx.conditional_format.set_statistical requires one coherent bounded rank or average final state.',
      }
    }
    const target = resolveXlsxConditionalFormatTarget(
      'xlsx.conditional_format.set_statistical',
      sheet,
      arguments_.range as string,
      services,
    )
    if (!target.ok) return target
    const result = setWorkbookStatisticalConditionalFormat(target.worksheet, target.range, {
      ruleId,
      kind,
      direction,
      rank,
      percent,
      inclusive,
      format,
      stopIfTrue,
    })
    if (!result) {
      return {
        ok: false,
        error: 'execution_failed',
        message:
          'xlsx.conditional_format.set_statistical can update only an existing rank or average rule.',
      }
    }
    return {
      ok: true,
      output: {
        sheet,
        range: target.address,
        ...result,
        kind,
        direction,
        rank,
        percent,
        inclusive,
        stopIfTrue,
      },
    }
  },
  'xlsx.conditional_format.set_formula': (arguments_, services) => {
    const sheet = arguments_.sheet as string
    const ruleId = arguments_.ruleId as string | null
    const formula = arguments_.formula as string
    const format = arguments_.format as WorkbookConditionalFormatStyle
    const stopIfTrue = arguments_.stopIfTrue as boolean
    if (
      formula.length < 2 ||
      formula.length > 8_192 ||
      !formula.startsWith('=') ||
      (ruleId !== null && (ruleId.length === 0 || ruleId.length > 128)) ||
      !isXlsxConditionalFormatStyleValid(format)
    ) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message:
          'xlsx.conditional_format.set_formula requires a bounded equals-prefixed formula, visible format, and valid rule identity.',
      }
    }
    const target = resolveXlsxConditionalFormatTarget(
      'xlsx.conditional_format.set_formula',
      sheet,
      arguments_.range as string,
      services,
    )
    if (!target.ok) return target
    const result = setWorkbookFormulaConditionalFormat(target.worksheet, target.range, {
      ruleId,
      formula,
      format,
      stopIfTrue,
    })
    if (!result) {
      return {
        ok: false,
        error: 'execution_failed',
        message:
          'xlsx.conditional_format.set_formula can update only an existing custom-formula rule.',
      }
    }
    return {
      ok: true,
      output: { sheet, range: target.address, ...result, formula, stopIfTrue },
    }
  },
  'xlsx.conditional_format.set_visual': (arguments_, services) => {
    const sheet = arguments_.sheet as string
    const input: WorkbookConditionalFormatVisualInput = {
      ruleId: arguments_.ruleId as string | null,
      kind: arguments_.kind as WorkbookConditionalFormatVisualKind,
      colors: arguments_.colors as string[],
      thresholds: arguments_.thresholds as WorkbookConditionalFormatThreshold[],
      iconSet: arguments_.iconSet as string | null,
      showValue: arguments_.showValue as boolean | null,
      reverse: arguments_.reverse as boolean | null,
      gradient: arguments_.gradient as boolean | null,
      stopIfTrue: arguments_.stopIfTrue as boolean,
    }
    if (
      (input.ruleId !== null && (input.ruleId.length === 0 || input.ruleId.length > 128)) ||
      !isWorkbookVisualConditionalFormatValid(input)
    ) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message:
          'xlsx.conditional_format.set_visual requires one coherent base-OOXML color-scale, data-bar, or icon-set final state.',
      }
    }
    const target = resolveXlsxConditionalFormatTarget(
      'xlsx.conditional_format.set_visual',
      sheet,
      arguments_.range as string,
      services,
    )
    if (!target.ok) return target
    const result = setWorkbookVisualConditionalFormat(target.worksheet, target.range, input)
    if (!result) {
      return {
        ok: false,
        error: 'execution_failed',
        message:
          'xlsx.conditional_format.set_visual can update only an existing visual conditional-format rule.',
      }
    }
    return {
      ok: true,
      output: {
        sheet,
        range: target.address,
        ...result,
        kind: input.kind,
        stopIfTrue: input.stopIfTrue,
      },
    }
  },
  'xlsx.range.set_filter': async (arguments_, services) => {
    const sheet = arguments_.sheet as string
    const address = arguments_.range as string
    const enabled = arguments_.enabled as boolean
    const runtime = services.runtime()
    if (!runtime) throw new Error('Open an XLSX workbook first.')
    const worksheet = xlsxWorksheet(runtime, sheet)
    const targetRange = worksheet.getRange(address)
    if (targetRange.getHeight() < 2) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'xlsx.range.set_filter requires at least two rows.',
      }
    }
    const result = await applyWorkbookFilter(runtime, worksheet, targetRange.getRange(), enabled)
    if (result === 'range_conflict') {
      return {
        ok: false,
        error: 'execution_failed',
        message: 'xlsx.range.set_filter target does not match the active worksheet filter.',
      }
    }
    if (result === 'failed') {
      return {
        ok: false,
        error: 'execution_failed',
        message: 'Univer did not complete xlsx.range.set_filter.',
      }
    }
    return {
      ok: true,
      output: { sheet, range: address.toUpperCase(), enabled },
    }
  },
  'xlsx.range.clear_filter_criteria': async (arguments_, services) => {
    const sheet = arguments_.sheet as string
    const address = arguments_.range as string
    const runtime = services.runtime()
    if (!runtime) throw new Error('Open an XLSX workbook first.')
    const worksheet = xlsxWorksheet(runtime, sheet)
    const targetRange = worksheet.getRange(address)
    const result = await clearWorkbookFilterCriteria(runtime, worksheet, targetRange.getRange())
    if (result === 'missing') {
      return {
        ok: false,
        error: 'execution_failed',
        message: 'xlsx.range.clear_filter_criteria requires an active worksheet filter.',
      }
    }
    if (result === 'range_conflict') {
      return {
        ok: false,
        error: 'execution_failed',
        message:
          'xlsx.range.clear_filter_criteria target does not match the active worksheet filter.',
      }
    }
    if (result === 'failed') {
      return {
        ok: false,
        error: 'execution_failed',
        message: 'Univer did not complete xlsx.range.clear_filter_criteria.',
      }
    }
    return {
      ok: true,
      output: { sheet, range: address.toUpperCase(), cleared: true },
    }
  },
  'xlsx.range.set_filter_values': async (arguments_, services) => {
    const sheet = arguments_.sheet as string
    const address = arguments_.range as string
    const column = arguments_.column as string
    const values = arguments_.values as string[]
    const includeBlank = arguments_.includeBlank as boolean
    const location = xlsxColumnLocation(column)
    if (!location) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'xlsx.range.set_filter_values requires an A1 column label.',
      }
    }
    const runtime = services.runtime()
    if (!runtime) throw new Error('Open an XLSX workbook first.')
    const worksheet = xlsxWorksheet(runtime, sheet)
    const result = await setWorkbookFilterValues(
      runtime,
      worksheet,
      parseRange(address.toUpperCase()),
      location.index,
      values,
      includeBlank,
    )
    if (result === 'missing') {
      return {
        ok: false,
        error: 'execution_failed',
        message: 'xlsx.range.set_filter_values requires an active worksheet filter.',
      }
    }
    if (result === 'range_conflict') {
      return {
        ok: false,
        error: 'execution_failed',
        message: 'xlsx.range.set_filter_values target does not match the active worksheet filter.',
      }
    }
    if (result === 'column_outside') {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'xlsx.range.set_filter_values column must be inside the filter range.',
      }
    }
    if (result === 'failed') {
      return {
        ok: false,
        error: 'execution_failed',
        message: 'Univer did not complete xlsx.range.set_filter_values.',
      }
    }
    return {
      ok: true,
      output: {
        sheet,
        range: address.toUpperCase(),
        column: location.normalized,
        selectedValues: values.length,
        includeBlank,
      },
    }
  },
  'xlsx.range.set_custom_filter': async (arguments_, services) => {
    const sheet = arguments_.sheet as string
    const address = arguments_.range as string
    const column = arguments_.column as string
    const conjunction = arguments_.conjunction as 'and' | 'or'
    const conditions = arguments_.conditions as WorkbookCustomFilterCondition[]
    const location = xlsxColumnLocation(column)
    if (!location) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'xlsx.range.set_custom_filter requires an A1 column label.',
      }
    }
    const runtime = services.runtime()
    if (!runtime) throw new Error('Open an XLSX workbook first.')
    const worksheet = xlsxWorksheet(runtime, sheet)
    const result = await setWorkbookCustomFilter(
      runtime,
      worksheet,
      parseRange(address.toUpperCase()),
      location.index,
      conjunction,
      conditions,
    )
    if (result === 'missing') {
      return {
        ok: false,
        error: 'execution_failed',
        message: 'xlsx.range.set_custom_filter requires an active worksheet filter.',
      }
    }
    if (result === 'range_conflict') {
      return {
        ok: false,
        error: 'execution_failed',
        message: 'xlsx.range.set_custom_filter target does not match the active worksheet filter.',
      }
    }
    if (result === 'column_outside') {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'xlsx.range.set_custom_filter column must be inside the filter range.',
      }
    }
    if (result === 'failed') {
      return {
        ok: false,
        error: 'execution_failed',
        message: 'Univer did not complete xlsx.range.set_custom_filter.',
      }
    }
    return {
      ok: true,
      output: {
        sheet,
        range: address.toUpperCase(),
        column: location.normalized,
        conjunction,
        conditions: conditions.length,
      },
    }
  },
  'xlsx.range.set_alignment': (arguments_, services) => {
    const sheet = arguments_.sheet as string
    const address = arguments_.range as string
    const alignment = arguments_.alignment as XlsxAlignment
    const fields = arguments_.fields as string[]
    const declaredFields = new Set(fields)
    const alignmentFields = Object.keys(alignment)
    if (
      declaredFields.size !== fields.length ||
      alignmentFields.length !== declaredFields.size ||
      alignmentFields.some((field) => !declaredFields.has(field)) ||
      fields.some(
        (field) => !isXlsxAlignmentField(field) || !hasValidXlsxAlignmentValue(alignment, field),
      )
    ) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'xlsx.range.set_alignment requires unique fields with matching explicit values.',
      }
    }
    const targetRange = xlsxWorksheet(services.runtime(), sheet).getRange(address)
    targetRange.setValue({
      s: xlsxAlignmentPatch(alignment, fields as XlsxAlignmentField[]),
    } as unknown as ICellData)
    targetRange.activate()
    return {
      ok: true,
      output: { sheet, range: address.toUpperCase(), fields },
    }
  },
  'xlsx.range.set_font': (arguments_, services) => {
    const sheet = arguments_.sheet as string
    const address = arguments_.range as string
    const font = arguments_.font as XlsxFont
    const fields = arguments_.fields as string[]
    const declaredFields = new Set(fields)
    const fontFields = Object.keys(font)
    if (
      declaredFields.size !== fields.length ||
      fontFields.length !== declaredFields.size ||
      fontFields.some((field) => !declaredFields.has(field)) ||
      fields.some((field) => !isXlsxFontField(field) || !hasValidXlsxFontValue(font, field))
    ) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'xlsx.range.set_font requires unique fields with matching explicit values.',
      }
    }
    const targetRange = xlsxWorksheet(services.runtime(), sheet).getRange(address)
    targetRange.setValue({
      s: xlsxFontPatch(font, fields as XlsxFontField[]),
    } as unknown as ICellData)
    targetRange.activate()
    return {
      ok: true,
      output: { sheet, range: address.toUpperCase(), fields },
    }
  },
  'xlsx.range.set_fill': (arguments_, services) => {
    const sheet = arguments_.sheet as string
    const address = arguments_.range as string
    const color = arguments_.color as string | null
    if (color !== null && !isXlsxColor(color)) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'xlsx.range.set_fill requires a #RRGGBB color or null.',
      }
    }
    const targetRange = xlsxWorksheet(services.runtime(), sheet).getRange(address)
    targetRange.setValue({
      s: { bg: color === null ? null : { rgb: color } },
    } as unknown as ICellData)
    targetRange.activate()
    return { ok: true, output: { sheet, range: address.toUpperCase() } }
  },
  'xlsx.range.set_border': (arguments_, services) => {
    const sheet = arguments_.sheet as string
    const address = arguments_.range as string
    const border = arguments_.border as {
      preset: string
      lineStyle?: string
      color?: string
    }
    const borderType = BORDER_COMMAND_TYPES[border.preset]
    const clearsBorder = border.preset === 'none'
    if (
      !borderType ||
      (clearsBorder
        ? border.lineStyle !== undefined || border.color !== undefined
        : !xlsxBorderLineStyles[border.lineStyle ?? ''] || !isXlsxColor(border.color))
    ) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message:
          'xlsx.range.set_border requires either preset none alone or an explicit line style and #RRGGBB color.',
      }
    }
    const targetRange = xlsxWorksheet(services.runtime(), sheet).getRange(address)
    targetRange.setBorder(
      borderType,
      clearsBorder ? BorderStyleTypes.NONE : xlsxBorderLineStyles[border.lineStyle as string]!,
      clearsBorder ? '#000000' : (border.color as string),
    )
    targetRange.activate()
    return {
      ok: true,
      output: { sheet, range: address.toUpperCase(), preset: border.preset },
    }
  },
  'xlsx.range.apply_cell_style': (arguments_, services) => {
    const sheet = arguments_.sheet as string
    const address = arguments_.range as string
    const preset = arguments_.preset as string
    const style = xlsxCellStylePatch(preset)
    if (!style) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'xlsx.range.apply_cell_style requires a known named preset.',
      }
    }
    const targetRange = xlsxWorksheet(services.runtime(), sheet).getRange(address)
    targetRange.setValue({ s: style } as unknown as ICellData)
    targetRange.activate()
    return {
      ok: true,
      output: { sheet, range: address.toUpperCase(), preset },
    }
  },
  'xlsx.range.set_number_format': (arguments_, services) => {
    const sheet = arguments_.sheet as string
    const address = arguments_.range as string
    const pattern = arguments_.pattern as string
    if (pattern.length === 0 || pattern.length > 255) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'xlsx.range.set_number_format requires a 1–255 character pattern.',
      }
    }
    const targetRange = xlsxWorksheet(services.runtime(), sheet).getRange(address)
    targetRange.setNumberFormat(pattern)
    targetRange.activate()
    return {
      ok: true,
      output: { sheet, range: address.toUpperCase(), pattern },
    }
  },
  'xlsx.range.merge': (arguments_, services) => {
    const sheet = arguments_.sheet as string
    const address = arguments_.range as string
    const mode = arguments_.mode as 'across' | 'cells' | 'center' | 'unmerge'
    if (mode !== 'across' && mode !== 'cells' && mode !== 'center' && mode !== 'unmerge') {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'xlsx.range.merge does not support this mode yet.',
      }
    }
    const targetRange = xlsxWorksheet(services.runtime(), sheet).getRange(address)
    if (mode === 'across') targetRange.mergeAcross()
    else if (mode === 'unmerge') targetRange.breakApart()
    else {
      targetRange.merge()
      if (mode === 'center') targetRange.setHorizontalAlignment('center')
    }
    targetRange.activate()
    return {
      ok: true,
      output: { sheet, range: address.toUpperCase(), mode },
    }
  },
  'xlsx.range.clear': (arguments_, services) => {
    const sheet = arguments_.sheet as string
    const address = arguments_.range as string
    const scope = arguments_.scope as 'all' | 'contents' | 'formats'
    const targetRange = xlsxWorksheet(services.runtime(), sheet).getRange(address)
    if (scope === 'contents') targetRange.clearContent()
    else if (scope === 'formats') targetRange.clearFormat()
    else if (scope === 'all') targetRange.clear()
    else {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'xlsx.range.clear does not support this scope yet.',
      }
    }
    targetRange.activate()
    return {
      ok: true,
      output: { sheet, range: address.toUpperCase(), scope },
    }
  },
  'xlsx.range.fill': async (arguments_, services) => {
    const sheet = arguments_.sheet as string
    const address = arguments_.range as string
    const direction = arguments_.direction as 'down' | 'right'
    const runtime = services.runtime()
    if (!runtime) throw new Error('Open an XLSX workbook first.')
    const targetRange = xlsxWorksheet(runtime, sheet).getRange(address)
    if (direction !== 'down' && direction !== 'right') {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'xlsx.range.fill does not support this direction yet.',
      }
    }
    const extent = direction === 'down' ? targetRange.getHeight() : targetRange.getWidth()
    if (extent < 2) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: `xlsx.range.fill ${direction} requires at least two ${direction === 'down' ? 'rows' : 'columns'}.`,
      }
    }
    targetRange.activate()
    const completed = await runtime.univerAPI.executeCommand(
      direction === 'down' ? 'sheet.command.copy-down' : 'sheet.command.copy-right',
    )
    if (!completed) {
      return {
        ok: false,
        error: 'execution_failed',
        message: 'Univer did not complete xlsx.range.fill.',
      }
    }
    return {
      ok: true,
      output: { sheet, range: address.toUpperCase(), direction },
    }
  },
  'xlsx.range.sort': async (arguments_, services) => {
    const sheet = arguments_.sheet as string
    const address = arguments_.range as string
    const direction = arguments_.direction as 'asc' | 'desc'
    const runtime = services.runtime()
    if (!runtime) throw new Error('Open an XLSX workbook first.')
    const workbook = xlsxWorkbook(runtime)
    const worksheet = xlsxWorksheet(runtime, sheet)
    const targetRange = worksheet.getRange(address)
    if (targetRange.getHeight() < 2) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'xlsx.range.sort requires at least two rows.',
      }
    }
    const startRow = targetRange.getRow()
    const startColumn = targetRange.getColumn()
    const completed = await runtime.univerAPI.executeCommand('sheet.command.sort-range', {
      unitId: workbook.getId(),
      subUnitId: worksheet.getSheetId(),
      range: {
        startRow,
        endRow: startRow + targetRange.getHeight() - 1,
        startColumn,
        endColumn: startColumn + targetRange.getWidth() - 1,
      },
      orderRules: [{ type: direction, colIndex: startColumn }],
      hasTitle: false,
    })
    if (!completed) {
      return {
        ok: false,
        error: 'execution_failed',
        message: 'Univer did not complete xlsx.range.sort.',
      }
    }
    targetRange.activate()
    return {
      ok: true,
      output: { sheet, range: address.toUpperCase(), direction },
    }
  },
  'xlsx.range.sort_custom': async (arguments_, services) => {
    const sheet = arguments_.sheet as string
    const address = arguments_.range as string
    const keys = arguments_.keys as Array<{ column: string; direction: 'asc' | 'desc' }>
    const hasHeader = arguments_.hasHeader as boolean
    const runtime = services.runtime()
    if (!runtime) throw new Error('Open an XLSX workbook first.')
    const workbook = xlsxWorkbook(runtime)
    const worksheet = xlsxWorksheet(runtime, sheet)
    const targetRange = worksheet.getRange(address)
    if (targetRange.getHeight() < 2) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'xlsx.range.sort_custom requires at least two rows.',
      }
    }
    const startRow = targetRange.getRow()
    const startColumn = targetRange.getColumn()
    const endColumn = startColumn + targetRange.getWidth() - 1
    const normalizedKeys: Array<{ column: string; direction: 'asc' | 'desc'; index: number }> = []
    const keyColumns = new Set<number>()
    for (const key of keys) {
      const location = xlsxColumnLocation(key.column)
      if (!location || location.index < startColumn || location.index > endColumn) {
        return {
          ok: false,
          error: 'invalid_arguments',
          message: 'xlsx.range.sort_custom keys must be A1 columns inside the target range.',
        }
      }
      if (keyColumns.has(location.index)) {
        return {
          ok: false,
          error: 'invalid_arguments',
          message: 'xlsx.range.sort_custom requires unique key columns.',
        }
      }
      keyColumns.add(location.index)
      normalizedKeys.push({
        column: location.normalized,
        direction: key.direction,
        index: location.index,
      })
    }
    const completed = await runtime.univerAPI.executeCommand('sheet.command.sort-range', {
      unitId: workbook.getId(),
      subUnitId: worksheet.getSheetId(),
      range: {
        startRow,
        endRow: startRow + targetRange.getHeight() - 1,
        startColumn,
        endColumn,
      },
      orderRules: normalizedKeys.map((key) => ({ type: key.direction, colIndex: key.index })),
      hasTitle: hasHeader,
    })
    if (!completed) {
      return {
        ok: false,
        error: 'execution_failed',
        message: 'Univer did not complete xlsx.range.sort_custom.',
      }
    }
    targetRange.activate()
    return {
      ok: true,
      output: {
        sheet,
        range: address.toUpperCase(),
        keys: normalizedKeys.map(({ column, direction }) => ({ column, direction })),
        hasHeader,
      },
    }
  },
  'xlsx.range.remove_duplicates': (arguments_, services) => {
    const sheet = arguments_.sheet as string
    const address = arguments_.range as string
    const hasHeader = arguments_.hasHeader as boolean
    const worksheet = xlsxWorksheet(services.runtime(), sheet)
    if (services.isSheetDataComplete && !services.isSheetDataComplete(worksheet.getSheetId())) {
      return {
        ok: false,
        error: 'execution_failed',
        message: 'xlsx.range.remove_duplicates requires fully loaded sheet data.',
      }
    }
    const targetRange = worksheet.getRange(address)
    if (targetRange.getHeight() < 2) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'xlsx.range.remove_duplicates requires at least two rows.',
      }
    }
    const values = targetRange.getValues().map((row) => row.map((value) => value ?? null))
    const { rows, removed } = dedupeRows(values, hasHeader)
    while (rows.length < values.length) {
      rows.push(Array.from({ length: targetRange.getWidth() }, () => null))
    }
    const startRow = targetRange.getRow()
    const startColumn = targetRange.getColumn()
    for (const [offset, row] of rows.entries()) {
      const current = values[offset]
      if (current && row.every((value, index) => value === current[index])) continue
      worksheet
        .getRange(startRow + offset, startColumn, 1, targetRange.getWidth())
        .setValues([[...row]] as unknown as ICellData[][])
    }
    targetRange.activate()
    return {
      ok: true,
      output: { sheet, range: address.toUpperCase(), removed },
    }
  },
  'xlsx.range.set_protection': (arguments_, services) => {
    const sheet = arguments_.sheet as string
    const address = arguments_.range as string
    const protection = arguments_.protection as XlsxProtection
    const fields = arguments_.fields as string[]
    const declaredFields = new Set(fields)
    const protectionFields = Object.keys(protection)
    if (
      declaredFields.size !== fields.length ||
      protectionFields.length !== declaredFields.size ||
      protectionFields.some((field) => !declaredFields.has(field)) ||
      fields.some(
        (field) => !isXlsxProtectionField(field) || typeof protection[field] !== 'boolean',
      )
    ) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'xlsx.range.set_protection requires unique fields with matching explicit values.',
      }
    }
    const state = xlsxOperationState(services)
    const worksheet = xlsxWorksheet(services.runtime(), sheet)
    const targetRange = worksheet.getRange(address)
    if (targetRange.getHeight() * targetRange.getWidth() > 10_000) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'xlsx.range.set_protection supports at most 10,000 cells per operation.',
      }
    }
    applyWorkbookCellProtection(state, worksheet, targetRange, {
      ...(declaredFields.has('locked') ? { locked: protection.locked } : {}),
      ...(declaredFields.has('hidden') ? { hidden: protection.hidden } : {}),
    })
    targetRange.activate()
    services.setPendingEdits?.(journalSize(state.editJournal))
    return {
      ok: true,
      output: { sheet, range: address.toUpperCase(), fields },
    }
  },
  'xlsx.hyperlink.set': (arguments_, services) => {
    const sheet = arguments_.sheet as string
    const address = arguments_.address as string
    const normalizedAddress = address.toUpperCase()
    const target = normalizeLinkTarget(arguments_.target as string)
    if (target === null) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'xlsx.hyperlink.set target must be a URL or an internal reference like Sheet1!A1.',
      }
    }
    const state = xlsxOperationState(services)
    const worksheet = xlsxWorksheet(services.runtime(), sheet)
    applyWorkbookHyperlink(state, worksheet, { address: normalizedAddress, target })
    worksheet.getRange(address).activate()
    services.setPendingEdits?.(journalSize(state.editJournal))
    return {
      ok: true,
      output: { sheet, address: normalizedAddress, target },
    }
  },
  'xlsx.hyperlink.remove': (arguments_, services) => {
    const sheet = arguments_.sheet as string
    const address = arguments_.address as string
    const normalizedAddress = address.toUpperCase()
    const state = xlsxOperationState(services)
    const worksheet = xlsxWorksheet(services.runtime(), sheet)
    applyWorkbookHyperlink(state, worksheet, { address: normalizedAddress, target: null })
    worksheet.getRange(address).activate()
    services.setPendingEdits?.(journalSize(state.editJournal))
    return {
      ok: true,
      output: { sheet, address: normalizedAddress, removed: true },
    }
  },
  'xlsx.table.add': (arguments_, services) => {
    const sheet = arguments_.sheet as string
    const address = arguments_.range as string
    const normalizedAddress = address.toUpperCase()
    const style = arguments_.style as string
    const runtime = services.runtime()
    if (!runtime) throw new Error('Open an XLSX workbook first.')
    const state = xlsxOperationState(services)
    const worksheet = xlsxWorksheet(runtime, sheet)
    const tableIndex = state.editJournal.tableAdds.length
    applyWorkbookTableAdd(runtime, state, {
      op: 'add_table',
      sheetId: worksheet.getSheetId(),
      range: normalizedAddress,
      style,
      bandedRows: true,
    })
    const table = state.editJournal.tableAdds[tableIndex]
    if (!table) throw new Error('The table was not recorded for save.')
    worksheet.getRange(address).activate()
    services.setPendingEdits?.(journalSize(state.editJournal))
    return {
      ok: true,
      output: { sheet, range: normalizedAddress, name: table.name, style },
    }
  },
  'xlsx.table.insert_rows': async (arguments_, services) => {
    const sheet = arguments_.sheet as string
    const tableName = arguments_.table as string
    const row = arguments_.row as number
    const count = arguments_.count as number
    const { state, worksheet, table } = xlsxSessionTable(services, sheet, tableName)
    const dataRows = table.area.endRow - table.area.startRow
    if (row > dataRows + 1) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: `xlsx.table.insert_rows row must be within 1..${dataRows + 1}.`,
      }
    }
    if (!services.mutateTable) throw new Error('XLSX table mutation is unavailable.')
    await services.mutateTable({
      op: 'add_table_row',
      sheetId: worksheet.getSheetId(),
      tableName,
      row,
      count,
    })
    services.setPendingEdits?.(journalSize(state.editJournal))
    return { ok: true, output: { sheet, table: tableName, row, count } }
  },
  'xlsx.table.delete_rows': async (arguments_, services) => {
    const sheet = arguments_.sheet as string
    const tableName = arguments_.table as string
    const row = arguments_.row as number
    const count = arguments_.count as number
    const { state, worksheet, table } = xlsxSessionTable(services, sheet, tableName)
    const dataRows = table.area.endRow - table.area.startRow
    if (row + count - 1 > dataRows || dataRows - count < 1) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'xlsx.table.delete_rows must stay inside the data area and retain one data row.',
      }
    }
    if (!services.mutateTable) throw new Error('XLSX table mutation is unavailable.')
    await services.mutateTable({
      op: 'delete_table_row',
      sheetId: worksheet.getSheetId(),
      tableName,
      row,
      count,
    })
    services.setPendingEdits?.(journalSize(state.editJournal))
    return { ok: true, output: { sheet, table: tableName, row, count } }
  },
  'xlsx.table.insert_columns': async (arguments_, services) => {
    const sheet = arguments_.sheet as string
    const tableName = arguments_.table as string
    const column = arguments_.column as number
    const columnName = arguments_.columnName as string
    const count = arguments_.count as number
    const { state, worksheet, table } = xlsxSessionTable(services, sheet, tableName)
    const tableColumns = table.area.endColumn - table.area.startColumn + 1
    if (column > tableColumns + 1) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: `xlsx.table.insert_columns column must be within 1..${tableColumns + 1}.`,
      }
    }
    if (columnName.trim().length === 0 || columnName.length > 255) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'xlsx.table.insert_columns requires a 1..255 character columnName.',
      }
    }
    const generatedNames = Array.from({ length: count }, (_, index) =>
      index === 0 ? columnName : `${columnName}${index + 1}`,
    )
    const existing = new Set(table.columnNames.map((name) => name.toLowerCase()))
    if (
      generatedNames.some((name) => name.length > 255 || existing.has(name.toLowerCase())) ||
      new Set(generatedNames.map((name) => name.toLowerCase())).size !== generatedNames.length
    ) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'xlsx.table.insert_columns generated names must be bounded and unique.',
      }
    }
    if (!services.mutateTable) throw new Error('XLSX table mutation is unavailable.')
    await services.mutateTable({
      op: 'add_table_column',
      sheetId: worksheet.getSheetId(),
      tableName,
      column,
      columnName,
      count,
    })
    services.setPendingEdits?.(journalSize(state.editJournal))
    return {
      ok: true,
      output: { sheet, table: tableName, column, columnName, count },
    }
  },
  'xlsx.table.delete_columns': async (arguments_, services) => {
    const sheet = arguments_.sheet as string
    const tableName = arguments_.table as string
    const column = arguments_.column as number
    const count = arguments_.count as number
    const { state, worksheet, table } = xlsxSessionTable(services, sheet, tableName)
    const tableColumns = table.area.endColumn - table.area.startColumn + 1
    if (column + count - 1 > tableColumns || tableColumns - count < 1) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'xlsx.table.delete_columns must stay inside the table and retain one column.',
      }
    }
    if (!services.mutateTable) throw new Error('XLSX table mutation is unavailable.')
    await services.mutateTable({
      op: 'delete_table_column',
      sheetId: worksheet.getSheetId(),
      tableName,
      column,
      count,
    })
    services.setPendingEdits?.(journalSize(state.editJournal))
    return { ok: true, output: { sheet, table: tableName, column, count } }
  },
  'xlsx.table.convert_to_range': async (arguments_, services) => {
    const sheet = arguments_.sheet as string
    const tableName = arguments_.table as string
    const { state, worksheet } = xlsxSessionTable(services, sheet, tableName)
    if (!services.mutateTable) throw new Error('XLSX table mutation is unavailable.')
    await services.mutateTable({
      op: 'convert_table_to_range',
      sheetId: worksheet.getSheetId(),
      tableName,
    })
    services.setPendingEdits?.(journalSize(state.editJournal))
    return { ok: true, output: { sheet, table: tableName, converted: true } }
  },
  'xlsx.sparkline.add': (arguments_, services) => {
    const sheet = arguments_.sheet as string
    const sourceAddress = normalizeXlsxCellRange(arguments_.sourceRange as string)
    const targetAddress = normalizeXlsxCellRange(arguments_.targetRange as string)
    const type = arguments_.type as WorkbookSparklineType
    if (!sourceAddress || !targetAddress) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'xlsx.sparkline.add requires valid in-sheet A1 source and target ranges.',
      }
    }
    const runtime = services.runtime()
    if (!runtime) throw new Error('Open an XLSX workbook first.')
    const state = xlsxOperationState(services)
    if (!state.file) throw new Error('Open a file-backed XLSX workbook first.')
    const worksheet = xlsxWorksheet(runtime, sheet)
    const source = worksheet.getRange(sourceAddress)
    const target = worksheet.getRange(targetAddress)
    const streamingState = xlsxStreamingState(services)
    if (!xlsxRangeDataAvailable(streamingState, worksheet, source)) {
      return {
        ok: false,
        error: 'execution_failed',
        message: 'xlsx.sparkline.add requires its source range to finish streaming.',
      }
    }
    if (!xlsxRangeDataAvailable(streamingState, worksheet, target)) {
      return {
        ok: false,
        error: 'execution_failed',
        message: 'xlsx.sparkline.add requires its target range to finish streaming.',
      }
    }
    const result = applyWorkbookSparklineAdd(
      runtime,
      { editJournal: state.editJournal, file: state.file },
      worksheet,
      source,
      target,
      type,
      (sheetId) => {
        services.setPendingEdits?.(journalSize(state.editJournal))
        services.refreshSparklines?.(sheetId)
      },
    )
    if (!result.ok) {
      const message = {
        source_width: 'xlsx.sparkline.add requires at least two source columns.',
        shape: 'xlsx.sparkline.add requires a one-column target with the same row count.',
        limit: 'xlsx.sparkline.add supports at most 200 target cells.',
        overlap: 'xlsx.sparkline.add source and target ranges must not overlap.',
        occupied: 'xlsx.sparkline.add cannot replace an existing sparkline.',
      }[result.reason]
      return { ok: false, error: 'invalid_arguments', message }
    }
    return {
      ok: true,
      output: {
        sheet,
        sourceRange: sourceAddress,
        targetRange: targetAddress,
        type,
        count: result.count,
      },
    }
  },
  'xlsx.range.set_text_style': (arguments_, services) => {
    const sheet = arguments_.sheet as string
    const address = arguments_.range as string
    const style = arguments_.style as XlsxTextStyle
    const fields = arguments_.fields as string[]
    const declaredFields = new Set(fields)
    const styleFields = Object.keys(style)
    if (
      declaredFields.size !== fields.length ||
      styleFields.length !== declaredFields.size ||
      styleFields.some((field) => !declaredFields.has(field)) ||
      fields.some(
        (field) => !isXlsxTextStyleField(field) || !hasValidXlsxTextStyleValue(style, field),
      )
    ) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'xlsx.range.set_text_style requires unique fields with matching explicit values.',
      }
    }
    const targetRange = xlsxWorksheet(services.runtime(), sheet).getRange(address)
    targetRange.setValue({
      s: xlsxTextStylePatch(style, fields as XlsxTextStyleField[]),
    } as unknown as ICellData)
    targetRange.activate()
    return {
      ok: true,
      output: { sheet, range: address.toUpperCase(), fields },
    }
  },
  'xlsx.outline.set_level': (arguments_, services) => {
    const sheet = arguments_.sheet as string
    const axis = arguments_.axis as WorkbookOutlineAxis
    const start = arguments_.start as number
    const count = arguments_.count as number
    const level = arguments_.level as number
    const runtime = services.runtime()
    if (!runtime) throw new Error('Open an XLSX workbook first.')
    const state = xlsxOperationState(services)
    if (!state.outline) throw new Error('Open a file-backed XLSX workbook first.')
    const worksheet = xlsxWorksheet(runtime, sheet)
    const maximum = axis === 'rows' ? worksheet.getMaxRows() : worksheet.getMaxColumns()
    if (start + count - 1 > maximum) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'xlsx.outline.set_level must stay inside the worksheet axis bounds.',
      }
    }
    if (
      axis === 'rows' &&
      !streamedWorkbookRectangleAvailable(
        xlsxStreamingState(services),
        worksheet.getSheetId(),
        start - 1,
        start + count - 2,
        0,
        0,
      )
    ) {
      return {
        ok: false,
        error: 'execution_failed',
        message: 'xlsx.outline.set_level requires its row metadata to finish streaming.',
      }
    }
    applyWorkbookOutlineLevel(
      runtime,
      state as Pick<LazyWorkbookState, 'editJournal' | 'outline'>,
      worksheet,
      axis,
      start - 1,
      count,
      level,
      services.setPendingEdits,
    )
    return { ok: true, output: { sheet, axis, start, count, level } }
  },
  'xlsx.outline.set_detail_visibility': (arguments_, services) => {
    const sheet = arguments_.sheet as string
    const axis = arguments_.axis as WorkbookOutlineAxis
    const start = arguments_.start as number
    const count = arguments_.count as number
    const hidden = arguments_.hidden as boolean
    const runtime = services.runtime()
    if (!runtime) throw new Error('Open an XLSX workbook first.')
    const state = xlsxOperationState(services)
    if (!state.outline) throw new Error('Open a file-backed XLSX workbook first.')
    const worksheet = xlsxWorksheet(runtime, sheet)
    const maximum = axis === 'rows' ? worksheet.getMaxRows() : worksheet.getMaxColumns()
    // The following row/column is the explicit summary item.
    if (start + count > maximum) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message:
          'xlsx.outline.set_detail_visibility must leave a following summary item inside the worksheet bounds.',
      }
    }
    if (
      axis === 'rows' &&
      !streamedWorkbookRectangleAvailable(
        xlsxStreamingState(services),
        worksheet.getSheetId(),
        start - 1,
        start + count - 1,
        0,
        0,
      )
    ) {
      return {
        ok: false,
        error: 'execution_failed',
        message:
          'xlsx.outline.set_detail_visibility requires its row metadata to finish streaming.',
      }
    }
    applyWorkbookOutlineDetailVisibility(
      runtime,
      state as Pick<LazyWorkbookState, 'editJournal' | 'outline'>,
      worksheet,
      axis,
      start - 1,
      count,
      hidden,
      services.setPendingEdits,
    )
    return { ok: true, output: { sheet, axis, start, count, hidden } }
  },
  'xlsx.row.set_height': (arguments_, services) => {
    const sheet = arguments_.sheet as string
    const row = arguments_.row as number
    const count = arguments_.count as number
    const heightPoints = arguments_.heightPoints as number
    const worksheet = xlsxWorksheet(services.runtime(), sheet)
    if (row + count - 1 > worksheet.getMaxRows()) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'xlsx.row.set_height must stay inside the worksheet row bounds.',
      }
    }
    applyWorkbookRowHeight(worksheet, row - 1, count, heightPoints)
    return { ok: true, output: { sheet, row, count, heightPoints } }
  },
  'xlsx.row.set_visibility': (arguments_, services) => {
    const sheet = arguments_.sheet as string
    const row = arguments_.row as number
    const count = arguments_.count as number
    const visible = arguments_.visible as boolean
    const worksheet = xlsxWorksheet(services.runtime(), sheet)
    if (row - 1 + count > worksheet.getMaxRows()) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'xlsx.row.set_visibility exceeds the worksheet row boundary.',
      }
    }
    if (visible) worksheet.showRows(row - 1, count)
    else worksheet.hideRows(row - 1, count)
    return { ok: true, output: { sheet, row, count, visible } }
  },
  'xlsx.row.move': async (arguments_, services) => {
    const sheet = arguments_.sheet as string
    const row = arguments_.row as number
    const count = arguments_.count as number
    const beforeRow = arguments_.beforeRow as number
    const runtime = services.runtime()
    const workbook = xlsxWorkbook(runtime)
    const worksheet = xlsxWorksheet(runtime, sheet)
    if (
      row + count - 1 > worksheet.getMaxRows() ||
      beforeRow > worksheet.getMaxRows() ||
      (beforeRow >= row && beforeRow <= row + count)
    ) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message:
          'xlsx.row.move requires an in-sheet source span and a distinct beforeRow outside or after the source boundary.',
      }
    }
    if (!services.isSheetDataComplete?.(worksheet.getSheetId())) {
      throw new Error('xlsx.row.move requires a fully loaded worksheet.')
    }
    const state = services.state?.()
    if (
      state?.file?.sheets.find((candidate) => candidate.id === worksheet.getSheetId())?.pivotRanges
        .length
    ) {
      throw new Error('xlsx.row.move cannot mutate a worksheet containing PivotTables.')
    }
    const lastColumn = worksheet.getMaxColumns() - 1
    const moved = await runtime!.univerAPI.executeCommand('sheet.command.move-rows', {
      unitId: workbook.getId(),
      subUnitId: worksheet.getSheetId(),
      fromRange: {
        startRow: row - 1,
        endRow: row + count - 2,
        startColumn: 0,
        endColumn: lastColumn,
      },
      toRange: {
        startRow: beforeRow - 1,
        endRow: beforeRow - 1,
        startColumn: 0,
        endColumn: lastColumn,
      },
    })
    if (!moved) throw new Error('Univer did not complete xlsx.row.move.')
    return { ok: true, output: { sheet, row, count, beforeRow } }
  },
  'xlsx.row.insert': (arguments_, services) => {
    const sheet = arguments_.sheet as string
    const row = arguments_.row as number
    const count = arguments_.count as number
    xlsxWorksheet(services.runtime(), sheet).insertRowsBefore(row - 1, count)
    return { ok: true, output: { sheet, row, count } }
  },
  'xlsx.sheet.set_freeze': (arguments_, services) => {
    const sheet = arguments_.sheet as string
    const frozenRows = arguments_.frozenRows as number
    const frozenColumns = arguments_.frozenColumns as number
    const worksheet = xlsxWorksheet(services.runtime(), sheet)
    if (frozenRows >= worksheet.getMaxRows() || frozenColumns >= worksheet.getMaxColumns()) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'xlsx.sheet.set_freeze must leave at least one unfrozen row and column.',
      }
    }
    applyWorkbookFreeze(worksheet, frozenRows, frozenColumns)
    return { ok: true, output: { sheet, frozenRows, frozenColumns } }
  },
  'xlsx.sheet.set_formula_view': (arguments_, services) => {
    const sheet = arguments_.sheet as string
    const enabled = arguments_.enabled as boolean
    const runtime = services.runtime()
    if (!runtime) throw new Error('Open an XLSX workbook first.')
    const worksheet = xlsxWorksheet(runtime, sheet)
    applyWorkbookFormulaView(
      runtime,
      xlsxFormulaViewState(services),
      worksheet.getSheetId(),
      enabled,
      services.setPendingEdits,
      (step) => pushWorkbookUndo(runtime, step),
    )
    return { ok: true, output: { sheet, enabled } }
  },
  'xlsx.sheet.set_fit_to_pages': (arguments_, services) => {
    const sheet = arguments_.sheet as string
    const widthPages = arguments_.widthPages as number
    const heightPages = arguments_.heightPages as number
    const runtime = services.runtime()
    if (!runtime) throw new Error('Open an XLSX workbook first.')
    const state = xlsxOperationState(services)
    const worksheet = xlsxWorksheet(runtime, sheet)
    applyWorkbookFitToPages(
      runtime,
      state,
      worksheet.getSheetId(),
      widthPages,
      heightPages,
      services.setPendingEdits,
    )
    return { ok: true, output: { sheet, widthPages, heightPages } }
  },
  'xlsx.sheet.set_gridlines': (arguments_, services) => {
    const sheet = arguments_.sheet as string
    const visible = arguments_.visible as boolean
    applyWorkbookGridlineVisibility(xlsxWorksheet(services.runtime(), sheet), visible)
    return { ok: true, output: { sheet, visible } }
  },
  'xlsx.sheet.set_header_footer': (arguments_, services) => {
    const sheet = arguments_.sheet as string
    const header = arguments_.header as HeaderFooterParts | null
    const footer = arguments_.footer as HeaderFooterParts | null
    if (!isBoundedXlsxHeaderFooterParts(header) || !isBoundedXlsxHeaderFooterParts(footer)) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message:
          'xlsx.sheet.set_header_footer requires null or at least one left, center, or right section of at most 1,024 characters.',
      }
    }
    const runtime = services.runtime()
    if (!runtime) throw new Error('Open an XLSX workbook first.')
    const state = xlsxOperationState(services)
    const worksheet = xlsxWorksheet(runtime, sheet)
    applyWorkbookHeaderFooter(
      runtime,
      state,
      worksheet.getSheetId(),
      header,
      footer,
      services.setPendingEdits,
    )
    return { ok: true, output: { sheet, header, footer } }
  },
  'xlsx.sheet.set_page_margins': (arguments_, services) => {
    const sheet = arguments_.sheet as string
    const margins = arguments_.margins as WorkbookPageMargins
    const runtime = services.runtime()
    if (!runtime) throw new Error('Open an XLSX workbook first.')
    const state = xlsxOperationState(services)
    const worksheet = xlsxWorksheet(runtime, sheet)
    applyWorkbookPageMargins(
      runtime,
      state,
      worksheet.getSheetId(),
      margins,
      services.setPendingEdits,
    )
    return { ok: true, output: { sheet, margins } }
  },
  'xlsx.sheet.set_page_orientation': (arguments_, services) => {
    const sheet = arguments_.sheet as string
    const orientation = arguments_.orientation as WorkbookPageOrientation
    const runtime = services.runtime()
    if (!runtime) throw new Error('Open an XLSX workbook first.')
    const state = xlsxOperationState(services)
    const worksheet = xlsxWorksheet(runtime, sheet)
    applyWorkbookPageOrientation(
      runtime,
      state,
      worksheet.getSheetId(),
      orientation,
      services.setPendingEdits,
    )
    return { ok: true, output: { sheet, orientation } }
  },
  'xlsx.sheet.set_paper_size': (arguments_, services) => {
    const sheet = arguments_.sheet as string
    const paperSize = arguments_.paperSize as WorkbookPaperSize
    const runtime = services.runtime()
    if (!runtime) throw new Error('Open an XLSX workbook first.')
    const state = xlsxOperationState(services)
    const worksheet = xlsxWorksheet(runtime, sheet)
    applyWorkbookPaperSize(
      runtime,
      state,
      worksheet.getSheetId(),
      paperSize,
      services.setPendingEdits,
    )
    return { ok: true, output: { sheet, paperSize } }
  },
  'xlsx.sheet.set_print_gridlines': (arguments_, services) => {
    const sheet = arguments_.sheet as string
    const enabled = arguments_.enabled as boolean
    const runtime = services.runtime()
    if (!runtime) throw new Error('Open an XLSX workbook first.')
    const state = xlsxOperationState(services)
    const worksheet = xlsxWorksheet(runtime, sheet)
    applyWorkbookPrintGridlines(
      runtime,
      state,
      worksheet.getSheetId(),
      enabled,
      services.setPendingEdits,
    )
    return { ok: true, output: { sheet, enabled } }
  },
  'xlsx.sheet.set_print_headings': (arguments_, services) => {
    const sheet = arguments_.sheet as string
    const enabled = arguments_.enabled as boolean
    const runtime = services.runtime()
    if (!runtime) throw new Error('Open an XLSX workbook first.')
    const state = xlsxOperationState(services)
    const worksheet = xlsxWorksheet(runtime, sheet)
    applyWorkbookPrintHeadings(
      runtime,
      state,
      worksheet.getSheetId(),
      enabled,
      services.setPendingEdits,
    )
    return { ok: true, output: { sheet, enabled } }
  },
  'xlsx.sheet.set_print_area': (arguments_, services) => {
    const sheet = arguments_.sheet as string
    const range = arguments_.range as string | null
    const normalizedRange = normalizeXlsxPrintArea(range)
    if (normalizedRange === undefined) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'xlsx.sheet.set_print_area requires an explicit A1 cell range or null.',
      }
    }
    const runtime = services.runtime()
    if (!runtime) throw new Error('Open an XLSX workbook first.')
    const state = xlsxOperationState(services)
    const worksheet = xlsxWorksheet(runtime, sheet)
    applyWorkbookPrintArea(
      runtime,
      state,
      worksheet.getSheetId(),
      normalizedRange,
      services.setPendingEdits,
    )
    return { ok: true, output: { sheet, range: normalizedRange } }
  },
  'xlsx.sheet.set_print_titles': (arguments_, services) => {
    const sheet = arguments_.sheet as string
    const rows = arguments_.rows as string | null
    const normalizedRows = normalizeXlsxPrintTitles(rows)
    if (normalizedRows === undefined) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message:
          'xlsx.sheet.set_print_titles requires an ascending explicit row span of at most 21 rows or null.',
      }
    }
    const runtime = services.runtime()
    if (!runtime) throw new Error('Open an XLSX workbook first.')
    const state = xlsxOperationState(services)
    const worksheet = xlsxWorksheet(runtime, sheet)
    applyWorkbookPrintTitles(
      runtime,
      state,
      worksheet.getSheetId(),
      normalizedRows,
      services.setPendingEdits,
    )
    return { ok: true, output: { sheet, rows: normalizedRows } }
  },
  'xlsx.sheet.set_print_scale': (arguments_, services) => {
    const sheet = arguments_.sheet as string
    const scalePercent = arguments_.scalePercent as number
    const runtime = services.runtime()
    if (!runtime) throw new Error('Open an XLSX workbook first.')
    const state = xlsxOperationState(services)
    const worksheet = xlsxWorksheet(runtime, sheet)
    applyWorkbookPrintScale(
      runtime,
      state,
      worksheet.getSheetId(),
      scalePercent,
      services.setPendingEdits,
    )
    return { ok: true, output: { sheet, scalePercent } }
  },
  'xlsx.sheet.set_protection': (arguments_, services) => {
    const sheet = arguments_.sheet as string
    const nextProtected = arguments_.protected as boolean
    const state = xlsxOperationState(services)
    const worksheet = xlsxWorksheet(services.runtime(), sheet)
    const sheetId = worksheet.getSheetId()
    const guard = applyWorkbookSheetProtection(state, sheetId, nextProtected)
    if (guard) {
      return {
        ok: false,
        error: 'execution_failed',
        message: guard,
      }
    }
    services.setPendingEdits?.(journalSize(state.editJournal))
    return { ok: true, output: { sheet, protected: nextProtected } }
  },
  'xlsx.row.delete': (arguments_, services) => {
    const sheet = arguments_.sheet as string
    const row = arguments_.row as number
    const count = arguments_.count as number
    xlsxWorksheet(services.runtime(), sheet).deleteRows(row - 1, count)
    return { ok: true, output: { sheet, row, count } }
  },
  'xlsx.sheet.add': (arguments_, services) => {
    const name = arguments_.name as string
    if (!isXlsxSheetName(name)) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'xlsx.sheet.add requires a non-empty sheet name of at most 31 characters.',
      }
    }
    xlsxWorkbook(services.runtime()).insertSheet(name)
    return { ok: true, output: { name } }
  },
  'xlsx.sheet.duplicate': (arguments_, services) => {
    const sheet = arguments_.sheet as string
    const name = arguments_.name as string
    if (!isXlsxSheetName(name)) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'xlsx.sheet.duplicate requires a valid XLSX worksheet name.',
      }
    }
    if (!services.duplicateSheet) {
      throw new Error('xlsx.sheet.duplicate is unavailable in this renderer.')
    }
    const workbook = xlsxWorkbook(services.runtime())
    if (workbook.getSheets().some((candidate) => candidate.getSheetName() === name)) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: `An XLSX worksheet named "${name}" already exists.`,
      }
    }
    const source = xlsxWorksheet(services.runtime(), sheet)
    if (!services.isSheetDataComplete?.(source.getSheetId())) {
      throw new Error('xlsx.sheet.duplicate requires a fully loaded source worksheet.')
    }
    const state = services.state?.()
    const sourceMeta = state?.file?.sheets.find((candidate) => candidate.id === source.getSheetId())
    if (sourceMeta && sourceMeta.pivotRanges.length > 0) {
      throw new Error('XLSX worksheets containing PivotTables cannot be duplicated safely.')
    }
    services.duplicateSheet(source, name)
    return { ok: true, output: { sheet, name } }
  },
  'xlsx.sheet.set_tab_color': async (arguments_, services) => {
    const sheet = arguments_.sheet as string
    const color = arguments_.color as string | null
    if (color !== null && !isXlsxColor(color)) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'xlsx.sheet.set_tab_color requires an exact #RRGGBB color or null.',
      }
    }
    const runtime = services.runtime()
    const workbook = xlsxWorkbook(runtime)
    const worksheet = xlsxWorksheet(runtime, sheet)
    const changed = await runtime!.univerAPI.executeCommand('sheet.command.set-tab-color', {
      unitId: workbook.getId(),
      subUnitId: worksheet.getSheetId(),
      value: color,
    })
    if (!changed) throw new Error('Univer did not complete xlsx.sheet.set_tab_color.')
    return { ok: true, output: { sheet, color } }
  },
  'xlsx.sheet.set_visibility': (arguments_, services) => {
    const sheet = arguments_.sheet as string
    const visible = arguments_.visible as boolean
    const worksheet = xlsxWorksheet(services.runtime(), sheet)
    if (visible) worksheet.showSheet()
    else worksheet.hideSheet()
    return { ok: true, output: { sheet, visible } }
  },
  'xlsx.sheet.rename': (arguments_, services) => {
    const sheet = arguments_.sheet as string
    const name = arguments_.name as string
    if (!isXlsxSheetName(name)) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'xlsx.sheet.rename requires a non-empty name of at most 31 characters.',
      }
    }
    xlsxWorksheet(services.runtime(), sheet).setName(name)
    return { ok: true, output: { sheet, name } }
  },
  'xlsx.sheet.delete': (arguments_, services) => {
    const sheet = arguments_.sheet as string
    const workbook = xlsxWorkbook(services.runtime())
    if (workbook.getSheets().length <= 1) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'A workbook must retain at least one worksheet.',
      }
    }
    const target = xlsxWorksheet(services.runtime(), sheet)
    if (!workbook.deleteSheet(target)) throw new Error(`Unable to delete worksheet: ${sheet}`)
    return { ok: true, output: { sheet } }
  },
  'xlsx.sheet.move': (arguments_, services) => {
    const sheet = arguments_.sheet as string
    const position = arguments_.position as number
    const workbook = xlsxWorkbook(services.runtime())
    if (position > workbook.getSheets().length) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'xlsx.sheet.move requires a valid 1-based position.',
      }
    }
    const target = xlsxWorksheet(services.runtime(), sheet)
    workbook.moveSheet(target, position - 1)
    return { ok: true, output: { sheet, position } }
  },
  'xlsx.document.load_staged': async (arguments_, services) => {
    if (!isArrayBuffer(arguments_.data)) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: '$.data must be a hydrated ArrayBuffer.',
      }
    }
    if (
      !(arguments_.name as string).toLowerCase().endsWith('.xlsx') ||
      arguments_.data.byteLength !== arguments_.size
    ) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'xlsx.document.load_staged requires a valid staged XLSX descriptor.',
      }
    }
    await services.loadStaged({
      name: arguments_.name as string,
      data: arguments_.data,
    })
    return {
      ok: true,
      output: { opened: true, fileName: arguments_.name as string },
      checkpointRecovery: false,
    }
  },
  'xlsx.document.save': async (_arguments, services) => {
    const saved = await services.save()
    if (!saved.ok) {
      return {
        ok: false,
        error: 'execution_failed',
        message: saved.message,
      }
    }
    return {
      ok: true,
      output: { saved: true, fileName: saved.fileName },
      checkpointRecovery: false,
    }
  },
} satisfies Record<XlsxOperationId, XlsxOperationHandler>

export async function executeXlsxOperation(
  command: XlsxOperationCommand,
  services: XlsxOperationServices,
): Promise<XlsxOperationExecution> {
  const descriptor = descriptorsByName.get(command.operation)
  if (!descriptor) return { handled: false }

  const validation = validateJsonSchemaValue(descriptor.inputSchema, command.arguments)
  if (!validation.ok) {
    return {
      handled: true,
      operationId: descriptor.id,
      ok: false,
      error: 'invalid_arguments',
      message: validation.error.message,
    }
  }

  try {
    const execution = await handlers[descriptor.id](command.arguments, services)
    return {
      handled: true,
      operationId: descriptor.id,
      ...execution,
    }
  } catch (error) {
    return {
      handled: true,
      operationId: descriptor.id,
      ok: false,
      error: 'execution_failed',
      message: error instanceof Error ? error.message : 'The workbook was not saved.',
    }
  }
}
