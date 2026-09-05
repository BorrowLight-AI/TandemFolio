// Browser-only XLSX package adapter used by the restored community renderer.
import JSZip from 'jszip'

import { applyDefinedNamesState } from '../gateway/xlsx-defined-names'
import { applyChartEdit as applyChartEditToXml } from '../gateway/xlsx-chart'
import { applyCfRules, type CfWireRule } from '../gateway/xlsx-cf'
import { applyDvRules, type DvWireRule } from '../gateway/xlsx-dv'
import {
  applyFilterState as applyFilterStateToXml,
  type SheetFilterState,
} from '../gateway/xlsx-filter'
import {
  applyHyperlinkEdits,
  ensureRelationshipNamespace,
  type HyperlinkEdit,
} from '../gateway/xlsx-hyperlinks'
import {
  applyPageSetupState,
  applyPrintAreas,
  type HeaderFooterParts,
  type SheetPageSetupState,
} from '../gateway/xlsx-page-setup'
import { parsePivotDefinition, setPivotRefreshOnLoad } from '../gateway/xlsx-pivot'
import { applyPivotAdditions, type PivotAddition } from '../gateway/xlsx-pivot-add'
import { applyPivotLayoutExpansions, type PivotRefreshUpdate } from '../gateway/xlsx-pivot-expand'
import { applySheetNotes, readSheetNotes, type SheetNote } from '../gateway/xlsx-notes'
import {
  applySheetProtection as applySheetProtectionToXml,
  applyWorkbookProtection as applyWorkbookProtectionToXml,
  applyProtectedRanges as applyProtectedRangesToXml,
  type ProtectedRangeState,
} from '../gateway/xlsx-protection'
import {
  applyThemeState,
  readThemeState,
  type WorkbookThemeState,
} from '../gateway/xlsx-theme'
import { applySparklineAdditions, type SparklineGroupAdd } from '../gateway/xlsx-sparkline'
import { StylesheetEditor } from '../gateway/xlsx-styles'
import { applyTableAdditions, type TableAddition } from '../gateway/xlsx-table-add'
import {
  applyStructuralOps,
  shiftChartReferences,
  shiftCrossSheetFormulas,
  shiftDefinedNames,
  shiftDrawingAnchors,
  shiftTablePart,
  shiftVmlObjectAnchors,
  type StructuralOp,
} from '../gateway/xlsx-structure'
import {
  addWorksheetOverride,
  addWorksheetRelationship,
  applySheetPlanToWorkbookXml,
  assertNoSheetScopedDefinedNames,
  buildWorksheetPartXml,
  maxRelationshipId,
  maxSheetIdInWorkbook,
  parseSheetElements,
  prepareClonedSheetRels,
  removePartOverride,
  removeRelationshipById,
  renameSheetReferencesInDefinedNames,
  renameSheetReferencesInWorksheet,
  sanitizeClonedWorksheetXml,
  stripPageSetupRelIds,
  validateSheetName,
  worksheetReferencesSheet,
  definedNamesReferenceSheet,
} from '../gateway/xlsx-sheets'
import {
  applyVisualAdditions,
  type MutablePackage,
  type VisualAddition,
} from '../gateway/xlsx-drawing-add'
import { applyVisualEdits as applyVisualEditsToPackage } from '../gateway/xlsx-drawing-edit'
import { readPivotStyleMetadata, readTableStyleMetadata } from './xlsx-pivot-style'
import type { SheetPivotAddition } from '../gateway/xlsx-gateway'
import type {
  WorkbookCellStyle,
  WorkbookChartEdit,
  WorkbookFile,
  WorkbookPagePrintSettings,
  WorkbookRangeResult,
  WorkbookRichRun,
  WorkbookStyleEdit,
  WorkbookVisualEdit,
  WorkbookVisualObject,
} from '../shared/desktop-api'

export type CellScalar = string | number | boolean | null

export interface BrowserCell {
  address: string
  value: CellScalar
  formula?: string
  styleIndex?: number
  rich?: WorkbookRichRun[]
}

export interface BrowserSheet {
  name: string
  hidden: boolean
  path: string
  cells: Map<string, BrowserCell>
  rows: WorkbookRangeResult['rows']
  columnWidths: WorkbookFile['sheets'][number]['columnWidths']
  defaultRowHeight: number | null
  defaultColumnWidth: number | null
  baseColumnWidth?: number
  dataValidations: WorkbookRangeResult['dataValidations']
  conditionalRules: WorkbookRangeResult['conditionalRules']
  merges: string[]
  autoFilterRef: string | null
  hyperlinks: Map<string, string>
  tables: WorkbookFile['sheets'][number]['tables']
  sparklines: WorkbookFile['sheets'][number]['sparklines']
  comments: WorkbookFile['sheets'][number]['comments']
  pivotRanges: WorkbookFile['sheets'][number]['pivotRanges']
  pivotTables: WorkbookFile['sheets'][number]['pivotTables']
  zoomScale?: number
  sheetProtection: { readonly protected: boolean; readonly hasPassword: boolean } | null
  protectedRanges: Array<ProtectedRangeState & { readonly hasPassword: boolean }>
}

export interface BrowserDefinedName {
  readonly name: string
  readonly formula: string
  readonly sheetIndex?: number | undefined
}

export type BrowserVisual = Omit<WorkbookVisualObject, 'sheetId'> & {
  readonly sheetName: string
}

interface StructureSnapshot {
  readonly sheetXml: Map<string, string>
  readonly metadataXml: Map<string, string>
  readonly sheets: BrowserSheet[]
  readonly dirtyPaths: Set<string>
  readonly removedPaths: Set<string>
}

const MAX_ENTRIES = 10_000
const MAX_UNCOMPRESSED_BYTES = 256 * 1024 * 1024

function xmlAttribute(attributes: string, name: string): string | undefined {
  const escaped = name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
  return new RegExp(`(?:^|\\s)${escaped}="([^"]*)"`).exec(attributes)?.[1]
}

function decodeXml(value: string): string {
  return value
    .replace(/&#x([0-9a-f]+);/gi, (_, hex: string) =>
      String.fromCodePoint(Number.parseInt(hex, 16)),
    )
    .replace(/&#([0-9]+);/g, (_, decimal: string) => String.fromCodePoint(Number(decimal)))
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&amp;/g, '&')
}

function escapeXml(value: string): string {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
}

function textContent(xml: string, tag: string): string | undefined {
  const body = new RegExp(`<${tag}\\b[^>]*>([\\s\\S]*?)</${tag}>`).exec(xml)?.[1]
  return body === undefined ? undefined : decodeXml(body.replace(/<[^>]+>/g, ''))
}

function elementAttribute(xml: string, tag: string, name: string): string | undefined {
  const attributes = new RegExp(`<${tag}\\b([^>]*)`).exec(xml)?.[1]
  return attributes === undefined ? undefined : xmlAttribute(attributes, name)
}

function integerAttribute(xml: string, tag: string, name: string): number | undefined {
  const value = elementAttribute(xml, tag, name)
  if (value === undefined) return undefined
  const parsed = Number(value)
  return Number.isInteger(parsed) ? parsed : undefined
}

function numberAttribute(xml: string, tag: string, name: string): number | undefined {
  const value = elementAttribute(xml, tag, name)
  if (value === undefined) return undefined
  const parsed = Number(value)
  return Number.isFinite(parsed) ? parsed : undefined
}

function trueAttribute(xml: string, tag: string, name: string): boolean | undefined {
  const value = elementAttribute(xml, tag, name)
  if (value === undefined) return undefined
  return value === '1' || value === 'true'
}

function decodeHeaderFooterParts(value: string | undefined): HeaderFooterParts | null {
  if (!value) return null
  const markers = [...value.matchAll(/&([LCR])/g)]
  if (markers.length === 0) return { center: value }
  const parts: { left?: string; center?: string; right?: string } = {}
  for (let index = 0; index < markers.length; index += 1) {
    const marker = markers[index]!
    const start = (marker.index ?? 0) + marker[0].length
    const end = markers[index + 1]?.index ?? value.length
    const section = value.slice(start, end)
    if (marker[1] === 'L' && section !== '') parts.left = section
    if (marker[1] === 'C' && section !== '') parts.center = section
    if (marker[1] === 'R' && section !== '') parts.right = section
  }
  return Object.keys(parts).length === 0 ? null : parts
}

function sheetScopedRange(
  workbookXml: string,
  name: '_xlnm.Print_Area' | '_xlnm.Print_Titles',
  sheetIndex: number,
): string | null | undefined {
  for (const match of workbookXml.matchAll(/<definedName\b([^>]*)>([\s\S]*?)<\/definedName>/g)) {
    const attributes = match[1] ?? ''
    if (
      xmlAttribute(attributes, 'name') !== name ||
      Number(xmlAttribute(attributes, 'localSheetId')) !== sheetIndex
    ) {
      continue
    }
    const reference = decodeXml((match[2] ?? '').replace(/<[^>]+>/g, ''))
    const separator = reference.lastIndexOf('!')
    if (separator === -1) return null
    return reference.slice(separator + 1).replace(/\$/g, '')
  }
  return undefined
}

function readPageSetupState(
  worksheetXml: string,
  workbookXml: string,
  sheetName: string,
  sheetIndex: number,
): SheetPageSetupState {
  const orientation = elementAttribute(worksheetXml, 'pageSetup', 'orientation')
  const margins = (() => {
    const left = elementAttribute(worksheetXml, 'pageMargins', 'left')
    const right = elementAttribute(worksheetXml, 'pageMargins', 'right')
    const top = elementAttribute(worksheetXml, 'pageMargins', 'top')
    const bottom = elementAttribute(worksheetXml, 'pageMargins', 'bottom')
    if (left === '0.7' && right === '0.7' && top === '0.75' && bottom === '0.75') return 'normal'
    if (left === '1' && right === '1' && top === '1' && bottom === '1') return 'wide'
    if (left === '0.25' && right === '0.25' && top === '0.75' && bottom === '0.75') {
      return 'narrow'
    }
    return undefined
  })()
  const printGridlines = elementAttribute(worksheetXml, 'printOptions', 'gridLines')
  const printHeadings = elementAttribute(worksheetXml, 'printOptions', 'headings')
  const showGridlines = elementAttribute(worksheetXml, 'sheetView', 'showGridLines')
  const showHeadings = elementAttribute(worksheetXml, 'sheetView', 'showRowColHeaders')
  const rightToLeft = elementAttribute(worksheetXml, 'sheetView', 'rightToLeft')
  const tabColorArgb = elementAttribute(worksheetXml, 'tabColor', 'rgb')
  const manualBreaks = (tag: 'rowBreaks' | 'colBreaks'): number[] => {
    const section = new RegExp(`<${tag}\\b[^>]*>([\\s\\S]*?)</${tag}>`).exec(worksheetXml)?.[1]
    if (section === undefined) return []
    return [...section.matchAll(/<brk\b([^>]*)\/?\s*>/g)].flatMap((match) => {
      const attributes = match[1] ?? ''
      const manual = xmlAttribute(attributes, 'man')
      const id = Number(xmlAttribute(attributes, 'id'))
      return (manual === '1' || manual === 'true') && Number.isInteger(id) && id > 0 ? [id] : []
    })
  }
  return {
    sheetName,
    ...(tabColorArgb && /^[0-9A-Fa-f]{8}$/.test(tabColorArgb)
      ? { tabColor: `#${tabColorArgb.slice(-6).toUpperCase()}` }
      : {}),
    ...(orientation === 'portrait' || orientation === 'landscape' ? { orientation } : {}),
    ...(integerAttribute(worksheetXml, 'pageSetup', 'paperSize') === undefined
      ? {}
      : { paperSize: integerAttribute(worksheetXml, 'pageSetup', 'paperSize') }),
    ...(integerAttribute(worksheetXml, 'pageSetup', 'scale') === undefined
      ? {}
      : { scale: integerAttribute(worksheetXml, 'pageSetup', 'scale') }),
    ...(integerAttribute(worksheetXml, 'pageSetup', 'fitToWidth') === undefined
      ? {}
      : { fitToWidth: integerAttribute(worksheetXml, 'pageSetup', 'fitToWidth') }),
    ...(integerAttribute(worksheetXml, 'pageSetup', 'fitToHeight') === undefined
      ? {}
      : { fitToHeight: integerAttribute(worksheetXml, 'pageSetup', 'fitToHeight') }),
    ...(margins === undefined ? {} : { margins }),
    ...(printGridlines === undefined ? {} : { printGridlines: printGridlines === '1' }),
    ...(printHeadings === undefined ? {} : { printHeadings: printHeadings === '1' }),
    showGridlines: showGridlines !== '0',
    showHeadings: showHeadings !== '0',
    rightToLeft: rightToLeft === '1' || rightToLeft === 'true',
    printArea: sheetScopedRange(workbookXml, '_xlnm.Print_Area', sheetIndex),
    printTitles: sheetScopedRange(workbookXml, '_xlnm.Print_Titles', sheetIndex),
    header: decodeHeaderFooterParts(textContent(worksheetXml, 'oddHeader')),
    footer: decodeHeaderFooterParts(textContent(worksheetXml, 'oddFooter')),
    rowBreaks: manualBreaks('rowBreaks'),
    colBreaks: manualBreaks('colBreaks'),
  }
}

interface BrowserRichString {
  readonly text: string
  readonly rich?: WorkbookRichRun[]
}

function runFlag(properties: string, tag: string): boolean {
  const match = new RegExp(`<${tag}\\b([^>]*)\\/?\\s*>`).exec(properties)
  if (!match) return false
  const value = xmlAttribute(match[1] ?? '', 'val')
  return value !== '0' && value !== 'false'
}

function parseRichString(xml: string): BrowserRichString {
  const runs: WorkbookRichRun[] = []
  for (const match of xml.matchAll(/<r\b[^>]*>([\s\S]*?)<\/r>/g)) {
    const body = match[1] ?? ''
    const properties = /<rPr\b[^>]*>([\s\S]*?)<\/rPr>/.exec(body)?.[1] ?? ''
    const text = [...body.matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/g)]
      .map((part) => decodeXml(part[1] ?? ''))
      .join('')
    const rgb = elementAttribute(properties, 'color', 'rgb')
    const size = Number(elementAttribute(properties, 'sz', 'val'))
    const family = elementAttribute(properties, 'rFont', 'val')
    const vertAlign = elementAttribute(properties, 'vertAlign', 'val')
    runs.push({
      text,
      bold: runFlag(properties, 'b'),
      italic: runFlag(properties, 'i'),
      underline: runFlag(properties, 'u'),
      strikethrough: runFlag(properties, 'strike'),
      ...(rgb && /^(?:[0-9A-Fa-f]{2})?[0-9A-Fa-f]{6}$/.test(rgb)
        ? { color: `#${rgb.slice(-6).toUpperCase()}` }
        : {}),
      ...(Number.isFinite(size) && size > 0 ? { size } : {}),
      ...(family ? { family: decodeXml(family) } : {}),
      ...(vertAlign === 'subscript' || vertAlign === 'superscript' ? { vertAlign } : {}),
    })
  }
  if (runs.length > 0) return { text: runs.map((run) => run.text).join(''), rich: runs }
  return {
    text: [...xml.matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/g)]
      .map((part) => decodeXml(part[1] ?? ''))
      .join(''),
  }
}

function readSharedStrings(xml: string | null): BrowserRichString[] {
  if (!xml) return []
  return [...xml.matchAll(/<si\b[^>]*>([\s\S]*?)<\/si>/g)].map((match) =>
    parseRichString(match[1] ?? ''),
  )
}

function parseCell(xml: string, sharedStrings: BrowserRichString[]): BrowserCell | null {
  const open = /^<c\b([^>]*)>/.exec(xml) ?? /^<c\b([^>]*)\/>/.exec(xml)
  const attributes = open?.[1] ?? ''
  const address = xmlAttribute(attributes, 'r')?.toUpperCase()
  if (!address) return null
  const type = xmlAttribute(attributes, 't')
  const style = xmlAttribute(attributes, 's')
  const formula = textContent(xml, 'f')
  const raw = textContent(xml, 'v')
  let value: CellScalar = null
  let rich: WorkbookRichRun[] | undefined
  if (type === 'inlineStr') {
    const parsed = parseRichString(xml)
    value = parsed.text
    rich = parsed.rich
  } else if (type === 's') {
    const shared = sharedStrings[Number(raw)]
    value = shared?.text ?? ''
    rich = shared?.rich
  } else if (type === 'b') {
    value = raw === '1'
  } else if (type === 'str') {
    value = raw ?? ''
  } else if (raw !== undefined && raw !== '') {
    const numeric = Number(raw)
    value = Number.isFinite(numeric) ? numeric : raw
  }
  return {
    address,
    value,
    ...(formula !== undefined ? { formula } : {}),
    ...(style !== undefined && Number.isInteger(Number(style))
      ? { styleIndex: Number(style) }
      : {}),
    ...(rich === undefined ? {} : { rich }),
  }
}

function parseConditionalRules(xml: string): WorkbookRangeResult['conditionalRules'] {
  const rules: WorkbookRangeResult['conditionalRules'] = []
  for (const section of xml.matchAll(
    /<conditionalFormatting\b([^>]*)>([\s\S]*?)<\/conditionalFormatting>/g,
  )) {
    const refs = decodeXml(xmlAttribute(section[1] ?? '', 'sqref') ?? '')
      .split(/\s+/)
      .filter(Boolean)
    const ranges = refs.flatMap((ref) => {
      try {
        return [tableArea(ref.replace(/\$/g, ''))]
      } catch {
        return []
      }
    })
    if (ranges.length === 0) continue
    for (const match of (section[2] ?? '').matchAll(
      /<cfRule\b([^>]*?)(?:\/>|>([\s\S]*?)<\/cfRule>)/g,
    )) {
      const attributes = match[1] ?? ''
      const body = match[2] ?? ''
      const ruleType = xmlAttribute(attributes, 'type') ?? ''
      const priority = Number(xmlAttribute(attributes, 'priority'))
      if (!ruleType || !Number.isInteger(priority)) continue
      const formulas = [...body.matchAll(/<formula\b[^>]*>([\s\S]*?)<\/formula>/g)].map(
        (formula) => decodeXml(formula[1] ?? ''),
      )
      const scaleBody =
        /<(?:colorScale|dataBar|iconSet)\b[^>]*>([\s\S]*?)<\/(?:colorScale|dataBar|iconSet)>/.exec(
          body,
        )?.[1] ?? ''
      const cfvos = [...scaleBody.matchAll(/<cfvo\b([^>]*)\/>/g)].map((cfvo) => {
        const attrs = cfvo[1] ?? ''
        const kind = xmlAttribute(attrs, 'type') ?? 'num'
        const value = xmlAttribute(attrs, 'val')
        const gte = xmlAttribute(attrs, 'gte')
        return {
          kind,
          ...(value === undefined ? {} : { value: decodeXml(value) }),
          ...(gte === undefined ? {} : { gte: gte === '1' || gte === 'true' }),
        }
      })
      const colors = [...scaleBody.matchAll(/<color\b([^>]*)\/>/g)].flatMap((color) => {
        const argb = xmlAttribute(color[1] ?? '', 'rgb')
        return argb && /^[0-9a-f]{6,8}$/i.test(argb)
          ? [`#${argb.slice(-6).toUpperCase()}`]
          : []
      })
      const operator = xmlAttribute(attributes, 'operator')
      const dxfId = Number(xmlAttribute(attributes, 'dxfId'))
      const rank = Number(xmlAttribute(attributes, 'rank'))
      const scaleAttributes =
        new RegExp(`<${ruleType}\b([^>]*)>`).exec(body)?.[1] ??
        /<(?:colorScale|dataBar|iconSet)\b([^>]*)>/.exec(body)?.[1] ??
        ''
      const minLength = Number(xmlAttribute(scaleAttributes, 'minLength'))
      const maxLength = Number(xmlAttribute(scaleAttributes, 'maxLength'))
      const effectiveMinLength =
        ruleType === 'dataBar' && !Number.isInteger(minLength) ? 10 : minLength
      const effectiveMaxLength =
        ruleType === 'dataBar' && !Number.isInteger(maxLength) ? 90 : maxLength
      rules.push({
        ranges,
        ruleType,
        formulas,
        priority,
        percent: booleanXmlAttribute(attributes, 'percent'),
        bottom: booleanXmlAttribute(attributes, 'bottom'),
        cfvos,
        colors,
        iconReverse: booleanXmlAttribute(scaleAttributes, 'reverse'),
        showValue: xmlAttribute(scaleAttributes, 'showValue') !== '0',
        ...(operator ? { operator } : {}),
        ...(Number.isInteger(dxfId) && dxfId >= 0 ? { dxfIndex: dxfId } : {}),
        ...(Number.isInteger(rank) && rank >= 0 ? { rank } : {}),
        ...(xmlAttribute(attributes, 'text') ? { text: decodeXml(xmlAttribute(attributes, 'text')!) } : {}),
        ...(xmlAttribute(scaleAttributes, 'iconSet')
          ? { iconSetName: xmlAttribute(scaleAttributes, 'iconSet') }
          : {}),
        ...(booleanXmlAttribute(attributes, 'stopIfTrue') ? { stopIfTrue: true } : {}),
        ...(Number.isInteger(effectiveMinLength) && effectiveMinLength >= 0 && effectiveMinLength <= 100
          ? { minLength: effectiveMinLength }
          : {}),
        ...(Number.isInteger(effectiveMaxLength) && effectiveMaxLength >= 0 && effectiveMaxLength <= 100
          ? { maxLength: effectiveMaxLength }
          : {}),
      })
    }
  }
  return rules
}

function booleanXmlAttribute(attributes: string, name: string): boolean {
  const value = xmlAttribute(attributes, name)
  return value === '1' || value === 'true'
}

function parseSheet(
  xml: string,
  name: string,
  path: string,
  sharedStrings: BrowserRichString[],
  relsXml: string | null = null,
  tables: WorkbookFile['sheets'][number]['tables'] = [],
  comments: WorkbookFile['sheets'][number]['comments'] = [],
  pivots: Pick<BrowserSheet, 'pivotRanges' | 'pivotTables'> = {
    pivotRanges: [],
    pivotTables: [],
  },
  hidden = false,
): BrowserSheet {
  const cells = new Map<string, BrowserCell>()
  for (const match of xml.matchAll(/<c\b[^>]*(?:\/>|>[\s\S]*?<\/c>)/g)) {
    const cell = parseCell(match[0], sharedStrings)
    if (cell) cells.set(cell.address, cell)
  }
  const rows: WorkbookRangeResult['rows'] = []
  for (const match of xml.matchAll(/<row\b([^>]*)\/?\s*>/g)) {
    const attributes = match[1] ?? ''
    const rowNumber = Number(xmlAttribute(attributes, 'r'))
    if (!Number.isInteger(rowNumber) || rowNumber < 1 || rowNumber > 1_048_576) continue
    const heightValue = Number(xmlAttribute(attributes, 'ht'))
    const outlineLevel = Number(xmlAttribute(attributes, 'outlineLevel'))
    const hidden = xmlAttribute(attributes, 'hidden')
    const customHeight = xmlAttribute(attributes, 'customHeight')
    const collapsed = xmlAttribute(attributes, 'collapsed')
    const styleIndex = Number(xmlAttribute(attributes, 's'))
    rows.push({
      row: rowNumber - 1,
      hidden: hidden === '1' || hidden === 'true',
      ...(customHeight === '1' || customHeight === 'true' ? { customHeight: true } : {}),
      ...(Number.isFinite(heightValue) && heightValue >= 0 ? { height: heightValue } : {}),
      ...(Number.isInteger(outlineLevel) && outlineLevel >= 1 && outlineLevel <= 7
        ? { outlineLevel }
        : {}),
      ...(collapsed === '1' || collapsed === 'true' ? { collapsed: true } : {}),
      ...(Number.isInteger(styleIndex) && styleIndex >= 0 ? { styleIndex } : {}),
    })
  }
  const columnWidths: WorkbookFile['sheets'][number]['columnWidths'] = []
  for (const match of xml.matchAll(/<col\b([^>]*)\/?\s*>/g)) {
    const attributes = match[1] ?? ''
    const minimum = Number(xmlAttribute(attributes, 'min'))
    const maximum = Number(xmlAttribute(attributes, 'max'))
    if (
      !Number.isInteger(minimum) ||
      !Number.isInteger(maximum) ||
      minimum < 1 ||
      maximum < minimum ||
      maximum > 16_384
    ) {
      continue
    }
    const width = Number(xmlAttribute(attributes, 'width'))
    const outlineLevel = Number(xmlAttribute(attributes, 'outlineLevel'))
    const hidden = xmlAttribute(attributes, 'hidden')
    const collapsed = xmlAttribute(attributes, 'collapsed')
    const styleIndex = Number(xmlAttribute(attributes, 'style'))
    columnWidths.push({
      startColumn: minimum - 1,
      endColumn: maximum - 1,
      hidden: hidden === '1' || hidden === 'true',
      ...(Number.isFinite(width) && width >= 0 ? { width } : {}),
      ...(Number.isInteger(outlineLevel) && outlineLevel >= 1 && outlineLevel <= 7
        ? { outlineLevel }
        : {}),
      ...(collapsed === '1' || collapsed === 'true' ? { collapsed: true } : {}),
      ...(Number.isInteger(styleIndex) && styleIndex >= 0 ? { styleIndex } : {}),
    })
  }
  const defaultRowHeight = numberAttribute(xml, 'sheetFormatPr', 'defaultRowHeight') ?? null
  const defaultColumnWidth = numberAttribute(xml, 'sheetFormatPr', 'defaultColWidth') ?? null
  const baseColumnWidth = integerAttribute(xml, 'sheetFormatPr', 'baseColWidth')
  const rawZoomScale = integerAttribute(xml, 'sheetView', 'zoomScale')
  const zoomScale =
    rawZoomScale !== undefined && rawZoomScale >= 10 && rawZoomScale <= 400
      ? rawZoomScale
      : undefined
  const dataValidations: WorkbookRangeResult['dataValidations'] = []
  for (const match of xml.matchAll(
    /<dataValidation\b([^>]*?)(?:\/>|>([\s\S]*?)<\/dataValidation>)/g,
  )) {
    const attributes = match[1] ?? ''
    const refs = decodeXml(xmlAttribute(attributes, 'sqref') ?? '')
      .split(/\s+/)
      .filter(Boolean)
    const ranges = refs.flatMap((ref) => {
      try {
        return [tableArea(ref.replace(/\$/g, ''))]
      } catch {
        return []
      }
    })
    if (ranges.length === 0) continue
    const body = match[2] ?? ''
    const formula1 = textContent(body, 'formula1')
    const formula2 = textContent(body, 'formula2')
    const rawBoolean = (name: string): boolean => {
      const value = xmlAttribute(attributes, name)
      return value === '1' || value === 'true'
    }
    const optionalText = (name: string): Record<string, string> => {
      const value = xmlAttribute(attributes, name)
      return value === undefined ? {} : { [name]: decodeXml(value) }
    }
    const ruleType = xmlAttribute(attributes, 'type') ?? 'none'
    const operator = xmlAttribute(attributes, 'operator')
    const errorStyle = xmlAttribute(attributes, 'errorStyle')
    dataValidations.push({
      ranges,
      ruleType,
      formulas: [formula1, formula2].filter((value): value is string => value !== undefined),
      allowBlank: rawBoolean('allowBlank'),
      suppressDropdown: rawBoolean('showDropDown'),
      showInputMessage: rawBoolean('showInputMessage'),
      showErrorMessage: rawBoolean('showErrorMessage'),
      ...(operator === undefined ? {} : { operator }),
      ...(errorStyle === undefined ? {} : { errorStyle }),
      ...optionalText('errorTitle'),
      ...optionalText('error'),
      ...optionalText('promptTitle'),
      ...optionalText('prompt'),
    })
  }
  const conditionalRules = parseConditionalRules(xml)
  const merges = [...xml.matchAll(/<mergeCell\b[^>]*\bref="([^"]+)"[^>]*\/>/g)].flatMap((match) =>
    match[1] ? [match[1]] : [],
  )
  const autoFilterRef = elementAttribute(xml, 'autoFilter', 'ref')?.replace(/\$/g, '') ?? null
  const relationshipTargets = new Map<string, string>()
  for (const match of relsXml?.matchAll(/<Relationship\b([^>]*)\/>/g) ?? []) {
    const attributes = match[1] ?? ''
    const id = xmlAttribute(attributes, 'Id')
    const target = xmlAttribute(attributes, 'Target')
    if (id && target) relationshipTargets.set(id, decodeXml(target))
  }
  const hyperlinks = new Map<string, string>()
  for (const match of xml.matchAll(/<hyperlink\b([^>]*)\/>/g)) {
    const attributes = match[1] ?? ''
    const ref = xmlAttribute(attributes, 'ref')?.replace(/\$/g, '')
    if (!ref || ref.includes(':')) continue
    const address = parseAddress(ref)
    const location = xmlAttribute(attributes, 'location')
    const relationshipId = xmlAttribute(attributes, 'r:id')
    const target =
      location === undefined
        ? relationshipId === undefined
          ? undefined
          : relationshipTargets.get(relationshipId)
        : `#${decodeXml(location)}`
    if (target) hyperlinks.set(`${address.row - 1}:${address.column - 1}`, target)
  }
  const protectionAttributes = /<sheetProtection\b([^>]*)\/?\s*>/.exec(xml)?.[1]
  const sheetProtection =
    protectionAttributes === undefined
      ? null
      : {
          protected:
            xmlAttribute(protectionAttributes, 'sheet') === '1' ||
            xmlAttribute(protectionAttributes, 'sheet') === 'true',
          hasPassword:
            xmlAttribute(protectionAttributes, 'password') !== undefined ||
            xmlAttribute(protectionAttributes, 'hashValue') !== undefined,
        }
  const protectedRanges = [...xml.matchAll(/<protectedRange\b([^>]*)\/?\s*>/g)].flatMap(
    (match) => {
      const attributes = match[1] ?? ''
      const name = xmlAttribute(attributes, 'name')
      const sqref = xmlAttribute(attributes, 'sqref')
      if (!name || !sqref) return []
      return [
        {
          name: decodeXml(name),
          sqref: decodeXml(sqref),
          hasPassword:
            xmlAttribute(attributes, 'password') !== undefined ||
            xmlAttribute(attributes, 'hashValue') !== undefined ||
            xmlAttribute(attributes, 'securityDescriptor') !== undefined,
        },
      ]
    },
  )
  const sparklines: WorkbookFile['sheets'][number]['sparklines'] = []
  for (const groupMatch of xml.matchAll(
    /<x14:sparklineGroup\b([^>]*)>([\s\S]*?)<\/x14:sparklineGroup>/g,
  )) {
    const attributes = groupMatch[1] ?? ''
    const body = groupMatch[2] ?? ''
    const rawType = xmlAttribute(attributes, 'type')
    const type = rawType === 'column' || rawType === 'stacked' ? rawType : 'line'
    const seriesArgb = elementAttribute(body, 'x14:colorSeries', 'rgb')
    const negativeArgb = elementAttribute(body, 'x14:colorNegative', 'rgb')
    const cells = [...body.matchAll(/<x14:sparkline\b[^>]*>([\s\S]*?)<\/x14:sparkline>/g)].flatMap(
      (cellMatch) => {
        const cellBody = cellMatch[1] ?? ''
        const sourceRef = textContent(cellBody, 'xm:f')
        const cell = textContent(cellBody, 'xm:sqref')
        return sourceRef && cell ? [{ cell, sourceRef }] : []
      },
    )
    if (cells.length === 0) continue
    sparklines.push({
      type,
      ...(seriesArgb ? { color: `#${seriesArgb.slice(-6)}` } : {}),
      ...(negativeArgb ? { negativeColor: `#${negativeArgb.slice(-6)}` } : {}),
      cells,
    })
  }
  return {
    name,
    hidden,
    path,
    cells,
    rows,
    columnWidths,
    defaultRowHeight,
    defaultColumnWidth,
    ...(baseColumnWidth === undefined ? {} : { baseColumnWidth }),
    ...(zoomScale === undefined ? {} : { zoomScale }),
    dataValidations,
    conditionalRules,
    merges,
    autoFilterRef,
    hyperlinks,
    tables,
    sparklines,
    comments,
    pivotRanges: pivots.pivotRanges,
    pivotTables: pivots.pivotTables,
    sheetProtection,
    protectedRanges,
  }
}

function worksheetRelationshipsPath(worksheetPath: string): string {
  return worksheetPath.replace(/^(.*\/)([^/]+)$/, '$1_rels/$2.rels')
}

function resolvePartTarget(fromPart: string, target: string): string {
  if (target.startsWith('/')) return target.slice(1)
  const resolved = fromPart.split('/').slice(0, -1)
  for (const segment of target.split('/')) {
    if (segment === '..') resolved.pop()
    else if (segment !== '.' && segment !== '') resolved.push(segment)
  }
  return resolved.join('/')
}

interface PartRelationship {
  readonly id: string
  readonly target: string
  readonly type?: string | undefined
}

function readRelationships(xml: string): PartRelationship[] {
  return [...xml.matchAll(/<Relationship\b([^>]*)\/?\s*>/g)].flatMap((match) => {
    const attributes = match[1] ?? ''
    const id = xmlAttribute(attributes, 'Id')
    const target = xmlAttribute(attributes, 'Target')
    if (!id || !target) return []
    const type = xmlAttribute(attributes, 'Type')
    return [{ id, target: decodeXml(target), ...(type === undefined ? {} : { type }) }]
  })
}

function xmlElementBody(xml: string, tag: string): string | undefined {
  const escaped = tag.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
  return new RegExp(`<${escaped}\\b[^>]*>([\\s\\S]*?)</${escaped}>`).exec(xml)?.[1]
}

function xmlTexts(xml: string, tag: string): string[] {
  const escaped = tag.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
  return [...xml.matchAll(new RegExp(`<${escaped}\\b[^>]*>([\\s\\S]*?)</${escaped}>`, 'g'))].map(
    (match) => decodeXml((match[1] ?? '').replace(/<[^>]+>/g, '')),
  )
}

function chartPointValues(xml: string): string[] {
  return [
    ...xml.matchAll(/<c:pt\b[^>]*>[\s\S]*?<c:v\b[^>]*>([\s\S]*?)<\/c:v>[\s\S]*?<\/c:pt>/g),
  ].map((match) => decodeXml((match[1] ?? '').replace(/<[^>]+>/g, '')))
}

function chartNumericPoints(xml: string): { values: number[]; blanks: number[] } {
  const declaredCount = integerAttribute(xml, 'c:ptCount', 'val')
  const points = [...xml.matchAll(/<c:pt\b([^>]*)>([\s\S]*?)<\/c:pt>/g)].flatMap(
    (match) => {
      const index = Number(xmlAttribute(match[1] ?? '', 'idx'))
      if (!Number.isInteger(index) || index < 0) return []
      const raw = textContent(match[2] ?? '', 'c:v')
      const value = raw === undefined || raw.trim() === '' ? undefined : Number(raw)
      return [{ index, value: value !== undefined && Number.isFinite(value) ? value : undefined }]
    },
  )
  const highestIndex = points.reduce((highest, point) => Math.max(highest, point.index), -1)
  const count = Math.max(declaredCount ?? 0, highestIndex + 1)
  const values = Array<number>(count).fill(0)
  const present = new Set<number>()
  const blanks: number[] = []
  for (const point of points) {
    if (point.index >= count) continue
    present.add(point.index)
    if (point.value === undefined) blanks.push(point.index)
    else values[point.index] = point.value
  }
  for (let index = 0; index < count; index += 1) {
    if (!present.has(index)) blanks.push(index)
  }
  blanks.sort((left, right) => left - right)
  return { values, blanks }
}

function chartSeries(chartXml: string): NonNullable<WorkbookVisualObject['chart']>['series'] {
  return [...chartXml.matchAll(/<c:ser\b[^>]*>([\s\S]*?)<\/c:ser>/g)].map((match) => {
    const body = match[1] ?? ''
    const tx = xmlElementBody(body, 'c:tx') ?? ''
    const categoriesBody = xmlElementBody(body, 'c:cat') ?? xmlElementBody(body, 'c:xVal') ?? ''
    const valuesBody = xmlElementBody(body, 'c:val') ?? xmlElementBody(body, 'c:yVal') ?? ''
    const name = textContent(tx, 'c:v') ?? xmlTexts(tx, 'a:t').join('')
    const numericPoints = chartNumericPoints(valuesBody)
    const categories = chartPointValues(categoriesBody)
    const valuesRef = textContent(valuesBody, 'c:f')
    const categoriesRef = textContent(categoriesBody, 'c:f')
    const seriesShape = xmlElementBody(body, 'c:spPr') ?? ''
    const lineShape = xmlElementBody(seriesShape, 'a:ln')
    const fillShape =
      lineShape === undefined
        ? seriesShape
        : seriesShape.replace(/<a:ln\b[^>]*>[\s\S]*?<\/a:ln>/, '')
    const rawColor = elementAttribute(fillShape, 'a:srgbClr', 'val')
    const rawLineColor = lineShape && elementAttribute(lineShape, 'a:srgbClr', 'val')
    const lineColor =
      lineShape === undefined
        ? undefined
        : /<a:noFill\b/.test(lineShape)
          ? 'none'
          : rawLineColor && /^[0-9A-Fa-f]{6}$/.test(rawLineColor)
            ? `#${rawLineColor}`
            : undefined
    const explosionPct = integerAttribute(body, 'c:explosion', 'val')
    return {
      name,
      categories,
      values: numericPoints.values,
      ...(numericPoints.blanks.length === 0 ? {} : { blanks: numericPoints.blanks }),
      ...(valuesRef === undefined ? {} : { valuesRef }),
      ...(categoriesRef === undefined ? {} : { categoriesRef }),
      ...(rawColor && /^[0-9A-Fa-f]{6}$/.test(rawColor) ? { color: `#${rawColor}` } : {}),
      ...(lineColor === undefined ? {} : { lineColor }),
      ...(explosionPct === undefined ? {} : { explosionPct }),
    }
  })
}

function chartMetadata(chartXml: string): NonNullable<WorkbookVisualObject['chart']> | null {
  const plotArea = xmlElementBody(chartXml, 'c:plotArea')
  if (plotArea === undefined) return null
  const chartTypes = [
    ...plotArea.matchAll(
      /<c:(barChart|lineChart|areaChart|pieChart|doughnutChart|scatterChart|radarChart)\b/g,
    ),
  ].map((match) => match[1]!)
  if (chartTypes.length === 0) return null
  const chartTitle = xmlElementBody(chartXml, 'c:title')
  const title =
    chartTitle === undefined
      ? ''
      : xmlTexts(chartTitle, 'a:t').join('') || textContent(chartTitle, 'c:v') || ''
  const legendBody = xmlElementBody(chartXml, 'c:legend')
  const legendPosition =
    legendBody === undefined ? undefined : elementAttribute(legendBody, 'c:legendPos', 'val')
  const legend =
    legendBody === undefined
      ? ('none' as const)
      : (({ r: 'right', b: 'bottom', t: 'top', l: 'left' } as const)[legendPosition ?? 'r'] ??
        ('right' as const))
  const labels = xmlElementBody(plotArea, 'c:dLbls')
  const show = (tag: string) => elementAttribute(labels ?? '', tag, 'val') === '1'
  const dataLabels =
    labels === undefined
      ? undefined
      : show('c:showCatName') && show('c:showPercent')
        ? ('category-percent' as const)
        : show('c:showPercent')
          ? ('percent' as const)
          : show('c:showVal')
            ? ('value' as const)
            : ('none' as const)
  const rawPosition = elementAttribute(labels ?? '', 'c:dLblPos', 'val')
  const dataLabelPosition = (
    { ctr: 'center', inEnd: 'inside-end', outEnd: 'outside-end' } as const
  )[rawPosition ?? '']
  const dataLabelFormat = elementAttribute(labels ?? '', 'c:numFmt', 'formatCode')
  const categoryAxis = xmlElementBody(plotArea, 'c:catAx')
  const valueAxis = xmlElementBody(plotArea, 'c:valAx')
  const categoryTitleBody = categoryAxis && xmlElementBody(categoryAxis, 'c:title')
  const valueTitleBody = valueAxis && xmlElementBody(valueAxis, 'c:title')
  const categoryTitle = categoryTitleBody ? xmlTexts(categoryTitleBody, 'a:t').join('') : undefined
  const valueTitle = valueTitleBody ? xmlTexts(valueTitleBody, 'a:t').join('') : undefined
  const minValue =
    valueAxis === undefined ? undefined : Number(elementAttribute(valueAxis, 'c:min', 'val'))
  const maxValue =
    valueAxis === undefined ? undefined : Number(elementAttribute(valueAxis, 'c:max', 'val'))
  const majorUnit =
    valueAxis === undefined ? undefined : Number(elementAttribute(valueAxis, 'c:majorUnit', 'val'))
  const valueAxisFormat =
    valueAxis === undefined ? undefined : elementAttribute(valueAxis, 'c:numFmt', 'formatCode')
  const valueAxisBounds = {
    ...(Number.isFinite(minValue) ? { min: minValue } : {}),
    ...(Number.isFinite(maxValue) ? { max: maxValue } : {}),
    ...(typeof majorUnit === 'number' && Number.isFinite(majorUnit) && majorUnit > 0
      ? { majorUnit }
      : {}),
    ...(valueAxisFormat === undefined ? {} : { numFmt: decodeXml(valueAxisFormat) }),
  }
  const rawGrouping = elementAttribute(plotArea, 'c:grouping', 'val')
  const grouping =
    rawGrouping === 'clustered' ||
    rawGrouping === 'stacked' ||
    rawGrouping === 'percentStacked' ||
    rawGrouping === 'standard'
      ? rawGrouping
      : undefined
  const gapWidthPct = integerAttribute(plotArea, 'c:gapWidth', 'val')
  const holeSizePct = integerAttribute(plotArea, 'c:holeSize', 'val')
  const rawDispBlanksAs = elementAttribute(chartXml, 'c:dispBlanksAs', 'val')
  const dispBlanksAs =
    rawDispBlanksAs === 'gap' || rawDispBlanksAs === 'zero' || rawDispBlanksAs === 'span'
      ? rawDispBlanksAs
      : undefined
  return {
    chartTypes,
    ...(elementAttribute(plotArea, 'c:barDir', 'val') === undefined
      ? {}
      : { barDirection: elementAttribute(plotArea, 'c:barDir', 'val') }),
    title,
    series: chartSeries(plotArea),
    ...(dispBlanksAs === undefined ? {} : { dispBlanksAs }),
    legend,
    ...(dataLabels === undefined ? {} : { dataLabels }),
    ...(dataLabelPosition === undefined ? {} : { dataLabelPosition }),
    ...(dataLabelFormat === undefined ? {} : { dataLabelFormat: decodeXml(dataLabelFormat) }),
    ...(grouping === undefined ? {} : { grouping }),
    ...(categoryTitle === undefined && valueTitle === undefined
      ? {}
      : {
          axisTitles: {
            ...(categoryTitle === undefined ? {} : { category: categoryTitle }),
            ...(valueTitle === undefined ? {} : { value: valueTitle }),
          },
        }),
    ...(valueAxis === undefined ? {} : { gridlines: /<c:majorGridlines[\s/>]/.test(valueAxis) }),
    ...(Object.keys(valueAxisBounds).length === 0 ? {} : { valueAxis: valueAxisBounds }),
    ...(gapWidthPct === undefined ? {} : { gapWidthPct }),
    ...(holeSizePct === undefined ? {} : { holeSizePct }),
  }
}

function drawingAnchor(anchorXml: string): WorkbookVisualObject['anchor'] {
  const marker = (tag: 'xdr:from' | 'xdr:to') => {
    const body = xmlElementBody(anchorXml, tag) ?? xmlElementBody(anchorXml, tag.slice(4)) ?? ''
    const number = (child: string): number => Number(textContent(body, child) ?? 0)
    return {
      row: number('xdr:row'),
      column: number('xdr:col'),
      rowOffset: number('xdr:rowOff'),
      columnOffset: number('xdr:colOff'),
    }
  }
  const from = marker('xdr:from')
  const explicitTo =
    xmlElementBody(anchorXml, 'xdr:to') !== undefined || xmlElementBody(anchorXml, 'to') !== undefined
  const to = explicitTo
    ? marker('xdr:to')
    : { row: from.row + 15, column: from.column + 7, rowOffset: 0, columnOffset: 0 }
  return {
    fromRow: from.row,
    fromColumn: from.column,
    fromRowOffset: from.rowOffset,
    fromColumnOffset: from.columnOffset,
    toRow: to.row,
    toColumn: to.column,
    toRowOffset: to.rowOffset,
    toColumnOffset: to.columnOffset,
    ...(explicitTo ? { explicitTo: true } : {}),
  }
}

function imageMediaType(path: string): string | undefined {
  const extension = path.split('.').pop()?.toLowerCase()
  return extension === 'png'
    ? 'image/png'
    : extension === 'jpg' || extension === 'jpeg'
      ? 'image/jpeg'
      : extension === 'gif'
        ? 'image/gif'
        : extension === 'svg'
          ? 'image/svg+xml'
          : extension === 'bmp'
            ? 'image/bmp'
            : extension === 'webp'
              ? 'image/webp'
              : extension === 'emf'
                ? 'image/x-emf'
                : extension === 'wmf'
                  ? 'image/x-wmf'
                  : extension === 'emz'
                    ? 'image/x-emz'
                    : extension === 'wmz'
                      ? 'image/x-wmz'
                      : undefined
}

interface OleObjectRecord {
  readonly shapeId: number
  readonly progId: string
  mediaPath?: string | undefined
  anchor?: WorkbookVisualObject['anchor'] | undefined
  lineColor?: string | undefined
  fillColor?: string | undefined
}

function vmlBoolean(value: string | undefined): boolean | undefined {
  if (value === 't' || value === 'true' || value === '1') return true
  if (value === 'f' || value === 'false' || value === '0') return false
  return undefined
}

function vmlColor(value: string | undefined, fallback: string): string {
  const head = value?.trim().split(/\s+/, 1)[0] ?? ''
  return /^#[0-9a-f]{6}$/i.test(head) ? head.toUpperCase() : fallback
}

function vmlAnchor(value: string | undefined): WorkbookVisualObject['anchor'] | undefined {
  const numbers = value?.split(',').map((part) => Number(part.trim())) ?? []
  if (numbers.length !== 8 || numbers.some((entry) => !Number.isFinite(entry) || entry < 0)) {
    return undefined
  }
  const [fromColumn, fromColumnPx, fromRow, fromRowPx, toColumn, toColumnPx, toRow, toRowPx] =
    numbers as [number, number, number, number, number, number, number, number]
  return {
    fromRow,
    fromColumn,
    fromRowOffset: fromRowPx * 9525,
    fromColumnOffset: fromColumnPx * 9525,
    toRow,
    toColumn,
    toRowOffset: toRowPx * 9525,
    toColumnOffset: toColumnPx * 9525,
    explicitTo: true,
  }
}

async function readOleObjects(
  worksheetXml: string,
  worksheetPath: string,
  relationships: readonly PartRelationship[],
  readPart: (path: string) => Promise<string | null>,
): Promise<OleObjectRecord[]> {
  if (!worksheetXml.includes('oleObject')) return []
  const records = new Map<number, OleObjectRecord>()
  for (const match of worksheetXml.matchAll(
    /<oleObject\b([^>]*?)(?:\/>|>([\s\S]*?)<\/oleObject>)/g,
  )) {
    const attributes = match[1] ?? ''
    const shapeId = Number(xmlAttribute(attributes, 'shapeId'))
    const progId = xmlAttribute(attributes, 'progId')
    if (!Number.isInteger(shapeId) || shapeId < 0 || !progId) continue
    const body = match[2] ?? ''
    const objectPr = /<objectPr\b([^>]*)>([\s\S]*?)<\/objectPr>/.exec(body)
    const previewId = objectPr ? xmlAttribute(objectPr[1] ?? '', 'r:id') : undefined
    const previewRelationship = relationships.find((entry) => entry.id === previewId)
    const anchorXml = objectPr ? xmlElementBody(objectPr[0], 'anchor') : undefined
    const next = records.get(shapeId) ?? { shapeId, progId: decodeXml(progId) }
    if (!next.mediaPath && previewRelationship) {
      next.mediaPath = resolvePartTarget(worksheetPath, previewRelationship.target)
    }
    if (!next.anchor && anchorXml) next.anchor = drawingAnchor(anchorXml)
    records.set(shapeId, next)
  }

  const legacyId = elementAttribute(worksheetXml, 'legacyDrawing', 'r:id')
  const legacyRelationship = relationships.find((entry) => entry.id === legacyId)
  if (!legacyRelationship) return [...records.values()]
  const vmlPath = resolvePartTarget(worksheetPath, legacyRelationship.target)
  const vmlXml = await readPart(vmlPath)
  if (!vmlXml) return [...records.values()]
  const vmlRelsXml = await readPart(worksheetRelationshipsPath(vmlPath))
  const vmlRelationships = vmlRelsXml ? readRelationships(vmlRelsXml) : []
  const inherited = new Map<string, { filled?: boolean; stroked?: boolean }>()
  for (const match of vmlXml.matchAll(/<v:shapetype\b([^>]*)[\s\S]*?<\/v:shapetype>|<v:shapetype\b([^>]*)\/>/g)) {
    const attributes = match[1] ?? match[2] ?? ''
    const id = xmlAttribute(attributes, 'id')
    if (!id) continue
    const filled = vmlBoolean(xmlAttribute(attributes, 'filled'))
    const stroked = vmlBoolean(xmlAttribute(attributes, 'stroked'))
    inherited.set(`#${id}`, {
      ...(filled === undefined ? {} : { filled }),
      ...(stroked === undefined ? {} : { stroked }),
    })
  }
  for (const match of vmlXml.matchAll(/<v:shape\b([^>]*)>([\s\S]*?)<\/v:shape>/g)) {
    const attributes = match[1] ?? ''
    const body = match[2] ?? ''
    const shapeId = Number(/_s([0-9]+)$/.exec(xmlAttribute(attributes, 'id') ?? '')?.[1])
    const record = records.get(shapeId)
    if (!record) continue
    const defaults = inherited.get(xmlAttribute(attributes, 'type') ?? '')
    const filled = vmlBoolean(xmlAttribute(attributes, 'filled')) ?? defaults?.filled ?? true
    const stroked = vmlBoolean(xmlAttribute(attributes, 'stroked')) ?? defaults?.stroked ?? true
    record.fillColor = filled
      ? vmlColor(xmlAttribute(attributes, 'fillcolor'), '#FFFFFF')
      : undefined
    record.lineColor = stroked
      ? vmlColor(xmlAttribute(attributes, 'strokecolor'), '#000000')
      : undefined
    if (!record.anchor) record.anchor = vmlAnchor(textContent(body, 'x:Anchor'))
    if (!record.mediaPath) {
      const imageId =
        elementAttribute(body, 'v:imagedata', 'o:relid') ??
        elementAttribute(body, 'v:imagedata', 'r:id')
      const imageRelationship = vmlRelationships.find((entry) => entry.id === imageId)
      if (imageRelationship) record.mediaPath = resolvePartTarget(vmlPath, imageRelationship.target)
    }
  }
  return [...records.values()]
}

async function readSheetVisuals(
  worksheetXml: string,
  worksheetPath: string,
  sheetName: string,
  relsXml: string | null,
  readPart: (path: string) => Promise<string | null>,
): Promise<BrowserVisual[]> {
  if (!relsXml) return []
  const sheetRelationships = readRelationships(relsXml)
  const oleObjects = await readOleObjects(
    worksheetXml,
    worksheetPath,
    sheetRelationships,
    readPart,
  )
  const drawingRelationshipId = elementAttribute(worksheetXml, 'drawing', 'r:id')
  const drawingRelationship = sheetRelationships.find(
    (relationship) =>
      relationship.id === drawingRelationshipId || relationship.type?.endsWith('/drawing'),
  )
  if (!drawingRelationship) {
    return oleObjects.flatMap((ole, index) => {
      if (!ole.anchor) return []
      const mediaType = ole.mediaPath ? imageMediaType(ole.mediaPath) : undefined
      return [{
        id: `ole-${index + 1}`,
        sheetName,
        kind: 'ole' as const,
        anchor: ole.anchor,
        progId: ole.progId,
        ...(ole.mediaPath && mediaType ? { mediaPath: ole.mediaPath, mediaType } : {}),
        ...(ole.lineColor === undefined ? {} : { lineColor: ole.lineColor }),
        ...(ole.fillColor === undefined ? {} : { fillColor: ole.fillColor }),
      }]
    })
  }
  const drawingPath = resolvePartTarget(worksheetPath, drawingRelationship.target)
  const drawingXml = await readPart(drawingPath)
  if (!drawingXml) return []
  const drawingRelsXml = await readPart(worksheetRelationshipsPath(drawingPath))
  const drawingRelationships = drawingRelsXml ? readRelationships(drawingRelsXml) : []
  const visuals: BrowserVisual[] = []
  const fallbackSlots = new Map<number, number>()
  const anchors = [
    ...drawingXml.matchAll(
      /<xdr:(twoCellAnchor|oneCellAnchor|absoluteAnchor)\b[^>]*>[\s\S]*?<\/xdr:\1>/g,
    ),
  ]
  for (const [drawingIndex, match] of anchors.entries()) {
    const anchorXml = match[0]
    const shapeXml = xmlElementBody(anchorXml, 'xdr:sp')
    if (shapeXml !== undefined) {
      const shapeType = elementAttribute(shapeXml, 'a:prstGeom', 'prst')
      if (!shapeType) continue
      const name = elementAttribute(shapeXml, 'xdr:cNvPr', 'name')
      const nativeId = Number(elementAttribute(shapeXml, 'xdr:cNvPr', 'id'))
      const rawColor = elementAttribute(shapeXml, 'a:srgbClr', 'val')
      const text = xmlTexts(shapeXml, 'a:t').join('')
      visuals.push({
        id: `file-shape-${drawingPath.replace(/[^A-Za-z0-9_-]/g, '-')}-${drawingIndex}`,
        sheetName,
        kind: 'shape',
        anchor: drawingAnchor(anchorXml),
        drawingPath,
        drawingIndex,
        shapeType,
        ...(name === undefined ? {} : { name: decodeXml(name) }),
        ...(rawColor && /^[0-9A-Fa-f]{6}$/.test(rawColor) ? { fillColor: `#${rawColor}` } : {}),
        ...(text === '' ? {} : { text }),
      })
      if (
        Number.isInteger(nativeId) &&
        oleObjects.some((candidate) => candidate.shapeId === nativeId)
      ) {
        fallbackSlots.set(nativeId, visuals.length - 1)
      }
      continue
    }
    const pictureXml = xmlElementBody(anchorXml, 'xdr:pic')
    if (pictureXml !== undefined) {
      const relationshipId = elementAttribute(pictureXml, 'a:blip', 'r:embed')
      const relationship = drawingRelationships.find(
        (candidate) => candidate.id === relationshipId && candidate.type?.endsWith('/image'),
      )
      if (!relationship) continue
      const mediaPath = resolvePartTarget(drawingPath, relationship.target)
      const mediaType = imageMediaType(mediaPath)
      if (!mediaType) continue
      const name = elementAttribute(pictureXml, 'xdr:cNvPr', 'name')
      visuals.push({
        id: `file-image-${mediaPath.replace(/[^A-Za-z0-9_-]/g, '-')}-${drawingIndex}`,
        sheetName,
        kind: 'image',
        anchor: drawingAnchor(anchorXml),
        mediaPath,
        mediaType,
        drawingPath,
        drawingIndex,
        ...(name === undefined ? {} : { name: decodeXml(name) }),
      })
      continue
    }
    const relationshipId = elementAttribute(anchorXml, 'c:chart', 'r:id')
    if (!relationshipId) continue
    const relationship = drawingRelationships.find(
      (candidate) => candidate.id === relationshipId && candidate.type?.endsWith('/chart'),
    )
    if (!relationship) continue
    const chartPath = resolvePartTarget(drawingPath, relationship.target)
    const chartXml = await readPart(chartPath)
    if (!chartXml) continue
    const chart = chartMetadata(chartXml)
    if (!chart) continue
    const name = elementAttribute(anchorXml, 'xdr:cNvPr', 'name')
    visuals.push({
      id: `file-chart-${chartPath.replace(/[^A-Za-z0-9_-]/g, '-')}`,
      sheetName,
      kind: 'chart',
      anchor: drawingAnchor(anchorXml),
      chart,
      chartPath,
      drawingPath,
      drawingIndex,
      ...(name === undefined ? {} : { name: decodeXml(name) }),
    })
  }
  for (const [index, ole] of oleObjects.entries()) {
    const slot = fallbackSlots.get(ole.shapeId)
    const anchor = ole.anchor ?? (slot === undefined ? undefined : visuals[slot]?.anchor)
    if (!anchor) continue
    const mediaType = ole.mediaPath ? imageMediaType(ole.mediaPath) : undefined
    const visual: BrowserVisual = {
      id: `ole-${index + 1}`,
      sheetName,
      kind: 'ole',
      anchor,
      progId: ole.progId,
      ...(ole.mediaPath && mediaType ? { mediaPath: ole.mediaPath, mediaType } : {}),
      ...(ole.lineColor === undefined ? {} : { lineColor: ole.lineColor }),
      ...(ole.fillColor === undefined ? {} : { fillColor: ole.fillColor }),
    }
    if (slot === undefined) visuals.push(visual)
    else visuals[slot] = visual
  }
  const fallbackIndexes = new Set(fallbackSlots.values())
  return visuals.filter((visual, index) => !fallbackIndexes.has(index) || visual.kind === 'ole')
}

function tableArea(ref: string): WorkbookFile['sheets'][number]['tables'][number]['range'] {
  const parts = ref.replace(/\$/g, '').split(':')
  if (parts.length > 2 || !parts[0]) throw new Error(`Invalid table range: ${ref}`)
  const start = parseAddress(parts[0])
  const end = parseAddress(parts[1] ?? parts[0])
  if (end.row < start.row || end.column < start.column) {
    throw new Error(`Invalid table range: ${ref}`)
  }
  return {
    startRow: start.row - 1,
    startColumn: start.column - 1,
    endRow: end.row - 1,
    endColumn: end.column - 1,
  }
}

async function readSheetTables(
  worksheetXml: string,
  worksheetPath: string,
  relsXml: string | null,
  readPart: (path: string) => Promise<string | null>,
  themeColors?: readonly string[],
): Promise<WorkbookFile['sheets'][number]['tables']> {
  if (!relsXml) return []
  const relationshipTargets = new Map<string, string>()
  for (const match of relsXml.matchAll(/<Relationship\b([^>]*)\/?\s*>/g)) {
    const attributes = match[1] ?? ''
    const id = xmlAttribute(attributes, 'Id')
    const target = xmlAttribute(attributes, 'Target')
    if (id && target) relationshipTargets.set(id, decodeXml(target))
  }
  const tables: WorkbookFile['sheets'][number]['tables'] = []
  for (const match of worksheetXml.matchAll(/<tablePart\b([^>]*)\/?\s*>/g)) {
    const relationshipId = xmlAttribute(match[1] ?? '', 'r:id')
    const target = relationshipId ? relationshipTargets.get(relationshipId) : undefined
    if (!target) continue
    const tableXml = await readPart(resolvePartTarget(worksheetPath, target))
    if (!tableXml) continue
    const ref = elementAttribute(tableXml, 'table', 'ref')
    if (!ref) continue
    const headerRowCount = integerAttribute(tableXml, 'table', 'headerRowCount') ?? 1
    const totalsRowCount = integerAttribute(tableXml, 'table', 'totalsRowCount') ?? 0
    const name =
      elementAttribute(tableXml, 'table', 'displayName') ??
      elementAttribute(tableXml, 'table', 'name')
    const columns = [...tableXml.matchAll(/<tableColumn\b([^>]*)\/?\s*>/g)].flatMap((entry) => {
      const columnName = xmlAttribute(entry[1] ?? '', 'name')
      return columnName === undefined ? [] : [decodeXml(columnName)]
    })
    const styleName = elementAttribute(tableXml, 'tableStyleInfo', 'name')
    const showRowStripes = elementAttribute(tableXml, 'tableStyleInfo', 'showRowStripes')
    const showColumnStripes = elementAttribute(tableXml, 'tableStyleInfo', 'showColumnStripes')
    const autoFilterXml = xmlElementBody(tableXml, 'autoFilter')
    const filterActive =
      autoFilterXml !== undefined &&
      /<(?:[A-Za-z_][\w.-]*:)?(?:filters|customFilters|dynamicFilter|top10|colorFilter|iconFilter)\b/.test(
        autoFilterXml,
      )
    tables.push({
      range: tableArea(ref),
      headerRowCount,
      showRowStripes: showRowStripes === '1' || showRowStripes === 'true',
      showColumnStripes: showColumnStripes === '1' || showColumnStripes === 'true',
      ...(filterActive ? { filterActive: true } : {}),
      ...(name ? { name: decodeXml(name) } : {}),
      ...(columns.length > 0 ? { columns } : {}),
      ...(totalsRowCount > 0 ? { totalsRowCount } : {}),
      ...(styleName ? { styleName: decodeXml(styleName) } : {}),
      ...readTableStyleMetadata(styleName ? decodeXml(styleName) : undefined, themeColors),
    })
  }
  return tables
}

async function readSheetPivots(
  worksheetPath: string,
  relsXml: string | null,
  readPart: (path: string) => Promise<string | null>,
  themeColors?: readonly string[],
): Promise<Pick<BrowserSheet, 'pivotRanges' | 'pivotTables'>> {
  if (!relsXml) return { pivotRanges: [], pivotTables: [] }
  const pivotTables: WorkbookFile['sheets'][number]['pivotTables'] = []
  const pivotRanges: WorkbookFile['sheets'][number]['pivotRanges'] = []
  for (const relationship of readRelationships(relsXml)) {
    if (!relationship.type?.endsWith('/pivotTable')) continue
    const path = resolvePartTarget(worksheetPath, relationship.target)
    const pivotXml = await readPart(path)
    if (!pivotXml) continue
    const outputRef = elementAttribute(pivotXml, 'location', 'ref')?.replace(/\$/g, '')
    if (!outputRef) continue
    let range: WorkbookFile['sheets'][number]['pivotRanges'][number]
    try {
      range = tableArea(outputRef)
    } catch {
      continue
    }
    const pivotRelsXml = await readPart(worksheetRelationshipsPath(path))
    const cacheRelationship = pivotRelsXml
      ? readRelationships(pivotRelsXml).find((candidate) =>
          candidate.type?.endsWith('/pivotCacheDefinition'),
        )
      : undefined
    pivotTables.push({
      path,
      cachePath: cacheRelationship ? resolvePartTarget(path, cacheRelationship.target) : null,
      outputRef,
      ...readPivotStyleMetadata(pivotXml, themeColors),
    })
    pivotRanges.push(range)
    if (pivotTables.length >= 100) break
  }
  return { pivotRanges, pivotTables }
}

function columnNumber(label: string): number {
  let value = 0
  for (const char of label) value = value * 26 + char.charCodeAt(0) - 64
  return value
}

function columnLabel(column: number): string {
  let label = ''
  for (let current = column; current > 0; current = Math.floor((current - 1) / 26)) {
    label = String.fromCharCode(((current - 1) % 26) + 65) + label
  }
  return label
}

function parseAddress(address: string): { row: number; column: number } {
  const match = /^([A-Z]+)([1-9][0-9]*)$/i.exec(address)
  if (!match?.[1] || !match[2]) throw new Error(`Invalid cell address: ${address}`)
  return { column: columnNumber(match[1].toUpperCase()), row: Number(match[2]) }
}

function cellXml(cell: BrowserCell): string {
  const style = cell.styleIndex === undefined ? '' : ` s="${cell.styleIndex}"`
  if (cell.value === null && !cell.formula) {
    return cell.styleIndex === undefined ? '' : `<c r="${cell.address}"${style}/>`
  }
  const formula = cell.formula ? `<f>${escapeXml(cell.formula)}</f>` : ''
  if (typeof cell.value === 'string') {
    if (formula)
      return `<c r="${cell.address}"${style} t="str">${formula}<v>${escapeXml(cell.value)}</v></c>`
    return `<c r="${cell.address}"${style} t="inlineStr"><is><t xml:space="preserve">${escapeXml(cell.value)}</t></is></c>`
  }
  if (typeof cell.value === 'boolean') {
    return `<c r="${cell.address}"${style} t="b">${formula}<v>${cell.value ? 1 : 0}</v></c>`
  }
  return `<c r="${cell.address}"${style}>${formula}<v>${cell.value ?? ''}</v></c>`
}

function patchCell(xml: string, cell: BrowserCell): string {
  const pattern = new RegExp(`<c\\b[^>]*\\br="${cell.address}"[^>]*(?:\\/>|>[\\s\\S]*?<\\/c>)`)
  const next = cellXml(cell)
  if (pattern.test(xml)) return xml.replace(pattern, () => next)
  if (!next) return xml
  const { row } = parseAddress(cell.address)
  const rowPattern = new RegExp(`(<row\\b[^>]*\\br="${row}"[^>]*>)([\\s\\S]*?)(<\\/row>)`)
  if (rowPattern.test(xml)) {
    return xml.replace(
      rowPattern,
      (_all, start: string, body: string, end: string) => `${start}${body}${next}${end}`,
    )
  }
  if (xml.includes('<sheetData/>')) {
    return xml.replace('<sheetData/>', () => `<sheetData><row r="${row}">${next}</row></sheetData>`)
  }
  if (!xml.includes('</sheetData>')) throw new Error('Worksheet has no sheetData element.')
  return xml.replace('</sheetData>', () => `<row r="${row}">${next}</row></sheetData>`)
}

async function zipText(zip: JSZip, path: string): Promise<string> {
  const entry = zip.file(path)
  if (!entry) throw new Error(`Workbook is missing ${path}.`)
  return entry.async('text')
}

function relationshipTarget(relsXml: string, relationshipId: string): string {
  for (const match of relsXml.matchAll(/<Relationship\b([^>]*)\/>/g)) {
    const attributes = match[1] ?? ''
    if (xmlAttribute(attributes, 'Id') === relationshipId) {
      const target = xmlAttribute(attributes, 'Target')
      if (target) return target
    }
  }
  throw new Error(`Workbook relationship is missing: ${relationshipId}`)
}

export class BrowserWorkbook {
  readonly #dirtyPaths = new Set<string>()
  readonly #sheetXml = new Map<string, string>()
  readonly #zip: JSZip
  readonly #stylesheet: StylesheetEditor | null
  readonly #sharedStrings: BrowserRichString[]
  readonly #metadataXml: Map<string, string>
  readonly #removedPaths = new Set<string>()
  readonly #structureUndo: StructureSnapshot[] = []
  readonly #structureRedo: StructureSnapshot[] = []

  constructor(
    readonly name: string,
    readonly sheets: BrowserSheet[],
    readonly visuals: BrowserVisual[],
    zip: JSZip,
    sheetXml: ReadonlyMap<string, string>,
    stylesheet: StylesheetEditor | null,
    sharedStrings: BrowserRichString[],
    metadataXml: ReadonlyMap<string, string>,
  ) {
    this.#zip = zip
    this.#sheetXml = new Map(sheetXml)
    this.#stylesheet = stylesheet
    this.#sharedStrings = sharedStrings
    this.#metadataXml = new Map(metadataXml)
  }

  get dirty(): boolean {
    return this.#dirtyPaths.size > 0 || this.#removedPaths.size > 0
  }

  styleCatalog(): WorkbookCellStyle[] {
    return this.#stylesheet?.styleCatalog() ?? []
  }

  dxfCatalog(): WorkbookCellStyle[] {
    return this.#stylesheet?.dxfCatalog() ?? []
  }

  async readBinary(path: string): Promise<Uint8Array> {
    if (path.startsWith('/') || path.split('/').includes('..') || this.#removedPaths.has(path)) {
      throw new Error('Workbook media path is invalid.')
    }
    const entry = this.#zip.file(path)
    if (!entry) throw new Error(`Workbook is missing ${path}.`)
    return entry.async('uint8array')
  }

  async readPivotDefinition(path: string, cachePath: string) {
    for (const candidate of [path, cachePath]) {
      if (
        candidate.startsWith('/') ||
        candidate.split('/').includes('..') ||
        !/^xl\/[A-Za-z0-9._/-]+\.xml$/.test(candidate)
      ) {
        throw new Error('Workbook pivot path is invalid.')
      }
    }
    return parsePivotDefinition(
      await this.#readPackageText(path),
      await this.#readPackageText(cachePath),
    )
  }

  async applyPivots(
    additions: readonly SheetPivotAddition[],
    refreshPaths: readonly string[],
    refreshUpdates: readonly PivotRefreshUpdate[],
  ): Promise<void> {
    if (additions.length === 0 && refreshPaths.length === 0 && refreshUpdates.length === 0) return
    const overlayText = new Map<string, string>()
    const overlayBinary = new Map<string, Uint8Array>()
    const removed = new Set<string>()
    const paths = async (): Promise<readonly string[]> =>
      [
        ...new Set([
          ...Object.values(this.#zip.files)
            .filter((entry) => !entry.dir)
            .map((entry) => entry.name),
          ...this.#sheetXml.keys(),
          ...this.#metadataXml.keys(),
          ...overlayText.keys(),
          ...overlayBinary.keys(),
        ]),
      ].filter((path) => !removed.has(path))
    const pkg: MutablePackage = {
      paths,
      has: async (path) =>
        !removed.has(path) &&
        (overlayText.has(path) ||
          overlayBinary.has(path) ||
          this.#sheetXml.has(path) ||
          this.#metadataXml.has(path) ||
          this.#zip.file(path) !== null),
      readText: async (path) => {
        if (removed.has(path)) throw new Error(`Workbook is missing ${path}.`)
        return overlayText.get(path) ?? (await this.#readPackageText(path))
      },
      write: (path, content) => {
        removed.delete(path)
        overlayText.set(path, content)
      },
      add: (path, content) => {
        removed.delete(path)
        overlayText.set(path, content)
      },
      addBinary: (path, bytes) => {
        removed.delete(path)
        overlayBinary.set(path, bytes)
      },
      remove: (path) => {
        overlayText.delete(path)
        overlayBinary.delete(path)
        removed.add(path)
      },
    }

    const resolvedUpdates = refreshUpdates.map((update) => ({
      ...update,
      worksheetPath:
        update.worksheetPath ??
        (update.sheetName === undefined ? undefined : this.#sheet(update.sheetName).path),
    }))
    if (resolvedUpdates.some((update) => update.worksheetPath === undefined)) {
      throw new Error('A pivot refresh update needs a worksheet path or sheet name.')
    }
    await applyPivotLayoutExpansions(
      pkg,
      resolvedUpdates as readonly PivotRefreshUpdate[],
      new Set<string>(),
    )
    for (const cachePath of refreshPaths) {
      pkg.write(cachePath, setPivotRefreshOnLoad(await pkg.readText(cachePath)))
    }
    if (additions.length > 0) {
      const resolvedAdditions: PivotAddition[] = additions.map(({ sheetName, ...addition }) => ({
        ...addition,
        worksheetPath: this.#sheet(sheetName).path,
      }))
      const workbookPath = 'xl/workbook.xml'
      pkg.write(
        workbookPath,
        await applyPivotAdditions(
          pkg,
          resolvedAdditions,
          await pkg.readText(workbookPath),
          new Set<string>(),
        ),
      )
    }

    for (const path of removed) {
      this.#sheetXml.delete(path)
      this.#metadataXml.delete(path)
      this.#dirtyPaths.delete(path)
      this.#removedPaths.add(path)
    }
    for (const [path, content] of overlayText) {
      if (this.#sheetXml.has(path)) this.#sheetXml.set(path, content)
      else this.#metadataXml.set(path, content)
      this.#removedPaths.delete(path)
      this.#dirtyPaths.add(path)
    }
    for (const [path, bytes] of overlayBinary) {
      this.#zip.file(path, bytes)
      this.#removedPaths.delete(path)
    }
  }

  cell(sheetName: string, address: string): BrowserCell | undefined {
    return this.#sheet(sheetName).cells.get(address.toUpperCase())
  }

  definedNames(): BrowserDefinedName[] {
    const workbookXml = this.#metadataXml.get('xl/workbook.xml')!
    return [...workbookXml.matchAll(/<definedName\b([^>]*)>([\s\S]*?)<\/definedName>/g)].flatMap(
      (match) => {
        const attributes = match[1] ?? ''
        const name = xmlAttribute(attributes, 'name')
        if (!name || name.startsWith('_xlnm') || /\bhidden="(?:1|true)"/.test(attributes)) {
          return []
        }
        const localSheetId = xmlAttribute(attributes, 'localSheetId')
        const sheetIndex = localSheetId === undefined ? undefined : Number(localSheetId)
        if (sheetIndex !== undefined && !Number.isInteger(sheetIndex)) return []
        return [
          {
            name: decodeXml(name),
            formula: decodeXml((match[2] ?? '').replace(/<[^>]+>/g, '')),
            ...(sheetIndex === undefined ? {} : { sheetIndex }),
          },
        ]
      },
    )
  }

  replaceDefinedNames(names: readonly BrowserDefinedName[]): void {
    const workbookPath = 'xl/workbook.xml'
    this.#metadataXml.set(
      workbookPath,
      applyDefinedNamesState(this.#metadataXml.get(workbookPath)!, {
        names,
        preserveNames: [],
      }),
    )
    this.#dirtyPaths.add(workbookPath)
  }

  theme(): WorkbookThemeState | null {
    const xml = this.#metadataXml.get('xl/theme/theme1.xml')
    return xml === undefined ? null : readThemeState(xml)
  }

  applyTheme(state: WorkbookThemeState): void {
    const path = 'xl/theme/theme1.xml'
    const xml = this.#metadataXml.get(path)
    if (xml === undefined) throw new Error('The workbook has no editable theme part.')
    this.#metadataXml.set(path, applyThemeState(xml, state))
    this.#dirtyPaths.add(path)
  }

  workbookProtection(): { readonly lockStructure: boolean; readonly hasPassword: boolean } | null {
    const xml = this.#metadataXml.get('xl/workbook.xml')!
    const attributes = /<workbookProtection\b([^>]*)\/?\s*>/.exec(xml)?.[1]
    if (attributes === undefined) return null
    return {
      lockStructure:
        xmlAttribute(attributes, 'lockStructure') === '1' ||
        xmlAttribute(attributes, 'lockStructure') === 'true',
      hasPassword:
        xmlAttribute(attributes, 'workbookPassword') !== undefined ||
        xmlAttribute(attributes, 'workbookHashValue') !== undefined,
    }
  }

  date1904(): boolean {
    const xml = this.#metadataXml.get('xl/workbook.xml')!
    const attributes = /<workbookPr\b([^>]*)\/?\s*>/.exec(xml)?.[1]
    if (attributes === undefined) return false
    const value = xmlAttribute(attributes, 'date1904')
    return value === '1' || value === 'true'
  }

  setWorkbookProtection(lockStructure: boolean): void {
    const path = 'xl/workbook.xml'
    this.#metadataXml.set(
      path,
      applyWorkbookProtectionToXml(this.#metadataXml.get(path)!, lockStructure),
    )
    this.#dirtyPaths.add(path)
  }

  pageSetup(sheetName: string): SheetPageSetupState {
    const sheet = this.#sheet(sheetName)
    return readPageSetupState(
      this.#sheetXml.get(sheet.path)!,
      this.#metadataXml.get('xl/workbook.xml')!,
      sheet.name,
      this.sheets.indexOf(sheet),
    )
  }

  pagePrintSettings(sheetName: string): WorkbookPagePrintSettings {
    const xml = this.#sheetXml.get(this.#sheet(sheetName).path)!
    const orientation = elementAttribute(xml, 'pageSetup', 'orientation')
    const margins = {
      left: numberAttribute(xml, 'pageMargins', 'left'),
      right: numberAttribute(xml, 'pageMargins', 'right'),
      top: numberAttribute(xml, 'pageMargins', 'top'),
      bottom: numberAttribute(xml, 'pageMargins', 'bottom'),
      header: numberAttribute(xml, 'pageMargins', 'header'),
      footer: numberAttribute(xml, 'pageMargins', 'footer'),
    }
    const completeMargins = Object.values(margins).every((value) => value !== undefined)
    const header = (name: string): string | undefined => textContent(xml, name)
    return {
      ...(orientation === 'portrait' || orientation === 'landscape' ? { orientation } : {}),
      ...(integerAttribute(xml, 'pageSetup', 'paperSize') === undefined
        ? {}
        : { paperSize: integerAttribute(xml, 'pageSetup', 'paperSize') }),
      ...(integerAttribute(xml, 'pageSetup', 'scale') === undefined
        ? {}
        : { scale: integerAttribute(xml, 'pageSetup', 'scale') }),
      ...(integerAttribute(xml, 'pageSetup', 'fitToWidth') === undefined
        ? {}
        : { fitToWidth: integerAttribute(xml, 'pageSetup', 'fitToWidth') }),
      ...(integerAttribute(xml, 'pageSetup', 'fitToHeight') === undefined
        ? {}
        : { fitToHeight: integerAttribute(xml, 'pageSetup', 'fitToHeight') }),
      ...(trueAttribute(xml, 'pageSetUpPr', 'fitToPage') === undefined
        ? {}
        : { fitToPage: trueAttribute(xml, 'pageSetUpPr', 'fitToPage') }),
      ...(completeMargins
        ? {
            margins: margins as {
              left: number
              right: number
              top: number
              bottom: number
              header: number
              footer: number
            },
          }
        : {}),
      ...(trueAttribute(xml, 'printOptions', 'gridLines') === undefined
        ? {}
        : { printGridlines: trueAttribute(xml, 'printOptions', 'gridLines') }),
      ...(trueAttribute(xml, 'printOptions', 'headings') === undefined
        ? {}
        : { printHeadings: trueAttribute(xml, 'printOptions', 'headings') }),
      ...(header('oddHeader') === undefined ? {} : { oddHeader: header('oddHeader') }),
      ...(header('oddFooter') === undefined ? {} : { oddFooter: header('oddFooter') }),
      ...(trueAttribute(xml, 'headerFooter', 'differentOddEven') === undefined
        ? {}
        : { differentOddEven: trueAttribute(xml, 'headerFooter', 'differentOddEven') }),
      ...(trueAttribute(xml, 'headerFooter', 'differentFirst') === undefined
        ? {}
        : { differentFirst: trueAttribute(xml, 'headerFooter', 'differentFirst') }),
      ...(trueAttribute(xml, 'headerFooter', 'scaleWithDoc') === false
        ? { headerFooterFixedSize: true }
        : {}),
      ...(header('evenHeader') === undefined ? {} : { evenHeader: header('evenHeader') }),
      ...(header('evenFooter') === undefined ? {} : { evenFooter: header('evenFooter') }),
      ...(header('firstHeader') === undefined ? {} : { firstHeader: header('firstHeader') }),
      ...(header('firstFooter') === undefined ? {} : { firstFooter: header('firstFooter') }),
    }
  }

  applyPageSetup(
    sheetName: string,
    patch: Omit<SheetPageSetupState, 'sheetName'>,
  ): SheetPageSetupState {
    const sheet = this.#sheet(sheetName)
    this.#sheetXml.set(
      sheet.path,
      applyPageSetupState(this.#sheetXml.get(sheet.path)!, { sheetName, ...patch }),
    )
    this.#dirtyPaths.add(sheet.path)
    if (patch.printArea !== undefined || patch.printTitles !== undefined) {
      const workbookPath = 'xl/workbook.xml'
      this.#metadataXml.set(
        workbookPath,
        applyPrintAreas(this.#metadataXml.get(workbookPath)!, [
          {
            sheetName,
            printArea: patch.printArea,
            printTitles: patch.printTitles,
          },
        ]),
      )
      this.#dirtyPaths.add(workbookPath)
    }
    return this.pageSetup(sheetName)
  }

  applyFilterState(sheetName: string, state: Omit<SheetFilterState, 'sheetName'>): void {
    const sheet = this.#sheet(sheetName)
    const xml = applyFilterStateToXml(this.#sheetXml.get(sheet.path)!, {
      sheetName,
      ...state,
    })
    this.#sheetXml.set(sheet.path, xml)
    const index = this.sheets.indexOf(sheet)
    this.sheets[index] = parseSheet(
      xml,
      sheet.name,
      sheet.path,
      this.#sharedStrings,
      this.#metadataXml.get(worksheetRelationshipsPath(sheet.path)) ?? null,
      sheet.tables,
      sheet.comments,
      { pivotRanges: sheet.pivotRanges, pivotTables: sheet.pivotTables },
      sheet.hidden,
    )
    this.#dirtyPaths.add(sheet.path)
  }

  setSheetProtection(sheetName: string, protected_: boolean): void {
    const sheet = this.#sheet(sheetName)
    const xml = applySheetProtectionToXml(this.#sheetXml.get(sheet.path)!, protected_)
    this.#sheetXml.set(sheet.path, xml)
    const index = this.sheets.indexOf(sheet)
    this.sheets[index] = parseSheet(
      xml,
      sheet.name,
      sheet.path,
      this.#sharedStrings,
      this.#metadataXml.get(worksheetRelationshipsPath(sheet.path)) ?? null,
      sheet.tables,
      sheet.comments,
      { pivotRanges: sheet.pivotRanges, pivotTables: sheet.pivotTables },
      sheet.hidden,
    )
    this.#dirtyPaths.add(sheet.path)
  }

  protectedRanges(sheetName: string): Array<ProtectedRangeState & { readonly hasPassword: boolean }> {
    return this.#sheet(sheetName).protectedRanges.map((range) => ({ ...range }))
  }

  setProtectedRanges(sheetName: string, ranges: readonly ProtectedRangeState[]): void {
    const sheet = this.#sheet(sheetName)
    const xml = applyProtectedRangesToXml(this.#sheetXml.get(sheet.path)!, ranges)
    this.#sheetXml.set(sheet.path, xml)
    const index = this.sheets.indexOf(sheet)
    this.sheets[index] = parseSheet(
      xml,
      sheet.name,
      sheet.path,
      this.#sharedStrings,
      this.#metadataXml.get(worksheetRelationshipsPath(sheet.path)) ?? null,
      sheet.tables,
      sheet.comments,
      { pivotRanges: sheet.pivotRanges, pivotTables: sheet.pivotTables },
      sheet.hidden,
    )
    this.#dirtyPaths.add(sheet.path)
  }

  applySparklines(sheetName: string, additions: readonly SparklineGroupAdd[]): void {
    if (additions.length === 0) return
    const sheet = this.#sheet(sheetName)
    const xml = applySparklineAdditions(this.#sheetXml.get(sheet.path)!, additions)
    this.#sheetXml.set(sheet.path, xml)
    const index = this.sheets.indexOf(sheet)
    this.sheets[index] = parseSheet(
      xml,
      sheet.name,
      sheet.path,
      this.#sharedStrings,
      this.#metadataXml.get(worksheetRelationshipsPath(sheet.path)) ?? null,
      sheet.tables,
      sheet.comments,
      { pivotRanges: sheet.pivotRanges, pivotTables: sheet.pivotTables },
      sheet.hidden,
    )
    this.#dirtyPaths.add(sheet.path)
  }

  applyDataValidations(sheetName: string, rules: readonly DvWireRule[]): void {
    const sheet = this.#sheet(sheetName)
    const xml = applyDvRules(this.#sheetXml.get(sheet.path)!, rules)
    this.#sheetXml.set(sheet.path, xml)
    const index = this.sheets.indexOf(sheet)
    this.sheets[index] = parseSheet(
      xml,
      sheet.name,
      sheet.path,
      this.#sharedStrings,
      this.#metadataXml.get(worksheetRelationshipsPath(sheet.path)) ?? null,
      sheet.tables,
      sheet.comments,
      { pivotRanges: sheet.pivotRanges, pivotTables: sheet.pivotTables },
      sheet.hidden,
    )
    this.#dirtyPaths.add(sheet.path)
  }

  applyConditionalFormats(sheetName: string, rules: readonly CfWireRule[]): void {
    if (!this.#stylesheet) {
      throw new Error(
        'The workbook stylesheet is missing — conditional formatting cannot be saved.',
      )
    }
    const sheet = this.#sheet(sheetName)
    const xml = applyCfRules(this.#sheetXml.get(sheet.path)!, rules, this.#stylesheet)
    this.#sheetXml.set(sheet.path, xml)
    this.#dirtyPaths.add(sheet.path)
    if (this.#stylesheet.changed) this.#dirtyPaths.add('xl/styles.xml')
  }

  applyHyperlinks(sheetName: string, edits: readonly HyperlinkEdit[]): void {
    const sheet = this.#sheet(sheetName)
    const relsPath = worksheetRelationshipsPath(sheet.path)
    const patch = applyHyperlinkEdits(
      this.#sheetXml.get(sheet.path)!,
      this.#metadataXml.get(relsPath) ?? null,
      edits,
    )
    const worksheetXml = patch.relsChanged
      ? ensureRelationshipNamespace(patch.worksheetXml)
      : patch.worksheetXml
    this.#sheetXml.set(sheet.path, worksheetXml)
    this.#dirtyPaths.add(sheet.path)
    if (patch.relsXml !== null) {
      this.#metadataXml.set(relsPath, patch.relsXml)
      if (patch.relsChanged) this.#dirtyPaths.add(relsPath)
    }
    const index = this.sheets.indexOf(sheet)
    this.sheets[index] = parseSheet(
      worksheetXml,
      sheet.name,
      sheet.path,
      this.#sharedStrings,
      patch.relsXml,
      sheet.tables,
      sheet.comments,
      { pivotRanges: sheet.pivotRanges, pivotTables: sheet.pivotTables },
      sheet.hidden,
    )
  }

  async applyTables(
    additions: readonly (Omit<TableAddition, 'worksheetPath'> & { readonly sheetName: string })[],
  ): Promise<void> {
    if (additions.length === 0) return
    const overlayText = new Map<string, string>()
    const overlayBinary = new Map<string, Uint8Array>()
    const removed = new Set<string>()
    const paths = async (): Promise<readonly string[]> =>
      [
        ...new Set([
          ...Object.values(this.#zip.files)
            .filter((entry) => !entry.dir)
            .map((entry) => entry.name),
          ...this.#sheetXml.keys(),
          ...this.#metadataXml.keys(),
          ...overlayText.keys(),
          ...overlayBinary.keys(),
        ]),
      ].filter((path) => !removed.has(path))
    const pkg: MutablePackage = {
      paths,
      has: async (path) =>
        !removed.has(path) &&
        (overlayText.has(path) ||
          overlayBinary.has(path) ||
          this.#sheetXml.has(path) ||
          this.#metadataXml.has(path) ||
          this.#zip.file(path) !== null),
      readText: async (path) => {
        if (removed.has(path)) throw new Error(`Workbook is missing ${path}.`)
        const overlay = overlayText.get(path)
        if (overlay !== undefined) return overlay
        const sheet = this.#sheetXml.get(path)
        if (sheet !== undefined) return sheet
        const metadata = this.#metadataXml.get(path)
        if (metadata !== undefined) return metadata
        return zipText(this.#zip, path)
      },
      write: (path, content) => {
        removed.delete(path)
        overlayText.set(path, content)
      },
      add: (path, content) => {
        removed.delete(path)
        overlayText.set(path, content)
      },
      addBinary: (path, bytes) => {
        removed.delete(path)
        overlayBinary.set(path, bytes)
      },
      remove: (path) => {
        overlayText.delete(path)
        overlayBinary.delete(path)
        removed.add(path)
      },
    }
    const touchedEntries = new Set<string>()
    await applyTableAdditions(
      pkg,
      additions.map(({ sheetName, ...addition }) => ({
        ...addition,
        worksheetPath: this.#sheet(sheetName).path,
      })),
      touchedEntries,
    )

    for (const path of removed) {
      this.#sheetXml.delete(path)
      this.#metadataXml.delete(path)
      this.#dirtyPaths.delete(path)
      this.#removedPaths.add(path)
    }
    for (const [path, content] of overlayText) {
      if (this.#sheetXml.has(path)) this.#sheetXml.set(path, content)
      else this.#metadataXml.set(path, content)
      this.#removedPaths.delete(path)
      this.#dirtyPaths.add(path)
    }
    for (const [path, bytes] of overlayBinary) {
      this.#zip.file(path, bytes)
      this.#removedPaths.delete(path)
    }
    for (const sheet of [...this.sheets]) {
      const xml = this.#sheetXml.get(sheet.path)
      if (xml === undefined) continue
      const relsXml = this.#metadataXml.get(worksheetRelationshipsPath(sheet.path)) ?? null
      const tables = await readSheetTables(xml, sheet.path, relsXml, async (path) => {
        const metadata = this.#metadataXml.get(path)
        if (metadata !== undefined) return metadata
        return this.#zip.file(path) ? zipText(this.#zip, path) : null
      })
      const index = this.sheets.indexOf(sheet)
      this.sheets[index] = parseSheet(
        xml,
        sheet.name,
        sheet.path,
        this.#sharedStrings,
        relsXml,
        tables,
        sheet.comments,
        { pivotRanges: sheet.pivotRanges, pivotTables: sheet.pivotTables },
        sheet.hidden,
      )
    }
  }

  async applyVisuals(
    additions: readonly (Omit<VisualAddition, 'worksheetPath'> & { readonly sheetName: string })[],
  ): Promise<void> {
    if (additions.length === 0) return
    const overlayText = new Map<string, string>()
    const overlayBinary = new Map<string, Uint8Array>()
    const removed = new Set<string>()
    const paths = async (): Promise<readonly string[]> =>
      [
        ...new Set([
          ...Object.values(this.#zip.files)
            .filter((entry) => !entry.dir)
            .map((entry) => entry.name),
          ...this.#sheetXml.keys(),
          ...this.#metadataXml.keys(),
          ...overlayText.keys(),
          ...overlayBinary.keys(),
        ]),
      ].filter((path) => !removed.has(path))
    const pkg: MutablePackage = {
      paths,
      has: async (path) =>
        !removed.has(path) &&
        (overlayText.has(path) ||
          overlayBinary.has(path) ||
          this.#sheetXml.has(path) ||
          this.#metadataXml.has(path) ||
          this.#zip.file(path) !== null),
      readText: async (path) => {
        if (removed.has(path)) throw new Error(`Workbook is missing ${path}.`)
        const overlay = overlayText.get(path)
        if (overlay !== undefined) return overlay
        const sheet = this.#sheetXml.get(path)
        if (sheet !== undefined) return sheet
        const metadata = this.#metadataXml.get(path)
        if (metadata !== undefined) return metadata
        return zipText(this.#zip, path)
      },
      write: (path, content) => {
        removed.delete(path)
        overlayText.set(path, content)
      },
      add: (path, content) => {
        removed.delete(path)
        overlayText.set(path, content)
      },
      addBinary: (path, bytes) => {
        removed.delete(path)
        overlayBinary.set(path, bytes)
      },
      remove: (path) => {
        overlayText.delete(path)
        overlayBinary.delete(path)
        removed.add(path)
      },
    }
    await applyVisualAdditions(
      pkg,
      additions.map(({ sheetName, ...addition }) => ({
        ...addition,
        worksheetPath: this.#sheet(sheetName).path,
      })),
      new Set<string>(),
    )
    for (const path of removed) {
      this.#sheetXml.delete(path)
      this.#metadataXml.delete(path)
      this.#dirtyPaths.delete(path)
      this.#removedPaths.add(path)
    }
    for (const [path, content] of overlayText) {
      if (this.#sheetXml.has(path)) this.#sheetXml.set(path, content)
      else this.#metadataXml.set(path, content)
      this.#removedPaths.delete(path)
      this.#dirtyPaths.add(path)
    }
    for (const [path, bytes] of overlayBinary) {
      this.#zip.file(path, bytes)
      this.#removedPaths.delete(path)
    }
  }

  async applyChartEdits(edits: readonly WorkbookChartEdit[]): Promise<void> {
    for (const edit of edits) {
      const source = this.#metadataXml.get(edit.chartPath)
      const chartXml =
        source ?? (this.#zip.file(edit.chartPath) ? await zipText(this.#zip, edit.chartPath) : null)
      if (chartXml === null) throw new Error(`Workbook is missing ${edit.chartPath}.`)
      this.#metadataXml.set(edit.chartPath, applyChartEditToXml(chartXml, edit))
      this.#dirtyPaths.add(edit.chartPath)
    }
  }

  async applyVisualEdits(edits: readonly WorkbookVisualEdit[]): Promise<void> {
    if (edits.length === 0) return
    const overlayText = new Map<string, string>()
    const removed = new Set<string>()
    const pkg: MutablePackage = {
      paths: async () =>
        [
          ...new Set([
            ...Object.values(this.#zip.files)
              .filter((entry) => !entry.dir)
              .map((entry) => entry.name),
            ...this.#sheetXml.keys(),
            ...this.#metadataXml.keys(),
            ...overlayText.keys(),
          ]),
        ].filter((path) => !removed.has(path)),
      has: async (path) =>
        !removed.has(path) &&
        (overlayText.has(path) ||
          this.#sheetXml.has(path) ||
          this.#metadataXml.has(path) ||
          this.#zip.file(path) !== null),
      readText: async (path) => {
        if (removed.has(path)) throw new Error(`Workbook is missing ${path}.`)
        return (
          overlayText.get(path) ??
          this.#sheetXml.get(path) ??
          this.#metadataXml.get(path) ??
          (await zipText(this.#zip, path))
        )
      },
      write: (path, content) => {
        removed.delete(path)
        overlayText.set(path, content)
      },
      add: (path, content) => {
        removed.delete(path)
        overlayText.set(path, content)
      },
      addBinary: () => {
        throw new Error('Visual edits cannot add binary parts.')
      },
      remove: (path) => {
        overlayText.delete(path)
        removed.add(path)
      },
    }
    await applyVisualEditsToPackage(pkg, edits, new Set<string>())
    for (const path of removed) {
      this.#sheetXml.delete(path)
      this.#metadataXml.delete(path)
      this.#dirtyPaths.delete(path)
      this.#removedPaths.add(path)
    }
    for (const [path, content] of overlayText) {
      if (this.#sheetXml.has(path)) this.#sheetXml.set(path, content)
      else this.#metadataXml.set(path, content)
      this.#removedPaths.delete(path)
      this.#dirtyPaths.add(path)
    }
  }

  async applyNotes(sheetName: string, notes: readonly SheetNote[]): Promise<void> {
    const sheet = this.#sheet(sheetName)
    const overlayText = new Map<string, string>()
    const removed = new Set<string>()
    const paths = async (): Promise<readonly string[]> =>
      [
        ...new Set([
          ...Object.values(this.#zip.files)
            .filter((entry) => !entry.dir)
            .map((entry) => entry.name),
          ...this.#sheetXml.keys(),
          ...this.#metadataXml.keys(),
          ...overlayText.keys(),
        ]),
      ].filter((path) => !removed.has(path))
    const pkg = {
      paths,
      has: async (path: string) =>
        !removed.has(path) &&
        (overlayText.has(path) ||
          this.#sheetXml.has(path) ||
          this.#metadataXml.has(path) ||
          this.#zip.file(path) !== null),
      readText: async (path: string) => {
        if (removed.has(path)) throw new Error(`Workbook is missing ${path}.`)
        const overlay = overlayText.get(path)
        if (overlay !== undefined) return overlay
        const worksheet = this.#sheetXml.get(path)
        if (worksheet !== undefined) return worksheet
        const metadata = this.#metadataXml.get(path)
        if (metadata !== undefined) return metadata
        return zipText(this.#zip, path)
      },
      write: (path: string, content: string) => {
        removed.delete(path)
        overlayText.set(path, content)
      },
      add: (path: string, content: string) => {
        removed.delete(path)
        overlayText.set(path, content)
      },
      remove: (path: string) => {
        overlayText.delete(path)
        removed.add(path)
      },
    }
    await applySheetNotes(pkg, sheet.path, notes, new Set<string>())
    for (const path of removed) {
      this.#sheetXml.delete(path)
      this.#metadataXml.delete(path)
      this.#dirtyPaths.delete(path)
      this.#removedPaths.add(path)
    }
    for (const [path, content] of overlayText) {
      if (this.#sheetXml.has(path)) this.#sheetXml.set(path, content)
      else this.#metadataXml.set(path, content)
      this.#removedPaths.delete(path)
      this.#dirtyPaths.add(path)
    }
    const index = this.sheets.indexOf(sheet)
    this.sheets[index] = parseSheet(
      this.#sheetXml.get(sheet.path)!,
      sheet.name,
      sheet.path,
      this.#sharedStrings,
      this.#metadataXml.get(worksheetRelationshipsPath(sheet.path)) ?? null,
      sheet.tables,
      notes.map((note) => ({ ...note })),
      { pivotRanges: sheet.pivotRanges, pivotTables: sheet.pivotTables },
      sheet.hidden,
    )
  }

  setCellValue(sheetName: string, address: string, value: CellScalar): void {
    const sheet = this.#sheet(sheetName)
    const normalized = address.toUpperCase()
    parseAddress(normalized)
    const previous = sheet.cells.get(normalized)
    const cell: BrowserCell = {
      address: normalized,
      value,
      ...(previous?.styleIndex !== undefined ? { styleIndex: previous.styleIndex } : {}),
    }
    sheet.cells.set(normalized, cell)
    this.#sheetXml.set(sheet.path, patchCell(this.#sheetXml.get(sheet.path)!, cell))
    this.#dirtyPaths.add(sheet.path)
  }

  setCellFormula(sheetName: string, address: string, formula: string, value: CellScalar): void {
    const sheet = this.#sheet(sheetName)
    const normalized = address.toUpperCase()
    parseAddress(normalized)
    const previous = sheet.cells.get(normalized)
    const cell: BrowserCell = {
      address: normalized,
      value,
      formula: formula.startsWith('=') ? formula.slice(1) : formula,
      ...(previous?.styleIndex !== undefined ? { styleIndex: previous.styleIndex } : {}),
    }
    sheet.cells.set(normalized, cell)
    this.#sheetXml.set(sheet.path, patchCell(this.#sheetXml.get(sheet.path)!, cell))
    this.#dirtyPaths.add(sheet.path)
  }

  setRangeValues(sheetName: string, range: string, values: CellScalar[][]): number {
    const [startText, endText = startText] = range.toUpperCase().split(':')
    if (!startText || !endText) throw new Error(`Invalid cell range: ${range}`)
    const start = parseAddress(startText)
    const end = parseAddress(endText)
    if (end.row < start.row || end.column < start.column)
      throw new Error(`Invalid cell range: ${range}`)
    const expectedRows = end.row - start.row + 1
    const expectedColumns = end.column - start.column + 1
    if (values.length !== expectedRows || values.some((row) => row.length !== expectedColumns)) {
      throw new Error(`Range ${range} requires a ${expectedRows}×${expectedColumns} value matrix.`)
    }
    let changed = 0
    for (let rowOffset = 0; rowOffset < expectedRows; rowOffset += 1) {
      for (let columnOffset = 0; columnOffset < expectedColumns; columnOffset += 1) {
        this.setCellValue(
          sheetName,
          `${columnLabel(start.column + columnOffset)}${start.row + rowOffset}`,
          values[rowOffset]![columnOffset]!,
        )
        changed += 1
      }
    }
    return changed
  }

  setRangeStyle(sheetName: string, range: string, style: WorkbookStyleEdit): number {
    if (!this.#stylesheet) {
      throw new Error('The workbook stylesheet is missing; formatting cannot be saved.')
    }
    const [startText, endText = startText] = range.toUpperCase().split(':')
    if (!startText || !endText) throw new Error(`Invalid cell range: ${range}`)
    const start = parseAddress(startText)
    const end = parseAddress(endText)
    if (end.row < start.row || end.column < start.column) {
      throw new Error(`Invalid cell range: ${range}`)
    }
    const sheet = this.#sheet(sheetName)
    let changed = 0
    for (let row = start.row; row <= end.row; row += 1) {
      for (let column = start.column; column <= end.column; column += 1) {
        const address = `${columnLabel(column)}${row}`
        const previous = sheet.cells.get(address) ?? { address, value: null }
        const cell: BrowserCell = {
          ...previous,
          styleIndex: this.#stylesheet.resolveStyle(previous.styleIndex ?? 0, style),
        }
        sheet.cells.set(address, cell)
        this.#sheetXml.set(sheet.path, patchCell(this.#sheetXml.get(sheet.path)!, cell))
        this.#dirtyPaths.add(sheet.path)
        changed += 1
      }
    }
    this.#dirtyPaths.add('xl/styles.xml')
    return changed
  }

  applyStructuralOperation(sheetName: string, operation: StructuralOp): void {
    this.#structureUndo.push(this.#captureStructure())
    this.#structureRedo.length = 0
    const sheet = this.#sheet(sheetName)
    const xml = applyStructuralOps(this.#sheetXml.get(sheet.path)!, [operation], sheetName)
    this.#sheetXml.set(sheet.path, xml)
    const relsPath = worksheetRelationshipsPath(sheet.path)
    const relsXml = this.#metadataXml.get(relsPath)
    if (relsXml) {
      const relationships = readRelationships(relsXml)
      const parts: Array<{ id: string; kind: 'drawing' | 'table' | 'vml' }> = []
      const drawingId = elementAttribute(xml, 'drawing', 'r:id')
      if (drawingId) parts.push({ id: drawingId, kind: 'drawing' })
      for (const match of xml.matchAll(/<tablePart\b([^>]*)\/?\s*>/g)) {
        const id = xmlAttribute(match[1] ?? '', 'r:id')
        if (id) parts.push({ id, kind: 'table' })
      }
      const legacyId = elementAttribute(xml, 'legacyDrawing', 'r:id')
      if (legacyId && xml.includes('<oleObjects')) parts.push({ id: legacyId, kind: 'vml' })
      for (const part of parts) {
        const relationship = relationships.find((entry) => entry.id === part.id)
        if (!relationship) {
          throw new Error(`${sheet.path} references missing ${part.kind} relationship ${part.id}.`)
        }
        const path = resolvePartTarget(sheet.path, relationship.target)
        const previous = this.#metadataXml.get(path)
        if (previous === undefined) {
          throw new Error(`${path} was not hydrated for structural editing.`)
        }
        const shifted =
          part.kind === 'drawing'
            ? shiftDrawingAnchors(previous, [operation])
            : part.kind === 'vml'
              ? shiftVmlObjectAnchors(previous, [operation])
              : shiftTablePart(previous, [operation])
        if (shifted === previous) continue
        this.#metadataXml.set(path, shifted)
        this.#dirtyPaths.add(path)
      }
    }
    const index = this.sheets.indexOf(sheet)
    this.sheets[index] = parseSheet(
      xml,
      sheet.name,
      sheet.path,
      this.#sharedStrings,
      this.#metadataXml.get(worksheetRelationshipsPath(sheet.path)) ?? null,
      sheet.tables,
      sheet.comments,
      { pivotRanges: sheet.pivotRanges, pivotTables: sheet.pivotTables },
      sheet.hidden,
    )
    this.#dirtyPaths.add(sheet.path)
    for (const candidate of this.sheets) {
      if (candidate.path === sheet.path) continue
      const previous = this.#sheetXml.get(candidate.path)!
      const shifted = shiftCrossSheetFormulas(previous, sheetName, [operation])
      if (shifted === previous) continue
      this.#sheetXml.set(candidate.path, shifted)
      const candidateIndex = this.sheets.indexOf(candidate)
      this.sheets[candidateIndex] = parseSheet(
        shifted,
        candidate.name,
        candidate.path,
        this.#sharedStrings,
        this.#metadataXml.get(worksheetRelationshipsPath(candidate.path)) ?? null,
        candidate.tables,
        candidate.comments,
        { pivotRanges: candidate.pivotRanges, pivotTables: candidate.pivotTables },
        candidate.hidden,
      )
      this.#dirtyPaths.add(candidate.path)
    }
    for (const [path, previous] of this.#metadataXml) {
      if (!path.startsWith('xl/charts/') || !path.endsWith('.xml')) continue
      const shifted = shiftChartReferences(previous, sheetName, [operation])
      if (shifted === previous) continue
      this.#metadataXml.set(path, shifted)
      this.#dirtyPaths.add(path)
    }
    const workbookPath = 'xl/workbook.xml'
    const workbookXml = this.#metadataXml.get(workbookPath)!
    const shiftedWorkbook = shiftDefinedNames(workbookXml, sheetName, [operation])
    if (shiftedWorkbook !== workbookXml) {
      this.#metadataXml.set(workbookPath, shiftedWorkbook)
      this.#dirtyPaths.add(workbookPath)
    }
  }

  addSheet(name: string): BrowserSheet {
    validateSheetName(name)
    if (this.sheets.some((sheet) => sheet.name.toLowerCase() === name.toLowerCase())) {
      throw new Error(`A worksheet named "${name}" already exists.`)
    }
    this.#structureUndo.push(this.#captureStructure())
    this.#structureRedo.length = 0
    const workbookPath = 'xl/workbook.xml'
    const relationshipsPath = 'xl/_rels/workbook.xml.rels'
    const contentTypesPath = '[Content_Types].xml'
    const workbookXml = this.#metadataXml.get(workbookPath)!
    const relationshipsXml = this.#metadataXml.get(relationshipsPath)!
    const usedPartNumbers = this.sheets.flatMap((sheet) => {
      const match = /\/sheet([0-9]+)\.xml$/.exec(sheet.path)
      return match?.[1] ? [Number(match[1])] : []
    })
    const partNumber = Math.max(0, ...usedPartNumbers) + 1
    const path = `xl/worksheets/sheet${partNumber}.xml`
    const allocation = {
      name,
      path,
      sheetId: maxSheetIdInWorkbook(workbookXml) + 1,
      relationshipId: `rId${maxRelationshipId(relationshipsXml) + 1}`,
    }
    const order = [...this.sheets.map((sheet) => sheet.name), name]
    this.#metadataXml.set(
      workbookPath,
      applySheetPlanToWorkbookXml(
        workbookXml,
        { renames: [], additions: [{ name }], removals: [], order },
        [allocation],
      ),
    )
    this.#metadataXml.set(
      relationshipsPath,
      addWorksheetRelationship(
        relationshipsXml,
        allocation.relationshipId,
        `worksheets/sheet${partNumber}.xml`,
      ),
    )
    this.#metadataXml.set(
      contentTypesPath,
      addWorksheetOverride(this.#metadataXml.get(contentTypesPath)!, path),
    )
    const xml = buildWorksheetPartXml()
    this.#sheetXml.set(path, xml)
    const sheet = parseSheet(xml, name, path, this.#sharedStrings)
    this.sheets.push(sheet)
    this.#dirtyPaths.add(path)
    this.#dirtyPaths.add(workbookPath)
    this.#dirtyPaths.add(relationshipsPath)
    this.#dirtyPaths.add(contentTypesPath)
    return sheet
  }

  duplicateSheet(sourceSheetName: string, name: string): BrowserSheet {
    validateSheetName(name)
    if (this.sheets.some((sheet) => sheet.name.toLowerCase() === name.toLowerCase())) {
      throw new Error(`A worksheet named "${name}" already exists.`)
    }
    const source = this.#sheet(sourceSheetName)
    const workbookPath = 'xl/workbook.xml'
    const relationshipsPath = 'xl/_rels/workbook.xml.rels'
    const contentTypesPath = '[Content_Types].xml'
    const workbookXml = this.#metadataXml.get(workbookPath)!
    assertNoSheetScopedDefinedNames(workbookXml, sourceSheetName)

    const sourceRelsPath = worksheetRelationshipsPath(source.path)
    const sourceRelsXml = this.#metadataXml.get(sourceRelsPath) ?? null
    const preparedRels =
      sourceRelsXml === null ? null : prepareClonedSheetRels(sourceRelsXml, sourceSheetName)
    let xml = sanitizeClonedWorksheetXml(this.#sheetXml.get(source.path)!)
    if (preparedRels?.droppedPrinterSettings) xml = stripPageSetupRelIds(xml)

    this.#structureUndo.push(this.#captureStructure())
    this.#structureRedo.length = 0
    const relationshipsXml = this.#metadataXml.get(relationshipsPath)!
    const usedPartNumbers = this.sheets.flatMap((sheet) => {
      const match = /\/sheet([0-9]+)\.xml$/.exec(sheet.path)
      return match?.[1] ? [Number(match[1])] : []
    })
    const partNumber = Math.max(0, ...usedPartNumbers) + 1
    const path = `xl/worksheets/sheet${partNumber}.xml`
    const allocation = {
      name,
      path,
      sheetId: maxSheetIdInWorkbook(workbookXml) + 1,
      relationshipId: `rId${maxRelationshipId(relationshipsXml) + 1}`,
    }
    this.#metadataXml.set(
      workbookPath,
      applySheetPlanToWorkbookXml(
        workbookXml,
        {
          renames: [],
          additions: [{ name, sourceSheetName }],
          removals: [],
          order: [...this.sheets.map((sheet) => sheet.name), name],
          orderChanged: true,
        },
        [allocation],
      ),
    )
    this.#metadataXml.set(
      relationshipsPath,
      addWorksheetRelationship(
        relationshipsXml,
        allocation.relationshipId,
        `worksheets/sheet${partNumber}.xml`,
      ),
    )
    this.#metadataXml.set(
      contentTypesPath,
      addWorksheetOverride(this.#metadataXml.get(contentTypesPath)!, path),
    )
    const relsPath = worksheetRelationshipsPath(path)
    if (preparedRels?.relsXml) {
      this.#metadataXml.set(relsPath, preparedRels.relsXml)
      this.#dirtyPaths.add(relsPath)
    }
    this.#sheetXml.set(path, xml)
    const sheet = parseSheet(xml, name, path, this.#sharedStrings, preparedRels?.relsXml ?? null)
    this.sheets.push(sheet)
    this.#dirtyPaths.add(path)
    this.#dirtyPaths.add(workbookPath)
    this.#dirtyPaths.add(relationshipsPath)
    this.#dirtyPaths.add(contentTypesPath)
    this.#dropCalculationChain()
    return sheet
  }

  setSheetHidden(sheetName: string, hidden: boolean): void {
    const sheet = this.#sheet(sheetName)
    if (sheet.hidden === hidden) return
    this.#structureUndo.push(this.#captureStructure())
    this.#structureRedo.length = 0
    const workbookPath = 'xl/workbook.xml'
    this.#metadataXml.set(
      workbookPath,
      applySheetPlanToWorkbookXml(
        this.#metadataXml.get(workbookPath)!,
        {
          renames: [],
          additions: [],
          removals: [],
          order: this.sheets.map((candidate) => candidate.name),
          hiddenChanges: [{ sheetName, hidden }],
        },
        [],
      ),
    )
    const index = this.sheets.indexOf(sheet)
    this.sheets[index] = { ...sheet, hidden }
    this.#dirtyPaths.add(workbookPath)
  }

  renameSheet(sheetName: string, newName: string): void {
    validateSheetName(newName)
    const sheet = this.#sheet(sheetName)
    if (
      this.sheets.some(
        (candidate) =>
          candidate !== sheet && candidate.name.toLowerCase() === newName.toLowerCase(),
      )
    ) {
      throw new Error(`A worksheet named "${newName}" already exists.`)
    }
    if (sheet.name === newName) return
    this.#structureUndo.push(this.#captureStructure())
    this.#structureRedo.length = 0
    for (const candidate of this.sheets) {
      const previous = this.#sheetXml.get(candidate.path)!
      const renamed = renameSheetReferencesInWorksheet(previous, sheetName, newName)
      if (renamed !== previous) {
        this.#sheetXml.set(candidate.path, renamed)
        this.#dirtyPaths.add(candidate.path)
      }
    }
    const workbookPath = 'xl/workbook.xml'
    const originalWorkbook = this.#metadataXml.get(workbookPath)!
    const referencesRenamed = renameSheetReferencesInDefinedNames(
      originalWorkbook,
      sheetName,
      newName,
    )
    const order = this.sheets.map((candidate) => (candidate === sheet ? newName : candidate.name))
    this.#metadataXml.set(
      workbookPath,
      applySheetPlanToWorkbookXml(
        referencesRenamed,
        {
          renames: [{ sheetName, newName }],
          additions: [],
          removals: [],
          order,
        },
        [],
      ),
    )
    const index = this.sheets.indexOf(sheet)
    this.sheets[index] = parseSheet(
      this.#sheetXml.get(sheet.path)!,
      newName,
      sheet.path,
      this.#sharedStrings,
      this.#metadataXml.get(worksheetRelationshipsPath(sheet.path)) ?? null,
      sheet.tables,
      sheet.comments,
      { pivotRanges: sheet.pivotRanges, pivotTables: sheet.pivotTables },
      sheet.hidden,
    )
    this.#dirtyPaths.add(workbookPath)
  }

  deleteSheet(sheetName: string): void {
    if (this.sheets.length <= 1) throw new Error('A workbook needs at least one worksheet.')
    const sheet = this.#sheet(sheetName)
    const workbookPath = 'xl/workbook.xml'
    const relationshipsPath = 'xl/_rels/workbook.xml.rels'
    const contentTypesPath = '[Content_Types].xml'
    const workbookXml = this.#metadataXml.get(workbookPath)!
    const sheetIndex = this.sheets.indexOf(sheet)
    for (const candidate of this.sheets) {
      if (
        candidate !== sheet &&
        worksheetReferencesSheet(this.#sheetXml.get(candidate.path)!, sheetName)
      ) {
        throw new Error(`Another worksheet references "${sheetName}".`)
      }
    }
    if (definedNamesReferenceSheet(workbookXml, sheetName, new Set([sheetIndex]))) {
      throw new Error(`A workbook defined name references "${sheetName}".`)
    }
    const relsPath = sheet.path.replace(/^(xl\/worksheets\/)([^/]+)$/, '$1_rels/$2.rels')
    if (this.#zip.file(relsPath)) {
      throw new Error(`Worksheet "${sheetName}" owns related package parts and cannot be deleted.`)
    }
    const element = parseSheetElements(workbookXml).find(
      (candidate) => candidate.name === sheetName,
    )
    const relationshipId = element?.relationshipId
    if (!relationshipId) throw new Error(`Worksheet "${sheetName}" has no workbook relationship.`)

    this.#structureUndo.push(this.#captureStructure())
    this.#structureRedo.length = 0
    const order = this.sheets
      .filter((candidate) => candidate !== sheet)
      .map((candidate) => candidate.name)
    this.#metadataXml.set(
      workbookPath,
      applySheetPlanToWorkbookXml(
        workbookXml,
        { renames: [], additions: [], removals: [sheetName], order, orderChanged: true },
        [],
      ),
    )
    this.#metadataXml.set(
      relationshipsPath,
      removeRelationshipById(this.#metadataXml.get(relationshipsPath)!, relationshipId),
    )
    this.#metadataXml.set(
      contentTypesPath,
      removePartOverride(this.#metadataXml.get(contentTypesPath)!, sheet.path),
    )
    this.sheets.splice(sheetIndex, 1)
    this.#sheetXml.delete(sheet.path)
    this.#dirtyPaths.delete(sheet.path)
    this.#removedPaths.add(sheet.path)
    this.#dirtyPaths.add(workbookPath)
    this.#dirtyPaths.add(relationshipsPath)
    this.#dirtyPaths.add(contentTypesPath)
  }

  moveSheet(sheetName: string, position: number): void {
    if (!Number.isInteger(position) || position < 1 || position > this.sheets.length) {
      throw new Error(`Sheet position must be between 1 and ${this.sheets.length}.`)
    }
    const sheet = this.#sheet(sheetName)
    const current = this.sheets.indexOf(sheet)
    const target = position - 1
    if (current === target) return
    this.#structureUndo.push(this.#captureStructure())
    this.#structureRedo.length = 0
    this.sheets.splice(current, 1)
    this.sheets.splice(target, 0, sheet)
    const workbookPath = 'xl/workbook.xml'
    this.#metadataXml.set(
      workbookPath,
      applySheetPlanToWorkbookXml(
        this.#metadataXml.get(workbookPath)!,
        {
          renames: [],
          additions: [],
          removals: [],
          order: this.sheets.map((candidate) => candidate.name),
          orderChanged: true,
        },
        [],
      ),
    )
    this.#dirtyPaths.add(workbookPath)
  }

  undoStructuralOperation(): boolean {
    const previous = this.#structureUndo.pop()
    if (!previous) return false
    this.#structureRedo.push(this.#captureStructure())
    this.#restoreStructure(previous)
    return true
  }

  redoStructuralOperation(): boolean {
    const next = this.#structureRedo.pop()
    if (!next) return false
    this.#structureUndo.push(this.#captureStructure())
    this.#restoreStructure(next)
    return true
  }

  async checkpoint(): Promise<Uint8Array> {
    for (const path of this.#removedPaths) this.#zip.remove(path)
    for (const path of this.#dirtyPaths) {
      if (path === 'xl/styles.xml' && this.#stylesheet) {
        this.#zip.file(path, this.#stylesheet.serialize())
      } else if (this.#metadataXml.has(path)) {
        this.#zip.file(path, this.#metadataXml.get(path)!)
      } else {
        this.#zip.file(path, this.#sheetXml.get(path)!)
      }
    }
    return this.#zip.generateAsync({
      type: 'uint8array',
      compression: 'DEFLATE',
      compressionOptions: { level: 6 },
    })
  }

  async save(): Promise<Uint8Array> {
    const bytes = await this.checkpoint()
    this.#dirtyPaths.clear()
    this.#removedPaths.clear()
    return bytes
  }

  #sheet(name: string): BrowserSheet {
    const sheet = this.sheets.find((candidate) => candidate.name === name)
    if (!sheet) throw new Error(`Unknown worksheet: ${name}`)
    return sheet
  }

  #dropCalculationChain(): void {
    const path = 'xl/calcChain.xml'
    if (!this.#zip.file(path) && !this.#metadataXml.has(path)) return
    this.#metadataXml.delete(path)
    this.#dirtyPaths.delete(path)
    this.#removedPaths.add(path)
    const contentTypesPath = '[Content_Types].xml'
    this.#metadataXml.set(
      contentTypesPath,
      removePartOverride(this.#metadataXml.get(contentTypesPath)!, path),
    )
    this.#dirtyPaths.add(contentTypesPath)
    const relationshipsPath = 'xl/_rels/workbook.xml.rels'
    const relationshipsXml = this.#metadataXml.get(relationshipsPath)!
    const nextRelationships = relationshipsXml.replace(
      /<Relationship\b[^>]*Target="calcChain\.xml"[^>]*\/>/,
      '',
    )
    if (nextRelationships !== relationshipsXml) {
      this.#metadataXml.set(relationshipsPath, nextRelationships)
      this.#dirtyPaths.add(relationshipsPath)
    }
  }

  async #readPackageText(path: string): Promise<string> {
    const text = this.#sheetXml.get(path) ?? this.#metadataXml.get(path)
    if (text !== undefined) return text
    if (this.#removedPaths.has(path)) throw new Error(`Workbook is missing ${path}.`)
    return zipText(this.#zip, path)
  }

  #captureStructure(): StructureSnapshot {
    return {
      sheetXml: new Map(this.#sheetXml),
      metadataXml: new Map(this.#metadataXml),
      sheets: this.sheets.map((sheet) => ({
        ...sheet,
        cells: new Map(
          [...sheet.cells].map(([address, cell]) => [
            address,
            {
              ...cell,
              ...(cell.rich === undefined
                ? {}
                : { rich: cell.rich.map((run) => ({ ...run })) }),
            },
          ]),
        ),
        rows: sheet.rows.map((row) => ({ ...row })),
        columnWidths: sheet.columnWidths.map((column) => ({ ...column })),
        dataValidations: sheet.dataValidations.map((rule) => ({
          ...rule,
          ranges: rule.ranges.map((range) => ({ ...range })),
          formulas: [...rule.formulas],
        })),
        merges: [...sheet.merges],
        hyperlinks: new Map(sheet.hyperlinks),
        tables: sheet.tables.map((table) => ({ ...table, range: { ...table.range } })),
        sparklines: sheet.sparklines.map((group) => ({
          ...group,
          cells: group.cells.map((cell) => ({ ...cell })),
        })),
      })),
      dirtyPaths: new Set(this.#dirtyPaths),
      removedPaths: new Set(this.#removedPaths),
    }
  }

  #restoreStructure(snapshot: StructureSnapshot): void {
    this.#sheetXml.clear()
    for (const [path, xml] of snapshot.sheetXml) this.#sheetXml.set(path, xml)
    this.#metadataXml.clear()
    for (const [path, xml] of snapshot.metadataXml) this.#metadataXml.set(path, xml)
    this.sheets.splice(0, this.sheets.length, ...snapshot.sheets)
    this.#dirtyPaths.clear()
    for (const path of snapshot.dirtyPaths) this.#dirtyPaths.add(path)
    this.#removedPaths.clear()
    for (const path of snapshot.removedPaths) this.#removedPaths.add(path)
  }
}

export async function openBrowserWorkbook(
  data: ArrayBuffer | Uint8Array,
  name: string,
): Promise<BrowserWorkbook> {
  const zip = await JSZip.loadAsync(data, { checkCRC32: true })
  const files = Object.values(zip.files).filter((entry) => !entry.dir)
  if (files.length > MAX_ENTRIES) throw new Error('Workbook contains too many ZIP entries.')
  let total = 0
  for (const file of files) {
    total += (await file.async('uint8array')).byteLength
    if (total > MAX_UNCOMPRESSED_BYTES)
      throw new Error('Workbook exceeds the uncompressed size limit.')
    if (file.name.startsWith('/') || file.name.split('/').includes('..')) {
      throw new Error('Workbook contains an unsafe ZIP path.')
    }
  }

  const workbookXml = await zipText(zip, 'xl/workbook.xml')
  const relsXml = await zipText(zip, 'xl/_rels/workbook.xml.rels')
  const contentTypesXml = await zipText(zip, '[Content_Types].xml')
  const metadataXml = new Map([
    ['xl/workbook.xml', workbookXml],
    ['xl/_rels/workbook.xml.rels', relsXml],
    ['[Content_Types].xml', contentTypesXml],
  ])
  const themeRelationship = [...relsXml.matchAll(/<Relationship\b([^>]*)\/?\s*>/g)].find(
    (match) => /\/theme$/.test(xmlAttribute(match[1] ?? '', 'Type') ?? ''),
  )
  const themeTarget = themeRelationship ? xmlAttribute(themeRelationship[1] ?? '', 'Target') : undefined
  let themeColors: readonly string[] | undefined
  if (themeTarget) {
    const themePath = `xl/${themeTarget.replace(/^\/?xl\//, '').replace(/^\.\//, '')}`
    if (zip.file(themePath)) {
      const themeXml = await zipText(zip, themePath)
      metadataXml.set(themePath, themeXml)
      themeColors = readThemeState(themeXml)?.colors?.values
    }
  }
  const sharedStrings = readSharedStrings(
    zip.file('xl/sharedStrings.xml') ? await zipText(zip, 'xl/sharedStrings.xml') : null,
  )
  const sheets: BrowserSheet[] = []
  const visuals: BrowserVisual[] = []
  const sheetXml = new Map<string, string>()
  for (const match of workbookXml.matchAll(/<sheet\b([^>]*)\/>/g)) {
    const attributes = match[1] ?? ''
    const sheetName = xmlAttribute(attributes, 'name')
    const relationshipId = xmlAttribute(attributes, 'r:id')
    const state = xmlAttribute(attributes, 'state')
    if (!sheetName || !relationshipId) continue
    const target = relationshipTarget(relsXml, relationshipId).replace(/^\/?xl\//, '')
    const path = `xl/${target.replace(/^\.\//, '')}`
    const xml = await zipText(zip, path)
    const sheetRelsPath = worksheetRelationshipsPath(path)
    const sheetRelsXml = zip.file(sheetRelsPath) ? await zipText(zip, sheetRelsPath) : null
    const readMetadataPart = async (partPath: string): Promise<string | null> => {
      if (!zip.file(partPath)) return null
      const content = await zipText(zip, partPath)
      metadataXml.set(partPath, content)
      return content
    }
    const tables = await readSheetTables(
      xml,
      path,
      sheetRelsXml,
      readMetadataPart,
      themeColors,
    )
    const comments = await readSheetNotes(path, sheetRelsXml, readMetadataPart)
    const pivots = await readSheetPivots(
      path,
      sheetRelsXml,
      readMetadataPart,
      themeColors,
    )
    visuals.push(
      ...(await readSheetVisuals(xml, path, decodeXml(sheetName), sheetRelsXml, readMetadataPart)),
    )
    sheetXml.set(path, xml)
    if (sheetRelsXml !== null) metadataXml.set(sheetRelsPath, sheetRelsXml)
    sheets.push(
      parseSheet(
        xml,
        decodeXml(sheetName),
        path,
        sharedStrings,
        sheetRelsXml,
        tables,
        comments,
        pivots,
        state === 'hidden' || state === 'veryHidden',
      ),
    )
  }
  if (sheets.length === 0) throw new Error('Workbook contains no readable worksheets.')
  const stylesXml = zip.file('xl/styles.xml') ? await zipText(zip, 'xl/styles.xml') : null
  return new BrowserWorkbook(
    name,
    sheets,
    visuals,
    zip,
    sheetXml,
    stylesXml === null ? null : new StylesheetEditor(stylesXml),
    sharedStrings,
    metadataXml,
  )
}
