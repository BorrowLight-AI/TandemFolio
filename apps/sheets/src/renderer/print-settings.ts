// Modified by TandemFolio contributors: browser-workbook print metadata and structural mapping.
import { columnLabel, parseRange } from '../domain/cell-address'
import type { WorkbookPagePrintSettings } from '../shared/desktop-api'
import type { HeaderFooterParts, PageSetupJournalState, StructuralJournalOp } from './edit-journal'
import { fileRangeToScreenRange, fileToScreen } from './view-transform'

export interface PrintMargins {
  readonly left: number
  readonly right: number
  readonly top: number
  readonly bottom: number
  readonly header: number
  readonly footer: number
}

export interface HeaderFooterPair {
  readonly header: HeaderFooterParts | null
  readonly footer: HeaderFooterParts | null
}

export type HeaderFooterPictureSlot = NonNullable<
  WorkbookPagePrintSettings['headerFooterPictures']
>[number]
export type FilePagePrintSettings = WorkbookPagePrintSettings

export interface EffectivePageSetup {
  readonly orientation: 'portrait' | 'landscape'
  readonly paperSize: number
  readonly scale: number
  readonly fitToWidth: number
  readonly fitToHeight: number
  readonly fitToPage: boolean
  readonly margins: PrintMargins
  readonly printGridlines: boolean
  readonly printHeadings: boolean
  readonly printAreas: readonly string[]
  readonly printTitles: string | null
  readonly header: HeaderFooterParts | null
  readonly footer: HeaderFooterParts | null
  readonly firstPage: HeaderFooterPair | null
  readonly evenPages: HeaderFooterPair | null
  readonly headerFooterScaleWithDoc: boolean
  readonly headerFooterPictures: readonly HeaderFooterPictureSlot[]
}

const MARGIN_PRESETS: Record<'normal' | 'wide' | 'narrow', PrintMargins> = {
  normal: { left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 },
  wide: { left: 1, right: 1, top: 1, bottom: 1, header: 0.5, footer: 0.5 },
  narrow: { left: 0.25, right: 0.25, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 },
}

const MAX_MARGIN_INCHES = 3

export interface FilePrintNames {
  readonly printArea?: string | undefined
  readonly printTitles?: string | undefined
}

function mapAreasToScreen(areas: readonly string[], ops: readonly StructuralJournalOp[]): string[] {
  if (ops.length === 0) return [...areas]
  const mapped: string[] = []
  for (const area of areas) {
    let bounds
    try {
      bounds = parseRange(area)
    } catch {
      continue
    }
    const screen = fileRangeToScreenRange(ops, bounds)
    if (screen === null) continue
    mapped.push(
      `${columnLabel(screen.startColumn)}${screen.startRow + 1}` +
        `:${columnLabel(screen.endColumn)}${screen.endRow + 1}`,
    )
  }
  return mapped
}

function mapTitleRowsToScreen(
  titles: string | null,
  ops: readonly StructuralJournalOp[],
): string | null {
  if (titles === null || ops.length === 0) return titles
  const match = /^([0-9]{1,7}):([0-9]{1,7})$/.exec(titles)
  if (!match) return null
  const rows: number[] = []
  for (let row = Number(match[1]) - 1; row <= Number(match[2]) - 1; row += 1) {
    const screen = fileToScreen(ops, 'row', row)
    if (screen !== null) rows.push(screen)
  }
  if (rows.length === 0) return null
  const start = Math.min(...rows)
  const end = Math.max(...rows)
  if (end - start > 20) return null
  return `${start + 1}:${end + 1}`
}

export function resolveEffectivePageSetup(
  journal: PageSetupJournalState,
  file: FilePagePrintSettings | null,
  names: FilePrintNames | null,
  ops: readonly StructuralJournalOp[] = [],
): EffectivePageSetup {
  const clampMargin = (value: number): number => Math.min(Math.max(value, 0), MAX_MARGIN_INCHES)
  const fileMargins = file?.margins
  const margins =
    journal.margins !== undefined
      ? MARGIN_PRESETS[journal.margins]
      : fileMargins !== undefined
        ? {
            left: clampMargin(fileMargins.left),
            right: clampMargin(fileMargins.right),
            top: clampMargin(fileMargins.top),
            bottom: clampMargin(fileMargins.bottom),
            header: clampMargin(fileMargins.header),
            footer: clampMargin(fileMargins.footer),
          }
        : MARGIN_PRESETS.normal
  const fitToPage = journal.fitToPage ?? file?.fitToPage ?? false
  const fitToWidth = journal.fitToWidth ?? file?.fitToWidth ?? (file?.fitToPage === true ? 1 : 0)
  const fitToHeight = journal.fitToHeight ?? file?.fitToHeight ?? (file?.fitToPage === true ? 1 : 0)
  const printAreas =
    journal.printArea !== undefined
      ? journal.printArea === null
        ? []
        : [journal.printArea]
      : mapAreasToScreen(printAreasFromFormula(names?.printArea), ops)
  const printTitles =
    journal.printTitles !== undefined
      ? journal.printTitles
      : mapTitleRowsToScreen(printTitleRowsFromFormula(names?.printTitles), ops)
  const header =
    journal.header !== undefined
      ? journal.header
      : file?.oddHeader !== undefined
        ? decodeHeaderFooter(file.oddHeader)
        : null
  const footer =
    journal.footer !== undefined
      ? journal.footer
      : file?.oddFooter !== undefined
        ? decodeHeaderFooter(file.oddFooter)
        : null
  const decodeOptional = (encoded: string | undefined): HeaderFooterParts | null =>
    encoded === undefined ? null : decodeHeaderFooter(encoded)
  return {
    orientation: journal.orientation ?? file?.orientation ?? 'portrait',
    paperSize: journal.paperSize ?? file?.paperSize ?? 9,
    scale: journal.scale ?? file?.scale ?? 100,
    fitToWidth,
    fitToHeight,
    fitToPage,
    margins,
    printGridlines: journal.printGridlines ?? file?.printGridlines ?? false,
    printHeadings: journal.printHeadings ?? file?.printHeadings ?? false,
    printAreas,
    printTitles,
    header,
    footer,
    firstPage:
      file?.differentFirst === true
        ? { header: decodeOptional(file.firstHeader), footer: decodeOptional(file.firstFooter) }
        : null,
    evenPages:
      file?.differentOddEven === true
        ? { header: decodeOptional(file.evenHeader), footer: decodeOptional(file.evenFooter) }
        : null,
    headerFooterScaleWithDoc: file?.headerFooterFixedSize !== true,
    headerFooterPictures: file?.headerFooterPictures ?? [],
  }
}

function splitAreas(formula: string): string[] {
  const parts: string[] = []
  let current = ''
  let quoted = false
  for (const character of formula) {
    if (character === "'") quoted = !quoted
    if (character === ',' && !quoted) {
      parts.push(current)
      current = ''
    } else {
      current += character
    }
  }
  parts.push(current)
  return parts
}

function plainReference(part: string): string {
  return part
    .slice(part.lastIndexOf('!') + 1)
    .replace(/\$/g, '')
    .trim()
}

export function printAreasFromFormula(formula: string | undefined): string[] {
  if (formula === undefined || formula === '') return []
  const areas: string[] = []
  for (const part of splitAreas(formula)) {
    const reference = plainReference(part).toUpperCase()
    if (/^[A-Z]{1,3}[0-9]{1,7}$/.test(reference)) {
      areas.push(`${reference}:${reference}`)
      continue
    }
    if (/^[A-Z]{1,3}[0-9]{1,7}:[A-Z]{1,3}[0-9]{1,7}$/.test(reference)) {
      areas.push(reference)
      continue
    }
    return []
  }
  return areas
}

export function printTitleRowsFromFormula(formula: string | undefined): string | null {
  if (formula === undefined || formula === '') return null
  for (const part of splitAreas(formula)) {
    const match = /^([0-9]{1,7}):([0-9]{1,7})$/.exec(plainReference(part))
    if (!match) continue
    const start = Number(match[1])
    const end = Number(match[2])
    if (start <= end && end - start <= 20) return `${start}:${end}`
  }
  return null
}

export function decodeHeaderFooter(encoded: string): HeaderFooterParts | null {
  const sections = { L: '', C: '', R: '' }
  let current: 'L' | 'C' | 'R' = 'C'
  let index = 0
  while (index < encoded.length) {
    const character = encoded[index] ?? ''
    if (character !== '&') {
      sections[current] += character
      index += 1
      continue
    }
    const next = encoded[index + 1]
    if (next === undefined) break
    if (next === '&') {
      sections[current] += '&&'
      index += 2
      continue
    }
    if (next === '"') {
      const close = encoded.indexOf('"', index + 2)
      index = close === -1 ? encoded.length : close + 1
      continue
    }
    if (/[0-9]/.test(next)) {
      index += 2
      while (index < encoded.length && /[0-9]/.test(encoded[index] ?? '')) index += 1
      continue
    }
    const code = next.toUpperCase()
    if (code === 'K') {
      index += 8
      continue
    }
    if (code === 'L' || code === 'C' || code === 'R') {
      current = code
      index += 2
      continue
    }
    if (['P', 'N', 'D', 'T', 'F', 'A', 'G'].includes(code)) sections[current] += `&${code}`
    index += 2
  }
  if (sections.L === '' && sections.C === '' && sections.R === '') return null
  return {
    ...(sections.L === '' ? {} : { left: sections.L }),
    ...(sections.C === '' ? {} : { center: sections.C }),
    ...(sections.R === '' ? {} : { right: sections.R }),
  }
}
