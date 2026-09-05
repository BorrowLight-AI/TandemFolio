import type { WorkbookFile } from '../shared/desktop-api'

type PivotTable = WorkbookFile['sheets'][number]['pivotTables'][number]
type PivotStyle = Omit<PivotTable, 'path' | 'cachePath' | 'outputRef'>
type Table = WorkbookFile['sheets'][number]['tables'][number]
type TableStyle = Omit<
  Table,
  | 'range'
  | 'headerRowCount'
  | 'showRowStripes'
  | 'showColumnStripes'
  | 'name'
  | 'columns'
  | 'totalsRowCount'
  | 'styleName'
  | 'filterActive'
>

const DEFAULT_THEME = [
  '#FFFFFF',
  '#000000',
  '#E7E6E6',
  '#44546A',
  '#4472C4',
  '#ED7D31',
  '#A5A5A5',
  '#FFC000',
  '#5B9BD5',
  '#70AD47',
] as const

function attribute(attributes: string, name: string): string | undefined {
  const escaped = name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
  return new RegExp(`(?:^|\\s)${escaped}="([^"]*)"`).exec(attributes)?.[1]
}

function elementAttributes(xml: string, name: string): string | undefined {
  return new RegExp(`<(?:[A-Za-z_][\\w.-]*:)?${name}\\b([^>]*)`).exec(xml)?.[1]
}

function elementBody(xml: string, name: string): string | undefined {
  return new RegExp(
    `<(?:[A-Za-z_][\\w.-]*:)?${name}\\b[^>]*>([\\s\\S]*?)<\\/(?:[A-Za-z_][\\w.-]*:)?${name}>`,
  ).exec(xml)?.[1]
}

function enabled(value: string | undefined, fallback = false): boolean {
  return value === undefined ? fallback : value === '1' || value === 'true'
}

function rgbToHsl(red: number, green: number, blue: number): [number, number, number] {
  const r = red / 255
  const g = green / 255
  const b = blue / 255
  const max = Math.max(r, g, b)
  const min = Math.min(r, g, b)
  const luminance = (max + min) / 2
  if (max === min) return [0, 0, luminance]
  const delta = max - min
  const saturation = luminance > 0.5 ? delta / (2 - max - min) : delta / (max + min)
  const hue =
    max === r
      ? ((g - b) / delta + (g < b ? 6 : 0)) * 60
      : max === g
        ? ((b - r) / delta + 2) * 60
        : ((r - g) / delta + 4) * 60
  return [hue, saturation, luminance]
}

function hslToHex(hue: number, saturation: number, luminance: number): string {
  const channel = (offset: number): number => {
    const k = (offset + hue / 30) % 12
    const a = saturation * Math.min(luminance, 1 - luminance)
    return Math.round(
      (luminance - a * Math.max(-1, Math.min(k - 3, 9 - k, 1))) * 255,
    )
  }
  const hex = (value: number): string => value.toString(16).padStart(2, '0').toUpperCase()
  return `#${hex(channel(0))}${hex(channel(8))}${hex(channel(4))}`
}

function tint(hex: string, amount: number): string {
  const value = hex.replace(/^#/, '')
  if (!/^[0-9A-Fa-f]{6}$/.test(value)) return hex.toUpperCase()
  const [hue, saturation, luminance] = rgbToHsl(
    Number.parseInt(value.slice(0, 2), 16),
    Number.parseInt(value.slice(2, 4), 16),
    Number.parseInt(value.slice(4, 6), 16),
  )
  const next = amount < 0 ? luminance * (1 + amount) : luminance * (1 - amount) + amount
  return hslToHex(hue, saturation, Math.min(Math.max(next, 0), 1))
}

function pivotPalette(styleName: string | undefined, theme: readonly string[]): PivotStyle {
  const match = /^(?:PivotStyle)(Light|Medium|Dark)(\d+)$/.exec(styleName ?? '')
  if (!match) return {}
  const family = match[1]?.toLowerCase()
  const number = Number(match[2])
  if (!Number.isInteger(number) || number < 1 || number > 28) return {}
  const dark1 = theme[1] ?? DEFAULT_THEME[1]
  const accentIndex = (number - 1) % 7
  const neutral = accentIndex === 0
  const base = neutral
    ? dark1
    : (theme[accentIndex + 3] ?? DEFAULT_THEME[accentIndex + 3] ?? DEFAULT_THEME[4])
  const band = (accentTint: number, neutralTint: number): string =>
    tint(base, neutral ? neutralTint : accentTint)
  const light = (): string => band(0.8, 0.85)
  const mid = (): string => band(0.6, 0.75)
  const deep = (): string => band(0.4, 0.65)
  const shade = (): string => band(-0.25, 0.5)
  const solid = (): string => band(0, 0.5)
  const gray = (amount: number): string => tint(dark1, amount)
  const variant = Math.floor((number - 1) / 7)
  const white = '#FFFFFF'

  if (family === 'light' && variant === 0) {
    return {
      headerBold: true,
      stripeFill: light(),
      columnStripeFill: light(),
      subheadingBold: true,
      subheading2FontColor: solid(),
      subheading2Bold: number !== 2,
      subtotalBold: true,
      totalRowFill: white,
      totalRowBold: true,
    }
  }
  if (family === 'light' && variant === 1) {
    return {
      headerFontColor: dark1,
      headerBold: true,
      wholeTableFontColor: neutral ? dark1 : shade(),
      subheadingFill: light(),
      subheadingFontColor: dark1,
      subheadingBold: true,
      subheading2FontColor: dark1,
      subheading2Bold: true,
      subtotalFill: light(),
      subtotalFontColor: dark1,
      subtotalBold: true,
      totalRowFontColor: dark1,
      totalRowBold: true,
    }
  }
  if (family === 'light' && variant === 2) {
    return {
      headerFill: light(),
      headerBold: true,
      stripeFill: gray(0.85),
      columnStripeFill: gray(0.85),
      subheadingBold: true,
      subheading2Bold: true,
      subtotalBold: true,
      totalRowFill: light(),
      totalRowBold: true,
    }
  }
  if (family === 'light') {
    return {
      headerBold: neutral,
      wholeTableFontColor: neutral ? dark1 : shade(),
      stripeFill: light(),
      columnStripeFill: light(),
      firstColumnBold: neutral,
      subtotalBold: neutral,
      totalRowBold: neutral,
    }
  }
  if (family === 'medium' && variant === 0) {
    return {
      headerFill: shade(),
      headerFontColor: white,
      headerBold: false,
      firstHeaderCellBold: true,
      subheadingFill: deep(),
      subheadingFontColor: white,
      subheadingBold: false,
      subheading2Fill: light(),
      subheading2Bold: false,
      subtotalFill: deep(),
      subtotalFontColor: white,
      subtotalBold: true,
      totalRowBold: true,
    }
  }
  if (family === 'medium' && variant === 1) {
    return {
      headerFill: solid(),
      headerFontColor: white,
      headerBold: true,
      subheadingFill: light(),
      subheadingBold: true,
      subheading2Bold: true,
      subtotalFill: mid(),
      subtotalBold: true,
      totalRowBold: true,
    }
  }
  if (family === 'medium' && variant === 2) {
    return {
      headerFill: dark1,
      headerFontColor: white,
      headerBold: false,
      wholeTableFill: band(0.8, 0.95),
      stripeFill: light(),
      columnStripeFill: light(),
      subheadingBold: true,
      subheading2FontColor: gray(0.5),
      subheading2Bold: true,
      subtotalBold: true,
      totalRowFill: dark1,
      totalRowFontColor: white,
      totalRowBold: false,
    }
  }
  if (family === 'medium') {
    return {
      headerBold: true,
      wholeTableFill: light(),
      wholeTableFontColor: shade(),
      secondRowStripeFill: mid(),
      secondColumnStripeFill: mid(),
      firstColumnFill: mid(),
      firstColumnBold: true,
      subheadingFontColor: dark1,
      subheadingBold: true,
      subtotalFontColor: dark1,
      subtotalBold: true,
      totalRowBold: true,
    }
  }
  if (family === 'dark' && variant === 0) {
    const dark = band(-0.5, 0.5)
    return {
      headerFill: dark,
      headerFontColor: white,
      headerBold: true,
      wholeTableFill: mid(),
      secondRowStripeFill: deep(),
      subheadingFill: light(),
      subheadingBold: true,
      subheading2Bold: true,
      subtotalBold: true,
      totalRowFill: dark,
      totalRowFontColor: white,
      totalRowBold: true,
    }
  }
  if (family === 'dark' && variant === 1) {
    return {
      headerFill: gray(0.25),
      headerFontColor: white,
      headerBold: true,
      wholeTableFill: light(),
      subheadingFill: mid(),
      subheadingBold: true,
      subheading2Bold: true,
      subtotalFill: mid(),
      subtotalBold: true,
      totalRowFill: gray(0.25),
      totalRowFontColor: white,
      totalRowBold: true,
    }
  }
  if (family === 'dark' && variant === 2) {
    return {
      headerFill: dark1,
      headerFontColor: white,
      headerBold: true,
      wholeTableFill: band(0, 0.55),
      wholeTableFontColor: light(),
      subheadingFill: shade(),
      subheadingFontColor: white,
      subheadingBold: true,
      subheading2FontColor: white,
      subheading2Bold: true,
      subtotalFontColor: white,
      subtotalBold: true,
      totalRowFill: dark1,
      totalRowFontColor: white,
      totalRowBold: true,
    }
  }
  return {
    headerFill: shade(),
    headerBold: false,
    firstHeaderCellFontColor: white,
    firstHeaderCellBold: true,
    wholeTableFill: solid(),
    wholeTableFontColor: light(),
    secondRowStripeFill: band(0.4, 0.55),
    secondColumnStripeFill: band(0.4, 0.55),
    firstColumnFill: shade(),
    firstColumnBold: false,
    subheadingFontColor: white,
    subheadingBold: true,
    subheading2FontColor: white,
    subheading2Bold: true,
    subtotalBold: neutral,
    totalRowFontColor: white,
    totalRowBold: true,
  }
}

function pivotRowKinds(xml: string): string {
  const pivotFields = [...(elementBody(xml, 'pivotFields') ?? '').matchAll(
    /<(?:[A-Za-z_][\w.-]*:)?pivotField\b([^>]*)/g,
  )].map((match) => attribute(match[1] ?? '', 'outline') !== '0')
  const rowFieldOutline = [...(elementBody(xml, 'rowFields') ?? '').matchAll(
    /<(?:[A-Za-z_][\w.-]*:)?field\b([^>]*)/g,
  )].map((match) => {
    const index = Number(attribute(match[1] ?? '', 'x'))
    return Number.isInteger(index) ? (pivotFields[index] ?? true) : true
  })
  return [...(elementBody(xml, 'rowItems') ?? '').matchAll(
    /<(?:[A-Za-z_][\w.-]*:)?i\b([^>]*)/g,
  )]
    .map((match) => {
      const attributes = match[1] ?? ''
      const depth = Number(attribute(attributes, 'r') ?? 0)
      const kind = attribute(attributes, 't') ?? 'data'
      if (kind === 'grand') return 'g'
      if (kind === 'blank') return 'b'
      if (kind !== 'data') return 't'
      const outer = depth + 1 < rowFieldOutline.length && (rowFieldOutline[depth] ?? true)
      return outer ? (depth === 0 ? 's' : 'S') : 'd'
    })
    .join('')
}

/** Extracts the file-native pivot layout and built-in style bands. */
export function readPivotStyleMetadata(
  pivotXml: string,
  themeColors: readonly string[] = DEFAULT_THEME,
): PivotStyle {
  const location = elementAttributes(pivotXml, 'location') ?? ''
  const root = elementAttributes(pivotXml, 'pivotTableDefinition') ?? ''
  const style = elementAttributes(pivotXml, 'pivotTableStyleInfo')
  const styleName = style ? attribute(style, 'name') : undefined
  const showRowStripes = enabled(style ? attribute(style, 'showRowStripes') : undefined)
  const showColumnStripes = enabled(style ? attribute(style, 'showColStripes') : undefined)
  const palette = pivotPalette(styleName, themeColors)
  if (!showRowStripes) {
    delete palette.stripeFill
    delete palette.secondRowStripeFill
  }
  if (!showColumnStripes) {
    delete palette.columnStripeFill
    delete palette.secondColumnStripeFill
  }
  return {
    ...palette,
    ...(styleName ? { styled: true } : {}),
    firstDataRow: Number(attribute(location, 'firstDataRow') ?? 1),
    firstDataCol: Number(attribute(location, 'firstDataCol') ?? 1),
    rowGrandTotals: enabled(attribute(root, 'rowGrandTotals'), true),
    rowKinds: pivotRowKinds(pivotXml),
  }
}

/** Resolves Excel's built-in table style families against the workbook theme. */
export function readTableStyleMetadata(
  styleName: string | undefined,
  themeColors: readonly string[] = DEFAULT_THEME,
): TableStyle {
  if (!styleName) return {}
  const match = /^TableStyle(Light|Medium|Dark)(\d+)$/.exec(styleName)
  const family = match?.[1]?.toLowerCase() ?? 'medium'
  const number = match ? Number(match[2]) : 2
  const dark1 = themeColors[1] ?? DEFAULT_THEME[1]
  const accent = (index: number): string => {
    const normalized = ((index - 1) % 6 + 6) % 6
    return themeColors[normalized + 4] ?? DEFAULT_THEME[normalized + 4] ?? DEFAULT_THEME[4]
  }
  const accentIndex = ((number - 1) % 7 + 7) % 7
  const neutral = accentIndex === 0
  const base = neutral ? dark1 : accent(accentIndex)
  const band = (color: string, fromDark1: boolean, amount: number): string =>
    tint(color, fromDark1 ? amount + 0.05 : amount)
  const variant = Math.floor(Math.max(number - 1, 0) / 7)
  const white = '#FFFFFF'

  if (family === 'light' && variant === 0) {
    return {
      headerFontColor: tint(base, -0.25),
      stripeFill: band(base, neutral, 0.8),
      totalRowFontColor: tint(base, -0.25),
      totalRowBorderColor: base,
      totalRowBorderStyle: 'thin',
      borderColor: base,
    }
  }
  if (family === 'light' && variant === 1) {
    return {
      headerFill: base,
      headerFontColor: white,
      totalRowBorderColor: base,
      totalRowBorderStyle: 'double',
    }
  }
  if (family === 'light') {
    return {
      headerFontColor: dark1,
      stripeFill: band(base, neutral, 0.8),
      totalRowBorderColor: base,
      totalRowBorderStyle: 'double',
      wholeTableBorderColor: base,
      wholeTableBorderStyle: 'thin',
      innerHorizontalBorderColor: base,
      innerHorizontalBorderStyle: 'thin',
      innerVerticalBorderColor: base,
      innerVerticalBorderStyle: 'thin',
    }
  }
  if (family === 'medium' && variant === 1) {
    return {
      headerFill: base,
      headerFontColor: white,
      stripeFill: band(base, neutral, 0.6),
      secondRowStripeFill: band(base, neutral, 0.8),
      wholeTableFill: band(base, neutral, 0.8),
      totalRowFill: base,
      totalRowFontColor: white,
    }
  }
  if (family === 'medium' && variant === 2) {
    return {
      headerFill: base,
      headerFontColor: white,
      stripeFill: tint(dark1, 0.85),
      totalRowBorderColor: dark1,
      totalRowBorderStyle: 'double',
    }
  }
  if (family === 'medium' && variant === 3) {
    return {
      headerFill: band(base, neutral, 0.8),
      headerFontColor: dark1,
      stripeFill: band(base, neutral, 0.6),
      wholeTableFill: band(base, neutral, 0.8),
      totalRowFill: band(base, neutral, 0.8),
      totalRowBorderColor: base,
      totalRowBorderStyle: 'medium',
    }
  }
  if (family === 'medium') {
    return {
      headerFill: base,
      headerFontColor: white,
      stripeFill: band(base, neutral, 0.8),
      totalRowBorderColor: base,
      totalRowBorderStyle: 'double',
    }
  }
  if (variant === 1) {
    const bandBase = neutral ? dark1 : accent(2 * (number - 8) - 1)
    const headerBase = neutral ? dark1 : accent(2 * (number - 8))
    return {
      headerFill: headerBase,
      headerFontColor: white,
      stripeFill: band(bandBase, neutral, 0.6),
      secondRowStripeFill: band(bandBase, neutral, 0.8),
      wholeTableFill: band(bandBase, neutral, 0.8),
      totalRowFill: band(bandBase, neutral, 0.8),
      totalRowBorderColor: dark1,
      totalRowBorderStyle: 'double',
    }
  }
  return {
    headerFill: dark1,
    headerFontColor: white,
    wholeTableFill: neutral ? tint(dark1, 0.45) : base,
    stripeFill: neutral ? tint(dark1, 0.25) : tint(base, -0.25),
    totalRowFill: neutral ? tint(dark1, 0.15) : tint(base, -0.5),
    bodyFontColor: white,
    totalRowFontColor: white,
  }
}
