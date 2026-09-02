import type { ThemeColors, ThemeFonts } from '@genoffice/docx-engine'

export const docxDocumentDesignFields = [
  'pageColor',
  'watermark',
  'themeFonts',
  'themeColors',
] as const

export type DocxDocumentDesignField = (typeof docxDocumentDesignFields)[number]

export interface DocxDocumentDesignState {
  readonly pageColor: string | null
  readonly watermark: string | null
  readonly themeFonts: ThemeFonts | null
  readonly themeColors: ThemeColors | null
}

export interface SetDocxDocumentDesignInput {
  readonly fields: readonly DocxDocumentDesignField[]
  readonly pageColor?: string | null
  readonly watermark?: string | null
  readonly themeFonts?: ThemeFonts
  readonly themeColors?: ThemeColors
}

export type SetDocxDocumentDesignResult =
  | {
      readonly ok: true
      readonly state: DocxDocumentDesignState
      readonly fields: readonly DocxDocumentDesignField[]
      readonly changedFields: readonly DocxDocumentDesignField[]
      readonly changed: boolean
    }
  | {
      readonly ok: false
      readonly error: 'invalid_arguments'
      readonly message: string
    }

const HEX = /^[0-9A-F]{6}$/
const FIELD_SET = new Set<string>(docxDocumentDesignFields)
const COLOR_FIELDS = [
  'dk2',
  'lt2',
  'accent1',
  'accent2',
  'accent3',
  'accent4',
  'accent5',
  'accent6',
] as const

function boundedString(value: unknown, minimum = 1, maximum = 255): value is string {
  return (
    typeof value === 'string' &&
    Array.from(value).length >= minimum &&
    Array.from(value).length <= maximum
  )
}

function validThemeFonts(value: unknown): value is ThemeFonts {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return false
  const fonts = value as Record<string, unknown>
  if (!boundedString(fonts.major) || !boundedString(fonts.minor)) return false
  const allowed = new Set(['major', 'minor', 'eastAsia'])
  if (Object.keys(fonts).some((key) => !allowed.has(key))) return false
  return fonts.eastAsia === undefined || boundedString(fonts.eastAsia)
}

function validThemeColors(value: unknown): value is ThemeColors {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return false
  const colors = value as Record<string, unknown>
  const allowed = new Set(['name', ...COLOR_FIELDS])
  if (Object.keys(colors).some((key) => !allowed.has(key))) return false
  if (!boundedString(colors.name)) return false
  return COLOR_FIELDS.every((field) => typeof colors[field] === 'string' && HEX.test(colors[field]))
}

function sameValue(left: unknown, right: unknown): boolean {
  return JSON.stringify(left) === JSON.stringify(right)
}

/** Resolve one explicit, masked document-design update before any renderer state is committed. */
export function applyDocxDocumentDesign(
  current: DocxDocumentDesignState,
  input: SetDocxDocumentDesignInput,
): SetDocxDocumentDesignResult {
  if (!Array.isArray(input?.fields) || input.fields.length < 1 || input.fields.length > 4) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: 'DOCX document design requires 1–4 explicit fields.',
    }
  }
  const fields = [...input.fields] as DocxDocumentDesignField[]
  if (
    fields.some((field) => !FIELD_SET.has(field)) ||
    new Set(fields).size !== fields.length ||
    docxDocumentDesignFields.some(
      (field) => Object.hasOwn(input, field) !== fields.includes(field),
    )
  ) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: 'DOCX document design fields must be unique and exactly match supplied values.',
    }
  }
  if (
    (fields.includes('pageColor') &&
      input.pageColor !== null &&
      (typeof input.pageColor !== 'string' || !HEX.test(input.pageColor))) ||
    (fields.includes('watermark') &&
      input.watermark !== null &&
      !boundedString(input.watermark)) ||
    (fields.includes('themeFonts') && !validThemeFonts(input.themeFonts)) ||
    (fields.includes('themeColors') && !validThemeColors(input.themeColors))
  ) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: 'DOCX document design values are malformed or outside their bounded schemas.',
    }
  }

  const state: DocxDocumentDesignState = { ...current }
  for (const field of fields) {
    ;(state as unknown as Record<string, unknown>)[field] = input[field]
  }
  const changedFields = fields.filter((field) => !sameValue(current[field], state[field]))
  return { ok: true, state, fields, changedFields, changed: changedFields.length > 0 }
}
