import type { Editor } from '@tiptap/core'
import {
  PAGE_MARK,
  TOTAL_PAGES_MARK,
  type HeaderFooter,
  type Run,
  type SectionInfo,
  type SectionSettings,
} from '@genoffice/docx-engine'
import { hfWithoutPageMarks } from './hf-dom'

export interface SectionSettingsOverride {
  readonly sectionIndex: number
  readonly settings: SectionSettings
  readonly titlePg?: boolean
  readonly evenOddHeaders?: boolean
  readonly pageNumbering?: SectionPageNumbering
  readonly headerFooterEdits?: readonly HeaderFooterEdit[]
}

export const DOCX_PAGE_NUMBER_FORMATS = [
  'decimal',
  'numberInDash',
  'lowerLetter',
  'upperLetter',
  'lowerRoman',
  'upperRoman',
  'chineseCounting',
] as const

export type DocxPageNumberFormat = (typeof DOCX_PAGE_NUMBER_FORMATS)[number]

export function isDocxPageNumberFormat(value: unknown): value is DocxPageNumberFormat {
  return DOCX_PAGE_NUMBER_FORMATS.includes(value as DocxPageNumberFormat)
}

export interface SectionPageNumbering {
  readonly format: DocxPageNumberFormat
  readonly start: number | null
}

export type HeaderFooterKind = 'header' | 'footer'
export type HeaderFooterVariant = 'default' | 'first' | 'even'

export interface HeaderFooterEdit {
  readonly kind: HeaderFooterKind
  readonly variant: HeaderFooterVariant
  readonly value: HeaderFooter
}

export function applySectionSettingsOverrides(
  sections: readonly SectionInfo[],
  overrides: readonly SectionSettingsOverride[],
): SectionInfo[] {
  const byIndex = new Map(overrides.map((override) => [override.sectionIndex, override]))
  return sections.map((section, sectionIndex) => {
    const override = byIndex.get(sectionIndex)
    const pageNumbering = override?.pageNumbering
    return {
      ...section,
      settings: override?.settings ?? section.settings,
      titlePg: override?.titlePg ?? section.titlePg,
      ...(pageNumbering
        ? {
            pageNumberFmt:
              pageNumbering.format === 'decimal' ? undefined : pageNumbering.format,
            pageNumberStart: pageNumbering.start ?? undefined,
          }
        : {}),
    }
  })
}

export function collectSectionSettingsOverrides(editor: Editor): SectionSettingsOverride[] {
  const overrides: SectionSettingsOverride[] = []
  for (let index = 0; index < editor.state.doc.childCount; index += 1) {
    const override = editor.state.doc.child(index).attrs
      .sectionSettingsOverride as SectionSettingsOverride | null
    if (override) overrides.push(override)
  }
  return overrides
}

export interface SetSectionOrientationInput {
  readonly sectionIndex: number
  readonly orientation: SectionSettings['orientation']
}

export interface SectionMarginsInput {
  readonly topTwips: number
  readonly rightTwips: number
  readonly bottomTwips: number
  readonly leftTwips: number
}

export interface SetSectionMarginsInput {
  readonly sectionIndex: number
  readonly margins: SectionMarginsInput
}

export interface SetSectionPageSizeInput {
  readonly sectionIndex: number
  readonly widthTwips: number
  readonly heightTwips: number
}

export interface SetSectionColumnsInput {
  readonly sectionIndex: number
  readonly count: number
  readonly spacingTwips: number
}

export interface SetSectionPageBorderInput {
  readonly sectionIndex: number
  readonly enabled: boolean
}

export interface SetSectionDifferentFirstPageInput {
  readonly sectionIndex: number
  readonly enabled: boolean
}

export interface SetDocumentDifferentOddEvenPagesInput {
  readonly enabled: boolean
}

export interface SetSectionPageNumberingInput extends SectionPageNumbering {
  readonly sectionIndex: number
}

export interface SetHeaderFooterTextInput {
  readonly sectionIndex: number
  readonly kind: HeaderFooterKind
  readonly variant: HeaderFooterVariant
  readonly text: string
}

export interface SetHeaderFooterPageNumberInput {
  readonly sectionIndex: number
  readonly kind: HeaderFooterKind
  readonly variant: HeaderFooterVariant
  readonly enabled: boolean
  readonly alignment: 'left' | 'center' | 'right'
}

export interface HeaderFooterSegmentInput {
  readonly type: 'text' | 'page' | 'total_pages'
  readonly text: string
  readonly bold?: boolean
  readonly italic?: boolean
  readonly underline?: boolean
  readonly strike?: boolean
  readonly color?: string | null
  readonly fontFamily?: string | null
  readonly sizeHalfPoints?: number | null
}

export interface HeaderFooterParagraphInput {
  readonly alignment: 'left' | 'center' | 'right' | 'justify'
  readonly segments: readonly HeaderFooterSegmentInput[]
}

export interface SetHeaderFooterParagraphsInput {
  readonly sectionIndex: number
  readonly kind: HeaderFooterKind
  readonly variant: HeaderFooterVariant
  readonly paragraphs: readonly HeaderFooterParagraphInput[]
}

export type SetSectionLayoutResult =
  | { readonly ok: true; readonly changed: boolean }
  | { readonly ok: false; readonly error: 'invalid_arguments'; readonly message: string }

function sectionAnchor(
  editor: Editor,
  section: SectionInfo,
): { readonly pos: number; readonly attrs: Record<string, unknown> } | null {
  let pos = 0
  let fallback: { readonly pos: number; readonly attrs: Record<string, unknown> } | null = null
  for (let index = 0; index < editor.state.doc.childCount; index += 1) {
    const node = editor.state.doc.child(index)
    const docxIndex = node.attrs.docxIndex
    if (
      typeof docxIndex === 'number' &&
      docxIndex >= section.firstBlockIndex &&
      docxIndex <= section.lastBlockIndex
    ) {
      fallback = { pos, attrs: node.attrs as Record<string, unknown> }
      if (docxIndex === section.lastBlockIndex) return fallback
    }
    pos += node.nodeSize
  }
  return fallback
}

export function setSectionLayoutOverride(
  editor: Editor,
  sectionIndex: number,
  settings: SectionSettings,
  sections: readonly SectionInfo[],
): SetSectionLayoutResult {
  const section = sections[sectionIndex]
  if (!section) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: `DOCX section index ${sectionIndex} is invalid for ${sections.length} section(s).`,
    }
  }
  const anchor = sectionAnchor(editor, section)
  if (!anchor) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: `DOCX section ${sectionIndex} has no addressable visible block.`,
    }
  }
  const current = anchor.attrs.sectionSettingsOverride as SectionSettingsOverride | null
  const effective = current?.sectionIndex === sectionIndex ? current.settings : section.settings
  if (JSON.stringify(effective) === JSON.stringify(settings)) return { ok: true, changed: false }

  editor.view.dispatch(
    editor.state.tr.setNodeMarkup(anchor.pos, undefined, {
      ...anchor.attrs,
      sectionSettingsOverride: {
        ...(current?.sectionIndex === sectionIndex ? current : {}),
        sectionIndex,
        settings,
      },
    }),
  )
  return { ok: true, changed: true }
}

export function setSectionOrientation(
  editor: Editor,
  input: SetSectionOrientationInput,
  sections: readonly SectionInfo[],
): SetSectionLayoutResult {
  const section = sections[input.sectionIndex]
  if (!section) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: `DOCX section index ${input.sectionIndex} is invalid for ${sections.length} section(s).`,
    }
  }
  const anchor = sectionAnchor(editor, section)
  const current = anchor?.attrs.sectionSettingsOverride as SectionSettingsOverride | null | undefined
  const settings = current?.sectionIndex === input.sectionIndex ? current.settings : section.settings
  if (settings.orientation === input.orientation) return { ok: true, changed: false }
  return setSectionLayoutOverride(
    editor,
    input.sectionIndex,
    {
      ...settings,
      orientation: input.orientation,
      pageWidth: settings.pageHeight,
      pageHeight: settings.pageWidth,
    },
    sections,
  )
}

export function setSectionMargins(
  editor: Editor,
  input: SetSectionMarginsInput,
  sections: readonly SectionInfo[],
): SetSectionLayoutResult {
  const section = sections[input.sectionIndex]
  if (!section) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: `DOCX section index ${input.sectionIndex} is invalid for ${sections.length} section(s).`,
    }
  }
  const anchor = sectionAnchor(editor, section)
  const current = anchor?.attrs.sectionSettingsOverride as SectionSettingsOverride | null | undefined
  const settings = current?.sectionIndex === input.sectionIndex ? current.settings : section.settings
  const { topTwips, rightTwips, bottomTwips, leftTwips } = input.margins
  if (
    topTwips + bottomTwips >= settings.pageHeight ||
    leftTwips + rightTwips >= settings.pageWidth
  ) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: 'DOCX section margins must leave positive page width and height.',
    }
  }
  return setSectionLayoutOverride(
    editor,
    input.sectionIndex,
    {
      ...settings,
      marginTop: topTwips,
      marginRight: rightTwips,
      marginBottom: bottomTwips,
      marginLeft: leftTwips,
    },
    sections,
  )
}

export function setSectionPageSize(
  editor: Editor,
  input: SetSectionPageSizeInput,
  sections: readonly SectionInfo[],
): SetSectionLayoutResult & { readonly orientation?: SectionSettings['orientation'] } {
  const section = sections[input.sectionIndex]
  if (!section) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: `DOCX section index ${input.sectionIndex} is invalid for ${sections.length} section(s).`,
    }
  }
  const anchor = sectionAnchor(editor, section)
  const current = anchor?.attrs.sectionSettingsOverride as SectionSettingsOverride | null | undefined
  const settings = current?.sectionIndex === input.sectionIndex ? current.settings : section.settings
  if (
    input.widthTwips <= settings.marginLeft + settings.marginRight ||
    input.heightTwips <= settings.marginTop + settings.marginBottom
  ) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: 'DOCX page size must exceed the current horizontal and vertical margins.',
    }
  }
  const orientation =
    input.widthTwips > input.heightTwips
      ? 'landscape'
      : input.widthTwips < input.heightTwips
        ? 'portrait'
        : settings.orientation
  const result = setSectionLayoutOverride(
    editor,
    input.sectionIndex,
    {
      ...settings,
      pageWidth: input.widthTwips,
      pageHeight: input.heightTwips,
      orientation,
    },
    sections,
  )
  return result.ok ? { ...result, orientation } : result
}

export function setSectionColumns(
  editor: Editor,
  input: SetSectionColumnsInput,
  sections: readonly SectionInfo[],
): SetSectionLayoutResult {
  const section = sections[input.sectionIndex]
  if (!section) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: `DOCX section index ${input.sectionIndex} is invalid for ${sections.length} section(s).`,
    }
  }
  const anchor = sectionAnchor(editor, section)
  const current = anchor?.attrs.sectionSettingsOverride as SectionSettingsOverride | null | undefined
  const settings = current?.sectionIndex === input.sectionIndex ? current.settings : section.settings
  const textWidth = settings.pageWidth - settings.marginLeft - settings.marginRight
  if ((input.count - 1) * input.spacingTwips >= textWidth) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: 'DOCX section column spacing must leave positive text width.',
    }
  }
  return setSectionLayoutOverride(
    editor,
    input.sectionIndex,
    { ...settings, columns: input.count, colSpace: input.spacingTwips },
    sections,
  )
}

export function setSectionPageBorder(
  editor: Editor,
  input: SetSectionPageBorderInput,
  sections: readonly SectionInfo[],
): SetSectionLayoutResult {
  const section = sections[input.sectionIndex]
  if (!section) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: `DOCX section index ${input.sectionIndex} is invalid for ${sections.length} section(s).`,
    }
  }
  const anchor = sectionAnchor(editor, section)
  const current = anchor?.attrs.sectionSettingsOverride as SectionSettingsOverride | null | undefined
  const settings = current?.sectionIndex === input.sectionIndex ? current.settings : section.settings
  return setSectionLayoutOverride(
    editor,
    input.sectionIndex,
    { ...settings, pageBorder: input.enabled },
    sections,
  )
}

export function setSectionDifferentFirstPage(
  editor: Editor,
  input: SetSectionDifferentFirstPageInput,
  sections: readonly SectionInfo[],
): SetSectionLayoutResult {
  const section = sections[input.sectionIndex]
  if (!section) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: `DOCX section index ${input.sectionIndex} is invalid for ${sections.length} section(s).`,
    }
  }
  const anchor = sectionAnchor(editor, section)
  if (!anchor) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: `DOCX section ${input.sectionIndex} has no addressable visible block.`,
    }
  }
  const current = anchor.attrs.sectionSettingsOverride as SectionSettingsOverride | null
  const effectiveTitlePg =
    current?.sectionIndex === input.sectionIndex && current.titlePg !== undefined
      ? current.titlePg
      : section.titlePg
  if (effectiveTitlePg === input.enabled) return { ok: true, changed: false }
  editor.view.dispatch(
    editor.state.tr.setNodeMarkup(anchor.pos, undefined, {
      ...anchor.attrs,
      sectionSettingsOverride: {
        ...(current?.sectionIndex === input.sectionIndex ? current : {}),
        sectionIndex: input.sectionIndex,
        settings:
          current?.sectionIndex === input.sectionIndex ? current.settings : section.settings,
        titlePg: input.enabled,
      },
    }),
  )
  return { ok: true, changed: true }
}

export function setDocumentDifferentOddEvenPages(
  editor: Editor,
  input: SetDocumentDifferentOddEvenPagesInput,
  sections: readonly SectionInfo[],
  currentValue: boolean,
): SetSectionLayoutResult {
  const section = sections[0]
  if (!section) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: 'DOCX document header/footer state is unavailable.',
    }
  }
  const anchor = sectionAnchor(editor, section)
  if (!anchor) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: 'DOCX document has no addressable visible block.',
    }
  }
  const current = anchor.attrs.sectionSettingsOverride as SectionSettingsOverride | null
  const effective = current?.evenOddHeaders ?? currentValue
  if (effective === input.enabled) return { ok: true, changed: false }
  editor.view.dispatch(
    editor.state.tr.setNodeMarkup(anchor.pos, undefined, {
      ...anchor.attrs,
      sectionSettingsOverride: {
        ...(current ?? {}),
        sectionIndex: 0,
        settings: current?.settings ?? section.settings,
        evenOddHeaders: input.enabled,
      },
    }),
  )
  return { ok: true, changed: true }
}

export function setSectionPageNumbering(
  editor: Editor,
  input: SetSectionPageNumberingInput,
  sections: readonly SectionInfo[],
): SetSectionLayoutResult {
  const section = sections[input.sectionIndex]
  if (!section) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: `DOCX section index ${input.sectionIndex} is invalid for ${sections.length} section(s).`,
    }
  }
  const anchor = sectionAnchor(editor, section)
  if (!anchor) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: `DOCX section ${input.sectionIndex} has no addressable visible block.`,
    }
  }
  const current = anchor.attrs.sectionSettingsOverride as SectionSettingsOverride | null
  const effective: SectionPageNumbering = current?.pageNumbering ?? {
    format: (section.pageNumberFmt ?? 'decimal') as DocxPageNumberFormat,
    start: section.pageNumberStart ?? null,
  }
  if (effective.format === input.format && effective.start === input.start) {
    return { ok: true, changed: false }
  }
  editor.view.dispatch(
    editor.state.tr.setNodeMarkup(anchor.pos, undefined, {
      ...anchor.attrs,
      sectionSettingsOverride: {
        ...(current?.sectionIndex === input.sectionIndex ? current : {}),
        sectionIndex: input.sectionIndex,
        settings:
          current?.sectionIndex === input.sectionIndex ? current.settings : section.settings,
        pageNumbering: { format: input.format, start: input.start },
      },
    }),
  )
  return { ok: true, changed: true }
}

export function setHeaderFooterValue(
  editor: Editor,
  input: Omit<SetHeaderFooterTextInput, 'text'> & { readonly value: HeaderFooter },
  sections: readonly SectionInfo[],
  currentValue: HeaderFooter | null = null,
): SetSectionLayoutResult {
  const section = sections[input.sectionIndex]
  if (!section) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: `DOCX section index ${input.sectionIndex} is invalid for ${sections.length} section(s).`,
    }
  }
  const anchor = sectionAnchor(editor, section)
  if (!anchor) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: `DOCX section ${input.sectionIndex} has no addressable visible block.`,
    }
  }
  if (input.value.text.length > 65536) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: 'DOCX header/footer text must not exceed 65,536 characters.',
    }
  }
  const current = anchor.attrs.sectionSettingsOverride as SectionSettingsOverride | null
  const edits = current?.headerFooterEdits ?? []
  const existing = edits.find(
    (edit) => edit.kind === input.kind && edit.variant === input.variant,
  )?.value
  if (JSON.stringify(existing ?? currentValue) === JSON.stringify(input.value)) {
    return { ok: true, changed: false }
  }
  const nextEdit: HeaderFooterEdit = {
    kind: input.kind,
    variant: input.variant,
    value: input.value,
  }
  editor.view.dispatch(
    editor.state.tr.setNodeMarkup(anchor.pos, undefined, {
      ...anchor.attrs,
      sectionSettingsOverride: {
        ...(current?.sectionIndex === input.sectionIndex ? current : {}),
        sectionIndex: input.sectionIndex,
        settings:
          current?.sectionIndex === input.sectionIndex ? current.settings : section.settings,
        headerFooterEdits: [
          ...edits.filter(
            (edit) => edit.kind !== input.kind || edit.variant !== input.variant,
          ),
          nextEdit,
        ],
      },
    }),
  )
  return { ok: true, changed: true }
}

export function setHeaderFooterText(
  editor: Editor,
  input: SetHeaderFooterTextInput,
  sections: readonly SectionInfo[],
  currentValue: HeaderFooter | null = null,
): SetSectionLayoutResult {
  return setHeaderFooterValue(
    editor,
    {
      sectionIndex: input.sectionIndex,
      kind: input.kind,
      variant: input.variant,
      value: { text: input.text },
    },
    sections,
    currentValue,
  )
}

export function setHeaderFooterPageNumber(
  editor: Editor,
  input: SetHeaderFooterPageNumberInput,
  sections: readonly SectionInfo[],
  currentValue: HeaderFooter | null = null,
): SetSectionLayoutResult {
  const currentOverride = collectSectionSettingsOverrides(editor)
    .find((override) => override.sectionIndex === input.sectionIndex)
    ?.headerFooterEdits?.find(
      (edit) => edit.kind === input.kind && edit.variant === input.variant,
    )?.value
  const effective = currentOverride ?? currentValue ?? { text: '' }
  const value: HeaderFooter = input.enabled
    ? {
        text: PAGE_MARK,
        pageNumber: true,
        paras: [{ align: input.alignment, runs: [{ text: PAGE_MARK }] }],
      }
    : hfWithoutPageMarks(effective)
  return setHeaderFooterValue(
    editor,
    {
      sectionIndex: input.sectionIndex,
      kind: input.kind,
      variant: input.variant,
      value,
    },
    sections,
    effective,
  )
}

export function setHeaderFooterParagraphs(
  editor: Editor,
  input: SetHeaderFooterParagraphsInput,
  sections: readonly SectionInfo[],
  currentValue: HeaderFooter | null = null,
): SetSectionLayoutResult {
  if (input.paragraphs.length < 1 || input.paragraphs.length > 64) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: 'DOCX header/footer content must contain 1 to 64 paragraphs.',
    }
  }
  let textLength = 0
  let hasPageNumber = false
  const paras: NonNullable<HeaderFooter['paras']> = []
  for (const paragraph of input.paragraphs) {
    if (paragraph.segments.length > 256) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'Each DOCX header/footer paragraph supports at most 256 segments.',
      }
    }
    const runs: Run[] = []
    for (const segment of paragraph.segments) {
      if (segment.type !== 'text' && segment.text !== '') {
        return {
          ok: false,
          error: 'invalid_arguments',
          message: 'DOCX header/footer field segments must use an empty text value.',
        }
      }
      if (segment.color != null && !/^[0-9A-Fa-f]{6}$/.test(segment.color)) {
        return {
          ok: false,
          error: 'invalid_arguments',
          message: 'DOCX header/footer run colors must be six hexadecimal digits.',
        }
      }
      if (segment.fontFamily != null && segment.fontFamily.length > 256) {
        return {
          ok: false,
          error: 'invalid_arguments',
          message: 'DOCX header/footer font names must not exceed 256 characters.',
        }
      }
      const text =
        segment.type === 'page'
          ? PAGE_MARK
          : segment.type === 'total_pages'
            ? TOTAL_PAGES_MARK
            : segment.text
      textLength += text.length
      hasPageNumber ||= segment.type === 'page'
      runs.push({
        text,
        ...(segment.bold ? { bold: true } : {}),
        ...(segment.italic ? { italic: true } : {}),
        ...(segment.underline ? { underline: true } : {}),
        ...(segment.strike ? { strike: true } : {}),
        ...(segment.color ? { color: segment.color.toUpperCase() } : {}),
        ...(segment.fontFamily
          ? { font: segment.fontFamily, fontAscii: segment.fontFamily }
          : {}),
        ...(segment.sizeHalfPoints ? { sizeHalfPoints: segment.sizeHalfPoints } : {}),
      })
    }
    paras.push({ align: paragraph.alignment, runs })
  }
  if (textLength > 65536) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: 'DOCX header/footer text must not exceed 65,536 characters.',
    }
  }
  const text = paras.flatMap((paragraph) => paragraph.runs).map((run) => run.text).join('')
  return setHeaderFooterValue(
    editor,
    {
      sectionIndex: input.sectionIndex,
      kind: input.kind,
      variant: input.variant,
      value: { text, pageNumber: hasPageNumber, paras },
    },
    sections,
    currentValue,
  )
}
