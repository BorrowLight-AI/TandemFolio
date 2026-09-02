import type { Editor } from '@tiptap/core'
import type { SourceInfo } from '@genoffice/docx-engine'
import { closeHistory } from '@tiptap/pm/history'
import { Fragment } from '@tiptap/pm/model'

const SOURCE_TYPES = new Set(['Book', 'JournalArticle', 'InternetSite', 'Report', 'Misc'])

export interface UpsertDocxSourceInput {
  readonly source: {
    readonly tag: string
    readonly type: string
    readonly author: string
    readonly title: string
    readonly year: string
    readonly publisher?: string | null
    readonly url?: string | null
  }
}

export type UpsertDocxSourceResult =
  | {
      readonly ok: true
      readonly tag: string
      readonly created: boolean
      readonly changed: boolean
      readonly sources: readonly SourceInfo[]
    }
  | {
      readonly ok: false
      readonly error: 'invalid_arguments' | 'execution_failed'
      readonly message: string
    }

export interface InsertDocxCitationInput {
  readonly range: { readonly from: number; readonly to: number }
  readonly sourceTag: string
  readonly displayText: string
}

export type InsertDocxCitationResult =
  | {
      readonly ok: true
      readonly from: number
      readonly to: number
      readonly sourceTag: string
      readonly changed: true
    }
  | {
      readonly ok: false
      readonly error: 'invalid_arguments' | 'execution_failed'
      readonly message: string
    }

export interface InsertDocxBibliographyInput {
  readonly afterBlockIndex: number
  readonly heading: string
  readonly entries: readonly {
    readonly sourceTag: string
    readonly text: string
  }[]
}

export type InsertDocxBibliographyResult =
  | {
      readonly ok: true
      readonly afterBlockIndex: number
      readonly entries: number
      readonly insertedBlocks: number
    }
  | {
      readonly ok: false
      readonly error: 'invalid_arguments' | 'execution_failed'
      readonly message: string
    }

function unicodeLength(value: unknown): number {
  return typeof value === 'string' ? Array.from(value).length : -1
}

function validOptionalString(value: unknown, maxLength: number): boolean {
  return value === undefined || value === null || unicodeLength(value) <= maxLength
}

function normalizeSource(input: UpsertDocxSourceInput['source']): SourceInfo {
  return {
    tag: input.tag,
    type: input.type,
    author: input.author,
    title: input.title,
    year: input.year,
    ...(input.publisher ? { publisher: input.publisher } : {}),
    ...(input.url ? { url: input.url } : {}),
  }
}

/** Read the Undo-owned final source snapshot, falling back to the parsed customXml part. */
export function collectDocxSources(
  editor: Editor,
  parsedSources: readonly SourceInfo[],
): SourceInfo[] {
  const raw = editor.state.doc.firstChild?.attrs.sourceStateOverride
  if (typeof raw !== 'string') return [...parsedSources]
  try {
    const value = JSON.parse(raw) as unknown
    return Array.isArray(value) ? (value as SourceInfo[]) : [...parsedSources]
  } catch {
    return [...parsedSources]
  }
}

/** Add or replace one stable bibliography tag through a native-history snapshot. */
export function upsertDocxSource(
  editor: Editor,
  currentSources: readonly SourceInfo[],
  input: UpsertDocxSourceInput,
): UpsertDocxSourceResult {
  const source = input?.source
  if (
    !source ||
    unicodeLength(source.tag) < 1 ||
    unicodeLength(source.tag) > 255 ||
    !/^[A-Za-z0-9_.-]+$/.test(source.tag) ||
    !SOURCE_TYPES.has(source.type) ||
    unicodeLength(source.author) < 0 ||
    unicodeLength(source.author) > 4096 ||
    unicodeLength(source.title) < 1 ||
    unicodeLength(source.title) > 4096 ||
    unicodeLength(source.year) < 0 ||
    unicodeLength(source.year) > 32 ||
    !validOptionalString(source.publisher, 4096) ||
    !validOptionalString(source.url, 4096)
  ) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message:
        'DOCX source upsert requires one bounded stable tag, finite type, and bounded fields.',
    }
  }

  const normalized = normalizeSource(source)
  const index = currentSources.findIndex((entry) => entry.tag === normalized.tag)
  const created = index === -1
  if (created && currentSources.length >= 1024) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: 'DOCX source upsert supports at most 1024 sources.',
    }
  }
  const sources = created
    ? [...currentSources, normalized]
    : currentSources.map((entry, entryIndex) => (entryIndex === index ? normalized : entry))
  if (Array.from(JSON.stringify(sources)).length > 65_536) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: 'DOCX source snapshot exceeds 65536 Unicode characters.',
    }
  }
  if (!created && JSON.stringify(currentSources[index]) === JSON.stringify(normalized)) {
    return { ok: true, tag: normalized.tag, created: false, changed: false, sources }
  }

  try {
    const anchor = editor.state.doc.firstChild
    if (!anchor) throw new Error('document has no source-state anchor')
    const transaction = closeHistory(
      editor.state.tr.setNodeMarkup(0, undefined, {
        ...anchor.attrs,
        sourceStateOverride: JSON.stringify(sources),
      }),
    )
    editor.view.dispatch(transaction)
    return { ok: true, tag: normalized.tag, created, changed: true, sources }
  } catch (error) {
    return {
      ok: false,
      error: 'execution_failed',
      message: `The mounted DOCX editor rejected source upsert: ${
        error instanceof Error ? error.message : String(error)
      }`,
    }
  }
}

/** Insert the retained renderer's plain-text citation for one stable source tag. */
export function insertDocxCitation(
  editor: Editor,
  sources: readonly SourceInfo[],
  input: InsertDocxCitationInput,
): InsertDocxCitationResult {
  const { from, to } = input?.range ?? {}
  if (
    !Number.isInteger(from) ||
    !Number.isInteger(to) ||
    from < 1 ||
    to < from ||
    to > editor.state.doc.content.size ||
    unicodeLength(input.sourceTag) < 1 ||
    unicodeLength(input.sourceTag) > 255 ||
    unicodeLength(input.displayText) < 1 ||
    unicodeLength(input.displayText) > 4096 ||
    !sources.some((source) => source.tag === input.sourceTag)
  ) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message:
        'DOCX citation insertion requires an exact inline range and one existing bounded source tag.',
    }
  }
  const $from = editor.state.doc.resolve(from)
  const $to = editor.state.doc.resolve(to)
  if (!$from.sameParent($to) || !$from.parent.inlineContent) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: 'DOCX citation insertion requires one inline text block.',
    }
  }
  try {
    editor.view.dispatch(closeHistory(editor.state.tr.insertText(input.displayText, from, to)))
    return { ok: true, from, to, sourceTag: input.sourceTag, changed: true }
  } catch (error) {
    return {
      ok: false,
      error: 'execution_failed',
      message: `The mounted DOCX editor rejected citation insertion: ${
        error instanceof Error ? error.message : String(error)
      }`,
    }
  }
}

/** Insert one explicit bibliography snapshot after a stable top-level block boundary. */
export function insertDocxBibliography(
  editor: Editor,
  sources: readonly SourceInfo[],
  input: InsertDocxBibliographyInput,
): InsertDocxBibliographyResult {
  const entries = input?.entries
  if (
    !Number.isInteger(input?.afterBlockIndex) ||
    input.afterBlockIndex < -1 ||
    input.afterBlockIndex >= editor.state.doc.childCount ||
    unicodeLength(input.heading) < 1 ||
    unicodeLength(input.heading) > 4096 ||
    !Array.isArray(entries) ||
    entries.length < 1 ||
    entries.length > 1024 ||
    entries.some(
      (entry) =>
        unicodeLength(entry?.sourceTag) < 1 ||
        unicodeLength(entry.sourceTag) > 255 ||
        unicodeLength(entry?.text) < 1 ||
        unicodeLength(entry.text) > 4096 ||
        !sources.some((source) => source.tag === entry.sourceTag),
    ) ||
    unicodeLength(JSON.stringify({ heading: input.heading, entries })) > 65_536
  ) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message:
        'DOCX bibliography insertion requires one stable block boundary and bounded entries backed by existing source tags.',
    }
  }

  try {
    let position = 0
    for (let index = 0; index <= input.afterBlockIndex; index += 1) {
      position += editor.state.doc.child(index).nodeSize
    }
    const heading = editor.schema.nodes.docHeading.create(
      { docxIndex: null, styleId: null, externalChanged: false, level: 1 },
      editor.schema.text(input.heading),
    )
    const paragraphs = entries.map((entry) =>
      editor.schema.nodes.docParagraph.create(
        { docxIndex: null, styleId: null, externalChanged: false },
        editor.schema.text(entry.text),
      ),
    )
    editor.view.dispatch(
      closeHistory(editor.state.tr.insert(position, Fragment.fromArray([heading, ...paragraphs]))),
    )
    return {
      ok: true,
      afterBlockIndex: input.afterBlockIndex,
      entries: entries.length,
      insertedBlocks: entries.length + 1,
    }
  } catch (error) {
    return {
      ok: false,
      error: 'execution_failed',
      message: `The mounted DOCX editor rejected bibliography insertion: ${
        error instanceof Error ? error.message : String(error)
      }`,
    }
  }
}
