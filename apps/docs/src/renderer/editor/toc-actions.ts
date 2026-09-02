import type { Editor } from '@tiptap/core'
import { Fragment, type Node as PmNode } from '@tiptap/pm/model'
import { closeHistory } from '@tiptap/pm/history'
import { generateTocFieldXml, type Block, type TocEntry } from '@genoffice/docx-engine'

export interface RefreshDocxTocEntryInput {
  readonly level: number
  readonly text: string
  readonly pageNumber: number | null
}

export interface RefreshDocxTocInput {
  readonly tocBlockIndex: number
  readonly entries: readonly RefreshDocxTocEntryInput[]
}

export type RefreshDocxTocResult =
  | {
      readonly ok: true
      readonly entries: number
      readonly replacedBlocks: number
    }
  | {
      readonly ok: false
      readonly error: 'invalid_arguments' | 'execution_failed'
      readonly message: string
    }

export interface DocxTocLabels {
  readonly toc: string
  readonly pageBreak: string
}

const DEFAULT_LABELS: DocxTocLabels = { toc: 'TOC field', pageBreak: 'Page break' }
const PAGE_BREAK_PARAGRAPH_XML = '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'

function invalid(message: string): RefreshDocxTocResult {
  return { ok: false, error: 'invalid_arguments', message }
}

/** Resolve retained source XML or an editor-generated fragment for one top-level node. */
export function docxXmlOfNode(
  node: { readonly attrs: Readonly<Record<string, unknown>> },
  blocks: readonly Block[],
): string {
  if (node.attrs.genXml) return String(node.attrs.genXml)
  const index = node.attrs.docxIndex
  if (index === null || index === undefined) return ''
  return blocks.find((block) => block.docxIndex === index)?.originalXml ?? ''
}

export function createDocxTocFieldNodes(
  entries: readonly TocEntry[],
  label = DEFAULT_LABELS.toc,
): Array<Record<string, unknown>> {
  return generateTocFieldXml([...entries]).map((xml, index) => ({
    type: 'docProtected',
    attrs: {
      docxIndex: null,
      blockType: 'passthrough',
      label,
      genXml: xml,
      fieldDisplay: {
        kind: 'tocLine',
        left: entries[index].text,
        right: entries[index].pageNo !== undefined ? String(entries[index].pageNo) : '',
        level: entries[index].level,
      },
    },
  }))
}

export function findFirstDocxTocBlockIndex(
  editor: Editor,
  blocks: readonly Block[],
): number | null {
  let found: number | null = null
  editor.state.doc.forEach((node, _offset, index) => {
    if (found === null && /<w:instrText[^>]*>\s*TOC[\s\\]/.test(docxXmlOfNode(node, blocks))) {
      found = index
    }
  })
  return found
}

/** Replace one exact existing TOC field region through native history. */
export function refreshDocxToc(
  editor: Editor,
  blocks: readonly Block[],
  input: RefreshDocxTocInput,
  labels: DocxTocLabels = DEFAULT_LABELS,
): RefreshDocxTocResult {
  const { doc } = editor.state
  if (
    !input ||
    !Number.isInteger(input.tocBlockIndex) ||
    input.tocBlockIndex < 0 ||
    input.tocBlockIndex >= doc.childCount
  ) {
    return invalid(
      `DOCX TOC block ${input?.tocBlockIndex ?? 'unknown'} is invalid for ${doc.childCount} block(s).`,
    )
  }
  if (!Array.isArray(input.entries) || input.entries.length < 1 || input.entries.length > 1024) {
    return invalid('DOCX TOC refresh requires 1 through 1024 entries.')
  }
  let totalCharacters = 0
  const entries: TocEntry[] = []
  for (const entry of input.entries) {
    const textLength = typeof entry?.text === 'string' ? Array.from(entry.text).length : 0
    if (!Number.isInteger(entry?.level) || entry.level < 1 || entry.level > 9) {
      return invalid('DOCX TOC entry levels must be integers from 1 through 9.')
    }
    if (textLength < 1 || textLength > 4096) {
      return invalid('DOCX TOC entry text must contain 1 through 4096 Unicode characters.')
    }
    if (
      entry.pageNumber !== null &&
      (!Number.isInteger(entry.pageNumber) ||
        entry.pageNumber < 1 ||
        entry.pageNumber > 2_147_483_647)
    ) {
      return invalid('DOCX TOC page numbers must be positive integers or null.')
    }
    totalCharacters += textLength
    if (totalCharacters > 65_536) {
      return invalid('DOCX TOC entries exceed the 65536 character aggregate limit.')
    }
    entries.push({
      level: entry.level,
      text: entry.text,
      ...(entry.pageNumber === null ? {} : { pageNo: entry.pageNumber }),
    })
  }

  let from = 0
  for (let index = 0; index < input.tocBlockIndex; index += 1) from += doc.child(index).nodeSize
  const firstXml = docxXmlOfNode(doc.child(input.tocBlockIndex), blocks)
  if (!/<w:instrText[^>]*>\s*TOC[\s\\]/.test(firstXml)) {
    return invalid(`DOCX block ${input.tocBlockIndex} is not the start of a TOC field.`)
  }

  let depth = 0
  let to = -1
  let endIndex = -1
  let offset = from
  let keepPageBreak = false
  for (let index = input.tocBlockIndex; index < doc.childCount; index += 1) {
    const node = doc.child(index)
    const xml = docxXmlOfNode(node, blocks)
    depth += (xml.match(/w:fldCharType="begin"/g) ?? []).length
    depth -= (xml.match(/w:fldCharType="end"/g) ?? []).length
    offset += node.nodeSize
    if (depth <= 0) {
      to = offset
      endIndex = index
      keepPageBreak = /<w:br\s[^>]*w:type="page"/.test(xml)
      break
    }
  }
  if (to < 0 || endIndex < input.tocBlockIndex) {
    return invalid(`DOCX TOC at block ${input.tocBlockIndex} has no matching field end.`)
  }

  const jsonNodes = createDocxTocFieldNodes(entries, labels.toc)
  if (keepPageBreak) {
    jsonNodes.push({
      type: 'docProtected',
      attrs: {
        docxIndex: null,
        blockType: 'passthrough',
        label: labels.pageBreak,
        genXml: PAGE_BREAK_PARAGRAPH_XML,
        fieldDisplay: { kind: 'pageBreak' },
      },
    })
  }

  try {
    const nodes: PmNode[] = jsonNodes.map((node) => editor.state.schema.nodeFromJSON(node))
    editor.view.dispatch(
      closeHistory(editor.state.tr.replaceWith(from, to, Fragment.fromArray(nodes))),
    )
    return {
      ok: true,
      entries: entries.length,
      replacedBlocks: endIndex - input.tocBlockIndex + 1,
    }
  } catch (error) {
    return {
      ok: false,
      error: 'execution_failed',
      message: `The mounted DOCX editor rejected the TOC refresh: ${
        error instanceof Error ? error.message : String(error)
      }`,
    }
  }
}
