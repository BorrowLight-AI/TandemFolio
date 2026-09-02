import type { Editor } from '@tiptap/core'
import {
  DEFAULT_SECTION,
  applySectionStartType,
  applySectionSettings,
  type Block,
  type SaveBlock,
  type SectionInfo,
  type SectionSettings,
} from '@genoffice/docx-engine'

export type DocxSectionStartType = SectionInfo['startType']

export interface InsertSectionBreakInput {
  readonly afterBlockIndex: number
  readonly startType: DocxSectionStartType
}

export interface SectionBreakSource {
  readonly sectPrXml: string
}

export type ResolveSectionBreakSource = (
  afterBlockIndex: number,
) => SectionBreakSource | null

export type InsertSectionBreakResult =
  | {
      readonly ok: true
      readonly insertedBlockIndex: number
      readonly startType: DocxSectionStartType
    }
  | { readonly ok: false; readonly error: 'invalid_arguments'; readonly message: string }

export interface SectionBreakStartMarker {
  readonly saveBlockIndex: number
  readonly startType: DocxSectionStartType
}

export interface MaterializedSectionBreakStarts {
  readonly saveBlocks: SaveBlock[]
  readonly trailingStartType: DocxSectionStartType | null
}

const EMPTY_SECT_PR =
  '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>'

function blockPositionAfter(editor: Editor, afterBlockIndex: number): number {
  let pos = 0
  for (let index = 0; index <= afterBlockIndex; index += 1) {
    pos += editor.state.doc.child(index).nodeSize
  }
  return pos
}

/** Resolve the section whose settings must be copied into a new break paragraph. */
export function resolveSectionBreakSource(
  editor: Editor,
  afterBlockIndex: number,
  sections: readonly SectionInfo[],
  fallbackSettings: SectionSettings | null,
): SectionBreakSource {
  let docxIndex: number | null = null
  if (afterBlockIndex >= 0 && afterBlockIndex < editor.state.doc.childCount) {
    for (let index = afterBlockIndex; index >= 0; index -= 1) {
      const candidate = editor.state.doc.child(index).attrs.docxIndex
      if (typeof candidate === 'number') {
        docxIndex = candidate
        break
      }
    }
  }
  if (docxIndex === null) {
    for (let index = Math.max(0, afterBlockIndex + 1); index < editor.state.doc.childCount; index += 1) {
      const candidate = editor.state.doc.child(index).attrs.docxIndex
      if (typeof candidate === 'number') {
        docxIndex = candidate
        break
      }
    }
  }

  const section =
    (docxIndex === null
      ? sections[0]
      : sections.find(
          (candidate) =>
            docxIndex! >= candidate.firstBlockIndex && docxIndex! <= candidate.lastBlockIndex,
        )) ?? sections[sections.length - 1]
  if (section?.sectPrXml) return { sectPrXml: section.sectPrXml }
  return {
    sectPrXml: applySectionSettings(EMPTY_SECT_PR, fallbackSettings ?? DEFAULT_SECTION),
  }
}

/** Insert one undo-owned native section-break node at a stable top-level boundary. */
export function insertSectionBreakAfterBlock(
  editor: Editor,
  input: InsertSectionBreakInput,
  resolveSource: ResolveSectionBreakSource,
): InsertSectionBreakResult {
  const { doc } = editor.state
  if (input.afterBlockIndex < -1 || input.afterBlockIndex >= doc.childCount) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: `DOCX section-break boundary after block ${input.afterBlockIndex} is invalid for ${doc.childCount} block(s).`,
    }
  }

  const source = resolveSource(input.afterBlockIndex)
  if (!source?.sectPrXml.includes('<w:sectPr')) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: 'DOCX section-break source state is unavailable.',
    }
  }

  const node = editor.state.schema.nodeFromJSON({
    type: 'docProtected',
    attrs: {
      docxIndex: null,
      blockType: 'passthrough',
      label: 'Section break paragraph',
      previewText: '',
      genXml: `<w:p><w:pPr>${source.sectPrXml}</w:pPr></w:p>`,
      sectionStartType: input.startType,
    },
  })
  const pos = blockPositionAfter(editor, input.afterBlockIndex)
  editor.view.dispatch(editor.state.tr.insert(pos, node))
  return {
    ok: true,
    insertedBlockIndex: input.afterBlockIndex + 1,
    startType: input.startType,
  }
}

function patchBlockSectionStartType(xml: string, startType: DocxSectionStartType): string | null {
  const match = /<w:sectPr\b[^>]*>[\s\S]*?<\/w:sectPr>/.exec(xml)
  if (!match) return null
  return xml.replace(match[0], applySectionStartType(match[0], startType))
}

/**
 * OOXML stores a section's start type on that section's terminating sectPr.
 * New break nodes own the desired type for Undo, then save projects it onto
 * the next visible sectPr or the hidden trailing sectPr.
 */
export function materializeSectionBreakStarts(
  sourceBlocks: readonly SaveBlock[],
  markers: readonly SectionBreakStartMarker[],
  originalBlocks: readonly Block[],
): MaterializedSectionBreakStarts {
  const saveBlocks = sourceBlocks.map((block) => ({ ...block })) as SaveBlock[]
  const originalXml = new Map(
    originalBlocks.flatMap((block) =>
      block.docxIndex === null || block.originalXml === undefined
        ? []
        : [[block.docxIndex, block.originalXml] as const],
    ),
  )
  let trailingStartType: DocxSectionStartType | null = null

  for (const marker of markers) {
    let projected = false
    for (let index = marker.saveBlockIndex + 1; index < saveBlocks.length; index += 1) {
      const block = saveBlocks[index]
      const xml =
        block.kind === 'xml'
          ? block.xml
          : block.kind === 'original'
            ? originalXml.get(block.docxIndex)
            : undefined
      if (!xml) continue
      const patched = patchBlockSectionStartType(xml, marker.startType)
      if (patched === null) continue
      const docxIndex =
        block.kind === 'original'
          ? block.docxIndex
          : block.kind === 'xml'
            ? block.docxIndex
            : undefined
      saveBlocks[index] = {
        kind: 'xml',
        xml: patched,
        ...(docxIndex !== undefined ? { docxIndex } : {}),
        ...(block.revision ? { revision: block.revision } : {}),
      }
      projected = true
      break
    }
    if (!projected) trailingStartType = marker.startType
  }

  return { saveBlocks, trailingStartType }
}
