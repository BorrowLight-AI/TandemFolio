import type { Editor } from '@tiptap/core'
import { buildWordArtParagraphXml, type TextboxDisplay } from '@genoffice/docx-engine'
import { WORDART_PRESETS, wordArtSolidColor } from '@genoffice/ui'

export const DOCX_WORDART_PRESET_IDS = [
  'blue',
  'gold',
  'red',
  'purple',
  'green-italic',
  'white-orange',
  'white-red',
  'gold-brown',
  'sky-navy',
  'navy-white',
  'black-gold',
  'silver-dark',
] as const

export type DocxWordArtPresetId = (typeof DOCX_WORDART_PRESET_IDS)[number]

export interface InsertDocxWordArtInput {
  readonly afterBlockIndex: number
  readonly preset: DocxWordArtPresetId
  readonly text: string
  readonly widthEmu: number
  readonly heightEmu: number
  readonly drawingId: number
}

export type InsertDocxWordArtResult =
  | {
      readonly ok: true
      readonly blockIndex: number
      readonly preset: DocxWordArtPresetId
      readonly drawingId: number
      readonly changed: true
    }
  | {
      readonly ok: false
      readonly error: 'invalid_arguments' | 'execution_failed'
      readonly message: string
    }

function drawingIds(editor: Editor): Set<number> {
  const ids = new Set<number>()
  editor.state.doc.descendants((node) => {
    const xml = typeof node.attrs?.genXml === 'string' ? node.attrs.genXml : ''
    for (const match of xml.matchAll(/<wp:docPr\b[^>]*\bid="([0-9]+)"/g)) ids.add(Number(match[1]))
  })
  return ids
}

export function nextDocxWordArtDrawingId(editor: Editor): number {
  const ids = drawingIds(editor)
  let candidate = 1
  while (ids.has(candidate)) candidate += 1
  return candidate
}

/** Insert one explicit WordArt block after a revision-scoped top-level boundary. */
export function insertDocxWordArtAfterBlock(
  editor: Editor,
  input: InsertDocxWordArtInput,
): InsertDocxWordArtResult {
  const { doc } = editor.state
  const preset = WORDART_PRESETS.find((candidate) => candidate.id === input?.preset)
  if (
    !input ||
    !Number.isInteger(input.afterBlockIndex) ||
    input.afterBlockIndex < -1 ||
    input.afterBlockIndex >= doc.childCount ||
    !preset ||
    typeof input.text !== 'string' ||
    Array.from(input.text).length < 1 ||
    Array.from(input.text).length > 4096 ||
    !Number.isInteger(input.widthEmu) ||
    input.widthEmu < 9_525 ||
    input.widthEmu > 20_000_000 ||
    !Number.isInteger(input.heightEmu) ||
    input.heightEmu < 9_525 ||
    input.heightEmu > 20_000_000 ||
    !Number.isInteger(input.drawingId) ||
    input.drawingId < 1 ||
    input.drawingId > 2_147_483_647 ||
    drawingIds(editor).has(input.drawingId)
  ) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message:
        'DOCX WordArt requires a valid boundary, finite preset, bounded text/geometry, and unique drawing ID.',
    }
  }

  const solidHex = wordArtSolidColor(preset).replace('#', '')
  const xml = buildWordArtParagraphXml({
    text: input.text,
    colorHex: solidHex,
    italic: preset.italic,
    widthEmu: input.widthEmu,
    heightEmu: input.heightEmu,
    id: input.drawingId,
  })
  const textbox: TextboxDisplay = {
    widthPx: Math.round(input.widthEmu / 9525),
    heightPx: Math.round(input.heightEmu / 9525),
    wordArtId: preset.id,
    paras: [
      {
        runs: [
          {
            text: input.text,
            color: solidHex,
            bold: true,
            italic: preset.italic,
            sizeHalfPoints: 72,
          },
        ],
        align: 'center',
      },
    ],
  }
  let position = 0
  for (let index = 0; index <= input.afterBlockIndex; index += 1) {
    position += doc.child(index).nodeSize
  }
  const inserted = editor.chain().focus().insertContentAt(position, {
    type: 'docProtected',
    attrs: {
      docxIndex: null,
      blockType: 'passthrough',
      label: `WordArt(${preset.id})`,
      genXml: xml,
      textboxes: [textbox],
    },
  }).run()
  if (!inserted) {
    return {
      ok: false,
      error: 'execution_failed',
      message: 'The DOCX WordArt block could not be inserted.',
    }
  }
  return {
    ok: true,
    blockIndex: input.afterBlockIndex + 1,
    preset: input.preset,
    drawingId: input.drawingId,
    changed: true,
  }
}
