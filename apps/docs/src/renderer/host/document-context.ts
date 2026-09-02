import type { Editor } from '@tiptap/core'

export interface DocxBlockSummary {
  index: number
  type: string
  text: string
  headingLevel?: number
}

export interface DocxSelectionContext {
  from: number
  to: number
  empty: boolean
  activeBlockIndex: number | null
  totalBlocks: number
  blocks: DocxBlockSummary[]
}

export function buildDocxSelectionContext(editor: Editor): DocxSelectionContext {
  const { from, to, empty } = editor.state.selection
  const allBlocks: Array<DocxBlockSummary & { start: number; end: number }> = []
  editor.state.doc.forEach((node, offset, index) => {
    allBlocks.push({
      index,
      type: node.type.name,
      text: node.textContent.slice(0, 500),
      ...(node.type.name === 'docHeading' && Number.isInteger(node.attrs.level)
        ? { headingLevel: node.attrs.level as number }
        : {}),
      start: offset + 1,
      end: offset + node.nodeSize - 1,
    })
  })
  const activeBlockIndex =
    allBlocks.find((block) => from <= block.end && to >= block.start)?.index ?? null
  const center = activeBlockIndex ?? 0
  const start = Math.max(0, center - 5)
  const blocks = allBlocks
    .slice(start, start + 11)
    .map(({ start: _start, end: _end, ...block }) => block)

  return { from, to, empty, activeBlockIndex, totalBlocks: allBlocks.length, blocks }
}
