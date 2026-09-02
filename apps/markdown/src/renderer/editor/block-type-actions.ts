import type { Editor } from '@tiptap/core'
import type { Node as ProseMirrorNode } from '@tiptap/pm/model'

export const markdownBlockTypes = [
  'paragraph',
  'heading_1',
  'heading_2',
  'heading_3',
  'heading_4',
  'heading_5',
  'heading_6',
  'quote',
  'code_block',
] as const

export type MarkdownBlockType = (typeof markdownBlockTypes)[number]

export type MarkdownBlockTypeResult =
  | { readonly ok: true; readonly changed: boolean }
  | { readonly ok: false; readonly message: string }

interface AddressedTextBlock {
  readonly position: number
  readonly node: ProseMirrorNode
  readonly currentType: MarkdownBlockType | 'wrapped'
  readonly quote?: {
    readonly position: number
    readonly node: ProseMirrorNode
    readonly childIndex: number
  }
}

function addressableTextBlocks(editor: Editor): AddressedTextBlock[] {
  const blocks: AddressedTextBlock[] = []
  editor.state.doc.descendants((node, position) => {
    if (!node.isTextblock) return
    const resolved = editor.state.doc.resolve(Math.min(position + 1, editor.state.doc.content.size))
    let wrapper: 'quote' | 'wrapped' | null = null
    let quote: AddressedTextBlock['quote']
    for (let depth = resolved.depth; depth > 0; depth -= 1) {
      const ancestor = resolved.node(depth)
      if (ancestor.type.name === 'blockquote') {
        wrapper = 'quote'
        quote = {
          position: resolved.before(depth),
          node: ancestor,
          childIndex: resolved.index(depth),
        }
        break
      }
      if (
        ancestor.type.name === 'bulletList' ||
        ancestor.type.name === 'orderedList' ||
        ancestor.type.name === 'taskList'
      ) {
        wrapper = 'wrapped'
        break
      }
    }
    const currentType: AddressedTextBlock['currentType'] = wrapper
      ? wrapper
      : node.type.name === 'heading'
        ? (`heading_${String(node.attrs.level)}` as MarkdownBlockType)
        : node.type.name === 'codeBlock'
          ? 'code_block'
          : 'paragraph'
    blocks.push({ position, node, currentType, ...(quote ? { quote } : {}) })
  })
  return blocks
}

export type MarkdownTextBlockAddressResult =
  | { readonly ok: true; readonly position: number }
  | { readonly ok: false; readonly message: string }

export function resolveMarkdownTextBlockPosition(
  editor: Editor,
  textBlockIndex: number,
): MarkdownTextBlockAddressResult {
  const blocks = addressableTextBlocks(editor)
  const target = blocks[textBlockIndex]
  if (!target) {
    return {
      ok: false,
      message: `Markdown text block ${textBlockIndex} is invalid for ${blocks.length} addressable text block(s).`,
    }
  }
  return { ok: true, position: target.position }
}

export function getMarkdownTextBlockIndexAtSelection(editor: Editor): number | null {
  const head = editor.state.selection.head
  let found: number | null = null
  let index = 0
  editor.state.doc.descendants((node, position) => {
    if (!node.isTextblock) return
    if (found === null && head >= position && head <= position + node.nodeSize) found = index
    index += 1
  })
  return found
}

export function getMarkdownTextBlockIndexAtPosition(
  editor: Editor,
  targetPosition: number,
): number | null {
  let found: number | null = null
  let index = 0
  editor.state.doc.descendants((node, position) => {
    if (!node.isTextblock) return
    if (
      found === null &&
      targetPosition >= position &&
      targetPosition <= position + node.nodeSize
    ) {
      found = index
    }
    index += 1
  })
  return found
}

/** Set one revision-scoped, document-order text block to an explicit final type. */
export function setMarkdownBlockType(
  editor: Editor,
  input: { readonly textBlockIndex: number; readonly type: MarkdownBlockType },
): MarkdownBlockTypeResult {
  const blocks = addressableTextBlocks(editor)
  const target = blocks[input.textBlockIndex]
  if (!target) {
    return {
      ok: false,
      message: `Markdown text block ${input.textBlockIndex} is invalid for ${blocks.length} addressable text block(s).`,
    }
  }
  if (target.currentType === input.type) return { ok: true, changed: false }

  const inlineNode = (() => {
    if (input.type === 'paragraph') {
      return editor.schema.nodes.paragraph!.create(null, target.node.content)
    }
    if (input.type === 'code_block') {
      return editor.schema.nodes.codeBlock!.create({ language: null }, target.node.content)
    }
    if (input.type.startsWith('heading_')) {
      return editor.schema.nodes.heading!.create(
        { level: Number(input.type.slice('heading_'.length)) },
        target.node.content,
      )
    }
    return editor.schema.nodes.paragraph!.create(null, target.node.content)
  })()

  if (target.quote && input.type !== 'quote') {
    if (target.quote.node.child(target.quote.childIndex) !== target.node) {
      return {
        ok: false,
        message: `Markdown text block ${input.textBlockIndex} is nested inside a quote and cannot be converted directly.`,
      }
    }
    const before: ProseMirrorNode[] = []
    const after: ProseMirrorNode[] = []
    target.quote.node.forEach((child, _offset, index) => {
      if (index < target.quote!.childIndex) before.push(child)
      if (index > target.quote!.childIndex) after.push(child)
    })
    const replacement: ProseMirrorNode[] = []
    if (before.length) replacement.push(editor.schema.nodes.blockquote!.create(null, before))
    replacement.push(inlineNode)
    if (after.length) replacement.push(editor.schema.nodes.blockquote!.create(null, after))
    const transaction = editor.state.tr.replaceWith(
      target.quote.position,
      target.quote.position + target.quote.node.nodeSize,
      replacement,
    )
    editor.view.dispatch(transaction.scrollIntoView())
    return { ok: true, changed: true }
  }

  if (target.currentType !== 'wrapped') {
    const replacement =
      input.type === 'quote'
        ? editor.schema.nodes.blockquote!.create(null, inlineNode)
        : inlineNode
    editor.view.dispatch(
      editor.state.tr
        .replaceWith(target.position, target.position + target.node.nodeSize, replacement)
        .scrollIntoView(),
    )
    return { ok: true, changed: true }
  }

  editor.commands.setTextSelection(Math.min(target.position + 1, editor.state.doc.content.size))
  let chain = editor.chain().focus()
  chain = chain.clearNodes()
  if (input.type === 'quote') chain = chain.setParagraph().setBlockquote()
  else if (input.type === 'paragraph') chain = chain.setParagraph()
  else if (input.type === 'code_block') chain = chain.setCodeBlock()
  else {
    chain = chain.setHeading({
      level: Number(input.type.slice('heading_'.length)) as 1 | 2 | 3 | 4 | 5 | 6,
    })
  }
  if (!chain.run()) {
    return {
      ok: false,
      message: `Markdown text block ${input.textBlockIndex} cannot be converted.`,
    }
  }
  return { ok: true, changed: true }
}
