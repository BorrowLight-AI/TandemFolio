import type { AnyExtension } from '@tiptap/core'
import StarterKit from '@tiptap/starter-kit'
import { Markdown } from '@tiptap/markdown'
import { TableCell, TableHeader, TableRow } from '@tiptap/extension-table'
import { TaskItem } from '@tiptap/extension-list'
import { CodeBlock } from '@tiptap/extension-code-block'
import { ReactNodeViewRenderer } from '@tiptap/react'
import { Placeholder } from '@tiptap/extensions'
import { CodeBlockView } from './CodeBlockView'
import { LocalImage } from './localImage'
import { BlockDragHandle } from './blockDragHandle'
import { BlockKeymap } from './blockKeymap'
import { SlashCommand } from './slashCommand'
import { LinearOrderedList, LinearTable, LinearTaskList } from './linearMarkdownExtensions'
import type { SlashController, SlashItem } from './slashCommand'
import { t } from '../i18n/locale'

export interface BuildExtensionsOptions {
  slashController: SlashController
  slashItems: () => SlashItem[]
}

export function buildExtensions(options: BuildExtensionsOptions): AnyExtension[] {
  return [
    StarterKit.configure({
      // LocalImage replaces the plain image; Link keeps the renderer's authored URL behavior.
      link: { openOnClick: false },
      // replaced by the NodeView-enhanced variant below (language picker + copy)
      codeBlock: false,
      // replaced by the linear-time Markdown tokenizer guard below
      orderedList: false,
      // underline would serialize as `++text++` — not part of GFM
      underline: false,
    }),
    CodeBlock.extend({
      addNodeView() {
        return ReactNodeViewRenderer(CodeBlockView)
      },
    }),
    // 4-space nesting: the default 2 spaces is below the content column of
    // ordered items ("1. " = 3), so strict CommonMark parsers (GitHub) would
    // flatten sub-lists in the saved file. 4 is safe for every marker width.
    Markdown.configure({ indentation: { style: 'space', size: 4 } }),
    // column widths are not expressible in GFM tables — no resizable columns;
    // the wrapper div gives wide tables a horizontal scrollbar
    LinearTable.configure({ resizable: false, renderWrapper: true }),
    TableCell,
    TableHeader,
    TableRow,
    LinearTaskList,
    TaskItem.configure({ nested: true }),
    LinearOrderedList,
    LocalImage,
    BlockDragHandle,
    BlockKeymap,
    Placeholder.configure({ placeholder: () => t('placeholder') }),
    SlashCommand.configure({
      controller: options.slashController,
      items: options.slashItems,
    }),
  ]
}
