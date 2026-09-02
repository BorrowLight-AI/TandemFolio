import { Extension } from '@tiptap/core'
import type { Editor, Range } from '@tiptap/core'
import { Suggestion } from '@tiptap/suggestion'
import type { SuggestionProps } from '@tiptap/suggestion'
import type { StringKey } from '../i18n/locale'
import { t } from '../i18n/locale'
import {
  getMarkdownTextBlockIndexAtSelection,
  setMarkdownBlockType,
  type MarkdownBlockType,
} from './block-type-actions'
import { setMarkdownListType, type MarkdownListType } from './list-actions'
import { insertMarkdownDivider, insertMarkdownTable } from './structure-actions'

export interface SlashItem {
  id: string
  labelKey: StringKey
  /** extra match terms besides id and the localized label */
  keywords: string[]
  run: (editor: Editor, range: Range) => void
}

export interface SlashMenuState {
  items: SlashItem[]
  clientRect: DOMRect | null
  command: (item: SlashItem) => void
}

/** App-side sink for the suggestion lifecycle; the React menu renders from it */
export interface SlashController {
  onOpen(state: SlashMenuState): void
  onUpdate(state: SlashMenuState): void
  onKeyDown(event: KeyboardEvent): boolean
  onClose(): void
}

function chain(editor: Editor, range: Range) {
  return editor.chain().focus().deleteRange(range)
}

function applySlashBlockType(editor: Editor, range: Range, type: MarkdownBlockType): void {
  editor.chain().focus().deleteRange(range).run()
  const textBlockIndex = getMarkdownTextBlockIndexAtSelection(editor)
  if (textBlockIndex !== null) setMarkdownBlockType(editor, { textBlockIndex, type })
}

function applySlashListType(editor: Editor, range: Range, type: MarkdownListType): void {
  editor.chain().focus().deleteRange(range).run()
  const textBlockIndex = getMarkdownTextBlockIndexAtSelection(editor)
  if (textBlockIndex !== null) setMarkdownListType(editor, { textBlockIndex, type })
}

function applySlashTable(editor: Editor, range: Range): void {
  editor.chain().focus().deleteRange(range).run()
  insertMarkdownTable(editor, {
    position: editor.state.selection.from,
    rows: 3,
    columns: 3,
    headerRow: true,
  })
}

function applySlashDivider(editor: Editor, range: Range): void {
  editor.chain().focus().deleteRange(range).run()
  insertMarkdownDivider(editor, { position: editor.state.selection.from })
}

export function buildSlashItems(extra?: { insertImage?: () => void }): SlashItem[] {
  const items: SlashItem[] = [
    {
      id: 'paragraph',
      labelKey: 'styleParagraph',
      keywords: ['text', 'p'],
      run: (e, r) => applySlashBlockType(e, r, 'paragraph'),
    },
    {
      id: 'h1',
      labelKey: 'styleH1',
      keywords: ['heading', '#'],
      run: (e, r) => applySlashBlockType(e, r, 'heading_1'),
    },
    {
      id: 'h2',
      labelKey: 'styleH2',
      keywords: ['heading', '##'],
      run: (e, r) => applySlashBlockType(e, r, 'heading_2'),
    },
    {
      id: 'h3',
      labelKey: 'styleH3',
      keywords: ['heading', '###'],
      run: (e, r) => applySlashBlockType(e, r, 'heading_3'),
    },
    {
      id: 'bullet',
      labelKey: 'bulletList',
      keywords: ['list', 'ul', '-'],
      run: (e, r) => applySlashListType(e, r, 'bullet'),
    },
    {
      id: 'ordered',
      labelKey: 'orderedList',
      keywords: ['list', 'ol', '1.'],
      run: (e, r) => applySlashListType(e, r, 'ordered'),
    },
    {
      id: 'task',
      labelKey: 'taskList',
      keywords: ['todo', 'checkbox', '[]'],
      run: (e, r) => applySlashListType(e, r, 'task'),
    },
    {
      id: 'quote',
      labelKey: 'styleQuote',
      keywords: ['blockquote', '>'],
      run: (e, r) => applySlashBlockType(e, r, 'quote'),
    },
    {
      id: 'code',
      labelKey: 'styleCodeBlock',
      keywords: ['codeblock', '```'],
      run: (e, r) => applySlashBlockType(e, r, 'code_block'),
    },
    {
      id: 'table',
      labelKey: 'insertTable',
      keywords: ['grid'],
      run: applySlashTable,
    },
    {
      id: 'hr',
      labelKey: 'insertHr',
      keywords: ['divider', 'rule', '---'],
      run: applySlashDivider,
    },
  ]
  if (extra?.insertImage) {
    items.push({
      id: 'image',
      labelKey: 'insertImage',
      keywords: ['picture', 'img', 'photo'],
      run: (e, r) => {
        chain(e, r).run()
        extra.insertImage!()
      },
    })
  }
  return items
}

export function filterSlashItems(items: SlashItem[], query: string): SlashItem[] {
  const q = query.trim().toLowerCase()
  if (!q) return items
  return items.filter(
    (item) =>
      item.id.includes(q) ||
      t(item.labelKey).toLowerCase().includes(q) ||
      item.keywords.some((k) => k.includes(q)),
  )
}

interface SlashOptions {
  controller: SlashController | null
  items: () => SlashItem[]
}

export const SlashCommand = Extension.create<SlashOptions>({
  name: 'slashCommand',

  addOptions() {
    return { controller: null, items: () => [] }
  },

  addProseMirrorPlugins() {
    const getController = () => this.options.controller
    const toState = (props: SuggestionProps<SlashItem>): SlashMenuState => ({
      items: props.items,
      clientRect: props.clientRect?.() ?? null,
      command: (item) => props.command(item),
    })
    return [
      Suggestion<SlashItem>({
        editor: this.editor,
        char: '/',
        pluginKey: undefined,
        allowSpaces: false,
        command: ({ editor, range, props }) => props.run(editor, range),
        items: ({ query }) => filterSlashItems(this.options.items(), query),
        render: () => ({
          onStart: (props) => getController()?.onOpen(toState(props)),
          onUpdate: (props) => getController()?.onUpdate(toState(props)),
          onKeyDown: (props) => getController()?.onKeyDown(props.event) ?? false,
          onExit: () => getController()?.onClose(),
        }),
      }),
    ]
  },
})
