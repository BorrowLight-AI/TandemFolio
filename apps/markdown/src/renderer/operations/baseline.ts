export type MarkdownCommandDisposition =
  'typed-operation' | 'typed-ingress' | 'native-input' | 'missing'

export interface MarkdownRetainedCommandAuditEntry {
  readonly family: string
  readonly retainedSources: readonly string[]
  readonly disposition: MarkdownCommandDisposition
  readonly operationIds: readonly string[]
  readonly rationale: string
}

/** Machine-checked mapping from retained Markdown command producers to executable routes. */
export const markdownRetainedCommandAudit = [
  {
    family: 'document-ingress',
    retainedSources: ['Ribbon.onOpen', 'browser-files.openMarkdownFile'],
    disposition: 'typed-ingress',
    operationIds: ['markdown.document.load_staged'],
    rationale: 'The UI picker and office_open_local_file converge on the same loadText service.',
  },
  {
    family: 'direct-text-and-selection',
    retainedSources: ['ProseMirror input', 'clipboard text paste/drop', 'native selection'],
    disposition: 'native-input',
    operationIds: [
      'markdown.selection.set',
      'markdown.text.insert',
      'markdown.text.replace_selection',
    ],
    rationale:
      'Finite Agent primitives reproduce arbitrary selection, insertion, deletion, and replacement post-state.',
  },
  {
    family: 'inline-marks',
    retainedSources: ['Ribbon mark controls', 'BlockKeymap Mod-k'],
    disposition: 'typed-operation',
    operationIds: ['markdown.text.set_marks'],
    rationale: 'All retained mark gestures converge on one explicit range/final-state action.',
  },
  {
    family: 'block-shape-and-order',
    retainedSources: ['Ribbon block style', 'Slash menu', 'BlockKeymap', 'BlockDragHandle'],
    disposition: 'typed-operation',
    operationIds: [
      'markdown.block.set_type',
      'markdown.block.update',
      'markdown.code_block.set_language',
    ],
    rationale: 'Addressed block mutations use the shared ProseMirror transaction seams.',
  },
  {
    family: 'lists',
    retainedSources: ['Ribbon list controls', 'Slash menu'],
    disposition: 'typed-operation',
    operationIds: ['markdown.list.set_type'],
    rationale: 'Toggle gestures are normalized to one explicit addressed final list state.',
  },
  {
    family: 'tables-and-divider',
    retainedSources: ['Ribbon insert controls', 'Slash menu', 'TableMenu'],
    disposition: 'typed-operation',
    operationIds: ['markdown.table.insert', 'markdown.table.update', 'markdown.divider.insert'],
    rationale: 'Insertion and relative table mutations use bounded shared actions.',
  },
  {
    family: 'images',
    retainedSources: ['Ribbon picker', 'Slash menu', 'clipboard image paste/drop'],
    disposition: 'typed-operation',
    operationIds: ['markdown.image.insert', 'markdown.image.insert_staged'],
    rationale: 'Public paths become bounded staged bytes before the shared image action runs.',
  },
  {
    family: 'frontmatter',
    retainedSources: ['FrontmatterPanel'],
    disposition: 'typed-operation',
    operationIds: ['markdown.frontmatter.set'],
    rationale: 'UI and Registry replace the complete raw YAML envelope through one service.',
  },
  {
    family: 'history',
    retainedSources: ['Ribbon undo/redo', 'native shortcuts'],
    disposition: 'typed-operation',
    operationIds: ['markdown.history.undo', 'markdown.history.redo'],
    rationale: 'Both producers use the mounted TipTap history stack.',
  },
  {
    family: 'persistence-output-and-preferences',
    retainedSources: ['Ribbon/QAT', 'Shift-Save shortcut', 'autosave timer'],
    disposition: 'typed-operation',
    operationIds: [
      'markdown.document.save',
      'markdown.document.save_as',
      'markdown.document.export_docx',
      'markdown.document.open_print_dialog',
      'markdown.document.set_auto_save',
    ],
    rationale:
      'UI and Registry share the same browser persistence, export, print, and preference services.',
  },
] as const satisfies readonly MarkdownRetainedCommandAuditEntry[]
