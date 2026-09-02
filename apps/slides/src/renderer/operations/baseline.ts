export type PptxCommandDisposition = 'typed-operation' | 'typed-ingress' | 'host-effect' | 'missing'

export interface PptxRetainedCommandAuditEntry {
  readonly family: string
  readonly retainedSources: readonly string[]
  readonly disposition: PptxCommandDisposition
  readonly operationIds: readonly string[]
  readonly rationale: string
}

/** Machine-checked mapping from retained PPTX command producers to executable routes. */
export const pptxRetainedCommandAudit = [
  {
    family: 'document-lifecycle',
    retainedSources: ['File menu', 'QAT save', 'browser picker', 'MCP staged load'],
    disposition: 'typed-ingress',
    operationIds: [
      'pptx.document.create_blank',
      'pptx.document.load_staged',
      'pptx.document.save',
      'pptx.document.save_as',
    ],
    rationale: 'User and Agent lifecycle routes share one BrowserPresentation and package seam.',
  },
  {
    family: 'selection-and-history',
    retainedSources: ['Konva selection', 'QAT undo/redo', 'keyboard shortcuts'],
    disposition: 'typed-operation',
    operationIds: ['pptx.selection.set', 'pptx.history.undo', 'pptx.history.redo'],
    rationale: 'Selection and history use the mounted session state and native snapshot stack.',
  },
  {
    family: 'text-and-paragraphs',
    retainedSources: ['TextEditOverlay', 'Find/Replace', 'Ribbon font and paragraph controls'],
    disposition: 'typed-operation',
    operationIds: [
      'pptx.text.replace_all',
      'pptx.text.replace_selection',
      'pptx.text.set_font',
      'pptx.text.set_paragraphs',
      'pptx.text.set_vertical_anchor',
      'pptx.paragraph.set_format',
    ],
    rationale:
      'Top-level and explicitly parent-addressed group children share retained text mappers.',
  },
  {
    family: 'objects-and-connectors',
    retainedSources: ['Insert/Ribbon', 'canvas transforms', 'Arrange', 'clipboard', 'Format pane'],
    disposition: 'typed-operation',
    operationIds: [
      'pptx.connector.set_endpoints',
      'pptx.object.add',
      'pptx.object.copy_to',
      'pptx.object.delete',
      'pptx.object.duplicate',
      'pptx.object.group',
      'pptx.object.move_selection',
      'pptx.object.reorder',
      'pptx.object.set_fill',
      'pptx.object.set_flip',
      'pptx.object.set_image_fill',
      'pptx.object.set_stroke',
      'pptx.object.set_transform',
      'pptx.object.set_transforms',
      'pptx.object.ungroup',
    ],
    rationale:
      'All persisted object changes use explicit ids, bounded final state, and native undo.',
  },
  {
    family: 'slides-and-layout',
    retainedSources: ['thumbnail pane', 'New Slide menu', 'Design', 'Transitions', 'Header/Footer'],
    disposition: 'typed-operation',
    operationIds: [
      'pptx.slide.add_blank',
      'pptx.slide.add_with_layout',
      'pptx.slide.apply_header_footer',
      'pptx.slide.copy_to',
      'pptx.slide.delete',
      'pptx.slide.duplicate',
      'pptx.slide.move',
      'pptx.slide.set_advance_times',
      'pptx.slide.set_background',
      'pptx.slide.set_hidden',
      'pptx.slide.set_layout',
      'pptx.slide.set_size',
      'pptx.slide.set_transition',
    ],
    rationale:
      'Clipboard gestures are normalized to explicit source/destination operations for Agent replay.',
  },
  {
    family: 'images-media-and-ink',
    retainedSources: ['Insert image/media/3D', 'picture format', 'Ink toolbar'],
    disposition: 'typed-operation',
    operationIds: [
      'pptx.image.add_bytes',
      'pptx.image.replace_bytes',
      'pptx.ink.add',
      'pptx.media.add_bytes',
      'pptx.model3d.add_bytes',
      'pptx.picture.set_crop',
      'pptx.picture.set_opacity',
    ],
    rationale:
      'File pickers become bounded bytes before entering shared retained package primitives.',
  },
  {
    family: 'tables',
    retainedSources: ['Insert Table', 'table canvas handles', 'Table Design/Layout controls'],
    disposition: 'typed-operation',
    operationIds: [
      'pptx.table.add',
      'pptx.table.edit_structure',
      'pptx.table.merge_cells',
      'pptx.table.set_cell_anchor',
      'pptx.table.set_cell_content',
      'pptx.table.set_column_width',
      'pptx.table.set_row_height',
      'pptx.table.set_style',
    ],
    rationale: 'Table mutations share the retained table engine and package-level undo snapshots.',
  },
  {
    family: 'rich-insert-and-theme',
    retainedSources: ['Chart dialog', 'SmartArt gallery', 'Animations pane', 'Theme gallery'],
    disposition: 'typed-operation',
    operationIds: [
      'pptx.animation.set',
      'pptx.chart.add',
      'pptx.chart.update',
      'pptx.smartart.add',
      'pptx.theme.apply',
    ],
    rationale:
      'Bounded data contracts feed the retained chart, SmartArt, animation, and theme engines.',
  },
  {
    family: 'document-metadata',
    retainedSources: ['Comments pane', 'Notes pane', 'Sections', 'Link dialog'],
    disposition: 'typed-operation',
    operationIds: [
      'pptx.comment.add',
      'pptx.comment.delete',
      'pptx.hyperlink.set',
      'pptx.notes.set',
      'pptx.section.add',
      'pptx.section.move',
      'pptx.section.remove',
      'pptx.section.rename',
    ],
    rationale:
      'Addressed metadata mutations share the same package, revision, history, and save route.',
  },
  {
    family: 'master-editing',
    retainedSources: ['MasterView canvas and Format pane'],
    disposition: 'typed-operation',
    operationIds: [
      'pptx.master.object.delete',
      'pptx.master.object.set_fill',
      'pptx.master.object.set_stroke',
      'pptx.master.object.set_transform',
      'pptx.master.text.set_paragraphs',
    ],
    rationale:
      'Master and layout part paths plus object ids make all retained master edits replayable.',
  },
  {
    family: 'browser-output-and-show',
    retainedSources: ['Export PDF/images', 'Print', 'PresenterView', 'AudienceView'],
    disposition: 'host-effect',
    operationIds: [],
    rationale:
      'These routes do not mutate the presentation; browser download, print, and channel adapters own them.',
  },
] as const satisfies readonly PptxRetainedCommandAuditEntry[]
