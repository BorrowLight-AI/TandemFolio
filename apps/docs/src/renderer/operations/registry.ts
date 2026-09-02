import type { Editor } from '@tiptap/core'
import { validateJsonSchemaValue } from '@tandemfolio/operation-contract'

import {
  executeCommands,
  validateEnvelope,
  type ApplyListPreset,
  type CreateParagraphBullets,
  type CommandContext,
  type CommandEnvelope,
  type DeleteBlocks,
  type DeleteParagraphBullets,
  type InsertToc,
  type MoveBlocks,
  type SetParagraphDirection,
  type SetHeadingLevel,
  type SetListLevel,
  type UpdateImageProperties,
  type UpdateParagraphStyle,
  type UpdateTextStyle,
} from '../editor/commands'
import {
  clearDocxCharacterFormat,
  setDocxCharacterStyle,
  setDocxCharacterFormat,
  transformDocxTextCase,
  type DocxCharacterFormatInput,
  type DocxCharacterRangeInput,
  type DocxCharacterStyleInput,
  type DocxTextCaseInput,
} from '../editor/character-format'
import { setDocxTextLink, type SetDocxTextLinkInput } from '../editor/link-actions'
import { insertDocxCaption, type InsertDocxCaptionInput } from '../editor/caption-actions'
import {
  addDocxComment,
  deleteDocxComment,
  replyDocxComment,
  setDocxCommentResolved,
  type AddDocxCommentInput,
  type DeleteDocxCommentInput,
  type ReplyDocxCommentInput,
  type SetDocxCommentResolvedInput,
} from '../editor/comments'
import {
  insertDocxIndex,
  markDocxIndexEntry,
  type InsertDocxIndexInput,
  type MarkDocxIndexEntryInput,
} from '../editor/index-actions'
import { insertDocxText, type InsertDocxTextInput } from '../editor/text-actions'
import {
  deleteDocxNote,
  insertDocxNote,
  updateDocxNote,
  type DeleteDocxNoteInput,
  type DocxNoteKind,
  type InsertDocxNoteInput,
  type UpdateDocxNoteInput,
} from '../editor/note-actions'
import {
  insertDocxBibliography,
  insertDocxCitation,
  upsertDocxSource,
  type InsertDocxBibliographyInput,
  type InsertDocxCitationInput,
  type UpsertDocxSourceInput,
} from '../editor/source-actions'
import { refreshDocxToc, type RefreshDocxTocInput } from '../editor/toc-actions'
import { setDocxBookmark, type SetDocxBookmarkInput } from '../editor/bookmark-actions'
import {
  insertDocxCoverPage,
  type InsertDocxCoverPageInput,
} from '../editor/cover-pages'
import {
  insertDocxCrossReference,
  insertDocxField,
  updateDocxFields,
  type InsertDocxCrossReferenceInput,
  type InsertDocxFieldInput,
  type UpdateDocxFieldsInput,
} from '../editor/field-actions'
import type {
  ContinueNumberingInput,
  NumberingActionResult,
  RestartNumberingInput,
} from '../numbering-actions'
import {
  deleteTableColumns,
  deleteTableRows,
  deleteTopLevelTable,
  insertTableColumns,
  insertTableRows,
  insertTopLevelTableAfterBlock,
  mergeTableCells,
  setTableCellBorders,
  setTableCellFormat,
  setTableRowHeight,
  setTableStyle,
  splitTableCell,
  type DeleteTableColumnsInput,
  type DeleteTableRowsInput,
  type InsertTableColumnsInput,
  type InsertTableInput,
  type InsertTableRowsInput,
  type MergeTableCellsInput,
  type SetTableCellBordersInput,
  type SetTableCellFormatInput,
  type SetTableRowHeightInput,
  type SetTableStyleInput,
  type SplitTableCellInput,
} from '../editor/table-actions'
import { setTableColumnWidths, type SetTableColumnWidthsInput } from '../editor/table-sizing'
import { insertPageBreakAfterBlock, type InsertPageBreakInput } from '../editor/page-actions'
import { setDocxDropCap, type SetDocxDropCapInput } from '../editor/drop-cap-actions'
import {
  insertDocxWordArtAfterBlock,
  type InsertDocxWordArtInput,
} from '../editor/wordart-actions'
import {
  insertDocxLineAfterBlock,
  insertDocxShapeAfterBlock,
  type InsertDocxLineInput,
  type InsertDocxShapeInput,
} from '../editor/shape-actions'
import {
  insertDocxTextboxAfterBlock,
  setDocxTextboxContentAtBlock,
  type InsertDocxTextboxInput,
  type SetDocxTextboxContentInput,
} from '../editor/textbox-actions'
import {
  insertDocxChartAfterBlock,
  updateDocxChartAtBlock,
  type InsertDocxChartInput,
  type UpdateDocxChartInput,
} from '../editor/chart-actions'
import {
  insertDocxEquation,
  updateDocxEquation,
  type InsertDocxEquationInput,
  type UpdateDocxEquationInput,
} from '../editor/equation-actions'
import {
  removeDocxObjectAtBlock,
  setDocxObjectOffsetAtBlock,
  setDocxObjectSizeAtBlock,
  setDocxObjectStyleAtBlock,
  type SetDocxObjectOffsetInput,
  type SetDocxObjectSizeInput,
  type SetDocxObjectStyleInput,
  type RemoveDocxObjectInput,
} from '../editor/object-actions'
import {
  insertDocxImageAfterBlock,
  removeDocxImageAtBlock,
  replaceDocxImageAtBlock,
  setDocxImageCropAtBlock,
  setDocxImageMarginPositionAtBlock,
  setDocxImageOffsetPositionAtBlock,
  setDocxImageTransformAtBlock,
  setDocxImageWrapAtBlock,
  stagedDocxImageMediaType,
  type InsertDocxImageInput,
  type ReplaceDocxImageInput,
  type SetDocxImageCropInput,
  type SetDocxImageMarginPositionInput,
  type SetDocxImageOffsetPositionInput,
  type SetDocxImageTransformInput,
  type SetDocxImageWrapInput,
} from '../editor/image-actions'
import {
  insertSectionBreakAfterBlock,
  type InsertSectionBreakInput,
  type ResolveSectionBreakSource,
} from '../editor/section-actions'
import {
  setDocumentDifferentOddEvenPages,
  setHeaderFooterText,
  setHeaderFooterPageNumber,
  setHeaderFooterParagraphs,
  setSectionColumns,
  setSectionDifferentFirstPage,
  setSectionMargins,
  setSectionOrientation,
  setSectionPageBorder,
  setSectionPageNumbering,
  setSectionPageSize,
  type SetDocumentDifferentOddEvenPagesInput,
  type SetHeaderFooterTextInput,
  type SetHeaderFooterPageNumberInput,
  type SetHeaderFooterParagraphsInput,
  type SetSectionColumnsInput,
  type SetSectionDifferentFirstPageInput,
  type SetSectionMarginsInput,
  type SetSectionOrientationInput,
  type SetSectionPageBorderInput,
  type SetSectionPageNumberingInput,
  type SetSectionPageSizeInput,
} from '../editor/section-layout'
import {
  applyDocxRevisionDecision,
  setDocxRevisionTracking,
  type ApplyDocxRevisionDecisionInput,
  type SetDocxRevisionTrackingInput,
} from '../editor/revisions'
import { setDocxProtection, type SetDocxProtectionInput } from '../editor/protection-actions'
import { applyDocxInk, type ApplyDocxInkInput, type InkAnnotation } from '../editor/ink'
import { compareDocxBytes, type CompareEntry } from '../editor/compare'
import {
  applyDocxDocumentDesign,
  type DocxDocumentDesignField,
  type DocxDocumentDesignState,
  type SetDocxDocumentDesignInput,
} from '../editor/document-design'
import type {
  Block,
  CommentInfo,
  DocProtection,
  HeaderFooter,
  NoteInfo,
  SectionInfo,
  SourceInfo,
} from '@genoffice/docx-engine'
import { docxOperationCatalog } from './catalog'

type DocxOperationDescriptor = (typeof docxOperationCatalog.operations)[number]
type DocxOperationId = DocxOperationDescriptor['id']

interface DocxOperationCommand {
  readonly operation: string
  readonly arguments: Record<string, unknown>
}

export interface DocxOperationServices {
  readonly allocateListNumId?: (kind: 'bullet' | 'ordered') => string | null
  readonly createListDef?: (levels: ApplyListPreset['levels']) => string | null
  readonly continueList?: (input: ContinueNumberingInput) => NumberingActionResult
  readonly hasCharacterStyle?: (styleId: string) => boolean
  readonly hasTableStyle?: (styleId: string) => boolean
  readonly resolveSectionBreakSource?: ResolveSectionBreakSource
  readonly getSectionLayoutState?: () => readonly SectionInfo[]
  readonly getDifferentOddEvenPages?: () => boolean
  readonly getSourceBlocks?: () => readonly Block[]
  readonly getNotes?: (kind: DocxNoteKind) => readonly NoteInfo[]
  readonly commitInsertedNote?: (kind: DocxNoteKind, id: string) => void
  readonly commitUpdatedNote?: (kind: DocxNoteKind, id: string, original: NoteInfo) => void
  readonly commitDeletedNote?: (kind: DocxNoteKind, id: string, original: NoteInfo) => void
  readonly getSources?: () => readonly SourceInfo[]
  readonly commitSources?: (sources: readonly SourceInfo[]) => void
  readonly getComments?: () => readonly CommentInfo[]
  readonly commitComments?: (comments: readonly CommentInfo[]) => void
  readonly commitRevisionTracking?: (enabled: boolean) => void
  readonly getProtection?: () => DocProtection | null
  readonly commitProtection?: (protection: DocProtection | null) => void
  readonly getInkAnnotations?: () => readonly InkAnnotation[]
  readonly commitInkAnnotations?: (annotations: readonly InkAnnotation[]) => void
  readonly commitCompareResult?: (result: {
    readonly otherName: string
    readonly entries: CompareEntry[]
  }) => void
  readonly getDocumentDesign?: () => DocxDocumentDesignState
  readonly commitDocumentDesign?: (
    state: DocxDocumentDesignState,
    fields: readonly DocxDocumentDesignField[],
  ) => void
  readonly getHeaderFooterValue?: (input: {
    readonly sectionIndex: number
    readonly kind: 'header' | 'footer'
    readonly variant: 'default' | 'first' | 'even'
  }) => HeaderFooter | null
  readonly restartList?: (input: RestartNumberingInput) => NumberingActionResult
  readonly loadStaged: (input: {
    readonly blobId: string
    readonly name: string
    readonly data: ArrayBuffer
  }) => void | Promise<void>
  readonly save: () => Promise<{ readonly fileName: string }>
}

export type DocxOperationExecution =
  | { readonly handled: false }
  | {
      readonly handled: true
      readonly operationId: DocxOperationId
      readonly ok: true
      readonly output?: Readonly<Record<string, unknown>>
      readonly checkpointRecovery?: false
    }
  | {
      readonly handled: true
      readonly operationId: DocxOperationId
      readonly ok: false
      readonly error: 'invalid_arguments' | 'execution_failed'
      readonly message: string
    }

interface DocxOperationHandlerSuccess {
  readonly ok: true
  readonly output?: Readonly<Record<string, unknown>>
  readonly checkpointRecovery?: false
}

interface DocxOperationHandlerFailure {
  readonly ok: false
  readonly error: 'invalid_arguments' | 'execution_failed'
  readonly message: string
}

type DocxOperationHandlerResult = DocxOperationHandlerSuccess | DocxOperationHandlerFailure

type DocxOperationHandler = (
  editor: Editor,
  arguments_: Record<string, unknown>,
  services: DocxOperationServices,
) => DocxOperationHandlerResult | Promise<DocxOperationHandlerResult>

function isArrayBuffer(value: unknown): value is ArrayBuffer {
  return Object.prototype.toString.call(value) === '[object ArrayBuffer]'
}

interface NativeEnvelopeOptions {
  readonly failureMessage: string
  readonly context?: CommandContext
  readonly failureError?:
    | DocxOperationHandlerFailure['error']
    | ((message: string) => DocxOperationHandlerFailure['error'])
  readonly includeSkippedDeleted?: boolean
}

function executeNativeEnvelope(
  editor: Editor,
  envelope: CommandEnvelope,
  options: NativeEnvelopeOptions,
): DocxOperationHandlerResult {
  const validationError = validateEnvelope(envelope)
  if (validationError) {
    return { ok: false, error: 'invalid_arguments', message: validationError }
  }

  const outcome = executeCommands(editor, envelope, options.context)
  if (!outcome.ok) {
    const message = outcome.error ?? options.failureMessage
    const failureError = options.failureError ?? 'execution_failed'
    return {
      ok: false,
      error: typeof failureError === 'function' ? failureError(message) : failureError,
      message,
    }
  }

  const result = outcome.results[0]
  const output: Record<string, unknown> = {
    summary: outcome.summary,
    matched: result.matched,
    changed: result.changed,
    skippedProtected: result.skippedProtected,
  }
  if (options.includeSkippedDeleted) output.skippedDeleted = result.skippedDeleted ?? 0
  return { ok: true, output }
}

const handlers = {
  'docx.text.insert': (editor, arguments_) => {
    const result = insertDocxText(editor, arguments_ as unknown as InsertDocxTextInput)
    if (!result.ok) return result
    return { ok: true }
  },
  'docx.text.replace_selection': (editor, arguments_) => {
    if (
      !editor
        .chain()
        .focus()
        .deleteSelection()
        .insertContent(arguments_.text as string)
        .run()
    ) {
      return {
        ok: false,
        error: 'execution_failed',
        message: 'The DOCX editor rejected selection replacement.',
      }
    }
    return { ok: true }
  },
  'docx.text.set_link': (editor, arguments_) => {
    const input = arguments_ as unknown as SetDocxTextLinkInput
    const result = setDocxTextLink(editor, input)
    if (!result.ok) return result
    return {
      ok: true,
      output: {
        summary: `${input.href === null ? 'Removed' : 'Set'} DOCX link over range ${result.from}..${result.to}`,
        from: result.from,
        to: result.to,
        href: input.href?.trim() ?? null,
        changed: result.changed,
      },
    }
  },
  'docx.note.insert': (editor, arguments_, services) => {
    if (!services.getNotes || !services.commitInsertedNote) {
      return {
        ok: false,
        error: 'execution_failed',
        message: 'The mounted DOCX session does not expose shared note state.',
      }
    }
    const input = arguments_ as unknown as InsertDocxNoteInput
    const result = insertDocxNote(editor, services.getNotes(input.kind), input)
    if (!result.ok) return result
    services.commitInsertedNote(result.kind, result.noteId)
    return {
      ok: true,
      output: {
        summary: `Inserted DOCX ${result.kind} ${result.noteId} at range ${result.from}..${result.to}`,
        from: result.from,
        to: result.to,
        kind: result.kind,
        noteId: result.noteId,
        number: result.number,
        changed: result.changed,
      },
    }
  },
  'docx.source.upsert': (editor, arguments_, services) => {
    if (!services.getSources || !services.commitSources) {
      return {
        ok: false,
        error: 'execution_failed',
        message: 'The mounted DOCX session does not expose shared bibliography state.',
      }
    }
    const result = upsertDocxSource(
      editor,
      services.getSources(),
      arguments_ as unknown as UpsertDocxSourceInput,
    )
    if (!result.ok) return result
    if (result.changed) services.commitSources(result.sources)
    return {
      ok: true,
      output: {
        summary: `${result.created ? 'Added' : result.changed ? 'Updated' : 'Kept'} DOCX source ${result.tag}`,
        tag: result.tag,
        created: result.created,
        changed: result.changed,
      },
    }
  },
  'docx.bibliography.insert': (editor, arguments_, services) => {
    if (!services.getSources) {
      return {
        ok: false,
        error: 'execution_failed',
        message: 'The mounted DOCX session does not expose shared bibliography state.',
      }
    }
    const result = insertDocxBibliography(
      editor,
      services.getSources(),
      arguments_ as unknown as InsertDocxBibliographyInput,
    )
    if (!result.ok) return result
    return {
      ok: true,
      output: {
        summary: `Inserted ${result.entries} DOCX bibliography entries after block ${result.afterBlockIndex}`,
        afterBlockIndex: result.afterBlockIndex,
        entries: result.entries,
        insertedBlocks: result.insertedBlocks,
      },
    }
  },
  'docx.caption.insert': (editor, arguments_) => {
    const result = insertDocxCaption(editor, arguments_ as unknown as InsertDocxCaptionInput)
    if (!result.ok) return result
    return {
      ok: true,
      output: {
        summary: `Inserted DOCX caption ${result.label} ${result.number} after block ${result.afterBlockIndex}`,
        afterBlockIndex: result.afterBlockIndex,
        label: result.label,
        number: result.number,
        changed: result.changed,
      },
    }
  },
  'docx.comment.add': (editor, arguments_, services) => {
    if (!services.getComments || !services.commitComments) {
      return {
        ok: false,
        error: 'execution_failed',
        message: 'The mounted DOCX session does not expose shared comment state.',
      }
    }
    const result = addDocxComment(
      editor,
      services.getComments(),
      arguments_ as unknown as AddDocxCommentInput,
    )
    if (!result.ok) return result
    services.commitComments(result.comments)
    return {
      ok: true,
      output: {
        summary: `Added DOCX comment ${result.id} at range ${result.from}..${result.to}`,
        id: result.id,
        from: result.from,
        to: result.to,
        changed: result.changed,
      },
    }
  },
  'docx.comment.delete': (editor, arguments_, services) => {
    if (!services.getComments || !services.commitComments) {
      return {
        ok: false,
        error: 'execution_failed',
        message: 'The mounted DOCX session does not expose shared comment state.',
      }
    }
    const result = deleteDocxComment(
      editor,
      services.getComments(),
      arguments_ as unknown as DeleteDocxCommentInput,
    )
    if (!result.ok) return result
    services.commitComments(result.comments)
    return {
      ok: true,
      output: {
        summary: `Deleted ${result.deleted} DOCX comment record(s) for ${result.id}`,
        id: result.id,
        deleted: result.deleted,
        anchors: result.anchors,
        changed: result.changed,
      },
    }
  },
  'docx.comment.reply': (editor, arguments_, services) => {
    if (!services.getComments || !services.commitComments) {
      return {
        ok: false,
        error: 'execution_failed',
        message: 'The mounted DOCX session does not expose shared comment state.',
      }
    }
    const result = replyDocxComment(
      editor,
      services.getComments(),
      arguments_ as unknown as ReplyDocxCommentInput,
    )
    if (!result.ok) return result
    services.commitComments(result.comments)
    return {
      ok: true,
      output: {
        summary: `Added DOCX comment reply ${result.id} to parent ${result.parentId}`,
        id: result.id,
        parentId: result.parentId,
        references: result.references,
        changed: result.changed,
      },
    }
  },
  'docx.comment.set_resolved': (editor, arguments_, services) => {
    if (!services.getComments || !services.commitComments) {
      return {
        ok: false,
        error: 'execution_failed',
        message: 'The mounted DOCX session does not expose shared comment state.',
      }
    }
    const result = setDocxCommentResolved(
      editor,
      services.getComments(),
      arguments_ as unknown as SetDocxCommentResolvedInput,
    )
    if (!result.ok) return result
    if (result.changed) services.commitComments(result.comments)
    return {
      ok: true,
      output: {
        summary: `${result.resolved ? 'Resolved' : 'Reopened'} DOCX comment thread ${result.id}`,
        id: result.id,
        resolved: result.resolved,
        affected: result.affected,
        changed: result.changed,
      },
    }
  },
  'docx.index.mark': (editor, arguments_) => {
    const result = markDocxIndexEntry(editor, arguments_ as unknown as MarkDocxIndexEntryInput)
    if (!result.ok) return result
    return {
      ok: true,
      output: {
        summary: `Marked DOCX index term ${result.term} at range ${result.from}..${result.to}`,
        from: result.from,
        to: result.to,
        term: result.term,
        changed: result.changed,
      },
    }
  },
  'docx.index.insert': (editor, arguments_) => {
    const result = insertDocxIndex(editor, arguments_ as unknown as InsertDocxIndexInput)
    if (!result.ok) return result
    return {
      ok: true,
      output: {
        summary: `Inserted DOCX index with ${result.entries} entries after block ${result.afterBlockIndex}`,
        afterBlockIndex: result.afterBlockIndex,
        entries: result.entries,
        insertedBlocks: result.insertedBlocks,
      },
    }
  },
  'docx.citation.insert': (editor, arguments_, services) => {
    if (!services.getSources) {
      return {
        ok: false,
        error: 'execution_failed',
        message: 'The mounted DOCX session does not expose shared bibliography state.',
      }
    }
    const result = insertDocxCitation(
      editor,
      services.getSources(),
      arguments_ as unknown as InsertDocxCitationInput,
    )
    if (!result.ok) return result
    return {
      ok: true,
      output: {
        summary: `Inserted DOCX citation for source ${result.sourceTag}`,
        from: result.from,
        to: result.to,
        sourceTag: result.sourceTag,
        changed: result.changed,
      },
    }
  },
  'docx.note.delete': (editor, arguments_, services) => {
    if (!services.getNotes || !services.commitDeletedNote) {
      return {
        ok: false,
        error: 'execution_failed',
        message: 'The mounted DOCX session does not expose shared note state.',
      }
    }
    const input = arguments_ as unknown as DeleteDocxNoteInput
    const result = deleteDocxNote(editor, services.getNotes(input.kind), input)
    if (!result.ok) return result
    services.commitDeletedNote(result.kind, result.noteId, result.original)
    return {
      ok: true,
      output: {
        summary: `Deleted DOCX ${result.kind} ${result.noteId}`,
        kind: result.kind,
        noteId: result.noteId,
        references: result.references,
        renumbered: result.renumbered,
        changed: result.changed,
      },
    }
  },
  'docx.note.update': (editor, arguments_, services) => {
    if (!services.getNotes || !services.commitUpdatedNote) {
      return {
        ok: false,
        error: 'execution_failed',
        message: 'The mounted DOCX session does not expose shared note state.',
      }
    }
    const input = arguments_ as unknown as UpdateDocxNoteInput
    const result = updateDocxNote(editor, services.getNotes(input.kind), input)
    if (!result.ok) return result
    if (result.changed) services.commitUpdatedNote(result.kind, result.noteId, result.original)
    return {
      ok: true,
      output: {
        summary: `${result.changed ? 'Updated' : 'Kept'} DOCX ${result.kind} ${result.noteId}`,
        kind: result.kind,
        noteId: result.noteId,
        references: result.references,
        changed: result.changed,
      },
    }
  },
  'docx.bookmark.set': (editor, arguments_) => {
    const result = setDocxBookmark(editor, arguments_ as unknown as SetDocxBookmarkInput)
    if (!result.ok) return result
    return {
      ok: true,
      output: {
        summary: `${result.enabled ? 'Enabled' : 'Disabled'} DOCX bookmark ${result.name} on block ${result.blockIndex}`,
        blockIndex: result.blockIndex,
        name: result.name,
        enabled: result.enabled,
        changed: result.changed,
      },
    }
  },
  'docx.cross_reference.insert': (editor, arguments_) => {
    const result = insertDocxCrossReference(
      editor,
      arguments_ as unknown as InsertDocxCrossReferenceInput,
    )
    if (!result.ok) return result
    return {
      ok: true,
      output: {
        summary: `Inserted DOCX cross-reference to ${result.bookmarkName} over range ${result.from}..${result.to}`,
        from: result.from,
        to: result.to,
        bookmarkName: result.bookmarkName,
        changed: result.changed,
      },
    }
  },
  'docx.field.insert': (editor, arguments_) => {
    const result = insertDocxField(editor, arguments_ as unknown as InsertDocxFieldInput)
    if (!result.ok) return result
    return {
      ok: true,
      output: {
        summary: `Inserted DOCX ${result.instruction} field over range ${result.from}..${result.to}`,
        from: result.from,
        to: result.to,
        instruction: result.instruction,
        changed: result.changed,
      },
    }
  },
  'docx.field.update': (editor, arguments_) => {
    const result = updateDocxFields(editor, arguments_ as unknown as UpdateDocxFieldsInput)
    if (!result.ok) return result
    return {
      ok: true,
      output: {
        summary: `Updated ${result.changed} DOCX field cache(s)`,
        matched: result.matched,
        changed: result.changed,
      },
    }
  },
  'docx.history.undo': (editor) => {
    if (!editor.can().undo() || !editor.chain().focus().undo().run()) {
      return {
        ok: false,
        error: 'execution_failed',
        message: 'docx.history.undo requires an available undo entry.',
      }
    }
    return { ok: true, output: { undone: true } }
  },
  'docx.revision.set_tracking': (editor, arguments_, services) => {
    if (!services.commitRevisionTracking) {
      return {
        ok: false,
        error: 'execution_failed',
        message: 'The mounted DOCX session does not expose shared revision-tracking state.',
      }
    }
    const result = setDocxRevisionTracking(
      editor,
      arguments_ as unknown as SetDocxRevisionTrackingInput,
    )
    if (!result.ok) return result
    services.commitRevisionTracking(result.enabled)
    return {
      ok: true,
      output: {
        summary: `${result.enabled ? 'Enabled' : 'Disabled'} DOCX revision tracking`,
        enabled: result.enabled,
        changed: result.changed,
      },
      checkpointRecovery: false,
    }
  },
  'docx.revision.apply_decision': (editor, arguments_) => {
    const result = applyDocxRevisionDecision(
      editor,
      arguments_ as unknown as ApplyDocxRevisionDecisionInput,
    )
    if (!result.ok) return result
    return {
      ok: true,
      output: {
        summary: `${result.decision === 'accept' ? 'Accepted' : 'Rejected'} ${result.matched} DOCX revision(s) in ${result.scope} scope`,
        decision: result.decision,
        scope: result.scope,
        matched: result.matched,
        remaining: result.remaining,
        changed: result.changed,
      },
    }
  },
  'docx.history.redo': (editor) => {
    if (!editor.can().redo() || !editor.chain().focus().redo().run()) {
      return {
        ok: false,
        error: 'execution_failed',
        message: 'docx.history.redo requires an available redo entry.',
      }
    }
    return { ok: true, output: { redone: true } }
  },
  'docx.document.insert_page_break': (editor, arguments_) => {
    const input = arguments_ as unknown as InsertPageBreakInput
    const result = insertPageBreakAfterBlock(editor, input)
    if (!result.ok) return result
    return {
      ok: true,
      output: {
        summary: `Inserted DOCX page break at block ${result.insertedBlockIndex}`,
        insertedBlockIndex: result.insertedBlockIndex,
      },
    }
  },
  'docx.cover_page.insert': (editor, arguments_) => {
    const result = insertDocxCoverPage(editor, arguments_ as unknown as InsertDocxCoverPageInput)
    if (!result.ok) return result
    return {
      ok: true,
      output: {
        summary: `Inserted DOCX ${result.preset} cover page with ${result.insertedBlocks} blocks`,
        preset: result.preset,
        insertedBlocks: result.insertedBlocks,
        changed: result.changed,
      },
    }
  },
  'docx.document.set_protection': async (_editor, arguments_, services) => {
    if (!services.getProtection || !services.commitProtection) {
      return {
        ok: false,
        error: 'execution_failed',
        message: 'The mounted DOCX session does not expose shared protection state.',
      }
    }
    const result = await setDocxProtection(
      services.getProtection(),
      arguments_ as unknown as SetDocxProtectionInput,
    )
    if (!result.ok) return result
    if (result.changed) services.commitProtection(result.protection)
    return {
      ok: true,
      output: {
        summary: `${result.enabled ? 'Enabled' : 'Disabled'} DOCX document protection`,
        enabled: result.enabled,
        passwordProtected: result.passwordProtected,
        changed: result.changed,
      },
    }
  },
  'docx.document.set_design': (_editor, arguments_, services) => {
    if (!services.getDocumentDesign || !services.commitDocumentDesign) {
      return {
        ok: false,
        error: 'execution_failed',
        message: 'The mounted DOCX session does not expose shared document-design state.',
      }
    }
    const result = applyDocxDocumentDesign(
      services.getDocumentDesign(),
      arguments_ as unknown as SetDocxDocumentDesignInput,
    )
    if (!result.ok) return result
    if (result.changed) services.commitDocumentDesign(result.state, result.changedFields)
    return {
      ok: true,
      output: {
        summary: `Set DOCX document design fields: ${result.fields.join(', ')}`,
        fields: result.fields,
        changedFields: result.changedFields,
        changed: result.changed,
      },
    }
  },
  'docx.ink.apply': (_editor, arguments_, services) => {
    if (!services.getInkAnnotations || !services.commitInkAnnotations) {
      return {
        ok: false,
        error: 'execution_failed',
        message: 'The mounted DOCX session does not expose shared ink state.',
      }
    }
    const result = applyDocxInk(
      services.getInkAnnotations(),
      arguments_ as unknown as ApplyDocxInkInput,
    )
    if (!result.ok) return result
    if (result.changed) services.commitInkAnnotations(result.annotations)
    return {
      ok: true,
      output: {
        summary: `Applied DOCX ink ${result.action}: ${result.added} added, ${result.deleted} deleted`,
        action: result.action,
        added: result.added,
        deleted: result.deleted,
        count: result.count,
        changed: result.changed,
      },
    }
  },
  'docx.section.insert_break': (editor, arguments_, services) => {
    const input = arguments_ as unknown as InsertSectionBreakInput
    const result = insertSectionBreakAfterBlock(
      editor,
      input,
      services.resolveSectionBreakSource ?? (() => null),
    )
    if (!result.ok) return result
    return {
      ok: true,
      output: {
        summary: `Inserted ${result.startType} DOCX section break at block ${result.insertedBlockIndex}`,
        insertedBlockIndex: result.insertedBlockIndex,
        startType: result.startType,
      },
    }
  },
  'docx.section.set_orientation': (editor, arguments_, services) => {
    const input = arguments_ as unknown as SetSectionOrientationInput
    const result = setSectionOrientation(editor, input, services.getSectionLayoutState?.() ?? [])
    if (!result.ok) return result
    return {
      ok: true,
      output: {
        sectionIndex: input.sectionIndex,
        orientation: input.orientation,
        changed: result.changed,
      },
    }
  },
  'docx.section.set_margins': (editor, arguments_, services) => {
    const input = arguments_ as unknown as SetSectionMarginsInput
    const result = setSectionMargins(editor, input, services.getSectionLayoutState?.() ?? [])
    if (!result.ok) return result
    return {
      ok: true,
      output: {
        sectionIndex: input.sectionIndex,
        margins: input.margins,
        changed: result.changed,
      },
    }
  },
  'docx.section.set_page_size': (editor, arguments_, services) => {
    const input = arguments_ as unknown as SetSectionPageSizeInput
    const result = setSectionPageSize(editor, input, services.getSectionLayoutState?.() ?? [])
    if (!result.ok) return result
    return {
      ok: true,
      output: {
        sectionIndex: input.sectionIndex,
        widthTwips: input.widthTwips,
        heightTwips: input.heightTwips,
        orientation: result.orientation!,
        changed: result.changed,
      },
    }
  },
  'docx.section.set_columns': (editor, arguments_, services) => {
    const input = arguments_ as unknown as SetSectionColumnsInput
    const result = setSectionColumns(editor, input, services.getSectionLayoutState?.() ?? [])
    if (!result.ok) return result
    return {
      ok: true,
      output: {
        sectionIndex: input.sectionIndex,
        count: input.count,
        spacingTwips: input.spacingTwips,
        changed: result.changed,
      },
    }
  },
  'docx.section.set_page_border': (editor, arguments_, services) => {
    const input = arguments_ as unknown as SetSectionPageBorderInput
    const result = setSectionPageBorder(editor, input, services.getSectionLayoutState?.() ?? [])
    if (!result.ok) return result
    return {
      ok: true,
      output: {
        sectionIndex: input.sectionIndex,
        enabled: input.enabled,
        changed: result.changed,
      },
    }
  },
  'docx.section.set_different_first_page': (editor, arguments_, services) => {
    const input = arguments_ as unknown as SetSectionDifferentFirstPageInput
    const result = setSectionDifferentFirstPage(
      editor,
      input,
      services.getSectionLayoutState?.() ?? [],
    )
    if (!result.ok) return result
    return {
      ok: true,
      output: {
        sectionIndex: input.sectionIndex,
        enabled: input.enabled,
        changed: result.changed,
      },
    }
  },
  'docx.document.set_different_odd_even_pages': (editor, arguments_, services) => {
    const input = arguments_ as unknown as SetDocumentDifferentOddEvenPagesInput
    const result = setDocumentDifferentOddEvenPages(
      editor,
      input,
      services.getSectionLayoutState?.() ?? [],
      services.getDifferentOddEvenPages?.() ?? false,
    )
    if (!result.ok) return result
    return { ok: true, output: { enabled: input.enabled, changed: result.changed } }
  },
  'docx.section.set_page_numbering': (editor, arguments_, services) => {
    const input = arguments_ as unknown as SetSectionPageNumberingInput
    const result = setSectionPageNumbering(editor, input, services.getSectionLayoutState?.() ?? [])
    if (!result.ok) return result
    return {
      ok: true,
      output: {
        sectionIndex: input.sectionIndex,
        format: input.format,
        start: input.start,
        changed: result.changed,
      },
    }
  },
  'docx.header_footer.set_text': (editor, arguments_, services) => {
    const input = arguments_ as unknown as SetHeaderFooterTextInput
    const target = {
      sectionIndex: input.sectionIndex,
      kind: input.kind,
      variant: input.variant,
    }
    const result = setHeaderFooterText(
      editor,
      input,
      services.getSectionLayoutState?.() ?? [],
      services.getHeaderFooterValue?.(target) ?? null,
    )
    if (!result.ok) return result
    return { ok: true, output: { ...target, changed: result.changed } }
  },
  'docx.header_footer.set_page_number': (editor, arguments_, services) => {
    const input = arguments_ as unknown as SetHeaderFooterPageNumberInput
    const target = {
      sectionIndex: input.sectionIndex,
      kind: input.kind,
      variant: input.variant,
    }
    const result = setHeaderFooterPageNumber(
      editor,
      input,
      services.getSectionLayoutState?.() ?? [],
      services.getHeaderFooterValue?.(target) ?? null,
    )
    if (!result.ok) return result
    return {
      ok: true,
      output: {
        ...target,
        enabled: input.enabled,
        alignment: input.alignment,
        changed: result.changed,
      },
    }
  },
  'docx.header_footer.set_paragraphs': (editor, arguments_, services) => {
    const input = arguments_ as unknown as SetHeaderFooterParagraphsInput
    const target = {
      sectionIndex: input.sectionIndex,
      kind: input.kind,
      variant: input.variant,
    }
    const result = setHeaderFooterParagraphs(
      editor,
      input,
      services.getSectionLayoutState?.() ?? [],
      services.getHeaderFooterValue?.(target) ?? null,
    )
    if (!result.ok) return result
    return {
      ok: true,
      output: { ...target, paragraphCount: input.paragraphs.length, changed: result.changed },
    }
  },
  'docx.text.set_character_format': (editor, arguments_) => {
    const input = arguments_ as unknown as DocxCharacterFormatInput
    const result = setDocxCharacterFormat(editor, input)
    if (!result.ok) {
      return { ok: false, error: 'invalid_arguments', message: result.message }
    }
    return {
      ok: true,
      output: { from: input.range.from, to: input.range.to, changed: result.changed },
    }
  },
  'docx.text.clear_character_format': (editor, arguments_) => {
    const input = arguments_ as unknown as DocxCharacterRangeInput
    const result = clearDocxCharacterFormat(editor, input)
    if (!result.ok) {
      return { ok: false, error: 'invalid_arguments', message: result.message }
    }
    return {
      ok: true,
      output: { from: input.range.from, to: input.range.to, changed: result.changed },
    }
  },
  'docx.text.transform_case': (editor, arguments_) => {
    const input = arguments_ as unknown as DocxTextCaseInput
    const result = transformDocxTextCase(editor, input)
    if (!result.ok) {
      return { ok: false, error: 'invalid_arguments', message: result.message }
    }
    return {
      ok: true,
      output: {
        from: input.range.from,
        to: input.range.to,
        mode: input.mode,
        changed: result.changed,
      },
    }
  },
  'docx.text.set_character_style': (editor, arguments_, services) => {
    const input = arguments_ as unknown as DocxCharacterStyleInput
    const styleId = input.styleId?.trim() ?? null
    if (styleId !== null) {
      if (
        !styleId ||
        styleId.length > 128 ||
        /[\u0000-\u001f\u007f]/.test(styleId) ||
        !services.hasCharacterStyle?.(styleId)
      ) {
        return {
          ok: false,
          error: 'invalid_arguments',
          message:
            'docx.text.set_character_style requires a character style from the current document.',
        }
      }
    }
    const result = setDocxCharacterStyle(editor, { ...input, styleId })
    if (!result.ok) {
      return { ok: false, error: 'invalid_arguments', message: result.message }
    }
    return {
      ok: true,
      output: {
        from: input.range.from,
        to: input.range.to,
        styleId,
        changed: result.changed,
      },
    }
  },
  'docx.list.apply': (editor, arguments_, services) => {
    const kind = arguments_.kind as 'bullet' | 'ordered'
    return executeNativeEnvelope(
      editor,
      {
        commands: [
          {
            createParagraphBullets: {
              target: arguments_.target as CreateParagraphBullets['target'],
              bulletPreset: kind === 'ordered' ? 'NUMBERED' : 'BULLET',
            },
          },
        ],
      },
      {
        failureMessage: 'The DOCX list operation failed.',
        context: { allocateNumId: services.allocateListNumId ?? (() => null) },
      },
    )
  },
  'docx.list.remove': (editor, arguments_) => {
    return executeNativeEnvelope(
      editor,
      {
        commands: [
          {
            deleteParagraphBullets: {
              target: arguments_.target as DeleteParagraphBullets['target'],
            },
          },
        ],
      },
      { failureMessage: 'The DOCX list-removal operation failed.' },
    )
  },
  'docx.list.set_level': (editor, arguments_) => {
    return executeNativeEnvelope(
      editor,
      {
        commands: [
          {
            setListLevel: {
              target: arguments_.target as SetListLevel['target'],
              level: arguments_.level as SetListLevel['level'],
            },
          },
        ],
      },
      { failureMessage: 'The DOCX list-level operation failed.' },
    )
  },
  'docx.list.apply_preset': (editor, arguments_, services) => {
    return executeNativeEnvelope(
      editor,
      {
        commands: [
          {
            applyListPreset: {
              target: arguments_.target as ApplyListPreset['target'],
              levels: arguments_.levels as ApplyListPreset['levels'],
            },
          },
        ],
      },
      {
        failureMessage: 'The DOCX list-preset operation failed.',
        context: { createListDef: services.createListDef },
      },
    )
  },
  'docx.list.restart': (_editor, arguments_, services) => {
    const input = arguments_ as unknown as RestartNumberingInput
    const result = services.restartList?.(input)
    if (!result) {
      return {
        ok: false,
        error: 'execution_failed',
        message: 'The DOCX numbering service is unavailable.',
      }
    }
    if (!result.ok) return result
    return {
      ok: true,
      output: {
        summary: `Restarted numbering at block ${input.blockIndex} from ${input.start} across ${result.changed} item(s)`,
        matched: 1,
        changed: result.changed,
        skippedProtected: 0,
      },
    }
  },
  'docx.list.continue': (_editor, arguments_, services) => {
    const input = arguments_ as unknown as ContinueNumberingInput
    if (input.previousBlockIndex >= input.blockIndex) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'docx.list.continue requires previousBlockIndex to be earlier than blockIndex.',
      }
    }
    const result = services.continueList?.(input)
    if (!result) {
      return {
        ok: false,
        error: 'execution_failed',
        message: 'The DOCX numbering service is unavailable.',
      }
    }
    if (!result.ok) return result
    return {
      ok: true,
      output: {
        summary: `Continued numbering at block ${input.blockIndex} from block ${input.previousBlockIndex} across ${result.changed} item(s)`,
        matched: 2,
        changed: result.changed,
        skippedProtected: 0,
      },
    }
  },
  'docx.table.insert': (editor, arguments_) => {
    const input = arguments_ as unknown as InsertTableInput
    const result = insertTopLevelTableAfterBlock(editor, input)
    if (!result.ok) return result
    return {
      ok: true,
      output: {
        summary: `Inserted a ${input.rows}×${input.columns} DOCX table after block ${input.afterBlockIndex}`,
        tableBlockIndex: result.tableBlockIndex,
        rows: input.rows,
        columns: input.columns,
      },
    }
  },
  'docx.table.delete': (editor, arguments_) => {
    const tableBlockIndex = arguments_.tableBlockIndex as number
    const result = deleteTopLevelTable(editor, tableBlockIndex)
    if (!result.ok) return result
    return {
      ok: true,
      output: {
        summary: `Deleted DOCX table at block ${tableBlockIndex}`,
        tableBlockIndex,
        deleted: true,
      },
    }
  },
  'docx.table.insert_rows': (editor, arguments_) => {
    const input = arguments_ as unknown as InsertTableRowsInput
    const result = insertTableRows(editor, input)
    if (!result.ok) return result
    return {
      ok: true,
      output: {
        summary: `Inserted ${input.count} DOCX row(s) at boundary ${input.rowIndex} in table block ${input.tableBlockIndex}`,
        tableBlockIndex: input.tableBlockIndex,
        rowIndex: input.rowIndex,
        insertedRows: input.count,
      },
    }
  },
  'docx.table.delete_rows': (editor, arguments_) => {
    const input = arguments_ as unknown as DeleteTableRowsInput
    const result = deleteTableRows(editor, input)
    if (!result.ok) return result
    return {
      ok: true,
      output: {
        summary: `Deleted ${input.count} DOCX row(s) from index ${input.rowIndex} in table block ${input.tableBlockIndex}`,
        tableBlockIndex: input.tableBlockIndex,
        rowIndex: input.rowIndex,
        deletedRows: input.count,
      },
    }
  },
  'docx.table.insert_columns': (editor, arguments_) => {
    const input = arguments_ as unknown as InsertTableColumnsInput
    const result = insertTableColumns(editor, input)
    if (!result.ok) return result
    return {
      ok: true,
      output: {
        summary: `Inserted ${input.count} DOCX column(s) at boundary ${input.columnIndex} in table block ${input.tableBlockIndex}`,
        tableBlockIndex: input.tableBlockIndex,
        columnIndex: input.columnIndex,
        insertedColumns: input.count,
      },
    }
  },
  'docx.table.delete_columns': (editor, arguments_) => {
    const input = arguments_ as unknown as DeleteTableColumnsInput
    const result = deleteTableColumns(editor, input)
    if (!result.ok) return result
    return {
      ok: true,
      output: {
        summary: `Deleted ${input.count} DOCX column(s) from index ${input.columnIndex} in table block ${input.tableBlockIndex}`,
        tableBlockIndex: input.tableBlockIndex,
        columnIndex: input.columnIndex,
        deletedColumns: input.count,
      },
    }
  },
  'docx.table.merge_cells': (editor, arguments_) => {
    const input = arguments_ as unknown as MergeTableCellsInput
    const result = mergeTableCells(editor, input)
    if (!result.ok) return result
    const mergedCells = (input.bottomRow - input.topRow) * (input.rightColumn - input.leftColumn)
    return {
      ok: true,
      output: {
        summary: `Merged ${mergedCells} DOCX logical cell(s) in table block ${input.tableBlockIndex}`,
        tableBlockIndex: input.tableBlockIndex,
        mergedCells,
      },
    }
  },
  'docx.table.split_cell': (editor, arguments_) => {
    const input = arguments_ as unknown as SplitTableCellInput
    const result = splitTableCell(editor, input)
    if (!result.ok) return result
    return {
      ok: true,
      output: {
        summary: `Split one DOCX merged cell into ${result.splitCells} logical cells in table block ${input.tableBlockIndex}`,
        tableBlockIndex: input.tableBlockIndex,
        splitCells: result.splitCells,
      },
    }
  },
  'docx.table.set_cell_format': (editor, arguments_) => {
    const input = arguments_ as unknown as SetTableCellFormatInput
    const result = setTableCellFormat(editor, input)
    if (!result.ok) return result
    return {
      ok: true,
      output: {
        summary: `Formatted ${result.matchedCells} DOCX cell(s) in table block ${input.tableBlockIndex}`,
        tableBlockIndex: input.tableBlockIndex,
        matchedCells: result.matchedCells,
        changedCells: result.changedCells,
      },
    }
  },
  'docx.table.set_cell_borders': (editor, arguments_) => {
    const input = arguments_ as unknown as SetTableCellBordersInput
    const result = setTableCellBorders(editor, input)
    if (!result.ok) return result
    return {
      ok: true,
      output: {
        summary: `Set ${input.mode} borders on ${result.matchedCells} DOCX cell(s) in table block ${input.tableBlockIndex}`,
        tableBlockIndex: input.tableBlockIndex,
        matchedCells: result.matchedCells,
        changedCells: result.changedCells,
      },
    }
  },
  'docx.table.set_style': (editor, arguments_, services) => {
    const rawInput = arguments_ as unknown as SetTableStyleInput
    const styleId = rawInput.styleId?.trim() ?? null
    if (
      styleId !== null &&
      (!styleId ||
        styleId.length > 128 ||
        /[\u0000-\u001f\u007f]/.test(styleId) ||
        !services.hasTableStyle?.(styleId))
    ) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'docx.table.set_style requires a table style from the current document.',
      }
    }
    const input = { ...rawInput, styleId }
    const result = setTableStyle(editor, input)
    if (!result.ok) return result
    return {
      ok: true,
      output: {
        summary: `Set DOCX table block ${input.tableBlockIndex} style to ${styleId ?? 'none'}`,
        tableBlockIndex: input.tableBlockIndex,
        styleId,
        changed: result.changed,
      },
    }
  },
  'docx.table.set_row_height': (editor, arguments_) => {
    const input = arguments_ as unknown as SetTableRowHeightInput
    const result = setTableRowHeight(editor, input)
    if (!result.ok) return result
    return {
      ok: true,
      output: {
        summary: `Set ${result.matchedRows} DOCX row height(s) in table block ${input.tableBlockIndex}`,
        tableBlockIndex: input.tableBlockIndex,
        matchedRows: result.matchedRows,
        changedRows: result.changedRows,
      },
    }
  },
  'docx.table.set_column_widths': (editor, arguments_) => {
    const input = arguments_ as unknown as SetTableColumnWidthsInput
    const result = setTableColumnWidths(editor, input)
    if (!result.ok) return result
    return {
      ok: true,
      output: {
        summary: `Set ${result.matchedColumns} DOCX column width(s) in table block ${input.tableBlockIndex}`,
        tableBlockIndex: input.tableBlockIndex,
        matchedColumns: result.matchedColumns,
        changedColumns: result.changedColumns,
      },
    }
  },
  'docx.block.delete': (editor, arguments_) => {
    return executeNativeEnvelope(
      editor,
      {
        commands: [
          {
            deleteBlocks: {
              target: arguments_.target as DeleteBlocks['target'],
            },
          },
        ],
      },
      {
        failureMessage: 'The DOCX block-delete operation failed.',
        includeSkippedDeleted: true,
      },
    )
  },
  'docx.block.move': (editor, arguments_) => {
    return executeNativeEnvelope(
      editor,
      {
        commands: [
          {
            moveBlocks: {
              blockIndexes: arguments_.blockIndexes as MoveBlocks['blockIndexes'],
              afterBlockIndex: arguments_.afterBlockIndex as MoveBlocks['afterBlockIndex'],
            },
          },
        ],
      },
      {
        failureMessage: 'The DOCX block-move operation failed.',
        failureError: 'invalid_arguments',
      },
    )
  },
  'docx.paragraph.set_heading_level': (editor, arguments_) => {
    return executeNativeEnvelope(
      editor,
      {
        commands: [
          {
            setHeadingLevel: {
              target: arguments_.target as SetHeadingLevel['target'],
              level: arguments_.level as SetHeadingLevel['level'],
            },
          },
        ],
      },
      { failureMessage: 'The DOCX heading-level operation failed.' },
    )
  },
  'docx.paragraph.set_drop_cap': (editor, arguments_) => {
    const result = setDocxDropCap(editor, arguments_ as unknown as SetDocxDropCapInput)
    if (!result.ok) return result
    return {
      ok: true,
      output: {
        summary: `Set DOCX drop cap ${result.mode} on block ${result.blockIndex}`,
        blockIndex: result.blockIndex,
        mode: result.mode,
        lines: result.lines,
        changed: result.changed,
      },
    }
  },
  'docx.wordart.insert': (editor, arguments_) => {
    const result = insertDocxWordArtAfterBlock(
      editor,
      arguments_ as unknown as InsertDocxWordArtInput,
    )
    if (!result.ok) return result
    return {
      ok: true,
      output: {
        summary: `Inserted DOCX ${result.preset} WordArt at block ${result.blockIndex}`,
        blockIndex: result.blockIndex,
        preset: result.preset,
        drawingId: result.drawingId,
        changed: result.changed,
      },
    }
  },
  'docx.paragraph.set_style': (editor, arguments_) => {
    return executeNativeEnvelope(
      editor,
      {
        commands: [
          {
            updateParagraphStyle: {
              target: arguments_.target as UpdateParagraphStyle['target'],
              style: arguments_.style as UpdateParagraphStyle['style'],
              fields: arguments_.fields as UpdateParagraphStyle['fields'],
            },
          },
        ],
      },
      { failureMessage: 'The DOCX paragraph-style operation failed.' },
    )
  },
  'docx.paragraph.set_direction': (editor, arguments_) => {
    return executeNativeEnvelope(
      editor,
      {
        commands: [
          {
            setParagraphDirection: {
              target: arguments_.target as SetParagraphDirection['target'],
              direction: arguments_.direction as SetParagraphDirection['direction'],
            },
          },
        ],
      },
      { failureMessage: 'The DOCX paragraph-direction operation failed.' },
    )
  },
  'docx.text.replace_all': (editor, arguments_) => {
    if (arguments_.containsText === '') {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: '$.containsText must not be empty.',
      }
    }
    return executeNativeEnvelope(
      editor,
      {
        commands: [
          {
            replaceAllText: {
              containsText: arguments_.containsText as string,
              replaceText: arguments_.replaceText as string,
              ...(arguments_.matchCase === undefined
                ? {}
                : { matchCase: arguments_.matchCase as boolean }),
            },
          },
        ],
      },
      {
        failureMessage: 'The DOCX replace-all operation failed.',
        includeSkippedDeleted: true,
      },
    )
  },
  'docx.toc.insert': (editor, arguments_) => {
    return executeNativeEnvelope(
      editor,
      {
        commands: [
          {
            insertToc: {
              afterBlockIndex: arguments_.afterBlockIndex as InsertToc['afterBlockIndex'],
            },
          },
        ],
      },
      {
        failureMessage: 'The DOCX table-of-contents insertion failed.',
        failureError: (message) =>
          message.startsWith('insertToc: afterBlockIndex out of range')
            ? 'invalid_arguments'
            : 'execution_failed',
      },
    )
  },
  'docx.toc.refresh': (editor, arguments_, services) => {
    const result = refreshDocxToc(
      editor,
      services.getSourceBlocks?.() ?? [],
      arguments_ as unknown as RefreshDocxTocInput,
    )
    if (!result.ok) return result
    return {
      ok: true,
      output: {
        summary: `Refreshed DOCX table of contents with ${result.entries} ${result.entries === 1 ? 'entry' : 'entries'}`,
        entries: result.entries,
        replacedBlocks: result.replacedBlocks,
      },
    }
  },
  'docx.text.set_style': (editor, arguments_) => {
    return executeNativeEnvelope(
      editor,
      {
        commands: [
          {
            updateTextStyle: {
              target: arguments_.target as UpdateTextStyle['target'],
              style: arguments_.style as UpdateTextStyle['style'],
              fields: arguments_.fields as UpdateTextStyle['fields'],
            },
          },
        ],
      },
      { failureMessage: 'The DOCX text-style operation failed.' },
    )
  },
  'docx.document.save': async (_editor, _arguments, services) => {
    try {
      const saved = await services.save()
      return {
        ok: true,
        output: { saved: true, fileName: saved.fileName },
        checkpointRecovery: false,
      }
    } catch {
      return {
        ok: false,
        error: 'execution_failed',
        message: 'The document was not saved.',
      }
    }
  },
  'docx.document.compare': (_editor, arguments_) => {
    const path = arguments_.path as string
    if (!path.startsWith('/') || path.length > 1_024 || !path.toLowerCase().endsWith('.docx')) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'docx.document.compare requires one absolute local DOCX path.',
      }
    }
    return {
      ok: false,
      error: 'execution_failed',
      message: 'docx.document.compare requires Broker staging.',
    }
  },
  'docx.document.compare_staged': async (_editor, arguments_, services) => {
    const data = arguments_.data
    if (
      !services.getSourceBlocks ||
      !services.commitCompareResult ||
      !isArrayBuffer(data) ||
      data.byteLength !== arguments_.size ||
      !(arguments_.blobId as string)
    ) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'docx.document.compare_staged requires valid mounted state and staged bytes.',
      }
    }
    const result = await compareDocxBytes(services.getSourceBlocks() as Block[], {
      name: arguments_.name as string,
      data,
    })
    if (!result.ok) return result
    services.commitCompareResult({ otherName: result.otherName, entries: result.entries })
    return {
      ok: true,
      output: {
        summary: `Compared DOCX with ${result.otherName}: ${result.added} added, ${result.removed} removed, ${result.changed} changed`,
        otherName: result.otherName,
        added: result.added,
        removed: result.removed,
        changed: result.changed,
        identical: result.identical,
      },
      checkpointRecovery: false,
    }
  },
  'docx.document.load_staged': async (_editor, arguments_, services) => {
    if (!isArrayBuffer(arguments_.data)) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: '$.data must be a hydrated ArrayBuffer.',
      }
    }
    if (
      !(arguments_.name as string).toLowerCase().endsWith('.docx') ||
      arguments_.data.byteLength !== arguments_.size
    ) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'docx.document.load_staged requires a valid staged DOCX descriptor.',
      }
    }
    try {
      await services.loadStaged({
        blobId: arguments_.blobId as string,
        name: arguments_.name as string,
        data: arguments_.data as ArrayBuffer,
      })
      return { ok: true, checkpointRecovery: false }
    } catch (error) {
      return {
        ok: false,
        error: 'execution_failed',
        message: error instanceof Error ? error.message : String(error),
      }
    }
  },
  'docx.image.insert': (_editor, arguments_) => {
    const path = arguments_.path as string
    if (!path.startsWith('/') || path.length > 1_024 || !/\.(png|jpe?g|gif)$/i.test(path)) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'docx.image.insert requires an absolute PNG, JPEG, or GIF path.',
      }
    }
    return {
      ok: false,
      error: 'execution_failed',
      message: 'docx.image.insert requires Broker staging.',
    }
  },
  'docx.image.insert_staged': (editor, arguments_) => {
    const data = arguments_.data
    const name = arguments_.name as string
    if (
      !isArrayBuffer(data) ||
      data.byteLength !== arguments_.size ||
      data.byteLength > 20 * 1024 * 1024 ||
      !(arguments_.blobId as string) ||
      name.length === 0 ||
      name.length > 255
    ) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'docx.image.insert_staged requires a valid bounded staged image descriptor.',
      }
    }
    const mime = stagedDocxImageMediaType(name, data)
    if (!mime) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'docx.image.insert_staged requires matching PNG, JPEG, or GIF bytes.',
      }
    }
    const input = {
      ...(arguments_ as unknown as Omit<InsertDocxImageInput, 'data' | 'mime' | 'name'>),
      data,
      mime,
      name,
    } satisfies InsertDocxImageInput
    const inserted = insertDocxImageAfterBlock(editor, input)
    if (!inserted.ok) return inserted
    return {
      ok: true,
      output: {
        summary: `Inserted DOCX image at block ${inserted.insertedBlockIndex}`,
        insertedBlockIndex: inserted.insertedBlockIndex,
        widthPx: input.widthPx,
        heightPx: input.heightPx,
        alignment: input.alignment,
      },
    }
  },
  'docx.image.replace': (_editor, arguments_) => {
    const path = arguments_.path as string
    if (!path.startsWith('/') || path.length > 1_024 || !/\.(png|jpe?g|gif)$/i.test(path)) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'docx.image.replace requires an absolute PNG, JPEG, or GIF path.',
      }
    }
    return {
      ok: false,
      error: 'execution_failed',
      message: 'docx.image.replace requires Broker staging.',
    }
  },
  'docx.image.replace_staged': (editor, arguments_) => {
    const data = arguments_.data
    const name = arguments_.name as string
    if (
      !isArrayBuffer(data) ||
      data.byteLength !== arguments_.size ||
      data.byteLength > 20 * 1024 * 1024 ||
      !(arguments_.blobId as string) ||
      name.length === 0 ||
      name.length > 255
    ) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'docx.image.replace_staged requires a valid bounded staged image descriptor.',
      }
    }
    const mime = stagedDocxImageMediaType(name, data)
    if (!mime) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'docx.image.replace_staged requires matching PNG, JPEG, or GIF bytes.',
      }
    }
    const input = {
      ...(arguments_ as unknown as Omit<ReplaceDocxImageInput, 'data' | 'mime'>),
      data,
      mime,
    } satisfies ReplaceDocxImageInput
    const replaced = replaceDocxImageAtBlock(editor, input)
    if (!replaced.ok) return replaced
    return {
      ok: true,
      output: {
        summary: `Replaced DOCX image at block ${replaced.imageBlockIndex}`,
        imageBlockIndex: replaced.imageBlockIndex,
        widthPx: input.widthPx,
        heightPx: input.heightPx,
      },
    }
  },
  'docx.image.remove': (editor, arguments_) => {
    const imageBlockIndex = arguments_.imageBlockIndex as number
    const result = removeDocxImageAtBlock(editor, imageBlockIndex)
    if (!result.ok) return result
    return {
      ok: true,
      output: {
        summary: `Removed DOCX image at block ${imageBlockIndex}`,
        imageBlockIndex,
        deleted: true,
      },
    }
  },
  'docx.image.set_crop': (editor, arguments_) => {
    const input = arguments_ as unknown as SetDocxImageCropInput
    const result = setDocxImageCropAtBlock(editor, input)
    if (!result.ok) return result
    return {
      ok: true,
      output: {
        summary: `Set DOCX image crop at block ${result.imageBlockIndex}`,
        imageBlockIndex: result.imageBlockIndex,
        left: input.left,
        top: input.top,
        right: input.right,
        bottom: input.bottom,
        changed: result.changed,
      },
    }
  },
  'docx.image.set_offset_position': (editor, arguments_) => {
    const input = arguments_ as unknown as SetDocxImageOffsetPositionInput
    const result = setDocxImageOffsetPositionAtBlock(editor, input)
    if (!result.ok) return result
    return {
      ok: true,
      output: {
        summary: `Set DOCX image offset position at block ${result.imageBlockIndex}`,
        imageBlockIndex: result.imageBlockIndex,
        wrap: input.wrap,
        offsetXEmu: input.offsetXEmu,
        offsetYEmu: input.offsetYEmu,
        changed: result.changed,
      },
    }
  },
  'docx.image.set_margin_position': (editor, arguments_) => {
    const input = arguments_ as unknown as SetDocxImageMarginPositionInput
    const result = setDocxImageMarginPositionAtBlock(editor, input)
    if (!result.ok) return result
    return {
      ok: true,
      output: {
        summary: `Set DOCX image margin position at block ${result.imageBlockIndex}`,
        imageBlockIndex: result.imageBlockIndex,
        horizontal: input.horizontal,
        vertical: input.vertical,
        changed: result.changed,
      },
    }
  },
  'docx.image.set_wrap': (editor, arguments_) => {
    const input = arguments_ as unknown as SetDocxImageWrapInput
    const result = setDocxImageWrapAtBlock(editor, input)
    if (!result.ok) return result
    return {
      ok: true,
      output: {
        summary: `Set DOCX image wrap at block ${result.imageBlockIndex}`,
        imageBlockIndex: result.imageBlockIndex,
        wrap: input.wrap,
        changed: result.changed,
      },
    }
  },
  'docx.image.set_transform': (editor, arguments_) => {
    const input = arguments_ as unknown as SetDocxImageTransformInput
    const result = setDocxImageTransformAtBlock(editor, input)
    if (!result.ok) return result
    return {
      ok: true,
      output: {
        summary: `Set DOCX image transform at block ${result.imageBlockIndex}`,
        imageBlockIndex: result.imageBlockIndex,
        rotationDegrees: input.rotationDegrees,
        flipHorizontal: input.flipHorizontal,
        flipVertical: input.flipVertical,
        changed: result.changed,
      },
    }
  },
  'docx.image.update': (editor, arguments_) => {
    return executeNativeEnvelope(
      editor,
      {
        commands: [
          {
            updateImageProperties: {
              target: arguments_.target as UpdateImageProperties['target'],
              properties: arguments_.properties as UpdateImageProperties['properties'],
              fields: arguments_.fields as UpdateImageProperties['fields'],
            },
          },
        ],
      },
      { failureMessage: 'The DOCX image update failed.' },
    )
  },
  'docx.shape.insert': (editor, arguments_) => {
    const input = arguments_ as unknown as InsertDocxShapeInput
    const result = insertDocxShapeAfterBlock(editor, input)
    if (!result.ok) return result
    return {
      ok: true,
      output: {
        summary: `Inserted DOCX ${input.preset} shape at block ${result.shapeBlockIndex}`,
        shapeBlockIndex: result.shapeBlockIndex,
        preset: input.preset,
        widthEmu: input.widthEmu,
        heightEmu: input.heightEmu,
      },
    }
  },
  'docx.line.insert': (editor, arguments_) => {
    const input = arguments_ as unknown as InsertDocxLineInput
    const result = insertDocxLineAfterBlock(editor, input)
    if (!result.ok) return result
    return {
      ok: true,
      output: {
        summary: `Inserted DOCX ${input.kind} connector at block ${result.lineBlockIndex}`,
        lineBlockIndex: result.lineBlockIndex,
        kind: input.kind,
        widthEmu: input.widthEmu,
        heightEmu: input.heightEmu,
      },
    }
  },
  'docx.textbox.insert': (editor, arguments_) => {
    const input = arguments_ as unknown as InsertDocxTextboxInput
    const result = insertDocxTextboxAfterBlock(editor, input)
    if (!result.ok) return result
    return {
      ok: true,
      output: {
        summary: `Inserted DOCX textbox at block ${result.textboxBlockIndex}`,
        textboxBlockIndex: result.textboxBlockIndex,
        widthEmu: input.widthEmu,
        heightEmu: input.heightEmu,
      },
    }
  },
  'docx.textbox.set_content': (editor, arguments_) => {
    const input = arguments_ as unknown as SetDocxTextboxContentInput
    const result = setDocxTextboxContentAtBlock(editor, input)
    if (!result.ok) return result
    return {
      ok: true,
      output: {
        summary: `Set DOCX textbox ${input.textboxIndex} content at block ${input.objectBlockIndex}`,
        objectBlockIndex: input.objectBlockIndex,
        textboxIndex: input.textboxIndex,
        paragraphCount: input.paragraphs.length,
        heightPx: result.heightPx,
        changed: result.changed,
      },
    }
  },
  'docx.chart.insert': (editor, arguments_) => {
    const input = arguments_ as unknown as InsertDocxChartInput
    const result = insertDocxChartAfterBlock(editor, input)
    if (!result.ok) return result
    return {
      ok: true,
      output: {
        summary: `Inserted DOCX ${input.kind} chart at block ${result.chartBlockIndex}`,
        chartBlockIndex: result.chartBlockIndex,
        kind: input.kind,
        widthPx: input.widthPx,
        heightPx: input.heightPx,
      },
    }
  },
  'docx.chart.update': (editor, arguments_) => {
    const input = arguments_ as unknown as UpdateDocxChartInput
    const result = updateDocxChartAtBlock(editor, input)
    if (!result.ok) return result
    return {
      ok: true,
      output: {
        summary: `Updated DOCX chart at block ${input.chartBlockIndex}`,
        chartBlockIndex: input.chartBlockIndex,
        fields: input.fields,
        changed: result.changed,
      },
    }
  },
  'docx.equation.insert': (editor, arguments_) => {
    const input = arguments_ as unknown as InsertDocxEquationInput
    const result = insertDocxEquation(editor, input)
    if (!result.ok) return result
    return {
      ok: true,
      output: {
        summary:
          input.placement === 'block'
            ? `Inserted block DOCX equation at block ${result.equationBlockIndex}`
            : `Inserted inline DOCX equation at range ${result.from}..${result.to}`,
        placement: input.placement,
        equationBlockIndex: result.equationBlockIndex,
        from: result.from,
        to: result.to,
      },
    }
  },
  'docx.equation.update': (editor, arguments_) => {
    const input = arguments_ as unknown as UpdateDocxEquationInput
    const result = updateDocxEquation(editor, input)
    if (!result.ok) return result
    const location =
      input.placement === 'block'
        ? `block ${input.equationBlockIndex}`
        : `range ${input.from}..${input.to}`
    return {
      ok: true,
      output: {
        summary: `Updated ${input.placement} DOCX equation at ${location} from ${input.mode}`,
        placement: input.placement,
        mode: input.mode,
        equationBlockIndex: input.equationBlockIndex,
        from: input.from,
        to: input.to,
        changed: result.changed,
      },
    }
  },
  'docx.object.remove': (editor, arguments_) => {
    const input = arguments_ as unknown as RemoveDocxObjectInput
    const result = removeDocxObjectAtBlock(editor, input)
    if (!result.ok) return result
    return {
      ok: true,
      output: {
        summary: `Removed DOCX ${result.objectKind} at block ${input.objectBlockIndex}`,
        objectBlockIndex: input.objectBlockIndex,
        objectKind: result.objectKind,
        replacementParagraphInserted: result.replacementParagraphInserted,
      },
    }
  },
  'docx.object.set_size': (editor, arguments_) => {
    const input = arguments_ as unknown as SetDocxObjectSizeInput
    const result = setDocxObjectSizeAtBlock(editor, input)
    if (!result.ok) return result
    return {
      ok: true,
      output: {
        summary: `Set DOCX ${result.objectKind} at block ${input.objectBlockIndex} to ${input.widthPx}×${input.heightPx} px`,
        objectBlockIndex: input.objectBlockIndex,
        objectKind: result.objectKind,
        widthPx: input.widthPx,
        heightPx: input.heightPx,
        changed: result.changed,
      },
    }
  },
  'docx.object.set_offset_position': (editor, arguments_) => {
    const input = arguments_ as unknown as SetDocxObjectOffsetInput
    const result = setDocxObjectOffsetAtBlock(editor, input)
    if (!result.ok) return result
    return {
      ok: true,
      output: {
        summary: `Positioned DOCX ${result.objectKind} at block ${input.objectBlockIndex} to (${input.offsetXEmu}, ${input.offsetYEmu}) EMU with ${input.wrap} wrap`,
        objectBlockIndex: input.objectBlockIndex,
        objectKind: result.objectKind,
        wrap: input.wrap,
        offsetXEmu: input.offsetXEmu,
        offsetYEmu: input.offsetYEmu,
        changed: result.changed,
      },
    }
  },
  'docx.object.set_style': (editor, arguments_) => {
    const input = arguments_ as unknown as SetDocxObjectStyleInput
    const result = setDocxObjectStyleAtBlock(editor, input)
    if (!result.ok) return result
    return {
      ok: true,
      output: {
        summary: `Set DOCX ${result.objectKind} style at block ${input.objectBlockIndex}`,
        objectBlockIndex: input.objectBlockIndex,
        objectKind: result.objectKind,
        fields: input.fields,
        changed: result.changed,
      },
    }
  },
} satisfies Record<DocxOperationId, DocxOperationHandler>

const descriptorsByName = new Map<string, DocxOperationDescriptor>()
for (const descriptor of docxOperationCatalog.operations) {
  descriptorsByName.set(descriptor.id, descriptor)
  for (const alias of descriptor.compatibilityAliases) {
    descriptorsByName.set(alias, descriptor)
  }
}

export async function executeDocxOperation(
  editor: Editor,
  command: DocxOperationCommand,
  services: DocxOperationServices,
): Promise<DocxOperationExecution> {
  const descriptor = descriptorsByName.get(command.operation)
  if (!descriptor) return { handled: false }

  const validation = validateJsonSchemaValue(descriptor.inputSchema, command.arguments)
  if (!validation.ok) {
    return {
      handled: true,
      operationId: descriptor.id,
      ok: false,
      error: 'invalid_arguments',
      message: validation.error.message,
    }
  }

  const result = await handlers[descriptor.id](editor, command.arguments, services)
  return {
    handled: true,
    operationId: descriptor.id,
    ...result,
  }
}
