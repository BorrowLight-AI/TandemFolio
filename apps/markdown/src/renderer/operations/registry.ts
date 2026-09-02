import type { Editor } from '@tiptap/core'
import { validateJsonSchemaValue } from '@tandemfolio/operation-contract'

import { markdownOperationCatalog } from './catalog'
import { runMarkdownHistoryAction } from '../editor/history-actions'
import { setMarkdownBlockType, type MarkdownBlockType } from '../editor/block-type-actions'
import { setMarkdownTextMarks, type MarkdownTextMarks } from '../editor/text-mark-actions'
import { setMarkdownListType, type MarkdownListType } from '../editor/list-actions'
import {
  insertMarkdownDivider,
  insertMarkdownTable,
  updateMarkdownTable,
  type MarkdownTableUpdateAction,
} from '../editor/structure-actions'
import {
  insertMarkdownImage,
  markdownImageDataUrl,
  stagedMarkdownImageMediaType,
} from '../editor/image-actions'
import { updateMarkdownBlock, type MarkdownBlockUpdateAction } from '../editor/block-actions'
import {
  setMarkdownCodeBlockLanguage,
  type MarkdownCodeBlockLanguage,
} from '../editor/code-block-actions'
import { setMarkdownSelection } from '../editor/selection-actions'

type MarkdownOperationDescriptor = (typeof markdownOperationCatalog.operations)[number]
type MarkdownOperationId = MarkdownOperationDescriptor['id']

interface MarkdownOperationHandlerSuccess {
  readonly ok: true
  readonly output?: Readonly<Record<string, unknown>>
  readonly checkpointRecovery?: false
}

interface MarkdownOperationHandlerFailure {
  readonly ok: false
  readonly error: 'invalid_arguments' | 'execution_failed'
  readonly message: string
}

type MarkdownOperationHandlerResult =
  MarkdownOperationHandlerSuccess | MarkdownOperationHandlerFailure

export interface MarkdownOperationServices {
  readonly loadStaged: (input: {
    readonly name: string
    readonly data: ArrayBuffer
    readonly assetRootId?: string
  }) => void | Promise<void>
  readonly save: (input: {
    readonly saveAs: boolean
  }) => Promise<{ readonly ok: true; readonly fileName: string } | { readonly ok: false }>
  readonly exportDocx: () => Promise<
    { readonly ok: true; readonly fileName: string } | { readonly ok: false }
  >
  readonly openPrintDialog: () => { readonly ok: true } | { readonly ok: false }
  readonly setAutoSave: (input: { readonly enabled: boolean }) => void
  readonly setFrontmatter: (input: { readonly yaml: string }) => void
}

type MarkdownOperationHandler = (
  editor: Editor,
  arguments_: Record<string, unknown>,
  services: MarkdownOperationServices,
) => MarkdownOperationHandlerResult | Promise<MarkdownOperationHandlerResult>

function isArrayBuffer(value: unknown): value is ArrayBuffer {
  return Object.prototype.toString.call(value) === '[object ArrayBuffer]'
}

export interface MarkdownOperationCommand {
  readonly operation: string
  readonly arguments: Record<string, unknown>
}

export type MarkdownOperationExecution =
  | { readonly handled: false }
  | {
      readonly handled: true
      readonly operationId: MarkdownOperationId
      readonly ok: true
      readonly output?: Readonly<Record<string, unknown>>
      readonly checkpointRecovery?: false
    }
  | {
      readonly handled: true
      readonly operationId: MarkdownOperationId
      readonly ok: false
      readonly error: 'invalid_arguments' | 'execution_failed'
      readonly message: string
    }

const handlers = {
  'markdown.text.insert': (editor, arguments_) => {
    editor
      .chain()
      .focus()
      .insertContent(arguments_.text as string)
      .run()
    return { ok: true }
  },
  'markdown.text.replace_selection': (editor, arguments_) => {
    editor
      .chain()
      .focus()
      .insertContent(arguments_.text as string)
      .run()
    return { ok: true }
  },
  'markdown.selection.set': (editor, arguments_) => {
    const result = setMarkdownSelection(editor, {
      from: arguments_.from as number,
      to: arguments_.to as number,
    })
    if (!result.ok) {
      return { ok: false, error: 'execution_failed', message: result.message }
    }
    return {
      ok: true,
      output: { from: result.from, to: result.to },
      checkpointRecovery: false,
    }
  },
  'markdown.text.set_marks': (editor, arguments_) => {
    const result = setMarkdownTextMarks(editor, {
      from: arguments_.from as number,
      to: arguments_.to as number,
      marks: arguments_.marks as unknown as MarkdownTextMarks,
    })
    if (result.ok) return { ok: true }
    return { ok: false, error: 'execution_failed', message: result.message }
  },
  'markdown.document.save': async (_editor, _arguments, services) => {
    const saved = await services.save({ saveAs: false })
    if (!saved.ok) {
      return {
        ok: false,
        error: 'execution_failed',
        message: 'The Markdown save was canceled or failed.',
      }
    }
    return {
      ok: true,
      output: { saved: true, fileName: saved.fileName },
    }
  },
  'markdown.document.save_as': async (_editor, _arguments, services) => {
    const saved = await services.save({ saveAs: true })
    if (!saved.ok) {
      return {
        ok: false,
        error: 'execution_failed',
        message: 'The Markdown Save As was canceled or failed.',
      }
    }
    return {
      ok: true,
      output: { saved: true, fileName: saved.fileName },
    }
  },
  'markdown.document.export_docx': async (_editor, _arguments, services) => {
    const exported = await services.exportDocx()
    if (!exported.ok) {
      return {
        ok: false,
        error: 'execution_failed',
        message: 'The Markdown DOCX export failed.',
      }
    }
    return {
      ok: true,
      output: { exported: true, fileName: exported.fileName },
    }
  },
  'markdown.document.open_print_dialog': (_editor, _arguments, services) => {
    const opened = services.openPrintDialog()
    if (!opened.ok) {
      return {
        ok: false,
        error: 'execution_failed',
        message: 'The host blocked the Markdown print dialog.',
      }
    }
    return { ok: true, output: { opened: true } }
  },
  'markdown.document.set_auto_save': (_editor, arguments_, services) => {
    const enabled = arguments_.enabled as boolean
    services.setAutoSave({ enabled })
    return { ok: true, output: { enabled }, checkpointRecovery: false }
  },
  'markdown.document.load_staged': async (_editor, arguments_, services) => {
    if (!isArrayBuffer(arguments_.data)) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: '$.data must be a hydrated ArrayBuffer.',
      }
    }
    await services.loadStaged({
      name: arguments_.name as string,
      data: arguments_.data as ArrayBuffer,
      ...(typeof arguments_.assetRootId === 'string'
        ? { assetRootId: arguments_.assetRootId }
        : {}),
    })
    return { ok: true, checkpointRecovery: false }
  },
  'markdown.list.set_type': (editor, arguments_) => {
    const result = setMarkdownListType(editor, {
      textBlockIndex: arguments_.textBlockIndex as number,
      type: arguments_.type as MarkdownListType,
    })
    if (result.ok) return { ok: true }
    return { ok: false, error: 'execution_failed', message: result.message }
  },
  'markdown.history.undo': (editor) => {
    if (runMarkdownHistoryAction(editor, 'undo')) return { ok: true }
    return {
      ok: false,
      error: 'execution_failed',
      message: 'markdown.history.undo requires an available undo entry.',
    }
  },
  'markdown.block.set_type': (editor, arguments_) => {
    const result = setMarkdownBlockType(editor, {
      textBlockIndex: arguments_.textBlockIndex as number,
      type: arguments_.type as MarkdownBlockType,
    })
    if (result.ok) return { ok: true }
    return { ok: false, error: 'execution_failed', message: result.message }
  },
  'markdown.history.redo': (editor) => {
    if (runMarkdownHistoryAction(editor, 'redo')) return { ok: true }
    return {
      ok: false,
      error: 'execution_failed',
      message: 'markdown.history.redo requires an available redo entry.',
    }
  },
  'markdown.divider.insert': (editor, arguments_) => {
    const result = insertMarkdownDivider(editor, { position: arguments_.position as number })
    if (result.ok) return { ok: true }
    return { ok: false, error: 'execution_failed', message: result.message }
  },
  'markdown.table.insert': (editor, arguments_) => {
    const result = insertMarkdownTable(editor, {
      position: arguments_.position as number,
      rows: arguments_.rows as number,
      columns: arguments_.columns as number,
      headerRow: arguments_.headerRow as boolean,
    })
    if (result.ok) return { ok: true }
    return { ok: false, error: 'execution_failed', message: result.message }
  },
  'markdown.image.insert': (_editor, arguments_) => {
    const path = arguments_.path as string
    if (
      (!path.startsWith('/') && !/^[a-zA-Z]:[\\/]/.test(path)) ||
      !/\.(png|jpe?g|gif)$/i.test(path)
    ) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'markdown.image.insert requires an absolute PNG, JPEG, or GIF path.',
      }
    }
    return {
      ok: false,
      error: 'execution_failed',
      message: 'markdown.image.insert requires Broker staging.',
    }
  },
  'markdown.image.insert_staged': (editor, arguments_) => {
    const data = arguments_.data
    const name = arguments_.name as string
    if (
      !isArrayBuffer(data) ||
      data.byteLength !== arguments_.size ||
      data.byteLength === 0 ||
      data.byteLength > 20 * 1024 * 1024
    ) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'markdown.image.insert_staged requires a valid bounded staged image descriptor.',
      }
    }
    const mediaType = stagedMarkdownImageMediaType(name, data)
    if (!mediaType) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'markdown.image.insert_staged requires matching PNG, JPEG, or GIF bytes.',
      }
    }
    const result = insertMarkdownImage(editor, {
      position: arguments_.position as number,
      src: markdownImageDataUrl(mediaType, data),
      alt: arguments_.alt as string,
      title: arguments_.title as string | null,
    })
    if (result.ok) return { ok: true }
    return { ok: false, error: 'execution_failed', message: result.message }
  },
  'markdown.frontmatter.set': (_editor, arguments_, services) => {
    services.setFrontmatter({ yaml: arguments_.yaml as string })
    return { ok: true }
  },
  'markdown.table.update': (editor, arguments_) => {
    const result = updateMarkdownTable(editor, {
      position: arguments_.position as number,
      action: arguments_.action as MarkdownTableUpdateAction,
      headerRow: arguments_.headerRow as boolean | null,
    })
    if (result.ok) return { ok: true }
    return { ok: false, error: 'execution_failed', message: result.message }
  },
  'markdown.block.update': (editor, arguments_) => {
    const result = updateMarkdownBlock(editor, {
      blockIndex: arguments_.blockIndex as number,
      action: arguments_.action as MarkdownBlockUpdateAction,
      afterBlockIndex: arguments_.afterBlockIndex as number | null,
      content: arguments_.content as string | null,
    })
    if (result.ok) return { ok: true }
    return { ok: false, error: 'execution_failed', message: result.message }
  },
  'markdown.code_block.set_language': (editor, arguments_) => {
    const result = setMarkdownCodeBlockLanguage(editor, {
      textBlockIndex: arguments_.textBlockIndex as number,
      language: arguments_.language as MarkdownCodeBlockLanguage,
    })
    if (result.ok) return { ok: true }
    return { ok: false, error: 'execution_failed', message: result.message }
  },
} satisfies Record<MarkdownOperationId, MarkdownOperationHandler>

const descriptorsByName = new Map<string, MarkdownOperationDescriptor>()
for (const descriptor of markdownOperationCatalog.operations) {
  descriptorsByName.set(descriptor.id, descriptor)
  for (const alias of descriptor.compatibilityAliases) {
    descriptorsByName.set(alias, descriptor)
  }
}

export async function executeMarkdownOperation(
  editor: Editor,
  command: MarkdownOperationCommand,
  services: MarkdownOperationServices,
): Promise<MarkdownOperationExecution> {
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
