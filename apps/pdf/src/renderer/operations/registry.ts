import {
  validateJsonSchemaValue,
  type OperationExecutionErrorCode,
} from '@tandemfolio/operation-contract'
import type {
  AnnotDeleteInput,
  DrawingInput,
  FormValueInput,
  ImageEditInput,
  MarkupInput,
  MetadataInput,
  NoteEditInput,
  NoteReplyTarget,
  StaticFormFillRecord,
  TextEditInput,
  TextInsertInput,
} from '../../shared/ipc'
import type { HeaderFooterConfig, WatermarkConfig } from '../stamps'

import { pdfOperationCatalog } from './catalog'

type PdfOperationDescriptor = (typeof pdfOperationCatalog.operations)[number]
type PdfOperationId = PdfOperationDescriptor['id']
type PdfRegistryError = Extract<
  OperationExecutionErrorCode,
  'invalid_arguments' | 'execution_failed'
>

export interface PdfOperationCommand {
  readonly operation: string
  readonly arguments: Record<string, unknown>
}

export interface PdfOperationServices {
  readonly createBlank?: (input: { readonly confirmReplace: boolean }) => Promise<{
    readonly fileName: string
    readonly pageCount: number
  }>
  readonly addDrawing?: (drawing: DrawingInput) => string | Promise<string>
  readonly addImageEdit?: (
    edit: ImageEditInput,
    staticFill?: StaticFormFillRecord,
  ) => string | Promise<string>
  readonly addMarkup?: (markup: MarkupInput) => string | Promise<string>
  readonly addTextInsert?: (text: TextInsertInput) => string | Promise<string>
  readonly deleteSavedAnnotation?: (deletion: AnnotDeleteInput) => void | Promise<void>
  readonly updateSavedNote?: (edit: NoteEditInput) => void | Promise<void>
  readonly deletePage?: (pageIndex: number) => void | Promise<void>
  readonly deletePending?: (
    kind: 'markup' | 'drawing' | 'textEdit' | 'textInsert' | 'imageEdit',
    id: string,
  ) => void | Promise<void>
  readonly loadStaged?: (input: {
    readonly name: string
    readonly data: ArrayBuffer
  }) => void | Promise<void>
  readonly insertPagesStaged?: (input: {
    readonly name: string
    readonly data: ArrayBuffer
    readonly afterPageIndex: number
  }) => number | Promise<number>
  readonly insertBlankPage?: (afterPageIndex: number) => void | Promise<void>
  readonly replacePagesStaged?: (input: {
    readonly name: string
    readonly data: ArrayBuffer
    readonly pages: number[]
  }) => { removed: number; inserted: number } | Promise<{ removed: number; inserted: number }>
  readonly cropPages?: (
    pages: number[],
    rect: { l: number; t: number; r: number; b: number },
  ) => void | Promise<void>
  readonly save?: () => boolean | Promise<boolean>
  readonly redo?: () => void | Promise<void>
  readonly replaceText?: (text: TextEditInput) => string | Promise<string>
  readonly setFormValue?: (value: FormValueInput) => void | Promise<void>
  readonly setMetadata?: (metadata: MetadataInput) => void | Promise<void>
  readonly setPageOrder?: (pageOrder: number[]) => void | Promise<void>
  readonly setPageSize?: (width: number, height: number) => void | Promise<void>
  readonly setPageRotation?: (pageIndex: number, rotation: 0 | 90 | 180 | 270) => void | Promise<void>
  readonly setStamps?: (
    watermark: WatermarkConfig | null,
    headerFooter: HeaderFooterConfig | null,
  ) => void | Promise<void>
  readonly setStaticForm?: (
    fill: StaticFormFillRecord,
    image: string,
    oldRect?: [number, number, number, number],
  ) => void | Promise<void>
  readonly undo?: () => void | Promise<void>
  readonly updateDrawing?: (id: string, drawing: DrawingInput) => void | Promise<void>
  readonly updateTextInsert?: (id: string, text: TextInsertInput) => void | Promise<void>
}

type PdfOperationHandlerResult =
  | {
      readonly ok: true
      readonly output: Readonly<Record<string, unknown>>
    }
  | {
      readonly ok: false
      readonly error: PdfRegistryError
      readonly message: string
    }

type PdfOperationHandler = (
  arguments_: Record<string, unknown>,
  services: PdfOperationServices,
) => PdfOperationHandlerResult | Promise<PdfOperationHandlerResult>

type PdfOperationFailure = Extract<PdfOperationHandlerResult, { readonly ok: false }>

function unavailable(message: string): PdfOperationFailure {
  return { ok: false, error: 'execution_failed', message }
}

async function executeMutation(
  action: (() => void | Promise<void>) | undefined,
  success: Readonly<Record<string, unknown>>,
  unavailableMessage: string,
  failureMessage: string,
): Promise<PdfOperationHandlerResult> {
  if (!action) return unavailable(unavailableMessage)
  try {
    await action()
    return { ok: true, output: success }
  } catch {
    return unavailable(failureMessage)
  }
}

async function executeIdMutation(
  action: (() => string | Promise<string>) | undefined,
  unavailableMessage: string,
  failureMessage: string,
): Promise<PdfOperationHandlerResult> {
  if (!action) return unavailable(unavailableMessage)
  try {
    const id = await action()
    if (!id) return unavailable(failureMessage)
    return { ok: true, output: { id } }
  } catch {
    return unavailable(failureMessage)
  }
}

function drawingInput(value: Record<string, unknown>): DrawingInput | null {
  const kind = value.kind
  const pageIndex = value.pageIndex as number
  const color = value.color as [number, number, number] | undefined
  const width = value.width as number | undefined
  if (kind === 'ink' && color && width !== undefined && Array.isArray(value.paths)) {
    const paths = value.paths as number[][]
    if (paths.every((path) => path.length % 2 === 0)) {
      return { kind, pageIndex, color, width, paths, formFieldName: value.formFieldName as string }
    }
  }
  if ((kind === 'rect' || kind === 'ellipse') && color && width !== undefined && value.rect) {
    return {
      kind,
      pageIndex,
      color,
      width,
      rect: value.rect as [number, number, number, number],
    }
  }
  if ((kind === 'line' || kind === 'arrow') && color && width !== undefined && value.from && value.to) {
    return {
      kind,
      pageIndex,
      color,
      width,
      from: value.from as [number, number],
      to: value.to as [number, number],
    }
  }
  if (kind === 'image' && typeof value.image === 'string' && value.rect) {
    return {
      kind,
      pageIndex,
      image: value.image,
      rect: value.rect as [number, number, number, number],
      formFieldName: value.formFieldName as string,
    }
  }
  if (kind === 'note' && color && value.at && typeof value.contents === 'string') {
    return {
      kind,
      pageIndex,
      color,
      at: value.at as [number, number],
      contents: value.contents,
      ...(typeof value.author === 'string' ? { author: value.author } : {}),
      ...(typeof value.createdMs === 'number' ? { createdMs: value.createdMs } : {}),
      ...(typeof value.localId === 'string' ? { localId: value.localId } : {}),
      ...(value.replyToSaved
        ? { replyToSaved: value.replyToSaved as NoteReplyTarget }
        : {}),
      ...(typeof value.replyToLocalId === 'string'
        ? { replyToLocalId: value.replyToLocalId }
        : {}),
    }
  }
  return null
}

export type PdfOperationExecution =
  | { readonly handled: false }
  | ({ readonly handled: true; readonly operationId: PdfOperationId } & PdfOperationHandlerResult)

const handlers = {
  'pdf.document.create_blank': async (arguments_, services) => {
    if (!services.createBlank) return unavailable('The PDF blank-document factory is not ready.')
    try {
      const created = await services.createBlank({
        confirmReplace: arguments_.confirmReplace === true,
      })
      return { ok: true, output: { opened: true, ...created } }
    } catch (error) {
      return unavailable(error instanceof Error ? error.message : 'The blank PDF could not be created.')
    }
  },
  'pdf.annotation.delete_saved': async (arguments_, services) => {
    if (!services.deleteSavedAnnotation) {
      return {
        ok: false,
        error: 'execution_failed',
        message: 'The mounted PDF annotation controller is not ready.',
      }
    }
    const deletion: AnnotDeleteInput = {
      pageIndex: arguments_.pageIndex as number,
      objNum: arguments_.objNum as number,
      subtype: arguments_.subtype as AnnotDeleteInput['subtype'],
      rect: arguments_.rect as [number, number, number, number],
      contents: arguments_.contents as string | undefined,
    }
    try {
      await services.deleteSavedAnnotation(deletion)
      return { ok: true, output: { deleted: deletion.objNum } }
    } catch {
      return {
        ok: false,
        error: 'execution_failed',
        message: 'The saved PDF annotation could not be deleted.',
      }
    }
  },
  'pdf.note.update_saved': (arguments_, services) => {
    const edit = arguments_ as unknown as NoteEditInput
    return executeMutation(
      services.updateSavedNote ? () => services.updateSavedNote!(edit) : undefined,
      { updated: edit.objNum },
      'The mounted PDF comment controller is not ready.',
      'The saved PDF comment could not be updated.',
    )
  },
  'pdf.document.set_metadata': (arguments_, services) =>
    executeMutation(
      services.setMetadata
        ? () => services.setMetadata!(arguments_ as unknown as MetadataInput)
        : undefined,
      { updated: true },
      'The mounted PDF metadata controller is not ready.',
      'The PDF metadata could not be updated.',
    ),
  'pdf.drawing.add': (arguments_, services) => {
    const drawing = drawingInput(arguments_)
    if (!drawing) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'pdf.drawing.add requires the exact fields for its bounded drawing kind.',
      }
    }
    return executeIdMutation(
      services.addDrawing ? () => services.addDrawing!(drawing) : undefined,
      'The mounted PDF drawing controller is not ready.',
      'The PDF drawing could not be added.',
    )
  },
  'pdf.drawing.update': (arguments_, services) => {
    const drawing = drawingInput(arguments_.drawing as Record<string, unknown>)
    if (!drawing) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'pdf.drawing.update requires the exact fields for its bounded drawing kind.',
      }
    }
    const id = arguments_.id as string
    return executeMutation(
      services.updateDrawing ? () => services.updateDrawing!(id, drawing) : undefined,
      { id },
      'The mounted PDF drawing controller is not ready.',
      'The pending PDF drawing could not be updated.',
    )
  },
  'pdf.form.set_value': (arguments_, services) => {
    const kind = arguments_.kind as FormValueInput['kind']
    if (kind === 'checkbox' ? typeof arguments_.checked !== 'boolean' : typeof arguments_.value !== 'string') {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: `pdf.form.set_value requires ${kind === 'checkbox' ? 'checked' : 'value'} for ${kind}.`,
      }
    }
    const value: FormValueInput = {
      name: arguments_.name as string,
      kind,
      ...(kind === 'checkbox'
        ? { checked: arguments_.checked as boolean }
        : { value: arguments_.value as string }),
    }
    return executeMutation(
      services.setFormValue ? () => services.setFormValue!(value) : undefined,
      { field: value.name },
      'The mounted PDF form controller is not ready.',
      'The PDF form field could not be updated.',
    )
  },
  'pdf.history.redo': (_arguments, services) =>
    executeMutation(
      services.redo,
      { redone: true },
      'The mounted PDF history controller is not ready.',
      'The PDF edit could not be redone.',
    ),
  'pdf.image.delete': (arguments_, services) =>
    executeIdMutation(
      services.addImageEdit
        ? () =>
            services.addImageEdit!({
              kind: 'deleteImage',
              pageIndex: arguments_.pageIndex as number,
              oldRect: arguments_.oldRect as [number, number, number, number],
            })
        : undefined,
      'The mounted PDF image controller is not ready.',
      'The PDF image could not be deleted.',
    ),
  'pdf.image.insert': (arguments_, services) =>
    executeIdMutation(
      services.addImageEdit
        ? () =>
            services.addImageEdit!({
              kind: 'insertImage',
              pageIndex: arguments_.pageIndex as number,
              rect: arguments_.rect as [number, number, number, number],
              image: arguments_.image as string,
              layer: arguments_.layer as 'belowText' | 'aboveText',
              rotate: arguments_.rotate as number | undefined,
            })
        : undefined,
      'The mounted PDF image controller is not ready.',
      'The PDF image could not be inserted.',
    ),
  'pdf.image.replace': (arguments_, services) =>
    executeIdMutation(
      services.addImageEdit
        ? () =>
            services.addImageEdit!({
              kind: 'replaceImage',
              pageIndex: arguments_.pageIndex as number,
              oldRect: arguments_.oldRect as [number, number, number, number],
              rect: arguments_.rect as [number, number, number, number],
              image: arguments_.image as string,
              layer: arguments_.layer as 'belowText' | 'aboveText' | undefined,
              quarterTurns: arguments_.quarterTurns as number | undefined,
            })
        : undefined,
      'The mounted PDF image controller is not ready.',
      'The PDF image could not be replaced.',
    ),
  'pdf.image.transform': (arguments_, services) =>
    executeIdMutation(
      services.addImageEdit
        ? () =>
            services.addImageEdit!({
              kind: 'transformImage',
              pageIndex: arguments_.pageIndex as number,
              oldRect: arguments_.oldRect as [number, number, number, number],
              rect: arguments_.rect as [number, number, number, number],
              layer: arguments_.layer as 'belowText' | 'aboveText' | undefined,
              quarterTurns: arguments_.quarterTurns as number | undefined,
            })
        : undefined,
      'The mounted PDF image controller is not ready.',
      'The PDF image could not be transformed.',
    ),
  'pdf.markup.add': (arguments_, services) =>
    executeIdMutation(
      services.addMarkup
        ? () => services.addMarkup!(arguments_ as unknown as MarkupInput)
        : undefined,
      'The mounted PDF markup controller is not ready.',
      'The PDF markup could not be added.',
    ),
  'pdf.page.insert': () =>
    unavailable('pdf.page.insert requires Broker staging before renderer execution.'),
  'pdf.page.insert_staged': async (arguments_, services) => {
    const data = arguments_.data
    if (
      typeof arguments_.name !== 'string' ||
      !arguments_.name.toLowerCase().endsWith('.pdf') ||
      !Number.isInteger(arguments_.size) ||
      !isArrayBuffer(data) ||
      data.byteLength !== arguments_.size ||
      !hasPdfMagic(data)
    ) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'pdf.page.insert_staged requires matching bounded hydrated PDF bytes.',
      }
    }
    if (!services.insertPagesStaged) {
      return unavailable('The mounted PDF page-insert controller is not ready.')
    }
    try {
      const insertedCount = await services.insertPagesStaged({
        name: arguments_.name,
        data,
        afterPageIndex: arguments_.afterPageIndex as number,
      })
      if (!Number.isInteger(insertedCount) || insertedCount < 1) {
        return unavailable('The staged PDF did not contain an insertable page.')
      }
      return { ok: true, output: { insertedCount } }
    } catch {
      return unavailable('The staged PDF pages could not be inserted.')
    }
  },
  'pdf.page.delete': (arguments_, services) => {
    const pageIndex = arguments_.pageIndex as number
    return executeMutation(
      services.deletePage ? () => services.deletePage!(pageIndex) : undefined,
      { deletedPage: pageIndex },
      'The mounted PDF page controller is not ready.',
      'The PDF page could not be deleted.',
    )
  },
  'pdf.page.replace': () =>
    unavailable('pdf.page.replace requires Broker staging before renderer execution.'),
  'pdf.page.replace_staged': async (arguments_, services) => {
    const data = arguments_.data
    const pages = arguments_.pages as number[]
    if (
      typeof arguments_.name !== 'string' ||
      !arguments_.name.toLowerCase().endsWith('.pdf') ||
      !Number.isInteger(arguments_.size) ||
      !isArrayBuffer(data) ||
      data.byteLength !== arguments_.size ||
      !hasPdfMagic(data) ||
      new Set(pages).size !== pages.length
    ) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'pdf.page.replace_staged requires matching PDF bytes and unique pages.',
      }
    }
    if (!services.replacePagesStaged) {
      return unavailable('The mounted PDF page-replace controller is not ready.')
    }
    try {
      const result = await services.replacePagesStaged({
        name: arguments_.name,
        data,
        pages,
      })
      if (result.removed < 1 || result.inserted < 1) {
        return unavailable('The staged PDF did not contain replaceable pages.')
      }
      return { ok: true, output: result }
    } catch {
      return unavailable('The staged PDF pages could not replace the selected pages.')
    }
  },
  'pdf.page.insert_blank': (arguments_, services) => {
    const afterPageIndex = arguments_.afterPageIndex as number
    return executeMutation(
      services.insertBlankPage ? () => services.insertBlankPage!(afterPageIndex) : undefined,
      { insertedAfterPage: afterPageIndex },
      'The mounted PDF blank-page controller is not ready.',
      'The blank PDF page could not be inserted.',
    )
  },
  'pdf.page.crop': (arguments_, services) => {
    const pages = arguments_.pages as number[]
    const rect = arguments_.rect as { l: number; t: number; r: number; b: number }
    if (new Set(pages).size !== pages.length || rect.l >= rect.r || rect.t >= rect.b) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'pdf.page.crop requires unique pages and a non-empty fractional rectangle.',
      }
    }
    return executeMutation(
      services.cropPages ? () => services.cropPages!(pages, rect) : undefined,
      { croppedPages: pages.length },
      'The mounted PDF crop controller is not ready.',
      'The PDF pages could not be cropped.',
    )
  },
  'pdf.page.reorder': (arguments_, services) => {
    const pageOrder = arguments_.pageOrder as number[]
    if (new Set(pageOrder).size !== pageOrder.length) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'pdf.page.reorder requires unique original page indices.',
      }
    }
    return executeMutation(
      services.setPageOrder ? () => services.setPageOrder!(pageOrder) : undefined,
      { pageCount: pageOrder.length },
      'The mounted PDF page controller is not ready.',
      'The PDF pages could not be reordered.',
    )
  },
  'pdf.page.set_size': (arguments_, services) => {
    const width = arguments_.width as number
    const height = arguments_.height as number
    return executeMutation(
      services.setPageSize ? () => services.setPageSize!(width, height) : undefined,
      { width, height },
      'The mounted PDF page-size controller is not ready.',
      'The PDF page size could not be changed.',
    )
  },
  'pdf.page.set_rotation': (arguments_, services) => {
    const rotation = arguments_.rotation as 0 | 90 | 180 | 270
    return executeMutation(
      services.setPageRotation
        ? () => services.setPageRotation!(arguments_.pageIndex as number, rotation)
        : undefined,
      { rotation },
      'The mounted PDF page controller is not ready.',
      'The PDF page rotation could not be updated.',
    )
  },
  'pdf.pending.delete': (arguments_, services) => {
    const id = arguments_.id as string
    return executeMutation(
      services.deletePending
        ? () =>
            services.deletePending!(
              arguments_.kind as 'markup' | 'drawing' | 'textEdit' | 'textInsert' | 'imageEdit',
              id,
            )
        : undefined,
      { id },
      'The mounted PDF pending-edit controller is not ready.',
      'The pending PDF edit could not be deleted.',
    )
  },
  'pdf.stamp.set': (arguments_, services) =>
    executeMutation(
      services.setStamps
        ? () =>
            services.setStamps!(
              arguments_.watermark as WatermarkConfig | null,
              arguments_.headerFooter as HeaderFooterConfig | null,
            )
        : undefined,
      { updated: true },
      'The mounted PDF stamp controller is not ready.',
      'The PDF watermark or header/footer could not be updated.',
    ),
  'pdf.static_form.set': (arguments_, services) => {
    const fill: StaticFormFillRecord = {
      id: arguments_.id as string,
      kind: arguments_.kind as StaticFormFillRecord['kind'],
      pageIndex: arguments_.pageIndex as number,
      rect: arguments_.rect as [number, number, number, number],
      text: arguments_.text as string | undefined,
      fontSize: arguments_.fontSize as number | undefined,
      color: arguments_.color as string | undefined,
      align: arguments_.align as StaticFormFillRecord['align'],
    }
    return executeMutation(
      services.setStaticForm
        ? () =>
            services.setStaticForm!(
              fill,
              arguments_.image as string,
              arguments_.oldRect as [number, number, number, number] | undefined,
            )
        : undefined,
      { id: fill.id },
      'The mounted PDF static-form controller is not ready.',
      'The PDF static form fill could not be updated.',
    )
  },
  'pdf.text.insert': (arguments_, services) =>
    executeIdMutation(
      services.addTextInsert
        ? () => services.addTextInsert!(arguments_ as unknown as TextInsertInput)
        : undefined,
      'The mounted PDF text controller is not ready.',
      'The PDF text could not be inserted.',
    ),
  'pdf.text.replace': (arguments_, services) =>
    executeIdMutation(
      services.replaceText
        ? () => services.replaceText!(arguments_ as unknown as TextEditInput)
        : undefined,
      'The mounted PDF text controller is not ready.',
      'The PDF text could not be replaced.',
    ),
  'pdf.text.update_inserted': (arguments_, services) => {
    const id = arguments_.id as string
    return executeMutation(
      services.updateTextInsert
        ? () => services.updateTextInsert!(id, arguments_.text as unknown as TextInsertInput)
        : undefined,
      { id },
      'The mounted PDF text controller is not ready.',
      'The pending inserted PDF text could not be updated.',
    )
  },
  'pdf.document.load_staged': async (arguments_, services) => {
    if (
      typeof arguments_.name !== 'string' ||
      !arguments_.name.toLowerCase().endsWith('.pdf') ||
      !Number.isInteger(arguments_.size) ||
      !isArrayBuffer(arguments_.data) ||
      arguments_.data.byteLength !== arguments_.size
    ) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'pdf.document.load_staged requires a valid hydrated PDF descriptor.',
      }
    }
    if (!services.loadStaged) {
      return {
        ok: false,
        error: 'execution_failed',
        message: 'The mounted PDF load controller is not ready.',
      }
    }
    try {
      await services.loadStaged({ name: arguments_.name, data: arguments_.data })
      return { ok: true, output: { opened: arguments_.name } }
    } catch {
      return {
        ok: false,
        error: 'execution_failed',
        message: 'The staged PDF could not be loaded.',
      }
    }
  },
  'pdf.document.save': async (_arguments, services) => {
    try {
      if (!services.save || !(await services.save())) {
        return {
          ok: false,
          error: 'execution_failed',
          message: 'The PDF save was canceled or failed.',
        }
      }
    } catch {
      return {
        ok: false,
        error: 'execution_failed',
        message: 'The PDF save was canceled or failed.',
      }
    }
    return { ok: true, output: { saved: true } }
  },
  'pdf.history.undo': async (_arguments, services) => {
    if (!services.undo) {
      return {
        ok: false,
        error: 'execution_failed',
        message: 'The mounted PDF history controller is not ready.',
      }
    }
    try {
      await services.undo()
      return { ok: true, output: { undone: true } }
    } catch {
      return {
        ok: false,
        error: 'execution_failed',
        message: 'The PDF edit could not be undone.',
      }
    }
  },
} satisfies Record<PdfOperationId, PdfOperationHandler>

function isArrayBuffer(value: unknown): value is ArrayBuffer {
  return Object.prototype.toString.call(value) === '[object ArrayBuffer]'
}

function hasPdfMagic(data: ArrayBuffer): boolean {
  const bytes = new Uint8Array(data, 0, Math.min(data.byteLength, 5))
  return bytes.length === 5 && String.fromCharCode(...bytes) === '%PDF-'
}

export const pdfOperationHandlerIds = Object.freeze(
  Object.keys(handlers).sort() as PdfOperationId[],
)

const descriptorsByName = new Map<string, PdfOperationDescriptor>()
for (const descriptor of pdfOperationCatalog.operations) {
  descriptorsByName.set(descriptor.id, descriptor)
  for (const alias of descriptor.compatibilityAliases) {
    descriptorsByName.set(alias, descriptor)
  }
}

export function resolvePdfOperationId(operation: string): PdfOperationId | null {
  return descriptorsByName.get(operation)?.id ?? null
}

export async function executePdfOperation(
  command: PdfOperationCommand,
  services: PdfOperationServices,
): Promise<PdfOperationExecution> {
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

  const result = await handlers[descriptor.id](command.arguments, services)
  return { handled: true, operationId: descriptor.id, ...result }
}
