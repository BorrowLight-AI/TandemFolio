import { describe, expect, it, vi } from 'vitest'
import {
  validateJsonSchemaValue,
  validateOperationCatalogs,
} from '@tandemfolio/operation-contract'

import { createPdfCommunityCommandBridge } from '../src/renderer/host/community-command-bridge'
import type { PdfCommunityController } from '../src/renderer/host/community-command-bridge'
import {
  pdfRetainedProducerBaseline,
  validatePdfRetainedProducerBaseline,
} from '../src/renderer/operations/baseline'
import { pdfOperationCatalog } from '../src/renderer/operations/catalog'
import { executePdfOperation, pdfOperationHandlerIds } from '../src/renderer/operations/registry'

function controller(overrides: Partial<PdfCommunityController> = {}): PdfCommunityController {
  return {
    addDrawing: async () => 'drawing-id',
    addImageEdit: async () => 'image-id',
    addMarkup: async () => 'markup-id',
    addTextInsert: async () => 'text-insert-id',
    deletePage: async () => undefined,
    deletePending: async () => undefined,
    deleteSavedAnnotation: async () => undefined,
    redo: async () => undefined,
    recoverySnapshot: async () => null,
    recoveryVersion: () => 0,
    replaceText: async () => 'text-edit-id',
    save: async () => true,
    setFormValue: async () => undefined,
    setMetadata: async () => undefined,
    setPageOrder: async () => undefined,
    setPageRotation: async () => undefined,
    setStamps: async () => undefined,
    setStaticForm: async () => undefined,
    undo: async () => undefined,
    updateDrawing: async () => undefined,
    updateTextInsert: async () => undefined,
    ...overrides,
  }
}

describe('PDF operation registry', () => {
  it('maps every retained state-changing producer to an implemented Registry operation', () => {
    expect(pdfRetainedProducerBaseline.length).toBeGreaterThan(20)
    expect(validatePdfRetainedProducerBaseline()).toEqual([])
  })

  it('keeps every serializable descriptor paired with exactly one handler', () => {
    expect(validateOperationCatalogs([pdfOperationCatalog])).toEqual({ ok: true })
    expect(pdfOperationHandlerIds).toEqual(
      pdfOperationCatalog.operations.map((operation) => operation.id).sort(),
    )
  })

  it('owns the complete retained PDF mutation surface as 23 public and two internal operations', () => {
    expect(pdfOperationCatalog.operations).toHaveLength(25)
    expect(
      pdfOperationCatalog.operations.filter((operation) => operation.visibility === 'agent'),
    ).toHaveLength(23)
    expect(
      pdfOperationCatalog.operations
        .filter((operation) => operation.visibility === 'internal')
        .map((operation) => operation.id),
    ).toEqual(['pdf.page.insert_staged', 'pdf.document.load_staged'])
  })

  it('routes retained final-state families through the mounted controller', async () => {
    const calls: [string, unknown[]][] = []
    const record =
      <T extends unknown[], R>(name: string, result: R) =>
      (...args: T): R => {
        calls.push([name, args])
        return result
      }
    const services = controller({
      addDrawing: record('addDrawing', 'drawing-1'),
      addImageEdit: record('addImageEdit', 'image-1'),
      addMarkup: record('addMarkup', 'markup-1'),
      addTextInsert: record('addTextInsert', 'insert-1'),
      deletePage: record('deletePage', undefined),
      deletePending: record('deletePending', undefined),
      redo: record('redo', undefined),
      replaceText: record('replaceText', 'replace-1'),
      setFormValue: record('setFormValue', undefined),
      setMetadata: record('setMetadata', undefined),
      setPageOrder: record('setPageOrder', undefined),
      setPageRotation: record('setPageRotation', undefined),
      setStamps: record('setStamps', undefined),
      setStaticForm: record('setStaticForm', undefined),
      updateDrawing: record('updateDrawing', undefined),
      updateTextInsert: record('updateTextInsert', undefined),
    })
    const commands = [
      ['pdf.history.redo', {}, { redone: true }],
      [
        'pdf.markup.add',
        {
          pageIndex: 0,
          type: 'highlight',
          color: [1, 0.8, 0],
          quads: [[10, 20, 30, 20, 10, 10, 30, 10]],
        },
        { id: 'markup-1' },
      ],
      [
        'pdf.drawing.add',
        { kind: 'rect', pageIndex: 0, color: [1, 0, 0], width: 2, rect: [1, 2, 3, 4] },
        { id: 'drawing-1' },
      ],
      [
        'pdf.drawing.update',
        {
          id: 'drawing-1',
          drawing: { kind: 'note', pageIndex: 0, color: [1, 1, 0], at: [5, 6], contents: 'N' },
        },
        { id: 'drawing-1' },
      ],
      [
        'pdf.form.set_value',
        { name: 'approved', kind: 'checkbox', checked: true },
        { field: 'approved' },
      ],
      [
        'pdf.text.replace',
        { pageIndex: 0, rect: [1, 2, 3, 4], oldText: 'A', newText: 'B', fontSize: 12 },
        { id: 'replace-1' },
      ],
      [
        'pdf.text.insert',
        { pageIndex: 0, origin: [20, 30], text: 'Inserted', fontSize: 12, color: [0, 0, 0] },
        { id: 'insert-1' },
      ],
      [
        'pdf.text.update_inserted',
        {
          id: 'insert-1',
          text: { pageIndex: 0, origin: [21, 31], text: 'Updated', fontSize: 13, color: [1, 2, 3] },
        },
        { id: 'insert-1' },
      ],
      [
        'pdf.image.insert',
        { pageIndex: 0, rect: [1, 2, 3, 4], image: 'png', layer: 'belowText' },
        { id: 'image-1' },
      ],
      [
        'pdf.image.transform',
        { pageIndex: 0, oldRect: [1, 2, 3, 4], rect: [2, 3, 5, 6], quarterTurns: 1 },
        { id: 'image-1' },
      ],
      [
        'pdf.image.replace',
        { pageIndex: 0, oldRect: [1, 2, 3, 4], rect: [1, 2, 3, 4], image: 'png2' },
        { id: 'image-1' },
      ],
      ['pdf.image.delete', { pageIndex: 0, oldRect: [1, 2, 3, 4] }, { id: 'image-1' }],
      ['pdf.pending.delete', { kind: 'drawing', id: 'drawing-1' }, { id: 'drawing-1' }],
      [
        'pdf.static_form.set',
        { id: 'fill-1', kind: 'check', pageIndex: 0, rect: [1, 2, 3, 4], image: 'png' },
        { id: 'fill-1' },
      ],
      [
        'pdf.document.set_metadata',
        { title: 'T', author: 'A', subject: 'S', keywords: 'K' },
        { updated: true },
      ],
      ['pdf.page.set_rotation', { pageIndex: 0, rotation: 90 }, { rotation: 90 }],
      ['pdf.page.delete', { pageIndex: 2 }, { deletedPage: 2 }],
      ['pdf.page.reorder', { pageOrder: [2, 0, 1] }, { pageCount: 3 }],
      ['pdf.stamp.set', { watermark: null, headerFooter: null }, { updated: true }],
    ] as const

    for (const [operation, arguments_, output] of commands) {
      await expect(
        executePdfOperation({ operation, arguments: arguments_ }, services),
      ).resolves.toMatchObject({
        handled: true,
        operationId: operation,
        ok: true,
        output,
      })
    }

    expect(calls.map(([name]) => name)).toEqual([
      'redo',
      'addMarkup',
      'addDrawing',
      'updateDrawing',
      'setFormValue',
      'replaceText',
      'addTextInsert',
      'updateTextInsert',
      'addImageEdit',
      'addImageEdit',
      'addImageEdit',
      'addImageEdit',
      'deletePending',
      'setStaticForm',
      'setMetadata',
      'setPageRotation',
      'deletePage',
      'setPageOrder',
      'setStamps',
    ])
  })

  it('rejects conditional drawing/form and duplicate page-order states before mutation', async () => {
    const addDrawing = vi.fn(async () => 'drawing')
    const setFormValue = vi.fn(async () => undefined)
    const setPageOrder = vi.fn(async () => undefined)

    await expect(
      executePdfOperation(
        { operation: 'pdf.drawing.add', arguments: { kind: 'rect', pageIndex: 0 } },
        { addDrawing },
      ),
    ).resolves.toMatchObject({ ok: false, error: 'invalid_arguments' })
    await expect(
      executePdfOperation(
        { operation: 'pdf.form.set_value', arguments: { name: 'x', kind: 'checkbox' } },
        { setFormValue },
      ),
    ).resolves.toMatchObject({ ok: false, error: 'invalid_arguments' })
    await expect(
      executePdfOperation(
        { operation: 'pdf.page.reorder', arguments: { pageOrder: [0, 0] } },
        { setPageOrder },
      ),
    ).resolves.toMatchObject({ ok: false, error: 'invalid_arguments' })

    expect(addDrawing).not.toHaveBeenCalled()
    expect(setFormValue).not.toHaveBeenCalled()
    expect(setPageOrder).not.toHaveBeenCalled()
  })

  it('hydrates and inserts a bounded staged PDF only after magic and byte-length checks', async () => {
    const insertPagesStaged = vi.fn(async () => 2)
    const data = new TextEncoder().encode('%PDF-1.7\nfixture').buffer

    await expect(
      executePdfOperation(
        {
          operation: 'pdf.page.insert_staged',
          arguments: {
            blobId: 'insert-blob',
            name: 'insert.pdf',
            size: data.byteLength,
            data,
            afterPageIndex: 1,
          },
        },
        { insertPagesStaged },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'pdf.page.insert_staged',
      ok: true,
      output: { insertedCount: 2 },
    })
    expect(insertPagesStaged).toHaveBeenCalledWith({ name: 'insert.pdf', data, afterPageIndex: 1 })

    await expect(
      executePdfOperation(
        {
          operation: 'pdf.page.insert_staged',
          arguments: {
            blobId: 'bad',
            name: 'insert.pdf',
            size: 4,
            data: new TextEncoder().encode('nope').buffer,
            afterPageIndex: 0,
          },
        },
        { insertPagesStaged },
      ),
    ).resolves.toMatchObject({ ok: false, error: 'invalid_arguments' })
    expect(insertPagesStaged).toHaveBeenCalledTimes(1)
  })

  it('declares the exact PDF save contract without inventing a file name', () => {
    const descriptor = pdfOperationCatalog.operations.find(
      (operation) => operation.id === 'pdf.document.save',
    )
    expect(descriptor).toBeDefined()
    if (!descriptor) throw new Error('Missing pdf.document.save descriptor.')

    expect(descriptor).toMatchObject({
      id: 'pdf.document.save',
      format: 'pdf',
      family: 'document',
      visibility: 'agent',
      context: ['document'],
      effects: ['persistence'],
      risk: 'medium',
      mutates: true,
      undoable: false,
      atomic: true,
      compatibilityAliases: [],
    })
    expect(validateJsonSchemaValue(descriptor.inputSchema, {})).toEqual({ ok: true })
    expect(validateJsonSchemaValue(descriptor.inputSchema, { extra: true }).ok).toBe(false)
    expect(validateJsonSchemaValue(descriptor.outputSchema, { saved: true })).toEqual({ ok: true })
    expect(
      validateJsonSchemaValue(descriptor.outputSchema, { saved: true, fileName: 'x.pdf' }).ok,
    ).toBe(false)
  })

  it('executes the canonical save through the injected mounted-App save service', async () => {
    const save = vi.fn(async () => true)

    await expect(
      executePdfOperation({ operation: 'pdf.document.save', arguments: {} }, { save }),
    ).resolves.toEqual({
      handled: true,
      operationId: 'pdf.document.save',
      ok: true,
      output: { saved: true },
    })
    expect(save).toHaveBeenCalledTimes(1)
  })

  it('rejects every retired public PDF alias before entering mounted-App execution', async () => {
    const save = vi.fn(async () => true)

    for (const operation of [
      'save',
      'undo',
      'redo',
      'delete_saved_annotation',
      'open_local_file',
    ] as const) {
      await expect(executePdfOperation({ operation, arguments: {} }, { save })).resolves.toEqual({
        handled: false,
      })
    }
    expect(save).not.toHaveBeenCalled()
  })

  it('rejects extra save arguments before calling the persistence service', async () => {
    const save = vi.fn(async () => true)

    await expect(
      executePdfOperation({ operation: 'pdf.document.save', arguments: { extra: true } }, { save }),
    ).resolves.toMatchObject({
      handled: true,
      operationId: 'pdf.document.save',
      ok: false,
      error: 'invalid_arguments',
    })
    expect(save).not.toHaveBeenCalled()
  })

  it.each([
    ['cancellation', async () => false],
    ['failure', async () => Promise.reject(new Error('write failed'))],
  ])('maps save %s to execution_failed', async (_case, save) => {
    await expect(
      executePdfOperation({ operation: 'pdf.document.save', arguments: {} }, { save }),
    ).resolves.toEqual({
      handled: true,
      operationId: 'pdf.document.save',
      ok: false,
      error: 'execution_failed',
      message: 'The PDF save was canceled or failed.',
    })
  })

  it('routes canonical bridge saves and rejects the retired public alias', async () => {
    const bridge = createPdfCommunityCommandBridge()
    const save = vi.fn(async () => true)
    bridge.register(
      controller({
        deleteSavedAnnotation: async () => undefined,
        undo: async () => undefined,
        save,
      }),
    )

    await expect(bridge.execute('pdf.document.save', {})).resolves.toEqual({
      ok: true,
      output: { saved: true },
    })
    await expect(bridge.execute('save', {})).resolves.toMatchObject({
      ok: false,
      error: 'unsupported_operation',
    })
    expect(save).toHaveBeenCalledOnce()
  })

  it('maps a missing mounted staged-load service to execution_failed', async () => {
    const data = Uint8Array.from([0x25, 0x50, 0x44, 0x46]).buffer

    await expect(
      executePdfOperation(
        {
          operation: 'pdf.document.load_staged',
          arguments: {
            blobId: 'pdf-blob',
            name: 'loaded.pdf',
            size: data.byteLength,
            data,
          },
        },
        {},
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'pdf.document.load_staged',
      ok: false,
      error: 'execution_failed',
      message: 'The mounted PDF load controller is not ready.',
    })
  })

  it('deletes an exact saved annotation through the canonical mounted-App operation', async () => {
    const bridge = createPdfCommunityCommandBridge()
    const deletions: unknown[] = []
    bridge.register(
      controller({
        deleteSavedAnnotation: async (deletion) => {
          deletions.push(deletion)
        },
        undo: async () => undefined,
        save: async () => true,
      }),
    )
    const deletion = {
      pageIndex: 0,
      objNum: 17,
      subtype: 'highlight',
      rect: [40, 220, 180, 242],
    }

    await expect(bridge.execute('pdf.annotation.delete_saved', deletion)).resolves.toEqual({
      ok: true,
      output: { deleted: 17 },
    })
    expect(deletions).toEqual([deletion])
  })

  it('undoes the latest mounted-App edit through the canonical history operation', async () => {
    const bridge = createPdfCommunityCommandBridge()
    const history = ['saved annotation deletion']
    bridge.register(
      controller({
        deleteSavedAnnotation: async () => undefined,
        undo: async () => {
          history.pop()
        },
        save: async () => true,
      }),
    )

    await expect(bridge.execute('pdf.history.undo', {})).resolves.toEqual({
      ok: true,
      output: { undone: true },
    })
    expect(history).toEqual([])
  })
})
