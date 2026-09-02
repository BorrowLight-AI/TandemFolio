import { PDFArray, PDFDict, PDFDocument, PDFName } from 'pdf-lib'
import { describe, expect, it, vi } from 'vitest'
import { createBrowserPdfHost } from '../src/renderer/host/browser-pdf-api'
import { createPdfCommunityCommandBridge } from '../src/renderer/host/community-command-bridge'
import type { AnnotDeleteInput, SavePdfRequest } from '../src/shared/ipc'

async function annotatedFixture(): Promise<{
  bytes: Uint8Array
  highlight: AnnotDeleteInput
}> {
  const document = await PDFDocument.create()
  const page = document.addPage([420, 300])
  const annotations = document.context.obj([]) as PDFArray
  page.node.set(PDFName.of('Annots'), annotations)

  const add = (subtype: 'Highlight' | 'Underline', rect: [number, number, number, number]) => {
    const annotation = document.context.obj({
      Type: 'Annot',
      Subtype: subtype,
      Rect: rect,
      QuadPoints: [rect[0], rect[3], rect[2], rect[3], rect[0], rect[1], rect[2], rect[1]],
      C: [1, 0.87, 0.35],
      F: 4,
      P: page.ref,
    })
    const reference = document.context.register(annotation)
    annotations.push(reference)
    return reference
  }

  const highlightReference = add('Highlight', [40, 220, 180, 242])
  add('Underline', [40, 180, 180, 202])
  return {
    bytes: await document.save({ useObjectStreams: false }),
    highlight: {
      pageIndex: 0,
      objNum: highlightReference.objectNumber,
      subtype: 'highlight',
      rect: [40, 220, 180, 242],
    },
  }
}

function request(path: string, deletion: AnnotDeleteInput): SavePdfRequest {
  return {
    path,
    markups: [],
    annotDeletes: [deletion],
    drawings: [],
    formValues: [],
    stamps: [],
  }
}

describe('PDF browser save boundary', () => {
  it('deletes the addressed saved annotation and downloads reopenable PDF bytes', async () => {
    const downloaded = vi.fn<(fileName: string, data: ArrayBuffer) => void>()
    const host = createBrowserPdfHost({ download: downloaded })
    const fixture = await annotatedFixture()
    host.stageFile('review.pdf', fixture.bytes.buffer.slice(0) as ArrayBuffer)
    const path = await host.api.consumePending()
    expect(path).not.toBeNull()

    await expect(host.api.save(request(path!, fixture.highlight))).resolves.toMatchObject({
      ok: true,
    })

    expect(downloaded).toHaveBeenCalledOnce()
    expect(downloaded.mock.calls[0]?.[0]).toBe('review.pdf')
    const reopened = await PDFDocument.load(downloaded.mock.calls[0]![1])
    const annotations = reopened
      .getPage(0)
      .node.lookup(PDFName.of('Annots'), PDFArray)
      .asArray()
      .map((reference) => reopened.context.lookup(reference, PDFDict))
    expect(
      annotations.map((annotation) =>
        annotation.lookup(PDFName.of('Subtype'), PDFName).decodeText(),
      ),
    ).toEqual(['Underline'])
  })

  it('serializes pending PDF edits into recovery bytes without downloading or replacing the source', async () => {
    const downloaded = vi.fn<(fileName: string, data: ArrayBuffer) => void>()
    const host = createBrowserPdfHost({ download: downloaded })
    const fixture = await annotatedFixture()
    void host.stageFile('review.pdf', fixture.bytes.buffer.slice(0) as ArrayBuffer)
    const path = await host.api.consumePending()
    expect(path).not.toBeNull()

    const recovery = await host.api.createRecovery(request(path!, fixture.highlight))

    expect(downloaded).not.toHaveBeenCalled()
    const recovered = await PDFDocument.load(recovery)
    const recoveredAnnotations = recovered
      .getPage(0)
      .node.lookup(PDFName.of('Annots'), PDFArray)
      .asArray()
    expect(recoveredAnnotations).toHaveLength(1)
    const original = await PDFDocument.load(await host.api.readFile(path!))
    const originalAnnotations = original
      .getPage(0)
      .node.lookup(PDFName.of('Annots'), PDFArray)
      .asArray()
    expect(originalAnnotations).toHaveLength(2)
  })

  it('exposes the mounted PDF renderer recovery snapshot through the live adapter', async () => {
    const commandBridge = createPdfCommunityCommandBridge()
    const host = createBrowserPdfHost({ commandBridge, download: () => undefined })
    const fixture = await annotatedFixture()
    void host.stageFile('review.pdf', fixture.bytes.buffer.slice(0) as ArrayBuffer)
    await host.api.consumePending()
    host.api.setDirty(true)
    commandBridge.register({
      recoverySnapshot: async () => ({
        fileName: 'review.pdf',
        data: fixture.bytes.buffer.slice(0) as ArrayBuffer,
      }),
    } as never)

    const recovery = await host.adapter.recoverySnapshot?.()

    expect(recovery?.fileName).toBe('review.pdf')
    expect(new Uint8Array(recovery!.data)).toEqual(fixture.bytes)
  })

  it('inserts staged PDF pages through the live adapter and downloads reopenable merged bytes', async () => {
    const base = await PDFDocument.create()
    base.addPage([100, 100])
    const inserted = await PDFDocument.create()
    inserted.addPage([200, 200])
    inserted.addPage([300, 300])
    const baseBytes = await base.save({ useObjectStreams: false })
    const insertedBytes = await inserted.save({ useObjectStreams: false })
    const downloaded = vi.fn<(fileName: string, data: ArrayBuffer) => void>()
    const host = createBrowserPdfHost({ download: downloaded })
    host.stageFile('base.pdf', baseBytes.buffer.slice(0) as ArrayBuffer)
    await host.api.consumePending()
    const data = insertedBytes.buffer.slice(0) as ArrayBuffer

    await expect(
      host.adapter.execute({
        commandId: 'insert-pages',
        baseRevision: 0,
        operation: 'pdf.page.insert_staged',
        arguments: {
          blobId: 'insert-blob',
          name: 'insert.pdf',
          size: data.byteLength,
          data,
          afterPageIndex: 0,
        },
      }),
    ).resolves.toEqual({ ok: true, output: { insertedCount: 2 } })

    expect(downloaded).toHaveBeenCalledOnce()
    expect(downloaded.mock.calls[0]?.[0]).toBe('base.pdf')
    const reopened = await PDFDocument.load(downloaded.mock.calls[0]![1])
    expect(reopened.getPageCount()).toBe(3)
    expect(reopened.getPages().map((page) => page.getWidth())).toEqual([100, 200, 300])
  })
})
