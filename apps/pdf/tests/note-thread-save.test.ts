import {
  PDFArray,
  PDFDict,
  PDFDocument,
  PDFHexString,
  PDFName,
  PDFRef,
  PDFString,
} from 'pdf-lib'
import { describe, expect, it } from 'vitest'
import { applySaveRequest } from '../src/domain/save-pdf'
import type { DrawingInput, SavePdfRequest } from '../src/shared/ipc'

type NoteInput = Extract<DrawingInput, { kind: 'note' }>

async function blankPdf(): Promise<Uint8Array> {
  const document = await PDFDocument.create()
  document.addPage([595, 842])
  return document.save({ useObjectStreams: false })
}

function saveRequest(drawings: DrawingInput[]): SavePdfRequest {
  return {
    path: 'thread.pdf',
    markups: [],
    drawings,
    formValues: [],
    stamps: [],
  }
}

function note(overrides: Partial<NoteInput>): NoteInput {
  return {
    kind: 'note',
    pageIndex: 0,
    color: [1, 0.78, 0.13],
    at: [200, 500],
    contents: 'Comment',
    ...overrides,
  }
}

describe('saved PDF comment threads', () => {
  it('persists the author and timestamps of a new root comment', async () => {
    const saved = await applySaveRequest(
      await blankPdf(),
      saveRequest([
        note({ contents: '你好，批注', author: 'hong', createdMs: 1_786_511_400_000 }),
      ]),
    )

    const reopened = await PDFDocument.load(saved.bytes)
    const annotations = reopened.getPage(0).node.lookup(PDFName.of('Annots'), PDFArray)
    const annotation = annotations.lookup(0, PDFDict)
    const author = annotation.lookup(PDFName.of('T'))
    const created = annotation.lookup(PDFName.of('CreationDate'))
    const modified = annotation.lookup(PDFName.of('M'))

    expect(author instanceof PDFHexString && author.decodeText()).toBe('hong')
    expect(created instanceof PDFString && created.decodeText()).toMatch(/^D:20\d{12}[+-]/)
    expect(modified instanceof PDFString && modified.decodeText()).toMatch(/^D:20\d{12}[+-]/)
  })

  it('links a reply to its saved parent after reopen even when the object hint is stale', async () => {
    const rootSave = await applySaveRequest(
      await blankPdf(),
      saveRequest([note({ contents: 'Root comment' })]),
    )
    const withReply = await applySaveRequest(
      rootSave.bytes,
      saveRequest([
        note({
          contents: 'Reply comment',
          replyToSaved: {
            objNum: 999_999,
            rect: [200, 482, 220, 500],
            contents: 'Root comment',
          },
        }),
      ]),
    )

    const reopened = await PDFDocument.load(withReply.bytes)
    const annotations = reopened.getPage(0).node.lookup(PDFName.of('Annots'), PDFArray)
    const rootRef = annotations.get(0)
    const reply = annotations.lookup(1, PDFDict)

    expect(rootRef).toBeInstanceOf(PDFRef)
    expect(reply.get(PDFName.of('IRT'))).toBe(rootRef)
    expect(reply.lookup(PDFName.of('RT'))).toBe(PDFName.of('R'))
  })

  it('updates a saved comment in place so existing reply identities remain stable', async () => {
    const rootSave = await applySaveRequest(
      await blankPdf(),
      saveRequest([note({ contents: 'Before edit' })]),
    )
    const firstOpen = await PDFDocument.load(rootSave.bytes)
    const rootRef = firstOpen
      .getPage(0)
      .node.lookup(PDFName.of('Annots'), PDFArray)
      .get(0) as PDFRef
    const request = saveRequest([])
    request.noteEdits = [
      {
        pageIndex: 0,
        objNum: rootRef.objectNumber,
        rect: [200, 482, 220, 500],
        oldContents: 'Before edit',
        contents: 'After edit',
      },
    ]

    const updated = await applySaveRequest(rootSave.bytes, request)
    const reopened = await PDFDocument.load(updated.bytes)
    const annotations = reopened.getPage(0).node.lookup(PDFName.of('Annots'), PDFArray)
    const reopenedRef = annotations.get(0) as PDFRef
    const annotation = annotations.lookup(0, PDFDict)
    const contents = annotation.lookup(PDFName.of('Contents'))

    expect(reopenedRef.objectNumber).toBe(rootRef.objectNumber)
    expect(contents instanceof PDFHexString && contents.decodeText()).toBe('After edit')
  })

  it('deletes exactly one saved reply when every thread member shares the root rectangle', async () => {
    const threaded = await applySaveRequest(
      await blankPdf(),
      saveRequest([
        note({ contents: 'Root', localId: 'root' }),
        note({ contents: 'Reply', replyToLocalId: 'root' }),
      ]),
    )
    const request = saveRequest([])
    request.annotDeletes = [
      {
        pageIndex: 0,
        objNum: 999_999,
        subtype: 'note',
        rect: [200, 482, 220, 500],
        contents: 'Reply',
      },
    ]

    const deleted = await applySaveRequest(threaded.bytes, request)
    const reopened = await PDFDocument.load(deleted.bytes)
    const annotations = reopened.getPage(0).node.lookup(PDFName.of('Annots'), PDFArray)
    const contents = Array.from({ length: annotations.size() }, (_, index) => {
      const value = annotations.lookup(index, PDFDict).lookup(PDFName.of('Contents'))
      return value instanceof PDFString || value instanceof PDFHexString ? value.decodeText() : ''
    })

    expect(contents).toEqual(['Root'])
  })
})
