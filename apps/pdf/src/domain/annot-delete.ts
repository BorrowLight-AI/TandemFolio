import { PDFArray, PDFDict, PDFDocument, PDFName, PDFNumber, PDFRef } from 'pdf-lib'
import type { AnnotDeleteInput, MarkupType } from '../shared/ipc'

const SUBTYPE: Record<MarkupType, string> = {
  highlight: 'Highlight',
  underline: 'Underline',
  strikeout: 'StrikeOut',
}
const RECT_TOLERANCE = 2

function normalized(rect: readonly number[]): readonly number[] {
  return [
    Math.min(rect[0] ?? 0, rect[2] ?? 0),
    Math.min(rect[1] ?? 0, rect[3] ?? 0),
    Math.max(rect[0] ?? 0, rect[2] ?? 0),
    Math.max(rect[1] ?? 0, rect[3] ?? 0),
  ]
}

function rectsClose(left: readonly number[], right: readonly number[]): boolean {
  const a = normalized(left)
  const b = normalized(right)
  return a.every((value, index) => Math.abs(value - (b[index] ?? 0)) <= RECT_TOLERANCE)
}

function annotationIdentity(
  document: PDFDocument,
  annotations: PDFArray,
  index: number,
): { objectNumber: number | null; subtype: string; rect: number[] } | null {
  const raw = annotations.get(index)
  const annotation = annotations.lookupMaybe(index, PDFDict)
  if (!annotation) return null
  const subtype = annotation.lookupMaybe(PDFName.of('Subtype'), PDFName)?.decodeText()
  const rect = annotation.lookupMaybe(PDFName.of('Rect'), PDFArray)
  if (!subtype || !rect) return null
  return {
    objectNumber: raw instanceof PDFRef ? raw.objectNumber : null,
    subtype,
    rect: rect.asArray().map((entry) => document.context.lookup(entry, PDFNumber).asNumber()),
  }
}

function matchingIndex(
  document: PDFDocument,
  annotations: PDFArray,
  deletion: AnnotDeleteInput,
): number | null {
  const expectedSubtype = SUBTYPE[deletion.subtype]
  const identities = Array.from({ length: annotations.size() }, (_, index) => ({
    index,
    identity: annotationIdentity(document, annotations, index),
  }))
  const guardedObject = identities.find(
    ({ identity }) =>
      identity?.objectNumber === deletion.objNum &&
      identity.subtype === expectedSubtype &&
      rectsClose(identity.rect, deletion.rect),
  )
  if (guardedObject) return guardedObject.index
  return (
    identities.find(
      ({ identity }) =>
        identity?.subtype === expectedSubtype && rectsClose(identity.rect, deletion.rect),
    )?.index ?? null
  )
}

export async function applyAnnotDeletes(
  bytes: Uint8Array,
  deletions: AnnotDeleteInput[],
): Promise<Uint8Array> {
  if (deletions.length === 0) return bytes
  const document = await PDFDocument.load(bytes, { updateMetadata: false })
  let removed = 0
  for (const deletion of deletions) {
    const page = document.getPages()[deletion.pageIndex]
    if (!page) continue
    const annotations = page.node.lookupMaybe(PDFName.of('Annots'), PDFArray)
    if (!annotations) continue
    const index = matchingIndex(document, annotations, deletion)
    if (index === null) continue
    annotations.remove(index)
    removed += 1
  }
  return removed > 0 ? document.save({ useObjectStreams: false }) : bytes
}
