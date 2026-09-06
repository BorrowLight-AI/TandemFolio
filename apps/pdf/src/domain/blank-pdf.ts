import { PDFDocument } from 'pdf-lib'

const A4_SIZE: [number, number] = [595.28, 841.89]

/** Create the minimal browser-owned document used by the PDF new-document command. */
export async function blankPdfBytes(): Promise<Uint8Array> {
  const document = await PDFDocument.create()
  document.addPage(A4_SIZE)
  return document.save({ useObjectStreams: false })
}
