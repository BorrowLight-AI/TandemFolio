import { readFile } from 'node:fs/promises'
import { PDFDocument } from 'pdf-lib'
import JSZip from 'jszip'

export async function fixture(format: string): Promise<Buffer> {
  if (format === 'docx')
    return readFile('apps/docs/tests/pagination-corpus/docx/01-simple-english.docx')
  if (format === 'pptx')
    return readFile('packages/pptx-engine/tests/fixtures/01_standard_business.pptx')
  if (format === 'markdown')
    return Buffer.from('# Exact session\n\nRetained contents after restart.')
  if (format === 'pdf') {
    const pdf = await PDFDocument.create()
    pdf.addPage().drawText('Exact session persisted PDF')
    return Buffer.from(await pdf.save())
  }
  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>',
  )
  zip.file(
    '_rels/.rels',
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>',
  )
  zip.file(
    'xl/workbook.xml',
    '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="ExactSession" sheetId="1" r:id="rId1"/></sheets></workbook>',
  )
  zip.file(
    'xl/_rels/workbook.xml.rels',
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>',
  )
  zip.file(
    'xl/worksheets/sheet1.xml',
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Retained worksheet</t></is></c></row></sheetData></worksheet>',
  )
  return zip.generateAsync({ type: 'nodebuffer' })
}
