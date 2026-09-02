import { buildBlankDocx } from '@genoffice/docx-engine'
import { createBlankPptx } from '@genoffice/pptx-engine'
import JSZip from 'jszip'
import { PDFDocument, StandardFonts } from 'pdf-lib'

export type ReleaseFormat = 'docx' | 'markdown' | 'xlsx' | 'pptx' | 'pdf'
export type ReleaseFixtureSize = 'small' | 'medium' | 'large'

export interface ReleaseFixture {
  bytes: Uint8Array
  fileName: string
  mimeType: string
}

const FIXED_DATE = new Date('2000-01-01T00:00:00.000Z')

const scale = {
  small: {
    markdownLines: 30,
    docxParagraphs: 20,
    xlsxRows: 20,
    xlsxColumns: 8,
    slides: 2,
    pages: 2,
  },
  medium: {
    markdownLines: 2_000,
    docxParagraphs: 500,
    xlsxRows: 500,
    xlsxColumns: 20,
    slides: 20,
    pages: 30,
  },
  large: {
    markdownLines: 12_000,
    docxParagraphs: 2_000,
    xlsxRows: 2_000,
    xlsxColumns: 30,
    slides: 80,
    pages: 120,
  },
} as const

function xmlEscape(value: string): string {
  return value.replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;')
}

function normalizeZipDates(zip: JSZip): void {
  for (const entry of Object.values(zip.files)) entry.date = FIXED_DATE
}

async function markdownFixture(size: ReleaseFixtureSize): Promise<ReleaseFixture> {
  const lines = Array.from(
    { length: scale[size].markdownLines },
    (_, index) =>
      `## Section ${index + 1}\n\nDeterministic release fixture paragraph ${index + 1}.`,
  )
  return {
    bytes: new TextEncoder().encode(`# TandemFolio ${size} fixture\n\n${lines.join('\n\n')}`),
    fileName: `${size}.md`,
    mimeType: 'text/markdown',
  }
}

async function docxFixture(size: ReleaseFixtureSize): Promise<ReleaseFixture> {
  const zip = await JSZip.loadAsync(await buildBlankDocx())
  const path = 'word/document.xml'
  const documentXml = await zip.file(path)!.async('text')
  const paragraphs = Array.from(
    { length: scale[size].docxParagraphs },
    (_, index) =>
      `<w:p><w:r><w:t>${xmlEscape(`Deterministic ${size} paragraph ${index + 1}`)}</w:t></w:r></w:p>`,
  ).join('')
  zip.file(path, documentXml.replace(/<w:body>[\s\S]*?(<w:sectPr>)/, `<w:body>${paragraphs}$1`), {
    date: FIXED_DATE,
  })
  normalizeZipDates(zip)
  return {
    bytes: await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }),
    fileName: `${size}.docx`,
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  }
}

function xlsxColumnName(index: number): string {
  let value = index + 1
  let name = ''
  while (value > 0) {
    value -= 1
    name = String.fromCharCode(65 + (value % 26)) + name
    value = Math.floor(value / 26)
  }
  return name
}

async function xlsxFixture(size: ReleaseFixtureSize): Promise<ReleaseFixture> {
  const zip = new JSZip()
  const options = { date: FIXED_DATE }
  zip.file(
    '[Content_Types].xml',
    '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>',
    options,
  )
  zip.file(
    '_rels/.rels',
    '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>',
    options,
  )
  zip.file(
    'xl/workbook.xml',
    '<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>',
    options,
  )
  zip.file(
    'xl/_rels/workbook.xml.rels',
    '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>',
    options,
  )
  zip.file(
    'xl/styles.xml',
    '<?xml version="1.0"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>',
    options,
  )
  const rows = Array.from({ length: scale[size].xlsxRows }, (_, rowIndex) => {
    const cells = Array.from({ length: scale[size].xlsxColumns }, (_, columnIndex) => {
      const address = `${xlsxColumnName(columnIndex)}${rowIndex + 1}`
      return `<c r="${address}" t="inlineStr"><is><t>${size}-${rowIndex + 1}-${columnIndex + 1}</t></is></c>`
    }).join('')
    return `<row r="${rowIndex + 1}">${cells}</row>`
  }).join('')
  zip.file(
    'xl/worksheets/sheet1.xml',
    `<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>${rows}</sheetData></worksheet>`,
    options,
  )
  normalizeZipDates(zip)
  return {
    bytes: await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }),
    fileName: `${size}.xlsx`,
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  }
}

async function pptxFixture(size: ReleaseFixtureSize): Promise<ReleaseFixture> {
  const zip = await JSZip.loadAsync(await createBlankPptx())
  const count = scale[size].slides
  const slideXml = await zip.file('ppt/slides/slide1.xml')!.async('text')
  const slideRels = await zip.file('ppt/slides/_rels/slide1.xml.rels')!.async('text')
  const contentTypesPath = '[Content_Types].xml'
  const presentationPath = 'ppt/presentation.xml'
  const relsPath = 'ppt/_rels/presentation.xml.rels'
  let contentTypes = await zip.file(contentTypesPath)!.async('text')
  let presentation = await zip.file(presentationPath)!.async('text')
  let relationships = await zip.file(relsPath)!.async('text')
  const ids = Array.from(
    { length: count },
    (_, index) => `<p:sldId id="${256 + index}" r:id="rId${index + 2}"/>`,
  ).join('')
  presentation = presentation.replace(
    /<p:sldIdLst>[\s\S]*?<\/p:sldIdLst>/,
    `<p:sldIdLst>${ids}</p:sldIdLst>`,
  )
  for (let index = 2; index <= count; index += 1) {
    contentTypes = contentTypes.replace(
      '</Types>',
      `<Override PartName="/ppt/slides/slide${index}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>`,
    )
    relationships = relationships.replace(
      '</Relationships>',
      `<Relationship Id="rId${index + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${index}.xml"/></Relationships>`,
    )
    zip.file('ppt/slides/slide' + index + '.xml', slideXml, { date: FIXED_DATE })
    zip.file('ppt/slides/_rels/slide' + index + '.xml.rels', slideRels, { date: FIXED_DATE })
  }
  zip.file(contentTypesPath, contentTypes, { date: FIXED_DATE })
  zip.file(presentationPath, presentation, { date: FIXED_DATE })
  zip.file(relsPath, relationships, { date: FIXED_DATE })
  normalizeZipDates(zip)
  return {
    bytes: await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }),
    fileName: `${size}.pptx`,
    mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  }
}

async function pdfFixture(size: ReleaseFixtureSize): Promise<ReleaseFixture> {
  const document = await PDFDocument.create()
  document.setCreationDate(FIXED_DATE)
  document.setModificationDate(FIXED_DATE)
  document.setProducer('TandemFolio release fixture')
  document.setCreator('TandemFolio release fixture')
  const font = await document.embedFont(StandardFonts.Helvetica)
  for (let index = 0; index < scale[size].pages; index += 1) {
    const page = document.addPage([612, 792])
    page.drawText(`Deterministic ${size} page ${index + 1}`, { x: 54, y: 720, size: 16, font })
    for (let line = 0; line < 20; line += 1) {
      page.drawText(`Release baseline content ${index + 1}.${line + 1}`, {
        x: 54,
        y: 680 - line * 28,
        size: 10,
        font,
      })
    }
  }
  return {
    bytes: await document.save({ useObjectStreams: false }),
    fileName: `${size}.pdf`,
    mimeType: 'application/pdf',
  }
}

export async function createReleaseFixture(
  format: ReleaseFormat,
  size: ReleaseFixtureSize,
): Promise<ReleaseFixture> {
  if (format === 'markdown') return markdownFixture(size)
  if (format === 'docx') return docxFixture(size)
  if (format === 'xlsx') return xlsxFixture(size)
  if (format === 'pptx') return pptxFixture(size)
  return pdfFixture(size)
}
