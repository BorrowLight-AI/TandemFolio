import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'

import { openBrowserWorkbook } from '../src/host/browser-workbook'
import { BrowserWorkbookDesktopApi } from '../src/renderer/browser-desktop-api'

const PNG = Uint8Array.from([
  137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 13, 73, 72, 68, 82, 0, 0, 0, 1, 0, 0, 0,
  1, 8, 6, 0, 0, 0, 31, 21, 196, 137,
])

async function oleFixture(): Promise<Uint8Array> {
  const zip = new JSZip()
  zip.file('[Content_Types].xml', '<Types/>')
  zip.file(
    'xl/workbook.xml',
    '<workbook><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>',
  )
  zip.file(
    'xl/_rels/workbook.xml.rels',
    '<Relationships><Relationship Id="rId1" Target="worksheets/sheet1.xml"/></Relationships>',
  )
  zip.file(
    'xl/worksheets/sheet1.xml',
    '<worksheet><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>' +
      '<legacyDrawing r:id="rId2"/><oleObjects><oleObject progId="Word.Document.12" shapeId="1025" r:id="rId3">' +
      '<objectPr r:id="rId4"><anchor><from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></from>' +
      '<to><xdr:col>4</xdr:col><xdr:colOff>384810</xdr:colOff><xdr:row>3</xdr:row><xdr:rowOff>156210</xdr:rowOff></to>' +
      '</anchor></objectPr></oleObject></oleObjects></worksheet>',
  )
  zip.file(
    'xl/worksheets/_rels/sheet1.xml.rels',
    '<Relationships>' +
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" Target="../drawings/vmlDrawing1.vml"/>' +
      '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject" Target="../embeddings/oleObject1.bin"/>' +
      '<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>' +
      '</Relationships>',
  )
  zip.file(
    'xl/drawings/vmlDrawing1.vml',
    '<xml><v:shape id="_x0000_s1025" filled="t" fillcolor="window [65]" stroked="t" strokecolor="windowText [64]">' +
      '<v:imagedata o:relid="rId1"/><x:ClientData ObjectType="Pict"><x:Anchor>1, 0, 1, 0, 4, 40, 3, 16</x:Anchor></x:ClientData></v:shape></xml>',
  )
  zip.file(
    'xl/drawings/_rels/vmlDrawing1.vml.rels',
    '<Relationships><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/></Relationships>',
  )
  zip.file('xl/embeddings/oleObject1.bin', Uint8Array.from([1, 2, 3, 4]))
  zip.file('xl/media/image1.png', PNG)
  return zip.generateAsync({ type: 'uint8array' })
}

describe('browser OLE objects', () => {
  it('exposes preview bytes through the browser session media route', async () => {
    const api = new BrowserWorkbookDesktopApi(async () => null)
    const file = await api.openBuffer((await oleFixture()).buffer as ArrayBuffer, 'ole.xlsx')
    const visual = file.visuals[0]
    expect(visual).toMatchObject({ kind: 'ole', progId: 'Word.Document.12' })
    const media = await api.readWorkbookMedia({ sessionId: file.sessionId, visualId: visual!.id })
    expect(media.mediaType).toBe('image/png')
    expect(media.base64.length).toBeGreaterThan(0)
  })

  it('hydrates a read-only object and preserves its opaque package parts', async () => {
    const source = await oleFixture()
    const workbook = await openBrowserWorkbook(source, 'ole.xlsx')
    expect(workbook.visuals).toEqual([
      expect.objectContaining({
        kind: 'ole',
        progId: 'Word.Document.12',
        mediaPath: 'xl/media/image1.png',
        mediaType: 'image/png',
        lineColor: '#000000',
        fillColor: '#FFFFFF',
        anchor: expect.objectContaining({ fromColumn: 1, toColumn: 4, fromRow: 1, toRow: 3 }),
      }),
    ])
    expect(workbook.visuals[0]?.drawingPath).toBeUndefined()

    workbook.setCellValue('Data', 'A1', 2)
    const saved = await JSZip.loadAsync(await workbook.save())
    const original = await JSZip.loadAsync(source)
    for (const path of [
      'xl/embeddings/oleObject1.bin',
      'xl/drawings/vmlDrawing1.vml',
      'xl/media/image1.png',
    ]) {
      expect(await saved.file(path)?.async('uint8array')).toEqual(
        await original.file(path)?.async('uint8array'),
      )
    }
  })

  it('moves worksheet and VML anchors with structural row and column inserts', async () => {
    const workbook = await openBrowserWorkbook(await oleFixture(), 'ole.xlsx')
    workbook.applyStructuralOperation('Data', { kind: 'insert-rows', index: 0, count: 2 })
    workbook.applyStructuralOperation('Data', { kind: 'insert-cols', index: 2, count: 1 })
    const saved = await JSZip.loadAsync(await workbook.save())
    const sheet = await saved.file('xl/worksheets/sheet1.xml')?.async('text')
    const vml = await saved.file('xl/drawings/vmlDrawing1.vml')?.async('text')
    expect(sheet).toContain('<xdr:col>5</xdr:col>')
    expect(sheet).toContain('<xdr:row>5</xdr:row>')
    expect(vml).toContain('<x:Anchor>1, 0, 3, 0, 5, 40, 5, 16</x:Anchor>')
  })
})
