import { describe, expect, it } from 'vitest'

import { oleCaption, oleFrameStyle, oleRenderKind } from '../src/renderer/ole-visual'
import { isEditableFileVisual } from '../src/renderer/WorkbookVisuals'
import type { WorkbookVisualObject } from '../src/shared/desktop-api'

const anchor = {
  fromRow: 1,
  fromColumn: 1,
  toRow: 3,
  toColumn: 4,
  fromRowOffset: 0,
  fromColumnOffset: 0,
  toRowOffset: 156_210,
  toColumnOffset: 384_810,
  explicitTo: true,
}

describe('OLE visual helpers', () => {
  it('keeps embedded objects outside the editable drawing pipeline', () => {
    const ole = {
      id: 'ole-1',
      sheetId: 'sheet-1',
      kind: 'ole',
      anchor,
      progId: 'Package',
    } satisfies WorkbookVisualObject
    expect(isEditableFileVisual(ole)).toBe(false)
    expect(
      isEditableFileVisual({ ...ole, drawingPath: 'xl/drawings/drawing1.xml', drawingIndex: 0 }),
    ).toBe(false)
  })

  it('selects cached previews and falls back after a failed or absent preview', () => {
    expect(oleRenderKind({ mediaPath: 'xl/media/image1.emf' }, false)).toBe('preview')
    expect(oleRenderKind({}, false)).toBe('placeholder')
    expect(oleRenderKind({ mediaPath: 'xl/media/image1.emf' }, true)).toBe('placeholder')
  })

  it('maps known ProgIDs and preserves useful unknown names', () => {
    expect(oleCaption('Word.Document.12')).toBe('Microsoft Word Document')
    expect(oleCaption('Excel.Sheet.12')).toBe('Microsoft Excel Worksheet')
    expect(oleCaption('AcroExch.Document.11')).toBe('Adobe Acrobat Document')
    expect(oleCaption('Vendor.Widget.3')).toBe('Vendor.Widget')
    expect(oleCaption(undefined)).toBe('Embedded Object')
  })

  it('uses the legacy VML frame colors when present', () => {
    expect(oleFrameStyle({ lineColor: '#000000', fillColor: '#FFFFFF' })).toEqual({
      border: '1px solid #000000',
      background: '#FFFFFF',
    })
    expect(oleFrameStyle({ lineColor: 'none', fillColor: 'none' })).toEqual({})
  })
})
