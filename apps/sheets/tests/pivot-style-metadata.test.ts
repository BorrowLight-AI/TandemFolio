import { describe, expect, it } from 'vitest'

import { readPivotStyleMetadata, readTableStyleMetadata } from '../src/host/xlsx-pivot-style'

const OFFICE_2007_THEME = [
  '#FFFFFF',
  '#000000',
  '#EEECE1',
  '#1F497D',
  '#4F81BD',
  '#C0504D',
  '#9BBB59',
  '#8064A2',
  '#4BACC6',
  '#F79646',
]

describe('readPivotStyleMetadata', () => {
  it('resolves native bands and row kinds from pivot OOXML', () => {
    const xml = `
      <pivotTableDefinition rowGrandTotals="0">
        <location ref="$A$1:$D$7" firstDataRow="2" firstDataCol="1"/>
        <pivotFields count="2"><pivotField/><pivotField outline="1"/></pivotFields>
        <rowFields count="2"><field x="0"/><field x="1"/></rowFields>
        <rowItems count="5">
          <i r="0"/><i r="1"/><i r="1"/><i r="0" t="sum"/><i t="grand"/>
        </rowItems>
        <pivotTableStyleInfo name="PivotStyleDark23" showRowStripes="1" showColStripes="1"/>
      </pivotTableDefinition>`

    expect(readPivotStyleMetadata(xml, OFFICE_2007_THEME)).toMatchObject({
      styled: true,
      firstDataRow: 2,
      firstDataCol: 1,
      rowGrandTotals: false,
      rowKinds: 'sddtg',
      headerFill: '#376092',
      wholeTableFill: '#4F81BD',
      wholeTableFontColor: '#DCE6F2',
      secondRowStripeFill: '#95B3D7',
      secondColumnStripeFill: '#95B3D7',
      firstColumnFill: '#376092',
      firstHeaderCellFontColor: '#FFFFFF',
    })
  })

  it('does not emit stripe fills when the style disables them', () => {
    const xml = `
      <pivotTableDefinition>
        <location ref="A1:B3"/>
        <pivotTableStyleInfo name="PivotStyleMedium23" showRowStripes="0" showColStripes="0"/>
      </pivotTableDefinition>`
    const metadata = readPivotStyleMetadata(xml, OFFICE_2007_THEME)
    expect(metadata.secondRowStripeFill).toBeUndefined()
    expect(metadata.secondColumnStripeFill).toBeUndefined()
    expect(metadata.wholeTableFill).toBe('#DCE6F2')
  })
})

describe('readTableStyleMetadata', () => {
  it('resolves a themed Medium table palette', () => {
    expect(readTableStyleMetadata('TableStyleMedium9', OFFICE_2007_THEME)).toMatchObject({
      headerFill: '#4F81BD',
      headerFontColor: '#FFFFFF',
      stripeFill: '#B9CDE5',
      secondRowStripeFill: '#DCE6F2',
      wholeTableFill: '#DCE6F2',
      totalRowFill: '#4F81BD',
      totalRowFontColor: '#FFFFFF',
    })
  })

  it('keeps style None unpainted', () => {
    expect(readTableStyleMetadata(undefined, OFFICE_2007_THEME)).toEqual({})
  })
})
