import { describe, expect, it } from 'vitest'

import {
  applyProtectedRanges,
  applySheetProtection,
  applyWorkbookProtection,
  SheetProtectionError,
} from '../src/gateway/xlsx-protection'

const BARE = '<worksheet><sheetData/><autoFilter ref="A1:C4"/></worksheet>'

describe('applySheetProtection', () => {
  it('inserts the element after sheetData with Excel defaults', () => {
    expect(applySheetProtection(BARE, true)).toBe(
      '<worksheet><sheetData/><sheetProtection sheet="1" objects="1" scenarios="1"/>'
      + '<autoFilter ref="A1:C4"/></worksheet>',
    )
  })

  it('removes an unpassworded element and is a no-op without one', () => {
    const protectedXml = applySheetProtection(BARE, true)
    expect(applySheetProtection(protectedXml, false)).toBe(BARE)
    expect(applySheetProtection(BARE, false)).toBe(BARE)
  })

  it('re-enables the sheet attribute on an existing element, keeping others', () => {
    const xml = '<worksheet><sheetData/>'
      + '<sheetProtection sheet="0" formatCells="0" insertRows="0"/></worksheet>'
    expect(applySheetProtection(xml, true)).toContain(
      '<sheetProtection sheet="1" formatCells="0" insertRows="0"/>',
    )
    const already = applySheetProtection(xml, true)
    expect(applySheetProtection(already, true)).toBe(already)
  })

  it('fails closed when unprotecting a password-protected sheet', () => {
    for (const attrs of [
      'sheet="1" password="83AF"',
      'sheet="1" algorithmName="SHA-512" hashValue="x" saltValue="y" spinCount="100000"',
    ]) {
      const xml = `<worksheet><sheetData/><sheetProtection ${attrs}/></worksheet>`
      expect(() => applySheetProtection(xml, false)).toThrow(SheetProtectionError)
    }
  })

  it('handles the paired-tag form', () => {
    const xml = '<worksheet><sheetData/><sheetProtection sheet="1"></sheetProtection></worksheet>'
    expect(applySheetProtection(xml, false)).toBe('<worksheet><sheetData/></worksheet>')
  })
})

describe('applyWorkbookProtection', () => {
  const workbook = '<workbook><bookViews/><sheets/></workbook>'

  it('adds and removes an unpassworded structure lock', () => {
    const locked = applyWorkbookProtection(workbook, true)
    expect(locked).toContain('<workbookProtection lockStructure="1"/>')
    expect(applyWorkbookProtection(locked, false)).toBe(workbook)
  })

  it('fails closed for password-protected structure', () => {
    const locked = '<workbook><workbookProtection lockStructure="1" workbookPassword="ABCD"/><sheets/></workbook>'
    expect(() => applyWorkbookProtection(locked, false)).toThrow(SheetProtectionError)
  })
})

describe('applyProtectedRanges', () => {
  it('writes escaped allow-edit ranges in schema order and removes them', () => {
    const written = applyProtectedRanges(BARE, [{ name: 'Sales & Tax', sqref: 'A2:B9 D2' }])
    expect(written).toContain(
      '<protectedRanges><protectedRange sqref="A2:B9 D2" name="Sales &amp; Tax"/></protectedRanges><autoFilter',
    )
    expect(applyProtectedRanges(written, [])).toBe(BARE)
  })

  it('fails closed instead of replacing password or identity permissions', () => {
    for (const item of [
      '<protectedRanges><protectedRange name="x" sqref="A1" password="AB"/></protectedRanges>',
      '<protectedRanges><protectedRange name="x" sqref="A1"><securityDescriptor/></protectedRange></protectedRanges>',
    ]) {
      expect(() => applyProtectedRanges(BARE.replace('<autoFilter', `${item}<autoFilter`), [])).toThrow(
        SheetProtectionError,
      )
    }
  })
})
