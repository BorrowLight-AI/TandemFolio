import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'

import { applyCellEditsToXlsx } from '../src/gateway/xlsx-gateway'
import { withFutureFunctionMarkers } from '../src/gateway/future-functions'
import { buildEditFixture } from './fixture-builder'

describe('withFutureFunctionMarkers', () => {
  it('prefixes future functions for storage', () => {
    expect(withFutureFunctionMarkers('MINIFS(C7:C10,C7:C10,">0")')).toBe(
      '_xlfn.MINIFS(C7:C10,C7:C10,">0")',
    )
    expect(withFutureFunctionMarkers('IFS(A1>0,STDEV.P(B:B),TRUE,0)')).toBe(
      '_xlfn.IFS(A1>0,_xlfn.STDEV.P(B:B),TRUE,0)',
    )
  })

  it('marks worksheet-scope dynamic-array functions with _xlws', () => {
    expect(withFutureFunctionMarkers('SORT(FILTER(A:A,B:B>0))')).toBe(
      '_xlfn._xlws.SORT(_xlfn._xlws.FILTER(A:A,B:B>0))',
    )
  })

  it('canonicalizes lowercase and mixed-case calls', () => {
    expect(withFutureFunctionMarkers('minifs(A:A,A:A,">0")')).toBe('_xlfn.MINIFS(A:A,A:A,">0")')
    expect(withFutureFunctionMarkers('XLookup(1,A:A,B:B)')).toBe('_xlfn.XLOOKUP(1,A:A,B:B)')
    expect(withFutureFunctionMarkers('sum(A1:A3)')).toBe('sum(A1:A3)')
  })

  it('restores the implicit-intersection and spill storage forms', () => {
    expect(withFutureFunctionMarkers('SINGLE(A:A)+SUM(ANCHORARRAY(C3))')).toBe(
      '_xlfn.SINGLE(A:A)+SUM(_xlfn.ANCHORARRAY(C3))',
    )
  })

  it('leaves classic functions, strings, and existing markers alone', () => {
    expect(withFutureFunctionMarkers('SUM(A1:A3)+IF(B1,1,0)')).toBe('SUM(A1:A3)+IF(B1,1,0)')
    expect(withFutureFunctionMarkers('CONCATENATE("MINIFS(",A1,")")')).toBe(
      'CONCATENATE("MINIFS(",A1,")")',
    )
    expect(withFutureFunctionMarkers('_xlfn.MINIFS(A:A,A:A,">0")')).toBe(
      '_xlfn.MINIFS(A:A,A:A,">0")',
    )
  })
})

describe('future functions saved through the workbook gateway', () => {
  it('writes Excel storage markers while preserving the editor formula shape', async () => {
    const mutation = await applyCellEditsToXlsx(await buildEditFixture(), [
      {
        sheetName: 'Data',
        row: 1,
        column: 0,
        writeValue: true,
        cell: { value: null, formula: '=FILTER(A1:A3,A1:A3>0)' },
      },
    ])
    const zip = await JSZip.loadAsync(mutation.buffer)
    const sheet = await zip.file('xl/worksheets/sheet1.xml')?.async('text')
    expect(sheet).toContain('<f>_xlfn._xlws.FILTER(A1:A3,A1:A3&gt;0)</f>')
  })
})
