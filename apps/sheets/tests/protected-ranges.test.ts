import { describe, expect, it } from 'vitest'

import { mapProtectedRanges } from '../src/renderer/protected-ranges'

const range = (name: string, sqref: string) => ({ name, sqref, hasPassword: false })

describe('mapProtectedRanges', () => {
  it('shifts areas through structural edits', () => {
    expect(
      mapProtectedRanges([range('Data', 'B3:D6'), range('Cell', 'B2')], [
        { kind: 'insert-rows', index: 1, count: 2 },
        { kind: 'remove-cols', index: 0, count: 1 },
      ]),
    ).toEqual([range('Data', 'A5:C8'), range('Cell', 'A4')])
  })

  it('shrinks partial deletion and drops fully deleted areas', () => {
    expect(
      mapProtectedRanges(
        [range('Shrinks', 'A2:A5'), range('Gone', 'A3:A4'), range('Multi', 'A1 A3:A4')],
        [{ kind: 'remove-rows', index: 2, count: 2 }],
      ),
    ).toEqual([range('Shrinks', 'A2:A3'), range('Multi', 'A1')])
  })

  it('splits a partially moved range into exact runs', () => {
    expect(
      mapProtectedRanges([range('Data', 'A2:A4')], [
        { kind: 'move-rows', index: 1, count: 2, before: 6 },
      ]),
    ).toEqual([range('Data', 'A2 A5:A6')])
  })

  it('keeps unparseable references fail-closed', () => {
    expect(
      mapProtectedRanges([range('Odd', 'NotARef')], [{ kind: 'insert-rows', index: 0, count: 1 }]),
    ).toEqual([range('Odd', 'NotARef')])
  })
})
