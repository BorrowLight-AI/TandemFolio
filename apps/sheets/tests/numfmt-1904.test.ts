import { describe, expect, it } from 'vitest'

import { fixFormattedValue, isCalendarDatePattern } from '../src/renderer/numfmt-fix'
import { readPivotSourceGrid } from '../src/renderer/workbook-ops'

describe('isCalendarDatePattern', () => {
  it('matches calendar date and datetime patterns', () => {
    expect(isCalendarDatePattern('m/d/yyyy')).toBe(true)
    expect(isCalendarDatePattern('yyyy-mm-dd h:mm')).toBe(true)
    expect(isCalendarDatePattern('d-mmm-yy')).toBe(true)
  })

  it('rejects time-only and elapsed patterns whose magnitude must not shift', () => {
    expect(isCalendarDatePattern('[h]:mm')).toBe(false)
    expect(isCalendarDatePattern('[mm]:ss')).toBe(false)
    expect(isCalendarDatePattern('h:mm:ss')).toBe(false)
    expect(isCalendarDatePattern('[h]:mm:ss;@')).toBe(false)
  })

  it('rejects plain number and text patterns', () => {
    expect(isCalendarDatePattern('#,##0.00')).toBe(false)
    expect(isCalendarDatePattern('@')).toBe(false)
    expect(isCalendarDatePattern('General')).toBe(false)
  })
})

describe('1904 workbook date formatting', () => {
  it('shifts static calendar serials onto the 1904 epoch', () => {
    expect(fixFormattedValue('yyyy-mm-dd', 1, '1900-01-01', true)).toBe('1904-01-02')
    expect(fixFormattedValue('yyyy-mm-dd h:mm', 1.5, '1900-01-01 12:00', true)).toBe(
      '1904-01-02\u00a012:00',
    )
  })

  it('does not shift time-only formats', () => {
    expect(fixFormattedValue('h:mm:ss', 0.5, '12:00:00', true)).toBeNull()
    expect(fixFormattedValue('[h]:mm', 1.5, '36:00', true)).toBeNull()
  })

  it('feeds pivot grouping width-independent 1904 date labels', () => {
    const range = {
      getRawValues: () => [
        ['Date', 'Formula date', 'Hours'],
        [1, 1, 1.5],
      ],
      getNumberFormats: () => [
        ['General', 'General', 'General'],
        ['yyyy-mm-dd', 'yyyy-mm-dd', '[h]:mm'],
      ],
      getFormulas: () => [
        ['', '', ''],
        ['', '=DATE(1900,1,1)', ''],
      ],
    }

    expect(readPivotSourceGrid(range, true)[1]).toEqual([
      '1904-01-02',
      '1900-01-01',
      '36:00',
    ])
  })
})
