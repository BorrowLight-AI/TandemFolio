import { describe, expect, it } from 'vitest'

import { toUniverStyle } from '../src/renderer/univer-sync'
import type { WorkbookCellStyle } from '../src/shared/desktop-api'

const base: WorkbookCellStyle = {
  bold: false,
  italic: false,
  underline: false,
  strikethrough: false,
  wrapText: false,
  diagonalUp: false,
  diagonalDown: false,
}

describe('toUniverStyle fill isolation', () => {
  it('blocks a column-style fill from bleeding into a fill-less cell xf', () => {
    expect(toUniverStyle(base).bg).toEqual({ rgb: '' })
  })

  it('keeps explicit fills', () => {
    expect(toUniverStyle({ ...base, fillColor: '#FF0000' }).bg).toEqual({ rgb: '#FF0000' })
  })
})
