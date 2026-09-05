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

describe('toUniverStyle text rotation', () => {
  it('maps counterclockwise, clockwise-encoded, and stacked rotations', () => {
    expect(toUniverStyle({ ...base, textRotation: 90 }).tr).toEqual({ a: 90 })
    expect(toUniverStyle({ ...base, textRotation: 135 }).tr).toEqual({ a: -45 })
    expect(toUniverStyle({ ...base, textRotation: 255 }).tr).toEqual({ a: 0, v: 1 })
  })

  it('omits invalid or unrotated values', () => {
    expect(toUniverStyle(base)).not.toHaveProperty('tr')
    expect(toUniverStyle({ ...base, textRotation: 200 }).tr).toBeUndefined()
  })
})
