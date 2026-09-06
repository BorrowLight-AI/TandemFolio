import { describe, expect, it } from 'vitest'
import {
  COLOR_PRESETS,
  hexTo255,
  hsvToRgb,
  rgb255ToHex,
  rgbToHsv,
} from '../src/renderer/color-presets'

describe('PDF native color picker presets', () => {
  it('round-trips every preset through the HSV picker model', () => {
    for (const color of COLOR_PRESETS) {
      expect(rgb255ToHex(hsvToRgb(...rgbToHsv(...hexTo255(color))))).toBe(color.toLowerCase())
    }
  })
})
