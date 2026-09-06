import { describe, expect, it } from 'vitest'
import { mapDocFont } from '../src/renderer/doc-font'

describe('PDF document font mapping', () => {
  it('removes subset prefixes and maps CJK serif faces to local metric-compatible families', () => {
    expect(mapDocFont('ABCDEF+SimSun')).toEqual({
      css: expect.stringContaining('SimSun'),
    })
  })
})
