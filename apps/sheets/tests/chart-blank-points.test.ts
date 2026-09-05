import { describe, expect, it } from 'vitest'

import { lineSegments } from '../src/renderer/WorkbookVisuals'

describe('lineSegments', () => {
  it('plots one run when there are no blanks or the mode keeps zeros', () => {
    expect(lineSegments(3, undefined, 'gap')).toEqual([[0, 1, 2]])
    expect(lineSegments(3, [1], undefined)).toEqual([[0, 1, 2]])
    expect(lineSegments(3, [1], 'zero')).toEqual([[0, 1, 2]])
    expect(lineSegments(0, undefined, undefined)).toEqual([])
  })

  it('breaks the line at blank cells for dispBlanksAs=gap', () => {
    expect(lineSegments(6, [2, 3], 'gap')).toEqual([
      [0, 1],
      [4, 5],
    ])
    expect(lineSegments(4, [0, 3], 'gap')).toEqual([[1, 2]])
    expect(lineSegments(2, [0, 1], 'gap')).toEqual([])
  })

  it('bridges blank cells for dispBlanksAs=span', () => {
    expect(lineSegments(6, [2, 3], 'span')).toEqual([[0, 1, 4, 5]])
    expect(lineSegments(2, [0, 1], 'span')).toEqual([])
  })
})
