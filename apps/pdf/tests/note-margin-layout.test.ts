import { describe, expect, it } from 'vitest'
import { layoutMarginCards } from '../src/renderer/note-margin-layout'

describe('PDF comment margin layout', () => {
  it('pushes overlapping cards apart while keeping their page order', () => {
    expect(
      layoutMarginCards(
        [
          { key: 'a', pinY: 54 },
          { key: 'b', pinY: 74 },
        ],
        () => 80,
        null,
        12,
      ),
    ).toEqual(
      new Map([
        ['a', 40],
        ['b', 132],
      ]),
    )
  })
})
