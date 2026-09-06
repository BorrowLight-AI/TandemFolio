import { describe, expect, it } from 'vitest'

import { layoutText } from '../src/renderer/components/FindReplaceDialog'
import type { RenderTextLayout } from '@genoffice/pptx-render'

describe('PPTX find text reconstruction', () => {
  it('preserves the exact whitespace consumed at a wrapped line boundary', () => {
    const text = {
      lines: [
        {
          paraStart: true,
          trailingSpace: true,
          trailingText: '   ',
          runs: [{ text: 'alpha' }],
        },
        {
          paraStart: false,
          runs: [{ text: 'beta' }],
        },
      ],
    } as RenderTextLayout

    expect(layoutText(text)).toBe('alpha   beta')
  })
})
