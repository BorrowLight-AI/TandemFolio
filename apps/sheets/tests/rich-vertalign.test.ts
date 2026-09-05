import { BaselineOffset } from '@univerjs/core'
import { describe, expect, it } from 'vitest'

import { toRichTextDocument } from '../src/renderer/univer-sync'
import type { WorkbookRichRun } from '../src/shared/desktop-api'

const run = (text: string, extra: Partial<WorkbookRichRun> = {}): WorkbookRichRun => ({
  text,
  bold: false,
  italic: false,
  underline: false,
  strikethrough: false,
  ...extra,
})

describe('rich-run vertical alignment', () => {
  it('maps subscript and superscript runs to Univer baseline offsets', () => {
    const document = toRichTextDocument(
      'Cr2O3 g/cm3',
      [
        run('Cr'),
        run('2', { vertAlign: 'subscript', size: 14 }),
        run('O'),
        run('3', { vertAlign: 'subscript', size: 14 }),
        run(' g/cm'),
        run('3', { vertAlign: 'superscript' }),
      ],
      { fs: 11 },
    )
    const styles = document?.body?.textRuns?.map((entry) => entry.ts)
    expect(styles?.[1]?.va).toBe(BaselineOffset.SUBSCRIPT)
    expect(styles?.[1]?.fs).toBe(14)
    expect(styles?.[3]?.va).toBe(BaselineOffset.SUBSCRIPT)
    expect(styles?.[5]?.va).toBe(BaselineOffset.SUPERSCRIPT)
    expect(styles?.[0]?.va).toBeUndefined()
  })

  it('keeps a vertical-align-only run distinct from the base style', () => {
    const document = toRichTextDocument(
      'x2',
      [run('x'), run('2', { vertAlign: 'subscript' })],
      { fs: 11 },
    )
    expect(document?.body?.textRuns?.[1]?.ts?.va).toBe(BaselineOffset.SUBSCRIPT)
    expect(document?.body?.textRuns?.[1]?.ts?.fs).toBe(11)
  })
})
