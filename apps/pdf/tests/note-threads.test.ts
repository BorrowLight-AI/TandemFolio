import { describe, expect, it } from 'vitest'
import { buildNoteThreads } from '../src/renderer/note-threads'
import type { SavedNoteAnnot } from '../src/renderer/note-threads'

const saved = (overrides: Partial<SavedNoteAnnot>): SavedNoteAnnot => ({
  pageIndex: 0,
  objNum: 1,
  type: 'note',
  rect: [100, 682, 120, 700],
  color: [1, 0.78, 0.13],
  author: 'Alice',
  contents: 'Root',
  timeMs: 1_000,
  inReplyTo: null,
  ...overrides,
})

describe('PDF comment thread projection', () => {
  it('orders saved replies by creation time under their root', () => {
    const threads = buildNoteThreads(
      [
        saved({ objNum: 1 }),
        saved({ objNum: 3, contents: 'Late', inReplyTo: 1, timeMs: 3_000 }),
        saved({ objNum: 2, contents: 'Early', inReplyTo: 1, timeMs: 2_000 }),
      ],
      [],
    )

    expect(threads).toHaveLength(1)
    expect(threads[0]!.replies.map((reply) => reply.contents)).toEqual(['Early', 'Late'])
  })
})
