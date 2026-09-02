import { describe, expect, it } from 'vitest'
import { parseDocx } from '@genoffice/docx-engine'
import {
  compareDocxBytes,
  compareParagraphs,
  summarize,
} from '../src/renderer/editor/compare'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'

describe('compareParagraphs', () => {
  it('reports identical documents as all same', () => {
    const entries = compareParagraphs(['a', 'b'], ['a', 'b'])
    expect(entries.every((e) => e.kind === 'same')).toBe(true)
    expect(summarize(entries)).toEqual({ added: 0, removed: 0, changed: 0 })
  })

  it('detects an added paragraph', () => {
    const entries = compareParagraphs(['a', 'c'], ['a', 'b', 'c'])
    expect(entries.map((e) => e.kind)).toEqual(['same', 'added', 'same'])
    expect(entries[1].right).toBe('b')
  })

  it('detects a removed paragraph', () => {
    const entries = compareParagraphs(['a', 'b', 'c'], ['a', 'c'])
    expect(entries.map((e) => e.kind)).toEqual(['same', 'removed', 'same'])
    expect(entries[1].left).toBe('b')
  })

  it('merges adjacent remove+add into changed', () => {
    const entries = compareParagraphs(['title', 'old content', 'ending'], ['title', 'new content', 'ending'])
    expect(entries.map((e) => e.kind)).toEqual(['same', 'changed', 'same'])
    expect(entries[1]).toMatchObject({ left: 'old content', right: 'new content' })
  })

  it('handles empty documents', () => {
    expect(compareParagraphs([], [])).toEqual([])
    expect(compareParagraphs([], ['x'])[0].kind).toBe('added')
    expect(compareParagraphs(['x'], [])[0].kind).toBe('removed')
  })
})

describe('staged DOCX comparison', () => {
  it('parses bounded staged bytes and returns deterministic panel state and counts', async () => {
    const current = await parseDocx(
      await buildDocx({
        bodyXml:
          '<w:p><w:r><w:t>Title</w:t></w:r></w:p><w:p><w:r><w:t>Old</w:t></w:r></w:p>',
      }),
    )
    const other = await buildDocx({
      bodyXml:
        '<w:p><w:r><w:t>Title</w:t></w:r></w:p><w:p><w:r><w:t>New</w:t></w:r></w:p>',
    })
    const result = await compareDocxBytes(current.blocks, {
      name: 'other.docx',
      data: other.buffer.slice(other.byteOffset, other.byteOffset + other.byteLength),
    })
    expect(result).toMatchObject({
      ok: true,
      otherName: 'other.docx',
      added: 0,
      removed: 0,
      changed: 1,
      identical: false,
    })
    if (!result.ok) throw new Error(result.message)
    expect(result.entries).toMatchObject([
      { kind: 'same', left: 'Title', right: 'Title' },
      { kind: 'changed', left: 'Old', right: 'New' },
    ])
  })
})
