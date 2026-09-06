import { describe, expect, it } from 'vitest'

import {
  extractMarkdownImageSources,
  prepareMarkdownAssetSave,
} from '../src/renderer/host/asset-save-plan'

const PNG_DATA = 'data:image/png;base64,iVBORw0KGgo='

describe('Markdown asset save plan', () => {
  it('extracts rendered Markdown and raw HTML image sources while ignoring code', () => {
    const markdown = [
      '![chart](images/chart.png)',
      `<IMG alt="raw" SRC='images&#47;raw.png'>`,
      '`![inline](ignored.png)`',
      '```md',
      '![fenced](ignored-too.png)',
      '```',
    ].join('\n')

    expect(extractMarkdownImageSources(markdown)).toEqual(['images/chart.png', 'images/raw.png'])
  })

  it('rewrites repeated data images to one deterministic companion file', async () => {
    const prepared = await prepareMarkdownAssetSave(
      `![first](${PNG_DATA})\n![second](${PNG_DATA})`,
      [],
      { copyRelative: false },
    )

    expect(prepared.companionFiles).toHaveLength(1)
    expect(prepared.companionFiles[0]?.relativePath).toMatch(/^assets\/image-[a-f0-9]{12}\.png$/)
    expect(prepared.text).not.toContain('data:image')
    expect(prepared.text.match(/assets\/image-/g)).toHaveLength(2)
  })

  it('copies safe relative Markdown and HTML sources for Save As but rejects traversal', async () => {
    const chart = Uint8Array.from([1, 2, 3]).buffer
    const raw = Uint8Array.from([4, 5, 6]).buffer
    const text = [
      '![chart](images/chart.png)',
      `<img src='images&#47;raw.png'>`,
      '![unsafe](../outside.png)',
    ].join('\n')
    const prepared = await prepareMarkdownAssetSave(
      text,
      [
        { source: 'images/chart.png', mime: 'image/png', data: chart },
        { source: 'images/raw.png', mime: 'image/png', data: raw },
      ],
      { copyRelative: true },
    )

    expect(prepared.companionFiles).toHaveLength(2)
    expect(prepared.text).toContain('![chart](assets/chart-')
    expect(prepared.text).toContain(`<img src='assets/raw-`)
    expect(prepared.text).toContain('![unsafe](../outside.png)')
  })

  it('keeps an existing content-addressed asset path stable across repeated Save As', async () => {
    const data = Uint8Array.from([1, 2, 3]).buffer
    const source = 'assets/chart-039058c6f2c0.png'
    const prepared = await prepareMarkdownAssetSave(
      `![chart](${source})`,
      [{ source, mime: 'image/png', data }],
      { copyRelative: true },
    )

    expect(prepared.text).toBe(`![chart](${source})`)
    expect(prepared.companionFiles).toEqual([{ relativePath: source, data }])
    expect(prepared.rewrites.size).toBe(0)
  })
})
