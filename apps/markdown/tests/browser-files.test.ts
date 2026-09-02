import { describe, expect, it } from 'vitest'

import { readLoadedMarkdownAsset, type LoadedMarkdown } from '../src/renderer/host/browser-files'

const TINY_PNG = Uint8Array.from(
  atob(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M/wHwAF/gL+Xw4bWQAAAABJRU5ErkJggg==',
  ),
  (character) => character.charCodeAt(0),
)

describe('browser Markdown asset reader', () => {
  it('resolves an authored relative path against the selected Markdown directory', async () => {
    const file = new File([TINY_PNG], 'pixel.png', { type: 'image/png' })
    const loaded: LoadedMarkdown = {
      fileName: 'notes.md',
      text: '![pixel](../assets/pixel.png)',
      assetBasePath: 'docs',
      assetFiles: new Map([['assets/pixel.png', file]]),
    }

    await expect(readLoadedMarkdownAsset(loaded, '../assets/pixel.png')).resolves.toEqual({
      mime: 'image/png',
      data: TINY_PNG.buffer,
    })
  })

  it('rejects paths outside the selected directory and remote URLs', async () => {
    const loaded: LoadedMarkdown = {
      fileName: 'notes.md',
      text: '',
      assetBasePath: '',
      assetFiles: new Map(),
    }
    await expect(readLoadedMarkdownAsset(loaded, '../secret.png')).resolves.toBeNull()
    await expect(readLoadedMarkdownAsset(loaded, 'https://example.com/a.png')).resolves.toBeNull()
  })
})
