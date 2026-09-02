import { afterEach, describe, expect, it } from 'vitest'
import { mkdtemp, mkdir, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import { LocalFileStore } from '../src/local-file-store'

const temporaryDirectories: string[] = []

afterEach(async () => {
  await Promise.all(temporaryDirectories.splice(0).map((path) => rm(path, { recursive: true })))
})

describe('LocalFileStore Markdown asset roots', () => {
  it('reads a bounded session-owned relative image in chunks', async () => {
    const directory = await mkdtemp(join(tmpdir(), 'tandemfolio-markdown-assets-'))
    temporaryDirectories.push(directory)
    await mkdir(join(directory, 'assets'))
    const markdownPath = join(directory, 'notes.md')
    await writeFile(markdownPath, '# Notes')
    const png = Buffer.from(
      'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M/wHwAF/gL+Xw4bWQAAAABJRU5ErkJggg==',
      'base64',
    )
    await writeFile(join(directory, 'assets', 'pixel.png'), png)
    const store = new LocalFileStore()
    const rootId = store.registerAssetRoot('session-1', markdownPath)

    await expect(
      store.readLocalAsset('session-1', rootId, 'assets/pixel.png', 0, 8),
    ).resolves.toMatchObject({
      offset: 0,
      size: png.length,
      mime: 'image/png',
      data: png.subarray(0, 8).toString('base64'),
      nextOffset: 8,
      eof: false,
    })
  })

  it('rejects remote URLs and cross-session asset roots', async () => {
    const directory = await mkdtemp(join(tmpdir(), 'tandemfolio-markdown-assets-'))
    temporaryDirectories.push(directory)
    const markdownPath = join(directory, 'notes.md')
    await writeFile(markdownPath, '# Notes')
    const store = new LocalFileStore()
    const rootId = store.registerAssetRoot('session-1', markdownPath)

    await expect(
      store.readLocalAsset('session-1', rootId, 'https://example.com/pixel.png', 0, 8),
    ).rejects.toThrow('Only local Markdown image paths can use the asset bridge.')
    await expect(store.readLocalAsset('session-2', rootId, 'pixel.png', 0, 8)).rejects.toThrow(
      `Unknown Markdown asset root: ${rootId}`,
    )
  })
})
