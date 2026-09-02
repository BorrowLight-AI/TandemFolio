import { readFile } from 'node:fs/promises'
import { fileURLToPath } from 'node:url'
import { SessionError } from './session-store'

const FONT_FILES = new Set([
  'Caladea-Bold.ttf',
  'Caladea-BoldItalic.ttf',
  'Caladea-Italic.ttf',
  'Caladea-Regular.ttf',
  'Carlito-Bold.ttf',
  'Carlito-BoldItalic.ttf',
  'Carlito-Italic.ttf',
  'Carlito-Regular.ttf',
  'GenOfficeSansKR-Regular-subset.woff2',
  'GenOfficeSansKR-Regular-subset.otf',
  'GenOfficeSerifKR-Regular-subset.woff2',
  'LiberationMono-Bold.ttf',
  'LiberationMono-BoldItalic.ttf',
  'LiberationMono-Italic.ttf',
  'LiberationMono-Regular.ttf',
  'LiberationSans-Bold.ttf',
  'LiberationSans-BoldItalic.ttf',
  'LiberationSans-Italic.ttf',
  'LiberationSans-Regular.ttf',
  'LiberationSerif-Bold.ttf',
  'LiberationSerif-BoldItalic.ttf',
  'LiberationSerif-Italic.ttf',
  'LiberationSerif-Regular.ttf',
  'NotoNaskhArabic-Regular-subset.woff2',
  'NotoNaskhArabic-Regular-subset.ttf',
  'NotoSansArabic-Regular-subset.woff2',
  'NotoSansCJKsc-Regular-subset.woff2',
  'NotoSansCJKsc-Regular-subset.otf',
  'NotoSerifCJKsc-Regular-subset.woff2',
])

export interface FontAssetChunk {
  fileName: string
  offset: number
  data: string
  nextOffset: number
  eof: boolean
}

export class FontAssetStore {
  readonly #cache = new Map<string, Buffer>()

  async read(fileName: string, offset: number, length: number): Promise<FontAssetChunk> {
    if (!FONT_FILES.has(fileName)) {
      throw new SessionError('invalid_arguments', `Unknown bundled font asset: ${fileName}`)
    }
    const data = await this.#load(fileName)
    if (offset > data.length) {
      throw new SessionError('invalid_arguments', 'Font chunk offset exceeds the asset size.')
    }
    const end = Math.min(offset + length, data.length)
    return {
      fileName,
      offset,
      data: data.subarray(offset, end).toString('base64'),
      nextOffset: end,
      eof: end >= data.length,
    }
  }

  async #load(fileName: string): Promise<Buffer> {
    const cached = this.#cache.get(fileName)
    if (cached) return cached

    const candidates = [
      new URL(`../assets/fonts/${fileName}`, import.meta.url),
      new URL(`../../docs/src/renderer/fonts/${fileName}`, import.meta.url),
    ]
    for (const candidate of candidates) {
      try {
        const data = await readFile(fileURLToPath(candidate))
        this.#cache.set(fileName, data)
        return data
      } catch (error) {
        if ((error as NodeJS.ErrnoException).code !== 'ENOENT') throw error
      }
    }
    throw new SessionError('execution_failed', `Bundled font asset is missing: ${fileName}`)
  }
}
