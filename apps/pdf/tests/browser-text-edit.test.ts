import { readFile } from 'node:fs/promises'
import { resolve } from 'node:path'
import { describe, expect, it } from 'vitest'
import { browserFontCoversText } from '../src/domain/browser-text-edit'

describe('browser PDF edit-font coverage', () => {
  it('accepts mapped Latin text and rejects an unsupported emoji before PDFium authors it', async () => {
    const font = new Uint8Array(
      await readFile(resolve(process.cwd(), '../docs/src/renderer/fonts/LiberationSans-Regular.ttf')),
    )
    expect(browserFontCoversText(font, 'Bounded text €')).toBe(true)
    expect(browserFontCoversText(font, 'Pay 🦄 now')).toBe(false)
  })
})
