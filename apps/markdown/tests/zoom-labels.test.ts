import { describe, expect, it } from 'vitest'

import { strings } from '../src/renderer/i18n/strings'

const locales = Object.keys(strings) as Array<keyof typeof strings>

describe('Markdown zoom labels', () => {
  it.each(locales)('locale %s labels every zoom control', (locale) => {
    const table = strings[locale] as Record<string, unknown>
    for (const key of ['zoom', 'zoomIn', 'zoomOut']) {
      expect(typeof table[key]).toBe('string')
      expect((table[key] as string).trim().length).toBeGreaterThan(0)
    }
  })
})
