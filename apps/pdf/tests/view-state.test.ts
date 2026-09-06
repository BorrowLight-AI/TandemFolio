import { describe, expect, it } from 'vitest'
import { loadViewState, saveViewState } from '../src/renderer/view-state'

function memoryStorage() {
  const values = new Map<string, string>()
  return {
    getItem: (key: string) => values.get(key) ?? null,
    setItem: (key: string, value: string) => values.set(key, value),
  }
}

describe('per-file PDF reading position', () => {
  it('restores the page, fractional row offset, zoom and fit mode for the same file', () => {
    const storage = memoryStorage()
    saveViewState(
      'browser://document-a.pdf',
      { page: 7, frac: 0.42, scale: 1.25, fitMode: null },
      storage,
      1_000,
    )

    expect(loadViewState('browser://document-a.pdf', storage)).toEqual({
      page: 7,
      frac: 0.42,
      scale: 1.25,
      fitMode: null,
    })
    expect(loadViewState('browser://document-b.pdf', storage)).toBeNull()
  })
})
