import { beforeEach, describe, expect, it } from 'vitest'

import { createBrowserPdfHost } from '../src/renderer/host/browser-pdf-api'

describe('PDF browser signature library', () => {
  beforeEach(() => localStorage.clear())

  it('persists, deduplicates, lists, and removes reusable native signatures', async () => {
    const host = createBrowserPdfHost()
    const signature = { kind: 'strokes' as const, paths: [[1, 2, 3, 4]], width: 420, height: 150 }

    const first = await host.api.addSavedSignature(signature)
    const second = await host.api.addSavedSignature(signature)
    expect(first).toHaveLength(1)
    expect(second).toHaveLength(1)
    await expect(host.api.listSavedSignatures()).resolves.toEqual(second)

    await expect(host.api.removeSavedSignature(second[0]!.id)).resolves.toEqual([])
    await expect(host.api.listSavedSignatures()).resolves.toEqual([])
  })
})
