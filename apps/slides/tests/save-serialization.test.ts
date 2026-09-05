// Modified by TandemFolio contributors: browser-host regression coverage (2026-09-05).
import { afterEach, describe, expect, it, vi } from 'vitest'

// file-actions pulls export-render → pptx-render → konva → native `canvas`
// (missing in CI). Mock the UI-only modules so this unit test exercises
// only the save-serialization queue.
vi.mock('../src/renderer/export-render', () => ({ renderSlidesToPngBase64: vi.fn() }))
vi.mock('../src/renderer/components/toast-bus', () => ({ showToast: vi.fn() }))
vi.mock('../src/renderer/i18n/locale', () => ({ t: (k: string) => k }))

import { save, saveAs } from '../src/renderer/file-actions'
import type { ActionCtx } from '../src/renderer/action-context'

function ctx(): ActionCtx {
  return {
    flushActiveEditRef: { current: () => Promise.resolve() },
    editingActiveRef: { current: false },
    flushNotes: () => Promise.resolve(),
    slides: [],
    current: 0,
    setSlides: () => {},
    setSelectedIds: () => {},
    setEnteredGroupId: () => {},
    setEditing: () => {},
    setEditingCell: () => {},
    setDirty: () => {},
    setPath: () => {},
    setStatus: () => {},
    path: '/test/deck.pptx',
  } as unknown as ActionCtx
}

describe('slides save serialization', () => {
  afterEach(() => vi.unstubAllGlobals())
  it('queues concurrent saves so only one renderer write runs at a time', async () => {
    let inFlight = 0
    let maxConcurrent = 0
    const mockSave = vi.fn(async () => {
      inFlight += 1
      maxConcurrent = Math.max(maxConcurrent, inFlight)
      // Simulate a slow write.
      await new Promise((resolve) => setTimeout(resolve, 50))
      inFlight -= 1
      return { ok: true }
    })
    vi.stubGlobal('document', { activeElement: null })
    vi.stubGlobal('window', { slidesApi: { save: mockSave } })

    // Fire three saves simultaneously — they must queue, not overlap.
    const [a, b, c] = await Promise.all([save(ctx()), save(ctx(), true), save(ctx())])

    expect(a).toBe(true)
    expect(b).toBe(true)
    expect(c).toBe(true)
    expect(mockSave).toHaveBeenCalledTimes(3)
    expect(maxConcurrent).toBe(1)

    vi.unstubAllGlobals()
  })

  it('reports failure without blocking the queue for the next caller', async () => {
    let calls = 0
    const mockSave = vi.fn(async () => {
      calls += 1
      if (calls === 1) return { ok: false, error: 'disk full' }
      return { ok: true }
    })
    vi.stubGlobal('document', { activeElement: null })
    vi.stubGlobal('window', { slidesApi: { save: mockSave } })

    const first = save(ctx(), true)
    const second = save(ctx(), true)
    const [r1, r2] = await Promise.all([first, second])

    expect(r1).toBe(false)
    expect(r2).toBe(true)
    expect(mockSave).toHaveBeenCalledTimes(2)

    vi.unstubAllGlobals()
  })
  it('queues Save As behind Save and continues after a rejected write', async () => {
    let release!: () => void
    const blocked = new Promise<void>((resolve) => {
      release = resolve
    })
    const calls: string[] = []
    vi.stubGlobal('document', { activeElement: null })
    vi.stubGlobal('window', {
      slidesApi: {
        save: async () => {
          calls.push('save')
          await blocked
          throw new Error('write rejected')
        },
        saveAs: async (name: string) => {
          calls.push(name)
          return { ok: true, path: '/test/' + name }
        },
      },
    })
    const saving = save(ctx())
    const rejected = expect(saving).rejects.toThrow('write rejected')
    const copying = saveAs(ctx(), 'copy.pptx')
    await vi.waitFor(() => expect(calls).toEqual(['save']))
    release()
    await rejected
    await copying
    expect(calls).toEqual(['save', 'copy.pptx'])
  })
})
