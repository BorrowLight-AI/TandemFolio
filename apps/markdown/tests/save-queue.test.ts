import { describe, expect, it } from 'vitest'

import { createMarkdownSaveQueue } from '../src/renderer/host/save-queue'

describe('Markdown save queue', () => {
  it('serializes overlapping save and Save As requests in arrival order', async () => {
    const events: string[] = []
    let finishFirst!: () => void
    const firstBlocked = new Promise<void>((resolve) => {
      finishFirst = resolve
    })
    let calls = 0
    const save = createMarkdownSaveQueue(async (saveAs) => {
      calls += 1
      events.push(`start:${saveAs}`)
      if (calls === 1) await firstBlocked
      events.push(`finish:${saveAs}`)
      return { ok: true as const, fileName: saveAs ? 'copy.md' : 'notes.md' }
    })

    const first = save(false)
    const second = save(true)
    await Promise.resolve()
    expect(events).toEqual(['start:false'])
    finishFirst()

    await expect(Promise.all([first, second])).resolves.toEqual([
      { ok: true, fileName: 'notes.md' },
      { ok: true, fileName: 'copy.md' },
    ])
    expect(events).toEqual(['start:false', 'finish:false', 'start:true', 'finish:true'])
  })

  it('continues with the next request after a failed write', async () => {
    let calls = 0
    const save = createMarkdownSaveQueue(async () => {
      calls += 1
      if (calls === 1) throw new Error('disk failure')
      return { ok: true as const, fileName: 'recovered.md' }
    })

    await expect(save(false)).rejects.toThrow('disk failure')
    await expect(save(false)).resolves.toEqual({ ok: true, fileName: 'recovered.md' })
  })
})
