/**
 * A dirty workbook gets a crash-recovery copy through the
 * same save pipeline, but recovery mode must never touch the opened file, prompt,
 * or clear the journal — the whole point is that unsaved work survives a crash.
 */
import { describe, expect, it, vi, beforeEach } from 'vitest'
import { handleSave, type SaveContext } from '../src/renderer/save-actions'
import { createEditJournal, recordSetRangeValues } from '../src/renderer/edit-journal'

const saveWorkbookEdits = vi.fn()
const writeWorkbookRecovery = vi.fn()

beforeEach(() => {
  saveWorkbookEdits.mockReset().mockResolvedValue({ canceled: true })
  writeWorkbookRecovery.mockReset().mockResolvedValue({ ok: true })
  ;(globalThis as unknown as { window: unknown }).window = {
    desktopApi: { saveWorkbookEdits, writeWorkbookRecovery },
  }
})

function ctxWith(opts: { dirty: boolean; needsSaveAs?: boolean }): {
  ctx: SaveContext
  messages: string[]
} {
  const journal = createEditJournal()
  if (opts.dirty) recordSetRangeValues(journal, 'sheet-1', { 0: { 0: { v: 'edited' } } })
  const messages: string[] = []
  return {
    messages,
    ctx: {
      univerRef: { current: null },
      lazyWorkbookRef: {
        current: {
          editJournal: journal,
          recalc: {
            timer: null,
            generation: 0,
            failed: false,
            formulaCells: new Map(),
            overlay: new Map(),
          },
          file: {
            sessionId: '11111111-1111-4111-8111-111111111111',
            needsSaveAs: !!opts.needsSaveAs,
          },
        },
      } as never,
      setMessage: (m: string) => messages.push(m),
      openLazyWorkbook: () => {},
    },
  }
}

describe('handleSave recovery mode', () => {
  it('writes a recovery copy and never the opened file', async () => {
    const { ctx, messages } = ctxWith({ dirty: true })
    await handleSave(ctx, 'recovery')
    expect(writeWorkbookRecovery).toHaveBeenCalledTimes(1)
    expect(saveWorkbookEdits).not.toHaveBeenCalled()
    // silent: no status messages for a background copy
    expect(messages).toEqual([])
  })

  it('returns the serialized recovery bytes to the live-session adapter', async () => {
    const data = Uint8Array.from([80, 75, 3, 4]).buffer
    writeWorkbookRecovery.mockResolvedValue({ ok: true, fileName: 'budget.xlsx', data })
    const { ctx } = ctxWith({ dirty: true })

    await expect(handleSave(ctx, 'recovery')).resolves.toEqual({
      ok: true,
      fileName: 'budget.xlsx',
      data,
    })
  })

  it('sends the same payload shape the save pipeline gets', async () => {
    const { ctx } = ctxWith({ dirty: true })
    await handleSave(ctx, 'recovery')
    const payload = writeWorkbookRecovery.mock.calls[0]![0] as Record<string, unknown>
    expect(payload.sessionId).toBe('11111111-1111-4111-8111-111111111111')
    // 'recovery' is a renderer-side mode; the main process still sees a normal save request
    expect(payload.mode).toBe('save')
    expect(Array.isArray(payload.edits)).toBe(true)
    expect((payload.edits as unknown[]).length).toBe(1)
    expect(payload).toHaveProperty('structuralOps')
    expect(payload).toHaveProperty('sparklineAdditions')
  })

  it('does nothing when there is nothing pending', async () => {
    const { ctx, messages } = ctxWith({ dirty: false })
    await handleSave(ctx, 'recovery')
    expect(writeWorkbookRecovery).not.toHaveBeenCalled()
    expect(messages).toEqual([])
  })

  it('a failed copy is swallowed (best-effort, never surfaces)', async () => {
    writeWorkbookRecovery.mockRejectedValue(new Error('disk full'))
    const { ctx, messages } = ctxWith({ dirty: true })
    await expect(handleSave(ctx, 'recovery')).resolves.toBeUndefined()
    expect(messages).toEqual([])
  })
})

describe('handleSave explicit save result', () => {
  it('returns the persisted file identity after save-as succeeds', async () => {
    saveWorkbookEdits.mockResolvedValue({
      canceled: false,
      file: {
        sessionId: '22222222-2222-4222-8222-222222222222',
        name: 'budget-copy.xlsx',
        entryCount: 12,
      },
      touchedEntries: ['xl/worksheets/sheet1.xml'],
    })
    const { ctx } = ctxWith({ dirty: false })

    await expect(handleSave(ctx, 'save-as')).resolves.toEqual({
      ok: true,
      fileName: 'budget-copy.xlsx',
    })
  })

  it('returns an explicit failure when save-as is canceled', async () => {
    saveWorkbookEdits.mockResolvedValue({ canceled: true })
    const { ctx, messages } = ctxWith({ dirty: false })

    const result = await handleSave(ctx, 'save-as')
    expect(result).toEqual({
      ok: false,
      message: messages.at(-1),
    })
  })

  it('returns an explicit failure when the save pipeline throws', async () => {
    saveWorkbookEdits.mockRejectedValue(new Error('disk unavailable'))
    const { ctx, messages } = ctxWith({ dirty: false })

    const result = await handleSave(ctx, 'save-as')
    expect(result).toEqual({
      ok: false,
      message: messages.at(-1),
    })
  })
})
