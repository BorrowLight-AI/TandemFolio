import { describe, expect, it } from 'vitest'

import { createEditJournal } from '../src/renderer/edit-journal'
import { executeXlsxOperation } from '../src/renderer/operations/registry'

describe('native workbook theme operation', () => {
  it('uses the same theme state and save journal as the Page Layout ribbon', async () => {
    const editJournal = createEditJournal()
    const state = {
      editJournal,
      file: {
        themeColors: Array.from({ length: 12 }, () => '#000000'),
        themeFonts: { major: 'Old Major', minor: 'Old Minor' },
        styles: [],
      },
    }
    let pendingEdits = 0
    const runtime = {
      univerAPI: {
        getActiveWorkbook: () => ({ getId: () => undefined }),
      },
    }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.document.set_theme',
          arguments: { mode: 'theme', scheme: 'forest' },
        },
        {
          runtime: () => runtime as never,
          state: () => state as never,
          setPendingEdits: (count) => {
            pendingEdits = count
          },
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.document.set_theme',
      ok: true,
      output: { mode: 'theme', scheme: 'forest', name: 'Forest' },
    })
    expect(state.file.themeColors[4]).toBe('#217346')
    expect(state.file.themeFonts).toEqual({ major: 'Candara', minor: 'Candara' })
    expect(editJournal.theme).toMatchObject({ colors: { name: 'Forest' }, fonts: { name: 'Candara' } })
    expect(pendingEdits).toBe(2)
  })
})
