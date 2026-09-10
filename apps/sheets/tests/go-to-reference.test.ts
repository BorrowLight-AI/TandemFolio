import { describe, expect, it, vi } from 'vitest'

import {
  goToReference,
  type DataToolsContext,
} from '../src/renderer/data-tools-actions'

describe('goToReference', () => {
  it('commits an active cell edit before moving the workbook selection', async () => {
    let finishEditing!: () => void
    const editingFinished = new Promise<void>((resolve) => {
      finishEditing = resolve
    })
    const range = {
      getSheetId: () => 'sheet-1',
      getRow: () => 1,
      getColumn: () => 1,
    }
    const worksheet = {
      getRange: vi.fn(() => range),
      scrollToCell: vi.fn(),
    }
    const setActiveRange = vi.fn()
    const workbook = {
      getActiveSheet: () => worksheet,
      getDefinedNames: () => [],
      getSheetBySheetId: () => worksheet,
      isCellEditing: () => true,
      endEditingAsync: vi.fn(() => editingFinished),
      setActiveRange,
    }
    const focus = vi.fn()
    const ctx = {
      univerRef: {
        current: {
          univerAPI: { getActiveWorkbook: () => workbook },
          univer: { __getInjector: () => ({ get: () => ({ focus }) }) },
        },
      },
      lazyWorkbookRef: { current: null },
      setMessage: vi.fn(),
    } as unknown as DataToolsContext

    const result = Promise.resolve(goToReference(ctx, 'B2:D8'))

    expect(setActiveRange).not.toHaveBeenCalled()
    finishEditing()
    await expect(result).resolves.toBeNull()
    expect(setActiveRange).toHaveBeenCalledWith(range)
    expect(worksheet.scrollToCell).toHaveBeenCalledWith(1, 1)
    expect(focus).toHaveBeenCalledOnce()
  })
})
