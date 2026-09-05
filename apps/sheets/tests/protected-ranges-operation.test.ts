import { describe, expect, it } from 'vitest'

import { createEditJournal } from '../src/renderer/edit-journal'
import { executeXlsxOperation } from '../src/renderer/operations/registry'

describe('native allow-edit ranges operation', () => {
  it('replaces the active sheet range set through the shared file journal', async () => {
    const editJournal = createEditJournal()
    const sheetProtectedRanges = new Map([
      ['sheet-budget', [{ name: 'Old', sqref: 'A1', hasPassword: false }]],
    ])
    const worksheet = { getSheetName: () => 'Budget', getSheetId: () => 'sheet-budget' }
    const workbook = {
      getId: () => undefined,
      getSheets: () => [worksheet],
      setActiveSheet: () => undefined,
    }
    let pending = 0

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.sheet.set_protected_ranges',
          arguments: {
            sheet: 'Budget',
            ranges: [{ name: 'Inputs', sqref: 'B2:B10' }],
          },
        },
        {
          runtime: () => ({ univerAPI: { getActiveWorkbook: () => workbook } }) as never,
          state: () => ({ editJournal, sheetProtectedRanges }) as never,
          setPendingEdits: (count) => {
            pending = count
          },
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.sheet.set_protected_ranges',
      ok: true,
      output: { sheet: 'Budget', ranges: 1 },
    })
    expect(sheetProtectedRanges.get('sheet-budget')).toEqual([
      { name: 'Inputs', sqref: 'B2:B10', hasPassword: false },
    ])
    expect(editJournal.protectedRangesDirty.has('sheet-budget')).toBe(true)
    expect(pending).toBe(1)
  })

  it('does not rewrite password-protected allow-edit ranges', async () => {
    const editJournal = createEditJournal()
    const sheetProtectedRanges = new Map([
      ['sheet-budget', [{ name: 'Secret', sqref: 'A1', hasPassword: true }]],
    ])
    const worksheet = { getSheetName: () => 'Budget', getSheetId: () => 'sheet-budget' }
    const workbook = {
      getId: () => undefined,
      getSheets: () => [worksheet],
      setActiveSheet: () => undefined,
    }
    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.sheet.set_protected_ranges',
          arguments: { sheet: 'Budget', ranges: [] },
        },
        {
          runtime: () => ({ univerAPI: { getActiveWorkbook: () => workbook } }) as never,
          state: () => ({ editJournal, sheetProtectedRanges }) as never,
        },
      ),
    ).resolves.toMatchObject({ handled: true, ok: false, error: 'execution_failed' })
    expect(editJournal.protectedRangesDirty.size).toBe(0)
  })
})
