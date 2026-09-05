import { describe, expect, it } from 'vitest'

import { createEditJournal } from '../src/renderer/edit-journal'
import { executeXlsxOperation } from '../src/renderer/operations/registry'

function runtime() {
  return { univerAPI: { getActiveWorkbook: () => ({ getId: () => undefined }) } }
}

describe('native workbook structure protection operation', () => {
  it('records a passwordless structure lock through the file journal', async () => {
    const editJournal = createEditJournal()
    let pending = 0
    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.document.set_protection',
          arguments: { lockStructure: true },
        },
        {
          runtime: () => runtime() as never,
          state: () => ({ editJournal, file: { styles: [] } }) as never,
          setPendingEdits: (count) => {
            pending = count
          },
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.document.set_protection',
      ok: true,
      output: { lockStructure: true },
    })
    expect(editJournal.workbookProtection.desired).toBe(true)
    expect(pending).toBe(1)
  })

  it('fails closed when asked to remove a password-protected structure lock', async () => {
    const editJournal = createEditJournal()
    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.document.set_protection',
          arguments: { lockStructure: false },
        },
        {
          runtime: () => runtime() as never,
          state: () => ({
            editJournal,
            file: {
              styles: [],
              workbookProtection: { lockStructure: true, hasPassword: true },
            },
          }) as never,
        },
      ),
    ).resolves.toMatchObject({
      handled: true,
      ok: false,
      error: 'execution_failed',
    })
    expect(editJournal.workbookProtection.desired).toBeNull()
  })
})
