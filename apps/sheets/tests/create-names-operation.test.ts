import { describe, expect, it, vi } from 'vitest'

import { xlsxOperationCatalog } from '../src/renderer/operations/catalog'
import { executeXlsxOperation } from '../src/renderer/operations/registry'

describe('xlsx.defined_name.create_from_selection', () => {
  it('is a public typed operation and delegates one explicit sheet range to the native kernel', async () => {
    expect(
      xlsxOperationCatalog.operations.some(
        (operation) => operation.id === 'xlsx.defined_name.create_from_selection',
      ),
    ).toBe(true)
    const worksheet = { getSheetId: () => 'sheet-1', getSheetName: () => 'Data' }
    const workbook = {
      getSheets: () => [worksheet],
      setActiveSheet: vi.fn(),
    }
    const createNamesFromSelection = vi.fn(() => ({ created: 2, skipped: 1 }))

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.defined_name.create_from_selection',
          arguments: { sheet: 'Data', range: 'A1:B3', labels: 'top' },
        },
        {
          runtime: () => ({ univerAPI: { getActiveWorkbook: () => workbook } }) as never,
          createNamesFromSelection,
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.defined_name.create_from_selection',
      ok: true,
      output: {
        sheet: 'Data',
        range: 'A1:B3',
        labels: 'top',
        created: 2,
        skipped: 1,
      },
    })
    expect(createNamesFromSelection).toHaveBeenCalledWith({
      sheetId: 'sheet-1',
      range: 'A1:B3',
      labels: 'top',
    })
  })
})
