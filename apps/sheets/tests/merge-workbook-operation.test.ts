import { describe, expect, it, vi } from 'vitest'

import { xlsxOperationCatalog } from '../src/renderer/operations/catalog'
import { executeXlsxOperation } from '../src/renderer/operations/registry'

describe('xlsx.workbook.merge_staged', () => {
  it('is internal and delegates hydrated bytes to the mounted merge kernel', async () => {
    const descriptor = xlsxOperationCatalog.operations.find(
      (operation) => operation.id === 'xlsx.workbook.merge_staged',
    )
    expect(descriptor?.visibility).toBe('internal')

    const data = new ArrayBuffer(12)
    const mergeStaged = vi.fn(async () => ({
      importedSheets: 2,
      files: 1,
      sheetNames: ['Data', 'Summary'],
    }))
    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.workbook.merge_staged',
          arguments: { blobId: 'blob-1', name: 'source.xlsx', size: 12, data },
        },
        { runtime: () => null, mergeStaged },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.workbook.merge_staged',
      ok: true,
      output: {
        merged: true,
        fileName: 'source.xlsx',
        importedSheets: 2,
        sheetNames: ['Data', 'Summary'],
      },
    })
    expect(mergeStaged).toHaveBeenCalledWith({ name: 'source.xlsx', data })
  })

  it('rejects mismatched staged byte lengths before calling the merge kernel', async () => {
    const mergeStaged = vi.fn()
    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.workbook.merge_staged',
          arguments: {
            blobId: 'blob-1',
            name: 'source.xlsx',
            size: 13,
            data: new ArrayBuffer(12),
          },
        },
        { runtime: () => null, mergeStaged },
      ),
    ).resolves.toMatchObject({
      handled: true,
      operationId: 'xlsx.workbook.merge_staged',
      ok: false,
      error: 'invalid_arguments',
    })
    expect(mergeStaged).not.toHaveBeenCalled()
  })
})
