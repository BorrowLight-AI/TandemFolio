import { BooleanNumber } from '@univerjs/core'
import { describe, expect, it } from 'vitest'

import { buildCustomFilters } from '../src/renderer/AdvancedFilterDialog'
import {
  handleApplyAdvancedFilter,
  type DataToolsContext,
} from '../src/renderer/data-tools-actions'

describe('buildCustomFilters', () => {
  it('maps two conditions with the and flag and operator values', () => {
    expect(
      buildCustomFilters(true, [
        { operator: 'greaterThanOrEqual', val: '10' },
        { operator: 'lessThan', val: '20' },
      ]),
    ).toEqual({
      and: BooleanNumber.TRUE,
      customFilters: [
        { val: '10', operator: 'greaterThanOrEqual' },
        { val: '20', operator: 'lessThan' },
      ],
    })
  })

  it('omits the and flag for OR pairs and single conditions', () => {
    expect(
      buildCustomFilters(false, [
        { operator: 'notEqual', val: 'x' },
        { operator: 'greaterThan', val: '3' },
      ]),
    ).toEqual({
      customFilters: [
        { val: 'x', operator: 'notEqual' },
        { val: '3', operator: 'greaterThan' },
      ],
    })
    expect(buildCustomFilters(true, [{ operator: 'lessThanOrEqual', val: '5' }])).toEqual({
      customFilters: [{ val: '5', operator: 'lessThanOrEqual' }],
    })
  })

  it('drops the explicit equal operator (numeric-only in Univer; OOXML default)', () => {
    expect(buildCustomFilters(true, [{ operator: 'equal', val: 'East' }])).toEqual({
      customFilters: [{ val: 'East' }],
    })
  })

  it('rejects zero and more-than-two conditions', () => {
    expect(() => buildCustomFilters(true, [])).toThrow(/至少需要一个条件/)
    expect(() =>
      buildCustomFilters(true, [
        { operator: 'equal', val: '1' },
        { operator: 'equal', val: '2' },
        { operator: 'equal', val: '3' },
      ]),
    ).toThrow(/最多支持两个条件/)
  })
})

describe('handleApplyAdvancedFilter', () => {
  it('applies the dialog criteria through the explicit workbook filter command', async () => {
    const commands: Array<{ id: string; params: unknown }> = []
    const messages: string[] = []
    const bounds = { startRow: 0, endRow: 20, startColumn: 0, endColumn: 2 }
    const worksheet = {
      getSheetId: () => 'sheet-1',
      getFilter: () => ({ getRange: () => ({ getRange: () => bounds }) }),
    }
    const workbook = {
      getId: () => 'workbook-1',
      getActiveSheet: () => worksheet,
    }
    const ctx = {
      univerRef: {
        current: {
          univerAPI: {
            getActiveWorkbook: () => workbook,
            executeCommand: async (id: string, params: unknown) => {
              commands.push({ id, params })
              return true
            },
          },
        },
      },
      setMessage: (message: string) => messages.push(message),
    } as unknown as DataToolsContext

    await expect(
      handleApplyAdvancedFilter(ctx, {
        colId: 1,
        and: true,
        filters: [
          { operator: 'greaterThanOrEqual', val: '10' },
          { operator: 'lessThan', val: '20' },
        ],
      }),
    ).resolves.toBeNull()
    expect(commands).toEqual([
      {
        id: 'sheet.command.set-filter-criteria',
        params: {
          unitId: 'workbook-1',
          subUnitId: 'sheet-1',
          col: 1,
          criteria: {
            colId: 1,
            customFilters: {
              and: BooleanNumber.TRUE,
              customFilters: [
                { operator: 'greaterThanOrEqual', val: '10' },
                { operator: 'lessThan', val: '20' },
              ],
            },
          },
        },
      },
    ])
    expect(messages).toHaveLength(1)
  })
})
