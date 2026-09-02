import {
  BooleanNumber,
  BorderStyleTypes,
  BorderType,
  HorizontalAlign,
  VerticalAlign,
  WrapStrategy,
} from '@univerjs/core'
import { LexerTreeBuilder } from '@univerjs/engine-formula'
import { describe, expect, it } from 'vitest'

import { createEditJournal } from '../src/renderer/edit-journal'
import { xlsxOperationCatalog } from '../src/renderer/operations/catalog'
import { executeXlsxOperation } from '../src/renderer/operations/registry'
import type { LazyWorkbookState } from '../src/renderer/univer-state'

describe('XLSX operation registry', () => {
  it('rejects every retired public XLSX alias before entering workbook execution', async () => {
    const services = {
      runtime: () => null,
      loadStaged: async () => undefined,
      save: async () => ({ ok: true as const, fileName: 'budget.xlsx' }),
    }

    for (const operation of [
      'set_cell_value',
      'set_cell_formula',
      'delete_columns',
      'insert_columns',
      'set_cols_hidden',
      'set_range_values',
      'find_replace',
      'add_table_row',
      'delete_table_row',
      'add_table_column',
      'delete_table_column',
      'delete_table',
      'set_rows_hidden',
      'insert_rows',
      'add_sheet',
      'duplicate_sheet',
      'set_sheet_hidden',
      'rename_sheet',
      'delete_sheet',
      'move_sheet',
      'delete_rows',
      'save',
    ] as const) {
      await expect(executeXlsxOperation({ operation, arguments: {} }, services)).resolves.toEqual({
        handled: false,
      })
    }
  })

  it('sets one cell through xlsx.cell.set_value and activates the edited range', async () => {
    const values = new Map<string, unknown>()
    let activeRange: string | null = null
    const worksheet = {
      getSheetName: () => 'Budget',
      getRange: (address: string) => ({
        setValue: (value: unknown) => values.set(address, value),
        activate: () => {
          activeRange = address
        },
      }),
    }
    const workbook = {
      getSheets: () => [worksheet],
      setActiveSheet: () => undefined,
    }
    const runtime = { univerAPI: { getActiveWorkbook: () => workbook } }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.cell.set_value',
          arguments: { sheet: 'Budget', address: 'b2', value: 42 },
        },
        {
          runtime: () => runtime,
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.cell.set_value',
      ok: true,
      output: { changed: 1, sheet: 'Budget', range: 'B2' },
    })
    expect(values.get('b2')).toBe(42)
    expect(activeRange).toBe('b2')
  })

  it('sets one explicit formula through the native formula write and activation route', async () => {
    let writtenFormula: string | null = null
    let activeRange: string | null = null
    const worksheet = {
      getSheetName: () => 'Budget',
      getRange: (address: string) => ({
        setFormula: (formula: string) => {
          writtenFormula = formula
        },
        activate: () => {
          activeRange = address
        },
      }),
    }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.cell.set_formula',
          arguments: { sheet: 'Budget', address: 'c4', formula: '=SUM(A1:A3)' },
        },
        {
          runtime: () => ({
            univerAPI: {
              getActiveWorkbook: () => ({
                getSheets: () => [worksheet],
                setActiveSheet: () => undefined,
              }),
            },
          }),
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.cell.set_formula',
      ok: true,
      output: { changed: 1, sheet: 'Budget', range: 'C4', formula: '=SUM(A1:A3)' },
    })
    expect(writtenFormula).toBe('=SUM(A1:A3)')
    expect(activeRange).toBe('C4')
  })

  it('duplicates one fully loaded worksheet to an explicit bounded name', async () => {
    const source = {
      getSheetName: () => 'Budget',
      getSheetId: () => 'sheet-budget',
    }
    const copy = {
      getSheetName: () => 'Budget Copy',
      getSheetId: () => 'sheet-copy',
    }
    const duplicates: unknown[] = []

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.sheet.duplicate',
          arguments: { sheet: 'Budget', name: 'Budget Copy' },
        },
        {
          runtime: () => ({
            univerAPI: {
              getActiveWorkbook: () => ({
                getSheets: () => [source],
                setActiveSheet: () => undefined,
              }),
            },
          }),
          isSheetDataComplete: () => true,
          duplicateSheet: (worksheet, name) => {
            duplicates.push({ worksheet, name })
            return copy
          },
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.sheet.duplicate',
      ok: true,
      output: { sheet: 'Budget', name: 'Budget Copy' },
    })
    expect(duplicates).toEqual([{ worksheet: source, name: 'Budget Copy' }])
  })

  it('sets one worksheet to an explicit visible final state', async () => {
    let shown = 0
    const worksheet = {
      getSheetName: () => 'Archive',
      getSheetId: () => 'sheet-archive',
      showSheet: () => {
        shown += 1
      },
    }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.sheet.set_visibility',
          arguments: { sheet: 'Archive', visible: true },
        },
        {
          runtime: () => ({
            univerAPI: {
              getActiveWorkbook: () => ({
                getSheets: () => [worksheet],
                setActiveSheet: () => undefined,
              }),
            },
          }),
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.sheet.set_visibility',
      ok: true,
      output: { sheet: 'Archive', visible: true },
    })
    expect(shown).toBe(1)
  })

  it('sets explicit row and column visibility through native axis commands', async () => {
    const calls: unknown[] = []
    const worksheet = {
      getSheetName: () => 'Budget',
      getSheetId: () => 'sheet-budget',
      hideRows: (start: number, count: number) => calls.push(['hideRows', start, count]),
      showColumns: (start: number, count: number) => calls.push(['showColumns', start, count]),
      getMaxRows: () => 1_000,
      getMaxColumns: () => 100,
    }
    const services = {
      runtime: () => ({
        univerAPI: {
          getActiveWorkbook: () => ({
            getSheets: () => [worksheet],
            setActiveSheet: () => undefined,
          }),
        },
      }),
      loadStaged: async () => undefined,
      save: async () => ({ ok: true as const, fileName: 'budget.xlsx' }),
    }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.row.set_visibility',
          arguments: { sheet: 'Budget', row: 3, count: 2, visible: false },
        },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.row.set_visibility',
      ok: true,
      output: { sheet: 'Budget', row: 3, count: 2, visible: false },
    })
    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.column.set_visibility',
          arguments: { sheet: 'Budget', column: 'C', count: 3, visible: true },
        },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.column.set_visibility',
      ok: true,
      output: { sheet: 'Budget', column: 'C', count: 3, visible: true },
    })
    expect(calls).toEqual([
      ['hideRows', 2, 2],
      ['showColumns', 2, 3],
    ])
  })

  it('routes the five bounded session-table lifecycle mutations through one table engine seam', async () => {
    const journal = createEditJournal()
    journal.tableAdds.push({
      sheetId: 'sheet-budget',
      area: { startRow: 0, startColumn: 0, endRow: 3, endColumn: 2 },
      name: 'SalesTable',
      columnNames: ['Region', 'Revenue', 'Cost'],
      bandedRows: true,
    })
    const worksheet = {
      getSheetName: () => 'Budget',
      getSheetId: () => 'sheet-budget',
    }
    const mutations: unknown[] = []
    const services = {
      runtime: () => ({
        univerAPI: {
          getActiveWorkbook: () => ({
            getSheets: () => [worksheet],
            setActiveSheet: () => undefined,
          }),
        },
      }),
      state: () => ({
        editJournal: journal,
        hyperlinkTargets: new Map(),
        sheetProtections: new Map(),
      }),
      mutateTable: async (operation: unknown) => {
        mutations.push(operation)
      },
      loadStaged: async () => undefined,
      save: async () => ({ ok: true as const, fileName: 'budget.xlsx' }),
    }

    const commands = [
      {
        operation: 'xlsx.table.insert_rows',
        arguments: { sheet: 'Budget', table: 'SalesTable', row: 2, count: 2 },
        output: { sheet: 'Budget', table: 'SalesTable', row: 2, count: 2 },
      },
      {
        operation: 'xlsx.table.delete_rows',
        arguments: { sheet: 'Budget', table: 'SalesTable', row: 2, count: 1 },
        output: { sheet: 'Budget', table: 'SalesTable', row: 2, count: 1 },
      },
      {
        operation: 'xlsx.table.insert_columns',
        arguments: {
          sheet: 'Budget',
          table: 'SalesTable',
          column: 2,
          columnName: 'Forecast',
          count: 1,
        },
        output: {
          sheet: 'Budget',
          table: 'SalesTable',
          column: 2,
          columnName: 'Forecast',
          count: 1,
        },
      },
      {
        operation: 'xlsx.table.delete_columns',
        arguments: { sheet: 'Budget', table: 'SalesTable', column: 2, count: 1 },
        output: { sheet: 'Budget', table: 'SalesTable', column: 2, count: 1 },
      },
      {
        operation: 'xlsx.table.convert_to_range',
        arguments: { sheet: 'Budget', table: 'SalesTable' },
        output: { sheet: 'Budget', table: 'SalesTable', converted: true },
      },
    ] as const

    for (const command of commands) {
      await expect(
        executeXlsxOperation(
          { operation: command.operation, arguments: { ...command.arguments } },
          services,
        ),
      ).resolves.toEqual({
        handled: true,
        operationId: command.operation,
        ok: true,
        output: command.output,
      })
    }
    expect(mutations).toEqual([
      {
        op: 'add_table_row',
        sheetId: 'sheet-budget',
        tableName: 'SalesTable',
        row: 2,
        count: 2,
      },
      {
        op: 'delete_table_row',
        sheetId: 'sheet-budget',
        tableName: 'SalesTable',
        row: 2,
        count: 1,
      },
      {
        op: 'add_table_column',
        sheetId: 'sheet-budget',
        tableName: 'SalesTable',
        column: 2,
        columnName: 'Forecast',
        count: 1,
      },
      {
        op: 'delete_table_column',
        sheetId: 'sheet-budget',
        tableName: 'SalesTable',
        column: 2,
        count: 1,
      },
      { op: 'convert_table_to_range', sheetId: 'sheet-budget', tableName: 'SalesTable' },
    ])
  })

  it('replaces bounded text values through one native range history seam', async () => {
    const replacements: unknown[] = []
    const worksheet = {
      getSheetName: () => 'Budget',
      getSheetId: () => 'sheet-budget',
      getRange: () => ({
        getRow: () => 0,
        getColumn: () => 0,
        getHeight: () => 4,
        getWidth: () => 2,
      }),
    }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.range.replace_text',
          arguments: {
            sheet: 'Budget',
            range: 'a1:b4',
            find: 'draft',
            replace: 'final',
            matchCase: false,
            wholeCell: false,
          },
        },
        {
          runtime: () => ({
            univerAPI: {
              getActiveWorkbook: () => ({
                getSheets: () => [worksheet],
                setActiveSheet: () => undefined,
              }),
            },
          }),
          replaceText: (input) => {
            replacements.push(input)
            return 2
          },
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.range.replace_text',
      ok: true,
      output: { sheet: 'Budget', range: 'A1:B4', changed: 2 },
    })
    expect(replacements).toEqual([
      {
        sheetId: 'sheet-budget',
        range: 'A1:B4',
        find: 'draft',
        replace: 'final',
        matchCase: false,
        wholeCell: false,
      },
    ])
  })

  it('sets a scalar matrix through xlsx.range.set_values and reports the changed cell count', async () => {
    let written: unknown[][] | null = null
    let activeRange: string | null = null
    const worksheet = {
      getSheetName: () => 'Budget',
      getRange: (address: string) => ({
        setValues: (values: unknown[][]) => {
          written = values
        },
        activate: () => {
          activeRange = address
        },
      }),
    }
    const workbook = {
      getSheets: () => [worksheet],
      setActiveSheet: () => undefined,
    }
    const runtime = { univerAPI: { getActiveWorkbook: () => workbook } }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.range.set_values',
          arguments: {
            sheet: 'Budget',
            range: 'a1:b2',
            values: [
              ['North', 10],
              ['South', 20],
            ],
          },
        },
        {
          runtime: () => runtime,
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.range.set_values',
      ok: true,
      output: { changed: 4, sheet: 'Budget', range: 'A1:B2' },
    })
    expect(written).toEqual([
      [{ v: 'North' }, { v: 10 }],
      [{ v: 'South' }, { v: 20 }],
    ])
    expect(activeRange).toBe('a1:b2')
  })

  it('copies computed scalar values between explicit worksheets and activates the destination', async () => {
    let written: unknown[][] | null = null
    let destinationActive = false
    const sourceWorksheet = {
      getSheetName: () => 'Budget',
      getSheetId: () => 'sheet-budget',
      getRange: (address: string) => ({
        getRow: () => 0,
        getColumn: () => 0,
        getHeight: () => 2,
        getWidth: () => 2,
        getValues: () => [
          [3, 'North'],
          [7, true],
        ],
        address,
      }),
    }
    const destinationWorksheet = {
      getSheetName: () => 'Archive',
      getSheetId: () => 'sheet-archive',
      getRange: (address: string) => ({
        getRow: () => 3,
        getColumn: () => 3,
        getHeight: () => 2,
        getWidth: () => 2,
        setValues: (values: unknown[][]) => {
          written = values
        },
        activate: () => {
          destinationActive = true
        },
        address,
      }),
    }
    const workbook = {
      getSheets: () => [sourceWorksheet, destinationWorksheet],
      setActiveSheet: () => undefined,
    }
    const runtime = { univerAPI: { getActiveWorkbook: () => workbook } }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.range.copy_values',
          arguments: {
            sourceSheet: 'Budget',
            sourceRange: 'a1:b2',
            destinationSheet: 'Archive',
            destinationRange: 'd4:e5',
          },
        },
        {
          runtime: () => runtime,
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.range.copy_values',
      ok: true,
      output: {
        sourceSheet: 'Budget',
        sourceRange: 'A1:B2',
        destinationSheet: 'Archive',
        destinationRange: 'D4:E5',
        copied: 4,
      },
    })
    expect(written).toEqual([
      [{ v: 3 }, { v: 'North' }],
      [{ v: 7 }, { v: true }],
    ])
    expect(destinationActive).toBe(true)
  })

  it('copies translated formulas and scalar cells between explicit worksheets in one write', async () => {
    let written: unknown[][] | null = null
    let destinationActive = false
    const sourceWorksheet = {
      getSheetName: () => 'Budget',
      getSheetId: () => 'sheet-budget',
      getRange: () => ({
        getRow: () => 0,
        getColumn: () => 0,
        getHeight: () => 2,
        getWidth: () => 2,
        getFormulas: () => [
          ['=A1+$A$1', ''],
          ['=SUM(A1:B1)', '=Sheet2!$C1'],
        ],
        getValues: () => [
          [3, 'North'],
          [7, true],
        ],
      }),
    }
    const destinationWorksheet = {
      getSheetName: () => 'Archive',
      getSheetId: () => 'sheet-archive',
      getRange: () => ({
        getRow: () => 3,
        getColumn: () => 3,
        getHeight: () => 2,
        getWidth: () => 2,
        setValues: (values: unknown[][]) => {
          written = values
        },
        activate: () => {
          destinationActive = true
        },
      }),
    }
    const workbook = {
      getSheets: () => [sourceWorksheet, destinationWorksheet],
      setActiveSheet: () => undefined,
    }
    const lexer = new LexerTreeBuilder()
    const runtime = {
      univerAPI: { getActiveWorkbook: () => workbook },
      univer: { __getInjector: () => ({ get: () => lexer }) },
    }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.range.copy_formulas',
          arguments: {
            sourceSheet: 'Budget',
            sourceRange: 'a1:b2',
            destinationSheet: 'Archive',
            destinationRange: 'd4:e5',
          },
        },
        {
          runtime: () => runtime,
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.range.copy_formulas',
      ok: true,
      output: {
        sourceSheet: 'Budget',
        sourceRange: 'A1:B2',
        destinationSheet: 'Archive',
        destinationRange: 'D4:E5',
        copied: 4,
      },
    })
    expect(written).toEqual([
      [{ f: '=D4+$A$1' }, { v: 'North' }],
      [{ f: '=SUM(D4:E4)' }, { f: '=Sheet2!$C4' }],
    ])
    expect(destinationActive).toBe(true)
  })

  it.each([
    'xlsx.range.copy_formulas',
    'xlsx.range.copy_formats',
    'xlsx.range.copy_without_borders',
  ] as const)(
    '%s rejects ranges with different dimensions before reading the source',
    async (operation) => {
      let sourceRead = false
      const worksheet = {
        getSheetName: () => 'Budget',
        getRange: (address: string) =>
          address === 'A1:B2'
            ? {
                getHeight: () => 2,
                getWidth: () => 2,
                getFormulas: () => {
                  sourceRead = true
                  return []
                },
              }
            : { getHeight: () => 1, getWidth: () => 1 },
      }
      const runtime = {
        univerAPI: {
          getActiveWorkbook: () => ({
            getSheets: () => [worksheet],
            setActiveSheet: () => undefined,
          }),
        },
      }

      await expect(
        executeXlsxOperation(
          {
            operation,
            arguments: {
              sourceSheet: 'Budget',
              sourceRange: 'A1:B2',
              destinationSheet: 'Budget',
              destinationRange: 'D1',
            },
          },
          {
            runtime: () => runtime,
            loadStaged: async () => undefined,
            save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
          },
        ),
      ).resolves.toMatchObject({
        handled: true,
        operationId: operation,
        ok: false,
        error: 'invalid_arguments',
        message: `${operation} requires equally shaped source and destination ranges.`,
      })
      expect(sourceRead).toBe(false)
    },
  )

  it.each([
    'xlsx.range.copy_formulas',
    'xlsx.range.copy_formats',
    'xlsx.range.copy_without_borders',
  ] as const)(
    '%s rejects operations larger than 20,000 cells before reading the source',
    async (operation) => {
      let sourceRead = false
      const worksheet = {
        getSheetName: () => 'Budget',
        getRange: () => ({
          getHeight: () => 20_001,
          getWidth: () => 1,
          getFormulas: () => {
            sourceRead = true
            return []
          },
        }),
      }
      const runtime = {
        univerAPI: {
          getActiveWorkbook: () => ({
            getSheets: () => [worksheet],
            setActiveSheet: () => undefined,
          }),
        },
      }

      await expect(
        executeXlsxOperation(
          {
            operation,
            arguments: {
              sourceSheet: 'Budget',
              sourceRange: 'A1:A20001',
              destinationSheet: 'Budget',
              destinationRange: 'C1:C20001',
            },
          },
          {
            runtime: () => runtime,
            loadStaged: async () => undefined,
            save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
          },
        ),
      ).resolves.toMatchObject({
        handled: true,
        operationId: operation,
        ok: false,
        error: 'invalid_arguments',
        message: `${operation} supports at most 20,000 cells per operation.`,
      })
      expect(sourceRead).toBe(false)
    },
  )

  it.each([
    'xlsx.range.copy_formulas',
    'xlsx.range.copy_formats',
    'xlsx.range.copy_without_borders',
  ] as const)(
    '%s rejects before reading a source range that is still streaming',
    async (operation) => {
      let sourceRead = false
      const source = {
        getRow: () => 9,
        getColumn: () => 0,
        getHeight: () => 2,
        getWidth: () => 1,
        getFormulas: () => {
          sourceRead = true
          return []
        },
      }
      const destination = {
        getRow: () => 0,
        getColumn: () => 3,
        getHeight: () => 2,
        getWidth: () => 1,
      }
      const worksheet = {
        getSheetName: () => 'Budget',
        getSheetId: () => 'sheet-budget',
        getRange: (address: string) => (address === 'A10:A11' ? source : destination),
      }
      const runtime = {
        univerAPI: {
          getActiveWorkbook: () => ({
            getSheets: () => [worksheet],
            setActiveSheet: () => undefined,
          }),
        },
      }
      const state = {
        file: { sheets: [{ id: 'sheet-budget', rowCount: 20, columnCount: 5 }] },
        loadedRanges: new Map([
          ['sheet-budget', { startRow: 0, endRow: 5, startColumn: 0, endColumn: 4 }],
        ]),
        flags: { preloadComplete: false },
      } as unknown as LazyWorkbookState

      await expect(
        executeXlsxOperation(
          {
            operation,
            arguments: {
              sourceSheet: 'Budget',
              sourceRange: 'A10:A11',
              destinationSheet: 'Budget',
              destinationRange: 'D1:D2',
            },
          },
          {
            runtime: () => runtime,
            state: () => state,
            loadStaged: async () => undefined,
            save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
          },
        ),
      ).resolves.toMatchObject({
        handled: true,
        operationId: operation,
        ok: false,
        error: 'execution_failed',
        message: `${operation} requires its source range to finish streaming.`,
      })
      expect(sourceRead).toBe(false)
    },
  )

  it.each([
    'xlsx.range.copy_formulas',
    'xlsx.range.copy_formats',
    'xlsx.range.copy_without_borders',
  ] as const)(
    '%s rejects before overwriting a destination range that is still streaming',
    async (operation) => {
      let sourceRead = false
      let destinationWritten = false
      const source = {
        getRow: () => 0,
        getColumn: () => 0,
        getHeight: () => 2,
        getWidth: () => 1,
        getFormulas: () => {
          sourceRead = true
          return [['=A1'], ['=A2']]
        },
      }
      const destination = {
        getRow: () => 9,
        getColumn: () => 3,
        getHeight: () => 2,
        getWidth: () => 1,
        setValues: () => {
          destinationWritten = true
        },
      }
      const worksheet = {
        getSheetName: () => 'Budget',
        getSheetId: () => 'sheet-budget',
        getRange: (address: string) => (address === 'A1:A2' ? source : destination),
      }
      const runtime = {
        univerAPI: {
          getActiveWorkbook: () => ({
            getSheets: () => [worksheet],
            setActiveSheet: () => undefined,
          }),
        },
      }
      const state = {
        file: { sheets: [{ id: 'sheet-budget', rowCount: 20, columnCount: 5 }] },
        loadedRanges: new Map([
          ['sheet-budget', { startRow: 0, endRow: 5, startColumn: 0, endColumn: 4 }],
        ]),
        flags: { preloadComplete: false },
      } as unknown as LazyWorkbookState

      await expect(
        executeXlsxOperation(
          {
            operation,
            arguments: {
              sourceSheet: 'Budget',
              sourceRange: 'A1:A2',
              destinationSheet: 'Budget',
              destinationRange: 'D10:D11',
            },
          },
          {
            runtime: () => runtime,
            state: () => state,
            loadStaged: async () => undefined,
            save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
          },
        ),
      ).resolves.toMatchObject({
        handled: true,
        operationId: operation,
        ok: false,
        error: 'execution_failed',
        message: `${operation} cannot overwrite a destination range that is still streaming.`,
      })
      expect(sourceRead).toBe(false)
      expect(destinationWritten).toBe(false)
    },
  )

  it.each([
    'xlsx.range.copy_formulas',
    'xlsx.range.copy_formats',
    'xlsx.range.copy_without_borders',
  ] as const)(
    '%s rejects addresses outside XLSX worksheet bounds before resolving ranges',
    async (operation) => {
      let rangeResolved = false
      const worksheet = {
        getSheetName: () => 'Budget',
        getRange: () => {
          rangeResolved = true
          return {}
        },
      }
      const runtime = {
        univerAPI: {
          getActiveWorkbook: () => ({
            getSheets: () => [worksheet],
            setActiveSheet: () => undefined,
          }),
        },
      }

      await expect(
        executeXlsxOperation(
          {
            operation,
            arguments: {
              sourceSheet: 'Budget',
              sourceRange: 'A0:A2',
              destinationSheet: 'Budget',
              destinationRange: 'XFE1:XFE2',
            },
          },
          {
            runtime: () => runtime,
            loadStaged: async () => undefined,
            save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
          },
        ),
      ).resolves.toMatchObject({
        handled: true,
        operationId: operation,
        ok: false,
        error: 'invalid_arguments',
        message: `${operation} requires valid in-sheet A1 source and destination ranges.`,
      })
      expect(rangeResolved).toBe(false)
    },
  )

  it('rejects copy-values ranges with different dimensions before writing the destination', async () => {
    let written: unknown[][] | null = null
    const worksheet = {
      getSheetName: () => 'Budget',
      getRange: (address: string) =>
        address === 'A1:B2'
          ? {
              getHeight: () => 2,
              getWidth: () => 2,
              getValues: () => [
                [1, 2],
                [3, 4],
              ],
            }
          : {
              getHeight: () => 1,
              getWidth: () => 1,
              setValues: (values: unknown[][]) => {
                written = values
              },
              activate: () => undefined,
            },
    }
    const workbook = {
      getSheets: () => [worksheet],
      setActiveSheet: () => undefined,
    }
    const runtime = { univerAPI: { getActiveWorkbook: () => workbook } }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.range.copy_values',
          arguments: {
            sourceSheet: 'Budget',
            sourceRange: 'A1:B2',
            destinationSheet: 'Budget',
            destinationRange: 'D1',
          },
        },
        {
          runtime: () => runtime,
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.range.copy_values',
      ok: false,
      error: 'invalid_arguments',
      message: 'xlsx.range.copy_values requires equally shaped source and destination ranges.',
    })
    expect(written).toBeNull()
  })

  it('rejects copy-values operations larger than 20,000 cells before reading the source', async () => {
    let sourceRead = false
    let destinationWritten = false
    const worksheet = {
      getSheetName: () => 'Budget',
      getRange: (address: string) => ({
        getHeight: () => 20_001,
        getWidth: () => 1,
        getValues: () => {
          sourceRead = true
          return []
        },
        setValues: () => {
          destinationWritten = true
        },
        activate: () => undefined,
        address,
      }),
    }
    const workbook = {
      getSheets: () => [worksheet],
      setActiveSheet: () => undefined,
    }
    const runtime = { univerAPI: { getActiveWorkbook: () => workbook } }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.range.copy_values',
          arguments: {
            sourceSheet: 'Budget',
            sourceRange: 'A1:A20001',
            destinationSheet: 'Budget',
            destinationRange: 'C1:C20001',
          },
        },
        {
          runtime: () => runtime,
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.range.copy_values',
      ok: false,
      error: 'invalid_arguments',
      message: 'xlsx.range.copy_values supports at most 20,000 cells per operation.',
    })
    expect(sourceRead).toBe(false)
    expect(destinationWritten).toBe(false)
  })

  it('rejects copy-values before reading a source range that is still streaming', async () => {
    let sourceRead = false
    let destinationWritten = false
    const source = {
      getRow: () => 9,
      getColumn: () => 0,
      getHeight: () => 2,
      getWidth: () => 2,
      getValues: () => {
        sourceRead = true
        return [
          [1, 2],
          [3, 4],
        ]
      },
    }
    const destination = {
      getRow: () => 0,
      getColumn: () => 3,
      getHeight: () => 2,
      getWidth: () => 2,
      setValues: () => {
        destinationWritten = true
      },
      activate: () => undefined,
    }
    const worksheet = {
      getSheetName: () => 'Budget',
      getSheetId: () => 'sheet-budget',
      getRange: (address: string) => (address === 'A10:B11' ? source : destination),
    }
    const workbook = {
      getSheets: () => [worksheet],
      setActiveSheet: () => undefined,
    }
    const runtime = { univerAPI: { getActiveWorkbook: () => workbook } }
    const state = {
      file: {
        sheets: [{ id: 'sheet-budget', rowCount: 20, columnCount: 5 }],
      },
      loadedRanges: new Map([
        ['sheet-budget', { startRow: 0, endRow: 5, startColumn: 0, endColumn: 4 }],
      ]),
      flags: { preloadComplete: false },
    } as unknown as LazyWorkbookState

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.range.copy_values',
          arguments: {
            sourceSheet: 'Budget',
            sourceRange: 'A10:B11',
            destinationSheet: 'Budget',
            destinationRange: 'D1:E2',
          },
        },
        {
          runtime: () => runtime,
          state: () => state,
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.range.copy_values',
      ok: false,
      error: 'execution_failed',
      message: 'xlsx.range.copy_values requires its source range to finish streaming.',
    })
    expect(sourceRead).toBe(false)
    expect(destinationWritten).toBe(false)
  })

  it('rejects copy-values before overwriting a destination range that is still streaming', async () => {
    let sourceRead = false
    let destinationWritten = false
    const source = {
      getRow: () => 0,
      getColumn: () => 0,
      getHeight: () => 2,
      getWidth: () => 2,
      getValues: () => {
        sourceRead = true
        return [
          [1, 2],
          [3, 4],
        ]
      },
    }
    const destination = {
      getRow: () => 9,
      getColumn: () => 3,
      getHeight: () => 2,
      getWidth: () => 2,
      setValues: () => {
        destinationWritten = true
      },
      activate: () => undefined,
    }
    const worksheet = {
      getSheetName: () => 'Budget',
      getSheetId: () => 'sheet-budget',
      getRange: (address: string) => (address === 'A1:B2' ? source : destination),
    }
    const workbook = {
      getSheets: () => [worksheet],
      setActiveSheet: () => undefined,
    }
    const runtime = { univerAPI: { getActiveWorkbook: () => workbook } }
    const state = {
      file: {
        sheets: [{ id: 'sheet-budget', rowCount: 20, columnCount: 5 }],
      },
      loadedRanges: new Map([
        ['sheet-budget', { startRow: 0, endRow: 5, startColumn: 0, endColumn: 4 }],
      ]),
      flags: { preloadComplete: false },
    } as unknown as LazyWorkbookState

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.range.copy_values',
          arguments: {
            sourceSheet: 'Budget',
            sourceRange: 'A1:B2',
            destinationSheet: 'Budget',
            destinationRange: 'D10:E11',
          },
        },
        {
          runtime: () => runtime,
          state: () => state,
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.range.copy_values',
      ok: false,
      error: 'execution_failed',
      message:
        'xlsx.range.copy_values cannot overwrite a destination range that is still streaming.',
    })
    expect(sourceRead).toBe(false)
    expect(destinationWritten).toBe(false)
  })

  it('rejects copy-values addresses outside the XLSX worksheet bounds before resolving ranges', async () => {
    let rangeResolved = false
    const worksheet = {
      getSheetName: () => 'Budget',
      getRange: () => {
        rangeResolved = true
        return {}
      },
    }
    const workbook = {
      getSheets: () => [worksheet],
      setActiveSheet: () => undefined,
    }
    const runtime = { univerAPI: { getActiveWorkbook: () => workbook } }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.range.copy_values',
          arguments: {
            sourceSheet: 'Budget',
            sourceRange: 'A0:B2',
            destinationSheet: 'Budget',
            destinationRange: 'XFE1:XFE2',
          },
        },
        {
          runtime: () => runtime,
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.range.copy_values',
      ok: false,
      error: 'invalid_arguments',
      message: 'xlsx.range.copy_values requires valid in-sheet A1 source and destination ranges.',
    })
    expect(rangeResolved).toBe(false)
  })

  it('sets horizontal alignment explicitly through xlsx.range.set_alignment', async () => {
    let written: unknown = null
    let activeRange: string | null = null
    const worksheet = {
      getSheetName: () => 'Budget',
      getRange: (address: string) => ({
        setValue: (value: unknown) => {
          written = value
        },
        activate: () => {
          activeRange = address
        },
      }),
    }
    const workbook = {
      getSheets: () => [worksheet],
      setActiveSheet: () => undefined,
    }
    const runtime = { univerAPI: { getActiveWorkbook: () => workbook } }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.range.set_alignment',
          arguments: {
            sheet: 'Budget',
            range: 'b2:c3',
            alignment: { horizontal: 'justify' },
            fields: ['horizontal'],
          },
        },
        {
          runtime: () => runtime,
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.range.set_alignment',
      ok: true,
      output: { sheet: 'Budget', range: 'B2:C3', fields: ['horizontal'] },
    })
    expect(written).toEqual({ s: { ht: HorizontalAlign.JUSTIFIED } })
    expect(activeRange).toBe('b2:c3')
  })

  it('sets vertical alignment explicitly through xlsx.range.set_alignment', async () => {
    let written: unknown = null
    const worksheet = {
      getSheetName: () => 'Budget',
      getRange: () => ({
        setValue: (value: unknown) => {
          written = value
        },
        activate: () => undefined,
      }),
    }
    const workbook = {
      getSheets: () => [worksheet],
      setActiveSheet: () => undefined,
    }
    const runtime = { univerAPI: { getActiveWorkbook: () => workbook } }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.range.set_alignment',
          arguments: {
            sheet: 'Budget',
            range: 'A1',
            alignment: { vertical: 'middle' },
            fields: ['vertical'],
          },
        },
        {
          runtime: () => runtime,
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.range.set_alignment',
      ok: true,
      output: { sheet: 'Budget', range: 'A1', fields: ['vertical'] },
    })
    expect(written).toEqual({ s: { vt: VerticalAlign.MIDDLE } })
  })

  it('sets wrapping explicitly through xlsx.range.set_alignment', async () => {
    let written: unknown = null
    const worksheet = {
      getSheetName: () => 'Budget',
      getRange: () => ({
        setValue: (value: unknown) => {
          written = value
        },
        activate: () => undefined,
      }),
    }
    const workbook = {
      getSheets: () => [worksheet],
      setActiveSheet: () => undefined,
    }
    const runtime = { univerAPI: { getActiveWorkbook: () => workbook } }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.range.set_alignment',
          arguments: {
            sheet: 'Budget',
            range: 'A1',
            alignment: { wrap: true },
            fields: ['wrap'],
          },
        },
        {
          runtime: () => runtime,
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.range.set_alignment',
      ok: true,
      output: { sheet: 'Budget', range: 'A1', fields: ['wrap'] },
    })
    expect(written).toEqual({ s: { tb: WrapStrategy.WRAP } })
  })

  it('sets bounded indentation explicitly through xlsx.range.set_alignment', async () => {
    let written: unknown = null
    const worksheet = {
      getSheetName: () => 'Budget',
      getRange: () => ({
        setValue: (value: unknown) => {
          written = value
        },
        activate: () => undefined,
      }),
    }
    const workbook = {
      getSheets: () => [worksheet],
      setActiveSheet: () => undefined,
    }
    const runtime = { univerAPI: { getActiveWorkbook: () => workbook } }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.range.set_alignment',
          arguments: {
            sheet: 'Budget',
            range: 'A1',
            alignment: { indent: 3 },
            fields: ['indent'],
          },
        },
        {
          runtime: () => runtime,
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.range.set_alignment',
      ok: true,
      output: { sheet: 'Budget', range: 'A1', fields: ['indent'] },
    })
    expect(written).toEqual({ s: { pd: { l: 24 } } })
  })

  it('sets stacked text rotation explicitly through xlsx.range.set_alignment', async () => {
    let written: unknown = null
    const worksheet = {
      getSheetName: () => 'Budget',
      getRange: () => ({
        setValue: (value: unknown) => {
          written = value
        },
        activate: () => undefined,
      }),
    }
    const workbook = {
      getSheets: () => [worksheet],
      setActiveSheet: () => undefined,
    }
    const runtime = { univerAPI: { getActiveWorkbook: () => workbook } }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.range.set_alignment',
          arguments: {
            sheet: 'Budget',
            range: 'A1',
            alignment: { rotation: { mode: 'stacked' } },
            fields: ['rotation'],
          },
        },
        {
          runtime: () => runtime,
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.range.set_alignment',
      ok: true,
      output: { sheet: 'Budget', range: 'A1', fields: ['rotation'] },
    })
    expect(written).toEqual({ s: { tr: { a: 0, v: BooleanNumber.TRUE } } })
  })

  it('sets bounded angle rotation explicitly through xlsx.range.set_alignment', async () => {
    let written: unknown = null
    const worksheet = {
      getSheetName: () => 'Budget',
      getRange: () => ({
        setValue: (value: unknown) => {
          written = value
        },
        activate: () => undefined,
      }),
    }
    const workbook = {
      getSheets: () => [worksheet],
      setActiveSheet: () => undefined,
    }
    const runtime = { univerAPI: { getActiveWorkbook: () => workbook } }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.range.set_alignment',
          arguments: {
            sheet: 'Budget',
            range: 'A1',
            alignment: { rotation: { mode: 'angle', degrees: -45 } },
            fields: ['rotation'],
          },
        },
        {
          runtime: () => runtime,
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.range.set_alignment',
      ok: true,
      output: { sheet: 'Budget', range: 'A1', fields: ['rotation'] },
    })
    expect(written).toEqual({ s: { tr: { a: -45 } } })
  })

  it('clears text rotation explicitly through xlsx.range.set_alignment', async () => {
    let written: unknown = null
    const worksheet = {
      getSheetName: () => 'Budget',
      getRange: () => ({
        setValue: (value: unknown) => {
          written = value
        },
        activate: () => undefined,
      }),
    }
    const workbook = {
      getSheets: () => [worksheet],
      setActiveSheet: () => undefined,
    }
    const runtime = { univerAPI: { getActiveWorkbook: () => workbook } }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.range.set_alignment',
          arguments: {
            sheet: 'Budget',
            range: 'A1',
            alignment: { rotation: { mode: 'none' } },
            fields: ['rotation'],
          },
        },
        {
          runtime: () => runtime,
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.range.set_alignment',
      ok: true,
      output: { sheet: 'Budget', range: 'A1', fields: ['rotation'] },
    })
    expect(written).toEqual({ s: { tr: null } })
  })

  it('sets multiple alignment fields through one atomic range mutation', async () => {
    const writes: unknown[] = []
    const worksheet = {
      getSheetName: () => 'Budget',
      getRange: () => ({
        setValue: (value: unknown) => {
          writes.push(value)
        },
        activate: () => undefined,
      }),
    }
    const workbook = {
      getSheets: () => [worksheet],
      setActiveSheet: () => undefined,
    }
    const runtime = { univerAPI: { getActiveWorkbook: () => workbook } }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.range.set_alignment',
          arguments: {
            sheet: 'Budget',
            range: 'A1:B2',
            alignment: {
              horizontal: 'center',
              vertical: 'bottom',
              wrap: false,
              indent: 0,
              rotation: { mode: 'angle', degrees: 45 },
            },
            fields: ['horizontal', 'vertical', 'wrap', 'indent', 'rotation'],
          },
        },
        {
          runtime: () => runtime,
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.range.set_alignment',
      ok: true,
      output: {
        sheet: 'Budget',
        range: 'A1:B2',
        fields: ['horizontal', 'vertical', 'wrap', 'indent', 'rotation'],
      },
    })
    expect(writes).toEqual([
      {
        s: {
          ht: HorizontalAlign.CENTER,
          vt: VerticalAlign.BOTTOM,
          tb: null,
          pd: null,
          tr: { a: 45 },
        },
      },
    ])
  })

  it('sets a font family through xlsx.range.set_font', async () => {
    let written: unknown = null
    const worksheet = {
      getSheetName: () => 'Budget',
      getRange: () => ({
        setValue: (value: unknown) => {
          written = value
        },
        activate: () => undefined,
      }),
    }
    const workbook = {
      getSheets: () => [worksheet],
      setActiveSheet: () => undefined,
    }
    const runtime = { univerAPI: { getActiveWorkbook: () => workbook } }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.range.set_font',
          arguments: {
            sheet: 'Budget',
            range: 'b2:c3',
            font: { family: 'Aptos' },
            fields: ['family'],
          },
        },
        {
          runtime: () => runtime,
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.range.set_font',
      ok: true,
      output: { sheet: 'Budget', range: 'B2:C3', fields: ['family'] },
    })
    expect(written).toEqual({ s: { ff: 'Aptos' } })
  })

  it('sets multiple font fields through one atomic range mutation', async () => {
    const writes: unknown[] = []
    const worksheet = {
      getSheetName: () => 'Budget',
      getRange: () => ({
        setValue: (value: unknown) => writes.push(value),
        activate: () => undefined,
      }),
    }
    const workbook = {
      getSheets: () => [worksheet],
      setActiveSheet: () => undefined,
    }
    const runtime = { univerAPI: { getActiveWorkbook: () => workbook } }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.range.set_font',
          arguments: {
            sheet: 'Budget',
            range: 'A1:B2',
            font: { family: null, size: 14, color: '#44546A' },
            fields: ['family', 'size', 'color'],
          },
        },
        {
          runtime: () => runtime,
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.range.set_font',
      ok: true,
      output: { sheet: 'Budget', range: 'A1:B2', fields: ['family', 'size', 'color'] },
    })
    expect(writes).toEqual([{ s: { ff: null, fs: 14, cl: { rgb: '#44546A' } } }])
  })

  it('rejects a mismatched xlsx.range.set_font field mask without mutating', async () => {
    let mutated = false
    const worksheet = {
      getSheetName: () => 'Budget',
      getRange: () => ({
        setValue: () => {
          mutated = true
        },
        activate: () => undefined,
      }),
    }
    const workbook = {
      getSheets: () => [worksheet],
      setActiveSheet: () => undefined,
    }
    const runtime = { univerAPI: { getActiveWorkbook: () => workbook } }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.range.set_font',
          arguments: {
            sheet: 'Budget',
            range: 'A1',
            font: { family: 'Aptos', color: '#44546A' },
            fields: ['family'],
          },
        },
        {
          runtime: () => runtime,
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toMatchObject({
      handled: true,
      operationId: 'xlsx.range.set_font',
      ok: false,
      error: 'invalid_arguments',
    })
    expect(mutated).toBe(false)
  })

  it('rejects an invalid xlsx.range.set_font color before mutation', async () => {
    let mutated = false
    const worksheet = {
      getSheetName: () => 'Budget',
      getRange: () => ({
        setValue: () => {
          mutated = true
        },
        activate: () => undefined,
      }),
    }
    const workbook = {
      getSheets: () => [worksheet],
      setActiveSheet: () => undefined,
    }
    const runtime = { univerAPI: { getActiveWorkbook: () => workbook } }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.range.set_font',
          arguments: {
            sheet: 'Budget',
            range: 'A1',
            font: { color: 'red' },
            fields: ['color'],
          },
        },
        {
          runtime: () => runtime,
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toMatchObject({
      handled: true,
      operationId: 'xlsx.range.set_font',
      ok: false,
      error: 'invalid_arguments',
    })
    expect(mutated).toBe(false)
  })

  it('sets and clears a fill through xlsx.range.set_fill', async () => {
    const writes: unknown[] = []
    const worksheet = {
      getSheetName: () => 'Budget',
      getRange: () => ({
        setValue: (value: unknown) => writes.push(value),
        activate: () => undefined,
      }),
    }
    const workbook = {
      getSheets: () => [worksheet],
      setActiveSheet: () => undefined,
    }
    const runtime = { univerAPI: { getActiveWorkbook: () => workbook } }
    const services = {
      runtime: () => runtime,
      loadStaged: async () => undefined,
      save: async () => ({ ok: true as const, fileName: 'budget.xlsx' }),
    }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.range.set_fill',
          arguments: { sheet: 'Budget', range: 'b2:c3', color: '#DDEBF7' },
        },
        services,
      ),
    ).resolves.toMatchObject({ handled: true, operationId: 'xlsx.range.set_fill', ok: true })
    await executeXlsxOperation(
      {
        operation: 'xlsx.range.set_fill',
        arguments: { sheet: 'Budget', range: 'b2:c3', color: null },
      },
      services,
    )
    expect(writes).toEqual([{ s: { bg: { rgb: '#DDEBF7' } } }, { s: { bg: null } }])
  })

  it('sets a border through xlsx.range.set_border using the native range route', async () => {
    const borders: unknown[][] = []
    const worksheet = {
      getSheetName: () => 'Budget',
      getRange: () => ({
        setBorder: (...values: unknown[]) => borders.push(values),
        activate: () => undefined,
      }),
    }
    const workbook = {
      getSheets: () => [worksheet],
      setActiveSheet: () => undefined,
    }
    const runtime = { univerAPI: { getActiveWorkbook: () => workbook } }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.range.set_border',
          arguments: {
            sheet: 'Budget',
            range: 'a1:b2',
            border: { preset: 'outer', lineStyle: 'medium', color: '#4472C4' },
          },
        },
        {
          runtime: () => runtime,
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.range.set_border',
      ok: true,
      output: { sheet: 'Budget', range: 'A1:B2', preset: 'outer' },
    })
    expect(borders).toEqual([[BorderType.OUTSIDE, BorderStyleTypes.MEDIUM, '#4472C4']])
  })

  it('applies a named cell style through one xlsx.range.apply_cell_style mutation', async () => {
    const writes: unknown[] = []
    const worksheet = {
      getSheetName: () => 'Budget',
      getRange: () => ({
        setValue: (value: unknown) => writes.push(value),
        activate: () => undefined,
      }),
    }
    const workbook = {
      getSheets: () => [worksheet],
      setActiveSheet: () => undefined,
    }
    const runtime = { univerAPI: { getActiveWorkbook: () => workbook } }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.range.apply_cell_style',
          arguments: { sheet: 'Budget', range: 'a1:b2', preset: 'input' },
        },
        {
          runtime: () => runtime,
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.range.apply_cell_style',
      ok: true,
      output: { sheet: 'Budget', range: 'A1:B2', preset: 'input' },
    })
    const edge = { s: BorderStyleTypes.THIN, cl: { rgb: '#7F7F7F' } }
    expect(writes).toEqual([
      {
        s: {
          bg: { rgb: '#FFCC99' },
          cl: { rgb: '#3F3F76' },
          bd: { t: edge, r: edge, b: edge, l: edge },
        },
      },
    ])
  })

  it('sets an exact number format through xlsx.range.set_number_format', async () => {
    let numberFormat: string | null = null
    let activeRange: string | null = null
    const worksheet = {
      getSheetName: () => 'Budget',
      getRange: (address: string) => ({
        setNumberFormat: (pattern: string) => {
          numberFormat = pattern
        },
        activate: () => {
          activeRange = address
        },
      }),
    }
    const workbook = {
      getSheets: () => [worksheet],
      setActiveSheet: () => undefined,
    }
    const runtime = { univerAPI: { getActiveWorkbook: () => workbook } }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.range.set_number_format',
          arguments: { sheet: 'Budget', range: 'b2:c3', pattern: 'h:mm:ss AM/PM' },
        },
        {
          runtime: () => runtime,
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.range.set_number_format',
      ok: true,
      output: { sheet: 'Budget', range: 'B2:C3', pattern: 'h:mm:ss AM/PM' },
    })
    expect(numberFormat).toBe('h:mm:ss AM/PM')
    expect(activeRange).toBe('b2:c3')
  })

  it.each(['', 'x'.repeat(256)])(
    'rejects an invalid number format pattern before mutating the range',
    async (pattern) => {
      let mutated = false
      const worksheet = {
        getSheetName: () => 'Budget',
        getRange: () => ({
          setNumberFormat: () => {
            mutated = true
          },
          activate: () => {
            mutated = true
          },
        }),
      }
      const workbook = {
        getSheets: () => [worksheet],
        setActiveSheet: () => undefined,
      }
      const runtime = { univerAPI: { getActiveWorkbook: () => workbook } }

      await expect(
        executeXlsxOperation(
          {
            operation: 'xlsx.range.set_number_format',
            arguments: { sheet: 'Budget', range: 'A1', pattern },
          },
          {
            runtime: () => runtime,
            loadStaged: async () => undefined,
            save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
          },
        ),
      ).resolves.toMatchObject({
        handled: true,
        operationId: 'xlsx.range.set_number_format',
        ok: false,
        error: 'invalid_arguments',
      })
      expect(mutated).toBe(false)
    },
  )

  it('merges each row through xlsx.range.merge across mode', async () => {
    const mutations: string[] = []
    const worksheet = {
      getSheetName: () => 'Budget',
      getRange: () => ({
        mergeAcross: () => mutations.push('across'),
        activate: () => mutations.push('activate'),
      }),
    }
    const workbook = {
      getSheets: () => [worksheet],
      setActiveSheet: () => undefined,
    }
    const runtime = { univerAPI: { getActiveWorkbook: () => workbook } }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.range.merge',
          arguments: { sheet: 'Budget', range: 'a1:c2', mode: 'across' },
        },
        {
          runtime: () => runtime,
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.range.merge',
      ok: true,
      output: { sheet: 'Budget', range: 'A1:C2', mode: 'across' },
    })
    expect(mutations).toEqual(['across', 'activate'])
  })

  it('merges one range through xlsx.range.merge cells mode', async () => {
    let merged = false
    const worksheet = {
      getSheetName: () => 'Budget',
      getRange: () => ({
        merge: () => {
          merged = true
        },
        activate: () => undefined,
      }),
    }
    const workbook = {
      getSheets: () => [worksheet],
      setActiveSheet: () => undefined,
    }
    const runtime = { univerAPI: { getActiveWorkbook: () => workbook } }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.range.merge',
          arguments: { sheet: 'Budget', range: 'A1:C2', mode: 'cells' },
        },
        {
          runtime: () => runtime,
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toMatchObject({ handled: true, operationId: 'xlsx.range.merge', ok: true })
    expect(merged).toBe(true)
  })

  it('unmerges a range through xlsx.range.merge unmerge mode', async () => {
    let unmerged = false
    const worksheet = {
      getSheetName: () => 'Budget',
      getRange: () => ({
        breakApart: () => {
          unmerged = true
        },
        activate: () => undefined,
      }),
    }
    const workbook = {
      getSheets: () => [worksheet],
      setActiveSheet: () => undefined,
    }
    const runtime = { univerAPI: { getActiveWorkbook: () => workbook } }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.range.merge',
          arguments: { sheet: 'Budget', range: 'A1:C2', mode: 'unmerge' },
        },
        {
          runtime: () => runtime,
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toMatchObject({ handled: true, operationId: 'xlsx.range.merge', ok: true })
    expect(unmerged).toBe(true)
  })

  it('preserves merge-and-center semantics through xlsx.range.merge center mode', async () => {
    const mutations: string[] = []
    const worksheet = {
      getSheetName: () => 'Budget',
      getRange: () => ({
        merge: () => mutations.push('merge'),
        setHorizontalAlignment: (alignment: string) => mutations.push(alignment),
        activate: () => mutations.push('activate'),
      }),
    }
    const workbook = {
      getSheets: () => [worksheet],
      setActiveSheet: () => undefined,
    }
    const runtime = { univerAPI: { getActiveWorkbook: () => workbook } }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.range.merge',
          arguments: { sheet: 'Budget', range: 'A1:C2', mode: 'center' },
        },
        {
          runtime: () => runtime,
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.range.merge',
      ok: true,
      output: { sheet: 'Budget', range: 'A1:C2', mode: 'center' },
    })
    expect(mutations).toEqual(['merge', 'center', 'activate'])
  })

  it('clears range contents through xlsx.range.clear contents scope', async () => {
    const mutations: string[] = []
    const worksheet = {
      getSheetName: () => 'Budget',
      getRange: () => ({
        clearContent: () => mutations.push('contents'),
        activate: () => mutations.push('activate'),
      }),
    }
    const workbook = {
      getSheets: () => [worksheet],
      setActiveSheet: () => undefined,
    }
    const runtime = { univerAPI: { getActiveWorkbook: () => workbook } }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.range.clear',
          arguments: { sheet: 'Budget', range: 'b2:c3', scope: 'contents' },
        },
        {
          runtime: () => runtime,
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.range.clear',
      ok: true,
      output: { sheet: 'Budget', range: 'B2:C3', scope: 'contents' },
    })
    expect(mutations).toEqual(['contents', 'activate'])
  })

  it('clears range formatting through xlsx.range.clear formats scope', async () => {
    const mutations: string[] = []
    const worksheet = {
      getSheetName: () => 'Budget',
      getRange: () => ({
        clearFormat: () => mutations.push('formats'),
        activate: () => mutations.push('activate'),
      }),
    }
    const workbook = {
      getSheets: () => [worksheet],
      setActiveSheet: () => undefined,
    }
    const runtime = { univerAPI: { getActiveWorkbook: () => workbook } }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.range.clear',
          arguments: { sheet: 'Budget', range: 'b2:c3', scope: 'formats' },
        },
        {
          runtime: () => runtime,
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.range.clear',
      ok: true,
      output: { sheet: 'Budget', range: 'B2:C3', scope: 'formats' },
    })
    expect(mutations).toEqual(['formats', 'activate'])
  })

  it('clears all range data and formatting through xlsx.range.clear all scope', async () => {
    const mutations: string[] = []
    const worksheet = {
      getSheetName: () => 'Budget',
      getRange: () => ({
        clear: () => mutations.push('all'),
        activate: () => mutations.push('activate'),
      }),
    }
    const workbook = {
      getSheets: () => [worksheet],
      setActiveSheet: () => undefined,
    }
    const runtime = { univerAPI: { getActiveWorkbook: () => workbook } }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.range.clear',
          arguments: { sheet: 'Budget', range: 'b2:c3', scope: 'all' },
        },
        {
          runtime: () => runtime,
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.range.clear',
      ok: true,
      output: { sheet: 'Budget', range: 'B2:C3', scope: 'all' },
    })
    expect(mutations).toEqual(['all', 'activate'])
  })

  it('fills a selected range downward through xlsx.range.fill', async () => {
    const mutations: string[] = []
    const worksheet = {
      getSheetName: () => 'Budget',
      getRange: () => ({
        getHeight: () => 3,
        activate: () => mutations.push('activate'),
      }),
    }
    const workbook = {
      getSheets: () => [worksheet],
      setActiveSheet: () => undefined,
    }
    const runtime = {
      univerAPI: {
        getActiveWorkbook: () => workbook,
        executeCommand: async (command: string) => {
          mutations.push(command)
          return true
        },
      },
    }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.range.fill',
          arguments: { sheet: 'Budget', range: 'a1:a3', direction: 'down' },
        },
        {
          runtime: () => runtime,
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.range.fill',
      ok: true,
      output: { sheet: 'Budget', range: 'A1:A3', direction: 'down' },
    })
    expect(mutations).toEqual(['activate', 'sheet.command.copy-down'])
  })

  it('fills a selected range rightward through xlsx.range.fill', async () => {
    const mutations: string[] = []
    const worksheet = {
      getSheetName: () => 'Budget',
      getRange: () => ({
        getWidth: () => 3,
        activate: () => mutations.push('activate'),
      }),
    }
    const workbook = {
      getSheets: () => [worksheet],
      setActiveSheet: () => undefined,
    }
    const runtime = {
      univerAPI: {
        getActiveWorkbook: () => workbook,
        executeCommand: async (command: string) => {
          mutations.push(command)
          return true
        },
      },
    }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.range.fill',
          arguments: { sheet: 'Budget', range: 'a1:c1', direction: 'right' },
        },
        {
          runtime: () => runtime,
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.range.fill',
      ok: true,
      output: { sheet: 'Budget', range: 'A1:C1', direction: 'right' },
    })
    expect(mutations).toEqual(['activate', 'sheet.command.copy-right'])
  })

  it('rejects a downward fill without a destination row before executing a command', async () => {
    let mutated = false
    const worksheet = {
      getSheetName: () => 'Budget',
      getRange: () => ({
        getHeight: () => 1,
        activate: () => {
          mutated = true
        },
      }),
    }
    const workbook = {
      getSheets: () => [worksheet],
      setActiveSheet: () => undefined,
    }
    const runtime = {
      univerAPI: {
        getActiveWorkbook: () => workbook,
        executeCommand: async () => {
          mutated = true
          return true
        },
      },
    }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.range.fill',
          arguments: { sheet: 'Budget', range: 'A1', direction: 'down' },
        },
        {
          runtime: () => runtime,
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toMatchObject({
      handled: true,
      operationId: 'xlsx.range.fill',
      ok: false,
      error: 'invalid_arguments',
    })
    expect(mutated).toBe(false)
  })

  it('sorts an explicit range ascending by its first column through xlsx.range.sort', async () => {
    const mutations: unknown[] = []
    const targetRange = {
      getRow: () => 1,
      getColumn: () => 2,
      getHeight: () => 3,
      getWidth: () => 2,
      activate: () => mutations.push('activate'),
    }
    const worksheet = {
      getSheetName: () => 'Budget',
      getSheetId: () => 'sheet-budget',
      getRange: () => targetRange,
    }
    const workbook = {
      getId: () => 'workbook-budget',
      getSheets: () => [worksheet],
      setActiveSheet: () => undefined,
    }
    const runtime = {
      univerAPI: {
        getActiveWorkbook: () => workbook,
        executeCommand: async (command: string, parameters: unknown) => {
          mutations.push([command, parameters])
          return true
        },
      },
    }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.range.sort',
          arguments: { sheet: 'Budget', range: 'c2:d4', direction: 'asc' },
        },
        {
          runtime: () => runtime,
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.range.sort',
      ok: true,
      output: { sheet: 'Budget', range: 'C2:D4', direction: 'asc' },
    })
    expect(mutations).toEqual([
      [
        'sheet.command.sort-range',
        {
          unitId: 'workbook-budget',
          subUnitId: 'sheet-budget',
          range: { startRow: 1, endRow: 3, startColumn: 2, endColumn: 3 },
          orderRules: [{ type: 'asc', colIndex: 2 }],
          hasTitle: false,
        },
      ],
      'activate',
    ])
  })

  it('rejects xlsx.range.sort when the target has no second row', async () => {
    let mutated = false
    const worksheet = {
      getSheetName: () => 'Budget',
      getSheetId: () => 'sheet-budget',
      getRange: () => ({
        getRow: () => 0,
        getColumn: () => 0,
        getHeight: () => 1,
        getWidth: () => 2,
        activate: () => {
          mutated = true
        },
      }),
    }
    const workbook = {
      getId: () => 'workbook-budget',
      getSheets: () => [worksheet],
      setActiveSheet: () => undefined,
    }
    const runtime = {
      univerAPI: {
        getActiveWorkbook: () => workbook,
        executeCommand: async () => {
          mutated = true
          return true
        },
      },
    }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.range.sort',
          arguments: { sheet: 'Budget', range: 'A1:B1', direction: 'desc' },
        },
        {
          runtime: () => runtime,
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toMatchObject({
      handled: true,
      operationId: 'xlsx.range.sort',
      ok: false,
      error: 'invalid_arguments',
    })
    expect(mutated).toBe(false)
  })

  it('sorts an explicit range by ordered column keys through xlsx.range.sort_custom', async () => {
    const mutations: unknown[] = []
    const targetRange = {
      getRow: () => 1,
      getColumn: () => 2,
      getHeight: () => 4,
      getWidth: () => 3,
      activate: () => mutations.push('activate'),
    }
    const worksheet = {
      getSheetName: () => 'Budget',
      getSheetId: () => 'sheet-budget',
      getRange: () => targetRange,
    }
    const workbook = {
      getId: () => 'workbook-budget',
      getSheets: () => [worksheet],
      setActiveSheet: () => undefined,
    }
    const runtime = {
      univerAPI: {
        getActiveWorkbook: () => workbook,
        executeCommand: async (command: string, parameters: unknown) => {
          mutations.push([command, parameters])
          return true
        },
      },
    }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.range.sort_custom',
          arguments: {
            sheet: 'Budget',
            range: 'c2:e5',
            keys: [
              { column: 'e', direction: 'desc' },
              { column: 'c', direction: 'asc' },
            ],
            hasHeader: true,
          },
        },
        {
          runtime: () => runtime,
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.range.sort_custom',
      ok: true,
      output: {
        sheet: 'Budget',
        range: 'C2:E5',
        keys: [
          { column: 'E', direction: 'desc' },
          { column: 'C', direction: 'asc' },
        ],
        hasHeader: true,
      },
    })
    expect(mutations).toEqual([
      [
        'sheet.command.sort-range',
        {
          unitId: 'workbook-budget',
          subUnitId: 'sheet-budget',
          range: { startRow: 1, endRow: 4, startColumn: 2, endColumn: 4 },
          orderRules: [
            { type: 'desc', colIndex: 4 },
            { type: 'asc', colIndex: 2 },
          ],
          hasTitle: true,
        },
      ],
      'activate',
    ])
  })

  it('rejects an xlsx.range.sort_custom key outside the target range before mutating', async () => {
    let mutated = false
    const worksheet = {
      getSheetName: () => 'Budget',
      getSheetId: () => 'sheet-budget',
      getRange: () => ({
        getRow: () => 1,
        getColumn: () => 2,
        getHeight: () => 4,
        getWidth: () => 3,
        activate: () => {
          mutated = true
        },
      }),
    }
    const workbook = {
      getId: () => 'workbook-budget',
      getSheets: () => [worksheet],
      setActiveSheet: () => undefined,
    }
    const runtime = {
      univerAPI: {
        getActiveWorkbook: () => workbook,
        executeCommand: async () => {
          mutated = true
          return true
        },
      },
    }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.range.sort_custom',
          arguments: {
            sheet: 'Budget',
            range: 'C2:E5',
            keys: [{ column: 'F', direction: 'asc' }],
            hasHeader: false,
          },
        },
        {
          runtime: () => runtime,
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toMatchObject({
      handled: true,
      operationId: 'xlsx.range.sort_custom',
      ok: false,
      error: 'invalid_arguments',
    })
    expect(mutated).toBe(false)
  })

  it('rejects duplicate xlsx.range.sort_custom keys before mutating', async () => {
    let mutated = false
    const worksheet = {
      getSheetName: () => 'Budget',
      getSheetId: () => 'sheet-budget',
      getRange: () => ({
        getRow: () => 0,
        getColumn: () => 0,
        getHeight: () => 3,
        getWidth: () => 2,
        activate: () => {
          mutated = true
        },
      }),
    }
    const workbook = {
      getId: () => 'workbook-budget',
      getSheets: () => [worksheet],
      setActiveSheet: () => undefined,
    }
    const runtime = {
      univerAPI: {
        getActiveWorkbook: () => workbook,
        executeCommand: async () => {
          mutated = true
          return true
        },
      },
    }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.range.sort_custom',
          arguments: {
            sheet: 'Budget',
            range: 'A1:B3',
            keys: [
              { column: 'A', direction: 'asc' },
              { column: 'a', direction: 'desc' },
            ],
            hasHeader: false,
          },
        },
        {
          runtime: () => runtime,
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toMatchObject({
      handled: true,
      operationId: 'xlsx.range.sort_custom',
      ok: false,
      error: 'invalid_arguments',
    })
    expect(mutated).toBe(false)
  })

  it('rejects xlsx.range.sort_custom when a header leaves no sortable row', async () => {
    let mutated = false
    const worksheet = {
      getSheetName: () => 'Budget',
      getSheetId: () => 'sheet-budget',
      getRange: () => ({
        getRow: () => 0,
        getColumn: () => 0,
        getHeight: () => 1,
        getWidth: () => 2,
        activate: () => {
          mutated = true
        },
      }),
    }
    const workbook = {
      getId: () => 'workbook-budget',
      getSheets: () => [worksheet],
      setActiveSheet: () => undefined,
    }
    const runtime = {
      univerAPI: {
        getActiveWorkbook: () => workbook,
        executeCommand: async () => {
          mutated = true
          return true
        },
      },
    }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.range.sort_custom',
          arguments: {
            sheet: 'Budget',
            range: 'A1:B1',
            keys: [{ column: 'A', direction: 'asc' }],
            hasHeader: true,
          },
        },
        {
          runtime: () => runtime,
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toMatchObject({
      handled: true,
      operationId: 'xlsx.range.sort_custom',
      ok: false,
      error: 'invalid_arguments',
    })
    expect(mutated).toBe(false)
  })

  it('removes duplicate rows through xlsx.range.remove_duplicates while preserving the header', async () => {
    const writes: unknown[] = []
    const targetRange = {
      getRow: () => 4,
      getColumn: () => 1,
      getHeight: () => 4,
      getWidth: () => 2,
      getValues: () => [
        ['Name', 'Score'],
        ['Alpha', 1],
        ['ALPHA', 1],
        ['Beta', 2],
      ],
      activate: () => writes.push('activate'),
    }
    const worksheet = {
      getSheetName: () => 'Budget',
      getSheetId: () => 'sheet-budget',
      getRange: (...arguments_: unknown[]) => {
        if (typeof arguments_[0] === 'string') return targetRange
        return {
          setValues: (values: unknown) => writes.push([arguments_, values]),
        }
      },
    }
    const workbook = {
      getSheets: () => [worksheet],
      setActiveSheet: () => undefined,
    }
    const runtime = { univerAPI: { getActiveWorkbook: () => workbook } }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.range.remove_duplicates',
          arguments: { sheet: 'Budget', range: 'b5:c8', hasHeader: true },
        },
        {
          runtime: () => runtime,
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.range.remove_duplicates',
      ok: true,
      output: { sheet: 'Budget', range: 'B5:C8', removed: 1 },
    })
    expect(writes).toEqual([
      [[6, 1, 1, 2], [['Beta', 2]]],
      [[7, 1, 1, 2], [[null, null]]],
      'activate',
    ])
  })

  it('rejects xlsx.range.remove_duplicates when the target has no second row', async () => {
    let read = false
    const worksheet = {
      getSheetName: () => 'Budget',
      getSheetId: () => 'sheet-budget',
      getRange: () => ({
        getHeight: () => 1,
        getValues: () => {
          read = true
          return [['Only row']]
        },
        getWidth: () => 1,
        getRow: () => 0,
        getColumn: () => 0,
        activate: () => {
          read = true
        },
      }),
    }
    const workbook = {
      getSheets: () => [worksheet],
      setActiveSheet: () => undefined,
    }
    const runtime = { univerAPI: { getActiveWorkbook: () => workbook } }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.range.remove_duplicates',
          arguments: { sheet: 'Budget', range: 'A1', hasHeader: false },
        },
        {
          runtime: () => runtime,
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toMatchObject({
      handled: true,
      operationId: 'xlsx.range.remove_duplicates',
      ok: false,
      error: 'invalid_arguments',
    })
    expect(read).toBe(false)
  })

  it('rejects xlsx.range.remove_duplicates before reading partially streamed sheet data', async () => {
    let read = false
    const worksheet = {
      getSheetName: () => 'Budget',
      getSheetId: () => 'sheet-budget',
      getRange: () => ({
        getHeight: () => 2,
        getValues: () => {
          read = true
          return [['A'], ['A']]
        },
      }),
    }
    const workbook = {
      getSheets: () => [worksheet],
      setActiveSheet: () => undefined,
    }
    const runtime = { univerAPI: { getActiveWorkbook: () => workbook } }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.range.remove_duplicates',
          arguments: { sheet: 'Budget', range: 'A1:A2', hasHeader: false },
        },
        {
          runtime: () => runtime,
          isSheetDataComplete: () => false,
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.range.remove_duplicates',
      ok: false,
      error: 'execution_failed',
      message: 'xlsx.range.remove_duplicates requires fully loaded sheet data.',
    })
    expect(read).toBe(false)
  })

  it('inserts one SUM formula below each selected column through xlsx.formula.insert_aggregate', async () => {
    const formulas = new Map<string, string>()
    const sourceRange = {
      getRow: () => 1,
      getColumn: () => 1,
      getHeight: () => 3,
      getWidth: () => 2,
    }
    const worksheet = {
      getSheetName: () => 'Budget',
      getSheetId: () => 'sheet-budget',
      getRange: (...arguments_: unknown[]) => {
        if (typeof arguments_[0] === 'string') return sourceRange
        const [row, column] = arguments_ as [number, number]
        return {
          setFormula: (formula: string) => formulas.set(`${row}:${column}`, formula),
        }
      },
    }
    const workbook = {
      getSheets: () => [worksheet],
      setActiveSheet: () => undefined,
    }
    const runtime = { univerAPI: { getActiveWorkbook: () => workbook } }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.formula.insert_aggregate',
          arguments: { sheet: 'Budget', range: 'b2:c4', function: 'SUM' },
        },
        {
          runtime: () => runtime,
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.formula.insert_aggregate',
      ok: true,
      output: {
        sheet: 'Budget',
        range: 'B2:C4',
        targetRange: 'B5:C5',
        function: 'SUM',
        inserted: 2,
      },
    })
    expect([...formulas.entries()]).toEqual([
      ['4:1', '=SUM(B2:B4)'],
      ['4:2', '=SUM(C2:C4)'],
    ])
  })

  it('supports the remaining retained aggregate functions through the same operation', async () => {
    const formulas: string[] = []
    const sourceRange = {
      getRow: () => 0,
      getColumn: () => 0,
      getHeight: () => 2,
      getWidth: () => 1,
    }
    const worksheet = {
      getSheetName: () => 'Budget',
      getSheetId: () => 'sheet-budget',
      getRange: (...arguments_: unknown[]) =>
        typeof arguments_[0] === 'string'
          ? sourceRange
          : { setFormula: (formula: string) => formulas.push(formula) },
    }
    const workbook = {
      getSheets: () => [worksheet],
      setActiveSheet: () => undefined,
    }
    const runtime = { univerAPI: { getActiveWorkbook: () => workbook } }

    for (const aggregate of ['AVERAGE', 'COUNT', 'MAX', 'MIN'] as const) {
      await expect(
        executeXlsxOperation(
          {
            operation: 'xlsx.formula.insert_aggregate',
            arguments: { sheet: 'Budget', range: 'A1:A2', function: aggregate },
          },
          {
            runtime: () => runtime,
            loadStaged: async () => undefined,
            save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
          },
        ),
      ).resolves.toMatchObject({
        handled: true,
        operationId: 'xlsx.formula.insert_aggregate',
        ok: true,
        output: { function: aggregate, inserted: 1, targetRange: 'A3:A3' },
      })
    }
    expect(formulas).toEqual(['=AVERAGE(A1:A2)', '=COUNT(A1:A2)', '=MAX(A1:A2)', '=MIN(A1:A2)'])
  })

  it('rejects xlsx.formula.insert_aggregate before writing into an unstreamed target row', async () => {
    const formulas: string[] = []
    const sourceRange = {
      getRow: () => 1,
      getColumn: () => 0,
      getHeight: () => 3,
      getWidth: () => 2,
    }
    const worksheet = {
      getSheetName: () => 'Budget',
      getSheetId: () => 'sheet-budget',
      getRange: (...arguments_: unknown[]) =>
        typeof arguments_[0] === 'string'
          ? sourceRange
          : { setFormula: (formula: string) => formulas.push(formula) },
    }
    const workbook = {
      getSheets: () => [worksheet],
      setActiveSheet: () => undefined,
    }
    const runtime = { univerAPI: { getActiveWorkbook: () => workbook } }
    const state = {
      file: {
        sheets: [{ id: 'sheet-budget', rowCount: 10, columnCount: 5 }],
      },
      loadedRanges: new Map([
        ['sheet-budget', { startRow: 0, endRow: 3, startColumn: 0, endColumn: 4 }],
      ]),
      flags: { preloadComplete: false },
    } as unknown as LazyWorkbookState

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.formula.insert_aggregate',
          arguments: { sheet: 'Budget', range: 'A2:B4', function: 'AVERAGE' },
        },
        {
          runtime: () => runtime,
          state: () => state,
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.formula.insert_aggregate',
      ok: false,
      error: 'execution_failed',
      message: 'xlsx.formula.insert_aggregate cannot overwrite a row that is still streaming.',
    })
    expect(formulas).toEqual([])
  })

  it('fills only empty target cells from retained examples through xlsx.range.flash_fill', async () => {
    let written: unknown[][] | null = null
    const sourceRange = {
      getRow: () => 0,
      getColumn: () => 2,
      getHeight: () => 3,
      getWidth: () => 1,
    }
    const grid = [
      ['John', 'Sales', 'John·Sales'],
      ['Mary', 'Marketing', ''],
      ['Bob', 'Finance', 'Bob·Finance'],
    ]
    const worksheet = {
      getSheetName: () => 'Budget',
      getSheetId: () => 'sheet-budget',
      getRange: (...arguments_: unknown[]) => {
        if (typeof arguments_[0] === 'string') return sourceRange
        const [row, column, rowCount, columnCount] = arguments_ as [number, number, number, number]
        if (row === 0 && column === 0 && rowCount === 3 && columnCount === 3) {
          return { getValues: () => grid }
        }
        return {
          setValues: (values: unknown[][]) => {
            written = values
          },
        }
      },
    }
    const workbook = {
      getSheets: () => [worksheet],
      setActiveSheet: () => undefined,
    }
    const runtime = { univerAPI: { getActiveWorkbook: () => workbook } }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.range.flash_fill',
          arguments: { sheet: 'Budget', range: 'c1:c3' },
        },
        {
          runtime: () => runtime,
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.range.flash_fill',
      ok: true,
      output: {
        sheet: 'Budget',
        range: 'C1:C3',
        targetRange: 'C1:C3',
        filled: 1,
      },
    })
    expect(written).toEqual([['John·Sales'], ['Mary·Marketing'], ['Bob·Finance']])
  })

  it('bounds single-cell flash-fill probing to 1,000 rows and reports the expanded target', async () => {
    let probeRows = 0
    let written: unknown[][] | null = null
    const sourceRange = {
      getRow: () => 0,
      getColumn: () => 2,
      getHeight: () => 1,
      getWidth: () => 1,
    }
    const grid = [
      ['John', 'Sales', 'John·Sales'],
      ['Mary', 'Marketing', ''],
      ['Bob', 'Finance', ''],
    ]
    const worksheet = {
      getSheetName: () => 'Budget',
      getSheetId: () => 'sheet-budget',
      getMaxRows: () => 2_000,
      getRange: (...arguments_: unknown[]) => {
        if (typeof arguments_[0] === 'string') return sourceRange
        const [row, column, rowCount, columnCount] = arguments_ as [number, number, number, number]
        if (row === 0 && column === 1 && columnCount === 1) {
          probeRows = rowCount
          return { getValues: () => [['Sales'], ['Marketing'], ['Finance'], ['']] }
        }
        if (row === 0 && column === 0 && rowCount === 3 && columnCount === 3) {
          return { getValues: () => grid }
        }
        return {
          setValues: (values: unknown[][]) => {
            written = values
          },
        }
      },
    }
    const runtime = {
      univerAPI: {
        getActiveWorkbook: () => ({
          getSheets: () => [worksheet],
          setActiveSheet: () => undefined,
        }),
      },
    }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.range.flash_fill',
          arguments: { sheet: 'Budget', range: 'c1' },
        },
        {
          runtime: () => runtime,
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toMatchObject({
      handled: true,
      operationId: 'xlsx.range.flash_fill',
      ok: true,
      output: { range: 'C1', targetRange: 'C1:C3', filled: 2 },
    })
    expect(probeRows).toBe(1_000)
    expect(written).toEqual([['John·Sales'], ['Mary·Marketing'], ['Bob·Finance']])
  })

  it('rejects single-cell flash fill before reading source rows that are still streaming', async () => {
    let read = false
    let written = false
    const sourceRange = {
      getRow: () => 0,
      getColumn: () => 2,
      getHeight: () => 1,
      getWidth: () => 1,
    }
    const worksheet = {
      getSheetName: () => 'Budget',
      getSheetId: () => 'sheet-budget',
      getMaxRows: () => 2_000,
      getRange: (...arguments_: unknown[]) =>
        typeof arguments_[0] === 'string'
          ? sourceRange
          : {
              getValues: () => {
                read = true
                return [['Sales'], ['Marketing']]
              },
              setValues: () => {
                written = true
              },
            },
    }
    const runtime = {
      univerAPI: {
        getActiveWorkbook: () => ({
          getSheets: () => [worksheet],
          setActiveSheet: () => undefined,
        }),
      },
    }
    const state = {
      file: {
        sheets: [{ id: 'sheet-budget', rowCount: 10, columnCount: 5 }],
      },
      loadedRanges: new Map([
        ['sheet-budget', { startRow: 0, endRow: 0, startColumn: 0, endColumn: 4 }],
      ]),
      flags: { preloadComplete: false },
    } as unknown as LazyWorkbookState

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.range.flash_fill',
          arguments: { sheet: 'Budget', range: 'C1' },
        },
        {
          runtime: () => runtime,
          state: () => state,
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.range.flash_fill',
      ok: false,
      error: 'execution_failed',
      message: 'xlsx.range.flash_fill requires its source rows to finish streaming.',
    })
    expect(read).toBe(false)
    expect(written).toBe(false)
  })

  it('rejects explicit flash-fill ranges before reading partially streamed inference data', async () => {
    let read = false
    let written = false
    const sourceRange = {
      getRow: () => 0,
      getColumn: () => 2,
      getHeight: () => 3,
      getWidth: () => 1,
    }
    const worksheet = {
      getSheetName: () => 'Budget',
      getSheetId: () => 'sheet-budget',
      getRange: (...arguments_: unknown[]) =>
        typeof arguments_[0] === 'string'
          ? sourceRange
          : {
              getValues: () => {
                read = true
                return []
              },
              setValues: () => {
                written = true
              },
            },
    }
    const runtime = {
      univerAPI: {
        getActiveWorkbook: () => ({
          getSheets: () => [worksheet],
          setActiveSheet: () => undefined,
        }),
      },
    }
    const state = {
      file: {
        sheets: [{ id: 'sheet-budget', rowCount: 10, columnCount: 5 }],
      },
      loadedRanges: new Map([
        ['sheet-budget', { startRow: 0, endRow: 1, startColumn: 0, endColumn: 4 }],
      ]),
      flags: { preloadComplete: false },
    } as unknown as LazyWorkbookState

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.range.flash_fill',
          arguments: { sheet: 'Budget', range: 'C1:C3' },
        },
        {
          runtime: () => runtime,
          state: () => state,
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toMatchObject({
      handled: true,
      operationId: 'xlsx.range.flash_fill',
      ok: false,
      error: 'execution_failed',
      message: 'xlsx.range.flash_fill requires its source rows to finish streaming.',
    })
    expect(read).toBe(false)
    expect(written).toBe(false)
  })

  it('rejects flash fill in the first column before reading worksheet data', async () => {
    let accessed = false
    const sourceRange = {
      getRow: () => 0,
      getColumn: () => 0,
      getHeight: () => 3,
      getWidth: () => 1,
    }
    const worksheet = {
      getSheetName: () => 'Budget',
      getSheetId: () => 'sheet-budget',
      getRange: (...arguments_: unknown[]) => {
        if (typeof arguments_[0] === 'string') return sourceRange
        accessed = true
        return { getValues: () => [], setValues: () => undefined }
      },
    }
    const runtime = {
      univerAPI: {
        getActiveWorkbook: () => ({
          getSheets: () => [worksheet],
          setActiveSheet: () => undefined,
        }),
      },
    }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.range.flash_fill',
          arguments: { sheet: 'Budget', range: 'A1:A3' },
        },
        {
          runtime: () => runtime,
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.range.flash_fill',
      ok: false,
      error: 'invalid_arguments',
      message: 'xlsx.range.flash_fill requires a source column to the left of the target.',
    })
    expect(accessed).toBe(false)
  })

  it('splits one text column by comma through xlsx.range.text_to_columns', async () => {
    const commands: unknown[] = []
    const targetRange = {
      getRow: () => 1,
      getColumn: () => 2,
      getHeight: () => 3,
      getWidth: () => 1,
    }
    const worksheet = {
      getSheetName: () => 'Budget',
      getSheetId: () => 'sheet-budget',
      getRange: () => targetRange,
    }
    const workbook = {
      getId: () => 'workbook-budget',
      getSheets: () => [worksheet],
      setActiveSheet: () => undefined,
    }
    const runtime = {
      univerAPI: {
        getActiveWorkbook: () => workbook,
        executeCommand: async (command: string, parameters: unknown) => {
          commands.push([command, parameters])
          return true
        },
      },
    }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.range.text_to_columns',
          arguments: { sheet: 'Budget', range: 'c2:c4', delimiter: 'comma' },
        },
        {
          runtime: () => runtime,
          isSheetDataComplete: () => true,
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.range.text_to_columns',
      ok: true,
      output: { sheet: 'Budget', range: 'C2:C4', delimiter: 'comma' },
    })
    expect(commands).toEqual([
      [
        'sheet.command.split-text-to-columns',
        {
          unitId: 'workbook-budget',
          subUnitId: 'sheet-budget',
          range: { startRow: 1, endRow: 3, startColumn: 2, endColumn: 2 },
          delimiter: 2,
        },
      ],
    ])
  })

  it.each([
    ['tab', 1],
    ['semicolon', 4],
    ['space', 8],
  ] as const)('maps the retained %s delimiter to Univer flag %s', async (delimiter, flag) => {
    let commandParameters: unknown
    const targetRange = {
      getRow: () => 0,
      getColumn: () => 0,
      getHeight: () => 2,
      getWidth: () => 1,
    }
    const worksheet = {
      getSheetName: () => 'Budget',
      getSheetId: () => 'sheet-budget',
      getRange: () => targetRange,
    }
    const workbook = {
      getId: () => 'workbook-budget',
      getSheets: () => [worksheet],
      setActiveSheet: () => undefined,
    }
    const runtime = {
      univerAPI: {
        getActiveWorkbook: () => workbook,
        executeCommand: async (_command: string, parameters: unknown) => {
          commandParameters = parameters
          return true
        },
      },
    }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.range.text_to_columns',
          arguments: { sheet: 'Budget', range: 'A1:A2', delimiter },
        },
        {
          runtime: () => runtime,
          isSheetDataComplete: () => true,
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toMatchObject({
      handled: true,
      operationId: 'xlsx.range.text_to_columns',
      ok: true,
      output: { delimiter },
    })
    expect(commandParameters).toMatchObject({ delimiter: flag })
  })

  it('rejects multi-column text-to-columns targets before executing a command', async () => {
    let mutated = false
    const worksheet = {
      getSheetName: () => 'Budget',
      getSheetId: () => 'sheet-budget',
      getRange: () => ({
        getRow: () => 0,
        getColumn: () => 0,
        getHeight: () => 3,
        getWidth: () => 2,
      }),
    }
    const workbook = {
      getId: () => 'workbook-budget',
      getSheets: () => [worksheet],
      setActiveSheet: () => undefined,
    }
    const runtime = {
      univerAPI: {
        getActiveWorkbook: () => workbook,
        executeCommand: async () => {
          mutated = true
          return true
        },
      },
    }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.range.text_to_columns',
          arguments: { sheet: 'Budget', range: 'A1:B3', delimiter: 'tab' },
        },
        {
          runtime: () => runtime,
          isSheetDataComplete: () => true,
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.range.text_to_columns',
      ok: false,
      error: 'invalid_arguments',
      message: 'xlsx.range.text_to_columns requires a single-column range.',
    })
    expect(mutated).toBe(false)
  })

  it('rejects text to columns before executing against partially streamed sheet data', async () => {
    let mutated = false
    const worksheet = {
      getSheetName: () => 'Budget',
      getSheetId: () => 'sheet-budget',
      getRange: () => ({
        getRow: () => 0,
        getColumn: () => 0,
        getHeight: () => 3,
        getWidth: () => 1,
      }),
    }
    const workbook = {
      getId: () => 'workbook-budget',
      getSheets: () => [worksheet],
      setActiveSheet: () => undefined,
    }
    const runtime = {
      univerAPI: {
        getActiveWorkbook: () => workbook,
        executeCommand: async () => {
          mutated = true
          return true
        },
      },
    }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.range.text_to_columns',
          arguments: { sheet: 'Budget', range: 'A1:A3', delimiter: 'space' },
        },
        {
          runtime: () => runtime,
          isSheetDataComplete: () => false,
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.range.text_to_columns',
      ok: false,
      error: 'execution_failed',
      message: 'xlsx.range.text_to_columns requires fully loaded sheet data.',
    })
    expect(mutated).toBe(false)
  })

  it('reports execution failure when Univer does not complete text to columns', async () => {
    const worksheet = {
      getSheetName: () => 'Budget',
      getSheetId: () => 'sheet-budget',
      getRange: () => ({
        getRow: () => 0,
        getColumn: () => 0,
        getHeight: () => 2,
        getWidth: () => 1,
      }),
    }
    const workbook = {
      getId: () => 'workbook-budget',
      getSheets: () => [worksheet],
      setActiveSheet: () => undefined,
    }
    const runtime = {
      univerAPI: {
        getActiveWorkbook: () => workbook,
        executeCommand: async () => false,
      },
    }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.range.text_to_columns',
          arguments: { sheet: 'Budget', range: 'A1:A2', delimiter: 'semicolon' },
        },
        {
          runtime: () => runtime,
          isSheetDataComplete: () => true,
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.range.text_to_columns',
      ok: false,
      error: 'execution_failed',
      message: 'Univer did not complete xlsx.range.text_to_columns.',
    })
  })

  it('sets a normalized hyperlink through xlsx.hyperlink.set and records the save journal', async () => {
    const journal = createEditJournal()
    const writes: unknown[] = []
    let activeRange: string | null = null
    let pendingEdits = 0
    const worksheet = {
      getSheetName: () => 'Budget',
      getSheetId: () => 'sheet-budget',
      getRange: (address: string) => ({
        setValue: (value: unknown) => writes.push(value),
        activate: () => {
          activeRange = address
        },
      }),
    }
    const workbook = {
      getSheets: () => [worksheet],
      setActiveSheet: () => undefined,
    }
    const runtime = { univerAPI: { getActiveWorkbook: () => workbook } }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.hyperlink.set',
          arguments: { sheet: 'Budget', address: 'b2', target: 'example.com/report' },
        },
        {
          runtime: () => runtime,
          state: () => ({
            editJournal: journal,
            hyperlinkTargets: new Map([['sheet-budget', new Map<string, string>()]]),
            sheetProtections: new Map(),
          }),
          setPendingEdits: (count: number) => {
            pendingEdits = count
          },
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.hyperlink.set',
      ok: true,
      output: {
        sheet: 'Budget',
        address: 'B2',
        target: 'https://example.com/report',
      },
    })
    expect(journal.hyperlinks.get('sheet-budget')?.get('1:1')).toBe('https://example.com/report')
    expect(writes).toEqual([{ s: { cl: { rgb: '#0563C1' }, ul: { s: BooleanNumber.TRUE } } }])
    expect(activeRange).toBe('b2')
    expect(pendingEdits).toBe(1)
  })

  it('removes a hyperlink through xlsx.hyperlink.remove and clears its link appearance', async () => {
    const journal = createEditJournal()
    const targets = new Map([['1:1', 'https://example.com/report']])
    const writes: unknown[] = []
    const worksheet = {
      getSheetName: () => 'Budget',
      getSheetId: () => 'sheet-budget',
      getRange: () => ({
        setValue: (value: unknown) => writes.push(value),
        activate: () => writes.push('activate'),
      }),
    }
    const workbook = {
      getSheets: () => [worksheet],
      setActiveSheet: () => undefined,
    }
    const runtime = { univerAPI: { getActiveWorkbook: () => workbook } }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.hyperlink.remove',
          arguments: { sheet: 'Budget', address: 'B2' },
        },
        {
          runtime: () => runtime,
          state: () => ({
            editJournal: journal,
            hyperlinkTargets: new Map([['sheet-budget', targets]]),
            sheetProtections: new Map(),
          }),
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.hyperlink.remove',
      ok: true,
      output: { sheet: 'Budget', address: 'B2', removed: true },
    })
    expect(journal.hyperlinks.get('sheet-budget')?.get('1:1')).toBeNull()
    expect(targets.has('1:1')).toBe(false)
    expect(writes).toEqual([{ s: { ul: null, cl: null } }, 'activate'])
  })

  it('creates a styled table through xlsx.table.add and records the same table for save', async () => {
    const journal = createEditJournal()
    const renderedTables: unknown[] = []
    let activeRange: string | null = null
    const worksheet = {
      getSheetName: () => 'Budget',
      getSheetId: () => 'sheet-budget',
      getRange: (...arguments_: unknown[]) =>
        typeof arguments_[0] === 'string'
          ? {
              activate: () => {
                activeRange = arguments_[0] as string
              },
            }
          : { getValues: () => [['Region', 'Amount']] },
      addTable: (name: string, range: unknown) => {
        renderedTables.push({ name, range })
        return Promise.resolve()
      },
    }
    const workbook = {
      getSheets: () => [worksheet],
      getSheetBySheetId: () => worksheet,
      setActiveSheet: () => undefined,
    }
    const runtime = { univerAPI: { getActiveWorkbook: () => workbook } }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.table.add',
          arguments: {
            sheet: 'Budget',
            range: 'a1:b3',
            style: 'TableStyleMedium4',
          },
        },
        {
          runtime: () => runtime,
          state: () => ({
            editJournal: journal,
            hyperlinkTargets: new Map(),
            sheetProtections: new Map(),
          }),
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.table.add',
      ok: true,
      output: {
        sheet: 'Budget',
        range: 'A1:B3',
        name: 'Table1',
        style: 'TableStyleMedium4',
      },
    })
    expect(journal.tableAdds).toEqual([
      expect.objectContaining({
        sheetId: 'sheet-budget',
        name: 'Table1',
        style: 'TableStyleMedium4',
        bandedRows: true,
        columnNames: ['Region', 'Amount'],
        area: { startRow: 0, startColumn: 0, endRow: 2, endColumn: 1 },
      }),
    ])
    expect(renderedTables).toEqual([
      {
        name: 'Table1',
        range: { startRow: 0, startColumn: 0, endRow: 2, endColumn: 1 },
      },
    ])
    expect(activeRange).toBe('a1:b3')
  })

  it('adds explicit row-aligned sparklines through xlsx.sparkline.add', async () => {
    const journal = createEditJournal()
    let targetActivated = false
    let refreshes = 0
    const worksheet = {
      getSheetName: () => 'Budget',
      getSheetId: () => 'sheet-budget',
      getRange: (address: string) => {
        if (address === 'A1:C2') {
          return {
            getRow: () => 0,
            getColumn: () => 0,
            getHeight: () => 2,
            getWidth: () => 3,
          }
        }
        return {
          getRow: () => 0,
          getColumn: () => 3,
          getHeight: () => 2,
          getWidth: () => 1,
          activate: () => {
            targetActivated = true
          },
        }
      },
    }
    const workbook = {
      getId: () => undefined,
      getSheets: () => [worksheet],
      setActiveSheet: () => undefined,
    }
    const runtime = { univerAPI: { getActiveWorkbook: () => workbook } }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.sparkline.add',
          arguments: {
            sheet: 'Budget',
            sourceRange: 'a1:c2',
            targetRange: 'd1:d2',
            type: 'column',
          },
        },
        {
          runtime: () => runtime,
          state: () => ({
            editJournal: journal,
            hyperlinkTargets: new Map(),
            sheetProtections: new Map(),
            file: {
              sheets: [{ id: 'sheet-budget', name: 'Budget', sparklines: [] }],
            } as unknown as LazyWorkbookState['file'],
          }),
          refreshSparklines: () => {
            refreshes += 1
          },
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.sparkline.add',
      ok: true,
      output: {
        sheet: 'Budget',
        sourceRange: 'A1:C2',
        targetRange: 'D1:D2',
        type: 'column',
        count: 2,
      },
    })
    expect(journal.sparklineAdds).toEqual([
      expect.objectContaining({
        sheetId: 'sheet-budget',
        type: 'column',
        cells: [
          { cell: 'D1', sourceRef: "'Budget'!$A$1:$C$1" },
          { cell: 'D2', sourceRef: "'Budget'!$A$2:$C$2" },
        ],
      }),
    ])
    expect(targetActivated).toBe(true)
    expect(refreshes).toBe(1)
  })

  it('sets an explicit final row-outline level through xlsx.outline.set_level', async () => {
    const journal = createEditJournal()
    const outline = new Map([
      [
        'sheet-budget',
        {
          rows: new Map([[1, { level: 1, collapsed: false, hidden: false }]]),
          cols: new Map(),
        },
      ],
    ])
    const worksheet = {
      getSheetName: () => 'Budget',
      getSheetId: () => 'sheet-budget',
      getMaxRows: () => 100,
      getMaxColumns: () => 26,
    }
    const workbook = {
      getId: () => undefined,
      getSheets: () => [worksheet],
      setActiveSheet: () => undefined,
    }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.outline.set_level',
          arguments: { sheet: 'Budget', axis: 'rows', start: 2, count: 3, level: 2 },
        },
        {
          runtime: () => ({ univerAPI: { getActiveWorkbook: () => workbook } }) as never,
          state: () => ({
            editJournal: journal,
            hyperlinkTargets: new Map(),
            sheetProtections: new Map(),
            outline,
          }),
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.outline.set_level',
      ok: true,
      output: { sheet: 'Budget', axis: 'rows', start: 2, count: 3, level: 2 },
    })
    expect(journal.structuralOps.get('sheet-budget')).toEqual([
      { kind: 'set-rows-outline', start: 1, end: 3, level: 2 },
    ])
    expect(outline.get('sheet-budget')?.rows).toEqual(
      new Map([
        [1, { level: 2, collapsed: false, hidden: false }],
        [2, { level: 2, collapsed: false, hidden: false }],
        [3, { level: 2, collapsed: false, hidden: false }],
      ]),
    )
  })

  it('sets explicit final outline-detail visibility through xlsx.outline.set_detail_visibility', async () => {
    const journal = createEditJournal()
    const outline = new Map([
      [
        'sheet-budget',
        {
          rows: new Map([
            [1, { level: 1, collapsed: false, hidden: false }],
            [2, { level: 1, collapsed: false, hidden: false }],
            [3, { level: 0, collapsed: false, hidden: false }],
          ]),
          cols: new Map(),
        },
      ],
    ])
    const hidden: Array<{ start: number; count: number }> = []
    const worksheet = {
      getSheetName: () => 'Budget',
      getSheetId: () => 'sheet-budget',
      getMaxRows: () => 100,
      getMaxColumns: () => 26,
      hideRows: (start: number, count: number) => hidden.push({ start, count }),
      showRows: () => undefined,
    }
    const workbook = {
      getId: () => undefined,
      getSheets: () => [worksheet],
      setActiveSheet: () => undefined,
    }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.outline.set_detail_visibility',
          arguments: { sheet: 'Budget', axis: 'rows', start: 2, count: 2, hidden: true },
        },
        {
          runtime: () => ({ univerAPI: { getActiveWorkbook: () => workbook } }) as never,
          state: () => ({
            editJournal: journal,
            hyperlinkTargets: new Map(),
            sheetProtections: new Map(),
            outline,
          }),
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.outline.set_detail_visibility',
      ok: true,
      output: { sheet: 'Budget', axis: 'rows', start: 2, count: 2, hidden: true },
    })
    expect(hidden).toEqual([{ start: 1, count: 2 }])
    expect(journal.structuralOps.get('sheet-budget')).toEqual([
      { kind: 'set-rows-hidden', start: 1, end: 2, hidden: true },
      { kind: 'set-rows-outline', start: 3, end: 3, level: 0, collapsed: true },
    ])
  })

  it('sets an explicit final checkbox-validation state through xlsx.range.set_checkbox', async () => {
    const journal = createEditJournal()
    const validations: unknown[] = []
    let active = false
    const range = {
      getRow: () => 1,
      getColumn: () => 1,
      getHeight: () => 2,
      getWidth: () => 2,
      setDataValidation: (value: unknown) => validations.push(value),
      activate: () => {
        active = true
      },
    }
    const worksheet = {
      getSheetName: () => 'Budget',
      getSheetId: () => 'sheet-budget',
      getRange: () => range,
    }
    const workbook = {
      getSheets: () => [worksheet],
      setActiveSheet: () => undefined,
    }
    const built = { type: 'checkbox' }
    const runtime = {
      univerAPI: {
        getActiveWorkbook: () => workbook,
        newDataValidation: () => ({
          requireCheckbox: () => ({ build: () => built }),
        }),
      },
    }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.range.set_checkbox',
          arguments: { sheet: 'Budget', range: 'b2:c3', enabled: true },
        },
        {
          runtime: () => runtime as never,
          state: () => ({
            editJournal: journal,
            hyperlinkTargets: new Map(),
            sheetProtections: new Map(),
            appliedDvSheets: new Set(['sheet-budget']),
          }),
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.range.set_checkbox',
      ok: true,
      output: { sheet: 'Budget', range: 'B2:C3', enabled: true },
    })
    expect(validations).toEqual([built])
    expect(active).toBe(true)
  })

  it('sets and removes explicit list validation through the native validation history route', async () => {
    const journal = createEditJournal()
    const validations: unknown[] = []
    let active = 0
    const range = {
      getRow: () => 1,
      getColumn: () => 1,
      getHeight: () => 2,
      getWidth: () => 2,
      setDataValidation: (value: unknown) => validations.push(value),
      activate: () => {
        active += 1
      },
    }
    const worksheet = {
      getSheetName: () => 'Budget',
      getSheetId: () => 'sheet-budget',
      getRange: () => range,
    }
    const workbook = { getSheets: () => [worksheet], setActiveSheet: () => undefined }
    const built = { type: 'list', values: ['Open', 'Closed'] }
    const allowBlank: boolean[] = []
    const runtime = {
      univerAPI: {
        getActiveWorkbook: () => workbook,
        newDataValidation: () => ({
          requireValueInList: (values: string[], multiple: boolean, showDropdown: boolean) => {
            expect(values).toEqual(['Open', 'Closed'])
            expect(multiple).toBe(false)
            expect(showDropdown).toBe(true)
            return {
              setAllowBlank: (enabled: boolean) => {
                allowBlank.push(enabled)
                return { build: () => built }
              },
            }
          },
        }),
      },
    }
    const services = {
      runtime: () => runtime as never,
      state: () => ({
        editJournal: journal,
        hyperlinkTargets: new Map(),
        sheetProtections: new Map(),
        appliedDvSheets: new Set(['sheet-budget']),
      }),
      loadStaged: async () => undefined,
      save: async () => ({ ok: true as const, fileName: 'budget.xlsx' }),
    }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.range.set_list_validation',
          arguments: {
            sheet: 'Budget',
            range: 'b2:c3',
            values: ['Open', 'Closed'],
            allowBlank: true,
            showDropdown: true,
          },
        },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.range.set_list_validation',
      ok: true,
      output: {
        sheet: 'Budget',
        range: 'B2:C3',
        valueCount: 2,
        allowBlank: true,
        showDropdown: true,
      },
    })
    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.range.remove_data_validation',
          arguments: { sheet: 'Budget', range: 'B2:C3' },
        },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.range.remove_data_validation',
      ok: true,
      output: { sheet: 'Budget', range: 'B2:C3', removed: true },
    })
    expect(validations).toEqual([built, null])
    expect(allowBlank).toEqual([true])
    expect(active).toBe(2)
  })

  it('rejects list validation that cannot round-trip through an inline XLSX list source', async () => {
    let mutated = false
    const range = {
      getRow: () => 0,
      getColumn: () => 0,
      getHeight: () => 1,
      getWidth: () => 1,
      setDataValidation: () => {
        mutated = true
      },
    }
    const worksheet = {
      getSheetName: () => 'Budget',
      getSheetId: () => 'sheet-budget',
      getRange: () => range,
    }
    const services = {
      runtime: () =>
        ({
          univerAPI: {
            getActiveWorkbook: () => ({
              getSheets: () => [worksheet],
              setActiveSheet: () => undefined,
            }),
          },
        }) as never,
      state: () => ({
        editJournal: createEditJournal(),
        hyperlinkTargets: new Map(),
        sheetProtections: new Map(),
        appliedDvSheets: new Set(['sheet-budget']),
      }),
      loadStaged: async () => undefined,
      save: async () => ({ ok: true as const, fileName: 'budget.xlsx' }),
    }
    for (const values of [[''], ['North,East'], ['x'.repeat(256)], Array(100).fill('abc')]) {
      await expect(
        executeXlsxOperation(
          {
            operation: 'xlsx.range.set_list_validation',
            arguments: {
              sheet: 'Budget',
              range: 'A1',
              values,
              allowBlank: false,
              showDropdown: true,
            },
          },
          services,
        ),
      ).resolves.toMatchObject({ handled: true, ok: false, error: 'invalid_arguments' })
    }
    expect(mutated).toBe(false)
  })

  it('sets range-backed list validation with explicit target and source ranges', async () => {
    const validations: unknown[] = []
    let active = false
    const target = {
      getRow: () => 0,
      getColumn: () => 0,
      getHeight: () => 10,
      getWidth: () => 1,
      setDataValidation: (rule: unknown) => validations.push(rule),
      activate: () => {
        active = true
      },
    }
    const source = {
      getRow: () => 0,
      getColumn: () => 3,
      getHeight: () => 3,
      getWidth: () => 1,
    }
    const targetSheet = {
      getSheetName: () => 'Budget',
      getSheetId: () => 'sheet-budget',
      getRange: () => target,
    }
    const sourceSheet = {
      getSheetName: () => 'Lists',
      getSheetId: () => 'sheet-lists',
      getRange: () => source,
    }
    const built = { type: 'listRef', source }
    const runtime = {
      univerAPI: {
        getActiveWorkbook: () => ({
          getSheets: () => [targetSheet, sourceSheet],
          setActiveSheet: () => undefined,
        }),
        newDataValidation: () => ({
          requireValueInRange: (range: unknown, multiple: boolean, showDropdown: boolean) => {
            expect(range).toBe(source)
            expect(multiple).toBe(false)
            expect(showDropdown).toBe(true)
            return {
              setAllowBlank: (allowBlank: boolean) => {
                expect(allowBlank).toBe(true)
                return { build: () => built }
              },
            }
          },
        }),
      },
    }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.range.set_list_reference_validation',
          arguments: {
            sheet: 'Budget',
            range: 'A1:A10',
            sourceSheet: 'Lists',
            sourceRange: 'D1:D3',
            allowBlank: true,
            showDropdown: true,
          },
        },
        {
          runtime: () => runtime as never,
          state: () => ({
            editJournal: createEditJournal(),
            hyperlinkTargets: new Map(),
            sheetProtections: new Map(),
            appliedDvSheets: new Set(['sheet-budget']),
          }),
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.range.set_list_reference_validation',
      ok: true,
      output: {
        sheet: 'Budget',
        range: 'A1:A10',
        sourceSheet: 'Lists',
        sourceRange: 'D1:D3',
        allowBlank: true,
        showDropdown: true,
      },
    })
    expect(validations).toEqual([built])
    expect(active).toBe(true)
  })

  it('rejects two-dimensional or oversized validation list source ranges before mutation', async () => {
    let mutated = false
    const target = {
      getRow: () => 0,
      getColumn: () => 0,
      getHeight: () => 1,
      getWidth: () => 1,
      setDataValidation: () => {
        mutated = true
      },
    }
    const source = {
      getRow: () => 0,
      getColumn: () => 0,
      getHeight: () => 1001,
      getWidth: () => 2,
    }
    const worksheet = (name: string, id: string, range: unknown) => ({
      getSheetName: () => name,
      getSheetId: () => id,
      getRange: () => range,
    })
    const services = {
      runtime: () =>
        ({
          univerAPI: {
            getActiveWorkbook: () => ({
              getSheets: () => [
                worksheet('Budget', 'sheet-budget', target),
                worksheet('Lists', 'sheet-lists', source),
              ],
              setActiveSheet: () => undefined,
            }),
          },
        }) as never,
      state: () => ({
        editJournal: createEditJournal(),
        hyperlinkTargets: new Map(),
        sheetProtections: new Map(),
        appliedDvSheets: new Set(['sheet-budget']),
      }),
      loadStaged: async () => undefined,
      save: async () => ({ ok: true as const, fileName: 'budget.xlsx' }),
    }
    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.range.set_list_reference_validation',
          arguments: {
            sheet: 'Budget',
            range: 'A1',
            sourceSheet: 'Lists',
            sourceRange: 'A1:B1001',
            allowBlank: false,
            showDropdown: true,
          },
        },
        services,
      ),
    ).resolves.toMatchObject({ handled: true, ok: false, error: 'invalid_arguments' })
    expect(mutated).toBe(false)
  })

  it('sets and bounds custom-formula validation through the native builder route', async () => {
    const validations: unknown[] = []
    const range = {
      getRow: () => 0,
      getColumn: () => 0,
      getHeight: () => 5,
      getWidth: () => 1,
      setDataValidation: (rule: unknown) => validations.push(rule),
      activate: () => undefined,
    }
    const worksheet = {
      getSheetName: () => 'Budget',
      getSheetId: () => 'sheet-budget',
      getRange: () => range,
    }
    const services = {
      runtime: () =>
        ({
          univerAPI: {
            getActiveWorkbook: () => ({
              getSheets: () => [worksheet],
              setActiveSheet: () => undefined,
            }),
            newDataValidation: () => ({
              requireFormulaSatisfied: (formula: string) => ({
                setAllowBlank: (allowBlank: boolean) => ({
                  build: () => ({ type: 'formula', formula, allowBlank }),
                }),
              }),
            }),
          },
        }) as never,
      state: () => ({
        editJournal: createEditJournal(),
        hyperlinkTargets: new Map(),
        sheetProtections: new Map(),
        appliedDvSheets: new Set(['sheet-budget']),
      }),
      loadStaged: async () => undefined,
      save: async () => ({ ok: true as const, fileName: 'budget.xlsx' }),
    }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.range.set_custom_formula_validation',
          arguments: {
            sheet: 'Budget',
            range: 'A1:A5',
            formula: '=AND(A1>0,A1<10)',
            allowBlank: false,
          },
        },
        services,
      ),
    ).resolves.toMatchObject({
      handled: true,
      operationId: 'xlsx.range.set_custom_formula_validation',
      ok: true,
    })
    expect(validations).toEqual([
      { type: 'formula', formula: '=AND(A1>0,A1<10)', allowBlank: false },
    ])

    for (const formula of ['A1>0', `=${'A'.repeat(8_192)}`]) {
      await expect(
        executeXlsxOperation(
          {
            operation: 'xlsx.range.set_custom_formula_validation',
            arguments: { sheet: 'Budget', range: 'A1', formula, allowBlank: false },
          },
          services,
        ),
      ).resolves.toMatchObject({ handled: true, ok: false, error: 'invalid_arguments' })
    }
  })

  it('sets all retained scalar comparison validation kinds through one bounded operation', async () => {
    const validations: Array<{ rule: Record<string, unknown> }> = []
    const range = {
      getRow: () => 0,
      getColumn: () => 0,
      getHeight: () => 5,
      getWidth: () => 1,
      setDataValidation: (validation: { rule: Record<string, unknown> }) =>
        validations.push(validation),
      activate: () => undefined,
    }
    const worksheet = {
      getSheetName: () => 'Budget',
      getSheetId: () => 'sheet-budget',
      getRange: () => range,
    }
    const services = {
      runtime: () =>
        ({
          univerAPI: {
            getActiveWorkbook: () => ({
              getSheets: () => [worksheet],
              setActiveSheet: () => undefined,
            }),
          },
        }) as never,
      state: () => ({
        editJournal: createEditJournal(),
        hyperlinkTargets: new Map(),
        sheetProtections: new Map(),
        appliedDvSheets: new Set(['sheet-budget']),
      }),
      loadStaged: async () => undefined,
      save: async () => ({ ok: true as const, fileName: 'budget.xlsx' }),
    }
    const comparisons = [
      { kind: 'whole', operator: 'greaterThan', operand1: 0, allowBlank: false },
      {
        kind: 'decimal',
        operator: 'between',
        operand1: 1.5,
        operand2: 9.5,
        allowBlank: true,
      },
      {
        kind: 'date',
        operator: 'notEqual',
        operand1: '2026-01-01',
        allowBlank: false,
      },
      { kind: 'time', operator: 'lessThan', operand1: '17:30', allowBlank: true },
      {
        kind: 'textLength',
        operator: 'between',
        operand1: 1,
        operand2: 20,
        allowBlank: false,
      },
    ]
    for (const comparison of comparisons) {
      await expect(
        executeXlsxOperation(
          {
            operation: 'xlsx.range.set_comparison_validation',
            arguments: { sheet: 'Budget', range: 'A1:A5', ...comparison },
          },
          services,
        ),
      ).resolves.toMatchObject({
        handled: true,
        operationId: 'xlsx.range.set_comparison_validation',
        ok: true,
        output: { sheet: 'Budget', range: 'A1:A5', ...comparison },
      })
    }
    expect(validations.map(({ rule }) => ({ ...rule, uid: undefined }))).toEqual([
      {
        uid: undefined,
        type: 'whole',
        operator: 'greaterThan',
        formula1: '0',
        allowBlank: false,
      },
      {
        uid: undefined,
        type: 'decimal',
        operator: 'between',
        formula1: '1.5',
        formula2: '9.5',
        allowBlank: true,
      },
      {
        uid: undefined,
        type: 'date',
        operator: 'notEqual',
        formula1: '2026-01-01',
        allowBlank: false,
      },
      {
        uid: undefined,
        type: 'time',
        operator: 'lessThan',
        formula1: '17:30',
        allowBlank: true,
      },
      {
        uid: undefined,
        type: 'textLength',
        operator: 'between',
        formula1: '1',
        formula2: '20',
        allowBlank: false,
      },
    ])
  })

  it('rejects comparison validation operands that do not match kind and operator semantics', async () => {
    let mutated = false
    const range = {
      getRow: () => 0,
      getColumn: () => 0,
      getHeight: () => 1,
      getWidth: () => 1,
      setDataValidation: () => {
        mutated = true
      },
    }
    const worksheet = {
      getSheetName: () => 'Budget',
      getSheetId: () => 'sheet-budget',
      getRange: () => range,
    }
    const services = {
      runtime: () =>
        ({
          univerAPI: {
            getActiveWorkbook: () => ({
              getSheets: () => [worksheet],
              setActiveSheet: () => undefined,
            }),
          },
        }) as never,
      state: () => ({
        editJournal: createEditJournal(),
        hyperlinkTargets: new Map(),
        sheetProtections: new Map(),
        appliedDvSheets: new Set(['sheet-budget']),
      }),
      loadStaged: async () => undefined,
      save: async () => ({ ok: true as const, fileName: 'budget.xlsx' }),
    }
    for (const comparison of [
      { kind: 'decimal', operator: 'between', operand1: 1 },
      { kind: 'decimal', operator: 'equal', operand1: 1, operand2: 2 },
      { kind: 'whole', operator: 'equal', operand1: 1.5 },
      { kind: 'date', operator: 'equal', operand1: '2026-02-30' },
      { kind: 'time', operator: 'equal', operand1: '25:00' },
      { kind: 'textLength', operator: 'equal', operand1: -1 },
      { kind: 'textLength', operator: 'equal', operand1: 32_768 },
      { kind: 'whole', operator: 'notBetween', operand1: 10, operand2: 1 },
    ]) {
      await expect(
        executeXlsxOperation(
          {
            operation: 'xlsx.range.set_comparison_validation',
            arguments: { sheet: 'Budget', range: 'A1', allowBlank: false, ...comparison },
          },
          services,
        ),
      ).resolves.toMatchObject({ handled: true, ok: false, error: 'invalid_arguments' })
    }
    expect(mutated).toBe(false)
  })

  it('sets and clears bounded validation messages through the existing native rule', async () => {
    const options: unknown[] = []
    let active = 0
    const validation = { setOptions: (value: unknown) => options.push(value) }
    const range = {
      getRow: () => 0,
      getColumn: () => 0,
      getHeight: () => 5,
      getWidth: () => 1,
      getDataValidation: () => validation,
      activate: () => {
        active += 1
      },
    }
    const worksheet = {
      getSheetName: () => 'Budget',
      getSheetId: () => 'sheet-budget',
      getRange: () => range,
    }
    const services = {
      runtime: () =>
        ({
          univerAPI: {
            getActiveWorkbook: () => ({
              getSheets: () => [worksheet],
              setActiveSheet: () => undefined,
            }),
          },
        }) as never,
      state: () => ({
        editJournal: createEditJournal(),
        hyperlinkTargets: new Map(),
        sheetProtections: new Map(),
        appliedDvSheets: new Set(['sheet-budget']),
      }),
      loadStaged: async () => undefined,
      save: async () => ({ ok: true as const, fileName: 'budget.xlsx' }),
    }
    const setArguments = {
      sheet: 'Budget',
      range: 'A1:A5',
      inputTitle: 'Choose status',
      inputMessage: 'Use one of the approved values.',
      errorStyle: 'warning',
      errorTitle: 'Unexpected status',
      errorMessage: 'This value is not in the approved list.',
    }
    await expect(
      executeXlsxOperation(
        { operation: 'xlsx.range.set_validation_messages', arguments: setArguments },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.range.set_validation_messages',
      ok: true,
      output: setArguments,
    })
    const clearArguments = {
      sheet: 'Budget',
      range: 'A1:A5',
      inputTitle: null,
      inputMessage: null,
      errorStyle: 'none',
      errorTitle: null,
      errorMessage: null,
    }
    await expect(
      executeXlsxOperation(
        { operation: 'xlsx.range.set_validation_messages', arguments: clearArguments },
        services,
      ),
    ).resolves.toMatchObject({ ok: true, output: clearArguments })
    expect(options).toEqual([
      {
        showInputMessage: true,
        promptTitle: 'Choose status',
        prompt: 'Use one of the approved values.',
        showErrorMessage: true,
        errorStyle: 2,
        errorTitle: 'Unexpected status',
        error: 'This value is not in the approved list.',
      },
      {
        showInputMessage: false,
        promptTitle: undefined,
        prompt: undefined,
        showErrorMessage: false,
        errorStyle: undefined,
        errorTitle: undefined,
        error: undefined,
      },
    ])
    expect(active).toBe(2)
  })

  it('rejects invalid validation message bounds and ranges without an existing rule', async () => {
    let changed = false
    const range = {
      getRow: () => 0,
      getColumn: () => 0,
      getHeight: () => 1,
      getWidth: () => 1,
      getDataValidation: () => null,
      activate: () => {
        changed = true
      },
    }
    const worksheet = {
      getSheetName: () => 'Budget',
      getSheetId: () => 'sheet-budget',
      getRange: () => range,
    }
    const services = {
      runtime: () =>
        ({
          univerAPI: {
            getActiveWorkbook: () => ({
              getSheets: () => [worksheet],
              setActiveSheet: () => undefined,
            }),
          },
        }) as never,
      state: () => ({
        editJournal: createEditJournal(),
        hyperlinkTargets: new Map(),
        sheetProtections: new Map(),
        appliedDvSheets: new Set(['sheet-budget']),
      }),
      loadStaged: async () => undefined,
      save: async () => ({ ok: true as const, fileName: 'budget.xlsx' }),
    }
    for (const arguments_ of [
      {
        sheet: 'Budget',
        range: 'A1',
        inputTitle: 'x'.repeat(33),
        inputMessage: null,
        errorStyle: 'none',
        errorTitle: null,
        errorMessage: null,
      },
      {
        sheet: 'Budget',
        range: 'A1',
        inputTitle: null,
        inputMessage: null,
        errorStyle: 'none',
        errorTitle: 'Must clear',
        errorMessage: null,
      },
    ]) {
      await expect(
        executeXlsxOperation(
          { operation: 'xlsx.range.set_validation_messages', arguments: arguments_ },
          services,
        ),
      ).resolves.toMatchObject({ handled: true, ok: false, error: 'invalid_arguments' })
    }
    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.range.set_validation_messages',
          arguments: {
            sheet: 'Budget',
            range: 'A1',
            inputTitle: null,
            inputMessage: null,
            errorStyle: 'stop',
            errorTitle: null,
            errorMessage: null,
          },
        },
        services,
      ),
    ).resolves.toMatchObject({ handled: true, ok: false, error: 'execution_failed' })
    expect(changed).toBe(false)
  })

  it('creates and updates one bounded comparison conditional-format rule through native history', async () => {
    const added: Array<Record<string, unknown>> = []
    const updated: Array<{ ruleId: string; rule: Record<string, unknown> }> = []
    const existing = {
      cfId: 'cf-existing',
      ranges: [{ startRow: 0, endRow: 4, startColumn: 0, endColumn: 0 }],
      stopIfTrue: false,
      rule: { type: 'highlightCell', subType: 'number' },
    }
    const makeBuilder = () => {
      const built = {
        cfId: `cf-${added.length + 1}`,
        ranges: [] as unknown[],
        stopIfTrue: false,
        rule: { type: 'highlightCell', subType: 'number' },
      }
      const builder = {
        whenNumberGreaterThan: (value: number) => {
          built.rule = { ...built.rule, operator: 'greaterThan', value } as never
          return builder
        },
        whenNumberBetween: (first: number, second: number) => {
          built.rule = { ...built.rule, operator: 'between', value: first, value2: second } as never
          return builder
        },
        setBackground: (color: string) => {
          built.rule = { ...built.rule, background: color } as never
          return builder
        },
        setFontColor: (color: string) => {
          built.rule = { ...built.rule, fontColor: color } as never
          return builder
        },
        setBold: (enabled: boolean) => {
          built.rule = { ...built.rule, bold: enabled } as never
          return builder
        },
        setItalic: (enabled: boolean) => {
          built.rule = { ...built.rule, italic: enabled } as never
          return builder
        },
        setRanges: (ranges: unknown[]) => {
          built.ranges = ranges
          return builder
        },
        build: () => built,
      }
      return builder
    }
    const range = {
      getRow: () => 0,
      getColumn: () => 0,
      getHeight: () => 5,
      getWidth: () => 1,
      getRange: () => ({ startRow: 0, endRow: 4, startColumn: 0, endColumn: 0 }),
      activate: () => undefined,
    }
    const worksheet = {
      getSheetName: () => 'Budget',
      getSheetId: () => 'sheet-budget',
      getRange: () => range,
      newConditionalFormattingRule: makeBuilder,
      getConditionalFormattingRules: () => [existing, ...added],
      addConditionalFormattingRule: (rule: Record<string, unknown>) => added.push(rule),
      setConditionalFormattingRule: (ruleId: string, rule: Record<string, unknown>) =>
        updated.push({ ruleId, rule }),
    }
    const services = {
      runtime: () =>
        ({
          univerAPI: {
            getActiveWorkbook: () => ({
              getSheets: () => [worksheet],
              setActiveSheet: () => undefined,
            }),
          },
        }) as never,
      state: () => ({
        editJournal: createEditJournal(),
        hyperlinkTargets: new Map(),
        sheetProtections: new Map(),
        appliedCfSheets: new Set(['sheet-budget']),
      }),
      loadStaged: async () => undefined,
      save: async () => ({ ok: true as const, fileName: 'budget.xlsx' }),
    }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.conditional_format.set_comparison',
          arguments: {
            sheet: 'Budget',
            range: 'A1:A5',
            ruleId: null,
            operator: 'greaterThan',
            operand1: 10,
            format: { fillColor: '#FFF2CC', bold: true },
            stopIfTrue: true,
          },
        },
        services,
      ),
    ).resolves.toMatchObject({
      handled: true,
      operationId: 'xlsx.conditional_format.set_comparison',
      ok: true,
      output: { sheet: 'Budget', range: 'A1:A5', ruleId: 'cf-1', created: true },
    })
    expect(added[0]).toMatchObject({ cfId: 'cf-1', stopIfTrue: true })

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.conditional_format.set_comparison',
          arguments: {
            sheet: 'Budget',
            range: 'B2:B6',
            ruleId: 'cf-existing',
            operator: 'between',
            operand1: 1,
            operand2: 9,
            format: { fontColor: '#9C0006', italic: true },
            stopIfTrue: false,
          },
        },
        services,
      ),
    ).resolves.toMatchObject({
      handled: true,
      ok: true,
      output: { ruleId: 'cf-existing', created: false, range: 'B2:B6' },
    })
    expect(updated).toEqual([
      {
        ruleId: 'cf-existing',
        rule: expect.objectContaining({ cfId: 'cf-existing', stopIfTrue: false }),
      },
    ])
  })

  it('removes only an explicitly identified conditional-format rule', async () => {
    const deleted: string[] = []
    const worksheet = {
      getSheetName: () => 'Budget',
      getSheetId: () => 'sheet-budget',
      getConditionalFormattingRules: () => [
        {
          cfId: 'cf-1',
          ranges: [{ startRow: 0, endRow: 4, startColumn: 0, endColumn: 0 }],
          rule: { type: 'highlightCell', subType: 'number' },
        },
      ],
      deleteConditionalFormattingRule: (ruleId: string) => deleted.push(ruleId),
    }
    const services = {
      runtime: () =>
        ({
          univerAPI: {
            getActiveWorkbook: () => ({
              getSheets: () => [worksheet],
              setActiveSheet: () => undefined,
            }),
          },
        }) as never,
      state: () => ({
        editJournal: createEditJournal(),
        hyperlinkTargets: new Map(),
        sheetProtections: new Map(),
        appliedCfSheets: new Set(['sheet-budget']),
      }),
      loadStaged: async () => undefined,
      save: async () => ({ ok: true as const, fileName: 'budget.xlsx' }),
    }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.conditional_format.remove',
          arguments: { sheet: 'Budget', ruleId: 'cf-1' },
        },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.conditional_format.remove',
      ok: true,
      output: { sheet: 'Budget', ruleId: 'cf-1', removed: true },
    })
    expect(deleted).toEqual(['cf-1'])

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.conditional_format.remove',
          arguments: { sheet: 'Budget', ruleId: 'missing' },
        },
        services,
      ),
    ).resolves.toMatchObject({ handled: true, ok: false, error: 'execution_failed' })
    expect(deleted).toEqual(['cf-1'])
  })

  it('rejects ambiguous comparison conditional formats before mutation', async () => {
    let mutated = false
    const worksheet = {
      getSheetName: () => 'Budget',
      getSheetId: () => 'sheet-budget',
      getRange: () => ({
        getRow: () => 0,
        getColumn: () => 0,
        getHeight: () => 1,
        getWidth: () => 1,
        getRange: () => ({ startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }),
      }),
      newConditionalFormattingRule: () => {
        mutated = true
        return {}
      },
      getConditionalFormattingRules: () => [],
    }
    const services = {
      runtime: () =>
        ({
          univerAPI: {
            getActiveWorkbook: () => ({
              getSheets: () => [worksheet],
              setActiveSheet: () => undefined,
            }),
          },
        }) as never,
      state: () => ({
        editJournal: createEditJournal(),
        hyperlinkTargets: new Map(),
        sheetProtections: new Map(),
        appliedCfSheets: new Set(['sheet-budget']),
      }),
      loadStaged: async () => undefined,
      save: async () => ({ ok: true as const, fileName: 'budget.xlsx' }),
    }
    const common = {
      sheet: 'Budget',
      range: 'A1',
      ruleId: null,
      format: { fillColor: '#FF0000' },
      stopIfTrue: false,
    }
    for (const arguments_ of [
      { ...common, operator: 'between', operand1: 1 },
      { ...common, operator: 'equal', operand1: 1, operand2: 2 },
      { ...common, operator: 'between', operand1: 9, operand2: 1 },
      { ...common, operator: 'equal', operand1: 1, format: {} },
      { ...common, operator: 'equal', operand1: 1, ruleId: 'x'.repeat(129) },
    ]) {
      await expect(
        executeXlsxOperation(
          { operation: 'xlsx.conditional_format.set_comparison', arguments: arguments_ },
          services,
        ),
      ).resolves.toMatchObject({ handled: true, ok: false, error: 'invalid_arguments' })
    }
    expect(mutated).toBe(false)
  })

  it('sets retained text, blank, and duplicate conditional-format predicates through one operation', async () => {
    const added: Array<Record<string, unknown>> = []
    const updated: string[] = []
    const existing = {
      cfId: 'cf-text',
      ranges: [{ startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }],
      rule: { type: 'highlightCell', subType: 'text' },
    }
    const makeBuilder = () => {
      const built = {
        cfId: `cf-${added.length + 1}`,
        ranges: [] as unknown[],
        rule: { type: 'highlightCell', subType: 'text' },
      }
      const predicate = (subType: string, operator?: string, value?: string) => {
        built.rule = {
          type: 'highlightCell',
          subType,
          ...(operator === undefined ? {} : { operator }),
          ...(value === undefined ? {} : { value }),
        } as never
        return builder
      }
      const builder = {
        whenTextContains: (value: string) => predicate('text', 'containsText', value),
        whenTextDoesNotContain: (value: string) => predicate('text', 'notContainsText', value),
        whenTextStartsWith: (value: string) => predicate('text', 'beginsWith', value),
        whenTextEndsWith: (value: string) => predicate('text', 'endsWith', value),
        whenTextEqualTo: (value: string) => predicate('text', 'equal', value),
        whenCellEmpty: () => predicate('text', 'containsBlanks'),
        whenCellNotEmpty: () => predicate('text', 'notContainsBlanks'),
        setDuplicateValues: () => predicate('duplicateValues'),
        setUniqueValues: () => predicate('uniqueValues'),
        setBackground: () => builder,
        setFontColor: () => builder,
        setBold: () => builder,
        setItalic: () => builder,
        setRanges: (ranges: unknown[]) => {
          built.ranges = ranges
          return builder
        },
        build: () => built,
      }
      return builder
    }
    const range = {
      getRow: () => 0,
      getColumn: () => 0,
      getHeight: () => 3,
      getWidth: () => 1,
      getRange: () => ({ startRow: 0, endRow: 2, startColumn: 0, endColumn: 0 }),
      activate: () => undefined,
    }
    const worksheet = {
      getSheetName: () => 'Budget',
      getSheetId: () => 'sheet-budget',
      getRange: () => range,
      newConditionalFormattingRule: makeBuilder,
      getConditionalFormattingRules: () => [existing, ...added],
      addConditionalFormattingRule: (rule: Record<string, unknown>) => added.push(rule),
      setConditionalFormattingRule: (ruleId: string) => updated.push(ruleId),
    }
    const services = {
      runtime: () =>
        ({
          univerAPI: {
            getActiveWorkbook: () => ({
              getSheets: () => [worksheet],
              setActiveSheet: () => undefined,
            }),
          },
        }) as never,
      state: () => ({
        editJournal: createEditJournal(),
        hyperlinkTargets: new Map(),
        sheetProtections: new Map(),
        appliedCfSheets: new Set(['sheet-budget']),
      }),
      loadStaged: async () => undefined,
      save: async () => ({ ok: true as const, fileName: 'budget.xlsx' }),
    }
    const predicates = [
      ['textContains', 'Approved'],
      ['textNotContains', 'Draft'],
      ['textStartsWith', 'FY'],
      ['textEndsWith', '.00'],
      ['textEqual', 'Approved'],
      ['textNotEqual', 'Draft'],
      ['blank', null],
      ['nonBlank', null],
      ['error', null],
      ['nonError', null],
      ['duplicate', null],
      ['unique', null],
    ] as const
    for (const [predicate, text] of predicates) {
      await expect(
        executeXlsxOperation(
          {
            operation: 'xlsx.conditional_format.set_highlight',
            arguments: {
              sheet: 'Budget',
              range: 'A1:A3',
              ruleId: null,
              predicate,
              text,
              format: { fillColor: '#FCE4D6' },
              stopIfTrue: false,
            },
          },
          services,
        ),
      ).resolves.toMatchObject({
        handled: true,
        operationId: 'xlsx.conditional_format.set_highlight',
        ok: true,
        output: { predicate, text, created: true },
      })
    }
    expect(added).toHaveLength(12)

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.conditional_format.set_highlight',
          arguments: {
            sheet: 'Budget',
            range: 'B1:B3',
            ruleId: 'cf-text',
            predicate: 'unique',
            text: null,
            format: { fontColor: '#006100' },
            stopIfTrue: true,
          },
        },
        services,
      ),
    ).resolves.toMatchObject({ ok: true, output: { ruleId: 'cf-text', created: false } })
    expect(updated).toEqual(['cf-text'])
  })

  it('rejects highlight conditional-format predicate/text mismatches before mutation', async () => {
    let mutated = false
    const worksheet = {
      getSheetName: () => 'Budget',
      getSheetId: () => 'sheet-budget',
      getRange: () => ({
        getRow: () => 0,
        getColumn: () => 0,
        getHeight: () => 1,
        getWidth: () => 1,
      }),
      newConditionalFormattingRule: () => {
        mutated = true
        return {}
      },
    }
    const services = {
      runtime: () =>
        ({
          univerAPI: {
            getActiveWorkbook: () => ({
              getSheets: () => [worksheet],
              setActiveSheet: () => undefined,
            }),
          },
        }) as never,
      state: () => ({
        editJournal: createEditJournal(),
        hyperlinkTargets: new Map(),
        sheetProtections: new Map(),
        appliedCfSheets: new Set(['sheet-budget']),
      }),
      loadStaged: async () => undefined,
      save: async () => ({ ok: true as const, fileName: 'budget.xlsx' }),
    }
    const common = {
      sheet: 'Budget',
      range: 'A1',
      ruleId: null,
      format: { fillColor: '#FF0000' },
      stopIfTrue: false,
    }
    for (const arguments_ of [
      { ...common, predicate: 'textContains', text: null },
      { ...common, predicate: 'textContains', text: '' },
      { ...common, predicate: 'textContains', text: 'x'.repeat(256) },
      { ...common, predicate: 'duplicate', text: 'unexpected' },
    ]) {
      await expect(
        executeXlsxOperation(
          { operation: 'xlsx.conditional_format.set_highlight', arguments: arguments_ },
          services,
        ),
      ).resolves.toMatchObject({ handled: true, ok: false, error: 'invalid_arguments' })
    }
    expect(mutated).toBe(false)
  })

  it('sets rank and average conditional formats through one explicit statistical contract', async () => {
    const added: Array<Record<string, unknown>> = []
    const makeBuilder = () => {
      const built = {
        cfId: `cf-${added.length + 1}`,
        ranges: [] as unknown[],
        rule: { type: 'highlightCell', subType: 'rank' } as Record<string, unknown>,
      }
      const builder = {
        setRank: (config: Record<string, unknown>) => {
          built.rule = { type: 'highlightCell', subType: 'rank', ...config }
          return builder
        },
        setAverage: (operator: string) => {
          built.rule = { type: 'highlightCell', subType: 'average', operator }
          return builder
        },
        setBackground: () => builder,
        setFontColor: () => builder,
        setBold: () => builder,
        setItalic: () => builder,
        setRanges: (ranges: unknown[]) => {
          built.ranges = ranges
          return builder
        },
        build: () => built,
      }
      return builder
    }
    const worksheet = {
      getSheetName: () => 'Budget',
      getSheetId: () => 'sheet-budget',
      getRange: () => ({
        getRow: () => 0,
        getColumn: () => 0,
        getHeight: () => 10,
        getWidth: () => 1,
        getRange: () => ({ startRow: 0, endRow: 9, startColumn: 0, endColumn: 0 }),
        activate: () => undefined,
      }),
      newConditionalFormattingRule: makeBuilder,
      getConditionalFormattingRules: () => added,
      addConditionalFormattingRule: (rule: Record<string, unknown>) => added.push(rule),
    }
    const services = {
      runtime: () =>
        ({
          univerAPI: {
            getActiveWorkbook: () => ({
              getSheets: () => [worksheet],
              setActiveSheet: () => undefined,
            }),
          },
        }) as never,
      state: () => ({
        editJournal: createEditJournal(),
        hyperlinkTargets: new Map(),
        sheetProtections: new Map(),
        appliedCfSheets: new Set(['sheet-budget']),
      }),
      loadStaged: async () => undefined,
      save: async () => ({ ok: true as const, fileName: 'budget.xlsx' }),
    }
    const rank = {
      sheet: 'Budget',
      range: 'A1:A10',
      ruleId: null,
      kind: 'rank',
      direction: 'bottom',
      rank: 10,
      percent: true,
      inclusive: null,
      format: { fillColor: '#FFC7CE' },
      stopIfTrue: false,
    }
    const average = {
      sheet: 'Budget',
      range: 'A1:A10',
      ruleId: null,
      kind: 'average',
      direction: 'above',
      rank: null,
      percent: null,
      inclusive: true,
      format: { fontColor: '#006100' },
      stopIfTrue: true,
    }
    for (const arguments_ of [rank, average]) {
      await expect(
        executeXlsxOperation(
          { operation: 'xlsx.conditional_format.set_statistical', arguments: arguments_ },
          services,
        ),
      ).resolves.toMatchObject({
        handled: true,
        operationId: 'xlsx.conditional_format.set_statistical',
        ok: true,
        output: {
          kind: arguments_.kind,
          direction: arguments_.direction,
          rank: arguments_.rank,
          percent: arguments_.percent,
          inclusive: arguments_.inclusive,
          created: true,
        },
      })
    }
    expect(added.map((rule) => rule.rule)).toEqual([
      {
        type: 'highlightCell',
        subType: 'rank',
        isBottom: true,
        isPercent: true,
        value: 10,
      },
      { type: 'highlightCell', subType: 'average', operator: 'greaterThanOrEqual' },
    ])
  })

  it('rejects inconsistent statistical conditional-format fields before mutation', async () => {
    const services = {
      runtime: () => null,
      loadStaged: async () => undefined,
      save: async () => ({ ok: true as const, fileName: 'budget.xlsx' }),
    }
    const common = {
      sheet: 'Budget',
      range: 'A1',
      ruleId: null,
      format: { fillColor: '#FF0000' },
      stopIfTrue: false,
    }
    for (const arguments_ of [
      {
        ...common,
        kind: 'rank',
        direction: 'top',
        rank: null,
        percent: false,
        inclusive: null,
      },
      {
        ...common,
        kind: 'rank',
        direction: 'above',
        rank: 10,
        percent: false,
        inclusive: null,
      },
      {
        ...common,
        kind: 'average',
        direction: 'below',
        rank: 10,
        percent: null,
        inclusive: false,
      },
    ]) {
      await expect(
        executeXlsxOperation(
          { operation: 'xlsx.conditional_format.set_statistical', arguments: arguments_ },
          services,
        ),
      ).resolves.toMatchObject({ handled: true, ok: false, error: 'invalid_arguments' })
    }
  })

  it('sets a bounded custom-formula conditional format through the shared lifecycle', async () => {
    const added: Array<Record<string, unknown>> = []
    const makeBuilder = () => {
      const built = {
        cfId: 'cf-formula',
        ranges: [] as unknown[],
        rule: {} as Record<string, unknown>,
      }
      const builder = {
        whenFormulaSatisfied: (formula: string) => {
          built.rule = { type: 'highlightCell', subType: 'formula', value: formula }
          return builder
        },
        setBackground: () => builder,
        setFontColor: () => builder,
        setBold: () => builder,
        setItalic: () => builder,
        setRanges: (ranges: unknown[]) => {
          built.ranges = ranges
          return builder
        },
        build: () => built,
      }
      return builder
    }
    const worksheet = {
      getSheetName: () => 'Budget',
      getSheetId: () => 'sheet-budget',
      getRange: () => ({
        getRow: () => 0,
        getColumn: () => 0,
        getHeight: () => 10,
        getWidth: () => 2,
        getRange: () => ({ startRow: 0, endRow: 9, startColumn: 0, endColumn: 1 }),
        activate: () => undefined,
      }),
      newConditionalFormattingRule: makeBuilder,
      getConditionalFormattingRules: () => added,
      addConditionalFormattingRule: (rule: Record<string, unknown>) => added.push(rule),
    }
    const services = {
      runtime: () =>
        ({
          univerAPI: {
            getActiveWorkbook: () => ({
              getSheets: () => [worksheet],
              setActiveSheet: () => undefined,
            }),
          },
        }) as never,
      state: () => ({
        editJournal: createEditJournal(),
        hyperlinkTargets: new Map(),
        sheetProtections: new Map(),
        appliedCfSheets: new Set(['sheet-budget']),
      }),
      loadStaged: async () => undefined,
      save: async () => ({ ok: true as const, fileName: 'budget.xlsx' }),
    }
    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.conditional_format.set_formula',
          arguments: {
            sheet: 'Budget',
            range: 'A1:B10',
            ruleId: null,
            formula: '=$B1>100',
            format: { fillColor: '#C6EFCE' },
            stopIfTrue: true,
          },
        },
        services,
      ),
    ).resolves.toMatchObject({
      handled: true,
      operationId: 'xlsx.conditional_format.set_formula',
      ok: true,
      output: {
        ruleId: 'cf-formula',
        formula: '=$B1>100',
        created: true,
        stopIfTrue: true,
      },
    })
    expect(added[0]?.rule).toEqual({
      type: 'highlightCell',
      subType: 'formula',
      value: '=$B1>100',
    })
  })

  it('rejects malformed custom-formula conditional formats before workbook access', async () => {
    const services = {
      runtime: () => null,
      loadStaged: async () => undefined,
      save: async () => ({ ok: true as const, fileName: 'budget.xlsx' }),
    }
    for (const formula of ['A1>0', '=', `=${'A'.repeat(8_192)}`]) {
      await expect(
        executeXlsxOperation(
          {
            operation: 'xlsx.conditional_format.set_formula',
            arguments: {
              sheet: 'Budget',
              range: 'A1',
              ruleId: null,
              formula,
              format: { fillColor: '#FF0000' },
              stopIfTrue: false,
            },
          },
          services,
        ),
      ).resolves.toMatchObject({ handled: true, ok: false, error: 'invalid_arguments' })
    }
  })

  it('sets color-scale, data-bar, and saveable icon-set rules through one visual contract', async () => {
    const added: Array<Record<string, unknown>> = []
    const makeBuilder = () => {
      const built = {
        cfId: `cf-${added.length + 1}`,
        ranges: [] as unknown[],
        rule: {} as Record<string, unknown>,
      }
      const builder = {
        setColorScale: (config: unknown) => {
          built.rule = { type: 'colorScale', config }
          return builder
        },
        setDataBar: (config: unknown) => {
          built.rule = { type: 'dataBar', config }
          return builder
        },
        setIconSet: (config: { iconConfigs: unknown; isShowValue: boolean }) => {
          built.rule = {
            type: 'iconSet',
            config: config.iconConfigs,
            isShowValue: config.isShowValue,
          }
          return builder
        },
        setRanges: (ranges: unknown[]) => {
          built.ranges = ranges
          return builder
        },
        build: () => built,
      }
      return builder
    }
    const worksheet = {
      getSheetName: () => 'Budget',
      getSheetId: () => 'sheet-budget',
      getRange: () => ({
        getRow: () => 0,
        getColumn: () => 0,
        getHeight: () => 10,
        getWidth: () => 1,
        getRange: () => ({ startRow: 0, endRow: 9, startColumn: 0, endColumn: 0 }),
        activate: () => undefined,
      }),
      newConditionalFormattingRule: makeBuilder,
      getConditionalFormattingRules: () => added,
      addConditionalFormattingRule: (rule: Record<string, unknown>) => added.push(rule),
    }
    const services = {
      runtime: () =>
        ({
          univerAPI: {
            getActiveWorkbook: () => ({
              getSheets: () => [worksheet],
              setActiveSheet: () => undefined,
            }),
          },
        }) as never,
      state: () => ({
        editJournal: createEditJournal(),
        hyperlinkTargets: new Map(),
        sheetProtections: new Map(),
        appliedCfSheets: new Set(['sheet-budget']),
      }),
      loadStaged: async () => undefined,
      save: async () => ({ ok: true as const, fileName: 'budget.xlsx' }),
    }
    const threshold = (type: string, value: number | string | null) => ({
      type,
      value,
      inclusive: null,
    })
    const cases = [
      {
        kind: 'colorScale',
        colors: ['#F8696B', '#FFEB84', '#63BE7B'],
        thresholds: [threshold('min', null), threshold('percentile', 50), threshold('max', null)],
        iconSet: null,
        showValue: null,
        reverse: null,
        gradient: null,
      },
      {
        kind: 'dataBar',
        colors: ['#638EC6'],
        thresholds: [threshold('min', null), threshold('max', null)],
        iconSet: null,
        showValue: false,
        reverse: null,
        gradient: true,
      },
      {
        kind: 'iconSet',
        colors: [],
        thresholds: [
          { type: 'percent', value: 0, inclusive: true },
          { type: 'percent', value: 33, inclusive: true },
          { type: 'percent', value: 67, inclusive: true },
        ],
        iconSet: '3TrafficLights1',
        showValue: true,
        reverse: false,
        gradient: null,
      },
    ]
    for (const visual of cases) {
      await expect(
        executeXlsxOperation(
          {
            operation: 'xlsx.conditional_format.set_visual',
            arguments: {
              sheet: 'Budget',
              range: 'A1:A10',
              ruleId: null,
              ...visual,
              stopIfTrue: false,
            },
          },
          services,
        ),
      ).resolves.toMatchObject({
        handled: true,
        operationId: 'xlsx.conditional_format.set_visual',
        ok: true,
        output: { kind: visual.kind, created: true },
      })
    }
    expect(added.map((rule) => (rule.rule as { type: string }).type)).toEqual([
      'colorScale',
      'dataBar',
      'iconSet',
    ])
  })

  it('rejects lossy or incoherent visual conditional-format states before workbook access', async () => {
    const services = {
      runtime: () => null,
      loadStaged: async () => undefined,
      save: async () => ({ ok: true as const, fileName: 'budget.xlsx' }),
    }
    const common = {
      sheet: 'Budget',
      range: 'A1',
      ruleId: null,
      stopIfTrue: false,
    }
    for (const visual of [
      {
        kind: 'colorScale',
        colors: ['#FF0000'],
        thresholds: [{ type: 'min', value: null, inclusive: null }],
        iconSet: null,
        showValue: null,
        reverse: null,
        gradient: null,
      },
      {
        kind: 'dataBar',
        colors: ['#638EC6'],
        thresholds: [
          { type: 'min', value: null, inclusive: null },
          { type: 'max', value: null, inclusive: null },
        ],
        iconSet: null,
        showValue: true,
        reverse: null,
        gradient: false,
      },
      {
        kind: 'iconSet',
        colors: [],
        thresholds: [
          { type: 'percent', value: 0, inclusive: true },
          { type: 'percent', value: 50, inclusive: true },
        ],
        iconSet: '3Stars',
        showValue: true,
        reverse: false,
        gradient: null,
      },
    ]) {
      await expect(
        executeXlsxOperation(
          {
            operation: 'xlsx.conditional_format.set_visual',
            arguments: { ...common, ...visual },
          },
          services,
        ),
      ).resolves.toMatchObject({ handled: true, ok: false, error: 'invalid_arguments' })
    }
  })

  it('clears conditional formats by explicit range or worksheet scope', async () => {
    let rangeClears = 0
    let sheetClears = 0
    const range = {
      getRow: () => 0,
      getColumn: () => 0,
      getHeight: () => 5,
      getWidth: () => 1,
      getConditionalFormattingRules: () => [{ cfId: 'cf-1' }, { cfId: 'cf-2' }],
      clearConditionalFormatRules: () => {
        rangeClears += 1
      },
      activate: () => undefined,
    }
    const worksheet = {
      getSheetName: () => 'Budget',
      getSheetId: () => 'sheet-budget',
      getRange: () => range,
      getConditionalFormattingRules: () => [{ cfId: 'cf-1' }, { cfId: 'cf-2' }, { cfId: 'cf-3' }],
      clearConditionalFormatRules: () => {
        sheetClears += 1
      },
    }
    const services = {
      runtime: () =>
        ({
          univerAPI: {
            getActiveWorkbook: () => ({
              getSheets: () => [worksheet],
              setActiveSheet: () => undefined,
            }),
          },
        }) as never,
      state: () => ({
        editJournal: createEditJournal(),
        hyperlinkTargets: new Map(),
        sheetProtections: new Map(),
        appliedCfSheets: new Set(['sheet-budget']),
      }),
      loadStaged: async () => undefined,
      save: async () => ({ ok: true as const, fileName: 'budget.xlsx' }),
    }
    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.conditional_format.clear',
          arguments: { sheet: 'Budget', scope: 'range', range: 'A1:A5' },
        },
        services,
      ),
    ).resolves.toMatchObject({
      handled: true,
      operationId: 'xlsx.conditional_format.clear',
      ok: true,
      output: { sheet: 'Budget', scope: 'range', range: 'A1:A5', cleared: 2 },
    })
    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.conditional_format.clear',
          arguments: { sheet: 'Budget', scope: 'sheet', range: null },
        },
        services,
      ),
    ).resolves.toMatchObject({
      handled: true,
      ok: true,
      output: { sheet: 'Budget', scope: 'sheet', range: null, cleared: 3 },
    })
    expect({ rangeClears, sheetClears }).toEqual({ rangeClears: 1, sheetClears: 1 })
  })

  it('rejects inconsistent conditional-format clear scope before workbook access', async () => {
    const services = {
      runtime: () => null,
      loadStaged: async () => undefined,
      save: async () => ({ ok: true as const, fileName: 'budget.xlsx' }),
    }
    for (const arguments_ of [
      { sheet: 'Budget', scope: 'range', range: null },
      { sheet: 'Budget', scope: 'sheet', range: 'A1' },
    ]) {
      await expect(
        executeXlsxOperation(
          { operation: 'xlsx.conditional_format.clear', arguments: arguments_ },
          services,
        ),
      ).resolves.toMatchObject({ handled: true, ok: false, error: 'invalid_arguments' })
    }
  })

  it('sets conditional-format priority by explicit one-based final position', async () => {
    const rules = ['cf-a', 'cf-b', 'cf-c'].map((cfId) => ({ cfId }))
    const moves: Array<{ ruleId: string; targetRuleId: string; placement: string }> = []
    const worksheet = {
      getSheetName: () => 'Budget',
      getSheetId: () => 'sheet-budget',
      getConditionalFormattingRules: () => rules,
      moveConditionalFormattingRule: (
        ruleId: string,
        targetRuleId: string,
        placement: 'before' | 'after',
      ) => {
        moves.push({ ruleId, targetRuleId, placement })
        const movingIndex = rules.findIndex((rule) => rule.cfId === ruleId)
        const [moving] = rules.splice(movingIndex, 1)
        const targetIndex = rules.findIndex((rule) => rule.cfId === targetRuleId)
        rules.splice(targetIndex + (placement === 'after' ? 1 : 0), 0, moving!)
      },
    }
    const services = {
      runtime: () =>
        ({
          univerAPI: {
            getActiveWorkbook: () => ({
              getSheets: () => [worksheet],
              setActiveSheet: () => undefined,
            }),
          },
        }) as never,
      state: () => ({
        editJournal: createEditJournal(),
        hyperlinkTargets: new Map(),
        sheetProtections: new Map(),
        appliedCfSheets: new Set(['sheet-budget']),
      }),
      loadStaged: async () => undefined,
      save: async () => ({ ok: true as const, fileName: 'budget.xlsx' }),
    }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.conditional_format.set_priority',
          arguments: { sheet: 'Budget', ruleId: 'cf-c', position: 1 },
        },
        services,
      ),
    ).resolves.toMatchObject({
      handled: true,
      ok: true,
      output: { sheet: 'Budget', ruleId: 'cf-c', position: 1, moved: true },
    })
    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.conditional_format.set_priority',
          arguments: { sheet: 'Budget', ruleId: 'cf-a', position: 2 },
        },
        services,
      ),
    ).resolves.toMatchObject({ ok: true, output: { position: 2, moved: false } })
    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.conditional_format.set_priority',
          arguments: { sheet: 'Budget', ruleId: 'cf-c', position: 3 },
        },
        services,
      ),
    ).resolves.toMatchObject({ ok: true, output: { position: 3, moved: true } })
    expect(moves).toEqual([
      { ruleId: 'cf-c', targetRuleId: 'cf-a', placement: 'before' },
      { ruleId: 'cf-c', targetRuleId: 'cf-b', placement: 'after' },
    ])
  })

  it('rejects missing rules and out-of-bounds conditional-format positions', async () => {
    const worksheet = {
      getSheetName: () => 'Budget',
      getSheetId: () => 'sheet-budget',
      getConditionalFormattingRules: () => [{ cfId: 'cf-a' }, { cfId: 'cf-b' }],
      moveConditionalFormattingRule: () => {
        throw new Error('must not mutate')
      },
    }
    const services = {
      runtime: () =>
        ({
          univerAPI: {
            getActiveWorkbook: () => ({
              getSheets: () => [worksheet],
              setActiveSheet: () => undefined,
            }),
          },
        }) as never,
      state: () => ({
        editJournal: createEditJournal(),
        hyperlinkTargets: new Map(),
        sheetProtections: new Map(),
        appliedCfSheets: new Set(['sheet-budget']),
      }),
      loadStaged: async () => undefined,
      save: async () => ({ ok: true as const, fileName: 'budget.xlsx' }),
    }
    for (const arguments_ of [
      { sheet: 'Budget', ruleId: 'missing', position: 1 },
      { sheet: 'Budget', ruleId: 'cf-a', position: 3 },
    ]) {
      await expect(
        executeXlsxOperation(
          { operation: 'xlsx.conditional_format.set_priority', arguments: arguments_ },
          services,
        ),
      ).resolves.toMatchObject({ handled: true, ok: false, error: 'invalid_arguments' })
    }
  })

  it('sets explicit cell-protection fields through xlsx.range.set_protection', async () => {
    const journal = createEditJournal()
    let activeRange = false
    const worksheet = {
      getSheetName: () => 'Budget',
      getSheetId: () => 'sheet-budget',
      getRange: () => ({
        getRow: () => 1,
        getColumn: () => 2,
        getHeight: () => 2,
        getWidth: () => 1,
        activate: () => {
          activeRange = true
        },
      }),
    }
    const workbook = {
      getSheets: () => [worksheet],
      setActiveSheet: () => undefined,
    }
    const runtime = { univerAPI: { getActiveWorkbook: () => workbook } }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.range.set_protection',
          arguments: {
            sheet: 'Budget',
            range: 'c2:c3',
            protection: { locked: false, hidden: true },
            fields: ['locked', 'hidden'],
          },
        },
        {
          runtime: () => runtime,
          state: () => ({
            editJournal: journal,
            hyperlinkTargets: new Map(),
            sheetProtections: new Map(),
          }),
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.range.set_protection',
      ok: true,
      output: {
        sheet: 'Budget',
        range: 'C2:C3',
        fields: ['locked', 'hidden'],
      },
    })
    expect(journal.cells.get('sheet-budget')?.get('1:2')?.style).toEqual({
      protectionLocked: false,
      protectionHidden: true,
    })
    expect(journal.cells.get('sheet-budget')?.get('2:2')?.style).toEqual({
      protectionLocked: false,
      protectionHidden: true,
    })
    expect(activeRange).toBe(true)
  })

  it('rejects xlsx.range.set_protection above the 10,000-cell save boundary', async () => {
    const journal = createEditJournal()
    const worksheet = {
      getSheetName: () => 'Budget',
      getSheetId: () => 'sheet-budget',
      getRange: () => ({
        getHeight: () => 101,
        getWidth: () => 100,
      }),
    }
    const workbook = {
      getSheets: () => [worksheet],
      setActiveSheet: () => undefined,
    }
    const runtime = { univerAPI: { getActiveWorkbook: () => workbook } }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.range.set_protection',
          arguments: {
            sheet: 'Budget',
            range: 'A1:CV101',
            protection: { locked: true },
            fields: ['locked'],
          },
        },
        {
          runtime: () => runtime,
          state: () => ({
            editJournal: journal,
            hyperlinkTargets: new Map(),
            sheetProtections: new Map(),
          }),
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.range.set_protection',
      ok: false,
      error: 'invalid_arguments',
      message: 'xlsx.range.set_protection supports at most 10,000 cells per operation.',
    })
    expect(journal.cells.size).toBe(0)
  })

  it('sets explicit worksheet protection through xlsx.sheet.set_protection', async () => {
    const journal = createEditJournal()
    let pendingEdits = 0
    const worksheet = {
      getSheetName: () => 'Budget',
      getSheetId: () => 'sheet-budget',
    }
    const workbook = {
      getSheets: () => [worksheet],
      setActiveSheet: () => undefined,
    }
    const runtime = { univerAPI: { getActiveWorkbook: () => workbook } }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.sheet.set_protection',
          arguments: { sheet: 'Budget', protected: true },
        },
        {
          runtime: () => runtime,
          state: () => ({
            editJournal: journal,
            hyperlinkTargets: new Map(),
            sheetProtections: new Map([['sheet-budget', { protected: false, hasPassword: false }]]),
          }),
          setPendingEdits: (count: number) => {
            pendingEdits = count
          },
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.sheet.set_protection',
      ok: true,
      output: { sheet: 'Budget', protected: true },
    })
    expect(journal.sheetProtection.get('sheet-budget')).toBe(true)
    expect(pendingEdits).toBe(1)
  })

  it('sets italic explicitly through xlsx.range.set_text_style as one range mutation', async () => {
    let written: unknown = null
    let activeRange: string | null = null
    const worksheet = {
      getSheetName: () => 'Budget',
      getRange: (address: string) => ({
        setValue: (value: unknown) => {
          written = value
        },
        activate: () => {
          activeRange = address
        },
      }),
    }
    const workbook = {
      getSheets: () => [worksheet],
      setActiveSheet: () => undefined,
    }
    const runtime = { univerAPI: { getActiveWorkbook: () => workbook } }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.range.set_text_style',
          arguments: {
            sheet: 'Budget',
            range: 'b2:c3',
            style: { italic: true },
            fields: ['italic'],
          },
        },
        {
          runtime: () => runtime,
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.range.set_text_style',
      ok: true,
      output: { sheet: 'Budget', range: 'B2:C3', fields: ['italic'] },
    })
    expect(written).toEqual({ s: { it: BooleanNumber.TRUE } })
    expect(activeRange).toBe('b2:c3')
  })

  it('clears bold explicitly through xlsx.range.set_text_style', async () => {
    let written: unknown = null
    const worksheet = {
      getSheetName: () => 'Budget',
      getRange: () => ({
        setValue: (value: unknown) => {
          written = value
        },
        activate: () => undefined,
      }),
    }
    const workbook = {
      getSheets: () => [worksheet],
      setActiveSheet: () => undefined,
    }
    const runtime = { univerAPI: { getActiveWorkbook: () => workbook } }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.range.set_text_style',
          arguments: {
            sheet: 'Budget',
            range: 'A1',
            style: { bold: false },
            fields: ['bold'],
          },
        },
        {
          runtime: () => runtime,
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.range.set_text_style',
      ok: true,
      output: { sheet: 'Budget', range: 'A1', fields: ['bold'] },
    })
    expect(written).toEqual({ s: { bl: null } })
  })

  it('sets single underline explicitly through xlsx.range.set_text_style', async () => {
    let written: unknown = null
    const worksheet = {
      getSheetName: () => 'Budget',
      getRange: () => ({
        setValue: (value: unknown) => {
          written = value
        },
        activate: () => undefined,
      }),
    }
    const workbook = {
      getSheets: () => [worksheet],
      setActiveSheet: () => undefined,
    }
    const runtime = { univerAPI: { getActiveWorkbook: () => workbook } }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.range.set_text_style',
          arguments: {
            sheet: 'Budget',
            range: 'A1:B2',
            style: { underline: 'single' },
            fields: ['underline'],
          },
        },
        {
          runtime: () => runtime,
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.range.set_text_style',
      ok: true,
      output: { sheet: 'Budget', range: 'A1:B2', fields: ['underline'] },
    })
    expect(written).toEqual({ s: { ul: { s: BooleanNumber.TRUE } } })
  })

  it('clears underline explicitly through xlsx.range.set_text_style', async () => {
    let written: unknown = null
    const worksheet = {
      getSheetName: () => 'Budget',
      getRange: () => ({
        setValue: (value: unknown) => {
          written = value
        },
        activate: () => undefined,
      }),
    }
    const workbook = {
      getSheets: () => [worksheet],
      setActiveSheet: () => undefined,
    }
    const runtime = { univerAPI: { getActiveWorkbook: () => workbook } }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.range.set_text_style',
          arguments: {
            sheet: 'Budget',
            range: 'A1',
            style: { underline: 'none' },
            fields: ['underline'],
          },
        },
        {
          runtime: () => runtime,
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.range.set_text_style',
      ok: true,
      output: { sheet: 'Budget', range: 'A1', fields: ['underline'] },
    })
    expect(written).toEqual({ s: { ul: null } })
  })

  it('sets double underline explicitly through xlsx.range.set_text_style', async () => {
    let written: unknown = null
    const worksheet = {
      getSheetName: () => 'Budget',
      getRange: () => ({
        setValue: (value: unknown) => {
          written = value
        },
        activate: () => undefined,
      }),
    }
    const workbook = {
      getSheets: () => [worksheet],
      setActiveSheet: () => undefined,
    }
    const runtime = { univerAPI: { getActiveWorkbook: () => workbook } }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.range.set_text_style',
          arguments: {
            sheet: 'Budget',
            range: 'A1',
            style: { underline: 'double' },
            fields: ['underline'],
          },
        },
        {
          runtime: () => runtime,
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.range.set_text_style',
      ok: true,
      output: { sheet: 'Budget', range: 'A1', fields: ['underline'] },
    })
    expect(written).toEqual({ s: { ul: { s: BooleanNumber.TRUE, t: 10 } } })
  })

  it('sets strike explicitly through xlsx.range.set_text_style', async () => {
    let written: unknown = null
    const worksheet = {
      getSheetName: () => 'Budget',
      getRange: () => ({
        setValue: (value: unknown) => {
          written = value
        },
        activate: () => undefined,
      }),
    }
    const workbook = {
      getSheets: () => [worksheet],
      setActiveSheet: () => undefined,
    }
    const runtime = { univerAPI: { getActiveWorkbook: () => workbook } }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.range.set_text_style',
          arguments: {
            sheet: 'Budget',
            range: 'A1',
            style: { strike: true },
            fields: ['strike'],
          },
        },
        {
          runtime: () => runtime,
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.range.set_text_style',
      ok: true,
      output: { sheet: 'Budget', range: 'A1', fields: ['strike'] },
    })
    expect(written).toEqual({ s: { st: { s: BooleanNumber.TRUE } } })
  })

  it('sets multiple text style fields through one atomic range mutation', async () => {
    const writes: unknown[] = []
    const worksheet = {
      getSheetName: () => 'Budget',
      getRange: () => ({
        setValue: (value: unknown) => {
          writes.push(value)
        },
        activate: () => undefined,
      }),
    }
    const workbook = {
      getSheets: () => [worksheet],
      setActiveSheet: () => undefined,
    }
    const runtime = { univerAPI: { getActiveWorkbook: () => workbook } }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.range.set_text_style',
          arguments: {
            sheet: 'Budget',
            range: 'C3:D4',
            style: {
              bold: true,
              italic: false,
              underline: 'double',
              strike: true,
            },
            fields: ['bold', 'italic', 'underline', 'strike'],
          },
        },
        {
          runtime: () => runtime,
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.range.set_text_style',
      ok: true,
      output: {
        sheet: 'Budget',
        range: 'C3:D4',
        fields: ['bold', 'italic', 'underline', 'strike'],
      },
    })
    expect(writes).toEqual([
      {
        s: {
          bl: BooleanNumber.TRUE,
          it: null,
          ul: { s: BooleanNumber.TRUE, t: 10 },
          st: { s: BooleanNumber.TRUE },
        },
      },
    ])
  })

  it('rejects text style values that are not covered by the declared field mask', async () => {
    let mutated = false
    const worksheet = {
      getSheetName: () => 'Budget',
      getRange: () => ({
        setValue: () => {
          mutated = true
        },
        activate: () => undefined,
      }),
    }
    const workbook = {
      getSheets: () => [worksheet],
      setActiveSheet: () => undefined,
    }
    const runtime = { univerAPI: { getActiveWorkbook: () => workbook } }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.range.set_text_style',
          arguments: {
            sheet: 'Budget',
            range: 'A1',
            style: { bold: true, italic: true },
            fields: ['bold'],
          },
        },
        {
          runtime: () => runtime,
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.range.set_text_style',
      ok: false,
      error: 'invalid_arguments',
      message: 'xlsx.range.set_text_style requires unique fields with matching explicit values.',
    })
    expect(mutated).toBe(false)
  })

  it('enables an explicit XLSX filter range through the native worksheet command', async () => {
    const commands: Array<{ id: string; params: unknown }> = []
    const range = {
      getHeight: () => 4,
      getRange: () => ({ startRow: 0, endRow: 3, startColumn: 0, endColumn: 2 }),
    }
    const worksheet = {
      getSheetName: () => 'Budget',
      getSheetId: () => 'sheet-1',
      getFilter: () => null,
      getRange: () => range,
    }
    const workbook = {
      getId: () => 'workbook-1',
      getSheets: () => [worksheet],
      setActiveSheet: () => undefined,
    }
    const runtime = {
      univerAPI: {
        getActiveWorkbook: () => workbook,
        executeCommand: async (id: string, params: unknown) => {
          commands.push({ id, params })
          return true
        },
      },
    }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.range.set_filter',
          arguments: { sheet: 'Budget', range: 'a1:c4', enabled: true },
        },
        {
          runtime: () => runtime,
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.range.set_filter',
      ok: true,
      output: { sheet: 'Budget', range: 'A1:C4', enabled: true },
    })
    expect(commands).toEqual([
      {
        id: 'sheet.command.set-filter-range',
        params: {
          unitId: 'workbook-1',
          subUnitId: 'sheet-1',
          range: { startRow: 0, endRow: 3, startColumn: 0, endColumn: 2 },
        },
      },
    ])
  })

  it('disables the matching XLSX filter range through the native worksheet command', async () => {
    const commands: Array<{ id: string; params: unknown }> = []
    const bounds = { startRow: 0, endRow: 3, startColumn: 0, endColumn: 2 }
    const range = {
      getHeight: () => 4,
      getRange: () => bounds,
    }
    const worksheet = {
      getSheetName: () => 'Budget',
      getSheetId: () => 'sheet-1',
      getFilter: () => ({ getRange: () => range }),
      getRange: () => range,
    }
    const workbook = {
      getId: () => 'workbook-1',
      getSheets: () => [worksheet],
      setActiveSheet: () => undefined,
    }
    const runtime = {
      univerAPI: {
        getActiveWorkbook: () => workbook,
        executeCommand: async (id: string, params: unknown) => {
          commands.push({ id, params })
          return true
        },
      },
    }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.range.set_filter',
          arguments: { sheet: 'Budget', range: 'A1:C4', enabled: false },
        },
        {
          runtime: () => runtime,
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.range.set_filter',
      ok: true,
      output: { sheet: 'Budget', range: 'A1:C4', enabled: false },
    })
    expect(commands).toEqual([
      {
        id: 'sheet.command.remove-sheet-filter',
        params: { unitId: 'workbook-1', subUnitId: 'sheet-1' },
      },
    ])
  })

  it('replays an already-satisfied XLSX filter state without toggling it again', async () => {
    const commands: string[] = []
    const bounds = { startRow: 0, endRow: 3, startColumn: 0, endColumn: 2 }
    const range = {
      getHeight: () => 4,
      getRange: () => bounds,
    }
    let filter: { getRange: () => typeof range } | null = { getRange: () => range }
    const worksheet = {
      getSheetName: () => 'Budget',
      getSheetId: () => 'sheet-1',
      getFilter: () => filter,
      getRange: () => range,
    }
    const workbook = {
      getId: () => 'workbook-1',
      getSheets: () => [worksheet],
      setActiveSheet: () => undefined,
    }
    const runtime = {
      univerAPI: {
        getActiveWorkbook: () => workbook,
        executeCommand: async (id: string) => {
          commands.push(id)
          return true
        },
      },
    }
    const services = {
      runtime: () => runtime,
      loadStaged: async () => undefined,
      save: async () => ({ ok: true as const, fileName: 'budget.xlsx' }),
    }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.range.set_filter',
          arguments: { sheet: 'Budget', range: 'A1:C4', enabled: true },
        },
        services,
      ),
    ).resolves.toMatchObject({ handled: true, ok: true })

    filter = null
    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.range.set_filter',
          arguments: { sheet: 'Budget', range: 'A1:C4', enabled: false },
        },
        services,
      ),
    ).resolves.toMatchObject({ handled: true, ok: true })

    expect(commands).toEqual([])
  })

  it('clears all criteria from the matching XLSX filter while retaining its range', async () => {
    const commands: Array<{ id: string; params: unknown }> = []
    const bounds = { startRow: 0, endRow: 3, startColumn: 0, endColumn: 2 }
    const range = {
      getHeight: () => 4,
      getRange: () => bounds,
    }
    const worksheet = {
      getSheetName: () => 'Budget',
      getSheetId: () => 'sheet-1',
      getFilter: () => ({ getRange: () => range }),
      getRange: () => range,
    }
    const workbook = {
      getId: () => 'workbook-1',
      getSheets: () => [worksheet],
      setActiveSheet: () => undefined,
    }
    const runtime = {
      univerAPI: {
        getActiveWorkbook: () => workbook,
        executeCommand: async (id: string, params: unknown) => {
          commands.push({ id, params })
          return true
        },
      },
    }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.range.clear_filter_criteria',
          arguments: { sheet: 'Budget', range: 'a1:c4' },
        },
        {
          runtime: () => runtime,
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.range.clear_filter_criteria',
      ok: true,
      output: { sheet: 'Budget', range: 'A1:C4', cleared: true },
    })
    expect(commands).toEqual([
      {
        id: 'sheet.command.clear-filter-criteria',
        params: { unitId: 'workbook-1', subUnitId: 'sheet-1' },
      },
    ])
  })

  it('sets an explicit value-list criterion on a column in the matching XLSX filter', async () => {
    const commands: Array<{ id: string; params: unknown }> = []
    const bounds = { startRow: 0, endRow: 3, startColumn: 0, endColumn: 2 }
    const range = { getRange: () => bounds }
    const worksheet = {
      getSheetName: () => 'Budget',
      getSheetId: () => 'sheet-1',
      getFilter: () => ({ getRange: () => range }),
    }
    const workbook = {
      getId: () => 'workbook-1',
      getSheets: () => [worksheet],
      setActiveSheet: () => undefined,
    }
    const runtime = {
      univerAPI: {
        getActiveWorkbook: () => workbook,
        executeCommand: async (id: string, params: unknown) => {
          commands.push({ id, params })
          return true
        },
      },
    }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.range.set_filter_values',
          arguments: {
            sheet: 'Budget',
            range: 'A1:C4',
            column: 'b',
            values: ['East', 'South'],
            includeBlank: false,
          },
        },
        {
          runtime: () => runtime,
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.range.set_filter_values',
      ok: true,
      output: {
        sheet: 'Budget',
        range: 'A1:C4',
        column: 'B',
        selectedValues: 2,
        includeBlank: false,
      },
    })
    expect(commands).toEqual([
      {
        id: 'sheet.command.set-filter-criteria',
        params: {
          unitId: 'workbook-1',
          subUnitId: 'sheet-1',
          col: 1,
          criteria: { colId: 1, filters: { filters: ['East', 'South'] } },
        },
      },
    ])
  })

  it('sets bounded custom conditions on a column in the matching XLSX filter', async () => {
    const commands: Array<{ id: string; params: unknown }> = []
    const bounds = { startRow: 0, endRow: 20, startColumn: 0, endColumn: 2 }
    const range = { getRange: () => bounds }
    const worksheet = {
      getSheetName: () => 'Budget',
      getSheetId: () => 'sheet-1',
      getFilter: () => ({ getRange: () => range }),
    }
    const workbook = {
      getId: () => 'workbook-1',
      getSheets: () => [worksheet],
      setActiveSheet: () => undefined,
    }
    const runtime = {
      univerAPI: {
        getActiveWorkbook: () => workbook,
        executeCommand: async (id: string, params: unknown) => {
          commands.push({ id, params })
          return true
        },
      },
    }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.range.set_custom_filter',
          arguments: {
            sheet: 'Budget',
            range: 'A1:C21',
            column: 'B',
            conjunction: 'and',
            conditions: [
              { operator: 'greaterThanOrEqual', value: '10' },
              { operator: 'lessThan', value: '20' },
            ],
          },
        },
        {
          runtime: () => runtime,
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.range.set_custom_filter',
      ok: true,
      output: {
        sheet: 'Budget',
        range: 'A1:C21',
        column: 'B',
        conjunction: 'and',
        conditions: 2,
      },
    })
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
  })

  it('rejects filter criteria when the requested range differs from the active filter', async () => {
    let commandExecuted = false
    const worksheet = {
      getSheetName: () => 'Budget',
      getSheetId: () => 'sheet-1',
      getFilter: () => ({
        getRange: () => ({
          getRange: () => ({
            startRow: 0,
            endRow: 3,
            startColumn: 0,
            endColumn: 2,
          }),
        }),
      }),
    }
    const runtime = {
      univerAPI: {
        getActiveWorkbook: () => ({
          getId: () => 'workbook-1',
          getSheets: () => [worksheet],
          setActiveSheet: () => undefined,
        }),
        executeCommand: async () => {
          commandExecuted = true
          return true
        },
      },
    }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.range.set_custom_filter',
          arguments: {
            sheet: 'Budget',
            range: 'A1:C5',
            column: 'B',
            conjunction: 'or',
            conditions: [{ operator: 'equal', value: 'East' }],
          },
        },
        {
          runtime: () => runtime,
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toMatchObject({
      handled: true,
      operationId: 'xlsx.range.set_custom_filter',
      ok: false,
      error: 'execution_failed',
    })
    expect(commandExecuted).toBe(false)
  })

  it('rejects xlsx.history.undo when the mounted workbook has no undo entry', async () => {
    await expect(
      executeXlsxOperation(
        { operation: 'xlsx.history.undo', arguments: {} },
        {
          runtime: () => ({ univerAPI: { undo: async () => false } }),
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.history.undo',
      ok: false,
      error: 'execution_failed',
      message: 'xlsx.history.undo requires an available undo entry.',
    })
  })

  it('rejects xlsx.history.redo when the mounted workbook has no redo entry', async () => {
    await expect(
      executeXlsxOperation(
        { operation: 'xlsx.history.redo', arguments: {} },
        {
          runtime: () => ({ univerAPI: { redo: async () => false } }),
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.history.redo',
      ok: false,
      error: 'execution_failed',
      message: 'xlsx.history.redo requires an available redo entry.',
    })
  })

  it.each([
    ['bold:on', 'xlsx.range.set_text_style'],
    ['align:left', 'xlsx.range.set_alignment'],
    ['valign:middle', 'xlsx.range.set_alignment'],
    ['wrap:on', 'xlsx.range.set_alignment'],
    ['indent:1', 'xlsx.range.set_alignment'],
    ['rotate:45', 'xlsx.range.set_alignment'],
    ['font-family:Aptos', 'xlsx.range.set_font'],
    ['font-size:14', 'xlsx.range.set_font'],
    ['font-color:#44546A', 'xlsx.range.set_font'],
    ['fill:#DDEBF7', 'xlsx.range.set_fill'],
    ['border:all:#4472C4', 'xlsx.range.set_border'],
    ['cell-style:good', 'xlsx.range.apply_cell_style'],
    ['format-painter', 'explicit xlsx.range style operations'],
    ['format:0.00%', 'xlsx.range.set_number_format'],
    ['decimal-inc', 'xlsx.range.set_number_format'],
    ['decimal-dec', 'xlsx.range.set_number_format'],
    ['merge:center', 'xlsx.range.merge'],
    ['clear-contents', 'xlsx.range.clear'],
    ['clear-formats', 'xlsx.range.clear'],
    ['clear-all', 'xlsx.range.clear'],
    ['fill-down', 'xlsx.range.fill'],
    ['fill-right', 'xlsx.range.fill'],
    ['sort:desc', 'xlsx.range.sort'],
    ['sort-custom:1:2a,3d', 'xlsx.range.sort_custom'],
    ['remove-duplicates:1', 'xlsx.range.remove_duplicates'],
    ['autofn:AVERAGE', 'xlsx.formula.insert_aggregate'],
    ['flash-fill', 'xlsx.range.flash_fill'],
    ['text-to-columns:4', 'xlsx.range.text_to_columns'],
    ['row-height:24.75', 'xlsx.row.set_height'],
    ['col-width:12.5', 'xlsx.column.set_width'],
    ['freeze-here', 'xlsx.sheet.set_freeze'],
    ['freeze-top-row', 'xlsx.sheet.set_freeze'],
    ['freeze-first-col', 'xlsx.sheet.set_freeze'],
    ['unfreeze', 'xlsx.sheet.set_freeze'],
    ['toggle-gridlines', 'xlsx.sheet.set_gridlines'],
    ['toggle-show-formulas', 'xlsx.sheet.set_formula_view'],
    ['page-layout:margins:wide', 'xlsx.sheet.set_page_margins'],
    ['page-layout:orientation:landscape', 'xlsx.sheet.set_page_orientation'],
    ['page-layout:paper:9', 'xlsx.sheet.set_paper_size'],
    ['page-layout:fit-width:2', 'xlsx.sheet.set_fit_to_pages'],
    ['page-layout:fit-height:3', 'xlsx.sheet.set_fit_to_pages'],
    ['page-layout:print-gridlines:1', 'xlsx.sheet.set_print_gridlines'],
    ['page-layout:print-headings:1', 'xlsx.sheet.set_print_headings'],
    ['page-layout:print-area:set', 'xlsx.sheet.set_print_area'],
    ['page-layout:print-area:clear', 'xlsx.sheet.set_print_area'],
    ['page-layout:print-titles:first-row', 'xlsx.sheet.set_print_titles'],
    ['page-layout:print-titles:set', 'xlsx.sheet.set_print_titles'],
    ['page-layout:print-titles:clear', 'xlsx.sheet.set_print_titles'],
    ['page-layout:scale:80', 'xlsx.sheet.set_print_scale'],
    ['link-set:https%3A%2F%2Fexample.com', 'xlsx.hyperlink.set'],
    ['link-remove', 'xlsx.hyperlink.remove'],
    ['format-as-table:TableStyleMedium2', 'xlsx.table.add'],
    ['sparkline:column', 'xlsx.sparkline.add'],
    ['outline-group:rows', 'xlsx.outline.set_level'],
    ['outline-ungroup:cols', 'xlsx.outline.set_level'],
    ['outline-hide-detail:rows', 'xlsx.outline.set_detail_visibility'],
    ['outline-show-detail:cols', 'xlsx.outline.set_detail_visibility'],
    ['insert-checkbox', 'xlsx.range.set_checkbox'],
    ['dv-open', 'explicit xlsx.range data-validation operations'],
    ['cf-open', 'explicit xlsx.conditional_format operations'],
    ['insert-symbol', 'xlsx.cell.set_value'],
    ['import-csv', 'xlsx.range.set_values'],
    ['insert-chart:line', 'xlsx.chart.add'],
    ['recommended-charts-open', 'xlsx.chart.add'],
    ['pivot-open', 'xlsx.pivot.add'],
    ['pivot-edit', 'xlsx.pivot.update'],
    ['insert-pivot-chart:column', 'xlsx.pivot.add_chart'],
    ['slicer-open', 'xlsx.pivot.set_member_filter'],
    ['timeline-open', 'xlsx.pivot.set_member_filter'],
    ['insert-picture', 'xlsx.image.add'],
    ['insert-screenshot', 'xlsx.image.add'],
    ['insert-icons', 'xlsx.image.add'],
    ['insert-equation', 'xlsx.image.add'],
    ['insert-function-open', 'xlsx.cell.set_formula'],
    ['replace', 'xlsx.range.replace_text'],
    ['subtotal-open', 'xlsx.range.insert_subtotals'],
    ['consolidate-open', 'xlsx.range.consolidate'],
    ['header-footer-open', 'xlsx.sheet.set_header_footer'],
    ['note-open', 'xlsx.note.set'],
    ['note-delete', 'xlsx.note.remove'],
    ['chart-type-line', 'xlsx.chart.update'],
    ['chart-legend:bottom', 'xlsx.chart.update'],
    ['chart-labels:value', 'xlsx.chart.update'],
    ['chart-grouping:stacked', 'xlsx.chart.update'],
    ['chart-layout:2', 'xlsx.chart.update'],
    ['chart-title:Revenue', 'xlsx.chart.update'],
    ['chart-axis-cat:Quarter', 'xlsx.chart.update'],
    ['chart-axis-val:Amount', 'xlsx.chart.update'],
    ['chart-delete', 'xlsx.chart.remove'],
    ['chart-colors:colorful', 'xlsx.chart.set_colors'],
    ['chart-switch-row-col', 'xlsx.chart.set_series'],
    ['chart-select-data', 'xlsx.chart.set_series'],
    ['chart-format-pane', 'xlsx.chart.update'],
    ['insert-shape:roundRect', 'xlsx.shape.add'],
    ['insert-textbox', 'xlsx.shape.add'],
    ['name-manager-open', 'xlsx.defined_name.set/remove'],
    ['cellprot:locked-on:hidden-off', 'xlsx.range.set_protection'],
    ['sheet-protect', 'xlsx.sheet.set_protection'],
    ['filter-toggle', 'xlsx.range.set_filter'],
    ['filter-clear', 'xlsx.range.clear_filter_criteria'],
    ['filter-advanced', 'xlsx.range.set_custom_filter'],
    ['insert-sheet', 'xlsx.sheet.add'],
    ['insert-row-here', 'xlsx.row.insert'],
    ['delete-row-here', 'xlsx.row.delete'],
    ['insert-col-here', 'xlsx.column.insert'],
    ['delete-col-here', 'xlsx.column.delete'],
    ['undo', 'xlsx.history.undo'],
    ['redo', 'xlsx.history.redo'],
    ['paste-special:value', 'xlsx.range.copy_values'],
    ['paste-special:formula', 'xlsx.range.copy_formulas'],
    ['paste-special:format', 'xlsx.range.copy_formats'],
    ['paste-special:col-width', 'xlsx.column.copy_widths'],
    ['paste-special:besides-border', 'xlsx.range.copy_without_borders'],
    ['copy', 'explicit source/destination XLSX operations'],
    ['cut', 'explicit source/destination XLSX operations'],
    ['paste', 'explicit source/destination XLSX operations'],
  ])('records retired Agent Ribbon command %s with replacement %s', (command, replacement) => {
    expect(command.length).toBeGreaterThan(0)
    expect(replacement).toMatch(/^(xlsx\.|explicit )/)
    if (/^xlsx\.[a-z0-9_]+(?:\.[a-z0-9_]+)+$/.test(replacement)) {
      expect(
        xlsxOperationCatalog.operations.some((operation) => operation.id === replacement),
      ).toBe(true)
    }
  })

  it.each([
    ['drag range', 'xlsx.range.move'],
    ['replace text', 'xlsx.range.replace_text'],
    ['move rows', 'xlsx.row.move'],
    ['hide rows', 'xlsx.row.set_visibility'],
    ['hide columns', 'xlsx.column.set_visibility'],
    ['duplicate worksheet', 'xlsx.sheet.duplicate'],
    ['hide worksheet', 'xlsx.sheet.set_visibility'],
    ['set worksheet tab color', 'xlsx.sheet.set_tab_color'],
    ['insert table rows', 'xlsx.table.insert_rows'],
    ['delete table rows', 'xlsx.table.delete_rows'],
    ['insert table columns', 'xlsx.table.insert_columns'],
    ['delete table columns', 'xlsx.table.delete_columns'],
    ['convert table to range', 'xlsx.table.convert_to_range'],
    ['refresh pivot', 'xlsx.pivot.refresh'],
  ])('maps retained native UI mutation %s to exact operation %s', (_gesture, operationId) => {
    expect(xlsxOperationCatalog.operations.some((operation) => operation.id === operationId)).toBe(
      true,
    )
  })

  it('sets and removes an explicitly addressed note through native Univer history commands', async () => {
    const commands: Array<{ operation: string; arguments: unknown }> = []
    const editJournal = createEditJournal()
    const pendingEdits: number[] = []
    let activeRange: string | null = null
    const worksheet = {
      getSheetName: () => 'Budget',
      getSheetId: () => 'sheet-budget',
      getRange: (address: string) => ({
        activate: () => {
          activeRange = address
        },
      }),
    }
    const workbook = {
      getId: () => 'workbook-1',
      getSheets: () => [worksheet],
      setActiveSheet: () => undefined,
    }
    const runtime = {
      univerAPI: {
        getActiveWorkbook: () => workbook,
        executeCommand: async (operation: string, arguments_: unknown) => {
          commands.push({ operation, arguments: arguments_ })
          return true
        },
      },
    }
    const services = {
      runtime: () => runtime,
      state: () => ({
        editJournal,
        hyperlinkTargets: new Map(),
        sheetProtections: new Map(),
      }),
      setPendingEdits: (count: number) => pendingEdits.push(count),
      loadStaged: async () => undefined,
      save: async () => ({ ok: true as const, fileName: 'budget.xlsx' }),
    }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.note.set',
          arguments: { sheet: 'Budget', address: 'c4', text: 'Check this total.' },
        },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.note.set',
      ok: true,
      output: { changed: 1, sheet: 'Budget', address: 'C4' },
    })
    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.note.remove',
          arguments: { sheet: 'Budget', address: 'c4' },
        },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.note.remove',
      ok: true,
      output: { changed: 1, sheet: 'Budget', address: 'C4' },
    })

    expect(activeRange).toBe('C4')
    expect([...editJournal.noteDirty]).toEqual(['sheet-budget'])
    expect(pendingEdits).toEqual([1, 1])
    expect(commands).toEqual([
      {
        operation: 'sheet.command.update-note',
        arguments: {
          unitId: 'workbook-1',
          sheetId: 'sheet-budget',
          row: 3,
          col: 2,
          note: { row: 3, col: 2, width: 220, height: 90, note: 'Check this total.' },
        },
      },
      {
        operation: 'sheet.command.delete-note',
        arguments: {
          unitId: 'workbook-1',
          sheetId: 'sheet-budget',
          row: 3,
          col: 2,
        },
      },
    ])
  })

  it('rejects an oversized note before invoking Univer', async () => {
    let mutated = false
    const worksheet = {
      getSheetName: () => 'Budget',
      getSheetId: () => 'sheet-budget',
      getRange: () => ({ activate: () => undefined }),
    }
    const workbook = {
      getId: () => 'workbook-1',
      getSheets: () => [worksheet],
      setActiveSheet: () => undefined,
    }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.note.set',
          arguments: { sheet: 'Budget', address: 'A1', text: 'x'.repeat(32_768) },
        },
        {
          runtime: () => ({
            univerAPI: {
              getActiveWorkbook: () => workbook,
              executeCommand: async () => {
                mutated = true
                return true
              },
            },
          }),
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toMatchObject({
      handled: true,
      operationId: 'xlsx.note.set',
      ok: false,
      error: 'invalid_arguments',
    })
    expect(mutated).toBe(false)
  })

  it('updates and removes an explicitly identified chart through the shared visual history seam', async () => {
    const editJournal = createEditJournal()
    editJournal.visualAdds.push({
      id: 'added-chart-1',
      sheetId: 'sheet-budget',
      kind: 'chart',
      anchor: {
        fromRow: 0,
        fromColumn: 3,
        fromRowOffset: 0,
        fromColumnOffset: 0,
        toRow: 15,
        toColumn: 10,
        toRowOffset: 0,
        toColumnOffset: 0,
      },
      chart: { chartTypes: ['lineChart'], title: 'Old', series: [] },
    })
    const edits: unknown[] = []
    const removals: string[] = []
    const services = {
      runtime: () => null,
      state: () => ({ editJournal, hyperlinkTargets: new Map(), sheetProtections: new Map() }),
      editChart: (editKey: string, edit: unknown) => edits.push({ editKey, edit }),
      removeVisual: (visualId: string) => removals.push(visualId),
      loadStaged: async () => undefined,
      save: async () => ({ ok: true as const, fileName: 'budget.xlsx' }),
    }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.chart.update',
          arguments: {
            chartId: 'added-chart-1',
            title: 'Revenue',
            type: 'bar',
            legend: 'bottom',
            dataLabels: 'value',
            grouping: 'stacked',
            categoryAxisTitle: 'Quarter',
            valueAxisTitle: null,
            gridlines: true,
            valueAxisMin: 0,
            valueAxisMax: 100,
            gapWidthPct: 80,
            holeSizePct: 40,
            explosionPct: 15,
            dataLabelPosition: 'outside-end',
            dataLabelFormat: '0.0%',
          },
        },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.chart.update',
      ok: true,
      output: { changed: 1, chartId: 'added-chart-1' },
    })
    await expect(
      executeXlsxOperation(
        { operation: 'xlsx.chart.remove', arguments: { chartId: 'added-chart-1' } },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.chart.remove',
      ok: true,
      output: { changed: 1, chartId: 'added-chart-1' },
    })

    expect(edits).toEqual([
      {
        editKey: 'added-chart-1',
        edit: {
          title: 'Revenue',
          chartType: 'bar',
          legend: 'bottom',
          dataLabels: 'value',
          grouping: 'stacked',
          axisTitles: { category: 'Quarter', value: null },
          gridlines: true,
          valueAxis: { min: 0, max: 100 },
          gapWidthPct: 80,
          holeSizePct: 40,
          explosionPct: 15,
          dataLabelPosition: 'outside-end',
          dataLabelFormat: '0.0%',
        },
      },
    ])
    expect(removals).toEqual(['added-chart-1'])
  })

  it('sets bounded explicit chart colors and complete series through the shared chart edit seam', async () => {
    const editJournal = createEditJournal()
    editJournal.visualAdds.push({
      id: 'added-chart-1',
      sheetId: 'sheet-budget',
      kind: 'chart',
      anchor: {
        fromRow: 0,
        fromColumn: 3,
        fromRowOffset: 0,
        fromColumnOffset: 0,
        toRow: 15,
        toColumn: 10,
        toRowOffset: 0,
        toColumnOffset: 0,
      },
      chart: { chartTypes: ['pieChart'], title: 'Revenue', series: [] },
    })
    const edits: unknown[] = []
    const services = {
      runtime: () => null,
      state: () => ({ editJournal, hyperlinkTargets: new Map(), sheetProtections: new Map() }),
      editChart: (editKey: string, edit: unknown) => edits.push({ editKey, edit }),
      loadStaged: async () => undefined,
      save: async () => ({ ok: true as const, fileName: 'budget.xlsx' }),
    }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.chart.set_colors',
          arguments: {
            chartId: 'added-chart-1',
            seriesColors: ['#4472C4', '#ED7D31'],
            pointColors: ['#70AD47', '#FFC000'],
          },
        },
        services,
      ),
    ).resolves.toMatchObject({
      handled: true,
      operationId: 'xlsx.chart.set_colors',
      ok: true,
      output: { changed: 1, chartId: 'added-chart-1' },
    })
    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.chart.set_series',
          arguments: {
            chartId: 'added-chart-1',
            series: [
              {
                name: 'Revenue',
                values: [10, 20],
                categories: ['Q1', 'Q2'],
                valuesRef: "'Budget'!$B$2:$B$3",
                categoriesRef: "'Budget'!$A$2:$A$3",
              },
            ],
          },
        },
        services,
      ),
    ).resolves.toMatchObject({
      handled: true,
      operationId: 'xlsx.chart.set_series',
      ok: true,
      output: { changed: 1, chartId: 'added-chart-1', seriesCount: 1 },
    })

    expect(edits).toEqual([
      {
        editKey: 'added-chart-1',
        edit: {
          seriesColors: { 0: '#4472C4', 1: '#ED7D31' },
          pointColors: { 0: { 0: '#70AD47', 1: '#FFC000' } },
        },
      },
      {
        editKey: 'added-chart-1',
        edit: {
          seriesSet: [
            {
              name: 'Revenue',
              values: [10, 20],
              categories: ['Q1', 'Q2'],
              valuesRef: "'Budget'!$B$2:$B$3",
              categoriesRef: "'Budget'!$A$2:$A$3",
            },
          ],
        },
      },
    ])
  })

  it('adds one bounded chart through the shared file-backed visual action', async () => {
    const additions: unknown[] = []
    const worksheet = {
      getSheetName: () => 'Budget',
      getSheetId: () => 'sheet-budget',
    }
    const workbook = {
      getSheets: () => [worksheet],
      setActiveSheet: () => undefined,
    }
    const runtime = { univerAPI: { getActiveWorkbook: () => workbook } }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.chart.add',
          arguments: {
            sheet: 'Budget',
            dataRange: 'a1:b4',
            type: 'line',
            anchorCell: 'd2',
          },
        },
        {
          runtime: () => runtime,
          addChart: async (input) => {
            additions.push(input)
            return 'added-chart-1'
          },
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.chart.add',
      ok: true,
      output: {
        sheet: 'Budget',
        dataRange: 'A1:B4',
        type: 'line',
        anchorCell: 'D2',
        visualId: 'added-chart-1',
      },
    })
    expect(additions).toEqual([
      {
        op: 'add_chart',
        sheetId: 'sheet-budget',
        dataRange: 'A1:B4',
        chartType: 'line',
        anchorCell: 'D2',
      },
    ])
  })

  it('rejects xlsx.chart.add ranges over 2,000 cells before visual mutation', async () => {
    let mutated = false
    const worksheet = {
      getSheetName: () => 'Budget',
      getSheetId: () => 'sheet-budget',
    }
    const workbook = {
      getSheets: () => [worksheet],
      setActiveSheet: () => undefined,
    }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.chart.add',
          arguments: { sheet: 'Budget', dataRange: 'A1:Z100', type: 'column' },
        },
        {
          runtime: () => ({ univerAPI: { getActiveWorkbook: () => workbook } }),
          addChart: async () => {
            mutated = true
            return 'unexpected'
          },
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.chart.add',
      ok: false,
      error: 'invalid_arguments',
      message: 'xlsx.chart.add supports at most 2,000 source cells.',
    })
    expect(mutated).toBe(false)
  })

  it('adds one bounded pivot through the shared workbook pivot action', async () => {
    const additions: unknown[] = []
    const sourceWorksheet = {
      getSheetName: () => 'Source',
      getSheetId: () => 'sheet-source',
    }
    const targetWorksheet = {
      getSheetName: () => 'Report',
      getSheetId: () => 'sheet-report',
    }
    const workbook = {
      getSheets: () => [sourceWorksheet, targetWorksheet],
      setActiveSheet: () => undefined,
    }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.pivot.add',
          arguments: {
            sourceSheet: 'Source',
            sourceRange: 'a1:d200',
            targetSheet: 'Report',
            targetCell: 'b3',
            name: 'RevenuePivot',
            rowFields: ['Region', 'Team'],
            columnFields: [],
            pageFields: ['Year'],
            values: [
              {
                field: 'Revenue',
                aggregation: 'sum',
                numberFormat: '#,##0.00',
                showAs: null,
              },
            ],
          },
        },
        {
          runtime: () => ({ univerAPI: { getActiveWorkbook: () => workbook } }),
          isSheetDataComplete: () => true,
          addPivot: async (operation) => {
            additions.push(operation)
            return 'RevenuePivot'
          },
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'pivot.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.pivot.add',
      ok: true,
      output: {
        sourceSheet: 'Source',
        sourceRange: 'A1:D200',
        targetSheet: 'Report',
        targetCell: 'B3',
        name: 'RevenuePivot',
      },
    })
    expect(additions).toEqual([
      {
        op: 'add_pivot',
        sheetId: 'sheet-source',
        sourceRange: 'A1:D200',
        targetSheetId: 'sheet-report',
        targetCell: 'B3',
        name: 'RevenuePivot',
        rowFields: ['Region', 'Team'],
        pageFields: ['Year'],
        values: [
          {
            field: 'Revenue',
            agg: 'sum',
            numFmt: '#,##0.00',
          },
        ],
      },
    ])
  })

  it('rejects incoherent pivot fields before invoking the workbook pivot action', async () => {
    let mutated = false
    const worksheets = ['Source', 'Report'].map((name) => ({
      getSheetName: () => name,
      getSheetId: () => `sheet-${name.toLowerCase()}`,
    }))

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.pivot.add',
          arguments: {
            sourceSheet: 'Source',
            sourceRange: 'A1:D20',
            targetSheet: 'Report',
            targetCell: 'A1',
            name: null,
            rowFields: ['Region'],
            columnFields: ['Year'],
            pageFields: [],
            values: [
              { field: 'Revenue', aggregation: 'sum', numberFormat: null, showAs: null },
              { field: 'Cost', aggregation: 'sum', numberFormat: null, showAs: null },
            ],
          },
        },
        {
          runtime: () => ({
            univerAPI: {
              getActiveWorkbook: () => ({
                getSheets: () => worksheets,
                setActiveSheet: () => undefined,
              }),
            },
          }),
          isSheetDataComplete: () => true,
          addPivot: async () => {
            mutated = true
            return 'unexpected'
          },
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'pivot.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.pivot.add',
      ok: false,
      error: 'invalid_arguments',
      message: 'xlsx.pivot.add requires exactly one values entry when columnFields are present.',
    })
    expect(mutated).toBe(false)
  })

  it('rejects pivot creation while either source or target sheet data is incomplete', async () => {
    let mutated = false
    const worksheets = ['Source', 'Report'].map((name) => ({
      getSheetName: () => name,
      getSheetId: () => `sheet-${name.toLowerCase()}`,
    }))

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.pivot.add',
          arguments: {
            sourceSheet: 'Source',
            sourceRange: 'A1:B20',
            targetSheet: 'Report',
            targetCell: 'A1',
            name: null,
            rowFields: ['Region'],
            columnFields: [],
            pageFields: [],
            values: [{ field: 'Revenue', aggregation: 'sum', numberFormat: null, showAs: null }],
          },
        },
        {
          runtime: () => ({
            univerAPI: {
              getActiveWorkbook: () => ({
                getSheets: () => worksheets,
                setActiveSheet: () => undefined,
              }),
            },
          }),
          isSheetDataComplete: (sheetId) => sheetId === 'sheet-source',
          addPivot: async () => {
            mutated = true
            return 'unexpected'
          },
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'pivot.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.pivot.add',
      ok: false,
      error: 'execution_failed',
      message: 'xlsx.pivot.add requires fully loaded source and target worksheets.',
    })
    expect(mutated).toBe(false)
  })

  it('refreshes every native pivot on one explicitly named worksheet', async () => {
    const refreshed: string[] = []
    const worksheet = {
      getSheetName: () => 'Report',
      getSheetId: () => 'sheet-report',
    }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.pivot.refresh',
          arguments: { sheet: 'Report' },
        },
        {
          runtime: () => ({
            univerAPI: {
              getActiveWorkbook: () => ({
                getSheets: () => [worksheet],
                setActiveSheet: () => undefined,
              }),
            },
          }),
          refreshPivots: (sheetId) => {
            refreshed.push(sheetId)
            return 2
          },
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'pivot.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.pivot.refresh',
      ok: true,
      output: { sheet: 'Report', refreshed: 2 },
    })
    expect(refreshed).toEqual(['sheet-report'])
  })

  it('replaces one explicitly identified pivot layout without using current selection', async () => {
    const updates: unknown[] = []

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.pivot.update',
          arguments: {
            pivotId: 'xl/pivotTables/pivotTable1.xml',
            targetCell: 'c4',
            rowFields: ['Region'],
            columnFields: [],
            pageFields: ['Year'],
            values: [
              {
                field: 'Revenue',
                aggregation: 'average',
                numberFormat: '0.00',
                showAs: 'percentOfTotal',
              },
            ],
          },
        },
        {
          runtime: () => null,
          updatePivot: (input) => {
            updates.push(input)
          },
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'pivot.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.pivot.update',
      ok: true,
      output: {
        pivotId: 'xl/pivotTables/pivotTable1.xml',
        targetCell: 'C4',
      },
    })
    expect(updates).toEqual([
      {
        pivotId: 'xl/pivotTables/pivotTable1.xml',
        targetCell: 'C4',
        rowFields: ['Region'],
        pageFields: ['Year'],
        values: [
          {
            field: 'Revenue',
            agg: 'average',
            numFmt: '0.00',
            showDataAs: 'percentOfTotal',
          },
        ],
      },
    ])
  })

  it('sets one explicit final member filter on an identified pivot field', async () => {
    const filters: unknown[] = []

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.pivot.set_member_filter',
          arguments: {
            pivotId: 'xl/pivotTables/pivotTable1.xml',
            field: 'Region',
            selectedValues: ['East'],
          },
        },
        {
          runtime: () => null,
          setPivotMemberFilter: (input) => {
            filters.push(input)
            return 1
          },
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'pivot.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.pivot.set_member_filter',
      ok: true,
      output: {
        pivotId: 'xl/pivotTables/pivotTable1.xml',
        field: 'Region',
        selectedCount: 1,
      },
    })
    expect(filters).toEqual([
      {
        pivotId: 'xl/pivotTables/pivotTable1.xml',
        field: 'Region',
        selectedValues: ['East'],
      },
    ])
  })

  it('adds an explicitly anchored bounded shape through the shared visual history seam', async () => {
    const additions: unknown[] = []
    const worksheet = { getSheetName: () => 'Budget', getSheetId: () => 'sheet-budget' }
    const services = {
      runtime: () => ({
        univerAPI: {
          getActiveWorkbook: () => ({
            getSheets: () => [worksheet],
            setActiveSheet: () => undefined,
          }),
        },
      }),
      addShape: async (operation: unknown) => {
        additions.push(operation)
        return 'added-shape-1'
      },
      loadStaged: async () => undefined,
      save: async () => ({ ok: true as const, fileName: 'budget.xlsx' }),
    }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.shape.add',
          arguments: {
            sheet: 'Budget',
            type: 'roundRect',
            anchorCell: 'D2',
            fillColor: '#DDEBF7',
            text: 'Forecast',
          },
        },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.shape.add',
      ok: true,
      output: {
        sheet: 'Budget',
        type: 'roundRect',
        anchorCell: 'D2',
        visualId: 'added-shape-1',
      },
    })
    expect(additions).toEqual([
      {
        op: 'add_shape',
        sheetId: 'sheet-budget',
        shapeType: 'roundRect',
        anchorCell: 'D2',
        fillColor: '#DDEBF7',
        text: 'Forecast',
      },
    ])
  })

  it('updates and removes an explicitly identified shape through the shared visual history seam', async () => {
    const editJournal = createEditJournal()
    const fileShape = {
      id: 'file-shape-1',
      sheetId: 'sheet-budget',
      kind: 'shape' as const,
      anchor: {
        fromRow: 1,
        fromColumn: 3,
        fromRowOffset: 0,
        fromColumnOffset: 0,
        toRow: 7,
        toColumn: 6,
        toRowOffset: 0,
        toColumnOffset: 0,
      },
      shapeType: 'roundRect',
      fillColor: '#DDEBF7',
      text: 'Forecast',
      drawingPath: 'xl/drawings/drawing1.xml',
      drawingIndex: 0,
    }
    const edits: unknown[] = []
    const removals: string[] = []
    const services = {
      runtime: () => null,
      state: () => ({
        editJournal,
        hyperlinkTargets: new Map(),
        sheetProtections: new Map(),
        file: { visuals: [fileShape] },
      }),
      updateShape: (shapeId: string, changes: unknown) => edits.push({ shapeId, changes }),
      removeVisual: (visualId: string) => removals.push(visualId),
      loadStaged: async () => undefined,
      save: async () => ({ ok: true as const, fileName: 'budget.xlsx' }),
    }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.shape.update',
          arguments: {
            shapeId: 'file-shape-1',
            anchorCell: 'F4',
            fillColor: '#4472C4',
            text: 'Updated forecast',
          },
        },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.shape.update',
      ok: true,
      output: { changed: 1, shapeId: 'file-shape-1' },
    })
    await expect(
      executeXlsxOperation(
        { operation: 'xlsx.shape.remove', arguments: { shapeId: 'file-shape-1' } },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.shape.remove',
      ok: true,
      output: { changed: 1, shapeId: 'file-shape-1' },
    })
    expect(edits).toEqual([
      {
        shapeId: 'file-shape-1',
        changes: {
          anchor: {
            fromRow: 3,
            fromColumn: 5,
            fromRowOffset: 0,
            fromColumnOffset: 0,
            toRow: 9,
            toColumn: 8,
            toRowOffset: 0,
            toColumnOffset: 0,
          },
          fillColor: '#4472C4',
          text: 'Updated forecast',
        },
      },
    ])
    expect(removals).toEqual(['file-shape-1'])
  })

  it('moves and removes an explicitly identified image through the shared visual history seam', async () => {
    const editJournal = createEditJournal()
    const fileImage = {
      id: 'file-image-1',
      sheetId: 'sheet-budget',
      kind: 'image' as const,
      anchor: {
        fromRow: 1,
        fromColumn: 3,
        fromRowOffset: 0,
        fromColumnOffset: 0,
        toRow: 7,
        toColumn: 6,
        toRowOffset: 0,
        toColumnOffset: 0,
      },
      mediaPath: 'xl/media/image1.png',
      mediaType: 'image/png',
      drawingPath: 'xl/drawings/drawing1.xml',
      drawingIndex: 0,
    }
    const moves: unknown[] = []
    const removals: string[] = []
    const services = {
      runtime: () => null,
      state: () => ({
        editJournal,
        hyperlinkTargets: new Map(),
        sheetProtections: new Map(),
        file: { visuals: [fileImage] },
      }),
      moveVisual: (visualId: string, anchor: unknown) => moves.push({ visualId, anchor }),
      removeVisual: (visualId: string) => removals.push(visualId),
      loadStaged: async () => undefined,
      save: async () => ({ ok: true as const, fileName: 'budget.xlsx' }),
    }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.image.move',
          arguments: { imageId: 'file-image-1', anchorCell: 'F4' },
        },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.image.move',
      ok: true,
      output: { changed: 1, imageId: 'file-image-1' },
    })
    await expect(
      executeXlsxOperation(
        { operation: 'xlsx.image.remove', arguments: { imageId: 'file-image-1' } },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.image.remove',
      ok: true,
      output: { changed: 1, imageId: 'file-image-1' },
    })
    expect(moves).toEqual([
      {
        visualId: 'file-image-1',
        anchor: {
          fromRow: 3,
          fromColumn: 5,
          fromRowOffset: 0,
          fromColumnOffset: 0,
          toRow: 9,
          toColumn: 8,
          toRowOffset: 0,
          toColumnOffset: 0,
        },
      },
    ])
    expect(removals).toEqual(['file-image-1'])
  })

  it('sets, renames, and removes explicitly scoped defined names through the native workbook model', async () => {
    const names = new Map<string, { name: string; ref: string; scope: string }>()
    const key = (name: string, scope: string) => `${scope}!${name}`
    const defined = (entry: { name: string; ref: string; scope: string }) => ({
      getName: () => entry.name,
      getFormulaOrRefString: () => entry.ref,
      getLocalSheetId: () => entry.scope,
      setName: (name: string) => {
        names.delete(key(entry.name, entry.scope))
        entry.name = name
        names.set(key(entry.name, entry.scope), entry)
      },
      setRef: (ref: string) => {
        entry.ref = ref
      },
      delete: () => names.delete(key(entry.name, entry.scope)),
    })
    const workbook = {
      getSheets: () => [
        { getSheetName: () => 'Budget', getSheetId: () => 'sheet-budget' },
        { getSheetName: () => 'Forecast', getSheetId: () => 'sheet-forecast' },
      ],
      setActiveSheet: () => undefined,
      getDefinedNames: () => [...names.values()].map(defined),
      newDefinedNameBuilder: () => ({
        load: (entry: { name: string; formulaOrRefString: string; localSheetId: string }) => ({
          build: () => entry,
        }),
      }),
      insertDefinedNameBuilder: (entry: {
        name: string
        formulaOrRefString: string
        localSheetId: string
      }) => {
        const stored = {
          name: entry.name,
          ref: entry.formulaOrRefString,
          scope: entry.localSheetId,
        }
        names.set(key(stored.name, stored.scope), stored)
      },
    }
    const services = {
      runtime: () => ({ univerAPI: { getActiveWorkbook: () => workbook } }),
      loadStaged: async () => undefined,
      save: async () => ({ ok: true as const, fileName: 'budget.xlsx' }),
    }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.defined_name.set',
          arguments: { name: 'Revenue', formula: '=Budget!$B$2:$B$9' },
        },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.defined_name.set',
      ok: true,
      output: {
        changed: 1,
        name: 'Revenue',
        formula: 'Budget!$B$2:$B$9',
        scope: 'workbook',
      },
    })
    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.defined_name.set',
          arguments: {
            previousName: 'Revenue',
            name: 'ForecastRevenue',
            formula: '=Forecast!$C$2:$C$9',
          },
        },
        services,
      ),
    ).resolves.toMatchObject({ ok: true, operationId: 'xlsx.defined_name.set' })
    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.defined_name.set',
          arguments: {
            name: 'LocalRate',
            formula: '=Forecast!$D$2',
            scopeSheet: 'Forecast',
          },
        },
        services,
      ),
    ).resolves.toMatchObject({ ok: true, operationId: 'xlsx.defined_name.set' })
    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.defined_name.remove',
          arguments: { name: 'LocalRate', scopeSheet: 'Forecast' },
        },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.defined_name.remove',
      ok: true,
      output: { changed: 1, name: 'LocalRate', scope: 'Forecast' },
    })
    expect([...names.values()]).toEqual([
      {
        name: 'ForecastRevenue',
        ref: 'Forecast!$C$2:$C$9',
        scope: 'AllDefaultWorkbook',
      },
    ])
  })

  it('fails closed for invalid, missing, or unknown-scope defined names', async () => {
    let inserts = 0
    const workbook = {
      getSheets: () => [{ getSheetName: () => 'Budget', getSheetId: () => 'sheet-budget' }],
      setActiveSheet: () => undefined,
      getDefinedNames: () => [],
      newDefinedNameBuilder: () => ({ load: () => ({ build: () => ({}) }) }),
      insertDefinedNameBuilder: () => {
        inserts += 1
      },
    }
    const services = {
      runtime: () => ({ univerAPI: { getActiveWorkbook: () => workbook } }),
      loadStaged: async () => undefined,
      save: async () => ({ ok: true as const, fileName: 'budget.xlsx' }),
    }
    for (const command of [
      {
        operation: 'xlsx.defined_name.set',
        arguments: { name: 'A1', formula: '=Budget!$A$1' },
      },
      {
        operation: 'xlsx.defined_name.set',
        arguments: { name: 'ValidName', formula: '=Budget!$A$1', scopeSheet: 'Missing' },
      },
      {
        operation: 'xlsx.defined_name.set',
        arguments: { name: 'Next', previousName: 'Gone', formula: '=Budget!$A$1' },
      },
      {
        operation: 'xlsx.defined_name.remove',
        arguments: { name: 'Gone' },
      },
    ]) {
      await expect(executeXlsxOperation(command, services)).resolves.toMatchObject({
        handled: true,
        ok: false,
      })
    }
    expect(inserts).toBe(0)
  })

  it('hydrates and adds one bounded staged image through the shared visual action', async () => {
    const data = new TextEncoder().encode('GIF87a').buffer
    const additions: unknown[] = []
    const worksheet = { getSheetName: () => 'Budget', getSheetId: () => 'sheet-budget' }
    const workbook = { getSheets: () => [worksheet], setActiveSheet: () => undefined }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.image.add_staged',
          arguments: {
            blobId: 'image-blob',
            name: 'logo.gif',
            size: data.byteLength,
            data,
            sheet: 'Budget',
            anchorCell: 'c3',
          },
        },
        {
          runtime: () => ({ univerAPI: { getActiveWorkbook: () => workbook } }),
          addImage: async (input) => {
            additions.push(input)
            return 'added-image-1'
          },
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.image.add_staged',
      ok: true,
      output: {
        sheet: 'Budget',
        anchorCell: 'C3',
        visualId: 'added-image-1',
      },
    })
    expect(additions).toEqual([
      {
        sheetId: 'sheet-budget',
        anchorCell: 'C3',
        name: 'logo.gif',
        mediaType: 'image/gif',
        data,
      },
    ])
  })

  it('sets an explicit XLSX row span height in points', async () => {
    const heights: Array<{ startRow: number; count: number; heightPixels: number }> = []
    const worksheet = {
      getSheetName: () => 'Budget',
      getMaxRows: () => 1_048_576,
      setRowHeightsForced: (startRow: number, count: number, heightPixels: number) => {
        heights.push({ startRow, count, heightPixels })
      },
    }
    const workbook = {
      getSheets: () => [worksheet],
      setActiveSheet: () => undefined,
    }
    const runtime = { univerAPI: { getActiveWorkbook: () => workbook } }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.row.set_height',
          arguments: { sheet: 'Budget', row: 3, count: 2, heightPoints: 24.75 },
        },
        {
          runtime: () => runtime,
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.row.set_height',
      ok: true,
      output: { sheet: 'Budget', row: 3, count: 2, heightPoints: 24.75 },
    })
    expect(heights).toEqual([{ startRow: 2, count: 2, heightPixels: 33 }])
  })

  it('rejects a row-height span outside the worksheet before mutation', async () => {
    let mutated = false
    const worksheet = {
      getSheetName: () => 'Budget',
      getMaxRows: () => 4,
      setRowHeightsForced: () => {
        mutated = true
      },
    }
    const workbook = {
      getSheets: () => [worksheet],
      setActiveSheet: () => undefined,
    }
    const runtime = { univerAPI: { getActiveWorkbook: () => workbook } }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.row.set_height',
          arguments: { sheet: 'Budget', row: 4, count: 2, heightPoints: 20 },
        },
        {
          runtime: () => runtime,
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.row.set_height',
      ok: false,
      error: 'invalid_arguments',
      message: 'xlsx.row.set_height must stay inside the worksheet row bounds.',
    })
    expect(mutated).toBe(false)
  })

  it('sets an explicit XLSX column span width in characters', async () => {
    const widths: Array<{ startColumn: number; count: number; widthPixels: number }> = []
    const worksheet = {
      getSheetName: () => 'Budget',
      getMaxColumns: () => 16_384,
      setColumnWidths: (startColumn: number, count: number, widthPixels: number) => {
        widths.push({ startColumn, count, widthPixels })
      },
    }
    const workbook = {
      getSheets: () => [worksheet],
      setActiveSheet: () => undefined,
    }
    const runtime = { univerAPI: { getActiveWorkbook: () => workbook } }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.column.set_width',
          arguments: { sheet: 'Budget', column: 'C', count: 2, widthCharacters: 12.5 },
        },
        {
          runtime: () => runtime,
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.column.set_width',
      ok: true,
      output: { sheet: 'Budget', column: 'C', count: 2, widthCharacters: 12.4296875 },
    })
    expect(widths).toEqual([{ startColumn: 2, count: 2, widthPixels: 92 }])
  })

  it('rejects a column-width span outside the worksheet before mutation', async () => {
    let mutated = false
    const worksheet = {
      getSheetName: () => 'Budget',
      getMaxColumns: () => 4,
      setColumnWidths: () => {
        mutated = true
      },
    }
    const workbook = {
      getSheets: () => [worksheet],
      setActiveSheet: () => undefined,
    }
    const runtime = { univerAPI: { getActiveWorkbook: () => workbook } }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.column.set_width',
          arguments: { sheet: 'Budget', column: 'D', count: 2, widthCharacters: 12.5 },
        },
        {
          runtime: () => runtime,
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.column.set_width',
      ok: false,
      error: 'invalid_arguments',
      message: 'xlsx.column.set_width must stay inside the worksheet column bounds.',
    })
    expect(mutated).toBe(false)
  })

  it('rejects a copy-width source span outside its worksheet before reading widths', async () => {
    let sourceRead = false
    const sourceWorksheet = {
      getSheetName: () => 'Source',
      getMaxColumns: () => 4,
      getColumnWidth: () => {
        sourceRead = true
        return 80
      },
    }
    const destinationWorksheet = {
      getSheetName: () => 'Destination',
      getMaxColumns: () => 16,
    }
    const workbook = {
      getSheets: () => [sourceWorksheet, destinationWorksheet],
      setActiveSheet: () => undefined,
    }
    const runtime = { univerAPI: { getActiveWorkbook: () => workbook } }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.column.copy_widths',
          arguments: {
            sourceSheet: 'Source',
            sourceColumn: 'D',
            destinationSheet: 'Destination',
            destinationColumn: 'A',
            count: 2,
          },
        },
        {
          runtime: () => runtime,
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.column.copy_widths',
      ok: false,
      error: 'invalid_arguments',
      message: 'xlsx.column.copy_widths must stay inside both worksheet column bounds.',
    })
    expect(sourceRead).toBe(false)
  })

  it('rejects a copy-width destination span outside its worksheet before reading widths', async () => {
    let sourceRead = false
    const sourceWorksheet = {
      getSheetName: () => 'Source',
      getMaxColumns: () => 16,
      getColumnWidth: () => {
        sourceRead = true
        return 80
      },
    }
    const destinationWorksheet = {
      getSheetName: () => 'Destination',
      getMaxColumns: () => 4,
    }
    const workbook = {
      getSheets: () => [sourceWorksheet, destinationWorksheet],
      setActiveSheet: () => undefined,
    }
    const runtime = { univerAPI: { getActiveWorkbook: () => workbook } }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.column.copy_widths',
          arguments: {
            sourceSheet: 'Source',
            sourceColumn: 'A',
            destinationSheet: 'Destination',
            destinationColumn: 'D',
            count: 2,
          },
        },
        {
          runtime: () => runtime,
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.column.copy_widths',
      ok: false,
      error: 'invalid_arguments',
      message: 'xlsx.column.copy_widths must stay inside both worksheet column bounds.',
    })
    expect(sourceRead).toBe(false)
  })

  it('rejects copy-width targets that are not A1 column labels', async () => {
    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.column.copy_widths',
          arguments: {
            sourceSheet: 'Source',
            sourceColumn: 'A1',
            destinationSheet: 'Destination',
            destinationColumn: 'D',
            count: 2,
          },
        },
        {
          runtime: () => null,
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.column.copy_widths',
      ok: false,
      error: 'invalid_arguments',
      message: 'xlsx.column.copy_widths requires A1 source and destination column labels.',
    })
  })

  it('sets an explicit XLSX worksheet freeze split', async () => {
    const freezes: Array<{
      startRow: number
      startColumn: number
      xSplit: number
      ySplit: number
    }> = []
    const worksheet = {
      getSheetName: () => 'Budget',
      getMaxRows: () => 1_048_576,
      getMaxColumns: () => 16_384,
      setFreeze: (freeze: {
        startRow: number
        startColumn: number
        xSplit: number
        ySplit: number
      }) => freezes.push(freeze),
    }
    const workbook = {
      getSheets: () => [worksheet],
      setActiveSheet: () => undefined,
    }
    const runtime = { univerAPI: { getActiveWorkbook: () => workbook } }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.sheet.set_freeze',
          arguments: { sheet: 'Budget', frozenRows: 2, frozenColumns: 1 },
        },
        {
          runtime: () => runtime,
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.sheet.set_freeze',
      ok: true,
      output: { sheet: 'Budget', frozenRows: 2, frozenColumns: 1 },
    })
    expect(freezes).toEqual([{ startRow: 2, startColumn: 1, xSplit: 1, ySplit: 2 }])
  })

  it('removes the XLSX worksheet freeze when both split counts are zero', async () => {
    let canceled = 0
    const worksheet = {
      getSheetName: () => 'Budget',
      getMaxRows: () => 1_048_576,
      getMaxColumns: () => 16_384,
      setFreeze: () => undefined,
      cancelFreeze: () => {
        canceled += 1
      },
    }
    const workbook = {
      getSheets: () => [worksheet],
      setActiveSheet: () => undefined,
    }
    const runtime = { univerAPI: { getActiveWorkbook: () => workbook } }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.sheet.set_freeze',
          arguments: { sheet: 'Budget', frozenRows: 0, frozenColumns: 0 },
        },
        {
          runtime: () => runtime,
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toMatchObject({ handled: true, ok: true })
    expect(canceled).toBe(1)
  })

  it('rejects a freeze split that leaves no scrollable worksheet pane', async () => {
    let mutated = false
    const worksheet = {
      getSheetName: () => 'Budget',
      getMaxRows: () => 4,
      getMaxColumns: () => 3,
      setFreeze: () => {
        mutated = true
      },
      cancelFreeze: () => {
        mutated = true
      },
    }
    const workbook = {
      getSheets: () => [worksheet],
      setActiveSheet: () => undefined,
    }
    const runtime = { univerAPI: { getActiveWorkbook: () => workbook } }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.sheet.set_freeze',
          arguments: { sheet: 'Budget', frozenRows: 4, frozenColumns: 1 },
        },
        {
          runtime: () => runtime,
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.sheet.set_freeze',
      ok: false,
      error: 'invalid_arguments',
      message: 'xlsx.sheet.set_freeze must leave at least one unfrozen row and column.',
    })
    expect(mutated).toBe(false)
  })

  it('sets the explicit XLSX worksheet gridline visibility', async () => {
    const hiddenStates: boolean[] = []
    const worksheet = {
      getSheetName: () => 'Budget',
      setHiddenGridlines: (hidden: boolean) => hiddenStates.push(hidden),
    }
    const workbook = {
      getSheets: () => [worksheet],
      setActiveSheet: () => undefined,
    }
    const runtime = { univerAPI: { getActiveWorkbook: () => workbook } }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.sheet.set_gridlines',
          arguments: { sheet: 'Budget', visible: false },
        },
        {
          runtime: () => runtime,
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.sheet.set_gridlines',
      ok: true,
      output: { sheet: 'Budget', visible: false },
    })
    expect(hiddenStates).toEqual([true])
  })

  it('sets the explicit XLSX worksheet formula view state', async () => {
    const editJournal = createEditJournal()
    const contextValues = new Map<unknown, unknown>()
    const worksheet = {
      getSheetName: () => 'Budget',
      getSheetId: () => 'sheet-1',
    }
    const workbook = {
      getId: () => undefined,
      getSheets: () => [worksheet],
      setActiveSheet: () => undefined,
    }
    const runtime = {
      univerAPI: { getActiveWorkbook: () => workbook },
      univer: {
        __getInjector: () => ({
          get: () => ({
            getContextValue: (key: unknown) => contextValues.get(key),
            setContextValue: (key: unknown, value: unknown) => contextValues.set(key, value),
          }),
        }),
      },
    }
    const state = {
      editJournal,
      hyperlinkTargets: new Map(),
      sheetProtections: new Map(),
      showFormulaSheets: new Set<string>(),
    }
    let pendingEdits = 0

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.sheet.set_formula_view',
          arguments: { sheet: 'Budget', enabled: true },
        },
        {
          runtime: () => runtime,
          state: () => state,
          setPendingEdits: (count) => {
            pendingEdits = count
          },
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.sheet.set_formula_view',
      ok: true,
      output: { sheet: 'Budget', enabled: true },
    })
    expect(state.showFormulaSheets).toEqual(new Set(['sheet-1']))
    expect(editJournal.pageSetup.get('sheet-1')).toEqual({ showFormulas: true })
    expect(pendingEdits).toBe(1)
  })

  it('sets the explicit XLSX worksheet page orientation', async () => {
    const editJournal = createEditJournal()
    const worksheet = {
      getSheetName: () => 'Budget',
      getSheetId: () => 'sheet-1',
    }
    const workbook = {
      getId: () => undefined,
      getSheets: () => [worksheet],
      setActiveSheet: () => undefined,
    }
    const runtime = { univerAPI: { getActiveWorkbook: () => workbook } }
    const state = {
      editJournal,
      hyperlinkTargets: new Map(),
      sheetProtections: new Map(),
      file: {
        sheets: [{ id: 'sheet-1', name: 'Budget', pageSetup: { orientation: 'portrait' } }],
      } as unknown as LazyWorkbookState['file'],
    }
    let pendingEdits = 0

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.sheet.set_page_orientation',
          arguments: { sheet: 'Budget', orientation: 'landscape' },
        },
        {
          runtime: () => runtime,
          state: () => state,
          setPendingEdits: (count) => {
            pendingEdits = count
          },
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.sheet.set_page_orientation',
      ok: true,
      output: { sheet: 'Budget', orientation: 'landscape' },
    })
    expect(editJournal.pageSetup.get('sheet-1')).toEqual({ orientation: 'landscape' })
    expect(pendingEdits).toBe(1)
  })

  it('sets explicit XLSX header and footer sections through page-setup history', async () => {
    const editJournal = createEditJournal()
    const worksheet = {
      getSheetName: () => 'Budget',
      getSheetId: () => 'sheet-1',
    }
    const workbook = {
      getId: () => undefined,
      getSheets: () => [worksheet],
      setActiveSheet: () => undefined,
    }
    const state = {
      editJournal,
      hyperlinkTargets: new Map(),
      sheetProtections: new Map(),
      file: {
        sheets: [{ id: 'sheet-1', name: 'Budget', pageSetup: {} }],
      } as unknown as LazyWorkbookState['file'],
    }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.sheet.set_header_footer',
          arguments: {
            sheet: 'Budget',
            header: { left: 'Quarterly', center: '&F' },
            footer: null,
          },
        },
        {
          runtime: () => ({ univerAPI: { getActiveWorkbook: () => workbook } }),
          state: () => state,
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.sheet.set_header_footer',
      ok: true,
      output: {
        sheet: 'Budget',
        header: { left: 'Quarterly', center: '&F' },
        footer: null,
      },
    })
    expect(editJournal.pageSetup.get('sheet-1')).toEqual({
      header: { left: 'Quarterly', center: '&F' },
      footer: null,
    })
  })

  it('routes explicit XLSX subtotals through the retained subtotal action', async () => {
    const worksheet = {
      getSheetName: () => 'Budget',
      getSheetId: () => 'sheet-1',
      getRange: () => ({
        getRow: () => 1,
        getColumn: () => 0,
        getHeight: () => 5,
        getWidth: () => 3,
      }),
    }
    const calls: unknown[] = []

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.range.insert_subtotals',
          arguments: {
            sheet: 'Budget',
            range: 'A2:C6',
            groupColumn: 'A',
            valueColumn: 'C',
            aggregation: 'sum',
          },
        },
        {
          runtime: () => ({
            univerAPI: {
              getActiveWorkbook: () => ({
                getSheets: () => [worksheet],
                setActiveSheet: () => undefined,
              }),
            },
          }),
          createSubtotals: (input) => {
            calls.push(input)
            return 2
          },
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.range.insert_subtotals',
      ok: true,
      output: {
        sheet: 'Budget',
        range: 'A2:C6',
        groupColumn: 'A',
        valueColumn: 'C',
        aggregation: 'sum',
        groups: 2,
        insertedRows: 3,
      },
    })
    expect(calls).toEqual([
      {
        sheetId: 'sheet-1',
        range: 'A2:C6',
        groupColumn: 0,
        valueColumn: 2,
        aggregation: 'sum',
      },
    ])
  })

  it('routes explicit XLSX consolidation through the retained consolidation action', async () => {
    const sheet = (name: string, id: string) => ({
      getSheetName: () => name,
      getSheetId: () => id,
      getRange: () => ({
        getRow: () => 1,
        getColumn: () => 0,
        getHeight: () => 7,
        getWidth: () => 3,
      }),
    })
    const budget = sheet('Budget', 'sheet-budget')
    const actual = sheet('Actual', 'sheet-actual')
    const forecast = sheet('Forecast', 'sheet-forecast')
    const calls: unknown[] = []

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.range.consolidate',
          arguments: {
            targetSheet: 'Budget',
            targetCell: 'E2',
            sources: [
              { sheet: 'Actual', range: 'A2:C8' },
              { sheet: 'Forecast', range: 'A2:C8' },
            ],
            aggregation: 'sum',
            leftLabels: true,
          },
        },
        {
          runtime: () => ({
            univerAPI: {
              getActiveWorkbook: () => ({
                getSheets: () => [budget, actual, forecast],
                setActiveSheet: () => undefined,
              }),
            },
          }),
          consolidate: (input) => {
            calls.push(input)
            return { rows: 4, columns: 3 }
          },
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.range.consolidate',
      ok: true,
      output: {
        targetSheet: 'Budget',
        targetCell: 'E2',
        sources: 2,
        aggregation: 'sum',
        leftLabels: true,
        rows: 4,
        columns: 3,
      },
    })
    expect(calls).toEqual([
      {
        targetSheetId: 'sheet-budget',
        targetCell: 'E2',
        sources: [
          { sheetId: 'sheet-actual', range: 'A2:C8' },
          { sheetId: 'sheet-forecast', range: 'A2:C8' },
        ],
        aggregation: 'sum',
        leftLabels: true,
      },
    ])
  })

  it('routes explicit XLSX PivotChart insertion through the retained visual action', async () => {
    const calls: unknown[] = []
    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.pivot.add_chart',
          arguments: { pivotId: 'xl/pivotTables/pivotTable1.xml', type: 'column' },
        },
        {
          runtime: () => null,
          addPivotChart: async (input) => {
            calls.push(input)
            return { chartId: 'added-pivotchart-1', truncated: false }
          },
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.pivot.add_chart',
      ok: true,
      output: {
        pivotId: 'xl/pivotTables/pivotTable1.xml',
        type: 'column',
        chartId: 'added-pivotchart-1',
        truncated: false,
      },
    })
    expect(calls).toEqual([{ pivotId: 'xl/pivotTables/pivotTable1.xml', type: 'column' }])
  })

  it('moves an explicit XLSX range through Univer native history', async () => {
    const commands: unknown[] = []
    const source = { getSheetName: () => 'Input', getSheetId: () => 'sheet-input' }
    const destination = { getSheetName: () => 'Output', getSheetId: () => 'sheet-output' }
    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.range.move',
          arguments: {
            sourceSheet: 'Input',
            sourceRange: 'A2:B4',
            destinationSheet: 'Output',
            destinationRange: 'D5:E7',
          },
        },
        {
          runtime: () => ({
            univerAPI: {
              getActiveWorkbook: () => ({
                getId: () => 'workbook-1',
                getSheets: () => [source, destination],
                setActiveSheet: () => undefined,
              }),
              executeCommand: async (id: string, params: unknown) => {
                commands.push({ id, params })
                return true
              },
            },
          }),
          isSheetDataComplete: () => true,
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.range.move',
      ok: true,
      output: {
        sourceSheet: 'Input',
        sourceRange: 'A2:B4',
        destinationSheet: 'Output',
        destinationRange: 'D5:E7',
        moved: 6,
      },
    })
    expect(commands).toEqual([
      {
        id: 'sheet.command.move-range',
        params: {
          fromUnitId: 'workbook-1',
          fromSubUnitId: 'sheet-input',
          fromRange: { startRow: 1, endRow: 3, startColumn: 0, endColumn: 1 },
          toUnitId: 'workbook-1',
          toSubUnitId: 'sheet-output',
          toRange: { startRow: 4, endRow: 6, startColumn: 3, endColumn: 4 },
        },
      },
    ])
  })

  it('moves an explicit XLSX row span through Univer native structural history', async () => {
    const commands: unknown[] = []
    const worksheet = {
      getSheetName: () => 'Budget',
      getSheetId: () => 'sheet-budget',
      getMaxRows: () => 100,
      getMaxColumns: () => 12,
    }
    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.row.move',
          arguments: { sheet: 'Budget', row: 2, count: 2, beforeRow: 7 },
        },
        {
          runtime: () => ({
            univerAPI: {
              getActiveWorkbook: () => ({
                getId: () => 'workbook-1',
                getSheets: () => [worksheet],
                setActiveSheet: () => undefined,
              }),
              executeCommand: async (id: string, params: unknown) => {
                commands.push({ id, params })
                return true
              },
            },
          }),
          isSheetDataComplete: () => true,
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.row.move',
      ok: true,
      output: { sheet: 'Budget', row: 2, count: 2, beforeRow: 7 },
    })
    expect(commands).toEqual([
      {
        id: 'sheet.command.move-rows',
        params: {
          unitId: 'workbook-1',
          subUnitId: 'sheet-budget',
          fromRange: { startRow: 1, endRow: 2, startColumn: 0, endColumn: 11 },
          toRange: { startRow: 6, endRow: 6, startColumn: 0, endColumn: 11 },
        },
      },
    ])
  })

  it('sets an explicit XLSX worksheet tab color through Univer native history', async () => {
    const commands: unknown[] = []
    const worksheet = { getSheetName: () => 'Budget', getSheetId: () => 'sheet-budget' }
    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.sheet.set_tab_color',
          arguments: { sheet: 'Budget', color: '#112233' },
        },
        {
          runtime: () => ({
            univerAPI: {
              getActiveWorkbook: () => ({
                getId: () => 'workbook-1',
                getSheets: () => [worksheet],
                setActiveSheet: () => undefined,
              }),
              executeCommand: async (id: string, params: unknown) => {
                commands.push({ id, params })
                return true
              },
            },
          }),
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.sheet.set_tab_color',
      ok: true,
      output: { sheet: 'Budget', color: '#112233' },
    })
    expect(commands).toEqual([
      {
        id: 'sheet.command.set-tab-color',
        params: { unitId: 'workbook-1', subUnitId: 'sheet-budget', value: '#112233' },
      },
    ])
  })

  it('sets the explicit XLSX worksheet page margins', async () => {
    const editJournal = createEditJournal()
    const worksheet = {
      getSheetName: () => 'Budget',
      getSheetId: () => 'sheet-1',
    }
    const workbook = {
      getId: () => undefined,
      getSheets: () => [worksheet],
      setActiveSheet: () => undefined,
    }
    const runtime = { univerAPI: { getActiveWorkbook: () => workbook } }
    const state = {
      editJournal,
      hyperlinkTargets: new Map(),
      sheetProtections: new Map(),
      file: {
        sheets: [{ id: 'sheet-1', name: 'Budget', pageSetup: { margins: 'normal' } }],
      } as unknown as LazyWorkbookState['file'],
    }
    let pendingEdits = 0

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.sheet.set_page_margins',
          arguments: { sheet: 'Budget', margins: 'wide' },
        },
        {
          runtime: () => runtime,
          state: () => state,
          setPendingEdits: (count) => {
            pendingEdits = count
          },
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.sheet.set_page_margins',
      ok: true,
      output: { sheet: 'Budget', margins: 'wide' },
    })
    expect(editJournal.pageSetup.get('sheet-1')).toEqual({ margins: 'wide' })
    expect(pendingEdits).toBe(1)
  })

  it('sets the explicit XLSX worksheet paper size', async () => {
    const editJournal = createEditJournal()
    const worksheet = {
      getSheetName: () => 'Budget',
      getSheetId: () => 'sheet-1',
    }
    const workbook = {
      getId: () => undefined,
      getSheets: () => [worksheet],
      setActiveSheet: () => undefined,
    }
    const runtime = { univerAPI: { getActiveWorkbook: () => workbook } }
    const state = {
      editJournal,
      hyperlinkTargets: new Map(),
      sheetProtections: new Map(),
      file: {
        sheets: [{ id: 'sheet-1', name: 'Budget', pageSetup: { paperSize: 1 } }],
      } as unknown as LazyWorkbookState['file'],
    }
    let pendingEdits = 0

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.sheet.set_paper_size',
          arguments: { sheet: 'Budget', paperSize: 9 },
        },
        {
          runtime: () => runtime,
          state: () => state,
          setPendingEdits: (count) => {
            pendingEdits = count
          },
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.sheet.set_paper_size',
      ok: true,
      output: { sheet: 'Budget', paperSize: 9 },
    })
    expect(editJournal.pageSetup.get('sheet-1')).toEqual({ paperSize: 9 })
    expect(pendingEdits).toBe(1)
  })

  it('sets both XLSX fit-to-page axes as one explicit final state', async () => {
    const editJournal = createEditJournal()
    const worksheet = {
      getSheetName: () => 'Budget',
      getSheetId: () => 'sheet-1',
    }
    const workbook = {
      getId: () => undefined,
      getSheets: () => [worksheet],
      setActiveSheet: () => undefined,
    }
    const runtime = { univerAPI: { getActiveWorkbook: () => workbook } }
    const state = {
      editJournal,
      hyperlinkTargets: new Map(),
      sheetProtections: new Map(),
      file: {
        sheets: [
          {
            id: 'sheet-1',
            name: 'Budget',
            pageSetup: { fitToWidth: 1, fitToHeight: 0, fitToPage: true },
          },
        ],
      } as unknown as LazyWorkbookState['file'],
    }
    let pendingEdits = 0

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.sheet.set_fit_to_pages',
          arguments: { sheet: 'Budget', widthPages: 2, heightPages: 3 },
        },
        {
          runtime: () => runtime,
          state: () => state,
          setPendingEdits: (count) => {
            pendingEdits = count
          },
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.sheet.set_fit_to_pages',
      ok: true,
      output: { sheet: 'Budget', widthPages: 2, heightPages: 3 },
    })
    expect(editJournal.pageSetup.get('sheet-1')).toEqual({
      fitToWidth: 2,
      fitToHeight: 3,
      fitToPage: true,
    })
    expect(pendingEdits).toBe(1)
  })

  it('disables XLSX fit-to-page when both axes are automatic', async () => {
    const editJournal = createEditJournal()
    const worksheet = {
      getSheetName: () => 'Budget',
      getSheetId: () => 'sheet-1',
    }
    const workbook = {
      getId: () => undefined,
      getSheets: () => [worksheet],
      setActiveSheet: () => undefined,
    }
    const runtime = { univerAPI: { getActiveWorkbook: () => workbook } }
    const state = {
      editJournal,
      hyperlinkTargets: new Map(),
      sheetProtections: new Map(),
      file: {
        sheets: [
          {
            id: 'sheet-1',
            name: 'Budget',
            pageSetup: { fitToWidth: 2, fitToHeight: 3, fitToPage: true },
          },
        ],
      } as unknown as LazyWorkbookState['file'],
    }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.sheet.set_fit_to_pages',
          arguments: { sheet: 'Budget', widthPages: 0, heightPages: 0 },
        },
        {
          runtime: () => runtime,
          state: () => state,
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toMatchObject({
      handled: true,
      ok: true,
      output: { sheet: 'Budget', widthPages: 0, heightPages: 0 },
    })
    expect(editJournal.pageSetup.get('sheet-1')).toEqual({
      fitToWidth: 0,
      fitToHeight: 0,
      fitToPage: false,
    })
  })

  it('sets XLSX print scale and disables fit-to-page as one final state', async () => {
    const editJournal = createEditJournal()
    const worksheet = {
      getSheetName: () => 'Budget',
      getSheetId: () => 'sheet-1',
    }
    const workbook = {
      getId: () => undefined,
      getSheets: () => [worksheet],
      setActiveSheet: () => undefined,
    }
    const runtime = { univerAPI: { getActiveWorkbook: () => workbook } }
    const state = {
      editJournal,
      hyperlinkTargets: new Map(),
      sheetProtections: new Map(),
      file: {
        sheets: [
          {
            id: 'sheet-1',
            name: 'Budget',
            pageSetup: { scale: 100, fitToPage: true },
          },
        ],
      } as unknown as LazyWorkbookState['file'],
    }
    let pendingEdits = 0

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.sheet.set_print_scale',
          arguments: { sheet: 'Budget', scalePercent: 80 },
        },
        {
          runtime: () => runtime,
          state: () => state,
          setPendingEdits: (count) => {
            pendingEdits = count
          },
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.sheet.set_print_scale',
      ok: true,
      output: { sheet: 'Budget', scalePercent: 80 },
    })
    expect(editJournal.pageSetup.get('sheet-1')).toEqual({ scale: 80, fitToPage: false })
    expect(pendingEdits).toBe(1)
  })

  it('sets the explicit XLSX worksheet print-gridline state', async () => {
    const editJournal = createEditJournal()
    const worksheet = {
      getSheetName: () => 'Budget',
      getSheetId: () => 'sheet-1',
    }
    const workbook = {
      getId: () => undefined,
      getSheets: () => [worksheet],
      setActiveSheet: () => undefined,
    }
    const runtime = { univerAPI: { getActiveWorkbook: () => workbook } }
    const state = {
      editJournal,
      hyperlinkTargets: new Map(),
      sheetProtections: new Map(),
      file: {
        sheets: [{ id: 'sheet-1', name: 'Budget', pageSetup: { printGridlines: false } }],
      } as unknown as LazyWorkbookState['file'],
    }
    let pendingEdits = 0

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.sheet.set_print_gridlines',
          arguments: { sheet: 'Budget', enabled: true },
        },
        {
          runtime: () => runtime,
          state: () => state,
          setPendingEdits: (count) => {
            pendingEdits = count
          },
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.sheet.set_print_gridlines',
      ok: true,
      output: { sheet: 'Budget', enabled: true },
    })
    expect(editJournal.pageSetup.get('sheet-1')).toEqual({ printGridlines: true })
    expect(pendingEdits).toBe(1)
  })

  it('disables XLSX worksheet print gridlines explicitly', async () => {
    const editJournal = createEditJournal()
    const worksheet = {
      getSheetName: () => 'Budget',
      getSheetId: () => 'sheet-1',
    }
    const workbook = {
      getId: () => undefined,
      getSheets: () => [worksheet],
      setActiveSheet: () => undefined,
    }
    const runtime = { univerAPI: { getActiveWorkbook: () => workbook } }
    const state = {
      editJournal,
      hyperlinkTargets: new Map(),
      sheetProtections: new Map(),
      file: {
        sheets: [{ id: 'sheet-1', name: 'Budget', pageSetup: { printGridlines: true } }],
      } as unknown as LazyWorkbookState['file'],
    }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.sheet.set_print_gridlines',
          arguments: { sheet: 'Budget', enabled: false },
        },
        {
          runtime: () => runtime,
          state: () => state,
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toMatchObject({
      handled: true,
      ok: true,
      output: { sheet: 'Budget', enabled: false },
    })
    expect(editJournal.pageSetup.get('sheet-1')).toEqual({ printGridlines: false })
  })

  it('sets the explicit XLSX worksheet print-heading state', async () => {
    const editJournal = createEditJournal()
    const worksheet = {
      getSheetName: () => 'Budget',
      getSheetId: () => 'sheet-1',
    }
    const workbook = {
      getId: () => undefined,
      getSheets: () => [worksheet],
      setActiveSheet: () => undefined,
    }
    const runtime = { univerAPI: { getActiveWorkbook: () => workbook } }
    const state = {
      editJournal,
      hyperlinkTargets: new Map(),
      sheetProtections: new Map(),
      file: {
        sheets: [{ id: 'sheet-1', name: 'Budget', pageSetup: { printHeadings: false } }],
      } as unknown as LazyWorkbookState['file'],
    }
    let pendingEdits = 0

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.sheet.set_print_headings',
          arguments: { sheet: 'Budget', enabled: true },
        },
        {
          runtime: () => runtime,
          state: () => state,
          setPendingEdits: (count) => {
            pendingEdits = count
          },
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.sheet.set_print_headings',
      ok: true,
      output: { sheet: 'Budget', enabled: true },
    })
    expect(editJournal.pageSetup.get('sheet-1')).toEqual({ printHeadings: true })
    expect(pendingEdits).toBe(1)
  })

  it('disables XLSX worksheet print headings explicitly', async () => {
    const editJournal = createEditJournal()
    const worksheet = {
      getSheetName: () => 'Budget',
      getSheetId: () => 'sheet-1',
    }
    const workbook = {
      getId: () => undefined,
      getSheets: () => [worksheet],
      setActiveSheet: () => undefined,
    }
    const runtime = { univerAPI: { getActiveWorkbook: () => workbook } }
    const state = {
      editJournal,
      hyperlinkTargets: new Map(),
      sheetProtections: new Map(),
      file: {
        sheets: [{ id: 'sheet-1', name: 'Budget', pageSetup: { printHeadings: true } }],
      } as unknown as LazyWorkbookState['file'],
    }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.sheet.set_print_headings',
          arguments: { sheet: 'Budget', enabled: false },
        },
        {
          runtime: () => runtime,
          state: () => state,
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toMatchObject({
      handled: true,
      ok: true,
      output: { sheet: 'Budget', enabled: false },
    })
    expect(editJournal.pageSetup.get('sheet-1')).toEqual({ printHeadings: false })
  })

  it('sets an explicit XLSX worksheet print area', async () => {
    const editJournal = createEditJournal()
    const worksheet = {
      getSheetName: () => 'Budget',
      getSheetId: () => 'sheet-1',
    }
    const workbook = {
      getId: () => undefined,
      getSheets: () => [worksheet],
      setActiveSheet: () => undefined,
    }
    const runtime = { univerAPI: { getActiveWorkbook: () => workbook } }
    const state = {
      editJournal,
      hyperlinkTargets: new Map(),
      sheetProtections: new Map(),
      file: {
        sheets: [{ id: 'sheet-1', name: 'Budget', pageSetup: {} }],
      } as unknown as LazyWorkbookState['file'],
    }
    let pendingEdits = 0

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.sheet.set_print_area',
          arguments: { sheet: 'Budget', range: 'b2:d8' },
        },
        {
          runtime: () => runtime,
          state: () => state,
          setPendingEdits: (count) => {
            pendingEdits = count
          },
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.sheet.set_print_area',
      ok: true,
      output: { sheet: 'Budget', range: 'B2:D8' },
    })
    expect(editJournal.pageSetup.get('sheet-1')).toEqual({ printArea: 'B2:D8' })
    expect(pendingEdits).toBe(1)
  })

  it('clears the XLSX worksheet print area explicitly', async () => {
    const editJournal = createEditJournal()
    const worksheet = {
      getSheetName: () => 'Budget',
      getSheetId: () => 'sheet-1',
    }
    const workbook = {
      getId: () => undefined,
      getSheets: () => [worksheet],
      setActiveSheet: () => undefined,
    }
    const runtime = { univerAPI: { getActiveWorkbook: () => workbook } }
    const state = {
      editJournal,
      hyperlinkTargets: new Map(),
      sheetProtections: new Map(),
      file: {
        sheets: [{ id: 'sheet-1', name: 'Budget', pageSetup: { printArea: 'A1:C10' } }],
      } as unknown as LazyWorkbookState['file'],
    }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.sheet.set_print_area',
          arguments: { sheet: 'Budget', range: null },
        },
        {
          runtime: () => runtime,
          state: () => state,
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toMatchObject({
      handled: true,
      ok: true,
      output: { sheet: 'Budget', range: null },
    })
    expect(editJournal.pageSetup.get('sheet-1')).toEqual({ printArea: null })
  })

  it.each(['Budget!A1:C10', 'XFE1:XFE2'])(
    'rejects invalid explicit XLSX print area %s',
    async (range) => {
      const editJournal = createEditJournal()
      const worksheet = {
        getSheetName: () => 'Budget',
        getSheetId: () => 'sheet-1',
      }
      const workbook = {
        getId: () => undefined,
        getSheets: () => [worksheet],
        setActiveSheet: () => undefined,
      }

      await expect(
        executeXlsxOperation(
          {
            operation: 'xlsx.sheet.set_print_area',
            arguments: { sheet: 'Budget', range },
          },
          {
            runtime: () => ({ univerAPI: { getActiveWorkbook: () => workbook } }),
            state: () => ({
              editJournal,
              hyperlinkTargets: new Map(),
              sheetProtections: new Map(),
            }),
            loadStaged: async () => undefined,
            save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
          },
        ),
      ).resolves.toEqual({
        handled: true,
        operationId: 'xlsx.sheet.set_print_area',
        ok: false,
        error: 'invalid_arguments',
        message: 'xlsx.sheet.set_print_area requires an explicit A1 cell range or null.',
      })
      expect(editJournal.pageSetup.size).toBe(0)
    },
  )

  it('sets explicit XLSX worksheet print-title rows', async () => {
    const editJournal = createEditJournal()
    const worksheet = {
      getSheetName: () => 'Budget',
      getSheetId: () => 'sheet-1',
    }
    const workbook = {
      getId: () => undefined,
      getSheets: () => [worksheet],
      setActiveSheet: () => undefined,
    }
    const runtime = { univerAPI: { getActiveWorkbook: () => workbook } }
    const state = {
      editJournal,
      hyperlinkTargets: new Map(),
      sheetProtections: new Map(),
      file: {
        sheets: [{ id: 'sheet-1', name: 'Budget', pageSetup: {} }],
      } as unknown as LazyWorkbookState['file'],
    }
    let pendingEdits = 0

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.sheet.set_print_titles',
          arguments: { sheet: 'Budget', rows: '2:8' },
        },
        {
          runtime: () => runtime,
          state: () => state,
          setPendingEdits: (count) => {
            pendingEdits = count
          },
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.sheet.set_print_titles',
      ok: true,
      output: { sheet: 'Budget', rows: '2:8' },
    })
    expect(editJournal.pageSetup.get('sheet-1')).toEqual({ printTitles: '2:8' })
    expect(pendingEdits).toBe(1)
  })

  it('clears XLSX worksheet print-title rows explicitly', async () => {
    const editJournal = createEditJournal()
    const worksheet = {
      getSheetName: () => 'Budget',
      getSheetId: () => 'sheet-1',
    }
    const workbook = {
      getId: () => undefined,
      getSheets: () => [worksheet],
      setActiveSheet: () => undefined,
    }
    const runtime = { univerAPI: { getActiveWorkbook: () => workbook } }
    const state = {
      editJournal,
      hyperlinkTargets: new Map(),
      sheetProtections: new Map(),
      file: {
        sheets: [{ id: 'sheet-1', name: 'Budget', pageSetup: { printTitles: '1:2' } }],
      } as unknown as LazyWorkbookState['file'],
    }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.sheet.set_print_titles',
          arguments: { sheet: 'Budget', rows: null },
        },
        {
          runtime: () => runtime,
          state: () => state,
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toMatchObject({
      handled: true,
      ok: true,
      output: { sheet: 'Budget', rows: null },
    })
    expect(editJournal.pageSetup.get('sheet-1')).toEqual({ printTitles: null })
  })

  it.each(['0:1', '3:1', '1:22', '1048576:1048577'])(
    'rejects invalid explicit XLSX print-title rows %s',
    async (rows) => {
      const editJournal = createEditJournal()
      const worksheet = {
        getSheetName: () => 'Budget',
        getSheetId: () => 'sheet-1',
      }
      const workbook = {
        getId: () => undefined,
        getSheets: () => [worksheet],
        setActiveSheet: () => undefined,
      }

      await expect(
        executeXlsxOperation(
          {
            operation: 'xlsx.sheet.set_print_titles',
            arguments: { sheet: 'Budget', rows },
          },
          {
            runtime: () => ({ univerAPI: { getActiveWorkbook: () => workbook } }),
            state: () => ({
              editJournal,
              hyperlinkTargets: new Map(),
              sheetProtections: new Map(),
            }),
            loadStaged: async () => undefined,
            save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
          },
        ),
      ).resolves.toEqual({
        handled: true,
        operationId: 'xlsx.sheet.set_print_titles',
        ok: false,
        error: 'invalid_arguments',
        message:
          'xlsx.sheet.set_print_titles requires an ascending explicit row span of at most 21 rows or null.',
      })
      expect(editJournal.pageSetup.size).toBe(0)
    },
  )

  it('inserts rows through xlsx.row.insert using a 1-based row target', async () => {
    const inserted: Array<{ rowIndex: number; count: number }> = []
    const worksheet = {
      getSheetName: () => 'Budget',
      insertRowsBefore: (rowIndex: number, count: number) => {
        inserted.push({ rowIndex, count })
      },
    }
    const workbook = {
      getSheets: () => [worksheet],
      setActiveSheet: () => undefined,
    }
    const runtime = { univerAPI: { getActiveWorkbook: () => workbook } }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.row.insert',
          arguments: { sheet: 'Budget', row: 3, count: 2 },
        },
        {
          runtime: () => runtime,
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.row.insert',
      ok: true,
      output: { sheet: 'Budget', row: 3, count: 2 },
    })
    expect(inserted).toEqual([{ rowIndex: 2, count: 2 }])
  })

  it('deletes rows through xlsx.row.delete using a 1-based row target', async () => {
    const deleted: Array<{ rowIndex: number; count: number }> = []
    const worksheet = {
      getSheetName: () => 'Budget',
      deleteRows: (rowIndex: number, count: number) => {
        deleted.push({ rowIndex, count })
      },
    }
    const workbook = {
      getSheets: () => [worksheet],
      setActiveSheet: () => undefined,
    }
    const runtime = { univerAPI: { getActiveWorkbook: () => workbook } }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.row.delete',
          arguments: { sheet: 'Budget', row: 4, count: 3 },
        },
        {
          runtime: () => runtime,
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.row.delete',
      ok: true,
      output: { sheet: 'Budget', row: 4, count: 3 },
    })
    expect(deleted).toEqual([{ rowIndex: 3, count: 3 }])
  })

  it('inserts columns through xlsx.column.insert using an A1 column label', async () => {
    const inserted: Array<{ columnIndex: number; count: number }> = []
    const worksheet = {
      getSheetName: () => 'Budget',
      insertColumnsBefore: (columnIndex: number, count: number) => {
        inserted.push({ columnIndex, count })
      },
    }
    const workbook = {
      getSheets: () => [worksheet],
      setActiveSheet: () => undefined,
    }
    const runtime = { univerAPI: { getActiveWorkbook: () => workbook } }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.column.insert',
          arguments: { sheet: 'Budget', column: 'C', count: 2 },
        },
        {
          runtime: () => runtime,
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.column.insert',
      ok: true,
      output: { sheet: 'Budget', column: 'C', count: 2 },
    })
    expect(inserted).toEqual([{ columnIndex: 2, count: 2 }])
  })

  it('deletes columns through xlsx.column.delete using an A1 column label', async () => {
    const deleted: Array<{ columnIndex: number; count: number }> = []
    const worksheet = {
      getSheetName: () => 'Budget',
      deleteColumns: (columnIndex: number, count: number) => {
        deleted.push({ columnIndex, count })
      },
    }
    const workbook = {
      getSheets: () => [worksheet],
      setActiveSheet: () => undefined,
    }
    const runtime = { univerAPI: { getActiveWorkbook: () => workbook } }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.column.delete',
          arguments: { sheet: 'Budget', column: 'D', count: 2 },
        },
        {
          runtime: () => runtime,
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.column.delete',
      ok: true,
      output: { sheet: 'Budget', column: 'D', count: 2 },
    })
    expect(deleted).toEqual([{ columnIndex: 3, count: 2 }])
  })

  it.each(['xlsx.column.delete', 'xlsx.column.insert'] as const)(
    'rejects an invalid %s label before mutating the worksheet',
    async (operation) => {
      let mutated = false
      const worksheet = {
        getSheetName: () => 'Budget',
        deleteColumns: () => {
          mutated = true
        },
        insertColumnsBefore: () => {
          mutated = true
        },
      }
      const workbook = {
        getSheets: () => [worksheet],
        setActiveSheet: () => undefined,
      }
      const runtime = { univerAPI: { getActiveWorkbook: () => workbook } }

      await expect(
        executeXlsxOperation(
          {
            operation,
            arguments: { sheet: 'Budget', column: 'A1', count: 1 },
          },
          {
            runtime: () => runtime,
            loadStaged: async () => undefined,
            save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
          },
        ),
      ).resolves.toEqual({
        handled: true,
        operationId: operation,
        ok: false,
        error: 'invalid_arguments',
        message: `${operation} requires an A1 column label.`,
      })
      expect(mutated).toBe(false)
    },
  )

  it('rejects an invalid xlsx.sheet.add name before mutating the workbook', async () => {
    let inserted = false
    const workbook = {
      insertSheet: () => {
        inserted = true
      },
    }
    const runtime = { univerAPI: { getActiveWorkbook: () => workbook } }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.sheet.add',
          arguments: { name: ' ' },
        },
        {
          runtime: () => runtime,
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.sheet.add',
      ok: false,
      error: 'invalid_arguments',
      message: 'xlsx.sheet.add requires a non-empty sheet name of at most 31 characters.',
    })
    expect(inserted).toBe(false)
  })

  it('adds a worksheet through xlsx.sheet.add', async () => {
    const inserted: string[] = []
    const workbook = {
      insertSheet: (name: string) => {
        inserted.push(name)
      },
    }
    const runtime = { univerAPI: { getActiveWorkbook: () => workbook } }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.sheet.add',
          arguments: { name: 'Forecast' },
        },
        {
          runtime: () => runtime,
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.sheet.add',
      ok: true,
      output: { name: 'Forecast' },
    })
    expect(inserted).toEqual(['Forecast'])
  })

  it('renames the specified worksheet through xlsx.sheet.rename', async () => {
    let renamedTo: string | null = null
    const worksheet = {
      getSheetName: () => 'Budget',
      setName: (name: string) => {
        renamedTo = name
      },
    }
    const workbook = {
      getSheets: () => [worksheet],
      setActiveSheet: () => undefined,
    }
    const runtime = { univerAPI: { getActiveWorkbook: () => workbook } }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.sheet.rename',
          arguments: { sheet: 'Budget', name: 'Plan' },
        },
        {
          runtime: () => runtime,
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.sheet.rename',
      ok: true,
      output: { sheet: 'Budget', name: 'Plan' },
    })
    expect(renamedTo).toBe('Plan')
  })

  it('refuses to delete the only worksheet through xlsx.sheet.delete', async () => {
    let deleted = false
    const worksheet = { getSheetName: () => 'Budget' }
    const workbook = {
      getSheets: () => [worksheet],
      setActiveSheet: () => undefined,
      deleteSheet: () => {
        deleted = true
        return true
      },
    }
    const runtime = { univerAPI: { getActiveWorkbook: () => workbook } }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.sheet.delete',
          arguments: { sheet: 'Budget' },
        },
        {
          runtime: () => runtime,
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.sheet.delete',
      ok: false,
      error: 'invalid_arguments',
      message: 'A workbook must retain at least one worksheet.',
    })
    expect(deleted).toBe(false)
  })

  it('deletes the specified worksheet through xlsx.sheet.delete', async () => {
    const budget = { getSheetName: () => 'Budget' }
    const forecast = { getSheetName: () => 'Forecast' }
    let deleted: typeof budget | typeof forecast | null = null
    const workbook = {
      getSheets: () => [budget, forecast],
      setActiveSheet: () => undefined,
      deleteSheet: (worksheet: typeof budget | typeof forecast) => {
        deleted = worksheet
        return true
      },
    }
    const runtime = { univerAPI: { getActiveWorkbook: () => workbook } }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.sheet.delete',
          arguments: { sheet: 'Forecast' },
        },
        {
          runtime: () => runtime,
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.sheet.delete',
      ok: true,
      output: { sheet: 'Forecast' },
    })
    expect(deleted).toBe(forecast)
  })

  it('rejects an out-of-range xlsx.sheet.move position without mutating the workbook', async () => {
    let moved = false
    const budget = { getSheetName: () => 'Budget' }
    const forecast = { getSheetName: () => 'Forecast' }
    const workbook = {
      getSheets: () => [budget, forecast],
      setActiveSheet: () => undefined,
      moveSheet: () => {
        moved = true
      },
    }
    const runtime = { univerAPI: { getActiveWorkbook: () => workbook } }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.sheet.move',
          arguments: { sheet: 'Forecast', position: 3 },
        },
        {
          runtime: () => runtime,
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.sheet.move',
      ok: false,
      error: 'invalid_arguments',
      message: 'xlsx.sheet.move requires a valid 1-based position.',
    })
    expect(moved).toBe(false)
  })

  it('moves the specified worksheet through xlsx.sheet.move using a 1-based position', async () => {
    const budget = { getSheetName: () => 'Budget' }
    const forecast = { getSheetName: () => 'Forecast' }
    let moved: { worksheet: typeof budget | typeof forecast; index: number } | null = null
    const workbook = {
      getSheets: () => [budget, forecast],
      setActiveSheet: () => undefined,
      moveSheet: (worksheet: typeof budget | typeof forecast, index: number) => {
        moved = { worksheet, index }
      },
    }
    const runtime = { univerAPI: { getActiveWorkbook: () => workbook } }

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.sheet.move',
          arguments: { sheet: 'Forecast', position: 1 },
        },
        {
          runtime: () => runtime,
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.sheet.move',
      ok: true,
      output: { sheet: 'Forecast', position: 1 },
    })
    expect(moved).toEqual({ worksheet: forecast, index: 0 })
  })

  it('loads staged XLSX bytes through xlsx.document.load_staged without checkpointing recovery', async () => {
    const data = new Uint8Array([80, 75, 3, 4]).buffer as ArrayBuffer
    let loaded: { name: string; data: ArrayBuffer } | null = null

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.document.load_staged',
          arguments: {
            blobId: 'xlsx-blob',
            name: 'budget.xlsx',
            size: data.byteLength,
            data,
          },
        },
        {
          runtime: () => null,
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
          loadStaged: async (staged) => {
            loaded = staged
          },
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.document.load_staged',
      ok: true,
      output: { opened: true, fileName: 'budget.xlsx' },
      checkpointRecovery: false,
    })
    expect(loaded).toEqual({ name: 'budget.xlsx', data })
  })

  it('rejects a staged XLSX load without hydrated ArrayBuffer bytes', async () => {
    let loaded = false

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.document.load_staged',
          arguments: {
            blobId: 'xlsx-blob',
            name: 'budget.xlsx',
            size: 4,
            data: {},
          },
        },
        {
          runtime: () => null,
          loadStaged: async () => {
            loaded = true
          },
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.document.load_staged',
      ok: false,
      error: 'invalid_arguments',
      message: '$.data must be a hydrated ArrayBuffer.',
    })
    expect(loaded).toBe(false)
  })

  it('rejects staged XLSX bytes whose size does not match the descriptor', async () => {
    const data = new Uint8Array([80, 75, 3, 4]).buffer as ArrayBuffer
    let loaded = false

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.document.load_staged',
          arguments: {
            blobId: 'xlsx-blob',
            name: 'budget.xlsx',
            size: data.byteLength + 1,
            data,
          },
        },
        {
          runtime: () => null,
          loadStaged: async () => {
            loaded = true
          },
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.document.load_staged',
      ok: false,
      error: 'invalid_arguments',
      message: 'xlsx.document.load_staged requires a valid staged XLSX descriptor.',
    })
    expect(loaded).toBe(false)
  })

  it('rejects staged workbook bytes with a non-XLSX file name', async () => {
    const data = new Uint8Array([80, 75, 3, 4]).buffer as ArrayBuffer
    let loaded = false

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.document.load_staged',
          arguments: {
            blobId: 'xlsx-blob',
            name: 'budget.xls',
            size: data.byteLength,
            data,
          },
        },
        {
          runtime: () => null,
          loadStaged: async () => {
            loaded = true
          },
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.document.load_staged',
      ok: false,
      error: 'invalid_arguments',
      message: 'xlsx.document.load_staged requires a valid staged XLSX descriptor.',
    })
    expect(loaded).toBe(false)
  })

  it('reports a staged XLSX load failure as execution_failed', async () => {
    const data = new Uint8Array([80, 75, 3, 4]).buffer as ArrayBuffer

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.document.load_staged',
          arguments: {
            blobId: 'xlsx-blob',
            name: 'budget.xlsx',
            size: data.byteLength,
            data,
          },
        },
        {
          runtime: () => null,
          loadStaged: async () => {
            throw new Error('Workbook parse failed.')
          },
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.document.load_staged',
      ok: false,
      error: 'execution_failed',
      message: 'Workbook parse failed.',
    })
  })

  it('rejects the retired open_local_file transport alias without loading XLSX bytes', async () => {
    const data = new Uint8Array([80, 75, 3, 4]).buffer as ArrayBuffer
    let loaded = false

    await expect(
      executeXlsxOperation(
        {
          operation: 'open_local_file',
          arguments: {
            blobId: 'xlsx-blob',
            name: 'legacy.XLSX',
            size: data.byteLength,
            data,
          },
        },
        {
          runtime: () => null,
          loadStaged: async () => {
            loaded = true
          },
          save: async () => ({ ok: true, fileName: 'budget.xlsx' }),
        },
      ),
    ).resolves.toEqual({ handled: false })
    expect(loaded).toBe(false)
  })

  it('saves through xlsx.document.save and returns the persisted file identity', async () => {
    let saveCalls = 0

    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.document.save',
          arguments: {},
        },
        {
          runtime: () => null,
          loadStaged: async () => undefined,
          save: async () => {
            saveCalls += 1
            return { ok: true, fileName: 'budget.xlsx' }
          },
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.document.save',
      ok: true,
      output: { saved: true, fileName: 'budget.xlsx' },
      checkpointRecovery: false,
    })
    expect(saveCalls).toBe(1)
  })

  it('reports a canceled workbook save as execution_failed', async () => {
    await expect(
      executeXlsxOperation(
        {
          operation: 'xlsx.document.save',
          arguments: {},
        },
        {
          runtime: () => null,
          loadStaged: async () => undefined,
          save: async () => ({ ok: false, message: 'Save canceled.' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'xlsx.document.save',
      ok: false,
      error: 'execution_failed',
      message: 'Save canceled.',
    })
  })
})
