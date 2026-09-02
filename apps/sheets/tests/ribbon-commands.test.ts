import { describe, expect, it } from 'vitest'

import { HorizontalAlign } from '@univerjs/core'

import { createEditJournal } from '../src/renderer/edit-journal'
import {
  handleRibbonCommand,
  parseStyleCommand,
  type RibbonCommandContext,
} from '../src/renderer/ribbon-actions'
import { t } from '../src/renderer/i18n/locale'
import type { LazyWorkbookState } from '../src/renderer/univer-state'

describe('parseStyleCommand', () => {
  it('splits plain name:argument commands', () => {
    expect(parseStyleCommand('fill:#00aa55')).toEqual({
      name: 'fill',
      argument: '#00aa55',
      extra: '',
    })
    expect(parseStyleCommand('align:right')).toEqual({
      name: 'align',
      argument: 'right',
      extra: '',
    })
  })

  it('handles commands without an argument', () => {
    expect(parseStyleCommand('bold')).toEqual({ name: 'bold', argument: '', extra: '' })
  })

  it('keeps colons inside number-format patterns intact', () => {
    expect(parseStyleCommand('format:h:mm:ss AM/PM')).toEqual({
      name: 'format',
      argument: 'h:mm:ss AM/PM',
      extra: '',
    })
    expect(parseStyleCommand('format:hh:mm:ss')).toEqual({
      name: 'format',
      argument: 'hh:mm:ss',
      extra: '',
    })
  })

  it('still splits the extra segment for three-part commands', () => {
    expect(parseStyleCommand('border:all:#112233')).toEqual({
      name: 'border',
      argument: 'all',
      extra: '#112233',
    })
    expect(parseStyleCommand('cellprot:locked-on:hidden-off')).toEqual({
      name: 'cellprot',
      argument: 'locked-on',
      extra: 'hidden-off',
    })
    expect(parseStyleCommand('cellprot:locked-on:')).toEqual({
      name: 'cellprot',
      argument: 'locked-on',
      extra: '',
    })
    expect(parseStyleCommand('sort-custom:1:2a,3d')).toEqual({
      name: 'sort-custom',
      argument: '1',
      extra: '2a,3d',
    })
  })
})

/// Mirrors transformFacadeHorizontalAlignment in @univerjs/sheets' facade:
/// only these three values are accepted, and 'normal' means RIGHT — anything
/// else (such as a raw 'right') hits the throwing default branch.
function facadeHorizontalAlignment(value: string): HorizontalAlign {
  switch (value) {
    case 'left':
      return HorizontalAlign.LEFT
    case 'center':
      return HorizontalAlign.CENTER
    case 'normal':
      return HorizontalAlign.RIGHT
    default:
      throw new Error(`Invalid horizontal alignment: ${value}`)
  }
}

interface DispatchHarness {
  ctx: RibbonCommandContext
  model: { ht?: HorizontalAlign; numberFormat?: string }
  messages: string[]
}

function makeDispatchHarness(): DispatchHarness {
  const model: DispatchHarness['model'] = {}
  const messages: string[] = []
  const range = {
    getCellStyleData: () => ({}),
    setHorizontalAlignment: (value: string) => {
      model.ht = facadeHorizontalAlignment(value)
    },
    setNumberFormat: (pattern: string) => {
      model.numberFormat = pattern
    },
  }
  const workbook = {
    getActiveSheet: () => undefined,
    getActiveRange: () => range,
  }
  const ctx = {
    univerRef: { current: { univerAPI: { getActiveWorkbook: () => workbook } } },
    setMessage: (message: string) => {
      messages.push(message)
    },
  } as unknown as RibbonCommandContext
  return { ctx, model, messages }
}

describe('handleRibbonCommand alignment', () => {
  it.each([
    ['align:left', HorizontalAlign.LEFT],
    ['align:center', HorizontalAlign.CENTER],
    ['align:right', HorizontalAlign.RIGHT],
  ])('%s lands in the model', (command, expected) => {
    const { ctx, model, messages } = makeDispatchHarness()
    handleRibbonCommand(ctx, command)
    expect(model.ht).toBe(expected)
    // The dispatcher reports the applied-style message, not a caught error.
    expect(messages).toHaveLength(1)
    expect(messages[0]).not.toMatch(/Invalid horizontal alignment/)
  })
})

describe('handleRibbonCommand number format', () => {
  it('passes patterns containing colons through unclipped', () => {
    const { ctx, model } = makeDispatchHarness()
    handleRibbonCommand(ctx, 'format:h:mm:ss AM/PM')
    expect(model.numberFormat).toBe('h:mm:ss AM/PM')
  })
})

describe('handleRibbonCommand aggregate formula', () => {
  it('inserts one formula below each selected column through the shared workbook seam', () => {
    const formulas = new Map<string, string>()
    const sourceRange = {
      getCellStyleData: () => ({}),
      getRow: () => 1,
      getColumn: () => 1,
      getHeight: () => 3,
      getWidth: () => 2,
    }
    const worksheet = {
      getSheetId: () => 'sheet-budget',
      getRange: (row: number, column: number) => ({
        setFormula: (formula: string) => formulas.set(`${row}:${column}`, formula),
      }),
    }
    const messages: string[] = []
    const ctx = {
      univerRef: {
        current: {
          univerAPI: {
            getActiveWorkbook: () => ({
              getActiveSheet: () => worksheet,
              getActiveRange: () => sourceRange,
            }),
          },
        },
      },
      lazyWorkbookRef: { current: null },
      setMessage: (message: string) => messages.push(message),
    } as unknown as RibbonCommandContext

    handleRibbonCommand(ctx, 'autofn:COUNT')

    expect([...formulas.entries()]).toEqual([
      ['4:1', '=COUNT(B2:B4)'],
      ['4:2', '=COUNT(C2:C4)'],
    ])
    expect(messages).toHaveLength(1)
  })
})

describe('handleRibbonCommand flash fill', () => {
  it('fills only empty target cells from retained examples', () => {
    let written: unknown[][] | null = null
    const selectedRange = {
      getCellStyleData: () => ({}),
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
      getSheetId: () => 'sheet-budget',
      getRange: (...arguments_: unknown[]) => {
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
    const messages: string[] = []
    const ctx = {
      univerRef: {
        current: {
          univerAPI: {
            getActiveWorkbook: () => ({
              getActiveSheet: () => worksheet,
              getActiveRange: () => selectedRange,
            }),
          },
        },
      },
      lazyWorkbookRef: { current: null },
      setMessage: (message: string) => messages.push(message),
    } as unknown as RibbonCommandContext

    handleRibbonCommand(ctx, 'flash-fill')

    expect(written).toEqual([['John·Sales'], ['Mary·Marketing'], ['Bob·Finance']])
    expect(messages).toHaveLength(1)
  })

  it('rejects partially streamed source rows before reading them', () => {
    let read = false
    const selectedRange = {
      getCellStyleData: () => ({}),
      getRow: () => 0,
      getColumn: () => 2,
      getHeight: () => 3,
      getWidth: () => 1,
    }
    const worksheet = {
      getSheetId: () => 'sheet-budget',
      getRange: () => ({
        getValues: () => {
          read = true
          return []
        },
      }),
    }
    const state = {
      file: { sheets: [{ id: 'sheet-budget', rowCount: 10, columnCount: 5 }] },
      loadedRanges: new Map([
        ['sheet-budget', { startRow: 0, endRow: 1, startColumn: 0, endColumn: 4 }],
      ]),
      flags: { preloadComplete: false },
    } as unknown as LazyWorkbookState
    const messages: string[] = []
    const ctx = {
      univerRef: {
        current: {
          univerAPI: {
            getActiveWorkbook: () => ({
              getActiveSheet: () => worksheet,
              getActiveRange: () => selectedRange,
            }),
          },
        },
      },
      lazyWorkbookRef: { current: state },
      setMessage: (message: string) => messages.push(message),
    } as unknown as RibbonCommandContext

    handleRibbonCommand(ctx, 'flash-fill')

    expect(read).toBe(false)
    expect(messages).toEqual([t('appAreaStreaming')])
  })
})

describe('handleRibbonCommand text to columns', () => {
  it('splits one selected column with the retained delimiter flag', () => {
    const commands: Array<{ command: string; parameters: unknown }> = []
    const selectedRange = {
      getCellStyleData: () => ({}),
      getRow: () => 1,
      getColumn: () => 2,
      getHeight: () => 3,
      getWidth: () => 1,
    }
    const worksheet = { getSheetId: () => 'sheet-budget' }
    const workbook = {
      getId: () => 'workbook-budget',
      getActiveSheet: () => worksheet,
      getActiveRange: () => selectedRange,
    }
    const messages: string[] = []
    const ctx = {
      univerRef: {
        current: {
          univerAPI: {
            getActiveWorkbook: () => workbook,
            executeCommand: (command: string, parameters: unknown) => {
              commands.push({ command, parameters })
              return Promise.resolve(true)
            },
          },
        },
      },
      lazyWorkbookRef: { current: null },
      setMessage: (message: string) => messages.push(message),
    } as unknown as RibbonCommandContext

    handleRibbonCommand(ctx, 'text-to-columns:4')

    expect(commands).toHaveLength(1)
    expect(commands[0]).toMatchObject({
      command: 'sheet.command.split-text-to-columns',
      parameters: {
        range: { startRow: 1, endRow: 3, startColumn: 2, endColumn: 2 },
        delimiter: 4,
      },
    })
    expect(messages).toHaveLength(1)
  })
})

describe('handleRibbonCommand row height', () => {
  it('sets the selected rows to the retained point height', () => {
    const heights: Array<{ startRow: number; count: number; heightPixels: number }> = []
    const worksheet = {
      setRowHeightsForced: (startRow: number, count: number, heightPixels: number) => {
        heights.push({ startRow, count, heightPixels })
      },
    }
    const workbook = {
      getActiveSheet: () => worksheet,
      getActiveRange: () => ({
        getRange: () => ({ startRow: 1, endRow: 2, startColumn: 0, endColumn: 1 }),
      }),
    }
    const ctx = {
      univerRef: { current: { univerAPI: { getActiveWorkbook: () => workbook } } },
      setMessage: () => undefined,
    } as unknown as RibbonCommandContext

    handleRibbonCommand(ctx, 'row-height:24.75')

    expect(heights).toEqual([{ startRow: 1, count: 2, heightPixels: 33 }])
  })
})

describe('handleRibbonCommand column width', () => {
  it('sets the selected columns to the retained character width', () => {
    const widths: Array<{ startColumn: number; count: number; widthPixels: number }> = []
    const worksheet = {
      setColumnWidths: (startColumn: number, count: number, widthPixels: number) => {
        widths.push({ startColumn, count, widthPixels })
      },
    }
    const workbook = {
      getActiveSheet: () => worksheet,
      getActiveRange: () => ({
        getRange: () => ({ startRow: 0, endRow: 1, startColumn: 2, endColumn: 3 }),
      }),
    }
    const ctx = {
      univerRef: { current: { univerAPI: { getActiveWorkbook: () => workbook } } },
      setMessage: () => undefined,
    } as unknown as RibbonCommandContext

    handleRibbonCommand(ctx, 'col-width:12.5')

    expect(widths).toEqual([{ startColumn: 2, count: 2, widthPixels: 92 }])
  })
})

describe('handleRibbonCommand freeze panes', () => {
  it.each([
    ['freeze-here', { startRow: 2, startColumn: 1, xSplit: 1, ySplit: 2 }],
    ['freeze-top-row', { startRow: 1, startColumn: -1, xSplit: 0, ySplit: 1 }],
    ['freeze-first-col', { startRow: -1, startColumn: 1, xSplit: 1, ySplit: 0 }],
  ] as const)('sets %s through the retained worksheet freeze route', (command, expected) => {
    const freezes: Array<{
      startRow: number
      startColumn: number
      xSplit: number
      ySplit: number
    }> = []
    const worksheet = {
      setFreeze: (freeze: {
        startRow: number
        startColumn: number
        xSplit: number
        ySplit: number
      }) => freezes.push(freeze),
    }
    const workbook = {
      getActiveSheet: () => worksheet,
      getActiveRange: () => ({ getRow: () => 2, getColumn: () => 1 }),
    }
    const ctx = {
      univerRef: { current: { univerAPI: { getActiveWorkbook: () => workbook } } },
      setMessage: () => undefined,
    } as unknown as RibbonCommandContext

    handleRibbonCommand(ctx, command)

    expect(freezes).toEqual([expected])
  })

  it('removes frozen panes through the retained worksheet route', () => {
    let canceled = 0
    const worksheet = {
      cancelFreeze: () => {
        canceled += 1
      },
    }
    const workbook = { getActiveSheet: () => worksheet }
    const ctx = {
      univerRef: { current: { univerAPI: { getActiveWorkbook: () => workbook } } },
      setMessage: () => undefined,
    } as unknown as RibbonCommandContext

    handleRibbonCommand(ctx, 'unfreeze')

    expect(canceled).toBe(1)
  })
})

describe('handleRibbonCommand worksheet gridlines', () => {
  it('toggles to the opposite explicit gridline visibility', () => {
    const hiddenStates: boolean[] = []
    const worksheet = {
      hasHiddenGridLines: () => false,
      getSheetId: () => 'sheet-1',
      setHiddenGridlines: (hidden: boolean) => hiddenStates.push(hidden),
    }
    const workbook = { getActiveSheet: () => worksheet }
    const ctx = {
      univerRef: { current: { univerAPI: { getActiveWorkbook: () => workbook } } },
      lazyWorkbookRef: { current: null },
      setMessage: () => undefined,
    } as unknown as RibbonCommandContext

    handleRibbonCommand(ctx, 'toggle-gridlines')

    expect(hiddenStates).toEqual([true])
  })

  it('toggles to the opposite explicit formula-view state', () => {
    const editJournal = createEditJournal()
    const contextValues = new Map<unknown, unknown>()
    const worksheet = {
      getSheetId: () => 'sheet-1',
    }
    const workbook = { getId: () => undefined, getActiveSheet: () => worksheet }
    const state = {
      editJournal,
      showFormulaSheets: new Set<string>(),
    }
    let pendingEdits = 0
    const ctx = {
      univerRef: {
        current: {
          univerAPI: { getActiveWorkbook: () => workbook },
          univer: {
            __getInjector: () => ({
              get: () => ({
                getContextValue: (key: unknown) => contextValues.get(key),
                setContextValue: (key: unknown, value: unknown) => contextValues.set(key, value),
              }),
            }),
          },
        },
      },
      lazyWorkbookRef: { current: state },
      setPendingEdits: (count: number) => {
        pendingEdits = count
      },
      setMessage: () => undefined,
    } as unknown as RibbonCommandContext

    handleRibbonCommand(ctx, 'toggle-show-formulas')

    expect(state.showFormulaSheets).toEqual(new Set(['sheet-1']))
    expect(editJournal.pageSetup.get('sheet-1')).toEqual({ showFormulas: true })
    expect(pendingEdits).toBe(1)
  })
})

describe('handleRibbonCommand worksheet filters', () => {
  it('resolves the selected range before enabling the worksheet filter', () => {
    const commands: Array<{ id: string; params: unknown }> = []
    const selection = { startRow: 0, endRow: 3, startColumn: 0, endColumn: 2 }
    const worksheet = {
      getSheetId: () => 'sheet-1',
      getFilter: () => null,
    }
    const workbook = {
      getId: () => 'workbook-1',
      getActiveSheet: () => worksheet,
      getActiveRange: () => ({ getRange: () => selection }),
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
      setMessage: () => undefined,
    } as unknown as RibbonCommandContext

    handleRibbonCommand(ctx, 'filter-toggle')

    expect(commands).toEqual([
      {
        id: 'sheet.command.set-filter-range',
        params: { unitId: 'workbook-1', subUnitId: 'sheet-1', range: selection },
      },
    ])
  })

  it('clears criteria from the explicit active filter while retaining its range', () => {
    const commands: Array<{ id: string; params: unknown }> = []
    const selection = { startRow: 0, endRow: 3, startColumn: 0, endColumn: 2 }
    const worksheet = {
      getSheetId: () => 'sheet-1',
      getFilter: () => ({ getRange: () => ({ getRange: () => selection }) }),
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
      setMessage: () => undefined,
    } as unknown as RibbonCommandContext

    handleRibbonCommand(ctx, 'filter-clear')

    expect(commands).toEqual([
      {
        id: 'sheet.command.clear-filter-criteria',
        params: { unitId: 'workbook-1', subUnitId: 'sheet-1' },
      },
    ])
  })
})
