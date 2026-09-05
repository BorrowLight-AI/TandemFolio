import { describe, expect, it } from 'vitest'

import { createEditJournal } from '../src/renderer/edit-journal'
import {
  handlePageLayoutCommand,
  type PageLayoutContext,
} from '../src/renderer/page-layout-actions'

describe('page layout actions', () => {
  it('applies a native theme preset to the live file state and save journal', () => {
    const editJournal = createEditJournal()
    const worksheet = { getSheetId: () => 'sheet-1' }
    const workbook = { getId: () => undefined, getActiveSheet: () => worksheet }
    let pendingEdits = 0
    const state = {
      editJournal,
      file: {
        themeColors: Array.from({ length: 12 }, () => '#000000'),
        themeFonts: { major: 'Old Major', minor: 'Old Minor' },
        styles: [],
      },
    }
    const ctx = {
      univerRef: { current: { univerAPI: { getActiveWorkbook: () => workbook } } },
      lazyWorkbookRef: { current: state },
      setMessage: () => undefined,
      setPendingEdits: (count: number) => {
        pendingEdits = count
      },
    } as unknown as PageLayoutContext

    handlePageLayoutCommand(ctx, 'theme:forest')

    expect(state.file.themeColors[4]).toBe('#217346')
    expect(state.file.themeFonts).toEqual({ major: 'Candara', minor: 'Candara' })
    expect(editJournal.theme).toMatchObject({
      colors: { name: 'Forest' },
      fonts: { name: 'Candara', major: 'Candara', minor: 'Candara' },
    })
    expect(pendingEdits).toBe(2)
  })

  it('sets the selected worksheet orientation to an explicit final state', () => {
    const editJournal = createEditJournal()
    const worksheet = { getSheetId: () => 'sheet-1' }
    const workbook = { getId: () => undefined, getActiveSheet: () => worksheet }
    let pendingEdits = 0
    const ctx = {
      univerRef: { current: { univerAPI: { getActiveWorkbook: () => workbook } } },
      lazyWorkbookRef: { current: { editJournal } },
      setMessage: () => undefined,
      setPendingEdits: (count: number) => {
        pendingEdits = count
      },
    } as unknown as PageLayoutContext

    handlePageLayoutCommand(ctx, 'orientation:landscape')

    expect(editJournal.pageSetup.get('sheet-1')).toEqual({ orientation: 'landscape' })
    expect(pendingEdits).toBe(1)
  })

  it('keeps the effective width when the Ribbon changes only fit-to-page height', () => {
    const editJournal = createEditJournal()
    const worksheet = { getSheetId: () => 'sheet-1' }
    const workbook = { getId: () => undefined, getActiveSheet: () => worksheet }
    const ctx = {
      univerRef: { current: { univerAPI: { getActiveWorkbook: () => workbook } } },
      lazyWorkbookRef: {
        current: {
          editJournal,
          file: {
            sheets: [
              {
                id: 'sheet-1',
                name: 'Budget',
                pageSetup: { fitToWidth: 2, fitToHeight: 3 },
              },
            ],
          },
        },
      },
      setMessage: () => undefined,
      setPendingEdits: () => undefined,
    } as unknown as PageLayoutContext

    handlePageLayoutCommand(ctx, 'fit-height:4')

    expect(editJournal.pageSetup.get('sheet-1')).toEqual({
      fitToWidth: 2,
      fitToHeight: 4,
      fitToPage: true,
    })
  })

  it('changes fixed print scale and fit mode as one renderer-owned undo unit', () => {
    const editJournal = createEditJournal()
    let registeredCommand:
      | {
          handler: (
            accessor: unknown,
            params?: { token: number; direction: 'undo' | 'redo' },
          ) => boolean
        }
      | undefined
    let undoItem:
      | {
          undoMutations: Array<{
            params: { token: number; direction: 'undo' | 'redo' }
          }>
        }
      | undefined
    const worksheet = { getSheetId: () => 'sheet-1' }
    const workbook = { getId: () => 'workbook-1', getActiveSheet: () => worksheet }
    const commandServices = {
      registerCommand: (command: NonNullable<typeof registeredCommand>) => {
        registeredCommand = command
      },
      pushUndoRedo: (item: NonNullable<typeof undoItem>) => {
        undoItem = item
      },
    }
    const ctx = {
      univerRef: {
        current: {
          univerAPI: { getActiveWorkbook: () => workbook },
          univer: { __getInjector: () => ({ get: () => commandServices }) },
        },
      },
      lazyWorkbookRef: {
        current: {
          editJournal,
          file: {
            sheets: [
              {
                id: 'sheet-1',
                name: 'Budget',
                pageSetup: { scale: 100, fitToPage: true },
              },
            ],
          },
        },
      },
      setMessage: () => undefined,
      setPendingEdits: () => undefined,
    } as unknown as PageLayoutContext

    handlePageLayoutCommand(ctx, 'scale:80')

    expect(editJournal.pageSetup.get('sheet-1')).toEqual({ scale: 80, fitToPage: false })
    expect(registeredCommand).toBeDefined()
    expect(undoItem).toBeDefined()
    registeredCommand!.handler(undefined, undoItem!.undoMutations[0]!.params)
    expect(editJournal.pageSetup.has('sheet-1')).toBe(false)
  })

  it('resolves the selected print area before entering renderer-owned undo', () => {
    const editJournal = createEditJournal()
    let registeredCommand:
      | {
          handler: (
            accessor: unknown,
            params?: { token: number; direction: 'undo' | 'redo' },
          ) => boolean
        }
      | undefined
    let undoItem:
      | {
          undoMutations: Array<{
            params: { token: number; direction: 'undo' | 'redo' }
          }>
        }
      | undefined
    const worksheet = { getSheetId: () => 'sheet-1' }
    const workbook = {
      getId: () => 'workbook-1',
      getActiveSheet: () => worksheet,
      getActiveRange: () => ({
        getColumn: () => 1,
        getRow: () => 1,
        getWidth: () => 3,
        getHeight: () => 7,
      }),
    }
    const commandServices = {
      registerCommand: (command: NonNullable<typeof registeredCommand>) => {
        registeredCommand = command
      },
      pushUndoRedo: (item: NonNullable<typeof undoItem>) => {
        undoItem = item
      },
    }
    const ctx = {
      univerRef: {
        current: {
          univerAPI: { getActiveWorkbook: () => workbook },
          univer: { __getInjector: () => ({ get: () => commandServices }) },
        },
      },
      lazyWorkbookRef: { current: { editJournal } },
      setMessage: () => undefined,
      setPendingEdits: () => undefined,
    } as unknown as PageLayoutContext

    handlePageLayoutCommand(ctx, 'print-area:set')

    expect(editJournal.pageSetup.get('sheet-1')).toEqual({ printArea: 'B2:D8' })
    expect(registeredCommand).toBeDefined()
    expect(undoItem).toBeDefined()
    registeredCommand!.handler(undefined, undoItem!.undoMutations[0]!.params)
    expect(editJournal.pageSetup.has('sheet-1')).toBe(false)
  })

  it('resolves selected print-title rows before entering renderer-owned undo', () => {
    const editJournal = createEditJournal()
    let registeredCommand:
      | {
          handler: (
            accessor: unknown,
            params?: { token: number; direction: 'undo' | 'redo' },
          ) => boolean
        }
      | undefined
    let undoItem:
      | {
          undoMutations: Array<{
            params: { token: number; direction: 'undo' | 'redo' }
          }>
        }
      | undefined
    const worksheet = { getSheetId: () => 'sheet-1' }
    const workbook = {
      getId: () => 'workbook-1',
      getActiveSheet: () => worksheet,
      getActiveRange: () => ({
        getRow: () => 1,
        getHeight: () => 7,
      }),
    }
    const commandServices = {
      registerCommand: (command: NonNullable<typeof registeredCommand>) => {
        registeredCommand = command
      },
      pushUndoRedo: (item: NonNullable<typeof undoItem>) => {
        undoItem = item
      },
    }
    const ctx = {
      univerRef: {
        current: {
          univerAPI: { getActiveWorkbook: () => workbook },
          univer: { __getInjector: () => ({ get: () => commandServices }) },
        },
      },
      lazyWorkbookRef: { current: { editJournal } },
      setMessage: () => undefined,
      setPendingEdits: () => undefined,
    } as unknown as PageLayoutContext

    handlePageLayoutCommand(ctx, 'print-titles:set')

    expect(editJournal.pageSetup.get('sheet-1')).toEqual({ printTitles: '2:8' })
    expect(registeredCommand).toBeDefined()
    expect(undoItem).toBeDefined()
    registeredCommand!.handler(undefined, undoItem!.undoMutations[0]!.params)
    expect(editJournal.pageSetup.has('sheet-1')).toBe(false)
  })

  it('rejects a Ribbon print-title selection wider than the print pipeline limit', () => {
    const editJournal = createEditJournal()
    const messages: string[] = []
    const worksheet = { getSheetId: () => 'sheet-1' }
    const workbook = {
      getId: () => 'workbook-1',
      getActiveSheet: () => worksheet,
      getActiveRange: () => ({
        getRow: () => 0,
        getHeight: () => 22,
      }),
    }
    const ctx = {
      univerRef: { current: { univerAPI: { getActiveWorkbook: () => workbook } } },
      lazyWorkbookRef: { current: { editJournal } },
      setMessage: (message: string) => messages.push(message),
      setPendingEdits: () => undefined,
    } as unknown as PageLayoutContext

    handlePageLayoutCommand(ctx, 'print-titles:set')

    expect(editJournal.pageSetup.size).toBe(0)
    expect(messages.at(-1)).toContain('21')
  })

  it('inserts a manual row and column break before the active cell', () => {
    const editJournal = createEditJournal()
    const worksheet = {
      getSheetId: () => 'sheet-1',
      getMaxRows: () => 1_000,
      getMaxColumns: () => 100,
    }
    const workbook = {
      getId: () => undefined,
      getActiveSheet: () => worksheet,
      getActiveRange: () => ({
        getRow: () => 5,
        getColumn: () => 2,
        getHeight: () => 1,
        getWidth: () => 1,
      }),
    }
    const ctx = {
      univerRef: { current: { univerAPI: { getActiveWorkbook: () => workbook } } },
      lazyWorkbookRef: {
        current: {
          editJournal,
          file: { sheets: [{ id: 'sheet-1', name: 'Budget', pageSetup: {} }] },
        },
      },
      setMessage: () => undefined,
      setPendingEdits: () => undefined,
    } as unknown as PageLayoutContext

    handlePageLayoutCommand(ctx, 'breaks:insert')

    expect(editJournal.pageSetup.get('sheet-1')).toMatchObject({ rowBreaks: [5], colBreaks: [2] })
  })

  it('inserts only a row break for a full-row selection', () => {
    const editJournal = createEditJournal()
    const worksheet = {
      getSheetId: () => 'sheet-1',
      getMaxRows: () => 1_000,
      getMaxColumns: () => 100,
    }
    const workbook = {
      getId: () => undefined,
      getActiveSheet: () => worksheet,
      getActiveRange: () => ({
        getRow: () => 5,
        getColumn: () => 0,
        getHeight: () => 1,
        getWidth: () => 100,
      }),
    }
    const ctx = {
      univerRef: { current: { univerAPI: { getActiveWorkbook: () => workbook } } },
      lazyWorkbookRef: {
        current: {
          editJournal,
          file: { sheets: [{ id: 'sheet-1', name: 'Budget', pageSetup: {} }] },
        },
      },
      setMessage: () => undefined,
      setPendingEdits: () => undefined,
    } as unknown as PageLayoutContext

    handlePageLayoutCommand(ctx, 'breaks:insert')

    expect(editJournal.pageSetup.get('sheet-1')).toMatchObject({ rowBreaks: [5], colBreaks: [] })
  })

  it('inserts only a column break for a full-column selection', () => {
    const editJournal = createEditJournal()
    const worksheet = {
      getSheetId: () => 'sheet-1',
      getMaxRows: () => 1_000,
      getMaxColumns: () => 100,
    }
    const workbook = {
      getId: () => undefined,
      getActiveSheet: () => worksheet,
      getActiveRange: () => ({
        getRow: () => 0,
        getColumn: () => 2,
        getHeight: () => 1_000,
        getWidth: () => 1,
      }),
    }
    const ctx = {
      univerRef: { current: { univerAPI: { getActiveWorkbook: () => workbook } } },
      lazyWorkbookRef: {
        current: {
          editJournal,
          file: { sheets: [{ id: 'sheet-1', name: 'Budget', pageSetup: {} }] },
        },
      },
      setMessage: () => undefined,
      setPendingEdits: () => undefined,
    } as unknown as PageLayoutContext

    handlePageLayoutCommand(ctx, 'breaks:insert')

    expect(editJournal.pageSetup.get('sheet-1')).toMatchObject({ rowBreaks: [], colBreaks: [2] })
  })

  it('does not create an empty page-break edit at the top-left cell', () => {
    const editJournal = createEditJournal()
    const messages: string[] = []
    const worksheet = {
      getSheetId: () => 'sheet-1',
      getMaxRows: () => 1_000,
      getMaxColumns: () => 100,
    }
    const workbook = {
      getId: () => undefined,
      getActiveSheet: () => worksheet,
      getActiveRange: () => ({
        getRow: () => 0,
        getColumn: () => 0,
        getHeight: () => 1,
        getWidth: () => 1,
      }),
    }
    const ctx = {
      univerRef: { current: { univerAPI: { getActiveWorkbook: () => workbook } } },
      lazyWorkbookRef: {
        current: {
          editJournal,
          file: { sheets: [{ id: 'sheet-1', name: 'Budget', pageSetup: {} }] },
        },
      },
      setMessage: (message: string) => messages.push(message),
      setPendingEdits: () => undefined,
    } as unknown as PageLayoutContext

    handlePageLayoutCommand(ctx, 'breaks:insert')

    expect(editJournal.pageSetup.size).toBe(0)
    expect(messages.at(-1)).toContain('第 1 行')
  })
})
