import { BooleanNumber } from '@univerjs/core'
import { describe, expect, it } from 'vitest'

import { handleRibbonCommand, type RibbonCommandContext } from '../src/renderer/ribbon-actions'
import { applyPageSetupState } from '../src/gateway/xlsx-page-setup'
import { workbookPageSetupStateSchema } from '../src/shared/desktop-api'

function makeHarness() {
  const config = {
    rowHeader: { hidden: BooleanNumber.FALSE },
    columnHeader: { hidden: BooleanNumber.FALSE },
  }
  const skeleton = { rowHeaderWidth: 46 }
  const viewLeftTop = { width: 46 }
  const viewMain = { left: 46 }
  const viewColumnRight = {
    left: 46,
    setViewportSize(props: { left?: number }) {
      if (props.left !== undefined) this.left = props.left
    },
  }
  const scene = {
    getViewport: (key: string) =>
      ({ viewMain, viewColumnRight, viewLeftTop })[key as 'viewMain'],
    makeDirty: () => {},
  }
  const render = { scene, with: () => ({ getCurrentSkeleton: () => skeleton }) }
  const worksheet = {
    getSheetId: () => 'sheet1',
    getSheet: () => ({ getConfig: () => config }),
  }
  const workbook = { getId: () => 'wb1', getActiveSheet: () => worksheet }
  const univerAPI = {
    getActiveWorkbook: () => workbook,
    executeCommand: async (id: string, params: { size: number }) => {
      if (id === 'sheet.command.set-row-header-width') {
        skeleton.rowHeaderWidth = params.size
        const delta = params.size - (viewLeftTop.width || 46)
        viewMain.left += delta
        viewColumnRight.setViewportSize({ left: viewColumnRight.left + delta })
      }
      return true
    },
  }
  const ctx = {
    univerRef: {
      current: {
        univerAPI,
        univer: { __getInjector: () => ({ get: () => ({ getRenderById: () => render }) }) },
      },
    },
    lazyWorkbookRef: { current: null },
    setMessage: () => {},
    setPendingEdits: () => {},
  } as unknown as RibbonCommandContext
  return { ctx, config, viewMain, viewColumnRight }
}

const settle = () => new Promise((resolve) => setTimeout(resolve, 0))

describe('toggle-headings viewport offsets', () => {
  it('hides and re-shows without drifting the grid', async () => {
    const { ctx, config, viewMain, viewColumnRight } = makeHarness()

    handleRibbonCommand(ctx, 'toggle-headings')
    await settle()
    expect(config.rowHeader.hidden).toBe(BooleanNumber.TRUE)
    expect(viewMain.left).toBe(0)
    expect(viewColumnRight.left).toBe(0)

    handleRibbonCommand(ctx, 'toggle-headings')
    await settle()
    expect(config.rowHeader.hidden).toBe(BooleanNumber.FALSE)
    expect(viewMain.left).toBe(46)
    expect(viewColumnRight.left).toBe(46)
  })

  it('persists the hidden state as sheetView metadata', () => {
    const xml = '<worksheet><sheetViews><sheetView workbookViewId="0"/></sheetViews></worksheet>'
    expect(applyPageSetupState(xml, { sheetName: 'Sheet1', showHeadings: false })).toContain(
      'showRowColHeaders="0"',
    )
    expect(
      applyPageSetupState(
        '<worksheet><sheetViews><sheetView workbookViewId="0" showRowColHeaders="0"/></sheetViews></worksheet>',
        { sheetName: 'Sheet1', showHeadings: true },
      ),
    ).not.toContain('showRowColHeaders')
  })

  it('accepts heading visibility on the browser save wire', () => {
    expect(
      workbookPageSetupStateSchema.parse({ sheetId: 'sheet1', showHeadings: false }),
    ).toEqual({ sheetId: 'sheet1', showHeadings: false })
  })
})
