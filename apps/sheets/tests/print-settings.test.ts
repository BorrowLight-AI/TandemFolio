import { describe, expect, it } from 'vitest'

import { resolveEffectivePageSetup } from '../src/renderer/print-settings'
import { buildSheetPrintPayload, type PrintWorksheet } from '../src/renderer/print-html'

describe('resolveEffectivePageSetup', () => {
  it('defaults an untouched worksheet to A4 portrait and normal margins', () => {
    expect(resolveEffectivePageSetup({}, null, null)).toMatchObject({
      orientation: 'portrait',
      paperSize: 9,
      scale: 100,
      fitToPage: false,
      fitToWidth: 0,
      fitToHeight: 0,
      margins: {
        left: 0.7,
        right: 0.7,
        top: 0.75,
        bottom: 0.75,
        header: 0.3,
        footer: 0.3,
      },
      printAreas: [],
      printTitles: null,
      header: null,
      footer: null,
    })
  })

  it('uses exact saved margins and first/even header variants when the session is untouched', () => {
    const setup = resolveEffectivePageSetup(
      {},
      {
        orientation: 'landscape',
        paperSize: 1,
        scale: 65,
        margins: { left: 0.98, right: 0.98, top: 5, bottom: 0.79, header: 0, footer: 0.51 },
        printGridlines: true,
        oddFooter: '&CPage &P of &N',
        differentOddEven: true,
        differentFirst: true,
        evenHeader: '&CEven',
        firstFooter: '&RFirst &D',
      },
      { printArea: "'S'!$A$1:$K$84", printTitles: "'S'!$17:$17" },
    )

    expect(setup).toMatchObject({
      orientation: 'landscape',
      paperSize: 1,
      scale: 65,
      margins: { left: 0.98, right: 0.98, top: 3, bottom: 0.79, header: 0, footer: 0.51 },
      printGridlines: true,
      printAreas: ['A1:K84'],
      printTitles: '17:17',
      footer: { center: 'Page &P of &N' },
      firstPage: { header: null, footer: { right: 'First &D' } },
      evenPages: { header: { center: 'Even' }, footer: null },
    })
  })

  it('projects exact margins and page-specific header/footer templates into print output', () => {
    const setup = resolveEffectivePageSetup(
      {},
      {
        margins: { left: 0.9, right: 0.8, top: 0.7, bottom: 0.6, header: 0.2, footer: 0.25 },
        oddHeader: '&L&A&CPage &P of &N',
        differentFirst: true,
        firstFooter: '&RFirst &D',
      },
      null,
    )
    const worksheet = {
      getLastRow: () => 0,
      getLastColumn: () => 0,
      getRowHeight: () => 20,
      getColumnWidth: () => 100,
      getMergedRanges: () => [],
      getRange: (...args: number[]) =>
        args.length === 2
          ? { getCellStyleData: () => null }
          : { getDisplayValues: () => [['Value']], getValues: () => [['Value']] },
    } as unknown as PrintWorksheet

    const payload = buildSheetPrintPayload(worksheet, setup, 'Book.pdf', 'Budget')

    expect(payload.margins).toEqual({ left: 0.9, right: 0.8, top: 0.7, bottom: 0.6 })
    expect(payload.headerTemplate).toContain('Budget')
    expect(payload.headerTemplate).toContain('class="pageNumber"')
    expect(payload.headerTemplate).toContain('class="totalPages"')
    expect(payload.firstPage?.footerTemplate).toContain('First')
  })

  it('prints every saved print area on a separate page', () => {
    const setup = resolveEffectivePageSetup({}, null, {
      printArea: "'Budget'!$A$1:$A$1,'Budget'!$C$3:$C$3",
    })
    const worksheet = {
      getLastRow: () => 2,
      getLastColumn: () => 2,
      getRowHeight: () => 20,
      getColumnWidth: () => 100,
      getMergedRanges: () => [],
      getRange: (...args: number[]) =>
        args.length === 2
          ? { getCellStyleData: () => null }
          : { getDisplayValues: () => [['Value']], getValues: () => [['Value']] },
    } as unknown as PrintWorksheet

    const payload = buildSheetPrintPayload(worksheet, setup, 'Book.pdf', 'Budget')

    expect(payload.html.match(/<table>/g)).toHaveLength(2)
    expect(payload.html).toContain('table + table { break-before: page; }')
  })
})
