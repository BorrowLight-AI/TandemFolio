import { describe, expect, it } from 'vitest'

import { listWorkbookConditionalFormats } from '../src/renderer/conditional-format-actions'
import { applyConditionalRules, buildConditionalRule } from '../src/renderer/univer-sync'

describe('conditional-format context', () => {
  it('marks empty worksheet conditional-format metadata as loaded', () => {
    const appliedCfSheets = new Set<string>()

    applyConditionalRules(
      { addConditionalFormattingRule: () => undefined } as never,
      { appliedCfSheets, file: { dxfStyles: [] } } as never,
      'sheet-budget',
      [],
    )

    expect(appliedCfSheets).toContain('sheet-budget')
  })

  it('publishes bounded session rule IDs and ranges for later update or removal', () => {
    const workbook = {
      getSheets: () => [
        {
          getSheetName: () => 'Budget',
          getConditionalFormattingRules: () => [
            {
              cfId: 'cf-number',
              ranges: [{ startRow: 0, endRow: 4, startColumn: 0, endColumn: 0 }],
              stopIfTrue: true,
              rule: { type: 'highlightCell', subType: 'number' },
            },
            {
              cfId: 'cf-scale',
              ranges: [{ startRow: 1, endRow: 9, startColumn: 2, endColumn: 3 }],
              rule: { type: 'colorScale' },
            },
          ],
        },
      ],
    }

    expect(listWorkbookConditionalFormats(workbook as never)).toEqual([
      {
        ruleId: 'cf-number',
        sheet: 'Budget',
        ranges: ['A1:A5'],
        kind: 'number',
        stopIfTrue: true,
      },
      {
        ruleId: 'cf-scale',
        sheet: 'Budget',
        ranges: ['C2:D10'],
        kind: 'colorScale',
        stopIfTrue: false,
      },
    ])
  })

  it('rehydrates text equality and error predicates as equivalent native formula rules', () => {
    const formulas: string[] = []
    const builder = {
      whenFormulaSatisfied: (formula: string) => {
        formulas.push(formula)
        return builder
      },
      setRanges: () => builder,
      build: () => ({ built: true }),
    }
    const worksheet = { newConditionalFormattingRule: () => builder }
    const range = [{ startRow: 1, endRow: 3, startColumn: 2, endColumn: 2 }]
    for (const rule of [
      { ruleType: 'cellIs', operator: 'equal', formulas: ['"Ready"'] },
      { ruleType: 'cellIs', operator: 'notEqual', formulas: ['"Draft"'] },
      { ruleType: 'containsErrors', formulas: [] },
      { ruleType: 'notContainsErrors', formulas: [] },
    ]) {
      expect(
        buildConditionalRule(worksheet as never, [], {
          ranges: range,
          priority: 1,
          stopIfTrue: false,
          dxfIndex: undefined,
          colors: [],
          cfvos: [],
          text: undefined,
          rank: undefined,
          bottom: false,
          percent: false,
          ...rule,
        } as never),
      ).toEqual({ built: true })
    }
    expect(formulas).toEqual(['=C2="Ready"', '=C2<>"Draft"', '=ISERROR(C2)', '=NOT(ISERROR(C2))'])
  })
})
