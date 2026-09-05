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

  it('rehydrates text equality and error predicates with equivalent native rules', () => {
    const formulas: string[] = []
    const builtRules: Record<string, unknown>[] = []
    let rule: Record<string, unknown> = {}
    const builder = {
      whenFormulaSatisfied: (formula: string) => {
        formulas.push(formula)
        rule = { subType: 'formula', value: formula }
        return builder
      },
      whenTextEqualTo: (value: string) => {
        rule = { subType: 'text', operator: 'equal', value }
        return builder
      },
      whenTextContains: (value: string) => {
        rule = { subType: 'text', operator: 'containsText', value }
        return builder
      },
      setRanges: () => builder,
      build: () => {
        builtRules.push(rule)
        return { built: true, rule }
      },
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
      ).toMatchObject({ built: true })
    }
    expect(formulas).toEqual(['=C2<>"Draft"'])
    expect(builtRules).toEqual([
      { subType: 'text', operator: 'equal', value: 'Ready' },
      { subType: 'formula', value: '=C2<>"Draft"' },
      { subType: 'text', operator: 'containsErrors' },
      { subType: 'text', operator: 'notContainsErrors' },
    ])
  })
})
