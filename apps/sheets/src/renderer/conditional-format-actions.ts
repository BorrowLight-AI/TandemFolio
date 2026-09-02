import {
  CFNumberOperator,
  CFTextOperator,
  CFValueType,
  type IConditionFormattingRule,
  type IValueConfig,
} from '@univerjs/preset-sheets-conditional-formatting'

import { formatAddress } from '../domain/cell-address'
import { OOXML_ICON_SETS, WORST_FIRST_ICON_SETS } from '../gateway/xlsx-cf'
import type { ActiveWorkbook, UniverWorksheet } from './univer-state'

export type WorkbookConditionalFormatComparisonOperator =
  | 'between'
  | 'notBetween'
  | 'equal'
  | 'notEqual'
  | 'greaterThan'
  | 'greaterThanOrEqual'
  | 'lessThan'
  | 'lessThanOrEqual'

export interface WorkbookConditionalFormatStyle {
  readonly fillColor?: string
  readonly fontColor?: string
  readonly bold?: boolean
  readonly italic?: boolean
}

export type WorkbookConditionalFormatHighlightPredicate =
  | 'textContains'
  | 'textNotContains'
  | 'textStartsWith'
  | 'textEndsWith'
  | 'textEqual'
  | 'textNotEqual'
  | 'blank'
  | 'nonBlank'
  | 'error'
  | 'nonError'
  | 'duplicate'
  | 'unique'

export type WorkbookConditionalFormatStatisticalKind = 'rank' | 'average'
export type WorkbookConditionalFormatStatisticalDirection = 'top' | 'bottom' | 'above' | 'below'

export type WorkbookConditionalFormatVisualKind = 'colorScale' | 'dataBar' | 'iconSet'

export interface WorkbookConditionalFormatThreshold {
  readonly type: 'min' | 'max' | 'num' | 'percent' | 'percentile' | 'formula'
  readonly value: number | string | null
  readonly inclusive: boolean | null
}

export interface WorkbookConditionalFormatVisualInput {
  readonly ruleId: string | null
  readonly kind: WorkbookConditionalFormatVisualKind
  readonly colors: readonly string[]
  readonly thresholds: readonly WorkbookConditionalFormatThreshold[]
  readonly iconSet: string | null
  readonly showValue: boolean | null
  readonly reverse: boolean | null
  readonly gradient: boolean | null
  readonly stopIfTrue: boolean
}

interface ComparisonConditionalFormatInput {
  readonly ruleId: string | null
  readonly operator: WorkbookConditionalFormatComparisonOperator
  readonly operand1: number
  readonly operand2?: number
  readonly format: WorkbookConditionalFormatStyle
  readonly stopIfTrue: boolean
}

interface HighlightConditionalFormatInput {
  readonly ruleId: string | null
  readonly predicate: WorkbookConditionalFormatHighlightPredicate
  readonly text: string | null
  readonly format: WorkbookConditionalFormatStyle
  readonly stopIfTrue: boolean
}

interface StatisticalConditionalFormatInput {
  readonly ruleId: string | null
  readonly kind: WorkbookConditionalFormatStatisticalKind
  readonly direction: WorkbookConditionalFormatStatisticalDirection
  readonly rank: number | null
  readonly percent: boolean | null
  readonly inclusive: boolean | null
  readonly format: WorkbookConditionalFormatStyle
  readonly stopIfTrue: boolean
}

type WorkbookRange = ReturnType<UniverWorksheet['getRange']>

function comparisonBuilder(
  worksheet: UniverWorksheet,
  operator: WorkbookConditionalFormatComparisonOperator,
  operand1: number,
  operand2: number | undefined,
) {
  const builder = worksheet.newConditionalFormattingRule()
  switch (operator) {
    case 'between':
      return builder.whenNumberBetween(operand1, operand2 as number)
    case 'notBetween':
      return builder.whenNumberNotBetween(operand1, operand2 as number)
    case 'equal':
      return builder.whenNumberEqualTo(operand1)
    case 'notEqual':
      return builder.whenNumberNotEqualTo(operand1)
    case 'greaterThan':
      return builder.whenNumberGreaterThan(operand1)
    case 'greaterThanOrEqual':
      return builder.whenNumberGreaterThanOrEqualTo(operand1)
    case 'lessThan':
      return builder.whenNumberLessThan(operand1)
    case 'lessThanOrEqual':
      return builder.whenNumberLessThanOrEqualTo(operand1)
  }
}

function withConditionalFormatStyle(
  builder: ReturnType<typeof comparisonBuilder>,
  format: WorkbookConditionalFormatStyle,
) {
  let styled = builder
  if (format.fillColor !== undefined) styled = styled.setBackground(format.fillColor)
  if (format.fontColor !== undefined) styled = styled.setFontColor(format.fontColor)
  if (format.bold === true) styled = styled.setBold(true)
  if (format.italic === true) styled = styled.setItalic(true)
  return styled
}

function highlightBuilder(
  worksheet: UniverWorksheet,
  predicate: WorkbookConditionalFormatHighlightPredicate,
  text: string | null,
) {
  const builder = worksheet.newConditionalFormattingRule()
  switch (predicate) {
    case 'textContains':
      return builder.whenTextContains(text as string)
    case 'textNotContains':
      return builder.whenTextDoesNotContain(text as string)
    case 'textStartsWith':
      return builder.whenTextStartsWith(text as string)
    case 'textEndsWith':
      return builder.whenTextEndsWith(text as string)
    case 'textEqual':
    case 'textNotEqual':
      return builder.whenTextEqualTo(text as string)
    case 'blank':
      return builder.whenCellEmpty()
    case 'nonBlank':
      return builder.whenCellNotEmpty()
    case 'error':
    case 'nonError':
      return builder.whenCellEmpty()
    case 'duplicate':
      return builder.setDuplicateValues()
    case 'unique':
      return builder.setUniqueValues()
  }
}

function commitConditionalFormat(
  worksheet: UniverWorksheet,
  range: WorkbookRange,
  ruleId: string | null,
  stopIfTrue: boolean,
  built: IConditionFormattingRule,
): { readonly ruleId: string; readonly created: boolean } {
  const committedRuleId = ruleId ?? built.cfId
  const rule: IConditionFormattingRule = {
    ...built,
    cfId: committedRuleId,
    stopIfTrue,
  }
  if (ruleId === null) worksheet.addConditionalFormattingRule(rule)
  else worksheet.setConditionalFormattingRule(ruleId, rule)
  range.activate()
  return { ruleId: committedRuleId, created: ruleId === null }
}

export function setWorkbookComparisonConditionalFormat(
  worksheet: UniverWorksheet,
  range: WorkbookRange,
  input: ComparisonConditionalFormatInput,
): { readonly ruleId: string; readonly created: boolean } | null {
  const existing =
    input.ruleId === null
      ? null
      : worksheet
          .getConditionalFormattingRules()
          .find((candidate) => candidate.cfId === input.ruleId)
  if (input.ruleId !== null) {
    const config = existing?.rule as { type?: string; subType?: string } | undefined
    if (config?.type !== 'highlightCell' || config.subType !== 'number') return null
  }

  const built = withConditionalFormatStyle(
    comparisonBuilder(worksheet, input.operator, input.operand1, input.operand2),
    input.format,
  )
    .setRanges([range.getRange()])
    .build()
  return commitConditionalFormat(worksheet, range, input.ruleId, input.stopIfTrue, built)
}

export function setWorkbookHighlightConditionalFormat(
  worksheet: UniverWorksheet,
  range: WorkbookRange,
  input: HighlightConditionalFormatInput,
): { readonly ruleId: string; readonly created: boolean } | null {
  if (input.ruleId !== null) {
    const existing = worksheet
      .getConditionalFormattingRules()
      .find((candidate) => candidate.cfId === input.ruleId)
    const config = existing?.rule as { type?: string; subType?: string } | undefined
    if (
      config?.type !== 'highlightCell' ||
      !['text', 'duplicateValues', 'uniqueValues'].includes(config.subType ?? '')
    ) {
      return null
    }
  }
  const built = withConditionalFormatStyle(
    highlightBuilder(worksheet, input.predicate, input.text),
    input.format,
  )
    .setRanges([range.getRange()])
    .build()
  const normalized = ['textNotEqual', 'error', 'nonError'].includes(input.predicate)
    ? ({
        ...built,
        rule: {
          ...built.rule,
          type: 'highlightCell',
          subType: 'text',
          operator:
            input.predicate === 'textNotEqual'
              ? CFTextOperator.notEqual
              : input.predicate === 'error'
                ? CFTextOperator.containsErrors
                : CFTextOperator.notContainsErrors,
          value: input.text ?? '',
        },
      } as IConditionFormattingRule)
    : built
  return commitConditionalFormat(worksheet, range, input.ruleId, input.stopIfTrue, normalized)
}

export function setWorkbookStatisticalConditionalFormat(
  worksheet: UniverWorksheet,
  range: WorkbookRange,
  input: StatisticalConditionalFormatInput,
): { readonly ruleId: string; readonly created: boolean } | null {
  if (input.ruleId !== null) {
    const existing = worksheet
      .getConditionalFormattingRules()
      .find((candidate) => candidate.cfId === input.ruleId)
    const config = existing?.rule as { type?: string; subType?: string } | undefined
    if (config?.type !== 'highlightCell' || !['rank', 'average'].includes(config.subType ?? '')) {
      return null
    }
  }
  const builder = worksheet.newConditionalFormattingRule()
  const statistical =
    input.kind === 'rank'
      ? builder.setRank({
          isBottom: input.direction === 'bottom',
          isPercent: input.percent as boolean,
          value: input.rank as number,
        })
      : builder.setAverage(
          input.direction === 'above'
            ? input.inclusive
              ? CFNumberOperator.greaterThanOrEqual
              : CFNumberOperator.greaterThan
            : input.inclusive
              ? CFNumberOperator.lessThanOrEqual
              : CFNumberOperator.lessThan,
        )
  const built = withConditionalFormatStyle(statistical, input.format)
    .setRanges([range.getRange()])
    .build()
  return commitConditionalFormat(worksheet, range, input.ruleId, input.stopIfTrue, built)
}

export function setWorkbookFormulaConditionalFormat(
  worksheet: UniverWorksheet,
  range: WorkbookRange,
  input: {
    readonly ruleId: string | null
    readonly formula: string
    readonly format: WorkbookConditionalFormatStyle
    readonly stopIfTrue: boolean
  },
): { readonly ruleId: string; readonly created: boolean } | null {
  if (input.ruleId !== null) {
    const existing = worksheet
      .getConditionalFormattingRules()
      .find((candidate) => candidate.cfId === input.ruleId)
    const config = existing?.rule as { type?: string; subType?: string } | undefined
    if (config?.type !== 'highlightCell' || config.subType !== 'formula') return null
  }
  const built = withConditionalFormatStyle(
    worksheet.newConditionalFormattingRule().whenFormulaSatisfied(input.formula),
    input.format,
  )
    .setRanges([range.getRange()])
    .build()
  return commitConditionalFormat(worksheet, range, input.ruleId, input.stopIfTrue, built)
}

function isValidThresholdValue(threshold: WorkbookConditionalFormatThreshold): boolean {
  if (threshold.type === 'min' || threshold.type === 'max') return threshold.value === null
  if (threshold.type === 'formula') {
    return (
      typeof threshold.value === 'string' &&
      threshold.value.length >= 2 &&
      threshold.value.length <= 8_192 &&
      threshold.value.startsWith('=')
    )
  }
  if (typeof threshold.value !== 'number' || !Number.isFinite(threshold.value)) return false
  return threshold.type === 'num' || (threshold.value >= 0 && threshold.value <= 100)
}

export function isWorkbookVisualConditionalFormatValid(
  input: WorkbookConditionalFormatVisualInput,
): boolean {
  if (
    input.colors.some((color) => !/^#[0-9A-F]{6}$/i.test(color)) ||
    input.thresholds.some((threshold) => !isValidThresholdValue(threshold))
  ) {
    return false
  }
  if (input.kind === 'colorScale') {
    return (
      (input.colors.length === 2 || input.colors.length === 3) &&
      input.thresholds.length === input.colors.length &&
      input.thresholds.every((threshold) => threshold.inclusive === null) &&
      input.iconSet === null &&
      input.showValue === null &&
      input.reverse === null &&
      input.gradient === null
    )
  }
  if (input.kind === 'dataBar') {
    return (
      input.colors.length === 1 &&
      input.thresholds.length === 2 &&
      input.thresholds.every((threshold) => threshold.inclusive === null) &&
      input.iconSet === null &&
      typeof input.showValue === 'boolean' &&
      input.reverse === null &&
      input.gradient === true
    )
  }
  const count = input.iconSet === null ? 0 : Number(input.iconSet[0])
  return (
    input.colors.length === 0 &&
    input.iconSet !== null &&
    OOXML_ICON_SETS.has(input.iconSet) &&
    input.thresholds.length === count &&
    input.thresholds[0]?.type === 'percent' &&
    input.thresholds[0]?.value === 0 &&
    input.thresholds.every((threshold) => typeof threshold.inclusive === 'boolean') &&
    typeof input.showValue === 'boolean' &&
    typeof input.reverse === 'boolean' &&
    input.gradient === null
  )
}

function toConditionalFormatValue(threshold: WorkbookConditionalFormatThreshold): IValueConfig {
  if (threshold.type === 'min') return { type: CFValueType.min }
  if (threshold.type === 'max') return { type: CFValueType.max }
  if (threshold.type === 'formula') {
    return { type: CFValueType.formula, value: threshold.value as string }
  }
  return { type: CFValueType[threshold.type], value: threshold.value as number }
}

export function setWorkbookVisualConditionalFormat(
  worksheet: UniverWorksheet,
  range: WorkbookRange,
  input: WorkbookConditionalFormatVisualInput,
): { readonly ruleId: string; readonly created: boolean } | null {
  if (input.ruleId !== null) {
    const existing = worksheet
      .getConditionalFormattingRules()
      .find((candidate) => candidate.cfId === input.ruleId)
    const config = existing?.rule as { type?: string } | undefined
    if (!config || !['colorScale', 'dataBar', 'iconSet'].includes(config.type ?? '')) return null
  }
  const builder = worksheet.newConditionalFormattingRule()
  const visual =
    input.kind === 'colorScale'
      ? builder.setColorScale(
          input.thresholds.map((threshold, index) => ({
            index,
            color: input.colors[index]!,
            value: toConditionalFormatValue(threshold),
          })),
        )
      : input.kind === 'dataBar'
        ? builder.setDataBar({
            min: toConditionalFormatValue(input.thresholds[0]!),
            max: toConditionalFormatValue(input.thresholds[1]!),
            isGradient: true,
            positiveColor: input.colors[0]!,
            nativeColor: input.colors[0]!,
            isShowValue: input.showValue as boolean,
          })
        : builder.setIconSet({
            iconConfigs: (() => {
              const count = input.thresholds.length
              const upIds = input.thresholds.map((_, index) => String(index))
              const downIds = input.thresholds.map((_, index) => String(count - 1 - index))
              const [naturalIds, reversedIds] = WORST_FIRST_ICON_SETS.has(input.iconSet as string)
                ? [upIds, downIds]
                : [downIds, upIds]
              const ids = input.reverse ? reversedIds : naturalIds
              return input.thresholds
                .map((threshold, index) => ({
                  iconType: input.iconSet as string,
                  iconId: ids[index]!,
                  operator: threshold.inclusive
                    ? CFNumberOperator.greaterThanOrEqual
                    : CFNumberOperator.greaterThan,
                  value: toConditionalFormatValue(threshold),
                }))
                .reverse()
            })(),
            isShowValue: input.showValue as boolean,
          } as Parameters<typeof builder.setIconSet>[0])
  const built = visual.setRanges([range.getRange()]).build()
  return commitConditionalFormat(worksheet, range, input.ruleId, input.stopIfTrue, built)
}

export function removeWorkbookConditionalFormat(
  worksheet: UniverWorksheet,
  ruleId: string,
): boolean {
  if (!worksheet.getConditionalFormattingRules().some((candidate) => candidate.cfId === ruleId)) {
    return false
  }
  worksheet.deleteConditionalFormattingRule(ruleId)
  return true
}

export function clearWorkbookConditionalFormats(
  target: {
    getConditionalFormattingRules(): readonly IConditionFormattingRule[]
    clearConditionalFormatRules(): unknown
  },
): number {
  const cleared = target.getConditionalFormattingRules().length
  target.clearConditionalFormatRules()
  return cleared
}

export function setWorkbookConditionalFormatPriority(
  worksheet: UniverWorksheet,
  ruleId: string,
  position: number,
): { readonly moved: boolean } | null {
  const rules = worksheet.getConditionalFormattingRules()
  const currentIndex = rules.findIndex((candidate) => candidate.cfId === ruleId)
  if (currentIndex < 0 || position < 1 || position > rules.length) return null
  const targetIndex = position - 1
  if (currentIndex === targetIndex) return { moved: false }

  const remaining = rules.filter((candidate) => candidate.cfId !== ruleId)
  const target = remaining[targetIndex]
  if (target) worksheet.moveConditionalFormattingRule(ruleId, target.cfId, 'before')
  else worksheet.moveConditionalFormattingRule(ruleId, remaining.at(-1)!.cfId, 'after')
  return { moved: true }
}

function conditionalFormatRange(range: IConditionFormattingRule['ranges'][number]): string {
  const start = formatAddress(range.startRow, range.startColumn)
  const end = formatAddress(range.endRow, range.endColumn)
  return start === end ? start : `${start}:${end}`
}

export function listWorkbookConditionalFormats(workbook: ActiveWorkbook | null) {
  if (!workbook) return []
  return workbook
    .getSheets()
    .flatMap((worksheet) =>
      worksheet.getConditionalFormattingRules().map((rule) => {
        const config = rule.rule as { type?: string; subType?: string }
        return {
          ruleId: rule.cfId,
          sheet: worksheet.getSheetName(),
          ranges: rule.ranges.map(conditionalFormatRange),
          kind: config.subType ?? config.type ?? 'unknown',
          stopIfTrue: rule.stopIfTrue === true,
        }
      }),
    )
    .slice(0, 100)
}
