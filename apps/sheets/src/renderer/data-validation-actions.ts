import type { UniverRuntime, UniverWorksheet } from './univer-state'

/** Applies one explicit checkbox-validation state through Univer's native DV command and history. */
export function applyWorkbookCheckbox(
  runtime: UniverRuntime,
  range: ReturnType<UniverWorksheet['getRange']>,
  enabled: boolean,
): void {
  range.setDataValidation(
    enabled ? runtime.univerAPI.newDataValidation().requireCheckbox().build() : null,
  )
  range.activate()
}

/** Applies an explicit inline-list rule through the same native DV command/history route. */
export function applyWorkbookListValidation(
  runtime: UniverRuntime,
  range: ReturnType<UniverWorksheet['getRange']>,
  input: {
    readonly values: readonly string[]
    readonly allowBlank: boolean
    readonly showDropdown: boolean
  },
): void {
  const rule = runtime.univerAPI
    .newDataValidation()
    .requireValueInList([...input.values], false, input.showDropdown)
    .setAllowBlank(input.allowBlank)
    .build()
  range.setDataValidation(rule)
  range.activate()
}

/** Applies a single-axis range-backed list through native DV history. */
export function applyWorkbookListReferenceValidation(
  runtime: UniverRuntime,
  target: ReturnType<UniverWorksheet['getRange']>,
  source: ReturnType<UniverWorksheet['getRange']>,
  input: {
    readonly allowBlank: boolean
    readonly showDropdown: boolean
  },
): void {
  const rule = runtime.univerAPI
    .newDataValidation()
    .requireValueInRange(source, false, input.showDropdown)
    .setAllowBlank(input.allowBlank)
    .build()
  target.setDataValidation(rule)
  target.activate()
}

/** Removes any rule from the explicit range through Univer's native DV history. */
export function removeWorkbookDataValidation(range: ReturnType<UniverWorksheet['getRange']>): void {
  range.setDataValidation(null)
  range.activate()
}

function applyBuiltValidation(
  range: ReturnType<UniverWorksheet['getRange']>,
  rule: Parameters<ReturnType<UniverWorksheet['getRange']>['setDataValidation']>[0],
): void {
  range.setDataValidation(rule)
  range.activate()
}

export function applyWorkbookCustomFormulaValidation(
  runtime: UniverRuntime,
  range: ReturnType<UniverWorksheet['getRange']>,
  formula: string,
  allowBlank: boolean,
): void {
  applyBuiltValidation(
    range,
    runtime.univerAPI
      .newDataValidation()
      .requireFormulaSatisfied(formula)
      .setAllowBlank(allowBlank)
      .build(),
  )
}

export type WorkbookComparisonValidationKind = 'whole' | 'decimal' | 'date' | 'time' | 'textLength'

export type WorkbookComparisonValidationOperator =
  | 'between'
  | 'notBetween'
  | 'equal'
  | 'notEqual'
  | 'greaterThan'
  | 'greaterThanOrEqual'
  | 'lessThan'
  | 'lessThanOrEqual'

let comparisonRuleSequence = 0

/**
 * Applies the exact wire model used by Univer's validation panel. FRange's
 * native facade reads only `.rule`, then dispatches the same add-rule command
 * and Undo entry used by builder-created validations.
 */
export function applyWorkbookComparisonValidation(
  range: ReturnType<UniverWorksheet['getRange']>,
  input: {
    readonly kind: WorkbookComparisonValidationKind
    readonly operator: WorkbookComparisonValidationOperator
    readonly formula1: string
    readonly formula2?: string
    readonly allowBlank: boolean
  },
): void {
  const validation = {
    rule: {
      uid: `registry-comparison-dv-${++comparisonRuleSequence}`,
      type: input.kind,
      operator: input.operator,
      formula1: input.formula1,
      ...(input.formula2 === undefined ? {} : { formula2: input.formula2 }),
      allowBlank: input.allowBlank,
    },
  }
  range.setDataValidation(
    validation as Parameters<ReturnType<UniverWorksheet['getRange']>['setDataValidation']>[0],
  )
  range.activate()
}

export interface WorkbookValidationMessages {
  readonly inputTitle: string | null
  readonly inputMessage: string | null
  readonly errorStyle: 'none' | 'stop' | 'warning' | 'information'
  readonly errorTitle: string | null
  readonly errorMessage: string | null
}

const VALIDATION_ERROR_STYLE = { stop: 1, warning: 2, information: 0 } as const

/** Updates an existing rule through FDataValidation's native update-options/Undo command. */
export function applyWorkbookValidationMessages(
  range: ReturnType<UniverWorksheet['getRange']>,
  messages: WorkbookValidationMessages,
): boolean {
  const validation = (
    range as unknown as {
      getDataValidation(): {
        setOptions(options: Record<string, unknown>): unknown
      } | null
    }
  ).getDataValidation()
  if (!validation) return false
  validation.setOptions({
    showInputMessage: messages.inputTitle !== null || messages.inputMessage !== null,
    promptTitle: messages.inputTitle ?? undefined,
    prompt: messages.inputMessage ?? undefined,
    showErrorMessage: messages.errorStyle !== 'none',
    errorStyle:
      messages.errorStyle === 'none' ? undefined : VALIDATION_ERROR_STYLE[messages.errorStyle],
    errorTitle: messages.errorTitle ?? undefined,
    error: messages.errorMessage ?? undefined,
  })
  range.activate()
  return true
}
