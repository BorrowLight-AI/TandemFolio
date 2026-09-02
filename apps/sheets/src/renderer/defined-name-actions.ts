import type { UniverRuntime } from './univer-state'
import { univerDefinedNames } from './univer-sync'

export type DefinedNameAction =
  | { kind: 'add'; name: string; ref: string; sheetId: string | null }
  | {
      kind: 'update'
      originalName: string
      scopeSheetId: string | null
      name: string
      ref: string
    }
  | { kind: 'remove'; name: string; scopeSheetId: string | null }

export interface DefinedNameSetInput {
  readonly name: string
  readonly formula: string
  readonly scopeSheetId: string | null
  readonly previousName?: string
}

const WORKBOOK_SCOPE = 'AllDefaultWorkbook'
const NAME_PATTERN = /^[A-Za-z_\\][A-Za-z0-9_.\\]*$/
const CELL_REF_PATTERN = /^(?:[A-Za-z]{1,3}[0-9]+|[Rr][0-9]*[Cc][0-9]*)$/

export function isValidDefinedName(name: string): boolean {
  const lower = name.toLowerCase()
  return (
    name.length > 0 &&
    name.length <= 255 &&
    NAME_PATTERN.test(name) &&
    !CELL_REF_PATTERN.test(name) &&
    lower !== 'true' &&
    lower !== 'false' &&
    !lower.startsWith('_xlnm')
  )
}

function normalizedScope(scopeSheetId: string | null | undefined): string {
  return scopeSheetId ?? WORKBOOK_SCOPE
}

function sameName(left: string, right: string): boolean {
  return left.toLowerCase() === right.toLowerCase()
}

function findDefinedName(runtime: UniverRuntime, name: string, scopeSheetId: string | null) {
  const scope = normalizedScope(scopeSheetId)
  return univerDefinedNames(runtime).find(
    (defined) =>
      sameName(defined.getName(), name) && normalizedScope(defined.getLocalSheetId()) === scope,
  )
}

function insertDefinedName(
  runtime: UniverRuntime,
  input: { readonly name: string; readonly formula: string; readonly scopeSheetId: string | null },
): void {
  const workbook = runtime.univerAPI.getActiveWorkbook()
  if (!workbook) throw new Error('Open an XLSX workbook first.')
  const wb = workbook as unknown as {
    newDefinedNameBuilder(): {
      load(param: Record<string, unknown>): { build(): unknown }
    }
    insertDefinedNameBuilder(param: unknown): void
  }
  wb.insertDefinedNameBuilder(
    wb
      .newDefinedNameBuilder()
      .load({
        name: input.name,
        formulaOrRefString: input.formula.replace(/^=/, ''),
        localSheetId: normalizedScope(input.scopeSheetId),
      })
      .build(),
  )
}

export function setWorkbookDefinedName(runtime: UniverRuntime, input: DefinedNameSetInput): void {
  const targetName = input.previousName ?? input.name
  const target = findDefinedName(runtime, targetName, input.scopeSheetId)
  if (!target) {
    if (input.previousName !== undefined) {
      throw new Error(`Unknown XLSX defined name: ${input.previousName}`)
    }
    insertDefinedName(runtime, input)
    return
  }
  if (target.getName() !== input.name) target.setName(input.name)
  target.setRef(input.formula.replace(/^=/, ''))
}

export function removeWorkbookDefinedName(
  runtime: UniverRuntime,
  name: string,
  scopeSheetId: string | null,
): void {
  const target = findDefinedName(runtime, name, scopeSheetId)
  if (!target) throw new Error(`Unknown XLSX defined name: ${name}`)
  target.delete()
}

export function applyWorkbookDefinedNameAction(
  runtime: UniverRuntime,
  action: DefinedNameAction,
): void {
  if (action.kind === 'add') {
    if (findDefinedName(runtime, action.name, action.sheetId)) {
      throw new Error(`The defined name "${action.name}" already exists in this scope.`)
    }
    insertDefinedName(runtime, {
      name: action.name,
      formula: action.ref,
      scopeSheetId: action.sheetId,
    })
    return
  }
  if (action.kind === 'remove') {
    removeWorkbookDefinedName(runtime, action.name, action.scopeSheetId)
    return
  }
  setWorkbookDefinedName(runtime, {
    name: action.name,
    formula: action.ref,
    scopeSheetId: action.scopeSheetId,
    previousName: action.originalName,
  })
}
