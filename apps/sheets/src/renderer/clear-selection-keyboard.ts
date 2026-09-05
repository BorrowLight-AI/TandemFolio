export const CLEAR_SELECTION_CONTENT_COMMAND = 'sheet.command.clear-selection-content'

const NATIVE_FIELD_SELECTOR = 'input, textarea, select'
const CONTENT_EDITABLE_SELECTOR = '[contenteditable="true"]'
const SHEET_CONTAINER_SELECTOR = '#univer-container'
export const SKIP_HOST_SELECTOR = [
  '[data-u-comp="formula-bar"]',
  '[data-u-comp="input"]',
  '[data-u-comp="textarea"]',
  '[data-u-comp="panel"]',
  '[data-u-comp="panel-field"]',
  '[data-u-comp="cell-popup"]',
  '[data-u-comp="defined-name"]',
  '[data-u-comp="defined-name-container"]',
  '[data-u-comp="select"]',
  '[data-u-comp="multiple-select"]',
  '[data-u-comp="sheets-dropdown-list"]',
  '[data-u-comp="gallery"]',
  '[data-u-comp="slide-tab-item"]',
  '.shape-editable',
  '.chart-editor',
  '.dialog-backdrop',
  '[role="dialog"]',
].join(', ')

export interface ClearSelectionKeyEvent {
  readonly key: string
  readonly metaKey: boolean
  readonly ctrlKey: boolean
  readonly altKey: boolean
  readonly shiftKey: boolean
  readonly defaultPrevented: boolean
  readonly isComposing: boolean
  readonly target: { closest(selector: string): unknown } | EventTarget | null
}

export function isClearSelectionHotkey(
  event: Pick<ClearSelectionKeyEvent, 'key' | 'metaKey' | 'ctrlKey' | 'altKey' | 'shiftKey'>,
): boolean {
  return (
    (event.key === 'Delete' || event.key === 'Backspace') &&
    !event.metaKey &&
    !event.ctrlKey &&
    !event.altKey &&
    !event.shiftKey
  )
}

export function shouldInterceptClearSelection(
  event: ClearSelectionKeyEvent,
  isCellEditing: boolean,
): boolean {
  if (!isClearSelectionHotkey(event) || event.defaultPrevented || event.isComposing || isCellEditing) {
    return false
  }
  const target = event.target
  if (!hasClosest(target)) return true
  if (target.closest(SKIP_HOST_SELECTOR) || target.closest(NATIVE_FIELD_SELECTOR)) return false
  if (target.closest(CONTENT_EDITABLE_SELECTOR) && !target.closest(SHEET_CONTAINER_SELECTOR)) return false
  return true
}

function hasClosest(value: ClearSelectionKeyEvent['target']): value is Element {
  return value !== null && typeof (value as Element).closest === 'function'
}
