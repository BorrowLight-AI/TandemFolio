import { describe, expect, it } from 'vitest'

import {
  isClearSelectionHotkey,
  shouldInterceptClearSelection,
  type ClearSelectionKeyEvent,
} from '../src/renderer/clear-selection-keyboard'

function event(key: string, hosts: readonly string[], extra: Partial<ClearSelectionKeyEvent> = {}) {
  return {
    key,
    metaKey: false,
    ctrlKey: false,
    altKey: false,
    shiftKey: false,
    defaultPrevented: false,
    isComposing: false,
    target: {
      closest(selector: string) {
        return hosts.some((host) => selector === host || selector.includes(host)) ? {} : null
      },
    },
    ...extra,
  } satisfies ClearSelectionKeyEvent
}

describe('clear-selection keyboard routing', () => {
  it('recognizes only unmodified Delete and Backspace', () => {
    expect(isClearSelectionHotkey(event('Delete', []))).toBe(true)
    expect(isClearSelectionHotkey(event('Backspace', []))).toBe(true)
    expect(isClearSelectionHotkey(event('Delete', [], { ctrlKey: true }))).toBe(false)
  })

  it('intercepts the hidden grid editor but skips active editors and dialogs', () => {
    expect(
      shouldInterceptClearSelection(
        event('Backspace', ['[contenteditable="true"]', '#univer-container']),
        false,
      ),
    ).toBe(true)
    expect(shouldInterceptClearSelection(event('Backspace', []), true)).toBe(false)
    expect(shouldInterceptClearSelection(event('Delete', ['[role="dialog"]']), false)).toBe(false)
    expect(
      shouldInterceptClearSelection(
        event('Backspace', ['[data-u-comp="slide-tab-item"]', '#univer-container']),
        false,
      ),
    ).toBe(false)
  })
})
