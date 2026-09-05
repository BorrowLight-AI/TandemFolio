import { describe, expect, it } from 'vitest'

import {
  dataValidationInputMessageHeight,
  getDataValidationInputMessage,
  shouldShowDataValidationDropdown,
} from '../src/renderer/data-validation-dropdown'

const validation = (type: string, rule: Record<string, unknown> = {}) => ({
  rule,
  getCriteriaType: () => type,
})

describe('active-cell data validation chrome', () => {
  it('shows a list dropdown unless the file suppresses it', () => {
    expect(shouldShowDataValidationDropdown(validation('list'))).toBe(true)
    expect(shouldShowDataValidationDropdown(validation('list', { showDropDown: false }))).toBe(
      false,
    )
    expect(shouldShowDataValidationDropdown(validation('decimal'))).toBe(false)
  })

  it('exposes enabled input messages and sizes multiline/CJK prompts', () => {
    const message = getDataValidationInputMessage(
      validation('list', {
        showInputMessage: true,
        promptTitle: '提示',
        prompt: '甲乙丙丁\nSecond',
      }),
    )
    expect(message).toEqual({ title: '提示', prompt: '甲乙丙丁\nSecond' })
    expect(dataValidationInputMessageHeight(message!)).toBeGreaterThan(16)
    expect(getDataValidationInputMessage(validation('list', { prompt: 'hidden' }))).toBeNull()
  })
})
