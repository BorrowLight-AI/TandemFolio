import { describe, expect, it } from 'vitest'
import { isDocDirty, type DocDirtyState } from '../src/renderer/doc-dirty'

function cleanState(): DocDirtyState {
  return {
    dirtyRef: { current: false },
    sectionDirty: false,
    sectionsDirty: [],
    trailingStartType: null,
    pageColorDirty: false,
    numberingDirty: false,
    styleUpserts: {},
    titlePgDirty: false,
    evenOddHfDirty: false,
    watermarkDirty: false,
    inksDirty: false,
    notesDirty: false,
    sourcesDirty: false,
    themeFontsDirty: false,
    themeColorsDirty: false,
    commentsDirty: false,
    protectionDirty: false,
  }
}

describe('isDocDirty', () => {
  it('is false for a fully clean state', () => {
    expect(isDocDirty(cleanState())).toBe(false)
  })

  it('is true when only the body changed', () => {
    expect(isDocDirty({ ...cleanState(), dirtyRef: { current: true } })).toBe(true)
  })

  // the regression: autosave ignored comment/protection edits; the recovery push ignored everything but the body
  it.each([
    ['commentsDirty', { commentsDirty: true }],
    ['protectionDirty', { protectionDirty: true }],
    ['sectionDirty', { sectionDirty: true }],
    ['pageColorDirty', { pageColorDirty: true }],
    ['titlePgDirty', { titlePgDirty: true }],
    ['evenOddHfDirty', { evenOddHfDirty: true }],
    ['watermarkDirty', { watermarkDirty: true }],
    ['inksDirty', { inksDirty: true }],
    ['notesDirty', { notesDirty: true }],
    ['sourcesDirty', { sourcesDirty: true }],
    ['themeFontsDirty', { themeFontsDirty: true }],
    ['themeColorsDirty', { themeColorsDirty: true }],
    ['numberingDirty', { numberingDirty: true }],
    ['sectionsDirty', { sectionsDirty: [1] }],
    ['styleUpserts', { styleUpserts: { Heading1: {} } }],
    ['trailingStartType', { trailingStartType: 'nextPage' }],
  ] as const)('is true when only %s changed', (_name, patch) => {
    expect(isDocDirty({ ...cleanState(), ...patch })).toBe(true)
  })
})
