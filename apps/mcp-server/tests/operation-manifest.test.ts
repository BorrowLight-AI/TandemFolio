import { describe, expect, it } from 'vitest'

import manifest from '../src/generated/operation-manifest.json'

describe('product operation manifest', () => {
  it('contains no compatibility aliases after staged-file transport becomes canonical', () => {
    const aliasedOperations = manifest.operations
      .filter((operation) => operation.compatibilityAliases.length > 0)
      .map((operation) => ({
        id: operation.id,
        visibility: operation.visibility,
        aliases: operation.compatibilityAliases,
      }))

    expect(aliasedOperations).toEqual([])
  })
})
