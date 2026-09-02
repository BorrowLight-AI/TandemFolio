import { describe, expect, it } from 'vitest'

import config from '../vite.renderer.config'

describe('XLSX product dependency boundary', () => {
  it('replaces Univer optional telemetry with the local inert service identifier', () => {
    const alias = config.resolve?.alias
    expect(alias).toMatchObject({
      '@univerjs/telemetry': expect.stringContaining('univer-optional-metrics.ts'),
    })
  })
})
