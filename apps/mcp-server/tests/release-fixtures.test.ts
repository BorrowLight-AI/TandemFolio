import { createHash } from 'node:crypto'
import { describe, expect, it } from 'vitest'
import { createReleaseFixture, type ReleaseFormat } from '../../../tools/release-gate/fixtures'

const formats: ReleaseFormat[] = ['docx', 'markdown', 'xlsx', 'pptx', 'pdf']

describe('release benchmark fixtures', () => {
  it.each(formats)(
    '%s fixtures are deterministic and increase in representative complexity',
    async (format) => {
      const first = await Promise.all(
        (['small', 'medium', 'large'] as const).map((size) => createReleaseFixture(format, size)),
      )
      const second = await Promise.all(
        (['small', 'medium', 'large'] as const).map((size) => createReleaseFixture(format, size)),
      )

      expect(first.map(({ bytes }) => bytes.byteLength)).toEqual(
        [...first.map(({ bytes }) => bytes.byteLength)].sort((left, right) => left - right),
      )
      expect(first.map(({ bytes }) => createHash('sha256').update(bytes).digest('hex'))).toEqual(
        second.map(({ bytes }) => createHash('sha256').update(bytes).digest('hex')),
      )
    },
  )
})
