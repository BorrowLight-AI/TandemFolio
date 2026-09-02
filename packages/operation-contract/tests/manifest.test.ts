import { describe, expect, it } from 'vitest'

import {
  OPERATION_MANIFEST_SCHEMA_VERSION,
  createOperationManifest,
  serializeOperationManifest,
  validateJsonSchemaValue,
  validateOperationCatalogs,
  type OperationCatalog,
  type OperationDescriptor,
} from '../src/index'

const objectSchema = {
  type: 'object',
  properties: {},
  required: [],
  additionalProperties: false,
} as const

function descriptor(
  id: OperationDescriptor['id'],
  format: OperationDescriptor['format'],
): OperationDescriptor {
  return {
    id,
    format,
    family: 'fixture',
    summary: `Fixture operation ${id}`,
    visibility: 'agent',
    inputSchema: objectSchema,
    outputSchema: objectSchema,
    risk: 'low',
    context: ['document'],
    effects: ['document'],
    mutates: true,
    undoable: true,
    atomic: true,
  }
}

describe('createOperationManifest', () => {
  it('produces one versioned manifest in stable operation-id order', () => {
    const catalogs: OperationCatalog[] = [
      {
        format: 'xlsx',
        operations: [descriptor('xlsx.fixture.second', 'xlsx')],
      },
      {
        format: 'markdown',
        operations: [descriptor('markdown.fixture.first', 'markdown')],
      },
    ]

    expect(createOperationManifest(catalogs)).toEqual({
      schemaVersion: OPERATION_MANIFEST_SCHEMA_VERSION,
      operations: [
        descriptor('markdown.fixture.first', 'markdown'),
        descriptor('xlsx.fixture.second', 'xlsx'),
      ],
    })
  })

  it('rejects duplicate operation ids across format catalogs', () => {
    const repeated = descriptor('markdown.fixture.repeated', 'markdown')

    expect(() =>
      createOperationManifest([
        { format: 'markdown', operations: [repeated] },
        { format: 'markdown', operations: [repeated] },
      ]),
    ).toThrowError('[duplicate_operation_id] markdown.fixture.repeated')
  })

  it('rejects an operation id whose prefix does not match its format', () => {
    expect(() =>
      createOperationManifest([
        {
          format: 'markdown',
          operations: [descriptor('xlsx.fixture.wrong-prefix', 'markdown')],
        },
      ]),
    ).toThrowError('[invalid_operation_id_prefix] xlsx.fixture.wrong-prefix: expected markdown.')
  })

  it('rejects an operation placed in another format catalog', () => {
    expect(() =>
      createOperationManifest([
        {
          format: 'xlsx',
          operations: [descriptor('markdown.fixture.wrong-catalog', 'markdown')],
        },
      ]),
    ).toThrowError(
      '[catalog_format_mismatch] markdown.fixture.wrong-catalog: expected xlsx, received markdown',
    )
  })

  it('rejects catalogs outside the five supported renderer formats', () => {
    const unsupportedCatalog = {
      format: 'word',
      operations: [
        {
          ...descriptor('markdown.fixture.unsupported-format', 'markdown'),
          id: 'word.fixture.unsupported-format',
          format: 'word',
        },
      ],
    } as unknown as OperationCatalog

    expect(() => createOperationManifest([unsupportedCatalog])).toThrowError(
      '[unsupported_operation_format] word',
    )
  })

  it('rejects a compatibility alias claimed by more than one operation', () => {
    expect(() =>
      createOperationManifest([
        {
          format: 'markdown',
          operations: [
            {
              ...descriptor('markdown.fixture.first', 'markdown'),
              compatibilityAliases: ['legacy_text'],
            },
            {
              ...descriptor('markdown.fixture.second', 'markdown'),
              compatibilityAliases: ['legacy_text'],
            },
          ],
        },
      ]),
    ).toThrowError('[duplicate_operation_alias] legacy_text')
  })

  it('rejects a compatibility alias that shadows an operation id', () => {
    expect(() =>
      createOperationManifest([
        {
          format: 'markdown',
          operations: [
            descriptor('markdown.fixture.first', 'markdown'),
            {
              ...descriptor('markdown.fixture.second', 'markdown'),
              compatibilityAliases: ['markdown.fixture.first'],
            },
          ],
        },
      ]),
    ).toThrowError('[duplicate_operation_alias] markdown.fixture.first')
  })

  it('rejects an operation with missing required metadata', () => {
    const { summary: _summary, ...incomplete } = descriptor(
      'markdown.fixture.incomplete',
      'markdown',
    )

    expect(() =>
      createOperationManifest([
        {
          format: 'markdown',
          operations: [incomplete as OperationDescriptor],
        },
      ]),
    ).toThrowError('[missing_operation_metadata] markdown.fixture.incomplete: summary')
  })

  it('rejects an operation descriptor larger than the bounded detail response budget', () => {
    const oversized = {
      ...descriptor('markdown.fixture.oversized-detail', 'markdown'),
      summary: 'x'.repeat(70_000),
    }

    expect(() =>
      createOperationManifest([{ format: 'markdown', operations: [oversized] }]),
    ).toThrowError(/\[operation_descriptor_too_large\] markdown\.fixture\.oversized-detail:/)
  })

  it('rejects JSON Schema keywords outside the supported contract subset', () => {
    expect(() =>
      createOperationManifest([
        {
          format: 'markdown',
          operations: [
            {
              ...descriptor('markdown.fixture.unsupported-schema', 'markdown'),
              inputSchema: {
                ...objectSchema,
                unsupportedKeyword: true,
              },
            },
          ],
        },
      ]),
    ).toThrowError(
      '[unsupported_schema_keyword] markdown.fixture.unsupported-schema: inputSchema.unsupportedKeyword',
    )
  })

  it('accepts and enforces enum constraints through the public schema interface', () => {
    const inputSchema = {
      type: 'object',
      properties: {
        scope: { type: 'string', enum: ['selection', 'document'] },
      },
      required: ['scope'],
      additionalProperties: false,
    } as const
    const constrained = {
      ...descriptor('docx.fixture.enum', 'docx'),
      inputSchema,
    }

    expect(() =>
      createOperationManifest([{ format: 'docx', operations: [constrained] }]),
    ).not.toThrow()
    expect(validateJsonSchemaValue(inputSchema, { scope: 'page' })).toEqual({
      ok: false,
      error: {
        path: '$.scope',
        message: '$.scope must be one of "selection", "document".',
      },
    })
  })

  it('validates every array item through the public schema interface', () => {
    const inputSchema = {
      type: 'object',
      properties: {
        blockIndexes: { type: 'array', items: { type: 'integer' } },
      },
      required: ['blockIndexes'],
      additionalProperties: false,
    } as const
    const constrained = {
      ...descriptor('docx.fixture.items', 'docx'),
      inputSchema,
    }

    expect(() =>
      createOperationManifest([{ format: 'docx', operations: [constrained] }]),
    ).not.toThrow()
    expect(validateJsonSchemaValue(inputSchema, { blockIndexes: [0, 1.5] })).toEqual({
      ok: false,
      error: {
        path: '$.blockIndexes[1]',
        message: '$.blockIndexes[1] must be an integer.',
      },
    })
  })

  it('enforces inclusive numeric bounds through the public schema interface', () => {
    const inputSchema = {
      type: 'object',
      properties: {
        level: { type: 'integer', minimum: 0, maximum: 6 },
      },
      required: ['level'],
      additionalProperties: false,
    } as const
    const constrained = {
      ...descriptor('docx.fixture.numeric-bounds', 'docx'),
      inputSchema,
    }

    expect(() =>
      createOperationManifest([{ format: 'docx', operations: [constrained] }]),
    ).not.toThrow()
    expect(validateJsonSchemaValue(inputSchema, { level: -1 })).toEqual({
      ok: false,
      error: {
        path: '$.level',
        message: '$.level must be greater than or equal to 0.',
      },
    })
    expect(validateJsonSchemaValue(inputSchema, { level: 7 })).toEqual({
      ok: false,
      error: {
        path: '$.level',
        message: '$.level must be less than or equal to 6.',
      },
    })
  })

  it('validates nullable object members through the public schema interface', () => {
    const inputSchema = {
      type: 'object',
      properties: {
        link: {
          type: ['object', 'null'],
          properties: { url: { type: 'string' } },
          required: ['url'],
          additionalProperties: false,
        },
      },
      required: ['link'],
      additionalProperties: false,
    } as const
    const constrained = {
      ...descriptor('docx.fixture.nullable-object', 'docx'),
      inputSchema,
    }

    expect(() =>
      createOperationManifest([{ format: 'docx', operations: [constrained] }]),
    ).not.toThrow()
    expect(validateJsonSchemaValue(inputSchema, { link: null })).toEqual({ ok: true })
    expect(validateJsonSchemaValue(inputSchema, { link: { url: 42 } })).toEqual({
      ok: false,
      error: {
        path: '$.link.url',
        message: '$.link.url must be a string.',
      },
    })
  })

  it('enforces minimum array length through the public schema interface', () => {
    const inputSchema = {
      type: 'object',
      properties: {
        blockIndexes: { type: 'array', minItems: 1, items: { type: 'integer' } },
      },
      required: ['blockIndexes'],
      additionalProperties: false,
    } as const
    const constrained = {
      ...descriptor('docx.fixture.min-items', 'docx'),
      inputSchema,
    }

    expect(() =>
      createOperationManifest([{ format: 'docx', operations: [constrained] }]),
    ).not.toThrow()
    expect(validateJsonSchemaValue(inputSchema, { blockIndexes: [] })).toEqual({
      ok: false,
      error: {
        path: '$.blockIndexes',
        message: '$.blockIndexes must contain at least 1 item.',
      },
    })
  })

  it('enforces maximum array length through the public schema interface', () => {
    const inputSchema = {
      type: 'object',
      properties: {
        conditions: { type: 'array', maxItems: 2, items: { type: 'string' } },
      },
      required: ['conditions'],
      additionalProperties: false,
    } as const
    const constrained = {
      ...descriptor('xlsx.fixture.max-items', 'xlsx'),
      inputSchema,
    }

    expect(() =>
      createOperationManifest([{ format: 'xlsx', operations: [constrained] }]),
    ).not.toThrow()
    expect(validateJsonSchemaValue(inputSchema, { conditions: ['a', 'b', 'c'] })).toEqual({
      ok: false,
      error: {
        path: '$.conditions',
        message: '$.conditions must contain at most 2 items.',
      },
    })
  })

  it('enforces inclusive string-length bounds through the public schema interface', () => {
    const inputSchema = {
      type: 'object',
      properties: {
        latex: { type: 'string', minLength: 1, maxLength: 5 },
      },
      required: ['latex'],
      additionalProperties: false,
    } as const
    const constrained = {
      ...descriptor('docx.fixture.string-bounds', 'docx'),
      inputSchema,
    }

    expect(() =>
      createOperationManifest([{ format: 'docx', operations: [constrained] }]),
    ).not.toThrow()
    expect(validateJsonSchemaValue(inputSchema, { latex: '' })).toEqual({
      ok: false,
      error: {
        path: '$.latex',
        message: '$.latex must contain at least 1 character.',
      },
    })
    expect(validateJsonSchemaValue(inputSchema, { latex: 'abcdef' })).toEqual({
      ok: false,
      error: {
        path: '$.latex',
        message: '$.latex must contain at most 5 characters.',
      },
    })
  })

  it('enforces string patterns through the public schema interface', () => {
    const inputSchema = {
      type: 'object',
      properties: {
        color: { type: 'string', pattern: '^[0-9A-F]{6}$' },
      },
      required: ['color'],
      additionalProperties: false,
    } as const
    const constrained = {
      ...descriptor('docx.fixture.string-pattern', 'docx'),
      inputSchema,
    }

    expect(() =>
      createOperationManifest([{ format: 'docx', operations: [constrained] }]),
    ).not.toThrow()
    expect(validateJsonSchemaValue(inputSchema, { color: 'ZZZZZZ' })).toEqual({
      ok: false,
      error: {
        path: '$.color',
        message: '$.color must match pattern ^[0-9A-F]{6}$.',
      },
    })
    expect(validateJsonSchemaValue(inputSchema, { color: 'A1B2C3' })).toEqual({ ok: true })
  })

  it('reports validation failures through a structured public result', () => {
    const repeated = descriptor('markdown.fixture.repeated', 'markdown')

    expect(
      validateOperationCatalogs([{ format: 'markdown', operations: [repeated, repeated] }]),
    ).toEqual({
      ok: false,
      error: {
        code: 'duplicate_operation_id',
        subject: 'markdown.fixture.repeated',
        message: '[duplicate_operation_id] markdown.fixture.repeated',
      },
    })
  })

  it('serializes equivalent manifests to identical bytes', () => {
    const first = descriptor('markdown.fixture.stable', 'markdown')
    const second: OperationDescriptor = {
      atomic: first.atomic,
      undoable: first.undoable,
      mutates: first.mutates,
      effects: first.effects,
      context: first.context,
      risk: first.risk,
      outputSchema: first.outputSchema,
      inputSchema: first.inputSchema,
      visibility: first.visibility,
      summary: first.summary,
      family: first.family,
      format: first.format,
      id: first.id,
    }

    expect(
      serializeOperationManifest(
        createOperationManifest([{ format: 'markdown', operations: [first] }]),
      ),
    ).toBe(
      serializeOperationManifest(
        createOperationManifest([{ format: 'markdown', operations: [second] }]),
      ),
    )
  })
})
