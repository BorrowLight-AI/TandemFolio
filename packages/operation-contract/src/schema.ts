import type { JsonSchema } from './descriptor'

const SUPPORTED_JSON_SCHEMA_KEYWORDS = new Set([
  'additionalProperties',
  'description',
  'enum',
  'items',
  'maxItems',
  'maxLength',
  'maximum',
  'minItems',
  'minLength',
  'minimum',
  'pattern',
  'properties',
  'required',
  'title',
  'type',
])

function isRecord(value: unknown): value is Readonly<Record<string, unknown>> {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}

export function findUnsupportedSchemaKeyword(schema: JsonSchema, path: string): string | null {
  for (const [keyword, value] of Object.entries(schema)) {
    const keywordPath = `${path}.${keyword}`
    if (!SUPPORTED_JSON_SCHEMA_KEYWORDS.has(keyword)) {
      return keywordPath
    }

    if (keyword === 'properties' && isRecord(value)) {
      for (const [name, nestedSchema] of Object.entries(value)) {
        if (!isRecord(nestedSchema)) continue
        const unsupported = findUnsupportedSchemaKeyword(nestedSchema, `${keywordPath}.${name}`)
        if (unsupported) return unsupported
      }
    }

    if (keyword === 'additionalProperties' && isRecord(value)) {
      const unsupported = findUnsupportedSchemaKeyword(value, keywordPath)
      if (unsupported) return unsupported
    }

    if (keyword === 'items' && isRecord(value)) {
      const unsupported = findUnsupportedSchemaKeyword(value, keywordPath)
      if (unsupported) return unsupported
    }
  }

  return null
}

export interface JsonSchemaValueValidationError {
  readonly path: string
  readonly message: string
}

export type JsonSchemaValueValidationResult =
  { readonly ok: true } | { readonly ok: false; readonly error: JsonSchemaValueValidationError }

function matchesType(type: unknown, value: unknown): boolean {
  if (Array.isArray(type)) return type.some((candidate) => matchesType(candidate, value))
  switch (type) {
    case 'object':
      return isRecord(value)
    case 'array':
      return Array.isArray(value)
    case 'string':
      return typeof value === 'string'
    case 'number':
      return typeof value === 'number' && Number.isFinite(value)
    case 'integer':
      return typeof value === 'number' && Number.isInteger(value)
    case 'boolean':
      return typeof value === 'boolean'
    case 'null':
      return value === null
    default:
      return false
  }
}

function includesType(type: unknown, candidate: string): boolean {
  return Array.isArray(type) ? type.includes(candidate) : type === candidate
}

function typeLabel(type: unknown): string {
  if (Array.isArray(type)) {
    return type.map((candidate) => typeLabel(candidate)).join(' or ')
  }
  return typeof type === 'string' ? type : 'supported JSON value'
}

function typeExpectation(type: unknown): string {
  const label = typeLabel(type)
  if (Array.isArray(type)) return label
  return `${label === 'integer' ? 'an' : 'a'} ${label}`
}

export function validateJsonSchemaValue(
  schema: JsonSchema,
  value: unknown,
  path = '$',
): JsonSchemaValueValidationResult {
  if (schema.type !== undefined && !matchesType(schema.type, value)) {
    return {
      ok: false,
      error: { path, message: `${path} must be ${typeExpectation(schema.type)}.` },
    }
  }

  if (Array.isArray(schema.enum) && !schema.enum.some((candidate) => Object.is(candidate, value))) {
    return {
      ok: false,
      error: {
        path,
        message: `${path} must be one of ${schema.enum.map((candidate) => JSON.stringify(candidate)).join(', ')}.`,
      },
    }
  }

  if (typeof value === 'number' && typeof schema.minimum === 'number' && value < schema.minimum) {
    return {
      ok: false,
      error: {
        path,
        message: `${path} must be greater than or equal to ${schema.minimum}.`,
      },
    }
  }
  if (typeof value === 'number' && typeof schema.maximum === 'number' && value > schema.maximum) {
    return {
      ok: false,
      error: {
        path,
        message: `${path} must be less than or equal to ${schema.maximum}.`,
      },
    }
  }

  if (typeof value === 'string') {
    const length = Array.from(value).length
    if (typeof schema.minLength === 'number' && length < schema.minLength) {
      const suffix = schema.minLength === 1 ? 'character' : 'characters'
      return {
        ok: false,
        error: {
          path,
          message: `${path} must contain at least ${schema.minLength} ${suffix}.`,
        },
      }
    }
    if (typeof schema.maxLength === 'number' && length > schema.maxLength) {
      const suffix = schema.maxLength === 1 ? 'character' : 'characters'
      return {
        ok: false,
        error: {
          path,
          message: `${path} must contain at most ${schema.maxLength} ${suffix}.`,
        },
      }
    }
    if (typeof schema.pattern === 'string' && !new RegExp(schema.pattern).test(value)) {
      return {
        ok: false,
        error: {
          path,
          message: `${path} must match pattern ${schema.pattern}.`,
        },
      }
    }
  }

  if (includesType(schema.type, 'array') && Array.isArray(value)) {
    if (typeof schema.minItems === 'number' && value.length < schema.minItems) {
      const suffix = schema.minItems === 1 ? 'item' : 'items'
      return {
        ok: false,
        error: {
          path,
          message: `${path} must contain at least ${schema.minItems} ${suffix}.`,
        },
      }
    }
    if (typeof schema.maxItems === 'number' && value.length > schema.maxItems) {
      const suffix = schema.maxItems === 1 ? 'item' : 'items'
      return {
        ok: false,
        error: {
          path,
          message: `${path} must contain at most ${schema.maxItems} ${suffix}.`,
        },
      }
    }
    if (!isRecord(schema.items)) return { ok: true }
    for (let index = 0; index < value.length; index += 1) {
      const result = validateJsonSchemaValue(schema.items, value[index], `${path}[${index}]`)
      if (!result.ok) return result
    }
  }

  if (!includesType(schema.type, 'object') || !isRecord(value)) return { ok: true }

  const properties = isRecord(schema.properties) ? schema.properties : {}
  const required = Array.isArray(schema.required) ? schema.required : []
  for (const property of required) {
    if (typeof property === 'string' && !(property in value)) {
      const propertyPath = `${path}.${property}`
      return {
        ok: false,
        error: { path: propertyPath, message: `${propertyPath} is required.` },
      }
    }
  }

  for (const [property, propertyValue] of Object.entries(value)) {
    const propertyPath = `${path}.${property}`
    const propertySchema = properties[property]
    if (!propertySchema) {
      if (schema.additionalProperties === false) {
        return {
          ok: false,
          error: { path: propertyPath, message: `${propertyPath} is not allowed.` },
        }
      }
      continue
    }
    if (!isRecord(propertySchema)) continue
    const result = validateJsonSchemaValue(propertySchema, propertyValue, propertyPath)
    if (!result.ok) return result
  }

  return { ok: true }
}
