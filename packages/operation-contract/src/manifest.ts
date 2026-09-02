import { OPERATION_FORMATS, type OperationCatalog, type OperationDescriptor } from './descriptor'
import { OperationContractError } from './errors'
import { findUnsupportedSchemaKeyword } from './schema'

export const OPERATION_MANIFEST_SCHEMA_VERSION = 1 as const
export const OPERATION_DESCRIPTOR_MAX_BYTES = 65_536

export interface OperationManifest {
  readonly schemaVersion: typeof OPERATION_MANIFEST_SCHEMA_VERSION
  readonly operations: readonly OperationDescriptor[]
}

const REQUIRED_OPERATION_FIELDS = [
  'id',
  'format',
  'family',
  'summary',
  'visibility',
  'inputSchema',
  'outputSchema',
  'risk',
  'context',
  'effects',
  'mutates',
  'undoable',
  'atomic',
] as const

function sortObjectKeys(value: unknown): unknown {
  if (Array.isArray(value)) return value.map(sortObjectKeys)
  if (typeof value !== 'object' || value === null) return value

  return Object.fromEntries(
    Object.entries(value)
      .sort(([left], [right]) => left.localeCompare(right))
      .map(([key, nestedValue]) => [key, sortObjectKeys(nestedValue)]),
  )
}

export function serializeOperationManifest(manifest: OperationManifest): string {
  return `${JSON.stringify(sortObjectKeys(manifest), null, 2)}\n`
}

export function createOperationManifest(catalogs: readonly OperationCatalog[]): OperationManifest {
  for (const catalog of catalogs) {
    if (!OPERATION_FORMATS.some((format) => format === catalog.format)) {
      throw new OperationContractError('unsupported_operation_format', catalog.format)
    }
    for (const operation of catalog.operations) {
      if (operation.format !== catalog.format) {
        throw new OperationContractError(
          'catalog_format_mismatch',
          `${operation.id}: expected ${catalog.format}, received ${operation.format}`,
        )
      }
    }
  }

  const operations = catalogs.flatMap((catalog) => catalog.operations)
  const operationIds = new Set<string>()
  const operationAliases = new Set<string>()

  for (const operation of operations) {
    const missingField = REQUIRED_OPERATION_FIELDS.find(
      (field) => operation[field] === undefined || operation[field] === null,
    )
    if (missingField) {
      throw new OperationContractError(
        'missing_operation_metadata',
        `${operation.id ?? '<unknown>'}: ${missingField}`,
      )
    }
    const descriptorBytes = new TextEncoder().encode(JSON.stringify(operation)).byteLength
    if (descriptorBytes > OPERATION_DESCRIPTOR_MAX_BYTES) {
      throw new OperationContractError(
        'operation_descriptor_too_large',
        `${operation.id}: ${descriptorBytes} bytes exceeds ${OPERATION_DESCRIPTOR_MAX_BYTES}`,
      )
    }
    for (const schemaField of ['inputSchema', 'outputSchema'] as const) {
      const unsupportedKeyword = findUnsupportedSchemaKeyword(operation[schemaField], schemaField)
      if (unsupportedKeyword) {
        throw new OperationContractError(
          'unsupported_schema_keyword',
          `${operation.id}: ${unsupportedKeyword}`,
        )
      }
    }
    if (!operation.id.startsWith(`${operation.format}.`)) {
      throw new OperationContractError(
        'invalid_operation_id_prefix',
        `${operation.id}: expected ${operation.format}.`,
      )
    }
    if (operationIds.has(operation.id)) {
      throw new OperationContractError('duplicate_operation_id', operation.id)
    }
    operationIds.add(operation.id)

    for (const alias of operation.compatibilityAliases ?? []) {
      const scopedAlias = `${operation.format}\0${alias}`
      if (
        operationAliases.has(scopedAlias) ||
        operations.some(
          (candidate) => candidate.format === operation.format && candidate.id === alias,
        )
      ) {
        throw new OperationContractError('duplicate_operation_alias', alias)
      }
      operationAliases.add(scopedAlias)
    }
  }

  return {
    schemaVersion: OPERATION_MANIFEST_SCHEMA_VERSION,
    operations: [...operations].sort((left, right) => left.id.localeCompare(right.id)),
  }
}
