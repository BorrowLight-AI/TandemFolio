import {
  OPERATION_MANIFEST_SCHEMA_VERSION,
  type OperationDescriptor,
  type OperationFormat,
  type OperationManifest,
} from '@tandemfolio/operation-contract'

import manifestSource from './generated/operation-manifest.json'

const manifest = manifestSource as unknown as OperationManifest

if (manifest.schemaVersion !== OPERATION_MANIFEST_SCHEMA_VERSION) {
  throw new Error(`Unsupported operation manifest version: ${String(manifest.schemaVersion)}`)
}

function descriptorsByVisibility(
  format: OperationFormat,
  visibility: OperationDescriptor['visibility'],
): readonly OperationDescriptor[] {
  return manifest.operations.filter(
    (descriptor) => descriptor.format === format && descriptor.visibility === visibility,
  )
}

export function getRegisteredOperationDescriptors(
  format: OperationFormat,
): readonly OperationDescriptor[] {
  return descriptorsByVisibility(format, 'agent')
}

export function getCanonicalRegisteredOperation(
  format: OperationFormat,
  operationId: string,
  visibility: OperationDescriptor['visibility'] = 'agent',
): OperationDescriptor | null {
  return (
    descriptorsByVisibility(format, visibility).find(
      (descriptor) => descriptor.id === operationId,
    ) ?? null
  )
}

export function resolveRegisteredOperation(
  format: OperationFormat,
  requestedOperation: string,
  visibility: OperationDescriptor['visibility'] = 'agent',
): OperationDescriptor | null {
  const canonical = getCanonicalRegisteredOperation(format, requestedOperation, visibility)
  if (canonical) return canonical
  for (const descriptor of descriptorsByVisibility(format, visibility)) {
    if (descriptor.compatibilityAliases?.includes(requestedOperation)) {
      return descriptor
    }
  }
  return null
}
