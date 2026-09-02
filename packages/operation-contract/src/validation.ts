import type { OperationCatalog } from './descriptor'
import { OperationContractError, type OperationContractErrorCode } from './errors'
import { createOperationManifest } from './manifest'

export interface OperationValidationError {
  readonly code: OperationContractErrorCode
  readonly subject: string
  readonly message: string
}

export type OperationValidationResult =
  { readonly ok: true } | { readonly ok: false; readonly error: OperationValidationError }

export function validateOperationCatalogs(
  catalogs: readonly OperationCatalog[],
): OperationValidationResult {
  try {
    createOperationManifest(catalogs)
    return { ok: true }
  } catch (error) {
    if (!(error instanceof OperationContractError)) throw error
    return {
      ok: false,
      error: {
        code: error.code,
        subject: error.subject,
        message: error.message,
      },
    }
  }
}
