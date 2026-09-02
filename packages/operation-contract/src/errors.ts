export type OperationContractErrorCode =
  | 'catalog_format_mismatch'
  | 'duplicate_operation_alias'
  | 'duplicate_operation_id'
  | 'invalid_operation_id_prefix'
  | 'missing_operation_metadata'
  | 'operation_descriptor_too_large'
  | 'unsupported_operation_format'
  | 'unsupported_schema_keyword'

export type OperationExecutionErrorCode =
  | 'session_not_found'
  | 'editor_offline'
  | 'revision_conflict'
  | 'command_in_flight'
  | 'command_not_found'
  | 'command_timeout'
  | 'operation_not_found'
  | 'operation_unavailable'
  | 'operation_schema_invalid'
  | 'transaction_not_atomic'
  | 'request_reused'
  | 'unsupported_operation'
  | 'invalid_arguments'
  | 'execution_failed'
  | 'internal_error'

export interface OperationExecutionError {
  readonly code: OperationExecutionErrorCode
  readonly message: string
  readonly details?: Readonly<Record<string, unknown>>
}

export class OperationContractError extends Error {
  readonly code: OperationContractErrorCode
  readonly subject: string

  constructor(code: OperationContractErrorCode, subject: string) {
    super(`[${code}] ${subject}`)
    this.name = 'OperationContractError'
    this.code = code
    this.subject = subject
  }
}
