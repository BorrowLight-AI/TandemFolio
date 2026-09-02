import type { OperationExecutionError } from './errors'
import type { OperationId } from './descriptor'

export interface OperationInvocation {
  readonly id: OperationId
  readonly arguments: Readonly<Record<string, unknown>>
}

export interface OperationTransaction {
  readonly sessionId: string
  readonly baseRevision: number
  readonly requestId: string
  readonly operations: readonly OperationInvocation[]
}

export interface OperationTransactionIdentity {
  readonly transactionId: string
  readonly requestId: string
  readonly baseRevision: number
}

export interface OperationExecutionResult {
  readonly id: OperationId
  readonly result: unknown
}

export interface OperationTransactionSuccess {
  readonly ok: true
  readonly transaction: OperationTransactionIdentity
  readonly result: {
    readonly revision: number
    readonly operations: readonly OperationExecutionResult[]
  }
}

export interface OperationTransactionFailure {
  readonly ok: false
  readonly error: OperationExecutionError
}

export type OperationTransactionOutcome = OperationTransactionSuccess | OperationTransactionFailure
