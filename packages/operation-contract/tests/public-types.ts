import type {
  OperationExecutionError,
  OperationTransaction,
  OperationTransactionOutcome,
} from '../src/index'

const transaction: OperationTransaction = {
  sessionId: 'session-1',
  baseRevision: 4,
  requestId: 'request-1',
  operations: [
    {
      id: 'markdown.fixture.insert_text',
      arguments: { text: 'hello' },
    },
  ],
}

const failure: OperationExecutionError = {
  code: 'operation_not_found',
  message: 'Unknown operation',
}

const outcome: OperationTransactionOutcome = {
  ok: false,
  error: failure,
}

void transaction
void outcome
