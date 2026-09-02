import { randomUUID } from 'node:crypto'

interface TransactionRecord<Response> {
  accepted: boolean
  fingerprint: string
  promise: Promise<Response>
  resolve: (response: Response) => void
  transactionId: string
}

export type TransactionJournalStart<Response> =
  | { kind: 'conflict' }
  | { kind: 'replay'; promise: Promise<Response> }
  | {
      kind: 'new'
      promise: Promise<Response>
      transactionId: string
      markAccepted: () => void
      complete: (response: Response) => void
      fail: (response: Response) => void
    }

function fingerprintPayload(value: unknown): string {
  const normalize = (candidate: unknown): unknown => {
    if (Array.isArray(candidate)) return candidate.map(normalize)
    if (candidate === null || typeof candidate !== 'object') return candidate
    return Object.fromEntries(
      Object.entries(candidate)
        .sort(([left], [right]) => left.localeCompare(right))
        .map(([key, nested]) => [key, normalize(nested)]),
    )
  }
  return JSON.stringify(normalize(value))
}

export class TransactionJournal<Response> {
  readonly #records = new Map<string, TransactionRecord<Response>>()

  start(sessionId: string, requestId: string, payload: unknown): TransactionJournalStart<Response> {
    const key = `${sessionId}\0${requestId}`
    const fingerprint = fingerprintPayload(payload)
    const existing = this.#records.get(key)
    if (existing) {
      return existing.fingerprint === fingerprint
        ? { kind: 'replay', promise: existing.promise }
        : { kind: 'conflict' }
    }

    let resolveTransaction!: (response: Response) => void
    const record: TransactionRecord<Response> = {
      accepted: false,
      fingerprint,
      promise: new Promise<Response>((resolve) => {
        resolveTransaction = resolve
      }),
      resolve: (response) => resolveTransaction(response),
      transactionId: randomUUID(),
    }
    this.#records.set(key, record)

    return {
      kind: 'new',
      promise: record.promise,
      transactionId: record.transactionId,
      markAccepted: () => {
        record.accepted = true
      },
      complete: (response) => record.resolve(response),
      fail: (response) => {
        if (!record.accepted) this.#records.delete(key)
        record.resolve(response)
      },
    }
  }
}
