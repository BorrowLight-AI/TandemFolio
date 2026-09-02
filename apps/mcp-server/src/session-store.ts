import { randomUUID } from 'node:crypto'

export interface QueuedCommand {
  commandId: string
  baseRevision: number
  operation: string
  arguments: Record<string, unknown>
}

export interface LiveSession {
  id: string
  format: 'docx' | 'markdown' | 'xlsx' | 'pptx' | 'pdf'
  revision: number
  connected: boolean
  fileName: string | null
  filePath: string | null
  dirty: boolean
  selection: Record<string, unknown> | null
  pending: QueuedCommand[]
}

export interface CommandCompletion {
  commandId: string
  ok: true
  revision: number
  output?: Record<string, unknown>
}

interface CommandCompletionRecord {
  promise: Promise<CommandCompletion>
  resolve: (completion: CommandCompletion) => void
  reject: (error: SessionError) => void
}

interface PollWaiter {
  resolve: (commands: QueuedCommand[]) => void
  timer: ReturnType<typeof setTimeout>
}

const LEGACY_VIEW_ID = 'legacy-view'

export class SessionError extends Error {
  constructor(
    readonly code:
      | 'session_not_found'
      | 'editor_offline'
      | 'editor_view_conflict'
      | 'revision_conflict'
      | 'command_in_flight'
      | 'command_not_found'
      | 'command_timeout'
      | 'request_reused'
      | 'transaction_not_atomic'
      | 'operation_not_found'
      | 'operation_schema_invalid'
      | 'operation_unavailable'
      | 'unsupported_operation'
      | 'invalid_arguments'
      | 'execution_failed',
    message: string,
  ) {
    super(message)
  }
}

export class SessionStore {
  readonly #sessions = new Map<string, LiveSession>()
  readonly #activeCommands = new Map<string, QueuedCommand>()
  readonly #commandCompletions = new Map<string, CommandCompletionRecord>()
  readonly #pollWaiters = new Map<string, PollWaiter>()
  readonly #activeViews = new Map<string, string>()

  create(format: LiveSession['format'] = 'docx'): LiveSession {
    return this.restore(randomUUID(), format)
  }

  restore(sessionId: string, format: LiveSession['format']): LiveSession {
    const existing = this.#sessions.get(sessionId)
    if (existing) {
      if (existing.format !== format) {
        throw new SessionError(
          'invalid_arguments',
          `Session ${sessionId} belongs to ${existing.format}, not ${format}.`,
        )
      }
      return existing
    }
    const session: LiveSession = {
      id: sessionId,
      format,
      revision: 0,
      connected: false,
      fileName: null,
      filePath: null,
      dirty: false,
      selection: null,
      pending: [],
    }
    this.#sessions.set(sessionId, session)
    return session
  }

  get(sessionId: string): LiveSession {
    const session = this.#sessions.get(sessionId)
    if (!session) throw new SessionError('session_not_found', `Unknown session: ${sessionId}`)
    return session
  }

  assertView(sessionId: string, viewId?: string): LiveSession {
    const session = this.get(sessionId)
    if (viewId === undefined) return session
    const activeViewId = this.#activeViews.get(sessionId)
    if (activeViewId !== viewId) {
      throw new SessionError(
        'editor_view_conflict',
        'This editor view does not own the editing session lease.',
      )
    }
    return session
  }

  connect(
    sessionId: string,
    context?: Partial<Pick<LiveSession, 'fileName' | 'dirty' | 'selection'>>,
    viewId = LEGACY_VIEW_ID,
  ): LiveSession {
    const session = this.get(sessionId)
    const activeViewId = this.#activeViews.get(sessionId)
    if (activeViewId && activeViewId !== viewId) {
      throw new SessionError(
        'editor_view_conflict',
        'This editing session is already connected to another mounted editor view.',
      )
    }
    this.#activeViews.set(sessionId, viewId)
    session.connected = true
    Object.assign(session, context)
    return session
  }

  disconnect(sessionId: string, viewId = LEGACY_VIEW_ID): void {
    const session = this.get(sessionId)
    const activeViewId = this.#activeViews.get(sessionId)
    if (activeViewId && activeViewId !== viewId) {
      throw new SessionError(
        'editor_view_conflict',
        'This editor view does not own the editing session lease.',
      )
    }
    session.connected = false
    this.#activeViews.delete(sessionId)
    this.#finishPoll(sessionId, [])
  }

  assertCanEnqueue(sessionId: string, baseRevision: number): LiveSession {
    const session = this.get(sessionId)
    if (!session.connected) {
      throw new SessionError('editor_offline', 'Open the TandemFolio editor before editing.')
    }
    if (session.revision !== baseRevision) {
      throw new SessionError(
        'revision_conflict',
        `Expected revision ${session.revision}, received ${baseRevision}. Refresh context and retry.`,
      )
    }
    if (session.pending.length > 0 || this.#activeCommands.has(sessionId)) {
      throw new SessionError(
        'command_in_flight',
        'Wait for the active editor command to finish before submitting another mutation.',
      )
    }
    return session
  }

  enqueue(
    sessionId: string,
    baseRevision: number,
    operation: string,
    args: Record<string, unknown>,
  ): QueuedCommand {
    const session = this.assertCanEnqueue(sessionId, baseRevision)
    const command = { commandId: randomUUID(), baseRevision, operation, arguments: args }
    let resolveCompletion!: (completion: CommandCompletion) => void
    let rejectCompletion!: (error: SessionError) => void
    const promise = new Promise<CommandCompletion>((resolve, reject) => {
      resolveCompletion = resolve
      rejectCompletion = reject
    })
    this.#commandCompletions.set(command.commandId, {
      promise,
      resolve: resolveCompletion,
      reject: rejectCompletion,
    })
    session.pending.push(command)
    const waiter = this.#pollWaiters.get(sessionId)
    if (waiter) this.#finishPoll(sessionId, this.#takePending(sessionId))
    return command
  }

  poll(
    sessionId: string,
    context?: Partial<Pick<LiveSession, 'fileName' | 'dirty' | 'selection'>>,
    viewId = LEGACY_VIEW_ID,
  ): QueuedCommand[] {
    this.connect(sessionId, context, viewId)
    return this.#takePending(sessionId)
  }

  #takePending(sessionId: string): QueuedCommand[] {
    const session = this.get(sessionId)
    const commands = session.pending.splice(0)
    if (commands[0]) this.#activeCommands.set(sessionId, commands[0])
    return commands
  }

  waitForPoll(
    sessionId: string,
    context: Partial<Pick<LiveSession, 'fileName' | 'dirty' | 'selection'>> | undefined,
    timeoutMs: number,
    viewId = LEGACY_VIEW_ID,
  ): Promise<QueuedCommand[]> {
    this.connect(sessionId, context, viewId)
    this.#finishPoll(sessionId, [])
    const commands = this.#takePending(sessionId)
    if (commands.length > 0 || timeoutMs === 0) return Promise.resolve(commands)

    return new Promise<QueuedCommand[]>((resolve) => {
      const timer = setTimeout(() => this.#finishPoll(sessionId, []), timeoutMs)
      this.#pollWaiters.set(sessionId, { resolve, timer })
    })
  }

  #finishPoll(sessionId: string, commands: QueuedCommand[]): void {
    const waiter = this.#pollWaiters.get(sessionId)
    if (!waiter) return
    clearTimeout(waiter.timer)
    this.#pollWaiters.delete(sessionId)
    waiter.resolve(commands)
  }

  waitForCommand(
    sessionId: string,
    commandId: string,
    timeoutMs: number | null = 15_000,
  ): Promise<CommandCompletion> {
    const session = this.get(sessionId)
    const command =
      session.pending.find((candidate) => candidate.commandId === commandId) ??
      this.#activeCommands.get(sessionId)
    if (command?.commandId !== commandId) {
      throw new SessionError(
        'command_not_found',
        `Command ${commandId} is not pending or active for this session.`,
      )
    }
    const completion = this.#commandCompletions.get(commandId)
    if (!completion) {
      throw new SessionError(
        'command_not_found',
        `Command ${commandId} has no completion record for this session.`,
      )
    }
    if (timeoutMs === null) return completion.promise

    return new Promise<CommandCompletion>((resolve, reject) => {
      const timer = setTimeout(() => {
        reject(
          new SessionError(
            'command_timeout',
            `Command ${commandId} was not acknowledged within ${timeoutMs} ms.`,
          ),
        )
      }, timeoutMs)
      completion.promise.then(
        (value) => {
          clearTimeout(timer)
          resolve(value)
        },
        (error) => {
          clearTimeout(timer)
          reject(error)
        },
      )
    })
  }

  reject(
    sessionId: string,
    commandId: string,
    code: 'unsupported_operation' | 'invalid_arguments' | 'execution_failed',
    message: string,
  ): LiveSession {
    const session = this.get(sessionId)
    const activeCommand = this.#activeCommands.get(sessionId)
    if (activeCommand?.commandId !== commandId) {
      throw new SessionError(
        'command_not_found',
        `Command ${commandId} is not the active command for this session.`,
      )
    }

    this.#activeCommands.delete(sessionId)
    const completion = this.#commandCompletions.get(commandId)
    this.#commandCompletions.delete(commandId)
    completion?.reject(new SessionError(code, message))
    return session
  }

  acknowledge(
    sessionId: string,
    commandId: string,
    revision: number,
    context: Partial<Pick<LiveSession, 'fileName' | 'dirty' | 'selection'>>,
    output?: Record<string, unknown>,
  ): LiveSession {
    const session = this.get(sessionId)
    const activeCommand = this.#activeCommands.get(sessionId)
    if (activeCommand?.commandId !== commandId) {
      throw new SessionError(
        'command_not_found',
        `Command ${commandId} is not the active command for this session.`,
      )
    }
    if (revision !== session.revision + 1) {
      throw new SessionError(
        'revision_conflict',
        `Acknowledgement must advance revision ${session.revision} to ${session.revision + 1}.`,
      )
    }
    session.revision = revision
    Object.assign(session, context)
    this.#activeCommands.delete(sessionId)
    const completion = this.#commandCompletions.get(commandId)
    this.#commandCompletions.delete(commandId)
    completion?.resolve({ commandId, ok: true, revision, ...(output ? { output } : {}) })
    return session
  }
}
