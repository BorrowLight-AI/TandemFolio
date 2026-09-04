import { mkdir, readFile, readdir, rename, rm, writeFile } from 'node:fs/promises'
import { join } from 'node:path'
import { createHash, randomUUID } from 'node:crypto'
import { SessionError, type LiveSession } from './session-store'

export interface RecoverySnapshot {
  sessionId: string
  format: LiveSession['format']
  fileName: string
  data: Buffer
}

interface RecoveryMetadata {
  version: 2
  sessionId: string
  format: LiveSession['format']
  fileName: string
  byteLength: number
  updatedAt: string
}

function recoveryStem(format: LiveSession['format'], sessionId: string): string {
  const identity = createHash('sha256').update(sessionId).digest('hex').slice(0, 32)
  return `${format}.${identity}`
}

export class RecoveryStore {
  readonly #now: () => number
  readonly #retentionMs: number
  readonly #maxBytes: number

  constructor(
    readonly directory: string,
    options: { now?: () => number; retentionMs?: number; maxBytes?: number } = {},
  ) {
    this.#now = options.now ?? Date.now
    this.#retentionMs = options.retentionMs ?? 7 * 86_400_000
    this.#maxBytes = options.maxBytes ?? 268_435_456
  }

  async save(snapshot: RecoverySnapshot): Promise<void> {
    if (snapshot.data.length > this.#maxBytes) {
      throw new SessionError(
        'invalid_arguments',
        `Recovery snapshot exceeds ${this.#maxBytes} bytes.`,
      )
    }
    await mkdir(this.directory, { recursive: true })
    const token = randomUUID()
    const stem = recoveryStem(snapshot.format, snapshot.sessionId)
    const binaryPath = join(this.directory, `${stem}.bin`)
    const metadataPath = join(this.directory, `${stem}.json`)
    const temporaryBinaryPath = join(this.directory, `${stem}.${token}.bin.tmp`)
    const temporaryMetadataPath = join(this.directory, `${stem}.${token}.json.tmp`)
    const metadata: RecoveryMetadata = {
      version: 2,
      sessionId: snapshot.sessionId,
      format: snapshot.format,
      fileName: snapshot.fileName,
      byteLength: snapshot.data.length,
      updatedAt: new Date(this.#now()).toISOString(),
    }

    try {
      await writeFile(temporaryBinaryPath, snapshot.data)
      await writeFile(temporaryMetadataPath, JSON.stringify(metadata))
      await rename(temporaryBinaryPath, binaryPath)
      await rename(temporaryMetadataPath, metadataPath)
    } finally {
      await Promise.all([
        rm(temporaryBinaryPath, { force: true }),
        rm(temporaryMetadataPath, { force: true }),
      ])
    }
  }

  async latest(
    format: LiveSession['format'],
    sessionId?: string,
  ): Promise<RecoverySnapshot | null> {
    if (sessionId) {
      const stem = recoveryStem(format, sessionId)
      return this.#read(format, join(this.directory, `${stem}.json`), sessionId)
    }
    try {
      const names = await readdir(this.directory)
      const candidates = await Promise.all(
        names
          .filter(
            (name) =>
              (name.startsWith(`${format}.`) && name.endsWith('.json')) ||
              name === `${format}.json`,
          )
          .map((name) => this.#read(format, join(this.directory, name))),
      )
      let latest: { snapshot: RecoverySnapshot; updatedAt: number } | null = null
      for (const candidate of candidates) {
        if (!candidate) continue
        const metadataName = `${recoveryStem(format, candidate.sessionId)}.json`
        let updatedAt = 0
        try {
          const metadata = JSON.parse(
            await readFile(join(this.directory, metadataName), 'utf8'),
          ) as { updatedAt?: string }
          updatedAt = Date.parse(metadata.updatedAt ?? '')
        } catch {
          // Legacy snapshots have no session-keyed metadata path.
        }
        if (!latest || updatedAt >= latest.updatedAt) latest = { snapshot: candidate, updatedAt }
      }
      return latest?.snapshot ?? null
    } catch (error) {
      if ((error as NodeJS.ErrnoException).code === 'ENOENT') return null
      throw error
    }
  }

  async clear(format: LiveSession['format'], sessionId?: string): Promise<void> {
    if (sessionId) {
      const stem = recoveryStem(format, sessionId)
      await Promise.all([
        rm(join(this.directory, `${stem}.bin`), { force: true }),
        rm(join(this.directory, `${stem}.json`), { force: true }),
      ])
      return
    }
    try {
      const names = await readdir(this.directory)
      await Promise.all(
        names
          .filter(
            (name) =>
              name === `${format}.bin` ||
              name === `${format}.json` ||
              (name.startsWith(`${format}.`) && (name.endsWith('.bin') || name.endsWith('.json'))),
          )
          .map((name) => rm(join(this.directory, name), { force: true })),
      )
    } catch (error) {
      if ((error as NodeJS.ErrnoException).code !== 'ENOENT') throw error
    }
  }

  async #read(
    format: LiveSession['format'],
    metadataPath: string,
    expectedSessionId?: string,
  ): Promise<RecoverySnapshot | null> {
    try {
      const metadata = JSON.parse(await readFile(metadataPath, 'utf8')) as
        | RecoveryMetadata
        | {
            version: 1
            format: LiveSession['format']
            fileName: string
            byteLength: number
            updatedAt: string
          }
      const sessionId = metadata.version === 2 ? metadata.sessionId : `legacy:${format}`
      if (expectedSessionId && sessionId !== expectedSessionId) return null
      const data = await readFile(metadataPath.replace(/\.json$/, '.bin'))
      const updatedAt = Date.parse(metadata.updatedAt)
      if (
        ![1, 2].includes(metadata.version) ||
        metadata.format !== format ||
        metadata.byteLength !== data.length ||
        typeof metadata.fileName !== 'string' ||
        typeof sessionId !== 'string' ||
        !Number.isFinite(updatedAt)
      ) {
        return null
      }
      if (this.#now() - updatedAt > this.#retentionMs) {
        await Promise.all([
          rm(metadataPath, { force: true }),
          rm(metadataPath.replace(/\.json$/, '.bin'), { force: true }),
        ])
        return null
      }
      return { sessionId, format, fileName: metadata.fileName, data }
    } catch (error) {
      if ((error as NodeJS.ErrnoException).code === 'ENOENT') return null
      throw error
    }
  }
}

interface RecoveryUpload {
  uploadId: string
  sessionId: string
  format: LiveSession['format']
  fileName: string
  data: Buffer
  nextOffset: number
}

export class RecoveryUploadStore {
  readonly #uploads = new Map<string, RecoveryUpload>()

  constructor(readonly recoveries: RecoveryStore) {}

  hasPending(sessionId: string): boolean {
    return [...this.#uploads.values()].some((upload) => upload.sessionId === sessionId)
  }

  invalidate(sessionId: string): void {
    for (const [id, upload] of this.#uploads) {
      if (upload.sessionId === sessionId) this.#uploads.delete(id)
    }
  }

  begin(sessionId: string, format: LiveSession['format'], fileName: string, size: number): string {
    const upload: RecoveryUpload = {
      uploadId: randomUUID(),
      sessionId,
      format,
      fileName,
      data: Buffer.alloc(size),
      nextOffset: 0,
    }
    this.#uploads.set(upload.uploadId, upload)
    return upload.uploadId
  }

  write(sessionId: string, uploadId: string, offset: number, base64: string): number {
    const upload = this.#get(sessionId, uploadId)
    if (offset !== upload.nextOffset) {
      throw new SessionError(
        'revision_conflict',
        `Recovery upload expected offset ${upload.nextOffset}, received ${offset}.`,
      )
    }
    const chunk = Buffer.from(base64, 'base64')
    if (offset + chunk.length > upload.data.length) {
      throw new SessionError('invalid_arguments', 'Recovery chunk exceeds the declared size.')
    }
    chunk.copy(upload.data, offset)
    upload.nextOffset += chunk.length
    return upload.nextOffset
  }

  async commit(sessionId: string, uploadId: string): Promise<void> {
    const upload = this.#get(sessionId, uploadId)
    if (upload.nextOffset !== upload.data.length) {
      throw new SessionError(
        'invalid_arguments',
        `Recovery upload is incomplete at ${upload.nextOffset} of ${upload.data.length} bytes.`,
      )
    }
    await this.recoveries.save({
      sessionId: upload.sessionId,
      format: upload.format,
      fileName: upload.fileName,
      data: upload.data,
    })
    this.#uploads.delete(uploadId)
  }

  #get(sessionId: string, uploadId: string): RecoveryUpload {
    const upload = this.#uploads.get(uploadId)
    if (!upload || upload.sessionId !== sessionId) {
      throw new SessionError('command_not_found', `Unknown recovery upload: ${uploadId}`)
    }
    return upload
  }
}
