import { createHash, randomUUID } from 'node:crypto'
import {
  access,
  mkdir,
  open,
  readFile,
  rename,
  rm,
  stat,
  writeFile,
  type FileHandle,
} from 'node:fs/promises'
import { basename, dirname, extname, isAbsolute, join } from 'node:path'
import { SessionError, type LiveSession } from './session-store'

export type DocumentSaveMode = 'save' | 'save-as' | 'export-copy'

interface DocumentBindingMetadata {
  version: 1
  sessionId: string
  format: LiveSession['format']
  path: string
}

interface DocumentSaveUpload {
  uploadId: string
  sessionId: string
  format: LiveSession['format']
  fileName: string
  mode: DocumentSaveMode
  targetPath: string
  temporaryPath: string
  size: number
  nextOffset: number
  handle: FileHandle
  closed: boolean
}

const formatExtensions: Record<LiveSession['format'], readonly string[]> = {
  docx: ['.docx'],
  markdown: ['.md', '.markdown'],
  xlsx: ['.xlsx'],
  pptx: ['.pptx'],
  pdf: ['.pdf'],
}

function sessionStem(sessionId: string): string {
  return createHash('sha256').update(sessionId).digest('hex').slice(0, 32)
}

function assertFileName(format: LiveSession['format'], fileName: string): void {
  if (
    fileName.length < 1 ||
    fileName.length > 240 ||
    fileName !== basename(fileName) ||
    fileName.includes('/') ||
    fileName.includes('\\') ||
    fileName.includes('\0') ||
    fileName === '.' ||
    fileName === '..'
  ) {
    throw new SessionError(
      'invalid_arguments',
      'Saved document names must be plain file names without path separators.',
    )
  }
  const extensions = formatExtensions[format]
  if (!extensions.includes(extname(fileName).toLowerCase())) {
    throw new SessionError(
      'invalid_arguments',
      `A ${format} document must use ${extensions.join(' or ')}.`,
    )
  }
}

function assertTargetPath(format: LiveSession['format'], path: string): void {
  if (!isAbsolute(path)) {
    throw new SessionError('invalid_arguments', 'Saved document paths must be absolute.')
  }
  assertFileName(format, basename(path))
}

async function pathExists(path: string): Promise<boolean> {
  try {
    await access(path)
    return true
  } catch (error) {
    if ((error as NodeJS.ErrnoException).code === 'ENOENT') return false
    throw error
  }
}

export class DocumentSaveStore {
  readonly #uploads = new Map<string, DocumentSaveUpload>()
  readonly #bindings = new Map<string, DocumentBindingMetadata>()
  readonly #reservedTargets = new Set<string>()

  constructor(
    readonly outputDirectory: string,
    readonly bindingDirectory: string,
    readonly maxBytes = 268_435_456,
  ) {}

  async bind(sessionId: string, format: LiveSession['format'], path: string): Promise<void> {
    assertTargetPath(format, path)
    await mkdir(this.bindingDirectory, { recursive: true })
    const metadata: DocumentBindingMetadata = { version: 1, sessionId, format, path }
    const metadataPath = this.#bindingPath(sessionId)
    const temporaryPath = `${metadataPath}.${randomUUID()}.tmp`
    try {
      await writeFile(temporaryPath, JSON.stringify(metadata), { flag: 'wx' })
      await rename(temporaryPath, metadataPath)
      this.#bindings.set(sessionId, metadata)
    } finally {
      await rm(temporaryPath, { force: true })
    }
  }

  async boundPath(sessionId: string, format: LiveSession['format']): Promise<string | null> {
    return (await this.#binding(sessionId, format))?.path ?? null
  }

  async begin(
    sessionId: string,
    format: LiveSession['format'],
    fileName: string,
    size: number,
    mode: DocumentSaveMode,
  ): Promise<{ uploadId: string; path: string }> {
    if (!Number.isInteger(size) || size < 0 || size > this.maxBytes) {
      throw new SessionError(
        'invalid_arguments',
        `Saved documents must be between 0 and ${this.maxBytes} bytes.`,
      )
    }
    const binding = mode === 'save' ? await this.#binding(sessionId, format) : null
    const persistedFileName = binding ? basename(binding.path) : fileName
    assertFileName(format, persistedFileName)
    const targetPath =
      binding?.path ?? (await this.#availableOutputPath(sessionId, format, fileName))
    const temporaryPath = join(
      dirname(targetPath),
      `.${basename(targetPath)}.${randomUUID()}.save.tmp`,
    )
    await mkdir(dirname(targetPath), { recursive: true })
    const handle = await open(temporaryPath, 'wx')
    try {
      if (await pathExists(targetPath)) {
        const target = await stat(targetPath)
        if (!target.isFile()) {
          throw new SessionError('invalid_arguments', 'The saved document target is not a file.')
        }
        await handle.chmod(target.mode)
      }
      const upload: DocumentSaveUpload = {
        uploadId: randomUUID(),
        sessionId,
        format,
        fileName: persistedFileName,
        mode,
        targetPath,
        temporaryPath,
        size,
        nextOffset: 0,
        handle,
        closed: false,
      }
      this.#uploads.set(upload.uploadId, upload)
      this.#reservedTargets.add(targetPath)
      return { uploadId: upload.uploadId, path: targetPath }
    } catch (error) {
      await handle.close()
      await rm(temporaryPath, { force: true })
      throw error
    }
  }

  async write(
    sessionId: string,
    uploadId: string,
    offset: number,
    base64: string,
  ): Promise<number> {
    const upload = this.#get(sessionId, uploadId)
    if (offset !== upload.nextOffset) {
      throw new SessionError(
        'revision_conflict',
        `Document save upload expected offset ${upload.nextOffset}, received ${offset}.`,
      )
    }
    if (
      base64.length > 262_144 ||
      !/^(?:[A-Za-z0-9+/]{4})*(?:[A-Za-z0-9+/]{2}==|[A-Za-z0-9+/]{3}=)?$/.test(base64)
    ) {
      throw new SessionError('invalid_arguments', 'Document save chunks must be bounded base64.')
    }
    const chunk = Buffer.from(base64, 'base64')
    if (chunk.length === 0 || offset + chunk.length > upload.size) {
      throw new SessionError('invalid_arguments', 'Document save chunk exceeds the declared size.')
    }
    let written = 0
    while (written < chunk.length) {
      const result = await upload.handle.write(
        chunk,
        written,
        chunk.length - written,
        offset + written,
      )
      if (result.bytesWritten < 1) {
        throw new SessionError(
          'execution_failed',
          'The document save target stopped accepting data.',
        )
      }
      written += result.bytesWritten
    }
    upload.nextOffset += chunk.length
    return upload.nextOffset
  }

  async commit(sessionId: string, uploadId: string): Promise<{ path: string; bound: boolean }> {
    const upload = this.#get(sessionId, uploadId)
    if (upload.nextOffset !== upload.size) {
      throw new SessionError(
        'invalid_arguments',
        `Document save upload is incomplete at ${upload.nextOffset} of ${upload.size} bytes.`,
      )
    }
    try {
      await upload.handle.sync()
      await upload.handle.close()
      upload.closed = true
      await rename(upload.temporaryPath, upload.targetPath)
      const bound = upload.mode !== 'export-copy'
      if (bound) await this.bind(upload.sessionId, upload.format, upload.targetPath)
      this.#release(upload)
      return { path: upload.targetPath, bound }
    } catch (error) {
      await this.#discard(upload)
      throw error
    }
  }

  async abort(sessionId: string, uploadId: string): Promise<void> {
    await this.#discard(this.#get(sessionId, uploadId))
  }

  #bindingPath(sessionId: string): string {
    return join(this.bindingDirectory, `${sessionStem(sessionId)}.json`)
  }

  async #binding(
    sessionId: string,
    format: LiveSession['format'],
  ): Promise<DocumentBindingMetadata | null> {
    const cached = this.#bindings.get(sessionId)
    if (cached) {
      if (cached.format !== format) {
        throw new SessionError(
          'invalid_arguments',
          'The saved document binding has another format.',
        )
      }
      return cached
    }
    try {
      const metadata = JSON.parse(
        await readFile(this.#bindingPath(sessionId), 'utf8'),
      ) as DocumentBindingMetadata
      if (
        metadata.version !== 1 ||
        metadata.sessionId !== sessionId ||
        metadata.format !== format ||
        typeof metadata.path !== 'string'
      ) {
        throw new SessionError('execution_failed', 'The saved document binding is invalid.')
      }
      assertTargetPath(format, metadata.path)
      this.#bindings.set(sessionId, metadata)
      return metadata
    } catch (error) {
      if ((error as NodeJS.ErrnoException).code === 'ENOENT') return null
      throw error
    }
  }

  async #availableOutputPath(
    sessionId: string,
    format: LiveSession['format'],
    fileName: string,
  ): Promise<string> {
    const directory = join(this.outputDirectory, `${format}-${sessionStem(sessionId)}`)
    await mkdir(directory, { recursive: true })
    const extension = extname(fileName)
    const stem = fileName.slice(0, -extension.length)
    for (let index = 0; index < 10_000; index += 1) {
      const candidate = join(directory, `${stem}${index === 0 ? '' : ` (${index})`}${extension}`)
      if (!this.#reservedTargets.has(candidate) && !(await pathExists(candidate))) return candidate
    }
    throw new SessionError('execution_failed', 'No available local document save target remains.')
  }

  #get(sessionId: string, uploadId: string): DocumentSaveUpload {
    const upload = this.#uploads.get(uploadId)
    if (!upload || upload.sessionId !== sessionId) {
      throw new SessionError('command_not_found', `Unknown document save upload: ${uploadId}`)
    }
    return upload
  }

  #release(upload: DocumentSaveUpload): void {
    this.#uploads.delete(upload.uploadId)
    this.#reservedTargets.delete(upload.targetPath)
  }

  async #discard(upload: DocumentSaveUpload): Promise<void> {
    if (!upload.closed) {
      await upload.handle.close().catch(() => undefined)
      upload.closed = true
    }
    await rm(upload.temporaryPath, { force: true })
    this.#release(upload)
  }
}
