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
  companions: Map<string, DocumentSaveCompanionUpload>
  removeCompanionPaths: string[]
}

export interface DocumentSaveCompanion {
  readonly relativePath: string
  readonly size: number
}

interface DocumentSaveCompanionUpload extends DocumentSaveCompanion {
  targetPath: string
  temporaryPath: string
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

function assertCompanion(companion: DocumentSaveCompanion): void {
  const segments = companion.relativePath.split('/')
  if (
    companion.relativePath.length < 1 ||
    companion.relativePath.length > 512 ||
    companion.relativePath.includes('\\') ||
    companion.relativePath.startsWith('/') ||
    segments.some((segment) => !segment || segment === '.' || segment === '..') ||
    !/\.(png|jpe?g|gif)$/i.test(companion.relativePath) ||
    !Number.isInteger(companion.size) ||
    companion.size < 1 ||
    companion.size > 20_971_520
  ) {
    throw new SessionError(
      'invalid_arguments',
      'Markdown companion assets must be bounded relative PNG, JPEG, or GIF paths.',
    )
  }
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
  readonly #ownedCompanions = new Map<string, Set<string>>()

  constructor(
    readonly outputDirectory: string,
    readonly bindingDirectory: string,
    readonly maxBytes = 268_435_456,
  ) {}

  hasPending(sessionId: string): boolean {
    return [...this.#uploads.values()].some((upload) => upload.sessionId === sessionId)
  }

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

  async unbind(sessionId: string): Promise<void> {
    for (const upload of this.#uploads.values()) {
      if (upload.sessionId === sessionId) await this.#discard(upload)
    }
    await rm(this.#bindingPath(sessionId), { force: true })
    this.#bindings.delete(sessionId)
    this.#ownedCompanions.delete(sessionId)
  }

  async begin(
    sessionId: string,
    format: LiveSession['format'],
    fileName: string,
    size: number,
    mode: DocumentSaveMode,
    companionFiles: readonly DocumentSaveCompanion[] = [],
    removeCompanionPaths: readonly string[] = [],
  ): Promise<{ uploadId: string; path: string }> {
    if (!Number.isInteger(size) || size < 0 || size > this.maxBytes) {
      throw new SessionError(
        'invalid_arguments',
        `Saved documents must be between 0 and ${this.maxBytes} bytes.`,
      )
    }
    if (
      companionFiles.length > 128 ||
      removeCompanionPaths.length > 128 ||
      ((companionFiles.length > 0 || removeCompanionPaths.length > 0) && format !== 'markdown')
    ) {
      throw new SessionError(
        'invalid_arguments',
        'Only Markdown saves may include up to 128 companion assets.',
      )
    }
    const companionPaths = new Set<string>()
    let totalSize = size
    for (const companion of companionFiles) {
      assertCompanion(companion)
      if (companionPaths.has(companion.relativePath)) {
        throw new SessionError(
          'invalid_arguments',
          'Markdown companion asset paths must be unique.',
        )
      }
      companionPaths.add(companion.relativePath)
      totalSize += companion.size
    }
    const uniqueRemovals = [...new Set(removeCompanionPaths)]
    for (const relativePath of uniqueRemovals) {
      assertCompanion({ relativePath, size: 1 })
      if (!/^assets\/[a-z0-9_.-]+-[a-f0-9]{12}\.(?:png|jpe?g|gif)$/i.test(relativePath)) {
        throw new SessionError(
          'invalid_arguments',
          'Only content-addressed Markdown companion assets may be collected.',
        )
      }
    }
    if (totalSize > this.maxBytes) {
      throw new SessionError(
        'invalid_arguments',
        `Saved document bundles must be at most ${this.maxBytes} bytes.`,
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
    const companions = new Map<string, DocumentSaveCompanionUpload>()
    try {
      if (await pathExists(targetPath)) {
        const target = await stat(targetPath)
        if (!target.isFile()) {
          throw new SessionError('invalid_arguments', 'The saved document target is not a file.')
        }
        await handle.chmod(target.mode)
      }
      for (const companion of companionFiles) {
        const companionTargetPath = join(dirname(targetPath), ...companion.relativePath.split('/'))
        const companionTemporaryPath = `${companionTargetPath}.${randomUUID()}.save.tmp`
        await mkdir(dirname(companionTargetPath), { recursive: true })
        const companionHandle = await open(companionTemporaryPath, 'wx')
        companions.set(companion.relativePath, {
          ...companion,
          targetPath: companionTargetPath,
          temporaryPath: companionTemporaryPath,
          nextOffset: 0,
          handle: companionHandle,
          closed: false,
        })
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
        companions,
        removeCompanionPaths: uniqueRemovals,
      }
      this.#uploads.set(upload.uploadId, upload)
      this.#reservedTargets.add(targetPath)
      return { uploadId: upload.uploadId, path: targetPath }
    } catch (error) {
      await handle.close()
      await rm(temporaryPath, { force: true })
      for (const companion of companions.values()) {
        await companion.handle.close().catch(() => undefined)
        await rm(companion.temporaryPath, { force: true })
      }
      throw error
    }
  }

  async write(
    sessionId: string,
    uploadId: string,
    offset: number,
    base64: string,
    relativePath?: string,
  ): Promise<number> {
    const upload = this.#get(sessionId, uploadId)
    const target = relativePath ? upload.companions.get(relativePath) : upload
    if (!target) {
      throw new SessionError(
        'invalid_arguments',
        `Unknown Markdown companion asset: ${relativePath}`,
      )
    }
    if (offset !== target.nextOffset) {
      throw new SessionError(
        'revision_conflict',
        `Document save upload expected offset ${target.nextOffset}, received ${offset}.`,
      )
    }
    if (
      base64.length > 262_144 ||
      !/^(?:[A-Za-z0-9+/]{4})*(?:[A-Za-z0-9+/]{2}==|[A-Za-z0-9+/]{3}=)?$/.test(base64)
    ) {
      throw new SessionError('invalid_arguments', 'Document save chunks must be bounded base64.')
    }
    const chunk = Buffer.from(base64, 'base64')
    if (chunk.length === 0 || offset + chunk.length > target.size) {
      throw new SessionError('invalid_arguments', 'Document save chunk exceeds the declared size.')
    }
    let written = 0
    while (written < chunk.length) {
      const result = await target.handle.write(
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
    target.nextOffset += chunk.length
    return target.nextOffset
  }

  async commit(sessionId: string, uploadId: string): Promise<{ path: string; bound: boolean }> {
    const upload = this.#get(sessionId, uploadId)
    if (upload.nextOffset !== upload.size) {
      throw new SessionError(
        'invalid_arguments',
        `Document save upload is incomplete at ${upload.nextOffset} of ${upload.size} bytes.`,
      )
    }
    for (const companion of upload.companions.values()) {
      if (companion.nextOffset !== companion.size) {
        throw new SessionError(
          'invalid_arguments',
          `Markdown companion ${companion.relativePath} is incomplete at ${companion.nextOffset} of ${companion.size} bytes.`,
        )
      }
    }
    const createdCompanions: string[] = []
    try {
      await upload.handle.sync()
      await upload.handle.close()
      upload.closed = true
      for (const companion of upload.companions.values()) {
        await companion.handle.sync()
        await companion.handle.close()
        companion.closed = true
        if (await pathExists(companion.targetPath)) {
          const existing = await stat(companion.targetPath)
          if (!existing.isFile()) {
            throw new SessionError(
              'invalid_arguments',
              `Markdown companion target is not a file: ${companion.relativePath}`,
            )
          }
          const [existingBytes, stagedBytes] = await Promise.all([
            readFile(companion.targetPath),
            readFile(companion.temporaryPath),
          ])
          if (!existingBytes.equals(stagedBytes)) {
            throw new SessionError(
              'revision_conflict',
              `Markdown companion target already contains different bytes: ${companion.relativePath}`,
            )
          }
          await rm(companion.temporaryPath, { force: true })
        }
      }
      for (const companion of upload.companions.values()) {
        if (!(await pathExists(companion.temporaryPath))) continue
        await rename(companion.temporaryPath, companion.targetPath)
        createdCompanions.push(companion.targetPath)
      }
      await rename(upload.temporaryPath, upload.targetPath)
      const bound = upload.mode !== 'export-copy'
      if (bound) await this.bind(upload.sessionId, upload.format, upload.targetPath)
      if (bound && upload.format === 'markdown') {
        const owned =
          upload.mode === 'save-as'
            ? new Set<string>()
            : new Set(this.#ownedCompanions.get(upload.sessionId) ?? [])
        for (const companion of upload.companions.values()) owned.add(companion.targetPath)
        for (const relativePath of upload.removeCompanionPaths) {
          const targetPath = join(dirname(upload.targetPath), ...relativePath.split('/'))
          if (!owned.has(targetPath)) continue
          try {
            const expected = /-([a-f0-9]{12})\.(?:png|jpe?g|gif)$/i.exec(targetPath)?.[1]
            const actual = createHash('sha256')
              .update(await readFile(targetPath))
              .digest('hex')
            if (expected && actual.startsWith(expected)) await rm(targetPath)
          } catch {
            // Missing or user-modified assets are preserved; ownership is relinquished.
          }
          owned.delete(targetPath)
        }
        this.#ownedCompanions.set(upload.sessionId, owned)
      }
      this.#release(upload)
      return { path: upload.targetPath, bound }
    } catch (error) {
      await Promise.all(createdCompanions.map((path) => rm(path, { force: true })))
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
    for (const companion of upload.companions.values()) {
      if (!companion.closed) await companion.handle.close().catch(() => undefined)
      await rm(companion.temporaryPath, { force: true })
    }
    this.#release(upload)
  }
}
