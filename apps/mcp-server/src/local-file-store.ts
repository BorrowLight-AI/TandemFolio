import { readFile } from 'node:fs/promises'
import { basename, dirname, extname, isAbsolute, resolve } from 'node:path'
import { randomUUID } from 'node:crypto'
import { SessionError, type LiveSession } from './session-store'

interface StagedLocalFile {
  blobId: string
  sessionId: string
  name: string
  data: Buffer
}

interface LocalAssetRoot {
  rootId: string
  sessionId: string
  directory: string
  assets: Map<string, { data: Buffer; mime: 'image/png' | 'image/jpeg' | 'image/gif' }>
}

export interface LocalFileDescriptor {
  blobId: string
  name: string
  size: number
}

export interface LocalFileChunk {
  offset: number
  data: string
  nextOffset: number
  eof: boolean
}

export interface LocalAssetChunk extends LocalFileChunk {
  size: number
  mime: 'image/png' | 'image/jpeg' | 'image/gif'
}

const formatExtensions: Record<LiveSession['format'], string[]> = {
  docx: ['.docx'],
  markdown: ['.md', '.markdown'],
  xlsx: ['.xlsx'],
  pptx: ['.pptx'],
  pdf: ['.pdf'],
}

export class LocalFileStore {
  readonly #files = new Map<string, StagedLocalFile>()
  readonly #assetRoots = new Map<string, LocalAssetRoot>()

  async stage(
    sessionId: string,
    format: LiveSession['format'],
    path: string,
  ): Promise<LocalFileDescriptor> {
    if (!isAbsolute(path)) {
      throw new SessionError('invalid_arguments', 'Local file paths must be absolute.')
    }
    const expectedExtensions = formatExtensions[format]
    if (!expectedExtensions.includes(extname(path).toLowerCase())) {
      throw new SessionError(
        'invalid_arguments',
        `A ${format} session can only open ${expectedExtensions.join(' or ')} files.`,
      )
    }
    return this.stageBuffer(sessionId, basename(path), await readFile(path))
  }

  async stageImage(sessionId: string, path: string): Promise<LocalFileDescriptor> {
    if (!isAbsolute(path)) {
      throw new SessionError('invalid_arguments', 'Local image paths must be absolute.')
    }
    const extension = extname(path).toLowerCase()
    if (!['.png', '.jpg', '.jpeg', '.gif'].includes(extension)) {
      throw new SessionError('invalid_arguments', 'Only PNG, JPEG, and GIF images can be inserted.')
    }
    const data = await readFile(path)
    validateLocalImage(extension, data)
    return this.stageBuffer(sessionId, basename(path), data)
  }

  registerAssetRoot(sessionId: string, documentPath: string): string {
    if (!isAbsolute(documentPath)) {
      throw new SessionError('invalid_arguments', 'Local file paths must be absolute.')
    }
    this.releaseAssetRoots(sessionId)
    const rootId = randomUUID()
    this.#assetRoots.set(rootId, {
      rootId,
      sessionId,
      directory: dirname(documentPath),
      assets: new Map(),
    })
    return rootId
  }

  async readLocalAsset(
    sessionId: string,
    rootId: string,
    authoredPath: string,
    offset: number,
    length: number,
  ): Promise<LocalAssetChunk> {
    const root = this.#assetRoots.get(rootId)
    if (!root || root.sessionId !== sessionId) {
      throw new SessionError('command_not_found', `Unknown Markdown asset root: ${rootId}`)
    }
    if (/^[a-z][a-z0-9+.-]*:/i.test(authoredPath)) {
      throw new SessionError(
        'invalid_arguments',
        'Only local Markdown image paths can use the asset bridge.',
      )
    }
    let decodedPath: string
    try {
      decodedPath = decodeURIComponent(authoredPath.split(/[?#]/, 1)[0] ?? '')
    } catch {
      throw new SessionError('invalid_arguments', 'The Markdown image path is malformed.')
    }
    const path = isAbsolute(decodedPath) ? decodedPath : resolve(root.directory, decodedPath)
    let asset = root.assets.get(path)
    if (!asset) {
      const extension = extname(path).toLowerCase()
      if (!['.png', '.jpg', '.jpeg', '.gif'].includes(extension)) {
        throw new SessionError(
          'invalid_arguments',
          'Only PNG, JPEG, and GIF Markdown assets are supported.',
        )
      }
      const data = await readFile(path)
      validateLocalImage(extension, data)
      asset = {
        data,
        mime:
          extension === '.png' ? 'image/png' : extension === '.gif' ? 'image/gif' : 'image/jpeg',
      }
      root.assets.set(path, asset)
    }
    const end = Math.min(offset + length, asset.data.length)
    return {
      offset,
      size: asset.data.length,
      mime: asset.mime,
      data: asset.data.subarray(offset, end).toString('base64'),
      nextOffset: end,
      eof: end >= asset.data.length,
    }
  }

  releaseAssetRoots(sessionId: string): void {
    for (const [rootId, root] of this.#assetRoots) {
      if (root.sessionId === sessionId) this.#assetRoots.delete(rootId)
    }
  }

  stageBuffer(sessionId: string, name: string, data: Buffer): LocalFileDescriptor {
    const file: StagedLocalFile = {
      blobId: randomUUID(),
      sessionId,
      name,
      data: Buffer.from(data),
    }
    this.#files.set(file.blobId, file)
    return { blobId: file.blobId, name: file.name, size: file.data.length }
  }

  read(sessionId: string, blobId: string, offset: number, length: number): LocalFileChunk {
    const file = this.#files.get(blobId)
    if (!file || file.sessionId !== sessionId) {
      throw new SessionError('command_not_found', `Unknown staged file: ${blobId}`)
    }
    const end = Math.min(offset + length, file.data.length)
    const chunk = {
      offset,
      data: file.data.subarray(offset, end).toString('base64'),
      nextOffset: end,
      eof: end >= file.data.length,
    }
    if (chunk.eof) this.release(blobId)
    return chunk
  }

  release(blobId: string): void {
    this.#files.delete(blobId)
  }
}

function validateLocalImage(extension: string, data: Buffer): void {
  if (data.length === 0 || data.length > 20 * 1024 * 1024) {
    throw new SessionError('invalid_arguments', 'Local images must be between 1 byte and 20 MB.')
  }
  const valid =
    (extension === '.png' &&
      data.subarray(0, 8).equals(Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]))) ||
    ((extension === '.jpg' || extension === '.jpeg') &&
      data.length >= 3 &&
      data[0] === 0xff &&
      data[1] === 0xd8 &&
      data[2] === 0xff) ||
    (extension === '.gif' &&
      data.length >= 6 &&
      (data.subarray(0, 6).toString('ascii') === 'GIF87a' ||
        data.subarray(0, 6).toString('ascii') === 'GIF89a'))
  if (!valid) {
    throw new SessionError('invalid_arguments', 'The local image bytes do not match its extension.')
  }
}
