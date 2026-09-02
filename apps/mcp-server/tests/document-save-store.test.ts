import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, describe, expect, it } from 'vitest'

import { DocumentSaveStore } from '../src/document-save-store'

const temporaryDirectories: string[] = []

afterEach(async () => {
  await Promise.all(
    temporaryDirectories
      .splice(0)
      .map((directory) => rm(directory, { recursive: true, force: true })),
  )
})

async function temporaryDirectory(prefix: string): Promise<string> {
  const directory = await mkdtemp(join(tmpdir(), prefix))
  temporaryDirectories.push(directory)
  return directory
}

async function upload(
  store: DocumentSaveStore,
  input: {
    sessionId: string
    format: 'docx' | 'markdown' | 'xlsx' | 'pptx' | 'pdf'
    fileName: string
    data: Buffer
    mode: 'save' | 'save-as' | 'export-copy'
  },
): Promise<string> {
  const begun = await store.begin(
    input.sessionId,
    input.format,
    input.fileName,
    input.data.length,
    input.mode,
  )
  const split = Math.floor(input.data.length / 2)
  const chunks = [input.data.subarray(0, split), input.data.subarray(split)].filter(
    (chunk) => chunk.length > 0,
  )
  let offset = 0
  for (const chunk of chunks) {
    offset = await store.write(input.sessionId, begun.uploadId, offset, chunk.toString('base64'))
  }
  return (await store.commit(input.sessionId, begun.uploadId)).path
}

describe('DocumentSaveStore', () => {
  it('atomically persists a new renderer-produced document and binds subsequent saves', async () => {
    const root = await temporaryDirectory('tandemfolio-document-save-')
    const outputDirectory = join(root, 'outputs')
    const bindingDirectory = join(root, 'bindings')
    const store = new DocumentSaveStore(outputDirectory, bindingDirectory)

    const firstPath = await upload(store, {
      sessionId: 'session-a',
      format: 'pptx',
      fileName: 'Quarterly Review.pptx',
      data: Buffer.from('first'),
      mode: 'save',
    })

    expect(firstPath).toMatch(/outputs[/\\]pptx-[a-f0-9]{32}[/\\]Quarterly Review\.pptx$/)
    await expect(readFile(firstPath)).resolves.toEqual(Buffer.from('first'))

    const secondPath = await upload(store, {
      sessionId: 'session-a',
      format: 'pptx',
      fileName: 'Quarterly Review.pptx',
      data: Buffer.from('second'),
      mode: 'save',
    })

    expect(secondPath).toBe(firstPath)
    await expect(readFile(firstPath)).resolves.toEqual(Buffer.from('second'))
  })

  it('overwrites only an explicitly opened local file and restores that binding after restart', async () => {
    const root = await temporaryDirectory('tandemfolio-document-save-')
    const sourcePath = join(root, 'source.docx')
    await writeFile(sourcePath, 'original')
    const outputDirectory = join(root, 'outputs')
    const bindingDirectory = join(root, 'bindings')

    await new DocumentSaveStore(outputDirectory, bindingDirectory).bind(
      'session-source',
      'docx',
      sourcePath,
    )
    const restarted = new DocumentSaveStore(outputDirectory, bindingDirectory)
    const savedPath = await upload(restarted, {
      sessionId: 'session-source',
      format: 'docx',
      fileName: 'source.docx',
      data: Buffer.from('edited'),
      mode: 'save',
    })

    expect(savedPath).toBe(sourcePath)
    await expect(readFile(sourcePath, 'utf8')).resolves.toBe('edited')
  })

  it('saves a renamed copy without mutating the opened source and rebinds future saves', async () => {
    const root = await temporaryDirectory('tandemfolio-document-save-')
    const sourcePath = join(root, 'source.md')
    await writeFile(sourcePath, '# source')
    const store = new DocumentSaveStore(join(root, 'outputs'), join(root, 'bindings'))
    await store.bind('session-markdown', 'markdown', sourcePath)

    const copyPath = await upload(store, {
      sessionId: 'session-markdown',
      format: 'markdown',
      fileName: 'renamed.md',
      data: Buffer.from('# copy'),
      mode: 'save-as',
    })

    expect(copyPath).not.toBe(sourcePath)
    await expect(readFile(sourcePath, 'utf8')).resolves.toBe('# source')
    await expect(readFile(copyPath, 'utf8')).resolves.toBe('# copy')

    const reboundPath = await upload(store, {
      sessionId: 'session-markdown',
      format: 'markdown',
      fileName: 'renamed.md',
      data: Buffer.from('# updated copy'),
      mode: 'save',
    })
    expect(reboundPath).toBe(copyPath)
    await expect(readFile(copyPath, 'utf8')).resolves.toBe('# updated copy')
  })

  it('exports a copy without changing the opened document binding', async () => {
    const root = await temporaryDirectory('tandemfolio-document-save-')
    const sourcePath = join(root, 'source.pdf')
    await writeFile(sourcePath, 'source')
    const store = new DocumentSaveStore(join(root, 'outputs'), join(root, 'bindings'))
    await store.bind('session-pdf', 'pdf', sourcePath)

    const copyPath = await upload(store, {
      sessionId: 'session-pdf',
      format: 'pdf',
      fileName: 'exported.pdf',
      data: Buffer.from('copy'),
      mode: 'export-copy',
    })
    const nextSavePath = await upload(store, {
      sessionId: 'session-pdf',
      format: 'pdf',
      fileName: 'source.pdf',
      data: Buffer.from('updated source'),
      mode: 'save',
    })

    expect(copyPath).not.toBe(sourcePath)
    expect(nextSavePath).toBe(sourcePath)
    await expect(readFile(copyPath, 'utf8')).resolves.toBe('copy')
    await expect(readFile(sourcePath, 'utf8')).resolves.toBe('updated source')
  })

  it('rejects traversal, mismatched formats, cross-session uploads, and out-of-order chunks', async () => {
    const root = await temporaryDirectory('tandemfolio-document-save-')
    const store = new DocumentSaveStore(join(root, 'outputs'), join(root, 'bindings'))

    await expect(store.begin('session-a', 'pdf', '../escape.pdf', 3, 'save')).rejects.toMatchObject(
      {
        code: 'invalid_arguments',
      },
    )
    await expect(store.begin('session-a', 'pdf', 'wrong.docx', 3, 'save')).rejects.toMatchObject({
      code: 'invalid_arguments',
    })

    const begun = await store.begin('session-a', 'xlsx', 'book.xlsx', 3, 'save')
    await expect(
      store.write('session-b', begun.uploadId, 0, Buffer.from('abc').toString('base64')),
    ).rejects.toMatchObject({ code: 'command_not_found' })
    await expect(
      store.write('session-a', begun.uploadId, 1, Buffer.from('abc').toString('base64')),
    ).rejects.toMatchObject({ code: 'revision_conflict' })
    await store.abort('session-a', begun.uploadId)
  })
})
