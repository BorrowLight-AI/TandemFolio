import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { createHash } from 'node:crypto'
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

  it('commits bounded Markdown companion assets beside the document', async () => {
    const root = await temporaryDirectory('tandemfolio-document-save-')
    const store = new DocumentSaveStore(join(root, 'outputs'), join(root, 'bindings'))
    const markdown = Buffer.from('![chart](assets/chart.png)')
    const image = Buffer.from([0x89, 0x50, 0x4e, 0x47])
    const begun = await store.begin(
      'session-assets',
      'markdown',
      'notes.md',
      markdown.length,
      'save',
      [{ relativePath: 'assets/chart.png', size: image.length }],
    )

    await store.write(
      'session-assets',
      begun.uploadId,
      0,
      image.toString('base64'),
      'assets/chart.png',
    )
    await store.write('session-assets', begun.uploadId, 0, markdown.toString('base64'))
    const saved = await store.commit('session-assets', begun.uploadId)

    await expect(readFile(saved.path)).resolves.toEqual(markdown)
    await expect(readFile(join(saved.path, '..', 'assets/chart.png'))).resolves.toEqual(image)
  })

  it('rejects companion traversal before creating save files', async () => {
    const root = await temporaryDirectory('tandemfolio-document-save-')
    const store = new DocumentSaveStore(join(root, 'outputs'), join(root, 'bindings'))

    await expect(
      store.begin('session-assets', 'markdown', 'notes.md', 1, 'save', [
        { relativePath: '../escape.png', size: 1 },
      ]),
    ).rejects.toMatchObject({ code: 'invalid_arguments' })
  })

  it('collects only content-addressed companion files owned by the live session', async () => {
    const root = await temporaryDirectory('tandemfolio-document-save-')
    const store = new DocumentSaveStore(join(root, 'outputs'), join(root, 'bindings'))
    const image = Buffer.from('owned image')
    const digest = createHash('sha256').update(image).digest('hex').slice(0, 12)
    const relativePath = `assets/image-${digest}.png`
    const markdown = Buffer.from(`![x](${relativePath})`)
    const first = await store.begin('session-gc', 'markdown', 'notes.md', markdown.length, 'save', [
      { relativePath, size: image.length },
    ])
    await store.write('session-gc', first.uploadId, 0, image.toString('base64'), relativePath)
    await store.write('session-gc', first.uploadId, 0, markdown.toString('base64'))
    const saved = await store.commit('session-gc', first.uploadId)

    const next = Buffer.from('# image removed')
    const second = await store.begin(
      'session-gc',
      'markdown',
      'notes.md',
      next.length,
      'save',
      [],
      [relativePath],
    )
    await store.write('session-gc', second.uploadId, 0, next.toString('base64'))
    await store.commit('session-gc', second.uploadId)

    await expect(readFile(saved.path, 'utf8')).resolves.toBe('# image removed')
    await expect(readFile(join(saved.path, '..', relativePath))).rejects.toMatchObject({
      code: 'ENOENT',
    })
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
