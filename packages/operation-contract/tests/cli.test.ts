import { execFile } from 'node:child_process'
import { mkdtemp, readFile, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { dirname, join } from 'node:path'
import { fileURLToPath, pathToFileURL } from 'node:url'
import { promisify } from 'node:util'

import { describe, expect, it } from 'vitest'

const execFileAsync = promisify(execFile)
const packageRoot = dirname(dirname(fileURLToPath(import.meta.url)))
const workspaceRoot = dirname(dirname(packageRoot))
const generatorPath = join(workspaceRoot, 'tools/generate-operation-manifest.ts')
const tsxPath = join(workspaceRoot, 'node_modules/tsx/dist/cli.mjs')
const fixturePath = join(packageRoot, 'tests/fixtures/catalogs.json')
const markdownCatalogPath = join(workspaceRoot, 'apps/markdown/src/renderer/operations/catalog.ts')
const productCatalogsPath = join(workspaceRoot, 'tools/operation-catalogs.ts')

describe('generate-operation-manifest CLI', () => {
  it('writes a generated manifest from serializable catalogs', async () => {
    const outputDirectory = await mkdtemp(join(tmpdir(), 'genoffice-operation-manifest-'))
    const outputPath = join(outputDirectory, 'operation-manifest.json')

    await execFileAsync(process.execPath, [
      tsxPath,
      generatorPath,
      '--input',
      fixturePath,
      '--output',
      outputPath,
    ])

    const manifest = JSON.parse(await readFile(outputPath, 'utf8')) as {
      schemaVersion: number
      operations: Array<{ id: string }>
    }
    expect(manifest).toMatchObject({
      schemaVersion: 1,
      operations: [{ id: 'markdown.fixture.insert_text' }],
    })
  })

  it('fails --check when the committed manifest has drifted', async () => {
    const outputDirectory = await mkdtemp(join(tmpdir(), 'genoffice-operation-manifest-'))
    const outputPath = join(outputDirectory, 'operation-manifest.json')
    await writeFile(outputPath, '{}\n', 'utf8')

    await expect(
      execFileAsync(process.execPath, [
        tsxPath,
        generatorPath,
        '--input',
        fixturePath,
        '--output',
        outputPath,
        '--check',
      ]),
    ).rejects.toMatchObject({
      stderr: `Operation manifest is out of date: ${outputPath}\n`,
    })
  })

  it('generates from a format-owned TypeScript catalog module', async () => {
    const outputDirectory = await mkdtemp(join(tmpdir(), 'genoffice-operation-manifest-'))
    const outputPath = join(outputDirectory, 'operation-manifest.json')

    await execFileAsync(process.execPath, [
      tsxPath,
      generatorPath,
      '--input',
      markdownCatalogPath,
      '--output',
      outputPath,
    ])

    const manifest = JSON.parse(await readFile(outputPath, 'utf8')) as {
      operations: Array<{ id: string; compatibilityAliases?: string[] }>
    }
    expect(manifest.operations).toHaveLength(22)
    expect(manifest.operations).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          id: 'markdown.document.load_staged',
          visibility: 'internal',
          compatibilityAliases: [],
        }),
        expect.objectContaining({
          id: 'markdown.document.save',
          compatibilityAliases: [],
        }),
        expect.objectContaining({
          id: 'markdown.selection.set',
          compatibilityAliases: [],
        }),
        expect.objectContaining({
          id: 'markdown.text.insert',
          compatibilityAliases: [],
        }),
        expect.objectContaining({
          id: 'markdown.text.replace_selection',
          compatibilityAliases: [],
        }),
      ]),
    )
    expect(manifest.operations.map((operation) => operation.id)).toEqual(
      [...manifest.operations.map((operation) => operation.id)].sort(),
    )
  })

  it('generates one product manifest from multiple format-owned catalogs', async () => {
    const outputDirectory = await mkdtemp(join(tmpdir(), 'genoffice-operation-manifest-'))
    const outputPath = join(outputDirectory, 'operation-manifest.json')

    await execFileAsync(process.execPath, [
      tsxPath,
      generatorPath,
      '--input',
      productCatalogsPath,
      '--output',
      outputPath,
    ])

    const manifest = JSON.parse(await readFile(outputPath, 'utf8')) as {
      operations: Array<{ id: string }>
    }
    const productCatalogs = (await import(pathToFileURL(productCatalogsPath).href))
      .default as Array<{
      operations: Array<{ id: string }>
    }>
    const expectedIds = productCatalogs
      .flatMap((catalog) => catalog.operations.map((operation) => operation.id))
      .sort((left, right) => left.localeCompare(right))
    expect(manifest.operations.map((operation) => operation.id)).toEqual(expectedIds)
  })
})
