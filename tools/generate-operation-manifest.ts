import { mkdir, readFile, writeFile } from 'node:fs/promises'
import { dirname, extname, resolve } from 'node:path'
import { pathToFileURL } from 'node:url'

import {
  createOperationManifest,
  serializeOperationManifest,
  type OperationCatalog,
} from '@tandemfolio/operation-contract'

interface CliOptions {
  readonly check: boolean
  readonly inputPath: string
  readonly outputPath: string
}

function parseOptions(arguments_: readonly string[]): CliOptions {
  const inputIndex = arguments_.indexOf('--input')
  const outputIndex = arguments_.indexOf('--output')
  const inputPath = inputIndex === -1 ? undefined : arguments_[inputIndex + 1]
  const outputPath = outputIndex === -1 ? undefined : arguments_[outputIndex + 1]

  if (!inputPath || !outputPath) {
    throw new Error(
      'Usage: generate-operation-manifest --input <catalogs.json|catalog.ts> --output <manifest.json>',
    )
  }

  return {
    check: arguments_.includes('--check'),
    inputPath: resolve(inputPath),
    outputPath: resolve(outputPath),
  }
}

async function loadCatalogs(inputPath: string): Promise<OperationCatalog[]> {
  if (extname(inputPath) === '.json') {
    return JSON.parse(await readFile(inputPath, 'utf8')) as OperationCatalog[]
  }

  const module = (await import(pathToFileURL(inputPath).href)) as {
    default?: OperationCatalog | OperationCatalog[]
  }
  if (!module.default) {
    throw new Error(`Operation catalog module must have a default export: ${inputPath}`)
  }
  return Array.isArray(module.default) ? module.default : [module.default]
}

async function main(): Promise<void> {
  const options = parseOptions(process.argv.slice(2))
  const catalogs = await loadCatalogs(options.inputPath)
  const serialized = serializeOperationManifest(createOperationManifest(catalogs))

  if (options.check) {
    const current = await readFile(options.outputPath, 'utf8').catch(() => null)
    if (current !== serialized) {
      throw new Error(`Operation manifest is out of date: ${options.outputPath}`)
    }
    return
  }

  await mkdir(dirname(options.outputPath), { recursive: true })
  await writeFile(options.outputPath, serialized, 'utf8')
}

main().catch((error: unknown) => {
  const message = error instanceof Error ? error.message : String(error)
  process.stderr.write(`${message}\n`)
  process.exitCode = 1
})
