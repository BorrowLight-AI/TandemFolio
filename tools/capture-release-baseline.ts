import { execFile, spawn, type ChildProcess } from 'node:child_process'
import { createHash } from 'node:crypto'
import { mkdir, readFile, writeFile } from 'node:fs/promises'
import { arch, cpus, platform, totalmem } from 'node:os'
import { dirname, join, relative, resolve } from 'node:path'
import { createRequire } from 'node:module'
import { promisify } from 'node:util'
import { fileURLToPath } from 'node:url'
import { chromium, type Browser, type BrowserContext, type Page } from '@playwright/test'
import { PDFDocument, StandardFonts } from 'pdf-lib'
import { RELEASE_EVIDENCE_SCHEMA_VERSION } from './release-gate/contract'
import { computeReleaseSourceFingerprint } from './release-gate/fingerprint'
import {
  createReleaseFixture,
  type ReleaseFixture,
  type ReleaseFixtureSize,
  type ReleaseFormat,
} from './release-gate/fixtures'
import { releaseMeasurementMatrix } from './release-gate/scenarios'
import { suggestedCeiling, summarizeSamples } from './release-gate/stats'

const require = createRequire(import.meta.url)
const execFileAsync = promisify(execFile)
const { PNG } = require('playwright-core/lib/utilsBundle') as {
  PNG: { sync: { read(value: Buffer): { width: number; height: number; data: Uint8Array } } }
}

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..')
const defaultOutput = join(root, 'release', 'release-evidence.candidate.json')
const baseUrl = 'http://127.0.0.1:4178'
const formats = [...releaseMeasurementMatrix.formats] as ReleaseFormat[]
const sizes = [...releaseMeasurementMatrix.fixtureSizes] as ReleaseFixtureSize[]
let upstreamVisualManifest: {
  upstreamCommit: string
  formats: Record<
    ReleaseFormat,
    {
      baselinePath: string
      baselineSha256: string
      maxDiffPixelRatio: number
      masks: unknown[]
    }
  >
}

const formatConfig: Record<
  ReleaseFormat,
  {
    documentSelector: string
    loadOperation: string
    interaction: (sample: number) => { operation: string; arguments: Record<string, unknown> }
  }
> = {
  docx: {
    documentSelector: '.app',
    loadOperation: 'docx.document.load_staged',
    interaction: (sample) => ({
      operation: 'docx.text.insert',
      arguments: { text: ` measured-${sample}` },
    }),
  },
  markdown: {
    documentSelector: '.doc-editor[aria-label="Markdown document"]',
    loadOperation: 'markdown.document.load_staged',
    interaction: (sample) => ({
      operation: 'markdown.text.insert',
      arguments: { text: ` measured-${sample}` },
    }),
  },
  xlsx: {
    documentSelector: '.app-shell',
    loadOperation: 'xlsx.document.load_staged',
    interaction: (sample) => ({
      operation: 'xlsx.cell.set_value',
      arguments: { sheet: 'Sheet1', address: 'A1', value: `Measured ${sample}` },
    }),
  },
  pptx: {
    documentSelector: '.app',
    loadOperation: 'pptx.document.load_staged',
    interaction: () => ({ operation: 'pptx.slide.add_blank', arguments: { afterSlideIndex: 0 } }),
  },
  pdf: {
    documentSelector: '.app',
    loadOperation: 'pdf.document.load_staged',
    interaction: (sample) => ({
      operation: 'pdf.page.set_rotation',
      arguments: { pageIndex: 0, rotation: sample % 2 === 0 ? 90 : 0 },
    }),
  },
}

interface CommandTiming {
  pollWaitMs: number
  hydrateMs: number
  executeMs: number
  ackTransportMs: number
  totalMs: number
  trace?: {
    operation: 'markdown.document.load_staged'
    phases: {
      decodeMs: number
      parseMs: number
      tiptapStateInstallMs: number
      reactCommitMs: number
    }
  }
}

interface XlsxColdStartTrace {
  operation: 'xlsx.editor.cold_start'
  phases: {
    bootstrapMs: number
    univerCreateMs: number
    worksheetInstallMs: number
    firstCommitMs: number
  }
  bootstrapPhases: {
    resourceReceiveMs: number
    moduleGraphReadyMs: number
    reactMountMs: number
  }
}

type XlsxColdStartPhase = keyof XlsxColdStartTrace['phases']
type XlsxBootstrapPhase = keyof XlsxColdStartTrace['bootstrapPhases']

interface MemoryPeaks {
  jsHeapBytes: number
  rendererRssBytes: number
}

interface SampleCounts {
  coldStart: number
  openDocument: number
  interaction: number
}

function argument(name: string): string | undefined {
  const index = process.argv.indexOf(name)
  return index === -1 ? undefined : process.argv[index + 1]
}

function sha256(value: Uint8Array): string {
  return createHash('sha256').update(value).digest('hex')
}

function rounded(value: number): number {
  return Number(value.toFixed(3))
}

async function health(): Promise<boolean> {
  try {
    return (await fetch(`${baseUrl}/health`)).ok
  } catch {
    return false
  }
}

async function waitForHealth(): Promise<void> {
  const deadline = Date.now() + 30_000
  while (Date.now() < deadline) {
    if (await health()) return
    await new Promise<void>((resolveDelay) => setTimeout(resolveDelay, 100))
  }
  throw new Error('Timed out waiting for the packaged-host benchmark server.')
}

async function startHost(): Promise<ChildProcess | null> {
  if (await health()) return null
  const child = spawn(process.execPath, ['tests/visual/serve-host.mjs'], {
    cwd: root,
    stdio: ['ignore', 'inherit', 'inherit'],
  })
  await waitForHealth()
  return child
}

async function descendantsWithRss(
  rootPid: number,
): Promise<Array<{ pid: number; rssBytes: number; command: string }>> {
  const { stdout: output } = await execFileAsync('ps', ['-axo', 'pid=,ppid=,rss=,command='], {
    encoding: 'utf8',
  })
  const rows = output
    .trim()
    .split('\n')
    .map((line) => /^\s*(\d+)\s+(\d+)\s+(\d+)\s+(.*)$/.exec(line))
    .filter((match): match is RegExpExecArray => match !== null)
    .map((match) => ({
      pid: Number(match[1]),
      ppid: Number(match[2]),
      rssBytes: Number(match[3]) * 1024,
      command: match[4],
    }))
  const descendants = new Set([rootPid])
  let changed = true
  while (changed) {
    changed = false
    for (const row of rows) {
      if (descendants.has(row.ppid) && !descendants.has(row.pid)) {
        descendants.add(row.pid)
        changed = true
      }
    }
  }
  return rows.filter((row) => descendants.has(row.pid))
}

async function rendererRss(browserPid: number): Promise<number> {
  return (await descendantsWithRss(browserPid))
    .filter((process) => process.command.includes('--type=renderer'))
    .reduce((sum, process) => sum + process.rssBytes, 0)
}

async function jsHeap(page: Page): Promise<number> {
  return page.evaluate(() => {
    const memory = (performance as Performance & { memory?: { usedJSHeapSize?: number } }).memory
    return memory?.usedJSHeapSize ?? 0
  })
}

function startMemorySampler(
  page: Page,
  browserPid: number,
  peaks: MemoryPeaks,
): () => Promise<void> {
  let active = true
  let pending: Promise<void> | null = null
  const sample = async () => {
    if (!active) return
    if (pending) return pending
    pending = (async () => {
      try {
        peaks.jsHeapBytes = Math.max(peaks.jsHeapBytes, await jsHeap(page))
        peaks.rendererRssBytes = Math.max(peaks.rendererRssBytes, await rendererRss(browserPid))
      } catch {
        // A page may close between the final sample and sampler shutdown.
      }
    })().finally(() => {
      pending = null
    })
    return pending
  }
  const timer = setInterval(() => void sample(), 25)
  void sample()
  return async () => {
    await sample()
    active = false
    clearInterval(timer)
  }
}

async function mountedPage(
  browser: Browser,
  format: ReleaseFormat,
): Promise<{
  context: BrowserContext
  page: Page
}> {
  const context = await browser.newContext({
    colorScheme: 'light',
    deviceScaleFactor: 1,
    locale: 'zh-CN',
    reducedMotion: 'reduce',
    viewport: { width: 1440, height: 1360 },
  })
  const page = await context.newPage()
  await page.goto(`${baseUrl}/?format=${format}&width=720&height=900`)
  await page.waitForFunction(() => {
    const state = window.__codexVisualHost
    return state?.initialized && state.firstPollAt != null
  })
  await page.frameLocator('#editor-frame').locator(formatConfig[format].documentSelector).waitFor()
  return { context, page }
}

function base64(bytes: Uint8Array): string {
  return Buffer.from(bytes).toString('base64')
}

async function executeCommand(
  page: Page,
  command: {
    commandId: string
    baseRevision: number
    operation: string
    arguments: Record<string, unknown>
  },
  fixture?: ReleaseFixture,
): Promise<CommandTiming> {
  await page.evaluate(
    ({ command, staged }) => {
      window.__codexVisualHost.enqueueCommand(
        command,
        staged ? { blobId: staged.blobId, base64: staged.base64 } : undefined,
      )
    },
    {
      command,
      staged: fixture
        ? { blobId: String(command.arguments.blobId), base64: base64(fixture.bytes) }
        : null,
    },
  )
  try {
    await page.waitForFunction(
      (commandId) =>
        window.__codexVisualHost.acknowledgements.some(
          (entry) => entry.commandId === commandId && entry.ok === true,
        ),
      command.commandId,
      { timeout: 120_000 },
    )
  } catch (error) {
    const diagnostics = await page
      .evaluate((commandId) => {
        const state = window.__codexVisualHost
        return {
          commandId,
          errors: state.errors,
          events: state.events.slice(-20),
          polls: state.polls,
          lastPollArguments: state.lastPollArguments,
          acknowledgement: state.acknowledgements.find((entry) => entry.commandId === commandId),
        }
      }, command.commandId)
      .catch(() => ({ commandId: command.commandId, pageClosed: true }))
    throw new Error(`Timed out waiting for ${command.commandId}: ${JSON.stringify(diagnostics)}`, {
      cause: error,
    })
  }
  return page.evaluate(
    (commandId) => window.__codexVisualHost.commandTimings[commandId] as CommandTiming,
    command.commandId,
  )
}

async function openFixture(
  page: Page,
  format: ReleaseFormat,
  fixture: ReleaseFixture,
  commandId: string,
): Promise<CommandTiming> {
  return executeCommand(
    page,
    {
      commandId,
      baseRevision: 0,
      operation: formatConfig[format].loadOperation,
      arguments: {
        blobId: `${commandId}-blob`,
        name: fixture.fileName,
        size: fixture.bytes.byteLength,
      },
    },
    fixture,
  )
}

function comparePng(
  actual: Buffer,
  expected: Buffer,
): { diffPixelRatio: number; dimensionsMatch: boolean } {
  const left = PNG.sync.read(actual)
  const right = PNG.sync.read(expected)
  if (left.width !== right.width || left.height !== right.height) {
    return { diffPixelRatio: 1, dimensionsMatch: false }
  }
  let different = 0
  const threshold = 0.2 * 255
  for (let index = 0; index < left.data.length; index += 4) {
    const delta = Math.max(
      Math.abs(left.data[index] - right.data[index]),
      Math.abs(left.data[index + 1] - right.data[index + 1]),
      Math.abs(left.data[index + 2] - right.data[index + 2]),
      Math.abs(left.data[index + 3] - right.data[index + 3]),
    )
    if (delta > threshold) different += 1
  }
  return { diffPixelRatio: different / (left.width * left.height), dimensionsMatch: true }
}

async function visualPdfFixture(): Promise<ReleaseFixture> {
  const document = await PDFDocument.create()
  const font = await document.embedFont(StandardFonts.Helvetica)
  const page = document.addPage([420, 300])
  page.drawText('TandemFolio PDF host gate', { x: 30, y: 250, size: 18, font })
  const bytes = await document.save()
  return { bytes, fileName: 'host-gate.pdf', mimeType: 'application/pdf' }
}

async function captureVisual(browser: Browser, format: ReleaseFormat, artifactRoot: string) {
  const { context, page } = await mountedPage(browser, format)
  try {
    if (format === 'pdf')
      await openFixture(page, format, await visualPdfFixture(), 'visual-pdf-open')
    await page
      .frameLocator('#editor-frame')
      .locator('html')
      .evaluate(() => document.fonts.ready)
    const actual = await page.locator('#editor-frame').screenshot({ animations: 'disabled' })
    const provenance = upstreamVisualManifest.formats[format]
    if (provenance.masks.length > 0) {
      throw new Error(`${format} visual masks require an explicit comparator implementation.`)
    }
    const baselinePath = join(root, provenance.baselinePath)
    const expected = await readFile(baselinePath)
    if (sha256(expected) !== provenance.baselineSha256) {
      throw new Error(`${format} pinned-upstream baseline hash does not match its manifest.`)
    }
    const comparison = comparePng(actual, expected)
    const actualPath = join(artifactRoot, `${format}-split-view.png`)
    await writeFile(actualPath, actual)
    const maxDiffPixelRatio = provenance.maxDiffPixelRatio
    return {
      passed: comparison.dimensionsMatch && comparison.diffPixelRatio <= maxDiffPixelRatio,
      upstreamCommit: upstreamVisualManifest.upstreamCommit,
      baselinePath: relative(root, baselinePath),
      actualPath: relative(root, actualPath),
      baselineSha256: sha256(expected),
      actualSha256: sha256(actual),
      dimensionsMatch: comparison.dimensionsMatch,
      diffPixelRatio: rounded(comparison.diffPixelRatio),
      maxDiffPixelRatio,
    }
  } finally {
    await context.close()
  }
}

function metricWithBudget(samples: number[], fixedCeiling?: number) {
  const measured = summarizeSamples(samples)
  return { measured, ceiling: fixedCeiling ?? suggestedCeiling(measured) }
}

function measurementsWithinBudgets(value: unknown): boolean {
  if (typeof value !== 'object' || value === null || Array.isArray(value)) return true
  const record = value as Record<string, unknown>
  if (typeof record.ceiling === 'number') {
    const measured =
      typeof record.measured === 'number'
        ? record.measured
        : typeof record.measured === 'object' &&
            record.measured !== null &&
            typeof (record.measured as { p95?: unknown }).p95 === 'number'
          ? (record.measured as { p95: number }).p95
          : undefined
    if (measured !== undefined && measured > record.ceiling) return false
  }
  return Object.values(record).every(measurementsWithinBudgets)
}

function timingBuckets() {
  return {
    pollWaitMs: [] as number[],
    hydrateMs: [] as number[],
    executeMs: [] as number[],
    ackTransportMs: [] as number[],
    totalMs: [] as number[],
  }
}

function recordTiming(buckets: ReturnType<typeof timingBuckets>, timing: CommandTiming): void {
  for (const key of Object.keys(buckets) as Array<keyof typeof buckets>) {
    buckets[key].push(Math.max(0, timing[key]))
  }
}

function stagedLoadPhaseBuckets() {
  return {
    small: {
      decodeMs: [] as number[],
      parseMs: [] as number[],
      tiptapStateInstallMs: [] as number[],
      reactCommitMs: [] as number[],
    },
    medium: {
      decodeMs: [] as number[],
      parseMs: [] as number[],
      tiptapStateInstallMs: [] as number[],
      reactCommitMs: [] as number[],
    },
    large: {
      decodeMs: [] as number[],
      parseMs: [] as number[],
      tiptapStateInstallMs: [] as number[],
      reactCommitMs: [] as number[],
    },
  }
}

function recordMarkdownStagedLoadTrace(
  buckets: ReturnType<typeof stagedLoadPhaseBuckets>,
  size: ReleaseFixtureSize,
  timing: CommandTiming,
): void {
  if (timing.trace?.operation !== 'markdown.document.load_staged') {
    throw new Error(`Markdown ${size} staged load did not publish its bounded phase trace.`)
  }
  for (const phase of Object.keys(buckets[size]) as Array<keyof (typeof buckets)[typeof size]>) {
    buckets[size][phase].push(Math.max(0, timing.trace.phases[phase]))
  }
}

function xlsxColdStartPhaseBuckets(): Record<XlsxColdStartPhase, number[]> {
  return {
    bootstrapMs: [],
    univerCreateMs: [],
    worksheetInstallMs: [],
    firstCommitMs: [],
  }
}

function xlsxBootstrapPhaseBuckets(): Record<XlsxBootstrapPhase, number[]> {
  return {
    resourceReceiveMs: [],
    moduleGraphReadyMs: [],
    reactMountMs: [],
  }
}

function recordXlsxColdStartTrace(
  buckets: ReturnType<typeof xlsxColdStartPhaseBuckets>,
  bootstrapBuckets: ReturnType<typeof xlsxBootstrapPhaseBuckets>,
  trace: XlsxColdStartTrace | undefined,
  hostColdStartMs: number,
): void {
  if (trace?.operation !== 'xlsx.editor.cold_start') {
    throw new Error('XLSX cold start did not publish its bounded phase trace.')
  }
  const phases = Object.keys(buckets) as XlsxColdStartPhase[]
  const tracedTotal = phases.reduce((sum, phase) => sum + trace.phases[phase], 0)
  if (tracedTotal > hostColdStartMs + 0.01) {
    throw new Error(
      `XLSX cold-start phases (${tracedTotal} ms) exceed the host measurement (${hostColdStartMs} ms).`,
    )
  }
  for (const phase of phases) buckets[phase].push(Math.max(0, trace.phases[phase]))
  const bootstrapPhases = Object.keys(bootstrapBuckets) as XlsxBootstrapPhase[]
  const bootstrapTotal = bootstrapPhases.reduce(
    (sum, phase) => sum + trace.bootstrapPhases[phase],
    0,
  )
  if (Math.abs(bootstrapTotal - trace.phases.bootstrapMs) > 0.01) {
    throw new Error(
      `XLSX bootstrap phases (${bootstrapTotal} ms) do not partition bootstrap (${trace.phases.bootstrapMs} ms).`,
    )
  }
  for (const phase of bootstrapPhases) {
    bootstrapBuckets[phase].push(Math.max(0, trace.bootstrapPhases[phase]))
  }
}

const markdownOpenDocumentCeilings: Record<ReleaseFixtureSize, number> = {
  small: 20,
  medium: 835.68,
  large: 5_000,
}

async function captureFormat(
  browser: Browser,
  browserPid: number,
  format: ReleaseFormat,
  counts: SampleCounts,
  artifactRoot: string,
) {
  process.stderr.write(`Measuring ${format}...\n`)
  const fixtures = Object.fromEntries(
    await Promise.all(
      sizes.map(async (size) => [size, await createReleaseFixture(format, size)] as const),
    ),
  ) as Record<ReleaseFixtureSize, ReleaseFixture>
  const peaks: MemoryPeaks = { jsHeapBytes: 0, rendererRssBytes: 0 }
  const coldStart: number[] = []
  const openDocument = { small: [] as number[], medium: [] as number[], large: [] as number[] }
  const interaction: number[] = []
  const ack = timingBuckets()
  const stagedLoadPhases = stagedLoadPhaseBuckets()
  const coldStartPhases = xlsxColdStartPhaseBuckets()
  const bootstrapPhases = xlsxBootstrapPhaseBuckets()
  const visual = await captureVisual(browser, format, artifactRoot)

  for (let sample = 0; sample < counts.coldStart; sample += 1) {
    const context = await browser.newContext({
      colorScheme: 'light',
      deviceScaleFactor: 1,
      locale: 'zh-CN',
      reducedMotion: 'reduce',
      viewport: { width: 1440, height: 1360 },
    })
    const page = await context.newPage()
    const stopMemory = startMemorySampler(page, browserPid, peaks)
    try {
      const startedAt = performance.now()
      await page.goto(`${baseUrl}/?format=${format}&width=720&height=900`)
      await page.waitForFunction(() => {
        const state = window.__codexVisualHost
        return state?.initialized && state.firstPollAt != null
      })
      await page
        .frameLocator('#editor-frame')
        .locator(formatConfig[format].documentSelector)
        .waitFor()
      if (format === 'xlsx') {
        const startup = await page.evaluate(() => {
          const state = window.__codexVisualHost
          return {
            hostColdStartMs: state.firstPollAt! - state.startedAt,
            trace: state.firstPollArguments?.startupTrace,
          }
        })
        recordXlsxColdStartTrace(
          coldStartPhases,
          bootstrapPhases,
          startup.trace as XlsxColdStartTrace | undefined,
          startup.hostColdStartMs,
        )
      }
      coldStart.push(performance.now() - startedAt)
    } finally {
      await stopMemory()
      await context.close()
    }
  }

  for (const size of sizes) {
    for (let sample = 0; sample < counts.openDocument; sample += 1) {
      const { context, page } = await mountedPage(browser, format)
      const stopMemory = startMemorySampler(page, browserPid, peaks)
      try {
        const timing = await openFixture(
          page,
          format,
          fixtures[size],
          `${format}-${size}-${sample}`,
        )
        openDocument[size].push(timing.totalMs)
        recordTiming(ack, timing)
        if (format === 'markdown') recordMarkdownStagedLoadTrace(stagedLoadPhases, size, timing)
      } finally {
        await stopMemory()
        await context.close()
      }
    }
  }

  {
    const { context, page } = await mountedPage(browser, format)
    const stopMemory = startMemorySampler(page, browserPid, peaks)
    try {
      const opening = await openFixture(page, format, fixtures.small, `${format}-interaction-open`)
      recordTiming(ack, opening)
      for (let sample = 0; sample < counts.interaction; sample += 1) {
        const request = formatConfig[format].interaction(sample)
        const timing = await executeCommand(page, {
          commandId: `${format}-interaction-${sample}`,
          baseRevision: sample + 1,
          operation: request.operation,
          arguments: request.arguments,
        })
        interaction.push(timing.totalMs)
        recordTiming(ack, timing)
      }
    } finally {
      await stopMemory()
      await context.close()
    }
  }

  const measurements = {
    fixtures: Object.fromEntries(
      sizes.map((size) => [
        size,
        { bytes: fixtures[size].bytes.byteLength, sha256: sha256(fixtures[size].bytes) },
      ]),
    ),
    visual,
    coldStartMs: metricWithBudget(coldStart, format === 'xlsx' ? 1_400 : undefined),
    openDocumentMs: Object.fromEntries(
      sizes.map((size) => [
        size,
        metricWithBudget(
          openDocument[size],
          format === 'markdown' ? markdownOpenDocumentCeilings[size] : undefined,
        ),
      ]),
    ),
    interactionMs: metricWithBudget(interaction),
    acknowledgement: Object.fromEntries(
      Object.entries(ack).map(([key, values]) => [key, metricWithBudget(values)]),
    ),
    ...(format === 'markdown'
      ? {
          stagedLoadPhases: Object.fromEntries(
            sizes.map((size) => [
              size,
              Object.fromEntries(
                Object.entries(stagedLoadPhases[size]).map(([phase, values]) => [
                  phase,
                  metricWithBudget(values),
                ]),
              ),
            ]),
          ),
        }
      : {}),
    ...(format === 'xlsx'
      ? {
          coldStartPhases: Object.fromEntries(
            Object.entries(coldStartPhases).map(([phase, values]) => [
              phase,
              metricWithBudget(values, phase === 'bootstrapMs' ? 500 : undefined),
            ]),
          ),
          bootstrapPhases: Object.fromEntries(
            Object.entries(bootstrapPhases).map(([phase, values]) => [
              phase,
              metricWithBudget(values),
            ]),
          ),
        }
      : {}),
    memory: {
      peakJsHeapBytes: {
        measured: peaks.jsHeapBytes,
        ceiling: Math.ceil(peaks.jsHeapBytes * 1.2),
      },
      peakRendererMemoryBytes: {
        measured: peaks.rendererRssBytes,
        ceiling: Math.ceil(peaks.rendererRssBytes * 1.2),
      },
    },
  }
  return {
    passed:
      visual.passed &&
      peaks.jsHeapBytes > 0 &&
      peaks.rendererRssBytes > 0 &&
      measurementsWithinBudgets(measurements),
    ...measurements,
  }
}

async function main(): Promise<void> {
  if (process.argv.includes('--describe')) {
    process.stdout.write(`${JSON.stringify(releaseMeasurementMatrix)}\n`)
    return
  }

  const output = resolve(argument('--output') ?? defaultOutput)
  const approved = process.argv.includes('--approve')
  const sampleOverride = Number(argument('--samples'))
  const counts: SampleCounts =
    Number.isInteger(sampleOverride) && sampleOverride > 0
      ? { coldStart: sampleOverride, openDocument: sampleOverride, interaction: sampleOverride }
      : { ...releaseMeasurementMatrix.defaultSamples }
  const requested = argument('--format')
  const selectedFormats = requested ? formats.filter((format) => format === requested) : formats
  if (selectedFormats.length === 0) throw new Error(`Unknown release format: ${requested}`)
  const canonicalSamples = Object.entries(releaseMeasurementMatrix.defaultSamples).every(
    ([metric, minimum]) => counts[metric as keyof SampleCounts] >= minimum,
  )
  if (approved && (!canonicalSamples || selectedFormats.length !== formats.length)) {
    throw new Error('Approved release evidence requires every format and at least 7/7/21 samples.')
  }
  upstreamVisualManifest = JSON.parse(
    await readFile(join(root, 'release/upstream-visual-manifest.json'), 'utf8'),
  )
  const artifactRoot = join(dirname(output), 'artifacts')
  await mkdir(artifactRoot, { recursive: true })

  const host = await startHost()
  try {
    const measuredFormats: Record<string, unknown> = {}
    let browserVersion = ''
    for (const format of selectedFormats) {
      const browserServer = await chromium.launchServer({
        headless: true,
        args: ['--enable-precise-memory-info'],
      })
      const browser = await chromium.connect(browserServer.wsEndpoint())
      browserVersion ||= browser.version()
      try {
        measuredFormats[format] = await captureFormat(
          browser,
          browserServer.process().pid,
          format,
          counts,
          artifactRoot,
        )
      } finally {
        await browser.close()
        await browserServer.close()
      }
    }
    const allFormatsPresent = selectedFormats.length === formats.length
    const allPassed =
      allFormatsPresent &&
      Object.values(measuredFormats).every((value) => (value as { passed: boolean }).passed)
    const upstream = JSON.parse(await readFile(join(root, 'upstream.config.json'), 'utf8')) as {
      repository: string
      baseline: string
    }
    const evidence = {
      schemaVersion: RELEASE_EVIDENCE_SCHEMA_VERSION,
      approved: approved && allPassed && canonicalSamples,
      ready: approved && allPassed && canonicalSamples,
      capturedAt: new Date().toISOString(),
      sourceFingerprint: computeReleaseSourceFingerprint(root),
      upstream: { repository: upstream.repository, commit: upstream.baseline },
      profile: {
        id: `${platform()}-${arch()}-chromium`,
        platform: platform(),
        arch: arch(),
        cpu: cpus()[0]?.model ?? 'unknown',
        logicalCpuCount: cpus().length,
        totalMemoryBytes: totalmem(),
        browser: 'chromium',
        browserVersion,
        viewport: { width: 1440, height: 1360, editorWidth: 720, editorHeight: 900 },
        samples: counts,
      },
      formats: measuredFormats,
    }
    await writeFile(output, `${JSON.stringify(evidence, null, 2)}\n`)
    process.stdout.write(
      `${JSON.stringify({ ok: true, output, approved: evidence.approved, ready: evidence.ready })}\n`,
    )
  } finally {
    if (host) host.kill('SIGTERM')
  }
}

void main().catch((error) => {
  process.stderr.write(
    `${error instanceof Error ? (error.stack ?? error.message) : String(error)}\n`,
  )
  process.exitCode = 1
})
