import assert from 'node:assert/strict'
import { cp, mkdtemp, readFile, rm } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { Client } from '@modelcontextprotocol/sdk/client/index.js'
import { StdioClientTransport } from '@modelcontextprotocol/sdk/client/stdio.js'

const expectedTools = [
  'office_create_session',
  'office_show_editor',
  'office_show_xlsx_editor',
  'office_show_pptx_editor',
  'office_show_pdf_editor',
  'office_show_markdown_editor',
  'office_get_capabilities',
  'office_get_context',
  'office_execute',
  'office_open_local_file',
  'office_editor_poll',
  'office_editor_begin_document_save',
  'office_editor_write_document_save_chunk',
  'office_editor_commit_document_save',
  'office_editor_abort_document_save',
  'office_editor_begin_recovery',
  'office_editor_write_recovery_chunk',
  'office_editor_commit_recovery',
  'office_editor_disconnect',
  'office_editor_read_local_asset_chunk',
  'office_editor_read_file_chunk',
  'office_editor_read_font_chunk',
  'office_editor_acknowledge',
]
const resourceUri = 'ui://tandemfolio/editor.html'
const markdownResourceUri = 'ui://tandemfolio/markdown.html'
const client = new Client({ name: 'tandemfolio-smoke', version: '0.1.0' })
const stagingRoot = await mkdtemp(join(tmpdir(), 'tandemfolio-plugin-smoke-'))
const installedPluginRoot = join(stagingRoot, 'tandemfolio')
const pluginSourceRoot =
  process.env.TANDEMFOLIO_PLUGIN_ROOT ?? join(process.cwd(), 'plugins/tandemfolio')
await cp(pluginSourceRoot, installedPluginRoot, { recursive: true })
const environment = Object.fromEntries(
  Object.entries(process.env).filter((entry) => entry[1] !== undefined),
)
const transport = new StdioClientTransport({
  command: 'node',
  args: ['./dist/server.js'],
  cwd: installedPluginRoot,
  env: {
    ...environment,
    TANDEMFOLIO_STATE_DIR: join(stagingRoot, 'state'),
    TANDEMFOLIO_OUTPUT_DIR: join(stagingRoot, 'outputs'),
  },
})

try {
  await client.connect(transport)
  const tools = await client.listTools()
  assert.deepEqual(
    tools.tools.map((tool) => tool.name),
    expectedTools,
  )
  const executeInputSchema = tools.tools.find((tool) => tool.name === 'office_execute')?.inputSchema
  assert.ok(executeInputSchema?.properties)
  assert.deepEqual(
    [...(executeInputSchema.required ?? [])].sort(),
    ['baseRevision', 'operations', 'requestId', 'sessionId'].sort(),
  )
  assert.equal(Object.hasOwn(executeInputSchema.properties, 'operation'), false)
  assert.equal(Object.hasOwn(executeInputSchema.properties, 'arguments'), false)

  const resources = await client.listResources()
  assert.ok(resources.resources.some((resource) => resource.uri === resourceUri))
  assert.ok(resources.resources.some((resource) => resource.uri === markdownResourceUri))

  const capabilitySummary = await client.callTool({
    name: 'office_get_capabilities',
    arguments: { format: 'xlsx' },
  })
  const summaryCapabilities = capabilitySummary.structuredContent?.capabilities
  assert.equal(summaryCapabilities?.discovery?.view, 'summary')
  assert.equal(summaryCapabilities?.operations, undefined)
  assert.ok(summaryCapabilities?.discovery?.operations?.length <= 20)
  assert.ok(Buffer.byteLength(JSON.stringify(capabilitySummary.structuredContent)) <= 8_192)

  const firstOperationId = summaryCapabilities?.discovery?.operations?.[0]?.id
  assert.equal(typeof firstOperationId, 'string')
  const capabilityDetail = await client.callTool({
    name: 'office_get_capabilities',
    arguments: { format: 'xlsx', view: 'detail', operation: firstOperationId },
  })
  const detailOperation = capabilityDetail.structuredContent?.capabilities?.discovery?.operation
  assert.equal(detailOperation?.id, firstOperationId)
  assert.equal(detailOperation?.visibility, 'agent')
  assert.equal(detailOperation?.inputSchema?.type, 'object')
  assert.deepEqual(detailOperation?.compatibilityAliases, [])

  const created = await client.callTool({
    name: 'office_create_session',
    arguments: { format: 'markdown', resume: 'none' },
  })
  const sessionId = created.structuredContent?.session?.id
  assert.equal(typeof sessionId, 'string')
  const shown = await client.callTool({
    name: 'office_show_markdown_editor',
    arguments: { sessionId },
  })
  const viewId = shown.structuredContent?.viewId
  assert.equal(typeof viewId, 'string')
  await client.callTool({
    name: 'office_editor_poll',
    arguments: { sessionId, viewId, selection: { from: 1, to: 1 } },
  })
  const duplicateShown = await client.callTool({
    name: 'office_show_markdown_editor',
    arguments: { sessionId },
  })
  const duplicateView = await client.callTool({
    name: 'office_editor_poll',
    arguments: { sessionId, viewId: duplicateShown.structuredContent?.viewId },
  })
  assert.equal(duplicateView.isError, true)
  assert.equal(duplicateView.structuredContent?.error, 'editor_view_conflict')

  const savedBytes = Buffer.from('# Packaged local save')
  const begunSave = await client.callTool({
    name: 'office_editor_begin_document_save',
    arguments: {
      sessionId,
      viewId,
      fileName: 'packaged-smoke.md',
      size: savedBytes.length,
      mode: 'save',
    },
  })
  const saveUploadId = begunSave.structuredContent?.uploadId
  assert.equal(typeof saveUploadId, 'string')
  const writtenSave = await client.callTool({
    name: 'office_editor_write_document_save_chunk',
    arguments: {
      sessionId,
      viewId,
      uploadId: saveUploadId,
      offset: 0,
      data: savedBytes.toString('base64'),
    },
  })
  assert.equal(writtenSave.structuredContent?.nextOffset, savedBytes.length)
  const committedSave = await client.callTool({
    name: 'office_editor_commit_document_save',
    arguments: { sessionId, viewId, uploadId: saveUploadId },
  })
  const savedPath = committedSave.structuredContent?.path
  assert.equal(typeof savedPath, 'string')
  assert.deepEqual(await readFile(savedPath), savedBytes)
  const savedContext = await client.callTool({
    name: 'office_get_context',
    arguments: { sessionId },
  })
  assert.equal(savedContext.structuredContent?.session?.filePath, savedPath)

  const retiredAlias = await client.callTool({
    name: 'office_execute',
    arguments: {
      sessionId,
      baseRevision: 0,
      requestId: 'packaged-smoke-retired-alias',
      operations: [{ id: 'insert_text', arguments: { text: 'legacy' } }],
    },
  })
  assert.equal(retiredAlias.isError, true)
  assert.equal(retiredAlias.structuredContent?.error, 'operation_not_found')
  const afterRetiredAlias = await client.callTool({
    name: 'office_editor_poll',
    arguments: { sessionId, viewId, waitMs: 0 },
  })
  assert.deepEqual(afterRetiredAlias.structuredContent?.commands, [])

  const transaction = {
    sessionId,
    baseRevision: 0,
    requestId: 'packaged-smoke-transaction',
    operations: [{ id: 'markdown.text.insert', arguments: { text: 'smoke' } }],
  }
  const executionPromise = client.callTool({ name: 'office_execute', arguments: transaction })
  const polled = await client.callTool({
    name: 'office_editor_poll',
    arguments: { sessionId, viewId, waitMs: 1_000 },
  })
  const command = polled.structuredContent?.commands?.[0]
  assert.equal(command?.operation, 'markdown.text.insert')
  assert.deepEqual(command?.arguments, { text: 'smoke' })
  await client.callTool({
    name: 'office_editor_acknowledge',
    arguments: {
      sessionId,
      viewId,
      commandId: command.commandId,
      revision: 1,
      dirty: true,
      output: { inserted: true },
    },
  })
  const execution = await executionPromise
  assert.equal(execution.structuredContent?.ok, true)
  assert.equal(execution.structuredContent?.transaction?.requestId, transaction.requestId)
  assert.equal(typeof execution.structuredContent?.transaction?.transactionId, 'string')
  assert.equal(execution.structuredContent?.result?.revision, 1)
  assert.deepEqual(execution.structuredContent?.result?.operations, [
    { id: 'markdown.text.insert', result: { inserted: true } },
  ])

  const replay = await client.callTool({ name: 'office_execute', arguments: transaction })
  assert.deepEqual(replay.structuredContent, execution.structuredContent)
  const afterReplay = await client.callTool({
    name: 'office_editor_poll',
    arguments: { sessionId, viewId, waitMs: 0 },
  })
  assert.deepEqual(afterReplay.structuredContent?.commands, [])

  const resource = await client.readResource({ uri: resourceUri })
  const content = resource.contents[0]
  assert.equal(content?.mimeType, 'text/html;profile=mcp-app')
  assert.equal(typeof content?.text, 'string')
  assert.ok((content?.text?.length ?? 0) > 100_000)
  assert.doesNotMatch(content?.text ?? '', /Run npm run build to package the editor/)
  assert.equal(content?.text?.match(/<\/script\s*>/gi)?.length, 1)

  process.stdout.write(
    `MCP smoke passed: ${tools.tools.length} tools, ${content?.text?.length ?? 0} HTML chars\n`,
  )
} finally {
  await client.close()
  await rm(stagingRoot, { recursive: true, force: true })
}
