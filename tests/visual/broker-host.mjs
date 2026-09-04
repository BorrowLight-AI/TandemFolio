// Isolated real-broker boundary for packaged browser regression tests.
import { mkdtemp, rm } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { Client } from '@modelcontextprotocol/sdk/client/index.js'
import { StdioClientTransport } from '@modelcontextprotocol/sdk/client/stdio.js'

const brokers = new Map()
async function connect(directory) {
  const client = new Client({ name: 'browser-regression', version: '1' })
  await client.connect(
    new StdioClientTransport({
      command: process.execPath,
      args: ['node_modules/tsx/dist/cli.mjs', 'apps/mcp-server/src/server.ts'],
      env: {
        ...process.env,
        TANDEMFOLIO_STATE_DIR: join(directory, 'state'),
        TANDEMFOLIO_OUTPUT_DIR: join(directory, 'outputs'),
      },
    }),
  )
  return client
}
export async function brokerRequest(id, action, request) {
  if (!/^[a-zA-Z0-9-]{1,80}$/.test(id)) throw new Error('Invalid test broker id')
  let broker = brokers.get(id)
  if (!broker) {
    const directory = await mkdtemp(join(tmpdir(), 'tandemfolio-browser-test-'))
    broker = { directory, client: await connect(directory) }
    brokers.set(id, broker)
  }
  if (action === 'close') {
    await broker.client.close()
    await rm(broker.directory, { recursive: true, force: true })
    brokers.delete(id)
    return { ok: true }
  }
  if (action === 'restart') {
    await broker.client.close()
    broker.client = await connect(broker.directory)
    return { ok: true }
  }
  return broker.client.callTool(request)
}
export async function closeBrokers() {
  await Promise.all([...brokers.keys()].map((id) => brokerRequest(id, 'close')))
}
