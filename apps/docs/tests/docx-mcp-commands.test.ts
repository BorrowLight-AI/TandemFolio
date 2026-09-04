import { describe, expect, it } from 'vitest'

import { createDocxRecoverySnapshot, loadStagedDocx } from '../src/renderer/host/docx-commands'
import type { OpenFileResult } from '../src/shared/lite-api'

describe('DOCX MCP commands', () => {
  it('loads reconstructed local bytes into the mounted document', async () => {
    const data = new Uint8Array([80, 75, 3, 4, 1, 2, 3, 4]).buffer
    let opened: OpenFileResult | undefined

    await loadStagedDocx({ blobId: 'blob-1', name: 'local.docx', data }, async (file) => {
      opened = file
    })

    expect(opened).toMatchObject({
      path: 'mcp-local:blob-1/local.docx',
      name: 'local.docx',
      data,
      hash: expect.stringMatching(/^[a-f0-9]{64}$/),
    })
  })

  it('checkpoints only the DOCX bytes returned by the format save pipeline', async () => {
    const backing = new Uint8Array([99, 80, 75, 3, 4, 88])
    const bytes = backing.subarray(1, 5)

    const recovery = await createDocxRecoverySnapshot('draft.docx', async () => bytes)

    expect(recovery?.fileName).toBe('draft.docx')
    expect(Array.from(new Uint8Array(recovery!.data))).toEqual([80, 75, 3, 4])
  })
})
