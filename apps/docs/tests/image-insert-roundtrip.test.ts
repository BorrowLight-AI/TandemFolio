import { Editor } from '@tiptap/core'
import { buildBlankDocx, parseDocx, saveDocx } from '@genoffice/docx-engine'
import { describe, expect, it } from 'vitest'

import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { executeDocxOperation } from '../src/renderer/operations/registry'

describe('DOCX staged image insertion round-trip', () => {
  it('persists registry-inserted bytes and reopens as a clean native image block', async () => {
    const parsed = await parseDocx(await buildBlankDocx())
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks) as never,
    })
    const data = Uint8Array.from(
      Buffer.from('R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7', 'base64'),
    ).buffer

    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.image.insert_staged',
          arguments: {
            blobId: 'image-blob',
            name: 'pixel.gif',
            size: data.byteLength,
            data,
            afterBlockIndex: 0,
            widthPx: 120,
            heightPx: 80,
            alignment: 'right',
          },
        },
        {
          loadStaged: async () => undefined,
          save: async () => ({ fileName: 'document.docx' }),
        },
      ),
    ).resolves.toMatchObject({ handled: true, ok: true })

    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    expect(plan.changedCount).toBe(1)
    const saved = await saveDocx(parsed, plan.saveBlocks)
    const reopened = await parseDocx(saved)
    expect(reopened.blocks[1]).toMatchObject({
      type: 'image',
      imageWidthPx: 120,
      imageHeightPx: 80,
      imageAlign: 'right',
    })
    expect(reopened.blocks[1].imageDataUrl).toMatch(/^data:image\/gif;base64,/)

    const reopenedEditor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(reopened.blocks) as never,
    })
    const reopenedPlan = pmDocToSavePlan(reopenedEditor.getJSON() as PmNode, reopened.blocks)
    expect(reopenedPlan.changedCount).toBe(0)
    expect(await saveDocx(reopened, reopenedPlan.saveBlocks)).toEqual(saved)

    editor.destroy()
    reopenedEditor.destroy()
  })
})
