import { Editor } from '@tiptap/core'
import { afterEach, describe, expect, it } from 'vitest'

import type { DocState, PendingNumbering } from '../src/renderer/doc-state'
import { editorExtensions } from '../src/renderer/editor/extensions'
import {
  continueNumberingAt,
  restartNumberingAt,
  type NumberingContext,
} from '../src/renderer/numbering-actions'

const editors: Editor[] = []

afterEach(() => {
  for (const editor of editors.splice(0)) editor.destroy()
})

describe('DOCX numbering actions', () => {
  it('clones the source definition and restarts only same-list items at and after the anchor', () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docListItem',
            attrs: { kind: 'ordered', numId: '7', ilvl: 0 },
            content: [{ type: 'text', text: 'Before' }],
          },
          {
            type: 'docListItem',
            attrs: { kind: 'ordered', numId: '7', ilvl: 1 },
            content: [{ type: 'text', text: 'Restart' }],
          },
          { type: 'docParagraph', content: [{ type: 'text', text: 'Gap' }] },
          {
            type: 'docListItem',
            attrs: { kind: 'ordered', numId: '7', ilvl: 1 },
            content: [{ type: 'text', text: 'After' }],
          },
          {
            type: 'docListItem',
            attrs: { kind: 'ordered', numId: '8', ilvl: 0 },
            content: [{ type: 'text', text: 'Other list' }],
          },
        ],
      },
    })
    editors.push(editor)

    const sourceDef = {
      numId: '7',
      abstractNumId: '3',
      levels: {
        0: { numFmt: 'decimal', lvlText: '%1.', start: 1, indentLeft: 720, hanging: 360 },
        1: { numFmt: 'lowerLetter', lvlText: '%2)', start: 1, indentLeft: 1440, hanging: 360 },
      },
      startOverrides: {},
    }
    const otherDef = { ...sourceDef, numId: '8', abstractNumId: '4' }
    const definitions = new Map([
      ['7', sourceDef],
      ['8', otherDef],
    ])
    ;(editor.storage.listNumbering as { defs: Map<string, typeof sourceDef> }).defs = definitions

    let pending: PendingNumbering = { newDefs: [], restartNums: [] }
    const context: NumberingContext = {
      editor,
      doc: { parsed: { numbering: definitions } } as unknown as DocState,
      pendingNumbering: pending,
      setPendingNumbering: (update) => {
        pending = update(pending)
        context.pendingNumbering = pending
      },
      numIdFloorRef: { current: 0 },
      setStatus: () => undefined,
    }

    expect(restartNumberingAt(context, { blockIndex: 1, start: 3 })).toEqual({
      ok: true,
      changed: 2,
    })
    expect(pending.restartNums).toEqual([
      { numId: '9', abstractNumId: '3', startOverrides: { 1: 3 } },
    ])
    expect(editor.state.doc.child(0).attrs.numId).toBe('7')
    expect(editor.state.doc.child(1).attrs.numId).toBe('9')
    expect(editor.state.doc.child(3).attrs.numId).toBe('9')
    expect(editor.state.doc.child(4).attrs.numId).toBe('8')

    expect(editor.commands.undo()).toBe(true)
    expect(editor.state.doc.child(1).attrs.numId).toBe('7')
    expect(editor.state.doc.child(3).attrs.numId).toBe('7')
  })

  it('continues only same-list items at and after an explicit anchor', () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docListItem',
            attrs: { kind: 'ordered', numId: '9', ilvl: 0 },
            content: [{ type: 'text', text: 'Previous list' }],
          },
          {
            type: 'docListItem',
            attrs: { kind: 'ordered', numId: '7', ilvl: 0 },
            content: [{ type: 'text', text: 'Continue' }],
          },
          { type: 'docParagraph', content: [{ type: 'text', text: 'Gap' }] },
          {
            type: 'docListItem',
            attrs: { kind: 'ordered', numId: '7', ilvl: 1 },
            content: [{ type: 'text', text: 'Same current list' }],
          },
          {
            type: 'docListItem',
            attrs: { kind: 'ordered', numId: '8', ilvl: 0 },
            content: [{ type: 'text', text: 'Other list' }],
          },
        ],
      },
    })
    editors.push(editor)
    const context = {
      editor,
      doc: null,
      pendingNumbering: { newDefs: [], restartNums: [] },
      setPendingNumbering: () => undefined,
      numIdFloorRef: { current: 0 },
      setStatus: () => undefined,
    } satisfies NumberingContext

    expect(continueNumberingAt(context, { blockIndex: 1, previousBlockIndex: 0 })).toEqual({
      ok: true,
      changed: 2,
    })
    expect(editor.state.doc.child(0).attrs.numId).toBe('9')
    expect(editor.state.doc.child(1).attrs.numId).toBe('9')
    expect(editor.state.doc.child(3).attrs.numId).toBe('9')
    expect(editor.state.doc.child(4).attrs.numId).toBe('8')

    expect(editor.commands.undo()).toBe(true)
    expect(editor.state.doc.child(1).attrs.numId).toBe('7')
    expect(editor.state.doc.child(3).attrs.numId).toBe('7')
  })
})
