import { Editor } from '@tiptap/core'
import { parseDocx, saveDocx, type NoteInfo } from '@genoffice/docx-engine'
import { describe, expect, it } from 'vitest'

import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'
import {
  collectDocxNotes,
  deleteDocxNote,
  insertDocxNote,
  updateDocxNote,
  type ManagedDocxNoteKey,
  type ManagedDocxNoteState,
} from '../src/renderer/editor/note-actions'

function createEditor(text = 'Alpha'): Editor {
  return new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: {
      type: 'doc',
      content: [
        {
          type: 'docParagraph',
          attrs: { docxIndex: 0 },
          content: [{ type: 'text', text }],
        },
      ],
    },
  })
}

function createEditorWithNote(kind: 'footnote' | 'endnote', id: string): Editor {
  return new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: {
      type: 'doc',
      content: [
        {
          type: 'docParagraph',
          attrs: { docxIndex: 0 },
          content: [
            { type: 'text', text: 'Alpha' },
            { type: 'docNoteRef', attrs: { kind, id, num: 1 } },
          ],
        },
      ],
    },
  })
}

function noteAttrs(editor: Editor): Array<Record<string, unknown>> {
  const attrs: Array<Record<string, unknown>> = []
  editor.state.doc.descendants((node) => {
    if (node.type.name === 'docNoteRef') attrs.push(node.attrs)
  })
  return attrs
}

describe('DOCX note actions', () => {
  it('inserts one explicit footnote as a native Undo/Redo unit', () => {
    const editor = createEditor()
    const input = {
      range: { from: 6, to: 6 },
      kind: 'footnote' as const,
      noteId: 2,
      text: 'Bounded footnote',
    }

    expect(insertDocxNote(editor, [], input)).toEqual({
      ok: true,
      from: 6,
      to: 6,
      kind: 'footnote',
      noteId: '2',
      number: 1,
      changed: true,
    })
    expect(noteAttrs(editor)).toMatchObject([
      {
        kind: 'footnote',
        id: '2',
        num: 1,
        noteText: 'Bounded footnote',
        managed: true,
      },
    ])

    const managed = new Map<ManagedDocxNoteKey, ManagedDocxNoteState>([
      ['footnote:2', { original: null, deleteWhenUnreferenced: true }],
    ])
    expect(collectDocxNotes(editor, [], [], managed).footnotes).toEqual([
      { id: '2', text: 'Bounded footnote' },
    ])
    expect(editor.commands.undo()).toBe(true)
    expect(collectDocxNotes(editor, [{ id: '2', text: 'Bounded footnote' }], [], managed)).toEqual({
      footnotes: [],
      endnotes: [],
    })
    expect(editor.commands.redo()).toBe(true)
    expect(collectDocxNotes(editor, [], [], managed).footnotes).toEqual([
      { id: '2', text: 'Bounded footnote' },
    ])
    editor.destroy()
  })

  it('rejects a duplicate note id and an invalid cross-block target without mutation', () => {
    const editor = createEditor()
    const existing: NoteInfo[] = [{ id: '2', text: 'Existing' }]
    expect(
      insertDocxNote(editor, existing, {
        range: { from: 1, to: 1 },
        kind: 'footnote',
        noteId: 2,
        text: 'Duplicate',
      }),
    ).toMatchObject({ ok: false, error: 'invalid_arguments' })
    expect(noteAttrs(editor)).toEqual([])

    expect(
      insertDocxNote(editor, [], {
        range: { from: 0, to: 7 },
        kind: 'endnote',
        noteId: 3,
        text: 'Bad range',
      }),
    ).toMatchObject({ ok: false, error: 'invalid_arguments' })
    expect(noteAttrs(editor)).toEqual([])
    editor.destroy()
  })

  it('persists an inserted endnote through save and reopen', async () => {
    const source = await buildDocx({
      bodyXml: '<w:p><w:r><w:t>Alpha</w:t></w:r></w:p>',
    })
    const parsed = await parseDocx(source)
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks) as never,
    })
    expect(
      insertDocxNote(editor, parsed.endnotes, {
        range: { from: 6, to: 6 },
        kind: 'endnote',
        noteId: 2,
        text: 'Replay-safe endnote',
      }),
    ).toMatchObject({ ok: true })

    const notes = collectDocxNotes(
      editor,
      parsed.footnotes,
      parsed.endnotes,
      new Map([['endnote:2', { original: null, deleteWhenUnreferenced: true }]]),
    )
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    const reopened = await parseDocx(
      await saveDocx(parsed, plan.saveBlocks, {
        footnotes: notes.footnotes,
        endnotes: notes.endnotes,
      }),
    )
    expect(reopened.endnotes).toEqual([{ id: '2', text: 'Replay-safe endnote' }])
    expect(reopened.blocks[0].runs).toEqual([
      { text: 'Alpha' },
      { text: '1', noteRef: { kind: 'endnote', id: '2' } },
    ])
    editor.destroy()
  })

  it('updates one stable note body and restores the original body through Undo/Redo', () => {
    const editor = createEditorWithNote('footnote', '2')
    const original: NoteInfo = { id: '2', text: 'Original body' }
    const managed = new Map<ManagedDocxNoteKey, ManagedDocxNoteState>([
      ['footnote:2', { original, deleteWhenUnreferenced: false }],
    ])

    expect(
      updateDocxNote(editor, [original], {
        kind: 'footnote',
        noteId: 2,
        text: 'Updated body',
      }),
    ).toEqual({
      ok: true,
      kind: 'footnote',
      noteId: '2',
      references: 1,
      changed: true,
      original,
    })
    expect(collectDocxNotes(editor, [original], [], managed).footnotes).toEqual([
      { id: '2', text: 'Updated body' },
    ])
    expect(editor.commands.undo()).toBe(true)
    expect(
      collectDocxNotes(editor, [{ id: '2', text: 'Updated body' }], [], managed).footnotes,
    ).toEqual([original])
    expect(editor.commands.redo()).toBe(true)
    expect(collectDocxNotes(editor, [original], [], managed).footnotes).toEqual([
      { id: '2', text: 'Updated body' },
    ])
    editor.destroy()
  })

  it('rejects updating a missing note identity without mutation', () => {
    const editor = createEditor()
    expect(
      updateDocxNote(editor, [], { kind: 'endnote', noteId: 7, text: 'Missing' }),
    ).toMatchObject({ ok: false, error: 'invalid_arguments' })
    expect(noteAttrs(editor)).toEqual([])
    editor.destroy()
  })

  it('persists an updated footnote body through save and reopen', async () => {
    const source = await buildDocx({
      bodyXml: '<w:p><w:r><w:t>Alpha</w:t></w:r><w:r><w:footnoteReference w:id="2"/></w:r></w:p>',
      extraRels:
        '<Relationship Id="rId40" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>',
      extraParts: [
        {
          path: 'word/footnotes.xml',
          xml:
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
            '<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
            '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>' +
            '<w:footnote w:id="2"><w:p><w:r><w:footnoteRef/></w:r><w:r><w:t>Original body</w:t></w:r></w:p></w:footnote>' +
            '</w:footnotes>',
          contentType:
            'application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml',
        },
      ],
    })
    const parsed = await parseDocx(source)
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks) as never,
    })
    const original = parsed.footnotes[0]
    expect(
      updateDocxNote(editor, parsed.footnotes, {
        kind: 'footnote',
        noteId: 2,
        text: 'Persisted update',
      }),
    ).toMatchObject({ ok: true, changed: true })
    const notes = collectDocxNotes(
      editor,
      parsed.footnotes,
      parsed.endnotes,
      new Map([['footnote:2', { original, deleteWhenUnreferenced: false }]]),
    )
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    const reopened = await parseDocx(
      await saveDocx(parsed, plan.saveBlocks, {
        footnotes: notes.footnotes,
        endnotes: notes.endnotes,
      }),
    )
    expect(reopened.footnotes[0].text).toBe('Persisted update')
    editor.destroy()
  })

  it('deletes a stable note, renumbers peers, and restores both through Undo', () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docParagraph',
            attrs: { docxIndex: 0 },
            content: [
              { type: 'text', text: 'A' },
              { type: 'docNoteRef', attrs: { kind: 'footnote', id: '2', num: 1 } },
              { type: 'text', text: 'B' },
              { type: 'docNoteRef', attrs: { kind: 'footnote', id: '3', num: 2 } },
            ],
          },
        ],
      },
    })
    const notes: NoteInfo[] = [
      { id: '2', text: 'First' },
      { id: '3', text: 'Second' },
    ]
    const managed = new Map<ManagedDocxNoteKey, ManagedDocxNoteState>([
      ['footnote:2', { original: notes[0], deleteWhenUnreferenced: true }],
    ])
    expect(deleteDocxNote(editor, notes, { kind: 'footnote', noteId: 2 })).toEqual({
      ok: true,
      kind: 'footnote',
      noteId: '2',
      references: 1,
      renumbered: 1,
      changed: true,
      original: notes[0],
    })
    expect(noteAttrs(editor)).toMatchObject([{ kind: 'footnote', id: '3', num: 1 }])
    expect(collectDocxNotes(editor, notes, [], managed).footnotes).toEqual([notes[1]])
    expect(editor.commands.undo()).toBe(true)
    expect(noteAttrs(editor)).toMatchObject([
      { kind: 'footnote', id: '2', num: 1 },
      { kind: 'footnote', id: '3', num: 2 },
    ])
    expect(collectDocxNotes(editor, [notes[1]], [], managed).footnotes).toEqual(notes)
    editor.destroy()
  })

  it('rejects deleting a missing note identity without mutation', () => {
    const editor = createEditor()
    expect(deleteDocxNote(editor, [], { kind: 'footnote', noteId: 9 })).toMatchObject({
      ok: false,
      error: 'invalid_arguments',
    })
    expect(noteAttrs(editor)).toEqual([])
    editor.destroy()
  })

  it('persists stable note deletion and peer renumbering through save and reopen', async () => {
    const source = await buildDocx({
      bodyXml:
        '<w:p><w:r><w:t>A</w:t></w:r><w:r><w:footnoteReference w:id="2"/></w:r>' +
        '<w:r><w:t>B</w:t></w:r><w:r><w:footnoteReference w:id="3"/></w:r></w:p>',
      extraRels:
        '<Relationship Id="rId40" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>',
      extraParts: [
        {
          path: 'word/footnotes.xml',
          xml:
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
            '<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
            '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>' +
            '<w:footnote w:id="2"><w:p><w:r><w:footnoteRef/></w:r><w:r><w:t>First</w:t></w:r></w:p></w:footnote>' +
            '<w:footnote w:id="3"><w:p><w:r><w:footnoteRef/></w:r><w:r><w:t>Second</w:t></w:r></w:p></w:footnote>' +
            '</w:footnotes>',
          contentType:
            'application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml',
        },
      ],
    })
    const parsed = await parseDocx(source)
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks) as never,
    })
    const removed = parsed.footnotes[0]
    expect(deleteDocxNote(editor, parsed.footnotes, { kind: 'footnote', noteId: 2 })).toMatchObject(
      {
        ok: true,
        renumbered: 1,
      },
    )
    const notes = collectDocxNotes(
      editor,
      parsed.footnotes,
      parsed.endnotes,
      new Map([['footnote:2', { original: removed, deleteWhenUnreferenced: true }]]),
    )
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    const reopened = await parseDocx(
      await saveDocx(parsed, plan.saveBlocks, {
        footnotes: notes.footnotes,
        endnotes: notes.endnotes,
      }),
    )
    expect(reopened.footnotes).toEqual([{ id: '3', text: 'Second' }])
    expect(reopened.blocks[0].runs).toEqual([
      { text: 'AB' },
      { text: '1', noteRef: { kind: 'footnote', id: '3' } },
    ])
    editor.destroy()
  })
})
