import { Editor } from '@tiptap/core'
import { parseDocx, saveDocx } from '@genoffice/docx-engine'
import { describe, expect, it } from 'vitest'

import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import {
  addDocxComment,
  collectDocxComments,
  deleteDocxComment,
  replyDocxComment,
  setDocxCommentResolved,
} from '../src/renderer/editor/comments'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'

const comment = {
  id: '1',
  author: 'Agent',
  initials: null,
  date: '2026-08-30T02:00:00Z',
  text: 'Review this text.',
} as const

function createEditor(): Editor {
  return new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: {
      type: 'doc',
      content: [
        {
          type: 'docParagraph',
          attrs: { docxIndex: 0 },
          content: [{ type: 'text', text: 'Alpha' }],
        },
      ],
    },
  })
}

describe('DOCX comment actions', () => {
  it('adds body metadata and an exact-range anchor as one Undo/Redo unit', () => {
    const editor = createEditor()
    expect(addDocxComment(editor, [], { range: { from: 1, to: 6 }, comment })).toMatchObject({
      ok: true,
      id: '1',
      from: 1,
      to: 6,
      changed: true,
    })
    expect(collectDocxComments(editor, [])).toEqual([
      { id: '1', author: 'Agent', date: comment.date, text: comment.text },
    ])
    expect(editor.state.doc.child(0).firstChild?.marks[0]?.attrs.ids).toBe('1')
    expect(editor.commands.undo()).toBe(true)
    expect(collectDocxComments(editor, [])).toEqual([])
    expect(editor.state.doc.child(0).firstChild?.marks).toHaveLength(0)
    expect(editor.commands.redo()).toBe(true)
    expect(collectDocxComments(editor, [])).toHaveLength(1)
    editor.destroy()
  })

  it('rejects a duplicate stable comment id without mutation', () => {
    const editor = createEditor()
    const existing = [{ id: '1', author: 'User', text: 'Existing' }]
    expect(addDocxComment(editor, existing, { range: { from: 1, to: 6 }, comment })).toMatchObject({
      ok: false,
      error: 'invalid_arguments',
    })
    expect(collectDocxComments(editor, existing)).toEqual(existing)
    expect(editor.state.doc.child(0).firstChild?.marks).toHaveLength(0)
    editor.destroy()
  })

  it('persists comment metadata and its anchor through save and reopen', async () => {
    const parsed = await parseDocx(
      await buildDocx({ bodyXml: '<w:p><w:r><w:t>Alpha</w:t></w:r></w:p>' }),
    )
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks) as never,
    })
    expect(
      addDocxComment(editor, parsed.comments, { range: { from: 1, to: 6 }, comment }),
    ).toMatchObject({
      ok: true,
    })
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    const reopened = await parseDocx(
      await saveDocx(parsed, plan.saveBlocks, {
        comments: collectDocxComments(editor, parsed.comments),
      }),
    )
    expect(reopened.comments).toMatchObject([
      { id: '1', author: 'Agent', date: comment.date, text: comment.text },
    ])
    expect(reopened.blocks[0].runs?.find((run) => run.commentIds)?.commentIds).toEqual(['1'])
    editor.destroy()
  })

  it('adds a stable reply on the parent anchor as one Undo unit', () => {
    const editor = createEditor()
    const parent = addDocxComment(editor, [], {
      range: { from: 1, to: 6 },
      comment,
    })
    if (!parent.ok) throw new Error(parent.message)
    const reply = {
      id: '2',
      author: 'Reviewer',
      initials: 'RV',
      date: '2026-08-30T02:01:00Z',
      text: 'Reply text.',
    }
    expect(
      replyDocxComment(editor, parent.comments, { parentId: '1', comment: reply }),
    ).toMatchObject({
      ok: true,
      id: '2',
      parentId: '1',
      references: 1,
      changed: true,
    })
    expect(collectDocxComments(editor, [])[1]).toMatchObject({ id: '2', parentId: '1' })
    expect(editor.state.doc.child(0).firstChild?.marks[0]?.attrs.ids).toBe('1 2')
    expect(editor.commands.undo()).toBe(true)
    expect(collectDocxComments(editor, [])).toEqual(parent.comments)
    expect(editor.state.doc.child(0).firstChild?.marks[0]?.attrs.ids).toBe('1')
    editor.destroy()
  })

  it('rejects a reply to a missing parent without mutation', () => {
    const editor = createEditor()
    expect(
      replyDocxComment(editor, [], {
        parentId: '9',
        comment: {
          id: '2',
          author: 'Reviewer',
          initials: null,
          date: '2026-08-30T02:01:00Z',
          text: 'Missing parent.',
        },
      }),
    ).toMatchObject({ ok: false, error: 'invalid_arguments' })
    expect(collectDocxComments(editor, [])).toEqual([])
    editor.destroy()
  })

  it('persists the reply relationship and shared anchor through save and reopen', async () => {
    const parsed = await parseDocx(
      await buildDocx({ bodyXml: '<w:p><w:r><w:t>Alpha</w:t></w:r></w:p>' }),
    )
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks) as never,
    })
    const parent = addDocxComment(editor, parsed.comments, {
      range: { from: 1, to: 6 },
      comment,
    })
    if (!parent.ok) throw new Error(parent.message)
    expect(
      replyDocxComment(editor, parent.comments, {
        parentId: '1',
        comment: {
          id: '2',
          author: 'Reviewer',
          initials: null,
          date: '2026-08-30T02:01:00Z',
          text: 'Reply text.',
        },
      }),
    ).toMatchObject({ ok: true })
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    const reopened = await parseDocx(
      await saveDocx(parsed, plan.saveBlocks, {
        comments: collectDocxComments(editor, parsed.comments),
      }),
    )
    expect(reopened.comments.find((entry) => entry.id === '2')).toMatchObject({ parentId: '1' })
    expect(reopened.blocks[0].runs?.find((run) => run.commentIds)?.commentIds).toEqual(['1', '2'])
    editor.destroy()
  })

  it('sets the parent thread to an explicit resolved state as one Undo unit', () => {
    const editor = createEditor()
    const comments = [
      { id: '1', author: 'User', text: 'Parent' },
      { id: '2', author: 'Reviewer', text: 'Reply', parentId: '1' },
    ]
    expect(setDocxCommentResolved(editor, comments, { id: '1', resolved: true })).toMatchObject({
      ok: true,
      id: '1',
      resolved: true,
      affected: 2,
      changed: true,
    })
    expect(collectDocxComments(editor, comments).map((entry) => entry.done)).toEqual([true, true])
    expect(editor.commands.undo()).toBe(true)
    expect(collectDocxComments(editor, comments)).toEqual(comments)
    editor.destroy()
  })

  it('rejects resolving a reply or missing thread id without mutation', () => {
    const editor = createEditor()
    const comments = [
      { id: '1', author: 'User', text: 'Parent' },
      { id: '2', author: 'Reviewer', text: 'Reply', parentId: '1' },
    ]
    expect(setDocxCommentResolved(editor, comments, { id: '2', resolved: true })).toMatchObject({
      ok: false,
      error: 'invalid_arguments',
    })
    expect(collectDocxComments(editor, comments)).toEqual(comments)
    editor.destroy()
  })

  it('persists resolved state through commentsExtended save and reopen', async () => {
    const parsed = await parseDocx(
      await buildDocx({ bodyXml: '<w:p><w:r><w:t>Alpha</w:t></w:r></w:p>' }),
    )
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks) as never,
    })
    const added = addDocxComment(editor, parsed.comments, {
      range: { from: 1, to: 6 },
      comment,
    })
    if (!added.ok) throw new Error(added.message)
    expect(
      setDocxCommentResolved(editor, added.comments, { id: '1', resolved: true }),
    ).toMatchObject({
      ok: true,
    })
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    const reopened = await parseDocx(
      await saveDocx(parsed, plan.saveBlocks, {
        comments: collectDocxComments(editor, parsed.comments),
      }),
    )
    expect(reopened.comments[0]).toMatchObject({ id: '1', done: true })
    editor.destroy()
  })

  it('deletes a parent, replies, and all anchor ids as one Undo unit', () => {
    const editor = createEditor()
    const parent = addDocxComment(editor, [], { range: { from: 1, to: 6 }, comment })
    if (!parent.ok) throw new Error(parent.message)
    const replied = replyDocxComment(editor, parent.comments, {
      parentId: '1',
      comment: {
        id: '2',
        author: 'Reviewer',
        initials: null,
        date: '2026-08-30T02:01:00Z',
        text: 'Reply',
      },
    })
    if (!replied.ok) throw new Error(replied.message)
    expect(deleteDocxComment(editor, replied.comments, { id: '1' })).toMatchObject({
      ok: true,
      id: '1',
      deleted: 2,
      anchors: 1,
      changed: true,
    })
    expect(collectDocxComments(editor, [])).toEqual([])
    expect(editor.state.doc.child(0).firstChild?.marks).toHaveLength(0)
    expect(editor.commands.undo()).toBe(true)
    expect(collectDocxComments(editor, [])).toEqual(replied.comments)
    expect(editor.state.doc.child(0).firstChild?.marks[0]?.attrs.ids).toBe('1 2')
    editor.destroy()
  })

  it('rejects a missing comment id without mutation', () => {
    const editor = createEditor()
    const comments = [{ id: '1', author: 'User', text: 'Parent' }]
    expect(deleteDocxComment(editor, comments, { id: '9' })).toMatchObject({
      ok: false,
      error: 'invalid_arguments',
    })
    expect(collectDocxComments(editor, comments)).toEqual(comments)
    editor.destroy()
  })

  it('persists cascaded comment deletion through save and reopen', async () => {
    const parsed = await parseDocx(
      await buildDocx({ bodyXml: '<w:p><w:r><w:t>Alpha</w:t></w:r></w:p>' }),
    )
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks) as never,
    })
    const added = addDocxComment(editor, parsed.comments, {
      range: { from: 1, to: 6 },
      comment,
    })
    if (!added.ok) throw new Error(added.message)
    expect(deleteDocxComment(editor, added.comments, { id: '1' })).toMatchObject({ ok: true })
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    const reopened = await parseDocx(
      await saveDocx(parsed, plan.saveBlocks, {
        comments: collectDocxComments(editor, parsed.comments),
      }),
    )
    expect(reopened.comments).toEqual([])
    expect(reopened.blocks[0].runs?.some((run) => run.commentIds)).toBe(false)
    editor.destroy()
  })
})
