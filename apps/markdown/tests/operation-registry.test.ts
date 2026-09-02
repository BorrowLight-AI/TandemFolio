import { afterEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'

import { buildExtensions } from '../src/renderer/editor/extensions'
import { executeMarkdownOperation } from '../src/renderer/operations/registry'
import { markdownOperationCatalog } from '../src/renderer/operations/catalog'
import { markdownRetainedCommandAudit } from '../src/renderer/operations/baseline'
import { parseDocText, serializeDocText } from '../src/renderer/markdown/docText'
import { setMarkdownFrontmatter } from '../src/renderer/editor/frontmatter-actions'
import { setMarkdownTextMarks } from '../src/renderer/editor/text-mark-actions'

const editors: Editor[] = []

afterEach(() => {
  for (const editor of editors.splice(0)) editor.destroy()
})

function createEditor(markdown: string): Editor {
  const editor = new Editor({
    extensions: buildExtensions({
      slashController: { onOpen() {}, onUpdate() {}, onKeyDown: () => false, onClose() {} },
      slashItems: () => [],
    }),
    content: '',
  })
  editor
    .chain()
    .setMeta('addToHistory', false)
    .setContent(markdown, { contentType: 'markdown' })
    .run()
  editors.push(editor)
  return editor
}

function textPosition(editor: Editor, text: string): number {
  let found = -1
  editor.state.doc.descendants((node, position) => {
    if (found < 0 && node.isTextblock && node.textContent === text) found = position + 1
  })
  return found
}

const services = {
  loadStaged: async () => undefined,
  save: async () => ({ ok: true as const, fileName: 'notes.md' }),
  exportDocx: async () => ({ ok: true as const, fileName: 'notes.docx' }),
  openPrintDialog: () => ({ ok: true as const }),
  setAutoSave: () => undefined,
  setFrontmatter: () => undefined,
}

describe('Markdown operation registry', () => {
  it('maps every retained command family and every registry descriptor without gaps', () => {
    expect(markdownRetainedCommandAudit.map((entry) => entry.disposition)).not.toContain('missing')
    const mapped = new Set(markdownRetainedCommandAudit.flatMap((entry) => entry.operationIds))
    expect([...mapped].sort()).toEqual(
      markdownOperationCatalog.operations.map((operation) => operation.id).sort(),
    )
  })

  it('sets an explicit Markdown text selection without adding an undo entry', async () => {
    const editor = createEditor('Hello world')

    await expect(
      executeMarkdownOperation(
        editor,
        { operation: 'markdown.selection.set', arguments: { from: 2, to: 7 } },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'markdown.selection.set',
      ok: true,
      output: { from: 2, to: 7 },
      checkpointRecovery: false,
    })
    expect(editor.state.selection.from).toBe(2)
    expect(editor.state.selection.to).toBe(7)
    expect(editor.commands.undo()).toBe(false)
  })

  it('applies an inline mark when the Ribbon receives a ProseMirror select-all range', () => {
    const editor = createEditor('First paragraph\n\nSecond paragraph')
    editor.commands.selectAll()
    const { from, to } = editor.state.selection

    const result = setMarkdownTextMarks(editor, {
      from,
      to,
      marks: { bold: true, italic: false, strike: false, code: false, link: null },
    })

    expect(result).toEqual({ ok: true })
    expect(editor.getMarkdown()).toBe('**First paragraph**\n\n**Second paragraph**')
  })

  it('rejects an invalid explicit Markdown text selection', async () => {
    const editor = createEditor('Hello')

    await expect(
      executeMarkdownOperation(
        editor,
        { operation: 'markdown.selection.set', arguments: { from: 5, to: 2 } },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'markdown.selection.set',
      ok: false,
      error: 'execution_failed',
      message: 'Markdown selection requires from <= to.',
    })
  })

  it('sets an addressed Markdown code-block language with undo and reopen fidelity', async () => {
    const editor = createEditor('```javascript\nconst value = 1\n```')

    await expect(
      executeMarkdownOperation(
        editor,
        {
          operation: 'markdown.code_block.set_language',
          arguments: { textBlockIndex: 0, language: 'python' },
        },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'markdown.code_block.set_language',
      ok: true,
    })
    expect(editor.getMarkdown()).toContain('```python')
    const reopened = createEditor(editor.getMarkdown())
    expect(reopened.state.doc.child(0).attrs.language).toBe('python')

    expect(editor.commands.undo()).toBe(true)
    expect(editor.state.doc.child(0).attrs.language).toBe('javascript')
  })

  it('maps the plaintext final state to a null code-block language', async () => {
    const editor = createEditor('```javascript\nvalue\n```')

    await executeMarkdownOperation(
      editor,
      {
        operation: 'markdown.code_block.set_language',
        arguments: { textBlockIndex: 0, language: 'plaintext' },
      },
      services,
    )
    expect(editor.state.doc.child(0).attrs.language).toBeNull()
    expect(editor.getMarkdown()).toContain('```\nvalue')
  })

  it('rejects a Markdown code language target that is not a code block', async () => {
    const editor = createEditor('Paragraph')

    await expect(
      executeMarkdownOperation(
        editor,
        {
          operation: 'markdown.code_block.set_language',
          arguments: { textBlockIndex: 0, language: 'typescript' },
        },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'markdown.code_block.set_language',
      ok: false,
      error: 'execution_failed',
      message: 'Markdown text block 0 is not a code block.',
    })
  })

  it('duplicates one explicit top-level Markdown block with native undo', async () => {
    const editor = createEditor('First\n\nSecond')

    await expect(
      executeMarkdownOperation(
        editor,
        {
          operation: 'markdown.block.update',
          arguments: {
            blockIndex: 1,
            action: 'duplicate',
            afterBlockIndex: null,
            content: null,
          },
        },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'markdown.block.update',
      ok: true,
    })
    expect(editor.state.doc.childCount).toBe(3)
    expect(editor.state.doc.child(2).textContent).toBe('Second')
    expect(editor.commands.undo()).toBe(true)
    expect(editor.state.doc.childCount).toBe(2)
  })

  it('moves one explicit Markdown block after a revision-scoped boundary', async () => {
    const editor = createEditor('First\n\nSecond\n\nThird')

    await executeMarkdownOperation(
      editor,
      {
        operation: 'markdown.block.update',
        arguments: {
          blockIndex: 2,
          action: 'move',
          afterBlockIndex: -1,
          content: null,
        },
      },
      services,
    )
    expect(
      Array.from(
        { length: editor.state.doc.childCount },
        (_, index) => editor.state.doc.child(index).textContent,
      ),
    ).toEqual(['Third', 'First', 'Second'])
  })

  it('deletes the sole Markdown block while preserving the document block invariant', async () => {
    const editor = createEditor('Only')

    await executeMarkdownOperation(
      editor,
      {
        operation: 'markdown.block.update',
        arguments: {
          blockIndex: 0,
          action: 'delete',
          afterBlockIndex: null,
          content: null,
        },
      },
      services,
    )
    expect(editor.state.doc.childCount).toBe(1)
    expect(editor.state.doc.child(0).type.name).toBe('paragraph')
    expect(editor.state.doc.child(0).textContent).toBe('')
  })

  it('adds explicit content below one addressed Markdown block', async () => {
    const editor = createEditor('First\n\nThird')

    await executeMarkdownOperation(
      editor,
      {
        operation: 'markdown.block.update',
        arguments: {
          blockIndex: 0,
          action: 'add_below',
          afterBlockIndex: null,
          content: 'Second',
        },
      },
      services,
    )
    expect(
      Array.from(
        { length: editor.state.doc.childCount },
        (_, index) => editor.state.doc.child(index).textContent,
      ),
    ).toEqual(['First', 'Second', 'Third'])
  })

  it('rejects fields that do not belong to the selected Markdown block action', async () => {
    const editor = createEditor('Only')

    await expect(
      executeMarkdownOperation(
        editor,
        {
          operation: 'markdown.block.update',
          arguments: {
            blockIndex: 0,
            action: 'duplicate',
            afterBlockIndex: -1,
            content: null,
          },
        },
        services,
      ),
    ).resolves.toMatchObject({
      handled: true,
      ok: false,
      message: 'Markdown block fields must match the selected action.',
    })
  })

  it('updates a table relative to an explicit cell position with undo and reopen fidelity', async () => {
    const editor = createEditor('| A | B |\n| --- | --- |\n| C | D |')
    const position = textPosition(editor, 'C')

    await expect(
      executeMarkdownOperation(
        editor,
        {
          operation: 'markdown.table.update',
          arguments: { position, action: 'add_row_after', headerRow: null },
        },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'markdown.table.update',
      ok: true,
    })
    const table = editor.getJSON().content?.find((node) => node.type === 'table')
    expect(table?.content).toHaveLength(3)
    const reopened = createEditor(editor.getMarkdown())
    expect(reopened.getJSON().content?.find((node) => node.type === 'table')?.content).toHaveLength(
      3,
    )

    expect(editor.commands.undo()).toBe(true)
    expect(editor.getJSON().content?.find((node) => node.type === 'table')?.content).toHaveLength(2)
  })

  it('sets the table header row to an explicit disabled final state', async () => {
    const editor = createEditor('| A | B |\n| --- | --- |\n| C | D |')

    await expect(
      executeMarkdownOperation(
        editor,
        {
          operation: 'markdown.table.update',
          arguments: {
            position: textPosition(editor, 'A'),
            action: 'set_header_row',
            headerRow: false,
          },
        },
        services,
      ),
    ).resolves.toMatchObject({ handled: true, ok: true })
    const table = Array.from({ length: editor.state.doc.childCount }, (_, index) =>
      editor.state.doc.child(index),
    ).find((node) => node.type.name === 'table')
    expect(table?.child(0).child(0).type.name).toBe('tableCell')
  })

  it('rejects a Markdown table update outside a table without mutating', async () => {
    const editor = createEditor('Outside')

    await expect(
      executeMarkdownOperation(
        editor,
        {
          operation: 'markdown.table.update',
          arguments: { position: 1, action: 'delete_table', headerRow: null },
        },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'markdown.table.update',
      ok: false,
      error: 'execution_failed',
      message: 'Markdown table update position 1 is not inside a table.',
    })
    expect(editor.getMarkdown()).toContain('Outside')
  })

  it('rejects a headerRow value for non-header Markdown table actions', async () => {
    const editor = createEditor('| A |\n| --- |\n| B |')

    await expect(
      executeMarkdownOperation(
        editor,
        {
          operation: 'markdown.table.update',
          arguments: {
            position: textPosition(editor, 'B'),
            action: 'delete_row',
            headerRow: false,
          },
        },
        services,
      ),
    ).resolves.toMatchObject({
      handled: true,
      ok: false,
      message:
        'Markdown table headerRow must be boolean only for set_header_row and null otherwise.',
    })
  })

  it('sets complete raw YAML frontmatter through the envelope and preserves it on reopen', async () => {
    const editor = createEditor('Body')
    const envelope = parseDocText('Body\n')

    await expect(
      executeMarkdownOperation(
        editor,
        {
          operation: 'markdown.frontmatter.set',
          arguments: { yaml: 'title: Typed\ntags:\n  - registry' },
        },
        {
          ...services,
          setFrontmatter: ({ yaml }) => {
            setMarkdownFrontmatter(envelope, yaml)
          },
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'markdown.frontmatter.set',
      ok: true,
    })

    const saved = serializeDocText(envelope, editor.getMarkdown())
    expect(saved).toContain('---\ntitle: Typed\ntags:\n  - registry\n---')
    expect(parseDocText(saved).frontmatter).toBe(envelope.frontmatter)
  })

  it('removes Markdown frontmatter through an explicit empty YAML final state', async () => {
    const editor = createEditor('Body')
    const envelope = parseDocText('---\ntitle: Old\n---\n\nBody\n')

    await executeMarkdownOperation(
      editor,
      { operation: 'markdown.frontmatter.set', arguments: { yaml: '' } },
      {
        ...services,
        setFrontmatter: ({ yaml }) => {
          setMarkdownFrontmatter(envelope, yaml)
        },
      },
    )
    expect(envelope.frontmatter).toBe('')
    expect(serializeDocText(envelope, editor.getMarkdown())).not.toContain('---')
  })

  it('requires Broker staging for a public Markdown image path', async () => {
    const editor = createEditor('Before')

    await expect(
      executeMarkdownOperation(
        editor,
        {
          operation: 'markdown.image.insert',
          arguments: {
            path: '/tmp/logo.png',
            position: 7,
            alt: 'Logo',
            title: null,
          },
        },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'markdown.image.insert',
      ok: false,
      error: 'execution_failed',
      message: 'markdown.image.insert requires Broker staging.',
    })
  })

  it('inserts hydrated staged image bytes with native undo and Markdown reopen fidelity', async () => {
    const editor = createEditor('Before')
    const data = Uint8Array.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a])
      .buffer as ArrayBuffer

    await expect(
      executeMarkdownOperation(
        editor,
        {
          operation: 'markdown.image.insert_staged',
          arguments: {
            blobId: 'markdown-image',
            name: 'logo.png',
            size: data.byteLength,
            data,
            position: 7,
            alt: 'Logo',
            title: 'Brand',
          },
        },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'markdown.image.insert_staged',
      ok: true,
    })
    const markdown = editor.getMarkdown()
    expect(markdown).toContain('![Logo](data:image/png;base64,')
    expect(markdown).toContain('"Brand"')
    const reopened = createEditor(markdown)
    expect(reopened.getJSON().content?.some((node) => node.type === 'image')).toBe(true)

    expect(editor.commands.undo()).toBe(true)
    expect(editor.getJSON().content?.some((node) => node.type === 'image')).toBe(false)
  })

  it('rejects staged Markdown image bytes that do not match the file extension', async () => {
    const editor = createEditor('Before')
    const data = new TextEncoder().encode('not-an-image').buffer as ArrayBuffer

    await expect(
      executeMarkdownOperation(
        editor,
        {
          operation: 'markdown.image.insert_staged',
          arguments: {
            blobId: 'fake-image',
            name: 'fake.png',
            size: data.byteLength,
            data,
            position: 7,
            alt: 'Fake',
            title: null,
          },
        },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'markdown.image.insert_staged',
      ok: false,
      error: 'invalid_arguments',
      message: 'markdown.image.insert_staged requires matching PNG, JPEG, or GIF bytes.',
    })
    expect(editor.getJSON().content?.some((node) => node.type === 'image')).toBe(false)
  })

  it('inserts a bounded Markdown table at an explicit position with undo and reopen fidelity', async () => {
    const editor = createEditor('Before')

    await expect(
      executeMarkdownOperation(
        editor,
        {
          operation: 'markdown.table.insert',
          arguments: { position: 7, rows: 2, columns: 3, headerRow: true },
        },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'markdown.table.insert',
      ok: true,
    })
    expect(editor.getJSON().content?.some((node) => node.type === 'table')).toBe(true)
    const serialized = editor.getMarkdown()
    expect(serialized).toContain('|')
    const reopened = createEditor(serialized)
    expect(reopened.getJSON().content?.some((node) => node.type === 'table')).toBe(true)

    expect(editor.commands.undo()).toBe(true)
    expect(editor.getJSON().content?.some((node) => node.type === 'table')).toBe(false)
  })

  it('inserts a Markdown divider at an explicit position with native undo', async () => {
    const editor = createEditor('Before')

    await expect(
      executeMarkdownOperation(
        editor,
        { operation: 'markdown.divider.insert', arguments: { position: 7 } },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'markdown.divider.insert',
      ok: true,
    })
    expect(editor.getJSON().content?.some((node) => node.type === 'horizontalRule')).toBe(true)
    expect(editor.commands.undo()).toBe(true)
    expect(editor.getJSON().content?.some((node) => node.type === 'horizontalRule')).toBe(false)
  })

  it('rejects an invalid Markdown structure insertion position without mutating', async () => {
    const editor = createEditor('Before')

    await expect(
      executeMarkdownOperation(
        editor,
        { operation: 'markdown.divider.insert', arguments: { position: 99 } },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'markdown.divider.insert',
      ok: false,
      error: 'execution_failed',
      message: 'Markdown insertion position 99 is invalid for document size 8.',
    })
    expect(editor.getMarkdown()).toContain('Before')
  })

  it('sets one addressed Markdown block to a task-list final state with undo and reopen fidelity', async () => {
    const editor = createEditor('Todo')

    await expect(
      executeMarkdownOperation(
        editor,
        {
          operation: 'markdown.list.set_type',
          arguments: { textBlockIndex: 0, type: 'task' },
        },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'markdown.list.set_type',
      ok: true,
    })
    expect(editor.getMarkdown()).toContain('- [ ] Todo')
    const reopened = createEditor(editor.getMarkdown())
    expect(reopened.getJSON().content?.[0]).toMatchObject({ type: 'taskList' })

    expect(editor.commands.undo()).toBe(true)
    expect(editor.getMarkdown()).not.toContain('- [ ] Todo')
  })

  it('sets one addressed list item to the none final state without deleting siblings', async () => {
    const editor = createEditor('- First\n- Second')

    await expect(
      executeMarkdownOperation(
        editor,
        {
          operation: 'markdown.list.set_type',
          arguments: { textBlockIndex: 1, type: 'none' },
        },
        services,
      ),
    ).resolves.toMatchObject({ handled: true, ok: true })
    expect(editor.getMarkdown()).toContain('- First')
    expect(editor.getMarkdown()).toContain('Second')
    expect(editor.getMarkdown()).not.toContain('- Second')
  })

  it('rejects an out-of-range Markdown list address without mutating', async () => {
    const editor = createEditor('Only')

    await expect(
      executeMarkdownOperation(
        editor,
        {
          operation: 'markdown.list.set_type',
          arguments: { textBlockIndex: 4, type: 'ordered' },
        },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'markdown.list.set_type',
      ok: false,
      error: 'execution_failed',
      message: 'Markdown text block 4 is invalid for 1 addressable text block(s).',
    })
    expect(editor.getMarkdown()).toContain('Only')
  })

  it('sets an explicit Markdown inline-mark final state with native undo and reopen fidelity', async () => {
    const editor = createEditor('Hello world')
    const arguments_ = {
      from: 1,
      to: 6,
      marks: {
        bold: true,
        italic: true,
        strike: false,
        code: false,
        link: 'https://example.com',
      },
    }

    await expect(
      executeMarkdownOperation(
        editor,
        { operation: 'markdown.text.set_marks', arguments: arguments_ },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'markdown.text.set_marks',
      ok: true,
    })
    const marks = editor.getJSON().content?.[0]?.content?.[0]?.marks
    expect(marks?.map((mark) => mark.type).sort()).toEqual(['bold', 'italic', 'link'])
    expect(marks?.find((mark) => mark.type === 'link')?.attrs).toMatchObject({
      href: 'https://example.com',
    })

    const reopened = createEditor(editor.getMarkdown())
    expect(
      reopened
        .getJSON()
        .content?.[0]?.content?.[0]?.marks?.map((mark) => mark.type)
        .sort(),
    ).toEqual(['bold', 'italic', 'link'])

    expect(editor.commands.undo()).toBe(true)
    expect(editor.getJSON().content?.[0]?.content?.[0]?.marks).toBeUndefined()
  })

  it('clears existing inline marks through an all-false final-state mask', async () => {
    const editor = createEditor('**Hello**')

    await executeMarkdownOperation(
      editor,
      {
        operation: 'markdown.text.set_marks',
        arguments: {
          from: 1,
          to: 6,
          marks: {
            bold: false,
            italic: false,
            strike: false,
            code: false,
            link: null,
          },
        },
      },
      services,
    )
    expect(editor.getJSON().content?.[0]?.content?.[0]?.marks).toBeUndefined()
  })

  it('rejects incompatible inline-code mark combinations without mutating', async () => {
    const editor = createEditor('Hello')

    await expect(
      executeMarkdownOperation(
        editor,
        {
          operation: 'markdown.text.set_marks',
          arguments: {
            from: 1,
            to: 6,
            marks: {
              bold: true,
              italic: false,
              strike: false,
              code: true,
              link: null,
            },
          },
        },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'markdown.text.set_marks',
      ok: false,
      error: 'execution_failed',
      message: 'Markdown inline code cannot be combined with bold, italic, strike, or link.',
    })
    expect(editor.getMarkdown()).toContain('Hello')
    expect(editor.getMarkdown()).not.toContain('`Hello`')
  })

  it('rejects an invalid Markdown mark range without mutating', async () => {
    const editor = createEditor('Hello')

    await expect(
      executeMarkdownOperation(
        editor,
        {
          operation: 'markdown.text.set_marks',
          arguments: {
            from: 1,
            to: 99,
            marks: {
              bold: true,
              italic: false,
              strike: false,
              code: false,
              link: null,
            },
          },
        },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'markdown.text.set_marks',
      ok: false,
      error: 'execution_failed',
      message: 'Markdown mark range 1..99 is invalid for document size 7.',
    })
  })

  it('sets one addressed Markdown text block type with native undo and reopen fidelity', async () => {
    const editor = createEditor('First\n\nSecond')

    await expect(
      executeMarkdownOperation(
        editor,
        {
          operation: 'markdown.block.set_type',
          arguments: { textBlockIndex: 1, type: 'heading_2' },
        },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'markdown.block.set_type',
      ok: true,
    })
    expect(editor.getMarkdown()).toContain('## Second')

    const reopened = createEditor(editor.getMarkdown())
    expect(reopened.getJSON().content?.[1]).toMatchObject({
      type: 'heading',
      attrs: { level: 2 },
    })

    expect(editor.commands.undo()).toBe(true)
    expect(editor.getMarkdown()).not.toContain('## Second')
  })

  it('normalizes an addressed quote to an explicit paragraph final state', async () => {
    const editor = createEditor('> Quoted')

    await expect(
      executeMarkdownOperation(
        editor,
        {
          operation: 'markdown.block.set_type',
          arguments: { textBlockIndex: 0, type: 'paragraph' },
        },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'markdown.block.set_type',
      ok: true,
    })
    expect(editor.getMarkdown()).toContain('Quoted')
    expect(editor.getMarkdown()).not.toContain('> Quoted')
  })

  it('converts one addressed list item without changing its sibling content', async () => {
    const editor = createEditor('- First\n- Second')

    await expect(
      executeMarkdownOperation(
        editor,
        {
          operation: 'markdown.block.set_type',
          arguments: { textBlockIndex: 1, type: 'heading_1' },
        },
        services,
      ),
    ).resolves.toMatchObject({ handled: true, ok: true })
    expect(editor.getMarkdown()).toContain('- First')
    expect(editor.getMarkdown()).toContain('# Second')
  })

  it('rejects an out-of-range Markdown text block address without mutating', async () => {
    const editor = createEditor('Only')

    await expect(
      executeMarkdownOperation(
        editor,
        {
          operation: 'markdown.block.set_type',
          arguments: { textBlockIndex: 1, type: 'code_block' },
        },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'markdown.block.set_type',
      ok: false,
      error: 'execution_failed',
      message: 'Markdown text block 1 is invalid for 1 addressable text block(s).',
    })
    expect(editor.getMarkdown()).toContain('Only')
  })

  it('shares the native TipTap history stack through markdown.history.undo and redo', async () => {
    const editor = createEditor('Hello')
    editor.commands.setTextSelection(6)
    editor.commands.insertContent(' world')

    await expect(
      executeMarkdownOperation(
        editor,
        { operation: 'markdown.history.undo', arguments: {} },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'markdown.history.undo',
      ok: true,
    })
    expect(editor.getMarkdown()).toContain('Hello')
    expect(editor.getMarkdown()).not.toContain(' world')

    await expect(
      executeMarkdownOperation(
        editor,
        { operation: 'markdown.history.redo', arguments: {} },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'markdown.history.redo',
      ok: true,
    })
    expect(editor.getMarkdown()).toContain('Hello world')
  })

  it('reports unavailable Markdown history entries without mutating', async () => {
    const editor = createEditor('Hello')

    await expect(
      executeMarkdownOperation(
        editor,
        { operation: 'markdown.history.undo', arguments: {} },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'markdown.history.undo',
      ok: false,
      error: 'execution_failed',
      message: 'markdown.history.undo requires an available undo entry.',
    })
    expect(editor.getMarkdown()).toContain('Hello')
  })

  it('loads staged Markdown bytes through markdown.document.load_staged without checkpointing recovery', async () => {
    const editor = createEditor('Old content')
    const data = new TextEncoder().encode('# Loaded').buffer as ArrayBuffer
    let loadedName: string | null = null
    let loadedAssetRoot: string | null = null

    await expect(
      executeMarkdownOperation(
        editor,
        {
          operation: 'markdown.document.load_staged',
          arguments: {
            blobId: 'markdown-blob',
            name: 'loaded.md',
            size: data.byteLength,
            assetRootId: 'asset-root-1',
            data,
          },
        },
        {
          ...services,
          save: async () => ({ ok: true, fileName: 'notes.md' }),
          loadStaged: async ({ name, data: stagedData, assetRootId }) => {
            loadedName = name
            loadedAssetRoot = assetRootId ?? null
            editor.commands.setContent(new TextDecoder().decode(stagedData), {
              contentType: 'markdown',
            })
          },
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'markdown.document.load_staged',
      ok: true,
      checkpointRecovery: false,
    })
    expect(loadedName).toBe('loaded.md')
    expect(loadedAssetRoot).toBe('asset-root-1')
    expect(editor.getMarkdown()).toContain('# Loaded')
  })

  it('rejects the retired open_local_file transport alias without loading Markdown', async () => {
    const editor = createEditor('Old content')
    const data = new TextEncoder().encode('Alias content').buffer as ArrayBuffer

    await expect(
      executeMarkdownOperation(
        editor,
        {
          operation: 'open_local_file',
          arguments: {
            blobId: 'markdown-blob',
            name: 'alias.md',
            size: data.byteLength,
            data,
          },
        },
        {
          ...services,
          save: async () => ({ ok: true, fileName: 'notes.md' }),
          loadStaged: async ({ data: stagedData }) => {
            editor.commands.setContent(new TextDecoder().decode(stagedData), {
              contentType: 'markdown',
            })
          },
        },
      ),
    ).resolves.toEqual({ handled: false })
    expect(editor.getMarkdown()).toContain('Old content')
  })

  it('rejects a staged Markdown load without hydrated ArrayBuffer bytes', async () => {
    const editor = createEditor('Old content')
    let loaded = false

    await expect(
      executeMarkdownOperation(
        editor,
        {
          operation: 'markdown.document.load_staged',
          arguments: {
            blobId: 'markdown-blob',
            name: 'invalid.md',
            size: 1,
            data: {},
          },
        },
        {
          ...services,
          save: async () => ({ ok: true, fileName: 'notes.md' }),
          loadStaged: async () => {
            loaded = true
          },
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'markdown.document.load_staged',
      ok: false,
      error: 'invalid_arguments',
      message: '$.data must be a hydrated ArrayBuffer.',
    })
    expect(loaded).toBe(false)
    expect(editor.getMarkdown()).toContain('Old content')
  })

  it('saves through markdown.document.save and returns the persisted file name', async () => {
    const editor = createEditor('Hello')

    await expect(
      executeMarkdownOperation(
        editor,
        {
          operation: 'markdown.document.save',
          arguments: {},
        },
        {
          ...services,
          loadStaged: async () => undefined,
          save: async () => ({ ok: true, fileName: 'notes.md' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'markdown.document.save',
      ok: true,
      output: { saved: true, fileName: 'notes.md' },
    })
  })

  it('saves through markdown.document.save_as using the shared forced-picker path', async () => {
    const editor = createEditor('Hello')
    const saveRequests: boolean[] = []

    await expect(
      executeMarkdownOperation(
        editor,
        {
          operation: 'markdown.document.save_as',
          arguments: {},
        },
        {
          ...services,
          save: async ({ saveAs }) => {
            saveRequests.push(saveAs)
            return { ok: true, fileName: 'copy.md' }
          },
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'markdown.document.save_as',
      ok: true,
      output: { saved: true, fileName: 'copy.md' },
    })
    expect(saveRequests).toEqual([true])
  })

  it('reports a canceled Markdown Save As as an execution failure', async () => {
    const editor = createEditor('Hello')

    await expect(
      executeMarkdownOperation(
        editor,
        { operation: 'markdown.document.save_as', arguments: {} },
        {
          ...services,
          save: async () => ({ ok: false }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'markdown.document.save_as',
      ok: false,
      error: 'execution_failed',
      message: 'The Markdown Save As was canceled or failed.',
    })
  })

  it('exports DOCX through markdown.document.export_docx using the shared exporter', async () => {
    const editor = createEditor('# Hello')
    let exports = 0

    await expect(
      executeMarkdownOperation(
        editor,
        { operation: 'markdown.document.export_docx', arguments: {} },
        {
          ...services,
          exportDocx: async () => {
            exports += 1
            return { ok: true, fileName: 'notes.docx' }
          },
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'markdown.document.export_docx',
      ok: true,
      output: { exported: true, fileName: 'notes.docx' },
    })
    expect(exports).toBe(1)
  })

  it('reports a failed Markdown DOCX export explicitly', async () => {
    const editor = createEditor('# Hello')

    await expect(
      executeMarkdownOperation(
        editor,
        { operation: 'markdown.document.export_docx', arguments: {} },
        {
          ...services,
          exportDocx: async () => ({ ok: false }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'markdown.document.export_docx',
      ok: false,
      error: 'execution_failed',
      message: 'The Markdown DOCX export failed.',
    })
  })

  it('opens the Markdown print dialog through the shared host action', async () => {
    const editor = createEditor('# Hello')
    let opened = 0

    await expect(
      executeMarkdownOperation(
        editor,
        { operation: 'markdown.document.open_print_dialog', arguments: {} },
        {
          ...services,
          openPrintDialog: () => {
            opened += 1
            return { ok: true }
          },
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'markdown.document.open_print_dialog',
      ok: true,
      output: { opened: true },
    })
    expect(opened).toBe(1)
  })

  it('reports a host-blocked Markdown print dialog explicitly', async () => {
    const editor = createEditor('# Hello')

    await expect(
      executeMarkdownOperation(
        editor,
        { operation: 'markdown.document.open_print_dialog', arguments: {} },
        { ...services, openPrintDialog: () => ({ ok: false }) },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'markdown.document.open_print_dialog',
      ok: false,
      error: 'execution_failed',
      message: 'The host blocked the Markdown print dialog.',
    })
  })

  it('sets the Markdown autosave preference to an explicit final state', async () => {
    const editor = createEditor('# Hello')
    const states: boolean[] = []

    await expect(
      executeMarkdownOperation(
        editor,
        { operation: 'markdown.document.set_auto_save', arguments: { enabled: true } },
        {
          ...services,
          setAutoSave: ({ enabled }) => states.push(enabled),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'markdown.document.set_auto_save',
      ok: true,
      output: { enabled: true },
      checkpointRecovery: false,
    })
    expect(states).toEqual([true])
  })

  it('reports a canceled Markdown save as an execution failure', async () => {
    const editor = createEditor('Hello')

    await expect(
      executeMarkdownOperation(
        editor,
        {
          operation: 'markdown.document.save',
          arguments: {},
        },
        {
          ...services,
          loadStaged: async () => undefined,
          save: async () => ({ ok: false }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'markdown.document.save',
      ok: false,
      error: 'execution_failed',
      message: 'The Markdown save was canceled or failed.',
    })
  })

  it('executes markdown.text.insert through the native TipTap undo route', async () => {
    const editor = createEditor('Hello')
    editor.commands.setTextSelection(6)

    expect(
      await executeMarkdownOperation(
        editor,
        {
          operation: 'markdown.text.insert',
          arguments: { text: ' world' },
        },
        services,
      ),
    ).toEqual({
      handled: true,
      operationId: 'markdown.text.insert',
      ok: true,
    })
    expect(editor.getMarkdown()).toContain('Hello world')

    expect(editor.commands.undo()).toBe(true)
    expect(editor.getMarkdown()).toContain('Hello')
    expect(editor.getMarkdown()).not.toContain(' world')
  })

  it('replaces the active selection through markdown.text.replace_selection and native undo', async () => {
    const editor = createEditor('Hello world')
    editor.commands.setTextSelection({ from: 7, to: 12 })

    expect(
      await executeMarkdownOperation(
        editor,
        {
          operation: 'markdown.text.replace_selection',
          arguments: { text: 'Markdown' },
        },
        services,
      ),
    ).toEqual({
      handled: true,
      operationId: 'markdown.text.replace_selection',
      ok: true,
    })
    expect(editor.getMarkdown()).toContain('Hello Markdown')

    expect(editor.commands.undo()).toBe(true)
    expect(editor.getMarkdown()).toContain('Hello world')
    expect(editor.getMarkdown()).not.toContain('Hello Markdown')
  })

  it('rejects retired public Markdown aliases without mutating the mounted editor', async () => {
    const editor = createEditor('Hello world')

    for (const [operation, arguments_] of [
      ['insert_text', { text: ' alias' }],
      ['replace_selection', { text: 'alias' }],
      ['save', {}],
    ] as const) {
      await expect(
        executeMarkdownOperation(editor, { operation, arguments: arguments_ }, services),
      ).resolves.toEqual({ handled: false })
    }
    expect(editor.getMarkdown()).toContain('Hello world')
  })

  it('rejects non-string text through the descriptor schema without mutating', async () => {
    const editor = createEditor('Hello')

    expect(
      await executeMarkdownOperation(
        editor,
        {
          operation: 'markdown.text.insert',
          arguments: { text: 42 },
        },
        services,
      ),
    ).toEqual({
      handled: true,
      operationId: 'markdown.text.insert',
      ok: false,
      error: 'invalid_arguments',
      message: '$.text must be a string.',
    })
    expect(editor.getMarkdown()).toContain('Hello')
  })
})
