import { Editor } from '@tiptap/core'
import {
  generateTocFieldXml,
  PAGE_MARK,
  TOTAL_PAGES_MARK,
  type NoteInfo,
  type SourceInfo,
} from '@genoffice/docx-engine'
import { afterEach, describe, expect, it } from 'vitest'

import { editorExtensions } from '../src/renderer/editor/extensions'
import {
  docxBlockTargetSchema,
  docxOperationCatalog,
  docxTextRangeSchema,
} from '../src/renderer/operations/catalog'
import { executeDocxOperation } from '../src/renderer/operations/registry'

const editors: Editor[] = []
const services = {
  allocateListNumId: (_kind: 'bullet' | 'ordered') => '42',
  loadStaged: async () => undefined,
  save: async () => ({ fileName: 'document.docx' }),
}

afterEach(() => {
  for (const editor of editors.splice(0)) editor.destroy()
})

function createEditor(text: string): Editor {
  return createEditorWithBlocks([text])
}

function createEditorWithBlocks(texts: readonly string[]): Editor {
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: {
      type: 'doc',
      content: texts.map((text, index) => ({
        type: 'docParagraph',
        attrs: { docxIndex: index },
        content: [{ type: 'text', text }],
      })),
    },
  })
  editors.push(editor)
  return editor
}

describe('DOCX operation registry', () => {
  it('declares bounded deterministic DOCX WordArt insertion', () => {
    const descriptor = docxOperationCatalog.operations.find(
      (operation) => operation.id === 'docx.wordart.insert',
    )
    expect(descriptor).toMatchObject({
      family: 'wordart',
      visibility: 'agent',
      inputSchema: {
        properties: {
          afterBlockIndex: { type: 'integer', minimum: -1, maximum: 2_147_483_647 },
          preset: { type: 'string' },
          text: { type: 'string', minLength: 1, maxLength: 4096 },
          widthEmu: { type: 'integer', minimum: 9_525, maximum: 20_000_000 },
          heightEmu: { type: 'integer', minimum: 9_525, maximum: 20_000_000 },
          drawingId: { type: 'integer', minimum: 1, maximum: 2_147_483_647 },
        },
        required: ['afterBlockIndex', 'preset', 'text', 'widthEmu', 'heightEmu', 'drawingId'],
        additionalProperties: false,
      },
      outputSchema: {
        required: ['summary', 'blockIndex', 'preset', 'drawingId', 'changed'],
        additionalProperties: false,
      },
      mutates: true,
      undoable: true,
      atomic: true,
    })
  })

  it('executes DOCX WordArt insertion through native history', async () => {
    const editor = createEditor('Original body')
    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.wordart.insert',
          arguments: {
            afterBlockIndex: 0,
            preset: 'blue',
            text: 'Registry WordArt',
            widthEmu: 2_700_000,
            heightEmu: 720_000,
            drawingId: 42,
          },
        },
        services,
      ),
    ).resolves.toMatchObject({
      handled: true,
      operationId: 'docx.wordart.insert',
      ok: true,
      output: { blockIndex: 1, preset: 'blue', drawingId: 42, changed: true },
    })
    expect(editor.commands.undo()).toBe(true)
    expect(editor.state.doc.childCount).toBe(1)
  })

  it('declares explicit-final-state DOCX drop-cap formatting', () => {
    const descriptor = docxOperationCatalog.operations.find(
      (operation) => operation.id === 'docx.paragraph.set_drop_cap',
    )
    expect(descriptor).toMatchObject({
      family: 'paragraph',
      visibility: 'agent',
      inputSchema: {
        properties: {
          blockIndex: { type: 'integer', minimum: 0, maximum: 2_147_483_647 },
          mode: { type: 'string', enum: ['none', 'drop', 'margin'] },
          lines: { type: ['integer', 'null'], minimum: 2, maximum: 10 },
        },
        required: ['blockIndex', 'mode', 'lines'],
        additionalProperties: false,
      },
      outputSchema: {
        required: ['summary', 'blockIndex', 'mode', 'lines', 'changed'],
        additionalProperties: false,
      },
      mutates: true,
      undoable: true,
      atomic: true,
    })
  })

  it('executes DOCX drop-cap formatting through native history', async () => {
    const editor = createEditor('Alpha paragraph')
    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.paragraph.set_drop_cap',
          arguments: { blockIndex: 0, mode: 'drop', lines: 3 },
        },
        services,
      ),
    ).resolves.toMatchObject({
      handled: true,
      operationId: 'docx.paragraph.set_drop_cap',
      ok: true,
      output: { blockIndex: 0, mode: 'drop', lines: 3, changed: true },
    })
    expect(editor.state.doc.child(0).attrs.dropCap).toBe('{"type":"drop","lines":3}')
    expect(editor.commands.undo()).toBe(true)
    expect(editor.state.doc.child(0).attrs.dropCap).toBeNull()
  })

  it('declares deterministic bounded DOCX cover-page insertion', () => {
    const descriptor = docxOperationCatalog.operations.find(
      (operation) => operation.id === 'docx.cover_page.insert',
    )
    expect(descriptor).toMatchObject({
      family: 'cover_page',
      visibility: 'agent',
      inputSchema: {
        properties: {
          preset: {
            type: 'string',
            enum: [
              'classic',
              'banded',
              'boxed',
              'sideline',
              'modern',
              'elegant',
              'minimal',
              'dark',
              'accent',
              'badge',
              'facet',
              'annual',
            ],
          },
          year: { type: 'integer', minimum: 1900, maximum: 9999 },
        },
        required: ['preset', 'title', 'subtitle', 'author', 'date', 'year'],
        additionalProperties: false,
      },
      outputSchema: {
        required: ['summary', 'preset', 'insertedBlocks', 'changed'],
        additionalProperties: false,
      },
      mutates: true,
      undoable: true,
      atomic: true,
    })
  })

  it('executes deterministic DOCX cover-page insertion through native history', async () => {
    const editor = createEditor('Original body')
    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.cover_page.insert',
          arguments: {
            preset: 'classic',
            title: 'Registry Architecture',
            subtitle: 'Typed operations',
            author: 'Agent',
            date: '30 August 2026',
            year: 2026,
          },
        },
        services,
      ),
    ).resolves.toMatchObject({
      handled: true,
      operationId: 'docx.cover_page.insert',
      ok: true,
      output: { preset: 'classic', insertedBlocks: 6, changed: true },
    })
    expect(editor.getText()).toContain('Registry Architecture')
    expect(editor.commands.undo()).toBe(true)
    expect(editor.getText()).toBe('Original body')
  })

  it('declares one bounded masked DOCX document-design contract', () => {
    const descriptor = docxOperationCatalog.operations.find(
      (operation) => operation.id === 'docx.document.set_design',
    )
    expect(descriptor).toMatchObject({
      family: 'document',
      visibility: 'agent',
      inputSchema: {
        properties: {
          fields: {
            type: 'array',
            minItems: 1,
            maxItems: 4,
            items: {
              type: 'string',
              enum: ['pageColor', 'watermark', 'themeFonts', 'themeColors'],
            },
          },
          pageColor: {
            type: ['string', 'null'],
            minLength: 6,
            maxLength: 6,
            pattern: '^[0-9A-F]{6}$',
          },
          watermark: { type: ['string', 'null'], minLength: 1, maxLength: 255 },
          themeFonts: { type: 'object' },
          themeColors: { type: 'object' },
        },
        required: ['fields'],
        additionalProperties: false,
      },
      outputSchema: {
        required: ['summary', 'fields', 'changedFields', 'changed'],
        additionalProperties: false,
      },
      mutates: true,
      undoable: false,
      atomic: true,
    })
  })

  it('executes DOCX document design against shared mounted state', async () => {
    let state: import('../src/renderer/editor/document-design').DocxDocumentDesignState = {
      pageColor: null,
      watermark: null,
      themeFonts: null,
      themeColors: null,
    }
    await expect(
      executeDocxOperation(
        createEditor('Alpha'),
        {
          operation: 'docx.document.set_design',
          arguments: {
            fields: ['pageColor', 'watermark'],
            pageColor: 'FFF9E6',
            watermark: 'DRAFT',
          },
        },
        {
          ...services,
          getDocumentDesign: () => state,
          commitDocumentDesign: (next) => {
            state = next
          },
        },
      ),
    ).resolves.toMatchObject({
      handled: true,
      operationId: 'docx.document.set_design',
      ok: true,
      output: {
        fields: ['pageColor', 'watermark'],
        changedFields: ['pageColor', 'watermark'],
        changed: true,
      },
    })
    expect(state).toMatchObject({ pageColor: 'FFF9E6', watermark: 'DRAFT' })
  })

  it('declares public-path and internal-staged DOCX comparison contracts', () => {
    const publicDescriptor = docxOperationCatalog.operations.find(
      (operation) => operation.id === 'docx.document.compare',
    )
    const internalDescriptor = docxOperationCatalog.operations.find(
      (operation) => operation.id === 'docx.document.compare_staged',
    )
    expect(publicDescriptor).toMatchObject({
      family: 'document',
      visibility: 'agent',
      inputSchema: {
        properties: { path: { type: 'string', minLength: 1, maxLength: 1024 } },
        required: ['path'],
        additionalProperties: false,
      },
      mutates: true,
      undoable: false,
      atomic: true,
    })
    expect(internalDescriptor).toMatchObject({
      family: 'document',
      visibility: 'internal',
      inputSchema: {
        required: ['blobId', 'name', 'size', 'data'],
        additionalProperties: false,
      },
      outputSchema: {
        required: ['summary', 'otherName', 'added', 'removed', 'changed', 'identical'],
        additionalProperties: false,
      },
    })
  })

  it('declares one bounded aggregate DOCX ink lifecycle contract', () => {
    const descriptor = docxOperationCatalog.operations.find(
      (operation) => operation.id === 'docx.ink.apply',
    )
    expect(descriptor).toMatchObject({
      family: 'ink',
      visibility: 'agent',
      inputSchema: {
        properties: {
          action: { type: 'string', enum: ['add', 'delete', 'clear'] },
          annotation: { type: 'object' },
          ids: { type: 'array', minItems: 1, maxItems: 1024 },
        },
        required: ['action'],
        additionalProperties: false,
      },
      outputSchema: {
        required: ['summary', 'action', 'added', 'deleted', 'count', 'changed'],
        additionalProperties: false,
      },
      mutates: true,
      undoable: false,
      atomic: true,
    })
  })

  it('executes DOCX ink addition against shared mounted state', async () => {
    let annotations: import('../src/renderer/editor/ink').InkAnnotation[] = []
    await expect(
      executeDocxOperation(
        createEditor('Alpha'),
        {
          operation: 'docx.ink.apply',
          arguments: {
            action: 'add',
            annotation: {
              id: 'ink-agent-1',
              anchorIndex: 0,
              tool: 'pen',
              color: 'C00000',
              width: 2,
              points: [{ x: 1, y: 2 }],
            },
          },
        },
        {
          ...services,
          getInkAnnotations: () => annotations,
          commitInkAnnotations: (next) => {
            annotations = [...next]
          },
        },
      ),
    ).resolves.toMatchObject({
      handled: true,
      operationId: 'docx.ink.apply',
      ok: true,
      output: { action: 'add', added: 1, deleted: 0, count: 1, changed: true },
    })
    expect(annotations).toHaveLength(1)
  })

  it('declares one explicit-final-state DOCX document-protection contract', () => {
    const descriptor = docxOperationCatalog.operations.find(
      (operation) => operation.id === 'docx.document.set_protection',
    )
    expect(descriptor).toMatchObject({
      family: 'document',
      visibility: 'agent',
      inputSchema: {
        properties: {
          enabled: { type: 'boolean' },
          password: { type: ['string', 'null'], minLength: 1, maxLength: 255 },
        },
        required: ['enabled', 'password'],
        additionalProperties: false,
      },
      outputSchema: {
        required: ['summary', 'enabled', 'passwordProtected', 'changed'],
        additionalProperties: false,
      },
      mutates: true,
      undoable: false,
      atomic: true,
    })
  })

  it('executes DOCX protection against shared mounted document state', async () => {
    let protection: import('@genoffice/docx-engine').DocProtection | null = null
    await expect(
      executeDocxOperation(
        createEditor('Alpha'),
        {
          operation: 'docx.document.set_protection',
          arguments: { enabled: true, password: null },
        },
        {
          ...services,
          getProtection: () => protection,
          commitProtection: (next) => {
            protection = next
          },
        },
      ),
    ).resolves.toMatchObject({
      handled: true,
      operationId: 'docx.document.set_protection',
      ok: true,
      output: { enabled: true, passwordProtected: false, changed: true },
    })
    expect(protection).toEqual({ edit: 'readOnly', enforced: true })
  })

  it('declares one bounded DOCX revision-decision contract', () => {
    const descriptor = docxOperationCatalog.operations.find(
      (operation) => operation.id === 'docx.revision.apply_decision',
    )
    expect(descriptor).toMatchObject({
      family: 'revision',
      visibility: 'agent',
      inputSchema: {
        properties: {
          decision: { type: 'string', enum: ['accept', 'reject'] },
          scope: { type: 'string', enum: ['current', 'all'] },
        },
        required: ['decision', 'scope'],
        additionalProperties: false,
      },
      outputSchema: {
        required: ['summary', 'decision', 'scope', 'matched', 'remaining', 'changed'],
        additionalProperties: false,
      },
      mutates: true,
      undoable: true,
      atomic: true,
    })
  })

  it('executes one bounded DOCX revision decision against mounted state', async () => {
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
              {
                type: 'text',
                text: 'Alpha',
                marks: [{ type: 'ins', attrs: { author: 'Agent', date: null } }],
              },
            ],
          },
        ],
      },
    })
    editors.push(editor)
    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.revision.apply_decision',
          arguments: { decision: 'accept', scope: 'all' },
        },
        services,
      ),
    ).resolves.toMatchObject({
      handled: true,
      operationId: 'docx.revision.apply_decision',
      ok: true,
      output: {
        decision: 'accept',
        scope: 'all',
        matched: 1,
        remaining: 0,
        changed: true,
      },
    })
  })

  it('declares one explicit-final-state DOCX revision-tracking contract', () => {
    const descriptor = docxOperationCatalog.operations.find(
      (operation) => operation.id === 'docx.revision.set_tracking',
    )
    expect(descriptor).toMatchObject({
      family: 'revision',
      visibility: 'agent',
      inputSchema: {
        properties: { enabled: { type: 'boolean' } },
        required: ['enabled'],
        additionalProperties: false,
      },
      outputSchema: {
        required: ['summary', 'enabled', 'changed'],
        additionalProperties: false,
      },
      mutates: true,
      undoable: false,
      atomic: true,
    })
  })

  it('executes DOCX revision tracking against the mounted recorder and shared UI state', async () => {
    const editor = createEditor('Alpha')
    let committed: boolean | null = null
    await expect(
      executeDocxOperation(
        editor,
        { operation: 'docx.revision.set_tracking', arguments: { enabled: true } },
        {
          ...services,
          commitRevisionTracking: (enabled: boolean) => {
            committed = enabled
          },
        },
      ),
    ).resolves.toMatchObject({
      handled: true,
      operationId: 'docx.revision.set_tracking',
      ok: true,
      output: { enabled: true, changed: true },
      checkpointRecovery: false,
    })
    expect((editor.storage.trackChanges as { enabled: boolean }).enabled).toBe(true)
    expect(committed).toBe(true)
  })

  it('declares one stable-id DOCX comment deletion contract', () => {
    const descriptor = docxOperationCatalog.operations.find(
      (operation) => operation.id === 'docx.comment.delete',
    )
    expect(descriptor).toMatchObject({
      family: 'comment',
      visibility: 'agent',
      inputSchema: {
        properties: {
          id: { type: 'string', minLength: 1, maxLength: 10, pattern: '^[0-9]+$' },
        },
        required: ['id'],
        additionalProperties: false,
      },
      outputSchema: {
        required: ['summary', 'id', 'deleted', 'anchors', 'changed'],
        additionalProperties: false,
      },
      mutates: true,
      undoable: true,
      atomic: true,
    })
  })

  it('executes stable-id DOCX comment deletion against mounted state', async () => {
    const editor = createEditor('Alpha')
    editor.commands.setTextSelection({ from: 1, to: 6 })
    const { addCommentToSelection } = await import('../src/renderer/editor/comments')
    addCommentToSelection(editor, '1')
    await expect(
      executeDocxOperation(
        editor,
        { operation: 'docx.comment.delete', arguments: { id: '1' } },
        {
          ...services,
          getComments: () => [{ id: '1', author: 'User', text: 'Parent' }],
          commitComments: () => undefined,
        },
      ),
    ).resolves.toMatchObject({
      handled: true,
      operationId: 'docx.comment.delete',
      ok: true,
      output: { id: '1', deleted: 1, anchors: 1, changed: true },
    })
  })

  it('declares one explicit-final-state DOCX comment resolution contract', () => {
    const descriptor = docxOperationCatalog.operations.find(
      (operation) => operation.id === 'docx.comment.set_resolved',
    )
    expect(descriptor).toMatchObject({
      family: 'comment',
      visibility: 'agent',
      inputSchema: {
        properties: {
          id: { type: 'string', minLength: 1, maxLength: 10, pattern: '^[0-9]+$' },
          resolved: { type: 'boolean' },
        },
        required: ['id', 'resolved'],
        additionalProperties: false,
      },
      outputSchema: {
        required: ['summary', 'id', 'resolved', 'affected', 'changed'],
        additionalProperties: false,
      },
      mutates: true,
      undoable: true,
      atomic: true,
    })
  })

  it('executes explicit DOCX comment resolution against mounted state', async () => {
    const editor = createEditor('Alpha')
    await expect(
      executeDocxOperation(
        editor,
        { operation: 'docx.comment.set_resolved', arguments: { id: '1', resolved: true } },
        {
          ...services,
          getComments: () => [{ id: '1', author: 'User', text: 'Parent' }],
          commitComments: () => undefined,
        },
      ),
    ).resolves.toMatchObject({
      handled: true,
      operationId: 'docx.comment.set_resolved',
      ok: true,
      output: { id: '1', resolved: true, affected: 1, changed: true },
    })
  })

  it('declares one stable-parent bounded DOCX comment-reply contract', () => {
    const descriptor = docxOperationCatalog.operations.find(
      (operation) => operation.id === 'docx.comment.reply',
    )
    expect(descriptor).toMatchObject({
      family: 'comment',
      visibility: 'agent',
      inputSchema: {
        properties: {
          parentId: { type: 'string', minLength: 1, maxLength: 10, pattern: '^[0-9]+$' },
          comment: {
            type: 'object',
            required: ['id', 'author', 'initials', 'date', 'text'],
            additionalProperties: false,
          },
        },
        required: ['parentId', 'comment'],
        additionalProperties: false,
      },
      outputSchema: {
        required: ['summary', 'id', 'parentId', 'references', 'changed'],
        additionalProperties: false,
      },
      mutates: true,
      undoable: true,
      atomic: true,
    })
  })

  it('executes DOCX comment reply against mounted parent state', async () => {
    const editor = createEditor('Alpha')
    editor.commands.setTextSelection({ from: 1, to: 6 })
    const parent = {
      id: '1',
      author: 'User',
      date: '2026-08-30T02:00:00Z',
      text: 'Parent',
    }
    const { addCommentToSelection } = await import('../src/renderer/editor/comments')
    addCommentToSelection(editor, '1')
    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.comment.reply',
          arguments: {
            parentId: '1',
            comment: {
              id: '2',
              author: 'Agent',
              initials: null,
              date: '2026-08-30T02:01:00Z',
              text: 'Reply',
            },
          },
        },
        {
          ...services,
          getComments: () => [parent],
          commitComments: () => undefined,
        },
      ),
    ).resolves.toMatchObject({
      handled: true,
      operationId: 'docx.comment.reply',
      ok: true,
      output: { id: '2', parentId: '1', references: 1, changed: true },
    })
  })

  it('declares one exact-range bounded DOCX comment-add contract', () => {
    const descriptor = docxOperationCatalog.operations.find(
      (operation) => operation.id === 'docx.comment.add',
    )
    expect(descriptor).toMatchObject({
      family: 'comment',
      visibility: 'agent',
      inputSchema: {
        properties: {
          range: docxTextRangeSchema,
          comment: {
            type: 'object',
            properties: {
              id: { type: 'string', minLength: 1, maxLength: 10, pattern: '^[0-9]+$' },
              author: { type: 'string', minLength: 1, maxLength: 255 },
              initials: { type: ['string', 'null'], maxLength: 16 },
              date: {
                type: 'string',
                minLength: 20,
                maxLength: 32,
                pattern: '^\\d{4}-\\d{2}-\\d{2}T\\d{2}:\\d{2}:\\d{2}(?:\\.\\d{1,3})?Z$',
              },
              text: { type: 'string', minLength: 1, maxLength: 65_536 },
            },
            required: ['id', 'author', 'initials', 'date', 'text'],
            additionalProperties: false,
          },
        },
        required: ['range', 'comment'],
        additionalProperties: false,
      },
      outputSchema: {
        required: ['summary', 'id', 'from', 'to', 'changed'],
        additionalProperties: false,
      },
      mutates: true,
      undoable: true,
      atomic: true,
    })
  })

  it('executes exact-range DOCX comment addition against mounted comment state', async () => {
    const editor = createEditor('Alpha')
    const committed: unknown[] = []
    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.comment.add',
          arguments: {
            range: { from: 1, to: 6 },
            comment: {
              id: '1',
              author: 'Agent',
              initials: null,
              date: '2026-08-30T02:00:00Z',
              text: 'Review this text.',
            },
          },
        },
        {
          ...services,
          getComments: () => [],
          commitComments: (comments) => committed.push(...comments),
        },
      ),
    ).resolves.toMatchObject({
      handled: true,
      operationId: 'docx.comment.add',
      ok: true,
      output: { id: '1', from: 1, to: 6, changed: true },
    })
    expect(committed).toHaveLength(1)
  })

  it('declares one stable-boundary bounded DOCX index insertion contract', () => {
    const descriptor = docxOperationCatalog.operations.find(
      (operation) => operation.id === 'docx.index.insert',
    )
    expect(descriptor).toMatchObject({
      family: 'index',
      visibility: 'agent',
      inputSchema: {
        properties: {
          afterBlockIndex: { type: 'integer', minimum: -1, maximum: 2_147_483_647 },
          label: { type: 'string', minLength: 1, maxLength: 255 },
          terms: {
            type: 'array',
            minItems: 1,
            maxItems: 1024,
            items: { type: 'string', minLength: 1, maxLength: 4096 },
          },
        },
        required: ['afterBlockIndex', 'label', 'terms'],
        additionalProperties: false,
      },
      outputSchema: {
        required: ['summary', 'afterBlockIndex', 'entries', 'insertedBlocks'],
        additionalProperties: false,
      },
      mutates: true,
      undoable: true,
      atomic: true,
    })
  })

  it('executes explicit index insertion at a stable boundary', async () => {
    const editor = createEditor('Alpha')
    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.index.insert',
          arguments: { afterBlockIndex: 0, label: 'Index', terms: ['Beta', 'Alpha'] },
        },
        services,
      ),
    ).resolves.toMatchObject({
      handled: true,
      operationId: 'docx.index.insert',
      ok: true,
      output: { afterBlockIndex: 0, entries: 2, insertedBlocks: 2 },
    })
  })

  it('declares one exact-range bounded DOCX index-mark contract', () => {
    const descriptor = docxOperationCatalog.operations.find(
      (operation) => operation.id === 'docx.index.mark',
    )
    expect(descriptor).toMatchObject({
      family: 'index',
      visibility: 'agent',
      inputSchema: {
        properties: {
          range: docxTextRangeSchema,
          term: {
            type: 'string',
            minLength: 1,
            maxLength: 4096,
            pattern: '^[^"\\u0000-\\u001F]+$',
          },
        },
        required: ['range', 'term'],
        additionalProperties: false,
      },
      outputSchema: {
        required: ['summary', 'from', 'to', 'term', 'changed'],
        additionalProperties: false,
      },
      mutates: true,
      undoable: true,
      atomic: true,
    })
  })

  it('executes exact-range DOCX index marking', async () => {
    const editor = createEditor('Alpha')
    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.index.mark',
          arguments: { range: { from: 1, to: 6 }, term: 'Alpha' },
        },
        services,
      ),
    ).resolves.toMatchObject({
      handled: true,
      operationId: 'docx.index.mark',
      ok: true,
      output: { from: 1, to: 6, term: 'Alpha', changed: true },
    })
  })

  it('declares one stable-boundary bounded DOCX caption insertion contract', () => {
    const descriptor = docxOperationCatalog.operations.find(
      (operation) => operation.id === 'docx.caption.insert',
    )
    expect(descriptor).toMatchObject({
      family: 'caption',
      visibility: 'agent',
      inputSchema: {
        properties: {
          afterBlockIndex: { type: 'integer', minimum: -1, maximum: 2_147_483_647 },
          label: { type: 'string', minLength: 1, maxLength: 255 },
          number: { type: 'integer', minimum: 1, maximum: 2_147_483_647 },
          text: { type: 'string', maxLength: 4096 },
        },
        required: ['afterBlockIndex', 'label', 'number', 'text'],
        additionalProperties: false,
      },
      outputSchema: {
        required: ['summary', 'afterBlockIndex', 'label', 'number', 'changed'],
        additionalProperties: false,
      },
      mutates: true,
      undoable: true,
      atomic: true,
    })
  })

  it('executes explicit caption insertion at a stable boundary', async () => {
    const editor = createEditor('Alpha')
    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.caption.insert',
          arguments: {
            afterBlockIndex: 0,
            label: 'Figure',
            number: 2,
            text: 'Registry architecture',
          },
        },
        services,
      ),
    ).resolves.toMatchObject({
      handled: true,
      operationId: 'docx.caption.insert',
      ok: true,
      output: { afterBlockIndex: 0, label: 'Figure', number: 2, changed: true },
    })
  })

  it('declares one stable-boundary bounded DOCX bibliography insertion contract', () => {
    const descriptor = docxOperationCatalog.operations.find(
      (operation) => operation.id === 'docx.bibliography.insert',
    )
    expect(descriptor).toMatchObject({
      family: 'bibliography',
      visibility: 'agent',
      inputSchema: {
        properties: {
          afterBlockIndex: { type: 'integer', minimum: -1, maximum: 2_147_483_647 },
          heading: { type: 'string', minLength: 1, maxLength: 4096 },
          entries: {
            type: 'array',
            minItems: 1,
            maxItems: 1024,
            items: {
              type: 'object',
              properties: {
                sourceTag: { type: 'string', minLength: 1, maxLength: 255 },
                text: { type: 'string', minLength: 1, maxLength: 4096 },
              },
              required: ['sourceTag', 'text'],
              additionalProperties: false,
            },
          },
        },
        required: ['afterBlockIndex', 'heading', 'entries'],
        additionalProperties: false,
      },
      outputSchema: {
        required: ['summary', 'afterBlockIndex', 'entries', 'insertedBlocks'],
        additionalProperties: false,
      },
      mutates: true,
      undoable: true,
      atomic: true,
    })
  })

  it('executes bibliography insertion against mounted source state', async () => {
    const editor = createEditor('Alpha')
    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.bibliography.insert',
          arguments: {
            afterBlockIndex: 0,
            heading: 'Bibliography',
            entries: [{ sourceTag: 'Wang2026', text: 'Wang, Wei. (2026). Registry.' }],
          },
        },
        {
          ...services,
          getSources: () => [
            {
              tag: 'Wang2026',
              type: 'Book',
              author: 'Wang, Wei',
              title: 'Registry',
              year: '2026',
            },
          ],
        },
      ),
    ).resolves.toMatchObject({
      handled: true,
      operationId: 'docx.bibliography.insert',
      ok: true,
      output: { afterBlockIndex: 0, entries: 1, insertedBlocks: 2 },
    })
  })

  it('declares one exact-range bounded DOCX citation insertion contract', () => {
    const descriptor = docxOperationCatalog.operations.find(
      (operation) => operation.id === 'docx.citation.insert',
    )
    expect(descriptor).toMatchObject({
      family: 'citation',
      visibility: 'agent',
      inputSchema: {
        properties: {
          range: docxTextRangeSchema,
          sourceTag: { type: 'string', minLength: 1, maxLength: 255 },
          displayText: { type: 'string', minLength: 1, maxLength: 4096 },
        },
        required: ['range', 'sourceTag', 'displayText'],
        additionalProperties: false,
      },
      outputSchema: {
        required: ['summary', 'from', 'to', 'sourceTag', 'changed'],
        additionalProperties: false,
      },
      mutates: true,
      undoable: true,
      atomic: true,
    })
  })

  it('executes citation insertion against mounted source state', async () => {
    const editor = createEditor('Alpha')
    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.citation.insert',
          arguments: {
            range: { from: 6, to: 6 },
            sourceTag: 'Wang2026',
            displayText: '(Wang, Wei, 2026)',
          },
        },
        {
          ...services,
          getSources: () => [
            {
              tag: 'Wang2026',
              type: 'Book',
              author: 'Wang, Wei',
              title: 'Registry Architecture',
              year: '2026',
            },
          ],
        },
      ),
    ).resolves.toMatchObject({
      handled: true,
      operationId: 'docx.citation.insert',
      ok: true,
      output: { from: 6, to: 6, sourceTag: 'Wang2026', changed: true },
    })
  })

  it('declares one stable-tag bounded DOCX source upsert contract', () => {
    const descriptor = docxOperationCatalog.operations.find(
      (operation) => operation.id === 'docx.source.upsert',
    )
    expect(descriptor).toMatchObject({
      family: 'source',
      visibility: 'agent',
      inputSchema: {
        properties: {
          source: {
            type: 'object',
            properties: {
              tag: { type: 'string', minLength: 1, maxLength: 255 },
              type: {
                type: 'string',
                enum: ['Book', 'JournalArticle', 'InternetSite', 'Report', 'Misc'],
              },
              author: { type: 'string', maxLength: 4096 },
              title: { type: 'string', minLength: 1, maxLength: 4096 },
              year: { type: 'string', maxLength: 32 },
              publisher: { type: ['string', 'null'], maxLength: 4096 },
              url: { type: ['string', 'null'], maxLength: 4096 },
            },
            required: ['tag', 'type', 'author', 'title', 'year', 'publisher', 'url'],
            additionalProperties: false,
          },
        },
        required: ['source'],
        additionalProperties: false,
      },
      outputSchema: {
        required: ['summary', 'tag', 'created', 'changed'],
        additionalProperties: false,
      },
      mutates: true,
      undoable: true,
      atomic: true,
    })
  })

  it('executes source upsert and commits shared source state', async () => {
    const editor = createEditor('Alpha')
    const source: SourceInfo = {
      tag: 'Wang2026',
      type: 'Book',
      author: 'Wang, Wei',
      title: 'Registry Architecture',
      year: '2026',
    }
    const committed: SourceInfo[][] = []
    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.source.upsert',
          arguments: { source: { ...source, publisher: null, url: null } },
        },
        {
          ...services,
          getSources: () => [],
          commitSources: (sources) => committed.push([...sources]),
        },
      ),
    ).resolves.toMatchObject({
      handled: true,
      operationId: 'docx.source.upsert',
      ok: true,
      output: { tag: 'Wang2026', created: true, changed: true },
    })
    expect(committed).toEqual([[source]])
  })

  it('declares one stable-id bounded DOCX note deletion contract', () => {
    const descriptor = docxOperationCatalog.operations.find(
      (operation) => operation.id === 'docx.note.delete',
    )
    expect(descriptor).toMatchObject({
      family: 'note',
      visibility: 'agent',
      inputSchema: {
        properties: {
          kind: { type: 'string', enum: ['footnote', 'endnote'] },
          noteId: { type: 'integer', minimum: 1, maximum: 2_147_483_647 },
        },
        required: ['kind', 'noteId'],
        additionalProperties: false,
      },
      outputSchema: {
        required: ['summary', 'kind', 'noteId', 'references', 'renumbered', 'changed'],
        additionalProperties: false,
      },
      mutates: true,
      undoable: true,
      atomic: true,
    })
  })

  it('executes stable note deletion through the shared state service', async () => {
    const editor = createEditor('Alpha')
    editor.commands.insertContentAt(6, {
      type: 'docNoteRef',
      attrs: { kind: 'endnote', id: '2', num: 1 },
    })
    const original = { id: '2', text: 'Original' }
    const committed: Array<{
      kind: 'footnote' | 'endnote'
      id: string
      original: NoteInfo
    }> = []
    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.note.delete',
          arguments: { kind: 'endnote', noteId: 2 },
        },
        {
          ...services,
          getNotes: () => [original],
          commitDeletedNote: (kind, id, note) => committed.push({ kind, id, original: note }),
        },
      ),
    ).resolves.toMatchObject({
      handled: true,
      operationId: 'docx.note.delete',
      ok: true,
      output: {
        kind: 'endnote',
        noteId: '2',
        references: 1,
        renumbered: 0,
        changed: true,
      },
    })
    expect(committed).toEqual([{ kind: 'endnote', id: '2', original }])
  })

  it('declares one explicit bounded DOCX note insertion contract', () => {
    const descriptor = docxOperationCatalog.operations.find(
      (operation) => operation.id === 'docx.note.insert',
    )
    expect(descriptor).toMatchObject({
      family: 'note',
      visibility: 'agent',
      inputSchema: {
        properties: {
          range: docxTextRangeSchema,
          kind: { type: 'string', enum: ['footnote', 'endnote'] },
          noteId: { type: 'integer', minimum: 1, maximum: 2_147_483_647 },
          text: { type: 'string', minLength: 1, maxLength: 65_536 },
        },
        required: ['range', 'kind', 'noteId', 'text'],
        additionalProperties: false,
      },
      outputSchema: {
        required: ['summary', 'from', 'to', 'kind', 'noteId', 'number', 'changed'],
        additionalProperties: false,
      },
      mutates: true,
      undoable: true,
      atomic: true,
    })
  })

  it('executes note insertion through the registry and commits its shared note state', async () => {
    const editor = createEditor('Alpha')
    const committed: Array<{ kind: 'footnote' | 'endnote'; id: string }> = []
    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.note.insert',
          arguments: {
            range: { from: 6, to: 6 },
            kind: 'footnote',
            noteId: 2,
            text: 'Registry note',
          },
        },
        {
          ...services,
          getNotes: () => [],
          commitInsertedNote: (kind, id) => committed.push({ kind, id }),
        },
      ),
    ).resolves.toMatchObject({
      handled: true,
      operationId: 'docx.note.insert',
      ok: true,
      output: {
        from: 6,
        to: 6,
        kind: 'footnote',
        noteId: '2',
        number: 1,
        changed: true,
      },
    })
    expect(committed).toEqual([{ kind: 'footnote', id: '2' }])
  })

  it('declares one stable-id bounded DOCX note update contract', () => {
    const descriptor = docxOperationCatalog.operations.find(
      (operation) => operation.id === 'docx.note.update',
    )
    expect(descriptor).toMatchObject({
      family: 'note',
      visibility: 'agent',
      inputSchema: {
        properties: {
          kind: { type: 'string', enum: ['footnote', 'endnote'] },
          noteId: { type: 'integer', minimum: 1, maximum: 2_147_483_647 },
          text: { type: 'string', minLength: 1, maxLength: 65_536 },
        },
        required: ['kind', 'noteId', 'text'],
        additionalProperties: false,
      },
      outputSchema: {
        required: ['summary', 'kind', 'noteId', 'references', 'changed'],
        additionalProperties: false,
      },
      mutates: true,
      undoable: true,
      atomic: true,
    })
  })

  it('executes stable note update through the shared state service', async () => {
    const editor = createEditor('Alpha')
    editor.commands.insertContentAt(6, {
      type: 'docNoteRef',
      attrs: { kind: 'footnote', id: '2', num: 1 },
    })
    const original = { id: '2', text: 'Original' }
    const committed: Array<{
      kind: 'footnote' | 'endnote'
      id: string
      original: NoteInfo
    }> = []
    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.note.update',
          arguments: { kind: 'footnote', noteId: 2, text: 'Updated' },
        },
        {
          ...services,
          getNotes: () => [original],
          commitUpdatedNote: (kind, id, note) => committed.push({ kind, id, original: note }),
        },
      ),
    ).resolves.toMatchObject({
      handled: true,
      operationId: 'docx.note.update',
      ok: true,
      output: {
        kind: 'footnote',
        noteId: '2',
        references: 1,
        changed: true,
      },
    })
    expect(committed).toEqual([{ kind: 'footnote', id: '2', original }])
  })

  it('bounds the shared DOCX text insertion contract', () => {
    const descriptor = docxOperationCatalog.operations.find(
      (operation) => operation.id === 'docx.text.insert',
    )
    expect(descriptor).toMatchObject({
      inputSchema: {
        properties: {
          text: { type: 'string', minLength: 1, maxLength: 65_536 },
        },
        required: ['text'],
        additionalProperties: false,
      },
    })
  })

  it('declares one stable-target bounded DOCX TOC refresh contract', () => {
    const descriptor = docxOperationCatalog.operations.find(
      (operation) => operation.id === 'docx.toc.refresh',
    )
    expect(descriptor).toMatchObject({
      family: 'toc',
      visibility: 'agent',
      inputSchema: {
        properties: {
          tocBlockIndex: { type: 'integer', minimum: 0, maximum: 2_147_483_647 },
          entries: {
            type: 'array',
            minItems: 1,
            maxItems: 1024,
            items: {
              type: 'object',
              properties: {
                level: { type: 'integer', minimum: 1, maximum: 9 },
                text: { type: 'string', minLength: 1, maxLength: 4096 },
                pageNumber: {
                  type: ['integer', 'null'],
                  minimum: 1,
                  maximum: 2_147_483_647,
                },
              },
              required: ['level', 'text', 'pageNumber'],
              additionalProperties: false,
            },
          },
        },
        required: ['tocBlockIndex', 'entries'],
        additionalProperties: false,
      },
      outputSchema: {
        required: ['summary', 'entries', 'replacedBlocks'],
        additionalProperties: false,
      },
      mutates: true,
      undoable: true,
      atomic: true,
    })
  })

  it('refreshes one exact TOC region and undoes it natively', async () => {
    const [oldXml] = generateTocFieldXml([{ level: 1, text: 'Old', pageNo: 1 }])
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docProtected',
            attrs: {
              docxIndex: null,
              blockType: 'passthrough',
              label: 'TOC field',
              genXml: oldXml,
              fieldDisplay: { kind: 'tocLine', left: 'Old', right: '1', level: 1 },
            },
          },
          {
            type: 'docHeading',
            attrs: { docxIndex: null, level: 1 },
            content: [{ type: 'text', text: 'New' }],
          },
        ],
      },
    })
    editors.push(editor)
    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.toc.refresh',
          arguments: {
            tocBlockIndex: 0,
            entries: [{ level: 1, text: 'New', pageNumber: 3 }],
          },
        },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.toc.refresh',
      ok: true,
      output: {
        summary: 'Refreshed DOCX table of contents with 1 entry',
        entries: 1,
        replacedBlocks: 1,
      },
    })
    expect(editor.state.doc.child(0).attrs.fieldDisplay).toMatchObject({
      kind: 'tocLine',
      left: 'New',
      right: '3',
      level: 1,
    })
    expect(editor.commands.undo()).toBe(true)
    expect(editor.state.doc.child(0).attrs.fieldDisplay).toMatchObject({ left: 'Old' })
  })

  it('rejects TOC refresh when the stable target is not a TOC field', async () => {
    const editor = createEditor('Alpha')
    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.toc.refresh',
          arguments: {
            tocBlockIndex: 0,
            entries: [{ level: 1, text: 'Alpha', pageNumber: null }],
          },
        },
        services,
      ),
    ).resolves.toMatchObject({ handled: true, ok: false, error: 'invalid_arguments' })
  })

  it('declares one bounded final-state DOCX bookmark contract', () => {
    const descriptor = docxOperationCatalog.operations.find(
      (operation) => operation.id === 'docx.bookmark.set',
    )

    expect(descriptor).toMatchObject({
      family: 'bookmark',
      visibility: 'agent',
      inputSchema: {
        properties: {
          blockIndex: { type: 'integer', minimum: 0, maximum: 2_147_483_647 },
          name: {
            type: 'string',
            minLength: 1,
            maxLength: 40,
            pattern: '^[A-Za-z_一-鿿][A-Za-z0-9_一-鿿]*$',
          },
          enabled: { type: 'boolean' },
        },
        required: ['blockIndex', 'name', 'enabled'],
        additionalProperties: false,
      },
      outputSchema: {
        required: ['summary', 'blockIndex', 'name', 'enabled', 'changed'],
        additionalProperties: false,
      },
      mutates: true,
      undoable: true,
      atomic: true,
    })
  })

  it('adds and removes a bookmark through one explicit final-state operation and native Undo', async () => {
    const editor = createEditor('Alpha')
    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.bookmark.set',
          arguments: { blockIndex: 0, name: 'Anchor_1', enabled: true },
        },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.bookmark.set',
      ok: true,
      output: {
        summary: 'Enabled DOCX bookmark Anchor_1 on block 0',
        blockIndex: 0,
        name: 'Anchor_1',
        enabled: true,
        changed: true,
      },
    })
    expect(editor.state.doc.child(0).attrs.bookmarks).toEqual(['Anchor_1'])
    expect(editor.commands.undo()).toBe(true)
    expect(editor.state.doc.child(0).attrs.bookmarks).toBeNull()

    await executeDocxOperation(
      editor,
      {
        operation: 'docx.bookmark.set',
        arguments: { blockIndex: 0, name: 'Anchor_1', enabled: true },
      },
      services,
    )
    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.bookmark.set',
          arguments: { blockIndex: 0, name: 'Anchor_1', enabled: false },
        },
        services,
      ),
    ).resolves.toMatchObject({
      handled: true,
      operationId: 'docx.bookmark.set',
      ok: true,
      output: { enabled: false, changed: true },
    })
    expect(editor.state.doc.child(0).attrs.bookmarks).toBeNull()
  })

  it('rejects duplicate bookmark names on different blocks', async () => {
    const editor = createEditorWithBlocks(['Alpha', 'Beta'])
    await executeDocxOperation(
      editor,
      {
        operation: 'docx.bookmark.set',
        arguments: { blockIndex: 0, name: 'Anchor', enabled: true },
      },
      services,
    )
    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.bookmark.set',
          arguments: { blockIndex: 1, name: 'Anchor', enabled: true },
        },
        services,
      ),
    ).resolves.toMatchObject({ handled: true, ok: false, error: 'invalid_arguments' })
  })

  it('declares one exact-range bounded DOCX cross-reference insertion contract', () => {
    const descriptor = docxOperationCatalog.operations.find(
      (operation) => operation.id === 'docx.cross_reference.insert',
    )

    expect(descriptor).toMatchObject({
      family: 'cross_reference',
      visibility: 'agent',
      inputSchema: {
        properties: {
          range: docxTextRangeSchema,
          bookmarkName: {
            type: 'string',
            minLength: 1,
            maxLength: 40,
            pattern: '^[A-Za-z_一-鿿][A-Za-z0-9_一-鿿]*$',
          },
          displayText: { type: 'string', minLength: 1, maxLength: 65_536 },
        },
        required: ['range', 'bookmarkName', 'displayText'],
        additionalProperties: false,
      },
      outputSchema: {
        required: ['summary', 'from', 'to', 'bookmarkName', 'changed'],
        additionalProperties: false,
      },
      mutates: true,
      undoable: true,
      atomic: true,
    })
  })

  it('inserts a cross-reference at an exact range and undoes it natively', async () => {
    const editor = createEditorWithBlocks(['Alpha', 'Beta'])
    await executeDocxOperation(
      editor,
      {
        operation: 'docx.bookmark.set',
        arguments: { blockIndex: 0, name: 'Anchor', enabled: true },
      },
      services,
    )
    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.cross_reference.insert',
          arguments: {
            range: { from: 8, to: 8 },
            bookmarkName: 'Anchor',
            displayText: 'Alpha',
          },
        },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.cross_reference.insert',
      ok: true,
      output: {
        summary: 'Inserted DOCX cross-reference to Anchor over range 8..13',
        from: 8,
        to: 13,
        bookmarkName: 'Anchor',
        changed: true,
      },
    })
    expect(editor.state.doc.nodeAt(8)?.marks[0]).toMatchObject({
      type: { name: 'refField' },
      attrs: { name: 'Anchor' },
    })
    expect(editor.commands.undo()).toBe(true)
    expect(editor.state.doc.child(1).textContent).toBe('Beta')
  })

  it('rejects a cross-reference whose bookmark target does not exist', async () => {
    const editor = createEditor('Alpha')
    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.cross_reference.insert',
          arguments: {
            range: { from: 1, to: 1 },
            bookmarkName: 'Missing',
            displayText: 'Missing',
          },
        },
        services,
      ),
    ).resolves.toMatchObject({ handled: true, ok: false, error: 'invalid_arguments' })
  })

  it('declares one exact-range bounded DOCX generic-field insertion contract', () => {
    const descriptor = docxOperationCatalog.operations.find(
      (operation) => operation.id === 'docx.field.insert',
    )

    expect(descriptor).toMatchObject({
      family: 'field',
      visibility: 'agent',
      inputSchema: {
        properties: {
          range: docxTextRangeSchema,
          instruction: {
            type: 'string',
            enum: ['DATE', 'TIME', 'PAGE', 'NUMPAGES', 'FILENAME'],
          },
          displayText: { type: 'string', minLength: 1, maxLength: 65_536 },
        },
        required: ['range', 'instruction', 'displayText'],
        additionalProperties: false,
      },
      outputSchema: {
        required: ['summary', 'from', 'to', 'instruction', 'changed'],
        additionalProperties: false,
      },
      mutates: true,
      undoable: true,
      atomic: true,
    })
  })

  it('inserts a generic field at an exact range and undoes it natively', async () => {
    const editor = createEditor('Alpha')
    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.field.insert',
          arguments: {
            range: { from: 1, to: 1 },
            instruction: 'PAGE',
            displayText: '1',
          },
        },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.field.insert',
      ok: true,
      output: {
        summary: 'Inserted DOCX PAGE field over range 1..2',
        from: 1,
        to: 2,
        instruction: 'PAGE',
        changed: true,
      },
    })
    expect(editor.state.doc.nodeAt(1)?.marks[0]).toMatchObject({
      type: { name: 'instrField' },
      attrs: { instr: 'PAGE' },
    })
    expect(editor.commands.undo()).toBe(true)
    expect(editor.state.doc.textContent).toBe('Alpha')
  })

  it('declares one bounded aggregate DOCX field-cache update contract', () => {
    const descriptor = docxOperationCatalog.operations.find(
      (operation) => operation.id === 'docx.field.update',
    )

    expect(descriptor).toMatchObject({
      family: 'field',
      visibility: 'agent',
      inputSchema: {
        properties: {
          updates: {
            type: 'array',
            minItems: 1,
            maxItems: 1024,
            items: {
              type: 'object',
              properties: {
                range: docxTextRangeSchema,
                instruction: {
                  type: 'string',
                  minLength: 1,
                  maxLength: 512,
                  pattern: '^[A-Za-z]+(?:\\s.*)?$',
                },
                displayText: { type: 'string', minLength: 1, maxLength: 65_536 },
              },
              required: ['range', 'instruction', 'displayText'],
              additionalProperties: false,
            },
          },
        },
        required: ['updates'],
        additionalProperties: false,
      },
      outputSchema: {
        required: ['summary', 'matched', 'changed'],
        additionalProperties: false,
      },
      mutates: true,
      undoable: true,
      atomic: true,
    })
  })

  it('updates exact field caches in one native Undo transaction', async () => {
    const editor = createEditor('Alpha')
    await executeDocxOperation(
      editor,
      {
        operation: 'docx.field.insert',
        arguments: {
          range: { from: 1, to: 1 },
          instruction: 'PAGE',
          displayText: '1',
        },
      },
      services,
    )
    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.field.update',
          arguments: {
            updates: [{ range: { from: 1, to: 2 }, instruction: 'PAGE', displayText: '12' }],
          },
        },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.field.update',
      ok: true,
      output: {
        summary: 'Updated 1 DOCX field cache(s)',
        matched: 1,
        changed: 1,
      },
    })
    expect(editor.state.doc.textContent).toBe('12Alpha')
    expect(editor.commands.undo()).toBe(true)
    expect(editor.state.doc.textContent).toBe('1Alpha')
  })

  it('rejects field-cache updates that do not exactly match a field mark', async () => {
    const editor = createEditor('Alpha')
    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.field.update',
          arguments: {
            updates: [{ range: { from: 1, to: 2 }, instruction: 'PAGE', displayText: '2' }],
          },
        },
        services,
      ),
    ).resolves.toMatchObject({ handled: true, ok: false, error: 'invalid_arguments' })
  })

  it('declares one bounded nullable exact-range DOCX link contract', () => {
    const descriptor = docxOperationCatalog.operations.find(
      (operation) => operation.id === 'docx.text.set_link',
    )

    expect(descriptor).toMatchObject({
      family: 'text',
      visibility: 'agent',
      inputSchema: {
        properties: {
          range: {
            type: 'object',
            properties: {
              from: { type: 'integer', minimum: 1 },
              to: { type: 'integer', minimum: 1 },
            },
            required: ['from', 'to'],
            additionalProperties: false,
          },
          href: { type: ['string', 'null'], minLength: 1, maxLength: 4096 },
          text: { type: ['string', 'null'], minLength: 1, maxLength: 65_536 },
        },
        required: ['range', 'href', 'text'],
        additionalProperties: false,
      },
      outputSchema: {
        required: ['summary', 'from', 'to', 'href', 'changed'],
        additionalProperties: false,
      },
      mutates: true,
      undoable: true,
      atomic: true,
    })
  })

  it('inserts, updates, and removes exact-range links through one final-state operation', async () => {
    const editor = createEditor('Alpha beta')
    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.text.set_link',
          arguments: {
            range: { from: 1, to: 6 },
            href: 'https://example.com/first',
            text: 'Linked',
          },
        },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.text.set_link',
      ok: true,
      output: {
        summary: 'Set DOCX link over range 1..7',
        from: 1,
        to: 7,
        href: 'https://example.com/first',
        changed: true,
      },
    })
    expect(editor.state.doc.textBetween(1, 7)).toBe('Linked')
    expect(
      editor.state.doc.nodeAt(1)?.marks.find((mark) => mark.type.name === 'link')?.attrs.href,
    ).toBe('https://example.com/first')
    expect(editor.commands.undo()).toBe(true)
    expect(editor.state.doc.textContent).toBe('Alpha beta')

    await executeDocxOperation(
      editor,
      {
        operation: 'docx.text.set_link',
        arguments: {
          range: { from: 1, to: 6 },
          href: 'https://example.com/second',
          text: null,
        },
      },
      services,
    )
    expect(
      editor.state.doc.nodeAt(1)?.marks.find((mark) => mark.type.name === 'link')?.attrs.href,
    ).toBe('https://example.com/second')
    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.text.set_link',
          arguments: { range: { from: 1, to: 6 }, href: null, text: null },
        },
        services,
      ),
    ).resolves.toMatchObject({ handled: true, operationId: 'docx.text.set_link', ok: true })
    expect(editor.state.doc.nodeAt(1)?.marks.some((mark) => mark.type.name === 'link')).toBe(false)
  })

  it('rejects link removal with replacement text', async () => {
    const editor = createEditor('Alpha')
    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.text.set_link',
          arguments: { range: { from: 1, to: 6 }, href: null, text: 'invalid' },
        },
        services,
      ),
    ).resolves.toMatchObject({
      handled: true,
      ok: false,
      error: 'invalid_arguments',
      message: 'DOCX link removal requires null text.',
    })
  })

  it('declares an exact-placement bounded DOCX equation insertion contract', () => {
    const descriptor = docxOperationCatalog.operations.find(
      (operation) => operation.id === 'docx.equation.insert',
    )

    expect(descriptor).toMatchObject({
      family: 'equation',
      visibility: 'agent',
      inputSchema: {
        properties: {
          placement: { type: 'string', enum: ['block', 'inline'] },
          latex: { type: 'string', minLength: 1, maxLength: 4096 },
          afterBlockIndex: {
            type: ['integer', 'null'],
            minimum: -1,
            maximum: 2_147_483_647,
          },
          from: { type: ['integer', 'null'], minimum: 1, maximum: 2_147_483_647 },
          to: { type: ['integer', 'null'], minimum: 1, maximum: 2_147_483_647 },
        },
        required: ['placement', 'latex', 'afterBlockIndex', 'from', 'to'],
        additionalProperties: false,
      },
      outputSchema: {
        required: ['summary', 'placement', 'equationBlockIndex', 'from', 'to'],
        additionalProperties: false,
      },
      mutates: true,
      undoable: true,
      atomic: true,
    })
  })

  it('declares one bounded dual-mode DOCX equation update contract', () => {
    const descriptor = docxOperationCatalog.operations.find(
      (operation) => operation.id === 'docx.equation.update',
    )

    expect(descriptor).toMatchObject({
      family: 'equation',
      visibility: 'agent',
      inputSchema: {
        properties: {
          placement: { type: 'string', enum: ['block', 'inline'] },
          mode: { type: 'string', enum: ['latex', 'tokens'] },
          latex: { type: ['string', 'null'], minLength: 1, maxLength: 4096 },
          tokens: {
            type: ['array', 'null'],
            minItems: 1,
            maxItems: 1024,
            items: { type: 'string', maxLength: 4096 },
          },
          equationBlockIndex: {
            type: ['integer', 'null'],
            minimum: 0,
            maximum: 2_147_483_647,
          },
          from: { type: ['integer', 'null'], minimum: 1, maximum: 2_147_483_647 },
          to: { type: ['integer', 'null'], minimum: 1, maximum: 2_147_483_647 },
        },
        required: ['placement', 'mode', 'latex', 'tokens', 'equationBlockIndex', 'from', 'to'],
        additionalProperties: false,
      },
      outputSchema: {
        required: ['summary', 'placement', 'mode', 'equationBlockIndex', 'from', 'to', 'changed'],
        additionalProperties: false,
      },
      mutates: true,
      undoable: true,
      atomic: true,
    })
  })

  it('updates one exact block equation from LaTeX and restores it with one Undo', async () => {
    const editor = createEditor('Anchor')
    await executeDocxOperation(
      editor,
      {
        operation: 'docx.equation.insert',
        arguments: {
          placement: 'block',
          latex: 'a+b',
          afterBlockIndex: 0,
          from: null,
          to: null,
        },
      },
      services,
    )
    const before = editor.state.doc.child(1).toJSON()

    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.equation.update',
          arguments: {
            placement: 'block',
            mode: 'latex',
            latex: 'x=\\frac{1}{2}',
            tokens: null,
            equationBlockIndex: 1,
            from: null,
            to: null,
          },
        },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.equation.update',
      ok: true,
      output: {
        summary: 'Updated block DOCX equation at block 1 from latex',
        placement: 'block',
        mode: 'latex',
        equationBlockIndex: 1,
        from: null,
        to: null,
        changed: true,
      },
    })
    expect(editor.state.doc.child(1).attrs.formulaDisplay).toMatchObject({
      latex: 'x=\\frac{1}{2}',
    })
    expect(editor.commands.undo()).toBe(true)
    expect(editor.state.doc.child(1).toJSON()).toEqual(before)
  })

  it('updates exact inline equation range and retained block tokens', async () => {
    const editor = createEditor('AB')
    await executeDocxOperation(
      editor,
      {
        operation: 'docx.equation.insert',
        arguments: {
          placement: 'inline',
          latex: 'a+b',
          afterBlockIndex: null,
          from: 2,
          to: 2,
        },
      },
      services,
    )
    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.equation.update',
          arguments: {
            placement: 'inline',
            mode: 'latex',
            latex: 'c^2',
            tokens: null,
            equationBlockIndex: null,
            from: 2,
            to: 3,
          },
        },
        services,
      ),
    ).resolves.toMatchObject({ handled: true, operationId: 'docx.equation.update', ok: true })
    expect(editor.state.doc.nodeAt(2)?.attrs.latex).toBe('c^2')

    await executeDocxOperation(
      editor,
      {
        operation: 'docx.equation.insert',
        arguments: {
          placement: 'block',
          latex: 'm+n',
          afterBlockIndex: 0,
          from: null,
          to: null,
        },
      },
      services,
    )
    const priorTokens = editor.state.doc.child(1).attrs.formulaDisplay.tokens as string[]
    const tokens = priorTokens.map((token, index) => (index === 0 ? 'z' : token))
    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.equation.update',
          arguments: {
            placement: 'block',
            mode: 'tokens',
            latex: null,
            tokens,
            equationBlockIndex: 1,
            from: null,
            to: null,
          },
        },
        services,
      ),
    ).resolves.toMatchObject({ handled: true, operationId: 'docx.equation.update', ok: true })
    expect(editor.state.doc.child(1).attrs.formulaDisplay.tokens).toEqual(tokens)
    expect(editor.state.doc.child(1).attrs.formulaDisplay.latex).toBeUndefined()
  })

  it('rejects mismatched equation modes and token shapes', async () => {
    const editor = createEditor('Anchor')
    await executeDocxOperation(
      editor,
      {
        operation: 'docx.equation.insert',
        arguments: {
          placement: 'block',
          latex: 'a+b',
          afterBlockIndex: 0,
          from: null,
          to: null,
        },
      },
      services,
    )
    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.equation.update',
          arguments: {
            placement: 'block',
            mode: 'latex',
            latex: 'x',
            tokens: ['x'],
            equationBlockIndex: 1,
            from: null,
            to: null,
          },
        },
        services,
      ),
    ).resolves.toMatchObject({
      handled: true,
      ok: false,
      error: 'invalid_arguments',
      message: 'LaTeX DOCX equation updates require latex and null tokens.',
    })

    const currentTokens = editor.state.doc.child(1).attrs.formulaDisplay.tokens as string[]
    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.equation.update',
          arguments: {
            placement: 'block',
            mode: 'tokens',
            latex: null,
            tokens: [...currentTokens, 'extra'],
            equationBlockIndex: 1,
            from: null,
            to: null,
          },
        },
        services,
      ),
    ).resolves.toMatchObject({
      handled: true,
      ok: false,
      error: 'invalid_arguments',
    })
  })

  it('declares one bounded cross-drawing DOCX object-size contract', () => {
    const descriptor = docxOperationCatalog.operations.find(
      (operation) => operation.id === 'docx.object.set_size',
    )

    expect(descriptor).toMatchObject({
      family: 'object',
      visibility: 'agent',
      inputSchema: {
        properties: {
          objectBlockIndex: { type: 'integer', minimum: 0, maximum: 2_147_483_647 },
          widthPx: { type: 'integer', minimum: 24, maximum: 4096 },
          heightPx: { type: 'integer', minimum: 8, maximum: 4096 },
        },
        required: ['objectBlockIndex', 'widthPx', 'heightPx'],
        additionalProperties: false,
      },
      outputSchema: {
        required: ['summary', 'objectBlockIndex', 'objectKind', 'widthPx', 'heightPx', 'changed'],
        additionalProperties: false,
      },
      mutates: true,
      undoable: true,
      atomic: true,
    })
  })

  it('declares one masked nullable DOCX object-style contract', () => {
    const descriptor = docxOperationCatalog.operations.find(
      (operation) => operation.id === 'docx.object.set_style',
    )

    expect(descriptor).toMatchObject({
      family: 'object',
      visibility: 'agent',
      inputSchema: {
        properties: {
          objectBlockIndex: { type: 'integer', minimum: 0, maximum: 2_147_483_647 },
          style: {
            type: 'object',
            properties: {
              fillHex: {
                type: ['string', 'null'],
                minLength: 6,
                maxLength: 6,
                pattern: '^[0-9A-F]{6}$',
              },
              borderHex: {
                type: ['string', 'null'],
                minLength: 6,
                maxLength: 6,
                pattern: '^[0-9A-F]{6}$',
              },
            },
            required: [],
            additionalProperties: false,
          },
          fields: {
            type: 'array',
            minItems: 1,
            maxItems: 2,
            items: { type: 'string', enum: ['fillHex', 'borderHex'] },
          },
        },
        required: ['objectBlockIndex', 'style', 'fields'],
        additionalProperties: false,
      },
      outputSchema: {
        required: ['summary', 'objectBlockIndex', 'objectKind', 'fields', 'changed'],
        additionalProperties: false,
      },
      mutates: true,
      undoable: true,
      atomic: true,
    })
  })

  it('sets masked shape fill/outline and restores prior style with one Undo', async () => {
    const editor = createEditor('Anchor')
    await executeDocxOperation(
      editor,
      {
        operation: 'docx.shape.insert',
        arguments: {
          afterBlockIndex: 0,
          preset: 'ellipse',
          widthEmu: 1_800_000,
          heightEmu: 1_080_000,
        },
      },
      services,
    )

    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.object.set_style',
          arguments: {
            objectBlockIndex: 1,
            style: { fillHex: 'A1B2C3', borderHex: null },
            fields: ['fillHex', 'borderHex'],
          },
        },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.object.set_style',
      ok: true,
      output: {
        summary: 'Set DOCX shape style at block 1',
        objectBlockIndex: 1,
        objectKind: 'shape',
        fields: ['fillHex', 'borderHex'],
        changed: true,
      },
    })
    expect(editor.state.doc.child(1).attrs.textboxes[0]).toMatchObject({ fill: 'A1B2C3' })
    expect(editor.state.doc.child(1).attrs.textboxes[0].borderColor).toBeUndefined()
    expect(editor.commands.undo()).toBe(true)
    expect(editor.state.doc.child(1).attrs.textboxes[0]).toMatchObject({
      fill: '4472C4',
      borderColor: '2F5496',
    })
  })

  it('rejects fill changes on stroke-only line objects', async () => {
    const editor = createEditor('Anchor')
    await executeDocxOperation(
      editor,
      {
        operation: 'docx.line.insert',
        arguments: {
          afterBlockIndex: 0,
          kind: 'lineBent',
          widthEmu: 1_800_000,
          heightEmu: 1_080_000,
        },
      },
      services,
    )

    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.object.set_style',
          arguments: {
            objectBlockIndex: 1,
            style: { fillHex: 'A1B2C3' },
            fields: ['fillHex'],
          },
        },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.object.set_style',
      ok: false,
      error: 'invalid_arguments',
      message: 'Stroke-only DOCX lines do not support fillHex.',
    })
  })

  it('declares one exact cross-object DOCX removal contract', () => {
    const descriptor = docxOperationCatalog.operations.find(
      (operation) => operation.id === 'docx.object.remove',
    )

    expect(descriptor).toMatchObject({
      family: 'object',
      visibility: 'agent',
      inputSchema: {
        properties: {
          objectBlockIndex: { type: 'integer', minimum: 0, maximum: 2_147_483_647 },
        },
        required: ['objectBlockIndex'],
        additionalProperties: false,
      },
      outputSchema: {
        properties: {
          objectKind: {
            type: 'string',
            enum: ['shape', 'line', 'textbox', 'chart', 'equation', 'diagram'],
          },
        },
        required: ['summary', 'objectBlockIndex', 'objectKind', 'replacementParagraphInserted'],
        additionalProperties: false,
      },
      mutates: true,
      undoable: true,
      atomic: true,
    })
  })

  it('removes one exact chart and restores it with one Undo', async () => {
    const editor = createEditor('Anchor')
    await executeDocxOperation(
      editor,
      {
        operation: 'docx.chart.insert',
        arguments: {
          afterBlockIndex: 0,
          kind: 'bar',
          title: 'Sales',
          categories: ['Q1'],
          series: [{ name: 'East', values: [10] }],
          widthPx: 576,
          heightPx: 336,
        },
      },
      services,
    )

    await expect(
      executeDocxOperation(
        editor,
        { operation: 'docx.object.remove', arguments: { objectBlockIndex: 1 } },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.object.remove',
      ok: true,
      output: {
        summary: 'Removed DOCX chart at block 1',
        objectBlockIndex: 1,
        objectKind: 'chart',
        replacementParagraphInserted: false,
      },
    })
    expect(editor.state.doc.childCount).toBe(1)
    expect(editor.commands.undo()).toBe(true)
    expect(editor.state.doc.child(1).attrs.chartDisplay).not.toBeNull()
  })

  it('rejects image deletion through the drawing-object family', async () => {
    const editor = createEditor('Anchor')
    editor
      .chain()
      .insertContentAt(editor.state.doc.content.size, {
        type: 'docProtected',
        attrs: { blockType: 'image', label: 'Image', imageDataUrl: 'data:image/png;base64,AA==' },
      })
      .run()

    await expect(
      executeDocxOperation(
        editor,
        { operation: 'docx.object.remove', arguments: { objectBlockIndex: 1 } },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.object.remove',
      ok: false,
      error: 'invalid_arguments',
      message: 'DOCX block 1 is not a removable drawing, chart, or block equation.',
    })
  })

  it('declares one explicit cross-drawing DOCX object-offset contract', () => {
    const descriptor = docxOperationCatalog.operations.find(
      (operation) => operation.id === 'docx.object.set_offset_position',
    )

    expect(descriptor).toMatchObject({
      family: 'object',
      visibility: 'agent',
      inputSchema: {
        properties: {
          objectBlockIndex: { type: 'integer', minimum: 0, maximum: 2_147_483_647 },
          wrap: {
            type: 'string',
            enum: ['square-left', 'square-right', 'topBottom', 'behind', 'front'],
          },
          offsetXEmu: { type: 'integer', minimum: -2_147_483_648, maximum: 2_147_483_647 },
          offsetYEmu: { type: 'integer', minimum: -2_147_483_648, maximum: 2_147_483_647 },
        },
        required: ['objectBlockIndex', 'wrap', 'offsetXEmu', 'offsetYEmu'],
        additionalProperties: false,
      },
      outputSchema: {
        required: [
          'summary',
          'objectBlockIndex',
          'objectKind',
          'wrap',
          'offsetXEmu',
          'offsetYEmu',
          'changed',
        ],
        additionalProperties: false,
      },
      mutates: true,
      undoable: true,
      atomic: true,
    })
  })

  it('positions one exact shape and restores its prior anchor with one Undo', async () => {
    const editor = createEditor('Anchor')
    await executeDocxOperation(
      editor,
      {
        operation: 'docx.shape.insert',
        arguments: {
          afterBlockIndex: 0,
          preset: 'ellipse',
          widthEmu: 1_800_000,
          heightEmu: 1_080_000,
        },
      },
      services,
    )

    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.object.set_offset_position',
          arguments: {
            objectBlockIndex: 1,
            wrap: 'front',
            offsetXEmu: -95_250,
            offsetYEmu: 190_500,
          },
        },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.object.set_offset_position',
      ok: true,
      output: {
        summary: 'Positioned DOCX shape at block 1 to (-95250, 190500) EMU with front wrap',
        objectBlockIndex: 1,
        objectKind: 'shape',
        wrap: 'front',
        offsetXEmu: -95_250,
        offsetYEmu: 190_500,
        changed: true,
      },
    })
    expect(editor.state.doc.child(1).attrs).toMatchObject({
      imageWrap: 'front',
      imageOffsetXEmu: -95_250,
      imageOffsetYEmu: 190_500,
      imagePosH: null,
      imagePosV: null,
    })
    expect(editor.commands.undo()).toBe(true)
    expect(editor.state.doc.child(1).attrs.imageOffsetXEmu).toBeNull()
    expect(editor.state.doc.child(1).attrs.imageOffsetYEmu).toBeNull()
  })

  it('resizes one exact shape and restores its prior dimensions with one Undo', async () => {
    const editor = createEditor('Anchor')
    await executeDocxOperation(
      editor,
      {
        operation: 'docx.shape.insert',
        arguments: {
          afterBlockIndex: 0,
          preset: 'ellipse',
          widthEmu: 1_800_000,
          heightEmu: 1_080_000,
        },
      },
      services,
    )
    const prior = editor.state.doc.child(1).attrs.textboxes[0]

    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.object.set_size',
          arguments: { objectBlockIndex: 1, widthPx: 320, heightPx: 180 },
        },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.object.set_size',
      ok: true,
      output: {
        summary: 'Set DOCX shape at block 1 to 320×180 px',
        objectBlockIndex: 1,
        objectKind: 'shape',
        widthPx: 320,
        heightPx: 180,
        changed: true,
      },
    })
    expect(editor.state.doc.child(1).attrs.textboxes[0]).toMatchObject({
      widthPx: 320,
      heightPx: 180,
      minHeightPx: 180,
    })
    expect(editor.commands.undo()).toBe(true)
    expect(editor.state.doc.child(1).attrs.textboxes[0]).toMatchObject({
      widthPx: prior.widthPx,
      heightPx: prior.heightPx,
    })
  })

  it('applies the retained narrower width limit when resizing a chart', async () => {
    const editor = createEditor('Anchor')
    await executeDocxOperation(
      editor,
      {
        operation: 'docx.chart.insert',
        arguments: {
          afterBlockIndex: 0,
          kind: 'bar',
          title: 'Sales',
          categories: ['Q1'],
          series: [{ name: 'East', values: [10] }],
          widthPx: 576,
          heightPx: 336,
        },
      },
      services,
    )

    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.object.set_size',
          arguments: { objectBlockIndex: 1, widthPx: 700, heightPx: 336 },
        },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.object.set_size',
      ok: false,
      error: 'invalid_arguments',
      message: 'DOCX chart widthPx must be an integer from 120 through 660.',
    })
  })

  it('inserts one display equation after a stable block and restores it with one Undo', async () => {
    const editor = createEditor('Anchor')
    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.equation.insert',
          arguments: {
            placement: 'block',
            latex: 'a^2 + b^2 = c^2',
            afterBlockIndex: 0,
            from: null,
            to: null,
          },
        },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.equation.insert',
      ok: true,
      output: {
        summary: 'Inserted block DOCX equation at block 1',
        placement: 'block',
        equationBlockIndex: 1,
        from: null,
        to: null,
      },
    })
    expect(editor.state.doc.child(1).attrs.formulaDisplay).toMatchObject({
      latex: 'a^2 + b^2 = c^2',
    })
    expect(editor.commands.undo()).toBe(true)
    expect(editor.state.doc.childCount).toBe(1)
  })

  it('replaces one exact inline range with an equation and restores it with one Undo', async () => {
    const editor = createEditor('Anchor')
    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.equation.insert',
          arguments: {
            placement: 'inline',
            latex: 'x^2',
            afterBlockIndex: null,
            from: 2,
            to: 4,
          },
        },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.equation.insert',
      ok: true,
      output: {
        summary: 'Inserted inline DOCX equation at range 2..4',
        placement: 'inline',
        equationBlockIndex: null,
        from: 2,
        to: 4,
      },
    })
    expect(editor.state.doc.child(0).child(1).type.name).toBe('docInlineMath')
    expect(editor.commands.undo()).toBe(true)
    expect(editor.state.doc.textContent).toBe('Anchor')
  })

  it('rejects mixed block and inline equation coordinates', async () => {
    const editor = createEditor('Anchor')
    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.equation.insert',
          arguments: {
            placement: 'block',
            latex: 'x',
            afterBlockIndex: 0,
            from: 1,
            to: 1,
          },
        },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.equation.insert',
      ok: false,
      error: 'invalid_arguments',
      message: 'Block DOCX equations require afterBlockIndex and null from/to coordinates.',
    })
  })

  it('declares a bounded stable-boundary DOCX chart insertion contract', () => {
    const descriptor = docxOperationCatalog.operations.find(
      (operation) => operation.id === 'docx.chart.insert',
    )

    expect(descriptor).toMatchObject({
      family: 'chart',
      visibility: 'agent',
      inputSchema: {
        properties: {
          afterBlockIndex: { type: 'integer', minimum: -1 },
          kind: { type: 'string', enum: ['bar', 'line', 'pie'] },
          title: { type: 'string' },
          categories: {
            type: 'array',
            minItems: 1,
            maxItems: 256,
            items: { type: 'string' },
          },
          series: {
            type: 'array',
            minItems: 1,
            maxItems: 64,
            items: {
              type: 'object',
              properties: {
                name: { type: 'string' },
                values: {
                  type: 'array',
                  minItems: 1,
                  maxItems: 256,
                  items: { type: ['number', 'null'], minimum: -1e12, maximum: 1e12 },
                },
              },
              required: ['name', 'values'],
              additionalProperties: false,
            },
          },
          widthPx: { type: 'integer', minimum: 120, maximum: 660 },
          heightPx: { type: 'integer', minimum: 80, maximum: 4096 },
        },
        required: [
          'afterBlockIndex',
          'kind',
          'title',
          'categories',
          'series',
          'widthPx',
          'heightPx',
        ],
        additionalProperties: false,
      },
      outputSchema: {
        required: ['summary', 'chartBlockIndex', 'kind', 'widthPx', 'heightPx'],
        additionalProperties: false,
      },
      mutates: true,
      undoable: true,
      atomic: true,
    })
  })

  it('inserts one bounded chart after a stable block and restores it with one Undo', async () => {
    const editor = createEditor('Anchor')
    const arguments_ = {
      afterBlockIndex: 0,
      kind: 'bar',
      title: 'Sales',
      categories: ['January', 'February'],
      series: [{ name: 'East', values: [10, 20] }],
      widthPx: 576,
      heightPx: 336,
    }

    await expect(
      executeDocxOperation(
        editor,
        { operation: 'docx.chart.insert', arguments: arguments_ },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.chart.insert',
      ok: true,
      output: {
        summary: 'Inserted DOCX bar chart at block 1',
        chartBlockIndex: 1,
        kind: 'bar',
        widthPx: 576,
        heightPx: 336,
      },
    })
    expect(editor.state.doc.child(1).attrs).toMatchObject({
      blockType: 'chart',
      genChart: {
        kind: 'bar',
        title: 'Sales',
        categories: ['January', 'February'],
        series: [{ name: 'East', values: [10, 20] }],
      },
      chartDisplay: { widthPx: 576, heightPx: 336 },
    })
    expect(editor.commands.undo()).toBe(true)
    expect(editor.state.doc.childCount).toBe(1)
  })

  it('rejects chart series whose values do not match the category count', async () => {
    const editor = createEditor('Anchor')
    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.chart.insert',
          arguments: {
            afterBlockIndex: 0,
            kind: 'line',
            title: 'Mismatch',
            categories: ['One', 'Two'],
            series: [{ name: 'Series', values: [1] }],
            widthPx: 576,
            heightPx: 336,
          },
        },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.chart.insert',
      ok: false,
      error: 'invalid_arguments',
      message: 'DOCX chart series 0 has 1 value(s), but 2 categories were supplied.',
    })
  })

  it('declares one masked bounded DOCX chart-update contract', () => {
    const descriptor = docxOperationCatalog.operations.find(
      (operation) => operation.id === 'docx.chart.update',
    )

    expect(descriptor).toMatchObject({
      family: 'chart',
      visibility: 'agent',
      inputSchema: {
        properties: {
          chartBlockIndex: { type: 'integer', minimum: 0, maximum: 2_147_483_647 },
          patch: {
            type: 'object',
            properties: {
              title: { type: 'string', maxLength: 512 },
              categories: {
                type: 'array',
                minItems: 1,
                maxItems: 256,
                items: { type: 'string', maxLength: 512 },
              },
              series: {
                type: 'array',
                minItems: 1,
                maxItems: 64,
                items: {
                  type: 'object',
                  properties: {
                    name: { type: ['string', 'null'], maxLength: 512 },
                    values: {
                      type: 'array',
                      minItems: 1,
                      maxItems: 256,
                      items: { type: ['number', 'null'], minimum: -1e12, maximum: 1e12 },
                    },
                  },
                  required: ['name', 'values'],
                  additionalProperties: false,
                },
              },
            },
            required: [],
            additionalProperties: false,
          },
          fields: {
            type: 'array',
            minItems: 1,
            maxItems: 3,
            items: { type: 'string', enum: ['title', 'categories', 'series'] },
          },
        },
        required: ['chartBlockIndex', 'patch', 'fields'],
        additionalProperties: false,
      },
      outputSchema: {
        required: ['summary', 'chartBlockIndex', 'fields', 'changed'],
        additionalProperties: false,
      },
      mutates: true,
      undoable: true,
      atomic: true,
    })
  })

  it('updates exact chart content and restores the prior matrix with one Undo', async () => {
    const editor = createEditor('Anchor')
    await executeDocxOperation(
      editor,
      {
        operation: 'docx.chart.insert',
        arguments: {
          afterBlockIndex: 0,
          kind: 'bar',
          title: 'Sales',
          categories: ['January', 'February'],
          series: [{ name: 'East', values: [10, 20] }],
          widthPx: 576,
          heightPx: 336,
        },
      },
      services,
    )

    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.chart.update',
          arguments: {
            chartBlockIndex: 1,
            patch: {
              title: 'Quarterly Sales',
              categories: ['Q1', 'Q2'],
              series: [{ name: 'North', values: [30, 40] }],
            },
            fields: ['title', 'categories', 'series'],
          },
        },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.chart.update',
      ok: true,
      output: {
        summary: 'Updated DOCX chart at block 1',
        chartBlockIndex: 1,
        fields: ['title', 'categories', 'series'],
        changed: true,
      },
    })
    expect(editor.state.doc.child(1).attrs.chartDisplay).toMatchObject({
      title: 'Quarterly Sales',
      categories: ['Q1', 'Q2'],
      series: [{ name: 'North', values: [30, 40] }],
    })
    expect(editor.commands.undo()).toBe(true)
    expect(editor.state.doc.child(1).attrs.chartDisplay).toMatchObject({
      title: 'Sales',
      categories: ['January', 'February'],
      series: [{ name: 'East', values: [10, 20] }],
    })
  })

  it('rejects writes into read-only chart cache gaps', async () => {
    const editor = createEditor('Anchor')
    await executeDocxOperation(
      editor,
      {
        operation: 'docx.chart.insert',
        arguments: {
          afterBlockIndex: 0,
          kind: 'line',
          title: 'Gaps',
          categories: ['One', 'Two'],
          series: [{ name: 'Series', values: [1, null] }],
          widthPx: 576,
          heightPx: 336,
        },
      },
      services,
    )

    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.chart.update',
          arguments: {
            chartBlockIndex: 1,
            patch: { series: [{ name: 'Series', values: [1, 2] }] },
            fields: ['series'],
          },
        },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.chart.update',
      ok: false,
      error: 'invalid_arguments',
      message: 'DOCX chart series 0 cache gap at value 1 is read-only.',
    })
  })

  it('declares public and internal bounded DOCX image insertion contracts', () => {
    const publicDescriptor = docxOperationCatalog.operations.find(
      (operation) => operation.id === 'docx.image.insert',
    )
    const stagedDescriptor = docxOperationCatalog.operations.find(
      (operation) => operation.id === 'docx.image.insert_staged',
    )

    expect(publicDescriptor).toMatchObject({
      family: 'image',
      visibility: 'agent',
      inputSchema: {
        properties: {
          path: { type: 'string' },
          afterBlockIndex: { type: 'integer', minimum: -1 },
          widthPx: { type: 'integer', minimum: 1, maximum: 4096 },
          heightPx: { type: 'integer', minimum: 1, maximum: 4096 },
          alignment: { type: 'string', enum: ['left', 'center', 'right'] },
        },
        required: ['path', 'afterBlockIndex', 'widthPx', 'heightPx', 'alignment'],
        additionalProperties: false,
      },
      outputSchema: {
        required: ['summary', 'insertedBlockIndex', 'widthPx', 'heightPx', 'alignment'],
        additionalProperties: false,
      },
      mutates: true,
      undoable: true,
      atomic: true,
    })
    expect(stagedDescriptor).toMatchObject({
      family: 'image',
      visibility: 'internal',
      inputSchema: {
        properties: {
          blobId: { type: 'string' },
          name: { type: 'string' },
          size: { type: 'integer', minimum: 1, maximum: 20_971_520 },
          data: { type: 'object' },
          afterBlockIndex: { type: 'integer', minimum: -1 },
          widthPx: { type: 'integer', minimum: 1, maximum: 4096 },
          heightPx: { type: 'integer', minimum: 1, maximum: 4096 },
          alignment: { type: 'string', enum: ['left', 'center', 'right'] },
        },
        required: [
          'blobId',
          'name',
          'size',
          'data',
          'afterBlockIndex',
          'widthPx',
          'heightPx',
          'alignment',
        ],
        additionalProperties: false,
      },
      mutates: true,
      undoable: true,
      atomic: true,
    })
  })

  it('inserts hydrated DOCX image bytes at an explicit boundary and undoes once', async () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          { type: 'docParagraph', content: [{ type: 'text', text: 'Before' }] },
          { type: 'docParagraph', content: [{ type: 'text', text: 'After' }] },
        ],
      },
    })
    editors.push(editor)
    const data = new TextEncoder().encode('GIF87a').buffer

    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.image.insert_staged',
          arguments: {
            blobId: 'image-blob',
            name: 'logo.gif',
            size: data.byteLength,
            data,
            afterBlockIndex: 0,
            widthPx: 320,
            heightPx: 180,
            alignment: 'center',
          },
        },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.image.insert_staged',
      ok: true,
      output: {
        summary: 'Inserted DOCX image at block 1',
        insertedBlockIndex: 1,
        widthPx: 320,
        heightPx: 180,
        alignment: 'center',
      },
    })
    expect(editor.state.doc.childCount).toBe(3)
    expect(editor.state.doc.child(1).attrs).toMatchObject({
      docxIndex: null,
      blockType: 'image',
      imageDataUrl: 'data:image/gif;base64,R0lGODdh',
      imageWidthPx: 320,
      imageHeightPx: 180,
      imageAlign: 'center',
      genImage: {
        base64: 'R0lGODdh',
        mime: 'image/gif',
        widthPx: 320,
        heightPx: 180,
        align: 'center',
      },
    })

    expect(editor.commands.undo()).toBe(true)
    expect(editor.state.doc.childCount).toBe(2)
    expect(editor.state.doc.textContent).toBe('BeforeAfter')
  })

  it('rejects mismatched staged DOCX image bytes without mutation', async () => {
    const editor = createEditor('Unchanged')
    const data = new TextEncoder().encode('GIF87a').buffer

    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.image.insert_staged',
          arguments: {
            blobId: 'image-blob',
            name: 'logo.png',
            size: data.byteLength,
            data,
            afterBlockIndex: 0,
            widthPx: 320,
            heightPx: 180,
            alignment: 'left',
          },
        },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.image.insert_staged',
      ok: false,
      error: 'invalid_arguments',
      message: 'docx.image.insert_staged requires matching PNG, JPEG, or GIF bytes.',
    })
    expect(editor.state.doc.childCount).toBe(1)
    expect(editor.state.doc.textContent).toBe('Unchanged')
  })

  it('declares public and internal bounded DOCX image replacement contracts', () => {
    const publicDescriptor = docxOperationCatalog.operations.find(
      (operation) => operation.id === 'docx.image.replace',
    )
    const stagedDescriptor = docxOperationCatalog.operations.find(
      (operation) => operation.id === 'docx.image.replace_staged',
    )

    expect(publicDescriptor).toMatchObject({
      family: 'image',
      visibility: 'agent',
      inputSchema: {
        properties: {
          path: { type: 'string' },
          imageBlockIndex: { type: 'integer', minimum: 0 },
          widthPx: { type: 'integer', minimum: 1, maximum: 4096 },
          heightPx: { type: 'integer', minimum: 1, maximum: 4096 },
        },
        required: ['path', 'imageBlockIndex', 'widthPx', 'heightPx'],
        additionalProperties: false,
      },
      outputSchema: {
        required: ['summary', 'imageBlockIndex', 'widthPx', 'heightPx'],
        additionalProperties: false,
      },
      mutates: true,
      undoable: true,
      atomic: true,
    })
    expect(stagedDescriptor).toMatchObject({
      family: 'image',
      visibility: 'internal',
      inputSchema: {
        properties: {
          blobId: { type: 'string' },
          name: { type: 'string' },
          size: { type: 'integer', minimum: 1, maximum: 20_971_520 },
          data: { type: 'object' },
          imageBlockIndex: { type: 'integer', minimum: 0 },
          widthPx: { type: 'integer', minimum: 1, maximum: 4096 },
          heightPx: { type: 'integer', minimum: 1, maximum: 4096 },
        },
        required: ['blobId', 'name', 'size', 'data', 'imageBlockIndex', 'widthPx', 'heightPx'],
        additionalProperties: false,
      },
      mutates: true,
      undoable: true,
      atomic: true,
    })
  })

  it('replaces one indexed original DOCX image and restores it with one Undo', async () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docProtected',
            attrs: {
              docxIndex: 7,
              blockType: 'image',
              label: 'Original image',
              imageDataUrl: 'data:image/png;base64,iVBORw0KGgo=',
              imageWidthPx: 400,
              imageHeightPx: 200,
              imageAlign: 'center',
              imageWrap: 'square-right',
              imageOffsetXEmu: 1200,
              imageOffsetYEmu: 2400,
              imageCrop: { l: 0.1, t: 0.2, r: 0, b: 0 },
              imageFillRect: { l: 0, t: 0, r: 0.1, b: 0 },
            },
          },
        ],
      },
    })
    editors.push(editor)
    const data = new TextEncoder().encode('GIF87a').buffer

    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.image.replace_staged',
          arguments: {
            blobId: 'replacement-blob',
            name: 'replacement.gif',
            size: data.byteLength,
            data,
            imageBlockIndex: 0,
            widthPx: 400,
            heightPx: 225,
          },
        },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.image.replace_staged',
      ok: true,
      output: {
        summary: 'Replaced DOCX image at block 0',
        imageBlockIndex: 0,
        widthPx: 400,
        heightPx: 225,
      },
    })
    expect(editor.state.doc.child(0).attrs).toMatchObject({
      docxIndex: 7,
      imageDataUrl: 'data:image/gif;base64,R0lGODdh',
      imageWidthPx: 400,
      imageHeightPx: 225,
      imageAlign: 'center',
      imageWrap: 'square-right',
      imageOffsetXEmu: 1200,
      imageOffsetYEmu: 2400,
      imageCrop: null,
      imageFillRect: null,
      imageReplace: { base64: 'R0lGODdh', mime: 'image/gif' },
    })

    expect(editor.commands.undo()).toBe(true)
    expect(editor.state.doc.child(0).attrs).toMatchObject({
      imageDataUrl: 'data:image/png;base64,iVBORw0KGgo=',
      imageWidthPx: 400,
      imageHeightPx: 200,
      imageCrop: { l: 0.1, t: 0.2, r: 0, b: 0 },
      imageFillRect: { l: 0, t: 0, r: 0.1, b: 0 },
      imageReplace: null,
    })
  })

  it('declares a bounded exact-index DOCX image wrap contract', () => {
    const descriptor = docxOperationCatalog.operations.find(
      (operation) => operation.id === 'docx.image.set_wrap',
    )

    expect(descriptor).toMatchObject({
      family: 'image',
      visibility: 'agent',
      inputSchema: {
        properties: {
          imageBlockIndex: { type: 'integer', minimum: 0 },
          wrap: {
            type: ['string', 'null'],
            enum: ['square-left', 'square-right', 'topBottom', 'behind', 'front', null],
          },
        },
        required: ['imageBlockIndex', 'wrap'],
        additionalProperties: false,
      },
      outputSchema: {
        required: ['summary', 'imageBlockIndex', 'wrap', 'changed'],
        additionalProperties: false,
      },
      mutates: true,
      undoable: true,
      atomic: true,
    })
  })

  it('sets image wrap while preserving position, then restores it with one Undo', async () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docProtected',
            attrs: {
              docxIndex: 4,
              blockType: 'image',
              imageDataUrl: 'data:image/gif;base64,R0lGODdh',
              imageWrap: 'front',
              imageOffsetXEmu: 1200,
              imageOffsetYEmu: 2400,
            },
          },
        ],
      },
    })
    editors.push(editor)

    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.image.set_wrap',
          arguments: { imageBlockIndex: 0, wrap: 'square-right' },
        },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.image.set_wrap',
      ok: true,
      output: {
        summary: 'Set DOCX image wrap at block 0',
        imageBlockIndex: 0,
        wrap: 'square-right',
        changed: true,
      },
    })
    expect(editor.state.doc.child(0).attrs).toMatchObject({
      imageWrap: 'square-right',
      imageOffsetXEmu: 1200,
      imageOffsetYEmu: 2400,
    })
    expect(editor.commands.undo()).toBe(true)
    expect(editor.state.doc.child(0).attrs.imageWrap).toBe('front')
  })

  it('sets an image inline and clears stale named and offset position state', async () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docProtected',
            attrs: {
              docxIndex: 4,
              blockType: 'image',
              imageDataUrl: 'data:image/gif;base64,R0lGODdh',
              imageWrap: 'square-left',
              imagePosH: 'center',
              imagePosV: 'bottom',
              imageOffsetXEmu: 1200,
              imageOffsetYEmu: 2400,
            },
          },
        ],
      },
    })
    editors.push(editor)

    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.image.set_wrap',
          arguments: { imageBlockIndex: 0, wrap: null },
        },
        services,
      ),
    ).resolves.toMatchObject({ handled: true, ok: true })
    expect(editor.state.doc.child(0).attrs).toMatchObject({
      imageWrap: null,
      imagePosH: null,
      imagePosV: null,
      imageOffsetXEmu: null,
      imageOffsetYEmu: null,
    })
  })

  it('declares a bounded exact-index DOCX margin-position contract', () => {
    const descriptor = docxOperationCatalog.operations.find(
      (operation) => operation.id === 'docx.image.set_margin_position',
    )

    expect(descriptor).toMatchObject({
      family: 'image',
      visibility: 'agent',
      inputSchema: {
        properties: {
          imageBlockIndex: { type: 'integer', minimum: 0 },
          horizontal: { type: 'string', enum: ['left', 'center', 'right'] },
          vertical: { type: 'string', enum: ['top', 'center', 'bottom'] },
        },
        required: ['imageBlockIndex', 'horizontal', 'vertical'],
        additionalProperties: false,
      },
      outputSchema: {
        required: ['summary', 'imageBlockIndex', 'horizontal', 'vertical', 'changed'],
        additionalProperties: false,
      },
      mutates: true,
      undoable: true,
      atomic: true,
    })
  })

  it('sets one original image margin position and restores it with one Undo', async () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docProtected',
            attrs: {
              docxIndex: 4,
              blockType: 'image',
              imageDataUrl: 'data:image/gif;base64,R0lGODdh',
              imageWrap: 'front',
              imageOffsetXEmu: 1200,
              imageOffsetYEmu: 2400,
            },
          },
        ],
      },
    })
    editors.push(editor)

    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.image.set_margin_position',
          arguments: { imageBlockIndex: 0, horizontal: 'right', vertical: 'bottom' },
        },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.image.set_margin_position',
      ok: true,
      output: {
        summary: 'Set DOCX image margin position at block 0',
        imageBlockIndex: 0,
        horizontal: 'right',
        vertical: 'bottom',
        changed: true,
      },
    })
    expect(editor.state.doc.child(0).attrs).toMatchObject({
      imageWrap: 'square-right',
      imagePosH: 'right',
      imagePosV: 'bottom',
      imageOffsetXEmu: null,
      imageOffsetYEmu: null,
    })
    expect(editor.commands.undo()).toBe(true)
    expect(editor.state.doc.child(0).attrs).toMatchObject({
      imageWrap: 'front',
      imagePosH: null,
      imagePosV: null,
      imageOffsetXEmu: 1200,
      imageOffsetYEmu: 2400,
    })
  })

  it('rejects margin positioning for an unsaved generated image', async () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docProtected',
            attrs: {
              docxIndex: null,
              blockType: 'image',
              imageDataUrl: 'data:image/gif;base64,R0lGODdh',
              genImage: {
                base64: 'R0lGODdh',
                mime: 'image/gif',
                widthPx: 100,
                heightPx: 100,
              },
            },
          },
        ],
      },
    })
    editors.push(editor)

    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.image.set_margin_position',
          arguments: { imageBlockIndex: 0, horizontal: 'center', vertical: 'top' },
        },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.image.set_margin_position',
      ok: false,
      error: 'invalid_arguments',
      message: 'DOCX image block 0 must be saved before setting a margin position.',
    })
  })

  it('declares a bounded exact-index DOCX offset-position contract', () => {
    const descriptor = docxOperationCatalog.operations.find(
      (operation) => operation.id === 'docx.image.set_offset_position',
    )

    expect(descriptor).toMatchObject({
      family: 'image',
      visibility: 'agent',
      inputSchema: {
        properties: {
          imageBlockIndex: { type: 'integer', minimum: 0 },
          wrap: {
            type: 'string',
            enum: [
              'square-left',
              'square-right',
              'tight-left',
              'tight-right',
              'through-left',
              'through-right',
              'topBottom',
              'behind',
              'front',
            ],
          },
          offsetXEmu: { type: 'integer', minimum: -2_147_483_648, maximum: 2_147_483_647 },
          offsetYEmu: { type: 'integer', minimum: -2_147_483_648, maximum: 2_147_483_647 },
        },
        required: ['imageBlockIndex', 'wrap', 'offsetXEmu', 'offsetYEmu'],
        additionalProperties: false,
      },
      outputSchema: {
        required: ['summary', 'imageBlockIndex', 'wrap', 'offsetXEmu', 'offsetYEmu', 'changed'],
        additionalProperties: false,
      },
      mutates: true,
      undoable: true,
      atomic: true,
    })
  })

  it('sets one generated image free position and restores it with one Undo', async () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docProtected',
            attrs: {
              docxIndex: null,
              blockType: 'image',
              imageDataUrl: 'data:image/gif;base64,R0lGODdh',
              imageWrap: 'square-right',
              imagePosH: 'right',
              imagePosV: 'top',
              genImage: {
                base64: 'R0lGODdh',
                mime: 'image/gif',
                widthPx: 100,
                heightPx: 100,
              },
            },
          },
        ],
      },
    })
    editors.push(editor)

    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.image.set_offset_position',
          arguments: {
            imageBlockIndex: 0,
            wrap: 'tight-left',
            offsetXEmu: -95_250,
            offsetYEmu: 190_500,
          },
        },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.image.set_offset_position',
      ok: true,
      output: {
        summary: 'Set DOCX image offset position at block 0',
        imageBlockIndex: 0,
        wrap: 'tight-left',
        offsetXEmu: -95_250,
        offsetYEmu: 190_500,
        changed: true,
      },
    })
    expect(editor.state.doc.child(0).attrs).toMatchObject({
      imageWrap: 'tight-left',
      imageOffsetXEmu: -95_250,
      imageOffsetYEmu: 190_500,
      imagePosH: null,
      imagePosV: null,
    })
    expect(editor.commands.undo()).toBe(true)
    expect(editor.state.doc.child(0).attrs).toMatchObject({
      imageWrap: 'square-right',
      imageOffsetXEmu: null,
      imageOffsetYEmu: null,
      imagePosH: 'right',
      imagePosV: 'top',
    })
  })

  it('declares a bounded exact-index DOCX image-transform contract', () => {
    const descriptor = docxOperationCatalog.operations.find(
      (operation) => operation.id === 'docx.image.set_transform',
    )

    expect(descriptor).toMatchObject({
      family: 'image',
      visibility: 'agent',
      inputSchema: {
        properties: {
          imageBlockIndex: { type: 'integer', minimum: 0 },
          rotationDegrees: { type: 'integer', minimum: 0, maximum: 359 },
          flipHorizontal: { type: 'boolean' },
          flipVertical: { type: 'boolean' },
        },
        required: ['imageBlockIndex', 'rotationDegrees', 'flipHorizontal', 'flipVertical'],
        additionalProperties: false,
      },
      outputSchema: {
        required: [
          'summary',
          'imageBlockIndex',
          'rotationDegrees',
          'flipHorizontal',
          'flipVertical',
          'changed',
        ],
        additionalProperties: false,
      },
      mutates: true,
      undoable: true,
      atomic: true,
    })
  })

  it('sets one generated image transform and restores it with one Undo', async () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docProtected',
            attrs: {
              docxIndex: null,
              blockType: 'image',
              imageDataUrl: 'data:image/gif;base64,R0lGODdh',
              imageRotDeg: 90,
              imageFlipH: true,
              imageFlipV: false,
              genImage: {
                base64: 'R0lGODdh',
                mime: 'image/gif',
                widthPx: 100,
                heightPx: 100,
              },
            },
          },
        ],
      },
    })
    editors.push(editor)

    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.image.set_transform',
          arguments: {
            imageBlockIndex: 0,
            rotationDegrees: 270,
            flipHorizontal: false,
            flipVertical: true,
          },
        },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.image.set_transform',
      ok: true,
      output: {
        summary: 'Set DOCX image transform at block 0',
        imageBlockIndex: 0,
        rotationDegrees: 270,
        flipHorizontal: false,
        flipVertical: true,
        changed: true,
      },
    })
    expect(editor.state.doc.child(0).attrs).toMatchObject({
      imageRotDeg: 270,
      imageFlipH: false,
      imageFlipV: true,
    })
    expect(editor.commands.undo()).toBe(true)
    expect(editor.state.doc.child(0).attrs).toMatchObject({
      imageRotDeg: 90,
      imageFlipH: true,
      imageFlipV: false,
    })
  })

  it('declares a bounded exact-index DOCX image-crop contract', () => {
    const descriptor = docxOperationCatalog.operations.find(
      (operation) => operation.id === 'docx.image.set_crop',
    )

    expect(descriptor).toMatchObject({
      family: 'image',
      visibility: 'agent',
      inputSchema: {
        properties: {
          imageBlockIndex: { type: 'integer', minimum: 0 },
          left: { type: 'number', minimum: 0, maximum: 0.99 },
          top: { type: 'number', minimum: 0, maximum: 0.99 },
          right: { type: 'number', minimum: 0, maximum: 0.99 },
          bottom: { type: 'number', minimum: 0, maximum: 0.99 },
        },
        required: ['imageBlockIndex', 'left', 'top', 'right', 'bottom'],
        additionalProperties: false,
      },
      outputSchema: {
        required: ['summary', 'imageBlockIndex', 'left', 'top', 'right', 'bottom', 'changed'],
        additionalProperties: false,
      },
      mutates: true,
      undoable: true,
      atomic: true,
    })
  })

  it('sets one generated image crop and restores it with one Undo', async () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docProtected',
            attrs: {
              docxIndex: null,
              blockType: 'image',
              imageDataUrl: 'data:image/gif;base64,R0lGODdh',
              imageCrop: null,
              imageFillRect: { l: 0, t: 0, r: 0.1, b: 0 },
              genImage: {
                base64: 'R0lGODdh',
                mime: 'image/gif',
                widthPx: 100,
                heightPx: 100,
              },
            },
          },
        ],
      },
    })
    editors.push(editor)

    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.image.set_crop',
          arguments: {
            imageBlockIndex: 0,
            left: 0.1,
            top: 0.2,
            right: 0.15,
            bottom: 0.05,
          },
        },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.image.set_crop',
      ok: true,
      output: {
        summary: 'Set DOCX image crop at block 0',
        imageBlockIndex: 0,
        left: 0.1,
        top: 0.2,
        right: 0.15,
        bottom: 0.05,
        changed: true,
      },
    })
    expect(editor.state.doc.child(0).attrs).toMatchObject({
      imageCrop: { l: 0.1, t: 0.2, r: 0.15, b: 0.05 },
      imageFillRect: null,
    })
    expect(editor.commands.undo()).toBe(true)
    expect(editor.state.doc.child(0).attrs).toMatchObject({
      imageCrop: null,
      imageFillRect: { l: 0, t: 0, r: 0.1, b: 0 },
    })
  })

  it('rejects a DOCX image crop with no remaining source area', async () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docProtected',
            attrs: {
              docxIndex: null,
              blockType: 'image',
              genImage: {
                base64: 'R0lGODdh',
                mime: 'image/gif',
                widthPx: 100,
                heightPx: 100,
              },
            },
          },
        ],
      },
    })
    editors.push(editor)

    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.image.set_crop',
          arguments: { imageBlockIndex: 0, left: 0.5, top: 0, right: 0.5, bottom: 0 },
        },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.image.set_crop',
      ok: false,
      error: 'invalid_arguments',
      message: 'DOCX image crop must retain positive horizontal and vertical source area.',
    })
  })

  it('declares a bounded exact-index DOCX image-removal contract', () => {
    const descriptor = docxOperationCatalog.operations.find(
      (operation) => operation.id === 'docx.image.remove',
    )

    expect(descriptor).toMatchObject({
      family: 'image',
      visibility: 'agent',
      inputSchema: {
        properties: { imageBlockIndex: { type: 'integer', minimum: 0 } },
        required: ['imageBlockIndex'],
        additionalProperties: false,
      },
      outputSchema: {
        required: ['summary', 'imageBlockIndex', 'deleted'],
        additionalProperties: false,
      },
      mutates: true,
      undoable: true,
      atomic: true,
    })
  })

  it('removes the only image with a replacement paragraph and restores it with one Undo', async () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docProtected',
            attrs: {
              docxIndex: null,
              blockType: 'image',
              imageDataUrl: 'data:image/gif;base64,R0lGODdh',
              genImage: {
                base64: 'R0lGODdh',
                mime: 'image/gif',
                widthPx: 100,
                heightPx: 100,
              },
            },
          },
        ],
      },
    })
    editors.push(editor)

    await expect(
      executeDocxOperation(
        editor,
        { operation: 'docx.image.remove', arguments: { imageBlockIndex: 0 } },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.image.remove',
      ok: true,
      output: {
        summary: 'Removed DOCX image at block 0',
        imageBlockIndex: 0,
        deleted: true,
      },
    })
    expect(editor.state.doc.childCount).toBe(1)
    expect(editor.state.doc.child(0).type.name).toBe('docParagraph')
    expect(editor.commands.undo()).toBe(true)
    expect(editor.state.doc.child(0).attrs.blockType).toBe('image')
  })

  it('declares a bounded stable-boundary DOCX preset-shape insertion contract', () => {
    const descriptor = docxOperationCatalog.operations.find(
      (operation) => operation.id === 'docx.shape.insert',
    )
    const presets = descriptor?.inputSchema.properties.preset.enum

    expect(descriptor).toMatchObject({
      family: 'shape',
      visibility: 'agent',
      inputSchema: {
        properties: {
          afterBlockIndex: { type: 'integer', minimum: -1 },
          preset: { type: 'string' },
          widthEmu: { type: 'integer', minimum: 9_525, maximum: 20_000_000 },
          heightEmu: { type: 'integer', minimum: 9_525, maximum: 20_000_000 },
        },
        required: ['afterBlockIndex', 'preset', 'widthEmu', 'heightEmu'],
        additionalProperties: false,
      },
      outputSchema: {
        required: ['summary', 'shapeBlockIndex', 'preset', 'widthEmu', 'heightEmu'],
        additionalProperties: false,
      },
      mutates: true,
      undoable: true,
      atomic: true,
    })
    expect(presets).toHaveLength(104)
    expect(presets).toEqual(expect.arrayContaining(['rect', 'diamond', 'flowChartProcess']))
    expect(presets).not.toContain('line')
  })

  it('inserts one preset shape after a stable block and restores it with one Undo', async () => {
    const editor = createEditor('Anchor')

    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.shape.insert',
          arguments: {
            afterBlockIndex: 0,
            preset: 'diamond',
            widthEmu: 1_800_000,
            heightEmu: 1_080_000,
          },
        },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.shape.insert',
      ok: true,
      output: {
        summary: 'Inserted DOCX diamond shape at block 1',
        shapeBlockIndex: 1,
        preset: 'diamond',
        widthEmu: 1_800_000,
        heightEmu: 1_080_000,
      },
    })
    expect(editor.state.doc.child(1).attrs).toMatchObject({
      blockType: 'passthrough',
      textboxes: [
        expect.objectContaining({
          prst: 'diamond',
          widthPx: Math.round(1_800_000 / 9_525),
          heightPx: Math.round(1_080_000 / 9_525),
        }),
      ],
    })
    expect(editor.commands.undo()).toBe(true)
    expect(editor.state.doc.childCount).toBe(1)
  })

  it('declares a bounded stable-boundary DOCX line insertion contract', () => {
    const descriptor = docxOperationCatalog.operations.find(
      (operation) => operation.id === 'docx.line.insert',
    )

    expect(descriptor).toMatchObject({
      family: 'line',
      visibility: 'agent',
      inputSchema: {
        properties: {
          afterBlockIndex: { type: 'integer', minimum: -1 },
          kind: {
            type: 'string',
            enum: ['line', 'lineArrow', 'lineArrowDouble', 'lineBent', 'lineCurved'],
          },
          widthEmu: { type: 'integer', minimum: 9_525, maximum: 20_000_000 },
          heightEmu: { type: 'integer', minimum: 9_525, maximum: 20_000_000 },
        },
        required: ['afterBlockIndex', 'kind', 'widthEmu', 'heightEmu'],
        additionalProperties: false,
      },
      outputSchema: {
        required: ['summary', 'lineBlockIndex', 'kind', 'widthEmu', 'heightEmu'],
        additionalProperties: false,
      },
      mutates: true,
      undoable: true,
      atomic: true,
    })
  })

  it('inserts one connector after a stable block and restores it with one Undo', async () => {
    const editor = createEditor('Anchor')

    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.line.insert',
          arguments: {
            afterBlockIndex: 0,
            kind: 'lineBent',
            widthEmu: 1_800_000,
            heightEmu: 1_080_000,
          },
        },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.line.insert',
      ok: true,
      output: {
        summary: 'Inserted DOCX lineBent connector at block 1',
        lineBlockIndex: 1,
        kind: 'lineBent',
        widthEmu: 1_800_000,
        heightEmu: 1_080_000,
      },
    })
    expect(editor.state.doc.child(1).attrs.textboxes[0]).toMatchObject({
      prst: 'lineBent',
      readOnly: true,
      widthPx: Math.round(1_800_000 / 9_525),
      heightPx: Math.round(1_080_000 / 9_525),
    })
    expect(editor.commands.undo()).toBe(true)
    expect(editor.state.doc.childCount).toBe(1)
  })

  it('rejects a straight DOCX line with a noncanonical grab height', async () => {
    const editor = createEditor('Anchor')
    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.line.insert',
          arguments: {
            afterBlockIndex: 0,
            kind: 'lineArrow',
            widthEmu: 1_800_000,
            heightEmu: 1_080_000,
          },
        },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.line.insert',
      ok: false,
      error: 'invalid_arguments',
      message: 'Straight DOCX lines require heightEmu 114300.',
    })
  })

  it('declares a bounded stable-boundary DOCX textbox insertion contract', () => {
    const descriptor = docxOperationCatalog.operations.find(
      (operation) => operation.id === 'docx.textbox.insert',
    )

    expect(descriptor).toMatchObject({
      family: 'textbox',
      visibility: 'agent',
      inputSchema: {
        properties: {
          afterBlockIndex: { type: 'integer', minimum: -1 },
          widthEmu: { type: 'integer', minimum: 9_525, maximum: 20_000_000 },
          heightEmu: { type: 'integer', minimum: 9_525, maximum: 20_000_000 },
        },
        required: ['afterBlockIndex', 'widthEmu', 'heightEmu'],
        additionalProperties: false,
      },
      outputSchema: {
        required: ['summary', 'textboxBlockIndex', 'widthEmu', 'heightEmu'],
        additionalProperties: false,
      },
      mutates: true,
      undoable: true,
      atomic: true,
    })
  })

  it('inserts one textbox after a stable block and restores it with one Undo', async () => {
    const editor = createEditor('Anchor')

    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.textbox.insert',
          arguments: {
            afterBlockIndex: 0,
            widthEmu: 1_800_000,
            heightEmu: 1_080_000,
          },
        },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.textbox.insert',
      ok: true,
      output: {
        summary: 'Inserted DOCX textbox at block 1',
        textboxBlockIndex: 1,
        widthEmu: 1_800_000,
        heightEmu: 1_080_000,
      },
    })
    expect(editor.state.doc.child(1).attrs.textboxes[0]).toMatchObject({
      fill: 'FFFFFF',
      borderColor: '000000',
      widthPx: Math.round(1_800_000 / 9_525),
      heightPx: Math.round(1_080_000 / 9_525),
      paras: [{ runs: [{ text: '' }] }],
    })
    expect(editor.commands.undo()).toBe(true)
    expect(editor.state.doc.childCount).toBe(1)
  })

  it('declares one bounded rich DOCX textbox-content contract', () => {
    const descriptor = docxOperationCatalog.operations.find(
      (operation) => operation.id === 'docx.textbox.set_content',
    )

    expect(descriptor).toMatchObject({
      family: 'textbox',
      visibility: 'agent',
      inputSchema: {
        properties: {
          objectBlockIndex: { type: 'integer', minimum: 0, maximum: 2_147_483_647 },
          textboxIndex: { type: 'integer', minimum: 0, maximum: 1023 },
          paragraphs: {
            type: 'array',
            minItems: 1,
            maxItems: 256,
            items: {
              type: 'object',
              properties: {
                align: {
                  type: 'string',
                  enum: ['left', 'center', 'right', 'justify', 'distribute'],
                },
                runs: {
                  type: 'array',
                  minItems: 1,
                  maxItems: 1024,
                  items: {
                    type: 'object',
                    properties: {
                      text: { type: 'string', maxLength: 4096 },
                      bold: { type: 'boolean' },
                      color: { type: 'string', pattern: '^[0-9A-F]{6}$' },
                    },
                    required: ['text'],
                    additionalProperties: false,
                  },
                },
              },
              required: ['runs'],
              additionalProperties: false,
            },
          },
          heightPx: { type: ['integer', 'null'], minimum: 8, maximum: 4096 },
        },
        required: ['objectBlockIndex', 'textboxIndex', 'paragraphs', 'heightPx'],
        additionalProperties: false,
      },
      outputSchema: {
        required: [
          'summary',
          'objectBlockIndex',
          'textboxIndex',
          'paragraphCount',
          'heightPx',
          'changed',
        ],
        additionalProperties: false,
      },
      mutates: true,
      undoable: true,
      atomic: true,
    })
  })

  it('sets rich textbox content at an exact nested identity and restores it with one Undo', async () => {
    const editor = createEditor('Anchor')
    await executeDocxOperation(
      editor,
      {
        operation: 'docx.textbox.insert',
        arguments: { afterBlockIndex: 0, widthEmu: 1_800_000, heightEmu: 1_080_000 },
      },
      services,
    )
    const before = editor.state.doc.child(1).attrs.textboxes[0]
    const paragraphs = [
      {
        align: 'center',
        runs: [
          { text: 'Rich ', bold: true, color: 'FF0000', sizeHalfPoints: 28 },
          { text: 'textbox', italic: true, font: 'Courier New' },
        ],
      },
      { runs: [{ text: 'Second line', underline: true }] },
    ]

    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.textbox.set_content',
          arguments: {
            objectBlockIndex: 1,
            textboxIndex: 0,
            paragraphs,
            heightPx: 240,
          },
        },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.textbox.set_content',
      ok: true,
      output: {
        summary: 'Set DOCX textbox 0 content at block 1',
        objectBlockIndex: 1,
        textboxIndex: 0,
        paragraphCount: 2,
        heightPx: 240,
        changed: true,
      },
    })
    expect(editor.state.doc.child(1).attrs.textboxes[0]).toMatchObject({
      paras: paragraphs,
      heightPx: 240,
    })
    expect(editor.commands.undo()).toBe(true)
    expect(editor.state.doc.child(1).attrs.textboxes[0]).toEqual(before)
  })

  it('rejects read-only textbox content without mutation', async () => {
    const editor = createEditor('Anchor')
    await executeDocxOperation(
      editor,
      {
        operation: 'docx.textbox.insert',
        arguments: { afterBlockIndex: 0, widthEmu: 1_800_000, heightEmu: 1_080_000 },
      },
      services,
    )
    const block = editor.state.doc.child(1)
    const position = editor.state.doc.child(0).nodeSize
    editor.view.dispatch(
      editor.state.tr.setNodeMarkup(position, undefined, {
        ...block.attrs,
        textboxes: [{ ...block.attrs.textboxes[0], readOnly: true }],
      }),
    )
    const before = editor.state.doc.toJSON()
    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.textbox.set_content',
          arguments: {
            objectBlockIndex: 1,
            textboxIndex: 0,
            paragraphs: [{ runs: [{ text: 'Blocked' }] }],
            heightPx: null,
          },
        },
        services,
      ),
    ).resolves.toMatchObject({
      handled: true,
      ok: false,
      error: 'invalid_arguments',
      message: 'The DOCX textbox target is read-only because its structure is flattened.',
    })
    expect(editor.state.doc.toJSON()).toEqual(before)
  })

  it('rejects aggregate textbox content beyond the renderer budget', async () => {
    const editor = createEditor('Anchor')
    await executeDocxOperation(
      editor,
      {
        operation: 'docx.textbox.insert',
        arguments: { afterBlockIndex: 0, widthEmu: 1_800_000, heightEmu: 1_080_000 },
      },
      services,
    )
    const before = editor.state.doc.toJSON()
    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.textbox.set_content',
          arguments: {
            objectBlockIndex: 1,
            textboxIndex: 0,
            paragraphs: [{ runs: Array.from({ length: 17 }, () => ({ text: 'x'.repeat(4096) })) }],
            heightPx: null,
          },
        },
        services,
      ),
    ).resolves.toMatchObject({
      handled: true,
      ok: false,
      error: 'invalid_arguments',
      message: 'DOCX textbox content is limited to 65536 total Unicode characters.',
    })
    expect(editor.state.doc.toJSON()).toEqual(before)
  })

  it('declares a bounded stable-position DOCX page-break contract', () => {
    const descriptor = docxOperationCatalog.operations.find(
      (operation) => operation.id === 'docx.document.insert_page_break',
    )

    expect(descriptor).toMatchObject({
      family: 'document',
      inputSchema: {
        properties: {
          afterBlockIndex: { type: 'integer', minimum: -1 },
        },
        required: ['afterBlockIndex'],
        additionalProperties: false,
      },
      outputSchema: {
        required: ['summary', 'insertedBlockIndex'],
        additionalProperties: false,
      },
      mutates: true,
      undoable: true,
      atomic: true,
    })
  })

  it('inserts a native DOCX page-break paragraph at an explicit boundary and undoes once', async () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docParagraph',
            content: [{ type: 'text', text: 'Before' }],
          },
          {
            type: 'docParagraph',
            content: [{ type: 'text', text: 'After' }],
          },
        ],
      },
    })
    editors.push(editor)

    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.document.insert_page_break',
          arguments: { afterBlockIndex: 0 },
        },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.document.insert_page_break',
      ok: true,
      output: {
        summary: 'Inserted DOCX page break at block 1',
        insertedBlockIndex: 1,
      },
    })
    expect(editor.state.doc.childCount).toBe(3)
    expect(editor.state.doc.child(0).textContent).toBe('Before')
    expect(editor.state.doc.child(1).type.name).toBe('docParagraph')
    expect(editor.state.doc.child(1).attrs.pageBreakBefore).toBe(true)
    expect(editor.state.doc.child(1).textContent).toBe('')
    expect(editor.state.doc.child(2).textContent).toBe('After')

    expect(editor.commands.undo()).toBe(true)
    expect(editor.state.doc.childCount).toBe(2)
    expect(editor.state.doc.textContent).toBe('BeforeAfter')
  })

  it('rejects an out-of-range DOCX page-break boundary without mutation', async () => {
    const editor = createEditor('Unchanged')

    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.document.insert_page_break',
          arguments: { afterBlockIndex: 1 },
        },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.document.insert_page_break',
      ok: false,
      error: 'invalid_arguments',
      message: 'DOCX page-break boundary after block 1 is invalid for 1 block(s).',
    })
    expect(editor.state.doc.childCount).toBe(1)
    expect(editor.state.doc.textContent).toBe('Unchanged')
  })

  it('declares a bounded stable-position DOCX section-break contract', () => {
    const descriptor = docxOperationCatalog.operations.find(
      (operation) => operation.id === 'docx.section.insert_break',
    )

    expect(descriptor).toMatchObject({
      family: 'section',
      inputSchema: {
        properties: {
          afterBlockIndex: { type: 'integer', minimum: -1 },
          startType: {
            type: 'string',
            enum: ['nextPage', 'continuous', 'evenPage', 'oddPage'],
          },
        },
        required: ['afterBlockIndex', 'startType'],
        additionalProperties: false,
      },
      outputSchema: {
        required: ['summary', 'insertedBlockIndex', 'startType'],
        additionalProperties: false,
      },
      risk: 'high',
      mutates: true,
      undoable: true,
      atomic: true,
    })
  })

  it('inserts a native DOCX section-break node with undo-owned start type', async () => {
    const editor = createEditor('Section body')

    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.section.insert_break',
          arguments: { afterBlockIndex: 0, startType: 'continuous' },
        },
        {
          ...services,
          resolveSectionBreakSource: () => ({
            sectPrXml:
              '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>',
          }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.section.insert_break',
      ok: true,
      output: {
        summary: 'Inserted continuous DOCX section break at block 1',
        insertedBlockIndex: 1,
        startType: 'continuous',
      },
    })
    expect(editor.state.doc.childCount).toBe(2)
    expect(editor.state.doc.child(1).type.name).toBe('docProtected')
    expect(editor.state.doc.child(1).attrs.label).toBe('Section break paragraph')
    expect(editor.state.doc.child(1).attrs.sectionStartType).toBe('continuous')
    expect(editor.state.doc.child(1).attrs.genXml).toContain('<w:sectPr>')

    expect(editor.commands.undo()).toBe(true)
    expect(editor.state.doc.childCount).toBe(1)
    expect(editor.state.doc.textContent).toBe('Section body')
  })

  it('rejects an invalid DOCX section-break boundary before resolving section state', async () => {
    const editor = createEditor('Unchanged')
    let resolved = false

    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.section.insert_break',
          arguments: { afterBlockIndex: 2, startType: 'oddPage' },
        },
        {
          ...services,
          resolveSectionBreakSource: () => {
            resolved = true
            return null
          },
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.section.insert_break',
      ok: false,
      error: 'invalid_arguments',
      message: 'DOCX section-break boundary after block 2 is invalid for 1 block(s).',
    })
    expect(resolved).toBe(false)
    expect(editor.state.doc.childCount).toBe(1)
  })

  it('declares a bounded explicit DOCX section-orientation contract', () => {
    const descriptor = docxOperationCatalog.operations.find(
      (operation) => operation.id === 'docx.section.set_orientation',
    )

    expect(descriptor).toMatchObject({
      family: 'section',
      inputSchema: {
        properties: {
          sectionIndex: { type: 'integer', minimum: 0 },
          orientation: { type: 'string', enum: ['portrait', 'landscape'] },
        },
        required: ['sectionIndex', 'orientation'],
        additionalProperties: false,
      },
      outputSchema: {
        required: ['sectionIndex', 'orientation', 'changed'],
        additionalProperties: false,
      },
      risk: 'medium',
      mutates: true,
      undoable: true,
      atomic: true,
    })
  })

  it('sets an exact DOCX section orientation in the native history', async () => {
    const editor = createEditor('Portrait section')
    const section = {
      settings: {
        pageWidth: 12240,
        pageHeight: 15840,
        marginTop: 1440,
        marginRight: 1440,
        marginBottom: 1440,
        marginLeft: 1440,
        pageBorder: false,
        orientation: 'portrait' as const,
        columns: 1,
      },
      startType: 'nextPage' as const,
      firstBlockIndex: 0,
      lastBlockIndex: 1,
      sectPrXml: '<w:sectPr/>',
      titlePg: false,
      headerRefs: {},
      footerRefs: {},
    }

    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.section.set_orientation',
          arguments: { sectionIndex: 0, orientation: 'landscape' },
        },
        { ...services, getSectionLayoutState: () => [section] },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.section.set_orientation',
      ok: true,
      output: { sectionIndex: 0, orientation: 'landscape', changed: true },
    })
    expect(editor.state.doc.child(0).attrs.sectionSettingsOverride).toMatchObject({
      sectionIndex: 0,
      settings: {
        orientation: 'landscape',
        pageWidth: 15840,
        pageHeight: 12240,
      },
    })

    expect(editor.commands.undo()).toBe(true)
    expect(editor.state.doc.child(0).attrs.sectionSettingsOverride).toBeNull()
  })

  it('rejects a missing DOCX section index without mutation', async () => {
    const editor = createEditor('Unchanged')

    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.section.set_orientation',
          arguments: { sectionIndex: 1, orientation: 'landscape' },
        },
        { ...services, getSectionLayoutState: () => [] },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.section.set_orientation',
      ok: false,
      error: 'invalid_arguments',
      message: 'DOCX section index 1 is invalid for 0 section(s).',
    })
    expect(editor.state.doc.child(0).attrs.sectionSettingsOverride).toBeNull()
  })

  it('declares a bounded exact DOCX section-margins contract', () => {
    const descriptor = docxOperationCatalog.operations.find(
      (operation) => operation.id === 'docx.section.set_margins',
    )

    expect(descriptor).toMatchObject({
      family: 'section',
      inputSchema: {
        properties: {
          sectionIndex: { type: 'integer', minimum: 0 },
          margins: {
            properties: {
              topTwips: { type: 'integer', minimum: 0, maximum: 31680 },
              rightTwips: { type: 'integer', minimum: 0, maximum: 31680 },
              bottomTwips: { type: 'integer', minimum: 0, maximum: 31680 },
              leftTwips: { type: 'integer', minimum: 0, maximum: 31680 },
            },
            required: ['topTwips', 'rightTwips', 'bottomTwips', 'leftTwips'],
            additionalProperties: false,
          },
        },
        required: ['sectionIndex', 'margins'],
        additionalProperties: false,
      },
      outputSchema: {
        required: ['sectionIndex', 'margins', 'changed'],
        additionalProperties: false,
      },
      mutates: true,
      undoable: true,
      atomic: true,
    })
  })

  it('sets exact DOCX section margins through the shared settings journal', async () => {
    const editor = createEditor('Margins')
    const section = {
      settings: {
        pageWidth: 12240,
        pageHeight: 15840,
        marginTop: 1440,
        marginRight: 1440,
        marginBottom: 1440,
        marginLeft: 1440,
        pageBorder: false,
        orientation: 'portrait' as const,
        columns: 1,
      },
      startType: 'nextPage' as const,
      firstBlockIndex: 0,
      lastBlockIndex: 1,
      sectPrXml: '<w:sectPr/>',
      titlePg: false,
      headerRefs: {},
      footerRefs: {},
    }
    const margins = {
      topTwips: 720,
      rightTwips: 1080,
      bottomTwips: 720,
      leftTwips: 1080,
    }

    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.section.set_margins',
          arguments: { sectionIndex: 0, margins },
        },
        { ...services, getSectionLayoutState: () => [section] },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.section.set_margins',
      ok: true,
      output: { sectionIndex: 0, margins, changed: true },
    })
    expect(editor.state.doc.child(0).attrs.sectionSettingsOverride.settings).toMatchObject({
      marginTop: 720,
      marginRight: 1080,
      marginBottom: 720,
      marginLeft: 1080,
    })
  })

  it('rejects DOCX section margins that consume the page body', async () => {
    const editor = createEditor('Unchanged')
    const section = {
      settings: {
        pageWidth: 12240,
        pageHeight: 15840,
        marginTop: 1440,
        marginRight: 1440,
        marginBottom: 1440,
        marginLeft: 1440,
        pageBorder: false,
        orientation: 'portrait' as const,
        columns: 1,
      },
      startType: 'nextPage' as const,
      firstBlockIndex: 0,
      lastBlockIndex: 1,
      sectPrXml: '<w:sectPr/>',
      titlePg: false,
      headerRefs: {},
      footerRefs: {},
    }

    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.section.set_margins',
          arguments: {
            sectionIndex: 0,
            margins: {
              topTwips: 8000,
              rightTwips: 1440,
              bottomTwips: 8000,
              leftTwips: 1440,
            },
          },
        },
        { ...services, getSectionLayoutState: () => [section] },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.section.set_margins',
      ok: false,
      error: 'invalid_arguments',
      message: 'DOCX section margins must leave positive page width and height.',
    })
    expect(editor.state.doc.child(0).attrs.sectionSettingsOverride).toBeNull()
  })

  it('declares a bounded exact DOCX page-size contract', () => {
    const descriptor = docxOperationCatalog.operations.find(
      (operation) => operation.id === 'docx.section.set_page_size',
    )

    expect(descriptor).toMatchObject({
      family: 'section',
      inputSchema: {
        properties: {
          sectionIndex: { type: 'integer', minimum: 0 },
          widthTwips: { type: 'integer', minimum: 1440, maximum: 31680 },
          heightTwips: { type: 'integer', minimum: 1440, maximum: 31680 },
        },
        required: ['sectionIndex', 'widthTwips', 'heightTwips'],
        additionalProperties: false,
      },
      outputSchema: {
        required: ['sectionIndex', 'widthTwips', 'heightTwips', 'orientation', 'changed'],
        additionalProperties: false,
      },
      mutates: true,
      undoable: true,
      atomic: true,
    })
  })

  it('sets exact DOCX page axes and derives landscape final state', async () => {
    const editor = createEditor('Page size')
    const section = {
      settings: {
        pageWidth: 12240,
        pageHeight: 15840,
        marginTop: 1440,
        marginRight: 1440,
        marginBottom: 1440,
        marginLeft: 1440,
        pageBorder: false,
        orientation: 'portrait' as const,
        columns: 1,
      },
      startType: 'nextPage' as const,
      firstBlockIndex: 0,
      lastBlockIndex: 1,
      sectPrXml: '<w:sectPr/>',
      titlePg: false,
      headerRefs: {},
      footerRefs: {},
    }

    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.section.set_page_size',
          arguments: { sectionIndex: 0, widthTwips: 16838, heightTwips: 11906 },
        },
        { ...services, getSectionLayoutState: () => [section] },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.section.set_page_size',
      ok: true,
      output: {
        sectionIndex: 0,
        widthTwips: 16838,
        heightTwips: 11906,
        orientation: 'landscape',
        changed: true,
      },
    })
    expect(editor.state.doc.child(0).attrs.sectionSettingsOverride.settings).toMatchObject({
      pageWidth: 16838,
      pageHeight: 11906,
      orientation: 'landscape',
    })
  })

  it('rejects DOCX page axes smaller than the current margins', async () => {
    const editor = createEditor('Unchanged')
    const section = {
      settings: {
        pageWidth: 12240,
        pageHeight: 15840,
        marginTop: 1440,
        marginRight: 1440,
        marginBottom: 1440,
        marginLeft: 1440,
        pageBorder: false,
        orientation: 'portrait' as const,
        columns: 1,
      },
      startType: 'nextPage' as const,
      firstBlockIndex: 0,
      lastBlockIndex: 1,
      sectPrXml: '<w:sectPr/>',
      titlePg: false,
      headerRefs: {},
      footerRefs: {},
    }

    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.section.set_page_size',
          arguments: { sectionIndex: 0, widthTwips: 2000, heightTwips: 2000 },
        },
        { ...services, getSectionLayoutState: () => [section] },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.section.set_page_size',
      ok: false,
      error: 'invalid_arguments',
      message: 'DOCX page size must exceed the current horizontal and vertical margins.',
    })
    expect(editor.state.doc.child(0).attrs.sectionSettingsOverride).toBeNull()
  })

  it('declares a bounded exact DOCX section-columns contract', () => {
    const descriptor = docxOperationCatalog.operations.find(
      (operation) => operation.id === 'docx.section.set_columns',
    )

    expect(descriptor).toMatchObject({
      family: 'section',
      inputSchema: {
        properties: {
          sectionIndex: { type: 'integer', minimum: 0 },
          count: { type: 'integer', minimum: 1, maximum: 16 },
          spacingTwips: { type: 'integer', minimum: 0, maximum: 31680 },
        },
        required: ['sectionIndex', 'count', 'spacingTwips'],
        additionalProperties: false,
      },
      outputSchema: {
        required: ['sectionIndex', 'count', 'spacingTwips', 'changed'],
        additionalProperties: false,
      },
      mutates: true,
      undoable: true,
      atomic: true,
    })
  })

  it('sets exact DOCX section column count and spacing through the shared journal', async () => {
    const editor = createEditor('Columns')
    const section = {
      settings: {
        pageWidth: 12240,
        pageHeight: 15840,
        marginTop: 1440,
        marginRight: 1440,
        marginBottom: 1440,
        marginLeft: 1440,
        pageBorder: false,
        orientation: 'portrait' as const,
        columns: 1,
        colSpace: 720,
      },
      startType: 'nextPage' as const,
      firstBlockIndex: 0,
      lastBlockIndex: 1,
      sectPrXml: '<w:sectPr/>',
      titlePg: false,
      headerRefs: {},
      footerRefs: {},
    }

    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.section.set_columns',
          arguments: { sectionIndex: 0, count: 3, spacingTwips: 480 },
        },
        { ...services, getSectionLayoutState: () => [section] },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.section.set_columns',
      ok: true,
      output: { sectionIndex: 0, count: 3, spacingTwips: 480, changed: true },
    })
    expect(editor.state.doc.child(0).attrs.sectionSettingsOverride.settings).toMatchObject({
      columns: 3,
      colSpace: 480,
    })
  })

  it('rejects DOCX section column gaps that consume the text width', async () => {
    const editor = createEditor('Unchanged')
    const section = {
      settings: {
        pageWidth: 12240,
        pageHeight: 15840,
        marginTop: 1440,
        marginRight: 1440,
        marginBottom: 1440,
        marginLeft: 1440,
        pageBorder: false,
        orientation: 'portrait' as const,
        columns: 1,
      },
      startType: 'nextPage' as const,
      firstBlockIndex: 0,
      lastBlockIndex: 1,
      sectPrXml: '<w:sectPr/>',
      titlePg: false,
      headerRefs: {},
      footerRefs: {},
    }

    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.section.set_columns',
          arguments: { sectionIndex: 0, count: 4, spacingTwips: 4000 },
        },
        { ...services, getSectionLayoutState: () => [section] },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.section.set_columns',
      ok: false,
      error: 'invalid_arguments',
      message: 'DOCX section column spacing must leave positive text width.',
    })
    expect(editor.state.doc.child(0).attrs.sectionSettingsOverride).toBeNull()
  })

  it('declares an explicit DOCX page-border final-state contract', () => {
    const descriptor = docxOperationCatalog.operations.find(
      (operation) => operation.id === 'docx.section.set_page_border',
    )

    expect(descriptor).toMatchObject({
      family: 'section',
      inputSchema: {
        properties: {
          sectionIndex: { type: 'integer', minimum: 0 },
          enabled: { type: 'boolean' },
        },
        required: ['sectionIndex', 'enabled'],
        additionalProperties: false,
      },
      outputSchema: {
        required: ['sectionIndex', 'enabled', 'changed'],
        additionalProperties: false,
      },
      mutates: true,
      undoable: true,
      atomic: true,
    })
  })

  it('sets explicit DOCX page-border state through the shared section journal', async () => {
    const editor = createEditor('Border')
    const section = {
      settings: {
        pageWidth: 12240,
        pageHeight: 15840,
        marginTop: 1440,
        marginRight: 1440,
        marginBottom: 1440,
        marginLeft: 1440,
        pageBorder: false,
        orientation: 'portrait' as const,
        columns: 1,
      },
      startType: 'nextPage' as const,
      firstBlockIndex: 0,
      lastBlockIndex: 1,
      sectPrXml: '<w:sectPr/>',
      titlePg: false,
      headerRefs: {},
      footerRefs: {},
    }

    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.section.set_page_border',
          arguments: { sectionIndex: 0, enabled: true },
        },
        { ...services, getSectionLayoutState: () => [section] },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.section.set_page_border',
      ok: true,
      output: { sectionIndex: 0, enabled: true, changed: true },
    })
    expect(editor.state.doc.child(0).attrs.sectionSettingsOverride.settings.pageBorder).toBe(true)
    expect(editor.commands.undo()).toBe(true)
    expect(editor.state.doc.child(0).attrs.sectionSettingsOverride).toBeNull()
  })

  it('declares an explicit indexed DOCX different-first-page contract', () => {
    const descriptor = docxOperationCatalog.operations.find(
      (operation) => operation.id === 'docx.section.set_different_first_page',
    )

    expect(descriptor).toMatchObject({
      family: 'header_footer',
      inputSchema: {
        properties: {
          sectionIndex: { type: 'integer', minimum: 0 },
          enabled: { type: 'boolean' },
        },
        required: ['sectionIndex', 'enabled'],
        additionalProperties: false,
      },
      outputSchema: {
        required: ['sectionIndex', 'enabled', 'changed'],
        additionalProperties: false,
      },
      mutates: true,
      undoable: true,
      atomic: true,
    })
  })

  it('sets different-first-page state on an explicit DOCX section and undoes once', async () => {
    const editor = createEditor('Header section')
    const section = {
      settings: {
        pageWidth: 12240,
        pageHeight: 15840,
        marginTop: 1440,
        marginRight: 1440,
        marginBottom: 1440,
        marginLeft: 1440,
        pageBorder: false,
        orientation: 'portrait' as const,
        columns: 1,
      },
      startType: 'nextPage' as const,
      firstBlockIndex: 0,
      lastBlockIndex: 1,
      sectPrXml: '<w:sectPr/>',
      titlePg: false,
      headerRefs: {},
      footerRefs: {},
    }

    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.section.set_different_first_page',
          arguments: { sectionIndex: 0, enabled: true },
        },
        { ...services, getSectionLayoutState: () => [section] },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.section.set_different_first_page',
      ok: true,
      output: { sectionIndex: 0, enabled: true, changed: true },
    })
    expect(editor.state.doc.child(0).attrs.sectionSettingsOverride).toMatchObject({
      sectionIndex: 0,
      titlePg: true,
    })
    expect(editor.commands.undo()).toBe(true)
    expect(editor.state.doc.child(0).attrs.sectionSettingsOverride).toBeNull()
  })

  it('declares an explicit DOCX odd/even-page header final-state contract', () => {
    const descriptor = docxOperationCatalog.operations.find(
      (operation) => operation.id === 'docx.document.set_different_odd_even_pages',
    )

    expect(descriptor).toMatchObject({
      family: 'header_footer',
      inputSchema: { properties: { enabled: { type: 'boolean' } }, required: ['enabled'] },
      outputSchema: {
        required: ['enabled', 'changed'],
        additionalProperties: false,
      },
      mutates: true,
      undoable: true,
      atomic: true,
    })
  })

  it('sets DOCX odd/even-page header state in the native journal and undoes once', async () => {
    const editor = createEditor('Odd/even headers')
    const section = {
      settings: {
        pageWidth: 12240,
        pageHeight: 15840,
        marginTop: 1440,
        marginRight: 1440,
        marginBottom: 1440,
        marginLeft: 1440,
        pageBorder: false,
        orientation: 'portrait' as const,
        columns: 1,
      },
      startType: 'nextPage' as const,
      firstBlockIndex: 0,
      lastBlockIndex: 1,
      sectPrXml: '<w:sectPr/>',
      titlePg: false,
      headerRefs: {},
      footerRefs: {},
    }

    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.document.set_different_odd_even_pages',
          arguments: { enabled: true },
        },
        {
          ...services,
          getSectionLayoutState: () => [section],
          getDifferentOddEvenPages: () => false,
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.document.set_different_odd_even_pages',
      ok: true,
      output: { enabled: true, changed: true },
    })
    expect(editor.state.doc.child(0).attrs.sectionSettingsOverride.evenOddHeaders).toBe(true)
    expect(editor.commands.undo()).toBe(true)
    expect(editor.state.doc.child(0).attrs.sectionSettingsOverride).toBeNull()
  })

  it('declares bounded DOCX section page-numbering final state', () => {
    const descriptor = docxOperationCatalog.operations.find(
      (operation) => operation.id === 'docx.section.set_page_numbering',
    )

    expect(descriptor).toMatchObject({
      family: 'header_footer',
      inputSchema: {
        properties: {
          sectionIndex: { type: 'integer', minimum: 0 },
          format: {
            type: 'string',
            enum: [
              'decimal',
              'numberInDash',
              'lowerLetter',
              'upperLetter',
              'lowerRoman',
              'upperRoman',
              'chineseCounting',
            ],
          },
          start: { type: ['integer', 'null'], minimum: 0, maximum: 999999 },
        },
        required: ['sectionIndex', 'format', 'start'],
        additionalProperties: false,
      },
      outputSchema: {
        required: ['sectionIndex', 'format', 'start', 'changed'],
        additionalProperties: false,
      },
      mutates: true,
      undoable: true,
      atomic: true,
    })
  })

  it('sets page numbering on an explicit DOCX section and undoes once', async () => {
    const editor = createEditor('Page numbering')
    const section = {
      settings: {
        pageWidth: 12240,
        pageHeight: 15840,
        marginTop: 1440,
        marginRight: 1440,
        marginBottom: 1440,
        marginLeft: 1440,
        pageBorder: false,
        orientation: 'portrait' as const,
        columns: 1,
      },
      startType: 'nextPage' as const,
      firstBlockIndex: 0,
      lastBlockIndex: 1,
      sectPrXml: '<w:sectPr/>',
      titlePg: false,
      headerRefs: {},
      footerRefs: {},
    }

    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.section.set_page_numbering',
          arguments: { sectionIndex: 0, format: 'upperRoman', start: 3 },
        },
        { ...services, getSectionLayoutState: () => [section] },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.section.set_page_numbering',
      ok: true,
      output: {
        sectionIndex: 0,
        format: 'upperRoman',
        start: 3,
        changed: true,
      },
    })
    expect(editor.state.doc.child(0).attrs.sectionSettingsOverride.pageNumbering).toEqual({
      format: 'upperRoman',
      start: 3,
    })
    expect(editor.commands.undo()).toBe(true)
    expect(editor.state.doc.child(0).attrs.sectionSettingsOverride).toBeNull()
  })

  it('declares explicit DOCX header/footer text targets', () => {
    const descriptor = docxOperationCatalog.operations.find(
      (operation) => operation.id === 'docx.header_footer.set_text',
    )

    expect(descriptor).toMatchObject({
      family: 'header_footer',
      inputSchema: {
        properties: {
          sectionIndex: { type: 'integer', minimum: 0 },
          kind: { type: 'string', enum: ['header', 'footer'] },
          variant: { type: 'string', enum: ['default', 'first', 'even'] },
          text: { type: 'string' },
        },
        required: ['sectionIndex', 'kind', 'variant', 'text'],
        additionalProperties: false,
      },
      outputSchema: {
        required: ['sectionIndex', 'kind', 'variant', 'changed'],
        additionalProperties: false,
      },
      mutates: true,
      undoable: true,
      atomic: true,
    })
  })

  it('sets explicit DOCX header text in the native journal and undoes once', async () => {
    const editor = createEditor('Header text')
    const section = {
      settings: {
        pageWidth: 12240,
        pageHeight: 15840,
        marginTop: 1440,
        marginRight: 1440,
        marginBottom: 1440,
        marginLeft: 1440,
        pageBorder: false,
        orientation: 'portrait' as const,
        columns: 1,
      },
      startType: 'nextPage' as const,
      firstBlockIndex: 0,
      lastBlockIndex: 1,
      sectPrXml: '<w:sectPr/>',
      titlePg: false,
      headerRefs: {},
      footerRefs: {},
    }

    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.header_footer.set_text',
          arguments: {
            sectionIndex: 0,
            kind: 'header',
            variant: 'first',
            text: 'Confidential',
          },
        },
        { ...services, getSectionLayoutState: () => [section] },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.header_footer.set_text',
      ok: true,
      output: {
        sectionIndex: 0,
        kind: 'header',
        variant: 'first',
        changed: true,
      },
    })
    expect(editor.state.doc.child(0).attrs.sectionSettingsOverride.headerFooterEdits).toEqual([
      {
        kind: 'header',
        variant: 'first',
        value: { text: 'Confidential' },
      },
    ])
    expect(editor.commands.undo()).toBe(true)
    expect(editor.state.doc.child(0).attrs.sectionSettingsOverride).toBeNull()
  })

  it('declares explicit DOCX page-number placement state', () => {
    const descriptor = docxOperationCatalog.operations.find(
      (operation) => operation.id === 'docx.header_footer.set_page_number',
    )

    expect(descriptor).toMatchObject({
      family: 'header_footer',
      inputSchema: {
        properties: {
          sectionIndex: { type: 'integer', minimum: 0 },
          kind: { type: 'string', enum: ['header', 'footer'] },
          variant: { type: 'string', enum: ['default', 'first', 'even'] },
          enabled: { type: 'boolean' },
          alignment: { type: 'string', enum: ['left', 'center', 'right'] },
        },
        required: ['sectionIndex', 'kind', 'variant', 'enabled', 'alignment'],
        additionalProperties: false,
      },
      outputSchema: {
        required: ['sectionIndex', 'kind', 'variant', 'enabled', 'alignment', 'changed'],
        additionalProperties: false,
      },
      mutates: true,
      undoable: true,
      atomic: true,
    })
  })

  it('places a DOCX page-number field in the native content journal and undoes once', async () => {
    const editor = createEditor('Page number')
    const section = {
      settings: {
        pageWidth: 12240,
        pageHeight: 15840,
        marginTop: 1440,
        marginRight: 1440,
        marginBottom: 1440,
        marginLeft: 1440,
        pageBorder: false,
        orientation: 'portrait' as const,
        columns: 1,
      },
      startType: 'nextPage' as const,
      firstBlockIndex: 0,
      lastBlockIndex: 1,
      sectPrXml: '<w:sectPr/>',
      titlePg: false,
      headerRefs: {},
      footerRefs: {},
    }

    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.header_footer.set_page_number',
          arguments: {
            sectionIndex: 0,
            kind: 'footer',
            variant: 'default',
            enabled: true,
            alignment: 'right',
          },
        },
        {
          ...services,
          getSectionLayoutState: () => [section],
          getHeaderFooterValue: () => ({ text: 'Legal' }),
        },
      ),
    ).resolves.toMatchObject({
      handled: true,
      operationId: 'docx.header_footer.set_page_number',
      ok: true,
      output: {
        sectionIndex: 0,
        kind: 'footer',
        variant: 'default',
        enabled: true,
        alignment: 'right',
        changed: true,
      },
    })
    expect(
      editor.state.doc.child(0).attrs.sectionSettingsOverride.headerFooterEdits[0],
    ).toMatchObject({
      kind: 'footer',
      variant: 'default',
      value: { pageNumber: true, paras: [{ align: 'right' }] },
    })
    expect(editor.commands.undo()).toBe(true)
    expect(editor.state.doc.child(0).attrs.sectionSettingsOverride).toBeNull()
  })

  it('declares bounded DOCX header/footer paragraph content', () => {
    const descriptor = docxOperationCatalog.operations.find(
      (operation) => operation.id === 'docx.header_footer.set_paragraphs',
    )

    expect(descriptor).toMatchObject({
      family: 'header_footer',
      inputSchema: {
        properties: {
          sectionIndex: { type: 'integer', minimum: 0 },
          kind: { type: 'string', enum: ['header', 'footer'] },
          variant: { type: 'string', enum: ['default', 'first', 'even'] },
          paragraphs: {
            type: 'array',
            minItems: 1,
            maxItems: 64,
            items: {
              properties: {
                alignment: {
                  type: 'string',
                  enum: ['left', 'center', 'right', 'justify'],
                },
                segments: { type: 'array', minItems: 0, maxItems: 256 },
              },
              required: ['alignment', 'segments'],
              additionalProperties: false,
            },
          },
        },
        required: ['sectionIndex', 'kind', 'variant', 'paragraphs'],
        additionalProperties: false,
      },
      outputSchema: {
        required: ['sectionIndex', 'kind', 'variant', 'paragraphCount', 'changed'],
        additionalProperties: false,
      },
      mutates: true,
      undoable: true,
      atomic: true,
    })
  })

  it('sets rich DOCX header/footer paragraphs with explicit field tokens and undoes once', async () => {
    const editor = createEditor('Rich header')
    const section = {
      settings: {
        pageWidth: 12240,
        pageHeight: 15840,
        marginTop: 1440,
        marginRight: 1440,
        marginBottom: 1440,
        marginLeft: 1440,
        pageBorder: false,
        orientation: 'portrait' as const,
        columns: 1,
      },
      startType: 'nextPage' as const,
      firstBlockIndex: 0,
      lastBlockIndex: 1,
      sectPrXml: '<w:sectPr/>',
      titlePg: false,
      headerRefs: {},
      footerRefs: {},
    }
    const arguments_ = {
      sectionIndex: 0,
      kind: 'footer',
      variant: 'even',
      paragraphs: [
        {
          alignment: 'right',
          segments: [
            { type: 'text', text: 'Page ', bold: true, color: '112233' },
            { type: 'page', text: '' },
            { type: 'text', text: ' / ' },
            { type: 'total_pages', text: '' },
          ],
        },
      ],
    }

    await expect(
      executeDocxOperation(
        editor,
        { operation: 'docx.header_footer.set_paragraphs', arguments: arguments_ },
        { ...services, getSectionLayoutState: () => [section] },
      ),
    ).resolves.toMatchObject({
      handled: true,
      operationId: 'docx.header_footer.set_paragraphs',
      ok: true,
      output: {
        sectionIndex: 0,
        kind: 'footer',
        variant: 'even',
        paragraphCount: 1,
        changed: true,
      },
    })
    expect(
      editor.state.doc.child(0).attrs.sectionSettingsOverride.headerFooterEdits[0].value,
    ).toEqual({
      text: `Page ${PAGE_MARK} / ${TOTAL_PAGES_MARK}`,
      pageNumber: true,
      paras: [
        {
          align: 'right',
          runs: [
            { text: 'Page ', bold: true, color: '112233' },
            { text: PAGE_MARK },
            { text: ' / ' },
            { text: TOTAL_PAGES_MARK },
          ],
        },
      ],
    })
    expect(editor.commands.undo()).toBe(true)
    expect(editor.state.doc.child(0).attrs.sectionSettingsOverride).toBeNull()
  })

  it('declares bounded DOCX undo and redo operation contracts', () => {
    const descriptors = Object.fromEntries(
      docxOperationCatalog.operations.map((operation) => [operation.id, operation]),
    )

    expect(descriptors['docx.history.undo']).toMatchObject({
      family: 'history',
      inputSchema: { properties: {}, required: [], additionalProperties: false },
      outputSchema: {
        properties: { undone: { type: 'boolean', enum: [true] } },
        required: ['undone'],
        additionalProperties: false,
      },
      mutates: true,
      undoable: true,
      atomic: true,
    })
    expect(descriptors['docx.history.redo']).toMatchObject({
      family: 'history',
      inputSchema: { properties: {}, required: [], additionalProperties: false },
      outputSchema: {
        properties: { redone: { type: 'boolean', enum: [true] } },
        required: ['redone'],
        additionalProperties: false,
      },
      mutates: true,
      undoable: true,
      atomic: true,
    })
  })

  it('undoes and redoes DOCX edits through the mounted native history', async () => {
    const editor = createEditor('Hello')
    editor.commands.setTextSelection(6)
    editor.commands.insertContent(' world')
    expect(editor.getText()).toContain('Hello world')

    await expect(
      executeDocxOperation(editor, { operation: 'docx.history.undo', arguments: {} }, services),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.history.undo',
      ok: true,
      output: { undone: true },
    })
    expect(editor.getText()).toContain('Hello')
    expect(editor.getText()).not.toContain(' world')

    await expect(
      executeDocxOperation(editor, { operation: 'docx.history.redo', arguments: {} }, services),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.history.redo',
      ok: true,
      output: { redone: true },
    })
    expect(editor.getText()).toContain('Hello world')
  })

  it.each([
    ['docx.history.undo', 'undo'],
    ['docx.history.redo', 'redo'],
  ] as const)('rejects %s when native history has no %s entry', async (operation, direction) => {
    const editor = createEditor('Unchanged')

    await expect(
      executeDocxOperation(editor, { operation, arguments: {} }, services),
    ).resolves.toEqual({
      handled: true,
      operationId: operation,
      ok: false,
      error: 'execution_failed',
      message: `docx.history.${direction} requires an available ${direction} entry.`,
    })
    expect(editor.getText()).toContain('Unchanged')
  })

  it('declares an exact bounded DOCX character-format contract', () => {
    const descriptor = docxOperationCatalog.operations.find(
      (operation) => operation.id === 'docx.text.set_character_format',
    )

    expect(descriptor).toMatchObject({
      family: 'text',
      inputSchema: {
        properties: {
          range: {
            properties: {
              from: { type: 'integer', minimum: 1 },
              to: { type: 'integer', minimum: 1 },
            },
            required: ['from', 'to'],
            additionalProperties: false,
          },
          format: {
            properties: {
              bold: { type: 'boolean' },
              italic: { type: 'boolean' },
              underline: { type: 'boolean' },
              strike: { type: 'boolean' },
              verticalAlign: {
                type: 'string',
                enum: ['baseline', 'superscript', 'subscript'],
              },
              color: { type: ['string', 'null'] },
              highlight: {
                type: ['string', 'null'],
                enum: [
                  'yellow',
                  'green',
                  'cyan',
                  'magenta',
                  'blue',
                  'red',
                  'darkBlue',
                  'darkCyan',
                  'darkGreen',
                  'darkMagenta',
                  'darkRed',
                  'darkYellow',
                  'darkGray',
                  'lightGray',
                  'black',
                  'white',
                  null,
                ],
              },
              fontFamily: { type: ['string', 'null'] },
              fontSizePoints: { type: ['number', 'null'], minimum: 1, maximum: 1638 },
            },
            additionalProperties: false,
          },
        },
        required: ['range', 'format', 'fields'],
        additionalProperties: false,
      },
      outputSchema: {
        required: ['from', 'to', 'changed'],
        additionalProperties: false,
      },
      mutates: true,
      undoable: true,
      atomic: true,
    })
  })

  it('sets exact DOCX character marks without changing text outside the explicit range', async () => {
    const editor = createEditor('Hello world')

    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.text.set_character_format',
          arguments: {
            range: { from: 7, to: 12 },
            format: {
              bold: true,
              italic: true,
              underline: false,
              strike: true,
              verticalAlign: 'superscript',
            },
            fields: ['bold', 'italic', 'underline', 'strike', 'verticalAlign'],
          },
        },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.text.set_character_format',
      ok: true,
      output: { from: 7, to: 12, changed: true },
    })

    const paragraph = editor.state.doc.child(0)
    const hello = paragraph.child(0)
    const world = paragraph.child(1)
    expect(hello.text).toBe('Hello ')
    expect(hello.marks).toEqual([])
    expect(world.text).toBe('world')
    expect(world.marks.map((mark) => mark.type.name)).toEqual([
      'bold',
      'italic',
      'strike',
      'docTextStyle',
    ])
    expect(world.marks.find((mark) => mark.type.name === 'docTextStyle')?.attrs.vertAlign).toBe(
      'superscript',
    )
    expect(editor.state.selection).toMatchObject({ from: 7, to: 12 })

    expect(editor.commands.undo()).toBe(true)
    expect(editor.state.doc.child(0).childCount).toBe(1)
    expect(editor.state.doc.child(0).child(0).marks).toEqual([])
  })

  it('clears selected DOCX character marks while preserving unlisted run style', async () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docParagraph',
            content: [
              {
                type: 'text',
                text: 'Styled',
                marks: [
                  { type: 'bold' },
                  { type: 'underline' },
                  {
                    type: 'docTextStyle',
                    attrs: { color: 'FF0000', vertAlign: 'subscript' },
                  },
                ],
              },
            ],
          },
        ],
      },
    })
    editors.push(editor)

    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.text.set_character_format',
          arguments: {
            range: { from: 1, to: 7 },
            format: { bold: false, verticalAlign: 'baseline' },
            fields: ['bold', 'verticalAlign'],
          },
        },
        services,
      ),
    ).resolves.toMatchObject({
      handled: true,
      operationId: 'docx.text.set_character_format',
      ok: true,
      output: { changed: true },
    })

    const marks = editor.state.doc.child(0).child(0).marks
    expect(marks.some((mark) => mark.type.name === 'bold')).toBe(false)
    expect(marks.some((mark) => mark.type.name === 'underline')).toBe(true)
    expect(marks.find((mark) => mark.type.name === 'docTextStyle')?.attrs).toMatchObject({
      color: 'FF0000',
      vertAlign: null,
    })
  })

  it('sets bounded DOCX font, size, color, and highlight without flattening run metadata', async () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docParagraph',
            content: [
              {
                type: 'text',
                text: 'A',
                marks: [
                  {
                    type: 'docTextStyle',
                    attrs: { color: 'AA0000', styleId: 'First', charSpacingTwips: 20 },
                  },
                ],
              },
              {
                type: 'text',
                text: 'B',
                marks: [
                  {
                    type: 'docTextStyle',
                    attrs: { color: '00AA00', styleId: 'Second', charSpacingTwips: 40 },
                  },
                ],
              },
            ],
          },
        ],
      },
    })
    editors.push(editor)

    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.text.set_character_format',
          arguments: {
            range: { from: 1, to: 3 },
            format: {
              color: '#112233',
              highlight: 'yellow',
              fontFamily: 'Arial',
              fontSizePoints: 12.5,
            },
            fields: ['color', 'highlight', 'fontFamily', 'fontSizePoints'],
          },
        },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.text.set_character_format',
      ok: true,
      output: { from: 1, to: 3, changed: true },
    })

    const first = editor.state.doc.child(0).child(0).marks[0].attrs
    const second = editor.state.doc.child(0).child(1).marks[0].attrs
    expect(first).toMatchObject({
      color: '112233',
      highlight: 'yellow',
      fontAscii: 'Arial',
      sizeHalfPoints: 25,
      styleId: 'First',
      charSpacingTwips: 20,
    })
    expect(second).toMatchObject({
      color: '112233',
      highlight: 'yellow',
      fontAscii: 'Arial',
      sizeHalfPoints: 25,
      styleId: 'Second',
      charSpacingTwips: 40,
    })

    expect(editor.commands.undo()).toBe(true)
    expect(editor.state.doc.child(0).child(0).marks[0].attrs).toMatchObject({
      color: 'AA0000',
      styleId: 'First',
      charSpacingTwips: 20,
    })
    expect(editor.state.doc.child(0).child(1).marks[0].attrs).toMatchObject({
      color: '00AA00',
      styleId: 'Second',
      charSpacingTwips: 40,
    })
  })

  it('clears exact DOCX font, size, color, and highlight fields', async () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docParagraph',
            content: [
              {
                type: 'text',
                text: 'A',
                marks: [
                  {
                    type: 'docTextStyle',
                    attrs: {
                      color: '112233',
                      highlight: 'yellow',
                      font: 'Noto Sans CJK SC',
                      fontAscii: 'Arial',
                      sizeHalfPoints: 25,
                      styleId: 'Keep',
                    },
                  },
                ],
              },
            ],
          },
        ],
      },
    })
    editors.push(editor)

    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.text.set_character_format',
          arguments: {
            range: { from: 1, to: 2 },
            format: { color: null, highlight: null, fontFamily: null, fontSizePoints: null },
            fields: ['color', 'highlight', 'fontFamily', 'fontSizePoints'],
          },
        },
        services,
      ),
    ).resolves.toMatchObject({
      handled: true,
      operationId: 'docx.text.set_character_format',
      ok: true,
      output: { changed: true },
    })

    expect(editor.state.doc.child(0).child(0).marks[0].attrs).toMatchObject({
      color: null,
      highlight: null,
      font: null,
      fontAscii: null,
      sizeHalfPoints: null,
      styleId: 'Keep',
    })
  })

  it.each([
    [{ color: '112233' }, ['color']],
    [{ fontFamily: '' }, ['fontFamily']],
    [{ fontFamily: 'x'.repeat(129) }, ['fontFamily']],
    [{ fontSizePoints: 12.25 }, ['fontSizePoints']],
  ] as const)('rejects invalid bounded DOCX character format %o', async (format, fields) => {
    const editor = createEditor('Unchanged')

    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.text.set_character_format',
          arguments: { range: { from: 1, to: 10 }, format, fields },
        },
        services,
      ),
    ).resolves.toMatchObject({
      handled: true,
      operationId: 'docx.text.set_character_format',
      ok: false,
      error: 'invalid_arguments',
    })
    expect(editor.state.doc.child(0).child(0).marks).toEqual([])
  })

  it('rejects mismatched DOCX character-format fields without mutation', async () => {
    const editor = createEditor('Unchanged')

    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.text.set_character_format',
          arguments: {
            range: { from: 1, to: 10 },
            format: { bold: true },
            fields: ['italic'],
          },
        },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.text.set_character_format',
      ok: false,
      error: 'invalid_arguments',
      message: 'docx.text.set_character_format fields must uniquely and exactly match format.',
    })
    expect(editor.state.doc.child(0).child(0).marks).toEqual([])
  })

  it('rejects an invalid DOCX character range without mutation', async () => {
    const editor = createEditor('Short')

    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.text.set_character_format',
          arguments: {
            range: { from: 2, to: 20 },
            format: { bold: true },
            fields: ['bold'],
          },
        },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.text.set_character_format',
      ok: false,
      error: 'invalid_arguments',
      message:
        'docx.text.set_character_format requires a non-empty text range inside the document.',
    })
    expect(editor.state.doc.child(0).child(0).marks).toEqual([])
  })

  it('declares an explicit bounded DOCX clear-character-format contract', () => {
    const descriptor = docxOperationCatalog.operations.find(
      (operation) => operation.id === 'docx.text.clear_character_format',
    )

    expect(descriptor).toMatchObject({
      family: 'text',
      inputSchema: {
        properties: {
          range: {
            properties: {
              from: { type: 'integer', minimum: 1 },
              to: { type: 'integer', minimum: 1 },
            },
            required: ['from', 'to'],
            additionalProperties: false,
          },
        },
        required: ['range'],
        additionalProperties: false,
      },
      outputSchema: {
        required: ['from', 'to', 'changed'],
        additionalProperties: false,
      },
      mutates: true,
      undoable: true,
      atomic: true,
    })
  })

  it('clears DOCX character formatting only inside the explicit range and undoes once', async () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docParagraph',
            content: [
              {
                type: 'text',
                text: 'A',
                marks: [{ type: 'bold' }, { type: 'docTextStyle', attrs: { color: 'FF0000' } }],
              },
              { type: 'text', text: 'B', marks: [{ type: 'italic' }] },
            ],
          },
        ],
      },
    })
    editors.push(editor)

    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.text.clear_character_format',
          arguments: { range: { from: 1, to: 2 } },
        },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.text.clear_character_format',
      ok: true,
      output: { from: 1, to: 2, changed: true },
    })

    expect(editor.state.doc.child(0).child(0).marks).toEqual([])
    expect(
      editor.state.doc
        .child(0)
        .child(1)
        .marks.map((mark) => mark.type.name),
    ).toEqual(['italic'])
    expect(editor.state.selection).toMatchObject({ from: 1, to: 2 })

    expect(editor.commands.undo()).toBe(true)
    expect(
      editor.state.doc
        .child(0)
        .child(0)
        .marks.map((mark) => mark.type.name),
    ).toEqual(['bold', 'docTextStyle'])
    expect(
      editor.state.doc
        .child(0)
        .child(1)
        .marks.map((mark) => mark.type.name),
    ).toEqual(['italic'])
  })

  it('rejects an invalid DOCX clear-character-format range without mutation', async () => {
    const editor = createEditor('Short')

    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.text.clear_character_format',
          arguments: { range: { from: 2, to: 20 } },
        },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.text.clear_character_format',
      ok: false,
      error: 'invalid_arguments',
      message:
        'docx.text.clear_character_format requires a non-empty text range inside the document.',
    })
    expect(editor.state.doc.child(0).child(0).marks).toEqual([])
  })

  it('declares a bounded DOCX text case-transform contract', () => {
    const descriptor = docxOperationCatalog.operations.find(
      (operation) => operation.id === 'docx.text.transform_case',
    )

    expect(descriptor).toMatchObject({
      family: 'text',
      inputSchema: {
        properties: {
          range: {
            properties: {
              from: { type: 'integer', minimum: 1 },
              to: { type: 'integer', minimum: 1 },
            },
            required: ['from', 'to'],
            additionalProperties: false,
          },
          mode: { type: 'string', enum: ['sentence', 'lower', 'upper', 'title'] },
        },
        required: ['range', 'mode'],
        additionalProperties: false,
      },
      outputSchema: {
        required: ['from', 'to', 'mode', 'changed'],
        additionalProperties: false,
      },
      mutates: true,
      undoable: true,
      atomic: true,
    })
  })

  it.each([
    ['sentence', 'Hello world. Next'],
    ['lower', 'hello world. next'],
    ['upper', 'HELLO WORLD. NEXT'],
    ['title', 'Hello World. Next'],
  ] as const)('transforms an exact DOCX range using %s case', async (mode, expected) => {
    const source = 'hELLO wORLD. nEXT'
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docParagraph',
            content: [
              {
                type: 'text',
                text: source,
                marks: [{ type: 'bold' }, { type: 'docTextStyle', attrs: { color: 'FF0000' } }],
              },
            ],
          },
        ],
      },
    })
    editors.push(editor)

    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.text.transform_case',
          arguments: { range: { from: 1, to: source.length + 1 }, mode },
        },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.text.transform_case',
      ok: true,
      output: { from: 1, to: source.length + 1, mode, changed: true },
    })
    expect(editor.state.doc.child(0).textContent).toBe(expected)
    expect(
      editor.state.doc
        .child(0)
        .child(0)
        .marks.map((mark) => mark.type.name),
    ).toEqual(['bold', 'docTextStyle'])

    expect(editor.commands.undo()).toBe(true)
    expect(editor.state.doc.child(0).textContent).toBe(source)
  })

  it('maps a length-changing Unicode DOCX case transform as one transaction', async () => {
    const editor = createEditor('straße')

    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.text.transform_case',
          arguments: { range: { from: 1, to: 7 }, mode: 'upper' },
        },
        services,
      ),
    ).resolves.toMatchObject({
      handled: true,
      operationId: 'docx.text.transform_case',
      ok: true,
      output: { changed: true },
    })
    expect(editor.state.doc.child(0).textContent).toBe('STRASSE')
    expect(editor.state.selection).toMatchObject({ from: 1, to: 8 })

    expect(editor.commands.undo()).toBe(true)
    expect(editor.state.doc.child(0).textContent).toBe('straße')
  })

  it('declares an explicit DOCX character-style contract', () => {
    const descriptor = docxOperationCatalog.operations.find(
      (operation) => operation.id === 'docx.text.set_character_style',
    )

    expect(descriptor).toMatchObject({
      family: 'text',
      inputSchema: {
        properties: {
          range: {
            properties: {
              from: { type: 'integer', minimum: 1 },
              to: { type: 'integer', minimum: 1 },
            },
            required: ['from', 'to'],
            additionalProperties: false,
          },
          styleId: { type: ['string', 'null'] },
        },
        required: ['range', 'styleId'],
        additionalProperties: false,
      },
      outputSchema: {
        required: ['from', 'to', 'styleId', 'changed'],
        additionalProperties: false,
      },
      mutates: true,
      undoable: true,
      atomic: true,
    })
  })

  it('applies a document-owned DOCX character style without flattening direct format', async () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docParagraph',
            content: [
              {
                type: 'text',
                text: 'Styled',
                marks: [
                  {
                    type: 'docTextStyle',
                    attrs: { color: 'FF0000', charSpacingTwips: 20 },
                  },
                ],
              },
            ],
          },
        ],
      },
    })
    editors.push(editor)

    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.text.set_character_style',
          arguments: { range: { from: 1, to: 7 }, styleId: 'Emphasis' },
        },
        { ...services, hasCharacterStyle: (styleId) => styleId === 'Emphasis' },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.text.set_character_style',
      ok: true,
      output: { from: 1, to: 7, styleId: 'Emphasis', changed: true },
    })
    expect(editor.state.doc.child(0).child(0).marks[0].attrs).toMatchObject({
      styleId: 'Emphasis',
      color: 'FF0000',
      charSpacingTwips: 20,
    })

    expect(editor.commands.undo()).toBe(true)
    expect(editor.state.doc.child(0).child(0).marks[0].attrs).toMatchObject({
      styleId: null,
      color: 'FF0000',
      charSpacingTwips: 20,
    })
  })

  it('removes a DOCX character style using the retained Ribbon mark-removal semantics', async () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docParagraph',
            content: [
              {
                type: 'text',
                text: 'Styled',
                marks: [
                  {
                    type: 'docTextStyle',
                    attrs: { styleId: 'Emphasis', color: 'FF0000' },
                  },
                ],
              },
            ],
          },
        ],
      },
    })
    editors.push(editor)

    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.text.set_character_style',
          arguments: { range: { from: 1, to: 7 }, styleId: null },
        },
        services,
      ),
    ).resolves.toMatchObject({
      handled: true,
      operationId: 'docx.text.set_character_style',
      ok: true,
      output: { styleId: null, changed: true },
    })
    expect(editor.state.doc.child(0).child(0).marks).toEqual([])
  })

  it('rejects a DOCX character style absent from the current document', async () => {
    const editor = createEditor('Unchanged')

    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.text.set_character_style',
          arguments: { range: { from: 1, to: 10 }, styleId: 'Missing' },
        },
        { ...services, hasCharacterStyle: () => false },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.text.set_character_style',
      ok: false,
      error: 'invalid_arguments',
      message:
        'docx.text.set_character_style requires a character style from the current document.',
    })
    expect(editor.state.doc.child(0).child(0).marks).toEqual([])
  })

  it('declares a bounded DOCX paragraph-direction contract', () => {
    const descriptor = docxOperationCatalog.operations.find(
      (operation) => operation.id === 'docx.paragraph.set_direction',
    )

    expect(descriptor).toMatchObject({
      family: 'paragraph',
      inputSchema: {
        properties: {
          target: docxBlockTargetSchema,
          direction: { type: 'string', enum: ['ltr', 'rtl'] },
        },
        required: ['target', 'direction'],
        additionalProperties: false,
      },
      outputSchema: {
        required: ['summary', 'matched', 'changed', 'skippedProtected'],
        additionalProperties: false,
      },
      mutates: true,
      undoable: true,
      atomic: true,
    })
  })

  it('sets DOCX paragraph direction and flips explicit alignment in one undo transaction', async () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docParagraph',
            attrs: { bidi: false, align: 'left' },
            content: [{ type: 'text', text: 'First' }],
          },
          {
            type: 'docHeading',
            attrs: { level: 2, bidi: false, align: 'right' },
            content: [{ type: 'text', text: 'Second' }],
          },
          {
            type: 'docParagraph',
            attrs: { bidi: false, align: 'center' },
            content: [{ type: 'text', text: 'Untouched' }],
          },
        ],
      },
    })
    editors.push(editor)

    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.paragraph.set_direction',
          arguments: { target: { blockIndexes: [0, 1] }, direction: 'rtl' },
        },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.paragraph.set_direction',
      ok: true,
      output: {
        summary: 'Updated paragraph direction in 2 block(s)',
        matched: 2,
        changed: 2,
        skippedProtected: 0,
      },
    })
    expect(editor.state.doc.child(0).attrs).toMatchObject({ bidi: true, align: 'right' })
    expect(editor.state.doc.child(1).attrs).toMatchObject({ bidi: true, align: 'left' })
    expect(editor.state.doc.child(2).attrs).toMatchObject({ bidi: false, align: 'center' })

    expect(editor.commands.undo()).toBe(true)
    expect(editor.state.doc.child(0).attrs).toMatchObject({ bidi: false, align: 'left' })
    expect(editor.state.doc.child(1).attrs).toMatchObject({ bidi: false, align: 'right' })
  })

  it('rejects a DOCX paragraph-direction target without a condition', async () => {
    const editor = createEditor('Unchanged')

    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.paragraph.set_direction',
          arguments: { target: {}, direction: 'rtl' },
        },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.paragraph.set_direction',
      ok: false,
      error: 'invalid_arguments',
      message: 'command #0 setParagraphDirection: target requires at least one condition',
    })
    expect(editor.state.doc.child(0).attrs.bidi).toBe(false)
  })

  it('declares a bounded absolute DOCX list-level contract', () => {
    const descriptor = docxOperationCatalog.operations.find(
      (operation) => operation.id === 'docx.list.set_level',
    )

    expect(descriptor).toMatchObject({
      family: 'list',
      inputSchema: {
        properties: {
          target: docxBlockTargetSchema,
          level: { type: 'integer', minimum: 0, maximum: 8 },
        },
        required: ['target', 'level'],
        additionalProperties: false,
      },
      outputSchema: {
        required: ['summary', 'matched', 'changed', 'skippedProtected'],
        additionalProperties: false,
      },
      mutates: true,
      undoable: true,
      atomic: true,
    })
  })

  it('sets multiple DOCX list items to one absolute level and undoes once', async () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docListItem',
            attrs: { kind: 'ordered', numId: '7', ilvl: 0 },
            content: [{ type: 'text', text: 'First' }],
          },
          {
            type: 'docListItem',
            attrs: { kind: 'ordered', numId: '7', ilvl: 1 },
            content: [{ type: 'text', text: 'Second' }],
          },
          { type: 'docParagraph', content: [{ type: 'text', text: 'Untouched' }] },
        ],
      },
    })
    editors.push(editor)

    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.list.set_level',
          arguments: { target: { blockIndexes: [0, 1, 2] }, level: 3 },
        },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.list.set_level',
      ok: true,
      output: {
        summary: 'Updated list level in 2 block(s)',
        matched: 3,
        changed: 2,
        skippedProtected: 0,
      },
    })
    expect(editor.state.doc.child(0).attrs.ilvl).toBe(3)
    expect(editor.state.doc.child(1).attrs.ilvl).toBe(3)
    expect(editor.state.doc.child(2).type.name).toBe('docParagraph')

    expect(editor.commands.undo()).toBe(true)
    expect(editor.state.doc.child(0).attrs.ilvl).toBe(0)
    expect(editor.state.doc.child(1).attrs.ilvl).toBe(1)
  })

  it('rejects a DOCX list-level target without a condition', async () => {
    const editor = createEditor('Unchanged')

    await expect(
      executeDocxOperation(
        editor,
        { operation: 'docx.list.set_level', arguments: { target: {}, level: 2 } },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.list.set_level',
      ok: false,
      error: 'invalid_arguments',
      message: 'command #0 setListLevel: target requires at least one condition',
    })
  })

  it('declares a bounded DOCX list-preset contract', () => {
    const descriptor = docxOperationCatalog.operations.find(
      (operation) => operation.id === 'docx.list.apply_preset',
    )

    expect(descriptor).toMatchObject({
      family: 'list',
      inputSchema: {
        properties: {
          target: docxBlockTargetSchema,
          levels: {
            type: 'array',
            minItems: 1,
            maxItems: 9,
            items: {
              properties: {
                numFmt: {
                  type: 'string',
                  enum: [
                    'decimal',
                    'bullet',
                    'lowerLetter',
                    'upperLetter',
                    'lowerRoman',
                    'upperRoman',
                    'chineseCountingThousand',
                  ],
                },
                lvlText: { type: 'string' },
                indentLeft: { type: 'integer', minimum: 0, maximum: 31680 },
                hanging: { type: 'integer', minimum: 0, maximum: 31680 },
                start: { type: 'integer', minimum: 1, maximum: 1000000 },
              },
              required: ['numFmt', 'lvlText', 'indentLeft'],
              additionalProperties: false,
            },
          },
        },
        required: ['target', 'levels'],
        additionalProperties: false,
      },
      mutates: true,
      undoable: true,
      atomic: true,
    })
  })

  it('creates and applies one DOCX list preset through the mounted numbering state', async () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          { type: 'docParagraph', content: [{ type: 'text', text: 'First' }] },
          {
            type: 'docListItem',
            attrs: { kind: 'bullet', numId: '7', ilvl: 1 },
            content: [{ type: 'text', text: 'Second' }],
          },
        ],
      },
    })
    editors.push(editor)
    const levels = [
      { numFmt: 'decimal', lvlText: '%1.', indentLeft: 720, hanging: 360, start: 1 },
      { numFmt: 'lowerLetter', lvlText: '%2)', indentLeft: 1440, hanging: 360, start: 1 },
    ]
    let created: unknown

    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.list.apply_preset',
          arguments: { target: { blockIndexes: [0, 1] }, levels },
        },
        {
          ...services,
          createListDef: (input) => {
            created = input
            return '99'
          },
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.list.apply_preset',
      ok: true,
      output: {
        summary: 'Applied a 2-level list preset to 2 block(s)',
        matched: 2,
        changed: 2,
        skippedProtected: 0,
      },
    })
    expect(created).toEqual(levels)
    expect(editor.state.doc.child(0).type.name).toBe('docListItem')
    expect(editor.state.doc.child(0).attrs).toMatchObject({
      kind: 'ordered',
      numId: '99',
      ilvl: 0,
    })
    expect(editor.state.doc.child(1).attrs).toMatchObject({
      kind: 'ordered',
      numId: '99',
      ilvl: 1,
    })

    expect(editor.commands.undo()).toBe(true)
    expect(editor.state.doc.child(0).type.name).toBe('docParagraph')
    expect(editor.state.doc.child(1).attrs).toMatchObject({ kind: 'bullet', numId: '7', ilvl: 1 })
  })

  it('rejects an empty DOCX list preset before creating numbering state', async () => {
    const editor = createEditor('Unchanged')
    let created = false

    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.list.apply_preset',
          arguments: { target: { blockIndexes: [0] }, levels: [] },
        },
        {
          ...services,
          createListDef: () => {
            created = true
            return '99'
          },
        },
      ),
    ).resolves.toMatchObject({
      handled: true,
      operationId: 'docx.list.apply_preset',
      ok: false,
      error: 'invalid_arguments',
    })
    expect(created).toBe(false)
  })

  it('declares an explicit bounded DOCX list-restart contract', () => {
    const descriptor = docxOperationCatalog.operations.find(
      (operation) => operation.id === 'docx.list.restart',
    )

    expect(descriptor).toMatchObject({
      family: 'list',
      inputSchema: {
        properties: {
          blockIndex: { type: 'integer', minimum: 0 },
          start: { type: 'integer', minimum: 1, maximum: 1000000 },
        },
        required: ['blockIndex', 'start'],
        additionalProperties: false,
      },
      outputSchema: {
        required: ['summary', 'matched', 'changed', 'skippedProtected'],
        additionalProperties: false,
      },
      mutates: true,
      undoable: true,
      atomic: true,
    })
  })

  it('restarts DOCX numbering at one explicit block and start value', async () => {
    const editor = createEditor('Unchanged')
    let received: unknown

    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.list.restart',
          arguments: { blockIndex: 4, start: 3 },
        },
        {
          ...services,
          restartList: (input) => {
            received = input
            return { ok: true, changed: 2 }
          },
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.list.restart',
      ok: true,
      output: {
        summary: 'Restarted numbering at block 4 from 3 across 2 item(s)',
        matched: 1,
        changed: 2,
        skippedProtected: 0,
      },
    })
    expect(received).toEqual({ blockIndex: 4, start: 3 })
  })

  it('rejects an invalid DOCX list-restart start before invoking numbering state', async () => {
    const editor = createEditor('Unchanged')
    let invoked = false

    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.list.restart',
          arguments: { blockIndex: 0, start: 0 },
        },
        {
          ...services,
          restartList: () => {
            invoked = true
            return { ok: true, changed: 1 }
          },
        },
      ),
    ).resolves.toMatchObject({
      handled: true,
      operationId: 'docx.list.restart',
      ok: false,
      error: 'invalid_arguments',
    })
    expect(invoked).toBe(false)
  })

  it('declares an explicit DOCX continue-numbering contract', () => {
    const descriptor = docxOperationCatalog.operations.find(
      (operation) => operation.id === 'docx.list.continue',
    )

    expect(descriptor).toMatchObject({
      family: 'list',
      inputSchema: {
        properties: {
          blockIndex: { type: 'integer', minimum: 0 },
          previousBlockIndex: { type: 'integer', minimum: 0 },
        },
        required: ['blockIndex', 'previousBlockIndex'],
        additionalProperties: false,
      },
      outputSchema: {
        required: ['summary', 'matched', 'changed', 'skippedProtected'],
        additionalProperties: false,
      },
      mutates: true,
      undoable: true,
      atomic: true,
    })
  })

  it('continues DOCX numbering from one explicit previous list block', async () => {
    const editor = createEditor('Unchanged')
    let received: unknown

    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.list.continue',
          arguments: { blockIndex: 4, previousBlockIndex: 1 },
        },
        {
          ...services,
          continueList: (input) => {
            received = input
            return { ok: true, changed: 2 }
          },
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.list.continue',
      ok: true,
      output: {
        summary: 'Continued numbering at block 4 from block 1 across 2 item(s)',
        matched: 2,
        changed: 2,
        skippedProtected: 0,
      },
    })
    expect(received).toEqual({ blockIndex: 4, previousBlockIndex: 1 })
  })

  it('rejects a non-previous DOCX list source before invoking numbering state', async () => {
    const editor = createEditor('Unchanged')
    let invoked = false

    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.list.continue',
          arguments: { blockIndex: 1, previousBlockIndex: 2 },
        },
        {
          ...services,
          continueList: () => {
            invoked = true
            return { ok: true, changed: 1 }
          },
        },
      ),
    ).resolves.toMatchObject({
      handled: true,
      operationId: 'docx.list.continue',
      ok: false,
      error: 'invalid_arguments',
    })
    expect(invoked).toBe(false)
  })

  it('sets DOCX text style through one native undo transaction', async () => {
    const editor = createEditor('Important text')

    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.text.set_style',
          arguments: {
            target: { blockIndexes: [0] },
            style: { color: 'FF0000', bold: true },
            fields: ['color', 'bold'],
          },
        },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.text.set_style',
      ok: true,
      output: {
        summary: 'Updated text style in 1 block(s)',
        matched: 1,
        changed: 1,
        skippedProtected: 0,
      },
    })
    const styledText = editor.state.doc.child(0).child(0)
    expect(styledText.marks.some((mark) => mark.type.name === 'bold')).toBe(true)
    expect(styledText.marks.find((mark) => mark.type.name === 'docTextStyle')?.attrs.color).toBe(
      'FF0000',
    )

    expect(editor.commands.undo()).toBe(true)
    const restoredText = editor.state.doc.child(0).child(0)
    expect(restoredText.marks.some((mark) => mark.type.name === 'bold')).toBe(false)
    expect(restoredText.marks.some((mark) => mark.type.name === 'docTextStyle')).toBe(false)
  })

  it('clears one DOCX text-style field without dropping unlisted formatting', async () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docParagraph',
            content: [
              {
                type: 'text',
                text: 'Styled text',
                marks: [
                  {
                    type: 'docTextStyle',
                    attrs: { color: 'FF0000', fontAscii: 'Arial' },
                  },
                ],
              },
            ],
          },
        ],
      },
    })
    editors.push(editor)

    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.text.set_style',
          arguments: {
            target: { blockIndexes: [0] },
            style: { color: null },
            fields: ['color'],
          },
        },
        services,
      ),
    ).resolves.toMatchObject({
      handled: true,
      operationId: 'docx.text.set_style',
      ok: true,
      output: { changed: 1 },
    })
    const cleared = editor.state.doc
      .child(0)
      .child(0)
      .marks.find((mark) => mark.type.name === 'docTextStyle')
    expect(cleared?.attrs).toMatchObject({ color: null, fontAscii: 'Arial' })

    expect(editor.commands.undo()).toBe(true)
    expect(
      editor.state.doc
        .child(0)
        .child(0)
        .marks.find((mark) => mark.type.name === 'docTextStyle')?.attrs,
    ).toMatchObject({ color: 'FF0000', fontAscii: 'Arial' })
  })

  it('sets DOCX paragraph style through one native undo transaction', async () => {
    const editor = createEditor('Centered paragraph')

    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.paragraph.set_style',
          arguments: {
            target: { blockIndexes: [0] },
            style: { align: 'center', spaceAfter: 240 },
            fields: ['align', 'spaceAfter'],
          },
        },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.paragraph.set_style',
      ok: true,
      output: {
        summary: 'Updated paragraph style in 1 block(s)',
        matched: 1,
        changed: 1,
        skippedProtected: 0,
      },
    })
    expect(editor.state.doc.child(0).attrs).toMatchObject({
      align: 'center',
      spaceAfter: 240,
      externalChanged: true,
    })

    expect(editor.commands.undo()).toBe(true)
    expect(editor.state.doc.child(0).attrs.align).toBeNull()
    expect(editor.state.doc.child(0).attrs.spaceAfter).toBeNull()
  })

  it('declares the retained ParagraphDialog fields with bounded geometry', () => {
    const descriptor = docxOperationCatalog.operations.find(
      (operation) => operation.id === 'docx.paragraph.set_style',
    )

    expect(descriptor).toMatchObject({
      inputSchema: {
        properties: {
          style: {
            properties: {
              lineSpacing: { type: ['number', 'null'], minimum: 0.06, maximum: 132 },
              lineRule: {
                type: ['string', 'null'],
                enum: ['auto', 'atLeast', 'exact', null],
              },
              lineRawTwips: {
                type: ['integer', 'null'],
                minimum: 20,
                maximum: 31680,
              },
              indentLeft: {
                type: ['number', 'null'],
                minimum: -31680,
                maximum: 31680,
              },
              indentRight: {
                type: ['number', 'null'],
                minimum: -31680,
                maximum: 31680,
              },
              indentFirstLine: {
                type: ['number', 'null'],
                minimum: -31680,
                maximum: 31680,
              },
              spaceBefore: {
                type: ['number', 'null'],
                minimum: 0,
                maximum: 31680,
              },
              spaceAfter: {
                type: ['number', 'null'],
                minimum: 0,
                maximum: 31680,
              },
            },
          },
          fields: {
            items: {
              enum: expect.arrayContaining(['lineRule', 'lineRawTwips']),
            },
          },
        },
      },
    })
  })

  it('sets exact DOCX line height and paragraph geometry as one undo unit', async () => {
    const editor = createEditor('Exact paragraph')

    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.paragraph.set_style',
          arguments: {
            target: { blockIndexes: [0] },
            style: {
              lineSpacing: null,
              lineRule: 'exact',
              lineRawTwips: 360,
              indentLeft: 720,
              indentFirstLine: -360,
              spaceBefore: 120,
              spaceAfter: 240,
            },
            fields: [
              'lineSpacing',
              'lineRule',
              'lineRawTwips',
              'indentLeft',
              'indentFirstLine',
              'spaceBefore',
              'spaceAfter',
            ],
          },
        },
        services,
      ),
    ).resolves.toMatchObject({
      handled: true,
      operationId: 'docx.paragraph.set_style',
      ok: true,
      output: { changed: 1 },
    })
    expect(editor.state.doc.child(0).attrs).toMatchObject({
      lineSpacing: null,
      lineRule: 'exact',
      lineRawTwips: 360,
      indentLeft: 720,
      indentFirstLine: -360,
      spaceBefore: 120,
      spaceAfter: 240,
      externalChanged: true,
    })

    expect(editor.commands.undo()).toBe(true)
    expect(editor.state.doc.child(0).attrs).toMatchObject({
      lineRule: null,
      lineRawTwips: null,
      indentLeft: null,
      indentFirstLine: null,
    })
  })

  it('rejects out-of-range DOCX exact line height before mutation', async () => {
    const editor = createEditor('Unchanged')

    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.paragraph.set_style',
          arguments: {
            target: { blockIndexes: [0] },
            style: { lineRule: 'exact', lineRawTwips: 31681 },
            fields: ['lineRule', 'lineRawTwips'],
          },
        },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.paragraph.set_style',
      ok: false,
      error: 'invalid_arguments',
      message: '$.style.lineRawTwips must be less than or equal to 31680.',
    })
    expect(editor.state.doc.child(0).attrs.lineRule).toBeNull()
  })

  it('declares a bounded DOCX paragraph tab-stop final state', () => {
    const descriptor = docxOperationCatalog.operations.find(
      (operation) => operation.id === 'docx.paragraph.set_style',
    )

    expect(descriptor).toMatchObject({
      inputSchema: {
        properties: {
          style: {
            properties: {
              tabStops: {
                type: ['array', 'null'],
                minItems: 1,
                maxItems: 64,
                items: {
                  properties: {
                    pos: { type: 'integer', minimum: 0, maximum: 31680 },
                    val: {
                      type: 'string',
                      enum: ['left', 'center', 'right', 'decimal', 'bar', 'clear'],
                    },
                    leader: {
                      type: 'string',
                      enum: ['none', 'dot', 'hyphen', 'underscore', 'heavy', 'middleDot'],
                    },
                  },
                  required: ['pos', 'val'],
                  additionalProperties: false,
                },
              },
            },
          },
          fields: { items: { enum: expect.arrayContaining(['tabStops']) } },
        },
      },
    })
  })

  it('sets a bounded DOCX paragraph tab-stop array and undoes once', async () => {
    const editor = createEditor('Tabbed paragraph')
    const tabStops = [
      { pos: 720, val: 'left', leader: 'dot' },
      { pos: 1440, val: 'decimal' },
    ]

    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.paragraph.set_style',
          arguments: {
            target: { blockIndexes: [0] },
            style: { tabStops },
            fields: ['tabStops'],
          },
        },
        services,
      ),
    ).resolves.toMatchObject({
      handled: true,
      operationId: 'docx.paragraph.set_style',
      ok: true,
      output: { changed: 1 },
    })
    expect(editor.state.doc.child(0).attrs.tabStops).toBe(JSON.stringify(tabStops))

    expect(editor.commands.undo()).toBe(true)
    expect(editor.state.doc.child(0).attrs.tabStops).toBeNull()
  })

  it('rejects unsorted DOCX paragraph tab stops before mutation', async () => {
    const editor = createEditor('Unchanged')

    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.paragraph.set_style',
          arguments: {
            target: { blockIndexes: [0] },
            style: {
              tabStops: [
                { pos: 1440, val: 'right' },
                { pos: 720, val: 'left' },
              ],
            },
            fields: ['tabStops'],
          },
        },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.paragraph.set_style',
      ok: false,
      error: 'invalid_arguments',
      message: 'command #0 updateParagraphStyle: tabStops must be strictly ordered by position',
    })
    expect(editor.state.doc.child(0).attrs.tabStops).toBeNull()
  })

  it('declares a bounded stable-position DOCX table-insert contract', () => {
    const descriptor = docxOperationCatalog.operations.find(
      (operation) => operation.id === 'docx.table.insert',
    )

    expect(descriptor).toMatchObject({
      family: 'table',
      inputSchema: {
        properties: {
          afterBlockIndex: { type: 'integer', minimum: -1 },
          rows: { type: 'integer', minimum: 1, maximum: 100 },
          columns: { type: 'integer', minimum: 1, maximum: 63 },
        },
        required: ['afterBlockIndex', 'rows', 'columns'],
        additionalProperties: false,
      },
      outputSchema: {
        properties: {
          summary: { type: 'string' },
          tableBlockIndex: { type: 'integer', minimum: 0 },
          rows: { type: 'integer' },
          columns: { type: 'integer' },
        },
        required: ['summary', 'tableBlockIndex', 'rows', 'columns'],
        additionalProperties: false,
      },
      mutates: true,
      undoable: true,
      atomic: true,
    })
  })

  it('inserts a native DOCX table after one explicit top-level block and undoes once', async () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          { type: 'docParagraph', content: [{ type: 'text', text: 'Before' }] },
          { type: 'docParagraph', content: [{ type: 'text', text: 'After' }] },
        ],
      },
    })
    editors.push(editor)

    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.table.insert',
          arguments: { afterBlockIndex: 0, rows: 2, columns: 3 },
        },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.table.insert',
      ok: true,
      output: {
        summary: 'Inserted a 2×3 DOCX table after block 0',
        tableBlockIndex: 1,
        rows: 2,
        columns: 3,
      },
    })
    expect(editor.state.doc.childCount).toBe(3)
    const table = editor.state.doc.child(1)
    expect(table.type.name).toBe('docTable')
    expect(table.childCount).toBe(2)
    expect(table.child(0).childCount).toBe(3)

    expect(editor.commands.undo()).toBe(true)
    expect(editor.state.doc.childCount).toBe(2)
    expect(editor.state.doc.child(1).textContent).toBe('After')
  })

  it('rejects a DOCX table exceeding the bounded cell budget before mutation', async () => {
    const editor = createEditor('Unchanged')

    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.table.insert',
          arguments: { afterBlockIndex: 0, rows: 100, columns: 63 },
        },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.table.insert',
      ok: false,
      error: 'invalid_arguments',
      message: 'docx.table.insert supports at most 4096 cells.',
    })
    expect(editor.state.doc.childCount).toBe(1)
  })

  it('declares an explicit stable DOCX table-delete contract', () => {
    const descriptor = docxOperationCatalog.operations.find(
      (operation) => operation.id === 'docx.table.delete',
    )

    expect(descriptor).toMatchObject({
      family: 'table',
      inputSchema: {
        properties: { tableBlockIndex: { type: 'integer', minimum: 0 } },
        required: ['tableBlockIndex'],
        additionalProperties: false,
      },
      outputSchema: {
        properties: {
          summary: { type: 'string' },
          tableBlockIndex: { type: 'integer', minimum: 0 },
          deleted: { type: 'boolean', enum: [true] },
        },
        required: ['summary', 'tableBlockIndex', 'deleted'],
        additionalProperties: false,
      },
      risk: 'high',
      mutates: true,
      undoable: true,
      atomic: true,
    })
  })

  it('deletes one explicit top-level DOCX table and restores it with one undo', async () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          { type: 'docParagraph', content: [{ type: 'text', text: 'Before' }] },
          {
            type: 'docTable',
            content: [
              {
                type: 'docTableRow',
                content: [
                  {
                    type: 'docTableCell',
                    content: [{ type: 'docParagraph', content: [{ type: 'text', text: 'Cell' }] }],
                  },
                ],
              },
            ],
          },
          { type: 'docParagraph', content: [{ type: 'text', text: 'After' }] },
        ],
      },
    })
    editors.push(editor)

    await expect(
      executeDocxOperation(
        editor,
        { operation: 'docx.table.delete', arguments: { tableBlockIndex: 1 } },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.table.delete',
      ok: true,
      output: {
        summary: 'Deleted DOCX table at block 1',
        tableBlockIndex: 1,
        deleted: true,
      },
    })
    expect(editor.state.doc.childCount).toBe(2)
    expect(editor.state.doc.child(1).textContent).toBe('After')

    expect(editor.commands.undo()).toBe(true)
    expect(editor.state.doc.childCount).toBe(3)
    expect(editor.state.doc.child(1).type.name).toBe('docTable')
    expect(editor.state.doc.child(1).textContent).toBe('Cell')
  })

  it('rejects a DOCX table-delete target that is not a table', async () => {
    const editor = createEditor('Unchanged')

    await expect(
      executeDocxOperation(
        editor,
        { operation: 'docx.table.delete', arguments: { tableBlockIndex: 0 } },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.table.delete',
      ok: false,
      error: 'invalid_arguments',
      message: 'DOCX block 0 is not a table.',
    })
    expect(editor.state.doc.childCount).toBe(1)
  })

  it('declares a bounded explicit DOCX row-insert contract', () => {
    const descriptor = docxOperationCatalog.operations.find(
      (operation) => operation.id === 'docx.table.insert_rows',
    )

    expect(descriptor).toMatchObject({
      family: 'table',
      inputSchema: {
        properties: {
          tableBlockIndex: { type: 'integer', minimum: 0 },
          rowIndex: { type: 'integer', minimum: 0 },
          count: { type: 'integer', minimum: 1, maximum: 100 },
        },
        required: ['tableBlockIndex', 'rowIndex', 'count'],
        additionalProperties: false,
      },
      outputSchema: {
        required: ['summary', 'tableBlockIndex', 'rowIndex', 'insertedRows'],
        additionalProperties: false,
      },
      mutates: true,
      undoable: true,
      atomic: true,
    })
  })

  it('inserts multiple native DOCX rows at one explicit boundary and undoes once', async () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docTable',
            content: [
              {
                type: 'docTableRow',
                content: [
                  {
                    type: 'docTableCell',
                    content: [{ type: 'docParagraph', content: [{ type: 'text', text: 'A' }] }],
                  },
                  {
                    type: 'docTableCell',
                    content: [{ type: 'docParagraph', content: [{ type: 'text', text: 'B' }] }],
                  },
                ],
              },
              {
                type: 'docTableRow',
                content: [
                  {
                    type: 'docTableCell',
                    content: [{ type: 'docParagraph', content: [{ type: 'text', text: 'C' }] }],
                  },
                  {
                    type: 'docTableCell',
                    content: [{ type: 'docParagraph', content: [{ type: 'text', text: 'D' }] }],
                  },
                ],
              },
            ],
          },
        ],
      },
    })
    editors.push(editor)

    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.table.insert_rows',
          arguments: { tableBlockIndex: 0, rowIndex: 1, count: 2 },
        },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.table.insert_rows',
      ok: true,
      output: {
        summary: 'Inserted 2 DOCX row(s) at boundary 1 in table block 0',
        tableBlockIndex: 0,
        rowIndex: 1,
        insertedRows: 2,
      },
    })
    expect(editor.state.doc.child(0).childCount).toBe(4)
    expect(editor.state.doc.child(0).child(1).childCount).toBe(2)
    expect(editor.state.doc.child(0).child(2).childCount).toBe(2)

    expect(editor.commands.undo()).toBe(true)
    expect(editor.state.doc.child(0).childCount).toBe(2)
    expect(editor.state.doc.child(0).textContent).toBe('ABCD')
  })

  it('rejects a DOCX row-insert boundary beyond the target table', async () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docTable',
            content: [
              {
                type: 'docTableRow',
                content: [
                  {
                    type: 'docTableCell',
                    content: [{ type: 'docParagraph' }],
                  },
                ],
              },
            ],
          },
        ],
      },
    })
    editors.push(editor)

    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.table.insert_rows',
          arguments: { tableBlockIndex: 0, rowIndex: 2, count: 1 },
        },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.table.insert_rows',
      ok: false,
      error: 'invalid_arguments',
      message: 'DOCX table block 0 has 1 row(s); boundary 2 is invalid.',
    })
    expect(editor.state.doc.childCount).toBe(1)
    expect(editor.state.doc.child(0).childCount).toBe(1)
  })

  it('declares a bounded explicit DOCX row-delete contract', () => {
    const descriptor = docxOperationCatalog.operations.find(
      (operation) => operation.id === 'docx.table.delete_rows',
    )

    expect(descriptor).toMatchObject({
      family: 'table',
      inputSchema: {
        properties: {
          tableBlockIndex: { type: 'integer', minimum: 0 },
          rowIndex: { type: 'integer', minimum: 0 },
          count: { type: 'integer', minimum: 1, maximum: 100 },
        },
        required: ['tableBlockIndex', 'rowIndex', 'count'],
        additionalProperties: false,
      },
      outputSchema: {
        required: ['summary', 'tableBlockIndex', 'rowIndex', 'deletedRows'],
        additionalProperties: false,
      },
      risk: 'high',
      mutates: true,
      undoable: true,
      atomic: true,
    })
  })

  it('deletes multiple native DOCX rows from one explicit index and undoes once', async () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docTable',
            content: ['A', 'B', 'C', 'D'].map((text) => ({
              type: 'docTableRow',
              content: [
                {
                  type: 'docTableCell',
                  content: [{ type: 'docParagraph', content: [{ type: 'text', text }] }],
                },
              ],
            })),
          },
        ],
      },
    })
    editors.push(editor)

    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.table.delete_rows',
          arguments: { tableBlockIndex: 0, rowIndex: 1, count: 2 },
        },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.table.delete_rows',
      ok: true,
      output: {
        summary: 'Deleted 2 DOCX row(s) from index 1 in table block 0',
        tableBlockIndex: 0,
        rowIndex: 1,
        deletedRows: 2,
      },
    })
    expect(editor.state.doc.child(0).childCount).toBe(2)
    expect(editor.state.doc.child(0).textContent).toBe('AD')

    expect(editor.commands.undo()).toBe(true)
    expect(editor.state.doc.child(0).childCount).toBe(4)
    expect(editor.state.doc.child(0).textContent).toBe('ABCD')
  })

  it.each([
    [{ tableBlockIndex: 0, rowIndex: 4, count: 1 }, 'row range [4, 5) is invalid'],
    [
      { tableBlockIndex: 0, rowIndex: 0, count: 4 },
      'cannot delete every row; use docx.table.delete instead',
    ],
  ])('rejects an invalid DOCX row-delete request without mutation', async (arguments_, reason) => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docTable',
            content: ['A', 'B', 'C', 'D'].map((text) => ({
              type: 'docTableRow',
              content: [
                {
                  type: 'docTableCell',
                  content: [{ type: 'docParagraph', content: [{ type: 'text', text }] }],
                },
              ],
            })),
          },
        ],
      },
    })
    editors.push(editor)

    await expect(
      executeDocxOperation(
        editor,
        { operation: 'docx.table.delete_rows', arguments: arguments_ },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.table.delete_rows',
      ok: false,
      error: 'invalid_arguments',
      message: `DOCX table block 0 ${reason}.`,
    })
    expect(editor.state.doc.child(0).childCount).toBe(4)
    expect(editor.state.doc.child(0).textContent).toBe('ABCD')
  })

  it('declares a bounded explicit DOCX column-insert contract', () => {
    const descriptor = docxOperationCatalog.operations.find(
      (operation) => operation.id === 'docx.table.insert_columns',
    )

    expect(descriptor).toMatchObject({
      family: 'table',
      inputSchema: {
        properties: {
          tableBlockIndex: { type: 'integer', minimum: 0 },
          columnIndex: { type: 'integer', minimum: 0 },
          count: { type: 'integer', minimum: 1, maximum: 63 },
        },
        required: ['tableBlockIndex', 'columnIndex', 'count'],
        additionalProperties: false,
      },
      outputSchema: {
        required: ['summary', 'tableBlockIndex', 'columnIndex', 'insertedColumns'],
        additionalProperties: false,
      },
      mutates: true,
      undoable: true,
      atomic: true,
    })
  })

  it('inserts multiple native DOCX columns at one explicit boundary and undoes once', async () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docTable',
            content: ['AB', 'CD'].map((row) => ({
              type: 'docTableRow',
              content: [...row].map((text) => ({
                type: 'docTableCell',
                content: [{ type: 'docParagraph', content: [{ type: 'text', text }] }],
              })),
            })),
          },
        ],
      },
    })
    editors.push(editor)

    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.table.insert_columns',
          arguments: { tableBlockIndex: 0, columnIndex: 1, count: 2 },
        },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.table.insert_columns',
      ok: true,
      output: {
        summary: 'Inserted 2 DOCX column(s) at boundary 1 in table block 0',
        tableBlockIndex: 0,
        columnIndex: 1,
        insertedColumns: 2,
      },
    })
    expect(editor.state.doc.child(0).child(0).childCount).toBe(4)
    expect(editor.state.doc.child(0).child(1).childCount).toBe(4)

    expect(editor.commands.undo()).toBe(true)
    expect(editor.state.doc.child(0).child(0).childCount).toBe(2)
    expect(editor.state.doc.child(0).child(1).childCount).toBe(2)
    expect(editor.state.doc.child(0).textContent).toBe('ABCD')
  })

  it('rejects a DOCX column-insert boundary beyond the target table', async () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docTable',
            content: [
              {
                type: 'docTableRow',
                content: [
                  { type: 'docTableCell', content: [{ type: 'docParagraph' }] },
                  { type: 'docTableCell', content: [{ type: 'docParagraph' }] },
                ],
              },
            ],
          },
        ],
      },
    })
    editors.push(editor)

    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.table.insert_columns',
          arguments: { tableBlockIndex: 0, columnIndex: 3, count: 1 },
        },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.table.insert_columns',
      ok: false,
      error: 'invalid_arguments',
      message: 'DOCX table block 0 has 2 column(s); boundary 3 is invalid.',
    })
    expect(editor.state.doc.child(0).child(0).childCount).toBe(2)
  })

  it('declares a bounded explicit DOCX column-delete contract', () => {
    const descriptor = docxOperationCatalog.operations.find(
      (operation) => operation.id === 'docx.table.delete_columns',
    )

    expect(descriptor).toMatchObject({
      family: 'table',
      inputSchema: {
        properties: {
          tableBlockIndex: { type: 'integer', minimum: 0 },
          columnIndex: { type: 'integer', minimum: 0 },
          count: { type: 'integer', minimum: 1, maximum: 63 },
        },
        required: ['tableBlockIndex', 'columnIndex', 'count'],
        additionalProperties: false,
      },
      outputSchema: {
        required: ['summary', 'tableBlockIndex', 'columnIndex', 'deletedColumns'],
        additionalProperties: false,
      },
      risk: 'high',
      mutates: true,
      undoable: true,
      atomic: true,
    })
  })

  it('deletes multiple native DOCX columns from one explicit index and undoes once', async () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docTable',
            content: ['ABCD', 'EFGH'].map((row) => ({
              type: 'docTableRow',
              content: [...row].map((text) => ({
                type: 'docTableCell',
                content: [{ type: 'docParagraph', content: [{ type: 'text', text }] }],
              })),
            })),
          },
        ],
      },
    })
    editors.push(editor)

    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.table.delete_columns',
          arguments: { tableBlockIndex: 0, columnIndex: 1, count: 2 },
        },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.table.delete_columns',
      ok: true,
      output: {
        summary: 'Deleted 2 DOCX column(s) from index 1 in table block 0',
        tableBlockIndex: 0,
        columnIndex: 1,
        deletedColumns: 2,
      },
    })
    expect(editor.state.doc.child(0).child(0).childCount).toBe(2)
    expect(editor.state.doc.child(0).child(1).childCount).toBe(2)
    expect(editor.state.doc.child(0).textContent).toBe('ADEH')

    expect(editor.commands.undo()).toBe(true)
    expect(editor.state.doc.child(0).child(0).childCount).toBe(4)
    expect(editor.state.doc.child(0).child(1).childCount).toBe(4)
    expect(editor.state.doc.child(0).textContent).toBe('ABCDEFGH')
  })

  it.each([
    [{ tableBlockIndex: 0, columnIndex: 4, count: 1 }, 'column range [4, 5) is invalid'],
    [
      { tableBlockIndex: 0, columnIndex: 0, count: 4 },
      'cannot delete every column; use docx.table.delete instead',
    ],
  ])(
    'rejects an invalid DOCX column-delete request without mutation',
    async (arguments_, reason) => {
      const editor = new Editor({
        element: document.createElement('div'),
        extensions: editorExtensions,
        content: {
          type: 'doc',
          content: [
            {
              type: 'docTable',
              content: [
                {
                  type: 'docTableRow',
                  content: ['A', 'B', 'C', 'D'].map((text) => ({
                    type: 'docTableCell',
                    content: [{ type: 'docParagraph', content: [{ type: 'text', text }] }],
                  })),
                },
              ],
            },
          ],
        },
      })
      editors.push(editor)

      await expect(
        executeDocxOperation(
          editor,
          { operation: 'docx.table.delete_columns', arguments: arguments_ },
          services,
        ),
      ).resolves.toEqual({
        handled: true,
        operationId: 'docx.table.delete_columns',
        ok: false,
        error: 'invalid_arguments',
        message: `DOCX table block 0 ${reason}.`,
      })
      expect(editor.state.doc.child(0).child(0).childCount).toBe(4)
      expect(editor.state.doc.child(0).textContent).toBe('ABCD')
    },
  )

  it('declares an explicit bounded DOCX cell-merge contract', () => {
    const descriptor = docxOperationCatalog.operations.find(
      (operation) => operation.id === 'docx.table.merge_cells',
    )

    expect(descriptor).toMatchObject({
      family: 'table',
      inputSchema: {
        properties: {
          tableBlockIndex: { type: 'integer', minimum: 0 },
          topRow: { type: 'integer', minimum: 0, maximum: 99 },
          leftColumn: { type: 'integer', minimum: 0, maximum: 62 },
          bottomRow: { type: 'integer', minimum: 1, maximum: 100 },
          rightColumn: { type: 'integer', minimum: 1, maximum: 63 },
        },
        required: ['tableBlockIndex', 'topRow', 'leftColumn', 'bottomRow', 'rightColumn'],
        additionalProperties: false,
      },
      outputSchema: {
        required: ['summary', 'tableBlockIndex', 'mergedCells'],
        additionalProperties: false,
      },
      mutates: true,
      undoable: true,
      atomic: true,
    })
  })

  it('merges one explicit DOCX cell rectangle through native table history and undoes once', async () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docTable',
            content: ['AB', 'CD'].map((row) => ({
              type: 'docTableRow',
              content: [...row].map((text) => ({
                type: 'docTableCell',
                content: [{ type: 'docParagraph', content: [{ type: 'text', text }] }],
              })),
            })),
          },
        ],
      },
    })
    editors.push(editor)

    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.table.merge_cells',
          arguments: {
            tableBlockIndex: 0,
            topRow: 0,
            leftColumn: 0,
            bottomRow: 2,
            rightColumn: 2,
          },
        },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.table.merge_cells',
      ok: true,
      output: {
        summary: 'Merged 4 DOCX logical cell(s) in table block 0',
        tableBlockIndex: 0,
        mergedCells: 4,
      },
    })
    const mergedCell = editor.state.doc.child(0).child(0).child(0)
    expect(mergedCell.attrs.colspan).toBe(2)
    expect(mergedCell.attrs.rowspan).toBe(2)
    expect(editor.state.doc.child(0).textContent).toBe('ABCD')

    expect(editor.commands.undo()).toBe(true)
    expect(editor.state.doc.child(0).child(0).childCount).toBe(2)
    expect(editor.state.doc.child(0).child(1).childCount).toBe(2)
  })

  it.each([
    [
      { tableBlockIndex: 0, topRow: 0, leftColumn: 0, bottomRow: 1, rightColumn: 1 },
      'must cover at least two logical cells',
    ],
    [
      { tableBlockIndex: 0, topRow: 0, leftColumn: 0, bottomRow: 3, rightColumn: 2 },
      'rectangle [0, 3) × [0, 2) is outside its 2×2 logical grid',
    ],
  ])(
    'rejects an invalid DOCX cell-merge rectangle without mutation',
    async (arguments_, reason) => {
      const editor = new Editor({
        element: document.createElement('div'),
        extensions: editorExtensions,
        content: {
          type: 'doc',
          content: [
            {
              type: 'docTable',
              content: ['AB', 'CD'].map((row) => ({
                type: 'docTableRow',
                content: [...row].map((text) => ({
                  type: 'docTableCell',
                  content: [{ type: 'docParagraph', content: [{ type: 'text', text }] }],
                })),
              })),
            },
          ],
        },
      })
      editors.push(editor)

      await expect(
        executeDocxOperation(
          editor,
          { operation: 'docx.table.merge_cells', arguments: arguments_ },
          services,
        ),
      ).resolves.toEqual({
        handled: true,
        operationId: 'docx.table.merge_cells',
        ok: false,
        error: 'invalid_arguments',
        message: `DOCX table block 0 ${reason}.`,
      })
      expect(editor.state.doc.child(0).child(0).childCount).toBe(2)
      expect(editor.state.doc.child(0).child(1).childCount).toBe(2)
    },
  )

  it('declares an explicit bounded DOCX cell-split contract', () => {
    const descriptor = docxOperationCatalog.operations.find(
      (operation) => operation.id === 'docx.table.split_cell',
    )

    expect(descriptor).toMatchObject({
      family: 'table',
      inputSchema: {
        properties: {
          tableBlockIndex: { type: 'integer', minimum: 0 },
          rowIndex: { type: 'integer', minimum: 0, maximum: 99 },
          columnIndex: { type: 'integer', minimum: 0, maximum: 62 },
        },
        required: ['tableBlockIndex', 'rowIndex', 'columnIndex'],
        additionalProperties: false,
      },
      outputSchema: {
        required: ['summary', 'tableBlockIndex', 'splitCells'],
        additionalProperties: false,
      },
      mutates: true,
      undoable: true,
      atomic: true,
    })
  })

  it('splits one explicit DOCX merged cell through native table history and undoes once', async () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docTable',
            content: [
              {
                type: 'docTableRow',
                content: [
                  {
                    type: 'docTableCell',
                    attrs: { colspan: 2, rowspan: 2 },
                    content: [
                      { type: 'docParagraph', content: [{ type: 'text', text: 'Merged' }] },
                    ],
                  },
                ],
              },
              { type: 'docTableRow' },
            ],
          },
        ],
      },
    })
    editors.push(editor)

    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.table.split_cell',
          arguments: { tableBlockIndex: 0, rowIndex: 1, columnIndex: 1 },
        },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.table.split_cell',
      ok: true,
      output: {
        summary: 'Split one DOCX merged cell into 4 logical cells in table block 0',
        tableBlockIndex: 0,
        splitCells: 4,
      },
    })
    expect(editor.state.doc.child(0).child(0).childCount).toBe(2)
    expect(editor.state.doc.child(0).child(1).childCount).toBe(2)
    expect(editor.state.doc.child(0).child(0).child(0).attrs.colspan).toBe(1)
    expect(editor.state.doc.child(0).child(0).child(0).attrs.rowspan).toBe(1)

    expect(editor.commands.undo()).toBe(true)
    expect(editor.state.doc.child(0).child(0).childCount).toBe(1)
    expect(editor.state.doc.child(0).child(0).child(0).attrs.colspan).toBe(2)
    expect(editor.state.doc.child(0).child(0).child(0).attrs.rowspan).toBe(2)
    expect(editor.state.doc.child(0).child(1).childCount).toBe(0)
  })

  it('rejects splitting an ordinary DOCX cell without mutation', async () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docTable',
            content: [
              {
                type: 'docTableRow',
                content: [
                  {
                    type: 'docTableCell',
                    content: [{ type: 'docParagraph', content: [{ type: 'text', text: 'Plain' }] }],
                  },
                ],
              },
            ],
          },
        ],
      },
    })
    editors.push(editor)

    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.table.split_cell',
          arguments: { tableBlockIndex: 0, rowIndex: 0, columnIndex: 0 },
        },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.table.split_cell',
      ok: false,
      error: 'invalid_arguments',
      message: 'DOCX table block 0 logical cell (0, 0) is not merged.',
    })
    expect(editor.state.doc.child(0).child(0).childCount).toBe(1)
    expect(editor.state.doc.child(0).textContent).toBe('Plain')
  })

  it('declares an exact bounded DOCX cell-format contract', () => {
    const descriptor = docxOperationCatalog.operations.find(
      (operation) => operation.id === 'docx.table.set_cell_format',
    )

    expect(descriptor).toMatchObject({
      family: 'table',
      inputSchema: {
        properties: {
          tableBlockIndex: { type: 'integer', minimum: 0 },
          topRow: { type: 'integer', minimum: 0, maximum: 99 },
          leftColumn: { type: 'integer', minimum: 0, maximum: 62 },
          bottomRow: { type: 'integer', minimum: 1, maximum: 100 },
          rightColumn: { type: 'integer', minimum: 1, maximum: 63 },
          format: {
            properties: {
              fill: { type: ['string', 'null'] },
              verticalAlignment: {
                type: 'string',
                enum: ['top', 'center', 'bottom'],
              },
            },
            required: [],
            additionalProperties: false,
          },
          fields: {
            type: 'array',
            minItems: 1,
            maxItems: 2,
            items: { type: 'string', enum: ['fill', 'verticalAlignment'] },
          },
        },
        required: [
          'tableBlockIndex',
          'topRow',
          'leftColumn',
          'bottomRow',
          'rightColumn',
          'format',
          'fields',
        ],
        additionalProperties: false,
      },
      outputSchema: {
        required: ['summary', 'tableBlockIndex', 'matchedCells', 'changedCells'],
        additionalProperties: false,
      },
      mutates: true,
      undoable: true,
      atomic: true,
    })
  })

  it('sets exact DOCX cell fill and vertical alignment in one native undo transaction', async () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docTable',
            content: ['AB', 'CD'].map((row) => ({
              type: 'docTableRow',
              content: [...row].map((text) => ({
                type: 'docTableCell',
                content: [{ type: 'docParagraph', content: [{ type: 'text', text }] }],
              })),
            })),
          },
        ],
      },
    })
    editors.push(editor)

    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.table.set_cell_format',
          arguments: {
            tableBlockIndex: 0,
            topRow: 0,
            leftColumn: 0,
            bottomRow: 2,
            rightColumn: 2,
            format: { fill: 'D9EAF7', verticalAlignment: 'center' },
            fields: ['fill', 'verticalAlignment'],
          },
        },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.table.set_cell_format',
      ok: true,
      output: {
        summary: 'Formatted 4 DOCX cell(s) in table block 0',
        tableBlockIndex: 0,
        matchedCells: 4,
        changedCells: 4,
      },
    })
    for (let row = 0; row < 2; row++) {
      for (let column = 0; column < 2; column++) {
        const cell = editor.state.doc.child(0).child(row).child(column)
        expect(cell.attrs.fill).toBe('D9EAF7')
        expect(cell.attrs.vAlign).toBe('center')
      }
    }

    expect(editor.commands.undo()).toBe(true)
    expect(editor.state.doc.child(0).child(0).child(0).attrs.fill).toBeNull()
    expect(editor.state.doc.child(0).child(0).child(0).attrs.vAlign).toBeNull()
  })

  it('rejects invalid DOCX cell format before mutation', async () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docTable',
            content: [
              {
                type: 'docTableRow',
                content: [{ type: 'docTableCell', content: [{ type: 'docParagraph' }] }],
              },
            ],
          },
        ],
      },
    })
    editors.push(editor)

    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.table.set_cell_format',
          arguments: {
            tableBlockIndex: 0,
            topRow: 0,
            leftColumn: 0,
            bottomRow: 1,
            rightColumn: 1,
            format: { fill: 'not-a-color' },
            fields: ['fill'],
          },
        },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.table.set_cell_format',
      ok: false,
      error: 'invalid_arguments',
      message: 'docx.table.set_cell_format fill must be null or six uppercase hexadecimal digits.',
    })
    expect(editor.state.doc.child(0).child(0).child(0).attrs.fill).toBeNull()
  })

  it('declares an explicit bounded DOCX cell-border contract', () => {
    const descriptor = docxOperationCatalog.operations.find(
      (operation) => operation.id === 'docx.table.set_cell_borders',
    )

    expect(descriptor).toMatchObject({
      family: 'table',
      inputSchema: {
        properties: {
          tableBlockIndex: { type: 'integer', minimum: 0 },
          topRow: { type: 'integer', minimum: 0, maximum: 99 },
          leftColumn: { type: 'integer', minimum: 0, maximum: 62 },
          bottomRow: { type: 'integer', minimum: 1, maximum: 100 },
          rightColumn: { type: 'integer', minimum: 1, maximum: 63 },
          mode: { type: 'string', enum: ['all', 'outer', 'inner', 'none'] },
          border: {
            type: ['object', 'null'],
            properties: {
              color: { type: 'string' },
              sizeEighths: { type: 'integer', minimum: 1, maximum: 96 },
            },
            required: ['color', 'sizeEighths'],
            additionalProperties: false,
          },
        },
        required: [
          'tableBlockIndex',
          'topRow',
          'leftColumn',
          'bottomRow',
          'rightColumn',
          'mode',
          'border',
        ],
        additionalProperties: false,
      },
      outputSchema: {
        required: ['summary', 'tableBlockIndex', 'matchedCells', 'changedCells'],
        additionalProperties: false,
      },
      mutates: true,
      undoable: true,
      atomic: true,
    })
  })

  it('sets all borders over one exact DOCX cell rectangle and undoes once', async () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docTable',
            content: ['AB', 'CD'].map((row) => ({
              type: 'docTableRow',
              content: [...row].map((text) => ({
                type: 'docTableCell',
                content: [{ type: 'docParagraph', content: [{ type: 'text', text }] }],
              })),
            })),
          },
        ],
      },
    })
    editors.push(editor)

    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.table.set_cell_borders',
          arguments: {
            tableBlockIndex: 0,
            topRow: 0,
            leftColumn: 0,
            bottomRow: 2,
            rightColumn: 2,
            mode: 'all',
            border: { color: '112233', sizeEighths: 8 },
          },
        },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.table.set_cell_borders',
      ok: true,
      output: {
        summary: 'Set all borders on 4 DOCX cell(s) in table block 0',
        tableBlockIndex: 0,
        matchedCells: 4,
        changedCells: 4,
      },
    })
    const borders = editor.state.doc.child(0).child(0).child(0).attrs.borders
    expect(borders).toEqual({
      top: { style: 'single', color: '112233', szEighths: 8 },
      bottom: { style: 'single', color: '112233', szEighths: 8 },
      left: { style: 'single', color: '112233', szEighths: 8 },
      right: { style: 'single', color: '112233', szEighths: 8 },
    })
    expect(editor.commands.undo()).toBe(true)
    expect(editor.state.doc.child(0).child(0).child(0).attrs.borders).toBeNull()
  })

  it('rejects an incoherent DOCX cell-border final state before mutation', async () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docTable',
            content: [
              {
                type: 'docTableRow',
                content: [{ type: 'docTableCell', content: [{ type: 'docParagraph' }] }],
              },
            ],
          },
        ],
      },
    })
    editors.push(editor)

    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.table.set_cell_borders',
          arguments: {
            tableBlockIndex: 0,
            topRow: 0,
            leftColumn: 0,
            bottomRow: 1,
            rightColumn: 1,
            mode: 'none',
            border: { color: '000000', sizeEighths: 4 },
          },
        },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.table.set_cell_borders',
      ok: false,
      error: 'invalid_arguments',
      message: 'docx.table.set_cell_borders mode none requires border: null.',
    })
    expect(editor.state.doc.child(0).child(0).child(0).attrs.borders).toBeNull()
  })

  it('declares an exact current-document DOCX table-style contract', () => {
    const descriptor = docxOperationCatalog.operations.find(
      (operation) => operation.id === 'docx.table.set_style',
    )

    expect(descriptor).toMatchObject({
      family: 'table',
      inputSchema: {
        properties: {
          tableBlockIndex: { type: 'integer', minimum: 0 },
          styleId: { type: ['string', 'null'] },
        },
        required: ['tableBlockIndex', 'styleId'],
        additionalProperties: false,
      },
      outputSchema: {
        required: ['summary', 'tableBlockIndex', 'styleId', 'changed'],
        additionalProperties: false,
      },
      mutates: true,
      undoable: true,
      atomic: true,
    })
  })

  it('sets one current-document DOCX table style and undoes once', async () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docTable',
            content: [
              {
                type: 'docTableRow',
                content: [{ type: 'docTableCell', content: [{ type: 'docParagraph' }] }],
              },
            ],
          },
        ],
      },
    })
    editors.push(editor)

    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.table.set_style',
          arguments: { tableBlockIndex: 0, styleId: 'TableGrid' },
        },
        { ...services, hasTableStyle: (styleId: string) => styleId === 'TableGrid' },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.table.set_style',
      ok: true,
      output: {
        summary: 'Set DOCX table block 0 style to TableGrid',
        tableBlockIndex: 0,
        styleId: 'TableGrid',
        changed: true,
      },
    })
    expect(editor.state.doc.child(0).attrs.tblStyleId).toBe('TableGrid')
    expect(editor.commands.undo()).toBe(true)
    expect(editor.state.doc.child(0).attrs.tblStyleId).toBeNull()
  })

  it('rejects a DOCX table style absent from the current document', async () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docTable',
            content: [
              {
                type: 'docTableRow',
                content: [{ type: 'docTableCell', content: [{ type: 'docParagraph' }] }],
              },
            ],
          },
        ],
      },
    })
    editors.push(editor)

    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.table.set_style',
          arguments: { tableBlockIndex: 0, styleId: 'MissingStyle' },
        },
        { ...services, hasTableStyle: () => false },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.table.set_style',
      ok: false,
      error: 'invalid_arguments',
      message: 'docx.table.set_style requires a table style from the current document.',
    })
    expect(editor.state.doc.child(0).attrs.tblStyleId).toBeNull()
  })

  it('declares an explicit bounded DOCX row-height contract', () => {
    const descriptor = docxOperationCatalog.operations.find(
      (operation) => operation.id === 'docx.table.set_row_height',
    )

    expect(descriptor).toMatchObject({
      family: 'table',
      inputSchema: {
        properties: {
          tableBlockIndex: { type: 'integer', minimum: 0 },
          rowIndex: { type: 'integer', minimum: 0, maximum: 99 },
          count: { type: 'integer', minimum: 1, maximum: 100 },
          heightTwips: { type: ['integer', 'null'], minimum: 1, maximum: 31680 },
        },
        required: ['tableBlockIndex', 'rowIndex', 'count', 'heightTwips'],
        additionalProperties: false,
      },
      outputSchema: {
        required: ['summary', 'tableBlockIndex', 'matchedRows', 'changedRows'],
        additionalProperties: false,
      },
      mutates: true,
      undoable: true,
      atomic: true,
    })
  })

  it('sets an exact DOCX row-height interval and undoes once', async () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docTable',
            content: ['A', 'B', 'C'].map((text) => ({
              type: 'docTableRow',
              content: [
                {
                  type: 'docTableCell',
                  content: [{ type: 'docParagraph', content: [{ type: 'text', text }] }],
                },
              ],
            })),
          },
        ],
      },
    })
    editors.push(editor)

    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.table.set_row_height',
          arguments: { tableBlockIndex: 0, rowIndex: 1, count: 2, heightTwips: 1440 },
        },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.table.set_row_height',
      ok: true,
      output: {
        summary: 'Set 2 DOCX row height(s) in table block 0',
        tableBlockIndex: 0,
        matchedRows: 2,
        changedRows: 2,
      },
    })
    expect(editor.state.doc.child(0).child(0).attrs.heightTwips).toBeNull()
    expect(editor.state.doc.child(0).child(1).attrs.heightTwips).toBe(1440)
    expect(editor.state.doc.child(0).child(2).attrs.heightTwips).toBe(1440)

    expect(editor.commands.undo()).toBe(true)
    expect(editor.state.doc.child(0).child(1).attrs.heightTwips).toBeNull()
    expect(editor.state.doc.child(0).child(2).attrs.heightTwips).toBeNull()
  })

  it('rejects a DOCX row-height interval beyond the target table', async () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docTable',
            content: [
              {
                type: 'docTableRow',
                content: [{ type: 'docTableCell', content: [{ type: 'docParagraph' }] }],
              },
            ],
          },
        ],
      },
    })
    editors.push(editor)

    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.table.set_row_height',
          arguments: { tableBlockIndex: 0, rowIndex: 1, count: 1, heightTwips: null },
        },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.table.set_row_height',
      ok: false,
      error: 'invalid_arguments',
      message: 'DOCX table block 0 row range [1, 2) is invalid.',
    })
    expect(editor.state.doc.child(0).child(0).attrs.heightTwips).toBeNull()
  })

  it('declares an explicit bounded DOCX column-width vector contract', () => {
    const descriptor = docxOperationCatalog.operations.find(
      (operation) => operation.id === 'docx.table.set_column_widths',
    )

    expect(descriptor).toMatchObject({
      family: 'table',
      inputSchema: {
        properties: {
          tableBlockIndex: { type: 'integer', minimum: 0 },
          widthsPx: {
            type: 'array',
            minItems: 1,
            maxItems: 63,
            items: { type: 'number', minimum: 40, maximum: 4096 },
          },
        },
        required: ['tableBlockIndex', 'widthsPx'],
        additionalProperties: false,
      },
      outputSchema: {
        required: ['summary', 'tableBlockIndex', 'matchedColumns', 'changedColumns'],
        additionalProperties: false,
      },
      mutates: true,
      undoable: true,
      atomic: true,
    })
  })

  it('sets one exact DOCX column-width vector and undoes once', async () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docTable',
            content: ['AB', 'CD'].map((row) => ({
              type: 'docTableRow',
              content: [...row].map((text) => ({
                type: 'docTableCell',
                content: [{ type: 'docParagraph', content: [{ type: 'text', text }] }],
              })),
            })),
          },
        ],
      },
    })
    editors.push(editor)

    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.table.set_column_widths',
          arguments: { tableBlockIndex: 0, widthsPx: [120, 180] },
        },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.table.set_column_widths',
      ok: true,
      output: {
        summary: 'Set 2 DOCX column width(s) in table block 0',
        tableBlockIndex: 0,
        matchedColumns: 2,
        changedColumns: 2,
      },
    })
    const table = editor.state.doc.child(0)
    expect(table.attrs.widthPx).toBe(300)
    expect(table.attrs.colWidthsPct).toEqual([40, 60])
    expect(table.child(0).child(0).attrs.colwidth).toEqual([120])
    expect(table.child(0).child(1).attrs.colwidth).toEqual([180])

    expect(editor.commands.undo()).toBe(true)
    expect(editor.state.doc.child(0).attrs.widthPx).toBeNull()
    expect(editor.state.doc.child(0).child(0).child(0).attrs.colwidth).toBeNull()
  })

  it('rejects a DOCX column-width vector that does not match the current grid', async () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docTable',
            content: [
              {
                type: 'docTableRow',
                content: [
                  { type: 'docTableCell', content: [{ type: 'docParagraph' }] },
                  { type: 'docTableCell', content: [{ type: 'docParagraph' }] },
                ],
              },
            ],
          },
        ],
      },
    })
    editors.push(editor)

    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.table.set_column_widths',
          arguments: { tableBlockIndex: 0, widthsPx: [200] },
        },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.table.set_column_widths',
      ok: false,
      error: 'invalid_arguments',
      message: 'DOCX table block 0 requires exactly 2 column width(s).',
    })
    expect(editor.state.doc.child(0).attrs.widthPx).toBeNull()
  })

  it('moves DOCX blocks in document order through one native undo transaction', async () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: ['A', 'B', 'C', 'D'].map((text) => ({
          type: 'docParagraph',
          content: [{ type: 'text', text }],
        })),
      },
    })
    editors.push(editor)

    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.block.move',
          arguments: { blockIndexes: [1, 2], afterBlockIndex: 3 },
        },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.block.move',
      ok: true,
      output: {
        summary: 'Moved 2 block(s)',
        matched: 2,
        changed: 2,
        skippedProtected: 0,
      },
    })
    expect(
      Array.from(
        { length: editor.state.doc.childCount },
        (_, i) => editor.state.doc.child(i).textContent,
      ),
    ).toEqual(['A', 'D', 'B', 'C'])

    expect(editor.commands.undo()).toBe(true)
    expect(
      Array.from(
        { length: editor.state.doc.childCount },
        (_, i) => editor.state.doc.child(i).textContent,
      ),
    ).toEqual(['A', 'B', 'C', 'D'])
  })

  it('updates DOCX image dimensions proportionally through one native undo transaction', async () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docProtected',
            attrs: {
              docxIndex: 7,
              blockType: 'image',
              label: 'Image',
              imageWidthPx: 400,
              imageHeightPx: 200,
            },
          },
        ],
      },
    })
    editors.push(editor)

    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.image.update',
          arguments: {
            target: { nodeType: 'image' },
            properties: { widthPx: 200 },
            fields: ['widthPx'],
          },
        },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.image.update',
      ok: true,
      output: {
        summary: 'Updated 1 image(s)',
        matched: 1,
        changed: 1,
        skippedProtected: 0,
      },
    })
    expect(editor.state.doc.child(0).attrs).toMatchObject({
      imageWidthPx: 200,
      imageHeightPx: 100,
    })

    expect(editor.commands.undo()).toBe(true)
    expect(editor.state.doc.child(0).attrs).toMatchObject({
      imageWidthPx: 400,
      imageHeightPx: 200,
    })
  })

  it('inserts a DOCX table of contents from current headings through one native undo transaction', async () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docHeading',
            attrs: { level: 1 },
            content: [{ type: 'text', text: 'Introduction' }],
          },
          {
            type: 'docHeading',
            attrs: { level: 2 },
            content: [{ type: 'text', text: 'Risks' }],
          },
          { type: 'docParagraph', content: [{ type: 'text', text: 'Body' }] },
        ],
      },
    })
    editors.push(editor)

    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.toc.insert',
          arguments: { afterBlockIndex: -1 },
        },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.toc.insert',
      ok: true,
      output: {
        summary: 'Inserted a table of contents with 2 entries',
        matched: 2,
        changed: 2,
        skippedProtected: 0,
      },
    })
    expect(editor.state.doc.childCount).toBe(5)
    expect(editor.state.doc.child(0).attrs).toMatchObject({
      blockType: 'passthrough',
      label: 'TOC field',
      fieldDisplay: { kind: 'tocLine', left: 'Introduction', level: 1 },
    })
    expect(editor.state.doc.child(1).attrs.fieldDisplay).toMatchObject({
      kind: 'tocLine',
      left: 'Risks',
      level: 2,
    })

    expect(editor.commands.undo()).toBe(true)
    expect(editor.state.doc.childCount).toBe(3)
    expect(editor.state.doc.child(0).type.name).toBe('docHeading')
  })

  it('applies DOCX list formatting through one native undo transaction', async () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          { type: 'docParagraph', content: [{ type: 'text', text: 'Introduction' }] },
          { type: 'docParagraph', content: [{ type: 'text', text: 'First risk' }] },
          { type: 'docParagraph', content: [{ type: 'text', text: 'Conclusion' }] },
        ],
      },
    })
    editors.push(editor)

    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.list.apply',
          arguments: { target: { blockIndexes: [1] }, kind: 'ordered' },
        },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.list.apply',
      ok: true,
      output: {
        summary: 'Converted 1 block(s) to list items',
        matched: 1,
        changed: 1,
        skippedProtected: 0,
      },
    })
    expect(editor.state.doc.child(1).type.name).toBe('docListItem')
    expect(editor.state.doc.child(1).attrs).toMatchObject({
      kind: 'ordered',
      numId: '42',
      ilvl: 0,
      externalChanged: true,
    })
    expect(editor.state.doc.child(1).textContent).toBe('First risk')

    expect(editor.commands.undo()).toBe(true)
    expect(editor.state.doc.child(1).type.name).toBe('docParagraph')
    expect(editor.state.doc.child(1).textContent).toBe('First risk')
  })

  it('keeps matching DOCX list items unchanged without allocating new numbering', async () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docListItem',
            attrs: { kind: 'ordered', numId: '7', ilvl: 2 },
            content: [{ type: 'text', text: 'Already ordered' }],
          },
        ],
      },
    })
    editors.push(editor)

    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.list.apply',
          arguments: { target: { blockIndexes: [0] }, kind: 'ordered' },
        },
        {
          allocateListNumId: () => null,
          loadStaged: async () => undefined,
          save: async () => ({ fileName: 'document.docx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.list.apply',
      ok: true,
      output: {
        summary: 'No blocks matched; the document was not changed.',
        matched: 1,
        changed: 0,
        skippedProtected: 0,
      },
    })
    expect(editor.state.doc.child(0).type.name).toBe('docListItem')
    expect(editor.state.doc.child(0).attrs).toMatchObject({
      kind: 'ordered',
      numId: '7',
      ilvl: 2,
    })
    expect(editor.state.doc.child(0).textContent).toBe('Already ordered')
  })

  it('rejects a DOCX list-apply target without a condition', async () => {
    const editor = createEditor('Keep this paragraph')

    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.list.apply',
          arguments: { target: {}, kind: 'bullet' },
        },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.list.apply',
      ok: false,
      error: 'invalid_arguments',
      message: 'command #0 createParagraphBullets: target requires at least one condition',
    })
    expect(editor.state.doc.child(0).type.name).toBe('docParagraph')
    expect(editor.state.doc.child(0).textContent).toBe('Keep this paragraph')
  })

  it('removes DOCX list formatting through one native undo transaction', async () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          { type: 'docParagraph', content: [{ type: 'text', text: 'Introduction' }] },
          {
            type: 'docListItem',
            attrs: { kind: 'bullet', numId: '7', ilvl: 1 },
            content: [{ type: 'text', text: 'Risk item' }],
          },
          { type: 'docParagraph', content: [{ type: 'text', text: 'Conclusion' }] },
        ],
      },
    })
    editors.push(editor)

    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.list.remove',
          arguments: { target: { blockIndexes: [1] } },
        },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.list.remove',
      ok: true,
      output: {
        summary: 'Converted 1 list item(s) to body text',
        matched: 1,
        changed: 1,
        skippedProtected: 0,
      },
    })
    expect(editor.state.doc.child(1).type.name).toBe('docParagraph')
    expect(editor.state.doc.child(1).textContent).toBe('Risk item')

    expect(editor.commands.undo()).toBe(true)
    expect(editor.state.doc.child(1).type.name).toBe('docListItem')
    expect(editor.state.doc.child(1).attrs).toMatchObject({ kind: 'bullet', numId: '7', ilvl: 1 })
    expect(editor.state.doc.child(1).textContent).toBe('Risk item')
  })

  it('rejects a DOCX list-removal target without a condition', async () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docListItem',
            attrs: { kind: 'bullet', numId: '7' },
            content: [{ type: 'text', text: 'Keep this item' }],
          },
        ],
      },
    })
    editors.push(editor)

    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.list.remove',
          arguments: { target: {} },
        },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.list.remove',
      ok: false,
      error: 'invalid_arguments',
      message: 'command #0 deleteParagraphBullets: target requires at least one condition',
    })
    expect(editor.state.doc.child(0).type.name).toBe('docListItem')
    expect(editor.state.doc.child(0).textContent).toBe('Keep this item')
  })

  it('deletes matching DOCX blocks through one native undo transaction', async () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          { type: 'docParagraph', content: [{ type: 'text', text: 'Introduction' }] },
          { type: 'docParagraph', content: [{ type: 'text', text: 'Risk Notes' }] },
          { type: 'docParagraph', content: [{ type: 'text', text: 'Conclusion' }] },
        ],
      },
    })
    editors.push(editor)

    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.block.delete',
          arguments: { target: { containsText: 'Risk Notes' } },
        },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.block.delete',
      ok: true,
      output: {
        summary: 'Deleted 1 block(s)',
        matched: 1,
        changed: 1,
        skippedProtected: 0,
        skippedDeleted: 0,
      },
    })
    expect(editor.state.doc.childCount).toBe(2)
    expect(editor.getText()).not.toContain('Risk Notes')

    expect(editor.commands.undo()).toBe(true)
    expect(editor.state.doc.childCount).toBe(3)
    expect(editor.state.doc.child(1).textContent).toBe('Risk Notes')
  })

  it('rejects a DOCX block-delete target without a condition', async () => {
    const editor = createEditor('Keep this block')

    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.block.delete',
          arguments: { target: {} },
        },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.block.delete',
      ok: false,
      error: 'invalid_arguments',
      message: 'command #0 deleteBlocks: target requires at least one condition',
    })
    expect(editor.state.doc.childCount).toBe(1)
    expect(editor.state.doc.child(0).textContent).toBe('Keep this block')
  })

  it('sets a DOCX heading level through one native undo transaction', async () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          { type: 'docParagraph', content: [{ type: 'text', text: 'Introduction' }] },
          { type: 'docParagraph', content: [{ type: 'text', text: 'Risk Notes' }] },
        ],
      },
    })
    editors.push(editor)

    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.paragraph.set_heading_level',
          arguments: { target: { blockIndexes: [1] }, level: 2 },
        },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.paragraph.set_heading_level',
      ok: true,
      output: {
        summary: 'Updated heading level in 1 block(s)',
        matched: 1,
        changed: 1,
        skippedProtected: 0,
      },
    })
    expect(editor.state.doc.child(1).type.name).toBe('docHeading')
    expect(editor.state.doc.child(1).attrs.level).toBe(2)

    expect(editor.commands.undo()).toBe(true)
    expect(editor.state.doc.child(1).type.name).toBe('docParagraph')
    expect(editor.state.doc.child(1).textContent).toBe('Risk Notes')
  })

  it('rejects a heading-level target without a condition', async () => {
    const editor = createEditor('Risk Notes')

    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.paragraph.set_heading_level',
          arguments: { target: {}, level: 2 },
        },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.paragraph.set_heading_level',
      ok: false,
      error: 'invalid_arguments',
      message: 'command #0 setHeadingLevel: target requires at least one condition',
    })
    expect(editor.state.doc.child(0).type.name).toBe('docParagraph')
  })

  it('replaces all DOCX text through one native undo transaction', async () => {
    const editor = createEditor('Acme one, Acme two')

    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.text.replace_all',
          arguments: { containsText: 'Acme', replaceText: 'Codex' },
        },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.text.replace_all',
      ok: true,
      output: {
        summary: 'replaced 2 occurrence(s) in 1 block(s)',
        matched: 1,
        changed: 1,
        skippedProtected: 0,
        skippedDeleted: 0,
      },
    })
    expect(editor.getText()).toContain('Codex one, Codex two')

    expect(editor.commands.undo()).toBe(true)
    expect(editor.getText()).toContain('Acme one, Acme two')
  })

  it('rejects an empty DOCX replace-all search without mutating', async () => {
    const editor = createEditor('Acme')

    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.text.replace_all',
          arguments: { containsText: '', replaceText: 'Codex' },
        },
        services,
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.text.replace_all',
      ok: false,
      error: 'invalid_arguments',
      message: '$.containsText must not be empty.',
    })
    expect(editor.getText()).toContain('Acme')
  })

  it('loads staged DOCX bytes through docx.document.load_staged without checkpointing recovery', async () => {
    const editor = createEditor('Old content')
    const data = new Uint8Array([80, 75, 3, 4, 1, 2, 3, 4]).buffer
    let loaded:
      { readonly blobId: string; readonly name: string; readonly data: ArrayBuffer } | undefined

    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.document.load_staged',
          arguments: {
            blobId: 'docx-blob',
            name: 'loaded.docx',
            size: data.byteLength,
            data,
          },
        },
        {
          save: async () => ({ fileName: 'document.docx' }),
          loadStaged: async (input) => {
            loaded = input
          },
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.document.load_staged',
      ok: true,
      checkpointRecovery: false,
    })
    expect(loaded).toEqual({ blobId: 'docx-blob', name: 'loaded.docx', data })
  })

  it('rejects a staged DOCX load without hydrated ArrayBuffer bytes', async () => {
    const editor = createEditor('Old content')
    let loaded = false

    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.document.load_staged',
          arguments: {
            blobId: 'docx-blob',
            name: 'invalid.docx',
            size: 1,
            data: {},
          },
        },
        {
          save: async () => ({ fileName: 'document.docx' }),
          loadStaged: async () => {
            loaded = true
          },
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.document.load_staged',
      ok: false,
      error: 'invalid_arguments',
      message: '$.data must be a hydrated ArrayBuffer.',
    })
    expect(loaded).toBe(false)
  })

  it('rejects a staged load whose file identity is not DOCX', async () => {
    const editor = createEditor('Old content')
    const data = new Uint8Array([80, 75, 3, 4]).buffer
    let loaded = false

    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.document.load_staged',
          arguments: {
            blobId: 'docx-blob',
            name: 'invalid.pdf',
            size: data.byteLength,
            data,
          },
        },
        {
          save: async () => ({ fileName: 'document.docx' }),
          loadStaged: async () => {
            loaded = true
          },
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.document.load_staged',
      ok: false,
      error: 'invalid_arguments',
      message: 'docx.document.load_staged requires a valid staged DOCX descriptor.',
    })
    expect(loaded).toBe(false)
  })

  it('rejects staged DOCX bytes that do not match the declared size', async () => {
    const editor = createEditor('Old content')
    const data = new Uint8Array([80, 75, 3, 4]).buffer
    let loaded = false

    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.document.load_staged',
          arguments: {
            blobId: 'docx-blob',
            name: 'invalid.docx',
            size: data.byteLength + 1,
            data,
          },
        },
        {
          save: async () => ({ fileName: 'document.docx' }),
          loadStaged: async () => {
            loaded = true
          },
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.document.load_staged',
      ok: false,
      error: 'invalid_arguments',
      message: 'docx.document.load_staged requires a valid staged DOCX descriptor.',
    })
    expect(loaded).toBe(false)
  })

  it('reports a staged DOCX load failure as execution_failed', async () => {
    const editor = createEditor('Old content')
    const data = new Uint8Array([80, 75, 3, 4]).buffer

    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.document.load_staged',
          arguments: {
            blobId: 'docx-blob',
            name: 'broken.docx',
            size: data.byteLength,
            data,
          },
        },
        {
          save: async () => ({ fileName: 'document.docx' }),
          loadStaged: async () => {
            throw new Error('failed to parse DOCX')
          },
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.document.load_staged',
      ok: false,
      error: 'execution_failed',
      message: 'failed to parse DOCX',
    })
  })

  it('rejects the retired open_local_file transport alias without loading DOCX bytes', async () => {
    const editor = createEditor('Old content')
    const data = new Uint8Array([80, 75, 3, 4]).buffer
    let loaded = false

    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'open_local_file',
          arguments: {
            blobId: 'docx-blob',
            name: 'alias.docx',
            size: data.byteLength,
            data,
          },
        },
        {
          save: async () => ({ fileName: 'document.docx' }),
          loadStaged: async () => {
            loaded = true
          },
        },
      ),
    ).resolves.toEqual({ handled: false })
    expect(loaded).toBe(false)
  })

  it('inserts through docx.text.insert and the native DOCX undo route', async () => {
    const editor = createEditor('Hello')
    editor.commands.setTextSelection(6)

    expect(
      await executeDocxOperation(
        editor,
        {
          operation: 'docx.text.insert',
          arguments: { text: ' world' },
        },
        services,
      ),
    ).toEqual({
      handled: true,
      operationId: 'docx.text.insert',
      ok: true,
    })
    expect(editor.getText()).toContain('Hello world')

    expect(editor.commands.undo()).toBe(true)
    expect(editor.getText()).toContain('Hello')
    expect(editor.getText()).not.toContain(' world')
  })

  it('replaces through docx.text.replace_selection and the native DOCX undo route', async () => {
    const editor = createEditor('Hello world')
    editor.commands.setTextSelection({ from: 7, to: 12 })

    expect(
      await executeDocxOperation(
        editor,
        {
          operation: 'docx.text.replace_selection',
          arguments: { text: 'Codex' },
        },
        services,
      ),
    ).toEqual({
      handled: true,
      operationId: 'docx.text.replace_selection',
      ok: true,
    })
    expect(editor.getText()).toContain('Hello Codex')
    expect(editor.getText()).not.toContain('world')

    expect(editor.commands.undo()).toBe(true)
    expect(editor.getText()).toContain('Hello world')
  })

  it('saves through docx.document.save and returns the persisted file identity', async () => {
    const editor = createEditor('Hello')

    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.document.save',
          arguments: {},
        },
        {
          loadStaged: async () => undefined,
          save: async () => ({ fileName: 'draft.docx' }),
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.document.save',
      ok: true,
      output: { saved: true, fileName: 'draft.docx' },
      checkpointRecovery: false,
    })
  })

  it('reports a DOCX save failure as execution_failed', async () => {
    const editor = createEditor('Hello')

    await expect(
      executeDocxOperation(
        editor,
        {
          operation: 'docx.document.save',
          arguments: {},
        },
        {
          loadStaged: async () => undefined,
          save: async () => {
            throw new Error('disk unavailable')
          },
        },
      ),
    ).resolves.toEqual({
      handled: true,
      operationId: 'docx.document.save',
      ok: false,
      error: 'execution_failed',
      message: 'The document was not saved.',
    })
  })

  it('rejects retired public DOCX aliases without mutating the mounted editor', async () => {
    const editor = createEditor('Hello world')

    for (const [operation, arguments_] of [
      ['insert_text', { text: ' alias' }],
      ['replace_selection', { text: 'alias' }],
      ['save', {}],
    ] as const) {
      await expect(
        executeDocxOperation(editor, { operation, arguments: arguments_ }, services),
      ).resolves.toEqual({ handled: false })
    }
    expect(editor.getText()).toContain('Hello world')
  })
})
