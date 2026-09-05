import { describe, expect, it, vi } from 'vitest'
import { createElement } from 'react'
import { createRoot } from 'react-dom/client'
import { act } from 'react'
import { Editor } from '@tiptap/core'
import { NodeSelection, TextSelection } from '@tiptap/pm/state'
import { editorExtensions } from '../src/renderer/editor/extensions'
import {
  EditorContextMenu,
  FontDialog,
  ParagraphDialog,
} from '../src/renderer/components/ContextMenu'

function createEditor(paraAttrs: Record<string, unknown> = {}): Editor {
  return new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: {
      type: 'doc',
      content: [
        {
          type: 'docParagraph',
          attrs: { docxIndex: 0, ...paraAttrs },
          content: [{ type: 'text', text: 'EVs market research' }],
        },
      ],
    },
  })
}

function createTableEditor(): Editor {
  return new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: {
      type: 'doc',
      content: [
        {
          type: 'docTable',
          attrs: { docxIndex: null },
          content: [
            {
              type: 'docTableRow',
              content: [
                {
                  type: 'docTableCell',
                  content: [{ type: 'docParagraph', content: [{ type: 'text', text: 'cell' }] }],
                },
              ],
            },
          ],
        },
      ],
    },
  })
}

function select(editor: Editor, from: number, to: number) {
  editor.view.dispatch(
    editor.state.tr.setSelection(TextSelection.create(editor.state.doc, from, to)),
  )
}

function pickDropdown(container: Element, dropdown: HTMLButtonElement, value: string) {
  act(() => dropdown.click())
  const item = container.querySelector<HTMLButtonElement>(`.gs-dd-item[data-value="${value}"]`)!
  act(() => item.click())
}

function render(element: React.ReactElement): { container: HTMLElement; unmount: () => void } {
  const container = document.createElement('div')
  document.body.appendChild(container)
  const root = createRoot(container)
  act(() => root.render(element))
  return {
    container,
    unmount: () => {
      act(() => root.unmount())
      container.remove()
    },
  }
}

const noop = () => {}

function menuProps(editor: Editor, overrides: Record<string, unknown> = {}) {
  return {
    editor,
    menu: { x: 10, y: 10 },
    onClose: noop,
    onFontDialog: noop,
    onParagraphDialog: noop,
    onLink: noop,
    onNewComment: noop,
    ...overrides,
  }
}

describe('EditorContextMenu', () => {
  it('disables selection-dependent items when nothing is selected', () => {
    const editor = createEditor()
    const { container, unmount } = render(createElement(EditorContextMenu, menuProps(editor)))
    const byLabel = (label: string) =>
      [...container.querySelectorAll<HTMLButtonElement>('.ctx-item')].find(
        (b) => b.querySelector('.ctx-label')?.textContent === label,
      )!
    expect(byLabel('剪切').disabled).toBe(true)
    expect(byLabel('复制').disabled).toBe(true)
    expect(byLabel('粘贴').disabled).toBe(false)
    expect(byLabel('字体…').disabled).toBe(false)
    expect(byLabel('段落…').disabled).toBe(false)
    expect(byLabel('新建批注').disabled).toBe(true)
    unmount()
    editor.destroy()
  })

  it('enables everything and routes New Comment / Font… when text is selected', () => {
    const editor = createEditor()
    select(editor, 1, 5)
    const onNewComment = vi.fn()
    const onFontDialog = vi.fn()
    const onClose = vi.fn()
    const { container, unmount } = render(
      createElement(EditorContextMenu, menuProps(editor, { onNewComment, onFontDialog, onClose })),
    )
    const byLabel = (label: string) =>
      [...container.querySelectorAll<HTMLButtonElement>('.ctx-item')].find(
        (b) => b.querySelector('.ctx-label')?.textContent === label,
      )!
    expect(byLabel('剪切').disabled).toBe(false)
    expect(byLabel('新建批注').disabled).toBe(false)
    act(() => byLabel('新建批注').click())
    expect(onNewComment).toHaveBeenCalledOnce()
    expect(onClose).toHaveBeenCalled()
    act(() => byLabel('字体…').click())
    expect(onFontDialog).toHaveBeenCalledOnce()
    unmount()
    editor.destroy()
  })

  it('offers native insert/delete/select/merge/split commands inside a table', () => {
    const editor = createTableEditor()
    editor.view.dispatch(
      editor.state.tr.setSelection(TextSelection.near(editor.state.doc.resolve(3))),
    )
    const { container, unmount } = render(createElement(EditorContextMenu, menuProps(editor)))
    const labels = [...container.querySelectorAll('.ctx-label')].map((node) => node.textContent)
    expect(labels).toContain('插入')
    expect(labels).toContain('删除')
    expect(labels).toContain('选择')
    expect(labels).toContain('合并单元格')
    expect(labels).toContain('拆分单元格')
    unmount()
    editor.destroy()
  })

  it('arranges a selected image through native history and floats an inline image', () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docProtected',
            attrs: {
              docxIndex: 0,
              blockType: 'image',
              imageDataUrl: 'data:image/gif;base64,R0lGODdh',
              imageWrap: null,
              imageZOrder: null,
            },
          },
        ],
      },
    })
    editor.view.dispatch(editor.state.tr.setSelection(NodeSelection.create(editor.state.doc, 0)))
    const { container, unmount } = render(createElement(EditorContextMenu, menuProps(editor)))
    const arrange = [...container.querySelectorAll<HTMLButtonElement>('.ctx-item')].find(
      (button) => button.querySelector('.ctx-label')?.textContent === '排列',
    )!
    act(() => arrange.dispatchEvent(new MouseEvent('mouseover', { bubbles: true })))
    const front = [...container.querySelectorAll<HTMLButtonElement>('.ctx-item')].find(
      (button) => button.querySelector('.ctx-label')?.textContent === '置于顶层',
    )!
    act(() => front.click())
    expect(editor.state.doc.child(0).attrs).toMatchObject({ imageWrap: 'front', imageZOrder: 1 })
    expect(editor.commands.undo()).toBe(true)
    expect(editor.state.doc.child(0).attrs).toMatchObject({ imageWrap: null, imageZOrder: null })
    unmount()
    editor.destroy()
  })
})

describe('FontDialog', () => {
  it('applies font marks to the selection on OK', () => {
    const editor = createEditor()
    select(editor, 1, 10)
    const { container, unmount } = render(createElement(FontDialog, { editor, onClose: noop }))
    const dropdowns = container.querySelectorAll<HTMLButtonElement>('.gs-dd-btn')
    pickDropdown(container, dropdowns[1]!, 'bold')
    const ok = [...container.querySelectorAll('button')].find((b) => b.textContent === '确定')!
    act(() => ok.click())
    expect(editor.isActive('bold')).toBe(true)
    expect(editor.getAttributes('docTextStyle').sizeHalfPoints).toBe(22)
    unmount()
    editor.destroy()
  })
})

describe('ParagraphDialog', () => {
  it('applies alignment and spacing to the paragraph on OK', () => {
    const editor = createEditor()
    select(editor, 2, 2)
    const { container, unmount } = render(createElement(ParagraphDialog, { editor, onClose: noop }))
    pickDropdown(container, container.querySelector<HTMLButtonElement>('.gs-dd-btn')!, 'center')
    const ok = [...container.querySelectorAll('button')].find((b) => b.textContent === '确定')!
    act(() => ok.click())
    expect(editor.getAttributes('docParagraph').align).toBe('center')
    unmount()
    editor.destroy()
  })

  it('resolves visual left/right against the paragraph direction (LTR)', () => {
    const editor = createEditor({ align: 'right' })
    select(editor, 2, 2)
    const { container, unmount } = render(createElement(ParagraphDialog, { editor, onClose: noop }))
    const alignDropdown = container.querySelector<HTMLButtonElement>('.gs-dd-btn')!
    expect(alignDropdown.dataset.value).toBe('right')
    pickDropdown(container, alignDropdown, 'left')
    const ok = [...container.querySelectorAll('button')].find((b) => b.textContent === '确定')!
    act(() => ok.click())
    // visual left is the start side in LTR → stored as null
    expect(editor.getAttributes('docParagraph').align).toBeNull()
    unmount()
    editor.destroy()
  })

  it('shows the start side as Right in RTL and stores visual left explicitly', () => {
    const editor = createEditor({ bidi: true })
    select(editor, 2, 2)
    const { container, unmount } = render(createElement(ParagraphDialog, { editor, onClose: noop }))
    const alignDropdown = container.querySelector<HTMLButtonElement>('.gs-dd-btn')!
    // unset align in an RTL paragraph renders right, so the dialog shows Right
    expect(alignDropdown.dataset.value).toBe('right')
    pickDropdown(container, alignDropdown, 'left')
    const ok = [...container.querySelectorAll('button')].find((b) => b.textContent === '确定')!
    act(() => ok.click())
    // visual left is the end side in RTL → stored explicitly
    expect(editor.getAttributes('docParagraph').align).toBe('left')
    unmount()
    editor.destroy()
  })

  it('clears the align attr when re-selecting the start side in RTL', () => {
    const editor = createEditor({ bidi: true, align: 'left' })
    select(editor, 2, 2)
    const { container, unmount } = render(createElement(ParagraphDialog, { editor, onClose: noop }))
    const alignDropdown = container.querySelector<HTMLButtonElement>('.gs-dd-btn')!
    expect(alignDropdown.dataset.value).toBe('left')
    pickDropdown(container, alignDropdown, 'right')
    const ok = [...container.querySelectorAll('button')].find((b) => b.textContent === '确定')!
    act(() => ok.click())
    expect(editor.getAttributes('docParagraph').align).toBeNull()
    unmount()
    editor.destroy()
  })
})
