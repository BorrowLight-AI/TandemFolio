import { Extension } from '@tiptap/core'
import { getMarkdownTopLevelBlockIndexAtSelection, updateMarkdownBlock } from './block-actions'
import { getMarkdownTextMarks, setMarkdownTextMarks } from './text-mark-actions'

declare module '@tiptap/core' {
  interface Commands<ReturnType> {
    blockKeymap: {
      /** duplicate the top-level block at the caret and move the caret into the copy */
      duplicateBlock: () => ReturnType
      /** delete the top-level block at the caret (the last block becomes an empty paragraph) */
      deleteBlock: () => ReturnType
      moveBlockUp: () => ReturnType
      moveBlockDown: () => ReturnType
    }
  }
}

/**
 * SecondBrain-parity block shortcuts: ⌘D duplicate, ⌘⇧⌫ delete,
 * ⌘⇧↑/↓ move the current top-level block.
 */
export const BlockKeymap = Extension.create({
  name: 'blockKeymap',

  addCommands() {
    return {
      duplicateBlock:
        () =>
        ({ dispatch, editor, tr }) => {
          const blockIndex = getMarkdownTopLevelBlockIndexAtSelection(editor)
          if (blockIndex === null) return false
          if (!dispatch) return true
          return updateMarkdownBlock(editor, {
            blockIndex,
            action: 'duplicate',
            afterBlockIndex: null,
            content: null,
          }, dispatch, tr).ok
        },
      deleteBlock:
        () =>
        ({ dispatch, editor, tr }) => {
          const blockIndex = getMarkdownTopLevelBlockIndexAtSelection(editor)
          if (blockIndex === null) return false
          if (!dispatch) return true
          return updateMarkdownBlock(editor, {
            blockIndex,
            action: 'delete',
            afterBlockIndex: null,
            content: null,
          }, dispatch, tr).ok
        },
      moveBlockUp:
        () =>
        ({ dispatch, editor, tr }) => {
          const blockIndex = getMarkdownTopLevelBlockIndexAtSelection(editor)
          if (blockIndex === null || blockIndex === 0) return false
          if (!dispatch) return true
          return updateMarkdownBlock(editor, {
            blockIndex,
            action: 'move',
            afterBlockIndex: blockIndex - 2,
            content: null,
          }, dispatch, tr).ok
        },
      moveBlockDown:
        () =>
        ({ dispatch, editor, tr }) => {
          const blockIndex = getMarkdownTopLevelBlockIndexAtSelection(editor)
          if (blockIndex === null || blockIndex + 1 >= editor.state.doc.childCount) return false
          if (!dispatch) return true
          return updateMarkdownBlock(editor, {
            blockIndex,
            action: 'move',
            afterBlockIndex: blockIndex + 1,
            content: null,
          }, dispatch, tr).ok
        },
    }
  },

  addKeyboardShortcuts() {
    return {
      'Mod-d': () => this.editor.commands.duplicateBlock(),
      'Mod-D': () => this.editor.commands.duplicateBlock(),
      'Mod-Shift-Backspace': () => this.editor.commands.deleteBlock(),
      'Mod-Shift-ArrowUp': () => this.editor.commands.moveBlockUp(),
      'Mod-Shift-ArrowDown': () => this.editor.commands.moveBlockDown(),
      'Mod-k': () => {
        const editor = this.editor
        editor.chain().focus().extendMarkRange('link').run()
        const { from, to } = editor.state.selection
        const marks = getMarkdownTextMarks(editor)
        if (editor.isActive('link')) {
          return setMarkdownTextMarks(editor, { from, to, marks: { ...marks, link: null } }).ok
        }
        const href = window.prompt('URL:')
        if (!href?.trim()) return true
        return setMarkdownTextMarks(editor, {
          from,
          to,
          marks: { ...marks, code: false, link: href.trim() },
        }).ok
      },
    }
  },
})
