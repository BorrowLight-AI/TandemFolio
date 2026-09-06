// Selectively ported from genspark-ai/genoffice@f2c3d0879df29622d5a447935d2b4aeac033544d.
import type { AnyExtension } from '@tiptap/core'
import { Extension } from '@tiptap/core'
import { BlockMath, InlineMath } from '@tiptap/extension-mathematics'
import { Plugin } from '@tiptap/pm/state'
import { openMathEditor } from './mathEdit'

/**
 * Keep inline math strict enough that adjacent currency amounts such as
 * "paid $5 and $10" remain ordinary Markdown text.
 */
const STRICT_INLINE_MATH_RE = /^\$(?!\s)([^$\n]*[^\s$])\$(?!\d)/

const StrictInlineMath = InlineMath.extend({
  markdownTokenizer: {
    name: 'inlineMath',
    level: 'inline',
    start: (source: string) => source.indexOf('$'),
    tokenize: (source: string) => {
      const match = STRICT_INLINE_MATH_RE.exec(source)
      if (!match) return undefined
      return { type: 'inlineMath', raw: match[0], latex: match[1].trim() }
    },
  },
})

const MathClickEdit = Extension.create({
  name: 'mathClickEdit',

  addProseMirrorPlugins() {
    const editor = this.editor
    return [
      new Plugin({
        props: {
          handleClickOn: (view, _position, node, nodePosition, event) => {
            if (node.type.name !== 'blockMath' && node.type.name !== 'inlineMath') return false
            if (!view.editable) return false
            const target = event.target as HTMLElement | null
            const anchor = target?.closest?.('.tiptap-mathematics-render') ?? target
            if (!anchor) return false
            openMathEditor(editor, {
              position: nodePosition,
              anchor: anchor.getBoundingClientRect(),
            })
            return true
          },
        },
      }),
    ]
  },
})

export function buildMathExtensions(): AnyExtension[] {
  return [BlockMath, StrictInlineMath, MathClickEdit]
}
