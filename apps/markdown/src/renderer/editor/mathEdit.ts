// Adapted from genspark-ai/genoffice Markdown math editing at upstream/main.
import type { Editor } from '@tiptap/core'
import katex from 'katex'
import { t } from '../i18n/locale'
import { insertMarkdownMath, setMarkdownMath } from './math-actions'

interface AnchorRect {
  left: number
  top: number
  bottom: number
}

interface OpenMathOptions {
  position: number | null
  anchor: AnchorRect
}

let activePopover: { element: HTMLElement; dispose: () => void } | null = null

function closeActive(): void {
  activePopover?.dispose()
  activePopover = null
}

export function openMathEditor(editor: Editor, options: OpenMathOptions): void {
  closeActive()
  const { position, anchor } = options
  const node = position === null ? null : editor.state.doc.nodeAt(position)
  const isBlock = node ? node.type.name === 'blockMath' : true
  const expectedType = node?.type.name ?? null

  const targetNode = () => {
    if (position === null) return null
    const target = editor.state.doc.nodeAt(position)
    return target && target.type.name === expectedType ? target : null
  }

  const popover = document.createElement('div')
  popover.className = 'md-math-pop'
  const input = document.createElement('textarea')
  input.className = 'md-math-input'
  input.rows = isBlock ? 4 : 2
  input.placeholder = t('mathPlaceholder')
  input.value = String(node?.attrs.latex ?? '')
  input.spellcheck = false
  const preview = document.createElement('div')
  preview.className = 'md-math-preview'
  const actions = document.createElement('div')
  actions.className = 'md-math-actions'
  const apply = document.createElement('button')
  apply.className = 'md-math-apply'
  apply.textContent = t('linkApply')
  actions.appendChild(apply)

  const close = (): void => {
    document.removeEventListener('pointerdown', onPointerDown, true)
    editor.off('transaction', onTransaction)
    popover.remove()
    if (activePopover?.element === popover) activePopover = null
  }

  if (position !== null) {
    const remove = document.createElement('button')
    remove.className = 'md-math-delete'
    remove.textContent = t('blockDelete')
    remove.addEventListener('click', () => {
      if (targetNode()) setMarkdownMath(editor, { position, latex: null })
      close()
    })
    actions.appendChild(remove)
  }

  const renderPreview = (): void => {
    const latex = input.value.trim()
    preview.classList.toggle('md-math-preview-empty', latex === '')
    if (!latex) {
      preview.textContent = ''
      return
    }
    katex.render(latex, preview, { throwOnError: false, displayMode: isBlock })
  }

  const applyLatex = (): void => {
    const latex = input.value.trim()
    editor.commands.focus()
    if (position === null) {
      if (latex) {
        insertMarkdownMath(editor, {
          position: editor.state.selection.from,
          display: 'block',
          latex,
        })
      }
    } else if (targetNode()) {
      setMarkdownMath(editor, { position, latex: latex || null })
    }
    close()
  }

  const onPointerDown = (event: PointerEvent): void => {
    if (!popover.contains(event.target as Node)) close()
  }
  const onTransaction = ({ transaction }: { transaction: { docChanged: boolean } }): void => {
    if (transaction.docChanged) close()
  }

  input.addEventListener('input', renderPreview)
  apply.addEventListener('click', applyLatex)
  popover.addEventListener('keydown', (event) => {
    if (event.key === 'Escape') {
      event.stopPropagation()
      close()
      editor.commands.focus()
    } else if (event.key === 'Enter' && (event.metaKey || event.ctrlKey)) {
      event.preventDefault()
      applyLatex()
    }
  })
  document.addEventListener('pointerdown', onPointerDown, true)
  editor.on('transaction', onTransaction)

  popover.append(input, preview, actions)
  document.body.appendChild(popover)
  activePopover = { element: popover, dispose: close }
  renderPreview()
  const margin = 8
  const left = Math.max(
    margin,
    Math.min(anchor.left, window.innerWidth - popover.offsetWidth - margin),
  )
  let top = anchor.bottom + 6
  if (top + popover.offsetHeight > window.innerHeight - margin) {
    top = Math.max(margin, anchor.top - popover.offsetHeight - 6)
  }
  popover.style.left = `${left}px`
  popover.style.top = `${top}px`
  input.focus()
  input.setSelectionRange(input.value.length, input.value.length)
}

export function openMathCreate(editor: Editor): void {
  const coordinates = editor.view.coordsAtPos(editor.state.selection.from)
  openMathEditor(editor, { position: null, anchor: coordinates })
}
