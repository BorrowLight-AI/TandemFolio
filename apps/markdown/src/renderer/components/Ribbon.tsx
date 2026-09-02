import { useEffect, useRef, useState } from 'react'
import type { ReactNode } from 'react'
import { EditorFileIcon, EditorFullscreenIcon, EditorSaveIcon } from '@genoffice/ui'
import type { Editor } from '@tiptap/core'
import { useEditorState } from '@tiptap/react'
import { useI18n } from '../i18n/locale'
import type { StringKey } from '../i18n/locale'
import { runMarkdownHistoryAction } from '../editor/history-actions'
import {
  getMarkdownTextBlockIndexAtSelection,
  setMarkdownBlockType,
  type MarkdownBlockType,
} from '../editor/block-type-actions'
import {
  getMarkdownTextMarks,
  setMarkdownTextMarks,
  type MarkdownTextMarks,
} from '../editor/text-mark-actions'
import { getActiveMarkdownListType, setMarkdownListType } from '../editor/list-actions'
import { insertMarkdownDivider, insertMarkdownTable } from '../editor/structure-actions'
import {
  IconBullets,
  IconExportDocx,
  IconHr,
  IconInlineCode,
  IconLink,
  IconNumbered,
  IconPicture,
  IconPrint,
  IconProperties,
  IconRedo,
  IconTable,
  IconTaskList,
  IconUndo,
} from './icons'

interface Props {
  editor: Editor | null
  disabled: boolean
  dirty: boolean
  onOpen: () => void
  onSave: () => void
  onExportDocx: () => void
  onPrintPdf: () => void
  autoSave: boolean
  onToggleAutoSave: (on: boolean) => void
  onInsertImage: () => void
  frontmatterOpen: boolean
  onToggleFrontmatter: () => void
  fullscreen: boolean
  onToggleFullscreen: () => void
}

type BlockStyle = 'paragraph' | 'h1' | 'h2' | 'h3' | 'h4' | 'h5' | 'h6' | 'quote' | 'codeBlock'

const STYLE_LABEL: Record<BlockStyle, StringKey> = {
  paragraph: 'styleParagraph',
  h1: 'styleH1',
  h2: 'styleH2',
  h3: 'styleH3',
  h4: 'styleH4',
  h5: 'styleH5',
  h6: 'styleH6',
  quote: 'styleQuote',
  codeBlock: 'styleCodeBlock',
}

function applyBlockStyle(editor: Editor, style: BlockStyle): void {
  const textBlockIndex = getMarkdownTextBlockIndexAtSelection(editor)
  if (textBlockIndex === null) return
  const type: MarkdownBlockType =
    style === 'codeBlock'
      ? 'code_block'
      : style === 'quote'
        ? 'quote'
        : style === 'paragraph'
          ? 'paragraph'
          : (`heading_${style.slice(1)}` as MarkdownBlockType)
  setMarkdownBlockType(editor, { textBlockIndex, type })
}

function applyMarkToggle(editor: Editor, mark: 'bold' | 'italic' | 'strike' | 'code'): void {
  const current = getMarkdownTextMarks(editor)
  let marks: MarkdownTextMarks = { ...current, [mark]: !current[mark] }
  if (mark === 'code' && marks.code) {
    marks = { bold: false, italic: false, strike: false, code: true, link: null }
  } else if (mark !== 'code' && marks[mark]) {
    marks = { ...marks, code: false }
  }
  const { from, to } = editor.state.selection
  setMarkdownTextMarks(editor, { from, to, marks })
}

function applyListType(editor: Editor, type: 'bullet' | 'ordered' | 'task'): void {
  const textBlockIndex = getMarkdownTextBlockIndexAtSelection(editor)
  if (textBlockIndex === null) return
  setMarkdownListType(editor, {
    textBlockIndex,
    type: getActiveMarkdownListType(editor) === type ? 'none' : type,
  })
}

function IconBtn({
  title,
  active,
  disabled,
  onClick,
  children,
}: {
  title: string
  active?: boolean
  disabled?: boolean
  onClick: () => void
  children: ReactNode
}) {
  return (
    <button
      type="button"
      className={`rb-btn${active ? ' active' : ''}`}
      data-tip={title}
      aria-label={title}
      disabled={disabled}
      onMouseDown={(e) => e.preventDefault()}
      onClick={onClick}
    >
      {children}
    </button>
  )
}

export function Ribbon({
  editor,
  disabled,
  dirty,
  onOpen,
  onSave,
  onExportDocx,
  onPrintPdf,
  autoSave,
  onToggleAutoSave,
  onInsertImage,
  frontmatterOpen,
  onToggleFrontmatter,
  fullscreen,
  onToggleFullscreen,
}: Props) {
  const { t } = useI18n()
  const [linkOpen, setLinkOpen] = useState(false)
  const [linkUrl, setLinkUrl] = useState('')
  const linkInputRef = useRef<HTMLInputElement>(null)

  const state = useEditorState({
    editor,
    selector: ({ editor: e }) => {
      if (!e) return null
      const style: BlockStyle = e.isActive('codeBlock')
        ? 'codeBlock'
        : e.isActive('blockquote')
          ? 'quote'
          : e.isActive('heading')
            ? (`h${e.getAttributes('heading').level}` as BlockStyle)
            : 'paragraph'
      return {
        style,
        empty: e.isEmpty,
        bold: e.isActive('bold'),
        italic: e.isActive('italic'),
        strike: e.isActive('strike'),
        code: e.isActive('code'),
        link: e.isActive('link'),
        bullet: e.isActive('bulletList'),
        ordered: e.isActive('orderedList'),
        task: e.isActive('taskList'),
        canUndo: e.can().undo(),
        canRedo: e.can().redo(),
      }
    },
  })

  useEffect(() => {
    if (linkOpen) linkInputRef.current?.focus()
  }, [linkOpen])

  const off = disabled || !editor || !state

  const openLink = () => {
    if (!editor) return
    setLinkUrl(String(editor.getAttributes('link').href ?? ''))
    setLinkOpen((v) => !v)
  }

  const applyLink = () => {
    if (!editor) return
    const url = linkUrl.trim()
    editor.chain().focus().extendMarkRange('link').run()
    const { from, to } = editor.state.selection
    const marks = getMarkdownTextMarks(editor)
    setMarkdownTextMarks(editor, {
      from,
      to,
      marks: { ...marks, code: url ? false : marks.code, link: url || null },
    })
    setLinkOpen(false)
  }

  // 20px inline-row rendering, same as the docs toolbar these icons come from
  // (pinned stroke paints 1.5px at this size per the suite-wide icon rules)
  const ICON = 20

  return (
    <div className="ribbon">
      {/* quick-access row above the toolbar (save / undo / redo / autosave), same as the docs QAT row */}
      <div className="ribbon-tabs">
        <button
          type="button"
          className="qa-btn"
          data-tip={t('open')}
          aria-label={t('open')}
          onClick={onOpen}
        >
          <EditorFileIcon />
        </button>
        <button
          type="button"
          className="qa-btn"
          data-tip={t('save')}
          aria-label={t('save')}
          disabled={off || !dirty}
          onMouseDown={(e) => e.preventDefault()}
          onClick={onSave}
        >
          <EditorSaveIcon />
        </button>
        <button
          type="button"
          className="qa-btn"
          data-tip={t('undo')}
          aria-label={t('undo')}
          disabled={off || !state?.canUndo}
          onMouseDown={(e) => e.preventDefault()}
          onClick={() => editor && runMarkdownHistoryAction(editor, 'undo')}
        >
          <IconUndo size={16} />
        </button>
        <button
          type="button"
          className="qa-btn"
          data-tip={t('redo')}
          aria-label={t('redo')}
          disabled={off || !state?.canRedo}
          onMouseDown={(e) => e.preventDefault()}
          onClick={() => editor && runMarkdownHistoryAction(editor, 'redo')}
        >
          <IconRedo size={16} />
        </button>
        <label className={`autosave-toggle${autoSave ? ' on' : ''}`} data-tip={t('autoSaveTip')}>
          <span className="autosave-knob" />
          <span className="autosave-text">{t('autoSave')}</span>
          <input
            type="checkbox"
            checked={autoSave}
            onChange={(e) => onToggleAutoSave(e.target.checked)}
          />
        </label>
        <button
          type="button"
          className="qa-btn"
          data-tip={t('exportDocx')}
          aria-label={t('exportDocx')}
          onClick={onExportDocx}
          disabled={off}
        >
          <IconExportDocx size={16} />
        </button>
        <button
          type="button"
          className="qa-btn"
          data-tip={t('printPdf')}
          aria-label={t('printPdf')}
          onClick={onPrintPdf}
          disabled={off}
        >
          <IconPrint size={16} />
        </button>
        <button
          type="button"
          className="qa-btn"
          data-tip={t(fullscreen ? 'exitFullscreen' : 'enterFullscreen')}
          aria-label={t(fullscreen ? 'exitFullscreen' : 'enterFullscreen')}
          onClick={onToggleFullscreen}
          aria-pressed={fullscreen}
        >
          <EditorFullscreenIcon exit={fullscreen} />
        </button>
      </div>

      <div className="ribbon-body">
        <div className="ribbon-group">
          <div className="ribbon-group-items">
            <select
              className="rb-style"
              value={state?.style ?? 'paragraph'}
              disabled={off}
              onChange={(e) => editor && applyBlockStyle(editor, e.target.value as BlockStyle)}
            >
              {(Object.keys(STYLE_LABEL) as BlockStyle[]).map((s) => (
                <option key={s} value={s}>
                  {t(STYLE_LABEL[s])}
                </option>
              ))}
            </select>
          </div>
        </div>

        <div className="rb-sep" />

        <div className="ribbon-group">
          <div className="ribbon-group-items">
            <IconBtn
              title={t('bold')}
              active={state?.bold}
              disabled={off}
              onClick={() => editor && applyMarkToggle(editor, 'bold')}
            >
              <b>B</b>
            </IconBtn>
            <IconBtn
              title={t('italic')}
              active={state?.italic}
              disabled={off}
              onClick={() => editor && applyMarkToggle(editor, 'italic')}
            >
              <i>I</i>
            </IconBtn>
            <IconBtn
              title={t('strike')}
              active={state?.strike}
              disabled={off}
              onClick={() => editor && applyMarkToggle(editor, 'strike')}
            >
              <s>ab</s>
            </IconBtn>
            <IconBtn
              title={t('inlineCode')}
              active={state?.code}
              disabled={off}
              onClick={() => editor && applyMarkToggle(editor, 'code')}
            >
              <IconInlineCode size={ICON} />
            </IconBtn>
            <span className="rb-link-anchor">
              <IconBtn title={t('link')} active={state?.link} disabled={off} onClick={openLink}>
                <IconLink size={ICON} />
              </IconBtn>
              {linkOpen && (
                <span className="rb-link-pop" onMouseDown={(e) => e.stopPropagation()}>
                  <input
                    ref={linkInputRef}
                    value={linkUrl}
                    placeholder={t('linkPlaceholder')}
                    onChange={(e) => setLinkUrl(e.target.value)}
                    onKeyDown={(e) => {
                      if (e.key === 'Enter') applyLink()
                      if (e.key === 'Escape') setLinkOpen(false)
                    }}
                  />
                  <button type="button" onClick={applyLink}>
                    {t('linkApply')}
                  </button>
                </span>
              )}
            </span>
          </div>
        </div>

        <div className="rb-sep" />

        <div className="ribbon-group">
          <div className="ribbon-group-items">
            <IconBtn
              title={t('bulletList')}
              active={state?.bullet}
              disabled={off}
              onClick={() => editor && applyListType(editor, 'bullet')}
            >
              <IconBullets size={ICON} />
            </IconBtn>
            <IconBtn
              title={t('orderedList')}
              active={state?.ordered}
              disabled={off}
              onClick={() => editor && applyListType(editor, 'ordered')}
            >
              <IconNumbered size={ICON} />
            </IconBtn>
            <IconBtn
              title={t('taskList')}
              active={state?.task}
              disabled={off}
              onClick={() => editor && applyListType(editor, 'task')}
            >
              <IconTaskList size={ICON} />
            </IconBtn>
          </div>
        </div>

        <div className="rb-sep" />

        <div className="ribbon-group">
          <div className="ribbon-group-items">
            <IconBtn
              title={t('insertTable')}
              disabled={off}
              onClick={() =>
                editor &&
                insertMarkdownTable(editor, {
                  position: editor.state.selection.from,
                  rows: 3,
                  columns: 3,
                  headerRow: true,
                })
              }
            >
              <IconTable size={ICON} />
            </IconBtn>
            <IconBtn title={t('insertImage')} disabled={off} onClick={onInsertImage}>
              <IconPicture size={ICON} />
            </IconBtn>
            <IconBtn
              title={t('insertHr')}
              disabled={off}
              onClick={() =>
                editor && insertMarkdownDivider(editor, { position: editor.state.selection.from })
              }
            >
              <IconHr size={ICON} />
            </IconBtn>
          </div>
        </div>

        <div className="rb-spacer" />

        <div className="ribbon-group">
          <div className="ribbon-group-items">
            <IconBtn
              title={t('fmProperties')}
              active={frontmatterOpen}
              disabled={disabled}
              onClick={onToggleFrontmatter}
            >
              <IconProperties size={ICON} />
            </IconBtn>
          </div>
        </div>
      </div>
    </div>
  )
}
