import { useState } from 'react'
import type { Editor } from '@tiptap/core'
import {
  bibliographyLine,
  citationText,
  type Block,
  type SourceInfo,
  type TocEntry,
} from '@genoffice/docx-engine'
import { PromptModal } from './PromptModal'
import { collectHeadings } from '../editor/headings'
import {
  createDocxTocFieldNodes,
  findFirstDocxTocBlockIndex,
  refreshDocxToc,
} from '../editor/toc-actions'
import { insertDocxBibliography, insertDocxCitation } from '../editor/source-actions'
import { insertDocxCaption } from '../editor/caption-actions'
import { insertDocxIndex, markDocxIndexEntry } from '../editor/index-actions'
import { t, useI18n, type StringKey } from '../i18n/locale'
import {
  IconBook,
  IconCaption,
  IconCaret,
  IconCitation,
  IconEndnote,
  IconFootnote,
  IconIndex,
  IconRefresh,
  IconToc,
} from './icons'

/** icon size for the big icon-over-label ribbon buttons (slides ribbon parity) */
import { BIG, TabProps, toggleDropdown } from './ribbon-tabs'

function collectTocEntries(editor: Editor): TocEntry[] {
  return collectHeadings(editor.state.doc).map(({ level, text }) => ({ level, text }))
}

/** Heading entries + real page numbers (headingPages and collectTocEntries share document order) */
function collectTocEntriesWithPages(
  editor: Editor,
  headingPages?: () => number[] | null,
): TocEntry[] {
  const entries = collectTocEntries(editor)
  const pages = headingPages?.()
  if (pages && pages.length === entries.length) {
    return entries.map((e, i) => ({ ...e, pageNo: pages[i] }))
  }
  return entries
}

/** PM nodes for a freshly generated TOC field (one docProtected per line) */
function tocFieldNodes(entries: TocEntry[]): Array<Record<string, unknown>> {
  return createDocxTocFieldNodes(entries, t('ribbonTocFieldLabel'))
}

/** OOXML of a top-level PM node: editor-generated fragment or original slice */
function xmlOfNode(node: { attrs: Record<string, unknown> }, blocks: Block[]): string {
  if (node.attrs.genXml) return String(node.attrs.genXml)
  const idx = node.attrs.docxIndex
  if (idx === null || idx === undefined) return ''
  return blocks.find((b) => b.docxIndex === idx)?.originalXml ?? ''
}

/**
 * Find the TOC field region (field begin ... matching end, tracked by fldChar
 * depth across top-level blocks) and replace it with a regenerated dirty TOC
 * field built from the current headings.
 */
function updateTocField(
  editor: Editor,
  blocks: Block[],
  headingPages?: () => number[] | null,
): void {
  const entries = collectTocEntriesWithPages(editor, headingPages)
  if (entries.length === 0) {
    window.alert(t('ribbonTocNoHeadings'))
    return
  }

  const tocBlockIndex = findFirstDocxTocBlockIndex(editor, blocks)
  if (tocBlockIndex === null) {
    window.alert(t('ribbonTocNotFound'))
    return
  }
  const result = refreshDocxToc(
    editor,
    blocks,
    {
      tocBlockIndex,
      entries: entries.map((entry) => ({
        level: entry.level,
        text: entry.text,
        pageNumber: entry.pageNo ?? null,
      })),
    },
    { toc: t('ribbonTocFieldLabel'), pageBreak: t('ribbonPageBreak') },
  )
  if (!result.ok) {
    window.alert(result.message)
    return
  }
  editor.commands.focus()
}

const CAPTION_LABEL_KEYS = [
  'ribbonCaptionFigure',
  'ribbonCaptionTable',
  'ribbonCaptionEquation',
] as const

/** Caption dialog: label + optional description, numbered by existing SEQ fields */
function CaptionModal({
  editor,
  blocks,
  onClose,
}: {
  editor: Editor
  blocks: Block[]
  onClose: () => void
}) {
  const { t } = useI18n()
  const CAPTION_LABELS = CAPTION_LABEL_KEYS.map((k) => t(k))
  const [label, setLabel] = useState<string>(() => t(CAPTION_LABEL_KEYS[0]))
  const [text, setText] = useState('')

  const nextNumber = (lbl: string): number => {
    let count = 0
    const re = new RegExp(`SEQ\\s+${lbl}[\\s\\\\]`)
    editor.state.doc.forEach((node) => {
      if (node.type.name !== 'docProtected') return
      if (re.test(xmlOfNode(node as never, blocks))) count += 1
    })
    return count + 1
  }

  const insert = () => {
    const number = nextNumber(label)
    const afterBlockIndex = Math.min(
      editor.state.selection.$head.index(0),
      editor.state.doc.childCount - 1,
    )
    editor.commands.focus()
    insertDocxCaption(editor, {
      afterBlockIndex,
      label,
      number,
      text: text.trim(),
    })
    onClose()
  }

  return (
    <div className="modal-backdrop" onMouseDown={(e) => e.target === e.currentTarget && onClose()}>
      <div className="modal">
        <h2>{t('ribbonCaptionInsertTitle')}</h2>
        <label>
          {t('ribbonCaptionLabel')}
          <select value={label} onChange={(e) => setLabel(e.target.value)}>
            {CAPTION_LABELS.map((l) => (
              <option key={l} value={l}>
                {l}
              </option>
            ))}
          </select>
        </label>
        <label>
          {t('ribbonCaption')}
          <input
            value={text}
            onChange={(e) => setText(e.target.value)}
            placeholder={t('ribbonCaptionPh', { label, n: nextNumber(label) })}
            onKeyDown={(e) => e.key === 'Enter' && insert()}
          />
        </label>
        <div className="modal-actions">
          <button className="btn-ghost" onClick={onClose}>
            {t('ribbonCancel')}
          </button>
          <button className="btn-primary" onClick={insert}>
            {t('ribbonInsert')}
          </button>
        </div>
      </div>
    </div>
  )
}

const SOURCE_TYPES: Array<{ key: string; nameKey: StringKey }> = [
  { key: 'Book', nameKey: 'ribbonSourceBook' },
  { key: 'JournalArticle', nameKey: 'ribbonSourceJournal' },
  { key: 'InternetSite', nameKey: 'ribbonSourceWebsite' },
  { key: 'Report', nameKey: 'ribbonSourceReport' },
  { key: 'Misc', nameKey: 'ribbonSourceMisc' },
]

/** Create Source dialog for the citation manager */
function SourceModal({
  sources,
  onAdd,
  onClose,
}: {
  sources: SourceInfo[]
  onAdd: (source: SourceInfo) => void
  onClose: () => void
}) {
  const { t } = useI18n()
  const [type, setType] = useState('Book')
  const [author, setAuthor] = useState('')
  const [title, setTitle] = useState('')
  const [year, setYear] = useState('')
  const [publisher, setPublisher] = useState('')
  const [url, setUrl] = useState('')

  const add = () => {
    if (!title.trim()) return
    // unique tag derived from author/title initials + year
    const base = (author.trim() || title.trim()).replace(/[\s,，、]+/g, '').slice(0, 8) || 'Src'
    let tag = `${base}${year.trim()}`
    let n = 1
    while (sources.some((s) => s.tag === tag)) tag = `${base}${year.trim()}_${++n}`
    onAdd({
      tag,
      type,
      author: author.trim(),
      title: title.trim(),
      year: year.trim(),
      ...(publisher.trim() ? { publisher: publisher.trim() } : {}),
      ...(url.trim() ? { url: url.trim() } : {}),
    })
    onClose()
  }

  return (
    <div className="modal-backdrop" onMouseDown={(e) => e.target === e.currentTarget && onClose()}>
      <div className="modal">
        <h2>{t('ribbonSourceCreateTitle')}</h2>
        <label>
          {t('ribbonSourceType')}
          <select value={type} onChange={(e) => setType(e.target.value)}>
            {SOURCE_TYPES.map((s) => (
              <option key={s.key} value={s.key}>
                {t(s.nameKey)}
              </option>
            ))}
          </select>
        </label>
        <label>
          {t('ribbonSourceAuthor')}
          <input
            value={author}
            onChange={(e) => setAuthor(e.target.value)}
            placeholder={t('ribbonSourceAuthorPh')}
          />
        </label>
        <label>
          {t('ribbonSourceTitle')}
          <input
            value={title}
            onChange={(e) => setTitle(e.target.value)}
            placeholder={t('ribbonSourceTitlePh')}
          />
        </label>
        <label>
          {t('ribbonSourceYear')}
          <input value={year} onChange={(e) => setYear(e.target.value)} placeholder="2026" />
        </label>
        <label>
          {type === 'JournalArticle'
            ? t('ribbonSourceJournalName')
            : type === 'InternetSite'
              ? t('ribbonSourceSiteName')
              : t('ribbonSourcePublisher')}
          <input value={publisher} onChange={(e) => setPublisher(e.target.value)} />
        </label>
        {type === 'InternetSite' && (
          <label>
            URL
            <input value={url} onChange={(e) => setUrl(e.target.value)} placeholder="https://…" />
          </label>
        )}
        <div className="modal-actions">
          <button className="btn-ghost" onClick={onClose}>
            {t('ribbonCancel')}
          </button>
          <button className="btn-primary" disabled={!title.trim()} onClick={add}>
            {t('ribbonAdd')}
          </button>
        </div>
      </div>
    </div>
  )
}

interface ReferencesTabProps extends TabProps {
  blocks: Block[]
  onInsertNote: (kind: 'footnote' | 'endnote') => void
  sources: SourceInfo[]
  onAddSource: (source: SourceInfo) => void
  /** TOC page-number backfill: docHeadings in document order → real page numbers */
  headingPages?: () => number[] | null
}

export function ReferencesTab({
  editor,
  hasDoc,
  blocks,
  dropdown,
  setDropdown,
  onInsertNote,
  sources,
  onAddSource,
  headingPages,
}: ReferencesTabProps) {
  const { t } = useI18n()
  const [captionOpen, setCaptionOpen] = useState(false)
  const [sourceOpen, setSourceOpen] = useState(false)

  const insertToc = () => {
    const entries = collectTocEntriesWithPages(editor, headingPages)
    if (entries.length === 0) {
      window.alert(t('ribbonTocNoHeadings'))
      return
    }
    editor
      .chain()
      .focus()
      .insertContent(tocFieldNodes(entries) as never)
      .run()
  }

  const insertCitation = (source: SourceInfo) => {
    const { from, to } = editor.state.selection
    insertDocxCitation(editor, sources, {
      range: { from, to },
      sourceTag: source.tag,
      displayText: citationText(source),
    })
    setDropdown(() => null)
  }

  const insertBibliography = () => {
    if (sources.length === 0) {
      window.alert(t('ribbonNoSources'))
      return
    }
    const afterBlockIndex = Math.min(
      editor.state.selection.$head.index(0),
      editor.state.doc.childCount - 1,
    )
    insertDocxBibliography(editor, sources, {
      afterBlockIndex,
      heading: t('ribbonBibliographyHeading'),
      entries: sources.map((source) => ({
        sourceTag: source.tag,
        text: bibliographyLine(source),
      })),
    })
    setDropdown(() => null)
  }

  const [xePrompt, setXePrompt] = useState<{
    initial: string
    range: { from: number; to: number }
  } | null>(null)

  const markIndexEntry = () => {
    const { from, to } = editor.state.selection
    setXePrompt({ initial: editor.state.doc.textBetween(from, to, ' ').trim(), range: { from, to } })
  }

  const submitIndexEntry = (term: string) => {
    if (!xePrompt) return
    editor.commands.focus()
    const result = markDocxIndexEntry(editor, { range: xePrompt.range, term })
    if (!result.ok) window.alert(result.message)
  }

  const insertIndex = () => {
    const terms: string[] = []
    editor.state.doc.descendants((node) => {
      if (node.type.name === 'docXeMark') terms.push(String(node.attrs.term))
    })
    if (terms.length === 0) {
      window.alert(t('ribbonNoIndexEntries'))
      return
    }
    const afterBlockIndex = Math.min(
      editor.state.selection.$head.index(0),
      editor.state.doc.childCount - 1,
    )
    editor.commands.focus()
    const result = insertDocxIndex(editor, {
      afterBlockIndex,
      label: t('ribbonIndexFieldLabel'),
      terms,
    })
    if (!result.ok) window.alert(result.message)
  }

  return (
    <>
      <div className="ribbon-group">
        <div className="ribbon-group-items">
          <button
            className="rb-big"
            disabled={!hasDoc}
            title={t('ribbonTocTip')}
            onClick={insertToc}
          >
            <span className="rb-big-icon">
              <IconToc size={BIG} />
            </span>
            <span>{t('ribbonToc')}</span>
          </button>
          <button
            className="rb-big"
            disabled={!hasDoc}
            title={t('ribbonTocUpdateTip')}
            onClick={() => updateTocField(editor, blocks, headingPages)}
          >
            <span className="rb-big-icon">
              <IconRefresh size={BIG} />
            </span>
            <span>{t('ribbonTocUpdate')}</span>
          </button>
        </div>
        <div className="ribbon-group-label">{t('ribbonToc')}</div>
      </div>

      <div className="ribbon-sep" />

      <div className="ribbon-group">
        <div className="ribbon-group-items">
          <button
            className="rb-big"
            disabled={!hasDoc}
            title={t('ribbonFootnoteTip')}
            onClick={() => onInsertNote('footnote')}
          >
            <span className="rb-big-icon">
              <IconFootnote size={BIG} />
            </span>
            <span>{t('ribbonFootnote')}</span>
          </button>
          <button
            className="rb-big"
            disabled={!hasDoc}
            title={t('ribbonEndnoteTip')}
            onClick={() => onInsertNote('endnote')}
          >
            <span className="rb-big-icon">
              <IconEndnote size={BIG} />
            </span>
            <span>{t('ribbonEndnote')}</span>
          </button>
        </div>
        <div className="ribbon-group-label">{t('ribbonGroupFootnotes')}</div>
      </div>

      <div className="ribbon-sep" />

      <div className="ribbon-group">
        <div className="ribbon-group-items">
          <div className="rb-split-wrap">
            <button
              className="rb-big"
              disabled={!hasDoc}
              title={t('ribbonCitationTip')}
              onClick={() => toggleDropdown(setDropdown, 'citation')}
            >
              <span className="rb-big-icon">
                <IconCitation size={BIG} />
                <IconCaret />
              </span>
              <span>{t('ribbonCitation')}</span>
            </button>
            {dropdown === 'citation' && (
              <div className="layout-menu">
                {sources.map((s) => (
                  <button key={s.tag} title={s.title} onClick={() => insertCitation(s)}>
                    {citationText(s)} {s.title.slice(0, 12)}
                  </button>
                ))}
                <button
                  onClick={() => {
                    setSourceOpen(true)
                    setDropdown(() => null)
                  }}
                >
                  {t('ribbonAddNewSource')}
                </button>
              </div>
            )}
          </div>
          <button
            className="rb-big"
            disabled={!hasDoc}
            title={t('ribbonBibliographyTip')}
            onClick={insertBibliography}
          >
            <span className="rb-big-icon">
              <IconBook size={BIG} />
            </span>
            <span>{t('ribbonBibliography')}</span>
          </button>
          <button
            className="rb-big"
            disabled={!hasDoc}
            title={t('ribbonCaptionTip')}
            onClick={() => setCaptionOpen(true)}
          >
            <span className="rb-big-icon">
              <IconCaption size={BIG} />
            </span>
            <span>{t('ribbonCaption')}</span>
          </button>
          <div className="rb-split-wrap">
            <button
              className="rb-big"
              disabled={!hasDoc}
              title={t('ribbonIndexTip')}
              onClick={() => toggleDropdown(setDropdown, 'index')}
            >
              <span className="rb-big-icon">
                <IconIndex size={BIG} />
                <IconCaret />
              </span>
              <span>{t('ribbonIndex')}</span>
            </button>
            {dropdown === 'index' && (
              <div className="layout-menu">
                <button
                  onClick={() => {
                    markIndexEntry()
                    setDropdown(() => null)
                  }}
                >
                  {t('ribbonMarkEntry')}
                </button>
                <button
                  onClick={() => {
                    insertIndex()
                    setDropdown(() => null)
                  }}
                >
                  {t('ribbonInsertIndex')}
                </button>
              </div>
            )}
          </div>
        </div>
        <div className="ribbon-group-label">{t('ribbonGroupCitationsIndex')}</div>
      </div>

      {captionOpen && (
        <CaptionModal editor={editor} blocks={blocks} onClose={() => setCaptionOpen(false)} />
      )}
      {sourceOpen && (
        <SourceModal sources={sources} onAdd={onAddSource} onClose={() => setSourceOpen(false)} />
      )}
      {xePrompt && (
        <PromptModal
          title={t('ribbonMarkEntryTitle')}
          placeholder={t('ribbonMarkEntryPh')}
          initial={xePrompt.initial}
          onSubmit={submitIndexEntry}
          onClose={() => setXePrompt(null)}
        />
      )}
    </>
  )
}

/* ================= Review ================= */
