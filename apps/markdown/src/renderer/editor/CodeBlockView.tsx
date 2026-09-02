import { useState } from 'react'
import { NodeViewContent, NodeViewWrapper } from '@tiptap/react'
import type { NodeViewProps } from '@tiptap/react'
import { t } from '../i18n/locale'
import {
  markdownCodeBlockLanguages,
  setMarkdownCodeBlockLanguage,
  type MarkdownCodeBlockLanguage,
} from './code-block-actions'
import { getMarkdownTextBlockIndexAtPosition } from './block-type-actions'

export function CodeBlockView({ node, getPos, editor }: NodeViewProps) {
  const [copied, setCopied] = useState(false)
  const language = String(node.attrs.language ?? '') || 'plaintext'

  const copy = () => {
    void navigator.clipboard.writeText(node.textContent).then(() => {
      setCopied(true)
      setTimeout(() => setCopied(false), 1500)
    })
  }

  return (
    <NodeViewWrapper className="md-codeblock">
      <div className="md-codeblock-bar" contentEditable={false}>
        <select
          className="md-codeblock-lang"
          value={
            markdownCodeBlockLanguages.includes(language as MarkdownCodeBlockLanguage)
              ? language
              : 'plaintext'
          }
          disabled={!editor.isEditable}
          onChange={(event) => {
            const position = getPos()
            if (typeof position !== 'number') return
            const textBlockIndex = getMarkdownTextBlockIndexAtPosition(editor, position)
            if (textBlockIndex === null) return
            setMarkdownCodeBlockLanguage(editor, {
              textBlockIndex,
              language: event.target.value as MarkdownCodeBlockLanguage,
            })
          }}
        >
          {markdownCodeBlockLanguages.map((lang) => (
            <option key={lang} value={lang}>
              {lang}
            </option>
          ))}
        </select>
        <button type="button" className="md-codeblock-copy" onClick={copy}>
          {copied ? t('codeCopied') : t('codeCopy')}
        </button>
      </div>
      <pre>
        <NodeViewContent<'code'> as="code" />
      </pre>
    </NodeViewWrapper>
  )
}
