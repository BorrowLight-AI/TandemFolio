import { useEffect, useRef, useState } from 'react'
import { NodeViewContent, NodeViewWrapper } from '@tiptap/react'
import type { NodeViewProps } from '@tiptap/react'
import { Dropdown } from '@genoffice/ui'
import { t } from '../i18n/locale'
import {
  markdownCodeBlockLanguages,
  setMarkdownCodeBlockLanguage,
  type MarkdownCodeBlockLanguage,
} from './code-block-actions'
import { getMarkdownTextBlockIndexAtPosition } from './block-type-actions'

export function CodeBlockView({ node, getPos, editor }: NodeViewProps) {
  const [copied, setCopied] = useState(false)
  const copyTimerRef = useRef<number | null>(null)
  const mountedRef = useRef(true)
  const language = String(node.attrs.language ?? '') || 'plaintext'

  useEffect(() => {
    mountedRef.current = true
    return () => {
      mountedRef.current = false
      if (copyTimerRef.current !== null) window.clearTimeout(copyTimerRef.current)
    }
  }, [])

  const copy = () => {
    void navigator.clipboard
      .writeText(node.textContent)
      .then(() => {
        if (!mountedRef.current) return
        if (copyTimerRef.current !== null) window.clearTimeout(copyTimerRef.current)
        setCopied(true)
        copyTimerRef.current = window.setTimeout(() => {
          copyTimerRef.current = null
          setCopied(false)
        }, 1500)
      })
      .catch(() => {})
  }

  return (
    <NodeViewWrapper className="md-codeblock">
      <div className="md-codeblock-bar" contentEditable={false}>
        <Dropdown
          className="md-codeblock-lang"
          value={
            markdownCodeBlockLanguages.includes(language as MarkdownCodeBlockLanguage)
              ? language
              : 'plaintext'
          }
          disabled={!editor.isEditable}
          options={markdownCodeBlockLanguages.map((lang) => ({ value: lang, label: lang }))}
          onPick={(language) => {
            const position = getPos()
            if (typeof position !== 'number') return
            const textBlockIndex = getMarkdownTextBlockIndexAtPosition(editor, position)
            if (textBlockIndex === null) return
            setMarkdownCodeBlockLanguage(editor, {
              textBlockIndex,
              language: language as MarkdownCodeBlockLanguage,
            })
          }}
        />
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
