import type { Editor } from '@tiptap/core'

import { resolveMarkdownTextBlockPosition } from './block-type-actions'

export const markdownCodeBlockLanguages = [
  'plaintext',
  'bash',
  'c',
  'cpp',
  'csharp',
  'css',
  'diff',
  'dockerfile',
  'go',
  'graphql',
  'html',
  'java',
  'javascript',
  'json',
  'kotlin',
  'lua',
  'markdown',
  'objectivec',
  'php',
  'python',
  'r',
  'ruby',
  'rust',
  'scala',
  'scss',
  'sql',
  'swift',
  'typescript',
  'xml',
  'yaml',
] as const

export type MarkdownCodeBlockLanguage = (typeof markdownCodeBlockLanguages)[number]

export type MarkdownCodeBlockLanguageResult =
  | { readonly ok: true; readonly changed: boolean }
  | { readonly ok: false; readonly message: string }

/** Set one addressed code block to an explicit language attribute. */
export function setMarkdownCodeBlockLanguage(
  editor: Editor,
  input: { readonly textBlockIndex: number; readonly language: MarkdownCodeBlockLanguage },
): MarkdownCodeBlockLanguageResult {
  const address = resolveMarkdownTextBlockPosition(editor, input.textBlockIndex)
  if (!address.ok) return address
  const node = editor.state.doc.nodeAt(address.position)
  if (!node || node.type.name !== 'codeBlock') {
    return {
      ok: false,
      message: `Markdown text block ${input.textBlockIndex} is not a code block.`,
    }
  }
  const language = input.language === 'plaintext' ? null : input.language
  if ((node.attrs.language ?? null) === language) return { ok: true, changed: false }
  editor.view.dispatch(
    editor.state.tr
      .setNodeMarkup(address.position, undefined, { ...node.attrs, language })
      .scrollIntoView(),
  )
  return { ok: true, changed: true }
}
