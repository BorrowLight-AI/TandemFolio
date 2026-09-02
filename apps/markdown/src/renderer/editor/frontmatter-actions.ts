import { buildFrontmatterRaw, type DocEnvelope } from '../markdown/docText'

export interface MarkdownFrontmatterResult {
  readonly changed: boolean
  readonly raw: string
}

/** Replace the complete YAML envelope while preserving body/EOL/BOM/EOF state. */
export function setMarkdownFrontmatter(
  envelope: DocEnvelope,
  yaml: string,
): MarkdownFrontmatterResult {
  const raw = buildFrontmatterRaw(yaml)
  const changed = envelope.frontmatter !== raw
  envelope.frontmatter = raw
  return { changed, raw }
}
