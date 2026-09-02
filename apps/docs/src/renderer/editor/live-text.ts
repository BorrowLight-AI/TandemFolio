import type { Node as ProseMirrorNode } from '@tiptap/pm/model'

const hasDeletionMark = (node: ProseMirrorNode) =>
  node.marks.some((mark) => mark.type.name === 'del')

/** Text as it reads after pending tracked deletions are applied. */
export function liveText(node: ProseMirrorNode): string {
  if ((node.attrs?.blockRevision as { kind?: string } | null)?.kind === 'del') return ''
  let output = ''
  const walk = (child: ProseMirrorNode): void => {
    if (hasDeletionMark(child)) return
    if (child.isText) output += child.text ?? ''
    else if (child.isLeaf) output += child.type.spec.leafText?.(child) ?? ''
    else child.forEach(walk)
  }
  node.forEach(walk)
  return output
}

/** Whether a complete block is currently represented only by deletion revisions. */
export function isTrackedDeleted(node: ProseMirrorNode): boolean {
  if ((node.attrs?.blockRevision as { kind?: string } | null)?.kind === 'del') return true
  let hasContent = false
  let hasLiveContent = false
  node.descendants((child) => {
    if (child.isText || (child.isInline && child.isLeaf)) {
      hasContent = true
      if (!hasDeletionMark(child)) hasLiveContent = true
    }
  })
  return hasContent && !hasLiveContent
}
