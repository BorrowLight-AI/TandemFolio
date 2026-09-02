type UndoTarget = {
  tagName?: string
  isContentEditable?: boolean
} | null

export function isTextUndoTarget(target: UndoTarget): boolean {
  return (
    !!target &&
    (target.tagName === 'INPUT' ||
      target.tagName === 'TEXTAREA' ||
      target.isContentEditable === true)
  )
}
