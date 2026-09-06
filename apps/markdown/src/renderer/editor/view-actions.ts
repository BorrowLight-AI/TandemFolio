export const MARKDOWN_MIN_ZOOM = 50
export const MARKDOWN_MAX_ZOOM = 200
export const MARKDOWN_ZOOM_STEP = 10

export function normalizeMarkdownZoom(percent: number): number {
  if (!Number.isFinite(percent)) return 100
  return Math.min(MARKDOWN_MAX_ZOOM, Math.max(MARKDOWN_MIN_ZOOM, Math.round(percent)))
}
