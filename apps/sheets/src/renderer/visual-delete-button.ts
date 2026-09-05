/**
 * Derived from genspark-ai/genoffice at
 * 360ce0625eaf748368e5535984b073f6fb2487b5 and adapted for TandemFolio's
 * mounted browser renderer.
 */
import { BooleanNumber } from '@univerjs/core'

export interface BoxRect {
  readonly left: number
  readonly top: number
  readonly right: number
  readonly bottom: number
}

export const VISUAL_DELETE_BUTTON_SIZE = 20
export const VISUAL_DELETE_BUTTON_GAP = 4

export function shouldShowVisualDeleteButton(state: {
  readonly selected: boolean
  readonly textEditing: boolean
}): boolean {
  return state.selected && !state.textEditing
}

export type VisualDeleteButtonPlacement = 'above' | 'below' | 'inside'

export interface VisualDeleteButtonPosition {
  readonly left: number
  readonly top: number
  readonly placement: VisualDeleteButtonPlacement
}

export function visualDeleteButtonPosition(
  frame: BoxRect,
  bounds: BoxRect,
  size = VISUAL_DELETE_BUTTON_SIZE,
  gap = VISUAL_DELETE_BUTTON_GAP,
): VisualDeleteButtonPosition {
  const above = frame.top - gap - size
  const below = frame.bottom + gap
  let placement: VisualDeleteButtonPlacement
  let top: number
  if (above >= bounds.top) {
    placement = 'above'
    top = above
  } else if (below + size <= bounds.bottom) {
    placement = 'below'
    top = below
  } else {
    placement = 'inside'
    top = Math.max(frame.top, bounds.top) + gap
  }
  const maxLeft = Math.max(bounds.left, bounds.right - size)
  const left = Math.min(Math.max(frame.right - size, bounds.left), maxLeft)
  return { left, top, placement }
}

export interface GridInsetSource {
  getZoom(): number
  getFreeze(): { readonly ySplit: number; readonly startRow: number }
  getRowHeight(row: number): number
  getSheet(): {
    getConfig(): {
      readonly columnHeader: { readonly height: number; readonly hidden?: BooleanNumber }
    }
    getRowVisible(row: number): boolean
  }
}

export function gridTopInset(worksheet: GridInsetSource): number {
  const zoom = worksheet.getZoom() || 1
  const sheet = worksheet.getSheet()
  const header = sheet.getConfig().columnHeader
  let inset = header.hidden === BooleanNumber.TRUE ? 0 : header.height
  const freeze = worksheet.getFreeze()
  for (let row = freeze.startRow - freeze.ySplit; row < freeze.startRow; row += 1) {
    if (row >= 0 && sheet.getRowVisible(row)) inset += worksheet.getRowHeight(row)
  }
  return inset * zoom
}
