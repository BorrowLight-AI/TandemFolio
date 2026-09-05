export interface PrintAreaHeights {
  readonly repeatedHeightPt: number
  readonly rowHeightsPt: readonly number[]
}

export interface FitToPageInput {
  readonly printableWidthPt: number
  readonly printableHeightPt: number
  readonly fitToWidth: number
  readonly fitToHeight: number
  readonly contentWidthPt: number
  readonly areas: readonly PrintAreaHeights[]
}

export const MIN_PRINT_SCALE = 0.1
export const MAX_PRINT_SCALE = 2
const SCALE_STEP = 0.005

export function fitToPageScale(input: FitToPageInput): number {
  let scale = 1
  if (input.fitToWidth > 0 && input.contentWidthPt > 0 && input.printableWidthPt > 0) {
    scale = Math.min(scale, (input.printableWidthPt * input.fitToWidth) / input.contentWidthPt)
  }
  if (input.fitToHeight > 0 && input.printableHeightPt > 0) {
    const areas = input.areas.filter(
      (area) => area.repeatedHeightPt + sumHeights(area.rowHeightsPt) > 0,
    )
    for (const area of areas) {
      const total = area.repeatedHeightPt + sumHeights(area.rowHeightsPt)
      scale = Math.min(scale, (input.printableHeightPt * input.fitToHeight) / total)
    }
    scale = Math.max(scale, MIN_PRINT_SCALE)
    while (
      scale > MIN_PRINT_SCALE &&
      areas.some(
        (area) => countPages([area], input.printableHeightPt / scale) > input.fitToHeight,
      )
    ) {
      scale = Math.max(MIN_PRINT_SCALE, scale - SCALE_STEP)
    }
  }
  return Math.min(1, Math.max(MIN_PRINT_SCALE, scale))
}

export function countPages(areas: readonly PrintAreaHeights[], capacityPt: number): number {
  let pages = 0
  for (const area of areas) {
    pages += 1
    let used = area.repeatedHeightPt
    for (const row of area.rowHeightsPt) {
      if (used + row > capacityPt && used > area.repeatedHeightPt) {
        pages += 1
        used = area.repeatedHeightPt
      }
      used += row
    }
  }
  return pages
}

function sumHeights(heights: readonly number[]): number {
  return heights.reduce((total, height) => total + height, 0)
}
