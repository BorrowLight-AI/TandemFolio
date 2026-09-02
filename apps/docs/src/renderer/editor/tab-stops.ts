import type { TabStop } from '@genoffice/docx-engine'

export const TAB_STOP_VALUES = ['left', 'center', 'right', 'decimal', 'bar', 'clear'] as const
export const TAB_STOP_LEADERS = [
  'none',
  'dot',
  'hyphen',
  'underscore',
  'heavy',
  'middleDot',
] as const

export function validateParagraphTabStops(value: unknown): string | null {
  if (value === null || value === undefined) return null
  if (!Array.isArray(value) || value.length < 1 || value.length > 64) {
    return 'tabStops must contain 1 through 64 items or be null'
  }
  let previous = -1
  for (const stop of value) {
    if (!stop || typeof stop !== 'object') return 'tabStops entries must be objects'
    const candidate = stop as Record<string, unknown>
    if (
      !Number.isInteger(candidate.pos) ||
      Number(candidate.pos) < 0 ||
      Number(candidate.pos) > 31680
    ) {
      return 'tabStops positions must be integers between 0 and 31680'
    }
    if (Number(candidate.pos) <= previous) {
      return 'tabStops must be strictly ordered by position'
    }
    if (!(TAB_STOP_VALUES as readonly unknown[]).includes(candidate.val)) {
      return 'tabStops contain an unsupported alignment'
    }
    if (
      candidate.leader !== undefined &&
      !(TAB_STOP_LEADERS as readonly unknown[]).includes(candidate.leader)
    ) {
      return 'tabStops contain an unsupported leader'
    }
    previous = Number(candidate.pos)
  }
  return null
}

export function serializeParagraphTabStops(stops: TabStop[] | null): string | null {
  return stops === null ? null : JSON.stringify(stops)
}
