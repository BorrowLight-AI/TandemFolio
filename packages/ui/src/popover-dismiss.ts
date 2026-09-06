/**
 * Unified dropdown/popover dismissal (the slides PR #505 model, generalized):
 *
 *  1. a press anywhere outside the popover closes it
 *  2. window blur closes it (app/window switch)
 * While any popover installed here is open, `<html>` carries the
 * `genoffice-popover-open` class for format-neutral styling.
 */
import { useEffect, useRef } from 'react'

/** open-popover refcount driving the html-level drag-region suspension class */
let openPopovers = 0
function bumpOpenPopovers(delta: 1 | -1): void {
  openPopovers = Math.max(0, openPopovers + delta)
  document.documentElement.classList.toggle('genoffice-popover-open', openPopovers > 0)
}

export interface PopoverDismissOptions {
  /**
   * Roots the press may land in without dismissing (the popover panel and the
   * trigger that toggles it). When provided, the outside-press listener runs
   * on capture-phase pointerdown with this containment guard. When omitted,
   * it runs on bubble-phase mousedown and the popover must protect itself
   * with `onMouseDown={(e) => e.stopPropagation()}` (the slides convention).
   */
  inside?: () => ReadonlyArray<Element | null | undefined>
}

/** Install browser-native dismissal listeners; returns a teardown function. */
export function installPopoverDismiss(
  close: () => void,
  options?: PopoverDismissOptions,
): () => void {
  const inside = options?.inside
  const onPress = (e: Event) => {
    if (inside) {
      const target = e.target as Node | null
      if (target) {
        for (const root of inside()) if (root && root.contains(target)) return
      }
    }
    close()
  }
  const onBlur = () => close()
  if (inside) window.addEventListener('pointerdown', onPress, true)
  else window.addEventListener('mousedown', onPress)
  window.addEventListener('blur', onBlur)
  bumpOpenPopovers(1)
  return () => {
    if (inside) window.removeEventListener('pointerdown', onPress, true)
    else window.removeEventListener('mousedown', onPress)
    window.removeEventListener('blur', onBlur)
    bumpOpenPopovers(-1)
  }
}

/** React binding: listeners live only while `open` is true. */
export function useDismissablePopover(
  open: boolean,
  close: () => void,
  options?: PopoverDismissOptions,
): void {
  const closeRef = useRef(close)
  closeRef.current = close
  const insideRef = useRef(options?.inside)
  insideRef.current = options?.inside
  const guarded = options?.inside != null
  useEffect(() => {
    if (!open) return
    return installPopoverDismiss(
      () => closeRef.current(),
      guarded ? { inside: () => insideRef.current?.() ?? [] } : undefined,
    )
  }, [open, guarded])
}
