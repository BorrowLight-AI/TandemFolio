/**
 * Derived from genspark-ai/genoffice at
 * 360ce0625eaf748368e5535984b073f6fb2487b5 and adapted for TandemFolio's
 * browser-only visual host.
 */
import { useEffect, useRef, type RefObject } from 'react'
import { createPortal } from 'react-dom'

import {
  type BoxRect,
  type GridInsetSource,
  VISUAL_DELETE_BUTTON_GAP,
  VISUAL_DELETE_BUTTON_SIZE,
  gridTopInset,
  visualDeleteButtonPosition,
} from './visual-delete-button'

function floatDomWrapperOf(host: HTMLElement): HTMLElement | null {
  for (let node = host.parentElement; node; node = node.parentElement) {
    if (node.style.overflow === 'hidden') return node
  }
  return null
}

function gridBounds(grid: HTMLElement | null, topInset: number): BoxRect {
  const rect = grid?.getBoundingClientRect()
  return rect
    ? { left: rect.left, top: rect.top + topInset, right: rect.right, bottom: rect.bottom }
    : { left: 0, top: topInset, right: window.innerWidth, bottom: window.innerHeight }
}

export function VisualDeleteButton({
  hostRef,
  worksheet,
  label,
  onDelete,
}: {
  readonly hostRef: RefObject<HTMLElement | null>
  readonly worksheet: GridInsetSource
  readonly label: string
  readonly onDelete: () => void
}): React.JSX.Element {
  const buttonRef = useRef<HTMLButtonElement>(null)

  useEffect(() => {
    const host = hostRef.current
    const button = buttonRef.current
    if (!host || !button) return
    const wrapper = floatDomWrapperOf(host)
    const grid = wrapper?.parentElement ?? null
    const measure = (): void => {
      const frame = (wrapper ?? host).getBoundingClientRect()
      if (frame.width <= 0 || frame.height <= 0) {
        button.dataset['ready'] = 'false'
        return
      }
      const position = visualDeleteButtonPosition(
        frame,
        gridBounds(grid, gridTopInset(worksheet)),
        VISUAL_DELETE_BUTTON_SIZE,
        VISUAL_DELETE_BUTTON_GAP,
      )
      button.style.left = `${position.left}px`
      button.style.top = `${position.top}px`
      button.dataset['placement'] = position.placement
      button.dataset['ready'] = 'true'
    }
    measure()
    const wrapperMoves = new MutationObserver(measure)
    if (wrapper) wrapperMoves.observe(wrapper, { attributes: true, attributeFilter: ['style'] })
    const resizes = new ResizeObserver(measure)
    resizes.observe(host)
    if (grid) resizes.observe(grid)
    window.addEventListener('resize', measure)
    return () => {
      wrapperMoves.disconnect()
      resizes.disconnect()
      window.removeEventListener('resize', measure)
    }
  }, [hostRef, worksheet])

  return createPortal(
    <button
      ref={buttonRef}
      type="button"
      className="shape-delete-button"
      tabIndex={-1}
      data-tip={label}
      aria-label={label}
      onMouseDown={(event) => event.preventDefault()}
      onPointerDown={(event) => event.stopPropagation()}
      onClick={(event) => {
        event.stopPropagation()
        onDelete()
      }}
    >
      ✕
    </button>,
    document.body,
  )
}
