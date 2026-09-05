import { useEffect, useLayoutEffect, useRef, useState } from 'react'
import { createPortal } from 'react-dom'
import { ColorPicker } from '@genoffice/ui'

import { t } from './i18n/locale'

function PortalPop({
  anchor,
  popRef,
  children,
}: {
  readonly anchor: HTMLElement
  readonly popRef: React.RefObject<HTMLDivElement | null>
  readonly children: React.ReactNode
}): React.JSX.Element {
  useLayoutEffect(() => {
    const panel = popRef.current
    if (!panel) return
    const position = (): void => {
      const rect = anchor.getBoundingClientRect()
      const left = Math.min(Math.max(8, rect.left), window.innerWidth - panel.offsetWidth - 8)
      const below = rect.bottom + 4
      const top =
        below + panel.offsetHeight <= window.innerHeight - 8
          ? below
          : Math.max(8, rect.top - panel.offsetHeight - 4)
      panel.style.left = `${left}px`
      panel.style.top = `${top}px`
    }
    position()
    window.addEventListener('resize', position)
    window.addEventListener('scroll', position, true)
    return () => {
      window.removeEventListener('resize', position)
      window.removeEventListener('scroll', position, true)
    }
  }, [anchor, popRef])

  return createPortal(
    <div ref={popRef} className="sheets-color-pop-portal">
      {children}
    </div>,
    document.body,
  )
}

export function ColorDropdown({
  label,
  'data-tip': tip,
  display,
  value,
  auto,
  disabled,
  portal,
  onPick,
}: {
  readonly label: string
  readonly 'data-tip'?: string
  readonly display?: React.ReactNode
  readonly value: string
  readonly auto?: string
  readonly disabled?: boolean
  readonly portal?: boolean
  readonly onPick: (hex: string | null) => void
}): React.JSX.Element {
  const [open, setOpen] = useState(false)
  const wrapRef = useRef<HTMLDivElement>(null)
  const popRef = useRef<HTMLDivElement>(null)

  useEffect(() => {
    if (!open) return
    const onDown = (event: PointerEvent): void => {
      const target = event.target as Node
      if (!wrapRef.current?.contains(target) && !popRef.current?.contains(target)) setOpen(false)
    }
    const onKey = (event: KeyboardEvent): void => {
      if (event.key === 'Escape') setOpen(false)
    }
    window.addEventListener('pointerdown', onDown, true)
    window.addEventListener('keydown', onKey)
    return () => {
      window.removeEventListener('pointerdown', onDown, true)
      window.removeEventListener('keydown', onKey)
    }
  }, [open])

  const picker = open ? (
    <ColorPicker
      className={portal ? undefined : 'sheets-color-pop'}
      value={value}
      strings={{
        auto,
        themeColors: t('appThemeColors'),
        standardColors: t('appStandardColors'),
        moreColors: t('appMoreColors'),
      }}
      onPick={(hex) => {
        onPick(hex ? hex.toLowerCase() : null)
        setOpen(false)
      }}
      moreInputProps={{
        onChange: (event) => onPick(event.currentTarget.value.toLowerCase()),
      }}
    />
  ) : null

  return (
    <div ref={wrapRef} className="menu-select">
      <button
        type="button"
        className={display ? 'color-tool' : 'color-well'}
        data-tip={tip}
        aria-label={label}
        aria-expanded={open}
        disabled={disabled}
        style={display ? undefined : { background: value }}
        onClick={() => setOpen((current) => !current)}
      >
        {display}
      </button>
      {picker &&
        (portal && wrapRef.current ? (
          <PortalPop anchor={wrapRef.current} popRef={popRef}>
            {picker}
          </PortalPop>
        ) : (
          picker
        ))}
    </div>
  )
}
