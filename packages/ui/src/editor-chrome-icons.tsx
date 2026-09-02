import type { ReactNode } from 'react'

export interface EditorChromeIconProps {
  size?: number
}

function pinnedStroke(size: number): number {
  const painted = size >= 20 ? 1.5 : size >= 13 ? 1.25 : 1.1
  return (painted * 16) / size
}

function EditorChromeIcon({
  size = 16,
  children,
}: EditorChromeIconProps & { children: ReactNode }) {
  return (
    <svg
      width={size}
      height={size}
      viewBox="0 0 16 16"
      fill="none"
      stroke="currentColor"
      strokeWidth={pinnedStroke(size)}
      strokeLinecap="round"
      strokeLinejoin="round"
      aria-hidden="true"
      focusable="false"
    >
      {children}
    </svg>
  )
}

export function EditorFileIcon(props: EditorChromeIconProps) {
  return (
    <EditorChromeIcon {...props}>
      <path d="M1.8 5.2h4l1.4 1.5h7l-1.3 6.4H3.1L1.8 5.2Z" />
      <path d="M2.7 5.2V3.4h4.1l1.3 1.4h4.5v1.9" />
    </EditorChromeIcon>
  )
}

export function EditorSaveIcon(props: EditorChromeIconProps) {
  return (
    <EditorChromeIcon {...props}>
      <path d="M2.5 2.2h8.7l2.3 2.3v9.3h-11V2.2Z" />
      <path d="M5 2.2v3.2h5.4V2.2M5 13.8V9h6v4.8" />
    </EditorChromeIcon>
  )
}

export function EditorFullscreenIcon({
  exit = false,
  ...props
}: EditorChromeIconProps & { exit?: boolean }) {
  return (
    <EditorChromeIcon {...props}>
      {exit ? (
        <path d="M6.3 1.8v4.5H1.8M9.7 1.8v4.5h4.5M6.3 14.2V9.7H1.8M9.7 14.2V9.7h4.5" />
      ) : (
        <path d="M6.3 1.8H1.8v4.5M9.7 1.8h4.5v4.5M6.3 14.2H1.8V9.7M9.7 14.2h4.5V9.7" />
      )}
    </EditorChromeIcon>
  )
}
