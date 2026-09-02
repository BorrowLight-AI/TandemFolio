export type HostVisualFormatId = 'docx' | 'markdown' | 'xlsx' | 'pptx' | 'pdf'

export interface ActiveHostVisualFormat {
  status: 'active'
  id: HostVisualFormatId
  label: string
  documentSelector: string
  viewportSelector: string
  statusSelector: string
  fullscreenButtonSelector: string
}

export interface PendingHostVisualFormat {
  status: 'awaiting-community-renderer'
  id: HostVisualFormatId
  label: string
  reason: string
}

export const hostVisualFormatGates: readonly (ActiveHostVisualFormat | PendingHostVisualFormat)[] =
  [
    {
      status: 'active',
      id: 'docx',
      label: 'DOCX',
      documentSelector: '.app',
      viewportSelector: '.editor-scroll',
      statusSelector: '.status-bar',
      fullscreenButtonSelector: '.mcp-fullscreen-button',
    },
    {
      status: 'active',
      id: 'markdown',
      label: 'Markdown',
      documentSelector: '.doc-editor[aria-label="Markdown document"]',
      viewportSelector: '.editor-scroll',
      statusSelector: '.status-bar',
      fullscreenButtonSelector: 'button[aria-pressed]',
    },
    {
      status: 'active',
      id: 'xlsx',
      label: 'XLSX',
      documentSelector: '.app-shell',
      viewportSelector: 'canvas[id^="univer-sheet-main-canvas_"]',
      statusSelector: '.status-bar',
      fullscreenButtonSelector: '.host-fullscreen[aria-pressed]',
    },
    {
      status: 'active',
      id: 'pptx',
      label: 'PPTX',
      documentSelector: '.app',
      viewportSelector: '.stage-wrap',
      statusSelector: '.status-bar',
      fullscreenButtonSelector: '.host-fullscreen[aria-pressed]',
    },
    {
      status: 'active',
      id: 'pdf',
      label: 'PDF',
      documentSelector: '.app',
      viewportSelector: '.pdf-page-content canvas',
      statusSelector: '.status-bar',
      fullscreenButtonSelector: '.host-fullscreen[aria-pressed]',
    },
  ]

export const activeHostVisualFormats = hostVisualFormatGates.filter(
  (format): format is ActiveHostVisualFormat => format.status === 'active',
)

export const pendingHostVisualFormats = hostVisualFormatGates.filter(
  (format): format is PendingHostVisualFormat => format.status === 'awaiting-community-renderer',
)
