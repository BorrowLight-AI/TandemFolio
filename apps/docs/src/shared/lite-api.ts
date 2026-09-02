import type { FaceVerticalMetrics } from '@genoffice/font-metrics'

export interface OpenFileResult {
  path: string
  name: string
  data: ArrayBuffer
  hash: string
}

export interface PickImageResult {
  base64: string
  mime: 'image/png' | 'image/jpeg' | 'image/gif'
  name: string
}

export interface DocsTabInfo {
  id: string
  title: string
  focused: boolean
}

export type MenuCommand =
  | 'new'
  | 'open'
  | 'open-path'
  | 'save'
  | 'save-as'
  | 'undo'
  | 'redo'
  | 'zoom-in'
  | 'zoom-out'
  | 'zoom-100'
  | 'zoom-page-width'
  | 'zoom-whole-page'
  | 'toggle-dark'
  | 'insert-table'
  | 'insert-image'
  | 'insert-page-break'
  | 'insert-link'
  | 'insert-equation'
  | 'insert-comment'
  | 'font-dialog'
  | 'paragraph-dialog'
  | 'bold'
  | 'italic'
  | 'underline'
  | 'align-left'
  | 'align-center'
  | 'align-right'
  | 'align-justify'
  | 'page-setup'
  | 'find'
  | 'print'
  | 'export-pdf'
  | 'word-count'

export type UiTheme = 'light' | 'dark' | 'system'

export interface DesktopApi {
  getLanguage(): Promise<'zh' | 'en' | 'ja' | 'ko' | 'fr' | 'de' | 'es' | 'th' | 'id' | 'ru' | 'ar'>
  onLanguageChanged(handler: (lang: Awaited<ReturnType<DesktopApi['getLanguage']>>) => void): () => void
  getTheme(): Promise<UiTheme>
  onThemeChanged(handler: (theme: UiTheme) => void): () => void
  openDocx(): Promise<OpenFileResult | null>
  openDocxPath(path: string): Promise<OpenFileResult | null>
  consumePendingOpenDocx(): Promise<OpenFileResult | null>
  consumeNewBlankDoc(): Promise<boolean>
  onOpenDocx(handler: (result: OpenFileResult) => void): () => void
  onRenamedDocx(handler: (paths: { oldPath: string; newPath: string }) => void): () => void
  saveDocx(path: string, data: ArrayBuffer, auto?: boolean): Promise<{ ok: boolean; error?: string; reason?: 'external-modified' }>
  writeRecoveryCopy(path: string, data: ArrayBuffer): Promise<{ ok: boolean }>
  onTeardown(handler: () => void): () => void
  saveDocxAs(defaultName: string, data: ArrayBuffer): Promise<{ ok: boolean; path?: string; error?: string }>
  saveDocxNew(defaultName: string, data: ArrayBuffer): Promise<{ ok: boolean; path?: string; error?: string }>
  getRecentFiles(): Promise<string[]>
  pickImage(): Promise<PickImageResult | null>
  fontMetrics(family: string): Promise<FaceVerticalMetrics | null>
  print(): Promise<void>
  exportPdf(defaultName: string, pageWidthTwips: number, pageHeightTwips: number, outPath?: string): Promise<{ ok: boolean; path?: string; error?: string }>
  printPdfBuffer(pageWidthTwips: number, pageHeightTwips: number): Promise<{ ok: boolean; base64?: string; error?: string }>
  saveMergedPdf(defaultName: string, base64Parts: string[], outPath?: string): Promise<{ ok: boolean; path?: string; error?: string }>
  openNewTab(openPath?: string | null): Promise<void>
  listDocsTabs(): Promise<DocsTabInfo[]>
  focusDocsTab(id: string): Promise<void>
  onMenuCommand(handler: (command: MenuCommand, payload?: string) => void): () => void
  onCloseCheck(handler: () => void): () => void
  reportCloseCheck(state: { dirty: boolean; autoSave: boolean; filePath?: string | null }): void
  onCloseSaveRequest(handler: () => void): () => void
  reportCloseSaveResult(ok: boolean): void
  reportViewMenuState(state: { darkCanvas: boolean }): void
}
