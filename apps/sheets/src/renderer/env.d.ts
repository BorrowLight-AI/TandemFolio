declare module '*.md?raw' {
  const content: string
  export default content
}

import type {
  LocalImageRequest,
  LocalImageResult,
  ScreenCaptureRequest,
  ScreenCaptureResult,
  ScreenSourcesResult,
  WorkbookFormulaCellsRequest,
  WorkbookFormulaCellsResult,
  WorkbookFile,
  WorkbookExportPdfRequest,
  WorkbookExportPdfResult,
  WorkbookMediaRequest,
  WorkbookMediaResult,
  WorkbookPivotDefinition,
  WorkbookPivotRequest,
  WorkbookRangeRequest,
  WorkbookRangeResult,
  WorkbookRecalcRequest,
  WorkbookRecalcResult,
  WorkbookSaveRequest,
  WorkbookSaveResult,
  MenuAction,
} from '../shared/desktop-api'

interface BrowserWorkbookHostApi {
  selectWorkbook(): Promise<WorkbookFile | null>
  readWorkbookRange(request: WorkbookRangeRequest): Promise<WorkbookRangeResult>
  readWorkbookFormulas(request: WorkbookFormulaCellsRequest): Promise<WorkbookFormulaCellsResult>
  recalcWorkbook(request: WorkbookRecalcRequest): Promise<WorkbookRecalcResult>
  readWorkbookMedia(request: WorkbookMediaRequest): Promise<WorkbookMediaResult>
  readPivotDefinition(request: WorkbookPivotRequest): Promise<WorkbookPivotDefinition>
  readLocalImage(request: LocalImageRequest): Promise<LocalImageResult>
  captureScreenSources(): Promise<ScreenSourcesResult>
  captureScreenSource(request: ScreenCaptureRequest): Promise<ScreenCaptureResult | null>
  saveWorkbookEdits(request: WorkbookSaveRequest): Promise<WorkbookSaveResult>
  writeWorkbookRecovery(request: WorkbookSaveRequest): Promise<
    | { ok: true; fileName: string; data: ArrayBuffer }
    | { ok: false }
  >
  exportPdf(request: WorkbookExportPdfRequest): Promise<WorkbookExportPdfResult>
  closeWorkbook(sessionId: string): Promise<void>
  openExternal(url: string): Promise<void>
  onThemeChanged(handler: (theme: 'light' | 'dark' | 'system') => void): () => void
  onMenuAction(handler: (action: MenuAction) => void): () => void
  notifyPendingEdits(count: number): void
  onCloseSaveRequest(handler: () => void): () => void
  reportCloseSaveResult(ok: boolean): void
  consumeNewBlankWorkbook(): Promise<boolean>
  hasQueuedWorkbook(): Promise<boolean>
  getPathForFile(file: File): string
}

declare global {
  interface Window {
    readonly desktopApi: BrowserWorkbookHostApi
    __genofficeXlsxEntryModuleReadyAt?: number
  }
}

export {}
