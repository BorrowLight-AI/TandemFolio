import { formatAddress, parseAddress, parseRange } from '../domain/cell-address'
import { saveLiveEditorFile } from '@tandemfolio/host-bridge'
import type { StructuralOp } from '../gateway/xlsx-structure'
import {
  openBrowserWorkbook,
  type BrowserSheet,
  type BrowserWorkbook,
} from '../host/browser-workbook'
import { workbookPivotDefinitionSchema } from '../shared/desktop-api'
import type {
  LocalImageRequest,
  LocalImageResult,
  MenuAction,
  ScreenCaptureRequest,
  ScreenCaptureResult,
  ScreenSourcesResult,
  WorkbookCellStyle,
  WorkbookExportPdfRequest,
  WorkbookExportPdfResult,
  WorkbookFile,
  WorkbookFormulaCellsRequest,
  WorkbookFormulaCellsResult,
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
} from '../shared/desktop-api'

interface BrowserSession {
  readonly workbook: BrowserWorkbook
  readonly file: WorkbookFile
  readonly sheetNames: Map<string, string>
  readonly handle: FileSystemFileHandle | null
}

const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'

type PickerWindow = Window &
  typeof globalThis & {
    showSaveFilePicker?: (options: Record<string, unknown>) => Promise<FileSystemFileHandle>
  }

const DEFAULT_STYLE: WorkbookCellStyle = {
  fontFamily: 'Calibri',
  fontSize: 11,
  bold: false,
  italic: false,
  underline: false,
  strikethrough: false,
  wrapText: false,
  diagonalUp: false,
  diagonalDown: false,
  numberFormat: 'General',
}

function toHex(bytes: Uint8Array): string {
  return [...bytes].map((byte) => byte.toString(16).padStart(2, '0')).join('')
}

function toBase64(bytes: Uint8Array): string {
  let binary = ''
  for (let offset = 0; offset < bytes.length; offset += 0x8000) {
    binary += String.fromCharCode(...bytes.subarray(offset, offset + 0x8000))
  }
  return btoa(binary)
}

async function sha256(data: ArrayBuffer): Promise<string> {
  return toHex(new Uint8Array(await crypto.subtle.digest('SHA-256', data)))
}

function dimensions(sheet: BrowserSheet): { rowCount: number; columnCount: number } {
  let rowCount = 100
  let columnCount = 26
  for (const cell of sheet.cells.values()) {
    const address = parseAddress(cell.address)
    rowCount = Math.max(rowCount, address.row + 1)
    columnCount = Math.max(columnCount, address.column + 1)
  }
  for (const row of sheet.rows) rowCount = Math.max(rowCount, row.row + 1)
  for (const column of sheet.columnWidths) {
    columnCount = Math.max(columnCount, column.endColumn + 1)
  }
  return { rowCount, columnCount }
}

function styleCatalog(workbook: BrowserWorkbook): WorkbookCellStyle[] {
  let maximum = 0
  for (const sheet of workbook.sheets) {
    for (const cell of sheet.cells.values()) maximum = Math.max(maximum, cell.styleIndex ?? 0)
  }
  return Array.from({ length: maximum + 1 }, () => ({ ...DEFAULT_STYLE }))
}

function xlsxName(name: string): string {
  return name.toLowerCase().endsWith('.xlsx') ? name : `${name}.xlsx`
}

async function pickSaveHandle(name: string): Promise<FileSystemFileHandle> {
  const picker = window as PickerWindow
  if (!picker.showSaveFilePicker) {
    throw new Error('This browser cannot write the workbook to a local file.')
  }
  return picker.showSaveFilePicker({
    suggestedName: xlsxName(name),
    types: [
      {
        description: 'Excel workbook',
        accept: {
          'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'],
        },
      },
    ],
  })
}

async function writeHandle(handle: FileSystemFileHandle, bytes: Uint8Array): Promise<File> {
  const writable = await handle.createWritable()
  const data = bytes.buffer.slice(
    bytes.byteOffset,
    bytes.byteOffset + bytes.byteLength,
  ) as ArrayBuffer
  await writable.write(data)
  await writable.close()
  return handle.getFile()
}

function createMemoryFileHandle(name: string): {
  readonly handle: FileSystemFileHandle
  readonly data: () => ArrayBuffer
} {
  let persisted = new ArrayBuffer(0)
  const handle = {
    kind: 'file' as const,
    name,
    createWritable: async () =>
      ({
        write: async (value: FileSystemWriteChunkType) => {
          if (value instanceof ArrayBuffer) persisted = value.slice(0)
          else if (ArrayBuffer.isView(value)) {
            persisted = value.buffer.slice(
              value.byteOffset,
              value.byteOffset + value.byteLength,
            ) as ArrayBuffer
          } else if (value instanceof Blob) persisted = await value.arrayBuffer()
          else throw new Error('Workbook recovery only accepts binary XLSX data.')
        },
        close: async () => undefined,
      }) as FileSystemWritableFileStream,
    getFile: async () =>
      ({ name, lastModified: Date.now(), arrayBuffer: async () => persisted.slice(0) }) as File,
  } as FileSystemFileHandle
  return { handle, data: () => persisted.slice(0) }
}

export class BrowserWorkbookDesktopApi {
  readonly #sessions = new Map<string, BrowserSession>()
  readonly #pickFile: () => Promise<File | null>

  constructor(pickFile: () => Promise<File | null>) {
    this.#pickFile = pickFile
  }

  async selectWorkbook(): Promise<WorkbookFile | null> {
    const file = await this.#pickFile()
    return file ? this.openBuffer(await file.arrayBuffer(), file.name) : null
  }

  async openBuffer(
    data: ArrayBuffer,
    name: string,
    handle: FileSystemFileHandle | null = null,
  ): Promise<WorkbookFile> {
    const workbook = await openBrowserWorkbook(data, name)
    const sessionId = crypto.randomUUID()
    const sheetNames = new Map<string, string>()
    const sheetIds = new Map<string, string>()
    const sheets = workbook.sheets.map((sheet, index) => {
      const id = `sheet-${index + 1}`
      sheetNames.set(id, sheet.name)
      sheetIds.set(sheet.name, id)
      const { sheetName: _sheetName, tabColor, ...pageSetup } = workbook.pageSetup(sheet.name)
      return {
        id,
        name: sheet.name,
        pageSetup,
        ...dimensions(sheet),
        columnWidths: sheet.columnWidths,
        defaultRowHeight: null,
        defaultColumnWidth: null,
        freeze: null,
        hidden: sheet.hidden,
        tabColor: tabColor ?? null,
        showGridLines: true,
        tables: sheet.tables,
        comments: sheet.comments,
        pivotRanges: sheet.pivotRanges,
        pivotTables: sheet.pivotTables,
        sparklines: sheet.sparklines,
      }
    })
    const file: WorkbookFile = {
      sessionId,
      name,
      sha256: await sha256(data),
      entryCount: 0,
      sheets,
      styles: styleCatalog(workbook),
      dxfStyles: [],
      visuals: workbook.visuals.flatMap(({ sheetName, ...visual }) => {
        const sheetId = sheetIds.get(sheetName)
        return sheetId === undefined ? [] : [{ ...visual, sheetId }]
      }),
      definedNames: workbook.definedNames(),
      readOnly: false,
      ...(handle ? {} : { needsSaveAs: true }),
    }
    this.#sessions.set(sessionId, { workbook, file, sheetNames, handle })
    return file
  }

  async readWorkbookRange(request: WorkbookRangeRequest): Promise<WorkbookRangeResult> {
    const session = this.#session(request.sessionId)
    const sheet = this.#sheet(session, request.sheetId)
    const cells = [...sheet.cells.values()].flatMap((cell) => {
      const address = parseAddress(cell.address)
      const { range } = request
      if (
        address.row < range.startRow ||
        address.row > range.endRow ||
        address.column < range.startColumn ||
        address.column > range.endColumn
      ) {
        return []
      }
      return [
        {
          row: address.row,
          column: address.column,
          value: cell.value,
          ...(cell.formula
            ? { formula: cell.formula.startsWith('=') ? cell.formula : `=${cell.formula}` }
            : {}),
          ...(cell.styleIndex === undefined ? {} : { styleIndex: cell.styleIndex }),
        },
      ]
    })
    const merges = sheet.merges
      .map((range) => parseRange(range))
      .filter(
        (range) =>
          range.endRow >= request.range.startRow &&
          range.startRow <= request.range.endRow &&
          range.endColumn >= request.range.startColumn &&
          range.startColumn <= request.range.endColumn,
      )
    const hyperlinks = [...sheet.hyperlinks].flatMap(([key, target]) => {
      const [row, column] = key.split(':').map(Number)
      if (
        row === undefined ||
        column === undefined ||
        row < request.range.startRow ||
        row > request.range.endRow ||
        column < request.range.startColumn ||
        column > request.range.endColumn
      ) {
        return []
      }
      return [{ row, column, target }]
    })
    const indexedThroughRow = Math.max(0, dimensions(sheet).rowCount - 1)
    return {
      cells,
      rows: sheet.rows.filter(
        (row) => row.row >= request.range.startRow && row.row <= request.range.endRow,
      ),
      merges,
      hyperlinks,
      conditionalRules: [],
      autoFilter: sheet.autoFilterRef ? parseRange(sheet.autoFilterRef) : null,
      dataValidations: sheet.dataValidations,
      sheetProtection: sheet.sheetProtection,
      indexedThroughRow,
      indexingComplete: true,
    }
  }

  async readWorkbookFormulas(
    request: WorkbookFormulaCellsRequest,
  ): Promise<WorkbookFormulaCellsResult> {
    const session = this.#session(request.sessionId)
    const sheet = this.#sheet(session, request.sheetId)
    return {
      cells: [...sheet.cells.values()].flatMap((cell) => {
        if (!cell.formula) return []
        const address = parseAddress(cell.address)
        return [
          { row: address.row, column: address.column, value: cell.value, formula: cell.formula },
        ]
      }),
      indexingComplete: true,
      truncated: false,
    }
  }

  async recalcWorkbook(_request: WorkbookRecalcRequest): Promise<WorkbookRecalcResult> {
    return { cells: [] }
  }

  async saveWorkbookEdits(request: WorkbookSaveRequest): Promise<WorkbookSaveResult> {
    const session = this.#session(request.sessionId)
    const { workbook } = session
    // Match the retained gateway contract: structural operations replay
    // first; journaled edits already use the resulting coordinate space.
    for (const operation of request.structuralOps) {
      const { sheetId, ...structural } = operation
      // Zod materialises optional keys as `undefined`; the package writer uses
      // exact optional properties, so remove those keys at this host boundary.
      const packageOperation = Object.fromEntries(
        Object.entries(structural).filter(([, value]) => value !== undefined),
      ) as StructuralOp
      workbook.applyStructuralOperation(this.#sheetName(session, sheetId), packageOperation)
    }
    for (const state of request.dvStates) {
      workbook.applyDataValidations(this.#sheetName(session, state.sheetId), state.rules)
    }
    for (const state of request.cfStates) {
      workbook.applyConditionalFormats(this.#sheetName(session, state.sheetId), state.rules)
    }
    for (const edit of request.edits) {
      const name = this.#sheetName(session, edit.sheetId)
      const address = formatAddress(edit.row, edit.column)
      if (edit.writeValue) {
        if (edit.formula) workbook.setCellFormula(name, address, edit.formula, edit.value)
        else workbook.setCellValue(name, address, edit.value)
      }
      if (edit.style || edit.styleReset) {
        workbook.setRangeStyle(name, address, edit.styleReset ? {} : (edit.style ?? {}))
      }
    }
    for (const operation of request.sheetOps) {
      if (operation.kind === 'rename-sheet') {
        const previous = this.#sheetName(session, operation.sheetId)
        workbook.renameSheet(previous, operation.newName)
        session.sheetNames.set(operation.sheetId, operation.newName)
      } else if (operation.kind === 'add-sheet') {
        workbook.addSheet(operation.name)
        session.sheetNames.set(operation.sheetId, operation.name)
      } else if (operation.kind === 'duplicate-sheet') {
        workbook.duplicateSheet(this.#sheetName(session, operation.sourceSheetId), operation.name)
        session.sheetNames.set(operation.sheetId, operation.name)
      } else if (operation.kind === 'remove-sheet') {
        workbook.deleteSheet(this.#sheetName(session, operation.sheetId))
        session.sheetNames.delete(operation.sheetId)
      } else if (operation.kind === 'set-sheet-hidden') {
        workbook.setSheetHidden(this.#sheetName(session, operation.sheetId), operation.hidden)
      } else if (operation.kind === 'reorder-sheets') {
        // The final order is applied below.
      }
    }
    request.sheetOrder.forEach((sheetId, index) => {
      const name = session.sheetNames.get(sheetId)
      if (name) workbook.moveSheet(name, index + 1)
    })
    for (const state of request.pageSetupStates) {
      const { sheetId, ...patch } = state
      workbook.applyPageSetup(this.#sheetName(session, sheetId), patch)
    }
    for (const state of request.filterStates) {
      const { sheetId, ...filterState } = state
      workbook.applyFilterState(this.#sheetName(session, sheetId), filterState)
    }
    for (const state of request.sheetProtections) {
      workbook.setSheetProtection(this.#sheetName(session, state.sheetId), state.protected)
    }
    const sparklineAdditions = new Map<
      string,
      Array<{
        type: 'line' | 'column' | 'stacked'
        color?: string
        cells: readonly { cell: string; sourceRef: string }[]
      }>
    >()
    for (const addition of request.sparklineAdditions) {
      const groups = sparklineAdditions.get(addition.sheetId) ?? []
      groups.push({
        type: addition.type,
        ...(addition.color === undefined ? {} : { color: addition.color }),
        cells: addition.cells,
      })
      sparklineAdditions.set(addition.sheetId, groups)
    }
    for (const [sheetId, additions] of sparklineAdditions) {
      workbook.applySparklines(this.#sheetName(session, sheetId), additions)
    }
    await workbook.applyChartEdits(request.chartEdits)
    await workbook.applyVisualEdits(request.visualEdits)
    for (const state of request.noteStates) {
      await workbook.applyNotes(this.#sheetName(session, state.sheetId), state.notes)
    }
    const hyperlinkEdits = new Map<
      string,
      Array<{ row: number; column: number; target: string | null }>
    >()
    for (const edit of request.hyperlinkEdits) {
      const edits = hyperlinkEdits.get(edit.sheetId) ?? []
      edits.push({ row: edit.row, column: edit.column, target: edit.target })
      hyperlinkEdits.set(edit.sheetId, edits)
    }
    for (const [sheetId, edits] of hyperlinkEdits) {
      workbook.applyHyperlinks(this.#sheetName(session, sheetId), edits)
    }
    await workbook.applyTables(
      request.tableAdditions.map((addition) => ({
        sheetName: this.#sheetName(session, addition.sheetId),
        area: addition.area,
        name: addition.name,
        columnNames: addition.columnNames,
        ...(addition.style === undefined ? {} : { style: addition.style }),
        bandedRows: addition.bandedRows,
      })),
    )
    await workbook.applyPivots(
      request.pivotAdditions.map(({ sheetId, sourceSheetId, ...addition }) => ({
        ...addition,
        sheetName: this.#sheetName(session, sheetId),
        sourceSheetName: this.#sheetName(session, sourceSheetId),
      })),
      request.pivotCacheRefreshPaths,
      request.pivotRefreshUpdates.map((update) => ({
        cachePath: update.cachePath,
        sheetName: this.#sheetName(session, update.sheetId),
        newOutputRef: update.newOutputRef,
        ...(update.pivotPath === undefined ? {} : { pivotPath: update.pivotPath }),
        ...(update.memberHiddenItems === undefined
          ? {}
          : { memberHiddenItems: update.memberHiddenItems }),
        ...(update.relayout === undefined
          ? {}
          : {
              relayout: (() => {
                const { sheetId: _sheetId, sourceSheetId, ...relayout } = update.relayout
                return {
                  ...relayout,
                  sourceSheetName: this.#sheetName(session, sourceSheetId),
                }
              })(),
            }),
      })),
    )
    await workbook.applyVisuals(
      request.visualAdditions.map((addition) => ({
        sheetName: this.#sheetName(session, addition.sheetId),
        anchor: addition.anchor,
        ...(addition.chart === undefined ? {} : { chart: addition.chart }),
        ...(addition.shape === undefined ? {} : { shape: addition.shape }),
        ...(addition.image === undefined ? {} : { image: addition.image }),
      })),
    )
    if (request.definedNamesState) {
      workbook.replaceDefinedNames(request.definedNamesState.names)
    }
    const bytes = await workbook.save()
    if ('parent' in window && window.parent !== window) {
      const fileName = xlsxName(session.file.name)
      const data = bytes.buffer.slice(
        bytes.byteOffset,
        bytes.byteOffset + bytes.byteLength,
      ) as ArrayBuffer
      const persisted = await saveLiveEditorFile({
        fileName,
        data,
        mode: request.mode === 'save-as' ? 'save-as' : 'save',
      })
      if (!persisted.ok) return { canceled: true }
      const file = await this.openBuffer(data, fileName)
      return { canceled: false, file, touchedEntries: ['browser-workbook'] }
    }
    let handle = session.handle
    try {
      if (request.mode === 'save-as' || !handle) handle = await pickSaveHandle(session.file.name)
    } catch (error) {
      if (error instanceof DOMException && error.name === 'AbortError') return { canceled: true }
      throw error
    }
    const persisted = await writeHandle(handle, bytes)
    const file = await this.openBuffer(await persisted.arrayBuffer(), persisted.name, handle)
    return { canceled: false, file, touchedEntries: ['browser-workbook'] }
  }

  async writeWorkbookRecovery(
    request: WorkbookSaveRequest,
  ): Promise<
    | { readonly ok: true; readonly fileName: string; readonly data: ArrayBuffer }
    | { readonly ok: false }
  > {
    const session = this.#session(request.sessionId)
    const base = await session.workbook.save()
    const memory = createMemoryFileHandle(xlsxName(session.file.name))
    const recoveryApi = new BrowserWorkbookDesktopApi(async () => null)
    const opened = await recoveryApi.openBuffer(
      base.buffer.slice(base.byteOffset, base.byteOffset + base.byteLength) as ArrayBuffer,
      memory.handle.name,
      memory.handle,
    )
    const saved = await recoveryApi.saveWorkbookEdits({
      ...request,
      sessionId: opened.sessionId,
      mode: 'save',
    })
    if (saved.canceled) return { ok: false }
    return { ok: true, fileName: saved.file.name, data: memory.data() }
  }

  async readWorkbookMedia(request: WorkbookMediaRequest): Promise<WorkbookMediaResult> {
    const session = this.#session(request.sessionId)
    const visual = session.file.visuals.find((candidate) => candidate.id === request.visualId)
    if (visual?.kind !== 'image' || !visual.mediaPath || !visual.mediaType) {
      throw new Error(`Unknown workbook image: ${request.visualId}`)
    }
    return {
      mediaType: visual.mediaType,
      base64: toBase64(await session.workbook.readBinary(visual.mediaPath)),
    }
  }

  async readPivotDefinition(request: WorkbookPivotRequest): Promise<WorkbookPivotDefinition> {
    return workbookPivotDefinitionSchema.parse(
      await this.#session(request.sessionId).workbook.readPivotDefinition(
        request.path,
        request.cachePath,
      ),
    )
  }

  async readLocalImage(_request: LocalImageRequest): Promise<LocalImageResult> {
    throw new Error('Use the browser image picker to insert a local image.')
  }

  async captureScreenSources(): Promise<ScreenSourcesResult> {
    return { status: 'denied', sources: [] }
  }

  async captureScreenSource(_request: ScreenCaptureRequest): Promise<ScreenCaptureResult | null> {
    return null
  }

  async exportPdf(_request: WorkbookExportPdfRequest): Promise<WorkbookExportPdfResult> {
    return { canceled: true }
  }

  async closeWorkbook(sessionId: string): Promise<void> {
    this.#sessions.delete(sessionId)
  }

  async openExternal(url: string): Promise<void> {
    const parsed = new URL(url)
    if (!['http:', 'https:', 'mailto:'].includes(parsed.protocol)) throw new Error('Unsafe URL.')
    window.open(parsed.href, '_blank', 'noopener,noreferrer')
  }

  onThemeChanged(handler: (theme: 'light' | 'dark' | 'system') => void): () => void {
    const media = window.matchMedia('(prefers-color-scheme: dark)')
    const listener = () => handler('system')
    media.addEventListener('change', listener)
    return () => media.removeEventListener('change', listener)
  }

  onMenuAction(handler: (action: MenuAction) => void): () => void {
    const listener = (event: KeyboardEvent) => {
      if (!(event.metaKey || event.ctrlKey)) return
      const key = event.key.toLowerCase()
      const action =
        key === 'o'
          ? 'open'
          : key === 's'
            ? event.shiftKey
              ? 'save-as'
              : 'save'
            : key === 'z'
              ? event.shiftKey
                ? 'redo'
                : 'undo'
              : null
      if (!action) return
      event.preventDefault()
      handler(action)
    }
    window.addEventListener('keydown', listener)
    return () => window.removeEventListener('keydown', listener)
  }

  notifyPendingEdits(_count: number): void {}
  onCloseSaveRequest(_handler: () => void): () => void {
    return () => undefined
  }
  reportCloseSaveResult(_ok: boolean): void {}
  async consumeNewBlankWorkbook(): Promise<boolean> {
    return false
  }
  async hasQueuedWorkbook(): Promise<boolean> {
    return false
  }
  getPathForFile(_file: File): string {
    return ''
  }

  #session(sessionId: string): BrowserSession {
    const session = this.#sessions.get(sessionId)
    if (!session) throw new Error('The browser workbook session is no longer available.')
    return session
  }

  #sheetName(session: BrowserSession, sheetId: string): string {
    const name = session.sheetNames.get(sheetId)
    if (!name) throw new Error(`Unknown worksheet: ${sheetId}`)
    return name
  }

  #sheet(session: BrowserSession, sheetId: string): BrowserSheet {
    const name = this.#sheetName(session, sheetId)
    const sheet = session.workbook.sheets.find((candidate) => candidate.name === name)
    if (!sheet) throw new Error(`Unknown worksheet: ${sheetId}`)
    return sheet
  }
}
