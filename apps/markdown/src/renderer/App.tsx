// Modified by TandemFolio contributors: exact-session document replacement and recovery.
import {
  useCallback,
  useEffect,
  useLayoutEffect,
  useMemo,
  useRef,
  useState,
  useSyncExternalStore,
} from 'react'
import { EditorContent, useEditor } from '@tiptap/react'
import type { Editor } from '@tiptap/core'
import {
  attachMcpLiveSession,
  replaceLiveEditorDocument,
  getLiveEditorActivity,
  getLiveEditorDisplayMode,
  readLiveEditorLocalAsset,
  subscribeLiveEditorActivity,
  subscribeLiveEditorDisplayMode,
  toggleLiveEditorFullscreen,
  type LiveEditorExecution,
  type LiveEditorExecutionTrace,
} from '@tandemfolio/host-bridge'
import { useI18n } from './i18n/locale'
import {
  frontmatterInner,
  parseDocText,
  serializeDocText,
  stripLegacyFencedDivs,
  type DocEnvelope,
} from './markdown/docText'
import { buildExtensions } from './editor/extensions'
import { buildSlashItems } from './editor/slashCommand'
import type { SlashController, SlashMenuState } from './editor/slashCommand'
import { Ribbon } from './components/Ribbon'
import { SlashMenu, type SlashMenuHandle } from './components/SlashMenu'
import { TableMenu } from './components/TableMenu'
import { FrontmatterPanel } from './components/FrontmatterPanel'
import { exportDocxBytes, loadMarkdownImageForDocx } from './export/docxExport'
import { buildPrintHtml } from './export/printHtml'
import {
  canOverwriteMarkdownFile,
  canPersistMarkdownCompanions,
  detachMarkdownFileHandle,
  downloadMarkdownExport,
  openMarkdownFile,
  pickImageDataUrl,
  readLoadedMarkdownAsset,
  saveMarkdownFile,
} from './host/browser-files'
import {
  extractMarkdownImageSources,
  prepareMarkdownAssetSave,
  rewriteMarkdownImageSources,
  type MarkdownAssetBytes,
} from './host/asset-save-plan'
import { executeMarkdownOperation } from './operations/registry'
import { createMarkdownSaveQueue } from './host/save-queue'
import { getMarkdownTextBlockIndexAtSelection } from './editor/block-type-actions'
import { insertMarkdownImage } from './editor/image-actions'
import { hydrateMarkdownLocalImages } from './editor/localImage'
import { setMarkdownFrontmatter } from './editor/frontmatter-actions'
import {
  readMarkdownAutoSavePreference,
  setMarkdownAutoSavePreference,
} from './editor/document-preference-actions'
import {
  MARKDOWN_MAX_ZOOM,
  MARKDOWN_MIN_ZOOM,
  MARKDOWN_ZOOM_STEP,
  normalizeMarkdownZoom,
} from './editor/view-actions'

type LoadStatus = 'loading' | 'ready' | 'error'
type SaveState = 'idle' | 'saving' | 'saved' | 'failed'

const EMPTY_ENVELOPE: DocEnvelope = {
  frontmatter: '',
  body: '',
  eol: '\n',
  trailingNewline: true,
  bom: false,
}

function markdownSelection(editor: Editor): Record<string, unknown> {
  const { from, to, empty } = editor.state.selection
  return {
    from,
    to,
    empty,
    text: empty ? '' : editor.state.doc.textBetween(from, to, '\n').slice(0, 4096),
    block: editor.state.selection.$from.parent.type.name,
    textBlockIndex: getMarkdownTextBlockIndexAtSelection(editor),
  }
}

function recoveryBytes(text: string): ArrayBuffer {
  const bytes = new TextEncoder().encode(text)
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer
}

export default function App() {
  const { t } = useI18n()
  const editorActive = useSyncExternalStore(
    subscribeLiveEditorActivity,
    getLiveEditorActivity,
    () => true,
  )
  const [status, setStatus] = useState<LoadStatus>('loading')
  const [fileName, setFileName] = useState<string | null>(null)
  const [dirty, setDirty] = useState(false)
  const [saveState, setSaveState] = useState<SaveState>('idle')
  const [slashState, setSlashState] = useState<SlashMenuState | null>(null)
  const [frontmatterOpen, setFrontmatterOpen] = useState(false)
  const [frontmatterText, setFrontmatterText] = useState('')
  const [autoSave, setAutoSave] = useState(readMarkdownAutoSavePreference)
  const [zoom, setZoom] = useState(100)
  const [loadCommitId, setLoadCommitId] = useState(0)
  const [, refreshDisplayMode] = useState(0)

  const statusRef = useRef<LoadStatus>('loading')
  const fileNameRef = useRef<string | null>(null)
  const dirtyRef = useRef(false)
  const recoveryVersionRef = useRef(0)
  const envelopeRef = useRef<DocEnvelope>({ ...EMPTY_ENVELOPE })
  const editorRef = useRef<Editor | null>(null)
  const slashMenuRef = useRef<SlashMenuHandle>(null)
  const scrollRef = useRef<HTMLDivElement>(null)
  const nextLoadCommitIdRef = useRef(0)
  const loadCommitWaitersRef = useRef(new Map<number, (durationMs: number) => void>())
  const assetBytesRef = useRef(new Map<string, MarkdownAssetBytes>())
  const assetRewritesRef = useRef(new Map<string, string>())

  const setDirtyState = useCallback((next: boolean) => {
    dirtyRef.current = next
    setDirty(next)
    if (next) setSaveState('idle')
  }, [])

  const markDirty = useCallback(() => {
    if (statusRef.current !== 'ready') return
    recoveryVersionRef.current += 1
    setDirtyState(true)
  }, [setDirtyState])

  const setAutoSaveState = useCallback((enabled: boolean) => {
    setMarkdownAutoSavePreference(enabled)
    setAutoSave(enabled)
  }, [])

  const setZoomState = useCallback(({ percent }: { percent: number }): number => {
    const normalized = normalizeMarkdownZoom(percent)
    setZoom(normalized)
    return normalized
  }, [])

  const zoomOut = useCallback(() => {
    setZoom((value) => normalizeMarkdownZoom(value - MARKDOWN_ZOOM_STEP))
  }, [])

  const zoomIn = useCallback(() => {
    setZoom((value) => normalizeMarkdownZoom(value + MARKDOWN_ZOOM_STEP))
  }, [])

  const insertImage = useCallback(() => {
    void pickImageDataUrl().then((image) => {
      const current = editorRef.current
      if (!image || !current) return
      insertMarkdownImage(current, {
        position: current.state.selection.from,
        src: image.src,
        alt: image.alt,
        title: null,
      })
    })
  }, [])

  const applyFrontmatter = useCallback(
    (yaml: string) => {
      setFrontmatterText(yaml)
      setMarkdownFrontmatter(envelopeRef.current, yaml)
      markDirty()
    },
    [markDirty],
  )

  const extensions = useMemo(() => {
    const controller: SlashController = {
      onOpen: setSlashState,
      onUpdate: setSlashState,
      onKeyDown: (event) => slashMenuRef.current?.handleKey(event) ?? false,
      onClose: () => setSlashState(null),
    }
    return buildExtensions({
      slashController: controller,
      slashItems: () => buildSlashItems({ insertImage }),
    })
  }, [insertImage])

  const editor = useEditor({
    extensions,
    content: '',
    autofocus: true,
    editorProps: { attributes: { class: 'doc-editor', 'aria-label': 'Markdown document' } },
    onUpdate: ({ transaction }) => {
      if (!transaction.getMeta('uiOnly')) markDirty()
    },
  })
  editorRef.current = editor
  fileNameRef.current = fileName

  const serializeCurrent = useCallback((): string => {
    const current = editorRef.current
    if (!current) return ''
    return serializeDocText(envelopeRef.current, current.getMarkdown())
  }, [])

  useLayoutEffect(() => {
    const committedAt = performance.now()
    for (const [commitId, resolve] of loadCommitWaitersRef.current) {
      if (commitId > loadCommitId) continue
      loadCommitWaitersRef.current.delete(commitId)
      resolve(committedAt)
    }
  }, [loadCommitId])

  useEffect(
    () => () => {
      const unmountedAt = performance.now()
      for (const resolve of loadCommitWaitersRef.current.values()) resolve(unmountedAt)
      loadCommitWaitersRef.current.clear()
    },
    [],
  )

  const awaitReactCommit = useCallback((): Promise<number> => {
    const commitId = ++nextLoadCommitIdRef.current
    const startedAt = performance.now()
    return new Promise((resolve) => {
      loadCommitWaitersRef.current.set(commitId, (committedAt) => resolve(committedAt - startedAt))
      setLoadCommitId(commitId)
    })
  }, [])

  const loadText = useCallback(
    async (
      name: string | null,
      raw: string,
      fromLocalHandle: boolean,
    ): Promise<Omit<LiveEditorExecutionTrace['phases'], 'decodeMs'>> => {
      const current = editorRef.current
      if (!current) throw new Error('The Markdown editor is not ready.')
      statusRef.current = 'loading'
      setStatus('loading')
      assetBytesRef.current.clear()
      assetRewritesRef.current.clear()

      const parseStartedAt = performance.now()
      const envelope = parseDocText(raw)
      const markdown = stripLegacyFencedDivs(envelope.body)
      if (!current.markdown) throw new Error('The Markdown parser is not ready.')
      const content = current.markdown.parse(markdown)
      const parseMs = performance.now() - parseStartedAt

      envelopeRef.current = envelope
      const stateInstallStartedAt = performance.now()
      current.chain().setMeta('addToHistory', false).setContent(content).run()
      const tiptapStateInstallMs = performance.now() - stateInstallStartedAt

      setFrontmatterText(frontmatterInner(envelope.frontmatter))
      setFrontmatterOpen(Boolean(envelope.frontmatter))
      if (!fromLocalHandle) detachMarkdownFileHandle()
      fileNameRef.current = name
      setFileName(name)
      setDirtyState(false)
      setSaveState('idle')
      statusRef.current = 'ready'
      setStatus('ready')
      const reactCommitMs = await awaitReactCommit()
      return { parseMs, tiptapStateInstallMs, reactCommitMs }
    },
    [awaitReactCommit, setDirtyState],
  )

  const hydrateAssets = useCallback(
    async (
      current: Editor,
      text: string,
      rootId: string,
      readAsset: (
        rootId: string,
        path: string,
      ) => Promise<{
        readonly mime: 'image/png' | 'image/jpeg' | 'image/gif'
        readonly data: ArrayBuffer
      } | null>,
    ): Promise<void> => {
      const readAndCache = async (assetRootId: string, source: string) => {
        const cached = assetBytesRef.current.get(source)
        if (cached) return { mime: cached.mime, data: cached.data }
        const asset = await readAsset(assetRootId, source).catch(() => null)
        if (asset) assetBytesRef.current.set(source, { source, ...asset })
        return asset
      }
      await hydrateMarkdownLocalImages(current, rootId, readAndCache)
      await Promise.all(
        [...new Set(extractMarkdownImageSources(text))].map((source) =>
          readAndCache(rootId, source),
        ),
      )
    },
    [],
  )

  useEffect(() => {
    if (!editor) return
    envelopeRef.current = { ...EMPTY_ENVELOPE }
    statusRef.current = 'ready'
    setStatus('ready')
  }, [editor])

  const performSave = useCallback(
    async (saveAs = false): Promise<{ ok: true; fileName: string } | { ok: false }> => {
      if (!editorRef.current || statusRef.current !== 'ready') {
        return { ok: false }
      }
      setSaveState('saving')
      try {
        const serializedText = serializeCurrent()
        const originalText = rewriteMarkdownImageSources(serializedText, assetRewritesRef.current)
        const suggestedName = fileNameRef.current ?? 'Untitled.md'
        const prepared = canPersistMarkdownCompanions(saveAs)
          ? await prepareMarkdownAssetSave(originalText, [...assetBytesRef.current.values()], {
              copyRelative: saveAs,
            })
          : { text: originalText, companionFiles: [], rewrites: new Map<string, string>() }
        const result = await saveMarkdownFile(
          prepared.text,
          suggestedName,
          saveAs,
          prepared.companionFiles,
        )
        if (!result.ok) {
          setSaveState('idle')
          return { ok: false }
        }
        for (const [source, target] of prepared.rewrites) {
          assetRewritesRef.current.set(source, target)
          const asset = assetBytesRef.current.get(source)
          if (asset) assetBytesRef.current.set(target, { ...asset, source: target })
        }
        fileNameRef.current = result.fileName
        setFileName(result.fileName)
        setDirtyState(false)
        setSaveState('saved')
        return { ok: true, fileName: result.fileName }
      } catch (error) {
        console.error('[markdown] save failed:', error)
        setSaveState('failed')
        return { ok: false }
      }
    },
    [serializeCurrent, setDirtyState],
  )

  const saveQueue = useMemo(() => createMarkdownSaveQueue(performSave), [performSave])
  const doSave = useCallback((saveAs = false) => saveQueue(saveAs), [saveQueue])

  const openFile = useCallback(() => {
    void openMarkdownFile()
      .then(async (loaded) => {
        if (!loaded) return
        await replaceLiveEditorDocument(async () => {
          await loadText(loaded.fileName, loaded.text, true)
        })
        if (loaded.assetFiles?.size) {
          await hydrateAssets(editor, loaded.text, 'browser-directory', (_rootId, path) =>
            readLoadedMarkdownAsset(loaded, path),
          )
        }
      })
      .catch((error) => {
        console.error('[markdown] open failed:', error)
        statusRef.current = 'error'
        setStatus('error')
      })
  }, [editor, hydrateAssets, loadText])

  const exportDocx = useCallback(async (): Promise<
    { ok: true; fileName: string } | { ok: false }
  > => {
    const current = editorRef.current
    if (!current) return { ok: false }
    try {
      const bytes = await exportDocxBytes(current.getJSON(), loadMarkdownImageForDocx)
      const stem = (fileNameRef.current ?? 'Untitled.md').replace(/\.(md|markdown)$/i, '')
      const exportedName = stem + '.docx'
      downloadMarkdownExport(
        bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer,
        exportedName,
        'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      )
      return { ok: true, fileName: exportedName }
    } catch (error) {
      console.error('[markdown] DOCX export failed:', error)
      return { ok: false }
    }
  }, [])

  const openPrintDialog = useCallback((): { ok: true } | { ok: false } => {
    const current = editorRef.current
    if (!current) return { ok: false }
    const title = (fileNameRef.current ?? 'Untitled.md').replace(/\.(md|markdown)$/i, '')
    const popup = window.open('', '_blank')
    if (!popup) return { ok: false }
    popup.document.open()
    popup.document.write(buildPrintHtml(current.view.dom, title))
    popup.document.close()
    popup.addEventListener('load', () => popup.print(), { once: true })
    return { ok: true }
  }, [])

  useEffect(
    () => subscribeLiveEditorDisplayMode(() => refreshDisplayMode((value) => value + 1)),
    [],
  )

  useEffect(() => {
    if (!editor) return
    return attachMcpLiveSession({
      snapshot: (revision) => ({
        revision,
        fileName: fileNameRef.current,
        dirty: dirtyRef.current,
        selection: markdownSelection(editor),
      }),
      recoverySnapshot: async (force) => {
        if (!force && !dirtyRef.current) return null
        return {
          fileName: fileNameRef.current ?? 'Untitled.md',
          data: recoveryBytes(serializeCurrent()),
        }
      },
      recoveryVersion: () => recoveryVersionRef.current,
      execute: async (command): Promise<LiveEditorExecution> => {
        try {
          let trace: LiveEditorExecutionTrace | undefined
          const registered = await executeMarkdownOperation(editor, command, {
            loadStaged: async ({ name, data, assetRootId }) => {
              const decodeStartedAt = performance.now()
              const raw = new TextDecoder().decode(data)
              const decodeMs = performance.now() - decodeStartedAt
              const phases = await loadText(name, raw, false)
              trace = {
                operation: 'markdown.document.load_staged',
                phases: { decodeMs, ...phases },
              }
              if (assetRootId) {
                await hydrateAssets(editor, raw, assetRootId, readLiveEditorLocalAsset)
              }
            },
            save: ({ saveAs }) => doSave(saveAs),
            exportDocx,
            openPrintDialog,
            setAutoSave: ({ enabled }) => setAutoSaveState(enabled),
            setFrontmatter: ({ yaml }) => applyFrontmatter(yaml),
            setZoom: setZoomState,
          })
          if (registered.handled) {
            if (!registered.ok) {
              return {
                ok: false,
                error: registered.error,
                message: registered.message,
              }
            }
            if (registered.output) {
              return {
                ok: true,
                output: { ...registered.output },
                ...(trace ? { trace } : {}),
              }
            }
            if (registered.checkpointRecovery === false) {
              return { ok: true, ...(trace ? { trace } : {}) }
            }
            return {
              ok: true,
              ...(trace ? { trace } : {}),
              recovery: {
                fileName: fileNameRef.current ?? 'Untitled.md',
                data: recoveryBytes(serializeCurrent()),
              },
            }
          }
          return {
            ok: false,
            error: 'unsupported_operation',
            message: 'Unsupported Markdown operation: ' + command.operation,
          }
        } catch (error) {
          return {
            ok: false,
            error: 'execution_failed',
            message: error instanceof Error ? error.message : String(error),
          }
        }
      },
    })
  }, [
    applyFrontmatter,
    doSave,
    editor,
    exportDocx,
    hydrateAssets,
    loadText,
    openPrintDialog,
    serializeCurrent,
    setAutoSaveState,
    setZoomState,
  ])

  useEffect(() => {
    if (!autoSave) return
    const saveIfNeeded = () => {
      if (dirtyRef.current && canOverwriteMarkdownFile() && !editorRef.current?.view.composing) {
        void doSave()
      }
    }
    const timer = window.setInterval(saveIfNeeded, 30_000)
    window.addEventListener('blur', saveIfNeeded)
    return () => {
      window.clearInterval(timer)
      window.removeEventListener('blur', saveIfNeeded)
    }
  }, [autoSave, doSave])

  useEffect(() => {
    const onKeyDown = (event: KeyboardEvent) => {
      if (!(event.metaKey || event.ctrlKey) || event.altKey) return
      const key = event.key.toLowerCase()
      if (key === 's') {
        event.preventDefault()
        void doSave(event.shiftKey)
      } else if (key === 'p' && !event.shiftKey) {
        event.preventDefault()
        openPrintDialog()
      } else if (key === '=' || key === '+') {
        event.preventDefault()
        zoomIn()
      } else if (key === '-' || key === '_') {
        event.preventDefault()
        zoomOut()
      } else if (key === '0') {
        event.preventDefault()
        setZoomState({ percent: 100 })
      }
    }
    window.addEventListener('keydown', onKeyDown, true)
    return () => window.removeEventListener('keydown', onKeyDown, true)
  }, [doSave, openPrintDialog, setZoomState, zoomIn, zoomOut])

  useEffect(() => {
    const onWheel = (event: WheelEvent) => {
      if (!event.ctrlKey && !event.metaKey) return
      if (!(event.target as HTMLElement | null)?.closest?.('.editor-scroll')) return
      event.preventDefault()
      setZoom((value) => normalizeMarkdownZoom(value - event.deltaY * 0.6))
    }
    window.addEventListener('wheel', onWheel, { passive: false })
    return () => window.removeEventListener('wheel', onWheel)
  }, [])

  const statusText =
    saveState === 'saving'
      ? t('saving')
      : saveState === 'failed'
        ? t('saveFailed')
        : dirty
          ? t('unsaved')
          : saveState === 'saved'
            ? t('savedOk')
            : ''

  if (status === 'error') {
    return (
      <main className="app">
        <div className="center-note">{t('loadError')}</div>
      </main>
    )
  }

  return (
    <main className="app">
      <Ribbon
        editor={editor}
        disabled={status !== 'ready'}
        dirty={dirty}
        onOpen={openFile}
        onSave={() => void doSave()}
        onExportDocx={() => void exportDocx()}
        onPrintPdf={() => void openPrintDialog()}
        autoSave={autoSave}
        onToggleAutoSave={setAutoSaveState}
        onInsertImage={insertImage}
        frontmatterOpen={frontmatterOpen}
        onToggleFrontmatter={() => setFrontmatterOpen((value) => !value)}
        fullscreen={getLiveEditorDisplayMode().mode === 'fullscreen'}
        onToggleFullscreen={() => void toggleLiveEditorFullscreen()}
      />
      {status === 'loading' && <div className="center-note">{t('loading')}</div>}
      <div className="app-main" hidden={status !== 'ready'}>
        <div className="app-content">
          {editorActive ? (
            <div className="editor-scroll" ref={scrollRef}>
              <div className="doc-page" style={{ zoom: zoom / 100 }}>
                {frontmatterOpen && (
                  <FrontmatterPanel value={frontmatterText} onChange={applyFrontmatter} />
                )}
                <EditorContent editor={editor} />
              </div>
            </div>
          ) : (
            <div className="editor-workspace-suspended" aria-hidden="true" />
          )}
          <footer className="status-bar" aria-live="polite">
            <div className="status-left">
              {fileName && <span className="status-item status-file">{fileName}</span>}
            </div>
            <div className="status-right">
              {statusText && (
                <span className={'status-save status-' + saveState}>{statusText}</span>
              )}
              <button
                type="button"
                className="zoom-btn"
                aria-label={t('zoomOut')}
                onClick={zoomOut}
                disabled={zoom <= MARKDOWN_MIN_ZOOM}
              >
                −
              </button>
              <input
                className="zoom-slider"
                type="range"
                min={MARKDOWN_MIN_ZOOM}
                max={MARKDOWN_MAX_ZOOM}
                step={MARKDOWN_ZOOM_STEP}
                value={zoom}
                aria-label={t('zoom')}
                onChange={(event) => setZoomState({ percent: Number(event.target.value) })}
              />
              <button
                type="button"
                className="zoom-btn"
                aria-label={t('zoomIn')}
                onClick={zoomIn}
                disabled={zoom >= MARKDOWN_MAX_ZOOM}
              >
                +
              </button>
              <span className="zoom-value">{zoom}%</span>
            </div>
          </footer>
        </div>
      </div>
      <SlashMenu ref={slashMenuRef} state={slashState} onDismiss={() => setSlashState(null)} />
      <TableMenu editor={editor} scrollRef={scrollRef} zoom={zoom} />
    </main>
  )
}
