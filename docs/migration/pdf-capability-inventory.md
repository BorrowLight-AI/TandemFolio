# PDF source-current native capability inventory

- Extraction baseline: `genspark-ai/genoffice@dc4d7e5927864498913b7ba42d0da06cc7cf628e`
- Reviewed candidate: `genspark-ai/genoffice@f2c3d0879df29622d5a447935d2b4aeac033544d`
- Scope: browser-safe community PDF renderer, format engine, focused tests, and TandemFolio browser/MCP adapters; `ee/` was not inspected
- Status: applicable native PDF migration and typed command parity complete; 2026-09-07 source-current release evidence approved and the release gate remains fail-closed for later drift

## Source boundary

The candidate changes native PDF behavior in ten snapshots between `d5558b6` and `2239cce`;
`99d376b` and the later `99dfbc1` change only the prohibited PDF AI panel. No PDF-native behavior was
added between the prior `360ce06` review point and the final `f2c3d08` candidate. The retained renderer
is the community visual editor. Its file calls are
adapted to `host/browser-pdf-api.ts`, while `community-command-bridge.ts` binds the typed Registry to
the same mounted React state. PDF.js remains the viewing authority; browser PDFium and PDF-lib are
format-owned save helpers, never a second mounted editor.

The following candidate-native renderer modules are retained or browser-adapted:

- `ColorPicker.tsx`, `PdfPage.tsx`, `PdfThumb.tsx`, and `icons.tsx` for the updated native editor chrome,
  page and thumbnail lifecycle, and accessible Office color selection;
- `NoteMargin.tsx`, `annotation-catalog.ts`, `note-margin-layout.ts`, and `note-threads.ts` for threaded
  margin comments;
- `PasswordDialog.tsx` for encrypted-file password retry and read-only viewing;
- `SignDropOverlay.tsx` and the updated `SignatureDialog.tsx` for signature placement and the local
  saved-signature library;
- `doc-font.ts`, `edit-state.ts`, `text-edit-preview.tsx`, and the updated text block/wrap/color modules
  for paragraph editing, movement, selection formatting, overflow preview, and font matching;
- `view-config.ts` and `view-state.ts` for display-mode options and per-file reading-position restore.

Candidate `src/main` algorithms for blank pages, comments, annotations, page transforms, signature
storage, and text editing were ported into `src/domain` or the browser host. Their Electron filesystem
and IPC wrappers were not copied.

## Retained native capabilities

| Capability group | Retained behavior | Mounted implementation and evidence |
| --- | --- | --- |
| Viewing and navigation | Page and thumbnail render cancellation/restart, thumbnail context menus, outline/thumbnails/comments sidebars, single/continuous/two-page modes, fit width/page/custom zoom, internal links, current-page navigation, and per-file scroll/zoom/sidebar restore | `PdfPage.tsx`, `PdfThumb.tsx`, `view-config.ts`, `view-state.ts`; lifecycle and view-state unit tests plus existing real-host width/offscreen scenarios |
| Search and print | Whole-document search with stable navigation and trailing-space-safe matches; browser print after flushing current edits | `search.ts`, renderer print path, and existing search/print tests; print remains a non-document host effect |
| Password-protected PDFs | Password prompt, reveal toggle, retry/error state, and cancellation; a successful encrypted open is deliberately read-only because the browser save engines cannot preserve encryption | `PasswordDialog.tsx` and the `App.tsx` open lifecycle; the host open acknowledgement remains pending across password retries and ends on success or explicit cancellation |
| Threaded comments | Root comments and replies with author/time metadata, margin collision layout, active-thread navigation, create/reply/edit/delete, exact same-rectangle reply identity, and save/reopen preservation of `/IRT` and `/RT` | `NoteMargin.tsx`, note helpers, `domain/save-pdf.ts`, `domain/annot-delete.ts`; `note-threads`, `note-margin-layout`, and `note-thread-save` tests |
| Markup, ink, notes, and signatures | Highlight/underline/strike toggle, ink and note placement, move/resize, saved annotation deletion, signature draw/type/image modes, signature-field placement, drag/drop placement, and a bounded deduplicated saved-signature library | Existing annotation/drawing state plus updated layers/dialogs and browser signature store; document mutations share Registry/Undo/save, local signature-library state is a host preference |
| Native text editing | Paragraph/block selection, exact run matching, object-preserving move, fragment replacement without dropping neighbours, deletion across multiple native objects, multiline layout, alignment/indent offsets, font/size/bold/italic/color and selection-level style runs, CJK/Korean/Arabic bundled-font coverage, unsupported-glyph rejection, overflow preview, and searchable saved output | `domain/browser-text-edit.ts` and renderer text modules; real PDFium-WASM tests cover movement, selection styling, fragment replacement, and deletion, while existing font/save-reopen tests cover multilingual output |
| Images and forms | Existing content-image insert, select, move, resize, rotate, layer, replace, crop, cutout, flip, opacity and delete; static text/check/cross fills; AcroForm controls and signature widgets | Updated image/form layers and dialogs retain the existing typed image/static-form/form routes and PDFium/PDF-lib save paths |
| Page organization | Insert another PDF, insert a matching blank page, delete, reorder/reverse, rotate one/all pages, replace selected pages from a PDF, crop one/all pages, and set all pages to an explicit paper size | Page UI and browser host share PDF-lib kernels. In-place structural writes preserve the prior bytes in a bounded mounted-document undo/redo journal and reload the same file |
| Page/document outputs | Extract page ranges, split a PDF into chunks, merge selected pages N-up with direction/separator options, split each page into 2/4/9 pages, merge multiple PDFs, export page images, Save As, and browser print | Output operations create separate files and therefore remain declared host effects; they do not replace the mounted document |
| Blank creation and file integrity | Create a one-page A4 document with explicit replacement confirmation; persistent page writes commit only after the browser save succeeds; failed persistence keeps mounted bytes and history unchanged | `domain/blank-pdf.ts`, browser host, blank/page operation tests, and the explicit-replacement gate |

## Registry and producer accounting

The PDF-owned catalog contains 32 descriptors: 29 Agent-visible operations and three internal
staged-byte routes. `pdfRetainedProducerBaseline` has no `missing` entry and every retained document
mutation resolves to a descriptor and executable handler.

| Family | Operations |
| --- | --- |
| Document | `pdf.document.create_blank`, `pdf.document.set_metadata`, `pdf.document.save`; internal `pdf.document.load_staged` |
| History | `pdf.history.undo`, `pdf.history.redo` |
| Markup and comments | `pdf.markup.add`, `pdf.note.update_saved`, `pdf.annotation.delete_saved`, `pdf.pending.delete` |
| Drawing and signature placement | `pdf.drawing.add`, `pdf.drawing.update` |
| Text | `pdf.text.insert`, `pdf.text.replace`, `pdf.text.update_inserted` |
| Images and static forms | `pdf.image.insert`, `pdf.image.transform`, `pdf.image.replace`, `pdf.image.delete`, `pdf.static_form.set` |
| AcroForm | `pdf.form.set_value` |
| Watermark/header/footer | `pdf.stamp.set` |
| Pages | `pdf.page.insert`, `pdf.page.insert_blank`, `pdf.page.delete`, `pdf.page.replace`, `pdf.page.crop`, `pdf.page.reorder`, `pdf.page.set_size`, `pdf.page.set_rotation`; internal `pdf.page.insert_staged`, `pdf.page.replace_staged` |

Public insertion/replacement accepts a bounded local path that the Broker hydrates into its matching
internal operation. Immediate insert, replace, blank, crop, and paper-size changes share the browser
host byte journal; `pdf.history.undo` and `pdf.history.redo` fall through to it when the mounted App
has no newer pending-state entry. Save/reopen preserves the resulting PDF bytes.

Signature-library changes, reading-position preferences, search/navigation/zoom/sidebar state,
Save As, image export, extraction, split, merge, N-up, print, and password entry do not mutate the
mounted PDF content and therefore do not require additional document operations.

## Explicit exclusions

| Candidate area | Disposition |
| --- | --- |
| `renderer/ai/**`, `AiAskPopover.tsx`, AI selection actions, PDF agent tools, AI navigation, and `99d376b` | Excluded by the user and product boundary; no model/provider/search flow is present |
| `src/main/ocr.ts` and `renderer/ocr-layer.tsx` | Excluded because this is a native executable sidecar/raster-recognition pipeline rather than original PDF document formatting |
| `auto-rename`, OS username lookup, and generated-document output | Excluded with account/AI-created-document lifecycle; comments use the local format-owned author label `GenOffice` |
| Electron `src/main`, preload, IPC, shell drag/drop, native PrintDialog, filesystem atomic-write wrappers | Replaced by the existing browser/MCP host, browser print, and mounted in-memory persistence transaction |
| Desktop system-font discovery | Replaced by lazy bundled browser font assets and cmap validation; retained text output stays portable and deterministic |
| Encrypted-PDF editing or encryption-preserving save | Not admitted because neither retained browser save engine can preserve the encryption envelope; encrypted files open read-only instead of silently stripping protection |
| `ee/` | Prohibited and not inspected |

## Verification evidence

The PDF workspace currently passes 35 test files and 332 assertions. Focused additions cover:

1. native threaded-comment creation, reply identity, edit, delete, and save/reopen metadata;
2. blank creation, page insertion/replacement/crop/resize, split/merge/N-up transforms, and PDF-lib reopen;
3. page-mutation persistence failure plus whole-document undo/redo;
4. real PDFium-WASM movement, selection styling, fragment replacement, multi-object deletion, and typed empty-text deletion;
5. signature-library bounding/deduplication, reading-position state, page lifecycle helpers, document-font matching, color presets, and operation parity.

The generated Product Manifest, root typecheck/build, packaged MCP smoke, resource limits, licenses,
and prohibited-dependency scan pass. All 16 PDF real-browser host scenarios pass, including
open/save/reopen, typed content-stream edits, shared Undo, narrow/split/fullscreen layout, exact-session
continuation, and offscreen canvas release/resume. Earlier R6 performance/visual evidence is historical;
`release:gate` writes `ready: true` for the approved 2026-09-07 formal five-format source-current
capture and returns to `ready: false` when evidence is stale, unapproved, or mismatched.
