# Migration ledger

## 2026-09-05 selective upstream synchronization

The extraction baseline remains `dc4d7e5927864498913b7ba42d0da06cc7cf628e`.
Reviewed candidate: `360ce0625eaf748368e5535984b073f6fb2487b5`.
See [commit dispositions, integrated DOCX capabilities, and cross-format backlog](upstream-review-2026-09-05.md).
This synchronization now retains all applicable DOCX-native editing, layout, rendering, and
round-trip work from the reviewed candidate, plus the earlier cross-format fixes. AI, enterprise,
account, telemetry, Electron/IPC, and desktop-only encryption/printing remain excluded:

| Source commit | Modified/admitted community paths |
| --- | --- |
| `945c370` | Modified `packages/docx-engine/src/parse.ts`; added `packages/docx-engine/tests/entity-decoding.test.ts` |
| `3548a80` | Modified `apps/slides/src/renderer/App.tsx`, `SlideCanvas.tsx`, `TextEditOverlay.tsx`; added `apps/slides/src/renderer/table-hit.ts` and `apps/slides/tests/table-hit.test.ts` (extended for flips/bleed) |
| `4846ace` | Modified `apps/slides/src/renderer/file-actions.ts`; added/adapted `apps/slides/tests/save-serialization.test.ts` for browser-host globals, Save As and rejected writes |
| `bc1dceb` | Modified `apps/sheets/src/renderer/univer-sync.ts`; added `apps/sheets/tests/read-sheet-range-batching.test.ts` |
| XLSX formula-reliability snapshot lineage through `9f971ed` | Retains the post-read structural-race guard for streamed formula closure and adapts the upstream formula-engine size gate to browser-open `Blob` bytes. Oversized workbooks and truncated formula indexes keep cached values without starting the fallback engine; AI planning guards and desktop sidecar tests remain excluded. |
| XLSX cross-save history snapshot lineage through `9f971ed` | Retains serializable Univer undo/redo across ordinary save/reopen Session replacement, drops closure-backed renderer mutations, clears load decoration, and skips unsafe sheet-identity/two-phase structural cases. Browser table registration now waits for its native mutations, retains file column labels, and suppresses the default theme over file-rendered style. |
| XLSX Conditional Formatting snapshot lineage through `9f971ed` | Retains formula-range folding/windowing and complete highlight-rule hydration, including error predicates, case-insensitive text, stop-if-true, dxf borders/number formats, nonnumeric operands, and native data-bar options. The optimization narrows only engine registration; original rule ranges, Undo, and OOXML save remain unchanged. |
| XLSX cross-highlight snapshot lineage through `9f971ed` | Retains the native active-cell row/column reading aid as one bounded Univer canvas extension per worksheet, with local preference persistence, theme refresh, streamed-sheet attachment, and complete disposal. It does not mutate document state or add AI search. |
| XLSX lazy-find snapshot lineage through `9f971ed` | Extends Univer's native Find/Replace session over unstreamed file rows, merging loaded and file-backed hits while honoring journal shadows, filters, sheet/order navigation, native replacement/Undo, and a 400,000-cell browser scan budget. Scan constants and localized status copy are format-local; no AI search source is retained. |
| XLSX error-checking snapshot lineage through `9f971ed` | Retains Formulas › Error Checking for loaded and streamed workbooks. The bounded browser scan observes session edits and structural mappings, includes computed journal formula results, loads distant targets before navigation, and never imports the upstream AI-search module. |
| XLSX Watch Window snapshot lineage through `9f971ed` | Retains the native non-modal formula inspection panel with stable sheet/cell identities, a bounded 20-cell watch set, live value/formula refresh, sheet-rename tolerance, and explicit removal. It is session-local read-only state and does not enter workbook revision or persistence. |
| XLSX Merge Workbooks snapshot lineage through `9f971ed` | Adapts the Electron multi-file source sessions to browser-owned secondary sessions and a standard multi-file picker. Imported sheets retain cached values, styles, merges, row heights, and column widths through native Univer commands and the existing save journal; `office_merge_local_workbook` supplies the typed staged-byte route. |
| XLSX embedded-OLE snapshot lineage through `9f971ed` | Adapts sidecar OLE discovery to the browser package reader: `oleObjects`, objectPr anchors, legacy VML fallback anchors/frames, cached raster or metafile previews, and friendly ProgID captions render read-only. Opaque embeddings and preview parts remain byte-preserved, while worksheet/VML/drawing/table/chart references shift with structural edits. |
| XLSX Save As toolbar snapshot lineage through `9f971ed` | Retains the dedicated quick-access Save As command for clean or dirty file-backed workbooks, routed through the existing browser package save-as target selection with localized copy and its distinct glyph. |
| XLSX defined-name snapshot lineage through `9f971ed` | Retains Use in Formula and top-row/left-column Create from Selection, including Unicode label normalization, lazy file-extent bounds, duplicate avoidance, native Univer name history, OOXML save, and typed `xlsx.defined_name.create_from_selection` parity. |
| XLSX snapshot lineage through `9f971ed` | Selectively admitted the browser-safe native spreadsheet work for OOXML namespace/relationship fidelity, Rich Data, shared and future formulas, themes, protection metadata, complete 1904 date-system parsing/display/pivot semantics, calculation control, cached results, IFS behavior, formula/delete guards, center-continuous, shrink-to-fit, fill-isolation, text-rotation, workbook-MDW/default-column-width layout, native number-format `*x` fill, cached/fixed/spacer row-height semantics, row/column default-style rendering, explicit empty-string/style-only cell streaming, active-cell data-validation chrome, selected-only non-overlapping visual deletion, structured table-reference thresholds, rich-text subscript/superscript open-edit-save fidelity, and sparse chart-cache, category-axis label, line-style, and legend rendering, browser-side `cellXfs`/dxf/Conditional Formatting hydration, blank coercion, priority/shadow semantics, formula-threshold folding, color-scale clamping, and native data-bar bounds, exact file print settings and page-specific header/footer layout, fit-to-page printing, manual page breaks and page-break preview, criteria comparison, selection clearing, and CSV serialization. Electron/preload/IPC, native sidecar, recovery shell, desktop PDF printing, AI, account, telemetry, and enterprise sources remain excluded. |
| DOCX snapshot lineage through `9f971ed` | Admitted the candidate DOCX engine, non-AI renderer modules, fonts, native tests, dropdown/popover primitives, and browser-authority adaptations. This covers parse/save fidelity, pagination/sections/columns, tables, drawings, notes, review UI, shortcuts, and formatting commands. |
| `93b8938` | Admitted and connected per-section widths/margins and percentage-table layout, including gap header/footer alignment. |

Apache-2.0 attribution remains intact. Newly adapted files carry prominent modification
notices. Mounted Session authority and Save target binding are unchanged. The new
`docx.image.set_z_order` and the three XLSX calculation routes raise the generated Registry to
103 DOCX / 123 XLSX / 347 total operations;
the user gesture and Agent command share one native transaction and Undo path. The previous approved release capture becomes
historical after these source changes; the source-current gate is not relaxed.
Generated plugin HTML resources were rebuilt from the workspace; the gate regenerated
`apps/mcp-server/src/generated/release-readiness.json` as not ready. These generated
outputs introduce no additional upstream source. Functional results are in the linked review.


The source baseline is the GenOffice community repository at commit `dc4d7e5927864498913b7ba42d0da06cc7cf628e`, licensed under Apache-2.0. Enterprise `ee/` source is outside the migration boundary and was never copied.

See [`provenance.md`](provenance.md) for the file-level extraction/modification record and [`roadmap.md`](roadmap.md) for future format gates.

The 2026-09-04 XLSX bootstrap refinement changes only owned packaging and browser
adapter loading: static dependencies are recorded during packaging, the critical
entry is identity text behind one importable Blob URL, dependency-free lazy modules
remain bytes through Blob construction, and ZIP/XML package I/O loads on first
`openBuffer`. It neither moves upstream areas nor removes retained behavior.
The 500 ms bootstrap, 1,400 ms total, and 11 MB entry gates remain fixed; current
release authorization is recorded in [`../../release/validation.md`](../../release/validation.md).

This table records current code and historical accepted milestone evidence, not
current release authorization. ADR 0003 rejects simplified
replacements as target architecture. All five formats have crossed their format-local retained
state-changing command parity audits, R6-01 closes the shared release gates through the
source-fingerprinted evidence defined by ADR 0005, R6-02 replaces fixed command polling under ADR
0006, R6-03 traces and bounds Markdown staged loading under ADR 0007, and R6-04 traces and bounds
XLSX cold start under ADR 0008. R6-05 decomposes and optimizes the remaining XLSX bootstrap
hotspot under ADR 0009 without changing renderer authority or capability. R6-06 implements ADR
0004's bounded summary/detail Manifest discovery, stable session-aware availability, and response
size gates without adding tools or resources. R6-07 implements the revision-guarded
`office_execute` transaction envelope, session-scoped idempotency journal, typed pre-dispatch
validation, and ACK-owned timeout/replay semantics. R6-08 removes the old `operation` + `arguments`
Adapter and legacy response branch after migrating all repository callers and tests; packaged smoke
guards the transaction-only public schema. R6-09 removes all thirty-six public-era aliases from the
five format registries and generated Manifest. Internal staged-file transport uses each format's
canonical `*.document.load_staged` operation; the remaining alias set is empty.

| Area                                                                                       | Upstream disposition           | TandemFolio status                                                  | Notes                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                          |
| ------------------------------------------------------------------------------------------ | ------------------------------ | ------------------------------------------------------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `apps/docs/src/renderer`                                                                   | Selectively retained           | Retained command parity complete; release recapture pending         | Browser file operations replace the preload boundary; the common Host Bridge routes acknowledged Agent mutations into the mounted editor and supplies display-mode and lazy bundled-font access. One hundred three DOCX descriptors cover the retained command inventory, including image stacking and the previously closed text, references, revision, design, structure, drawing, and persistence families. Registry and retained UI share native edit kernels, deterministic identities/content, Undo or renderer-owned journals, recovery, and save/reopen projection as appropriate. The 2026-09-05 candidate's applicable non-AI native renderer work is integrated. Prior R6 evidence is historical and the source-current gate generates `ready: false`. Prohibited product areas remain excluded. |
| `apps/sheets/src/renderer`                                                                 | Selectively retained/adapted   | Candidate-native XLSX migration in progress; release recapture pending | `main.tsx` mounts the community `App.tsx` directly. The 122-operation XLSX registry (120 Agent-visible, two internal) includes typed calculation-mode, recalculation, Goal Seek, native themes, workbook protection, allow-edit ranges, manual page breaks, and create-names-from-selection through mounted Univer/file-journal state. Candidate-native formula correctness, structural guards, native style hydration, shrink-to-fit rendering, printing, keyboard clearing, and CSV serialization are being admitted in vertical TDD slices. Source-current release evidence remains fail-closed while this migration changes the renderer. See [`xlsx-capability-inventory.md`](xlsx-capability-inventory.md). |
| `apps/sheets/src/host/browser-workbook.ts`                                                 | TandemFolio-owned host adapter | Browser OOXML package boundary; not a renderer                      | Supplies browser-only package open/save to the mounted community App across cells/formulas, layout, validation/CF, hyperlinks, tables, protection, sparklines/outlines, drawings/media, notes, Pivot parts, page setup, tab color, and worksheet duplication/visibility. Unsupported or lossy OOXML states fail explicitly rather than silently diverging.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                     |
| `apps/slides/src/renderer`, `src/main/edit-text.ts`                                        | Selectively retained/adapted   | Retained command parity and R6-01 release gate complete             | All 80 permitted renderer paths are present (61 byte-identical, 19 browser-host/product-boundary adapted); 24 AI modules/assets are excluded. The original App/Ribbon/Konva/text editor and Presenter/Audience views run through complete browser host adapters. The 74-operation PPTX registry (73 Agent-visible plus staged load) covers every retained mutation family through one `BrowserPresentation`, native history, recovery, renderer refresh, and package save seam. A machine-checked producer baseline has no missing entry; export/print/show are classified host effects. Approved R6-01 evidence generates `ready: true`. See [`pptx-capability-inventory.md`](pptx-capability-inventory.md).                                                                                                                                                                                                                                                                                                                                                                                                                                  |
| `apps/pdf/src/renderer`, `src/domain`                                                      | Selectively retained/adapted   | Retained command parity and R6-01 release gate complete             | All 33 pinned non-AI renderer files are present; seven AI files/assets are excluded. The 25-operation Registry (23 public, two internal) covers the complete retained producer baseline. Mounted App history or an explicitly declared host persistence route remains authoritative; browser PDFium restores searchable text/image content-stream editing with lazy allowlisted fonts. PDF-lib-safe annotation, drawing, form, stamp, metadata, and page families persist through save/reopen. Twenty-three test files / 283 assertions and eleven real-host scenarios pass. Approved R6-01 evidence generates `ready: true`. See [`pdf-capability-inventory.md`](pdf-capability-inventory.md).                                                                                                                                                                                                                                                                                                                                                                                                                                                |
| `apps/markdown/src/renderer`                                                               | Selectively retained           | Retained parity and R6-03 staged-load gate complete                 | Twenty-one pinned non-AI renderer files are present plus browser/operation adapters. Twenty public operations plus two internal routes have generated visibility and exact Broker/renderer validation. R6-03 ACKs decode/parse/TipTap-install/React-commit phases only after committed state is visible; TandemFolio candidate guards remove quadratic non-matching list/table tokenizer probes while retaining actual syntax behavior. Approved schema-v4 evidence retains the five-second canonical-large ceiling introduced in R6-03. See [`markdown-capability-inventory.md`](markdown-capability-inventory.md).                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                           |
| `packages/docx-engine`                                                                     | Retained                       | Product workspace                                                   | Format engine with provenance preserved; the retained section writer serializes explicit column spacing, while original/new image writers persist non-destructive source crop and materialize a missing minimal picture transform container before applying rotation/flip state.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                               |
| `packages/pptx-engine`, `packages/pptx-render`                                             | Retained and browser-adapted   | Product workspace                                                   | Browser-safe crypto/compression paths drive the PPTX visual editor; Node streaming save is isolated to the `node-file` subpath.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                |
| `packages/file-parse`, `packages/font-metrics`, `packages/i18n`, `packages/ui`             | Retained                       | Product workspace                                                   | Shared community infrastructure; AI-only UI components were removed from `packages/ui`. `file-parse` is retained only for later format review and is not reachable from DOCX.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                  |
| `packages/operation-contract`, `tools/{generate-operation-manifest,operation-catalogs}.ts` | TandemFolio-owned              | R1 foundation; R6-09 canonical public names complete                | Defines only format-neutral descriptors, transactions, errors, validation results, bounded schema validation including enum/item/range/string-length/pattern constraints, deterministic manifest generation, and drift checks. The product catalog entry aggregates all five serializable catalogs without importing renderer handlers. R6-09 removes the thirty-six public-era aliases; a cross-format test fixes the complete remaining alias set at empty, including staged-file transport.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                 |
| `packages/host-bridge`, `apps/mcp-server/src/{session-store,transaction-journal}.ts`       | TandemFolio-owned              | R6-10 exact Session/view isolation complete                         | All five formats use one shared live-session bridge for R6-02 wakeable bounded polling, staged-byte hydration, recovery-before-ACK, display mode, and lazy bundled fonts. R6-10 makes each mounted `(sessionId, viewId)` immutable, grants one Broker-enforced view lease per Session, waits for the lease-bearing show result before polling, defaults creation to an isolated blank Session, and restores known work only by exact id. Duplicate views cannot poll, acknowledge, transfer Session bytes, write recovery, or disconnect the owner. Markdown and XLSX retain strict traced first-poll readiness boundaries; transport retries preserve one-shot startup evidence until a poll succeeds. R6-07 separates command completion from each caller timeout and journals session-scoped fingerprints/results, so completed or in-flight replay does not redispatch and staged bytes remain available through final acknowledgements. R6-08 requires the public transaction envelope; R6-09 makes Registry ingress canonical-only. Public tools, five UI resource bindings, renderer authority, and one-revision-per-ACK are unchanged. |
| DOCX Electron `main`/`preload`, IPC, updater, and build trees                              | Imported as migration input    | Removed                                                             | Browser and MCP host adapters now own the DOCX runtime boundary.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                               |
| Standalone Electron shell, Electron E2E harness, and updater scripts                       | Imported as migration input    | Removed                                                             | These host-only areas remain outside the renderer product boundary. The Markdown renderer is no longer grouped with this removal.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                              |
| DOCX renderer AI folders and shared AI composer/markdown components                        | Imported as migration input    | Removed                                                             | Generic document commands and live-text helpers were re-homed under `renderer/editor`; provider and attachment workflows were discarded.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                       |
| AI/model, login/account, telemetry, Electron main/preload/IPC, old sidecars                | Imported as migration input    | Removed from current code and prohibited in all five target formats | Product dependency audit is empty for the deleted packages and Electron. Marketplace `policy.authentication` is required Codex install metadata, not application authentication code.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                          |
| `ee/`                                                                                      | Not imported                   | Prohibited                                                          | Enterprise source has a separate boundary and must never be copied or inspected for migration.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                 |

## Modification policy

Retained files continue to carry the repository-level `LICENSE` and `NOTICE`. Materially rewritten or newly created TandemFolio code uses the TandemFolio namespace. The ledger must be updated whenever a retained upstream area is removed, replaced, or admitted into the product build.

## 2026-09 editor reliability repair

### 2026-09-03 explicit document replacement

ADR 0015 distinguishes exact Session continuation from whole-document replacement. PPTX's existing
`document.create_blank` operation now requires `confirmReplace: true` before resetting a named,
dirty, or undo/redo-bearing presentation. The default rejection occurs before Save unbinding or
native replacement; explicit new retains the existing separate-output behavior. The shared Skill
preserves the current document for follow-up generation in all five formats. No upstream file or
other format's operation surface changes; owned paths and regression scope are in `provenance.md`.

### 2026-09-03 user-directed continuation

ADR 0014 partially supersedes ADR 0013's visibility-only eligibility and unbounded wait duration.
All five formats expose **在此继续编辑** on a waiting same-view replica, selecting that mount and
requesting a cooperative fresh checkpoint even when the old owner still reports active. Prepared
handoffs cannot be retargeted; input, pending command/upload, native restore, and revision fences
remain intact. Automatic waiting ends after 30 seconds. Yielded mounts remain suspended until
explicitly selected, while uncertain transport outcomes stay locked pending lease reconciliation.
An explicit selection received during an in-flight ordinary retry is retained for the next poll;
the waiting UI disables duplicate actions and reports the pending safe handoff until it settles.
This changes only shared owned transport/status code and tests, not format renderers, formal Save,
or the public tool inventory. See `provenance.md` for paths and `development.md` for validation.

### 2026-09-03 cooperative replay handoff

ADR 0013 partially supersedes terminal same-view conflicts: a visible replay requests a fresh native
checkpoint from the hidden owner, then restores under an exclusive transferred lease. The bridge
blocks editing until restore/ACK, retains location visibility, pauses hidden waiting/yielded views,
and stops automatic retry after checkpoint failure. Prepared handoff rejects Agent mutations and
formal saves; pending uploads/commands prevent preparation. Abandoned idle owners still use the
30-second liveness limit. No cross-Session adoption, forced active-owner takeover, new renderer, or
formal Save on switching is introduced. One app-only tool is added; public tools stay unchanged.
The owned-file and packaged five-format regression scope is recorded in `provenance.md`.

### 2026-09-03 exact resume and Save-target completion

ADR 0012 extends the prior repair across all five formats. Cold mount identity prevents replayed
show results from competing as one owner; rejected mounts stop polling and expose explicit retry.
Owner contact keeps a lease alive, and takeover of an abandoned lease requires 30 seconds idle and
no pending/active command. Exact restore prefers the Session's checkpoint, otherwise its bound
formal file, and never adopts another task's document. Native restore failures stay visible.

Browser-selected replacement detaches the previous Save target and retires old uploads before
native load; the bridge serializes any concurrently delivered command behind replacement and a
forced checkpoint. Committed bound Save retires older recovery; Export Copy does not. Each editor
displays/copies the committed absolute path. XLSX canonical Save preserves the existing target,
explicit clean saves persist, and recovery uses the same package serializer in memory without
entering the formal-save transport. PDF explicit clean saves now write instead of returning a false
success, and DOCX synthetic source identities preserve the filename required for first persistence.
The modified upstream and owned paths are listed in
[`provenance.md`](provenance.md). No legacy-data migration, renderer replacement, or release-gate
exception is included.

The new packaged-browser lifecycle suite covers all five native saved-file reopen paths, clean
browser imports and exact checkpoint resume, plus multiple rejected PPTX mounts. This focused repair
evidence does not replace the missing source-current release matrix.

### Earlier reliability slices

This repair pass preserves the existing renderer boundary and closes observed host/runtime defects:

- Markdown Quick Access export/open/print/fullscreen actions use localized SVG icon buttons, and
  exact ProseMirror whole-document selections normalize to editable content bounds.
- DOCX, Markdown, XLSX, PPTX, and PDF file/open, save, and fullscreen controls use the same
  format-neutral SVG chrome primitives with localized accessible names and no visible button text.
  DOCX anchors its File menu outside the horizontally scrollable tab row so Open, Save, and Save As
  remain visible and clickable instead of being clipped by the Ribbon. PPTX also removes the
  obsolete macOS Electron traffic-light inset so its File control starts at the embedded Ribbon's
  leading edge.
- PDF empty mounts stop waiting for a nonexistent pending file; late staged opens still dispatch to
  the mounted renderer. PDF recovery snapshots serialize pending live edits without changing the
  open source or triggering a download.
- XLSX standalone first Save obtains a browser file handle, writes and closes it, then reopens the
  resulting bytes before reporting success; later Save reuses the same handle. Recovery uses the
  same package writer against an in-memory clone, and empty conditional-format rule sets now mark
  metadata loaded.
- DOCX, Markdown, XLSX, PPTX, and PDF embedded saves use the lease-checked app-only local persistence
  protocol instead of sandbox-blocked anchors, iframe pickers, or the unavailable MCP Apps host
  download capability. Each renderer supplies its format-owned serialized bytes; the Broker accepts
  bounded ordered chunks, flushes a same-directory temporary file, and atomically renames it. An
  opened file binds normal Save to its exact path, new files receive Session-isolated output paths,
  Save As rebinds, PDF Export Copy does not, and `office_get_context.filePath` exposes the verified
  result. PPTX retains its editable Save As name; standalone Save continues to reuse its granted
  browser file handle.
- PPTX blank-slide operations explicitly request renderer refresh, so the reported slide count,
  thumbnails, active index, and subsequent object commands observe one updated deck.
- Recovery storage is session-scoped; ordinary creation is blank, exact reconnect cannot restore a
  different task's document, and newest-snapshot recovery is explicit only. Each of the five show
  tools grants a fresh view lease; a mounted iframe binds only from that lease-bearing result and
  cannot switch Session. Duplicate views cannot poll, acknowledge, transfer bytes, write recovery, or
  disconnect a healthy owner. Hidden editors retain one bounded command long-poll and suspend periodic recovery
  serialization. Nested viewport intersection becomes authoritative after its first observation, so
  visible-task editors that scroll offscreen suspend while responsive pinned-summary transitions
  that remain intersecting stay painted. Browser content-visibility lifecycle additionally pauses
  renderer-heavy work: DOCX/Markdown release document DOM, XLSX releases Univer Canvas backing
  stores, PDF cancels page/thumbnail renders and releases their canvases, and PPTX releases its Konva
  workspace, all without remounting the iframe or replacing the format-owned state/undo authority.
  All five formats expose monotonic persisted-state recovery versions, preventing unchanged dirty
  documents from repeated serialization. DOCX also flushes pending protected-table/textbox edits
  before releasing node views.
- DOCX, Markdown, and XLSX chrome now shrink to compact host panes; all five formats pass the shared
  280 px visibility/no-body-overflow gate, with wide Ribbon content scrolling internally.
- The bundled Skill now closes a new deliverable's first successful generation pass with exactly
  one canonical format-owned `*.document.save` transaction. Browser destination/permission prompts
  remain user-controlled, cancellation is not reported as delivery, and later edits are not saved
  implicitly unless requested or covered by the renderer's AutoSave preference.

## Current verification

These checks describe the current source tree and generated plugin. Format-local parity is complete;
the prior approved R6 evidence is no longer source-current after this responsive-host repair, so the
formal release gate intentionally fails closed. The last approved generated projection remains a
historical build input and cannot authorize a new package. Three formal recaptures passed DOCX,
Markdown, PPTX, PDF, and every visual comparison, but the busy local host measured XLSX bootstrap
p95 at 647.6, 536.8, and 609 ms against the unchanged 500 ms ceiling. Functional gates below pass;
release readiness remains pending a source-current formal capture on the canonical idle host profile.
R6-02 still closes fixed-poll command latency, R6-03 closes Markdown staged-load observability and
the canonical-large five-second budget, and R6-05 retains strict XLSX subphase evidence and fixed
resource/performance budgets.

- Typecheck: host bridge, five renderers, and MCP server pass the root typecheck graph.
- Operation registry foundation: fixture and real multi-format product manifests pass deterministic
  `--check`. The generated Manifest contains 347 operations: 103 DOCX, 22 Markdown, 123 XLSX, 74
  PPTX, and 25 PDF. All five retained-command producer mappings have no missing entry.
- Tests: the root workspaces execute 376 passing test files plus one environment-conditional skip,
  for 4,712 passing assertions and one skipped assertion. Workspace assertions are Operation
  Contract 24, Host Bridge 44, DOCX 1,680, Markdown 127, XLSX 1,967 plus one environment skip,
  PPTX 184, PDF 284, and MCP server 385. The separate DOCX engine suite adds 996 passing and one
  skipped assertion. These include wakeable-poll/startup-
  trace lifecycle, five-format isolated/exact Session recovery, exclusive view-lease enforcement,
  immutable iframe binding, schema-v4 release-evidence failure modes, installed-Skill/Manifest drift,
  and the empty-alias Product Manifest guard.
- Round trips: DOCX surgical fixtures including inserted/replaced-image persistence; XLSX
  formula/style/merge/chart/pivot/unknown-part retention; PPTX text/object/slide lifecycle with
  unknown-part retention; PDF text/image, annotation, generated-stamp, metadata/page, and AcroForm
  persistence.
- Product graph: no Electron, updater, GenOffice AI provider, agent-core, ai-search, or project-store dependency.
- Source boundary: Markdown, XLSX, all 80 permitted PPTX renderer files, and all 33 permitted PDF renderer files are restored. PPTX classifies all 104 pinned renderer paths as 61 byte-identical, 19 browser-host/product-boundary adapted, and 24 prohibited AI modules/assets. XLSX excludes exactly 20 `renderer/ai` files, 14 AI composer assets, and `i18n/strings-ai.ts`; the bundled optional Univer render-metrics token is locally inert and does not include the upstream telemetry module.
- Generated plugin: the artifact contains eleven Agent-visible tools, fifteen app-only transport tools,
  five resources, and self-contained DOCX/Markdown/XLSX/PPTX/PDF editor assets; the install-like real
  MCP smoke passes.
- Browser smoke: all five standalone entries have mount coverage. All five formats are active in the
  shared 25-scenario host-width matrix; PPTX and PDF retain their format-specific width suites. The
  complete 144-scenario Playwright run includes the cross-format icon-only chrome contract, the
  PPTX zero-leading-inset regression, the responsive pinned-summary paint regression,
  one-load offscreen release/resume and per-version recovery regressions for all five formats, the R6 metric tracer, four-phase committed Markdown load,
  one-shot four-phase XLSX cold start, canonical budgets, bounded-poll wakeup, format interaction,
  user/Agent state convergence, persistence, initial size-before-poll/fullscreen ordering, and one
  iframe mount across fullscreen entry and exit.
- PDF community evidence: 23 PDF test files / 283 tests pass. Eleven real-host scenarios cover user
  and staged open, shared delete/undo, typed text and image lifecycle, CJK/Korean/Arabic plus mixed
  color text, generated-stamp replacement/clear, save/reopen, and four host-width states. The
  25-operation Registry owns all retained mutation producers; public legacy names are rejected,
  staged byte routes remain internal, and approved R6-01 evidence generates `ready: true`.
- PPTX operation evidence: the Slides workspace passes four files / 169 tests. Seventy-four registry descriptors cover the retained mutation inventory with exact validation, shared renderer/history/recovery/package paths, canonical public dispatch, real-fixture host evidence, and a complete producer baseline. Approved R6-01 evidence generates `ready: true`.
- XLSX verification: the format test suite, community-renderer browser scenarios, Codex
  width/fullscreen matrix, and R6-05 aggregate/subphase startup tracer run against the same mounted App. The final
  command audit requires every exact replacement ID to exist in the XLSX catalog.
- Resource guardrails: generated resources measure DOCX 3,508,575 raw / 987,885 gzip, Markdown
  1,873,880 / 572,423, XLSX 12,818,734 / 4,804,389, PPTX 3,387,763 / 992,890, and PDF 6,631,726 /
  3,371,418 bytes. Raw ceilings are 3,650,000 / 2,500,000 / 21,000,000 / 4,000,000 / 7,000,000
  respectively. XLSX additionally caps the inflated initial executable entry at 11,000,000 bytes;
  optional modules remain compressed in the same HTML resource. PDF includes gzip-compressed PDFium WASM for retained browser content-stream
  editing; its edit fonts remain lazy external assets. Budgets are regression signals, never
  reasons to delete retained capabilities.
- Production dependency audit: as of 2026-08-28, `npm audit --omit=dev --audit-level=high` reports 45 high-severity transitive findings rooted in the pinned Univer dependency on vulnerable `nanoid`; npm reports no available fix. The operation-contract workspace and R2-01 through R2-110 integrations add no new external package to the pinned renderer dependency chain; R2-74 declares the already-installed `@univerjs/sheets` package directly because format-copy, column-width-copy, and outline actions consume its native mutation contracts.
