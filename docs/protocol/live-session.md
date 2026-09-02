# Live session protocol

- Status: Implemented five-format source contract; release projection pending source-current evidence
- Version: `0.3.28`
- Transport: MCP over stdio; MCP App calls proxied by the host
- UI resources: `ui://tandemfolio/editor.html`, `markdown.html`, `xlsx.html`, `pptx.html`, and `pdf.html`
- UI MIME type: `text/html;profile=mcp-app`

The format-owned architecture is documented in
[ADR 0004](../adr/0004-format-owned-operation-registries.md), the
[operation registry contract](operation-registry.md), and the
[implementation plan](../migration/operation-registry-plan.md). This file describes the implemented
revision-guarded transaction-only wire protocol.

## Purpose and authority

This file records what the current source contract implements. Under
[ADR 0003](../adr/0003-complete-community-renderers-and-mcp-parity.md), all five pinned renderer
producer baselines now map every retained state-changing command to typed MCP. R6-01 closes the
shared pinned-source visual, runtime, acknowledgement, memory, resource, package, smoke, license,
and repository gate. The current responsive-host repair invalidates the prior approved source
fingerprint, so the formal release gate remains fail-closed until a source-current capture passes the
fixed XLSX bootstrap budget; the last approved generated projection cannot authorize a new package.
R6-06 exposes that Registry through bounded summary pages and one-operation detail lookup rather
than returning every schema in one response. R6-07 adds caller request identity, exact revision
guarding, idempotent completed/in-flight replay, pre-dispatch typed validation, and an acknowledged
transaction result without changing renderer authority. R6-08 removes the old `operation` +
`arguments` compatibility Adapter after migrating repository callers and tests. The public tool
schema now requires the complete transaction envelope. R6-09 removes all thirty-six public-era
format-local aliases. Staged-file transport uses canonical format operations, so the Product
Manifest alias set is empty.

The user and Agent edit one mounted format renderer. The renderer owns document state, selection,
and undo/redo. The MCP broker owns session identity, one exclusive active view lease, connection
state, one queued/active command, the last acknowledged revision, the session-scoped request journal,
staged local bytes, opaque recovery transport, and the constrained opaque persistence sink from
[ADR 0011](../adr/0011-session-bound-local-document-persistence.md). [ADR 0010](../adr/0010-immutable-editor-view-leases-and-exact-session-resume.md)
defines the cross-format Session/view identity boundary.

The broker never parses or edits an Office document. Local recovery snapshots permitted by ADR 0002 are renderer-produced bytes and become editable only after they are loaded into a mounted renderer.

Only the format show tools declare UI resources. Context, file, capability, and mutation tools must not remount or reparent the iframe.

## Session shape

```ts
interface LiveSession {
  id: string
  format: 'docx' | 'markdown' | 'xlsx' | 'pptx' | 'pdf'
  revision: number
  connected: boolean
  fileName: string | null
  filePath: string | null
  dirty: boolean
  selection: Record<string, unknown> | null
  pending: QueuedCommand[]
}
```

`revision` starts at `0` and advances exactly once after a successful renderer acknowledgement. A negative acknowledgement clears the active command without advancing revision. At most one command may be pending or active.

## Lifecycle

```text
office_get_capabilities
        │
        ▼
office_create_session(resume = none | exact | latest)
        │
format show tool returns (sessionId, viewId) and renders one persistent iframe
        │
editor claims that lease with one immediate poll and current lightweight context
        ▼
connected=true; optional recovery bytes load in the renderer
        │
office_get_context → office_execute(requestId, baseRevision=current, operations=[one])
        │
enqueue wakes one bounded renderer poll → applies → checkpoints when dirty → acknowledges
        ▼
office_execute returns the transaction identity and result.revision=current+1
```

A mounted iframe's `(sessionId, viewId)` pair is immutable. It ignores later show notifications for a
different Session or view; switching documents requires teardown and a new mount. The Broker grants
the first polling view an exclusive lease and returns `editor_view_conflict` to another `viewId`.
Only the lease owner may poll, acknowledge, transfer Session-bound bytes, write recovery, or
disconnect, so tearing down a rejected duplicate cannot mark the healthy view offline.

The renderer registers the show-result handler before starting the MCP App handshake. Show input
alone does not bind or poll because it has no Broker-issued `viewId`; the result must supply both ids.
If `ui/initialize` is rejected while the host is changing responsive panes, the bridge closes that
transport and retries with a fresh MCP App after a bounded one-second attempt and 250 ms delay; the
first successful lease-bearing `office_editor_poll` remains the connection gate.

When the MCP App handshake completes, every format keeps the ext-apps automatic size lifecycle enabled. The SDK reports the iframe's initial width and content height to the host and continues reporting document/body changes through `ResizeObserver`. The renderer yields the initial animation frame so that size notification is sent before its first `office_editor_poll`; after that poll succeeds, it yields one stable layout frame before making the one-time fullscreen request. The host therefore has a mounted size and an active editor session before changing display mode.

The first poll uses `waitMs: 0`. XLSX first awaits its format-owned cold-start commit Promise and
attaches the one-shot R6-05 trace described below. After bootstrap, the shared DOCX/Markdown/XLSX/
PPTX/PDF Host Bridge keeps one `waitMs: 10000` request pending. Broker enqueue
resolves that waiter immediately;
an empty timeout re-arms it without delay. A transport failure alone uses a 500 ms retry. A newer
poll supersedes a lost older waiter, and editor disconnect releases the outstanding waiter. This is
the wakeable bounded-poll contract from [ADR 0006](../adr/0006-wakeable-live-session-command-delivery.md),
not a new public MCP tool or a second document authority.

Before the first nested viewport observation, the bridge uses page visibility as its activity
fallback. After the observation is known, intersection is authoritative: a visible task whose old
editor iframe has left the viewport becomes inactive, while a responsive pinned summary that
transiently reports `document.visibilityState === "hidden"` remains active when its editor still
intersects. Inactivity checkpoints a dirty session once, stops the two-second checkpoint timer, and
suppresses root paint. The root otherwise uses `content-visibility: auto`; its
`contentvisibilityautostatechange` event pauses or resumes renderer-heavy work at the browser's
pre-render boundary. DOCX and Markdown unmount only their document DOM while retaining their TipTap
editor objects and native undo state. DOCX first flushes pending protected-table/textbox edits into
the main ProseMirror document. XLSX zeros its Univer Canvas backing stores and restores them into the
same workbook runtime; PDF clears visible page/thumbnail render sets, cancels page render tasks, and
keeps the same loaded `PDFDocumentProxy`; PPTX releases its mounted workspace Canvas nodes and
reconstructs them from the same `BrowserPresentation` and React state. One bounded long poll remains
active, and no activity transition remounts the iframe or replaces native undo state.

XLSX additionally tracks one workbook-readiness Promise from file selection through parse, initial
visible/preload installation, and the next committed Univer frame. A command arriving during open
waits on that Promise, and a successful XLSX acknowledgement is emitted only after native
selection/history state has settled. This replaces timing luck from the former 500 ms cadence with
an explicit format-owned readiness boundary.

Host context is authoritative for display mode. If new host context arrives while a display-mode request is pending, the request's later result is stale and cannot overwrite the newer mode or availability state.

## Public Agent tools

### `office_get_capabilities`

Input:

```ts
{
  format?: 'docx' | 'markdown' | 'xlsx' | 'pptx' | 'pdf'
  view?: 'summary' | 'detail'
  operation?: string
  sessionId?: string
  family?: string
  limit?: number
  cursor?: string | null
}
```

`format` defaults to `docx`; `view` defaults to `summary`. Summary discovery returns release
readiness, display/local-file metadata, and at most twenty schema-free operation summaries. Each
summary contains only `id`, `family`, `summary`, `risk`, `context`, and `effects`, plus
`availability` when `sessionId` is supplied. `family` filters before pagination. `cursor` is the
last canonical operation id from the preceding page, and the response returns `{ limit, total,
nextCursor }`. Every maximum summary page is regression-gated below 8 KiB.

Detail discovery requires one exact canonical `visibility: agent` operation id and returns its
complete generated descriptor, including `inputSchema` and `outputSchema`. It rejects family,
cursor, and limit selectors. Compatibility aliases, internal staged operations, cross-format ids,
and invalid cursors return `invalid_arguments`. Generated descriptors are capped at 64 KiB.

When `sessionId` is present, discovery does not silently remove unavailable operations. It projects
`{ available, reason }`, where reason is `format_mismatch`, `editor_offline`,
`selection_required`, or `null`. Callers should use summary to choose an id, fetch detail for that
one id, then execute only after fresh context confirms the current revision. `ready` remains the
source-fingerprinted release gate defined by ADR 0005, not operation availability.

### `office_create_session`

Input:

```ts
{
  format?: 'docx' | 'markdown' | 'xlsx' | 'pptx' | 'pdf'
  resume?: 'none' | 'exact' | 'latest'
  sessionId?: string
}
```

`format` defaults to `docx` and `resume` defaults to `none`, so ordinary creation always returns an
isolated blank Session. `resume: "exact"` requires `sessionId`: it reuses that same in-memory Session
or stages only its exact `(format, sessionId)` recovery snapshot. A missing or format-mismatched exact
handle fails instead of selecting another document. `resume: "latest"` is explicit cross-session
disaster recovery and accepts no `sessionId`. The result includes `recoveryAvailable` and `reused`,
but recovery bytes are not returned to the Agent.

### Format show tools

Input: `{ "sessionId": "<uuid>" }`.

Use `office_show_editor` for DOCX, `office_show_markdown_editor` for Markdown, `office_show_xlsx_editor` for XLSX, `office_show_pptx_editor` for PPTX, and `office_show_pdf_editor` for PDF. Each is associated only with its matching resource and returns `{ sessionId, viewId, format, revision }`. The fresh `viewId` is the candidate lease for that one mount. After initial size reporting and the first successful lease-bearing session poll, the renderer requests fullscreen once when supported and provides a manual inline/fullscreen button. A show result does not prove the handshake completed; wait for `office_get_context.connected`.

For a follow-up edit, inspect the recorded exact `sessionId` first. If it is already connected, issue
data operations directly and do not call a show tool again. Show a known Session only when it is
disconnected. After Broker restart, recreate the same handle with `resume: "exact"` before showing it
once; never substitute `latest` for a known document.

### `office_get_context`

Input: `{ "sessionId": "<uuid>" }`.

Returns the last broker-acknowledged lightweight session state. `filePath` is the absolute local
target of the last bound Save, or `null` before persistence; an Export Copy does not change it. DOCX
selection currently uses zero-based ProseMirror `{ from, to, empty }` positions. Complete document
bytes and undo history are never model context.

### `office_execute`

Canonical input (the operation id must match the session format):

```json
{
  "sessionId": "<uuid>",
  "baseRevision": 4,
  "requestId": "<caller-generated-id>",
  "operations": [
    {
      "id": "markdown.text.replace_selection",
      "arguments": { "text": "replacement" }
    }
  ]
}
```

`requestId` is 1–128 characters and is scoped to the live session. `operations` currently contains
exactly one canonical Agent-visible Registry id: no format yet declares a multi-operation native
atomic/undo grouping. The Broker validates the structural envelope, request identity, connection,
exact revision, canonical id, exact input schema, current required context, and one-command rule
before enqueue. Success is renderer acknowledgement, not queue acceptance:

```json
{
  "ok": true,
  "transaction": {
    "transactionId": "<uuid>",
    "requestId": "<caller-generated-id>",
    "baseRevision": 4
  },
  "result": {
    "revision": 5,
    "operations": [
      {
        "id": "markdown.text.replace_selection",
        "result": {}
      }
    ]
  }
}
```

One positive acknowledgement advances the shared revision exactly once. A renderer failure returns
an MCP error with its typed code and does not advance revision.

An exact replay uses structural JSON identity, so object-key order does not matter. A replay of an
in-flight request joins the same execution; a replay of a completed request returns the same
response and `transactionId` without dispatch, even though its original `baseRevision` is now
stale. Reusing the same request id with a changed base revision, operation, or arguments returns
`request_reused`. Validation failures before enqueue do not reserve the request id.

The bundled TandemFolio Skill treats a request to create a new deliverable as authorization to
invoke the matching canonical `*.document.save` operation once after the first complete successful
generation pass. It refreshes context, uses a new request id for that persistence transaction, then
verifies the resulting `filePath`. In an embedded editor this uses the app-only local persistence
protocol; standalone browser mode may still require a file-picker permission. This client convention
does not turn a recovery snapshot into a saved delivery and does not automatically save every later
edit or an existing document opened only for inspection/editing.

`command_timeout` limits one tool caller's wait; it is not a final transaction outcome. The Broker
keeps the underlying command, request record, and any staged local bytes alive until the renderer
acknowledges or rejects it. Retry only the exact same envelope and request id. A later exact replay
returns the final acknowledged response without redispatch.

The retired `{ sessionId, baseRevision, operation, arguments }` form fails MCP input validation
before renderer dispatch because `requestId` and `operations` are required. `office_execute`
accepts exact canonical Agent-visible ids only. The five retained `open_local_file` compatibility
aliases are internal staged-file transports and are not reachable through this tool.

### `office_open_local_file`

Input:

```json
{ "sessionId": "<uuid>", "baseRevision": 0, "path": "/absolute/path/document.docx" }
```

The local server validates the absolute path and session-format extension and stages bytes under a session-scoped opaque id. Markdown, DOCX, XLSX, PPTX, and PDF queue their format-owned internal canonical `*.document.load_staged` operation. The shared host bridge recognizes canonical ids and their legacy transport aliases, reads bounded chunks, and supplies hydrated bytes to the renderer's existing format-owned load pipeline. Internal load operations are not advertised and cannot be called through `office_execute`. The tool waits for the same acknowledgement contract as `office_execute`. After acknowledgement the Broker binds future Save operations from that Session to the exact opened path. The path is visible as `office_get_context.session.filePath`; bytes are never placed in model context.

## App-only transport tools

These tools use `_meta.ui.visibility: ["app"]`:

- `office_editor_poll` requires app-only `sessionId` and `viewId`, with optional lightweight file,
  dirty, selection, startup-trace, and wait fields. `waitMs` defaults to `0` and is bounded to
  `0..10000`. Its first call claims the Session's view lease, synchronizes lightweight context, and
  atomically delivers its command immediately or by waking the Session's single pending waiter.
  `startupTrace` is accepted only as the strict XLSX shape below; the Broker validates and discards
  it rather than storing document or telemetry state.
- `office_editor_acknowledge` carries the same `sessionId` and `viewId`, then accepts `ok: true` plus
  the next revision/context, or `ok: false` plus `unsupported_operation`, `invalid_arguments`, or
  `execution_failed`. Successful acknowledgements may include app-only hydration, execution, and
  optional trace timing; the host combines these renderer-local durations with queue/poll and
  acknowledgement timestamps without exporting telemetry. The trace is a strict
  operation-discriminated object rather than an open phase map.

For internal `markdown.document.load_staged`, the optional trace has this exact shape:

```json
{
  "operation": "markdown.document.load_staged",
  "phases": {
    "decodeMs": 0,
    "parseMs": 0,
    "tiptapStateInstallMs": 0,
    "reactCommitMs": 0
  }
}
```

`decodeMs` covers staged bytes to UTF-8; `parseMs` covers document-envelope normalization and
Markdown-to-TipTap JSON; `tiptapStateInstallMs` covers mounted ProseMirror state/view installation;
and `reactCommitMs` covers the final filename/frontmatter/dirty/status React layout commit. Every
phase is a bounded non-negative millisecond duration. A positive staged-load acknowledgement is
sent only after that commit is observable. The trace contains no content, path, or remote telemetry.

The first successful XLSX poll may additionally carry this exact one-shot trace:

```json
{
  "operation": "xlsx.editor.cold_start",
  "phases": {
    "bootstrapMs": 0,
    "univerCreateMs": 0,
    "worksheetInstallMs": 0,
    "firstCommitMs": 0
  },
  "bootstrapPhases": {
    "resourceReceiveMs": 0,
    "moduleGraphReadyMs": 0,
    "reactMountMs": 0
  }
}
```

`bootstrapMs` covers iframe navigation through entry into the Univer-creation effect;
`univerCreateMs` covers retained preset registration; `worksheetInstallMs` covers initial workbook
and worksheet installation; and `firstCommitMs` covers remaining integration through an active
workbook, active sheet, and observable canvas. Durations are bounded and non-negative, and their
sum cannot exceed the host cold-start measurement. The three `bootstrapPhases` partition
`bootstrapMs` into HTML resource receipt, initial entry-module graph readiness, and entry-to-React
mount; their sum equals the aggregate within measurement precision. A transport failure retains
the trace for the retry; the first successful poll consumes it. Later polls omit it. See
[ADR 0008](../adr/0008-traced-bounded-xlsx-cold-start.md) and
[ADR 0009](../adr/0009-self-contained-deferred-xlsx-bootstrap.md).

- `office_editor_disconnect` marks the session offline on teardown only when the supplied `viewId`
  owns that Session's active lease.
- `office_editor_read_file_chunk` reads at most 256 KiB from a session-owned staged file.
- `office_editor_read_local_asset_chunk` reads at most 256 KiB from a session-bound Markdown asset root; only validated PNG/JPEG/GIF bytes are admitted.
- `office_editor_read_font_chunk` lazily reads allowlisted bundled font assets without inflating the initial editor HTML.
- `office_editor_begin_recovery`, `office_editor_write_recovery_chunk`, and `office_editor_commit_recovery` atomically persist renderer-produced recovery bytes.
- `office_editor_begin_document_save`, `office_editor_write_document_save_chunk`,
  `office_editor_commit_document_save`, and `office_editor_abort_document_save` transfer at most
  256 MiB of renderer-produced document bytes. Begin validates format extension, size, mode, and the
  active view lease; writes require exact ordered offsets; commit flushes and atomically renames the
  temporary file. `save` overwrites the Session binding, `save-as` writes and rebinds a collision-safe
  file, and `export-copy` writes without rebinding.

Every Session-bound app-only endpoint carries the show result's `viewId`: poll, acknowledgement,
disconnect, document persistence, recovery upload, staged-file reads, and local-asset reads all
validate the same lease.
Bundled-font reads are format resources rather than Session state and therefore carry no lease.

Dirty Agent mutations checkpoint before successful acknowledgement. All five renderers expose a
monotonic persisted-state recovery version to the Host Bridge. A successfully stored version,
including one carried by an Agent acknowledgement, is not serialized again until the renderer
reports a newer version. DOCX includes direct UI document-design state and pending protected
sub-editor input; Markdown tracks TipTap/frontmatter changes; XLSX tracks its workbook edit journal;
PPTX tracks retained history; PDF tracks page, annotation, form, metadata, text, and image edits.

Recovery is bounded to 256 MiB per snapshot, retains at most one current snapshot per
format/session pair, expires after seven days, and removes the current session's snapshot after an
acknowledged explicit save. `resume: "exact"` stages only the named Session's snapshot;
`resume: "latest"` explicitly selects the newest unexpired snapshot for the requested format and is
never used implicitly.
Temporary uploads are renamed atomically and incomplete temporary files are removed.

For a newly generated document, the default bound target is
`~/Documents/TandemFolio/<format>-<session-hash>/<file-name>`. Set
`TANDEMFOLIO_OUTPUT_DIR` on the MCP server to choose another output root. An explicitly opened
local file is the bound target for normal Save. Bindings are persisted by exact Session id under the
configured state directory so a broker restart cannot silently select another task's artifact.

## Implemented format operations

### DOCX

| Operation                                    | Arguments                                                                         | Effect                                                                            |
| -------------------------------------------- | --------------------------------------------------------------------------------- | --------------------------------------------------------------------------------- |
| `docx.block.delete`                          | `{ target }`                                                                      | Deletes matching blocks through one native transaction.                           |
| `docx.block.move`                            | `{ blockIndexes, afterBlockIndex }`                                               | Moves blocks in document order through one native transaction.                    |
| `docx.image.insert`                          | `{ path, afterBlockIndex, widthPx, heightPx, alignment }`                         | Stages and inserts a bounded local image at a stable block boundary.              |
| `docx.image.insert_staged` internal          | `{ blobId, name, size, data, afterBlockIndex, widthPx, heightPx, alignment }`     | Hydrates validated image bytes into one native image node.                        |
| `docx.image.remove`                          | `{ imageBlockIndex }`                                                             | Removes one exact image while preserving a non-empty document.                    |
| `docx.image.replace`                         | `{ path, imageBlockIndex, widthPx, heightPx }`                                    | Replaces one indexed image with explicit final geometry.                          |
| `docx.image.replace_staged` internal         | `{ blobId, name, size, data, imageBlockIndex, widthPx, heightPx }`                | Hydrates replacement bytes into the original/native image route.                  |
| `docx.image.set_crop`                        | `{ imageBlockIndex, left, top, right, bottom }`                                   | Sets or resets non-destructive source crop insets for one image.                  |
| `docx.image.set_margin_position`             | `{ imageBlockIndex, horizontal, vertical }`                                       | Sets one original image to a named margin-relative preset.                        |
| `docx.image.set_offset_position`             | `{ imageBlockIndex, wrap, offsetXEmu, offsetYEmu }`                               | Sets explicit floating wrap and signed EMU offsets for one image.                 |
| `docx.image.set_transform`                   | `{ imageBlockIndex, rotationDegrees, flipHorizontal, flipVertical }`              | Sets complete final rotation and mirror state for one image.                      |
| `docx.image.set_wrap`                        | `{ imageBlockIndex, wrap }`                                                       | Sets retained floating wrap or explicit inline state.                             |
| `docx.image.update`                          | `{ target, properties, fields }`                                                  | Updates image size/alignment through the retained image command.                  |
| `docx.list.apply`                            | `{ target, kind }`                                                                | Applies bullet/ordered formatting through one native transaction.                 |
| `docx.list.remove`                           | `{ target }`                                                                      | Converts matching list items to body paragraphs in one transaction.               |
| `docx.list.set_level`                        | `{ target, level }`                                                               | Sets an absolute `0..8` list level through shared bounds.                         |
| `docx.list.apply_preset`                     | `{ target, levels }`                                                              | Creates and applies a bounded 1–9-level numbering definition.                     |
| `docx.list.restart`                          | `{ blockIndex, start }`                                                           | Restarts one stable list anchor from an explicit bounded value.                   |
| `docx.list.continue`                         | `{ blockIndex, previousBlockIndex }`                                              | Continues one list from an explicitly identified earlier list block.              |
| `docx.history.undo`                          | `{}`                                                                              | Undoes one available entry in the mounted native TipTap history.                  |
| `docx.history.redo`                          | `{}`                                                                              | Redoes one available entry in the mounted native TipTap history.                  |
| `docx.paragraph.set_style`                   | `{ target, style, fields }`                                                       | Applies bounded paragraph geometry, line modes, and tab stops.                    |
| `docx.paragraph.set_direction`               | `{ target, direction }`                                                           | Sets LTR/RTL with shared logical alignment-flip semantics.                        |
| `docx.text.insert`                           | `{ "text": string (1..65536 Unicode chars) }`                                     | Inserts at the active caret/selection; the symbol palette shares this action.     |
| `docx.text.replace_selection`                | `{ "text": string }`                                                              | Replaces the active selection.                                                    |
| `docx.text.replace_all`                      | `{ containsText, replaceText, matchCase? }`                                       | Replaces matching text in one native whole-document transaction.                  |
| `docx.text.clear_character_format`           | `{ range }`                                                                       | Clears selected marks through the retained Ribbon's shared helper.                |
| `docx.text.set_character_format`             | `{ range, format, fields }`                                                       | Sets exact selected marks/font/size/color through a shared helper.                |
| `docx.text.set_character_style`              | `{ range, styleId }`                                                              | Applies a current-document character style through the shared helper.             |
| `docx.text.set_link`                         | `{ range, href, text }`                                                           | Inserts, updates, or removes one bounded exact-range hyperlink.                   |
| `docx.bookmark.set`                          | `{ blockIndex, name, enabled }`                                                   | Sets one bounded unique top-level bookmark to an explicit final state.            |
| `docx.note.insert`                           | `{ range, kind, noteId, text }`                                                   | Inserts an Undo-owned bounded footnote/endnote marker and body at an exact range. |
| `docx.note.update`                           | `{ kind, noteId, text }`                                                          | Sets bounded final note text by stable ID through the same Undo-owned atoms.      |
| `docx.note.delete`                           | `{ kind, noteId }`                                                                | Deletes one stable note and renumbers same-kind references in one Undo unit.      |
| `docx.source.upsert`                         | `{ source }`                                                                      | Adds or replaces a bounded source by stable tag through an Undo-owned snapshot.   |
| `docx.citation.insert`                       | `{ range, sourceTag, displayText }`                                               | Inserts bounded display text for one existing source at an exact inline range.    |
| `docx.bibliography.insert`                   | `{ afterBlockIndex, heading, entries }`                                           | Inserts explicit source-backed bibliography blocks after one stable boundary.     |
| `docx.caption.insert`                        | `{ afterBlockIndex, label, number, text }`                                        | Inserts one explicit dirty SEQ caption after a stable block boundary.             |
| `docx.index.mark`                            | `{ range, term }`                                                                 | Attaches one bounded hidden XE marker to an exact inline range.                   |
| `docx.index.insert`                          | `{ afterBlockIndex, label, terms }`                                               | Inserts one explicit cached INDEX field after a stable block boundary.            |
| `docx.comment.add`                           | `{ range, comment }`                                                              | Adds explicit comment metadata and exact-range anchors in one Undo unit.          |
| `docx.comment.reply`                         | `{ parentId, comment }`                                                           | Adds an explicit reply on every stable parent anchor in one Undo unit.            |
| `docx.comment.set_resolved`                  | `{ id, resolved }`                                                                | Sets parent/reply resolved state to one explicit final boolean.                   |
| `docx.cross_reference.insert`                | `{ range, bookmarkName, displayText }`                                            | Inserts one exact REF field targeting an existing bounded bookmark.               |
| `docx.field.insert`                          | `{ range, instruction, displayText }`                                             | Inserts one finite generic field with explicit cached display text.               |
| `docx.field.update`                          | `{ updates: [{ range, instruction, displayText }] }`                              | Updates bounded exact generic-field caches in one native Undo transaction.        |
| `docx.text.set_style`                        | `{ target, style, fields }`                                                       | Applies or clears masked text formatting in one native transaction.               |
| `docx.text.transform_case`                   | `{ range, mode }`                                                                 | Applies one of four case modes through a shared mapped transaction.               |
| `docx.paragraph.set_heading_level`           | `{ target, level }`                                                               | Promotes/demotes matching blocks through one native transaction.                  |
| `docx.toc.insert`                            | `{ afterBlockIndex }`                                                             | Inserts TOC field blocks generated from current headings.                         |
| `docx.toc.refresh`                           | `{ tocBlockIndex, entries }`                                                      | Replaces one exact TOC field region with bounded explicit final entries.          |
| `docx.table.insert`                          | `{ afterBlockIndex, rows, columns }`                                              | Inserts a bounded native table at a stable top-level boundary.                    |
| `docx.table.delete`                          | `{ tableBlockIndex }`                                                             | Deletes one explicitly indexed top-level native table.                            |
| `docx.table.insert_rows`                     | `{ tableBlockIndex, rowIndex, count }`                                            | Inserts bounded rows at an explicit rowspan-aware table boundary.                 |
| `docx.table.delete_rows`                     | `{ tableBlockIndex, rowIndex, count }`                                            | Deletes a bounded proper subset of rows from an explicit index.                   |
| `docx.table.insert_columns`                  | `{ tableBlockIndex, columnIndex, count }`                                         | Inserts bounded columns at an explicit span-aware table boundary.                 |
| `docx.table.delete_columns`                  | `{ tableBlockIndex, columnIndex, count }`                                         | Deletes a bounded proper subset of logical columns.                               |
| `docx.table.merge_cells`                     | `{ tableBlockIndex, topRow, leftColumn, bottomRow, rightColumn }`                 | Merges one exact half-open logical-cell rectangle.                                |
| `docx.table.split_cell`                      | `{ tableBlockIndex, rowIndex, columnIndex }`                                      | Splits the merged cell covering one exact logical coordinate.                     |
| `docx.table.set_cell_format`                 | `{ tableBlockIndex, topRow, leftColumn, bottomRow, rightColumn, format, fields }` | Sets masked fill/alignment over one exact cell rectangle.                         |
| `docx.table.set_cell_borders`                | `{ tableBlockIndex, topRow, leftColumn, bottomRow, rightColumn, mode, border }`   | Applies one bounded edge policy over an exact cell rectangle.                     |
| `docx.table.set_style`                       | `{ tableBlockIndex, styleId }`                                                    | Sets or clears one current-document table style identity.                         |
| `docx.table.set_row_height`                  | `{ tableBlockIndex, rowIndex, count, heightTwips }`                               | Sets or clears height over one bounded physical-row interval.                     |
| `docx.table.set_column_widths`               | `{ tableBlockIndex, widthsPx }`                                                   | Replaces one complete bounded logical-column width vector.                        |
| `docx.document.insert_page_break`            | `{ afterBlockIndex }`                                                             | Inserts a native page-break paragraph at a stable block boundary.                 |
| `docx.section.insert_break`                  | `{ afterBlockIndex, startType }`                                                  | Inserts an Undo-owned section break with a finite start type.                     |
| `docx.section.set_orientation`               | `{ sectionIndex, orientation }`                                                   | Sets portrait/landscape through the Undo-owned section journal.                   |
| `docx.section.set_margins`                   | `{ sectionIndex, margins }`                                                       | Sets four bounded twip margins while preserving positive body area.               |
| `docx.section.set_page_size`                 | `{ sectionIndex, widthTwips, heightTwips }`                                       | Sets bounded actual page axes and derives orientation.                            |
| `docx.shape.insert`                          | `{ afterBlockIndex, preset, widthEmu, heightEmu }`                                | Inserts one of 104 filled presets at a stable block boundary.                     |
| `docx.line.insert`                           | `{ afterBlockIndex, kind, widthEmu, heightEmu }`                                  | Inserts one of five stroke-only line/connector kinds.                             |
| `docx.textbox.insert`                        | `{ afterBlockIndex, widthEmu, heightEmu }`                                        | Inserts an empty textbox at a stable block boundary.                              |
| `docx.textbox.set_content`                   | `{ objectBlockIndex, textboxIndex, paragraphs, heightPx }`                        | Replaces one exact nested textbox with bounded rich paragraphs/final height.      |
| `docx.chart.insert`                          | `{ afterBlockIndex, kind, title, categories, series, widthPx, heightPx }`         | Inserts one bounded data-backed chart with explicit final extent.                 |
| `docx.chart.update`                          | `{ chartBlockIndex, patch, fields }`                                              | Updates masked existing title/category/series cache slots.                        |
| `docx.equation.insert`                       | `{ placement, latex, afterBlockIndex, from, to }`                                 | Inserts bounded LaTeX at an exact block boundary or inline range.                 |
| `docx.equation.update`                       | `{ placement, mode, latex, tokens, equationBlockIndex, from, to }`                | Rebuilds exact LaTeX or preserves one block equation's token shape.               |
| `docx.object.set_size`                       | `{ objectBlockIndex, widthPx, heightPx }`                                         | Sets one shape, line, textbox, or chart to an explicit bounded size.              |
| `docx.object.set_offset_position`            | `{ objectBlockIndex, wrap, offsetXEmu, offsetYEmu }`                              | Sets one textbox-backed drawing to an explicit floating anchor.                   |
| `docx.object.set_style`                      | `{ objectBlockIndex, style, fields }`                                             | Sets masked nullable fill/outline final state on one drawing.                     |
| `docx.object.remove`                         | `{ objectBlockIndex }`                                                            | Removes one exact drawing, chart, diagram, or block equation.                     |
| `docx.section.set_columns`                   | `{ sectionIndex, count, spacingTwips }`                                           | Sets exact bounded column count and gap with positive text width.                 |
| `docx.section.set_page_border`               | `{ sectionIndex, enabled }`                                                       | Sets explicit page-border state through the section journal.                      |
| `docx.section.set_different_first_page`      | `{ sectionIndex, enabled }`                                                       | Sets indexed `w:titlePg` state through the section journal.                       |
| `docx.document.set_different_odd_even_pages` | `{ enabled }`                                                                     | Sets document-wide odd/even header-footer variants through native history.        |
| `docx.section.set_page_numbering`            | `{ sectionIndex, format, start }`                                                 | Sets bounded page-number format and nullable start through the section journal.   |
| `docx.header_footer.set_text`                | `{ sectionIndex, kind, variant, text }`                                           | Sets explicit header/footer variant text through the Undo-owned content journal.  |
| `docx.header_footer.set_page_number`         | `{ sectionIndex, kind, variant, enabled, alignment }`                             | Sets or removes canonical PAGE placement through the content journal.             |
| `docx.header_footer.set_paragraphs`          | `{ sectionIndex, kind, variant, paragraphs }`                                     | Sets bounded rich paragraphs with styled text and PAGE/NUMPAGES tokens.           |
| `docx.document.save`                         | `{}`                                                                              | Runs the DOCX surgical-save pipeline and browser host save behavior.              |
| `docx.document.load_staged` internal         | `{ blobId, name, size, data }`                                                    | Loads host-hydrated DOCX bytes without a recovery checkpoint.                     |

All ninety-one DOCX operations in the table are generated-registry operations. R6-09 retires the
legacy `insert_text`, `replace_selection`, and `save` aliases. `open_local_file` remains only as the
internal staged-load transport alias and stays outside public discovery and transaction dispatch.
The staged-load descriptor is also hidden from Agent
discovery and rejected through direct `office_execute`; `office_open_local_file` is its public
transport entry. Public `docx.image.insert` and `docx.image.replace` are transformed by the Broker
into their hidden staged counterparts; binary bytes are chunk-hydrated only inside the mounted
renderer and never appear in public execution arguments. A successful save returns
`{ "saved": true, "fileName": string }`, acknowledges a clean document, and clears the stored DOCX
recovery snapshot; failure returns `execution_failed`. Replace-all is case-sensitive unless
`matchCase` is `false`; success returns its summary plus matched/changed block counts and protected
or tracked-deleted skip counts, checkpoints dirty recovery, and advances one revision. Empty search
text is rejected without mutation. Heading-level changes accept the existing DOCX target fields
(`nodeType`, `headingLevel`, `containsText`, `matchCase`, `blockIndexes`, and `scope`) and a level
from 0 through 6; success returns summary plus matched/changed/protected counts, checkpoints dirty
recovery, and advances one revision. The Broker rejects invalid enum, item, and numeric-bound input,
while a target with no condition is rejected at the renderer without mutation. Block deletion uses
the same DOCX target schema, returns
matched/changed/protected/tracked-deleted counts, and leaves one empty paragraph when every block is
deleted. It is high risk, undoable, recovery-checkpointed, and revision guarded. Conditionless
targets are rejected without mutation. Both list operations reuse the same target schema. List
application accepts only `bullet | ordered`, lazily allocates a renderer numbering definition when
a block actually changes, preserves text, and treats the same kind as a no-op without replacing its
numId. List removal converts only matching list items to body paragraphs. Both return
matched/changed/protected counts and remain medium risk with one native undo. Text and paragraph
style updates use non-empty field masks; nullable style values clear only the named properties.
Block moves require at least one source index, preserve source order, and are high risk. Image
updates preserve proportional scaling when only one dimension is supplied. TOC insertion reads the
mounted heading state and creates the retained field-block representation. The former
`batch_update` capability and renderer route have been removed; an attempted call is unsupported
rather than dispatched through a compatibility executor.

### Markdown

Markdown owns 22 generated Registry operations: twenty public text/selection/mark, block/list,
table/divider, image, frontmatter, history, persistence/output/preference operations plus two
internal staged document/image routes. Exact schemas are returned by `office_get_capabilities`;
the format-owned retained-command mapping requires every descriptor to belong to a retained UI,
typed ingress, or native-input family.

R6-09 retires the legacy `save`, `insert_text`, and `replace_selection` aliases; capability
discovery and renderer execution accept only canonical public ids. `open_local_file` remains only
as the internal staged-load transport alias. A successful save returns
`{ "saved": true, "fileName": string }`; cancellation or
failure returns `execution_failed` without a positive acknowledgement.

Internal `markdown.document.load_staged` additionally publishes the four-stage app-only trace
defined above. The same operation still installs the retained TipTap state, awaits the final React
commit, hydrates display-only local assets, and advances the shared session revision exactly once.

Context includes the active TipTap selection, bounded selected text, and active block type.
Frontmatter, original line endings, BOM state, trailing-newline behavior, browser/MCP local-image
hydration, and DOCX export are handled by format-owned layers. The complete mapping and evidence
are in [`../migration/markdown-capability-inventory.md`](../migration/markdown-capability-inventory.md).

### XLSX

| Operation                                  | Arguments                                                                             | Effect                                                                                                                              |
| ------------------------------------------ | ------------------------------------------------------------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------- |
| `xlsx.cell.set_value`                      | `{ "sheet": string, "address": string, "value": scalar }`                             | Changes one cell in the named worksheet.                                                                                            |
| `xlsx.range.set_values`                    | `{ "sheet": string, "range": string, "values": scalar[][] }`                          | Writes a bounded non-empty scalar matrix and activates the target range.                                                            |
| `xlsx.range.set_text_style`                | `{ "sheet", "range", "style", "fields" }`                                             | Explicitly sets bold, italic, underline, or strike fields as one Univer range mutation.                                             |
| `xlsx.range.set_alignment`                 | `{ "sheet", "range", "alignment", "fields" }`                                         | Explicitly sets masked horizontal, vertical, wrap, indent, and rotation fields in one mutation.                                     |
| `xlsx.range.set_font`                      | `{ "sheet", "range", "font", "fields" }`                                              | Explicitly sets or clears masked font family, size, and color fields in one mutation.                                               |
| `xlsx.range.set_fill`                      | `{ "sheet", "range", "color" }`                                                       | Sets or clears a bounded `#RRGGBB` range fill.                                                                                      |
| `xlsx.range.set_border`                    | `{ "sheet", "range", "border" }`                                                      | Sets or clears an exact border preset through Univer's native range border route.                                                   |
| `xlsx.range.apply_cell_style`              | `{ "sheet", "range", "preset" }`                                                      | Applies one bounded built-in cell-style preset as one native range mutation.                                                        |
| `xlsx.range.set_number_format`             | `{ "sheet", "range", "pattern" }`                                                     | Assigns one explicit 1–255-character number-format pattern through the native range route.                                          |
| `xlsx.range.merge`                         | `{ "sheet", "range", "mode" }`                                                        | Applies `cells`, `across`, `center`, or `unmerge`; center preserves the retained two-step UI behavior.                              |
| `xlsx.range.clear`                         | `{ "sheet", "range", "scope" }`                                                       | Clears `contents`, `formats`, or `all` through the native range facade.                                                             |
| `xlsx.range.copy_values`                   | `{ "sourceSheet", "sourceRange", "destinationSheet", "destinationRange" }`            | Copies computed scalar values between equal-shape ranges of at most 20,000 cells through one native undoable write.                 |
| `xlsx.range.copy_formulas`                 | `{ "sourceSheet", "sourceRange", "destinationSheet", "destinationRange" }`            | Copies formulas with translated references and scalar cells between equal-shape ranges through one native undoable write.           |
| `xlsx.range.copy_formats`                  | `{ "sourceSheet", "sourceRange", "destinationSheet", "destinationRange" }`            | Replaces cell formats between equal-shape ranges through one native Undo unit without changing destination values or formulas.      |
| `xlsx.range.copy_without_borders`          | `{ "sourceSheet", "sourceRange", "destinationSheet", "destinationRange" }`            | Copies cells and non-border formats between bounded equal-shape ranges while retaining destination borders in one native Undo unit. |
| `xlsx.range.fill`                          | `{ "sheet", "range", "direction" }`                                                   | Fills `down` or `right` after activating an explicit range with a destination row or column.                                        |
| `xlsx.range.sort`                          | `{ "sheet", "range", "direction" }`                                                   | Sorts an explicit range by its first column in ascending or descending order.                                                       |
| `xlsx.range.sort_custom`                   | `{ "sheet", "range", "keys", "hasHeader" }`                                           | Sorts by unique in-range A1 column keys with an explicit header flag.                                                               |
| `xlsx.range.remove_duplicates`             | `{ "sheet", "range", "hasHeader" }`                                                   | Removes duplicate rows, keeps the first occurrence, and reports the removed count.                                                  |
| `xlsx.range.set_filter`                    | `{ "sheet", "range", "enabled" }`                                                     | Sets the explicit final AutoFilter state for a range through native filter commands and shared Undo.                                |
| `xlsx.range.clear_filter_criteria`         | `{ "sheet", "range" }`                                                                | Clears all criteria while retaining the exact AutoFilter range through native Undo.                                                 |
| `xlsx.range.set_filter_values`             | `{ "sheet", "range", "column", "values", "includeBlank" }`                            | Sets a bounded value-list criterion on one absolute filter column.                                                                  |
| `xlsx.range.set_custom_filter`             | `{ "sheet", "range", "column", "conjunction", "conditions" }`                         | Sets one or two bounded custom conditions on one absolute filter column.                                                            |
| `xlsx.formula.insert_aggregate`            | `{ "sheet", "range", "function" }`                                                    | Inserts one bounded aggregate formula below each selected column after the streaming safety gate.                                   |
| `xlsx.history.undo`                        | `{}`                                                                                  | Undoes the latest mounted Univer history entry; fails when none is available.                                                       |
| `xlsx.history.redo`                        | `{}`                                                                                  | Redoes the latest mounted Univer history entry; fails when none is available.                                                       |
| `xlsx.range.flash_fill`                    | `{ "sheet", "range" }`                                                                | Infers from target examples, preserves non-empty cells, and fills empty targets in one native write.                                |
| `xlsx.range.text_to_columns`               | `{ "sheet", "range", "delimiter" }`                                                   | Splits one fully loaded text column using `tab`, `comma`, `semicolon`, or `space`; high-risk and undoable.                          |
| `xlsx.row.set_height`                      | `{ "sheet", "row", "count", "heightPoints" }`                                         | Sets a bounded 1-based row span to a final `0–409.5` point height through native undo.                                              |
| `xlsx.column.set_width`                    | `{ "sheet", "column", "count", "widthCharacters" }`                                   | Sets a bounded A1 column span and returns the actual quantized OOXML character width.                                               |
| `xlsx.column.copy_widths`                  | `{ "sourceSheet", "sourceColumn", "destinationSheet", "destinationColumn", "count" }` | Copies a bounded width vector through one native Undo unit without changing cell contents.                                          |
| `xlsx.sheet.set_freeze`                    | `{ "sheet", "frozenRows", "frozenColumns" }`                                          | Sets explicit frozen row/column counts; `0,0` unfreezes and native undo owns the change.                                            |
| `xlsx.sheet.set_gridlines`                 | `{ "sheet", "visible" }`                                                              | Sets final worksheet gridline visibility through native undo and the saved sheet-view journal.                                      |
| `xlsx.sheet.set_formula_view`              | `{ "sheet", "enabled" }`                                                              | Sets final per-sheet formula view through renderer-owned undo and the saved sheet-view journal.                                     |
| `xlsx.sheet.set_fit_to_pages`              | `{ "sheet", "widthPages", "heightPages" }`                                            | Sets explicit `0–1000` fit axes as one undo unit; `0,0` disables fit-to-page.                                                       |
| `xlsx.sheet.set_page_orientation`          | `{ "sheet", "orientation" }`                                                          | Sets `portrait` or `landscape` through exact renderer-owned undo and the saved page-setup journal.                                  |
| `xlsx.sheet.set_page_margins`              | `{ "sheet", "margins" }`                                                              | Sets `normal`, `wide`, or `narrow` through exact renderer-owned undo and saved OOXML margins.                                       |
| `xlsx.sheet.set_paper_size`                | `{ "sheet", "paperSize" }`                                                            | Sets one of seven visible paper-size codes through exact renderer-owned undo and saved OOXML page setup.                            |
| `xlsx.sheet.set_print_area`                | `{ "sheet", "range" }`                                                                | Sets a normalized explicit A1 print area or clears it with `null`; exact undo and saved defined-name state.                         |
| `xlsx.sheet.set_print_gridlines`           | `{ "sheet", "enabled" }`                                                              | Sets final printed-gridline state through exact renderer-owned undo and saved OOXML print options.                                  |
| `xlsx.sheet.set_print_headings`            | `{ "sheet", "enabled" }`                                                              | Sets final printed row/column-heading state through renderer-owned undo and saved OOXML print options.                              |
| `xlsx.sheet.set_print_scale`               | `{ "sheet", "scalePercent" }`                                                         | Sets a fixed integer print scale from 10 through 400 and disables fit-to-page in the same undo unit.                                |
| `xlsx.sheet.set_print_titles`              | `{ "sheet", "rows" }`                                                                 | Sets an ascending explicit row span of at most 21 rows or clears it with `null`; exact undo and saved `_xlnm.Print_Titles` state.   |
| `xlsx.range.set_protection`                | `{ "sheet", "range", "protection", "fields" }`                                        | Records explicit locked/hidden OOXML flags for at most 10,000 cells; file-side and non-undoable.                                    |
| `xlsx.range.set_checkbox`                  | `{ "sheet", "range", "enabled" }`                                                     | Sets or removes checkbox validation over at most 10,000 cells through native Undo and browser save/reopen.                          |
| `xlsx.range.set_list_validation`           | `{ "sheet", "range", "values", "allowBlank", "showDropdown" }`                        | Sets a bounded inline-list rule through native validation history and declarative save.                                             |
| `xlsx.range.set_list_reference_validation` | `{ "sheet", "range", "sourceSheet", "sourceRange", "allowBlank", "showDropdown" }`    | Sets a bounded single-axis range-backed list through native validation history.                                                     |
| `xlsx.range.remove_data_validation`        | `{ "sheet", "range" }`                                                                | Removes any validation rule from an explicit bounded range through the same shared route.                                           |
| `xlsx.range.set_comparison_validation`     | `{ "sheet", "range", "kind", "operator", "operand1", "operand2"?, "allowBlank" }`     | Sets one of five scalar comparison kinds and eight operators through native validation history.                                     |
| `xlsx.range.set_custom_formula_validation` | `{ "sheet", "range", "formula", "allowBlank" }`                                       | Sets a bounded equals-prefixed custom formula through native validation history.                                                    |
| `xlsx.range.set_validation_messages`       | `{ "sheet", "range", nullable prompt/error fields }`                                  | Sets or clears bounded prompt/error metadata on one existing rule through native update-options Undo.                               |
| `xlsx.hyperlink.set`                       | `{ "sheet", "address", "target" }`                                                    | Normalizes and records one URL or internal-sheet link plus the retained link appearance.                                            |
| `xlsx.hyperlink.remove`                    | `{ "sheet", "address" }`                                                              | Removes one hyperlink target and clears its link appearance.                                                                        |
| `xlsx.table.add`                           | `{ "sheet", "range", "style" }`                                                       | Creates a new table with one of the six retained Ribbon styles and reports its generated name.                                      |
| `xlsx.row.insert`                          | `{ "sheet": string, "row": integer, "count": integer }`                               | Inserts rows through the mounted Univer worksheet and undo journal.                                                                 |
| `xlsx.row.delete`                          | `{ "sheet": string, "row": integer, "count": integer }`                               | Deletes rows through the mounted Univer worksheet and undo journal.                                                                 |
| `xlsx.column.insert`                       | `{ "sheet": string, "column": string, "count": integer }`                             | Inserts columns before the named A1 column label.                                                                                   |
| `xlsx.column.delete`                       | `{ "sheet": string, "column": string, "count": integer }`                             | Deletes columns starting at the named A1 column label.                                                                              |
| `xlsx.sheet.add`                           | `{ "name": string }`                                                                  | Adds a worksheet through the mounted Univer workbook.                                                                               |
| `xlsx.sheet.rename`                        | `{ "sheet": string, "name": string }`                                                 | Renames a worksheet through the mounted state.                                                                                      |
| `xlsx.sheet.delete`                        | `{ "sheet": string }`                                                                 | Deletes a worksheet while retaining at least one sheet.                                                                             |
| `xlsx.sheet.move`                          | `{ "sheet": string, "position": integer }`                                            | Moves a worksheet to a 1-based tab position.                                                                                        |
| `xlsx.sheet.set_protection`                | `{ "sheet": string, "protected": boolean }`                                           | Records an explicit passwordless worksheet-protection state.                                                                        |
| `xlsx.sparkline.add`                       | `{ "sheet", "sourceRange", "targetRange", "type" }`                                   | Adds at most 200 row-aligned line/column/win-loss sparklines through shared Undo and native x14 save/reopen.                        |
| `xlsx.outline.set_level`                   | `{ "sheet", "axis", "start", "count", "level" }`                                      | Sets an absolute `0..7` row/column outline level through shared Undo and structural save/reopen.                                    |
| `xlsx.outline.set_detail_visibility`       | `{ "sheet", "axis", "start", "count", "hidden" }`                                     | Sets final detail visibility and the following summary item's collapsed state as one undo unit.                                     |
| `xlsx.defined_name.set`                    | `{ "name", "formula", "scopeSheet"?, "previousName"? }`                               | Upserts or atomically renames one workbook- or named-sheet-scoped Defined Name through native history.                              |
| `xlsx.defined_name.remove`                 | `{ "name", "scopeSheet"? }`                                                           | Removes one explicitly scoped Defined Name through the same native model and declarative save route.                                |
| `xlsx.document.load_staged` internal       | `{ blobId, name, size, data }`                                                        | Loads host-hydrated `.xlsx` bytes through the retained workbook-open path without a recovery checkpoint.                            |
| `xlsx.document.save`                       | `{}`                                                                                  | Saves a workbook copy through the source-preserving worksheet patcher.                                                              |

The ninety public fully qualified XLSX operations and two internal staged operations are generated
registry operations. R6-09 retires all twenty-two public-era XLSX aliases; `open_local_file`
remains only as the internal staged-load transport alias. Save success returns
`{ "saved": true, "fileName": string }`; cancellation or write failure returns `execution_failed`.
Staged load success returns `{ "opened": true, "fileName": string }`; malformed descriptors and
load failure do not advance revision. The former open-ended Agent `ribbon_command` route was
removed in R2-110. Text-style changes must use `xlsx.range.set_text_style`. Alignment,
font appearance, fill, border, cell-style, number-format, merge, clear, fill, basic/custom sort,
remove-duplicates, AutoFilter toggle/clear/advanced criteria, aggregate-formula, Flash Fill, hyperlink, format-as-table,
cell-protection, sheet-protection, fixed print-scale, print-area, and print-title families must
use their exact operations above. Context-relative sheet/row/column insert/delete strings must use
the exact structure operations above. Agent `undo` and `redo` Ribbon strings must use
`xlsx.history.undo` and `xlsx.history.redo`; visible controls continue to use that same mounted
Univer history. Agent `paste-special:value`, `paste-special:formula`, `paste-special:format`,
`paste-special:col-width`, and `paste-special:besides-border` must use
`xlsx.range.copy_values`, `xlsx.range.copy_formulas`, `xlsx.range.copy_formats`,
`xlsx.column.copy_widths`, and `xlsx.range.copy_without_borders`; the visible Ribbon commands
continue to use Univer's native clipboard. Relative decimal Ribbon gestures
are represented to Agents as an explicit final number-format assignment. Sort keys use absolute A1
column labels inside the target range; deduplication requires fully loaded sheet data.
Hyperlink and protection operations share the same file-side journals as the visible Ribbon;
`xlsx.table.add` shares its table engine. R2-78 through R2-81 connect hyperlink, table, and
protection journals to the browser package writer and reopen path. Native table metadata and
worksheet protection state are hydrated again on open; passworded unprotect remains fail-closed.
Agent `copy`, `cut`, and `paste` are likewise UI-only because their result depends on transient
clipboard and selection state; deterministic calls use explicit source/destination operations.
`format-painter` is UI-only because it depends on the next user selection. The former unqualified
Agent Ribbon route is absent.

R2-102 and R2-103 add comparison plus text/blank/duplicate Conditional Formatting create/update
and exact-ID removal. XLSX context publishes at most 100 session rule IDs with sheet/range/kind
metadata; operations remain independent of the selected cell. They use Univer's native CF history
and declarative desktop/browser package save path. Interactive `cf-open` is not accepted from the
Agent.

Context includes workbook, active worksheet, active range/cell, visible value/formula, worksheet names, and bounded dimensions.

### PPTX

| Family     | Public operations                                                                                        |
| ---------- | -------------------------------------------------------------------------------------------------------- |
| Document   | `pptx.document.create_blank`, `pptx.document.save`, `pptx.document.save_as`                              |
| History    | `pptx.history.undo`, `pptx.history.redo`                                                                 |
| Selection  | `pptx.selection.set`                                                                                     |
| Slides     | `pptx.slide.add_blank`, `pptx.slide.duplicate`, `pptx.slide.delete`, `pptx.slide.move`                   |
| Objects    | `pptx.object.delete`, `pptx.object.set_transform`, `pptx.object.move_selection`                          |
| Text       | `pptx.text.set_paragraphs`, `pptx.text.set_font`, `pptx.text.replace_all`, `pptx.text.replace_selection` |
| Paragraphs | `pptx.paragraph.set_format`                                                                              |

Internal `pptx.document.load_staged { blobId, name, size, data }` is the nineteenth descriptor and
is rejected by direct `office_execute`. Inputs use exact bounded schemas; geometry uses
document-native EMU. R6-09 retires the legacy `save`, `select_objects`,
`replace_selected_text`, and `move_selected_objects` aliases; `open_local_file` remains only as the
internal staged-load transport alias. User and
Agent routes share the mounted `BrowserPresentation`, renderer refresh, dirty state, recovery,
native undo/redo, and persistence authority. Discovery advertises the 18 public descriptors only.

Context includes active slide, selected object summaries, bounding boxes, text, notes, and slide count.

### PDF

The retained PDF community renderer accepts local files from the visible browser picker and
`office_open_local_file`. Its PDF-owned Registry contains 25 operations: 23 Agent-visible and two
internal staged-byte routes.

| Family                  | Public operations                                                                                         |
| ----------------------- | --------------------------------------------------------------------------------------------------------- |
| Document                | `pdf.document.set_metadata`, `pdf.document.save`                                                          |
| History                 | `pdf.history.undo`, `pdf.history.redo`                                                                    |
| Markup/pending          | `pdf.markup.add`, `pdf.annotation.delete_saved`, `pdf.pending.delete`                                     |
| Drawing/signature       | `pdf.drawing.add`, `pdf.drawing.update`                                                                   |
| Text                    | `pdf.text.insert`, `pdf.text.replace`, `pdf.text.update_inserted`                                         |
| Image/static form       | `pdf.image.insert`, `pdf.image.transform`, `pdf.image.replace`, `pdf.image.delete`, `pdf.static_form.set` |
| AcroForm                | `pdf.form.set_value`                                                                                      |
| Watermark/header/footer | `pdf.stamp.set`                                                                                           |
| Pages                   | `pdf.page.insert`, `pdf.page.delete`, `pdf.page.reorder`, `pdf.page.set_rotation`                         |

Internal `pdf.document.load_staged { blobId, name, size, data }` and
`pdf.page.insert_staged { blobId, name, size, data, afterPageIndex }` receive Broker-hydrated bytes
and cannot be called through `office_execute`. Public page insertion accepts a bounded absolute
local PDF path, stages it, then persists and reloads the merged active document; it is explicitly
non-undoable. Other mutations use mounted App history or their declared persistence semantics.
Staged document load resolves only after PDF.js parsing, page-size/base-rotation state, React
status, and the refreshed community command controller have committed for two animation frames.
A positive load acknowledgement therefore makes an immediately following typed mutation safe; it
cannot race the previous document controller.

R6-09 retires the legacy `delete_saved_annotation`, `undo`, and `save` aliases; `open_local_file`
remains only as the internal staged-load transport alias. Browser PDFium supplies searchable text and content-image
save paths; allowlisted script fonts are loaded lazily and cmap validation rejects unsupported
glyphs before mutation. Save success returns `{ "saved": true }`. PDF is included in the passing
R6-01 all-format readiness evidence.

## Errors

All tool failures set `isError: true` and return `{ ok: false, error, message }` as structured content.

| Code                       | Meaning                                                                    | Recovery                                                                                    |
| -------------------------- | -------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------- |
| `session_not_found`        | No live Session or exact recovery exists for the requested id.             | Use the recorded id with `resume: exact`; use `latest` only for explicit disaster recovery. |
| `editor_offline`           | No mounted renderer is connected.                                          | Show/reopen the editor and wait for `connected: true`.                                      |
| `editor_view_conflict`     | Another mounted view owns this Session's active lease.                     | Continue in the existing editor; do not mount or bind a duplicate view.                     |
| `revision_conflict`        | A new request's base revision or a chunk offset is stale.                  | Refresh context and rebuild a new intended mutation; exact known replays are exempt.        |
| `command_in_flight`        | Another mutation is pending or active.                                     | Wait for its result; replay only the exact accepted request when its outcome is uncertain.  |
| `command_not_found`        | A command, staged file, or recovery upload does not belong to the session. | Refresh/restart the affected workflow.                                                      |
| `command_timeout`          | This caller saw no acknowledgement within 15 seconds.                      | Keep the editor open; retry the exact same transaction request id and payload.              |
| `request_reused`           | A session request id was reused with a different transaction payload.      | Stop; use a new request id only for a newly grounded intended mutation.                     |
| `transaction_not_atomic`   | The requested group has no declared native atomic/undo route.              | Submit one operation; do not approximate a grouped transaction client-side.                 |
| `operation_not_found`      | The id is not an exact canonical Agent operation for the session format.   | Rediscover the canonical id; aliases and internal ids are not accepted.                     |
| `operation_schema_invalid` | Arguments do not match the operation's exact generated schema.             | Fetch detail discovery and correct the arguments without changing revision.                 |
| `operation_unavailable`    | Required live context, currently selection, is unavailable.                | Refresh context and establish the required live selection.                                  |
| `unsupported_operation`    | The format adapter does not implement the operation.                       | Read capabilities and choose a supported operation.                                         |
| `invalid_arguments`        | Tool or renderer arguments are invalid for the operation/format.           | Correct arguments without changing revision.                                                |
| `execution_failed`         | The mounted renderer rejected or failed the operation.                     | Preserve the message and current document state.                                            |
| `internal_error`           | An unexpected server error escaped the typed layer.                        | Inspect local server diagnostics.                                                           |

## Current limitations

These are current protocol or release limitations. They are not unexplained format-local retained
command gaps.

- Session metadata remains process-memory only; recovery stores opaque renderer snapshots keyed by
  format and session id.
- Recovery is local plaintext crash protection; encryption at rest and a user-facing recovery
  candidate browser are not implemented.
- DOCX and Markdown retained-command producer mappings are complete and both are included in the
  passing R6-01 release evidence.
- Standalone-browser overwrite depends on a granted `FileSystemFileHandle`; cancelling the picker is
  not a save. Inside an MCP Apps iframe, all five formats send generated bytes through the
  host-mediated `ui/download-file` boundary and report success only after the host accepts the
  download. PPTX Save reuses a standalone handle, while Save As always requests a new destination
  and adopts the selected or explicitly entered file name.
- XLSX mounts the pinned community `App` directly; all permitted renderer files and focused tests are restored, the substitute renderer is removed, and 114 format-owned operations cover the audited retained mutation surface through shared Univer/file-journal and browser save/reopen routes. Transient UI/navigation/clipboard arming and external export gestures are not document mutations. XLSX passes the shared packaged-host visual, performance/resource, MCP smoke, and repository gate.
- PPTX's 74-operation Registry covers every retained state-changing producer through its complete
  browser API, native history, recovery, and package save seam.
- PDF's 25-operation Registry covers every retained state-changing producer. Browser PDFium handles
  text/image content streams and PDF-lib-safe routes cover annotations, drawings, forms,
  stamps/signatures, metadata, and pages. PDF passes the same all-format release evidence.
- All five generated editor mounts have standalone visual smoke coverage and active four-state
  Codex-host pixel matrices. R6-01 adds pinned-source split-view provenance, deterministic
  small/medium/large opens, cold start, interaction, ACK decomposition, and peak-memory evidence.
- Background or UI-closed editing remains prohibited.
