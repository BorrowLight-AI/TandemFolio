# TandemFolio context

## Product goal

TandemFolio is a lightweight, local-first document editing plugin for MCP Apps hosts such as Codex. It preserves the pinned GenOffice community DOCX, XLSX, PPTX, PDF, and Markdown Web renderers while removing product AI, login, account, branding, Electron shell, preload, and IPC code from the shipped product.

The product has three parts:

1. A format-owned Web editor that can run inside an MCP App iframe or as a standalone browser page.
2. A thin live-session broker that owns connection state, active selection/context, and a monotonic revision number.
3. A Codex plugin containing the Skill, MCP server, and packaged UI resources.

## Current model

The first release is intentionally live-bound. The mounted editor remains authoritative for document state and undo/redo. The broker routes agent commands into that same editor session and reports a typed `editor_offline` error when no editor is connected. The renderer may checkpoint opaque local recovery bytes under ADR 0002 and explicitly commit renderer-produced document bytes through the session-bound local persistence sink in ADR 0011; the broker never parses or edits either. Background editing, editing after the UI closes, multi-tenant persistence, and unattended writes remain deferred.

ADR 0003 rejects simplified renderer replacements as the first-release target, and all parallel
replacement renderer directories have been removed. All five pinned non-AI renderer structures are
mounted through format-owned browser/MCP adapters. Their machine-checked producer baselines have no
unexplained state-changing gap: 102 DOCX, 22 Markdown, 114 XLSX, 74 PPTX, and 25 PDF operations feed
the generated product Manifest. Native renderer history or explicitly declared format-owned
journals remain authoritative, and persisted families have save/reopen evidence. The remaining work
is no longer a renderer-parity or release-evidence migration: R6-01 records approved pinned-source
visual, packaged-host performance/resource, MCP smoke, license/prohibited-dependency, and repository
gate evidence for all five formats. `ready` is generated from that approved, source-fingerprinted
evidence rather than maintained as a hand-written capability claim.

R6-02 keeps that authority model but removes the fixed 500 ms command cadence. Every mounted
renderer performs one immediate bootstrap poll and then holds one bounded app-only poll that the
broker resolves when a command is enqueued. A ten-second empty timeout re-arms the request, while
500 ms remains only as a retry after transport failure. This adds no public tool or UI resource.

R6-03 makes Markdown staged loading observable and bounded without replacing its renderer. The
app-only acknowledgement records decode, parse, TipTap state installation, and final React commit
durations, and is emitted only after committed content/status is visible. Constant-time candidate
guards remove quadratic non-matching ordered-list, task-list, and table tokenizer probes while
delegating real candidates to the retained tokenizer. Release evidence records all four phases and
gates the canonical large fixture at five seconds.

R6-04 applies the same evidence discipline to XLSX cold start. Its first app-only poll waits for
the retained Univer runtime, initial worksheet, active workbook/sheet, and observable canvas, then
publishes bootstrap, Univer-create, worksheet-install, and first-commit durations exactly once.
Release-evidence schema v3 stores seven-sample phase p95 values and enforces a fixed 1,400 ms XLSX
cold-start ceiling. The capture samples RSS asynchronously to remove benchmark observer blocking,
and the self-contained XLSX build targets the ES2022 baseline supported by the MCP Apps host;
presets, commands, history, persistence, resources, and renderer authority are unchanged.

R6-05 decomposes the remaining XLSX bootstrap hotspot into resource receipt, initial module-graph
readiness, and React mount. The complete split Vite output is gzip-compressed into an in-resource
module vault: the entry graph is inflated first and optional Univer locale/hyphenation modules are
inflated only when requested, without a sibling resource or network fetch. Release-evidence schema
v4 requires all three subphases, fixes bootstrap p95 at 500 ms, retains the 1,400 ms total XLSX
cold-start ceiling, and caps initial executable JavaScript at 11 MB. The first poll, one mounted
iframe, 114-operation Registry, languages, native history, and save/reopen authority are unchanged.

The 2026-09-04 refinement records the vault's static dependencies at build time,
keeps dependency-free modules as UTF-8 bytes through Blob construction, and embeds
the critical entry as identity text so its one importable Blob URL avoids startup
base64 and gzip work. XLSX ZIP/XML package I/O loads inside the existing asynchronous `openBuffer` only when needed;
the initial native blank worksheet does not need that adapter. All modules remain
embedded and loadable, and first-open latency is measured separately rather than
hidden behind the bootstrap result. No startup boundary or budget changes. The
source-current approved capture records XLSX bootstrap p95 at 432.5 ms.

R6-06 makes Registry discovery bounded as capability count grows. `office_get_capabilities` now
defaults to a schema-free summary of at most twenty Agent-visible operations, supports family
filtering and stable cursor pagination, and returns a full schema only for one exact canonical
operation requested through the detail view. An optional live session projects stable availability
reasons without hiding operations. Internal operations and compatibility aliases never enter public
discovery, every summary response is gated below 8 KiB, and every generated descriptor is capped at
64 KiB. The Agent-visible surface remains ten tools and five UI resources; app-only transport now
uses fifteen endpoints (including lease-checked document replacement reset and cooperative handoff).

R6-07 makes mutation submission revision-guarded and idempotently replayable. The public
`office_execute` transaction envelope carries a caller `requestId`, exact `baseRevision`, and one
canonical Agent-visible operation. A session-scoped journal joins in-flight duplicates and returns
the original completed result for exact replays without redispatch, while payload-changing request
reuse, stale revisions, invalid schemas, unavailable context, aliases/internal ids, and unsupported
multi-operation grouping fail with stable typed errors before renderer dispatch. Caller timeout does
not settle an accepted transaction: the underlying renderer acknowledgement remains authoritative,
including for staged local bytes, and a later exact replay returns the final result. R6-08 removes
the old `operation` + `arguments` input. The public tool schema requires the complete transaction
envelope, all repository callers and packaged smoke use it, and only exact canonical Agent-visible
ids can enter `office_execute`. R6-09 removes the remaining thirty-six public-era format-local
aliases from the five renderer registries. The generated Product Manifest now contains exactly five
compatibility aliases: one internal `open_local_file` staged-file transport alias per format. These
internal aliases remain outside public discovery and mutation ingress.

R6-10 makes visible editor identity explicit across DOCX, Markdown, XLSX, PPTX, and PDF. Ordinary
session creation defaults to an isolated blank document, while exact follow-up/restart recovery uses
the recorded `sessionId`; newest-snapshot recovery is opt-in only. Each show call grants a fresh
`viewId`, the Broker permits one exclusive view lease per Session, and the mounted iframe keeps its
initial `(sessionId, viewId)` binding for its lifetime. Connected follow-up edits reuse the existing
Session without another show call, preventing hidden duplicate editors and cross-task document
inheritance while preserving the mounted renderer as the only editable authority.

R6-11 replaces the unavailable MCP host-download Save boundary with the app-only local persistence
protocol from ADR 0011. All five renderers serialize their own bytes, the Broker validates the active
view lease and atomically writes those opaque bytes, and `office_get_context.filePath` reports the
bound result. An opened local file is overwritten only by its own Session; a new document receives a
Session-isolated output path; Save As rebinds; Export Copy does not. This adds no headless editing or
broker-side format semantics.

R6-12 completes exact recovery and Save-target lifecycle under ADR 0012. Each cold renderer mount
has its own identity in addition to the immutable Session/view pair. A live owner cannot be forcibly replaced;
an idle abandoned lease can be reclaimed after 30 seconds without contact, never with an in-flight
command. Exact recovery loads the Session's checkpoint or bound formal file through the native
renderer. Terminal connection errors stop polling and surface a manual retry. Browser file
replacement invalidates the old binding before native load, then checkpoints the new document;
Save displays the committed absolute path with a copy action. XLSX recovery stays in memory and
its canonical Save uses the existing bound target. No legacy data migration or second authority is
introduced; source-current release evidence remains required.

R6-13 refines replayed-view recovery under ADR 0013: a visible replay of the same logical view can
request a cooperative handoff from its hidden owner. Preparation fences Agent commands and formal
saves, then a fresh native checkpoint precedes exclusive lease transfer and native restore/ACK.
Visible owners remain authoritative. Waiting views use bounded retries only while visible; a failed
checkpoint retains the original and requires explicit retry. Restoring views cannot edit a blank
replica, and file location remains visible even during conflict. No other Session is adopted.

R6-14 refines ADR 0013 under ADR 0014: visibility is a work-scheduling hint, not evidence of which
host replica the user intends to edit. **在此继续编辑** selects the exact same-view candidate and
requests cooperative handoff even when the old mount still reports active. The same native
checkpoint/restore fence applies; this is not force takeover. Automatic waiting stops after 30
seconds, and yielded mounts require explicit continuation rather than reclaiming on visibility.
Uncertain transport outcomes keep the old mount locked until lease-checked reconciliation confirms
ownership. These rules apply to all five formats without changing their renderer authorities.

R6-15 adds explicit document-replacement intent under ADR 0015. A follow-up edit preserves the
current document and Save target, not merely its Session/view. PPTX `document.create_blank` rejects
resetting a named, edited, or history-bearing presentation unless `confirmReplace: true` records
explicit user replacement intent. Untouched initial blanks remain usable. Confirmed replacement
still detaches the old Save target; ordinary slide edits and saves keep it. The shared Skill applies
the continuation rule to every format without adding reset operations or changing native authority.

Every state-changing capability retained from a community renderer must be reachable through a typed MCP operation against the same mounted state and undo stack. A format is not complete until its pinned command inventory has matching retained UI and MCP routes with save/reopen tests. Capability discovery is a bounded projection of the generated Manifest; it does not redefine the product boundary.

## Capability model

TandemFolio separates the retained renderer command inventory from the operations that are currently safe and executable through MCP.

- The **baseline capability inventory** records every retained, prohibited, or host-adapted community renderer command and the evidence required for parity. It answers what the pinned product contains; it is not an execution registry.
- Each format owns an **executable operation registry** containing only implemented Agent operations for that renderer. The registry is the source for operation identifiers, exact argument schemas, context requirements, risk, revision behavior, acknowledgement, and the renderer adapter that reaches the native state/undo route.
- The MCP server exposes a generated **capability projection** of those registries. The projection supports discovery and validation but does not own format semantics or document state.
- An **operation transaction** is one acknowledged Agent mutation against an exact base revision and a session-scoped caller request id. Exact replays converge on the same result. It may contain multiple operations only when the mounted renderer can apply them atomically as one native undo unit; no current format declares such a grouping, so the implemented envelope accepts one operation.
- A **native command seam** is the format-owned renderer route shared by user gestures and Agent operations. DOM clicking, an open-ended command string, or a second editable document model is not a native command seam.

The public MCP tool set remains small and stable as renderer capability grows. Visible Ribbon controls, shortcuts, dialogs, and context actions may map many-to-one onto typed semantic operations; MCP parity never requires one public tool per visible command.

## Invariants

- The human and agent edit the same live document state.
- An active editor is mounted once; tool calls update it without remounting or iframe reparenting.
- A mounted editor has one immutable `(sessionId, viewId)` identity, and a Session has at most one
  active view lease across all five formats.
- Ordinary creation never consumes another Session's recovery; follow-up and restart recovery use an
  exact recorded `sessionId`, while newest-snapshot recovery is explicit disaster recovery only.
- The complete pinned non-AI DOCX, XLSX, PPTX, PDF, and Markdown renderer structures are the product UI baseline.
- Every retained state-changing renderer capability has a typed MCP invocation path before the format is declared complete.
- Every advertised Agent operation originates in its format-owned executable operation registry; MCP metadata, validation, Skill references, and parity reports are generated or checked from that source.
- Baseline capability inventory and executable operation status remain distinct so missing work cannot be advertised as implemented behavior.
- Transitional central catalogs, open-ended command strings, and renderer-level operation condition chains are migration scaffolds, not extension points for new capability work.
- Product builds exclude AI/model/login/account/Electron main/preload/IPC code.
- Enterprise `ee/` source is never imported.
- Apache-2.0 attribution and a file-level migration ledger are retained.
- The renderer remains usable as a standalone Web application for development and recovery.
- Local recovery snapshots are renderer-produced crash protection, never a second editable document authority.
- Explicit local persistence writes only opaque renderer-produced bytes to the exact Session-bound target.
