# ADR 0011: Persist renderer-produced documents through the session broker

- Status: Accepted
- Date: 2026-09-02
- Extends: [ADR 0001](0001-tandemfolio-plugin-and-live-editor.md), [ADR 0002](0002-local-recovery-snapshots.md), and [ADR 0010](0010-immutable-editor-view-leases-and-exact-session-resume.md)

## Context

An MCP App iframe cannot reliably save an Office artifact with a sandboxed anchor, and the current
Codex host does not advertise the MCP Apps `downloadFile` capability. Treating a host download as
the primary Save boundary therefore reports cancellation without writing the renderer's bytes.
Recovery checkpoints protect unsaved work, but they are not user-visible document delivery and must
not be presented as a successful Save.

The mounted format renderer must remain the only authority for document contents and undo history.
The local broker may persist opaque renderer-produced bytes, provided it neither parses nor mutates
them and the destination is constrained by the active editing Session.

## Decision

The embedded DOCX, Markdown, XLSX, PPTX, and PDF renderers persist explicit saves through an
app-only, lease-checked `begin → ordered chunks → commit` protocol.

- The renderer serializes the complete document with its existing format-owned save pipeline.
- Every upload carries the immutable `(sessionId, viewId)` lease and is limited to 256 MiB.
- Chunks are bounded, base64 encoded, ordered by exact byte offset, and owned by one Session.
- Commit flushes the temporary file and atomically renames it in the destination directory. A
  partial or aborted upload never replaces the prior document.
- Opening an explicit local file binds successful later Save operations to that exact path. The
  binding is recorded only after the mounted renderer acknowledges the open.
- A new document's first Save uses `TANDEMFOLIO_OUTPUT_DIR`, or
  `~/Documents/TandemFolio` by default, inside a format/session-isolated directory. Later Save
  operations overwrite that binding.
- Save As writes a collision-safe renamed file in the Session output directory and makes it the
  Session's subsequent Save target.
- Export Copy writes a collision-safe file without changing the active binding. PDF copy/export
  flows use this mode because the open source remains the active document.
- Successful bound commits expose the absolute `filePath` in `office_get_context`; document bytes
  remain outside model context.
- Bindings survive a broker restart and are keyed by the exact Session id. Another Session cannot
  reuse, upload to, or overwrite them implicitly.

The broker is a persistence sink, not a document runtime. It does not parse, generate, edit, merge,
or replay Office semantics. Editing still fails with `editor_offline` when the renderer is absent.
Standalone browser mode continues using a user-granted file handle or browser download fallback.

## Consequences

Explicit embedded Save no longer depends on a Codex download dialog and can provide a verifiable
absolute local path. The protocol adds local write authority to the broker, so filenames,
extensions, sizes, view leases, chunk order, and destination selection are validated before commit.
Session-scoped output directories prevent identically named artifacts from different tasks from
colliding.

Recovery and document persistence remain separate: recovery is crash protection, while a successful
document commit is delivery. Any future arbitrary-path Save As picker, cloud storage, collaborative
storage, headless serialization, or broker-side format mutation requires another decision.
