# ADR 0010: Bind each editor view immutably and resume exact sessions

- Status: Accepted
- Date: 2026-09-01
- Upstream baseline: `genspark-ai/genoffice@dc4d7e5927864498913b7ba42d0da06cc7cf628e`
- Extends: ADR 0001, ADR 0002, and ADR 0006
- Supersedes: ADR 0002's same-iframe session-rebinding preference

## Context

The live editor is both a visible MCP App view and the authority for one document's state. The former
host bridge nevertheless allowed an already-mounted iframe to accept a later show notification for a
different session. At the same time, session creation recovered the newest snapshot by default and
show tools had no view identity. A follow-up turn could therefore mount a second editor, bind it to a
different session, or restore an unrelated task's newest document while the intended editor remained
visible elsewhere in the task.

This affected DOCX, Markdown, XLSX, PPTX, and PDF because all five formats use the shared live-session
bridge. It also made duplicate hidden editors continue polling and serializing recovery state, which
increased resource use and allowed a stale view to interfere with the active one.

## Decision

One live document is addressed by its stable `sessionId`, and one mounted editor instance is addressed
by an ephemeral `viewId`:

- `office_create_session` defaults to `resume: "none"`. A normal create is always an isolated blank
  session and never imports another task's recovery snapshot.
- `resume: "exact"` requires a known `sessionId`. It reuses that in-memory Session or stages only the
  recovery snapshot with the exact `(format, sessionId)` identity. A missing exact snapshot returns
  `session_not_found` instead of selecting another document.
- `resume: "latest"` remains an explicit cross-session disaster-recovery action. It is never the
  default creation or follow-up path.
- Every format show tool returns a fresh `viewId`. The renderer begins session polling only after the
  show result supplies both `sessionId` and `viewId`; show input alone is not a lease.
- The Broker grants at most one active view lease for each Session. Session-bound app-only polling,
  acknowledgement, file/recovery transfer, and disconnect calls validate the lease. A different
  `viewId` receives `editor_view_conflict` and cannot disconnect or acknowledge for the owner.
- A mounted iframe's `(sessionId, viewId)` binding is immutable. Later show notifications for another
  Session or view are ignored. Reusing the iframe for another document requires teardown and a new
  mount; data tools never remount or reparent it.
- Follow-up edits first resolve and inspect the exact prior `sessionId`. A connected Session is edited
  directly without another show call. A disconnected known Session is shown once. After a Broker
  restart, the same id is restored with `resume: "exact"` and then shown once.

The `sessionId` is the protocol's document handle for the current task. Recording and choosing among
multiple handles remains a host/Skill responsibility; the Broker does not infer a document from the
newest global snapshot.

## Consequences

- Multiple documents may coexist only as distinct Session/view pairs; one iframe cannot silently
  inherit or replace another document.
- A stale or duplicate editor cannot steal a live Session, acknowledge its commands, write its
  recovery snapshot, or mark it offline.
- Exact follow-up and restart recovery preserve document identity across all five formats without
  introducing UI-closed editing or a second document authority.
- Hosts must deliver the show result before the renderer connects, and app-only session calls must
  carry the granted `viewId`.
- Existing clients that omit `viewId` remain a temporary internal compatibility path for tests and
  non-packaged callers; every packaged editor uses the lease-bearing path.
