# ADR 0012: Complete exact document resume and save-target lifecycle

- Status: Accepted
- Date: 2026-09-03
- Extends: ADR 0002, ADR 0006, ADR 0010, ADR 0011
- Supersedes in part: ADR 0010's snapshot-only exact restart recovery

## Context

A successful Save can retire recovery bytes, leaving exact restart recovery with no source even
though the Session's formal file exists. A cold iframe can replay a show result while the old view
lease remains connected. Structured poll errors were treated as successful empty responses, creating
a zero-delay retry loop. Browser imports also left the previous document's Save target attached.

## Decision

- Exact resume uses only the recorded Session: an applicable recovery snapshot first, otherwise its
  bound formal file. Missing/unreadable sources fail visibly; never select another Session's latest
  document. The Broker transfers opaque bytes to the mounted renderer and does not parse them.
- Each mounted bridge adds an ephemeral mount identity to its immutable Session/view identity. A
  cold mount restores before publishing blank context. Replayed bootstrap requests do not load twice.
  A different mount cannot take an active lease. Abandoned leases may be reclaimed only after a
  bounded liveness timeout and with no command in flight; active or uncertain mutations fail closed.
- Terminal view/session errors stop polling and display a recoverable connection error. Transient
  errors use the existing 500 ms retry. Hidden healthy views retain one bounded poll, not a hot loop.
- A browser-selected replacement document explicitly invalidates the old Save target and recovery
  before replacing renderer contents. It cannot infer the source's absolute path from a File object.
  Its next Save gets a Session-isolated collision-safe target. MCP path opens bind only after ACK.
- Saves report an absolute destination in the editor with a copy action. Export Copy reports its
  output without replacing the active Save target. Failed writes never report successful persistence.
- Saved bytes and checkpoints remain local crash protection/delivery, not another document authority.

## Consequences

All five format-owned adapters share transport and lifecycle safeguards. Browser import failures may
leave the existing renderer unbound (requiring a new Save), which is safer than overwriting the wrong
file. No public Agent tool is added, no arbitrary-path picker or legacy-data migration is introduced,
and the source-current release gate remains fail closed.
