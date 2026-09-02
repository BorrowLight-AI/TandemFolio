# ADR 0002: Persist local recovery snapshots without creating a second editor authority

- Status: Accepted
- Date: 2026-08-26
- Extends: ADR 0001
- Extended by: [ADR 0003](0003-complete-community-renderers-and-mcp-parity.md)
- Superseded in part by: [ADR 0010](0010-immutable-editor-view-leases-and-exact-session-resume.md)

## Context

ADR 0001 made the mounted format renderer authoritative and initially deferred durable document ownership. That boundary prevents hidden headless edits, but a Codex task change can restart the MCP server or remount the iframe. An unsaved live document then has no durable identity and may reopen as a blank document.

The community DOCX renderer already has a surgical serialization pipeline and crash-recovery behavior. XLSX, PPTX, and PDF will need the same host-level recovery property without moving their format semantics into the MCP broker.

## Decision

TandemFolio may persist local, opaque recovery snapshots produced by a mounted renderer.

- A recovery snapshot contains serialized format bytes plus minimal identity metadata.
- Snapshot identity is the pair `(format, sessionId)`. Reconnecting an existing session may only
  auto-stage that session's snapshot; it must not consume another task's unsaved document.
- The renderer creates the bytes through its existing format-owned save pipeline.
- The MCP transport transfers snapshots in bounded chunks and stores them only on the local machine.
- The broker never parses, edits, merges, or claims authority over snapshot contents.
- Restoring a snapshot always mounts the matching format renderer and loads the bytes there.
- Session metadata and model context remain small; document bytes are never returned through `office_get_context`.
- A clean explicit save may retire the corresponding recovery snapshot.
- `resume: latest` is an explicit cross-session recovery request and selects the newest unexpired
  snapshot for the requested format. Ordinary reconnect uses exact-session lookup instead.
- Recovery storage contains no account, authentication, telemetry, cloud-sync, or model-provider behavior.

ADR 0010 supersedes the former same-iframe session-rebinding preference. A mounted iframe now keeps
one immutable Session/view binding; durable exact-session recovery is the fallback for MCP server
restart or iframe remount.

## Consequences

Unsaved work can survive task boundaries without introducing headless editing. Checkpoint serialization and local I/O add latency and storage cost, so implementations must debounce user-edit checkpoints, bound chunk sizes, clean obsolete snapshots, and measure the cost per format.

Recovery is local crash protection, not collaborative storage, version history, unattended editing,
or multi-user session persistence. Storage may retain one current snapshot per format/session pair;
expiry and bounded cleanup prevent abandoned task snapshots from growing without limit. Any feature
that mutates a snapshot without mounting its renderer still requires a new superseding ADR.
