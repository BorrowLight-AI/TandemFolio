# ADR 0006: Deliver live-session commands through wakeable bounded polls

- Status: Accepted
- Date: 2026-08-30
- Upstream baseline: `genspark-ai/genoffice@dc4d7e5927864498913b7ba42d0da06cc7cf628e`
- Extends: ADR 0001 and ADR 0005

## Context

The mounted renderer previously called the app-only `office_editor_poll` transport every 500 ms.
That preserved live-renderer authority, but each ordinary Agent mutation inherited an avoidable
fixed wait. R6-01 acknowledgement evidence showed queue-to-poll delay as the largest common cost
across all five formats.

MCP Apps does not expose a session-addressed server-to-iframe notification primitive for this
transport. Advertising `tools/list_changed` for document commands would misuse capability-change
semantics and can cause global tool refreshes. Associating `office_execute` with a UI resource would
also violate the one-mounted-editor rule by allowing mutation calls to remount the iframe.

## Decision

`office_editor_poll` remains the single app-only command-delivery endpoint and accepts an optional
bounded `waitMs` value from `0` through `10_000` milliseconds.

Each renderer performs one immediate `waitMs: 0` bootstrap poll before requesting fullscreen. It
then keeps at most one `waitMs: 10_000` request pending. The session broker owns one waiter per
session and resolves it immediately when `enqueue` adds a command. An empty timeout is a bounded
fallback and is immediately re-armed. Only a transport failure uses the existing 500 ms retry
delay.

A newer poll supersedes a lost older waiter, and disconnect resolves the outstanding waiter before
marking the session offline. DOCX and the format-neutral Markdown/XLSX/PPTX/PDF bridge use the same
wakeable scheduler.

Because removing the cadence exposes renderer work immediately, a format must not rely on the old
delay for readiness. XLSX therefore tracks file selection, parse, initial Univer installation, and
one committed frame as an explicit Promise; commands and positive acknowledgements cross that
boundary only after native selection/history state is stable.

This change adds no public MCP tool, no UI resource binding, no capability-change notification,
and no new document authority.

## Consequences

- Agent commands wake the already-mounted renderer without waiting for a 500 ms cadence.
- Idle sessions hold one bounded request instead of continuously polling.
- Lost responses recover within ten seconds, while transient transport failures retain a 500 ms
  retry.
- The initial size, bootstrap-poll, and fullscreen ordering remains unchanged.
- Session revision, one-command-in-flight, renderer-owned state/history, and acknowledgement
  semantics remain unchanged.
