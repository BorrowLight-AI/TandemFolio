# Repository guidance

## Read before changing behavior

- Read `CONTEXT.md` before changing product boundaries.
- Read relevant accepted decisions in `docs/adr/` before implementation; a conflicting change needs a superseding ADR.
- Use `docs/README.md` to find the implemented protocol, status records, runbook, and source facts.

## Product rules

- Keep document behavior inside its format owner: `apps/docs`, `apps/sheets`, `apps/slides`, `apps/pdf`, or `apps/markdown`.
- Preserve the mounted renderer as document authority. Mount it once per editing session; data tools must not remount or reparent its iframe.
- User gestures and Agent commands must converge on the same session and monotonic document revision.
- Every retained state-changing editor command needs a typed MCP route through the renderer's own state and undo path.
- Do not introduce product AI/model flows, account or telemetry code, Electron main/preload code, IPC dependencies, or a second headless document authority.

## Source, license, and documentation

- The pinned community source is `genspark-ai/genoffice@dc4d7e5927864498913b7ba42d0da06cc7cf628e`, as recorded in `upstream.config.json`.
- Follow [project facts and attribution](docs/project-facts.md) for upstream-derived code, copyright, license, modification, and review requirements.
- Update `docs/migration/provenance.md` and `docs/migration/ledger.md` whenever retained or modified upstream-derived areas move.
- Update the relevant protocol, runbook, or status document in the same change when behavior or support status changes.

## Current delivery slice

All five formats mount their pinned community renderer sources through format-owned browser/MCP host adapters. The product remains pre-release until source-current release evidence passes; do not weaken that fail-closed gate or replace the live editor with a separate document runtime.
