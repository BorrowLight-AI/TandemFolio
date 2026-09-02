# ADR 0008: Trace and bound XLSX cold start

- Status: Accepted
- Date: 2026-08-31
- Upstream baseline: `genspark-ai/genoffice@dc4d7e5927864498913b7ba42d0da06cc7cf628e`
- Extends: ADR 0005, ADR 0006, and ADR 0007

## Context

The approved R6-03 packaged-host evidence measured XLSX cold start at 1,687.694 ms p95, the
largest of the five formats. The aggregate measurement could not distinguish bundle/bootstrap
work, Univer construction, initial worksheet installation, or the first observable canvas commit.
It also sampled renderer RSS by synchronously running `ps` every 25 ms in the Playwright process,
which delayed host-side observation and inflated the wall-clock result.

Removing Univer presets, acknowledging before the real worksheet exists, or replacing the pinned
renderer would violate the complete-renderer and shared-state decisions. The cold-start boundary
therefore needs format-owned phase evidence before any optimization or release claim.

## Decision

The first successful app-only `office_editor_poll` from XLSX may include this exact optional
`startupTrace`:

```json
{
  "operation": "xlsx.editor.cold_start",
  "phases": {
    "bootstrapMs": 0,
    "univerCreateMs": 0,
    "worksheetInstallMs": 0,
    "firstCommitMs": 0
  }
}
```

The four non-negative bounded durations are non-overlapping:

- `bootstrapMs`: iframe navigation start through entry-module evaluation and entry into the XLSX
  React effect that creates Univer;
- `univerCreateMs`: construction of Univer and registration of every retained community preset;
- `worksheetInstallMs`: creation and installation of the initial workbook and worksheet state;
- `firstCommitMs`: remaining synchronous integration and rendering through an observable active
  workbook, active worksheet, and Univer canvas.

The XLSX adapter resolves a one-shot startup Promise only after that first commit. The shared Host
Bridge waits for it before the bootstrap poll. A transport failure retries the same trace; the
first successful poll consumes it, and later polls omit it. The nested input is strict and bounded,
contains no workbook data or path, is app-only, and creates neither a public operation nor a
telemetry channel. The sum of the four phases must not exceed the host cold-start measurement.

Release-evidence schema v3 records nearest-rank summaries for all four phases from seven isolated
cold starts. XLSX cold-start p95 has a reviewed fixed ceiling of 1,400 ms. Each format receives a
fresh Chromium process, and the capture process samples RSS asynchronously and serially, so prior
format heap/process history and observer blocking cannot distort another format's cold start.
Capture itself marks an over-budget format failed before evidence can become approved. XLSX targets
ES2022 because the supported MCP Apps Chromium host already implements it;
unnecessary compatibility transforms are not added to the self-contained Univer bundle.

## Consequences

- XLSX cannot connect its live command session before the real mounted workbook and canvas commit.
- Regressions can be assigned to bootstrap, Univer creation, worksheet installation, first commit,
  or residual host overhead.
- The 1,400 ms budget cannot be met by deleting presets, commands, history, save/reopen behavior,
  languages, or the single mounted iframe.
- All five UI resources, session/revision semantics, native Univer history, and format-owned package
  persistence remain unchanged.
- The approved R6-04 capture must replace schema-v2 evidence before generated readiness can pass.
- The approved seven-sample XLSX p95 is 1,188.7 ms; bootstrap remains dominant at 613.5 ms p95.
