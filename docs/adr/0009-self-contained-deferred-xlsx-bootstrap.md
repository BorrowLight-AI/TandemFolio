# ADR 0009: Defer the self-contained XLSX module graph and bound bootstrap

- Status: Accepted
- Date: 2026-08-31
- Upstream baseline: `genspark-ai/genoffice@dc4d7e5927864498913b7ba42d0da06cc7cf628e`
- Extends: ADR 0005, ADR 0006, and ADR 0008

## Context

The approved R6-04 evidence reduced XLSX cold-start p95 to 1,188.7 ms, but its 613.5 ms
`bootstrapMs` phase remained the largest product-owned startup cost. The XLSX Vite build forced
all dynamic imports into one executable entry module. That made optional Univer locales,
formula-language catalogs, render hyphenation data, and other normally lazy modules part of the
initial parse/evaluation graph even though the first worksheet did not consume them.

An MCP App UI resource cannot assume that sibling hashed files will be served. Restoring ordinary
external chunks would therefore weaken the self-contained resource boundary. Removing retained
presets, languages, history, Registry operations, or persistence would violate ADR 0003 and ADR 0008. A finer trace is also required so a smaller aggregate cannot hide a shift from resource
delivery into module evaluation or React mounting.

## Decision

The XLSX startup trace keeps the R6-04 aggregate phases and adds an exact strict
`bootstrapPhases` object:

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

The bounded non-negative bootstrap subphases are:

- `resourceReceiveMs`: iframe navigation/time origin through completion of the HTML resource;
- `moduleGraphReadyMs`: resource completion through evaluation of the initial XLSX entry graph;
- `reactMountMs`: entry-body evaluation through entry into the React effect that creates Univer.

Their sum equals `bootstrapMs` within measurement precision. The first successful app-only poll
still consumes the trace exactly once; a transport failure retries the same trace. No public MCP
tool, Registry operation, UI resource, document content, or path is added.

The XLSX Vite build preserves dynamic-import boundaries. Packaging gzip-compresses every emitted
JavaScript module and embeds the base64 payloads in the existing XLSX HTML resource. A small
in-resource module vault inflates the entry module, resolves static dependencies to blob URLs, and
inflates/imports optional modules only when the retained renderer requests them. The runtime makes
no network request and creates no sibling MCP resource. The initial executable JavaScript has a
fixed 11,000,000-byte ceiling; every retained embedded module remains available.

Release-evidence schema v4 records nearest-rank summaries for the three bootstrap subphases from
seven isolated cold starts. The existing 1,400 ms XLSX cold-start p95 ceiling remains fixed, and
`bootstrapMs` gains a fixed 500 ms p95 ceiling. Candidate values cannot redefine either budget.
The first poll must still observe the active workbook, active worksheet, and Univer canvas.

## Consequences

- Optional locale and hyphenation modules no longer enter the first executable module graph, but
  remain embedded and loadable through the same renderer APIs.
- The XLSX editor stays one self-contained MCP HTML resource and one mounted iframe; blob URLs are
  an in-memory implementation detail rather than a transport or authority boundary.
- Startup regressions can be assigned to resource receipt, initial module graph, React mount,
  Univer construction, worksheet installation, first commit, or residual host overhead.
- Package tests reject external JavaScript/CSS references, missing module payloads, an oversized
  entry module, and malformed entry metadata.
- Native Univer history, all presets and languages, the 114-operation Registry, revisions, and
  package save/reopen behavior remain unchanged.
- Approved schema-v3 evidence is stale until the full schema-v4 capture passes and regenerates
  readiness against the new source fingerprint.
