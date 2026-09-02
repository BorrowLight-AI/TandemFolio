# ADR 0005: Gate release readiness with reproducible packaged-host evidence

- Status: Accepted
- Date: 2026-08-30
- Upstream baseline: `genspark-ai/genoffice@dc4d7e5927864498913b7ba42d0da06cc7cf628e`
- Extends: ADR 0003 and ADR 0004

## Context

All five format-owned registries now close their retained state-changing producer baselines, but
registry parity alone does not prove that the packaged renderers remain visually faithful or meet
usable runtime bounds. The previous `ready: false` constants correctly avoided an unsupported
release claim, but they did not define a reproducible path to `ready: true`.

Visual snapshots, asset ceilings, browser scenarios, and format tests already existed separately.
Cold first-interactive time, small/medium/large open time, interaction latency, acknowledgement
decomposition, and peak memory had no shared evidence contract. A developer could not establish
which exact source, upstream revision, fixtures, browser, host profile, or sample counts supported a
readiness claim.

## Decision

Release readiness is a build-time projection of one approved, machine-readable evidence bundle.
The MCP server never runs benchmarks at request time.

`npm run release:baseline:capture` builds the packaged plugin and captures a candidate bundle. A
candidate records:

- the pinned repository and full upstream commit;
- a SHA-256 fingerprint of release-relevant renderer, host, package, dependency-lock, measurement,
  packaged-resource, and visual-provenance inputs;
- the Chromium version, platform, architecture, CPU, logical CPU count, memory, viewport, and
  sample counts;
- deterministic small, medium, and large fixture hashes for every format;
- reviewed pinned-source visual provenance and current packaged-host pixel differences;
- cold first-interactive, document-open, interaction, acknowledgement, JS-heap, and Chromium
  renderer-RSS measurements plus approved ceilings.

An approved capture requires every format and at least seven cold starts, seven opens per fixture
size, and twenty-one interactions. Latency summaries use deterministic nearest-rank p95. The first
approved ceilings are p95 plus twenty percent; changing a ceiling is an explicit evidence review,
not an automatic response to a regression.

Acknowledgement timing separates queue-to-poll wait, staged-byte hydration, renderer execution,
and remaining acknowledgement transport. The renderer reports only local duration values through
the app-only acknowledgement route; no telemetry leaves the local MCP host. Peak JS heap is the
stable primary browser-memory measure. Peak RSS is the sum of Chromium renderer descendants and is
retained as a host diagnostic.

The visual manifest grounds every baseline in permitted community renderer paths at the pinned
commit. The current gate uses no masks. Prohibited AI, account, branding, telemetry, Electron, IPC,
and enterprise surfaces are removed before capture rather than hidden during comparison.

`npm run release:gate` fails closed when evidence is missing, malformed, unapproved, undersampled,
captured against another upstream commit, stale against the current source fingerprint, missing a
required format or measurement, over a visual/performance/memory ceiling, or references a missing
or hash-mismatched visual artifact. Every gate run generates
`apps/mcp-server/src/generated/release-readiness.json`: passing evidence writes `ready: true`, and
all rejected evidence writes an all-format `ready: false` projection.

Readiness is global for this release: all five format gates must pass before any format reports
`ready: true`. `npm run release:package` runs the gate before rebuilding and packaging the MCP
server. Ordinary development builds remain possible before evidence exists so that the evidence
itself can be captured.

## Consequences

- `office_get_capabilities.ready` is evidence-backed instead of hand-maintained.
- A relevant source or packaged-resource change invalidates the evidence until the canonical suite
  is captured and reviewed again.
- Runtime capability lookup remains cheap and deterministic.
- Slow paths remain visible. In particular, a large Markdown fixture is not hidden inside a single
  aggregate “editor latency” number.
- The gate measures the current local canonical host profile. Additional CI or hardware profiles
  may be added later without weakening this accepted profile.
- This decision does not change renderer authority, enable UI-closed editing, introduce a document
  Runtime, or admit enterprise source.
