# ADR 0017: Bound the self-contained XLSX resource for Codex

- Status: Accepted
- Date: 2026-09-07
- Upstream baseline: `genspark-ai/genoffice@dc4d7e5927864498913b7ba42d0da06cc7cf628e`
- Supersedes: the 2026-09-04 identity-entry refinement of ADR 0009

## Context

Codex measures MCP App HTML as a UTF-8 `Blob` and rejects resources larger than
10,000,000 bytes before creating the renderer iframe. The complete XLSX resource had grown to
19,639,844 bytes while the repository allowed 21,000,000 bytes, so
`office_show_xlsx_editor` returned a resource that the host refused to mount. The Session
therefore remained disconnected and the user saw “HTML exceeds the maximum supported size.”

The largest avoidable costs were the identity-encoded initial module and twelve CSS
`@font-face` rules that embedded the same four Liberation Sans files under Aptos, Calibri, and
Carlito. Removing presets or other retained XLSX behavior would violate ADR 0003.

## Decision

- The packaged XLSX HTML has a fixed 10,000,000-byte raw ceiling matching the host boundary.
- The initial module uses the existing gzip/base64 module vault. Every emitted JavaScript module
  remains embedded and importable from the same HTML resource.
- Aptos, Calibri, and Carlito share one format-owned four-style FontFace fallback module. It is
  loaded from the module vault and settles before Univer's first canvas measurement.
- The initial executable JavaScript keeps ADR 0009's 11,000,000-byte inflated ceiling. The
  500 ms bootstrap and 1,400 ms total cold-start p95 gates remain unchanged.
- Package tests decode the entry, validate the complete static module graph, and enforce both
  ceilings. No sibling resource or network request becomes part of XLSX startup.

## Consequences

- The current complete resource is 8,721,880 bytes and its initial entry inflates to 8,390,150
  bytes, leaving headroom under both fixed limits.
- Startup performs one additional in-memory gzip decode before importing the initial entry. This
  cost remains visible in the existing module-graph phase and must pass the unchanged timing gates.
- Native Univer state, history, all retained presets, locale/package modules, browser save/reopen,
  and typed MCP operations remain available in the same mounted renderer.
- Source-fingerprinted release evidence is stale until the canonical suite is recaptured and
  approved; readiness continues to fail closed.
