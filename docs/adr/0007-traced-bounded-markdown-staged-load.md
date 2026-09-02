# ADR 0007: Trace and bound Markdown staged document loading

- Status: Accepted
- Date: 2026-08-31
- Upstream baseline: `genspark-ai/genoffice@dc4d7e5927864498913b7ba42d0da06cc7cf628e`
- Extends: ADR 0005 and ADR 0006

## Context

The approved R6-02 packaged-host evidence measured the canonical 769,818-byte Markdown fixture at
about 22.2 seconds p95. Its aggregate acknowledgement split attributed almost all time to renderer
execution, but could not distinguish UTF-8 decoding, Markdown parsing, TipTap state installation,
or the final React commit.

Profiling showed that TipTap Markdown extension probes for ordered lists, task lists, and tables
split the entire remaining source at every Marked cursor. The canonical fixture has 12,000
headings and 12,000 paragraphs, so these non-matching probes made an otherwise linear parse path
quadratic. Replacing the parser, introducing a second document model, or acknowledging before the
mounted editor committed would violate the retained-renderer and shared-state boundaries.

## Decision

Successful app-only acknowledgements for `markdown.document.load_staged` may include this exact,
bounded trace under `timing.trace`:

```json
{
  "operation": "markdown.document.load_staged",
  "phases": {
    "decodeMs": 0,
    "parseMs": 0,
    "tiptapStateInstallMs": 0,
    "reactCommitMs": 0
  }
}
```

The four non-negative local durations mean:

- `decodeMs`: staged `ArrayBuffer` to UTF-8 text;
- `parseMs`: envelope and legacy normalization plus Markdown-to-TipTap JSON;
- `tiptapStateInstallMs`: TipTap JSON to the mounted ProseMirror state and view;
- `reactCommitMs`: file name, frontmatter, dirty/status state, and final React layout commit.

The renderer sends a positive acknowledgement only after the React layout commit is observable.
Asset hydration that follows document installation remains part of total renderer execution but is
not mislabeled as one of these four phases. The trace is app-only, contains no document content or
path, and is not a public MCP operation or telemetry channel.

The retained TipTap/Marked extensions remain authoritative. TandemFolio-owned wrappers add constant-time
first-line or first-two-line candidate guards for ordered lists, task lists, and tables, then call
the original upstream tokenizer unchanged when the source is a real candidate. All TipTap packages
used by Markdown are normalized to the declared 3.30.5 dependency set so tokenizer behavior is not
silently mixed across versions.

Release-evidence schema v2 records the four phase summaries for small, medium, and large Markdown
fixtures. Markdown staged-open p95 is gated by fixed ceilings of 20 ms, 835.68 ms, and 5,000 ms
respectively. These are reviewed product budgets, not ceilings derived from the capture that is
being approved.

## Consequences

- A regression can be assigned to decode, parse, TipTap installation, React commit, or residual
  execution rather than hidden in one aggregate duration.
- The canonical large fixture must open and commit before acknowledgement within five seconds.
- Markdown syntax, native TipTap state/history, save/reopen behavior, one mounted iframe, and
  session revision semantics remain unchanged.
- Candidate guards must retain parse/serialize/parse structural tests for actual ordered lists,
  task lists, and tables.
- This decision does not add a worker-owned editor, a parallel parser model, public tools,
  telemetry, or enterprise source.
