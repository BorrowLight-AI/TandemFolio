# TandemFolio documentation

This is the documentation home for users, contributors, and maintainers. The root [README](../README.md) is the public product introduction; this page points to the detailed source of truth.

## Start here

- [Getting started](getting-started.md) — run from a checkout, build the local plugin, and open it in a compatible MCP Apps host.
- [Windows/macOS releases](distribution.md) — download, verify, install, update, and publish a gated release bundle.
- [Project facts and attribution](project-facts.md) — current release posture, source provenance, license obligations, and modification records.
- [Development guide](development.md) — maintainer commands, packaging, testing, release evidence, and troubleshooting.

## Product behavior

- [Product context](../CONTEXT.md) — product goal, vocabulary, boundaries, and invariants.
- [Live-session protocol](protocol/live-session.md) — implemented resources, tools, sessions, revisions, acknowledgements, and local persistence.
- [Operation registry contract](protocol/operation-registry.md) — typed format-owned operations and capability discovery.
- [Performance baselines](performance-baselines.md) — measurable release constraints and evidence.
- [Latest release verification](../release/validation.md) — source-current capture, checks, and remaining release blockers.

## Architecture decisions

Accepted decisions live in [docs/adr](adr/). Read the relevant ADR before changing product boundaries or behavior. Key decisions cover the mounted renderer boundary, recovery, operation registries, release evidence, command delivery, startup budgets, editor identity, and local persistence.

## Source and migration records

- [Source provenance](migration/provenance.md) — file-level extraction and modification record.
- [Migration ledger](migration/ledger.md) — current retained, removed, and adapted areas.
- [Upstream synchronization](migration/upstream-sync.md) — safe review and selective-port workflow.
- [Format capability inventories](migration/) — detailed DOCX, Markdown, XLSX, PPTX, and PDF evidence.
- [Migration roadmap](migration/roadmap.md) — future milestones and release gates.

## Documentation precedence

1. A newer accepted or superseding ADR controls architecture.
2. `CONTEXT.md` controls the current product boundary when no ADR conflicts.
3. The live-session protocol describes implemented behavior; it must not silently promise roadmap work.
4. The migration ledger records current state; the roadmap describes intended future state.
5. `project-facts.md` is the public summary of provenance and distribution obligations. The provenance document remains the file-level authority.
