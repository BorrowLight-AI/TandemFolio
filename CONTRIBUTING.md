# Contributing to TandemFolio

Thank you for contributing. TandemFolio is a browser-hosted MCP App for collaborative, local-first document editing. Read `CONTEXT.md`, `AGENTS.md`, and the relevant ADR before changing product boundaries.

## Local setup

Requires Node.js 22.12+ and npm 10+.

```bash
npm install --ignore-scripts
npm run dev:editor
```

The root workspace list is intentional. Do not replace it with wildcards.

## Validation

Run the same gates as CI:

```bash
npm run licenses
npm run check
npm run smoke:mcp
```

For renderer changes, also run the relevant format workspace tests. File open/save changes need a round-trip or fidelity regression test.

## Source-derived changes

Read [project facts and attribution](docs/project-facts.md) before porting or changing upstream-derived code. Preserve original copyright notices, make modifications prominent, and update the ledger and provenance record in the same change.

Agent commands enter through the shared live-session host bridge and mutate the already-mounted editor. Do not introduce a second editable document state.

## Pull requests

Keep changes focused, document user-visible behavior, and list the checks you ran. Do not describe a pre-release capability as production-ready.
