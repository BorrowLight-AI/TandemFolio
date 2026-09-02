# ADR 0003: Restore complete community renderers and require MCP editing parity

- Status: Accepted
- Date: 2026-08-27
- Upstream baseline: `genspark-ai/genoffice@dc4d7e5927864498913b7ba42d0da06cc7cf628e`
- Extends: ADR 0001 and ADR 0002
- Extended by: [ADR 0004](0004-format-owned-operation-registries.md)
- Supersedes: the M3–M6 narrow vertical-slice strategy and every decision that treats a TandemFolio-owned renderer replacement as the target product

## Context

ADR 0001 chose a source-preserving product because renderer behavior cannot be reconstructed faithfully from screenshots or a small operation sample. The subsequent XLSX, PPTX, and PDF work did not follow that boundary: it replaced the pinned community renderer trees with simplified five- or six-file replacement renderers and exposed only a narrow command tracer. The pinned community renderer trees contain 111 XLSX files, 104 PPTX files, and 40 PDF files. The 30-file Markdown renderer was initially removed; 22 non-AI files have since been restored with one TandemFolio browser-host adapter, while the nine absent pinned files are its AI panel/skills/transport, AI highlight, and AI send-button assets.

Those slices are useful migration diagnostics, but they are not the product described by ADR 0001. Calling them migrated or ready hides missing Ribbon, dialog, selection, keyboard, clipboard, formatting, structure, review, view, import/export, and format-specific editing behavior. It also creates UI capabilities that an Agent cannot invoke through MCP.

## Decision

### Complete renderer boundary

The first release includes five pinned community renderers:

- DOCX from `apps/docs/src/renderer`;
- XLSX from `apps/sheets/src/renderer`;
- PPTX from `apps/slides/src/renderer`;
- PDF from `apps/pdf/src/renderer`;
- Markdown from `apps/markdown/src/renderer`.

For each format, the community renderer is the implementation source and UI baseline. Migration retains every non-prohibited renderer module, interaction route, asset, locale, format engine, and focused test needed by the original editing experience. Host adapters may replace Electron main/preload/IPC and file-system calls, but may not replace the renderer with a simplified UI or redesign format behavior.

The only product-level removals are GenOffice/Genspark AI, model-provider, login, account, telemetry, branding, updater, Electron shell/main/preload/IPC, and enterprise `ee/` behavior. A removed community capability must be proven to belong to one of those prohibited areas and recorded in the migration ledger. Bundle size, implementation effort, or lack of an existing MCP schema are not valid reasons to remove an editing capability.

The transitional XLSX, PPTX, and PDF replacement trees must be retired after the matching community renderer is connected; that retirement is now complete. The prior removal of the Markdown renderer is reversed as a product decision; its non-AI source has been selectively restored from the pinned community baseline, and its package/MCP parity work must be completed.

### MCP editing parity

Every retained editing capability available to a user in a migrated renderer must also be invocable by an Agent against the same mounted editor state through MCP. This includes format-specific creation, insertion, deletion, replacement, formatting, layout, structure, object, review, history, and save/export actions present in the pinned non-AI renderer.

MCP parity does not require one public MCP tool per Ribbon command. `office_execute` may dispatch a typed, discoverable operation registry. Each retained editing capability must nevertheless have:

1. a stable operation identifier and validated argument schema;
2. sufficient context or selection operations to address the same target as the UI;
3. execution through the renderer's existing command/state/undo route rather than DOM clicking or a second document model;
4. risk and revision semantics, positive or negative acknowledgement, and a deterministic result;
5. tests showing equivalent user and Agent post-state plus save/reopen fidelity.

There may be no undocumented “UI-only” editing capability in a format declared complete. Read-only presentation behavior does not require a mutation operation, but any state-changing command does. Prohibited AI/account/Electron behavior is excluded rather than represented as an unsupported editing gap.

### Completion rule

A format is complete only when a pinned-source inventory maps every non-prohibited community editing command to both its retained UI route and MCP operation, with no unexplained gaps. The inventory must cite source files and tests. Compilation, an empty-state smoke test, a small operation catalog, or preservation of untouched OOXML/PDF parts is insufficient.

Release documentation and the plugin Skill must not call a format fully migrated, complete, or ready until this rule passes. The implemented protocol may continue to document transitional tools, but must label them as current implementation rather than target support.

## Current implementation fact

At acceptance of this ADR:

- DOCX retains the community renderer, but its non-AI command-to-MCP inventory is incomplete.
- XLSX, PPTX, and PDF now mount their permitted pinned community renderer sources behind browser/MCP host adapters. Their remaining gaps are browser save/reopen coverage and typed MCP command parity, not replacement renderer UI.
- Markdown retains the pinned non-AI renderer source and is present in the generated plugin, but only a narrow text/save MCP catalog is registered.
- The MCP operation catalog exposes only a small subset of the retained or planned editing capabilities.

Therefore none of these gaps may be described as an intentional first-release limitation. They are migration work required by this decision.

## Consequences

The migration surface and MCP catalog become substantially larger, but the project stops maintaining parallel simplified editors and regains the source fidelity selected in ADR 0001. Format work must proceed from the pinned renderer inward: retain the original UI and command paths, remove only prohibited dependencies, attach browser/MCP host adapters, and then close MCP parity gaps with tests.

ADR 0001's live mounted renderer remains the sole document authority. ADR 0002's opaque local recovery rule remains unchanged. This decision does not authorize headless editing, a second document runtime, enterprise source, or reintroduction of product AI and account systems.
