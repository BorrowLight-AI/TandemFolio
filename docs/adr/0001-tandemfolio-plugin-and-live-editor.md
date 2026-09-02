# ADR 0001: Preserve the community renderer behind a live MCP App boundary

- Status: Accepted
- Date: 2026-08-26
- Upstream baseline: `genspark-ai/genoffice@dc4d7e5927864498913b7ba42d0da06cc7cf628e`
- Extended by: [ADR 0003](0003-complete-community-renderers-and-mcp-parity.md) and [ADR 0004](0004-format-owned-operation-registries.md)

## Context

Reimplementing GenOffice from screenshots or rendered output repeatedly drifts because layout, editor state, font metrics, selection, pagination, file semantics, and platform behavior are coupled beneath the visible UI. The community source already contains those decisions. A full format-neutral runtime migration would eventually create stronger headless capabilities, but it is too costly for the first useful release.

Codex and other MCP Apps hosts can render an interactive UI in a sandboxed iframe and exchange tool calls and model context with it. Editing canvases may use fullscreen display mode. Tool execution and UI rendering should be decoupled so ordinary mutations do not recreate the UI.

## Decision

TandemFolio will be a source-preserving fork of the Apache-2.0 community edition, distributed as a Codex plugin containing:

- a Skill that teaches the agent the editing workflow;
- an MCP server that exposes session, context, mutation, save, and show-editor tools;
- packaged renderer resources for MCP App iframe/fullscreen presentation;
- a standalone Web entry for local development and recovery.

The initial runtime is live-bound. A mounted format renderer owns document state and undo/redo. Agent commands are delivered to the same renderer through a thin session broker. Every accepted mutation advances a monotonic revision. If the renderer is unavailable, mutation tools return `editor_offline`; they do not silently edit a second copy.

Electron APIs are replaced behind `@tandemfolio/host-bridge`. The bridge selects an MCP Apps adapter when embedded and a browser adapter when standalone. Renderer structure is retained unless removal of excluded product behavior requires a local change.

Only `office_show_editor` is associated with the MCP UI resource. Data and mutation tools are intentionally UI-free, preventing the host from remounting the editor after every command.

## Consequences

This route minimizes visual drift, keeps existing format behavior, and gives Codex true what-you-see-is-what-the-agent-edits semantics with a much smaller migration surface.

The first version requires an open editor for mutations. Background/headless editing, UI-closed editing, durable server-side document ownership, broad automation, and multi-user sessions require a later superseding architecture.

## Rejected alternatives

- Screenshot-driven 1:1 reconstruction: too much hidden behavior and continuing visual drift.
- Full office runtime before a usable plugin: too much boundary and semantic migration cost.
- Electron as the visible product: cannot provide the desired native MCP App iframe/fullscreen experience.
- Treating MCP iframe widget state as the document authority: unsuitable for large documents, undo history, and binary fidelity.
