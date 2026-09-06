# PPTX community renderer capability inventory

- Extraction baseline: `genspark-ai/genoffice@dc4d7e5927864498913b7ba42d0da06cc7cf628e`
- Selectively reviewed PPTX source: `genspark-ai/genoffice@f2c3d0879df29622d5a447935d2b4aeac033544d`
- Governing decisions: [ADR 0003](../adr/0003-complete-community-renderers-and-mcp-parity.md), [ADR 0004](../adr/0004-format-owned-operation-registries.md), and [ADR 0005](../adr/0005-reproducible-release-evidence-gate.md)
- Evidence date: 2026-09-06
- Capability status: applicable non-AI PPTX-native work in the reviewed range is integrated; functional checks pass; release readiness remains fail-closed until source-current evidence is recaptured.

## Source disposition

The reviewed candidate expands `apps/slides/src/renderer` to 205 paths. Path count is provenance
evidence rather than a completion metric: the candidate also contains AI panels, Agent layouts,
provider lifecycle code, desktop font download/catalog behavior, and generated-page assets that are
outside the product boundary. This migration selectively admits the native presentation engine,
renderer, editor controls, and localized UI required by the mounted browser editor. It does not move
the extraction baseline or create a merge parent.

The retained and adapted source-current areas are:

- `packages/pptx-engine/src`: OOXML parse/generate, chart and ChartEx, diagram hierarchy,
  SmartArt fallback, embedded-font discovery, theme/default-text inheritance, group/picture/table
  editing, relationship cleanup, WordArt/effect properties, and stable object identity;
- `packages/pptx-render/src`: chart geometry, pattern and group fills, DrawingML effects, scene 3D,
  preset/custom geometry, text layout, vertical/warped text, RTL, and image effects;
- `apps/slides/src/renderer`: canvas/text/table interaction, format/background panes, shape gallery,
  color history, chart gallery, print HTML, embedded-font registration, keyboard and wheel navigation,
  the shared Save/Save As serialization queue, and the split 19-locale application/Ribbon/pane catalogs;
- TandemFolio-owned `renderer/host` and `renderer/operations`: browser-safe package ownership,
  typed MCP validation and dispatch, monotonic revision, native history, recovery, and save/reopen.

Explicitly excluded source includes `renderer/ai/**`, AI layout audit and generation tools, Agent-run
lifecycle code, account/provider/telemetry code, Electron main/preload/IPC, desktop font download
services, generated-page temporary files, and `ee/`. The desktop `PrintDialog` is not imported;
the mounted browser renderer already exposes equivalent native page, handout, and notes layouts
through its format-owned print path.

## Integrated native capabilities

| Capability group                  | Source-current behavior retained in the mounted editor                                                                                                                                                                                                                   |
| --------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| OOXML and round trip              | Relationship-safe deletion, unknown-part preservation, theme override/default text style inheritance, explicit-off run properties, paragraph tabs, symbol fonts, background inheritance, hidden shapes, placeholder picture geometry, group fills, and resource cleanup. |
| Charts                            | ChartEx parsing and fallback, sparse caches, category/axis/legend fidelity, line and marker styles, data-label placement, richer chart rendering, and native `bar3D`/`pie3D` insertion and update.                                                                       |
| SmartArt and diagrams             | Diagram hierarchy parsing, stable identity, fallback rendering, group-child editing, and preservation of unsupported diagram payloads on save.                                                                                                                           |
| Shapes, WordArt, and 3D           | Preset/custom geometry, adjustment handles, shape conversion, pattern/gradient/picture fills, complete line caps/joins/compound/gradient settings, shadow/glow/reflection/soft-edge effects, scene 3D, bevels, and WordArt text effects.                                 |
| Text and RTL                      | Horizontal, stacked, vertical, 270-degree, and WordArt vertical modes; wrap/autofit/insets; highlights; warped text; East Asian and Korean font classification; RTL paragraphs and RTL table direction; explicit formatting preservation during edits.                   |
| Tables and interaction            | Rotated/flipped table cell hit testing, RTL resize direction, native table-style RTL flag, robust selection/drag transforms, centered rotation pivots, and one-history-unit adjustment drags.                                                                            |
| Pictures and media                | Source-rectangle inset crop fidelity, blip effects, image-fill tile/stretch details, placeholder geometry, embedded media/3D preservation, and relationship cleanup after replacement or deletion.                                                                       |
| Fonts                             | Shared Carlito metric fallbacks load lazily through the host font asset bridge, while embedded OOXML font faces are discovered and registered through browser `FontFace` before first render; missing or malformed faces fail soft.                                      |
| Background, print, and navigation | Solid, gradient, image, reset, and hide-master-graphics background actions; printable slide/handout/notes HTML; current-slide presentation shortcut; bounded wheel page flipping; and live deck refresh in an open audience window.                                      |
| Localized editor UI               | Source-current format panes, background pane, shape and chart galleries, color controls, and application/Ribbon/pane strings across the retained 19 locales.                                                                                                             |

## Typed mutation parity

PPTX now owns 81 serializable/executable descriptors: 80 Agent-visible operations and internal
`pptx.document.load_staged`. Seven new routes expose source-current mutations through the same
`OpenedPptx`, renderer refresh, history, recovery, and save seam:

- `pptx.slide.set_background_gradient`
- `pptx.slide.set_background_image`
- `pptx.slide.reset_background`
- `pptx.slide.set_background_graphics_hidden`
- `pptx.object.set_effects`
- `pptx.object.set_geometry`
- `pptx.text.set_body_properties`

Existing routes were widened without aliases: object fill accepts alpha, complete gradients and path
focus; stroke accepts alpha, gradient, caps, joins and compound lines; paragraph and table style
accept RTL; chart add/update accept `bar3D` and `pie3D`. Preview adjustment drags coalesce into one
native undo unit. UI gestures and Agent commands therefore converge on the same document state and
saved OOXML bytes.

## Verification

- `@genoffice/pptx-engine`: 81 test files / 836 passing tests; typecheck passes.
- `@genoffice/pptx-render`: 10 test files / 245 passing tests; typecheck passes.
- `@genoffice/slides`: 21 test files / 266 passing tests; typecheck passes.
- Integration coverage opens real fixtures through `BrowserPresentation`, dispatches the new typed
  operations, verifies native Undo/Redo and monotonic revisions, saves, reopens, and checks the
  resulting geometry, effects, text-body, and background state.
- The generated Manifest contains 364 operations: 103 DOCX, 25 Markdown, 123 XLSX, 81 PPTX, and 32
  PDF. Registry descriptors, handlers, and retained-producer mappings have no missing entry.
- AI/Electron dependency scans remain empty for the admitted PPTX product graph.

The previous release projection predates these source changes. `ready` must remain false until the
five-format source-current performance, visual, packaged-host, license, and smoke evidence is
recaptured and approved; functional and round-trip tests do not replace that gate.
