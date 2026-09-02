# PPTX community renderer capability inventory

- Baseline: `genspark-ai/genoffice@dc4d7e5927864498913b7ba42d0da06cc7cf628e`
- Governing decisions: [ADR 0003](../adr/0003-complete-community-renderers-and-mcp-parity.md) and [ADR 0004](../adr/0004-format-owned-operation-registries.md)
- Status: permitted renderer restored; retained state-changing command parity complete through R2-308; R6-01 passed; `ready === true`
- Evidence date: 2026-08-30

## Pinned source accounting

The pinned `apps/slides/src/renderer` tree contains 104 files. Every path has an explicit
disposition:

| Classification                                      | Count | Current evidence                                                                                                                                                                                                                                                                                |
| --------------------------------------------------- | ----: | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Byte-identical pinned renderer files                |    61 | Hash equals the pinned blob at the same path.                                                                                                                                                                                                                                                   |
| Host/product-boundary-adapted pinned renderer files |    19 | `App.tsx`; `AnimationPane.tsx`, `CommentsPane.tsx`, `FormatPane.tsx`, `Ribbon.tsx`, `RibbonHomeTab.tsx`, `icons.tsx`, `ribbon-shared.tsx`; `export-render.tsx`; locale/string modules; `index.html`, `keyboard-actions.ts`, `main.tsx`, `show-actions.ts`, `styles.css`, and `undo-routing.ts`. |
| Prohibited pinned files intentionally absent        |    24 | Nine `renderer/ai/**` modules, 14 AI/chat/branding assets, and `i18n/strings-ai.ts`.                                                                                                                                                                                                            |

The 61 + 19 + 24 classification covers all pinned renderer paths. TandemFolio-owned support consists of
the three `renderer/host/**` files, three `renderer/operations/**` files, and `bidi-js.d.ts`.
The pure `apps/slides/src/main/edit-text.ts` mapper is restored byte-for-byte outside the renderer
tree and reused by the browser host. Electron `slides-main.ts` is behavior evidence only and is
not shipped or imported. No enterprise source was inspected or used.

## Mounted renderer and host boundary

`renderer/main.tsx` mounts the original community `App`, Ribbon, Konva canvas, dialogs, panes,
views, text editor, notes, comments, master view, ink, clipboard, keyboard, and context menus.
`?mode=audience` again mounts the original `AudienceView`; a same-origin browser channel shares
the presentation render snapshot, animations, transitions, Morph keys, absolute show state,
navigation, and ink with `PresenterView`.

`BrowserPresentation` is the one format-owned package/state adapter. User and Agent changes share
its package model, selection context, monotonically revised live session, dirty state, native
snapshot history, renderer refresh, recovery checkpoint, and save seam. `BrowserSlidesHost` is
now a compile-time-complete `SlidesApi`; the former `Partial` plus unsupported-method Proxy is
removed. AI-only URL image methods/history batching and the unused whole-section setter were
removed from the retained interface instead of being reimplemented as product capability.

Browser host effects are implemented without pretending to be document mutations:

- image and media pickers normalize files to bounded bytes before shared retained mutations;
- image/PDF export uses browser downloads, with directory-handle writes where available;
- print uses an isolated frame for full-page, handout, and notes layouts;
- Presenter/Audience uses same-origin channel transport; display swapping returns false where the
  browser cannot move windows across physical displays.

## Retained mutation parity

PPTX owns 74 serializable/executable descriptors: 73 Agent-visible operations and internal
`pptx.document.load_staged`. R2-239 through R2-253 established the first document, selection,
history, slide, object, text, and paragraph tracers. R2-254 through R2-308 close the retained
families:

- slides, layouts, size, background, hidden state, transitions, timings, header/footer, and
  explicit slide/object copy;
- explicit object lifecycle, z-order, group/ungroup, flip, fill/image-fill/stroke, batch geometry,
  connector endpoints, and parent-addressed group-child text/font/paragraph/transform edits;
- tables, pictures, charts, SmartArt, theme, animations, hyperlinks, notes, comments, sections,
  ink, embedded images/media/3D, and master/layout-part editing;
- picker-driven image insert/fill and internal element/slide clipboards converge on the same
  primitives as their bounded Agent operations.

`renderer/operations/baseline.ts` machine-checks every descriptor against a retained producer
family and has no `missing` disposition. Export/print/presenter routes are explicitly classified as
non-mutating `host-effect`. R6-09 retires the legacy `save`, `select_objects`,
`replace_selected_text`, and `move_selected_objects` aliases; `open_local_file` remains an internal
transport alias. The MCP server contains no handwritten PPTX schemas and the browser host contains
no operation-id dispatch branches.

Undo/package regressions found during the migration are now locked down: table preset style-part
injection occurs inside the history unit, so Undo removes the injected `tableStyles.xml`; Morph
keys read stable `cNvPr id`; master caches reset with history restoration; group-child operations
write their direct group slices without ungrouping.

## Verification

- Slides workspace: four test files / 168 passing tests.
- Registry/baseline: every descriptor has a handler and retained producer mapping; generated
  Manifest has 337 operations total, including 74 PPTX operations.
- Browser/engine evidence covers real-fixture open, shared UI/Agent primitives, native Undo/Redo,
  package checkpoint/save, master edits, clipboards, embedded resources, and package-part rollback.
- `tests/visual/pptx-community.spec.ts` retains 17 original-App and Codex-host scenarios, including
  open/navigation, text/geometry/formatting/deletion, slide lifecycle/history, save/reopen, and the
  four width/fullscreen states.
- Typecheck, deterministic Manifest generation/check, PPTX production build, plugin packaging,
  `git diff --check`, and the resource guard pass.
- Generated self-contained PPTX resource: 3,370,397 raw bytes / 986,896 gzip bytes against the
  4,000,000 raw-byte ceiling. The ceiling is a regression guard, never permission to remove
  retained capability.

## Shared release evidence

R6-01 closes the ADR 0003 cross-format gate with pinned-source visual provenance, canonical
runtime/resource samples, MCP smoke, license/product-boundary checks, and repository verification.
`office_get_capabilities(format: "pptx").ready` is generated as `true` with the other four formats.
