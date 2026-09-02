# Markdown renderer and capability inventory

- Pinned source: `genspark-ai/genoffice@dc4d7e5927864498913b7ba42d0da06cc7cf628e`
- Current state: community renderer restored and packaged; retained-command MCP parity plus R6-03 traced/bounded staged loading complete
- Capability flag: `ready: true`

This inventory distinguishes source restoration from release readiness. The original non-AI Markdown editing surface is back in the product graph, and its twenty public plus two internal operations are mapped by the format-owned retained-command audit. The approved shared packaged-host, resource, smoke, repository, and R6-03 staged-load gates generate `ready: true`; release-relevant source drift still fails closed until evidence is recaptured.

## Source classification

The pinned `apps/markdown/src/renderer` tree contains 30 files. TandemFolio retains or host-adapts the 21 non-AI renderer files and adds a browser file adapter plus the format-owned operation catalog and registry.

| Classification               | Files or areas                                                                              | Disposition                                                                                                            |
| ---------------------------- | ------------------------------------------------------------------------------------------- | ---------------------------------------------------------------------------------------------------------------------- |
| Retained community UI        | `components/{FrontmatterPanel,Ribbon,SlashMenu,TableMenu,icons}.tsx`                        | Original non-AI Ribbon, menus, and frontmatter editing remain available.                                               |
| Retained community editor    | `editor/{CodeBlockView,blockDragHandle,blockKeymap,extensions,localImage,slashCommand}.ts*` | TipTap/ProseMirror document state, history, block commands, tables, code blocks, and images remain renderer-owned.     |
| TandemFolio parser guard            | `editor/linearMarkdownExtensions.ts`                                                        | Adds constant-time candidate guards, then delegates actual ordered-list/task-list/table syntax to retained tokenizers. |
| Retained format behavior     | `markdown/docText.ts`, `export/{docxExport,printHtml}.ts`                                   | Markdown envelope round-trip, DOCX export, and print output remain format-owned.                                       |
| Host-adapted community entry | `App.tsx`, `main.tsx`, `env.d.ts`, `index.html`, i18n, and styles                           | Electron/AI calls are replaced by browser and MCP session adapters without replacing the editor.                       |
| TandemFolio browser host adapter    | `host/browser-files.ts`                                                                     | Browser file open/save/download and image selection.                                                                   |
| TandemFolio operation registry      | `operations/catalog.ts`, `operations/registry.ts`                                           | Format-owned descriptors, canonical public ids, one internal staged-load alias, exact validation, and TipTap handlers. |
| Prohibited AI source         | `ai/{AiPanel,markdown-skill,search-skill,tools,transport}.ts*`, `editor/aiHighlight.ts`     | Not restored because these files implement product AI/provider behavior.                                               |
| Prohibited AI assets         | `assets/{send-enter-off,send-enter-on,send-stop}.png`                                       | Not restored because they belong only to the removed AI composer.                                                      |
| Prohibited desktop host      | `apps/markdown/src/{main,preload}`, shared Markdown IPC, and Electron build config          | Replaced by the browser/MCP boundary; none is admitted to the product graph.                                           |

## Module ownership

| Module                                         | Owns                                                                                                      |
| ---------------------------------------------- | --------------------------------------------------------------------------------------------------------- |
| `apps/markdown/src/renderer`                   | TipTap state, selection, history, Markdown serialization, UI commands, and exports.                       |
| `packages/host-bridge`                         | Mounted-session polling, revision checks, acknowledgements, recovery transport, and display mode.         |
| `apps/mcp-server`                              | Session identity, typed capability schemas, local-file staging, and the `markdown.html` MCP App resource. |
| `tools/package-plugin.mjs`                     | Self-contained `assets/editors/markdown/index.html` generation.                                           |
| `plugins/tandemfolio/skills/tandemfolio` | Agent workflow and operation guidance; it must honor the live `ready` flag.                               |

## Current MCP routes

| Retained behavior                           | Typed route                                                         | State                                                                                                                                                                           |
| ------------------------------------------- | ------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Open `.md` or `.markdown` by absolute path  | `office_open_local_file` → internal `markdown.document.load_staged` | Complete internal registry tracer; chunk hydration, exact renderer validation, and revision acknowledgement.                                                                    |
| Insert text at the active TipTap selection  | `office_execute: markdown.text.insert`                              | Complete registry tracer; generated discovery, exact validation, native undo, and revision acknowledgement.                                                                     |
| Replace the active text selection           | `office_execute: markdown.text.replace_selection`                   | Complete registry tracer; generated discovery, exact validation, native undo, and revision acknowledgement.                                                                     |
| Set an explicit text selection              | `office_execute: markdown.selection.set`                            | R2-235 complete view-state tracer; bounded positions make arbitrary insert/delete/replace targeting reproducible without adding an Undo entry or recovery snapshot.             |
| Save Markdown                               | `office_execute: markdown.document.save`                            | Complete registry tracer; exact output, explicit failure, generated discovery, and revision acknowledgement.                                                                    |
| Save Markdown to a new destination          | `office_execute: markdown.document.save_as`                         | R2-231 complete persistence tracer; the retained Shift-Save shortcut and Registry share the forced picker path, return the final file name, and report cancellation explicitly. |
| Export DOCX                                 | `office_execute: markdown.document.export_docx`                     | R2-232 complete output tracer; UI/Registry share the DOCX engine and bounded PNG/JPEG/GIF loader, returning the exported file name or explicit failure.                         |
| Open Print / PDF dialog                     | `office_execute: markdown.document.open_print_dialog`               | R2-233 complete host-view tracer; success means the host dialog opened, while popup blocking is reported explicitly.                                                            |
| Set autosave preference                     | `office_execute: markdown.document.set_auto_save`                   | R2-234 complete explicit-final-state tracer; UI/Registry share the same persisted preference action and do not checkpoint document recovery.                                    |
| Undo or redo                                | `office_execute: markdown.history.undo` / `markdown.history.redo`   | R2-221 complete tracer; Ribbon and Registry share the native TipTap history action, unavailable entries fail explicitly, and live-session queue/acknowledgement is covered.     |
| Set paragraph/heading/quote/code-block type | `office_execute: markdown.block.set_type`                           | R2-222 complete tracer; revision-scoped `textBlockIndex` addresses flattened editable blocks, Ribbon/Slash/Registry share one action, and Undo/save-reopen are covered.         |
| Set bold/italic/strike/code/link marks      | `office_execute: markdown.text.set_marks`                           | R2-223 complete aggregate tracer; explicit PM range plus full final-state mask replaces five toggles, with UI convergence, conflict validation, Undo, and reopen evidence.      |
| Set bullet/ordered/task/no-list state       | `office_execute: markdown.list.set_type`                            | R2-224 complete aggregate tracer; addressed explicit final state replaces three toggles across Ribbon/Slash/Registry with Undo/reopen and failure coverage.                     |
| Insert table or divider                     | `office_execute: markdown.table.insert` / `markdown.divider.insert` | R2-225 complete position-addressed tracers; bounded table shape, shared Ribbon/Slash actions, native Undo, table reopen, and live-session evidence.                             |
| Insert a local image                        | `markdown.image.insert` → internal `markdown.image.insert_staged`   | R2-226 complete staged tracer; Broker and Renderer validate bounded PNG/JPEG/GIF bytes, all UI producers share data-URL insertion, and Undo/reopen are covered.                 |
| Set or remove raw YAML frontmatter          | `office_execute: markdown.frontmatter.set`                          | R2-227 complete envelope tracer; UI/Registry share complete YAML final state with dirty/recovery/save-reopen and explicit non-TipTap undo semantics.                            |
| Update a table relative to a cell           | `office_execute: markdown.table.update`                             | R2-228 aggregate tracer; eight bounded actions include explicit header final state and share one TableMenu/Registry action with Undo/reopen evidence.                           |
| Duplicate/delete/add/move top-level blocks  | `office_execute: markdown.block.update`                             | R2-229 aggregate tracer; menu, keyboard, plus, drag-drop, and Registry share one explicit block transaction with invariant/Undo evidence.                                       |
| Set code-block language                     | `office_execute: markdown.code_block.set_language`                  | R2-230 complete tracer; 30 finite final states, NodeView/Registry convergence, plaintext null mapping, native Undo, and reopen evidence.                                        |
| Read active selection and block type        | `office_get_context`                                                | Implemented as bounded context.                                                                                                                                                 |
| Inline/fullscreen display                   | show tool plus display controller                                   | Implemented without remounting for data operations.                                                                                                                             |
| Local recovery snapshot                     | session recovery transport                                          | Implemented with renderer-produced Markdown bytes.                                                                                                                              |

R6-09 retires the legacy `insert_text`, `replace_selection`, and `save` operation names. All public
Markdown discovery and renderer dispatch use canonical ids. Staged local-file loading is generated
with `visibility: internal`, remains unavailable through `office_execute`, and uses
`open_local_file` only as an internal transport alias. Markdown's composition root no longer
dispatches operation ids.

## Retained-command audit closure

`operations/baseline.ts` is the machine-checked mapping for every retained Markdown command
producer. It classifies document ingress as `typed-ingress`, native typing/deletion/text
paste-drop as `native-input` reproduced by explicit selection plus insert/replace, and all finite
format, structure, image, history, persistence, output, and preference commands as
`typed-operation`. Its test requires every one of the 22 Registry descriptors to be mapped and
rejects any `missing` disposition.

## Browser-host fidelity adaptations

- R2-236 replaces the removed `md-asset` protocol for MCP local open with a session-owned asset root, 256 KiB app-only chunk reads, 20 MiB image bounds, MIME/magic validation, and display-only data-URL hydration. Serialization preserves the authored relative or absolute path.
- R2-237 makes standalone browser Open directory-aware. It retains the selected Markdown file handle, resolves selected-tree relative images for display and DOCX export, and preserves authored paths on save. Browser security intentionally requires explicit directory authorization; MCP local open handles absolute filesystem paths through the session-bound bridge.
- R2-232 wires the same bounded image loader into UI and Agent DOCX export. Unavailable or invalid images degrade to alt text instead of disappearing silently.
- Print/PDF remains a host dialog rather than a headless PDF generator. R2-233 reports popup blocking as a deterministic operation failure.

These are host adaptations rather than unexplained capability gaps. R6-01 runs the shared
cross-format gates together and includes Markdown in the generated `ready: true` projection.

R6-03 makes staged-open performance part of that evidence. Successful internal load ACKs report
decode, parse, TipTap state installation, and final React layout commit; ACK follows the committed
filename/frontmatter/dirty/status state. The canonical 12,000-section fixture is guarded at five
seconds, with small and medium ceilings fixed at 20 ms and 835.68 ms. Ordered-list, task-list, and
table parse/serialize/parse tests protect the guarded retained tokenizers from semantic drift.

## Verification evidence

- Markdown unit suites cover all twenty public and two internal Registry contracts, the retained-command audit, explicit selection, persistence/output/preferences, local-image hydration, code-language NodeView convergence, invariants/Undo, and envelope/reopen behavior.
- MCP integration tests prove generated canonical discovery for all twenty public operations, internal visibility/staging, queue/acknowledgement, pre-enqueue validation, session-bound local-asset reads, and exact persistence/output results.
- Shared host-bridge tests prove canonical staged loads and session-bound local assets are hydrated before renderer execution while legacy transport names remain supported for unmigrated formats.
- The root typecheck, five-editor build, MCP package test, real stdio smoke test, and asset budget gate include Markdown.
- The packaged resource is `plugins/tandemfolio/assets/editors/markdown/index.html` and is exposed as `ui://tandemfolio/markdown.html`.
- A standalone in-app-browser smoke test mounted the Ribbon and TipTap surface, entered and formatted text, and reported no console errors.
- Packaged-host R6-03 tests assert the exact four-phase trace at acknowledgement time, visible committed content/status, and the fixed five-second canonical-large gate. Release evidence stores per-size nearest-rank phase summaries.

Markdown retained-command migration and the shared cross-format release gate are complete.
