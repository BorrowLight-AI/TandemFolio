# XLSX community renderer migration inventory

- Baseline: `genspark-ai/genoffice@dc4d7e5927864498913b7ba42d0da06cc7cf628e`
- Renderer-source status: complete for the permitted pinned community source set
- Capability status: 114 XLSX registry operations (112 Agent-visible, two internal) through R2-131; retained state-changing command parity is audited and the open-ended Agent Ribbon route is removed
- Product readiness: retained-command parity plus R6-05 traced/bounded bootstrap optimization complete;
  `office_get_capabilities({ format: "xlsx" }).ready === true` for source-current approved evidence

This file separates renderer-source restoration from product completion. Both the format-local
work and the shared ADR 0003/ADR 0005 release evidence now pass.

## Pinned renderer source accounting

The pinned `apps/sheets/src/renderer` tree contains 111 files. The current tree contains all 76
permitted pinned paths plus eleven TandemFolio host/operation files.

| Disposition                                               | Count | Evidence                                                                                          |
| --------------------------------------------------------- | ----: | ------------------------------------------------------------------------------------------------- |
| Byte-identical pinned files                               |    63 | `git show <baseline>:<path> \| cmp - <path>` returns equal.                                       |
| Browser-host/product-boundary adaptations of pinned files |    13 | Listed below; the original renderer structure and command state remain the implementation source. |
| Prohibited `renderer/ai/**` files                         |    20 | AI planners, prompts, tools, and product workflow.                                                |
| Prohibited AI composer assets                             |    14 | Attach/file/send/stop/app assets used only by the removed AI panel.                               |
| Prohibited AI locale catalog                              |     1 | `renderer/i18n/strings-ai.ts`.                                                                    |
| TandemFolio-only renderer host/operation files                   |    11 | Browser/operation adapters and focused shared action modules; none replaces the mounted renderer. |

The 13 adapted pinned files are:

- `App.tsx`, `ExcelShell.tsx`, `HeaderFooterDialog.tsx`;
- `edit-journal.ts`, `save-actions.ts`, `visual-actions.ts`;
- `env.d.ts`, `index.html`, `main.tsx`, `styles.css`;
- `i18n/locale.tsx`, `i18n/strings.ts`, `ribbon-icons.tsx`.

They remove prohibited AI/Electron coupling, install the browser/MCP host boundary, and connect
the original Univer state to browser open/save. No `ee/` path is part of this accounting and no
enterprise source was inspected.

Reproduce the retained/absent path classification:

```bash
comm -23 \
  <(git ls-tree -r --name-only dc4d7e5927864498913b7ba42d0da06cc7cf628e apps/sheets/src/renderer | sed 's#apps/sheets/src/renderer/##' | sort) \
  <(find apps/sheets/src/renderer -type f | sed 's#apps/sheets/src/renderer/##' | sort)
```

The expected output is exactly 35 prohibited paths: 20 under `ai/`, 14 under `assets/`, and
`i18n/strings-ai.ts`.

## Connected renderer and host boundary

`apps/sheets/src/renderer/main.tsx` mounts the restored pinned `App.tsx` directly. It retains the
original `ExcelShell`, Univer workbook, Ribbon tabs, formula bar, worksheet canvas, sheet tabs,
dialogs, keyboard routes, selection, clipboard/context menu, undo/redo, status bar, chart/data/
pivot/review/view surfaces, page layout, and header/footer UI.

The rejected substitute renderer has been removed; there is no parallel XLSX UI entry. Browser-only OOXML package I/O now lives at
`apps/sheets/src/host/browser-workbook.ts`; it is a host adapter used by the community App, not a
second renderer or document authority. `browser-desktop-api.ts` replaces Electron IPC with browser
file picker/download calls. User and Agent mutations converge on the same mounted Univer state,
edit journal, undo stack, and revision.

The optional Univer render-metrics token is aliased to an inert local identifier. No telemetry
implementation is registered.

## Original focused-test accounting

The pinned `apps/sheets/tests` tree contains 90 files: 89 test files and `fixture-builder.ts`.
All 81 permitted original test files plus the fixture builder are restored. Eight original tests
remain absent for recorded product-boundary reasons:

| Excluded tests                                                                                                              | Reason                                                                                                                                               |
| --------------------------------------------------------------------------------------------------------------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------- |
| `lazy-plan.test.ts`, `privacy-policy.test.ts`, `workbook-skill-tools.test.ts`                                               | Import prohibited renderer AI planners/policies/tools.                                                                                               |
| `close-guard.test.ts`, `xlsx-borders.test.ts`, `xlsx-recalc.test.ts`, `xlsx-sidecar.test.ts`, `xlsx-streaming-save.test.ts` | Require the removed Electron main process or XLSX sidecar. Browser-host behavior is covered by public browser tests instead of a fake desktop layer. |

The current suite contains 86 test files: 81 permitted original tests and five TandemFolio browser/product-
boundary/registry tests. `npm test -w @genoffice/sheets` executes 85 passing files and one environment-
conditional LibreOffice pivot suite; 1,345 assertions pass and one is skipped when `soffice` is not
available.

## Executable browser evidence

`tests/visual/xlsx-community.spec.ts` has 69 real-host scenarios. It proves:

- the packaged entry mounts the real `univer-sheet-main-canvas`, not a DOM imitation;
- browser-picker and Agent-staged local workbooks open in the same community App;
- user and Agent value, range, formula, and core style changes survive save/reopen;
- exact Agent undo and redo share the mounted Univer history with visible controls, advance one
  revision per successful entry, and preserve the redone value on save;
- explicit value copy materializes formula results and scalars between named ranges, shares visible
  UI Undo/Redo, and saves destination values without formulas;
- explicit formula copy translates relative, absolute, and mixed references between named ranges,
  copies scalar cells with formula-paste semantics, shares visible UI Undo/Redo, and saves
  destination formulas;
- user Ribbon formatting and typed Agent range styling converge on the same Univer state and undo
  stack; explicit alignment/font/fill/border/named-style operations survive save, while all migrated
  style Ribbon strings and UI-only `format-painter` are rejected through MCP;
- explicit number-format, merge, clear, and fill operations use the same mounted selection and undo
  stack; a combined tracer verifies live format context, visible clear undo/redo, and saved format,
  merge, filled-value, and cleared-content output;
- explicit AutoFilter enable/disable and the retained Ribbon Filter button share the same resolved
  range, native undo stack, declarative filter snapshot, browser package writer, and reopen path;
- explicit row height uses the same native row command and undo stack; its saved OOXML height
  survives reopen;
- explicit column width returns its actual quantized OOXML width; column size, freeze panes,
  gridline visibility, and formula view share native or renderer-owned undo and survive save/reopen;
- explicit print orientation, margin, paper-size, dual-axis fit-to-pages, fixed print scale,
  printed-gridline, printed-heading, print-area, and print-title settings share one page-setup
  journal/history seam; Agent and visible Ribbon output persist through save/reopen where controls
  exist, while Undo removes newly introduced overrides;
- row/column insert and delete survive save/reopen, including user undo;
- worksheet add, rename, delete, move, user add, and undo survive package save/reopen;
- defined names, Insert Function, Page Layout, and Header & Footer use restored community UI and
  persist through the browser package adapter.

`tests/visual/host-widths.spec.ts` adds five Codex-host scenes covering 420 px narrow sidebar,
720 px split view, 1332 px first fullscreen, fullscreen exit without iframe remount, and restored
host layout. Readiness
requires the real worksheet canvas, so an outer shell with a blank renderer cannot pass.

## Current typed MCP catalog

| Operation family    | Operations                                                                                                                                                                                                                                                                                                                | Shared renderer route                                                                                                                                                                                                |
| ------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Cell data           | `xlsx.cell.set_value`, `xlsx.cell.set_formula`, `xlsx.range.set_values`, `xlsx.range.replace_text`                                                                                                                                                                                                                        | Active Univer worksheet/range, explicit formula writes, shared bounded replacement, and original edit journal.                                                                                                       |
| Range relocation    | `xlsx.range.move`                                                                                                                                                                                                                                                                                                         | Native grid range move with explicit equal-shape source/destination ranges, Pivot/streaming guards, shared Undo, and package persistence.                                                                            |
| Range copy          | `xlsx.range.copy_values`, `xlsx.range.copy_formulas`, `xlsx.range.copy_formats`, `xlsx.range.copy_without_borders`                                                                                                                                                                                                        | Explicit same-shape source/destination ranges plus shared 20,000-cell and streaming guards; format-only copy preserves content, while except-border copy moves content/non-border styles and retains target borders. |
| Range styling       | `xlsx.range.set_text_style`, `xlsx.range.set_alignment`, `xlsx.range.set_font`                                                                                                                                                                                                                                            | Explicit masked native range mutations and the same undo journal.                                                                                                                                                    |
| Fill/border/style   | `xlsx.range.set_fill`, `xlsx.range.set_border`, `xlsx.range.apply_cell_style`                                                                                                                                                                                                                                             | Native range style/border routes; named presets collapse to one mutation.                                                                                                                                            |
| Number formatting   | `xlsx.range.set_number_format`                                                                                                                                                                                                                                                                                            | Explicit final pattern assignment through the native range facade.                                                                                                                                                   |
| Merge/clear/fill    | `xlsx.range.merge`, `xlsx.range.clear`, `xlsx.range.fill`                                                                                                                                                                                                                                                                 | Native range/command routes with exact modes, scopes, and directions.                                                                                                                                                |
| Sort/data cleanup   | `xlsx.range.sort`, `xlsx.range.sort_custom`, `xlsx.range.remove_duplicates`                                                                                                                                                                                                                                               | Native sort plus guarded, first-occurrence whole-row deduplication.                                                                                                                                                  |
| AutoFilter          | `xlsx.range.set_filter`, `xlsx.range.clear_filter_criteria`, `xlsx.range.set_filter_values`, `xlsx.range.set_custom_filter`                                                                                                                                                                                               | Explicit range/final-state and bounded value/custom criteria through shared native commands, Undo, declarative save, and reopen.                                                                                     |
| Checkbox validation | `xlsx.range.set_checkbox`                                                                                                                                                                                                                                                                                                 | Explicit bounded final state through native DV/Undo, declarative base-OOXML save, and checkbox hydration on reopen.                                                                                                  |
| Aggregate formulas  | `xlsx.formula.insert_aggregate`                                                                                                                                                                                                                                                                                           | Shared AutoSum-family formula placement with streamed-target protection.                                                                                                                                             |
| History             | `xlsx.history.undo`, `xlsx.history.redo`                                                                                                                                                                                                                                                                                  | Exact empty-input operations over the mounted Univer native history shared with visible UI controls.                                                                                                                 |
| Flash Fill          | `xlsx.range.flash_fill`                                                                                                                                                                                                                                                                                                   | Shared example inference, bounded probing, and one native range write.                                                                                                                                               |
| Text to Columns     | `xlsx.range.text_to_columns`                                                                                                                                                                                                                                                                                              | Shared native split command with four exact delimiter modes and undo.                                                                                                                                                |
| Row layout          | `xlsx.row.set_height`, `xlsx.row.set_visibility`, `xlsx.row.move`                                                                                                                                                                                                                                                         | Explicit bounded spans/final visibility and native row relocation through shared structural/axis journals.                                                                                                           |
| Column layout       | `xlsx.column.set_width`, `xlsx.column.set_visibility`, `xlsx.column.copy_widths`                                                                                                                                                                                                                                          | Explicit bounded column spans, final visibility, pixel/OOXML width quantization, and one native Undo unit for copied widths.                                                                                         |
| Sheet view          | `xlsx.sheet.set_freeze`, `xlsx.sheet.set_gridlines`, `xlsx.sheet.set_formula_view`                                                                                                                                                                                                                                        | Explicit final states through native/renderer undo and the page-setup journal.                                                                                                                                       |
| Page setup          | `xlsx.sheet.set_page_orientation`, `xlsx.sheet.set_page_margins`, `xlsx.sheet.set_paper_size`, `xlsx.sheet.set_fit_to_pages`, `xlsx.sheet.set_print_scale`, `xlsx.sheet.set_print_gridlines`, `xlsx.sheet.set_print_headings`, `xlsx.sheet.set_print_area`, `xlsx.sheet.set_print_titles`, `xlsx.sheet.set_header_footer` | Explicit print layout plus bounded final header/footer sections through one renderer-owned page-setup history and OOXML journal.                                                                                     |
| Hyperlinks          | `xlsx.hyperlink.set`, `xlsx.hyperlink.remove`                                                                                                                                                                                                                                                                             | Shared normalized target/link style; browser package save and streamed reopen reuse the retained worksheet/relationship gateway.                                                                                     |
| Table lifecycle     | `xlsx.table.add`, `xlsx.table.insert_rows`, `xlsx.table.delete_rows`, `xlsx.table.insert_columns`, `xlsx.table.delete_columns`, `xlsx.table.convert_to_range`                                                                                                                                                             | Retained table model/history with stable table names, bounded relative targets, declarative save, and reopen.                                                                                                        |
| Protection          | `xlsx.range.set_protection`, `xlsx.sheet.set_protection`                                                                                                                                                                                                                                                                  | Shared file-side journals; explicit fields/state, with password guard.                                                                                                                                               |
| Sparklines          | `xlsx.sparkline.add`                                                                                                                                                                                                                                                                                                      | Explicit row-aligned source/target ranges through the shared journal, renderer-owned Undo, retained x14 writer, and browser reopen hydration.                                                                        |
| Outline             | `xlsx.outline.set_level`, `xlsx.outline.set_detail_visibility`                                                                                                                                                                                                                                                            | Absolute bounded level and detail visibility over shared Ribbon actions, one renderer-owned Undo unit, structural OOXML save, and row/column metadata hydration on reopen.                                           |
| Chart lifecycle     | `xlsx.chart.add`, `xlsx.chart.update`, `xlsx.chart.remove`, `xlsx.chart.set_colors`, `xlsx.chart.set_series`                                                                                                                                                                                                              | Bounded visual journal and Undo routes; stable session/reopened chart IDs, final series/colors/advanced format state, and browser drawing/chart-part persistence.                                                    |
| Shape lifecycle     | `xlsx.shape.add`, `xlsx.shape.update`, `xlsx.shape.remove`                                                                                                                                                                                                                                                                | Bounded insertion/final-state edit/removal with explicit anchors, stable IDs, shared visual Undo, and native drawing save/reopen.                                                                                    |
| Image lifecycle     | `xlsx.image.add`, `xlsx.image.move`, `xlsx.image.remove`                                                                                                                                                                                                                                                                  | Hidden staged-byte insertion plus stable reopened IDs, exact package media reads, shared visual Undo, and drawing move/remove save.                                                                                  |
| Cell notes          | `xlsx.note.set`, `xlsx.note.remove`                                                                                                                                                                                                                                                                                       | Explicit cell targets through Univer native note history, declarative note snapshots, legacy comments/VML save, and reopen hydration.                                                                                |
| Data tools          | `xlsx.range.insert_subtotals`, `xlsx.range.consolidate`                                                                                                                                                                                                                                                                   | Explicit bounded source/target contracts over shared Subtotal/Consolidate actions, native writes, and package persistence.                                                                                           |
| Pivot lifecycle     | `xlsx.pivot.add`, `xlsx.pivot.refresh`, `xlsx.pivot.update`, `xlsx.pivot.set_member_filter`, `xlsx.pivot.add_chart`                                                                                                                                                                                                       | Stable pivot identities, bounded layout/filter/chart contracts, renderer-owned Undo where applicable, native Pivot OOXML/cache persistence, and reopen.                                                              |
| Row structure       | `xlsx.row.insert`, `xlsx.row.delete`                                                                                                                                                                                                                                                                                      | Univer worksheet structure commands and the same undo journal.                                                                                                                                                       |
| Column structure    | `xlsx.column.insert`, `xlsx.column.delete`                                                                                                                                                                                                                                                                                | Univer worksheet structure commands and the same undo journal.                                                                                                                                                       |
| Worksheet structure | `xlsx.sheet.add`, `xlsx.sheet.duplicate`, `xlsx.sheet.rename`, `xlsx.sheet.delete`, `xlsx.sheet.move`, `xlsx.sheet.set_visibility`, `xlsx.sheet.set_tab_color`                                                                                                                                                            | Active Univer workbook/sheet commands and original sheet journal; browser save/reopen covers duplicate, explicit visibility, and native tab color.                                                                   |
| Ribbon UI adapter   | `handleRibbonCommand` (renderer-local)                                                                                                                                                                                                                                                                                    | Visible UI only; Agent execution has no string-command route and uses typed operations.                                                                                                                              |
| Document loading    | `xlsx.document.load_staged` internal                                                                                                                                                                                                                                                                                      | Shared chunk hydration plus original browser open/workbook route.                                                                                                                                                    |
| Persistence         | `xlsx.document.save`                                                                                                                                                                                                                                                                                                      | Original save assembler plus browser package/download adapter.                                                                                                                                                       |

`office_open_local_file` stages bytes and queues internal `xlsx.document.load_staged`. The shared
host bridge hydrates the bytes, and the XLSX registry calls the same `openBuffer` →
`openLazyWorkbook` route as the visible picker. Legacy `open_local_file` remains only an internal
transport alias; neither id appears in Agent discovery or can be called through `office_execute`.

The generic Agent `ribbon_command` was removed in R2-110. Bold, italic,
underline, and strike strings are rejected and must use `xlsx.range.set_text_style`. Alignment,
font appearance, fill, border, and named cell-style strings are rejected in favor of their exact
range operations. Number-format, decimal, merge, clear, and fill strings are likewise rejected in
favor of explicit final-state operations. AutoSum-family strings are rejected in favor of the
bounded aggregate-formula operation. Flash Fill and Text to Columns strings are rejected in favor
of their exact range operations. Row Height and Column Width strings are rejected in favor of
explicit axis operations. Freeze-pane strings and the gridline toggle are rejected in favor of
explicit final sheet-view states. Page orientation, margins, paper size, fit-to-pages, fixed
print scale, printed gridlines/headings, print area, and print-title strings are rejected in favor
of their exact sheet operations. Hyperlink, Format as Table, cell-protection, and
sheet-protection strings are rejected in favor of their five exact operations. `filter-toggle` is
rejected in favor of the explicit AutoFilter range/final-state operation; `filter-clear` and
`filter-advanced` are rejected in favor of exact filter-criteria operations. The context-relative
`insert-sheet`, `insert-row-here`, `delete-row-here`, `insert-col-here`, and `delete-col-here`
strings are rejected in favor of existing exact worksheet/row/column operations. `format-painter`
is also rejected: it arms a one-shot UI
mode whose result depends on the next user selection, so it is not a deterministic Agent operation.
Agent paste-special value, formula, and format strings are rejected in favor of their explicit
bounded range-copy operations; visible Ribbon variants retain Univer's native clipboard route.
Agent `copy`, `cut`, and `paste` are rejected as clipboard/selection-state UI gestures and point to
explicit source/destination operations; visible Univer controls remain native.
The eight retained `outline-*` strings are rejected in favor of absolute level and explicit
detail-visibility operations over bounded 1-based row/column spans.
`insert-checkbox` is rejected in favor of the explicit bounded checkbox-validation operation.
`note-open` and `note-delete` are rejected for Agent execution in favor of explicit note set/remove;
previous/next/show remain non-mutating UI navigation.
Chart type, title, legend, label, grouping, quick-layout, axis-title, delete, palette,
switch-row/column, and Select Data strings are rejected in favor of explicit chart operations.
`insert-symbol` is rejected in favor of the existing explicit final cell-value operation; the
visible Symbol dialog remains a UI producer of that same native write.
`import-csv` is rejected as a file-picker gesture in favor of explicit bounded
`xlsx.range.set_values` or staged CSV open; the visible importer shares the native range action.
The former open-ended Agent route is absent; every retained state-changing producer has a stable
typed operation or an explicit UI-only classification.

`xlsx.document.save` is generated from the XLSX-owned catalog and dispatched by the XLSX registry.
Its exact empty input is validated by the Broker and renderer; success returns the persisted file
identity, while cancellation or write failure returns `execution_failed`. R6-09 removes the legacy
`save` alias; the former direct App/capability branches remain absent.

`xlsx.document.load_staged` is generated from the same catalog with internal visibility. The
renderer validates hydrated `ArrayBuffer` bytes, `.xlsx` identity, and descriptor length before
calling the retained loader. Success returns `{ opened: true, fileName }` without checkpointing a
freshly loaded workbook; parse/load failure returns `execution_failed`. The former direct
`open_local_file` App branch is removed.

`xlsx.cell.set_value`, `xlsx.range.set_values`, `xlsx.row.insert`, `xlsx.row.delete`, and
`xlsx.column.insert` are generated from the XLSX-owned catalog and dispatched by the same registry.
Their schemas bound scalar matrices, 1-based row positions, and structural counts before enqueue;
the renderer resolves the mounted workbook and calls native Univer range/structure methods. R6-09
removes their legacy aliases; the five former hand-written capability entries and direct App
branches remain absent.

`xlsx.column.delete` and `xlsx.sheet.add|rename|delete|move` complete the currently advertised
typed structure surface. Their handlers reuse the mounted Univer workbook, shared bounded column
schema, worksheet-name rules, final-sheet protection, and dynamic 1-based tab positions. R6-09
removes their legacy aliases; their five former capability entries plus composition-root branches
remain absent.

`xlsx.range.set_text_style` begins the `ribbon_command` decomposition. Its explicit field mask
supports boolean bold/italic/strike and `none | single | double` underline. Exact mask validation
prevents implicit toggles and unintended formatting changes; the renderer writes all requested
fields in one native range mutation. The visible Ribbon still uses the retained dispatcher, but
MCP text-style Ribbon strings are rejected so Agent execution has only one canonical path.

`xlsx.range.set_alignment`, `xlsx.range.set_font`, `xlsx.range.set_fill`,
`xlsx.range.set_border`, and `xlsx.range.apply_cell_style` complete the selected alignment, font,
fill/border, and named-style migration groups. Alignment covers horizontal/vertical placement,
wrap, bounded indent, angle/stacked rotation, and explicit clearing. Font covers family, bounded
size, and color with an exact field mask. Fill accepts `#RRGGBB` or explicit clear. Border accepts a
bounded preset and line style, with `none` as the only clear form. Named styles accept only the
retained gallery presets. Multi-field alignment/font/style changes are one native mutation; border
uses the existing native range border route. The real-host tracer proves shared selection,
undo/redo, revision, and save output for all five operations.

`xlsx.range.set_number_format`, `xlsx.range.merge`, `xlsx.range.clear`, and `xlsx.range.fill`
complete R2-37 through R2-40. Number formatting accepts one exact 1–255-character pattern, so
relative decimal UI gestures become deterministic Agent assignments. Merge exposes `cells`,
`across`, `center`, and `unmerge`; center keeps the retained merge-then-align behavior and is marked
non-atomic. Clear separates contents, formats, and all-data removal. Fill requires an explicit
multi-row or multi-column target and invokes the native copy-down/right command after activation.
The real-host tracer proves live selection, visible undo/redo, revision, and saved package output.

`xlsx.range.sort`, `xlsx.range.sort_custom`, and `xlsx.range.remove_duplicates` complete R2-41
through R2-43. Basic sort keeps the retained first-column behavior. Custom sort accepts ordered,
unique A1 column keys inside the target plus an explicit header flag. Both use Univer's native sort
command. Remove duplicates keeps the first case-insensitive whole-row match, reports its removed
count, rejects partially streamed source data before reading, and rewrites only changed rows so
unchanged formulas are not flattened. Real-host evidence covers shared revision, visible
undo/redo, and the saved row order/cleared duplicate row.

`xlsx.hyperlink.set`, `xlsx.hyperlink.remove`, `xlsx.table.add`,
`xlsx.range.set_protection`, and `xlsx.sheet.set_protection` complete R2-44 through R2-48. The two
hyperlink operations normalize the same URL/internal-reference forms as the Ribbon and share its
journal plus link appearance. Table creation uses the retained Format as Table engine and only the
six visible gallery styles; no unsupported existing-table restyle operation is advertised. Cell
protection has an exact `locked | hidden` mask and a 10,000-cell ceiling. Sheet protection takes an
explicit boolean and preserves the password-removal guard. These file-side operations are marked
non-undoable where Univer has no model. Hyperlink save/reopen is complete in R2-78; table and
protection browser package gates remain open.

`xlsx.formula.insert_aggregate` completes R2-49. Its bounded `function` enum replaces
`autofn:SUM|AVERAGE|COUNT|MAX|MIN`; the visible Ribbon and Agent operation share one format-owned
formula seam that places one formula below every selected column. Both routes reject a destination
row whose original file content has not streamed in before writing any formula. Multi-column
insertion preserves the retained per-column Univer mutations and is therefore truthfully marked
non-atomic while remaining natively undoable. Generic formula save/reopen coverage already exercises
this same Univer edit journal and browser OOXML patcher.

`xlsx.range.flash_fill` completes R2-50. Its exact `{ sheet, range }` contract replaces the old
`flash-fill` Ribbon string. The visible Ribbon and Agent operation share one format-owned seam that
reads up to six columns to the left, learns from at most three non-empty target examples, preserves
every existing target value, and applies inferred values to empty target cells through one native
`setValues` mutation. A single-cell target follows the immediately adjacent left column downward
with a hard 1,000-row probe ceiling; explicit multi-row targets remain explicit. Both the probe and
the final inference rectangle fail before reading when original-file rows are still streaming, so
unloaded cells cannot be mistaken for blanks. Existing generic range-value save/reopen evidence
covers the final Univer write.

`xlsx.range.text_to_columns` completes R2-51. Its exact `{ sheet, range, delimiter }` contract
replaces `text-to-columns:1|2|4|8` with the bounded `tab | comma | semicolon | space` enum. The
visible Ribbon and Agent operation share one format-owned seam that sends an explicitly grounded
workbook, worksheet, and single-column range through Univer's native split command. Both routes
retain the command's grouped undo behavior. The registry also requires fully loaded sheet data so
unstreamed cells cannot be read or overwritten. The retained command can overwrite destination
cells to the right and can insert columns when a split exceeds worksheet width; the descriptor is
therefore intentionally high-risk rather than silently changing the visible editor semantics.
Existing Univer range-journal save coverage applies to the resulting cell mutations.

`xlsx.row.set_height` completes R2-52. Its exact `{ sheet, row, count, heightPoints }` contract
replaces `row-height:*` with a 1-based contiguous row span, a 10,000-row ceiling, and the retained
`0–409.5` Excel-point range. The visible Ribbon and Agent operation share one format-owned seam
that converts points to pixels and invokes Univer's forced row-height command. The renderer rejects
spans beyond the mounted worksheet before mutation. Univer supplies one native undo unit, while the
existing axis-attribute listener records point heights for OOXML `ht`/`customHeight` persistence.
The real-host tracer proves Agent execution, user undo/redo, saved output for every targeted row,
and retention after reopen.

`xlsx.column.set_width` completes R2-53. Its exact
`{ sheet, column, count, widthCharacters }` contract replaces `col-width:*` with an explicit A1
start column, bounded contiguous count, and retained `0–255` Excel character-width range. Ribbon
and Agent callers share one native column mutation. Since Univer applies integer pixels and OOXML
stores 1/256-character widths, the operation returns the actual quantized width (for example,
`12.5` requests become `12.4296875`) that the axis journal saves. The renderer rejects spans beyond
the current worksheet before mutation. Real-host evidence proves user undo/redo, complete-span
`width`/`customWidth` output, and preservation after reopen.

`xlsx.sheet.set_freeze` completes R2-54. Its exact
`{ sheet, frozenRows, frozenColumns }` contract replaces `freeze-here`, `freeze-top-row`,
`freeze-first-col`, and `unfreeze` with one explicit final state. `0,0` cancels freezing; non-zero
counts identify the top-left scrollable cell without depending on Agent selection. The renderer
rejects any split that consumes the complete row or column extent. Ribbon and Agent callers share
the native freeze route, and the existing mutation listener keeps the page-setup journal aligned
with user undo/redo. Saved pane coordinates survive reopen.

`xlsx.sheet.set_gridlines` completes R2-55. Its exact `{ sheet, visible }` contract replaces the
state-dependent `toggle-gridlines` string. The visible Ribbon derives the next final boolean and
then shares the same native visibility seam as the Agent. Univer owns undo/redo, while the existing
page-setup journal writes `sheetView@showGridLines`; real-host evidence proves saved and reopened
visibility.

`xlsx.sheet.set_formula_view` completes R2-56. Its exact `{ sheet, enabled }` contract replaces the
state-dependent `toggle-show-formulas` string. The visible Ribbon derives the next boolean and then
shares the retained per-sheet formula projection and `sheetView@showFormulas` journal with the
Agent. The generalized `pushWorkbookUndo` adapter places this renderer-owned state change in
Univer's history; Undo and Redo reapply both the visible projection and persisted final state.
Real-host evidence proves removal of `showFormulas` after undo, saved `showFormulas="1"` after the
redo cycle, and preservation after reopen.

`xlsx.sheet.set_page_orientation` completes R2-57. Its exact `{ sheet, orientation }` contract
replaces `page-layout:orientation:portrait|landscape` with a bounded final state. Ribbon and Agent
callers share one page-layout seam and the existing OOXML page-setup journal. Renderer-owned history
restores the exact prior journal field, including absence when the source file had no override, and
Redo reapplies the requested final orientation. Real-host evidence proves removal of the
`orientation` attribute after undo, saved `orientation="landscape"` after the redo cycle, and
preservation after reopen.

`xlsx.sheet.set_page_margins` completes R2-58. Its exact `{ sheet, margins }` contract replaces
`page-layout:margins:normal|wide|narrow` with one bounded final-state preset. Ribbon and Agent
callers share the same XLSX-owned seam, while a private page-setup preset module owns file fallback,
journal field replacement, idempotence, and exact renderer history for both margins and orientation.
The retained OOXML writer expands `wide` to left/right/top/bottom `1` and header/footer `0.5`.
Real-host evidence proves removal of a newly introduced `pageMargins` element after Undo, the six
saved wide-margin values after Redo and through the visible Ribbon, and preservation after reopen.

`xlsx.sheet.set_paper_size` completes R2-59. Its exact `{ sheet, paperSize }` contract replaces the
seven retained `page-layout:paper:*` strings with an integer enum limited to Letter (1), Tabloid
(3), Legal (5), Executive (7), A3 (8), A4 (9), and A5 (11). Ribbon and Agent callers share
`applyWorkbookPaperSize` and the private page-setup preset module, so file fallback, idempotence,
journal replacement, and renderer-owned history remain one deep seam. Real-host evidence proves
removal of a newly introduced paper-size override after Undo, saved `paperSize="9"` after Redo and
through the visible A4 Ribbon preset, and preservation after reopen.

`xlsx.sheet.set_fit_to_pages` completes R2-60. Its exact
`{ sheet, widthPages, heightPages }` contract replaces both retained fit-width and fit-height
strings with bounded `0–1000` integer axes; `0` means Automatic. Agent calls explicitly provide
both axes, while each single-axis Ribbon control resolves the other axis from effective file/journal
state before entering the shared seam. The private page-setup patch module snapshots all three
affected journal fields—width, height, and `fitToPage`—as one exact Undo/Redo unit. Real-host
evidence proves removal of newly introduced fit overrides after Undo, saved `fitToWidth="2"`,
`fitToHeight="3"`, and `fitToPage="1"` after Redo and through the retained Ribbon, and preservation
after reopen. `0,0` explicitly disables fit-to-page.

`xlsx.sheet.set_print_gridlines` completes R2-61. Its exact `{ sheet, enabled }` contract replaces
retained `page-layout:print-gridlines:0|1` strings without conflating printed gridlines with
`xlsx.sheet.set_gridlines`, which controls worksheet display. Agent and Ribbon callers share
`applyWorkbookPrintGridlines` and the private page-setup patch seam, so file fallback, idempotence,
journal replacement, and renderer-owned history remain localized. Real-host evidence proves Agent
and visible Ribbon mutation, exact Undo/Redo, removal of a newly introduced
`printOptions@gridLines` override after Undo, explicit disable, and preservation after reopen.

`xlsx.sheet.set_print_headings` completes R2-62. Its exact `{ sheet, enabled }` contract replaces
retained `page-layout:print-headings:0|1` strings with a final printed row/column-heading state.
Agent and Ribbon callers share `applyWorkbookPrintHeadings` and the private page-setup patch seam,
so file fallback, idempotence, journal replacement, and renderer-owned history remain localized.
Real-host evidence proves Agent and visible Ribbon mutation, exact Undo/Redo, removal of a newly
introduced `printOptions@headings` override after Undo, explicit disable, and preservation after
reopen.

`xlsx.sheet.set_print_area` completes R2-63. Its exact `{ sheet, range }` contract replaces retained
`page-layout:print-area:set|clear` strings with a normalized explicit A1 cell range or `null`.
Agent input is validated at the registry seam; the retained Ribbon resolves the live selection
before both routes share `applyWorkbookPrintArea` and the private page-setup patch/history seam.
Real-host evidence proves Agent and visible Ribbon mutation, exact Undo/Redo, removal after Undo or
explicit clear, saved `_xlnm.Print_Area` absolute references, and preservation after reopen.

`xlsx.sheet.set_print_titles` completes R2-64. Its exact `{ sheet, rows }` contract replaces
retained `page-layout:print-titles:first-row|set|clear` strings with an ascending explicit row span
or `null`. The renderer rejects malformed, reversed, out-of-sheet, and more-than-21-row spans; the
visible Ribbon resolves Row 1 or the live selected rows before both routes share
`applyWorkbookPrintTitles` and the private page-setup patch/history seam. Real-host evidence proves
Agent and visible Ribbon mutation, exact Undo/Redo, removal after Undo or explicit clear, saved
`_xlnm.Print_Titles` absolute row references, and preservation after reopen.

`xlsx.sheet.set_print_scale` completes R2-65. Its exact `{ sheet, scalePercent }` contract replaces
retained `page-layout:scale:10..400` strings with an explicit bounded integer. Agent and the
retained page-layout dispatcher share `applyWorkbookPrintScale`, which records scale plus
`fitToPage: false` through the private multi-field page-setup patch/history seam as one Undo/Redo
unit. Real-host evidence starts from a fit-to-pages workbook and proves restoration after Undo,
fixed-scale output after Redo, removal of `pageSetUpPr@fitToPage`, saved `pageSetup@scale="80"`, and
preservation after reopen. The pinned Page Layout Ribbon does not expose a visible Scale control,
so no new UI was invented for this migration.

`xlsx.range.set_filter` completes R2-66. Its exact `{ sheet, range, enabled }` contract replaces
the retained `filter-toggle` string with an explicit target and final state. Replaying a satisfied
state is a no-op instead of another toggle; a conflicting active filter range fails closed. Agent
and Ribbon callers share `filter-actions.ts`, while the Ribbon resolves the active filter or smart
selection before entering Univer's native set/remove commands. Real-host evidence proves Agent and
visible Ribbon mutation, native Undo/Redo, removal after Undo, persisted
`<autoFilter ref="A1:C4">`, and preservation after reopen. The browser host now writes declarative
filter snapshots through the existing `xlsx-filter` gateway and restores saved filter ranges when
opening a workbook.

`xlsx.range.clear_filter_criteria`, `xlsx.range.set_filter_values`, and
`xlsx.range.set_custom_filter` complete R2-67 through R2-69. They retain an exact AutoFilter range,
address columns by absolute A1 label, bound value lists to 10,000 entries and custom conditions to
two, and reject missing/mismatched filters before mutation. Agent, Ribbon Clear, and the Advanced
Filter dialog share `filter-actions.ts` over Univer's native filter criteria commands and Undo
stack. Real-host evidence covers Agent value/custom criteria, Ribbon clear plus Undo, dialog custom
criteria, filtered row visibility, saved `<filterColumn>` output, and preservation after reopen.

## Remaining product-completion gates

R2-131's executable audit maps every retained state-changing Ribbon, dialog, native grid, visual,
table, Pivot, row/column, and worksheet gesture to an exact catalog operation. The browser adapter
now saves and reopens pivots, Conditional Formatting, drawings/media, notes, validation, tables,
page setup, row/range movement, sheet duplication/visibility, and tab color. File pickers, clipboard
arming, find/navigation, zoom, trace arrows, statistics, screen capture, and PDF export are transient
UI or external-output gestures rather than retained XLSX document mutations; deterministic Agent
work uses the explicit operations documented above.

R2-70 retires the five context-relative structural strings from Agent execution. The public
MCP-to-editor path now rejects them with their existing exact replacements before selection-relative
dispatch, while visible sheet/row structure controls retain their native Undo/Redo and save/reopen
behavior. No registry operation was added, so the XLSX operation count remains fifty-one.

R2-71 adds `xlsx.history.undo {}` and `xlsx.history.redo {}` as exact operations over the mounted
Univer history. Agent edits, visible UI controls, and the two registry handlers share one native
undo/redo stack. Empty stacks fail without advancing revision; successful calls each advance one
revision, and the real-host tracer proves edit → undo → redo → save on one mounted workbook. The old
Agent `ribbon_command` strings are rejected, raising the XLSX registry count to fifty-three.

R2-72 adds `xlsx.range.copy_values` as the first explicit clipboard replacement. It accepts exact
source/destination worksheets and same-shaped A1 ranges, copies computed scalar values rather than
formulas, supports cross-sheet targets, caps each operation at 20,000 cells, and rejects source or
destination rectangles that are still streaming. The real-host tracer proves formula-to-value
materialization, shared UI Undo/Redo, one Agent revision, and saved OOXML without a destination
formula. Agent `paste-special:value` is rejected, raising the XLSX registry count to fifty-four.

R2-73 adds `xlsx.range.copy_formulas` over the same explicit, bounded range-pair seam. It copies
formula cells with Univer-native row/column reference translation, copies non-formula cells as
scalars to match Paste Formulas behavior, and preserves cross-sheet targets. Relative, absolute,
and mixed references are covered by the real-host tracer; one destination matrix write shares UI
Undo/Redo and persists translated `<f>` elements. Agent `paste-special:formula` is rejected, raising
the XLSX registry count to fifty-five.

R2-74 adds `xlsx.range.copy_formats` through that same bounded range-pair seam. It replaces every
destination cell-style field from the source matrix while retaining destination scalars, formulas,
and rich content. The renderer commits one low-level Univer mutation plus its exact inverse as a
single native history item; the real-host tracer proves font/fill/alignment/wrap/number-format
copy, visible Undo/Redo, preserved content, and saved OOXML. Agent `paste-special:format` is
rejected, raising the XLSX registry count to fifty-six. Merged-cell topology stays explicit through
`xlsx.range.merge`.

R2-75 adds `xlsx.column.copy_widths` with explicit source/destination worksheets, normalized A1
start columns, and a bounded 1–10,000 count. Both spans are checked independently before any source
width is read. The renderer snapshots the complete source vector and applies one object-backed
native mutation with one inverse history item, so same-sheet overlap and cross-sheet copies remain
deterministic. The real-host tracer proves two distinct widths, visible Undo/Redo, unchanged scalar
and formula content, and saved OOXML equality. Agent `paste-special:col-width` is rejected, raising
the XLSX registry count to fifty-seven. Column-width metadata is eagerly installed and therefore
does not depend on cell-range streaming state.

R2-76 adds `xlsx.range.copy_without_borders` through the shared bounded range-pair seam. It copies
source scalar/rich content, translates relative and mixed formula references, and replaces all
non-border style fields while deliberately omitting `bd`, so Univer retains each destination
border. Inactive-sheet formulas use the read-only file formula index only as a missing-model
fallback, with current edit-journal content taking precedence. The real-host tracer proves
cross-sheet copy, visible Undo/Redo, saved translated formulas, source non-border formatting, and
the original target border. Agent `paste-special:besides-border` is rejected, raising the XLSX
registry count to fifty-eight.

R2-77 rejects Agent `copy`, `cut`, and `paste` because their behavior depends on transient
clipboard and current-selection state. The visible Ribbon controls still execute Univer's native
commands, while Agent callers use explicit source/destination operations. The real-host rejection
tracer covers all three strings; no descriptor is added, so the XLSX registry remains at
fifty-eight operations.

R2-78 connects the existing hyperlink journal to the browser package adapter. Worksheet link
elements and external relationship files are still written by `xlsx-hyperlinks.ts`; browser open
now indexes external targets and internal locations into the same streamed link context used by
the mounted App. The real-host tracer proves typed set/save, exact OOXML output, reopen, and user
click behavior; focused host tests cover set and remove across reopen. The registry count remains
fifty-eight.

R2-79 connects `xlsx.table.add` to the browser package adapter through a transactional overlay over
the retained table gateway. The real-host tracer proves native table part, relationship,
worksheet, and content-type output plus reopen; validation cannot partially commit package edits.
R2-80 resolves those table relationships during browser open and hydrates range/header/style and
stripe metadata into the retained Univer table path, so native and newly saved tables reopen with
table behavior rather than as plain cells.

R2-81 connects the existing cell/sheet protection operations to browser save/reopen. Cell flags
continue through the retained stylesheet editor, while sheet protection uses
`xlsx-protection.ts`; browser open records both the protected bit and password presence. The
real-host tracer proves exact cell XF and worksheet output, reopen, and an explicit unprotect
request over the hydrated state. The registry count remains fifty-eight.

R2-82 adds explicit `xlsx.sparkline.add` over same-height source and one-column target ranges. The
operation caps output at 200 cells, rejects overlap and occupied targets, and shares the retained
journal, renderer-owned Undo/Redo, and float-DOM projection with the Ribbon. Browser save delegates
to `xlsx-sparkline.ts`; reopen hydrates group types, colors, source refs, and host cells. The
real-host tracer proves visible Agent output, exact x14 XML, save, and reopen. Agent `sparkline:*`
strings are rejected, and the registry count rises to fifty-nine.

R2-83 adds `xlsx.outline.set_level` and `xlsx.outline.set_detail_visibility`. The first writes an
absolute level rather than replaying Group/Ungroup deltas; the second writes a final hidden state
and the immediately following summary item's collapsed state. Ribbon and Agent share one action
layer and one renderer-owned Univer history item. Browser open now hydrates row and column outline
metadata, and the real-host tracer proves collapse/save/reopen/expand/save. The eight legacy Agent
strings are rejected and the registry count rises to sixty-one.

R2-84 adds `xlsx.range.set_checkbox` with an explicit final boolean over at most 10,000 cells.
Ribbon and Agent share Univer's native validation command and Undo stack. The browser package
adapter now writes declarative validation snapshots through `xlsx-dv.ts`, parses base OOXML rules
on reopen, and maps the saved `list "1,0"` representation back to a checkbox. The legacy Agent
string is rejected and the registry count rises to sixty-two.

R2-85 retires Agent `insert-symbol` without adding a descriptor. The visible dialog still appends
the selected character, but its resolved final string and `xlsx.cell.set_value` now share one
native cell write/activation action. Deterministic Agent callers assign the explicit final value;
the registry remains at sixty-two operations.

R2-86 retires Agent `import-csv` without adding a descriptor. The visible browser importer retains
file selection, decoding, parsing, coercion, and its 50,000-cell UI bound, while the resolved matrix
and `xlsx.range.set_values` converge on one native range write. Deterministic Agent input uses an
explicit matrix or staged CSV open; the registry count remains sixty-two.

R2-87 adds `xlsx.chart.add` with nine explicit chart types and a 2,000-cell source bound. The exact
operation and retained chart pickers share the visual journal, renderer projection, and native
history route. The browser adapter now writes visual additions transactionally through the retained
drawing gateway; a real-host tracer proves rendering, exact chart/drawing OOXML, typed save, and
reopen. The registry count rises to sixty-three.

R2-88 adds public `xlsx.image.add` plus internal `xlsx.image.add_staged`. Broker staging keeps image
bytes out of Agent JSON and enforces absolute PNG/JPEG/GIF paths, matching magic bytes, and a 20 MB
limit. The four retained image producers and Agent path share the visual journal, Undo route, and
browser drawing writer. The registry count rises to sixty-five; staged transport is undiscoverable.

R2-89 adds `xlsx.note.set` and `xlsx.note.remove` over explicit single-cell addresses and a
32,767-character runtime bound. Both use Univer's native note commands/history and the retained
declarative journal. Browser open/save now parses and rewrites legacy comments and VML anchors; a
real-host tracer proves set, typed save, exact OOXML, reopen, native removal, and second save. The
registry count rises to sixty-seven.

R2-90 adds `xlsx.chart.update` and `xlsx.chart.remove` over the stable visual ID returned by chart
insertion or supplied by file metadata. Update accepts bounded final-state title, six convertible
types, legend, data labels, grouping, and axis titles; remove uses the existing visual edit seam.
Both retain renderer history. Browser save now applies file chart and visual edits; real-host
tracers prove session update/save and removal-before-save. The registry count rises to sixty-nine.

R2-91 adds explicit `xlsx.chart.set_colors` and `xlsx.chart.set_series`. Color arrays are bounded
to 24 series and 64 pie/doughnut points with exact hex validation. Complete series replacement is
bounded to 24 series and 1,000 finite values per series, with optional same-length categories and
range refs. The operations replace palette, switch-row/column, and Select Data Agent strings while
reusing chart history; real-host save proves final caches and color XML. The registry count rises
to seventy-one.

R2-92 adds no descriptor. Browser open now resolves worksheet drawing and chart relationships,
hydrates chart anchors/metadata/save locators, and assigns deterministic `file-chart-*` IDs that
are published in live selection context. `xlsx.chart.update` gains bounded advanced format-pane
properties for gridlines, value-axis min/max, gap width, hole size, pie explosion, and data-label
position/number format. A real-host tracer proves save, reopen, stable-ID update, and second save.

R2-93 adds `xlsx.shape.add` over the exact retained gallery preset enum plus `textbox`. The
operation bounds text to 1,000 characters, validates explicit hex fill and cell anchor, and reuses
the visual journal/history installed for UI shapes. Browser open now hydrates shape preset, text,
fill, anchor, drawing locator, and stable ID; the real-host tracer proves Undo/Redo and save/reopen.

R2-94 adds explicit shape update/remove over session and reopened IDs. Update accepts bounded text,
hex fill, and an explicit top-left cell while preserving frame size. The visual edit journal and
OOXML gateway now patch file-native shape text/fill/anchor, or remove the anchor, with exact
renderer Undo. Real-host evidence covers two saves, two reopens, and final removal.

R2-95 resolves drawing image relationships during browser open and serves exact package bytes to
the retained image renderer. Stable `file-image-*` IDs drive explicit move/remove operations; the
visual journal preserves frame size, Undo, and drawing output. Real-host evidence covers reopen
rendering, movement save, second reopen, removal, and final save.

R2-96 adds explicit Defined Name upsert/rename and removal with workbook or named-sheet scope.
`defined-name-actions.ts` is the shared action layer for the retained Name Manager and registry,
so both paths use Univer's native model/history mutations and the existing full-state save
snapshot. Browser workbook tests prove workbook- and sheet-scoped XML save, reopen, update, and
removal. Agent `name-manager-open` now fails closed with the exact operation replacements.

R2-97 adds bounded inline-list validation and general validation removal. The existing checkbox
action, new list/remove actions, and registry now share one target/streaming/metadata safety gate
and Univer's native validation history. Browser tests prove list OOXML save/reopen and rule
removal. Agent `dv-open` is rejected in favor of explicit validation operations.

R2-98 adds numeric-between, date-between, and custom-formula validation. Runtime validation rejects
non-finite/reversed numbers, impossible/reversed ISO dates, and malformed or oversized formulas
before mutation. All three share the same native validation/history and declarative OOXML mapping
already exercised by the focused validation and browser round-trip suites.

R2-99 adds range-backed list validation with explicit target/source sheet names and ranges. The
source must be a single row or column of at most 1,000 loaded cells; the target uses the shared
10,000-cell and DV-metadata guards. Browser save/reopen verifies the referenced list formula.

R2-100 consolidates the scalar comparison matrix into `xlsx.range.set_comparison_validation`.
Whole, decimal, date, time, and text-length kinds share eight operators, strict type/arity/order
validation, the native panel rule model, and FRange's Undo command. The R2-98 number/date-between
descriptors are removed rather than retained as duplicate tools; custom-formula validation remains
separate because it has different semantics.

R2-101 exposes the validation panel's prompt and error metadata as one explicit nullable final
state. It bounds Excel title/message lengths, requires consistent disabled error state, resolves an
existing rule, and reuses FDataValidation's native update-options/Undo plus declarative save mapper.

R2-102 starts Conditional Formatting with one comparison lifecycle contract rather than an
operator matrix. `ruleId: null` creates; a published session ID updates only an existing comparison
rule; exact-ID removal is separate. The 10,000-cell and streaming/metadata guards run before native
Univer add/set/delete commands. Context exposes up to 100 rule IDs, and the browser package adapter
now persists CF snapshots and DXFs through the same gateway used by desktop save.

R2-103 adds one highlight lifecycle contract for text contains/not-contains/starts-with/ends-with,
blank/non-blank, and duplicate/unique rules. An explicit nullable text operand eliminates hidden
arity, while the comparison tracer's style, rule identity, target loading, native Undo, and save
boundaries are reused.

R2-104 adds one statistical lifecycle contract for top/bottom amount or percent and above/below
average with optional equality. Explicit nulls make the two parameter families mutually exclusive;
all variants share the existing native add/set/Undo and CF/DXF save route.

R2-105 adds a bounded equals-prefixed custom-formula lifecycle operation over the same native and
persistence route.

R2-106 adds one visual lifecycle operation for bounded color scales, base data bars, and saveable
icon sets. Typed thresholds and kind-specific nullable fields reject lossy x14-only states before
mutation while reusing native commands and the tested CF serializer.

R2-107 adds text equal/not-equal and error/non-error to the existing highlight contract, including
equivalent reopen hydration. It closes the UI/save drift for solid and distinct-negative-color data
bars; those and the already-guarded date/average/icon variants are explicitly x14-only.

R2-108 adds explicit range/worksheet Conditional Formatting clearing, and R2-109 adds absolute
one-based rule priority over the bounded session rule list. Both retain native Univer history.
R2-110 removes the open-ended Agent Ribbon operation entirely; the visible Ribbon remains a UI
adapter, while Agent execution is restricted to the ninety public typed XLSX operations.

R2-111 through R2-130 finish the post-Ribbon capability audit. They connect Pivot definition
loading plus add/refresh/update/member-filter/PivotChart operations; explicit formula writes;
worksheet duplicate/visibility; row/column visibility; full table row/column/convert lifecycle;
bounded text replacement, subtotals, consolidation, header/footer; native range and row movement;
and worksheet tab color. The final browser-adapter slice removes its duplicate/visibility rejection
and proves those operations through save/reopen. R2-131 adds an executable retained-command audit,
regenerates the 138-operation product Manifest, and records 114 XLSX operations: 112 public and two
internal staged-load operations.

The XLSX mutation-parity gates are closed. R6-01 completes the repository-wide packaged-host
visual comparison, runtime/resource measurements, MCP smoke, and all-format acceptance. R6-04
adds a stricter format-local startup boundary: the first poll occurs only after an active workbook,
active `Sheet1`, and Univer canvas are observable, and carries the bounded bootstrap,
Univer-create, worksheet-install, and first-commit phases once. R6-05 partitions bootstrap into
resource-receive, initial-module-graph, and React-mount durations. A failed transport retries the
same trace; later successful polls omit it. Schema-v4 release evidence requires seven aggregate and
subphase samples, a fixed 500 ms bootstrap p95 ceiling, and the retained fixed 1,400 ms total XLSX
cold-start p95 ceiling. All retained presets, native Univer history, Registry operations, and
browser package save/reopen routes remain unchanged.

The restored self-contained XLSX resource is 6,485,773 raw bytes / 4,793,010 gzip bytes. Every
split JavaScript module is still compressed and embedded in that single HTML resource; the initial
entry graph inflates to about 10.08 MB and is capped at 11 MB, while optional locale/hyphenation
modules inflate only on demand. These budgets are regression guardrails and cannot justify deleting
renderer capabilities.
