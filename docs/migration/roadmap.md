# Complete renderer migration roadmap

- Status: Complete through R6-10
- Baseline: `genspark-ai/genoffice@dc4d7e5927864498913b7ba42d0da06cc7cf628e`
- Governing decisions: [ADR 0003](../adr/0003-complete-community-renderers-and-mcp-parity.md), [ADR 0004](../adr/0004-format-owned-operation-registries.md), [ADR 0007](../adr/0007-traced-bounded-markdown-staged-load.md), [ADR 0008](../adr/0008-traced-bounded-xlsx-cold-start.md), and [ADR 0010](../adr/0010-immutable-editor-view-leases-and-exact-session-resume.md)
- Strategy: source-preserving complete renderer restoration plus format-owned operation registries, delivered format by format

## Objective

Ship the complete pinned non-AI community DOCX, XLSX, PPTX, PDF, and Markdown renderers in one Codex-compatible TandemFolio plugin. Preserve their original editing behavior while replacing only Electron main/preload/IPC and file-system boundaries with the existing browser/MCP host adapters.

Every retained state-changing renderer command must be callable through MCP against the same mounted editor state, selection, undo stack, and monotonically increasing revision. The project does not introduce a separate document Runtime, hidden headless editor, enterprise source, or capabilities outside the pinned community applications.

## Corrected decision

The former M3–M6 strategy accepted narrow XLSX, PPTX, and PDF command tracers and TandemFolio-owned renderer replacements as runnable delivery slices, and removed Markdown from the product workspace. ADR 0003 rejects that strategy as a product endpoint.

The diagnostic replacement scaffolds have been removed now that each matching community renderer is connected. Work must not recreate or polish a simplified replacement in place of the pinned renderer. The earlier removal of Markdown was not an acceptable first-release limitation; its source and package restoration are now present and must be completed through MCP parity.

## Operation integration route

ADR 0004 changes how the remaining MCP parity work is implemented without changing renderer authority or the required feature scope. Before broad new capability families are added, the current executable surface must migrate into five format-owned operation registries according to [`operation-registry-plan.md`](operation-registry-plan.md).

The R1 format-neutral contract and manifest generator are complete as of 2026-08-28. R2-01 through R2-04 connect `markdown.text.insert`, `markdown.text.replace_selection`, `markdown.document.save`, and internal `markdown.document.load_staged` as the first complete registry vertical tracer. R2-05 through R2-08 add `docx.text.insert`, `docx.text.replace_selection`, `docx.document.save`, and internal `docx.document.load_staged` as the first cross-format proof: one product manifest aggregates format-owned catalogs, the broker scopes shared legacy aliases by session format, and the mounted DOCX renderer executes text mutations through TipTap history, save through its format-owned persistence/recovery path, and staged load through its format-owned open path. R2-09 through R2-13 split replace-all, heading-level, block-delete, and list apply/remove out of DOCX's shallow batch route with exact schemas and native undo. R2-14 through R2-18 complete that decomposition with `docx.text.set_style`, `docx.paragraph.set_style`, `docx.block.move`, `docx.image.update`, and `docx.toc.insert`; the common contract adds bounded nullable types and `minItems`, while the obsolete `batch_update` schema, App branch, and helper are deleted. R2-19 and R2-20 establish XLSX document operations: `xlsx.document.save` uses the retained save assembler/browser package boundary, and internal `xlsx.document.load_staged` uses the retained staged-open path. R2-21 through R2-25 add `xlsx.cell.set_value`, `xlsx.range.set_values`, `xlsx.row.insert`, `xlsx.row.delete`, and `xlsx.column.insert` through the mounted Univer runtime and native undo journal. R2-26 through R2-30 complete column deletion and the worksheet structure family as `xlsx.column.delete` plus `xlsx.sheet.add|rename|delete|move`, preserving name bounds, final-sheet protection, dynamic positions, and native undo. R2-31 starts decomposing `ribbon_command` with `xlsx.range.set_text_style`: bold, italic, strike, and bounded underline values use one explicit masked Univer mutation, while their old MCP Ribbon strings are rejected. R2-32 through R2-36 add exact alignment, font, fill, border, and named cell-style operations, reject their old MCP Ribbon strings, and classify `format-painter` as a transient UI-only mode. R2-37 through R2-40 add explicit number-format assignment, merge modes, clear scopes, and directional fill. R2-41 through R2-43 add basic sort, ordered multi-key/header sort, and guarded row deduplication, rejecting all migrated Ribbon strings while preserving mounted Univer state, native undo, and save output. R2-44 through R2-48 add exact hyperlink set/remove, table creation, cell-protection, and sheet-protection operations over the same retained renderer journals and table engine; their browser package save/reopen gate remains open. R2-49 adds `xlsx.formula.insert_aggregate` for the five retained AutoSum functions through a shared Ribbon/registry seam with streamed-target protection. R2-50 adds `xlsx.range.flash_fill` through a shared example-inference seam with empty-target-only writes, a 1,000-row single-cell probe ceiling, and pre-read streaming protection. R2-51 adds `xlsx.range.text_to_columns` for the four retained delimiter modes through Univer's native split/undo command, with single-column and full-load guards. R2-52 adds `xlsx.row.set_height` with an explicit bounded row span and final point height through the shared native row-sizing/undo/save route. R2-53 through R2-55 add quantized explicit column width, explicit freeze-pane counts, and explicit gridline visibility through shared native layout/view seams with undo and save/reopen evidence. R2-56 adds explicit per-sheet formula view through the retained projection/page-setup journal and generalizes renderer-owned custom history as `pushWorkbookUndo`. R2-57 adds bounded per-sheet print orientation through the shared page-layout journal and exact renderer-owned undo/redo. R2-58 adds bounded page-margin presets through the same journal and deepens their shared file/journal/history module. R2-59 adds the seven retained paper-size presets through the same deep seam with exact numeric validation, history, and saved OOXML output. R2-60 combines retained fit-width and fit-height commands into an explicit dual-axis final state, using the deepened multi-field page-setup patch/history seam and saved OOXML fit output. R2-61 separates printed-gridline output from worksheet display gridlines through an explicit boolean final state, the shared page-setup patch/history seam, and saved OOXML print options. R2-62 adds the corresponding explicit printed-heading state through the same seam and persisted `printOptions@headings`. R2-63 adds a nullable explicit A1 print area through the same renderer-owned history seam and persisted `_xlnm.Print_Area` defined name. R2-64 adds bounded nullable print-title rows through the shared history seam and persisted `_xlnm.Print_Titles` defined name. R2-65 adds bounded fixed print scale, disables fit-to-page in the same undo unit, and persists `pageSetup@scale`. R2-66 replaces the implicit filter toggle with an explicit range/final-state AutoFilter operation, shares Univer's native set/remove commands with the Ribbon, and connects filter snapshots to the browser package save/reopen path. R2-67 through R2-69 add exact all-criteria clear, bounded value-list criteria, and bounded two-condition custom criteria; the Contract adds `maxItems`, while Agent, Ribbon Clear, and Advanced Filter converge on the same native filter/Undo/save path. R2-70 rejects the five context-relative sheet/row/column structure strings before Agent dispatch and points callers to the existing exact operations without growing the registry. Fifty-one XLSX registry operations are generated and dispatched; `ribbon_command` remains transitional only for families not yet migrated.

R2-71 adds `xlsx.history.undo {}` and `xlsx.history.redo {}` over the mounted Univer native history,
retires the old Agent Ribbon strings, and raises the generated XLSX registry to fifty-three
operations. Visible history controls continue to share the same stack.

R2-72 adds bounded explicit `xlsx.range.copy_values` source/destination semantics, materializes
formula results as scalar values through one native undoable write, rejects partially streamed
ranges, and retires Agent `paste-special:value`. The generated XLSX registry now contains
fifty-four operations.

R2-73 adds bounded explicit `xlsx.range.copy_formulas` source/destination semantics, translates
relative and mixed formula references through Univer's mounted lexer, shares the value-copy range
and streaming boundary, and retires Agent `paste-special:formula`. The generated XLSX registry now
contains fifty-five operations.

R2-74 adds bounded explicit `xlsx.range.copy_formats` source/destination semantics, replaces the
complete destination cell-style matrix without changing values or formulas, shares the copy-family
range and streaming boundary, and commits one native Univer mutation/Undo unit. Agent
`paste-special:format` is retired while the visible Ribbon keeps its native clipboard route. The
generated XLSX registry now contains fifty-six operations.

R2-75 adds explicit `xlsx.column.copy_widths` source/destination column spans with a bounded
1–10,000 count and independent worksheet-column bounds. It copies the complete source width vector
through one native Univer mutation/Undo unit, preserves cell contents, supports cross-sheet targets,
and persists through the existing axis journal. Agent `paste-special:col-width` is retired while the
visible Ribbon keeps its native clipboard route. The generated XLSX registry now contains
fifty-seven operations.

R2-76 adds explicit `xlsx.range.copy_without_borders` source/destination ranges with the shared
worksheet-bounded A1, equal-shape, 20,000-cell, cross-sheet, and dual-streaming guards. It copies
cell content, translates formulas, replaces every non-border style field, and deliberately retains
the destination border through Univer's native style merge and one Undo/Redo unit. Agent
`paste-special:besides-border` is retired while the visible Ribbon keeps its native clipboard route.
The generated XLSX registry now contains fifty-eight operations.

R2-77 retires Agent `copy`, `cut`, and `paste` as clipboard/selection-state UI gestures. Their
visible Ribbon controls continue to use Univer's native clipboard commands, while deterministic
Agent work must use explicit source/destination operations. No descriptor is added, so the XLSX
registry remains at fifty-eight operations.

R2-78 connects the existing hyperlink edit journal to the browser package adapter. Typed hyperlink
set/remove now reuse the retained OOXML worksheet/relationship gateway, save through
`xlsx.document.save`, hydrate again on browser open, and drive the same user click behavior after
reopen. No descriptor is added.

R2-79 through R2-81 close the table and protection browser-package gates. Typed table additions
write native table parts transactionally; browser reopen hydrates native table metadata into the
retained Univer table path; cell protection persists through the stylesheet editor; and
passwordless sheet protection saves and rehydrates through the retained OOXML gateway. The XLSX
descriptor count remains fifty-eight.

R2-82 migrates the retained sparkline family to explicit source/target ranges and one exact type.
Ribbon and Agent share the existing journal, renderer-owned Undo, and float-DOM projection; browser
save/reopen uses the retained x14 gateway. The XLSX descriptor count is now fifty-nine.

R2-83 migrates all eight retained outline strings to absolute bounded level and explicit final
detail-visibility operations. Ribbon and Agent share one renderer-owned history action; browser
save/reopen hydrates row/column outline state. The XLSX descriptor count is now sixty-one.

R2-84 migrates checkbox insertion to an explicit bounded final-state operation and connects the
declarative data-validation journal to browser save/reopen. The XLSX descriptor count is now
sixty-two.

R2-85 classifies Symbol insertion as a retained UI text-picker over the existing explicit
`xlsx.cell.set_value` capability. Both routes share one native write; no descriptor is added.

R2-86 classifies CSV import as a retained browser file-picker/decoder over the existing explicit
`xlsx.range.set_values` capability. Both routes share one native matrix write; no descriptor is
added.

R2-87 migrates chart creation to bounded `xlsx.chart.add`. Ordinary and recommended-chart pickers
share the visual journal and Undo route; the browser package adapter now persists visual additions
through the retained transactional drawing writer. The XLSX descriptor count is now sixty-three.

R2-88 migrates picture/screenshot/icon/equation production to public path-based `xlsx.image.add`
and internal staged bytes. It reuses visual journal/Undo/browser media output and keeps Base64 out
of Agent arguments. The XLSX descriptor count is now sixty-five.

R2-89 migrates cell-note mutation to `xlsx.note.set` and `xlsx.note.remove`. Explicit addresses
replace popup/current-selection Agent intent; both handlers use Univer native history and the
existing note journal. Browser comments/VML save and reopen hydration are connected. The XLSX
descriptor count is now sixty-seven.

R2-90 adds explicit-ID `xlsx.chart.update` and `xlsx.chart.remove`. The operations reuse retained
chart/visual history and browser chart/visual edit persistence; title/type/legend/labels/grouping/
axis-title commands and delete leave `ribbon_command`. Palette and switch-row/column remain for
R2-91. The XLSX descriptor count is now sixty-nine.

R2-91 adds explicit bounded `xlsx.chart.set_colors` and `xlsx.chart.set_series`. Palette,
switch-row/column, and Select Data Agent strings leave `ribbon_command`; the retained UI still
derives its final arrays before using the same chart edit/history seam. The XLSX descriptor count
is now seventy-one.

R2-92 keeps the descriptor count at seventy-one and completes chart reopen identity in the
browser host. It hydrates worksheet drawing/chart relationships into stable file-chart IDs and
extends `xlsx.chart.update` with bounded advanced format-pane final states. A real-host tracer
proves typed save, reopen, context discovery, second edit, and second OOXML save.

R2-93 adds bounded `xlsx.shape.add` for the retained shape gallery and text boxes. Explicit sheet
and cell anchors replace draw-mode/current-selection Agent intent; the operation shares the visual
journal and Undo route. Browser save/reopen now hydrates native shapes with stable IDs. The XLSX
descriptor count is seventy-two.

R2-94 adds explicit `xlsx.shape.update` and `xlsx.shape.remove`. Session and reopened shape IDs
share one visual history route; the browser drawing gateway now persists file-native text, fill,
movement, and removal. The XLSX descriptor count is seventy-four.

R2-95 completes browser image hydration and adds explicit `xlsx.image.move/remove`. Reopened image
IDs resolve package media bytes through the browser adapter; movement and removal share visual
Undo and drawing save. The XLSX descriptor count is seventy-six.

R2-96 adds explicit `xlsx.defined_name.set/remove` over workbook or named-sheet scope. Set supports
upsert and atomic rename through `previousName`; both Agent and retained Name Manager paths use the
same Univer model/history and declarative browser save/reopen route. The XLSX descriptor count is
seventy-eight.

R2-97 adds explicit inline-list validation set and general validation removal. It bounds target
cells, option count, and the XLSX inline source, then reuses native validation history and the
declarative save/reopen route. The XLSX descriptor count is eighty.

R2-98 adds bounded numeric-between, ISO-date-between, and custom-formula validation operations over
the same native target/history/save seam. The XLSX descriptor count is eighty-three.

R2-99 adds explicit range-backed list validation with independent bounded source and target ranges,
native history, and browser formula save/reopen. The XLSX descriptor count is eighty-four.

R2-100 consolidates all retained whole/decimal/date/time/text-length comparisons and eight
operators into one bounded operation. It deletes the narrower number/date-between descriptors, so
the XLSX descriptor count falls to eighty-three while the supported matrix expands.

R2-101 adds explicit nullable validation prompt/error metadata over an existing rule through the
native update-options/Undo and declarative save seam. The XLSX descriptor count is eighty-four.

R2-102 starts Conditional Formatting with explicit comparison-rule create/update and exact-ID
removal. Context publishes bounded session rule IDs, targets require complete streaming and CF
metadata, and browser/desktop saves share the declarative CF/DXF mapper. The XLSX descriptor count
is eighty-six.

R2-103 adds one highlight-rule lifecycle operation covering four text predicates, blank/non-blank,
and duplicate/unique. Predicate/text semantics are explicit and bounded while the R2-102 native
history, context identity, and save boundaries are reused. The XLSX descriptor count is eighty-seven.

R2-104 adds one mutually exclusive rank/average lifecycle contract. Top/bottom amount or percent
and above/below inclusive-average states cannot be mixed, and they reuse the established native
history and persistence route. The XLSX descriptor count is eighty-eight.

R2-105 adds bounded custom-formula Conditional Formatting through the shared lifecycle. The XLSX
descriptor count is eighty-nine.

R2-106 adds one visual-rule lifecycle contract for two/three-color scales, base data bars, and the
whitelisted base-OOXML icon sets. Extended-only shapes remain rejected. The XLSX descriptor count
is ninety.

R2-107 completes saveable text/error predicates and tightens the shared UI/save guard for x14-only
date, average, icon, solid-bar, and negative-bar-color states. No descriptor is added.

R2-108 and R2-109 complete Conditional Formatting rule management with explicit range/worksheet
clear and absolute one-based rule priority. Both use the native Conditional Formatting command and
Undo routes; the generated XLSX registry contains ninety-two operations.

R2-110 removes the open-ended Agent `ribbon_command` capability and renderer dispatch branch.
Retained Ribbon UI remains mounted, but Agent mutations now enter XLSX only through bounded
format-owned operations.

R2-111 through R2-131 complete the XLSX post-Ribbon audit: native Pivot add/refresh/update/filter/
chart, explicit formulas, sheet duplicate/visibility/tab color, row/column visibility, row/range
move, table lifecycle, bounded replacement/subtotals/consolidation, and header/footer all use exact
registry operations and proven browser save/reopen seams. The executable retained-command map has
no unexplained state-changing gap. The generated product Manifest contains 114 XLSX operations,
including 112 Agent-visible and two internal staged operations.

R2-132 establishes the retained DOCX semantic-family inventory and its executable closure rule.
R2-133 closes native history first: `docx.history.undo` and `docx.history.redo` share the mounted
TipTap state and history used by the UI and fail closed when no matching entry is available. The
DOCX registry now contains sixteen operations, fifteen Agent-visible and one internal.

R2-134 adds the first exact character-range operation,
`docx.text.set_character_format`. It shares one format-owned ProseMirror transaction helper with
the retained Ribbon for selected text, preserves unlisted run style, and is one native Undo unit.
The DOCX registry now contains seventeen operations, sixteen Agent-visible and one internal; the
rest of the character-format family remains in progress.

R2-135 completes the direct font/size/color/highlight portion by deepening that operation rather
than growing the catalog. The renderer enforces visible font-name, half-point, hex-color, and named
highlight bounds; selected Ribbon actions share the same per-run merge transaction. Character
styles, Clear Formatting, and case transforms remain open.

R2-136 migrates Clear Formatting over the shared exact text-range validator and transaction helper.
The DOCX registry contains eighteen operations, seventeen Agent-visible and one internal;
character styles and case transforms remain open in the exact-character family.

R2-137 migrates all four retained Change Case modes through a shared mapped transaction. The DOCX
registry contains nineteen operations, eighteen Agent-visible and one internal; named character
styles are the final open branch of the exact-character family.

R2-138 adds document-catalog-validated character styles and closes the exact-character family.
The DOCX registry contains twenty operations, nineteen Agent-visible and one internal. Paragraph
and list detail is the next retained semantic family.

R2-139 adds target-aware paragraph direction through the Ribbon's shared logical-alignment rule.
The DOCX registry contains twenty-one operations, twenty Agent-visible and one internal; remaining
list preset/level/restart/continue details stay open.

R2-140 adds absolute list level and shares its `0..8` boundary rule with the Ribbon's relative
indent gesture. The DOCX registry contains twenty-two operations, twenty-one Agent-visible and one
internal; custom presets and restart/continue remain open.

R2-141 adds bounded custom multilevel-list presets through the same document-numbering allocation
service used by the retained dialog. The DOCX registry contains twenty-three operations,
twenty-two Agent-visible and one internal; restart and continue are the remaining list transitions.

R2-142 adds stable-anchor list restart with an explicit bounded start value through the same
numbering-definition clone and forward-rewrite action as the retained context menu. The DOCX
registry contains twenty-four operations, twenty-three Agent-visible and one internal; continue is
the remaining list transition.

R2-143 adds explicit-source list continuation and removes the Agent-side nearest-list inference.
The DOCX registry contains twenty-five operations, twenty-four Agent-visible and one internal; the
list-numbering family is closed and bounded paragraph details are next.

R2-144 deepens the existing paragraph-style operation with bounded ParagraphDialog line modes,
spacing, and indentation. The registry count remains twenty-five; Ruler tab stops are the remaining
paragraph-specific mutation.

R2-145 adds bounded final-state tab stops to that same operation and shares serialization with the
Ruler. The registry count remains twenty-five; the retained paragraph-format family is closed and
table lifecycle/formatting is next.

R2-146 starts table lifecycle with stable-boundary bounded insertion over the retained native table
model. The DOCX registry contains twenty-six operations, twenty-five Agent-visible and one internal;
table deletion and row/column/cell mutations remain open.

R2-147 adds exact top-level table deletion with native Undo and the valid-empty-document fallback.
The DOCX registry contains twenty-seven operations, twenty-six Agent-visible and one internal;
row/column/cell lifecycle remains open.

R2-148 adds explicit-boundary, rowspan-aware multi-row insertion in one transaction. The DOCX
registry contains twenty-eight operations, twenty-seven Agent-visible and one internal; row deletion
and column/cell lifecycle remain open.

R2-149 adds exact-index, rowspan-aware multi-row deletion in one transaction. The DOCX registry
contains twenty-nine operations, twenty-eight Agent-visible and one internal; full-table deletion
remains a distinct lifecycle operation, and column/cell lifecycle remains open.

R2-150 adds explicit-boundary, colspan-aware multi-column insertion in one transaction. The DOCX
registry contains thirty operations, twenty-nine Agent-visible and one internal; column deletion
and cell lifecycle remain open.

R2-151 adds exact-index, colspan-aware multi-column deletion in one transaction. The DOCX registry
contains thirty-one operations, thirty Agent-visible and one internal; row/column lifecycle is
closed and cell lifecycle remains open.

R2-152 adds exact half-open rectangle cell merge through the retained native command and fixes the
shared row schema for rows wholly covered by rowspan. The DOCX registry contains thirty-two
operations, thirty-one Agent-visible and one internal; split-cell and table formatting remain open.

R2-153 adds exact logical-coordinate split through the retained native command. The DOCX registry
contains thirty-three operations, thirty-two Agent-visible and one internal; merge/split lifecycle
is closed and table formatting remains open.

R2-154 starts exact table formatting with masked cell fill and vertical alignment over an explicit
rectangle; retained Ribbon selections share the same write kernel. The DOCX registry contains
thirty-four operations, thirty-three Agent-visible and one internal; borders, table styles, and
sizing remain open.

R2-155 adds exact rectangle border policies and replaces the Ribbon's private edge algorithm with
the shared registry kernel. The DOCX registry contains thirty-five operations, thirty-four
Agent-visible and one internal; table styles and sizing remain open.

R2-156 adds current-document table-style final state and routes Ribbon style cards through the
shared native attribute kernel. The DOCX registry contains thirty-six operations, thirty-five
Agent-visible and one internal; row and column sizing remain open.

R2-157 adds bounded physical-row height final state and shares its write kernel with the Ribbon.
The DOCX registry contains thirty-seven operations, thirty-six Agent-visible and one internal;
column sizing remains open.

R2-158 adds a bounded complete column-width vector through the Ribbon's existing grid write kernel.
The DOCX registry contains thirty-eight operations, thirty-seven Agent-visible and one internal;
the retained table family is closed.

R2-159 starts the page/section family with a stable-boundary native page-break paragraph. The DOCX
registry contains thirty-nine operations, thirty-eight Agent-visible and one internal; explicit
section breaks and page/section final-state settings remain open.

R2-160 adds finite, stable-boundary section-break insertion with Undo-owned start type and
save/reopen projection to the following section terminator. The DOCX registry contains forty
operations, thirty-nine Agent-visible and one internal; page/section settings remain open.

R2-161 adds exact indexed orientation and establishes the Undo-owned full section-settings
override journal shared by Ribbon and Agent. The DOCX registry contains forty-one operations,
forty Agent-visible and one internal; margins, paper, columns, and later section fields remain open.

R2-162 adds exact four-sided margins with bounded twips, positive-body validation, native Undo,
and save/reopen evidence through the same journal. The DOCX registry contains forty-two operations,
forty-one Agent-visible and one internal; paper, columns, and later section fields remain open.

R2-163 adds exact bounded page axes with margin-fit validation and derived orientation through the
same journal. The DOCX registry contains forty-three operations, forty-two Agent-visible and one
internal; columns and later section fields remain open.

R2-164 adds exact bounded column count and spacing, fixes the retained engine's hard-coded spacing
fallback, and proves save/reopen. The DOCX registry contains forty-four operations, forty-three
Agent-visible and one internal; later retained section fields remain open.

R2-165 adds explicit page-border state through the shared journal. The DOCX registry contains
forty-five operations, forty-four Agent-visible and one internal; every retained page/section UI
mutation is registry-owned and the family is closed.

R2-166 starts the header/footer family with indexed different-first-page state, native Undo, and
non-final/final `w:titlePg` save projection. The DOCX registry contains forty-six operations,
forty-five Agent-visible and one internal.

R2-167 adds explicit document-wide different-odd/even-page state through the same native journal
and `settings.xml/w:evenAndOddHeaders` save projection. The DOCX registry contains forty-seven
operations, forty-six Agent-visible and one internal.

R2-168 adds indexed page-number format and nullable start state through the same native journal,
removing the former final/non-final React dirty-state split. The DOCX registry contains forty-eight
operations, forty-seven Agent-visible and one internal.

R2-169 adds explicit section/kind/variant header-footer text through an Undo-owned content journal,
removes the former React dirty save path, and extends non-final engine writes to first/even variants.
The DOCX registry contains forty-nine operations, forty-eight Agent-visible and one internal.

R2-170 adds explicit header/footer PAGE-field enablement and alignment, replacing gallery strings
with canonical sentinel-backed final state. The DOCX registry contains fifty operations,
forty-nine Agent-visible and one internal.

R2-171 adds bounded rich header/footer paragraphs with styled text plus PAGE/NUMPAGES segments.
HeaderFooterArea and Agent values now share the same full-value journal, closing the retained
header/footer family at fifty-one operations (fifty Agent-visible and one internal).

R2-172 starts the image lifecycle with public stable-boundary `docx.image.insert` and internal
`docx.image.insert_staged`. Broker staging, host hydration, renderer magic validation, native Undo,
shared Ribbon/paste image-node construction, and real save/reopen evidence are complete. The DOCX
registry contains fifty-three operations (fifty-one Agent-visible and two internal).

R2-173 adds exact-index staged image replacement with explicit final dimensions. Original and
unsaved images share one replacement-attrs kernel with the retained UI, preserve placement/wrap,
clear stale crop windows, undo natively, and reopen with the replacement bytes. The DOCX registry
contains fifty-five operations (fifty-two Agent-visible and three internal).

R2-174 adds exact-index final wrap/inline state. All retained image wrap producers share one helper;
inline clears stale positioning, floating modes preserve it, and save/reopen plus native Undo are
covered. The DOCX registry contains fifty-six operations (fifty-three Agent-visible and three
internal).

R2-175 adds the original-image-only 3×3 margin position gallery as an exact indexed operation.
Named axes, derived square wrap, cleared free offsets, UI sharing, Undo, and save/reopen are covered.
The DOCX registry contains fifty-seven operations (fifty-four Agent-visible and three internal).

R2-176 adds explicit signed-EMU free positioning for original and generated images. The registry
and retained image-drag path share one final-state projection for wrap, offsets, and cleared named
axes; textbox/shape dragging remains object-owned. Native Undo and original-image save/reopen are
covered. The DOCX registry contains fifty-eight operations (fifty-five Agent-visible and three
internal).

R2-177 replaces image rotation deltas and flip toggles with one exact indexed final-state
operation. Registry and retained Rotate/Flip controls share the same canonical projection; native
Undo, generated-image state, and original-image save/reopen are covered. The saver also creates a
minimal missing picture transform container instead of silently losing the edit. The DOCX registry
contains fifty-nine operations (fifty-six Agent-visible and three internal).

R2-178 replaces destructive Crop-dialog byte baking with explicit non-destructive source insets.
The bounded operation supports all-zero reset, rejects empty retained regions, shares its attrs
writer with the retained UI, and persists `a:srcRect` for original and new images. Native Undo and
save/reopen are covered. The DOCX registry contains sixty operations (fifty-seven Agent-visible and
three internal); image removal is the remaining lifecycle mutation.

R2-179 adds exact indexed image removal with the document non-empty invariant in the same native
transaction. Native Undo and save/reopen deletion are covered. The DOCX registry contains
sixty-one operations (fifty-eight Agent-visible and three internal), and the retained image
lifecycle is closed.

R2-180 starts the object family with stable-boundary insertion for all 104 retained filled shape
presets. Registry, Gallery, and draw mode share one node builder; both EMU axes are bounded, and
native Undo plus save/reopen are covered. The DOCX registry contains sixty-two operations
(fifty-nine Agent-visible and three internal); five line/connector presets remain separate.

R2-181 adds stable-boundary insertion for all five line/connector presets. Straight kinds require
the canonical 114,300 EMU grab height; bent/curved kinds preserve both bounded axes. Registry,
Gallery, and draw mode share one stroke-only builder, with native Undo and save/reopen evidence.
The DOCX registry contains sixty-three operations (sixty Agent-visible and three internal).

R2-182 adds stable-boundary textbox insertion with two bounded EMU axes. Registry and Ribbon now
share one format-owned textbox node builder, native Undo, and save/reopen projection. The DOCX
registry contains sixty-four operations (sixty-one Agent-visible and three internal).

R2-183 adds stable-boundary chart insertion for bounded bar, line, and pie data matrices with an
explicit final extent. Registry and the retained Chart dialog share one format-owned node builder,
native Undo, and embedded workbook save/reopen projection. The DOCX registry contains sixty-five
operations (sixty-two Agent-visible and three internal).

R2-184 adds exact block/inline equation insertion with 1–4,096-character LaTeX, mutually exclusive
stable block/range coordinates, and Unicode-aware string bounds in Operation Contract. Gallery,
modal, and Registry share one LaTeX-to-OMML action, native Undo, and save/reopen projection. The
DOCX registry contains sixty-six operations (sixty-three Agent-visible and three internal).

R2-185 adds one aggregate cross-drawing size route for shapes, lines, textboxes/WordArt-like boxes,
and charts. Registry and the retained corner handle share a bounded position-based kernel, with
chart/straight-line invariants, an independent native Undo group, and save/reopen projection. The
DOCX registry contains sixty-seven operations (sixty-four Agent-visible and three internal).

R2-186 adds one aggregate offset-position route for textbox-backed drawings with explicit wrap and
signed 32-bit EMU offsets. Shape draw, move handles, and Registry share the mutation kernel; OOXML
numeric offsets now retain left/right wrap semantics and polygon bytes. The DOCX registry contains
sixty-eight operations (sixty-five Agent-visible and three internal).

R2-187 adds one aggregate masked object-style route for nullable fill and outline colors. The
retained Shape Format palette and Registry share the same native transaction; exact block identity,
uppercase six-digit colors, line invariants, native Undo, and save/reopen are covered. The DOCX
registry contains sixty-nine operations (sixty-six Agent-visible and three internal).

R2-188 adds one aggregate exact-removal route for drawing, chart, diagram, and block-equation
objects. Real object-mode Backspace/Delete and Registry share the kernel; image-family separation,
last-block replacement, native Undo, and save/reopen are covered. The DOCX registry contains
seventy operations (sixty-seven Agent-visible and three internal).

R2-189 adds a masked same-shape chart-content update route for existing title, category,
series-name, and numeric cache slots. The retained protected grid and Registry share one bounded
transaction; read-only gaps, native Undo, and original/generated save/reopen are covered. Combined
with object size/removal, retained chart lifecycle is closed. The DOCX registry contains seventy-one
operations (sixty-eight Agent-visible and three internal).

R2-190 adds one dual-mode exact equation-update route. Bounded LaTeX rebuilds target either a stable
block identity or one exact inline atom range, while same-shape token arrays preserve retained block
OMML structure. EquationModal, protected token editing, and Registry share one native-history
kernel; invalid coordinate/mode combinations, single-step Undo, and save/reopen are covered. The
DOCX registry contains seventy-two operations (sixty-nine Agent-visible and three internal).

R2-191 adds one exact nested rich-content route shared by plain textboxes, textbox-backed shapes,
WordArt, and multi-box drawings. Bounded paragraphs/runs, common rich/paragraph formatting, an
explicit nullable final height, flattened-content rejection, and aggregate text/run budgets flow
through one batch-capable native-history writer used by Registry and retained sub-editors. UI Undo
and original OOXML save/reopen are covered. The DOCX registry contains seventy-three operations
(seventy Agent-visible and three internal), closing the retained shape/textbox/equation content row.

R2-192 adds one nullable exact-range link operation for insertion/replacement, href changes, and
removal. Link modal, keyboard/context-menu entry, and Registry share one native-history action;
bounded validation, single-step Undo, external relationship allocation, and save/reopen are
covered. The superseded whole-block link field is removed from `docx.text.set_style`, so link state
has one canonical writer. The DOCX registry contains seventy-four operations (seventy-one
Agent-visible and three internal).

R2-193 adds one bounded final-state bookmark operation. A stable top-level paragraph/heading/list
item identity, unique 1–40-character name, and explicit enabled flag replace the Bookmark dialog's
direct node-attribute mutation. UI and Registry share one native-history action; idempotence,
duplicate rejection, Undo, and OOXML save/reopen are covered. The DOCX registry contains
seventy-five operations (seventy-two Agent-visible and three internal).

R2-194 adds exact-range cross-reference insertion with an existing bounded bookmark identity and
explicit cached display text. CrossRefModal and Registry share one native-history field action;
missing targets fail before mutation, and Undo plus REF-field save/reopen are covered. The DOCX
registry contains seventy-six operations (seventy-three Agent-visible and three internal).

R2-195 adds exact-range generic-field insertion for the retained DATE, TIME, PAGE, NUMPAGES, and
FILENAME set. Explicit cached display text keeps execution deterministic; App/Ribbon and Registry
share one native-history field action, with Undo and OOXML save/reopen evidence. The DOCX registry
contains seventy-seven operations (seventy-four Agent-visible and three internal).

R2-196 adds bounded aggregate generic-field cache updates with exact ranges, original instructions,
and explicit final display text. F9/context menu and Registry share one native-history transaction;
exact-mark rejection, multi-field Undo, and OOXML save/reopen are covered. The DOCX registry contains
seventy-eight operations (seventy-five Agent-visible and three internal).

R2-197 closes symbol insertion by deepening existing `docx.text.insert`, not adding another tool.
The shared 1–65,536-Unicode-character action serves Registry and the symbol palette through one
native-history transaction; empty input, Undo, and save/reopen are covered. Manifest/DOCX counts
remain 202/78, and the bookmark/cross-reference/generic-field/symbol family is closed.

R2-198 adds stable-region `docx.toc.refresh` with bounded explicit final entries. Original/generated
TOC field boundaries, trailing page-break preservation, retained UI sharing, single-step Undo, and
save/reopen are covered. The Manifest contains 203 operations and DOCX owns seventy-nine
(seventy-six Agent-visible, three internal), closing the TOC family.

R2-199 starts the notes family with explicit `docx.note.insert`. Stable note IDs and bounded bodies
are stored with the exact-range reference atom, allowing native Undo/Redo to reconcile note-part
state through the shared retained UI/Registry seam. Duplicate IDs, Broker validation, and
save/reopen are covered. The Manifest contains 204 operations and DOCX owns eighty (seventy-seven
Agent-visible, three internal).

R2-200 adds stable-ID `docx.note.update`. Bounded final text is stored on every matching reference
atom while the shared session keeps the original note body as the history baseline; native
Undo/Redo therefore restores both marker-owned metadata and note-part text. Retained UI, Registry,
Broker validation, and save/reopen are covered. The Manifest contains 205 operations and DOCX owns
eighty-one (seventy-eight Agent-visible, three internal).

R2-201 adds stable-ID `docx.note.delete`. All matching references are deleted and peers renumbered
in one native-history transaction; the shared baseline distinguishes deletion from update so Undo
restores both body and numbering. Retained UI, Registry, Broker validation, and save/reopen are
covered. The Manifest contains 206 operations and DOCX owns eighty-two (seventy-nine Agent-visible,
three internal), closing footnote/endnote lifecycle.

R2-202 starts bibliography-source migration with `docx.source.upsert`. Stable-tag add/update writes
one bounded final Sources snapshot to an Undo-owned document-block override; retained UI, Registry,
Broker validation, native Undo/Redo, and customXml save/reopen share the route. The Manifest contains
207 operations and DOCX owns eighty-three (eighty Agent-visible, three internal).

R2-203 adds exact-range `docx.citation.insert` with stable source-tag validation and explicit
bounded display text. Retained citation UI, Registry, native Undo, Broker validation, and
save/reopen share the renderer's existing plain-text citation semantics. The Manifest contains 208
operations and DOCX owns eighty-four (eighty-one Agent-visible, three internal).

R2-204 adds stable-boundary `docx.bibliography.insert` with an explicit bounded heading and
source-backed final entry lines. Retained Bibliography UI, Registry, native one-step Undo, Broker
validation, and save/reopen share one renderer action. The Manifest contains 209 operations and
DOCX owns eighty-five (eighty-two Agent-visible, three internal).

R2-205 adds stable-boundary `docx.caption.insert` with explicit bounded label, positive final
number, and optional display text. Retained Caption UI, Registry, native one-step Undo, Broker
validation, and dirty SEQ-field save/reopen share one action. The Manifest contains 210 operations
and DOCX owns eighty-six (eighty-three Agent-visible, three internal).

R2-206 adds exact-range `docx.index.mark` with one bounded round-trippable term. Retained Mark Entry
UI freezes the original range while its prompt is open, and shares native Undo, Broker validation,
and XE-field save/reopen with Registry. The Manifest contains 211 operations and DOCX owns
eighty-seven (eighty-four Agent-visible, three internal).

R2-207 adds stable-boundary `docx.index.insert` with a bounded explicit term snapshot. Retained
Index UI, Registry, one-step native Undo, Broker validation, and dirty INDEX-field save/reopen share
one action. Audit confirms retained source UI has add/upsert but no source-delete mutation, closing
the References and Index family. The Manifest contains 212 operations and DOCX owns eighty-eight
(eighty-five Agent-visible, three internal).

R2-208 adds exact-range `docx.comment.add` with explicit stable metadata. Anchor marks and the full
comment snapshot now share one native-history transaction, reconciling retained UI state and
`comments.xml` through Undo/Redo and save/reopen. The Manifest contains 213 operations and DOCX
owns eighty-nine (eighty-six Agent-visible, three internal).

R2-209 adds stable-parent `docx.comment.reply` with explicit bounded reply metadata. Parent anchor
augmentation and the final comments snapshot share one native-history transaction, including
Broker validation and `commentsExtended.xml` save/reopen. The Manifest contains 214 operations and
DOCX owns ninety (eighty-seven Agent-visible, three internal).

R2-210 adds explicit-final `docx.comment.set_resolved`. Parent/reply state shares one Undo-owned
snapshot with retained Resolve/Reopen UI, Broker validation, and extended-comment save/reopen. The
Manifest contains 215 operations and DOCX owns ninety-one (eighty-eight Agent-visible, three internal).

R2-211 adds stable-ID `docx.comment.delete`. Parent deletion cascades direct replies, reply deletion
remains local, and mark cleanup plus the final comments snapshot share one native-history
transaction. Retained Delete UI, Registry, Broker validation, Undo, and save/reopen share the route.
The Manifest contains 216 operations and DOCX owns ninety-two (eighty-nine Agent-visible, three
internal), closing the retained comment lifecycle.

R2-212 adds explicit-final `docx.revision.set_tracking`. Retained Ribbon and Registry share the
mounted recorder-state writer; repeated state is a no-op. The session policy is explicitly
non-undoable and recovery-free, while the next tracked edit remains native-history-owned and
persists as `w:ins`. The Manifest contains 217 operations and DOCX owns ninety-three (ninety
Agent-visible, three internal).

R2-213 adds bounded `docx.revision.apply_decision`, using two finite enums to cover accept/reject
for the current selection-resolved revision or all revisions without four separate tools. Retained
Review UI and Registry share one native-history transaction with Broker rejection and final
save/reopen evidence. The Manifest contains 218 operations and DOCX owns ninety-four (ninety-one
Agent-visible, three internal), closing the revisions family.

R2-214 adds explicit-final `docx.document.set_protection`. Passwords are bounded input-only values;
password-protected removal verifies the current credential and never returns it. Retained Review
UI and Registry share the same resolver and dirty save journal. The non-ProseMirror state is
explicitly non-undoable, with recovery and `settings.xml` enable/remove save/reopen evidence. The
Manifest contains 219 operations and DOCX owns ninety-five (ninety-two Agent-visible, three internal).

R2-215 adds aggregate `docx.ink.apply` for bounded vector add, stable-ID delete, and clear. It
cross-validates the finite action shape and never accepts arbitrary image/base64 injection.
Retained overlay/Ribbon and Registry share the renderer-owned non-undoable ink journal with dirty
recovery and ink-part save/reopen evidence. The Manifest contains 220 operations and DOCX owns
ninety-six (ninety-three Agent-visible, three internal), closing Ink.

R2-216 adds public `docx.document.compare { path }` plus internal staged-bytes
`docx.document.compare_staged`. The Broker owns local-path validation and opaque staging; the
mounted renderer owns DOCX parsing, deterministic paragraph diff, and the retained Compare panel
state. The recovery-free operation advances acknowledged session state without dirtying or
rewriting the document. The Manifest contains 222 operations and DOCX owns ninety-eight
(ninety-four Agent-visible, four internal), closing Compare.

R2-217 adds masked aggregate `docx.document.set_design` for bounded page color, watermark, theme
fonts, and theme colors. Its unique `fields` mask exactly identifies supplied final-state values,
so combined theme selection is atomic without exposing UI preset strings. Retained Design UI and
Registry share the resolver and dirty recovery-backed save journal; combined save/reopen evidence
covers all four fields. The Manifest contains 223 operations and DOCX owns ninety-nine
(ninety-five Agent-visible, four internal), closing Document design.

R2-218 adds deterministic `docx.cover_page.insert`. Twelve finite style presets are separated from
explicit bounded title, subtitle, author, date, and year content, eliminating locale/time drift on
replay. Retained Ribbon defaults call the same action; start-of-document insertion, one-step native
Undo, and custom-content save/reopen are covered. The Manifest contains 224 operations and DOCX
owns 100 (ninety-six Agent-visible, four internal), closing Cover pages.

R2-219 adds explicit-final `docx.paragraph.set_drop_cap { blockIndex, mode, lines }`. `none` pairs
with null lines; visible `drop`/`margin` modes require 2–10. Retained Ribbon resolves the current
top-level paragraph identity and shares one native transaction with Registry. Broker bounds,
deterministic no-op, one-step Undo, and save/reopen are covered. The Manifest contains 225
operations and DOCX owns 101 (ninety-seven Agent-visible, four internal), closing Drop caps.

R2-220 adds deterministic `docx.wordart.insert`. Explicit top-level boundary, finite preset, bounded
text/geometry, and unique drawing identity replace the retained Ribbon's localized/default/random
insertion details. Ribbon allocates the lowest free ID and shares the action; Broker preset
rejection, duplicate-ID rejection, one-step Undo, and text/color/geometry save/reopen are covered.
The Manifest contains 226 operations and DOCX owns 102 (ninety-eight Agent-visible, four internal).
The R2-132 DOCX retained-command inventory closed here; at this historical slice, shared release
gates remained pending. R6-01 later closed them.

R2-221 starts Markdown retained-command parity with `markdown.history.undo {}` and
`markdown.history.redo {}`. Retained Ribbon and Registry share one native TipTap history action;
empty-stack execution fails explicitly instead of acknowledging a false mutation. Generated
discovery and real MCP session queue/acknowledgement cover both routes. The Manifest contains 228
operations and Markdown owns six (five Agent-visible, one internal). Block-type parity is next.

R2-222 adds `markdown.block.set_type { textBlockIndex, type }` for explicit paragraph, H1–H6,
quote, and code-block final states. The revision-scoped address walks all editable text blocks in
document order, including nested retained structures; `office_get_context` exposes the current
address. Ribbon, Slash Menu, and Registry share the action. Native Undo, heading save/reopen,
quote normalization, runtime bounds, Broker enum rejection, and live-session acknowledgement are
covered. The Manifest contains 229 operations and Markdown owns seven (six Agent-visible, one
internal). Inline-mark parity is next.

R2-223 adds aggregate `markdown.text.set_marks { from, to, marks }`. The full required mask carries
bold, italic, strike, inline-code, and nullable link final state over one explicit ProseMirror
range; inline code conflicts fail explicitly. All five retained Ribbon controls call the same
action instead of direct toggles. Set/clear, native Undo, Markdown save/reopen, range/conflict
failure, generated discovery, nested link bounds, and real MCP acknowledgement are covered. The
Manifest contains 230 operations and Markdown owns eight (seven Agent-visible, one internal).
List parity is next.

R2-224 adds `markdown.list.set_type { textBlockIndex, type }` with explicit none/bullet/ordered/task
final states. It reuses the revision-scoped text-block address, preserves sibling items when one
item exits a list, and uses one native TipTap chain for conversion. Ribbon, Slash Menu, and
Registry share the action. Task-list save/reopen, Undo, runtime bounds, Broker enum rejection,
generated discovery, and live-session acknowledgement are covered. The Manifest contains 231
operations and Markdown owns nine (eight Agent-visible, one internal). Deterministic structure
insertion is next.

R2-225 adds `markdown.table.insert { position, rows, columns, headerRow }` and
`markdown.divider.insert { position }`. Revision-scoped ProseMirror positions preserve native
Ribbon and Slash insertion semantics; table dimensions are bounded to 100×100. Shared actions,
table save/reopen, one-step Undo, runtime position failure, Broker row bounds, generated discovery,
and two real-session acknowledgements are covered. The Manifest contains 233 operations and
Markdown owns eleven (ten Agent-visible, one internal). Staged image insertion is next.

R2-226 adds public `markdown.image.insert { path, position, alt, title }` and internal
`markdown.image.insert_staged { blobId, name, size, data, position, alt, title }`. Broker staging
and Renderer magic checks bound PNG/JPEG/GIF input to 20 MB; paths and bytes never cross the wrong
boundary. Hydrated bytes become a browser-safe data URL through the same action used by picker,
paste/drop, and Slash insertion. Undo, save/reopen, forged-byte rejection, internal hiding, and
real no-inline-payload staging are covered. The Manifest contains 235 operations and Markdown owns
thirteen (eleven Agent-visible, two internal). Frontmatter is next.

R2-227 adds `markdown.frontmatter.set { yaml }` as a complete raw-envelope final state. The shared
action preserves body, BOM, EOL, and EOF-newline state while empty YAML removes the envelope. It is
explicitly non-TipTap-undoable, matching the retained textarea/persistence boundary, but dirty,
recovery, save/reopen, generated discovery, 1 MiB Broker bounds, and live acknowledgement are
covered. The Manifest contains 236 operations and Markdown owns fourteen (twelve Agent-visible,
two internal). Table-relative commands are next.

R2-228 adds aggregate `markdown.table.update { position, action, headerRow }`. Eight finite actions
cover row/column insertion/deletion, table deletion, and explicit header-row final state without
eight public descriptors or a replay-unsafe toggle. TableMenu and Registry share the action.
Cell-relative execution, header no-op, native Undo, table reopen, non-table and field-relation
failure, Broker enum rejection, discovery, and live acknowledgement are covered. The Manifest
contains 237 operations and Markdown owns fifteen (thirteen Agent-visible, two internal). Block
operations are next.

R2-229 adds aggregate `markdown.block.update { blockIndex, action, afterBlockIndex, content }` for
duplicate, delete, add-below, and move. Move boundaries use original-revision top-level indexes;
action-specific nullable fields are runtime-checked. BlockKeymap, menu, plus control, native drag
drop, and Registry share one transaction builder; deleting the final block atomically restores an
empty paragraph. Undo, ordering, field-relation failure, Broker enum rejection, discovery, and
live acknowledgement are covered. The Manifest contains 238 operations and Markdown owns sixteen
(fourteen Agent-visible, two internal). Code-block language is next.

R2-230 adds `markdown.code_block.set_language { textBlockIndex, language }`. The 30 retained UI
choices are a finite contract; plaintext maps to a null language attribute. CodeBlock NodeView and
Registry share one native `setNodeMarkup` transaction. Non-code target failure, no-op semantics,
Undo, Markdown reopen, Broker enum rejection, discovery, and live acknowledgement are covered.
The Manifest contains 239 operations and Markdown owns seventeen (fifteen Agent-visible, two
internal). A retained-UI/direct-edit/export audit is next.

R2-231 adds `markdown.document.save_as {}` for the retained Shift-Save command. UI and Registry
share `doSave(true)`, which deliberately detaches from the current destination, returns the final
file name, and maps picker cancellation or persistence failure to an explicit negative
acknowledgement. Generated discovery and live-session queue/output evidence are covered. The
Manifest contains 240 operations and Markdown owns eighteen (sixteen Agent-visible, two internal).
DOCX export is next.

R2-232 through R2-234 close retained output and preference commands. Public
`markdown.document.export_docx {}`, `markdown.document.open_print_dialog {}`, and
`markdown.document.set_auto_save { enabled }` share the exact UI services. DOCX export now embeds
bounded valid PNG/JPEG/GIF bytes, print reports host blocking rather than claiming a PDF, and
autosave is an explicit persisted final state. The Manifest reaches 243 operations and Markdown
owns twenty-one (nineteen Agent-visible, two internal).

R2-235 adds `markdown.selection.set { from, to }`. Agent workflows can now reproduce arbitrary
native text selection before insert or replacement/deletion; selection-only transactions add no
Undo entry or recovery checkpoint. The Manifest contains 244 operations and Markdown owns
twenty-two (twenty Agent-visible, two internal).

R2-236 and R2-237 replace the removed `md-asset` protocol. MCP local open uses one session-bound
asset root, a 256 KiB app-only chunk transport, exact image validation, and display-only data-URL
hydration. Standalone browser Open authorizes a directory tree and resolves relative image files.
Both paths preserve authored Markdown paths and feed the same DOCX exporter.

R2-238 adds the machine-checked Markdown retained-command mapping. Every one of the 22 Registry
descriptors belongs to a typed UI/ingress/native-input family and no `missing` disposition remains.
Markdown retained-command parity is complete; shared packaged-host, resource, smoke, and root
release gates remain before `ready` changes.

R6-01 closes those shared gates. It adds an approved release-evidence contract, deterministic
small/medium/large fixtures for all formats, pinned-source visual provenance, 7/7/21 canonical
runtime samples, ACK decomposition, peak JS heap and renderer RSS, source/upstream fingerprints,
artifact integrity checks, and a generated all-format readiness projection. All five formats now
report `ready: true`; any release-relevant source drift fails closed until evidence is recaptured.

R6-02 replaces the five-format 500 ms command cadence with one shared wakeable bounded-poll
scheduler. The immediate bootstrap poll and size-before-poll/fullscreen ordering remain intact;
steady state holds one ten-second app-only request that Broker enqueue resolves immediately. Empty
timeouts re-arm, disconnect releases the waiter, newer polls supersede lost waiters, and 500 ms is
retained only after transport failure. No public tool, UI resource, or authority boundary changes.
The XLSX tracer also replaces a formerly timing-dependent open/edit race with an explicit
file-selection-to-Univer-ready Promise and a one-frame acknowledgement settlement boundary.

R6-03 traces and bounds Markdown staged loading. The internal load acknowledgement now reports
decode, parse, TipTap state-install, and React-commit phases and is positive only after committed
content and status are observable. Constant-time candidate guards prevent ordered-list, task-list,
and table extension probes from repeatedly splitting the remaining source while preserving the
retained tokenizer for actual syntax. Release-evidence schema v2 records every phase for all three
fixture sizes and enforces fixed Markdown open p95 ceilings of 20 ms, 835.68 ms, and 5,000 ms.

R6-04 traces and bounds XLSX cold start without weakening the complete retained renderer. The
first app-only poll waits for an active workbook, active worksheet, and Univer canvas, then carries
bootstrap, Univer-create, worksheet-install, and first-commit durations once, retrying the same
trace only after transport failure. Release-evidence schema v3 records seven-sample phase
summaries and enforces a fixed 1,400 ms p95 ceiling. The capture's RSS probe is asynchronous and
serialized to remove its former event-loop observer penalty, while the XLSX bundle targets the
host-supported ES2022 baseline. Every retained preset, command, native history path, and save/reopen
route remains present.

R6-05 decomposes and reduces the dominant XLSX bootstrap phase. The startup trace adds strict
resource-receive, initial-module-graph, and React-mount durations whose sum equals the retained
`bootstrapMs` aggregate. XLSX packaging preserves Vite's lazy boundaries while gzip/base64
embedding every emitted module in the existing HTML resource; an in-memory module vault inflates
the entry graph first and optional locale/hyphenation chunks on demand without network access.
Schema-v4 evidence keeps the 1,400 ms total XLSX cold-start ceiling, adds a fixed 500 ms bootstrap
p95 ceiling, and package tests cap initial executable JavaScript at 11 MB. All retained modules,
Registry operations, native history, persistence, and the first-commit-before-poll boundary remain.

R6-06 replaces the complete per-format Registry response with bounded Manifest discovery. Default
calls return at most twenty schema-free summaries with family filtering and stable cursor
pagination; detail calls return exactly one canonical Agent-visible descriptor. Optional
session-aware availability uses stable reasons without hiding operations. Internal operations and
aliases stay outside discovery, summary pages are gated below 8 KiB, and generated descriptors are
capped at 64 KiB. The Agent-visible MCP surface remains ten tools and five UI resources; thirteen
app-only endpoints own transport details.

R6-07 replaces new mutation calls with a revision-guarded transaction envelope carrying caller
`requestId`, exact `baseRevision`, and one canonical Agent-visible operation. The live Broker owns a
session-scoped journal: exact completed and in-flight replays converge on the original transaction,
payload-changing reuse fails, and canonical-id/schema/context/atomicity errors stop before renderer
dispatch. One positive renderer acknowledgement advances revision once. A caller timeout does not
settle the transaction or release its staged bytes, so an exact later replay returns the eventual
result without dispatching twice. The old `operation` + `arguments` form remains only as a temporary
compatibility Adapter through R6-07. R6-08 removes that Adapter, its public alias-resolution and
legacy response branches, and migrates all repository callers/tests to the required transaction
envelope. Packaged smoke now guards the absence of both retired input fields. R6-09 removes the
remaining thirty-six public-era format aliases from the five registries and generated Manifest;
only the five internal `open_local_file` staged-file transport aliases remain.

R6-10 separates stable document identity from mounted-view identity across all five formats. Normal
creation defaults to a blank isolated Session; exact follow-up and Broker-restart recovery retain the
known `sessionId`, while newest-snapshot recovery is explicit only. Each show result grants a fresh
`viewId`, the Broker enforces one active lease per Session, and one mounted iframe cannot rebind to a
different Session/view pair. Connected follow-up edits therefore reuse the existing editor without a
second show call.

PPTX-P1 and PPTX-P2 establish the PPTX document-operation tracer: public
`pptx.document.save` and internal `pptx.document.load_staged` are format-owned, generated,
Broker-validated, and dispatched through injected services into the mounted
`BrowserPresentation` persistence/load pipeline. The historical `save` alias is retired in R6-09;
`open_local_file` remains internal transport-only, and the replaced browser-host operation branches are absent. This does not
complete PPTX selection, mutation, or retained-command parity.

R2-239 through R2-253 migrate every PPTX mutation that the browser host currently executes. The
19-operation PPTX catalog now contains 18 public operations plus internal staged load: explicit
selection, selected text/movement compatibility, blank/duplicate/delete/move slide lifecycle,
native undo/redo, save/save-as/new blank, explicit delete and EMU transform, font and paragraph
formatting, scoped find/replace, and retained rich-text paragraphs. The browser host has no direct
operation-id branches and the MCP server has no handwritten PPTX schemas. Real-fixture tests prove
dirty state, renderer refresh, recovery, and native undo/redo. This closes registry parity for the
currently executable browser slice; it does not make unsupported community `SlidesApi` producers
complete or set PPTX `ready`.

R2-254 through R2-308 complete the retained PPTX state-changing command inventory. Seventy-four
format-owned operations (73 Agent-visible and one internal) now cover slides/layout, explicit
object and clipboard actions, parent-addressed group children, connector/picture/table/chart/
SmartArt/theme/animation, notes/comments/sections/hyperlinks, embedded image/media/3D/ink, and
master/layout-part edits. The browser host implements the complete retained `SlidesApi` at compile
time; picker inputs converge on bounded byte operations, Presenter/Audience share same-origin
browser state, and export/print/show are classified as non-mutating host effects. The machine-
checked PPTX retained-command baseline maps every descriptor with no `missing` disposition.
Format-local retained-command parity is closed. Those shared gates kept `ready`
false at this historical milestone; current authorization is recorded in
[`../../release/validation.md`](../../release/validation.md).

The following locations are migration scaffolds rather than extension points:

- the central hand-written format catalogs in `apps/mcp-server/src/capabilities.ts`;
- arbitrary `operation: string` acceptance before broker validation;
- renderer composition-root operation condition chains;
- the now-removed XLSX open-ended `ribbon_command` route (retained here only as migration history);
- manually duplicated Skill and protocol operation tables.

Existing behavior moves first and must remain behaviorally equivalent. Each replacement slice removes or demotes the old route it supersedes. After the registry foundation and compatibility migration pass, format parity resumes as vertical operation-family slices with descriptor, context, native command seam, undo, acknowledgement, save/reopen, generated documentation, and baseline mapping evidence in the same change.

## Current facts

| Format   | Pinned community renderer | Current implementation                                                                 | Decision status                                                   |
| -------- | ------------------------: | -------------------------------------------------------------------------------------- | ----------------------------------------------------------------- |
| DOCX     |                 130 files | Community renderer selectively retained                                                | Retained-command parity and R6-01 release gate complete; ready    |
| XLSX     |                 111 files | 76 permitted pinned files plus TandemFolio host/operation adapters; pinned App mounted | Renderer/mutation parity and R6-01 release gate complete; ready   |
| PPTX     |                 104 files | 61 identical + 19 adapted permitted files; original App mounted                        | Retained-command parity and R6-01 release gate complete; ready    |
| PDF      |                  40 files | All 33 permitted renderer files boot from `src/renderer`                               | Retained producer baseline and R6-01 release gate complete; ready |
| Markdown |         30 renderer files | 21 pinned non-AI files plus browser/operation adapters                                 | Retained-command parity and R6-01 release gate complete; ready    |

These counts are pinned-source evidence, not a completion metric by themselves. Completion is determined by the source and capability gates below.

## Shared acceptance gate

A format is complete only when all of the following pass:

1. **Pinned source inventory** — every renderer file and original focused test is classified as retained, host-adapted, or prohibited with source evidence.
2. **Complete renderer** — the original non-AI Ribbon, menus, dialogs, canvases, keyboard routes, clipboard behavior, selection, undo/redo, views, and format-specific editing surfaces remain available.
3. **Permitted removals only** — every removed capability is proven to be AI/model, login/account, telemetry, branding, updater, Electron shell/main/preload/IPC, or enterprise `ee/` behavior.
4. **Product graph** — prohibited packages and Electron runtime are absent from installed and bundled dependencies.
5. **Standalone host** — the renderer starts in a browser and provides format-appropriate open/save/export fallbacks.
6. **MCP App host** — one persistent iframe supports inline/fullscreen switching, responsive resize,
   local file open/save, and exact-session recovery. Its Session/view binding is immutable, one
   Session has one active view lease, and data operations never remount or reparent it.
7. **Shared state** — user gestures and MCP operations use the renderer's same command path, selection, undo stack, and revision.
8. **MCP capability parity** — every retained state-changing command has a stable operation id, schema, addressable context, risk/revision semantics, acknowledgement, and tests. No unexplained UI-only edit remains.
9. **Round trip** — representative open/edit/save/reopen fixtures preserve edited and untouched format data.
10. **Visual regression** — representative fixtures match the pinned renderer at stable viewport, zoom, fonts, theme, and device scale, excluding approved product removals.
11. **Performance** — cold mount, file open, interaction, command acknowledgement, save, memory, and bundle size are measured without deleting capabilities to meet an invented budget.
12. **Documentation** — protocol, capability inventory, ledger, provenance, limitations, Skill, and runbook are updated in the same change.

`office_get_capabilities.ready` is release readiness only after this gate passes. During migration, implemented transitional operations may be reported, but they must not imply renderer or capability completeness.

## Migration sequence

### M0 — Boundary and correction

Status: Complete.

- Preserve ADR 0001's mounted-renderer authority and ADR 0002's opaque local recovery.
- Accept ADR 0003 and remove the narrow-slice target decision.
- Record the former replacement strategy as rejected and retain the actual per-format MCP parity gaps.
- Keep the pinned community baseline and prohibit enterprise `ee/` inspection or import.

### M1 — DOCX complete capability audit

Status: Complete through R2-220 and the shared R6-01 release gate. The retained-command inventory, native history, revisions, document-protection, Ink, Compare, Document design, Cover-page, Drop-cap, WordArt, exact-character/symbol, references/index, comment lifecycle, link, bookmark/cross-reference/generic-field/TOC/note-lifecycle, list-numbering, paragraph-format, complete retained table, page/section, header/footer, image, chart lifecycle, shape/textbox/equation content, and object insertion/geometry/masked style/removal families are registry-owned. All former structured commands are complete tracers and the shallow batch route is removed.

- Reconcile the retained DOCX renderer against the pinned non-AI source.
- Restore any original behavior removed for reasons outside the prohibited boundary.
- Inventory all DOCX editing commands and expose each through the DOCX-owned executable operation registry.
- Preserve the verified canonical text/save paths, native undo or persistence semantics, recovery, and revision evidence while migrating adjacent operations.
- Inventory the remaining retained DOCX UI mutations beyond the former structured-command surface and add exact registry operations with native atomic-grouping rules where needed.
- Retain the current browser file, fullscreen, session, revision, recovery, and surgical-save adapters.

### M2 — XLSX complete renderer

Status: Complete through R2-131 and the shared R6-01 release gate.

Current evidence: all 76 permitted pinned renderer files are present; the 35 absent files are 20 AI
modules/prompts, 14 AI composer assets, and `strings-ai.ts`. `main.tsx` mounts the pinned `App.tsx`
directly, and the rejected substitute renderer directory is gone. All permitted original focused
tests are restored. The 114-operation XLSX registry covers all audited retained mutations through
the mounted Univer or renderer-owned file journal, and browser round trips cover the package-backed
families. See [`xlsx-capability-inventory.md`](xlsx-capability-inventory.md).

M2's renderer and mutation-parity work is complete. The format is not marked release-ready until
the shared packaged-host visual, performance/resource, MCP smoke, and repository acceptance gates
run together with the other formats.

- Restore the pinned Univer renderer, Ribbon, dialogs, worksheet state, clipboard, formula, formatting, structure, data, chart, pivot, review, view, and page-layout behavior.
- Remove only AI/account/Electron dependencies and adapt original open/save and host calls.
- Route every retained workbook mutation through MCP and the same Univer command/undo path.
- Replace `ribbon_command` with exact semantic operation families; do not add new Ribbon strings to the transitional route.
- Extend only the community App's browser host adapter and retained gateways; do not recreate a parallel renderer.

### M3 — PPTX complete renderer

Status: Complete through R2-308 and the shared R6-01 release gate.

Current evidence: all 80 permitted pinned renderer paths are present: 61 are byte-identical and 19 contain audited browser-host or prohibited-product cleanup adaptations; the remaining 24 pinned paths are prohibited AI modules/assets. The TandemFolio host adapters mount the original `App`, Ribbon, Konva canvas, dialogs, panes, views, text editor, and Presenter/Audience entry. Seventeen public browser scenarios remain, while four workspace test files / 168 assertions cover Registry and real-fixture browser primitives. Seventy-four PPTX-owned descriptors cover the retained mutation inventory, generated discovery/Broker validation, native Undo/Redo and package state. The producer baseline has no missing entry and the complete `SlidesApi` is compile-time enforced. Original focused-test classification and shared visual/performance/package release gates remain. See [`pptx-capability-inventory.md`](pptx-capability-inventory.md).

- Restore the pinned Konva renderer, Ribbon, dialogs, slide lifecycle, text, shape, picture, table, chart, design, transition, animation, slideshow, review, notes, media, master, ink, and view behavior.
- Preserve the retained PPTX parse/render engines and reconnect the original UI/state routes.
- Route every retained presentation mutation through MCP and the same command/undo path.
- Move current Agent command branching out of the renderer host Adapter and into the PPTX-owned executable operation registry before adding broad command families.
- Keep browser presentation state in `src/renderer/host/browser-presentation.ts` behind the original renderer contract.

### M4 — PDF complete renderer

Status: Format-local retained state-changing command parity and the shared R6-01 release gate complete.

Current evidence: all 33 permitted pinned renderer files are present at their original paths and
boot through `src/renderer`; seven AI files/assets remain intentionally excluded. The 25-operation
PDF registry contains 23 Agent-visible routes plus two staged-byte internals. Its machine-checked
producer baseline has no missing entry. Browser PDFium text/image mutation, PDF-lib-safe annotation,
form, stamp, metadata and page persistence, 23 test files / 280 assertions, and eleven real-host
scenarios cover the retained families. See
[`pdf-capability-inventory.md`](pdf-capability-inventory.md).

- Preserve the closed producer baseline and require every future retained state change to land as a
  typed Registry tracer with the matching native state/history and persistence evidence.
- Keep page insertion's immediate persistence/reload route explicitly non-undoable; do not imply an
  App-history guarantee that the retained host primitive cannot provide.
- Complete the shared visual, performance/resource, MCP smoke, license, provenance, and repository
  acceptance gates before changing PDF `ready`.

### M5 — Markdown complete renderer

Status: Complete through R2-238 and the shared R6-01 release gate.

- Keep the restored pinned Ribbon, frontmatter, slash/table menus, block editing, code blocks, local images, import/export, print, and DOCX export behavior.
- Preserve the audited removal of nine AI-only renderer files, the browser file adapter replacing Electron IPC, and the format-owned operation registry.
- Keep the verified Markdown session format, UI resource, local open/save/recovery, context, Skill guidance, generated packaged asset, and real MCP-client smoke in the release graph.
- Keep all 22 Registry descriptors mapped by the format-owned retained-command audit.
- Preserve session-bound/browser-directory local-image hydration without changing authored Markdown paths.

### M6 — Cross-format MCP parity and release

Status: Complete through R6-10.

- Generate or verify per-format capability inventories from the renderer command registries.
- Generate MCP manifests, installed-Skill references, and baseline-to-operation parity reports from the format registries.
- Remove central catalogs, open string dispatch, renderer composition-root operation branches, and expired compatibility aliases.
- Prove user-command and MCP-command post-state equivalence for every retained edit family.
- Keep default fullscreen plus manual inline/fullscreen switching and responsive host resize in all five editors.
- Run format-specific visual, round-trip, performance, package, license, and prohibited-dependency gates.
- Mark a format ready only when its inventory contains no unexplained retained capability gaps.
- Preserve the approved evidence/source fingerprint relationship; relevant source changes must
  recapture and reapprove the canonical suite before packaging.
- Preserve the R6-02 wakeable bounded-poll contract: one immediate bootstrap, one steady waiter,
  enqueue wakeup, bounded timeout re-arm, disconnect cleanup, and error-only 500 ms retry.
- Preserve the R6-05 XLSX first-commit boundary, one-shot strict startup trace, seven-sample
  aggregate/subphase evidence, fixed 500 ms bootstrap and 1,400 ms total p95 gates, embedded lazy
  module vault, and 11 MB initial-JavaScript ceiling without removing renderer capability.
- Preserve the R6-06 bounded discovery contract: schema-free pages of at most twenty operations,
  exact canonical detail lookup, stable cursors and availability reasons, Agent-only visibility,
  an 8 KiB summary-response gate, and a 64 KiB generated-descriptor ceiling.
- Preserve the R6-07 through R6-09 canonical transaction-only contract: session-scoped request identity, exact structural
  replay, revision guarding for new work, canonical Agent ids, pre-dispatch schema/context checks,
  one-operation atomicity, ACK-owned completion, staged-byte lifetime through final settlement, and
  schema-level rejection of the retired `operation` + `arguments` envelope, plus rejection of all
  thirty-six retired public-era aliases at renderer and packaged MCP boundaries.
- Preserve the R6-10 identity contract: blank isolated creation by default, exact-id follow-up and
  restart recovery, explicit-only newest recovery, one exclusive view lease per Session, immutable
  iframe binding, and no repeat show call for a connected Session.

## Release rule

R6-01 proves the complete renderer, MCP parity, and shared release gates for the current pinned
source; R6-02 removes the shared fixed command wait, R6-03 bounds Markdown staged loading, and
R6-04 bounds XLSX cold start while R6-05 splits and optimizes its dominant bootstrap phase without
changing renderer authority. R6-06 bounds Manifest discovery without expanding the tool surface,
R6-07 makes mutation submission revision-guarded and idempotently replayable, and R6-08 makes that
transaction envelope the only public mutation input. R6-09 makes canonical ids the only public
registry names and reduces the Manifest alias set to the five internal staged-file transports.
R6-10 prevents duplicate or cross-task editor inheritance with exact Session recovery and immutable
lease-bearing views across DOCX, Markdown, XLSX, PPTX, and PDF.
The generated release projection remains ready only against recaptured source-current evidence. Future source changes fail closed until the
same gate passes again. There is no accepted “unsupported because TandemFolio omitted it” category for a
state-changing command present in the pinned non-AI community renderer.
