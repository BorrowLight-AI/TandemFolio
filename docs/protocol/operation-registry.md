# Format-owned operation registry contract

- Status: Implemented first-release runtime contract through R6-09; format-local retained-command
  parity, shared release gates, bounded Manifest discovery, revision-guarded replay, and
  transaction-only canonical public execution pass for all five formats
- Governing decision: [ADR 0004](../adr/0004-format-owned-operation-registries.md)
- Authority model: mounted renderer remains authoritative under ADR 0001
- Applies to: DOCX, XLSX, PPTX, PDF, and Markdown

## Purpose

This contract defines how TandemFolio will expand from its transitional command tracer to complete MCP editing parity without expanding the public MCP tool set or creating a second document authority.

It separates four concerns:

1. the pinned baseline capability inventory, which defines required product scope;
2. format-owned executable operation registries, which define implemented Agent behavior;
3. generated capability projections, which let the MCP server discover and validate operations;
4. renderer Adapters, which execute operations through the same native command seam used by the UI.

This document records the accepted contract and its staged implementation. [`live-session.md`](live-session.md) remains the exact implemented wire contract.

## Implemented foundation

The R2 chronology below records the compatibility aliases used while each tracer replaced its
pre-registry route. R6-09 supersedes those transitional alias statements: all thirty-six
public-era aliases are now retired. Only five internal `open_local_file` staged-file transport
aliases remain, one per format.

R1 is implemented in `packages/operation-contract` and `tools/generate-operation-manifest.ts`. It provides the format-neutral descriptor and transaction types, stable error and validation results, deterministic manifest serialization, integrity checks, fixture generation, and `--check` drift detection.

R2-01 through R2-04 connect Markdown's current public and internal operation surface. Markdown owns the
serializable descriptors and renderer handlers for `markdown.text.insert`,
`markdown.text.replace_selection`, `markdown.document.save`, and the internal
`markdown.document.load_staged`; the generated manifest supplies
their MCP discovery schemas and broker-side validation. Text handlers execute through TipTap's
native history. The async document handler invokes the format-owned browser persistence service and
returns exact `{ saved, fileName }` output or `execution_failed`. The legacy `insert_text`,
`replace_selection`, and `save` names are accepted only as Agent compatibility aliases. The legacy
`open_local_file` transport name is an internal alias. All are normalized to canonical ids before
renderer execution. Internal visibility keeps staged load out of capability discovery and blocks it
from `office_execute`; only `office_open_local_file` and recovery transport may enqueue it. Markdown's
composition-root operation branches, the transitional `markdown-commands.ts` helper, and its
hand-written public capability entries have been deleted.

R2-05 adds the first cross-format tracer: DOCX owns `docx.text.insert`, which is projected from the
same product manifest, validated and canonicalized by the broker, and executed through the mounted
DOCX TipTap history seam before the existing recovery/revision acknowledgement. Its legacy
`insert_text` alias is scoped to DOCX, so Markdown may retain the same input-only alias without a
global collision. The direct DOCX insertion branch and hand-written capability schema are deleted.

R2-06 adds `docx.text.replace_selection` through the same generated projection and mounted-renderer
seam. Legacy `replace_selection` is an input-only DOCX alias, invalid canonical arguments are
rejected before enqueue, and the TipTap replacement enters the native undo history before the
existing recovery/revision acknowledgement. The direct App and central capability branches are
deleted, so both current DOCX text operations are now registry-owned.

R2-07 adds `docx.document.save` through an injected format-owned persistence service. It declares
exact empty input and `{ saved, fileName }` output, maps persistence failure to `execution_failed`,
and keeps save non-undoable. Legacy `save` is an input-only DOCX alias. Broker integration proves
canonical queueing, output propagation, exact validation, and removal of the format recovery
snapshot only after acknowledged success. The direct App and central capability save branches are
deleted.

R2-08 adds internal `docx.document.load_staged` through an injected format-owned load service. The
generated Manifest retains its exact staged descriptor and `open_local_file` alias but filters it
from Agent discovery and direct `office_execute`. The DOCX host adapter hydrates chunked bytes for
the canonical id; the renderer registry checks `ArrayBuffer`, `.docx` identity, and declared byte
length, maps loader failures, and suppresses a recovery checkpoint after successful load. The old
composition-root branch and command-shaped helper are deleted, so `batch_update` is DOCX's only
remaining transitional operation dispatch.

R2-09 migrates the first structured command out of that batch as `docx.text.replace_all`. The exact
descriptor declares document context/effects, medium risk, one native undo unit, optional
case-insensitive matching, and structured replacement/skipped-content counts. Its handler invokes
the existing `executeCommands` ProseMirror transaction instead of adding another replacement
implementation. Broker validation rejects malformed input before enqueue, renderer validation
rejects an empty search without mutation, and the composition root preserves both the operation
output and the usual dirty recovery snapshot. `batch_update.replaceAllText` is rejected and the
other nine batch variants remain transitional.

R2-10 adds the first bounded structured-target tracer. The common contract now supports and
enforces `enum`, array `items`, and inclusive `minimum`/`maximum` without learning DOCX target
semantics. The DOCX-owned `docx.paragraph.set_heading_level` descriptor uses those primitives for
its full block target and level 0–6 constraint; its handler delegates to the existing
`setHeadingLevel` transaction. Invalid enum, array-item, and range values stop at the Broker,
conditionless targets stop at the renderer seam, and valid changes preserve output, recovery,
revision, and one-step undo. `batch_update.setHeadingLevel` is rejected, leaving eight batch
variants transitional.

R2-11 extracts DOCX's complete block target as the reusable format-owned
`docxBlockTargetSchema` and consumes it in both heading-level mutation and the new high-risk
`docx.block.delete`. The delete handler invokes the existing native `deleteBlocks` transaction,
including tracked-deletion accounting and the required empty-paragraph fallback when every block
is deleted. Exact Broker validation, renderer semantic validation, output/recovery/revision, and
one-step undo remain on the same path. `batch_update.deleteBlocks` is rejected, leaving seven batch
variants transitional.

R2-12 reuses that same DOCX target contract in a third operation family. `docx.list.remove`
delegates to the retained `deleteParagraphBullets` command, converting only matching list items to
body paragraphs without changing their text. Exact Broker validation, renderer semantic
validation, deterministic output, recovery, one revision advance, and one-step undo remain on the
registry path. `batch_update.deleteParagraphBullets` is rejected, leaving six batch variants
transitional.

R2-13 completes the current list pair with `docx.list.apply`. Its canonical `{ target, kind }`
input narrows the old prefix-based `bulletPreset` string to `bullet | ordered`, while the handler
maps back to the retained `createParagraphBullets` transaction and the mounted App's numbering
allocator. Allocation is lazy, so a same-kind no-op keeps its existing numId without creating an
unused definition. Exact Broker validation, renderer semantic validation, output, recovery, one
revision advance, and one-step undo stay on the registry path. The old batch variant is rejected,
leaving five transitional variants.

R2-14 through R2-18 complete decomposition of the current DOCX structured-command surface.
`docx.text.set_style` and `docx.paragraph.set_style` preserve the native target/style/field-mask
behavior; `docx.block.move` preserves ordered block relocation and renderer-context checks;
`docx.image.update` preserves proportional dimension scaling and alignment; and `docx.toc.insert`
builds real TOC field blocks from current headings. The common contract grows only the bounded
nullable-type and `minItems`/`maxItems` mechanics required to validate these exact format-owned schemas. Every
operation queues canonically, returns deterministic output, advances one acknowledged revision,
checkpoints recovery, and uses one native undo unit. After the last tracer, the public
`batch_update` schema, DOCX App branch, and helper are deleted.

R2-19 starts the XLSX registry with `xlsx.document.save`. Its empty input and exact
`{ saved, fileName }` output come from the XLSX-owned catalog; the handler calls the retained save
assembler and browser package/local-persistence boundary used by the mounted community App. The save action
now distinguishes persisted success, cancellation, and write failure so the renderer can issue a
deterministic acknowledgement. Legacy `save` is an input-only format-scoped alias normalized by
the Broker. The direct App save branch and hand-written XLSX save capability entry are deleted.

R2-20 adds internal `xlsx.document.load_staged` through the same XLSX registry. Its generated
descriptor and `open_local_file` compatibility alias remain available only to the staged-file
transport: Agent discovery filters them and direct `office_execute` rejects them. After the shared
host bridge hydrates bounded chunks, the renderer validates `ArrayBuffer`, `.xlsx` identity, and
declared length before invoking the retained `openBuffer` → `openLazyWorkbook` path. Success returns
`{ opened, fileName }` without checkpointing the freshly loaded document; load failures are
acknowledged as `execution_failed`. The direct XLSX App branch is deleted.

R2-21 through R2-25 move the first five XLSX mutations behind that registry:
`xlsx.cell.set_value`, `xlsx.range.set_values`, `xlsx.row.insert`, `xlsx.row.delete`, and
`xlsx.column.insert`. Their exact schemas bound scalar matrices and structural counts before
enqueue; renderer handlers resolve the mounted Univer workbook and call its native range/structure
methods. Legacy names are input-only aliases, and the corresponding hand-written capability and
App branches are deleted. At R2-25, six XLSX routes remained transitional.

R2-26 through R2-30 migrate `xlsx.column.delete` and the complete
`xlsx.sheet.add|rename|delete|move` family. The renderer preserves bounded column targets,
worksheet-name rules, final-sheet protection, dynamic tab-position bounds, and native Univer
mutation/undo behavior. Their legacy names are input-only aliases; the matching central schemas,
composition-root helper, and App branches are deleted. Only `ribbon_command` remains transitional
for XLSX.

R2-31 starts decomposing the remaining open XLSX Ribbon route with
`xlsx.range.set_text_style`. Its exact `{ sheet, range, style, fields }` contract supports explicit
bold, italic, strike, and `none | single | double` underline values. The field mask must contain
unique entries and exactly match the supplied style keys. The renderer resolves the named range and
submits one Univer `setValue({ s })` mutation for the complete patch, preserving the mounted state,
selection, native undo, and one acknowledged revision. MCP `ribbon_command` calls in those four
families are rejected after migration, while visible Ribbon buttons continue to use the retained UI
dispatcher. Other Ribbon families remain transitional.

R2-32 through R2-36 migrate the next three Ribbon formatting groups into five exact operations:
`xlsx.range.set_alignment`, `xlsx.range.set_font`, `xlsx.range.set_fill`,
`xlsx.range.set_border`, and `xlsx.range.apply_cell_style`. Alignment and font use unique exact
field masks; rotations distinguish angle, stacked, and explicit clear; colors are bounded to
`#RRGGBB`; border presets require explicit line style/color unless clearing; and named cell styles
are a bounded preset enum. Multi-field alignment, font, and cell-style changes each submit one
native Univer mutation, while border uses Univer's native range border command. The visible Ribbon
continues to own user gestures, but the migrated Ribbon strings are rejected through MCP.
`format-painter` is deliberately not an Agent operation: it arms a transient UI mode whose result
depends on a later user selection, so Agents must express the intended destination formatting with
the explicit range operations.

R2-37 through R2-40 migrate the next three retained Ribbon groups into four semantic operations.
`xlsx.range.set_number_format` assigns one explicit bounded pattern; the relative `decimal-inc` and
`decimal-dec` UI gestures therefore become deterministic final-state assignments for Agents.
`xlsx.range.merge` exposes `cells`, `across`, `center`, and `unmerge`; cells, across, and unmerge
are direct Univer range actions, while `center` deliberately retains the UI's merge-then-align
sequence and is marked `atomic: false`. `xlsx.range.clear` separates contents, formats, and all-data
clearing behind an exact scope enum. `xlsx.range.fill` activates an explicit multi-row or multi-column
range before invoking Univer's copy-down or copy-right command, rejecting targets without a
destination. All four operations share mounted state, selection, edit journal, undo/redo, revision,
and browser save output. Their old MCP Ribbon strings are rejected after migration, while the visible
Ribbon continues to own the same user gestures.

R2-41 through R2-43 migrate the retained sorting and row-deduplication group into three semantic
operations. `xlsx.range.sort` preserves the visible Ribbon's first-column behavior with an exact
`asc | desc` direction. `xlsx.range.sort_custom` accepts ordered, unique A1 column keys that must
fall inside the explicit target and carries an explicit header flag. Both call Univer's native sort
command and require at least two rows. `xlsx.range.remove_duplicates` preserves the renderer's
case-insensitive whole-row comparison and first-occurrence behavior, rejects partially streamed
sheet data before reading, reports the removed count, and deliberately remains non-atomic because
it rewrites only rows that move or clear so unchanged formulas survive. Their old MCP Ribbon strings
are rejected; visible Ribbon gestures retain their original routes.

R2-44 through R2-48 migrate hyperlink set/remove, Format as Table, cell-protection flags, and
worksheet protection into `xlsx.hyperlink.set`, `xlsx.hyperlink.remove`, `xlsx.table.add`,
`xlsx.range.set_protection`, and `xlsx.sheet.set_protection`. Hyperlinks normalize the retained URL
and internal-sheet forms and share the Ribbon's journal/style seam. Table creation accepts only the
six retained gallery styles and deliberately does not advertise an unsupported “restyle existing
table” operation. Cell protection uses an exact `locked | hidden` field mask and rejects targets
above 10,000 cells before journaling. Worksheet protection replaces a toggle with an explicit
boolean and fails closed when password state prevents removal. These file-side operations are
truthfully non-undoable where Univer has no matching model. R2-78 through R2-81 connect hyperlink,
table, and protection writes to the browser package boundary and hydrate their native state on
reopen.

R2-49 migrates the retained AutoSum family into `xlsx.formula.insert_aggregate`. Its exact
`SUM | AVERAGE | COUNT | MAX | MIN` enum replaces the open `autofn:*` string, and both Ribbon and
registry callers use one XLSX-owned formula seam. The seam places one aggregate formula below each
selected column, rejects source ranges shorter than two rows, and refuses an unstreamed destination
row before any write. Multi-column insertion preserves the retained per-column native mutations,
so the descriptor is undoable but non-atomic.

R2-50 migrates retained Flash Fill into `xlsx.range.flash_fill`. The exact `{ sheet, range }`
contract replaces the open `flash-fill` string while both Ribbon and registry callers share one
XLSX-owned inference seam. It reads at most six source columns, learns from at most three examples,
preserves non-empty targets, and commits inferred empty targets as one native range write. A
single-cell target probes at most 1,000 adjacent-left rows; both that probe and explicit inference
rectangles reject partially streamed original-file data before any read or write.

R2-51 migrates retained Text to Columns into `xlsx.range.text_to_columns`. Its exact
`{ sheet, range, delimiter }` contract accepts only `tab | comma | semicolon | space`, requires one
source column and fully loaded worksheet data, and sends both Ribbon and registry callers through
Univer's native split command. The native command may overwrite destination cells to the right or
insert columns when necessary, so the descriptor is high-risk; its grouped native mutation remains
undoable and atomic.

R2-52 migrates retained Row Height into `xlsx.row.set_height`. Its exact
`{ sheet, row, count, heightPoints }` contract replaces the selection-dependent `row-height:*`
string with a bounded 1-based contiguous row target and a `0–409.5` point height. Ribbon and
registry callers share the same point-to-pixel native seam, while the renderer rejects spans beyond
the current worksheet. The retained Univer command is atomic and undoable, and its axis-attribute
journal survives save/reopen.

R2-53 migrates retained Column Width into `xlsx.column.set_width`. Its exact
`{ sheet, column, count, widthCharacters }` contract replaces `col-width:*`, bounds the requested
Excel width to `0–255`, rejects spans beyond the mounted worksheet, and shares the native column
mutation with Ribbon callers. Because Univer applies integer pixels, the deterministic result and
OOXML journal report the actual 1/256-character width after quantization.

R2-54 migrates the four retained freeze-pane strings into `xlsx.sheet.set_freeze`. Its exact
`{ sheet, frozenRows, frozenColumns }` final state uses `0,0` for unfreeze and rejects counts that
would leave no scrollable pane. Ribbon and registry callers share the native freeze seam and page-
setup journal, preserving one undo unit and saved/reopened pane coordinates.

R2-55 migrates retained gridline toggling into `xlsx.sheet.set_gridlines`. Its exact
`{ sheet, visible }` contract assigns final visibility rather than depending on current state. The
Ribbon derives its next explicit value before entering the same native seam; Univer undo and the
existing `sheetView@showGridLines` journal remain authoritative.

R2-56 migrates retained formula-view toggling into `xlsx.sheet.set_formula_view`. Its exact
`{ sheet, enabled }` contract assigns a final per-sheet formula-view state. Ribbon and registry
callers share the retained formula projection and `sheetView@showFormulas` journal, while the
generalized renderer-owned `pushWorkbookUndo` adapter places the change in Univer's undo history.
Undo/redo, saved output, and reopen all re-enter the same seam.

R2-57 migrates retained print-orientation selection into `xlsx.sheet.set_page_orientation`. Its
exact `{ sheet, orientation }` contract accepts only `portrait | landscape`. Ribbon and registry
callers share the format-owned page-layout journal seam, and renderer-owned history restores the
exact prior orientation override before Redo reapplies the final state. The saved
`pageSetup@orientation` value survives reopen.

R2-58 migrates retained margin-preset selection into `xlsx.sheet.set_page_margins`. Its exact
`{ sheet, margins }` contract accepts only `normal | wide | narrow`. Ribbon and registry callers
share the format-owned page-layout journal seam; a private page-setup preset module restores the
exact prior field state through renderer-owned Undo/Redo, while the retained writer expands each
preset into the six OOXML `pageMargins` values. Saved values survive reopen.

R2-59 migrates retained paper-size selection into `xlsx.sheet.set_paper_size`. Its exact
`{ sheet, paperSize }` contract accepts only the seven visible OOXML paper-size codes `1`, `3`, `5`,
`7`, `8`, `9`, and `11`. Ribbon and registry callers share the private page-setup preset seam;
renderer-owned Undo/Redo restores exact field presence, and the retained writer persists
`pageSetup@paperSize` through save and reopen.

R2-60 migrates retained fit-width and fit-height selection into `xlsx.sheet.set_fit_to_pages`. Its
exact `{ sheet, widthPages, heightPages }` contract accepts two explicit `0–1000` integer axes;
`0,0` disables fit-to-page. Ribbon single-axis gestures resolve the untouched effective axis before
entering the shared multi-field page-setup patch seam. One renderer-owned Undo/Redo unit restores
exact journal-field presence, and the retained writer persists both `pageSetup` axes plus
`pageSetUpPr@fitToPage` through save and reopen.

R2-61 migrates retained printed-gridline toggling into `xlsx.sheet.set_print_gridlines`. Its exact
`{ sheet, enabled }` contract assigns final print output without conflating it with worksheet display
gridlines. Ribbon and registry callers share the private page-setup patch seam; renderer-owned
Undo/Redo restores exact field presence, and the retained writer persists
`printOptions@gridLines` through save and reopen.

R2-62 migrates retained printed-heading toggling into `xlsx.sheet.set_print_headings`. Its exact
`{ sheet, enabled }` contract assigns final row/column-heading print output. Ribbon and registry
callers share the private page-setup patch seam; renderer-owned Undo/Redo restores exact field
presence, and the retained writer persists `printOptions@headings` through save and reopen.

R2-63 migrates retained print-area setting and clearing into `xlsx.sheet.set_print_area`. Its exact
`{ sheet, range }` contract accepts a normalized explicit A1 cell range or `null`; the retained
Ribbon resolves its current selection before entering the same private page-setup patch seam.
Renderer-owned Undo/Redo restores exact field presence, while the retained writer persists or
removes the sheet-scoped `_xlnm.Print_Area` defined name through save and reopen.

R2-64 migrates retained print-title Row 1, selected-row, and clear commands into
`xlsx.sheet.set_print_titles`. Its exact `{ sheet, rows }` contract accepts an ascending explicit
row span of at most 21 rows or `null`; the retained Ribbon resolves Row 1 or its current selected
rows before entering the same private page-setup patch seam. Renderer-owned Undo/Redo restores
exact field presence, while the retained writer persists or removes the sheet-scoped
`_xlnm.Print_Titles` defined name through save and reopen.

R2-65 migrates retained fixed print-scale command strings into `xlsx.sheet.set_print_scale`. Its
exact `{ sheet, scalePercent }` contract accepts an integer from 10 through 400. Agent and the
page-layout dispatcher share one multi-field patch that records scale and disables fit-to-page as
one renderer-owned Undo/Redo unit. The retained writer removes `pageSetUpPr@fitToPage`, persists
`pageSetup@scale`, and preserves the result through reopen. The pinned Ribbon contains no visible
Scale control, so no additional UI is introduced.

R2-66 migrates retained `filter-toggle` into `xlsx.range.set_filter`. Its exact
`{ sheet, range, enabled }` contract names the target and desired final state, making Agent retries
idempotent instead of replaying a toggle. Agent and Ribbon enter one format-owned seam over Univer's
native set/remove filter commands; the Ribbon resolves its current selection or active filter before
dispatch. A conflicting existing filter fails rather than silently moving it, and the browser package
adapter persists and reopens the resulting `autoFilter@ref` through the retained XLSX filter gateway.

R2-67 through R2-69 finish the retained worksheet-filter condition family.
`xlsx.range.clear_filter_criteria` clears every criterion but retains the exact AutoFilter range;
`xlsx.range.set_filter_values` applies a bounded 1–10,000 item value list and explicit blank state;
`xlsx.range.set_custom_filter` applies one or two conditions from the six saveable operators with
an explicit conjunction. The common Contract adds `maxItems`; the renderer additionally rejects a
missing filter, mismatched range, or column outside that range. Agent, Ribbon Clear, and the
Advanced Filter dialog share one deep filter seam over Univer's native Undo commands, while the
retained gateway persists criteria and filtered-row visibility through save/reopen.

R2-70 retires five selection-relative structure strings without adding parallel operations.
`insert-sheet`, `insert-row-here`, `delete-row-here`, `insert-col-here`, and `delete-col-here` are
rejected on the Agent `ribbon_command` path with pointers to the existing exact sheet/row/column
operations. Visible controls continue to use the retained workbook and worksheet structure routes,
so their native history and saved output do not change.

R2-71 migrates native workbook history into `xlsx.history.undo {}` and
`xlsx.history.redo {}`. Both operations call the mounted Univer API directly, share the same history
as the visible controls, and fail without advancing revision when the requested history entry is
unavailable. The old Agent `ribbon_command` strings are rejected; UI controls retain their native
route.

R2-72 migrates `paste-special:value` into
`xlsx.range.copy_values { sourceSheet, sourceRange, destinationSheet, destinationRange }`. The
operation materializes computed scalar values between equal-shape ranges of at most 20,000 cells,
supports cross-sheet destinations, and rejects invalid or partially streamed source/destination
rectangles before reading or writing. One native destination matrix write owns Undo/Redo and saved
output; the visible Ribbon keeps its retained Univer clipboard route.

R2-73 migrates `paste-special:formula` into
`xlsx.range.copy_formulas { sourceSheet, sourceRange, destinationSheet, destinationRange }`. It
shares value-copy's worksheet-bounded A1, equal-shape, 20,000-cell, cross-sheet, and streaming
guards. Univer's mounted formula lexer translates relative and mixed references by the exact
source-to-destination offset while preserving absolute components; non-formula cells copy scalar
values, matching the retained Paste Formulas behavior. One native destination matrix write owns
Undo/Redo and saved formula output; the visible Ribbon keeps its retained clipboard route.

R2-74 migrates `paste-special:format` into
`xlsx.range.copy_formats { sourceSheet, sourceRange, destinationSheet, destinationRange }`. It
shares the range-copy bounds and replaces all cell-style fields rather than merging them, while
preserving destination values, formulas, and rich content. One low-level Univer style mutation and
its exact inverse form a single native Undo/Redo item; saved output remains on the renderer journal.
Merged-cell topology stays explicit through `xlsx.range.merge`, and the visible Ribbon keeps its
retained clipboard route.

R2-75 migrates `paste-special:col-width` into
`xlsx.column.copy_widths { sourceSheet, sourceColumn, destinationSheet, destinationColumn, count }`.
Both column spans use normalized A1 labels, remain independently worksheet-bounded, and share the
Contract's 1–10,000 structural count limit. The renderer snapshots the source width vector before
one low-level Univer column-width mutation, registers its exact inverse as one native Undo/Redo
item, preserves cell contents, and persists through the axis journal. Column widths are eagerly
available worksheet metadata, so this operation does not pretend they depend on cell-range
streaming. The visible Ribbon keeps its retained clipboard route.

R2-76 migrates `paste-special:besides-border` into
`xlsx.range.copy_without_borders { sourceSheet, sourceRange, destinationSheet, destinationRange }`.
It reuses the worksheet-bounded A1, equal-shape, 20,000-cell, cross-sheet, and dual-streaming
guards. The renderer snapshots source cell content and non-border styles, translates formulas by
the exact destination offset, and omits the incoming `bd` style field so Univer's native style
merge retains every destination border. Non-active-sheet formulas use the read-only file formula
index only when the mounted model has no formula; current edit-journal content takes precedence, so
the mounted session remains authoritative. One native range write owns Undo/Redo and saved output.
The visible Ribbon keeps its retained clipboard route.

R2-77 classifies Agent `copy`, `cut`, and `paste` as nondeterministic clipboard/selection-state UI
gestures. `ribbon_command` rejects all three and directs Agent callers to explicit
source/destination operations; the retained Univer UI controls remain unchanged. This migration
adds no operation descriptor.

R2-78 connects the existing `xlsx.hyperlink.set/remove` journal to the browser package boundary.
The adapter reuses the retained worksheet/relationship gateway, indexes saved targets on reopen,
and feeds the mounted App's existing hyperlink context and click route. The registry and manifest
do not grow.

R2-79 through R2-81 connect the already registered table and protection families to the browser
package boundary. Table writes use the retained gateway behind a transactional overlay; browser
open resolves native table relationships and hydrates table metadata into the existing Univer
path. Cell protection persists through the stylesheet editor, while sheet protection uses the
retained OOXML gateway and rehydrates password/protected state. The registry and manifest do not
grow.

R2-82 migrates retained sparkline creation to
`xlsx.sparkline.add { sheet, sourceRange, targetRange, type }`. The exact operation bounds output to
200 row-aligned cells, requires a one-column non-overlapping target, rejects occupied hosts, and
shares one journal/history/projection action with the Ribbon. The browser package adapter uses the
retained x14 writer and hydrates saved sparkline groups again on open.

R2-83 migrates the eight retained outline strings into two explicit final-state operations:
`xlsx.outline.set_level { sheet, axis, start, count, level }` and
`xlsx.outline.set_detail_visibility { sheet, axis, start, count, hidden }`. Relative Ribbon
Group/Ungroup gestures resolve absolute levels before entering the shared action. Detail visibility,
the following summary item's collapsed state, the structural journal, and the live renderer enter
one Univer history item. Browser open hydrates row and column outline metadata so explicit expansion
after reopen does not depend on an empty session cache. The generated product manifest now contains
eighty-five operations, including sixty-one XLSX operations.

R2-84 migrates `insert-checkbox` to
`xlsx.range.set_checkbox { sheet, range, enabled }`. The exact operation bounds the target to
10,000 cells and uses Univer's native validation command/Undo stack. The browser adapter now writes
declarative validation snapshots through the retained gateway and parses base OOXML rules on
reopen, including the `list "1,0"` checkbox round trip. The generated product manifest now contains
eighty-six operations, including sixty-two XLSX operations.

R2-85 retires Agent `insert-symbol` in favor of existing `xlsx.cell.set_value`. The visible Symbol
dialog resolves the appended final string and uses the same native write/activation action as the
Registry handler. It remains a UI picker rather than a new operation family, so the manifest does
not grow.

R2-86 retires Agent `import-csv` as a browser picker/decoder in favor of bounded
`xlsx.range.set_values` or staged CSV open. The visible importer and Registry matrix operation
share one native range write/activation action; no new descriptor is added.

R2-87 adds `xlsx.chart.add { sheet, dataRange, type, anchorCell? }`. Its chart type is an enum and
its source is limited to 2,000 cells. The operation resolves the worksheet, calls the retained
visual-add action, enters mounted history as one unit, and persists through the browser drawing
gateway. Agent `insert-chart:*` and `recommended-charts-open` are rejected in favor of this final
intent. The generated product manifest contains eighty-seven operations, including sixty-three
XLSX operations.

R2-88 adds public `xlsx.image.add { sheet, path, anchorCell }` and internal
`xlsx.image.add_staged { blobId, name, size, data, sheet, anchorCell }`. The Broker validates and
stages at most 20 MB of matching PNG/JPEG/GIF bytes; the live bridge hydrates them outside Agent
JSON. Retained UI producers and the internal handler converge on the visual journal, Undo route,
and browser media writer. The manifest contains eighty-nine operations, including sixty-five XLSX
operations; staged transport is hidden from discovery and direct `office_execute`.

R2-89 adds `xlsx.note.set { sheet, address, text }` and
`xlsx.note.remove { sheet, address }`. The registry validates one explicit cell and bounds note
text to Excel's 32,767-character limit before invoking Univer native note commands. Both operations
enter the mounted history and declarative note journal. Browser save/reopen preserves legacy
comments and VML anchors. The manifest contains ninety-one operations, including sixty-seven XLSX
operations.

R2-90 adds `xlsx.chart.update { chartId, ...finalState }` and
`xlsx.chart.remove { chartId }`. Update accepts at least one bounded title, convertible type,
legend, data-label, grouping, or axis-title property. Both resolve an existing visual ID, use the
mounted chart/visual edit refs and native history, and browser save consumes chart/visual edit
journals. The manifest contains ninety-three operations, including sixty-nine XLSX operations.

R2-91 adds `xlsx.chart.set_colors { chartId, seriesColors?, pointColors? }` and
`xlsx.chart.set_series { chartId, series }`. It rejects empty/unbounded arrays, non-hex colors,
point colors on non-pie charts, non-finite values, oversized series, and mismatched category
vectors before mutation. Both call the same chart edit/history route. The manifest contains
ninety-five operations, including seventy-one XLSX operations.

R2-92 completes the browser-host chart reopen path without adding a new descriptor.
`xlsx.chart.update` now also accepts bounded gridline, value-axis bound, bar-gap, doughnut-hole,
pie-explosion, and data-label position/format final states. Browser open follows the worksheet,
drawing, and chart relationships to restore chart metadata, anchors, save locators, and a stable
`file-chart-*` ID; the live context publishes those IDs so a reopened chart can be updated and
saved again. The manifest remains at ninety-five operations, including seventy-one XLSX
operations.

R2-93 adds `xlsx.shape.add { sheet, type, anchorCell, fillColor?, text? }`. The `type` enum is the
retained Shape Gallery's bounded preset set plus `textbox`; anchors are explicit cells, colors are
exact hex values, and text is capped at 1,000 characters. Agent and retained UI insertion converge
on the visual journal and renderer-owned Undo route. Browser save writes native drawing shapes and
reopen restores their anchor, preset, fill, text, save locator, and stable `file-shape-*` ID. The
manifest contains ninety-six operations, including seventy-two XLSX operations.

R2-94 adds `xlsx.shape.update { shapeId, anchorCell?, fillColor?, text? }` and
`xlsx.shape.remove { shapeId }`. Stable session or reopened IDs resolve one shape; update requires
at least one bounded final-state field and preserves frame size when moving. File-native shape text
and fill now travel through the drawing edit journal instead of remaining session-only. Native
drawing XML is patched surgically and all operations share renderer Undo. The manifest contains
ninety-eight operations, including seventy-four XLSX operations.

R2-95 adds `xlsx.image.move { imageId, anchorCell }` and
`xlsx.image.remove { imageId }`. Browser open now resolves drawing image relationships into stable
`file-image-*` IDs and media locators; `readWorkbookMedia` returns the exact package bytes so
reopened images render. Move preserves frame size and remove uses the shared visual history/save
route. The manifest contains one hundred operations, including seventy-six XLSX operations.

R2-96 adds `xlsx.defined_name.set { name, formula, scopeSheet?, previousName? }` and
`xlsx.defined_name.remove { name, scopeSheet? }`. Scope is expressed by public worksheet name and
defaults to workbook scope; `previousName` identifies an existing name for atomic rename/update.
Names, formulas, and scopes are bounded before the shared native Univer action layer runs. The
Name Manager uses the same layer, and declarative workbook XML save/reopen proves persistence. The
manifest contains one hundred two operations, including seventy-eight XLSX operations.

R2-97 adds
`xlsx.range.set_list_validation { sheet, range, values, allowBlank, showDropdown }` and
`xlsx.range.remove_data_validation { sheet, range }`. Inline values are bounded to 100 non-empty,
comma-free options and a combined 255-character XLSX source; targets are capped at 10,000 cells.
Both operations share the native Univer validation/history route, streaming and metadata guards,
and declarative browser save/reopen/removal boundary. The manifest contains one hundred four
operations, including eighty XLSX operations.

R2-98 adds `xlsx.range.set_number_between_validation`,
`xlsx.range.set_date_between_validation`, and `xlsx.range.set_custom_formula_validation`. Their
contracts require explicit blank handling and bounded, semantically valid numeric/date/formula
inputs before the shared 10,000-cell validation guard runs. Each builds a native Univer validation
rule, so history and declarative XLSX persistence remain shared with the retained panel. The
manifest contains one hundred seven operations, including eighty-three XLSX operations.

R2-99 adds
`xlsx.range.set_list_reference_validation { sheet, range, sourceSheet, sourceRange, allowBlank, showDropdown }`.
The source is one row or column of at most 1,000 cells; the target retains the shared 10,000-cell
limit. Both ranges must finish streaming before the range-backed native list builder runs. Saved
OOXML retains the explicit source formula across reopen. The manifest contains one hundred eight
operations, including eighty-four XLSX operations.

R2-100 adds
`xlsx.range.set_comparison_validation { sheet, range, kind, operator, operand1, operand2?, allowBlank }`.
One semantic contract covers five scalar kinds and eight operators, with type/arity/order checks
before the shared target guard. It writes the exact native validation-panel rule model through
FRange's add-rule/Undo command. The specialized R2-98 number/date-between descriptors are removed;
the manifest therefore contains one hundred seven operations, including eighty-three XLSX
operations, while supporting a larger validation matrix.

R2-101 adds
`xlsx.range.set_validation_messages { sheet, range, inputTitle, inputMessage, errorStyle, errorTitle, errorMessage }`.
All message fields are explicit strings or `null`; titles are limited to 32 characters and bodies
to 255. The target must resolve one existing validation rule. FDataValidation's native
update-options command owns mutation and Undo, and the existing wire mapper persists prompt/error
attributes. The manifest contains one hundred eight operations, including eighty-four XLSX
operations.

R2-102 adds
`xlsx.conditional_format.set_comparison { sheet, range, ruleId, operator, operand1, operand2?, format, stopIfTrue }`
and `xlsx.conditional_format.remove { sheet, ruleId }`. A null `ruleId` creates a rule and returns
its session ID; a non-null ID replaces only an existing numeric-comparison rule. Workbook context
publishes at most 100 rule IDs with sheet, ranges, kind, and `stopIfTrue`, so update/remove never
depend on selection or hidden UI state. The target is limited to 10,000 fully loaded cells and
requires complete CF metadata. Native Univer add/set/delete commands own history, while desktop and
browser hosts share the declarative CF/DXF save mapper. The manifest contains one hundred ten
operations, including eighty-six XLSX operations.

R2-103 adds
`xlsx.conditional_format.set_highlight { sheet, range, ruleId, predicate, text, format, stopIfTrue }`.
One descriptor covers `textContains`, `textNotContains`, `textStartsWith`, `textEndsWith`, `blank`,
`nonBlank`, `duplicate`, and `unique`. Text predicates require a non-empty value of at most 255
characters; state predicates require `text: null`. Create/update uses the same explicit session ID,
loaded target, native history, and declarative CF/DXF persistence contract as R2-102. The manifest
contains one hundred eleven operations, including eighty-seven XLSX operations.

R2-104 adds
`xlsx.conditional_format.set_statistical { sheet, range, ruleId, kind, direction, rank, percent, inclusive, format, stopIfTrue }`.
For `kind: rank`, direction is top/bottom, rank is 1..1,000 (or 1..100 in percent mode), percent is
boolean, and inclusive is null. For `kind: average`, direction is above/below, rank and percent are
null, and inclusive is boolean. Invalid cross-kind combinations fail before workbook access. The
operation reuses the R2-102 lifecycle and save boundaries. The manifest contains one hundred twelve
operations, including eighty-eight XLSX operations.

R2-105 adds
`xlsx.conditional_format.set_formula { sheet, range, ruleId, formula, format, stopIfTrue }`.
Formulas must start with `=` and contain 2..8,192 characters; updates accept only an existing formula
rule. The manifest contains one hundred thirteen operations, including eighty-nine XLSX operations.

R2-106 adds
`xlsx.conditional_format.set_visual { sheet, range, ruleId, kind, colors, thresholds, iconSet, showValue, reverse, gradient, stopIfTrue }`.
Kind-specific validation admits only losslessly saveable base-OOXML color scales, gradient data
bars, and whitelisted icon sets with coherent threshold counts. Thresholds are typed as
min/max/number/percent/percentile/formula, and icon strictness is explicit. x14-only configurations
remain rejected. The manifest contains one hundred fourteen operations, including ninety XLSX
operations.

R2-107 expands `set_highlight` with text equality/inequality and error/non-error predicates, and
rehydrates their OOXML forms as equivalent native formula rules. The same dry-run validator used by
the UI and save path rejects date-occurring rules, equal/not-equal average rules, non-whitelisted or
custom-order icon sets, solid data bars, and distinct negative data-bar colors. These are x14-only
states rather than missing Agent tools. The descriptor count is unchanged.

R2-108 adds explicit Conditional Formatting clear scope. `scope: range` requires a bounded A1
range, while `scope: sheet` requires `range: null`; both return the number of rules cleared through
Univer's native history. R2-109 adds absolute one-based rule priority over the session-published
rule IDs. The product manifest contains one hundred sixteen operations, including ninety-two XLSX
operations.

R2-110 deletes the open-ended Agent `ribbon_command` contract and its renderer dispatch branch.
Unknown strings can no longer receive a successful acknowledgement. The visible Ribbon continues
to use its native dispatcher.

R2-111 through R2-130 complete the remaining audited XLSX persisted families: Pivot lifecycle and
PivotCharts, explicit formula writes, sheet duplicate/visibility/tab color, row/column visibility,
row/range movement, full table lifecycle, bounded replacement/subtotals/consolidation, and
header/footer. Every new route uses a stable target, exact bounded schema, mounted renderer action,
and browser save/reopen evidence. R2-131 adds executable retained-command mapping and regenerates
the product Manifest at 138 total operations, including 114 XLSX operations (112 Agent-visible and
two internal). No retained state-changing XLSX command remains unexplained; release readiness still
depends on shared non-registry gates.

R2-132 records the retained DOCX command-parity baseline and distinguishes state-changing semantic
families from file pickers, view state, navigation, and clipboard arming. R2-133 adds
`docx.history.undo {}` and `docx.history.redo {}` as the first post-audit tracers. Their renderer
handlers invoke the native history on the already mounted TipTap editor—the same authority used by
the Ribbon and quick-access controls—and return `{ undone: true }` or `{ redone: true }` only when
an entry was applied. Missing entries return `execution_failed`. The generated Manifest contains
140 operations, including sixteen DOCX operations (fifteen public and one internal).

R2-134 introduces the revision-scoped DOCX character range as a reusable exact target.
`docx.text.set_character_format { range, format, fields }` requires a non-empty in-document range
and a unique field mask exactly matching the supplied final values. Bold, italic, underline,
strike, and vertical alignment are applied in one native ProseMirror transaction; per-run
`docTextStyle` attributes outside the mask survive. The retained Ribbon uses the same helper for
non-empty selections, and one native Undo reverts the operation. The generated Manifest contains
141 operations, including seventeen DOCX operations (sixteen public and one internal).

R2-135 extends `docx.text.set_character_format` without adding an operation. Font family is a
trimmed 1–128-character name, font size is 1–1638pt in half-point steps, color is `#RRGGBB`, and
highlight is one of the sixteen retained DOCX names; nullable values clear those properties.
Renderer validation supplements the serializable schema where string shape and half-step
constraints are required. Per-run merging preserves every unlisted attribute, and selected Ribbon
font/color/size/highlight actions use the same helper.

R2-136 reuses the exact DOCX text-range contract in
`docx.text.clear_character_format { range }`. The handler and retained Ribbon share one native
transaction helper for non-empty selections; it clears marks inside the range, keeps text and
outside marks unchanged, and remains one Undo unit. The generated Manifest contains 142 operations,
including eighteen DOCX operations (seventeen public and one internal).

R2-137 adds `docx.text.transform_case { range, mode }`. Its four explicit modes replace the
retained Change Case menu, preserve source marks, and use transaction mapping for Unicode
conversions whose length changes. Ribbon and Agent call the same helper and share one native Undo
unit. The generated Manifest contains 143 operations, including nineteen DOCX operations (eighteen
public and one internal).

R2-138 adds `docx.text.set_character_style { range, styleId }` and closes exact character
formatting. A non-null ID must be a character style in the mounted document's injected style
catalog; `null` removes the style mark with retained Ribbon behavior. Application preserves direct
run formatting. Actual style cards share this helper, while built-in fallback cards map to the
existing explicit character-format operation. The generated Manifest contains 144 operations,
including twenty DOCX operations (nineteen public and one internal).

R2-139 adds `docx.paragraph.set_direction { target, direction }`. Direction is an explicit `ltr` or
`rtl` final state over the reusable DOCX block target. Registry and retained UI import the same
logical-alignment flip rule, so `bidi` and explicit left/right alignment cannot drift. One native
transaction changes all matches. The generated Manifest contains 145 operations, including
twenty-one DOCX operations (twenty public and one internal).

R2-140 adds `docx.list.set_level { target, level }`. The operation assigns an absolute bounded
`0..8` final list level, while retained Ribbon indent buttons remain relative UI gestures over the
same normalization rule. Non-list target matches are unchanged and all changed items share one
native Undo transaction. The generated Manifest contains 146 operations, including twenty-two
DOCX operations (twenty-one public and one internal).

R2-141 adds `docx.list.apply_preset { target, levels }`. `levels` contains 1–9 bounded numbering
levels and creates a new document numbering definition through the retained numbering service.
Eligible target blocks receive the returned `numId` in one native Undo transaction; validation
failure occurs before definition creation. The generated Manifest contains 147 operations,
including twenty-three DOCX operations (twenty-two public and one internal).

R2-142 adds `docx.list.restart { blockIndex, start }`. The stable block anchor and explicit bounded
start value are revision-scoped inputs; the renderer clones the source numbering definition and
rewrites the anchor plus later same-list items in one native Undo transaction. The generated
Manifest contains 148 operations, including twenty-four DOCX operations (twenty-three public and
one internal).

R2-154 adds `docx.table.set_cell_format` for masked fill and vertical-alignment final state over one
exact bounded logical rectangle. Unlisted cell attrs remain unchanged; Agent and Ribbon target
adapters share the same write kernel and native Undo transaction. The generated Manifest contains
158 operations, including thirty-four DOCX operations (thirty-three public and one internal).

R2-155 adds `docx.table.set_cell_borders` for all/outer/inner/none edge policies over one exact
cell rectangle. Color/width and the null clear state are bounded, while Registry and Ribbon share
the same geometry/write kernel and Undo transaction. The generated Manifest contains 159
operations, including thirty-five DOCX operations (thirty-four public and one internal).

R2-156 adds `docx.table.set_style { tableBlockIndex, styleId }`. Non-null style identity is
validated against current-document table styles; `null` clears it. Registry and Ribbon share one
native table-attribute kernel and Undo route. The generated Manifest contains 160 operations,
including thirty-six DOCX operations (thirty-five public and one internal).

R2-157 adds `docx.table.set_row_height { tableBlockIndex, rowIndex, count, heightTwips }`. One
bounded physical-row interval receives a nullable 1–31,680 twip final state through the shared
Ribbon/Registry kernel and native Undo route. The generated Manifest contains 161 operations,
including thirty-seven DOCX operations (thirty-six public and one internal).

R2-158 adds `docx.table.set_column_widths { tableBlockIndex, widthsPx }`. A bounded complete vector
matching current logical width synchronizes every spanning cell plus table pixel/percentage width
through the same kernel used after Ribbon fitting. The generated Manifest contains 162 operations,
including thirty-eight DOCX operations (thirty-seven public and one internal).

R2-159 adds `docx.document.insert_page_break { afterBlockIndex }`. `-1` means document start and
the greatest valid value is the current last top-level block. One native empty
`docParagraph { pageBreakBefore: true }` is inserted after that boundary as a single Undo unit;
out-of-range boundaries fail without dispatch. Registry and Ribbon selection adapters share the
same page-break writer. The generated Manifest contains 163 operations, including thirty-nine DOCX
operations (thirty-eight public and one internal).

R2-160 adds `docx.section.insert_break { afterBlockIndex, startType }`. `startType` is exactly one
of `nextPage|continuous|evenPage|oddPage`. The new protected section-break node carries the copied
current-section `sectPr` and desired following-section type in the native history. During save,
each pending type is applied to the next section terminator; the last uses the hidden trailing
`sectPr`. This makes consecutive insertions replayable and lets Undo remove the whole pending
mutation. The generated Manifest contains 164 operations, including forty DOCX operations
(thirty-nine public and one internal).

R2-161 adds `docx.section.set_orientation { sectionIndex, orientation }`. The indexed current
section must exist; orientation is `portrait|landscape`, and a real transition swaps width/height
without changing margins, columns, borders, or other section fields. A complete section settings
override lives on the section anchor block in native history and is projected to the appropriate
`sectPr` during save. The generated Manifest contains 165 operations, including forty-one DOCX
operations (forty public and one internal).

R2-162 adds `docx.section.set_margins { sectionIndex, margins }`. `topTwips`, `rightTwips`,
`bottomTwips`, and `leftTwips` are required integers in `0..31680`; their axis sums must remain
strictly smaller than the current page dimensions. The complete final settings snapshot shares the
Ribbon journal and save projection. The generated Manifest contains 166 operations, including
forty-two DOCX operations (forty-one public and one internal).

R2-163 adds `docx.section.set_page_size { sectionIndex, widthTwips, heightTwips }`. Both axes are
required integers in `1440..31680` and must remain larger than the current corresponding margin
sums. Width greater than height yields landscape, height greater than width yields portrait, and a
square preserves current orientation. The generated Manifest contains 167 operations, including
forty-three DOCX operations (forty-two public and one internal).

R2-164 adds `docx.section.set_columns { sectionIndex, count, spacingTwips }`. Count is an integer
in `1..16`, spacing is an integer in `0..31680`, and `(count - 1) * spacingTwips` must be smaller
than the current text width. Both `w:num` and `w:space` round-trip exactly through the retained
engine. The generated Manifest contains 168 operations, including forty-four DOCX operations
(forty-three public and one internal).

R2-165 adds `docx.section.set_page_border { sectionIndex, enabled }` as an explicit boolean final
state over the same Undo-owned complete section snapshot and `w:pgBorders` save projection. The
generated Manifest contains 169 operations, including forty-five DOCX operations (forty-four
public and one internal). This closes the retained page/section mutation family.

R2-166 adds `docx.section.set_different_first_page { sectionIndex, enabled }`. The exact boolean is
Undo-owned on the indexed section anchor and maps to presence/absence of `w:titlePg` in that
section's terminating properties. The generated Manifest contains 170 operations, including
forty-six DOCX operations (forty-five public and one internal).

R2-167 adds `docx.document.set_different_odd_even_pages { enabled }`. The exact document-wide
boolean is Undo-owned on the first section anchor and maps to presence/absence of
`settings.xml/w:evenAndOddHeaders`. The generated Manifest contains 171 operations, including
forty-seven DOCX operations (forty-six public and one internal).

R2-168 adds `docx.section.set_page_numbering { sectionIndex, format, start }`. `format` is one of
the seven retained dialog values; `start` is `null` or an integer in `0..999999`. Both fields are
Undo-owned on the indexed section anchor and map to `w:pgNumType` in that section's terminating
properties. The generated Manifest contains 172 operations, including forty-eight DOCX operations
(forty-seven public and one internal).

R2-169 adds `docx.header_footer.set_text { sectionIndex, kind, variant, text }`. Kind is
`header|footer`, variant is `default|first|even`, and the renderer caps text at 65,536 characters.
The full content value is Undo-owned on the indexed section anchor; save maps it to a variant-aware
non-final section part or the existing final-section writer. The generated Manifest contains 173
operations, including forty-nine DOCX operations (forty-eight public and one internal).

R2-170 adds
`docx.header_footer.set_page_number { sectionIndex, kind, variant, enabled, alignment }`.
Alignment is `left|center|right`; enabled writes one canonical PAGE-field paragraph, while disabled
removes PAGE fields without discarding other content. The generated Manifest contains 174
operations, including fifty DOCX operations (forty-nine public and one internal).

R2-171 adds `docx.header_footer.set_paragraphs { sectionIndex, kind, variant, paragraphs }`.
Paragraph arrays are bounded to 1–64 and segment arrays to 0–256. Segments are typed text, PAGE,
or NUMPAGES values with bounded run styling; renderer validation enforces aggregate text and lexical
constraints not expressible in the current base Schema. The generated Manifest contains 175
operations, including fifty-one DOCX operations (fifty public and one internal), closing the
retained header/footer family.

R2-172 adds public
`docx.image.insert { path, afterBlockIndex, widthPx, heightPx, alignment }` and internal
`docx.image.insert_staged`. Direct Agent calls contain only an absolute local path and bounded,
explicit layout state. The Broker stages at most 20 MiB of PNG/JPEG/GIF data under the session,
queues only its opaque descriptor, and releases it after acknowledgement. The mounted host hydrates
the ArrayBuffer; renderer validation repeats size/extension/magic checks before inserting one
Undo-owned native image node at the exact top-level boundary. Retained dialog/paste insertion uses
the same node builder. The generated Manifest contains 177 operations, including fifty-three DOCX
operations (fifty-one public and two internal).

R2-173 adds public `docx.image.replace { path, imageBlockIndex, widthPx, heightPx }` and internal
`docx.image.replace_staged`. The explicit top-level target and final geometry replace selection and
natural-size context. Broker staging and host hydration reuse the insertion transport, while the
renderer again validates byte length and media magic. A single replacement-attrs kernel selects
original `imageReplace` or pending `genImage`, preserves placement/wrap/format state, and clears old
crop windows. The retained UI and Agent share it with native Undo and save/reopen evidence. The
generated Manifest contains 179 operations, including fifty-five DOCX operations (fifty-two public
and three internal).

R2-174 adds `docx.image.set_wrap { imageBlockIndex, wrap }`. The public enum contains only retained
UI states plus explicit `null` inline state. Renderer validation targets one exact image; inline
clears both margin-relative and offset positioning, while other wrap values retain position.
ContextMenu, Layout, Picture Format, and Agent dispatch share the same attrs writer, one native Undo
unit, and the existing save/reopen projection. The generated Manifest contains 180 operations,
including fifty-six DOCX operations (fifty-three public and three internal).

R2-175 adds `docx.image.set_margin_position { imageBlockIndex, horizontal, vertical }`. The two
finite axes select one of nine retained Word presets; the renderer derives square wrap, stores both
named positions, clears offset coordinates, and rejects unsaved generated images that cannot yet
persist this state. Layout UI and Agent execution share one exact-index writer, native Undo, and
save/reopen/resave evidence. The generated Manifest contains 181 operations, including fifty-seven
DOCX operations (fifty-four public and three internal).

R2-176 adds
`docx.image.set_offset_position { imageBlockIndex, wrap, offsetXEmu, offsetYEmu }`. The operation
uses the complete retained non-null image-wrap enum and bounded signed 32-bit EMU axes. It clears
named positioning, targets original or generated images, and commits one native Undo unit.
Registry execution and retained image dragging share the same pure final-state projection;
textbox/shape drag remains a separate object concern. Original-image save/reopen is covered. The
generated Manifest contains 182 operations, including fifty-eight DOCX operations (fifty-five
public and three internal).

R2-177 adds
`docx.image.set_transform { imageBlockIndex, rotationDegrees, flipHorizontal, flipVertical }`.
The required final state replaces Ribbon-only deltas/toggles; rotation is an integer `0..359`, and
both mirror axes are booleans. Registry and retained controls share one canonical attrs writer and
native Undo. Saving materializes a minimal `pic:spPr/a:xfrm` when absent so a valid image cannot
silently discard the transform. Original/generated state and original-image save/reopen are
covered. The generated Manifest contains 183 operations, including fifty-nine DOCX operations
(fifty-six public and three internal).

R2-178 adds `docx.image.set_crop { imageBlockIndex, left, top, right, bottom }`. The four source
insets are bounded `0..0.99` numbers, their opposing pairs must leave positive area, and all zeros
mean explicit reset. Registry and retained Crop UI share a non-destructive attrs writer that clears
stale fill windows without replacing source bytes. Original and new image save routes write
`a:srcRect`; native Undo and save/reopen are covered. The generated Manifest contains 184
operations, including sixty DOCX operations (fifty-seven public and three internal).

R2-179 adds `docx.image.remove { imageBlockIndex }`. Only an exact top-level image is accepted;
deleting the sole block inserts an empty paragraph in the same native Undo transaction. Save/reopen
deletion is covered. The generated Manifest contains 185 operations, including sixty-one DOCX
operations (fifty-eight public and three internal), closing the retained image lifecycle.

R2-180 adds `docx.shape.insert { afterBlockIndex, preset, widthEmu, heightEmu }`. The preset enum
is projected from all 104 retained filled Gallery entries; five line/connector entries remain
outside this fill-bearing operation. Stable top-level insertion and both bounded EMU axes are
validated before Registry, Gallery, and draw mode converge on one node builder and native Undo.
Save/reopen is covered. The generated Manifest contains 186 operations, including sixty-two DOCX
operations (fifty-nine public and three internal).

R2-181 adds `docx.line.insert { afterBlockIndex, kind, widthEmu, heightEmu }` for five finite
stroke-only line/connector kinds. Straight kinds accept only the canonical 114,300 EMU grab
height; bent/curved connectors retain two bounded axes. Registry and retained Gallery/draw mode
share one node builder and native Undo. Save/reopen is covered. The generated Manifest contains
187 operations, including sixty-three DOCX operations (sixty public and three internal).

R2-182 adds `docx.textbox.insert { afterBlockIndex, widthEmu, heightEmu }`. The stable top-level
boundary and both 9,525–20,000,000 EMU axes are validated before the retained Ribbon and Registry
converge on one format-owned textbox node builder. Native Undo and save/reopen are covered. The
generated Manifest contains 188 operations, including sixty-four DOCX operations (sixty-one public
and three internal).

R2-183 adds bounded `docx.chart.insert` for bar, line, and pie data matrices. It carries explicit
title, category and series values, plus a 120–660 px by 80–4,096 px final extent. Cross-field
validation enforces equal category/value dimensions, one pie series, finite ±10^12 values, and a
4,096-value total budget. Registry and retained Chart dialog share one node builder and native Undo;
the embedded chart part and workbook reopen with the requested data and extent. The generated
Manifest contains 189 operations, including sixty-five DOCX operations (sixty-two public and three
internal).

R2-184 extends the supported bounded Schema subset with Unicode-aware `minLength` and `maxLength`,
then adds `docx.equation.insert { placement, latex, afterBlockIndex, from, to }`. Block placement
requires a stable boundary plus null range; inline placement requires a null boundary plus one
same-paragraph exact range. LaTeX is bounded to 1–4,096 characters and parsed before mutation.
Gallery, modal, and Registry share LaTeX-to-OMML construction and native Undo; both display and
inline save/reopen routes retain valid OMML. The generated Manifest contains 190 operations,
including sixty-six DOCX operations (sixty-three public and three internal).

R2-185 adds `docx.object.set_size { objectBlockIndex, widthPx, heightPx }` as one bounded aggregate
operation for shape, line, textbox/WordArt-like textbox, and chart corner resize. The renderer
enforces the narrower chart domain and fixed straight-line grab height after exact block
resolution. Registry and retained drag UI share one transaction kernel; changed sizes start a new
native history group and save/reopen through the existing drawing projection. The generated
Manifest contains 191 operations, including sixty-seven DOCX operations (sixty-four public and
three internal).

R2-186 adds `docx.object.set_offset_position { objectBlockIndex, wrap, offsetXEmu, offsetYEmu }`
for shape, line, textbox/WordArt-like textbox, and diagram nodes. Its finite wrap and signed 32-bit
EMU axes describe a replayable final anchor. Shape draw, move handles, and Registry share one
kernel; numeric-offset OOXML encodes side direction through `wrapText`, retains tight/through
polygons, and reopens negative offsets. The generated Manifest contains 192 operations, including
sixty-eight DOCX operations (sixty-five public and three internal).

R2-187 adds `docx.object.set_style { objectBlockIndex, style, fields }` for shape, line,
textbox/WordArt-like textbox, and diagram nodes. Its unique mask must exactly match supplied nullable
`fillHex`/`borderHex` final-state properties; string values match uppercase six-digit hex through the
Operation Contract `pattern` keyword, and lines reject fill. Shape Format and Registry share one
native-history kernel, and save/reopen preserves both set and removed properties. The generated
Manifest contains 193 operations, including sixty-nine DOCX operations (sixty-six public and three
internal).

R2-188 adds `docx.object.remove { objectBlockIndex }` for exact shape, line,
textbox/WordArt-like textbox, chart, diagram, and block-equation nodes. The operation rejects image
and generic protected nodes, reports whether it inserted the required replacement paragraph, and
uses one position-level transaction shared with real object-mode Backspace/Delete. Native Undo and
save/reopen are covered. The generated Manifest contains 194 operations, including seventy DOCX
operations (sixty-seven public and three internal).

R2-189 adds `docx.chart.update { chartBlockIndex, patch, fields }`. Its exact unique mask targets
existing title, category, series-name, and numeric cache slots without changing chart kind or matrix
shape. Text is bounded to 512 characters; values are finite ±10^12 numbers; absent series names and
null cache gaps remain read-only. The protected chart grid and Registry share one native-history
kernel. Original chart parts receive surgical cache patches; generated chart/workbook parts consume
the final display matrix. The generated Manifest contains 195 operations, including seventy-one
DOCX operations (sixty-eight public and three internal).

R2-190 adds
`docx.equation.update { placement, mode, latex, tokens, equationBlockIndex, from, to }`. LaTeX mode
requires non-null bounded `latex`, null `tokens`, and either a stable block index or exact inline
atom range. Token mode requires null `latex`, a bounded same-shape token array, and a retained block
target; total token content is capped at 4096 Unicode characters. Both modes use the same
native-history writer as the corresponding retained UI editor and persist through the existing
OMML save path. The generated Manifest contains 196 operations, including seventy-two DOCX
operations (sixty-nine public and three internal).

R2-191 adds
`docx.textbox.set_content { objectBlockIndex, textboxIndex, paragraphs, heightPx }`. Public JSON
expresses bounded rich runs and paragraph alignment/geometry; aggregate runtime limits cap total
runs at 4096 and text at 65,536 Unicode characters. `heightPx: null` preserves current fixed/autofit
state, while a number records an explicit final rendered height. Flattened complex boxes are
read-only. The exact same batch-capable native transaction serves Registry and retained nested
editors, including multi-box commits. The generated Manifest contains 197 operations, including
seventy-three DOCX operations (seventy public and three internal).

R2-192 adds `docx.text.set_link { range, href, text }`. A non-null `text` replaces the exact range
before applying a non-null href; null text preserves a non-empty text-only range for href update or
removal; null href requires null text. All strings are Unicode-bounded and coordinates remain in
one text-bearing block. LinkInsertModal and Registry use the same native-history action, and the
save path allocates/reopens external relationships. The older `docx.text.set_style` descriptor no
longer accepts its whole-block link field. The generated Manifest contains 198 operations,
including seventy-four DOCX operations (seventy-one public and three internal).

R2-193 adds `docx.bookmark.set { blockIndex, name, enabled }`. The target is one revision-scoped
top-level paragraph, heading, or list item; the unique name is pattern-checked and Unicode-bounded
to 40 characters. Explicit enabled state is idempotent. The retained Bookmark dialog and Registry
share one native-history action, while OOXML save/reopen emits the same bookmark pair. The generated
Manifest contains 199 operations, including seventy-five DOCX operations (seventy-two public and
three internal).

R2-194 adds `docx.cross_reference.insert { range, bookmarkName, displayText }`. The bounded exact
range stays in one text-bearing block, the named user bookmark must exist, and cached display text is
explicit. The retained dialog and Registry share one native-history REF-field insertion path. The
generated Manifest contains 200 operations, including seventy-six DOCX operations (seventy-three
public and three internal).

R2-195 adds `docx.field.insert { range, instruction, displayText }`. The instruction enum contains
DATE, TIME, PAGE, NUMPAGES, and FILENAME; the range stays inside one text-bearing block, and cached
display text is explicit. App/Ribbon and Registry share one native-history generic-field insertion
path. The generated Manifest contains 201 operations, including seventy-seven DOCX operations
(seventy-four public and three internal).

R2-196 adds `docx.field.update { updates }`. Each bounded item identifies an exact current field mark
by range plus full original instruction and provides explicit final cached text. Runtime accepts the
seven retained F9 keywords, including instructions with Word switches, and applies all validated
changes in one native-history transaction. The generated Manifest contains 202 operations,
including seventy-eight DOCX operations (seventy-five public and three internal).

R2-197 bounds existing `docx.text.insert.text` to 1–65,536 Unicode characters and routes both
Registry insertion and the retained symbol palette through one format-owned native-history action.
It intentionally adds no operation; the generated Manifest remains at 202 operations including
seventy-eight DOCX operations (seventy-five public and three internal).

R2-198 adds `docx.toc.refresh { tocBlockIndex, entries }`. The stable top-level block must begin one
TOC field, and bounded explicit entries replace the exact matched fldChar region while preserving a
tail page break. Retained UI and Registry share one native-history action. The generated Manifest
contains 203 operations, including seventy-nine DOCX operations (seventy-six public and three
internal).

R2-199 adds `docx.note.insert { range, kind, noteId, text }`. Its explicit positive ID and exact
inline target make insertion replay-safe; duplicate IDs and cross-block targets fail before
mutation. The bounded body is retained on the native-history reference atom, so Undo/Redo drives
both document marker and shared note-part membership. The generated Manifest contains 204
operations, including eighty DOCX operations (seventy-seven public and three internal).

R2-200 adds `docx.note.update { kind, noteId, text }`. A stable positive ID must resolve in the note
part and document references before mutation. The bounded body is written to the Undo-owned atoms,
and the mounted session preserves the original note as the baseline for exact Undo/Redo
reconciliation. The generated Manifest contains 205 operations, including eighty-one DOCX
operations (seventy-eight public and three internal).

R2-201 adds `docx.note.delete { kind, noteId }`. It removes every reference for one resolved stable
identity and deterministically renumbers same-kind peers in one native-history transaction. A
history-owned deletion flag lets note-part reconciliation distinguish deletion from an undone body
update. The generated Manifest contains 206 operations, including eighty-two DOCX operations
(seventy-nine public and three internal).

R2-202 adds `docx.source.upsert { source }`. Stable tag and finite source type select a bounded
create/update, while a first-block final Sources override makes the customXml metadata participate
in native Undo/Redo. The generated Manifest contains 207 operations, including eighty-three DOCX
operations (eighty public and three internal).

R2-203 adds `docx.citation.insert { range, sourceTag, displayText }`. The stable tag must exist in
the current source snapshot, and explicit bounded display text is inserted at one exact inline
range through native history. The generated Manifest contains 208 operations, including
eighty-four DOCX operations (eighty-one public and three internal).

R2-204 adds `docx.bibliography.insert { afterBlockIndex, heading, entries }`. A revision-scoped
top-level boundary and source-backed bounded final display lines replace selection-dependent
generation. The heading and all entry paragraphs insert in one native-history transaction. The
generated Manifest contains 209 operations, including eighty-five DOCX operations (eighty-two
public and three internal).

R2-205 adds `docx.caption.insert { afterBlockIndex, label, number, text }`. Bounded explicit final
caption values and a stable top-level boundary create one dirty SEQ field in a native-history
transaction. The generated Manifest contains 210 operations, including eighty-six DOCX operations
(eighty-three public and three internal).

R2-206 adds `docx.index.mark { range, term }`. An exact same-block inline range anchors one bounded
hidden XE marker; terms containing writer-normalized quotes or controls are rejected. The generated
Manifest contains 211 operations, including eighty-seven DOCX operations (eighty-four public and
three internal).

R2-207 adds `docx.index.insert { afterBlockIndex, label, terms }`. A bounded term snapshot becomes a
deduplicated sorted dirty INDEX cache inserted after one stable top-level boundary in one native
transaction. The generated Manifest contains 212 operations, including eighty-eight DOCX
operations (eighty-five public and three internal). Retained source UI audit found no deletion path.

R2-208 adds `docx.comment.add { range, comment }`. Explicit stable metadata and exact-range anchor
marks are committed with an Undo-owned final comments snapshot in one native transaction. The
generated Manifest contains 213 operations, including eighty-nine DOCX operations (eighty-six
public and three internal).

R2-209 adds `docx.comment.reply { parentId, comment }`. A stable parent anchor and explicit bounded
reply metadata are committed with the final comment snapshot in one native transaction. The
generated Manifest contains 214 operations, including ninety DOCX operations (eighty-seven public
and three internal).

R2-210 adds `docx.comment.set_resolved { id, resolved }`. A stable top-level thread ID and final
boolean update parent and replies in one Undo-owned metadata snapshot. The generated Manifest
contains 215 operations, including ninety-one DOCX operations (eighty-eight public and three internal).

R2-211 adds `docx.comment.delete { id }`. Stable top-level deletion cascades direct replies while
reply deletion remains local; affected anchors and the final comments snapshot share one native
transaction. The generated Manifest contains 216 operations, including ninety-two DOCX operations
(eighty-nine public and three internal), closing the retained comment lifecycle.

R2-212 adds `docx.revision.set_tracking { enabled }`. The explicit final state is applied directly
to the mounted recorder and synchronized to retained Ribbon state. It is a non-undoable,
recovery-free session policy; subsequent tracked document edits retain native Undo and OOXML
persistence. The generated Manifest contains 217 operations, including ninety-three DOCX
operations (ninety public and three internal).

R2-213 adds `docx.revision.apply_decision { decision, scope }`. Bounded accept/reject and
current/all enums cover the four retained decisions through one native transaction. The generated
Manifest contains 218 operations, including ninety-four DOCX operations (ninety-one public and
three internal), closing the retained revisions family.

R2-214 adds `docx.document.set_protection { enabled, password }`. Passwords are bounded input-only
values; protected removal verifies the current credential and results expose only boolean state.
The operation uses the renderer-owned save journal, is explicitly non-undoable, and has recovery
plus settings-part save/reopen evidence. The generated Manifest contains 219 operations, including
ninety-five DOCX operations (ninety-two public and three internal).

R2-215 adds `docx.ink.apply { action, annotation?, ids? }`. Finite add/delete/clear action shapes,
bounded vector geometry, and stable IDs provide one aggregate renderer-owned ink lifecycle route;
arbitrary image payload injection is not exposed. The generated Manifest contains 220 operations,
including ninety-six DOCX operations (ninety-three public and three internal), closing Ink.

R2-216 adds public `docx.document.compare { path }` and hidden
`docx.document.compare_staged { blobId, name, size, data }`. The Broker alone resolves the local
path and stages bounded DOCX bytes; the renderer alone parses the comparison package and commits
the deterministic paragraph-diff panel state shared with retained UI. The operation is
non-undoable and recovery-free, reports explicit counts/identity, and leaves document dirty state
unchanged. The generated Manifest contains 222 operations, including ninety-eight DOCX operations
(ninety-four public and four internal), closing Compare.

R2-217 adds `docx.document.set_design { fields, pageColor?, watermark?, themeFonts?,
themeColors? }`. The unique bounded field mask must exactly match supplied values, providing one
atomic final-state contract for all retained document design mutations without encoding Ribbon
presets. Renderer and Broker validate the same finite shape; retained UI and Agent share the dirty,
recovery-backed, non-undoable save journal. The generated Manifest contains 223 operations,
including ninety-nine DOCX operations (ninety-five public and four internal), closing Document
design.

R2-218 adds `docx.cover_page.insert { preset, title, subtitle, author, date, year }`. A finite
twelve-value style enum is deliberately separated from explicit bounded visible content, so Agent
replay never depends on renderer locale or wall-clock time. Retained UI and Registry share one
start-of-document native transaction with Undo and save/reopen evidence. The generated Manifest
contains 224 operations, including 100 DOCX operations (ninety-six public and four internal),
closing Cover pages.

R2-219 adds `docx.paragraph.set_drop_cap { blockIndex, mode, lines }`. An explicit top-level block
identity and finite none/drop/margin state replace implicit selection; null/2–10 line constraints
are cross-validated before the native paragraph transaction. Retained Ribbon and Registry share
no-op, Undo, and save/reopen semantics. The generated Manifest contains 225 operations, including
101 DOCX operations (ninety-seven public and four internal), closing Drop caps.

R2-220 adds `docx.wordart.insert { afterBlockIndex, preset, text, widthEmu, heightEmu, drawingId }`.
The finite preset, bounded visible/geometry data, explicit boundary, and unique OOXML drawing ID
make insertion replayable. Retained UI allocates the lowest free identity and calls the same native
transaction. The generated Manifest contains 226 operations, including 102 DOCX operations
(ninety-eight public and four internal), closing the DOCX retained-command inventory while release
readiness remains governed by the shared gates.

R2-143 adds `docx.list.continue { blockIndex, previousBlockIndex }`. Both explicit top-level block
identities are validated before the numbering action rebinds the current list tail to the selected
earlier list identity in one native Undo transaction. The generated Manifest contains 149
operations, including twenty-five DOCX operations (twenty-four public and one internal).

R2-145 extends `docx.paragraph.set_style` with a nullable, strictly ordered, 1–64-item `tabStops`
array. Positions and all categorical values are bounded before the shared adapter writes the
renderer-owned serialized attribute. The generated Manifest remains at 149 operations.

R2-146 adds `docx.table.insert { afterBlockIndex, rows, columns }`. The operation targets a stable
top-level block boundary, validates both dimensions and total cell count, and inserts the retained
native table model in one Undo transaction. The generated Manifest contains 150 operations,
including twenty-six DOCX operations (twenty-five public and one internal).

R2-147 adds `docx.table.delete { tableBlockIndex }`. Only an exact top-level native table is
accepted; deletion and the sole-table empty-paragraph fallback are one native Undo transaction.
The generated Manifest contains 151 operations, including twenty-seven DOCX operations
(twenty-six public and one internal).

R2-148 adds `docx.table.insert_rows { tableBlockIndex, rowIndex, count }`. Boundary/count and
resulting-cell budget validation precede one rowspan-aware native transaction. The generated
Manifest contains 152 operations, including twenty-eight DOCX operations (twenty-seven public and
one internal).

R2-149 adds `docx.table.delete_rows { tableBlockIndex, rowIndex, count }`. The explicit interval
must be in range and leave at least one row; full-table deletion is a separate operation. Repeated
rowspan-aware removal remains one native transaction and Undo entry. The generated Manifest
contains 153 operations, including twenty-nine DOCX operations (twenty-eight public and one
internal).

R2-150 adds `docx.table.insert_columns { tableBlockIndex, columnIndex, count }`. Boundary, maximum
resulting width, and resulting-cell validation precede one colspan-aware native transaction. The
generated Manifest contains 154 operations, including thirty DOCX operations (twenty-nine public
and one internal).

R2-151 adds `docx.table.delete_columns { tableBlockIndex, columnIndex, count }`. The explicit
interval must be in range and retain at least one column; repeated colspan-aware removal remains
one native transaction and Undo entry. The generated Manifest contains 155 operations, including
thirty-one DOCX operations (thirty public and one internal).

R2-152 adds `docx.table.merge_cells` with one bounded, exact half-open logical rectangle. The
renderer rejects rectangles outside the current grid or crossed by existing spans, synthesizes the
native CellSelection, and invokes the retained merge command in one Undo transaction. The shared
row schema permits rows wholly covered by rowspan. The generated Manifest contains 156 operations,
including thirty-two DOCX operations (thirty-one public and one internal).

R2-153 adds `docx.table.split_cell { tableBlockIndex, rowIndex, columnIndex }`. One bounded logical
coordinate resolves through the current TableMap to a merged cell; ordinary cells fail before
mutation. The retained native command restores the original span with one Undo. The generated
Manifest contains 157 operations, including thirty-three DOCX operations (thirty-two public and
one internal).

R2-144 extends the existing `docx.paragraph.set_style` schema with bounded `lineRule` and
`lineRawTwips` final state and applies upper/lower bounds to paragraph spacing and indentation.
This is a contract deepening rather than a new operation, so the generated Manifest remains at 149
operations, including twenty-five DOCX operations (twenty-four public and one internal).

PDF-P1 through PDF-P4 established the Registry seam for save, staged load, exact saved-annotation
deletion, and mounted history. The completed PDF parity slice expands that same seam to 25
descriptors: 23 Agent-visible operations and internal document/page staged-byte routes. Exact
bounded schemas cover metadata, undo/redo, markup, pending deletion, drawing/note/signature state,
text insertion/replacement/update, image lifecycle, static and AcroForms, generated stamps, and page
insert/delete/reorder/rotation. The machine-checked retained producer baseline has no missing entry.
Browser PDFium provides searchable content-stream text/image persistence; PDF-lib-safe families use
the retained save assembler. Legacy `save`, `open_local_file`, `delete_saved_annotation`, and
`undo` names remain input-only aliases. The generated product manifest owns discovery and Broker
validation while keeping staged operations internal.

PPTX-P1/P2 plus R2-239 through R2-253 migrate the initial executable PPTX browser slice into its
format-owned registry. Nineteen descriptors cover document creation/load/save/save-as,
selection, native undo/redo, slide lifecycle, explicit object deletion/EMU transform, selected
movement, rich text, font/paragraph formatting, and scoped find/replace. Internal staged load
validates hydrated bytes and remains hidden. Legacy `save`, `open_local_file`, `select_objects`,
`replace_selected_text`, and `move_selected_objects` are input-only aliases. Generated discovery,
normalization, validation and visibility replace central schemas and direct host dispatch.

R2-254 through R2-308 complete PPTX retained-command parity. The catalog now owns 74 descriptors:
73 Agent-visible operations plus internal staged load. Exact bounded contracts cover every
retained slide/layout, object, text, group-child, connector, picture, table, chart, SmartArt,
theme, animation, notes/comments/sections/hyperlink, embedded-resource, and master/layout mutation
family. The retained `SlidesApi` is compile-time complete, user picker paths reuse the same byte
primitives, native snapshot history and package persistence remain authoritative, and a
machine-checked producer baseline has no missing descriptor. Export, print, and presentation-show
routes are explicitly non-mutating host effects rather than Agent document operations.

R6-07 replaces the public single-operation input with a revision-guarded transaction envelope and
session-scoped idempotency journal. The renderer transport still receives one command because no
current format declares a multi-operation native atomic/undo grouping. Exact completed and
in-flight replays converge without redispatch. R6-08 removes the old input Adapter after migrating
all repository callers and tests; public execution now accepts only the complete transaction
envelope and exact canonical Agent-visible ids. R6-09 removes the remaining thirty-six public-era
aliases from the format registries and generated Product Manifest. A cross-format drift test now
requires the Manifest's complete alias set to be exactly five internal `open_local_file` aliases.
Format-local retained-command parity is closed for all five formats, and approved R6-01 evidence
closes the shared cross-format release gates. Readiness is generated from that evidence rather than
inferred from registry completeness alone.

## Module map

```text
Pinned community renderer
        │ source inventory
        ▼
Baseline capability inventory
        │ required mapping and evidence
        ▼
Format-owned executable operation registry
        ├── serializable descriptors
        ├── renderer handlers
        └── descriptor/handler/test checks
        │
        ├── generated MCP capability manifest
        ├── generated Skill operation reference
        └── generated parity report
        │
Agent ──┴─ office_get_capabilities / office_get_context / office_execute
                  │
                  ▼
          live-session broker
                  │ queued transaction
                  ▼
          mounted renderer Adapter
                  │
                  ▼
       native command/state/undo seam
                  │
                  ▼
          acknowledge revision + 1
```

The public MCP Interface is intentionally smaller than the operation catalog. Renderer UI placement is intentionally larger than the operation catalog. A Ribbon button, shortcut, context action, and dialog submit action may all map to the same semantic operation.

## Registry ownership

Each format owns its descriptors, schemas, target semantics, handlers, and focused tests. The planned locations are:

```text
apps/docs/src/renderer/operations/
apps/sheets/src/renderer/operations/
apps/slides/src/renderer/operations/
apps/pdf/src/renderer/operations/
apps/markdown/src/renderer/operations/
```

The final path may follow an existing renderer's naming conventions, but ownership may not move into `apps/mcp-server` or a cross-format document model.

A small shared Module may live under `packages/operation-contract`. Its Interface is limited to common descriptor types, transaction envelopes, validation results, standard errors, compatibility aliases, and manifest generation. Format-specific targets and schemas stay inside the owning format.

## Operation descriptor

The serializable descriptor has the following conceptual shape:

```ts
interface OperationDescriptor {
  id: `${Format}.${string}`
  format: 'docx' | 'xlsx' | 'pptx' | 'pdf' | 'markdown'
  family: string
  summary: string
  visibility: 'agent' | 'internal'
  inputSchema: JsonSchema
  outputSchema: JsonSchema
  risk: 'low' | 'medium' | 'high'
  context: Array<'document' | 'selection' | `${Format}.${string}`>
  effects: Array<'selection' | 'document' | 'persistence' | 'view'>
  mutates: boolean
  undoable: boolean
  atomic: boolean
  compatibilityAliases?: string[]
}
```

Normative rules:

- `id` is stable and fully qualified by format.
- `visibility` keeps staged-file and other renderer transport operations out of the Agent projection while still routing them through the format registry.
- `inputSchema` and `outputSchema` describe only that operation, not the full union for a format.
- `additionalProperties` is false unless a format has a documented preservation reason.
- `risk` describes semantic impact; it does not pretend that a preview exists.
- `context` declares what `office_get_context` must expose before execution. Format-specific requirements use format-owned namespaced values such as `xlsx.activeRange`; the common contract does not enumerate renderer targets.
- `effects` makes selection-only, document, persistence, and view behavior explicit for validation and evidence rules.
- `undoable` is true only when the operation enters the renderer's existing undo history.
- `atomic` is true only when the renderer can apply the operation without a partially visible intermediate state.
- aliases are transitional, resolve only inside the active session format, and require an explicit removal gate.

The executable registry binds every descriptor to exactly one renderer handler. A handler without a descriptor and a descriptor without a handler both fail the registry check.

## Baseline mapping

The baseline capability inventory is not the executable registry. It records every pinned command route with a disposition such as:

- `retained-state-change`;
- `retained-read-only`;
- `host-adapted`;
- `excluded-ai`;
- `excluded-account`;
- `excluded-electron`;
- `upstream-placeholder`.

Every `retained-state-change` entry must identify:

- source file and retained UI route;
- native command seam;
- one or more executable operation ids;
- selection/context requirements;
- user/Agent equivalence evidence;
- undo evidence when the UI command is undoable;
- save/reopen evidence when the command changes persisted document state.

The baseline inventory may contain incomplete mappings. The executable registry may contain only passing operations. Release readiness requires the baseline mapping to have no unexplained retained state-changing gaps.

## Capability discovery

R6-06 implements bounded Registry discovery without adding a public tool or resource.
`office_get_capabilities` supports schema-free summary pages and exact-detail lookup.

Summary example:

```json
{
  "format": "xlsx",
  "view": "summary",
  "family": "range",
  "cursor": null
}
```

Summary results contain ids, families, summaries, risk, context requirements, availability, and pagination metadata. They do not inline every operation schema.

Only `visibility: agent` descriptors appear in this projection. Internal descriptors, including the staged-file load command used by `office_open_local_file`, remain callable only by the broker/renderer transport path.

Detail example:

```json
{
  "format": "xlsx",
  "view": "detail",
  "operation": "xlsx.range.set_values"
}
```

Detail results contain the exact descriptor and schema. Supplying `sessionId` projects availability
against the active format, connection, and selection state. An unavailable operation returns the
stable reason `format_mismatch`, `editor_offline`, or `selection_required`; it is not silently
omitted. Internal descriptors and compatibility aliases are never discoverable. Summary pages are
limited to twenty operations and regression-gated below 8 KiB; every generated detail descriptor
is capped at 64 KiB.

The short `{ "format": "..." }` form is the canonical default-summary request. It no longer returns
the legacy complete operation-schema object.

## Operation transaction

The implemented R6-08 `office_execute` input is:

```json
{
  "sessionId": "<uuid>",
  "baseRevision": 4,
  "requestId": "<caller-generated-id>",
  "operations": [
    {
      "id": "xlsx.range.set_values",
      "arguments": {
        "sheet": "Sheet1",
        "range": "A1:C1",
        "values": [[1, 2, 3]]
      }
    }
  ]
}
```

Validation and replay order is:

1. MCP input validation requires the complete transaction envelope; the retired `operation` +
   `arguments` shape never enters the handler;
2. a multi-operation request is rejected as `transaction_not_atomic` before renderer dispatch
   because no current Registry declares an allowed grouping;
3. the session-scoped journal compares canonicalized structural JSON: an exact known request joins
   or returns its original result, while payload-changing reuse returns `request_reused`;
4. for a new request, the session exists, the renderer is connected, `baseRevision` equals the last
   acknowledgement, and no other command is pending or active;
5. the operation is an exact canonical Agent-visible id in the session format;
6. its arguments match the exact generated schema;
7. required live context is present;
8. the Broker stages any bounded local input and enqueues the single renderer command.

Object-key order is not transaction identity. Pre-enqueue failure releases the request id for a
corrected retry. Once enqueue succeeds, the journal retains the accepted result or failure for the
live Broker lifetime. The renderer Adapter validates again, invokes the native command seam,
checkpoints renderer-produced recovery bytes when required, and acknowledges the transaction.

Success advances revision exactly once:

```json
{
  "ok": true,
  "transaction": {
    "transactionId": "<uuid>",
    "requestId": "<caller-generated-id>",
    "baseRevision": 4
  },
  "result": {
    "revision": 5,
    "operations": [
      {
        "id": "xlsx.range.set_values",
        "result": { "changed": 3, "sheet": "Sheet1", "range": "A1:C1" }
      }
    ]
  }
}
```

An exact completed replay returns that same response and `transactionId`, even after revision has
advanced. An in-flight replay waits on the same promise and never creates a second renderer command.
A caller's 15-second `command_timeout` does not settle the accepted transaction: the underlying ACK
and staged-byte lifetime continue, and a later exact replay returns the final response.

A multi-operation transaction cannot partially acknowledge. Formats that cannot yet provide atomic
grouping accept only one operation. R6-07 therefore rejects every `operations.length > 1` request
before dispatch rather than approximating rollback outside the renderer.

The legacy `{ operation, arguments }` input and its `{ command, result }` response branch are
removed. The installed Skill, repository tests, internal callers, and packaged smoke use the
transaction form. Format-local renderer and staged-transport aliases do not create a public
execution path.

## Errors

The existing live-session errors remain. R6-07 implements stable validation distinctions where
callers can act differently:

| Code                       | Meaning                                                                                        |
| -------------------------- | ---------------------------------------------------------------------------------------------- |
| `operation_not_found`      | The fully qualified id is absent from the active format registry.                              |
| `operation_unavailable`    | The operation exists but current connection, selection, or document state does not satisfy it. |
| `operation_schema_invalid` | Arguments do not match the operation's exact schema.                                           |
| `transaction_not_atomic`   | The requested multi-operation combination cannot be one renderer transaction/undo unit.        |
| `request_reused`           | The request id was reused with a different payload.                                            |

These errors do not advance revision.

## Generated artifacts

The registry build produces or verifies:

```text
generated MCP capability manifest
generated installed-Skill operation reference
generated per-format executable-operation tables
generated baseline-to-operation parity report
descriptor ↔ handler ↔ schema ↔ test drift report
```

Generated output is never the authoring source. Format-owned registry definitions and baseline inventories remain authoritative.

## Compatibility and removal

The migration is a strangler replacement, not a second permanent command path.

| Transitional path                              | Compatibility treatment                                                                         | Removal gate                                                                           |
| ---------------------------------------------- | ----------------------------------------------------------------------------------------------- | -------------------------------------------------------------------------------------- |
| `office_execute.operation` + `arguments`       | Removed in R6-08 with the legacy response and public alias-resolution branch.                   | Complete: the schema requires the transaction form; tests and packaged smoke guard it. |
| Thirty-six public-era format aliases           | Removed from all five format registries and the generated Manifest in R6-09.                    | Complete: renderer rejection, Manifest-set, and packaged-smoke tests guard removal.    |
| Central `capabilities.ts` catalogs             | Compare against generated manifests during migration.                                           | Five format registries generate the complete server projection.                        |
| DOCX `batch_update`                            | Removed in R2-18 after all ten commands gained exact operations.                                | Complete: capability, App branch, and executor are absent.                             |
| XLSX `save`                                    | Canonical registry operation; legacy name removed in R6-09.                                     | Complete: direct branches and compatibility alias are absent.                          |
| XLSX `open_local_file`                         | Internal canonical registry operation; legacy name is transport-only.                           | Complete in R2-20: direct App branch is absent.                                        |
| PPTX `save`                                    | Canonical registry operation; legacy name removed in R6-09.                                     | Complete: direct branch, central schema, and compatibility alias are absent.           |
| PPTX `open_local_file`                         | Internal canonical registry operation; legacy name is transport-only.                           | Complete in PPTX-P2: direct host branch is absent.                                     |
| Markdown history controls                      | Canonical `markdown.history.undo` / `markdown.history.redo` operations.                         | Complete in R2-221: Ribbon and Registry share the mounted TipTap history action.       |
| Markdown block-type controls                   | Canonical addressed `markdown.block.set_type` final-state operation.                            | Complete in R2-222: Ribbon, Slash Menu, and Registry share one text-block action.      |
| Markdown inline-mark controls                  | Aggregate explicit-range `markdown.text.set_marks` final-state mask.                            | Complete in R2-223: five Ribbon controls and Registry share one bounded action.        |
| Markdown list controls                         | Addressed `markdown.list.set_type` with explicit final list state.                              | Complete in R2-224: Ribbon, Slash Menu, and Registry share one list action.            |
| Markdown table/divider insertion               | Explicit-position bounded table and divider insert operations.                                  | Complete in R2-225: Ribbon, Slash Menu, and Registry share both actions.               |
| Markdown local-image insertion                 | Public path contract transformed to hidden hydrated staged bytes.                               | Complete in R2-226: Broker/Renderer validate and all producers share data-URL insert.  |
| Markdown frontmatter editing                   | Complete raw-YAML `markdown.frontmatter.set` envelope final state.                              | Complete in R2-227: UI/Registry share recovery-backed envelope persistence.            |
| Markdown table-relative commands               | Aggregate bounded `markdown.table.update` with explicit header state.                           | Complete in R2-228: eight TableMenu actions converge on one descriptor/action.         |
| Markdown top-level block commands              | Aggregate explicit `markdown.block.update` transaction.                                         | Complete in R2-229: menu, keymap, plus, drag, and Registry share one builder.          |
| Markdown code-block language                   | Addressed finite `markdown.code_block.set_language` final state.                                | Complete in R2-230: NodeView and Registry share one native attribute transaction.      |
| Markdown Save As                               | Canonical `markdown.document.save_as` forced-destination persistence.                           | Complete in R2-231: Shift-Save and Registry share `doSave(true)` and exact output.     |
| Markdown DOCX export                           | Canonical `markdown.document.export_docx` output operation.                                     | Complete in R2-232: UI/Registry share bounded image-aware DOCX generation.             |
| Markdown print dialog                          | Canonical `markdown.document.open_print_dialog` host-view operation.                            | Complete in R2-233: popup success/failure is explicit and does not claim headless PDF. |
| Markdown autosave preference                   | Explicit `markdown.document.set_auto_save { enabled }` final state.                             | Complete in R2-234: UI/Registry share persisted preference state.                      |
| Markdown selection                             | Explicit bounded `markdown.selection.set { from, to }`.                                         | Complete in R2-235: arbitrary text targets without history/recovery pollution.         |
| Markdown local assets                          | Session-bound chunk bridge plus browser-directory hydration.                                    | Complete in R2-236–R2-237: display/export bytes preserve authored paths.               |
| Markdown retained-command audit                | Format-owned producer-to-operation mapping with no `missing` entry.                             | Complete in R2-238: all 22 descriptors are machine-checked.                            |
| XLSX first five mutation aliases               | Canonical registry operations; legacy names removed in R6-09.                                   | Complete: direct branches and compatibility aliases are absent.                        |
| XLSX column-delete and worksheet aliases       | Canonical registry operations; legacy names removed in R6-09.                                   | Complete: direct branches and compatibility aliases are absent.                        |
| XLSX text-style Ribbon commands                | Rejected in favor of `xlsx.range.set_text_style`.                                               | Complete in R2-31 for bold, italic, underline, and strike.                             |
| XLSX alignment/font/fill/border/style commands | Rejected in favor of exact `xlsx.range.*` operations.                                           | Complete in R2-32–R2-36; `format-painter` is explicitly UI-only.                       |
| XLSX number-format/merge/clear/fill commands   | Rejected in favor of semantic exact range operations.                                           | Complete in R2-37–R2-40; decimal gestures assign an explicit pattern.                  |
| XLSX sort and remove-duplicates commands       | Rejected in favor of explicit range/key/header operations.                                      | Complete in R2-41–R2-43 with native sort and guarded row deduplication.                |
| XLSX hyperlink commands                        | Rejected in favor of exact set/remove operations.                                               | Browser save/reopen complete in R2-78 over the retained gateway.                       |
| XLSX table/protection commands                 | Rejected in favor of explicit range/sheet operations.                                           | Registry and browser package save/reopen complete through R2-81.                       |
| XLSX sparkline commands                        | Rejected in favor of `xlsx.sparkline.add`.                                                      | Complete in R2-82 with shared Undo and x14 save/reopen.                                |
| XLSX outline commands                          | Rejected in favor of two absolute final-state outline operations.                               | Complete in R2-83 with one Undo unit and browser save/reopen hydration.                |
| XLSX checkbox command                          | Rejected in favor of `xlsx.range.set_checkbox`.                                                 | Complete in R2-84 with native Undo and browser DV save/reopen.                         |
| XLSX Conditional Formatting panel command      | Rejected in favor of explicit `xlsx.conditional_format.*` operations.                           | Comparison lifecycle complete in R2-102; remaining rule families are tracked.          |
| XLSX Symbol dialog command                     | Rejected in favor of explicit `xlsx.cell.set_value`.                                            | Complete in R2-85; retained as a UI picker over the shared cell action.                |
| XLSX CSV import command                        | Rejected in favor of `xlsx.range.set_values` or staged CSV open.                                | Complete in R2-86; retained as a UI picker/decoder over shared writes.                 |
| XLSX chart insertion commands                  | Rejected in favor of bounded `xlsx.chart.add`.                                                  | Complete in R2-87 with shared Undo and browser drawing save/reopen.                    |
| XLSX image producer commands                   | Rejected in favor of staged `xlsx.image.add`.                                                   | Complete in R2-88 with bounded staging and browser media save/reopen.                  |
| XLSX note mutation commands                    | Rejected in favor of explicit `xlsx.note.set/remove`.                                           | Complete in R2-89 with native Undo and comments/VML save/reopen.                       |
| XLSX chart edit/delete/format commands         | Rejected in favor of explicit-ID `xlsx.chart.update/remove`.                                    | Complete through R2-92, including stable reopened IDs and advanced format properties.  |
| XLSX chart palette/series commands             | Rejected in favor of explicit `xlsx.chart.set_colors/set_series`.                               | Complete in R2-91 with bounded final arrays and chart OOXML evidence.                  |
| XLSX shape/text-box insertion commands         | Rejected in favor of bounded `xlsx.shape.add`.                                                  | Complete in R2-93 with shared Undo and browser drawing save/reopen.                    |
| XLSX aggregate/Flash Fill/Text to Columns      | Rejected in favor of exact formula/range operations.                                            | Complete in R2-49–R2-51 with shared Ribbon/registry seams.                             |
| XLSX row-height command                        | Rejected in favor of `xlsx.row.set_height`.                                                     | Complete in R2-52 with explicit rows, native undo, and save/reopen.                    |
| XLSX column-width command                      | Rejected in favor of `xlsx.column.set_width`.                                                   | Complete in R2-53 with explicit columns and quantized saved width.                     |
| XLSX freeze-pane commands                      | Rejected in favor of `xlsx.sheet.set_freeze`.                                                   | Complete in R2-54 with explicit final splits, undo, and save/reopen.                   |
| XLSX gridline toggle                           | Rejected in favor of `xlsx.sheet.set_gridlines`.                                                | Complete in R2-55 with explicit visibility, undo, and save/reopen.                     |
| XLSX formula-view toggle                       | Rejected in favor of `xlsx.sheet.set_formula_view`.                                             | Complete in R2-56 with explicit state, shared undo, and save/reopen.                   |
| XLSX page-orientation commands                 | Rejected in favor of `xlsx.sheet.set_page_orientation`.                                         | Complete in R2-57 with bounded state, exact undo, and save/reopen.                     |
| XLSX page-margin commands                      | Rejected in favor of `xlsx.sheet.set_page_margins`.                                             | Complete in R2-58 with bounded presets, exact undo, and save/reopen.                   |
| XLSX paper-size commands                       | Rejected in favor of `xlsx.sheet.set_paper_size`.                                               | Complete in R2-59 with seven bounded presets, exact undo, and reopen.                  |
| XLSX fit-to-page commands                      | Rejected in favor of `xlsx.sheet.set_fit_to_pages`.                                             | Complete in R2-60 with explicit dual axes, exact undo, and reopen.                     |
| XLSX print-gridline command                    | Rejected in favor of `xlsx.sheet.set_print_gridlines`.                                          | Complete in R2-61 with explicit print state, exact undo, and reopen.                   |
| XLSX print-heading command                     | Rejected in favor of `xlsx.sheet.set_print_headings`.                                           | Complete in R2-62 with explicit print state, exact undo, and reopen.                   |
| XLSX print-area commands                       | Rejected in favor of `xlsx.sheet.set_print_area`.                                               | Complete in R2-63 with nullable A1 state, exact undo, and reopen.                      |
| XLSX print-title commands                      | Rejected in favor of `xlsx.sheet.set_print_titles`.                                             | Complete in R2-64 with bounded nullable rows, exact undo, and reopen.                  |
| XLSX fixed print-scale commands                | Rejected in favor of `xlsx.sheet.set_print_scale`.                                              | Complete in R2-65 with 10–400 bounds, fit disable, undo, and reopen.                   |
| XLSX AutoFilter toggle                         | Rejected in favor of `xlsx.range.set_filter`.                                                   | Complete in R2-66 with explicit final state, shared undo, and reopen.                  |
| XLSX filter clear/value/custom criteria        | Rejected in favor of three exact range filter operations.                                       | Complete in R2-67–R2-69 with bounds, shared undo, and reopen.                          |
| XLSX context-relative structure strings        | Rejected in favor of existing exact sheet/row/column operations.                                | Complete in R2-70 without adding duplicate registry operations.                        |
| XLSX history commands                          | Rejected in favor of `xlsx.history.undo` and `xlsx.history.redo`.                               | Complete in R2-71 over the mounted Univer native history.                              |
| XLSX paste-special value command               | Rejected in favor of `xlsx.range.copy_values`.                                                  | Complete in R2-72 with explicit bounded ranges and native undo.                        |
| XLSX paste-special formula command             | Rejected in favor of `xlsx.range.copy_formulas`.                                                | Complete in R2-73 with translated references and native undo.                          |
| XLSX paste-special format command              | Rejected in favor of `xlsx.range.copy_formats`.                                                 | Complete in R2-74 with full style replacement and content preservation.                |
| XLSX paste-special column-width command        | Rejected in favor of `xlsx.column.copy_widths`.                                                 | Complete in R2-75 with bounded spans, one native undo, and saved widths.               |
| XLSX paste-special except-border command       | Rejected in favor of `xlsx.range.copy_without_borders`.                                         | Complete in R2-76 with bounded ranges and preserved target borders.                    |
| XLSX copy/cut/paste clipboard strings          | Rejected in favor of explicit source/destination operations.                                    | Complete in R2-77; native UI clipboard controls remain available.                      |
| PDF legacy save/delete/undo aliases            | Canonical registry operations; legacy names removed in R6-09.                                   | Complete: Registry-only canonical dispatch rejects the retired names.                  |
| PDF local document/page byte inputs            | Internal staged operations; local paths are Broker transport only.                              | Complete: staged routes are hidden and direct Agent execution is rejected.             |
| XLSX `ribbon_command`                          | Removed after all retained families gained exact operations or explicit UI-only classification. | Complete: no open-ended Agent Ribbon route remains.                                    |
| Renderer `if`/`switch` dispatch                | Only the five internal staged-load transport aliases enter generic registry alias resolution.   | Public descriptor/handler coverage is canonical-only.                                  |
| Hand-written Skill/protocol operation tables   | Check against generated output.                                                                 | Installed references are generated from the executable registries.                     |

No new retained capability may be implemented only in a transitional path after the registry foundation milestone begins.
