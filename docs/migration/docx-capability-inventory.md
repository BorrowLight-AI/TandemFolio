# DOCX community renderer capability inventory

- Baseline: `genspark-ai/genoffice@dc4d7e5927864498913b7ba42d0da06cc7cf628e`
- Renderer authority: the mounted TipTap/ProseMirror community editor in `apps/docs/src/renderer`
- Current registry: 102 operations (98 Agent-visible, four internal staged operations)
- Capability status: retained state-changing command parity complete; R6-01 shared release gate passed; `ready === true`

This inventory is the R2-132 command-parity baseline. It treats visible buttons that resolve to the
same native transaction as one semantic family, and separates document mutations from transient UI,
navigation, clipboard arming, file pickers, print/export, zoom, and read-only views. A family is not
complete merely because a broader block command can approximate it: exact selection/identity,
native Undo, recovery, and save/reopen behavior must match the retained UI route.

## Already registry-owned

| Family                     | Operations                                                                                                                                                                                                                                                                                                                                                  | Current shared route                                                                                                                                                                     |
| -------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Text insertion/replacement | `docx.text.insert`, `docx.text.replace_selection`, `docx.text.replace_all`                                                                                                                                                                                                                                                                                  | Bounded shared insertion (including the symbol palette), mounted selection replacement, or retained structured-command transaction.                                                      |
| History                    | `docx.history.undo`, `docx.history.redo`                                                                                                                                                                                                                                                                                                                    | The same mounted TipTap history used by the retained Ribbon and quick-access controls; unavailable entries fail closed.                                                                  |
| Revisions                  | `docx.revision.set_tracking`, `docx.revision.apply_decision`                                                                                                                                                                                                                                                                                                | Explicit recorder state plus bounded accept/reject × current/all decisions share retained Review UI, mounted revision state, and native history.                                         |
| Block structure            | `docx.block.delete`, `docx.block.move`                                                                                                                                                                                                                                                                                                                      | Retained top-level block matcher and one ProseMirror transaction.                                                                                                                        |
| Heading/paragraph          | `docx.paragraph.set_heading_level`, `docx.paragraph.set_style`, `docx.paragraph.set_direction`                                                                                                                                                                                                                                                              | Bounded ParagraphDialog/Ribbon alignment, line mode, spacing, geometry and Ruler tab stops plus target-aware LTR/RTL through native transactions.                                        |
| Text style                 | `docx.text.set_style`, `docx.text.set_character_format`, `docx.text.clear_character_format`, `docx.text.transform_case`, `docx.text.set_character_style`                                                                                                                                                                                                    | Block-targeted style plus complete exact-range marks/font/size/color/highlight, Clear Formatting, four case modes, and document-owned character styles through Ribbon-shared helpers.    |
| Links                      | `docx.text.set_link`                                                                                                                                                                                                                                                                                                                                        | One exact range plus nullable href/text covers insert, replacement, href update, and removal through the Link modal's native transaction.                                                |
| Bookmarks                  | `docx.bookmark.set`                                                                                                                                                                                                                                                                                                                                         | Stable top-level block identity plus a bounded name and explicit enabled state serve the retained Bookmark dialog through native Undo/save projection.                                   |
| Cross-references           | `docx.cross_reference.insert`                                                                                                                                                                                                                                                                                                                               | Exact inline range, existing bookmark identity, and explicit cached display text serve the retained Cross-reference dialog through native Undo/save projection.                          |
| Generic fields             | `docx.field.insert`, `docx.field.update`                                                                                                                                                                                                                                                                                                                    | Exact insert ranges plus bounded aggregate cache updates serve retained field insertion/F9 through native Undo/save projection.                                                          |
| Notes                      | `docx.note.insert`, `docx.note.update`, `docx.note.delete`                                                                                                                                                                                                                                                                                                  | Exact-range insertion plus stable-ID body update/delete store final state on Undo-owned reference atoms and reconcile note parts through the same UI/Agent state seam; lifecycle closed. |
| Bibliography sources       | `docx.source.upsert`                                                                                                                                                                                                                                                                                                                                        | Stable-tag add/update stores one bounded final Sources override on the first block, shared by retained UI, native Undo, and customXml save projection.                                   |
| Citations                  | `docx.citation.insert`                                                                                                                                                                                                                                                                                                                                      | Existing stable source tag plus exact inline range and explicit display text serve retained plain-text citation insertion through native Undo/save projection.                           |
| Bibliographies             | `docx.bibliography.insert`                                                                                                                                                                                                                                                                                                                                  | Explicit source-backed entries insert after one stable top-level boundary as a single native Undo unit and persist as ordinary heading/paragraph blocks.                                 |
| Captions                   | `docx.caption.insert`                                                                                                                                                                                                                                                                                                                                       | Explicit label/number/text insert one dirty SEQ field after a stable block boundary through one native Undo/save projection.                                                             |
| Index entries              | `docx.index.mark`, `docx.index.insert`                                                                                                                                                                                                                                                                                                                      | Exact-range XE marking plus stable-boundary cached INDEX insertion share bounded native Undo/save projection; family closed.                                                             |
| Comments                   | `docx.comment.add`, `docx.comment.reply`, `docx.comment.set_resolved`, `docx.comment.delete`                                                                                                                                                                                                                                                                | The complete add/reply/resolve/delete lifecycle shares Undo-owned snapshots and exact anchor transactions; family closed.                                                                |
| Lists                      | `docx.list.apply`, `docx.list.remove`, `docx.list.set_level`, `docx.list.apply_preset`, `docx.list.restart`, `docx.list.continue`                                                                                                                                                                                                                           | Conversion, absolute level, bounded definitions, stable restart, and explicit-source continuation share the retained numbering state and transaction seam.                               |
| Tables                     | `docx.table.insert`, `docx.table.delete`, `docx.table.insert_rows`, `docx.table.delete_rows`, `docx.table.insert_columns`, `docx.table.delete_columns`, `docx.table.merge_cells`, `docx.table.split_cell`, `docx.table.set_cell_format`, `docx.table.set_cell_borders`, `docx.table.set_style`, `docx.table.set_row_height`, `docx.table.set_column_widths` | Lifecycle, structural edits, cell format/style, and complete row/column sizing share native table state and Undo; family closed.                                                         |
| Images                     | `docx.image.insert`, internal `docx.image.insert_staged`, `docx.image.remove`, `docx.image.replace`, internal `docx.image.replace_staged`, `docx.image.set_crop`, `docx.image.set_wrap`, `docx.image.set_margin_position`, `docx.image.set_offset_position`, `docx.image.set_transform`, `docx.image.update`                                                | Complete lifecycle, bytes/cutout, crop/reset, wrap/inline, position, transform, and dimensions/alignment share native image nodes, Undo, and save projection; family closed.             |
| Shapes/lines               | `docx.shape.insert`, `docx.line.insert`, `docx.textbox.set_content`                                                                                                                                                                                                                                                                                         | All 104 filled presets and five stroke-only kinds share insertion; textbox-backed shapes use the bounded rich-content route plus native Undo/save projection.                            |
| Textboxes                  | `docx.textbox.insert`, `docx.textbox.set_content`                                                                                                                                                                                                                                                                                                           | Stable insertion and exact nested rich-content replacement share the mounted textbox model, native Undo, and save projection.                                                            |
| Charts                     | `docx.chart.insert`, `docx.chart.update`                                                                                                                                                                                                                                                                                                                    | Bounded insertion and same-shape content updates share the retained chart UI, native Undo, and original/generated save projection; aggregate object size/remove close the lifecycle.     |
| Equations                  | `docx.equation.insert`, `docx.equation.update`                                                                                                                                                                                                                                                                                                              | Gallery/modal LaTeX rebuilds and retained block-token edits share exact block/inline targets, native Undo, and save projection with Registry.                                            |
| Object lifecycle           | `docx.object.set_size`, `docx.object.set_offset_position`, `docx.object.set_style`, `docx.object.remove`                                                                                                                                                                                                                                                    | Exact geometry/style and cross-object removal kernels serve retained object UI plus Registry, native Undo, and save projection; images retain their separate closed lifecycle.           |
| TOC                        | `docx.toc.insert`, `docx.toc.refresh`                                                                                                                                                                                                                                                                                                                       | Heading-generated insertion plus stable-region bounded refresh share native field nodes, Undo, and save projection; family closed.                                                       |
| Page structure             | `docx.document.insert_page_break`, `docx.section.insert_break`, `docx.section.set_orientation`, `docx.section.set_margins`, `docx.section.set_page_size`, `docx.section.set_columns`, `docx.section.set_page_border`                                                                                                                                        | Every retained page/section UI mutation shares Undo-owned state and save/reopen projection; family closed.                                                                               |
| Header/footer              | `docx.section.set_different_first_page`, `docx.document.set_different_odd_even_pages`, `docx.section.set_page_numbering`, `docx.header_footer.set_text`, `docx.header_footer.set_page_number`, `docx.header_footer.set_paragraphs`                                                                                                                          | Flags, indexed numbering, plain/rich content, and PAGE/NUMPAGES fields share the section-anchor journal, Ribbon controls, Undo, and save projection; family closed.                      |
| Document boundary          | `docx.document.save`, internal `docx.document.load_staged`                                                                                                                                                                                                                                                                                                  | Existing browser save/recovery and staged open paths.                                                                                                                                    |
| Document protection        | `docx.document.set_protection`                                                                                                                                                                                                                                                                                                                              | Explicit read-only final state, bounded password input, retained password verification, dirty recovery, and settings-part save/reopen share the Review UI and Registry route.            |
| Ink                        | `docx.ink.apply`                                                                                                                                                                                                                                                                                                                                            | Bounded vector add, stable-ID delete, and clear share the renderer-owned overlay state and authoritative ink save/reopen projection; family closed.                                      |
| Compare                    | `docx.document.compare`, internal `docx.document.compare_staged`                                                                                                                                                                                                                                                                                            | The Broker stages a bounded local DOCX as opaque bytes; retained UI and Registry share deterministic paragraph diff and one recovery-free comparison-panel commit.                       |
| Document design            | `docx.document.set_design`                                                                                                                                                                                                                                                                                                                                  | One masked aggregate action atomically sets bounded page color, watermark, theme fonts, and theme colors through retained UI/Registry state, dirty recovery, and save projection.        |
| Cover pages                | `docx.cover_page.insert`                                                                                                                                                                                                                                                                                                                                    | Twelve bounded style presets consume explicit replayable title/subtitle/author/date/year content and insert at document start in one native Undo/save projection.                        |
| Drop caps                  | `docx.paragraph.set_drop_cap`                                                                                                                                                                                                                                                                                                                               | One explicit top-level block, none/drop/margin final state, and bounded line count share the retained Ribbon's native paragraph transaction and save projection.                         |
| WordArt                    | `docx.wordart.insert`                                                                                                                                                                                                                                                                                                                                       | Explicit boundary, finite style, text, geometry, and unique drawing identity replace random/UI defaults through one native Undo and OOXML save/reopen projection.                        |

## Retained state-changing families still to migrate

None. R2-220 closes the inventory established by R2-132. DOCX remained `ready === false` until
R6-01 passed the shared visual, performance, source, round-trip, packaging, and release
documentation gates; approved evidence now generates `ready === true`.

## Non-document or external gestures

These remain renderer-local and do not become document-mutation tools: File/Open pickers, Copy,
format painter arming, Find/navigation panes, zoom/read/focus/split modes, show-formatting-marks,
statistics, print/PDF export, dialog opening, and selection-only navigation. Cut/Paste outcomes are
represented by deterministic explicit delete/insert/replace operations rather than clipboard state.

## Closure rule

Each row above closes only when its retained UI producer and Agent handler share the same mounted
state/transaction seam, exact targets and bounded schema are validated by Broker and renderer,
successful mutation advances one revision, native or renderer-owned Undo is proven where
applicable, and the saved DOCX reopens with edited and unrelated parts preserved. The audit becomes
complete when no retained state-changing source route remains unmapped.

R2-133 closes the History family with explicit empty-input contracts and exact success outputs.
Both operations invoke the mounted editor's native TipTap history, preserve the Ribbon's document
and selection semantics, and reject missing undo/redo entries without acknowledging a mutation.

R2-134 starts the Exact character formatting family with
`docx.text.set_character_format { range, format, fields }`. The explicit range uses the same
revision-scoped ProseMirror coordinates exposed in live session context. A unique exact field mask
sets or clears bold, italic, underline, strike, and baseline/superscript/subscript in one native
transaction while preserving unlisted per-run style. The retained Ribbon uses the same helper for
non-empty selections; native Undo restores the whole edit. The family stays open for the remaining
font/color/size, character-style, Clear Formatting, and case-transform routes.

R2-135 expands that same operation and shared helper with bounded font family, 1–1638pt half-point
font size, `#RRGGBB` text color, and the sixteen DOCX named highlight colors. `null` explicitly
clears these properties. Per-run style IDs, spacing, script-specific font slots, and every other
unlisted attribute remain distinct across heterogeneous selections. The registry count is
unchanged because this deepens the existing semantic operation rather than adding another tool.

R2-136 adds `docx.text.clear_character_format { range }` over the shared explicit text-range
validator. It applies the retained Ribbon's Clear Formatting semantics only inside that range,
keeps text and range-external marks intact, restores the edit with one native Undo, and makes the
Ribbon call the same helper for non-empty selections.

R2-137 adds `docx.text.transform_case { range, mode }` for sentence, lower, upper, and title case.
It preserves every text node's marks, maps positions when Unicode case conversion changes string
length, remains one native Undo transaction, and replaces the Ribbon's private implementation with
the same format-owned helper.

R2-138 closes the Exact character formatting family with
`docx.text.set_character_style { range, styleId }`. Non-null IDs must resolve to a character style
in the currently opened document; application preserves direct run formatting, while `null`
matches the retained Ribbon's character-style mark removal. Built-in Emphasis/Strong fallback
cards reuse `docx.text.set_character_format` semantics and do not introduce preset strings into the
contract.

R2-139 starts Paragraph/list details with
`docx.paragraph.set_direction { target, direction }`. It reuses the Ribbon's sole direction-flip
rule: setting LTR/RTL updates `bidi` and swaps explicit left/right alignment to preserve logical
start/end. Multiple matching paragraph-like blocks change in one native Undo transaction.

R2-140 adds `docx.list.set_level { target, level }` as the replayable absolute counterpart to the
Ribbon's relative indent controls. Both paths share the same `0..8` normalization rule; only
matching `docListItem` nodes change, and each target's original level returns with one Undo.

R2-141 adds `docx.list.apply_preset { target, levels }`. The 1–9-level schema bounds numbering
format, level text, indentation, hanging indent, and optional start value. Registry execution and
the retained custom-list dialog both allocate definitions through the same format-owned numbering
state; all matching paragraph/list blocks bind to the new `numId` in one native Undo transaction.
Empty or invalid presets fail before a numbering definition is created. The generated Manifest
contains 147 operations, including twenty-three DOCX operations (twenty-two Agent-visible and one
internal). Restart and continue remain the only open list-numbering transitions.

R2-142 adds `docx.list.restart { blockIndex, start }`. The stable top-level block index replaces
selection-relative Agent behavior, while the bounded start value replaces the UI-only fixed value.
The shared action clones the source `abstractNum` with a level-specific start override and rewrites
only same-`numId` items at and after the anchor in one native Undo transaction. The retained
context-menu route calls the same action with its current block and start `1`. The generated
Manifest contains 148 operations, including twenty-four DOCX operations (twenty-three
Agent-visible and one internal). Continue numbering is the remaining list transition.

R2-143 adds `docx.list.continue { blockIndex, previousBlockIndex }`. Both inputs are stable
top-level list-item indexes and the source must precede the current anchor. The action rewrites the
current `numId` only at and after the anchor to the explicitly selected previous `numId`, preserving
other lists and grouping the body change into one native Undo transaction. The retained context
menu resolves its legacy nearest-previous gesture before entering the same action. The generated
Manifest contains 149 operations, including twenty-five DOCX operations (twenty-four Agent-visible
and one internal). The list-numbering family is closed.

R2-144 deepens `docx.paragraph.set_style { target, style, fields }` without adding another tool.
The contract now includes the retained ParagraphDialog's `auto`, `atLeast`, and `exact` line modes,
raw 20–31,680-twip line height, 0.06–132 multiple spacing, and bounded paragraph indentation and
before/after spacing. The target-aware executor writes the same mounted node attributes as the
Ribbon and dialog in one native Undo transaction. The Manifest remains at 149 operations, including
twenty-five DOCX operations (twenty-four Agent-visible and one internal). Ruler tab stops are the
remaining paragraph-specific mutation.

R2-145 deepens the same `docx.paragraph.set_style` contract with `tabStops`. A non-null final state
contains 1–64 strictly increasing stops; positions are bounded to 0–31,680 twips, alignment and
leader values are finite enums, and `null` clears custom stops. Registry and Ruler share the format
adapter that serializes this public array into the renderer's internal node attribute. One
target-aware transaction remains one native Undo entry, the Manifest stays at 149 operations, and
the retained paragraph-format family is closed.

R2-146 starts the table family with `docx.table.insert { afterBlockIndex, rows, columns }`.
`afterBlockIndex=-1` means document start; row/column axes are bounded to 100/63 and a second
renderer guard caps the product at 4096 cells. Registry insertion uses a stable top-level block
boundary, while the retained UI preserves selection-sensitive nested-table behavior; both consume
the same format-owned empty-table model and native table schema. The generated Manifest contains
150 operations, including twenty-six DOCX operations (twenty-five Agent-visible and one internal).

R2-147 adds `docx.table.delete { tableBlockIndex }`. The revision-scoped index must identify a
top-level native table; non-table targets fail before mutation. Deletion is one native Undo entry,
and deleting the document's sole table atomically leaves one valid empty paragraph. The generated
Manifest contains 151 operations, including twenty-seven DOCX operations (twenty-six Agent-visible
and one internal). Nested-table deletion remains a separate cell-path concern.

R2-148 adds `docx.table.insert_rows { tableBlockIndex, rowIndex, count }`. `rowIndex` is an
explicit boundary from zero through current row count and `count` is bounded to 1–100; resulting
logical cells may not exceed 4096. Each addition recomputes the retained ProseMirror `TableMap`, so
rowspans remain authoritative, while all additions remain in one transaction and Undo entry. The
generated Manifest contains 152 operations, including twenty-eight DOCX operations (twenty-seven
Agent-visible and one internal).

R2-149 adds `docx.table.delete_rows { tableBlockIndex, rowIndex, count }`. The explicit zero-based
start and 1–100 count must identify a proper subset of the table rows; deleting every row fails
closed and delegates that distinct lifecycle outcome to `docx.table.delete`. Each staged deletion
recomputes the native `TableMap`, retaining rowspan repair, while one accumulated transaction
creates one Undo entry. The generated Manifest contains 153 operations, including twenty-nine DOCX
operations (twenty-eight Agent-visible and one internal).

R2-150 adds `docx.table.insert_columns { tableBlockIndex, columnIndex, count }`. The explicit
boundary may range from zero through current logical width; `count` is bounded to 1–63, final width
to 63, and final logical cells to 4096. Each staged insertion recomputes the native `TableMap` for
colspan-aware behavior and the accumulated transaction produces one Undo entry. The generated
Manifest contains 154 operations, including thirty DOCX operations (twenty-nine Agent-visible and
one internal).

R2-151 adds `docx.table.delete_columns { tableBlockIndex, columnIndex, count }`. The explicit
logical interval must be in range and leave at least one column; deleting every column fails closed
in favor of `docx.table.delete`. Recomputing the native `TableMap` between removals repairs
colspans while one accumulated transaction yields one Undo entry. The generated Manifest contains
155 operations, including thirty-one DOCX operations (thirty Agent-visible and one internal).

R2-152 adds
`docx.table.merge_cells { tableBlockIndex, topRow, leftColumn, bottomRow, rightColumn }`. The
bounded half-open rectangle must cover at least two logical cells, fit the current `TableMap`, and
not cross an existing merged-cell boundary. Registry execution synthesizes the exact native
`CellSelection` and invokes the retained merge command in one Undo transaction. The shared
`docTableRow` schema now permits zero physical cells when the whole row is covered by an earlier
rowspan, matching ProseMirror table semantics and fixing the retained UI path too. The generated
Manifest contains 156 operations, including thirty-two DOCX operations (thirty-one Agent-visible
and one internal).

R2-153 adds `docx.table.split_cell { tableBlockIndex, rowIndex, columnIndex }`. The bounded logical
coordinate resolves through the current `TableMap`, so any coordinate covered by the same merged
cell identifies that cell deterministically. Ordinary cells fail before mutation; merged cells use
the retained native split command, report the original `rowspan × colspan` logical-cell count, and
restore with one Undo. The generated Manifest contains 157 operations, including thirty-three DOCX
operations (thirty-two Agent-visible and one internal). Exact merge/split lifecycle is closed.

R2-154 adds
`docx.table.set_cell_format { tableBlockIndex, topRow, leftColumn, bottomRow, rightColumn, format, fields }`.
The exact half-open rectangle cannot cross an existing merged-cell boundary. A unique field mask
sets or clears six-digit uppercase fill and sets top/center/bottom vertical alignment while
preserving unlisted attrs. Registry explicit targets and Ribbon selection targets share the same
physical-cell write kernel and one native Undo transaction. The generated Manifest contains 158
operations, including thirty-four DOCX operations (thirty-three Agent-visible and one internal).

R2-155 adds
`docx.table.set_cell_borders { tableBlockIndex, topRow, leftColumn, bottomRow, rightColumn, mode, border }`.
The exact rectangle accepts `all`, `outer`, `inner`, or `none`; a drawing policy uses uppercase
six-digit color and 1–96 eighth-point width, while `none` requires `border: null`. Existing
unaddressed sides remain intact for outer/inner. Registry and all four Ribbon controls share the
same merged-cell-aware edge/write kernel and one Undo transaction. The generated Manifest contains
159 operations, including thirty-five DOCX operations (thirty-four Agent-visible and one internal).

R2-156 adds `docx.table.set_style { tableBlockIndex, styleId }`. Non-null IDs must resolve to a
table style in the current document; `null` explicitly removes styling, and repeated final state
returns `changed: false`. Registry stable targets and retained Ribbon style cards share the same
native table-attribute kernel and Undo route. The generated Manifest contains 160 operations,
including thirty-six DOCX operations (thirty-five Agent-visible and one internal). Table style is
closed; row/column sizing remains the last retained table family.

R2-157 adds `docx.table.set_row_height { tableBlockIndex, rowIndex, count, heightTwips }`. The
physical-row interval is bounded to 1–100 rows; height is 1–31,680 OOXML twips or `null` for auto.
Only `heightTwips` changes, preserving height rule and retained row metadata. Ribbon retains its
centimeter/page-box conversion but shares the same row write kernel and one Undo transaction. The
generated Manifest contains 161 operations, including thirty-seven DOCX operations (thirty-six
Agent-visible and one internal). Column sizing is the remaining table mutation.

R2-158 adds `docx.table.set_column_widths { tableBlockIndex, widthsPx }`. The input is a complete
1–63-entry vector matching current logical width; each column is 40–4096px and aggregate width is
at most 4096px. The format-owned sizing module synchronizes spanning-cell `colwidth`, table
`widthPx`, and percentage grid in one native Undo transaction. Ribbon continues to derive a vector
from selection and section width via `fitColumnWidths`, while Agent supplies the replayable final
vector; both use `writeGridWidths`. The generated Manifest contains 162 operations, including
thirty-eight DOCX operations (thirty-seven Agent-visible and one internal). The retained table
family is closed.

R2-159 starts the page/section family with
`docx.document.insert_page_break { afterBlockIndex }`. The explicit `-1..lastBlock` top-level
boundary inserts the renderer's native empty `docParagraph` with `pageBreakBefore: true` in one
Undo transaction; the retained Ribbon selection route and registry boundary route share that
writer. Invalid boundaries fail before dispatch. The generated Manifest contains 163 operations,
including thirty-nine DOCX operations (thirty-eight Agent-visible and one internal). Section
breaks and page/section settings remain open.

R2-160 adds
`docx.section.insert_break { afterBlockIndex, startType }` with the finite `nextPage`,
`continuous`, `evenPage`, and `oddPage` states. The inserted protected break node owns both copied
`sectPr` XML and the new section start type, so native Undo removes the complete pending mutation.
Save projects that state onto the following inserted/original `sectPr` or the hidden trailing
`sectPr`; consecutive inserted breaks and save/reopen are covered. Ribbon selection and Agent
stable-boundary adapters share the same insertion kernel. The generated Manifest contains 164
operations, including forty DOCX operations (thirty-nine Agent-visible and one internal).

R2-161 adds `docx.section.set_orientation { sectionIndex, orientation }`. Portrait/landscape is
an explicit final state; switching orientation swaps the current page axes while preserving all
other section settings. A full settings override is attached to the section's last visible native
block, making the change one ProseMirror Undo unit. Save applies overrides to either the indexed
non-final section terminator or the final hidden `sectPr`, and the Ribbon's orientation command now
uses this same journal. The generated Manifest contains 165 operations, including forty-one DOCX
operations (forty Agent-visible and one internal).

R2-162 adds `docx.section.set_margins { sectionIndex, margins }`. All four sides are explicit
0–31,680 twip integers, and renderer validation rejects horizontal or vertical sums that leave no
positive body area. Agent and retained preset/custom-margin UI write the same complete section
override; save/reopen preserves all four sides. The generated Manifest contains 166 operations,
including forty-two DOCX operations (forty-one Agent-visible and one internal).

R2-163 adds `docx.section.set_page_size { sectionIndex, widthTwips, heightTwips }`. Both actual
page axes are explicit 1,440–31,680 twip integers and must exceed the current margin sums. The
renderer derives landscape/portrait from the final axes (preserving orientation for a square), so
the operation is independent of Ribbon preset context. Retained paper presets already produce the
same full section override. The generated Manifest contains 167 operations, including forty-three
DOCX operations (forty-two Agent-visible and one internal).

R2-164 adds `docx.section.set_columns { sectionIndex, count, spacingTwips }`. The column count is
1–16, spacing is 0–31,680 twips, and total gaps must leave positive text width after margins. The
RED save/reopen test exposed and fixed the retained engine writer's hard-coded 425-twip fallback:
`SectionSettings.colSpace` is now written exactly for existing or newly created `<w:cols>`.
Ribbon count presets retain current spacing through the same journal. The generated Manifest
contains 168 operations, including forty-four DOCX operations (forty-three Agent-visible and one
internal).

R2-165 adds `docx.section.set_page_border { sectionIndex, enabled }`. The retained Design-tab
on/off control and Agent route share the full settings journal, native Undo, and existing
`w:pgBorders` save projection. This closes every retained UI write in the page/section family;
parsed-only `headerDist`, `footerDist`, `vAlign`, `docGrid`, and `textDirection` remain preserved
data rather than invented mutation tools. The generated Manifest contains 169 operations,
including forty-five DOCX operations (forty-four Agent-visible and one internal).

R2-166 starts the header/footer family with
`docx.section.set_different_first_page { sectionIndex, enabled }`. `titlePg` now lives beside the
full settings snapshot on the indexed section anchor, so the retained Ribbon toggle and Agent
operation share native Undo. Save writes or removes `w:titlePg` in either a non-final section-break
paragraph or the final hidden `sectPr`; save/reopen is covered. The generated Manifest contains
170 operations, including forty-six DOCX operations (forty-five Agent-visible and one internal).

R2-167 adds `docx.document.set_different_odd_even_pages { enabled }`. The document-wide boolean is
stored on the first section anchor in the same native history journal, while save projects it to
presence or absence of `w:evenAndOddHeaders` in `settings.xml`. The retained Ribbon checkbox and
Agent route share the same helper; save/reopen and Undo are covered. The generated Manifest contains
171 operations, including forty-seven DOCX operations (forty-six Agent-visible and one internal).

R2-168 adds `docx.section.set_page_numbering { sectionIndex, format, start }`. The format is one of
the retained seven-value gallery and `start` is an explicit `null` or bounded `0..999999` integer.
The old final/non-final React dirty-state split is removed: the dialog and Agent route now write the
same section-anchor journal, and save projects `w:pgNumType` to the indexed non-final or final
`sectPr`. Both paths reopen exactly and native Undo restores the prior state. The generated Manifest
contains 172 operations, including forty-eight DOCX operations (forty-seven Agent-visible and one
internal).

R2-169 adds `docx.header_footer.set_text { sectionIndex, kind, variant, text }`. Explicit section,
header/footer kind, and `default|first|even` variant replace implicit canvas context. Both Agent
plain text and the retained HeaderFooterArea/Ribbon content values use one `headerFooterEdits`
section-anchor journal. The former header/footer React dirty save channels are removed. The retained
engine now accepts variant-aware non-final `sectionHf` writes, while final variants keep their
existing surgical part writers; save/reopen and native Undo are covered. The generated Manifest
contains 173 operations, including forty-nine DOCX operations (forty-eight Agent-visible and one
internal).

R2-170 adds
`docx.header_footer.set_page_number { sectionIndex, kind, variant, enabled, alignment }`. Enabling
replaces the target with one left/center/right PAGE-field paragraph using the unambiguous private
sentinel; disabling removes PAGE fields while retaining other rich text. Ribbon gallery/removal and
Agent execution share the content-journal helper, native Undo, and save/reopen route. The generated
Manifest contains 174 operations, including fifty DOCX operations (forty-nine Agent-visible and one
internal).

R2-171 adds `docx.header_footer.set_paragraphs { sectionIndex, kind, variant, paragraphs }`.
The bounded model allows 1–64 paragraphs, at most 256 text/PAGE/NUMPAGES segments per paragraph,
and the retained run styles (bold, italic, underline, strike, color, font, and half-point size).
Renderer validation caps aggregate text at 65,536 characters and rejects malformed field/color/font
values. It maps to the same full-value content journal used by HeaderFooterArea, with native Undo and
rich save/reopen evidence. The generated Manifest contains 175 operations, including fifty-one DOCX
operations (fifty Agent-visible and one internal); the retained header/footer family is closed.

R2-172 starts the image lifecycle with public
`docx.image.insert { path, afterBlockIndex, widthPx, heightPx, alignment }` and internal
`docx.image.insert_staged`. The Broker validates and stages an absolute PNG/JPEG/GIF path as a
session-scoped blob of at most 20 MiB, never placing bytes in `office_execute`; the mounted DOCX
session hydrates the internal descriptor and revalidates byte count, extension, and magic before
one stable-boundary ProseMirror insertion. Agent insertion and retained dialog/paste insertion now
share the same native image-node builder. Native Undo removes the complete image, and a real GIF
save/reopen/resave proves surgical persistence. The generated Manifest contains 177 operations,
including fifty-three DOCX operations (fifty-one Agent-visible and two internal).

R2-173 adds public
`docx.image.replace { path, imageBlockIndex, widthPx, heightPx }` and internal
`docx.image.replace_staged`. The revision-scoped top-level index must identify an image; final
1–4096px dimensions make the replacement result explicit. The Broker reuses session image staging,
and the renderer repeats byte-count/extension/magic checks. Original images receive the retained
`imageReplace` surgical patch, while unsaved images update `genImage`; both preserve alignment,
wrap, rotation, flip, and floating position, and clear stale crop/fill windows exactly like the
retained Replace Picture UI. Ribbon/crop/cutout and Agent routes share the same replacement-attrs
kernel, native Undo restores the prior bytes and geometry, and real save/reopen is covered. The
generated Manifest contains 179 operations, including fifty-five DOCX operations (fifty-two
Agent-visible and three internal).

R2-197 closes the retained symbol palette without adding an operation. Existing `docx.text.insert`
now bounds input to 1–65,536 Unicode characters and dispatches through shared `insertDocxText` with
one native close-history transaction; the symbol palette calls the same action at the mounted active
selection. Empty/oversized input fails before mutation, Unicode symbol Undo and save/reopen are
covered, and the generated Manifest remains at 202 operations including seventy-eight DOCX
operations (seventy-five Agent-visible and three internal). Priority 9 is closed.

R2-198 adds `docx.toc.refresh { tocBlockIndex, entries }`. A stable top-level start identifies the
exact original or generated TOC fldChar region; 1–1024 final entries bound level, Unicode text, and
nullable page number under a 65,536-character aggregate cap. The action validates the matching field
end, preserves a trailing page break, and replaces the whole region in one native Undo transaction.
The retained Update TOC button and Registry share this action, with generated-to-original reopen,
Undo/redo, and refreshed save/reopen evidence. The generated Manifest contains 203 operations,
including seventy-nine DOCX operations (seventy-six Agent-visible and three internal), closing TOC.

R2-199 adds `docx.note.insert { range, kind, noteId, text }`. The exact inline range, finite note
kind, positive stable ID, and 1–65,536-character body are validated before one native-history
transaction. Newly created note text lives on the reference atom, so Undo/Redo reconciles both the
marker and note-part membership instead of leaving orphan React state. Retained footnote/endnote UI
and Registry share the action; duplicate IDs, invalid ranges, Broker validation, and real
save/reopen are covered. The generated Manifest contains 204 operations, including eighty DOCX
operations (seventy-seven Agent-visible and three internal).

R2-200 adds `docx.note.update { kind, noteId, text }`. The stable positive note identity must exist
in both the selected note part and at least one document reference. A bounded final body is written
to every matching reference atom in one transaction; the session retains the original body baseline
so Undo/Redo restores the correct note-part text across repeated updates. The retained edit dialog,
Registry, Broker, malformed-input rejection, and surgical save/reopen share evidence. The generated
Manifest contains 205 operations, including eighty-one DOCX operations (seventy-eight Agent-visible
and three internal).

R2-201 adds `docx.note.delete { kind, noteId }`. The stable positive identity must resolve in the
note part and document; every matching reference is removed and remaining same-kind markers are
renumbered in one native transaction. The shared note history record distinguishes inserted,
updated, and deleted final states, so Undo restores the correct body and numbering even after prior
updates. Retained UI, Registry, Broker rejection, and save/reopen are covered. The generated
Manifest contains 206 operations, including eighty-two DOCX operations (seventy-nine Agent-visible
and three internal), closing the retained note lifecycle.

R2-202 adds `docx.source.upsert { source }`. A bounded stable tag and finite Word source type
identify an add or replacement; modeled strings and the aggregate 1,024-source snapshot are
bounded. The final Sources array is stored as an Undo-owned first-block override, while the retained
dialog and Registry commit it to the same customXml save state. Create/update Undo/Redo, malformed
tags, Broker projection, and save/reopen are covered. The generated Manifest contains 207
operations, including eighty-three DOCX operations (eighty Agent-visible and three internal).

R2-203 adds `docx.citation.insert { range, sourceTag, displayText }`. The source tag must resolve in
the current Sources snapshot, and bounded explicit display text replaces one exact inline range in
a native transaction. The retained citation menu and Registry share the action without claiming a
richer Word CITATION field than the renderer retains. Missing sources, Broker validation, Undo, and
save/reopen are covered. The generated Manifest contains 208 operations, including eighty-four
DOCX operations (eighty-one Agent-visible and three internal).

R2-204 adds `docx.bibliography.insert { afterBlockIndex, heading, entries }`. Every bounded entry
names an existing stable source tag, and the explicit heading plus final display lines insert after
one revision-scoped top-level boundary in a single native transaction. The retained Bibliography
button and Registry share the action; malformed/unknown sources, one-step Undo, Broker projection,
and save/reopen are covered. The generated Manifest contains 209 operations, including eighty-five
DOCX operations (eighty-two Agent-visible and three internal).

R2-205 adds `docx.caption.insert { afterBlockIndex, label, number, text }`. The explicit positive
number makes the dirty SEQ field replayable, while the retained dialog keeps its existing preview
count and submits the final value. The bounded caption inserts after one revision-scoped block in
one native transaction. Invalid values, UI/Registry parity, one-step Undo, Broker projection, and
save/reopen are covered. The generated Manifest contains 210 operations, including eighty-six DOCX
operations (eighty-three Agent-visible and three internal).

R2-206 adds `docx.index.mark { range, term }`. The exact source range must remain within one inline
block; the hidden XE atom is inserted at its end, and bounded terms reject quotes/control characters
that the OOXML writer cannot round-trip exactly. Retained Mark Entry UI, Registry, native one-step
Undo, Broker projection/rejection, and save/reopen share the route. The generated Manifest contains
211 operations, including eighty-seven DOCX operations (eighty-four Agent-visible and three internal).

R2-207 adds `docx.index.insert { afterBlockIndex, label, terms }`. Bounded terms are normalized,
deduplicated, and deterministically sorted before the complete dirty INDEX cache inserts after one
stable top-level boundary in one transaction. Retained UI, Registry, one-step Undo, Broker
projection/rejection, and save/reopen share the action. Source UI audit confirms the retained
renderer exposes add/upsert only—there is no UI-only source deletion to migrate—so the References
and Index family is closed. The generated Manifest contains 212 operations, including eighty-eight
DOCX operations (eighty-five Agent-visible and three internal).

R2-208 adds `docx.comment.add { range, comment }`. The caller supplies a unique numeric ID,
bounded author/initials/text, and UTC timestamp. One native transaction attaches the exact-range
comment marks and writes the full final comment snapshot to the first block, so native Undo/Redo
reconciles React state and `comments.xml` together. Retained New Comment UI, Registry, duplicate-ID
rejection, Broker projection, and save/reopen are covered. The generated Manifest contains 213
operations, including eighty-nine DOCX operations (eighty-six Agent-visible and three internal).

R2-209 adds `docx.comment.reply { parentId, comment }`. A stable top-level parent identifies every
shared anchor range; the explicit reply metadata and all augmented marks commit with the final
comments snapshot in one native transaction. Missing parents/anchors and duplicate IDs fail closed.
Retained Reply UI, Registry, Broker, one-step Undo, and `commentsExtended.xml` save/reopen are
covered. The generated Manifest contains 214 operations, including ninety DOCX operations
(eighty-seven Agent-visible and three internal).

R2-210 adds `docx.comment.set_resolved { id, resolved }`. A stable top-level thread ID plus explicit
final boolean updates parent and replies in one Undo-owned comments snapshot without touching their
anchors. Repeated final state is a deterministic no-op. Retained Resolve/Reopen UI, Registry,
Broker validation, native Undo, and `commentsExtended.xml` save/reopen are covered. The generated
Manifest contains 215 operations, including ninety-one DOCX operations (eighty-eight Agent-visible
and three internal).

R2-211 adds `docx.comment.delete { id }`. A stable top-level ID deletes its full thread (parent plus
direct replies), while a reply ID deletes only that reply. All affected comment marks and the final
comment snapshot commit in one native-history transaction; missing IDs fail before mutation.
Retained Delete UI, Registry, Broker validation, one-step Undo, and comments-part save/reopen share
the route. The generated Manifest contains 216 operations, including ninety-two DOCX operations
(eighty-nine Agent-visible and three internal), closing the retained comment lifecycle.

R2-212 adds `docx.revision.set_tracking { enabled }`. The explicit final boolean updates the
mounted TrackChanges recorder immediately and synchronizes retained Ribbon state through the same
writer used by Registry execution. Repeated final state is a deterministic no-op. Because the
toggle is session recorder policy rather than document content, it is explicitly non-undoable and
does not checkpoint recovery; the next edit remains native-history-owned and its `w:ins`
save/reopen is covered. The generated Manifest contains 217 operations, including ninety-three
DOCX operations (ninety Agent-visible and three internal).

R2-213 adds
`docx.revision.apply_decision { decision: accept|reject, scope: current|all }`. One bounded schema
covers all four retained Review decisions without creating four tools. The current scope resolves
from the mounted selection; all scope snapshots every revision before applying one native-history
transaction. Empty documents fail closed. Retained UI, Registry, one-step Undo, Broker enum
validation, and accepted/rejected save/reopen share the route. The generated Manifest contains 218
operations, including ninety-four DOCX operations (ninety-one Agent-visible and three internal),
closing the retained revisions family.

R2-214 adds `docx.document.set_protection { enabled, password }`. One explicit boolean covers
passwordless/password-protected enablement and verified removal without exposing credentials in
results. Repeated final state is a deterministic no-op; incorrect passwords fail before shared
state changes. Retained Review modal and Registry use the same resolver. Protection is explicitly
non-undoable because the retained state lives in the document save journal rather than ProseMirror
history; dirty recovery and `settings.xml` enable/remove save/reopen are covered. The generated
Manifest contains 219 operations, including ninety-five DOCX operations (ninety-two Agent-visible
and three internal).

R2-215 adds `docx.ink.apply { action, annotation?, ids? }`. The finite action enum aggregates
vector-stroke add, stable-ID delete, and clear without one tool per gesture. IDs, anchors, colors,
widths, coordinate ranges, point counts, and delete batches are bounded; runtime cross-validates
the action-specific shape and rejects duplicate/missing identities before state changes. Opaque
reopened ink can be deleted but arbitrary base64 cannot be injected. Retained overlay/Ribbon and
Registry share the renderer-owned layer; the journal is explicitly non-undoable and has dirty
recovery plus ink-part save/reopen evidence. The generated Manifest contains 220 operations,
including ninety-six DOCX operations (ninety-three Agent-visible and three internal), closing Ink.

R2-216 adds public `docx.document.compare { path }` and internal
`docx.document.compare_staged { blobId, name, size, data }`. The Broker validates and stages the
local DOCX without exposing paths or bytes through the public Manifest; the renderer hydrates the
opaque payload, parses it through the existing DOCX engine, and commits the retained comparison
panel's deterministic paragraph diff. Comparison is an explicit view-state mutation, is
non-undoable and recovery-free, and does not dirty or rewrite the document package. Broker staging,
schema rejection, shared retained-UI action, and deterministic output are covered. The generated
Manifest contains 222 operations, including ninety-eight DOCX operations (ninety-four
Agent-visible and four internal), closing Compare.

R2-217 adds `docx.document.set_design { fields, pageColor?, watermark?, themeFonts?,
themeColors? }`. The 1–4-item unique mask must exactly match supplied values; page/watermark
nullability, font names, scheme names, and all document-writable theme colors are bounded before
state changes. A whole theme can therefore commit fonts and colors atomically without four public
tools or opaque preset strings. Retained Design Ribbon callbacks and Registry share the same
resolver/commit seam. The non-ProseMirror journal is explicitly non-undoable, remains dirty and
recovery-backed, and has combined page-color/watermark/theme save/reopen evidence. The generated
Manifest contains 223 operations, including ninety-nine DOCX operations (ninety-five
Agent-visible and four internal), closing Document design.

R2-218 adds `docx.cover_page.insert { preset, title, subtitle, author, date, year }`. Twelve finite
presets select styling only; every locale/date-sensitive visible value is explicit, bounded, and
therefore replayable. The retained Ribbon derives its localized defaults and calls the same action.
Insertion always targets document start, all generated paragraphs plus the page-break boundary are
one native Undo unit, and custom content survives save/reopen. The generated Manifest contains 224
operations, including 100 DOCX operations (ninety-six Agent-visible and four internal), closing
Cover pages.

R2-219 adds `docx.paragraph.set_drop_cap { blockIndex, mode, lines }`. The explicit top-level
paragraph/heading/list identity is revision-scoped; `none` requires null lines while `drop` and
`margin` require 2–10. Repeated final state is a no-op. Retained Ribbon resolves its current
top-level block then calls the same action, which commits one native transaction with Undo and
save/reopen evidence. The generated Manifest contains 225 operations, including 101 DOCX
operations (ninety-seven Agent-visible and four internal), closing Drop caps.

R2-220 adds `docx.wordart.insert { afterBlockIndex, preset, text, widthEmu, heightEmu, drawingId }`.
All twelve retained styles are finite; text and geometry are bounded; the OOXML drawing identity
is explicit and rejected when duplicated. Retained Ribbon allocates the smallest free identity and
calls the same action, deleting its random-ID/private insertion route. One native Undo transaction
and save/reopen evidence cover text, readable persisted style approximation, and geometry. The
generated Manifest contains 226 operations, including 102 DOCX operations (ninety-eight
Agent-visible and four internal). This closes the R2-132 retained state-changing command inventory;
release readiness remains gated separately.

R2-179 adds `docx.image.remove { imageBlockIndex }`. The revision-scoped top-level index must be
an image; deletion is one native transaction. Removing the only document block atomically inserts
an empty paragraph to preserve the ProseMirror `block+` invariant, matching retained selected-node
deletion. Native Undo and save/reopen deletion are covered. The generated Manifest contains 185
operations, including sixty-one DOCX operations (fifty-eight Agent-visible and three internal),
and the retained image lifecycle is closed.

R2-180 starts the object family with
`docx.shape.insert { afterBlockIndex, preset, widthEmu, heightEmu }`. Its enum is derived from the
shared Gallery and covers all 104 filled presets while reserving five line/connector kinds for a
separate stroke operation. Both axes are bounded to 9,525–20,000,000 EMU and insertion uses a
revision-scoped top-level boundary. Registry, Ribbon gallery, and draw mode share one native shape
node builder; native Undo and save/reopen are covered. The generated Manifest contains 186
operations, including sixty-two DOCX operations (fifty-nine Agent-visible and three internal).

R2-181 adds `docx.line.insert { afterBlockIndex, kind, widthEmu, heightEmu }` for the five retained
stroke-only line/connector kinds. Stable top-level insertion and both EMU axes are bounded.
Straight line/arrow kinds require the canonical 114,300 EMU grab height; invalid Agent input fails
instead of being silently normalized, while retained draw gestures still converge on that final
state. Registry, Gallery, and draw mode share one line node builder. Native Undo and save/reopen
are covered. The generated Manifest contains 187 operations, including sixty-three DOCX
operations (sixty Agent-visible and three internal).

R2-182 adds `docx.textbox.insert { afterBlockIndex, widthEmu, heightEmu }`. Both axes are bounded
to 9,525–20,000,000 EMU and the insertion boundary is revision-scoped. The retained Ribbon and
Registry share the format-owned textbox node builder instead of constructing parallel XML/display
models. One native Undo restores the prior document and save/reopen preserves the requested
dimensions. The generated Manifest contains 188 operations, including sixty-four DOCX operations
(sixty-one Agent-visible and three internal).

R2-183 adds
`docx.chart.insert { afterBlockIndex, kind, title, categories, series, widthPx, heightPx }` for
retained bar, line, and pie charts. Categories are bounded to 1–256, series to 1–64, the complete
matrix to 4,096 values, and numeric values to ±10^12; every series must match the category count
and pie charts require exactly one series. Explicit 120–660 px width and 80–4,096 px height remove
hidden dialog defaults. The Chart dialog and Registry share one format-owned node builder, one
native Undo unit, and embedded chart/workbook save/reopen projection. The generated Manifest
contains 189 operations, including sixty-five DOCX operations (sixty-two Agent-visible and three
internal).

R2-184 adds
`docx.equation.insert { placement, latex, afterBlockIndex, from, to }` for both retained display
and inline equation producers. `placement: block` requires one stable block boundary and null text
coordinates; `placement: inline` requires a null block boundary and one exact same-paragraph range.
Operation Contract now validates Unicode-aware `minLength`/`maxLength`, bounding LaTeX to 1–4,096
characters before Broker dispatch; the renderer trims, checks coordinates, and parses LaTeX before
mutating. Gallery, modal, and Registry share the same LaTeX-to-OMML action, native Undo, and
save/reopen route. The generated Manifest contains 190 operations, including sixty-six DOCX
operations (sixty-three Agent-visible and three internal).

R2-185 adds `docx.object.set_size { objectBlockIndex, widthPx, heightPx }` as one aggregate route
for retained shape, line, textbox/WordArt-like textbox, and chart corner resizing. General object
axes are bounded to 24–4,096 px by 8–4,096 px; charts retain their narrower 120–660 px by
80–4,096 px domain, and straight lines retain their fixed grab height. The UI drag plugin and
Registry share one position-based node mutation kernel. Every changed size explicitly starts a new
native history group, and generated-shape plus original-object save projection is covered. The
generated Manifest contains 191 operations, including sixty-seven DOCX operations (sixty-four
Agent-visible and three internal).

R2-186 adds
`docx.object.set_offset_position { objectBlockIndex, wrap, offsetXEmu, offsetYEmu }` for every
textbox-backed drawing, including shape, line, text box/WordArt-like box, and multi-box diagram.
Wrap is explicit and finite; both offsets are signed 32-bit EMU values. Shape draw, retained move
handles, and Registry share one position kernel, while images stay on their closed image lifecycle.
The OOXML writer/parser now preserves left/right square/tight/through semantics at numeric offsets
through `wrapText` without rewriting retained wrap polygons. Native Undo and negative-offset
save/reopen are covered. The generated Manifest contains 192 operations, including sixty-eight
DOCX operations (sixty-five Agent-visible and three internal).

R2-187 adds
`docx.object.set_style { objectBlockIndex, style, fields }` as one aggregate masked route for shape,
line, textbox/WordArt-like textbox, and diagram fill/outline state. `fields` must uniquely and
exactly name the supplied nullable `fillHex`/`borderHex` properties; colors are uppercase six-digit
hex values, and stroke-only lines reject fill changes. Operation Contract now supports bounded
string `pattern` validation. Shape Format and Registry share one protected-node mutation kernel,
changed calls form one native Undo unit, and fill/outline removal survives save/reopen. The generated
Manifest contains 193 operations, including sixty-nine DOCX operations (sixty-six Agent-visible and
three internal).

R2-188 adds `docx.object.remove { objectBlockIndex }` as one exact lifecycle route for shape, line,
textbox/WordArt-like textbox, chart, diagram, and block equation nodes. Images are rejected so their
separate closed lifecycle remains authoritative; generic protected fields are not broadened into the
object family. Object-mode Backspace/Delete and Registry share the same position-level transaction,
the sole-block case inserts an empty paragraph, and changed deletion starts an independent native
Undo unit. Real DOM keydown, chart removal, image rejection, and save/reopen are covered. The
generated Manifest contains 194 operations, including seventy DOCX operations (sixty-seven
Agent-visible and three internal).

R2-189 adds `docx.chart.update { chartBlockIndex, patch, fields }` as one masked final-state route
for existing title, category, series-name, and numeric cache slots. Field masks exactly match the
payload; title/category/series dimensions stay fixed, absent names and null cache gaps stay
read-only, strings are bounded to 512 characters, and finite values stay within ±10^12. The
protected chart grid and Registry share one transaction kernel with independent native Undo.
Original charts retain surgical cache-part patches while generated charts regenerate their chart
and workbook from the updated display model. Save/reopen is covered. The generated Manifest
contains 195 operations, including seventy-one DOCX operations (sixty-eight Agent-visible and three
internal); retained chart lifecycle is closed through aggregate object size/removal.

R2-190 adds
`docx.equation.update { placement, mode, latex, tokens, equationBlockIndex, from, to }` as one
bounded dual-mode route rather than separate LaTeX and token tools. `latex` rebuilds one exact
block equation or inline atom; `tokens` preserves the existing block OMML shape and requires the
same token count. Nullable payload and target fields are mutually exclusive at runtime, aggregate
token text is capped at 4096 Unicode characters, and the retained EquationModal plus protected
token editor share the same close-history transaction kernel with Registry. Block/inline execution,
invalid mode/shape rejection, native Undo, and both generated and structure-preserving save/reopen
are covered. The generated Manifest contains 196 operations, including seventy-two DOCX operations
(sixty-nine Agent-visible and three internal).

R2-191 adds
`docx.textbox.set_content { objectBlockIndex, textboxIndex, paragraphs, heightPx }`. The two-level
identity covers plain textboxes, textbox-backed shapes/WordArt, and multi-box retained drawings
without separate content tools. Paragraphs and runs have per-item plus aggregate bounds; common
paragraph geometry/alignment and rich run marks are explicit, uppercase colors are constrained,
and `heightPx` either records the UI's final auto-grown height or preserves current autofit/fixed
state. Flattened table/content-control boxes fail closed. Registry and all nested TipTap sub-editors
share one batch-capable close-history writer, so a multi-box UI commit is one native Undo unit.
Rich execution, invalid read-only/aggregate inputs, UI Undo, and original OOXML save/reopen are
covered. The generated Manifest contains 197 operations, including seventy-three DOCX operations
(seventy Agent-visible and three internal); the retained shape/textbox/equation content row is
closed.

R2-192 adds `docx.text.set_link { range, href, text }` as one nullable final-state route instead of
separate insert/update/remove tools. Non-null `text` replaces the exact range (including an empty
insertion point) before applying `href`; null `text` preserves a non-empty text-only range; null
`href` removes links and requires null `text`. Hrefs and replacement text are Unicode-bounded, and
ranges remain inside one text-bearing block. LinkInsertModal/keyboard/context-menu entry and
Registry share one close-history action; external relationships allocate through the existing
surgical save path and reopen as linked runs. The superseded whole-block `link` field is removed
from public `docx.text.set_style`, leaving one canonical link writer. The generated Manifest
contains 198 operations, including seventy-four DOCX operations (seventy-one Agent-visible and
three internal).

R2-193 adds `docx.bookmark.set { blockIndex, name, enabled }`. The revision-scoped top-level target
must be a paragraph, heading, or list item; names use the retained letter/underscore/CJK grammar,
are Unicode-bounded to 40 characters, and remain unique across the document. Explicit enabled state
makes add/remove replayable and idempotent. BookmarkModal and Registry share one close-history
action; duplicate/invalid targets fail before mutation, native Undo is single-step, and the existing
OOXML bookmark start/end projection reopens on the same block. The generated Manifest contains 199
operations, including seventy-five DOCX operations (seventy-two Agent-visible and three internal).

R2-194 adds `docx.cross_reference.insert { range, bookmarkName, displayText }`. The exact range must
remain inside one text-bearing block, the bounded named bookmark must exist, and explicit cached
display text makes the inserted REF field deterministic. CrossRefModal and Registry share one
close-history action; invalid/missing targets fail before mutation, native Undo is single-step, and
the retained REF OOXML projection reopens with the same bookmark and text. The generated Manifest
contains 200 operations, including seventy-six DOCX operations (seventy-three Agent-visible and
three internal).

R2-195 adds `docx.field.insert { range, instruction, displayText }`. The instruction is limited to
the retained DATE, TIME, PAGE, NUMPAGES, and FILENAME menu, the exact range remains inside one
text-bearing block, and explicit cached display text makes time/page/file-dependent insertion
deterministic. App/Ribbon and Registry share the field action with native single-step Undo and
generic-field OOXML save/reopen evidence. The generated Manifest contains 201 operations, including
seventy-seven DOCX operations (seventy-four Agent-visible and three internal).

R2-196 adds aggregate `docx.field.update { updates }`. Each 1–1024 entry carries an exact field
range, the bounded original instruction (including retained Word switches), and explicit final
cached text; runtime permits the seven field keywords the existing F9 evaluator understands and
caps aggregate display text at 65,536 Unicode characters. Every target must exactly match one
current `instrField` mark before any mutation. App F9/context-menu and Registry share one descending
close-history transaction; multi-field Undo and generic-field save/reopen are covered. The generated
Manifest contains 202 operations, including seventy-eight DOCX operations (seventy-five
Agent-visible and three internal).

R2-174 adds `docx.image.set_wrap { imageBlockIndex, wrap }`. The finite retained values are
`square-left`, `square-right`, `topBottom`, `behind`, `front`, and `null`; `null` is explicit inline
state and clears named/offset positioning, while a non-null change preserves current position.
Agent execution and all three retained ContextMenu/Layout/Picture Format producers share one
exact-image attrs kernel and native Undo. The existing image surgical-save route reopens the chosen
wrap exactly. The generated Manifest contains 180 operations, including fifty-six DOCX operations
(fifty-three Agent-visible and three internal).

R2-175 adds
`docx.image.set_margin_position { imageBlockIndex, horizontal, vertical }` for the retained 3×3
Word gallery. Horizontal is `left|center|right`, vertical is `top|center|bottom`; the shared writer
sets the matching square wrap, named positions, and clears free offsets in one Undo transaction.
Agent and Layout gallery share that writer. Because the current generated-image save model has no
named margin position, both routes expose the retained original-image-only guard and fail rather
than lose state. Save/reopen and clean resave are covered. The generated Manifest contains 181
operations, including fifty-seven DOCX operations (fifty-four Agent-visible and three internal).

R2-176 adds
`docx.image.set_offset_position { imageBlockIndex, wrap, offsetXEmu, offsetYEmu }` for explicit
floating-image placement. The finite wrap enum preserves square, tight, through, top/bottom,
behind, and front modes; both signed EMU axes are bounded to 32-bit integers. The operation clears
named positions and writes the final offsets in one native Undo transaction for original and
generated images. Registry execution and the retained image-drag path share the same pure attrs
projection, while textbox/shape dragging remains object-owned. Save/reopen of an original image is
covered. The generated Manifest contains 182 operations, including fifty-eight DOCX operations
(fifty-five Agent-visible and three internal).

R2-177 adds
`docx.image.set_transform { imageBlockIndex, rotationDegrees, flipHorizontal, flipVertical }`.
Rotation is an explicit integer from 0 through 359 and both mirror axes are required booleans, so
the full state is replayable instead of encoding Ribbon deltas/toggles. Registry execution and the
retained Rotate/Flip buttons share one canonical attrs projection; zero rotation becomes the
absence of `rot`. The surgical saver now materializes a minimal `pic:spPr/a:xfrm` when a valid
source image omitted it, closing a prior save-loss gap. Original/generated images, native Undo,
and original-image save/reopen are covered. The generated Manifest contains 183 operations,
including fifty-nine DOCX operations (fifty-six Agent-visible and three internal).

R2-178 adds
`docx.image.set_crop { imageBlockIndex, left, top, right, bottom }`. Each source inset is a bounded
`0..0.99` fraction and cross-field validation requires positive remaining width and height; all
zeros explicitly reset cropping. Registry execution and the retained Crop dialog now share a
non-destructive attrs writer, clear stale fill windows, and preserve the source bytes. Original
image patches and newly embedded images write `a:srcRect`. Original/generated nodes, native Undo,
and original-image save/reopen are covered. Cutout remains on the R2-173 replacement kernel. The
generated Manifest contains 184 operations, including sixty DOCX operations (fifty-seven
Agent-visible and three internal).
