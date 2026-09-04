# TandemFolio operation catalog

Always page `office_get_capabilities` in its default summary view to discover canonical ids, then
request `view: "detail"` for the one selected id and use that returned schema. A summary is
schema-free by design; this page explains intent and grounding, not the executable contract.
Execute the selected canonical id with a fresh caller `requestId` and a one-item
`operations: [{ id, arguments }]` array. After a caller timeout, replay that exact request id and
payload; do not create a second mutation while its final acknowledgement is unknown.

This is the current 337-operation catalog projection. All five retained state-changing producer
baselines have typed mappings and the R6-01 shared release gate passes, so `ready` is true. An
operation absent from summary discovery is unavailable and must not be inferred from visible UI.
Supplying the current `sessionId` also returns stable availability without hiding the operation.

Every canonical `*.document.save` runs the format-owned serializer. In Codex, success atomically
commits the renderer-produced bytes through the Session-bound local protocol; verify the resulting
absolute `office_get_context.session.filePath`. New documents use a Session-isolated directory under
`TANDEMFOLIO_OUTPUT_DIR` or `~/Documents/TandemFolio`. Opening a local file binds Save to that
exact path, while format Save As flows create and bind a renamed copy.

## DOCX

- `docx.block.delete { target }`: delete matching blocks as one native undo transaction. Uses the same target fields as heading-level mutation, returns matched/changed/protected/tracked-deleted counts, and leaves one empty paragraph if every block is deleted.
- `docx.block.move { blockIndexes, afterBlockIndex }`: move one or more non-negative block indexes after another block (`-1` means document start), preserving their document order in one native undo transaction.
- `docx.image.insert { path, afterBlockIndex, widthPx, heightPx, alignment }`: insert an absolute local PNG/JPEG/GIF path after a stable top-level block. Axes are explicit 1–4096px values; the Broker transports at most 20 MiB through a hidden session blob and the renderer inserts one native Undo-owned image node.
- `docx.image.remove { imageBlockIndex }`: remove one exact top-level image. Removing the sole block inserts an empty paragraph in the same native Undo transaction.
- `docx.image.replace { path, imageBlockIndex, widthPx, heightPx }`: replace one indexed native image with an absolute local PNG/JPEG/GIF and explicit final geometry. Placement/wrap/format state is preserved, stale crop windows are cleared, and native Undo restores the prior bytes.
- `docx.image.set_crop { imageBlockIndex, left, top, right, bottom }`: set non-destructive source crop insets as bounded fractions; opposing pairs must leave positive area. Four zeros reset the crop.
- `docx.image.set_margin_position { imageBlockIndex, horizontal, vertical }`: set one saved original image to one of nine named margin-relative positions. This derives square wrap, clears free offsets, and rejects unsaved generated images whose current save model cannot encode named positioning.
- `docx.image.set_offset_position { imageBlockIndex, wrap, offsetXEmu, offsetYEmu }`: set one original or generated image to an explicit retained floating wrap and bounded signed EMU offsets. This clears named margin positions and is directly replayable.
- `docx.image.set_transform { imageBlockIndex, rotationDegrees, flipHorizontal, flipVertical }`: set the complete final image transform. Rotation is an integer from 0 through 359 and both flip axes are explicit booleans.
- `docx.image.set_wrap { imageBlockIndex, wrap }`: set one indexed image to `square-left`, `square-right`, `topBottom`, `behind`, `front`, or explicit inline `null`. Inline clears stale named/offset positioning; floating changes preserve it.
- `docx.image.update { target, properties, fields }`: update matching image width, height, or alignment. Supplying one dimension preserves the original aspect ratio.
- `docx.shape.insert { afterBlockIndex, preset, widthEmu, heightEmu }`: insert one of 104 retained filled shape presets after a stable block. Dimensions are explicit bounded EMU values; line/connector presets use a separate operation family.
- `docx.line.insert { afterBlockIndex, kind, widthEmu, heightEmu }`: insert one of five retained stroke-only line/connector kinds. Straight kinds require the canonical 114,300 EMU grab height.
- `docx.textbox.insert { afterBlockIndex, widthEmu, heightEmu }`: insert one empty textbox after a stable block with explicit bounded EMU dimensions.
- `docx.chart.insert { afterBlockIndex, kind, title, categories, series, widthPx, heightPx }`: insert a bounded bar, line, or single-series pie chart with an explicit final extent; all series must match the category count.
- `docx.chart.update { chartBlockIndex, patch, fields }`: update masked final title/category/series cache slots on one exact chart while preserving matrix shape, read-only gaps, chart kind, and geometry.
- `docx.equation.insert { placement, latex, afterBlockIndex, from, to }`: insert 1–4,096-character LaTeX as a display block at a stable boundary or as an inline atom replacing one exact same-paragraph range; unused coordinates must be `null`.
- `docx.object.set_size { objectBlockIndex, widthPx, heightPx }`: set one exact shape, line, textbox/WordArt-like box, or chart to a bounded final size; charts and straight lines retain their narrower invariants.
- `docx.object.set_offset_position { objectBlockIndex, wrap, offsetXEmu, offsetYEmu }`: set one exact shape, line, textbox/WordArt-like box, or diagram to a finite floating wrap and signed 32-bit EMU offsets; images use the separate image operation.
- `docx.object.set_style { objectBlockIndex, style, fields }`: set masked nullable fill/outline colors on one exact shape, line, textbox/WordArt-like box, or diagram; uppercase six-digit hex is required and lines reject fill.
- `docx.object.remove { objectBlockIndex }`: remove one exact shape, line, textbox/WordArt-like box, chart, diagram, or block equation; images use their separate lifecycle and the sole-block case inserts a paragraph.
- `docx.list.apply { target, kind }`: apply `bullet` or `ordered` list formatting as one native undo transaction. Reuses the DOCX block target fields, preserves text, and returns matched/changed/protected counts. A same-kind target is left unchanged.
- `docx.list.remove { target }`: remove list formatting from matching list items as one native undo transaction. Reuses the DOCX block target fields, preserves item text, and returns matched/changed/protected counts.
- `docx.list.set_level { target, level }`: set matching DOCX list items to an absolute `0..8` level as one native Undo unit. Non-list matches are left unchanged.
- `docx.list.apply_preset { target, levels }`: create a bounded 1–9-level numbering definition and apply its new `numId` to matching paragraphs/list items as one native Undo unit. Each level declares an explicit format, level text, indentation, and optional hanging/start values.
- `docx.list.restart { blockIndex, start }`: restart the list anchored at one stable top-level block from an explicit `1..1,000,000` value. Later items sharing the source list identity move with the anchor in one native Undo unit.
- `docx.list.continue { blockIndex, previousBlockIndex }`: continue the list at one stable block from an explicitly identified earlier list block. Only later items sharing the current list identity are rebound, in one native Undo unit.
- `docx.history.undo {}` / `docx.history.redo {}`: undo or redo one available entry in the mounted TipTap history shared with the retained DOCX UI. An unavailable entry returns `execution_failed` without advancing revision.
- `docx.text.insert { text }`: insert at active caret/selection. The server still accepts legacy `insert_text` as an input-only DOCX compatibility alias.
- `docx.text.replace_selection { text }`: replace active selection. The server still accepts legacy `replace_selection` as an input-only DOCX compatibility alias.
- `docx.text.replace_all { containsText, replaceText, matchCase? }`: replace matching text across the active document as one native undo transaction. Matching is case-sensitive unless `matchCase` is `false`; success returns a summary plus matched/changed block and protected/tracked-deleted skip counts.
- `docx.text.clear_character_format { range }`: apply the retained Clear Formatting behavior to one non-empty revision-scoped ProseMirror range as a single native undo transaction; text and marks outside the range remain unchanged.
- `docx.text.set_character_format { range, format, fields }`: set exact bold, italic, underline, strike, baseline/superscript/subscript, font family, 1–1638pt half-point font size, `#RRGGBB` color, and named DOCX highlight final state over a non-empty revision-scoped ProseMirror text range. Nullable font/color/size/highlight fields clear their property. `fields` must uniquely and exactly match `format`; unlisted per-run style is preserved.
- `docx.text.set_character_style { range, styleId }`: apply a character style that exists in the current DOCX, or pass `null` to remove the character-style mark with the retained Ribbon semantics. Direct format survives application; the edit is one native Undo unit.
- `docx.text.transform_case { range, mode }`: transform one non-empty revision-scoped text range using `sentence`, `lower`, `upper`, or `title` case while preserving run marks and mapping Unicode length changes in one native undo transaction.
- `docx.text.set_style { target, style, fields }`: apply or clear only the named text-style fields while preserving unlisted formatting.
- `docx.paragraph.set_heading_level { target, level }`: promote matching blocks to heading levels 1–6 or demote them to body level 0 as one native undo transaction. `target` supports `nodeType`, `headingLevel`, `containsText`, `matchCase`, non-negative `blockIndexes`, and `scope`; at least one targeting condition is required.
- `docx.paragraph.set_direction { target, direction }`: set matching paragraph-like blocks to `ltr` or `rtl`; explicit left/right alignment flips through the same rule as the retained Ribbon so logical start/end is preserved in one native Undo unit.
- `docx.paragraph.set_style { target, style, fields }`: apply or clear only the named paragraph-style fields as one native undo transaction, including bounded line spacing, geometry, alignment, page-break, shading/borders, and a 1–64-item final tab-stop array.
- `docx.toc.insert { afterBlockIndex }`: insert real TOC field blocks generated from the mounted document's current headings; `-1` inserts at document start.
- `docx.table.insert { afterBlockIndex, rows, columns }`: insert a bounded native top-level table after one stable block (`-1` means document start); axes are capped at 100×63 and total cells at 4096.
- `docx.table.delete { tableBlockIndex }`: delete one explicitly indexed top-level native table as a single Undo unit; a sole-table document receives a valid empty paragraph.
- `docx.table.insert_rows { tableBlockIndex, rowIndex, count }`: insert 1–100 rows at an explicit native table boundary; rowspan-aware TableMap updates and the 4096-resulting-cell budget remain one Undo unit.
- `docx.table.delete_rows { tableBlockIndex, rowIndex, count }`: delete 1–100 rows from an explicit native table index in one rowspan-aware Undo unit; deleting every row is rejected in favor of `docx.table.delete`.
- `docx.table.insert_columns { tableBlockIndex, columnIndex, count }`: insert 1–63 columns at an explicit native table boundary; final width and logical-cell budgets are renderer-validated in one span-aware Undo unit.
- `docx.table.delete_columns { tableBlockIndex, columnIndex, count }`: delete 1–63 columns from an explicit logical index in one colspan-aware Undo unit; deleting every column is rejected in favor of `docx.table.delete`.
- `docx.table.merge_cells { tableBlockIndex, topRow, leftColumn, bottomRow, rightColumn }`: merge one exact half-open logical-cell rectangle through the native table command; existing spans may not cross its boundary.
- `docx.table.split_cell { tableBlockIndex, rowIndex, columnIndex }`: split the merged cell covering one bounded logical coordinate through the native table command; ordinary cells fail without mutation.
- `docx.table.set_cell_format { tableBlockIndex, topRow, leftColumn, bottomRow, rightColumn, format, fields }`: set masked fill and vertical-alignment final state over one exact half-open cell rectangle; Ribbon and Agent share the same write kernel.
- `docx.table.set_cell_borders { tableBlockIndex, topRow, leftColumn, bottomRow, rightColumn, mode, border }`: apply all/outer/inner/none over one exact rectangle with bounded color and eighth-point width; `none` requires `border: null`.
- `docx.table.set_style { tableBlockIndex, styleId }`: set one current-document table style by stable top-level table identity; `null` clears the style and unknown IDs fail closed.
- `docx.table.set_row_height { tableBlockIndex, rowIndex, count, heightTwips }`: set 1–100 explicit physical rows to a 1–31,680 twip height; `null` restores automatic height.
- `docx.table.set_column_widths { tableBlockIndex, widthsPx }`: replace the complete 1–63-column grid with bounded 40–4096px widths (total at most 4096px), synchronizing cells, table width, and percentages.
- `docx.document.insert_page_break { afterBlockIndex }`: insert one native page-break paragraph after a stable top-level block boundary (`-1` means document start).
- `docx.section.insert_break { afterBlockIndex, startType }`: insert an Undo-owned section break using `nextPage`, `continuous`, `evenPage`, or `oddPage`.
- `docx.section.set_orientation { sectionIndex, orientation }`: set exact portrait/landscape state while swapping the current page axes.
- `docx.section.set_margins { sectionIndex, margins }`: set four explicit bounded twip margins that leave positive body width and height.
- `docx.section.set_page_size { sectionIndex, widthTwips, heightTwips }`: set exact bounded page axes and derive orientation from their final relationship.
- `docx.section.set_columns { sectionIndex, count, spacingTwips }`: set exact bounded column count and gap while preserving positive text width.
- `docx.section.set_page_border { sectionIndex, enabled }`: set explicit page-border state on one indexed section.
- `docx.section.set_different_first_page { sectionIndex, enabled }`: set indexed different-first-page state and `w:titlePg` persistence.
- `docx.document.set_different_odd_even_pages { enabled }`: set document-wide odd/even header-footer variants and `w:evenAndOddHeaders` persistence.
- `docx.section.set_page_numbering { sectionIndex, format, start }`: set indexed page-number format and nullable start state.
- `docx.header_footer.set_text { sectionIndex, kind, variant, text }`: set explicit header/footer variant plain text through native Undo.
- `docx.header_footer.set_page_number { sectionIndex, kind, variant, enabled, alignment }`: set or remove canonical PAGE-field placement.
- `docx.header_footer.set_paragraphs { sectionIndex, kind, variant, paragraphs }`: set bounded rich content with styled text and PAGE/NUMPAGES tokens.
- `docx.document.save {}`: run the DOCX surgical save path; success returns `{ saved: true, fileName }`, commits the local file, and clears the stored recovery snapshot.

`office_open_local_file` routes DOCX bytes internally as `docx.document.load_staged`. Public
`docx.image.insert` and `docx.image.replace` are similarly transformed into their internal staged
operations after the Broker validates and stages the local image. These internal operations and
the `open_local_file` transport alias are not advertised by capability discovery and cannot be
called through `office_execute`.

Additional retained DOCX families:

- Links, bookmarks, and references: `docx.text.set_link`, `docx.bookmark.set`,
  `docx.cross_reference.insert`, `docx.caption.insert`, `docx.index.insert`,
  `docx.index.mark`, `docx.toc.refresh`.
- Sources and citations: `docx.source.upsert`, `docx.citation.insert`,
  `docx.bibliography.insert`.
- Comments, notes, and revisions: `docx.comment.add`, `docx.comment.reply`,
  `docx.comment.delete`, `docx.comment.set_resolved`, `docx.note.insert`,
  `docx.note.update`, `docx.note.delete`, `docx.revision.set_tracking`,
  `docx.revision.apply_decision`, `docx.document.compare`.
- Fields and rich objects: `docx.field.insert`, `docx.field.update`,
  `docx.equation.update`, `docx.textbox.set_content`, `docx.wordart.insert`,
  `docx.ink.apply`.
- Document presentation and protection: `docx.cover_page.insert`,
  `docx.paragraph.set_drop_cap`, `docx.document.set_design`,
  `docx.document.set_protection`.

Context includes the active selection plus bounded structural/document summaries. The 99 public
DOCX operations and three internal staged-byte routes cover the retained state-changing producer
baseline. DOCX is included in the passing R6-01 readiness projection.

## Markdown

- Text and selection: `markdown.text.insert`, `markdown.text.replace_selection`,
  `markdown.text.set_marks`, `markdown.selection.set`.
- Document/history: `markdown.document.save`, `markdown.document.save_as`,
  `markdown.document.export_docx`, `markdown.document.open_print_dialog`,
  `markdown.document.set_auto_save`, `markdown.history.undo`, `markdown.history.redo`.
- Blocks/structure: `markdown.block.set_type`, `markdown.block.update`,
  `markdown.list.set_type`, `markdown.divider.insert`, `markdown.table.insert`,
  `markdown.table.update`, `markdown.code_block.set_language`.
- Metadata/media: `markdown.frontmatter.set`, `markdown.image.insert`.

Context includes the active selection, bounded selected text, and active block type. YAML
frontmatter, original line endings, UTF-8 BOM state, trailing-newline behavior, and local images are
preserved by the format-owned round-trip layer. The 20 public and two internal Markdown operations
cover its retained producer baseline. Markdown is included in the passing R6-01 readiness projection.

## XLSX

- `xlsx.cell.set_value { sheet, address, value }`: set one scalar cell. Legacy `set_cell_value` is an input-only alias.
- `xlsx.range.set_values { sheet, range, values }`: set one bounded non-empty scalar matrix. Legacy `set_range_values` is an input-only alias.
- `xlsx.range.set_text_style { sheet, range, style, fields }`: explicitly set or clear bold, italic, strike, and `none | single | double` underline fields as one native range mutation. `fields` must uniquely and exactly name the supplied style values.
- `xlsx.range.set_alignment { sheet, range, alignment, fields }`: explicitly set or clear horizontal/vertical alignment, wrapping, bounded indent, and angle/stacked rotation fields. The unique `fields` mask must exactly match `alignment`.
- `xlsx.range.set_font { sheet, range, font, fields }`: explicitly set or clear font family, size (1–409), and `#RRGGBB` color fields. The unique `fields` mask must exactly match `font`.
- `xlsx.range.set_fill { sheet, range, color }`: set a `#RRGGBB` fill or clear it with `null`.
- `xlsx.range.set_border { sheet, range, border }`: set one bounded border preset with an explicit line style and color, or clear all borders with `{ preset: "none" }`.
- `xlsx.range.apply_cell_style { sheet, range, preset }`: apply one retained built-in named cell-style preset as a single native mutation.
- `xlsx.range.set_number_format { sheet, range, pattern }`: assign one explicit 1–255-character number-format pattern. Use the intended final pattern instead of relative decimal increment/decrement gestures.
- `xlsx.range.merge { sheet, range, mode }`: merge `cells`, merge each row `across`, merge then `center`, or `unmerge`. Center retains the visible UI's merge-then-align sequence and is non-atomic.
- `xlsx.range.clear { sheet, range, scope }`: clear `contents`, `formats`, or `all` through the mounted native range.
- `xlsx.range.copy_values { sourceSheet, sourceRange, destinationSheet, destinationRange }`: copy computed scalar values, not formulas, between equal-shape worksheet-bounded A1 ranges on the same or different worksheets. Each operation is limited to 20,000 cells, fails closed while either file-backed rectangle is still streaming, and writes one native undoable destination matrix.
- `xlsx.range.copy_formulas { sourceSheet, sourceRange, destinationSheet, destinationRange }`: copy formulas and scalar cells between equal-shape worksheet-bounded A1 ranges on the same or different worksheets. Formula references are translated by the mounted Univer lexer using the exact row/column offset; the operation shares the 20,000-cell limit, streaming guards, and one native undoable destination matrix with value copy.
- `xlsx.range.copy_formats { sourceSheet, sourceRange, destinationSheet, destinationRange }`: replace destination cell formats from an equal-shape source range without changing values or formulas. The operation supports cross-sheet ranges, shares the 20,000-cell and streaming guards, and commits one native undoable style matrix; merged-cell topology remains an explicit `xlsx.range.merge` concern.
- `xlsx.range.fill { sheet, range, direction }`: fill `down` or `right` from the leading row/column of an explicit target containing at least one destination row/column.
- `xlsx.range.sort { sheet, range, direction }`: sort an explicit range `asc` or `desc` by its first column through Univer's native sort command.
- `xlsx.range.sort_custom { sheet, range, keys, hasHeader }`: sort by ordered, unique A1 column keys inside the target; every key carries an explicit `asc | desc` direction.
- `xlsx.range.remove_duplicates { sheet, range, hasHeader }`: remove case-insensitive duplicate rows while keeping the first occurrence; returns `removed` and requires fully loaded source data.
- `xlsx.range.set_filter { sheet, range, enabled }`: set the explicit final AutoFilter state for a range through native filter commands and shared Undo. Repeating an already-satisfied state is a no-op; a different active filter range fails closed.
- `xlsx.range.clear_filter_criteria { sheet, range }`: clear every column criterion while retaining the exact active AutoFilter range.
- `xlsx.range.set_filter_values { sheet, range, column, values, includeBlank }`: replace one absolute filter column with a bounded 1–10,000 item value list and explicit blank inclusion.
- `xlsx.range.set_custom_filter { sheet, range, column, conjunction, conditions }`: replace one absolute filter column with one or two custom conditions joined by explicit `and` or `or`; operators are `equal`, `notEqual`, `greaterThan`, `greaterThanOrEqual`, `lessThan`, and `lessThanOrEqual`.
- `xlsx.formula.insert_aggregate { sheet, range, function }`: insert `SUM`, `AVERAGE`, `COUNT`, `MAX`, or `MIN` below every selected column; requires at least two source rows and rejects an unstreamed destination row before writing.
- `xlsx.history.undo {}` / `xlsx.history.redo {}`: undo or redo one available entry in the mounted Univer workbook history shared with visible UI controls. An unavailable entry returns `execution_failed` without advancing revision.
- `xlsx.range.flash_fill { sheet, range }`: infer a concatenation pattern from retained target examples and fill only empty target cells. Single-cell targets probe the adjacent left column for at most 1,000 rows, and partially streamed source rows are rejected before reading.
- `xlsx.range.text_to_columns { sheet, range, delimiter }`: split one fully loaded text column with `tab`, `comma`, `semicolon`, or `space`. The native command is undoable but high-risk because split output can overwrite cells to the right or add columns.
- `xlsx.row.set_height { sheet, row, count, heightPoints }`: set a bounded 1-based contiguous row span to a final `0–409.5` point height through the native undo and OOXML row-size journal.
- `xlsx.column.set_width { sheet, column, count, widthCharacters }`: set a bounded A1 column span to a requested `0–255` character width; the result reports the actual 1/256-character width after native pixel quantization.
- `xlsx.column.copy_widths { sourceSheet, sourceColumn, destinationSheet, destinationColumn, count }`: copy a bounded 1–10,000-column width vector between explicit A1 column spans on the same or different worksheets. Source and destination spans must remain within their worksheets; one native mutation owns Undo/Redo and saved OOXML while cell contents remain unchanged.
- `xlsx.sheet.set_freeze { sheet, frozenRows, frozenColumns }`: set explicit frozen row/column counts through native undo; `0,0` removes frozen panes.
- `xlsx.sheet.set_gridlines { sheet, visible }`: set final worksheet gridline visibility through native undo and the persisted sheet-view journal.
- `xlsx.sheet.set_formula_view { sheet, enabled }`: set final per-sheet formula view through renderer-owned undo and the persisted `sheetView@showFormulas` journal.
- `xlsx.sheet.set_fit_to_pages { sheet, widthPages, heightPages }`: set both `0–1000` print-fit axes explicitly as one renderer-owned undo unit; `0` means Automatic and `0,0` disables fit-to-page.
- `xlsx.sheet.set_page_margins { sheet, margins }`: set final `normal`, `wide`, or `narrow` print margins through renderer-owned undo and the persisted page-setup journal.
- `xlsx.sheet.set_page_orientation { sheet, orientation }`: set final `portrait` or `landscape` print orientation through renderer-owned undo and the persisted page-setup journal.
- `xlsx.sheet.set_paper_size { sheet, paperSize }`: set a final visible paper preset through renderer-owned undo and the persisted page-setup journal. Supported OOXML codes are Letter `1`, Tabloid `3`, Legal `5`, Executive `7`, A3 `8`, A4 `9`, and A5 `11`.
- `xlsx.sheet.set_print_gridlines { sheet, enabled }`: set final printed-gridline output through renderer-owned undo and the persisted `printOptions@gridLines` journal; this is distinct from worksheet display gridlines.
- `xlsx.sheet.set_print_headings { sheet, enabled }`: set whether row and column headings print through renderer-owned undo and the persisted `printOptions@headings` journal.
- `xlsx.sheet.set_print_scale { sheet, scalePercent }`: set a fixed integer print scale from `10` through `400`; the same renderer-owned undo unit disables fit-to-page and persists `pageSetup@scale`.
- `xlsx.sheet.set_print_area { sheet, range }`: set a normalized explicit A1 cell range as the worksheet print area, or pass `null` to clear it, through renderer-owned undo and the persisted `_xlnm.Print_Area` defined name.
- `xlsx.sheet.set_print_titles { sheet, rows }`: set an ascending explicit row span such as `"1:2"` as repeated print-title rows, or pass `null` to clear it. Spans are limited to 21 rows and use renderer-owned undo plus the persisted `_xlnm.Print_Titles` defined name.
- `xlsx.range.set_protection { sheet, range, protection, fields }`: set explicit `locked` and/or `hidden` OOXML flags for at most 10,000 cells. The exact field mask must match supplied boolean values; this file-side route is not natively undoable.
- `xlsx.hyperlink.set { sheet, address, target }` / `xlsx.hyperlink.remove { sheet, address }`: normalize or remove one cell hyperlink while sharing the Ribbon's save journal and link appearance.
- `xlsx.table.add { sheet, range, style }`: create a new table with one of `TableStyleLight1`, `TableStyleLight9`, `TableStyleMedium2`, `TableStyleMedium4`, `TableStyleMedium7`, or `TableStyleDark2`; returns the generated table name.
- `xlsx.sparkline.add { sheet, sourceRange, targetRange, type }`: add `line`, `column`, or `stacked` row-aligned sparklines from an explicit source to a same-height one-column target. The operation is limited to 200 target cells, rejects overlap and occupied hosts, shares the retained Undo path, and survives native x14 save/reopen.
- `xlsx.outline.set_level { sheet, axis, start, count, level }`: set an absolute `0..7` level over a bounded 1-based row or column span.
- `xlsx.outline.set_detail_visibility { sheet, axis, start, count, hidden }`: set final detail visibility and the immediately following summary item's collapsed state; shares one Undo unit and survives browser save/reopen.
- `xlsx.range.set_checkbox { sheet, range, enabled }`: set or remove checkbox validation over at most 10,000 cells through native Undo and browser data-validation save/reopen.
- `xlsx.range.set_list_validation`, `xlsx.range.set_list_reference_validation`, `xlsx.range.set_comparison_validation`, `xlsx.range.set_custom_formula_validation`, `xlsx.range.set_validation_messages`, and `xlsx.range.remove_data_validation`: manage bounded validation rules and messages over explicit ranges through native history and browser save/reopen.
- `xlsx.conditional_format.set_comparison`, `xlsx.conditional_format.set_highlight`, `xlsx.conditional_format.set_statistical`, `xlsx.conditional_format.set_formula`, and `xlsx.conditional_format.set_visual`: create or update one explicitly identified, losslessly saveable Conditional Formatting rule over at most 10,000 cells. Use `ruleId: null` to create and a rule ID from fresh context to update.
- `xlsx.conditional_format.remove { sheet, ruleId }`: remove one published Conditional Formatting rule by exact ID.
- `xlsx.conditional_format.clear { sheet, scope, range }`: clear rules from one explicit range, or use `scope: "sheet"` with `range: null` to clear the worksheet.
- `xlsx.conditional_format.set_priority { sheet, ruleId, position }`: place one published rule at an explicit one-based final priority from 1 through 100.
- `xlsx.chart.add { sheet, dataRange, type, anchorCell? }`: add one `column`, `bar`, `line`, `area`, `pie`, `doughnut`, `scatter`, `radar`, or `combo` chart from at most 2,000 source cells. It shares the retained visual journal and Undo route and persists through the browser drawing writer.
- `xlsx.chart.update { chartId, title?, type?, legend?, dataLabels?, grouping?, categoryAxisTitle?, valueAxisTitle? }`: apply at least one bounded final-state property to an identified chart through retained history. Convertible update types are `column`, `bar`, `line`, `area`, `pie`, and `doughnut`.
- `xlsx.chart.remove { chartId }`: remove an identified chart through the same visual journal/history route; file-backed removal cascades its drawing relationship and chart part during save.
- `xlsx.chart.set_colors { chartId, seriesColors?, pointColors? }`: assign explicit hex colors to at most 24 series or 64 pie/doughnut points.
- `xlsx.chart.set_series { chartId, series }`: replace the complete chart series set with 1–24 bounded series of at most 1,000 finite values each; optional categories must match value length.
- `xlsx.image.add { sheet, path, anchorCell }`: insert an absolute local PNG, JPEG, or GIF at an explicit cell. Broker staging validates magic bytes and a 20 MB limit and keeps Base64 outside Agent JSON; the internal renderer operation shares the retained visual journal, Undo route, and browser media writer.
- `xlsx.note.set { sheet, address, text }` / `xlsx.note.remove { sheet, address }`: set or remove one explicitly addressed legacy cell note through Univer native history. Note text is bounded to 32,767 characters and browser save/reopen preserves comments XML plus VML anchors.
- `xlsx.row.insert { sheet, row, count }` / `xlsx.row.delete { sheet, row, count }`: edit 1-based worksheet row structure through Univer. Legacy `insert_rows` and `delete_rows` are input-only aliases.
- `xlsx.column.insert { sheet, column, count }` / `xlsx.column.delete { sheet, column, count }`: edit columns from an A1 column label. Legacy `insert_columns` and `delete_columns` are input-only aliases.
- `xlsx.sheet.add { name }`, `xlsx.sheet.rename { sheet, name }`, `xlsx.sheet.delete { sheet }`, `xlsx.sheet.move { sheet, position }`: edit worksheet tabs through the mounted workbook state. Legacy `add_sheet`, `rename_sheet`, `delete_sheet`, and `move_sheet` are input-only aliases.
- `xlsx.sheet.set_protection { sheet, protected }`: set an explicit passwordless sheet-protection state; password-protected removal fails closed.
- `xlsx.cell.set_formula { sheet, address, formula }`: set one equals-prefixed explicit formula through the mounted native formula route.
- `xlsx.range.move { sourceSheet, sourceRange, destinationSheet, destinationRange }`: move an equal-shape bounded range through Univer's native move command; source and destination must differ and Pivot-backed ranges fail closed.
- `xlsx.range.copy_without_borders { sourceSheet, sourceRange, destinationSheet, destinationRange }`: copy content, formulas, and non-border formats while preserving destination borders.
- `xlsx.range.replace_text { sheet, range, find, replace, matchCase, wholeCell }`: replace bounded text in one explicit fully loaded range.
- `xlsx.range.insert_subtotals { sheet, range, groupColumn, valueColumn, aggregation }`: insert explicit grouped subtotal rows using `sum`, `count`, `average`, `max`, or `min` over at most 50,000 cells and 200 groups.
- `xlsx.range.consolidate { targetSheet, targetCell, sources, aggregation, leftLabels }`: consolidate 1–20 explicit source ranges into one target using a bounded aggregate and optional left labels.
- `xlsx.row.set_visibility { sheet, row, count, visible }` / `xlsx.column.set_visibility { sheet, column, count, visible }`: set final row or column visibility over explicit bounded spans.
- `xlsx.row.move { sheet, row, count, beforeRow }`: move a contiguous row span to an explicit final insertion boundary through native structural history.
- `xlsx.sheet.duplicate { sheet, name }`: duplicate one named worksheet to an explicit valid name. Unsupported source-owned satellite parts and sheet-scoped names fail closed; browser save/reopen preserves the clone.
- `xlsx.sheet.set_visibility { sheet, visible }`: set an explicit final worksheet visibility while preserving the at-least-one-visible-sheet invariant.
- `xlsx.sheet.set_tab_color { sheet, color }`: set a `#RRGGBB` tab color or clear it with `null` through native history and OOXML `sheetPr` persistence.
- `xlsx.sheet.set_header_footer { sheet, header, footer }`: replace bounded nullable left/center/right header and footer sections through the shared page-setup journal.
- `xlsx.table.insert_rows`, `xlsx.table.delete_rows`, `xlsx.table.insert_columns`, `xlsx.table.delete_columns`, and `xlsx.table.convert_to_range`: edit one stable named table with bounded table-relative row/column coordinates through retained table history and save/reopen.
- `xlsx.pivot.add { sourceSheet, sourceRange, targetSheet, targetCell, name, rowFields, columnFields, pageFields, values }`: create a native PivotTable from explicit fields and bounded ranges.
- `xlsx.pivot.refresh { sheet }`: recompute retained pivots on a named sheet and flag their native caches for refresh on open.
- `xlsx.pivot.update { pivotId, targetCell, rowFields, columnFields, pageFields, values }`: replace one stable PivotTable layout with explicit bounded final fields and target.
- `xlsx.pivot.set_member_filter { pivotId, field, selectedValues }`: set a Pivot field's selected members by stable Pivot identity and exact final values.
- `xlsx.pivot.add_chart { pivotId, type }`: add a `column`, `bar`, `line`, `pie`, `doughnut`, or `radar` chart bound to one stable Pivot identity.
- `xlsx.shape.add`, `xlsx.shape.update`, and `xlsx.shape.remove`: create, update, or remove bounded shapes/text boxes by explicit sheet/anchor or stable shape identity through shared visual history.
- `xlsx.image.move { imageId, anchorCell }` / `xlsx.image.remove { imageId }`: move or remove a stable session/reopened image identity through shared visual history. Internal `xlsx.image.add_staged` carries broker-staged bytes and is not Agent-discoverable.
- `xlsx.defined_name.set { name, formula, scope?, previousName? }` / `xlsx.defined_name.remove { name, scope? }`: upsert, rename, or remove workbook/sheet-scoped Defined Names through native model history and declarative save.
- `xlsx.document.save {}`: run the retained XLSX save assembler. Standalone browsers write and reopen
  a granted file handle; Codex-embedded editors atomically commit the generated package through the
  Session-bound local protocol. Success returns `{ saved: true, fileName }` only after the active
  boundary completes, while cancellation or write failure returns `execution_failed`.

`office_open_local_file` routes `.xlsx` bytes internally as `xlsx.document.load_staged`. This
operation and its `open_local_file` transport alias are hidden from capability discovery and cannot
be called through `office_execute`; success returns `{ opened: true, fileName }` after the mounted
renderer accepts the hydrated workbook bytes.

Use stable identities and exact workbook/sheet/range/cell targets from fresh context. The 112 public
XLSX operations cover every retained state-changing command audited through R2-131; two internal
operations carry staged workbook/image bytes. XLSX passes the repository-wide packaged-host
visual, performance/resource, MCP smoke, and all-format R6-01 release gates.

## PPTX

`pptx.document.create_blank` is a whole-document replacement, not a prerequisite for editing or
generating slides in the current file. An opened, saved, restored, or edited presentation rejects
replacement unless `confirmReplace: true` records explicit user replacement intent. Discover its
current detail schema before use. Ordinary continuation edits preserve the existing Save target.

- Document/history/selection: `pptx.document.create_blank`, `pptx.document.save`,
  `pptx.document.save_as`, `pptx.history.undo`, `pptx.history.redo`, `pptx.selection.set`.
- Slides/sections: `pptx.slide.add_blank`, `pptx.slide.add_with_layout`,
  `pptx.slide.copy_to`, `pptx.slide.duplicate`, `pptx.slide.delete`, `pptx.slide.move`,
  `pptx.slide.set_layout`, `pptx.slide.set_size`, `pptx.slide.set_background`,
  `pptx.slide.set_hidden`, `pptx.slide.set_transition`, `pptx.slide.set_advance_times`,
  `pptx.slide.apply_header_footer`, `pptx.section.add`, `pptx.section.rename`,
  `pptx.section.move`, `pptx.section.remove`.
- Objects/connectors: `pptx.object.add`, `pptx.object.copy_to`, `pptx.object.delete`,
  `pptx.object.duplicate`, `pptx.object.group`, `pptx.object.ungroup`,
  `pptx.object.reorder`, `pptx.object.move_selection`, `pptx.object.set_transform`,
  `pptx.object.set_transforms`, `pptx.object.set_flip`, `pptx.object.set_fill`,
  `pptx.object.set_image_fill`, `pptx.object.set_stroke`,
  `pptx.connector.set_endpoints`.
- Text/paragraph/hyperlink: `pptx.text.set_paragraphs`, `pptx.text.replace_selection`,
  `pptx.text.replace_all`, `pptx.text.set_font`, `pptx.text.set_vertical_anchor`,
  `pptx.paragraph.set_format`, `pptx.hyperlink.set`.
- Pictures/media/ink: `pptx.image.add_bytes`, `pptx.image.replace_bytes`,
  `pptx.picture.set_crop`, `pptx.picture.set_opacity`, `pptx.media.add_bytes`,
  `pptx.model3d.add_bytes`, `pptx.ink.add`.
- Tables/charts/SmartArt: `pptx.table.add`, `pptx.table.edit_structure`,
  `pptx.table.merge_cells`, `pptx.table.set_cell_anchor`,
  `pptx.table.set_cell_content`, `pptx.table.set_column_width`,
  `pptx.table.set_row_height`, `pptx.table.set_style`, `pptx.chart.add`,
  `pptx.chart.update`, `pptx.smartart.add`.
- Design/review/master: `pptx.theme.apply`, `pptx.animation.set`, `pptx.notes.set`,
  `pptx.comment.add`, `pptx.comment.delete`, `pptx.master.object.delete`,
  `pptx.master.object.set_fill`, `pptx.master.object.set_stroke`,
  `pptx.master.object.set_transform`, `pptx.master.text.set_paragraphs`.

`office_open_local_file` routes `.pptx` bytes internally as `pptx.document.load_staged`. This
operation and its `open_local_file` transport alias are hidden from capability discovery and cannot
be called through `office_execute`; success returns `{ opened: true, fileName }` after the mounted
`BrowserPresentation` accepts the hydrated and length-validated bytes.

Context includes active slide, selected object summaries, bounding positions, text, and notes. The
73 public and one internal PPTX operations cover its retained producer baseline. `ready` remains
true through the passing shared R6-01 release gate.

## PDF

- Document: `pdf.document.set_metadata`, `pdf.document.save`.
- History: `pdf.history.undo`, `pdf.history.redo`.
- Markup and pending objects: `pdf.markup.add`, `pdf.annotation.delete_saved`,
  `pdf.pending.delete`.
- Drawings, notes, and visual signatures: `pdf.drawing.add`, `pdf.drawing.update`.
- Searchable text: `pdf.text.insert`, `pdf.text.replace`, `pdf.text.update_inserted`.
- Images and bitmap-backed static forms: `pdf.image.insert`, `pdf.image.transform`,
  `pdf.image.replace`, `pdf.image.delete`, `pdf.static_form.set`.
- AcroForms and generated document stamps: `pdf.form.set_value`, `pdf.stamp.set`.
- Pages: `pdf.page.insert`, `pdf.page.delete`, `pdf.page.reorder`,
  `pdf.page.set_rotation`.

Use exact bounded identities, page indexes, rectangles/quads, colors, and final states from the live
schema. Public page insertion accepts an absolute local PDF path; the Broker stages it into internal
`pdf.page.insert_staged`, persists the merged file, and reloads it, so the operation is not
undoable. Local open similarly uses internal `pdf.document.load_staged`. Internal operations are
hidden from discovery and rejected through direct `office_execute`.

Legacy `delete_saved_annotation`, `undo`, and `save` are input-only aliases. The 23 public and two
internal PDF operations cover the retained state-changing producer baseline. PDF remains
`ready: true` through the passing shared R6-01 release gate.
