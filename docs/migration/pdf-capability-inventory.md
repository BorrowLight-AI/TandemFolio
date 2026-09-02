# PDF pinned renderer capability inventory

- Baseline: `genspark-ai/genoffice@dc4d7e5927864498913b7ba42d0da06cc7cf628e`
- Scope: community `apps/pdf/src/renderer` and its focused tests only; enterprise `ee/` is outside the permitted source boundary
- Status: format-local retained state-changing command parity complete; R6-01 passed; PDF is release-ready

## Source evidence

The pinned renderer contains 40 files. All 33 non-AI files are present at their original
`apps/pdf/src/renderer` paths. The seven intentionally absent paths are prohibited product AI:

- `ai/AiPanel.tsx`
- `ai/pdf-skill.ts`
- `ai/tools.ts`
- `ai/transport.ts`
- `assets/send-enter-off.png`
- `assets/send-enter-on.png`
- `assets/send-stop.png`

`App.tsx`, `main.tsx`, `index.html`, `styles.css`, and `i18n/strings.ts` are host-adapted: the
community PDF UI remains, while AI, branding, Electron preload, and IPC entry points are removed.
`host/browser-pdf-api.ts` is the TandemFolio browser/file adapter and
`host/community-command-bridge.ts` registers typed commands against the single mounted App state.
The obsolete replacement renderer is absent.

Browser text and image persistence uses the pinned community PDFium implementation through
`src/domain/pdfium-browser.ts`. The WASM is gzip-compressed into the self-contained renderer at
build time. Standard PDF fonts are used for bounded ASCII text; allowlisted OFL font assets are
loaded on demand for CJK, Korean, Arabic, or extended text. A browser-safe cmap check rejects text
that the chosen font cannot encode before mutation, rather than producing missing glyphs.

## Registry and producer accounting

The PDF-owned catalog and executable registry contain 25 descriptors: 23 Agent-visible operations
and two internal staged-byte operations. `pdfRetainedProducerBaseline` classifies 25 retained
producer families as Registry, non-document host effect, or view-only. It has no `missing` entry and
every Registry mapping resolves to a catalog operation.

| Family                      | Operations                                                                                                           |
| --------------------------- | -------------------------------------------------------------------------------------------------------------------- |
| Document                    | `pdf.document.set_metadata`, `pdf.document.save`; internal `pdf.document.load_staged`                                |
| History                     | `pdf.history.undo`, `pdf.history.redo`                                                                               |
| Markup and pending edits    | `pdf.markup.add`, `pdf.annotation.delete_saved`, `pdf.pending.delete`                                                |
| Drawings, notes, signatures | `pdf.drawing.add`, `pdf.drawing.update`                                                                              |
| Text                        | `pdf.text.insert`, `pdf.text.replace`, `pdf.text.update_inserted`                                                    |
| Images and static forms     | `pdf.image.insert`, `pdf.image.transform`, `pdf.image.replace`, `pdf.image.delete`, `pdf.static_form.set`            |
| AcroForm                    | `pdf.form.set_value`                                                                                                 |
| Watermark/header/footer     | `pdf.stamp.set`                                                                                                      |
| Pages                       | `pdf.page.insert`, `pdf.page.delete`, `pdf.page.reorder`, `pdf.page.set_rotation`; internal `pdf.page.insert_staged` |

R6-09 retires the legacy public `save`, `delete_saved_annotation`, and `undo` names.
`open_local_file` remains only as an internal staged-load transport alias. Discovery and
Broker-side validation come from the generated Product Manifest; internal operations are hidden
from discovery and rejected by direct `office_execute`.

Page insertion is an explicit exception to the mounted App history model: public
`pdf.page.insert { path, afterPageIndex }` stages a bounded local PDF, dispatches internal
`pdf.page.insert_staged`, persists the merged bytes, and reloads the same active document. Its
descriptor is therefore honestly `undoable:false`; user and Agent insertion share the same browser
host primitive. Save As, image export, page extraction, and print are non-document host effects.
Search, navigation, zoom, sidebars, view modes, and form focus are view-only.

## Verification evidence

The PDF workspace passes 23 test files and 280 assertions. This includes the restored non-AI
focused domains, operation catalog/handler coverage, producer-baseline validation, staged loading,
browser PDFium save paths, font coverage, generated-stamp replacement, and PDF-lib reopen checks.

`tests/visual/pdf-community.spec.ts` contains eleven real-host scenarios. Seven behavior tracers
prove:

1. user and MCP staged local open converge on the retained renderer;
2. user delete/undo and Agent annotation delete share mounted state and survive save/reopen;
3. Agent text and image insertion survives browser save and reopen;
4. text replacement and image transform/replace/delete survive sequential saves and reopens;
5. CJK, Korean, and Arabic text plus bounded selection-level colors remain searchable after save;
6. generated watermark/header/footer state can be set and explicitly cleared after reopen.
7. a positive staged-open acknowledgement is followed immediately by a typed page rotation,
   proving load commit/controller readiness without a timing delay.

Four additional scenarios cover the 420-pixel sidebar, 720-pixel split view, first fullscreen, and
fullscreen exit without iframe remount. PDF.js currently emits a non-fatal `standardFontDataUrl`
warning in some reopened fixtures; persisted content and extraction assertions still pass.

## Shared completion evidence

PDF has no unexplained retained state-changing producer gap. R6-01 closes the cross-format
packaged-host visual, pinned-source comparison, performance/resource, real MCP smoke,
license/prohibited-dependency, and repository acceptance gates. The generated capability reports
`ready: true`; future PDF work must still enter through the format-owned Registry.
