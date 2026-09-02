# Operation registry implementation plan

- Status: Complete through R6-09; all five retained renderer producer baselines, shared release
  gates, traced loading/startup evidence, bounded Manifest discovery, revision-guarded replay, and
  canonical transaction-only public execution pass
- Governing decision: [ADR 0004](../adr/0004-format-owned-operation-registries.md)
- Target contract: [Format-owned operation registry contract](../protocol/operation-registry.md)
- Baseline: `genspark-ai/genoffice@dc4d7e5927864498913b7ba42d0da06cc7cf628e`

## Outcome

Replace the transitional MCP capability catalog and per-renderer command tracers with five format-owned executable operation registries. Complete ADR 0003 parity by adding capabilities through those registries rather than extending the current central catalog, open string routes, or renderer composition-root condition chains.

The end state keeps:

- the mounted renderer as the only editable document authority;
- one persistent iframe per editing session;
- the existing thin live-session broker and monotonic revision;
- the current small public MCP tool surface;
- format-specific state, selection, undo, persistence, and command behavior in the five renderer applications.

It removes:

- manually duplicated central format schemas;
- arbitrary operation strings accepted before registry validation;
- permanent `ribbon_command` and shallow `batch_update` public contracts;
- format operation dispatch from large renderer composition roots;
- manually synchronized operation tables across server, Skill, protocol, and inventories.

## Starting-state evidence

| Area                | Current fact                                                                                                    | Structural problem                                                                                                   |
| ------------------- | --------------------------------------------------------------------------------------------------------------- | -------------------------------------------------------------------------------------------------------------------- |
| MCP tools           | `apps/mcp-server/src/server.ts` registers 18 stable tools.                                                      | Tool count is not the growth problem and should remain stable.                                                       |
| Capability catalog  | `apps/mcp-server/src/capabilities.ts` centrally describes all five formats.                                     | Format semantics and partial schemas lack Locality and can drift from renderer handlers.                             |
| Broker validation   | `office_execute` accepts `operation: string` and `arguments: Record<string, unknown>`.                          | Unsupported ids and schema errors are discovered after dispatch.                                                     |
| DOCX                | `batch_update` exposes only `commands: array` while the renderer owns multiple distinct commands.               | Discovery is shallow and atomicity is implicit.                                                                      |
| XLSX                | `ribbon_command` still forwards non-migrated command strings into the retained Ribbon dispatcher.               | The remaining route is not an exact typed Agent operation.                                                           |
| Renderer dispatch   | PPTX has no direct operation-id branches and its 74-operation registry covers every retained mutation producer. | Future browser handlers must land with an exact registry tracer rather than reintroducing composition-root dispatch. |
| DOCX transport      | Resolved: DOCX now uses the shared Host Bridge for polling, acknowledgement, display mode, and bundled fonts.   | The former format-local MCP session state machine is deleted, removing the dual-implementation drift risk.           |
| PPTX context        | The browser presentation Adapter and React App retain separate active-selection state.                          | UI and Agent targets do not yet converge on one native command seam.                                                 |
| PDF context         | The mounted controller publishes selection and routes 25 typed operations; its producer baseline has no gap.    | R6-01 supplies shared release evidence; no format-local state-changing producer is unexplained.                      |
| Capability evidence | Per-format inventories are primarily prose tables.                                                              | Baseline scope, executable status, source routes, and test evidence cannot be checked as one graph.                  |

All five pinned non-AI renderer sources are now mounted and the former replacement renderer directories are removed. This plan therefore changes command integration structure only; it does not create another renderer or migrate document authority.

## Implementation principles

1. **Structural migration before breadth.** Existing operations move into registries before new capability families are added.
2. **Replacement in the same slice.** A new registry route must delete or demote the transitional route it replaces; permanent dual dispatch is not accepted.
3. **Format Locality.** Descriptors, schemas, handlers, target semantics, and focused tests live with the owning renderer.
4. **Small common Interface.** Shared code contains envelopes and validation mechanics, never format semantics.
5. **Exact discovery.** Every advertised id has an exact schema; large whole-format unions and open command strings are not final contracts.
6. **Native execution.** Agent operations invoke the same renderer command/state/undo seam as user gestures.
7. **Evidence before advertisement.** A handler may exist privately before it passes, but it may not enter the generated MCP projection until its required tests pass.
8. **No new Runtime.** The broker remains live-bound and never parses or edits document bytes.

## Planned module layout

```text
packages/operation-contract/
  src/
    descriptor.ts
    transaction.ts
    validation.ts
    errors.ts
    manifest.ts

apps/<format>/src/renderer/operations/
  catalog/
    <family>.ts
  handlers/
    <family>.ts
  registry.ts
  context.ts

tools/
  generate-operation-manifest.ts
  operation-catalogs.ts
  check-operation-parity.mjs

apps/mcp-server/src/generated/
  operation-manifest.json

plugins/tandemfolio/skills/tandemfolio/references/
  operations.md

docs/migration/generated/
  <format>-operation-parity.md
```

Names may be adjusted to match the owning renderer, but the Module ownership and dependency direction are fixed:

```text
format registry definitions ──generate──> serializable manifest ──> MCP server
format registry handlers ───────────────> mounted renderer only
MCP server ─X─> renderer implementation imports
common contract ─X─> format semantics
```

## Milestones

### R0 — Architecture and planning baseline

Status: Complete.

Deliverables:

- ADR 0004;
- capability terminology in `CONTEXT.md`;
- target operation-registry protocol;
- this replacement plan;
- roadmap and documentation-map integration.

Exit gate:

- the accepted documents explicitly preserve mounted-renderer authority;
- the old central catalog and open command routes are named as migration scaffolds;
- new parity work has one documented entry path.

### R1 — Common contract and generator skeleton

Status: Complete as of 2026-08-28.

Goal: establish the shared Interface without changing behavior.

Deliverables:

- `packages/operation-contract` with descriptor, transaction, validation-result, error, and manifest types;
- a build-time generator that consumes serializable format catalogs without importing browser handlers;
- manifest schema/version and deterministic output ordering;
- registry integrity checks for duplicate ids, invalid format prefixes, unsupported schema keywords, alias collisions, and missing metadata;
- package/build/typecheck integration.

Exit gate:

- empty or fixture registries generate deterministic manifests;
- generator `--check` detects drift;
- no runtime tool or renderer behavior changes;
- common code contains no format-specific nouns or target schemas.

Implemented evidence:

- `packages/operation-contract` exports descriptor, transaction, validation-result, standard-error, and manifest contracts;
- `tools/generate-operation-manifest.ts` generates canonical JSON from serializable catalogs and rejects drift under `--check`;
- fixture catalogs exercise deterministic ordering, duplicate ids, format ownership and prefixes, aliases, required metadata, supported JSON Schema keywords, and structured validation results;
- root typecheck, test, and `check` commands include the new workspace and manifest drift gate;
- the MCP server still uses the transitional catalog, so this milestone changes no tool, renderer, session, or editing behavior.

### R2 — Migrate the current executable surface

Status: Complete. R2-01 through R2-52 completed on 2026-08-28; R2-53 through R2-76 plus the first
PDF/PPTX registry tracers completed on 2026-08-29; R2-77 through R2-308 and the final PDF parity
slice completed on 2026-08-30. Shared release gates remain.

The compatibility-alias statements in this R2 chronology describe each tracer's migration-time
state. R6-09 later retires all thirty-six public-era aliases; only the five internal
`open_local_file` staged-file transport aliases remain.

Goal: prove the architecture by moving every currently advertised operation without adding capability breadth.

R2-01 migrates only Markdown text insertion. `markdown.text.insert` is now owned by the
Markdown catalog and renderer registry, projected into the MCP capability response from the
generated manifest, validated before broker enqueue and again before renderer execution, and
applied through TipTap's native command/history route. The old `insert_text` name remains an
input-only compatibility alias and is normalized to the canonical id before dispatch. The direct
`insert_text` branch was removed from the Markdown composition root.

R2-02 applies the same replacement slice to selection replacement. The format-owned id is
`markdown.text.replace_selection`; legacy `replace_selection` remains an input-only alias. The
generated manifest now owns discovery and Broker validation for both Markdown text operations, and
the renderer registry owns both TipTap handlers. The old `markdown-commands.ts` helper and the final
text-operation condition branch in `App.tsx` are deleted.

R2-03 migrates Markdown persistence as `markdown.document.save`; legacy `save` remains an
input-only alias. Its descriptor declares exact empty input, exact `{ saved, fileName }` output,
document context, persistence effects, and non-undoable mutation semantics. The async registry
handler invokes the format-owned browser save service, reports cancellation or failure as
`execution_failed`, and returns the persisted filename on success. Generated discovery and Broker
normalization now cover all three public Markdown operations, and the direct `save` branch is
deleted from `App.tsx`. Markdown therefore has no public hand-written operation entry in the
central capability catalog; staged local-file loading was then the final internal transitional
route.

R2-04 migrates that final internal route as `markdown.document.load_staged`; legacy
`open_local_file` remains an internal compatibility alias. The descriptor is retained in the real
manifest with `visibility: internal`, so it can drive Broker normalization and renderer validation
without entering Agent capability discovery or becoming callable through `office_execute`. The
shared host bridge hydrates both the legacy transport name and any format-qualified
`*.document.load_staged` command before renderer execution. The Markdown registry validates the
hydrated `ArrayBuffer`, invokes the existing format-owned load pipeline, and suppresses a redundant
recovery checkpoint for the freshly loaded document. The last operation-specific branch is deleted
from `App.tsx`.

R2-05 proves that the same architecture crosses format boundaries by migrating DOCX insertion as
`docx.text.insert`; legacy `insert_text` remains an input-only compatibility alias scoped to the
DOCX format. `tools/operation-catalogs.ts` aggregates serializable format-owned catalogs for one
deterministic product manifest without moving either format's semantics into shared code. The
broker resolves the same legacy alias independently for DOCX and Markdown, validates the canonical
DOCX schema before enqueue, and queues only `docx.text.insert`. The DOCX registry executes the
existing TipTap insertion/history seam, after which `App.tsx` retains its existing renderer-owned
recovery checkpoint and live-session acknowledgement. The direct DOCX `insert_text` branch and its
hand-written capability entry are deleted; unrelated DOCX operations remain explicitly
transitional.

R2-06 migrates the adjacent DOCX selection replacement as `docx.text.replace_selection`; legacy
`replace_selection` remains an input-only compatibility alias scoped to DOCX. Its exact descriptor
joins the same product Manifest and the same format-owned renderer registry as insertion. The broker
normalizes the legacy name, rejects invalid canonical arguments before enqueue, and advertises only
the canonical id. The handler executes the existing TipTap
`deleteSelection().insertContent()` history seam before `App.tsx` performs the unchanged recovery
checkpoint and acknowledgement. The direct DOCX `replace_selection` branch and its hand-written
capability entry are deleted; `save`, `batch_update`, and internal local open remain transitional.

R2-07 migrates DOCX persistence as `docx.document.save`; legacy `save` remains an input-only DOCX
compatibility alias. Its descriptor declares exact empty input, exact `{ saved, fileName }` output,
document context, persistence effects, medium risk, and non-undoable atomic semantics. The DOCX
registry now has an async format-owned save service boundary: success returns the persisted file
identity, while cancellation or failure becomes `execution_failed`. The Broker normalizes the old
name, rejects extra arguments before enqueue, and clears the format recovery snapshot only after a
positive renderer acknowledgement carrying `saved: true`. `App.tsx` returns save output without
creating a redundant post-save recovery checkpoint. The direct App save branch and hand-written
central capability entry are deleted; only `batch_update` and internal local open remain in DOCX's
transitional command dispatch.

R2-08 migrates DOCX staged local loading as internal `docx.document.load_staged`; legacy
`open_local_file` remains an internal compatibility alias. The descriptor is retained in the
product Manifest with `visibility: internal`, so `office_open_local_file` can normalize to the
canonical id while Agent discovery and direct `office_execute` access remain blocked. The DOCX host
adapter chunk-hydrates the canonical command, the registry validates the hydrated `ArrayBuffer`,
`.docx` identity, and declared byte length, then invokes the injected format-owned load service.
Successful load suppresses a redundant recovery checkpoint; loader failures become deterministic
`execution_failed` acknowledgements. The old App branch and `openLocalFileCommand` helper are
deleted, leaving `batch_update` as DOCX's only transitional operation dispatch.

R2-09 starts decomposing that last transitional DOCX route. `batch_update.replaceAllText` is now
the agent-visible `docx.text.replace_all` operation with exact `{ containsText, replaceText,
matchCase? }` input and exact replacement-result output. Its registry handler wraps the existing
`executeCommands` replacement in a one-command envelope, preserving the mounted ProseMirror
transaction, tracked-deletion/protected-content accounting, and one-step native undo. Broker
validation rejects wrong types and extra fields before enqueue; an empty search is rejected by the
renderer without mutation. Because this mutator returns output, the DOCX composition root now
independently carries both registry output and the normal dirty recovery checkpoint. The old batch
variant is rejected with a canonical migration error, while the other nine structured batch
commands remain transitional.

R2-10 deepens the format-neutral schema contract only as required by the next DOCX tracer. The
bounded subset now admits and enforces `enum`, array `items`, and inclusive `minimum`/`maximum`
constraints at both Manifest integrity and runtime validation boundaries. DOCX consumes those
primitives in `docx.paragraph.set_heading_level`, whose exact `{ target, level }` input preserves
the existing block-index, node-type, heading-level, text, case, and selection/document targeting
surface. Its handler wraps native `setHeadingLevel` execution in one ProseMirror transaction,
returns deterministic matched/changed/protected counts, checkpoints normal dirty recovery, and is
restored by one undo. Broker validation rejects invalid enum values, array items, and numeric bounds
before enqueue; the renderer rejects a structurally valid but conditionless target without
mutation. `batch_update.setHeadingLevel` now returns its canonical migration error, leaving eight
structured commands transitional.

R2-11 migrates `batch_update.deleteBlocks` as the high-risk `docx.block.delete` operation. The
DOCX catalog now owns one reusable `docxBlockTargetSchema`, consumed by both block deletion and
heading-level changes without moving target semantics into the common contract. The delete handler
wraps the existing `deleteBlocks` command in one ProseMirror transaction, preserves the native
all-block deletion fallback to one empty paragraph, returns matched/changed/protected/tracked-
deleted counts, checkpoints normal dirty recovery, and is restored by one undo. Conditionless
targets are rejected without mutation, and unknown nested target fields are rejected by the Broker
before enqueue. `batch_update.deleteBlocks` now returns its canonical migration error, leaving seven
structured commands transitional.

R2-12 migrates `batch_update.deleteParagraphBullets` as `docx.list.remove`. Its exact `{ target }`
descriptor reuses the DOCX-owned `docxBlockTargetSchema`, proving the same target contract can serve
block, paragraph, and list families without entering the format-neutral operation contract. The
handler wraps the existing `deleteParagraphBullets` command in one ProseMirror transaction,
converts only matching list items to body paragraphs, preserves their text, returns deterministic
matched/changed/protected counts, checkpoints normal dirty recovery, and is restored by one undo.
Conditionless targets are rejected at the renderer seam; invalid bounded target values stop at the
Broker before enqueue. `batch_update.deleteParagraphBullets` now returns its canonical migration
error, leaving six structured commands transitional.

R2-13 migrates the complementary `batch_update.createParagraphBullets` command as
`docx.list.apply`. Its exact `{ target, kind }` descriptor reuses `docxBlockTargetSchema` and
replaces the transitional open `bulletPreset` prefix string with the bounded semantic enum
`bullet | ordered`. The handler maps that small Interface into the retained
`createParagraphBullets` command, lazily obtains a renderer-owned numbering definition only when a
matched block must change, and applies every target through one ProseMirror transaction. Text,
matched/changed/protected output, normal dirty recovery, revision advancement, and one-step undo are
preserved; applying the same kind to an existing list item is a successful no-op that retains its
numId and creates no orphan definition. `batch_update.createParagraphBullets` now returns its
canonical migration error, leaving five structured commands transitional.

R2-14 migrates `updateTextStyle` as `docx.text.set_style`. Its exact
`{ target, style, fields }` contract preserves the retained field-mask semantics, including
explicit `null` clearing for color, highlight, size, font, baseline offset, and link values. The
common schema subset gains bounded nullable type unions so both the Broker and renderer can
validate the contract without learning DOCX style vocabulary. The handler delegates to the native
`updateTextStyle` transaction, preserves unlisted marks, and remains one undo unit.

R2-15 applies the same pattern to `updateParagraphStyle` as `docx.paragraph.set_style`. Alignment,
spacing, indentation, page-break, shading, and border fields remain format-owned, targeted through
`docxBlockTargetSchema`, and applied only when named by the non-empty `fields` mask.

R2-16 migrates `moveBlocks` as the high-risk `docx.block.move` operation. Its exact
`{ blockIndexes, afterBlockIndex }` schema requires at least one non-negative source index and an
anchor of `-1` or greater. Renderer-context validation retains document-bound and non-overlap
checks; valid moves preserve source order and execute as one native undo transaction.

R2-17 migrates `updateImageProperties` as `docx.image.update`. The exact target/property/mask
contract retains proportional scaling when only one dimension is supplied, bounded positive
dimensions, nullable alignment, deterministic counts, normal recovery, and one-step undo against
the mounted protected image block.

R2-18 migrates `insertToc` as `docx.toc.insert`. It generates the retained DOCX TOC field blocks
from current heading state, distinguishes an invalid insertion anchor from unavailable heading
context, and inserts all entries in one undo transaction. With all ten structured commands now
registry-owned, the shallow `batch_update` capability, composition-root branch, executor, and
migration-error shell are deleted rather than retained as a second route.

R2-01 evidence:

- deterministic generation and drift checking consume the real format-owned TypeScript catalog;
- descriptor-to-handler coverage is compile-time exhaustive for the Markdown registry;
- focused renderer tests cover canonical execution, compatibility aliasing, exact argument
  rejection, document mutation, and one-step native undo;
- MCP integration tests cover generated discovery, alias normalization to the queued canonical id,
  and rejection of invalid arguments without enqueue or revision advance;
- no new operation family or public MCP tool was added.

R2-02 evidence:

- the real manifest exposes both canonical Markdown text ids and hides both compatibility aliases;
- focused renderer tests prove selection replacement and one-step native undo, plus legacy alias
  resolution;
- MCP integration proves `replace_selection` is normalized before enqueueing and acknowledged
  through the existing revision contract;
- deleting the legacy helper leaves no Markdown text-operation branch outside the registry.

R2-03 evidence:

- the real manifest exposes `markdown.document.save`, hides `save`, and records the exact output,
  persistence, risk, and undo metadata;
- focused renderer tests cover success output, cancellation/failure, and legacy alias resolution;
- MCP integration proves `save` is normalized to the canonical id before enqueue and that the
  existing revision/output acknowledgement contract is preserved;
- the direct Markdown save-dispatch branch and final hand-written public Markdown capability entry
  are deleted.

R2-04 evidence:

- the real manifest contains the internal canonical id and `open_local_file` alias while the live
  capability response continues to expose only the three Agent operations;
- focused renderer tests cover canonical load, internal alias resolution, cross-realm hydrated-byte
  validation, content replacement, and recovery-checkpoint suppression;
- MCP integration proves `office_open_local_file` queues the canonical id for Markdown while direct
  `office_execute` access is rejected before enqueue;
- shared host-bridge coverage proves canonical staged commands are chunk-hydrated before renderer
  execution, and `App.tsx` contains no operation-id dispatch branch.

R2-05 evidence:

- the product manifest aggregates the DOCX and Markdown catalogs and permits compatibility-alias
  reuse only across different format scopes;
- focused DOCX renderer tests cover canonical execution, legacy alias resolution, document
  mutation, and one-step native TipTap undo;
- MCP integration proves DOCX alias normalization, rejection of invalid canonical arguments before
  enqueue, generated capability discovery, and preservation of the existing revision contract;
- the direct DOCX insertion branch and central `insert_text` schema are removed without changing
  `replace_selection`, `batch_update`, save, local-open, recovery, or iframe transport behavior.

R2-06 evidence:

- the generated product manifest contains both canonical DOCX text operations and hides both
  compatibility aliases from capability discovery;
- focused DOCX renderer tests cover canonical selection replacement, legacy alias resolution,
  document mutation, and one-step native TipTap undo;
- MCP integration proves alias normalization, exact pre-enqueue argument rejection, no revision
  advance on rejection, and the existing successful acknowledgement contract;
- the direct DOCX replacement branch and central `replace_selection` schema are removed without
  changing insertion, save, batch, local-open, recovery, or iframe transport behavior.

R2-07 evidence:

- the generated product manifest exposes `docx.document.save` with exact empty input and exact
  persistence output while hiding the legacy `save` alias;
- focused DOCX registry tests cover canonical success output, persistence failure mapping, and
  legacy alias resolution through the injected format save boundary;
- MCP integration proves canonical queueing, exact argument rejection before enqueue, persisted
  identity propagation, clean revision acknowledgement, and recovery removal after success;
- the direct DOCX save branch and central `save` schema are removed without changing text, batch,
  local-open, renderer authority, or iframe lifecycle behavior.

R2-08 evidence:

- the generated product Manifest contains `docx.document.load_staged`, internal visibility, and the
  `open_local_file` compatibility alias without exposing either id through Agent discovery;
- focused DOCX registry tests cover canonical load, hydrated-byte, extension and length validation,
  loader failure mapping, alias resolution, and recovery-checkpoint suppression;
- MCP and DOCX host-session integration prove `office_open_local_file` queues the canonical id,
  chunk-hydrates it before renderer execution, and rejects direct `office_execute` access;
- the DOCX composition root contains no local-open operation branch, and the former command-shaped
  helper is replaced by a narrow validated `loadStagedDocx` host service.

R2-09 evidence:

- the generated product Manifest and capability projection expose only
  `docx.text.replace_all`, including exact required fields, optional case matching, result schema,
  document effect, medium risk, atomicity, and native undo metadata;
- focused registry tests prove whole-document mutation, deterministic output, empty-search
  rejection without mutation, and restoration of all replacements through one undo action;
- MCP integration proves canonical argument queueing, renderer output propagation, one revision
  advance, and rejection of extra arguments without enqueue or revision advance;
- `batch_update.replaceAllText` now returns a migration error and cannot bypass the registry; the
  transitional batch remains available only for its other nine structured commands.

R2-10 evidence:

- public operation-contract tests prove supported-keyword admission and runtime enforcement for
  enums, every array item, and inclusive lower/upper numeric bounds through
  `validateJsonSchemaValue`;
- the generated product Manifest and capability projection expose
  `docx.paragraph.set_heading_level` with the complete DOCX target shape, bounded heading levels,
  structured output, document/selection context, medium risk, atomicity, and native undo metadata;
- focused registry tests prove targeted paragraph promotion, deterministic output, one-step undo,
  and conditionless-target rejection without mutation;
- MCP integration proves canonical target queueing, output propagation, one revision advance, and
  pre-enqueue rejection of invalid enum, array-item, and numeric-range values;
- `batch_update.setHeadingLevel` can no longer bypass the registry, while the transitional batch
  remains available for its other eight structured commands.

R2-11 evidence:

- `docxBlockTargetSchema` is a DOCX-owned serializable schema reused by block deletion and heading
  mutation descriptors; the common contract remains format-neutral;
- the generated product Manifest and capability projection expose `docx.block.delete` with exact
  target input, deterministic deletion/skipped-content output, document/selection context, high
  risk, atomicity, and native undo metadata;
- focused registry tests prove targeted deletion, deterministic output, one-step undo, and
  conditionless-target rejection without mutation; existing native command tests retain the
  delete-all-to-empty-paragraph invariant;
- MCP integration proves canonical target queueing, output propagation, one revision advance, and
  pre-enqueue rejection of unknown nested target fields;
- `batch_update.deleteBlocks` can no longer bypass the registry, while the transitional batch
  remains available for its other seven structured commands.

R2-12 evidence:

- the generated product Manifest and capability projection expose `docx.list.remove` with the
  reusable exact target input, structured output, document/selection context, medium risk,
  atomicity, and native undo metadata;
- focused registry tests prove targeted list-format removal, text preservation, deterministic
  output, one-step undo, and conditionless-target rejection without mutation;
- MCP integration proves canonical target queueing, output propagation, one revision advance, and
  pre-enqueue rejection of an invalid block-index item through `docxBlockTargetSchema`;
- `batch_update.deleteParagraphBullets` can no longer bypass the registry, while the transitional
  batch remains available for its other six structured commands.

R2-13 evidence:

- the generated product Manifest and capability projection expose `docx.list.apply` with reusable
  exact target input, bounded list kind, structured output, document/selection context, medium
  risk, atomicity, and native undo metadata;
- focused registry tests prove ordered-list conversion, renderer numbering allocation, text
  preservation, deterministic output, one-step undo, same-kind no-op behavior, and
  conditionless-target rejection without mutation;
- MCP integration proves canonical `{ target, kind }` queueing, output propagation, one revision
  advance, and pre-enqueue rejection of legacy preset strings through the bounded enum;
- `batch_update.createParagraphBullets` can no longer bypass the registry, while the transitional
  batch remains available for its other five structured commands.

R2-14 through R2-18 evidence:

- the generated Manifest and Agent capability projection expose exact schemas for
  `docx.text.set_style`, `docx.paragraph.set_style`, `docx.block.move`, `docx.image.update`, and
  `docx.toc.insert`, while `batch_update` is absent;
- public contract tests prove nullable primitive/object validation and non-empty array enforcement;
- focused renderer tests prove field-mask preservation, explicit style clearing, ordered block
  movement, proportional image scaling, TOC-field generation, deterministic output, and one-step
  undo through the retained command seam;
- MCP integration tests prove canonical queueing, renderer-output propagation, one revision
  advance, and pre-enqueue rejection for malformed nested, enum, minimum, and `minItems` input;
- the DOCX composition root now dispatches every current operation through the registry, and the
  former batch executor plus its hand-written capability schema are deleted.

R2-19 evidence:

- the first XLSX-owned executable descriptor and handler expose `xlsx.document.save` with exact
  empty input, exact `{ saved, fileName }` output, document context, persistence effects, medium
  risk, atomicity, and non-undoable semantics;
- the handler delegates to the retained XLSX save assembler and browser package/download adapter;
  the save boundary now reports persisted identity, cancellation, and write failure explicitly;
- the generated product Manifest and Agent capability projection advertise only the canonical id,
  while legacy `save` remains an input-only XLSX compatibility alias normalized before enqueue;
- focused renderer tests prove success and negative acknowledgement, and MCP integration proves
  canonical queueing, output propagation, one revision advance, and rejection of extra arguments
  before renderer dispatch;
- the direct XLSX `save` branch and hand-written capability entry are deleted; other XLSX command
  families remain transitional and must migrate independently.

R2-20 evidence:

- the XLSX-owned catalog and one-to-one handler map now include internal
  `xlsx.document.load_staged` with exact staged descriptor input, exact `{ opened, fileName }`
  output, document/selection effects, medium risk, atomicity, and non-undoable semantics;
- `office_open_local_file` resolves the internal `open_local_file` compatibility alias by XLSX
  session format and queues only the canonical id, while capability discovery and direct
  `office_execute` access exclude the internal descriptor;
- the shared host bridge hydrates bounded staged chunks before registry execution; the XLSX seam
  validates a real `ArrayBuffer`, `.xlsx` identity, and declared byte length before invoking the
  retained `openBuffer` → `openLazyWorkbook` route;
- successful load returns the existing `{ opened: true, fileName }` acknowledgement without a
  redundant recovery checkpoint, and parser/load failures return `execution_failed`;
- focused registry and MCP integration tests cover canonical dispatch, legacy aliasing, malformed
  hydrated input, descriptor integrity, internal visibility, canonical queueing, output, and one
  revision advance; the direct XLSX `open_local_file` App branch is deleted.

R2-21 through R2-25 evidence:

- the XLSX-owned catalog and one-to-one handler map add `xlsx.cell.set_value`,
  `xlsx.range.set_values`, `xlsx.row.insert`, `xlsx.row.delete`, and `xlsx.column.insert` with exact
  input/output schemas, bounded scalar matrices, 1-based row targets, and shared structural-count
  limits;
- handlers resolve the named worksheet from the mounted Univer runtime, activate cell/range targets,
  and invoke native range, row, and column mutation methods so user and Agent edits keep one workbook
  authority and undo journal;
- legacy `set_cell_value`, `set_range_values`, `insert_rows`, `delete_rows`, and `insert_columns`
  remain input-only aliases and are normalized to canonical ids before enqueue;
- focused renderer tests prove value translation, target activation, 1-based to zero-based structural
  translation, normalized A1 column labels, and semantic rejection without mutation; MCP integration
  proves canonical queueing, exact output, one revision advance, and bounded pre-enqueue rejection;
- the five hand-written capability entries and matching XLSX App branches are deleted. Six
  transitional XLSX routes remain: `delete_columns`, four worksheet operations, and
  `ribbon_command`.

R2-26 through R2-30 evidence:

- `xlsx.column.delete` completes the column structure pair by reusing the bounded column mutation
  contract and native Univer deletion route; `xlsx.sheet.add`, `xlsx.sheet.rename`,
  `xlsx.sheet.delete`, and `xlsx.sheet.move` replace all four worksheet transition branches;
- sheet handlers operate only on the mounted active workbook, preserve the 31-character name
  boundary, prevent deletion of the final worksheet, and translate 1-based tab positions to native
  zero-based indexes;
- legacy `delete_columns`, `add_sheet`, `rename_sheet`, `delete_sheet`, and `move_sheet` names remain
  input-only aliases normalized before enqueue;
- focused renderer tests prove A1 column translation, name validation, final-sheet protection,
  dynamic move bounds, native target selection, and deterministic output; MCP integration proves
  canonical queueing, exact schemas, pre-enqueue rejection, output propagation, and one revision
  advance;
- the five hand-written capability entries, the `mcpSheet` composition-root helper, and matching App
  branches are deleted. `ribbon_command` is the only remaining transitional XLSX route.

R2-31 evidence:

- `xlsx.range.set_text_style` starts decomposing `ribbon_command` with an exact
  `{ sheet, range, style, fields }` descriptor for bold, italic, strike, and three-state underline;
- renderer validation requires unique field names and an exact field/value mask, so Agent calls are
  explicit assignments rather than selection-dependent toggles;
- the complete style patch is passed to one native Univer range mutation, preserving one mounted
  workbook authority, target activation, one native undo unit, and one acknowledged revision;
- generated capability and Broker tests prove each style field and the bounded underline enum;
  renderer tests prove explicit clearing, multi-field atomicity, and rejection without mutation;
- real-host browser evidence proves the canonical operation changes the active Univer state and is
  undone by the visible UI, while MCP text-style `ribbon_command` variants are rejected. The visible
  Ribbon dispatcher remains unchanged for user gestures and non-migrated families.

R2-32 through R2-36 evidence:

- `xlsx.range.set_alignment` replaces `align:*`, `valign:*`, `wrap:*`, `indent:*`, and
  `rotate:*` with an exact masked contract covering explicit clear, bounded indent, bounded angle,
  and stacked rotation; all requested fields land in one native style mutation;
- `xlsx.range.set_font` replaces `font-family:*`, `font-size:*`, and `font-color:*` with explicit
  masked values and bounded size/color validation, again using one native style mutation;
- `xlsx.range.set_fill` and `xlsx.range.set_border` replace their Ribbon string families with exact
  nullable color and border preset/line-style contracts; border uses the retained native range
  border command so Univer remains the undo authority;
- `xlsx.range.apply_cell_style` bounds the retained built-in style gallery and folds each preset
  into one native style mutation rather than replaying several facade setters;
- generated discovery and Broker tests prove all five schemas, bounds, and canonical queueing;
  renderer tests prove exact masks, clear behavior, atomic patches, native border dispatch, and
  rejection without mutation;
- real-host evidence applies all five operations to the mounted workbook, observes the same live
  selection state, undoes/redoes through the visible UI, and verifies the saved XLSX style output;
- MCP rejects migrated alignment/font/fill/border/cell-style Ribbon strings. `format-painter` is
  also rejected because its armed transient mode is not replayable without the next user gesture;
  Agents use explicit destination range operations instead.

R2-37 through R2-40 evidence:

- `xlsx.range.set_number_format` replaces `format:*`, `decimal-inc`, and `decimal-dec` with one
  explicit 1–255-character final pattern. Agents no longer depend on the current display value or
  relative decimal toggles;
- `xlsx.range.merge` replaces `merge:*` with `cells | across | center | unmerge`. Cells, across, and
  unmerge call the matching native range actions. Center preserves the retained merge-then-align UI
  sequence and is transparently marked non-atomic;
- `xlsx.range.clear` replaces `clear-contents`, `clear-formats`, and `clear-all` with an exact
  `contents | formats | all` scope, using the native range facade and shared undo journal;
- `xlsx.range.fill` replaces `fill-down` and `fill-right` with an explicit target and direction,
  rejects ranges without a destination row/column, activates the target, and invokes Univer's
  native copy command;
- generated discovery, Broker validation, and queue/acknowledgement tests cover all four operations
  and their bounded enums; renderer tests cover every merge/clear/fill mode plus invalid targets;
- real-host evidence observes the assigned number format in live context, proves clear undo/redo
  through the visible UI, and verifies filled data, merge structure, cleared content, and number
  format in the saved XLSX package;
- MCP rejects all migrated number-format, merge, clear, and fill Ribbon strings. Visible Ribbon
  gestures continue through their retained user path.

R2-41 through R2-43 evidence:

- `xlsx.range.sort` replaces `sort:asc` and `sort:desc` with an explicit target and direction,
  preserves the retained first-column behavior, rejects single-row targets, and awaits Univer's
  native `sheet.command.sort-range` result before acknowledging;
- `xlsx.range.sort_custom` replaces `sort-custom:*` with ordered, unique
  `{ column, direction }` keys plus `hasHeader`; keys normalize to A1 uppercase and must fall inside
  the explicit target before the same native sort command is called;
- `xlsx.range.remove_duplicates` replaces `remove-duplicates:*`, preserves the existing
  case-insensitive whole-row comparison and first-occurrence rule, reports `removed`, rejects
  single-row or partially streamed sources before reading, and rewrites only changed rows;
- generated discovery, Broker queueing, exact enum/boolean validation, renderer validation, and
  migrated-Ribbon rejection tests cover all three operations;
- real-host evidence executes basic sort, custom sort, and deduplication in one mounted workbook,
  observes one revision per acknowledgement, proves visible undo/redo for the cleared duplicate,
  and verifies sorted/deduplicated OOXML output;
- sort operations are native and atomic; remove-duplicates is transparently `atomic: false` because
  selective row rewrites preserve unchanged formula cells across the target.

R2-44 through R2-48 evidence:

- `xlsx.hyperlink.set` and `xlsx.hyperlink.remove` replace `link-set:*` and `link-remove`, share
  the Ribbon's target normalization, file journal, and link appearance, and return exact cell-level
  acknowledgements;
- `xlsx.table.add` replaces `format-as-table:*`, bounds style to the six visible gallery choices,
  reuses the retained table engine and generated naming, and deliberately omits an unsupported
  existing-table restyle operation;
- `xlsx.range.set_protection` replaces `cellprot:*` with an exact `locked | hidden` field mask and
  rejects targets above 10,000 cells before any file-side journal write;
- `xlsx.sheet.set_protection` replaces `sheet-protect` toggle semantics with an explicit boolean,
  preserves indexed-sheet and password-removal guards, and shares the same journal seam as the UI;
- generated discovery, Broker queue/acknowledgement, bounded-enum validation, renderer behavior,
  and migrated-Ribbon rejection tests cover all five operations;
- hyperlink, cell-protection, and sheet-protection writes are non-undoable file-side journals;
  table creation is also non-atomic. Their browser package save/reopen gate remains open and is not
  claimed by this registry slice.

R2-49 evidence:

- `xlsx.formula.insert_aggregate` replaces `autofn:SUM|AVERAGE|COUNT|MAX|MIN` with an exact
  `{ sheet, range, function }` contract and reports its source range, destination row, and inserted
  formula count;
- the retained Ribbon and registry handler share one XLSX-owned formula seam, preserving one
  formula below each selected column and the original per-column Univer undo behavior;
- source ranges require at least two rows, and a destination row backed by original file content
  must be streamed before the shared seam writes any formula;
- generated discovery, Broker queue/acknowledgement, bounded-enum validation, all five aggregate
  functions, and migrated-Ribbon rejection are covered through public interfaces;
- multi-column writes remain transparently `atomic: false`; formula persistence continues through
  the existing Univer edit journal and browser OOXML formula save/reopen route.

R2-50 evidence:

- `xlsx.range.flash_fill` replaces `flash-fill` with the exact `{ sheet, range }` contract and
  reports the requested range, effective target range, and number of newly filled cells;
- the retained Ribbon and registry handler share one XLSX-owned Flash Fill seam that reads at most
  six source columns, learns from at most three non-empty target examples, preserves every existing
  target value, and writes all inferred results through one native `setValues` mutation;
- a single-cell target follows the immediately adjacent left column and bounds its downward probe
  to 1,000 rows; an explicit multi-row range remains explicit;
- both probe and inference rectangles reject original-file rows that are still streaming before
  `getValues` or `setValues`, preventing unloaded cells from being mistaken for blanks;
- generated discovery, Broker queue/acknowledgement, exact-schema validation, Ribbon equivalence,
  bounded probing, streaming rejection, and migrated-Ribbon rejection are covered through public
  interfaces; generic range-value save/reopen evidence continues to cover the final native write.

R2-51 evidence:

- `xlsx.range.text_to_columns` replaces `text-to-columns:1|2|4|8` with the exact
  `{ sheet, range, delimiter }` contract and the bounded `tab | comma | semicolon | space` enum;
- the retained Ribbon and registry handler share one XLSX-owned seam that invokes Univer's native
  `sheet.command.split-text-to-columns` command with an explicit workbook, worksheet, and range;
- the registry rejects multi-column targets and partially streamed worksheets before mutation;
  Univer retains its visible behavior of overwriting cells to the right and inserting columns when
  the split result exceeds the current worksheet width, so the operation is explicitly high-risk;
- generated discovery, Broker queue/acknowledgement, enum validation, all four delimiter mappings,
  native-command failure, Ribbon equivalence, and migrated-Ribbon rejection are covered through
  public interfaces; Univer owns the grouped undo unit and the existing range-edit save journal.

R2-52 evidence:

- `xlsx.row.set_height` replaces `row-height:*` with the exact
  `{ sheet, row, count, heightPoints }` contract; rows are 1-based, the contiguous count is bounded
  to 10,000, and the final Excel point height is bounded to the retained `0–409.5` range;
- the retained Ribbon and registry handler share one XLSX-owned seam that converts points to Univer
  pixels and invokes `setRowHeightsForced` against the mounted worksheet;
- a cross-field renderer guard rejects spans past the worksheet row limit before mutation, while
  generated schema validation rejects malformed bounds before Broker enqueue;
- generated discovery, Broker queue/acknowledgement, schema bounds, explicit target execution,
  Ribbon equivalence, and migrated-Ribbon rejection are covered through public interfaces;
- real-host evidence proves Agent mutation, user undo/redo, saved `ht`/`customHeight` output for the
  complete row span, and preservation after reopen. Univer owns the atomic undo unit and the
  existing axis-attribute journal owns OOXML persistence.

R2-53 evidence:

- `xlsx.column.set_width` replaces `col-width:*` with the exact
  `{ sheet, column, count, widthCharacters }` contract; the A1 start column and contiguous count are
  explicit, counts remain bounded to 10,000, and requested width is bounded to Excel's `0–255`
  character range;
- Ribbon and registry callers share one XLSX-owned character-width-to-pixel seam and the retained
  native column mutation; the operation returns the actual 1/256-character OOXML width after pixel
  quantization rather than falsely echoing an unrepresentable request;
- a renderer cross-field guard rejects spans beyond the mounted worksheet before mutation, while
  schema bounds reject malformed widths and counts before Broker enqueue;
- real-host evidence proves Agent mutation, user undo/redo, quantized `width`/`customWidth` output
  for the complete column span, and preservation after reopen.

R2-54 evidence:

- `xlsx.sheet.set_freeze` replaces `freeze-here`, `freeze-top-row`, `freeze-first-col`, and
  `unfreeze` with the exact `{ sheet, frozenRows, frozenColumns }` final-state contract;
- `0,0` explicitly cancels freezing, one zero axis represents a top-row or first-column freeze,
  and non-zero counts identify the top-left scrollable cell without depending on Agent selection;
- Ribbon and registry callers share one XLSX-owned native freeze seam. The renderer rejects a split
  that would leave no scrollable row or column, while schema bounds enforce XLSX maxima;
- real-host evidence proves Agent mutation, user undo/redo, saved frozen-pane coordinates, and
  preservation after reopen through the existing page-setup journal.

R2-55 evidence:

- `xlsx.sheet.set_gridlines` replaces the selection-independent `toggle-gridlines` string with the
  exact `{ sheet, visible }` final-state contract;
- the retained Ribbon toggle derives its next explicit value and converges with the registry on one
  XLSX-owned native gridline seam, keeping Univer as the undo authority;
- generated discovery, Broker queue/acknowledgement, boolean validation, and migrated-Ribbon
  rejection are covered through public interfaces;
- real-host evidence proves Agent mutation, user undo/redo, saved `sheetView@showGridLines`, and
  preservation after reopen through the existing page-setup journal.

R2-56 evidence:

- `xlsx.sheet.set_formula_view` replaces the state-dependent `toggle-show-formulas` string with the
  exact `{ sheet, enabled }` final-state contract;
- the retained Ribbon derives its next boolean before converging with the registry on one
  XLSX-owned formula-view projection and page-setup journal seam;
- the shared `pushWorkbookUndo` adapter records renderer-owned state changes in Univer's history,
  so formula view interleaves with native cell edits and both Undo and Redo reapply the visible and
  persisted state through the same seam;
- generated discovery, Broker queue/acknowledgement, boolean validation, and migrated-Ribbon
  rejection are covered through public interfaces;
- real-host evidence proves Agent mutation, user undo/redo, saved `sheetView@showFormulas`, removal
  of that attribute after undo, and preservation after reopen.

R2-57 evidence:

- `xlsx.sheet.set_page_orientation` replaces `page-layout:orientation:portrait|landscape` with the
  exact `{ sheet, orientation }` final-state contract and a bounded `portrait | landscape` enum;
- retained Ribbon and registry callers share `applyWorkbookPageOrientation`, so both update the
  same format-owned page-setup journal instead of duplicating save semantics;
- the operation snapshots the effective file/journal orientation and registers exact Undo/Redo
  callbacks through `pushWorkbookUndo`; undo removes a newly introduced override instead of
  manufacturing a default value;
- generated discovery, Broker queue/acknowledgement, enum validation, and only the migrated
  orientation Ribbon-family rejection are covered through public interfaces; other `page-layout:*`
  commands remain transitional;
- real-host evidence proves Agent mutation, user undo/redo, saved `pageSetup@orientation`, removal
  of the override after undo, and preservation after reopen.

R2-58 evidence:

- `xlsx.sheet.set_page_margins` replaces `page-layout:margins:normal|wide|narrow` with the exact
  `{ sheet, margins }` final-state contract and a bounded `normal | wide | narrow` enum;
- retained Ribbon and registry callers share `applyWorkbookPageMargins`, and orientation/margins
  now hide their common file fallback, journal update, and history mechanics behind one private
  page-setup preset module;
- renderer-owned Undo restores the exact prior margins field presence/value, while Redo reapplies
  the requested preset and the existing OOXML writer expands it to all six `pageMargins` values;
- generated discovery, Broker queue/acknowledgement, enum validation, and only the migrated margins
  Ribbon-family rejection are covered through public interfaces; `page-layout:paper:*` and the
  other page-layout commands remain transitional;
- real-host evidence proves Agent mutation, user undo/redo, saved wide-margin values, removal of a
  newly introduced `pageMargins` element after undo, equivalent retained Ribbon output, and
  preservation after reopen.

R2-59 evidence:

- `xlsx.sheet.set_paper_size` replaces retained `page-layout:paper:1|3|5|7|8|9|11` strings with
  the exact `{ sheet, paperSize }` final-state contract and an integer enum bounded to the seven
  visible paper presets;
- retained Ribbon and registry callers share `applyWorkbookPaperSize`, while the private page-setup
  preset module continues to own file fallback, idempotence, journal replacement, and history;
- renderer-owned Undo restores the exact prior `paperSize` field presence/value, while Redo
  reapplies the requested preset and the existing OOXML writer persists `pageSetup@paperSize`;
- generated discovery, Broker queue/acknowledgement, numeric-enum validation, and only the migrated
  paper-size Ribbon-family rejection are covered through public interfaces;
- real-host evidence proves Agent mutation, user undo/redo, removal of a newly introduced page-setup
  override after Undo, equivalent retained Ribbon A4 output, and preservation after reopen.

R2-60 evidence:

- `xlsx.sheet.set_fit_to_pages` replaces retained `page-layout:fit-width:*` and
  `page-layout:fit-height:*` strings with the exact `{ sheet, widthPages, heightPages }` final-state
  contract; both axes are integers from `0` through `1000`, where `0` means Automatic;
- Agent callers must supply both axes, while a retained Ribbon single-axis gesture resolves the
  untouched axis from the effective file/journal state before entering `applyWorkbookFitToPages`;
- the private page-setup patch module now owns single- and multi-field file fallback, idempotence,
  journal replacement, and one exact Undo/Redo unit; `0,0` records `fitToPage: false`, while any
  positive axis records `fitToPage: true`;
- generated discovery, Broker queue/acknowledgement, numeric-bound validation, and rejection of
  both migrated fit-width and fit-height Ribbon families are covered through public interfaces;
- real-host evidence proves Agent mutation, exact user Undo/Redo, removal of newly introduced fit
  overrides after Undo, saved `fitToWidth="2"` / `fitToHeight="3"` plus `fitToPage="1"`, equivalent
  retained Ribbon output, and preservation after reopen.

R2-61 evidence:

- `xlsx.sheet.set_print_gridlines` replaces retained `page-layout:print-gridlines:0|1` strings with
  the exact `{ sheet, enabled }` final-state contract, separate from worksheet display gridlines;
- Agent and retained Ribbon callers share `applyWorkbookPrintGridlines`, which enters the private
  page-setup patch seam for file fallback, idempotence, journal replacement, and exact Undo/Redo;
- generated discovery, Broker queue/acknowledgement, boolean validation, and rejection of the
  migrated print-gridline Ribbon family are covered through public interfaces;
- real-host evidence proves Agent and visible Ribbon mutation, user Undo/Redo, removal of a newly
  introduced `printOptions@gridLines` override after Undo, explicit disable, and preservation after
  save/reopen.

R2-62 evidence:

- `xlsx.sheet.set_print_headings` replaces retained `page-layout:print-headings:0|1` strings with
  the exact `{ sheet, enabled }` final-state contract for printed row and column headings;
- Agent and retained Ribbon callers share `applyWorkbookPrintHeadings`, reusing the private
  page-setup patch seam for file fallback, idempotence, journal replacement, and exact Undo/Redo;
- generated discovery, Broker queue/acknowledgement, boolean validation, and rejection of the
  migrated print-heading Ribbon family are covered through public interfaces;
- real-host evidence proves Agent and visible Ribbon mutation, user Undo/Redo, removal of a newly
  introduced `printOptions@headings` override after Undo, explicit disable, and preservation after
  save/reopen.

R2-63 evidence:

- `xlsx.sheet.set_print_area` replaces retained `page-layout:print-area:set|clear` strings with the
  exact `{ sheet, range }` contract, where a normalized explicit A1 cell range sets the area and
  `null` clears it;
- the registry validates and normalizes Agent-provided A1 ranges, while the retained Ribbon resolves
  its current selection before both callers enter `applyWorkbookPrintArea` and the private
  page-setup patch seam;
- generated discovery, Broker queue/acknowledgement, nullable-type validation, and rejection of both
  migrated print-area Ribbon strings are covered through public interfaces;
- real-host evidence proves Agent and visible Ribbon mutation, user Undo/Redo, removal after Undo or
  explicit clear, persisted `_xlnm.Print_Area` absolute references, and preservation after reopen.

R2-64 evidence:

- `xlsx.sheet.set_print_titles` replaces retained
  `page-layout:print-titles:first-row|set|clear` strings with the exact `{ sheet, rows }` contract,
  where an ascending explicit row span sets repeated title rows and `null` clears them;
- the registry rejects malformed, reversed, out-of-sheet, or more-than-21-row spans, while the
  retained Ribbon resolves Row 1 or the live selected rows before both callers enter
  `applyWorkbookPrintTitles` and the private page-setup patch seam;
- generated discovery, Broker queue/acknowledgement, nullable-type validation, and rejection of all
  three migrated print-title Ribbon strings are covered through public interfaces;
- real-host evidence proves Agent and visible Ribbon mutation, user Undo/Redo, removal after Undo or
  explicit clear, persisted `_xlnm.Print_Titles` absolute row references, and preservation after
  reopen.

R2-65 evidence:

- `xlsx.sheet.set_print_scale` replaces retained `page-layout:scale:10..400` strings with the exact
  `{ sheet, scalePercent }` final-state contract and generated integer bounds;
- Agent and the retained page-layout dispatcher share `applyWorkbookPrintScale`, which records
  `scale` and `fitToPage: false` through the private multi-field page-setup patch seam as one exact
  Undo/Redo unit;
- generated discovery, Broker queue/acknowledgement, numeric-bound validation, and rejection of the
  migrated scale Ribbon-command family are covered through public interfaces;
- real-host evidence starts from a fit-to-pages workbook and proves Agent mutation, user Undo/Redo,
  restoration of the original fit mode after Undo, removal of `pageSetUpPr@fitToPage`, persisted
  `pageSetup@scale="80"`, and preservation after reopen. The pinned Page Layout Ribbon has no visible
  Scale control, so this slice does not invent one.

R2-66 evidence:

- `xlsx.range.set_filter` replaces retained `filter-toggle` with the exact
  `{ sheet, range, enabled }` final-state contract; repeated calls that already match the requested
  state succeed without toggling it again, and a different active filter range fails closed;
- Agent and the retained Ribbon share `filter-actions.ts`; the Ribbon resolves its current filter or
  smart selection before both callers enter Univer's native set/remove filter commands and undo
  stack;
- generated discovery, Broker queue/acknowledgement, boolean validation, and rejection of the
  migrated `filter-toggle` MCP string are covered through public interfaces;
- real-host evidence proves Agent mutation, visible Ribbon equivalence, user Undo/Redo, removal after
  Undo, persisted `<autoFilter ref="A1:C4">`, and preservation after reopen. The browser package
  adapter now routes filter snapshots through the retained `xlsx-filter` gateway and restores the
  saved filter range on open.

R2-67 through R2-69 evidence:

- `xlsx.range.clear_filter_criteria { sheet, range }` replaces retained `filter-clear`, requires an
  exact active AutoFilter range, clears every column criterion through Univer's native command and
  keeps the AutoFilter itself;
- `xlsx.range.set_filter_values { sheet, range, column, values, includeBlank }` adds a bounded
  1–10,000 item value-list route for the retained native filter dropdown semantics;
- `xlsx.range.set_custom_filter { sheet, range, column, conjunction, conditions }` replaces the
  mutation behind retained `filter-advanced`; one or two custom conditions use six saveable
  operators and an explicit `and | or` conjunction;
- the common Contract now supports `maxItems`, so Broker validation rejects oversized condition
  and value arrays before enqueueing. Renderer validation additionally rejects a missing filter,
  mismatched range, or out-of-range column;
- Agent, Ribbon Clear, the Advanced Filter dialog, and native filter UI all converge on
  `filter-actions.ts` and Univer's native filter commands/Undo stack. The dialog awaits command
  completion before closing;
- generated discovery, queue/acknowledgement, legacy `filter-clear`/`filter-advanced` rejection,
  saved value/custom criteria, filtered row visibility, Undo/Redo, and preservation after reopen
  are covered through public and real-host interfaces.

R2-70 evidence:

- Agent `ribbon_command` calls for `insert-sheet`, `insert-row-here`, `delete-row-here`,
  `insert-col-here`, and `delete-col-here` are rejected before the retained Ribbon dispatcher can
  resolve the active selection;
- each error names the existing exact replacement: `xlsx.sheet.add`, `xlsx.row.insert`,
  `xlsx.row.delete`, `xlsx.column.insert`, or `xlsx.column.delete`. No new operation descriptor or
  handler is added, so the generated product manifest remains at sixty-nine operations and XLSX at
  fifty-one;
- the visible sheet-add and row-insert controls still enter the retained workbook/worksheet route;
  real-host tests retain native Undo/Redo and save/reopen evidence;
- the public MCP-to-editor rejection tracer and the fast renderer replacement-map suite cover all
  five strings.

R2-71 evidence:

- `xlsx.history.undo {}` and `xlsx.history.redo {}` are exact empty-input operations owned by the
  XLSX catalog and renderer registry; successful results are `{ undone: true }` and
  `{ redone: true }`;
- both handlers call the mounted runtime's `univerAPI.undo()` / `redo()` directly, so Agent and
  visible UI controls share the same native workbook history instead of introducing a parallel
  history model;
- an empty native history stack returns `execution_failed` and does not advance session revision;
  successful undo and redo each acknowledge exactly one monotonically increasing revision;
- the real-host tracer edits a cell through `xlsx.cell.set_value`, undoes it, redoes it, and saves
  the redone value from the same mounted workbook. Broker discovery, exact empty-object validation,
  renderer validation, and old-string rejection are covered;
- Agent `ribbon_command` calls for `undo` and `redo` are rejected with pointers to the exact history
  operations. The visible toolbar controls retain their existing direct Univer route;
- the generated product manifest now contains seventy-one operations, including fifty-three XLSX
  operations.

R2-72 evidence:

- `xlsx.range.copy_values { sourceSheet, sourceRange, destinationSheet, destinationRange }` replaces
  the first clipboard-dependent Agent mutation with an exact source/destination contract;
- the handler reads computed scalar values from the mounted source range and writes one native
  `setValues` matrix to an equally shaped destination range. Formulas therefore become values while
  visible Ribbon paste-special remains on Univer's native clipboard route;
- source and destination may be on different worksheets. Both ranges are normalized within XLSX
  row/column limits, must have identical dimensions, and are capped at 20,000 cells;
- source reads and destination overwrites fail closed while either file-backed rectangle is still
  streaming. The checks run before any source read or destination write;
- the real-host tracer copies a formula result plus scalar values, shares UI Undo/Redo, advances one
  revision, and saves a destination cell with `<v>` but no `<f>`;
- Agent `ribbon_command` calls for `paste-special:value` are rejected in favor of the exact
  operation. Generated discovery, Broker validation/queueing, renderer validation, cross-sheet
  execution, and old-string rejection are covered;
- the generated product manifest now contains seventy-two operations, including fifty-four XLSX
  operations.

R2-73 evidence:

- `xlsx.range.copy_formulas { sourceSheet, sourceRange, destinationSheet, destinationRange }`
  replaces the formula-only clipboard dependency with an exact source/destination contract;
- formulas are translated by Univer's mounted formula lexer using the source-to-destination row and
  column offsets, preserving absolute components while moving relative and mixed references. As in
  the retained Paste Formulas command, non-formula source cells copy their scalar values;
- formula and value copy share one range-pair boundary: worksheet-bounded A1 normalization, equal
  dimensions, a 20,000-cell limit, cross-sheet targets, and source/destination streaming guards;
- the destination is committed by one native `setValues` matrix, so formulas, scalar cells,
  selection, UI Undo/Redo, and save remain on the mounted Univer route;
- the real-host tracer copies relative, absolute, and mixed references, proves visible Undo/Redo,
  advances one Agent revision, and saves translated destination `<f>` elements;
- Agent `ribbon_command` calls for `paste-special:formula` are rejected in favor of the exact
  operation. The visible Ribbon retains its native Univer clipboard command;
- the generated product manifest now contains seventy-three operations, including fifty-five XLSX
  operations (fifty-four Agent-visible and one internal).

R2-74 evidence:

- `xlsx.range.copy_formats { sourceSheet, sourceRange, destinationSheet, destinationRange }`
  replaces clipboard-dependent Agent format pasting with an exact source/destination contract;
- it shares worksheet-bounded A1 normalization, equal dimensions, the 20,000-cell limit,
  cross-sheet targets, and source/destination streaming guards with value and formula copy;
- every source cell contributes a complete style replacement: absent source fields explicitly clear
  destination fields instead of merging onto them. Values, formulas, and rich cell content remain
  unchanged; merged-cell topology remains owned by `xlsx.range.merge`;
- the format matrix uses Univer's set-range-values mutation directly, with the exact inverse and redo
  payload registered as one native history item. Destination formulas ride in that payload unchanged
  so renderer journaling cannot misclassify recalculation results as value edits;
- the real-host tracer copies font, fill, alignment, wrapping, and number format, proves visible
  Undo/Redo, preserves a destination scalar and formula, and verifies saved style/value/formula
  output;
- Agent `ribbon_command` calls for `paste-special:format` are rejected in favor of the exact
  operation. Generated discovery, Broker validation/queueing, renderer bounds, and old-string
  rejection are covered;
- the generated product manifest now contains seventy-four operations, including fifty-six XLSX
  operations (fifty-five Agent-visible and one internal).

R2-75 evidence:

- `xlsx.column.copy_widths { sourceSheet, sourceColumn, destinationSheet, destinationColumn, count }`
  replaces clipboard-dependent Agent column-width pasting with exact source/destination spans;
- source and destination use normalized A1 column labels, have independent worksheet-bound checks,
  and share the Contract's 1–10,000 structural count bound;
- the renderer snapshots every source width before applying one object-backed Univer
  set-column-width mutation, so overlapping and cross-sheet copies remain deterministic;
- one exact inverse/redo pair is registered in the mounted workbook history, while the existing
  axis journal persists the visible result without touching cell values or formulas;
- the real-host tracer copies two distinct widths, proves visible UI Undo/Redo, preserves a scalar
  and formula, and verifies saved target widths equal the source widths;
- Agent `ribbon_command` calls for `paste-special:col-width` are rejected in favor of the exact
  operation. Generated discovery, Broker count validation/queueing, renderer bounds, and old-string
  rejection are covered;
- the generated product manifest now contains seventy-five operations, including fifty-seven XLSX
  operations (fifty-six Agent-visible and one internal).

R2-76 evidence:

- `xlsx.range.copy_without_borders { sourceSheet, sourceRange, destinationSheet, destinationRange }`
  replaces clipboard-dependent Agent “paste except borders” with exact source/destination ranges;
- it reuses the range-copy family's worksheet-bounded A1 normalization, equal-shape constraint,
  20,000-cell ceiling, cross-sheet support, and independent source/destination streaming guards;
- source values, rich content, formulas, and non-border styles are snapshotted before one native
  range write. Formula references move by the exact row/column offset, all non-border style fields
  replace the destination, and omission of `bd` retains each destination border;
- inactive-sheet formulas use the read-only file index only as a missing-model fallback, while
  current edit-journal content wins. The mounted workbook remains the editable authority;
- the real-host tracer proves cross-sheet value/formula/style copying, visible UI Undo/Redo, saved
  formula translation, and saved target-border preservation. Generated discovery, Broker
  queue/schema rejection, renderer copy bounds, and old-string rejection are covered;
- Agent `ribbon_command` calls for `paste-special:besides-border` are rejected; the visible Ribbon
  keeps Univer's retained clipboard command;
- the current generated product manifest contains eighty-two operations, including fifty-eight
  XLSX operations (fifty-seven Agent-visible and one internal).

R2-77 evidence:

- Agent `ribbon_command` calls for `copy`, `cut`, and `paste` are rejected as clipboard/selection-
  state UI gestures and point callers to explicit source/destination XLSX operations;
- the retained Ribbon buttons continue to execute Univer's native clipboard commands, so the user
  workflow and its internal clipboard cache are unchanged;
- a real-host rejection tracer proves the three strings cannot receive positive Agent
  acknowledgement. No descriptor is added, so the generated product manifest remains at
  eighty-two operations, including fifty-eight XLSX operations.

R2-78 evidence:

- the browser workbook adapter now accepts the existing `hyperlinkEdits` save journal instead of
  rejecting it, and delegates worksheet/relationship XML changes to the retained hyperlink gateway;
- browser open parses external relationship targets and internal locations into the same streamed
  hyperlink context consumed by the mounted App. User clicks after reopen therefore resolve the
  saved target without a second link model;
- the real-host tracer proves `xlsx.hyperlink.set` → `xlsx.document.save` → OOXML relationship
  output → browser reopen → visible user click. Focused host tests cover set and remove across
  reopen. No descriptor is added, so the manifest remains at eighty-two operations.

R2-79 through R2-81 evidence:

- `xlsx.table.add` now saves through a transactional browser-package overlay over the retained
  table gateway. Native table parts, worksheet relationships, `tableParts`, and content-type
  overrides are committed only after validation succeeds;
- browser open resolves table relationships and hydrates range, header count, built-in style name,
  and row/column stripe flags into the retained Univer table path. Both file-native and newly saved
  tables reopen as tables rather than plain cells;
- cell protection continues through the neutral style journal and retained stylesheet editor;
  sheet protection now uses the retained OOXML gateway, and browser open hydrates the protected bit
  plus password presence so explicit unprotect remains correct and passworded files fail closed;
- focused host tests and real-host tracers prove table and protection typed save, exact OOXML,
  reopen, visible table styling, and explicit unprotect after reopen. No descriptor is added, so
  the manifest remains at eighty-two operations.

R2-82 evidence:

- `xlsx.sparkline.add { sheet, sourceRange, targetRange, type }` replaces the retained
  `sparkline:line|column|stacked` strings with explicit row-aligned source and one-column target
  ranges, a 200-row bound, overlap/collision checks, and an exact type enum;
- Ribbon and Agent use one XLSX-owned action that records the existing sparkline journal, pushes
  one renderer-owned Univer history item, refreshes the retained float-DOM projection, and
  activates the explicit target;
- the browser package adapter delegates x14 output to the retained sparkline gateway and parses
  saved group type, colors, source refs, and host cells back into the mounted renderer on reopen;
- focused registry/browser tests and a real-host tracer prove exact validation, journal state,
  visible output, typed save, OOXML output, and same-iframe reopen. The product manifest now has
  eighty-three operations, including fifty-nine XLSX operations.

R2-83 evidence:

- `xlsx.outline.set_level { sheet, axis, start, count, level }` replaces Group/Ungroup Agent
  gestures with an absolute `0..7` level over a 1-based bounded row or column span;
- `xlsx.outline.set_detail_visibility { sheet, axis, start, count, hidden }` replaces Hide/Show
  Detail with an explicit final state. The immediately following row/column is the summary item and
  is updated with the matching collapsed state;
- the retained Ribbon computes its relative level changes and then calls the same XLSX-owned
  actions. File-journal state, live visibility, and the outline map enter Univer history as one
  renderer-owned undo item instead of splitting the change across native and custom stacks;
- browser open now parses row height/hidden/outline/collapsed metadata and column
  width/hidden/outline/collapsed spans. Focused host tests and a real-host tracer prove collapse,
  exact OOXML, typed save, reopen hydration, explicit expansion, and a second save;
- all eight `outline-*` Agent strings are rejected in favor of the two exact operations. The
  product manifest now has eighty-five operations, including sixty-one XLSX operations.

R2-84 evidence:

- `xlsx.range.set_checkbox { sheet, range, enabled }` replaces `insert-checkbox` with a normalized
  explicit range, a 10,000-cell bound, a streaming/validation-metadata gate, and a replayable final
  boolean state;
- Ribbon and Agent call one XLSX-owned action over Univer's native data-validation command and Undo
  stack. `enabled: false` explicitly removes validation from the target range;
- the browser package adapter now accepts the declarative `dvStates` save journal, delegates output
  to the retained `xlsx-dv` writer, parses base OOXML validation rules on reopen, and restores the
  writer's `list "1,0"` checkbox representation to Univer checkbox state;
- focused Registry/browser tests and a real-host tracer prove enable, exact OOXML, save, reopen,
  explicit removal, and second save. The product manifest now has eighty-six operations, including
  sixty-two XLSX operations.

R2-85 evidence:

- the retained Symbol dialog is a UI text picker, not a separate document capability. Its final
  mutation is the same scalar cell assignment already exposed as `xlsx.cell.set_value`;
- the dialog's append calculation and the Registry handler now converge on one native cell
  write/activation action, retaining Univer Undo and the existing cell save/reopen route;
- Agent `insert-symbol` is rejected in favor of an explicit final cell value. Registry and
  real-host rejection tests prove the mapping; no descriptor is added, so the manifest remains at
  eighty-six operations, including sixty-two XLSX operations.

R2-86 evidence:

- `import-csv` is a browser file-picker/decoding gesture, not a deterministic Agent operation. An
  Agent that already owns parsed tabular values uses bounded `xlsx.range.set_values`; a CSV that
  should become its own workbook uses the staged-open transport;
- the retained visible import still performs charset/delimiter sniffing, numeric coercion, and its
  50,000-cell UI bound, but its resolved matrix and `xlsx.range.set_values` now share one native
  range-write/activation action and Univer Undo route;
- Agent `import-csv` is rejected in favor of the explicit matrix operation. Registry and real-host
  rejection coverage proves the mapping; no descriptor is added.

R2-87 evidence:

- `xlsx.chart.add { sheet, dataRange, type, anchorCell? }` exposes the retained chart mutation as
  an explicit operation with nine chart types and a 2,000-source-cell bound;
- the handler resolves worksheet identity and delegates to the existing visual journal, native
  history item, and renderer projection used by Ribbon chart insertion. `insert-chart:*` and the
  recommended-chart picker command are rejected for Agent callers in favor of final chart intent;
- the browser workbook now accepts visual-addition journal entries through the retained
  transactional drawing writer. Focused and real-host tests prove immediate rendering, exact chart
  and drawing OOXML, typed save, and successful reopen. The product manifest now has eighty-seven
  operations, including sixty-three XLSX operations.

PDF-P1 through PDF-P4 evidence:

- `pdf.document.save {}` is exact and returns `{ saved: true }` without inventing a filename. Both
  Agent and user saves reach the mounted App controller's existing persistence seam;
- internal `pdf.document.load_staged { blobId, name, size, data }` validates the hydrated staged
  payload and calls the mounted App's retained `consumePending`/`readFile` load pipeline. It remains
  absent from Agent discovery and direct `office_execute` calls are rejected;
- `pdf.annotation.delete_saved { pageIndex, objNum, subtype, rect }` validates the exact retained
  annotation family and calls the controller route that records the edit in mounted App history;
- `pdf.history.undo {}` calls the same mounted App undo action as the retained UI. Save, staged load,
  saved-annotation deletion, and undo each have one catalog descriptor and one executable handler;
- legacy `save`, `open_local_file`, `delete_saved_annotation`, and `undo` are input-only aliases.
  The replaced browser-host and command-bridge execution branches are deleted;
- the generated product manifest supplies PDF discovery and Broker pre-enqueue validation. MCP
  tests cover canonical discovery, legacy normalization, exact-schema rejection, acknowledged
  outputs and revisions, internal-operation rejection, and hydrated staged-file transport;
- PDF-focused registry/host tests, the full PDF workspace suite, and PDF/MCP typechecks pass. The
  Broker neither parses nor edits PDF bytes, and the renderer adds no second PDF state model, DOM
  clicking path, or open string extension point.

Completed PDF retained-command evidence:

- the PDF-owned catalog and executable registry now contain 25 descriptors: 23 Agent-visible
  operations plus internal document and page staged-byte routes;
- exact schemas cover document metadata/save, undo/redo, markup and pending deletion, drawings,
  notes/signatures, text, image lifecycle, static and AcroForms, generated stamps, and page
  lifecycle;
- browser PDFium restores searchable content-stream text/image persistence. Gzip-compressed WASM
  ships inside the editor, while allowlisted script fonts load lazily and browser cmap validation
  rejects unsupported glyphs before mutation;
- `pdfRetainedProducerBaseline` classifies all retained producer families and has no missing or
  unknown operation mapping. Save As/export/extract/print are non-document host effects; navigation,
  search, zoom, sidebars, view modes, and form focus are view-only;
- 23 PDF test files / 280 assertions and eleven real-host scenarios prove staged/user open, shared
  delete/undo, text/image lifecycle, CJK/Korean/Arabic plus mixed colors, generated-stamp clear,
  save/reopen, and four host-width states.

PPTX-P1 through PPTX-P2 evidence:

- `pptx.document.save {}` declares exact empty input, exact
  `{ saved: true, fileName: string }` output, document context, persistence effects, medium risk,
  non-undoable atomic semantics, and the legacy input-only `save` alias;
- internal `pptx.document.load_staged { blobId, name, size, data }` validates hydrated
  `ArrayBuffer` identity, `.pptx` extension, and declared byte length before calling the mounted
  `BrowserPresentation.open` path. `open_local_file` remains transport-only;
- the browser host injects the existing format-owned save and load services into the registry.
  Its former direct `save` and `open_local_file` operation branches are deleted, so user and Agent
  persistence converge without a second presentation model;
- the generated manifest owns PPTX discovery, alias normalization, exact Broker validation, and
  internal-operation rejection. Capability discovery advertises canonical save and hides staged
  load;
- focused registry tests cover descriptor/handler parity, canonical and legacy ids, exact input,
  success output, and cancellation/failure mapping. A real PPTX fixture proves the public host
  adapter loads into mounted presentation state, and Slides/MCP/Contract checks pass.

R2-239 through R2-253 complete the initial executable PPTX slice. Nineteen descriptors cover
document creation/load/save/save-as, selection, native undo/redo, slide lifecycle, explicit object
delete/EMU transform, selected movement, rich text, font/paragraph formatting, and scoped
find/replace. Legacy `select_objects`, `replace_selected_text`, and `move_selected_objects` resolve
only as input aliases. The browser adapter has no direct operation-id branches and the central MCP
catalog contains no PPTX-specific schemas. Real-fixture tests cover state, renderer refresh,
recovery and native history.

R2-254 through R2-308 complete the retained PPTX mutation inventory. Seventy-four descriptors
(73 Agent-visible plus internal staged load) cover slide/layout state, object and clipboard
actions, parent-addressed group children, connector/picture/table/chart/SmartArt/theme/animation,
notes/comments/sections/hyperlinks, embedded image/media/3D/ink, and master/layout-part edits. The
browser host implements the complete retained `SlidesApi` at compile time; picker actions converge
on bounded byte primitives; Presenter/Audience share browser presentation state; and export,
print, and show are classified as non-mutating host effects. A machine-checked retained-producer
baseline maps every descriptor with no `missing` disposition. Slides typecheck, 170 focused tests,
Manifest generation/check, production build, plugin packaging, and the resource guard pass. This
closes PPTX format-local parity while shared release gates keep `ready` false.

R2-88 adds public `xlsx.image.add { sheet, path, anchorCell }` and internal
`xlsx.image.add_staged`. Broker staging validates absolute PNG/JPEG/GIF paths, a 20 MB limit, and
magic bytes, then the live bridge hydrates bytes outside Agent JSON. Both retained UI producers and
the internal handler share the visual journal and Undo route. Broker, Registry, package, and real
host tests prove rendering, non-empty media/drawing OOXML, typed save, and reopen. The product
manifest now has eighty-nine operations, including sixty-five XLSX operations.

R2-89 adds `xlsx.note.set { sheet, address, text }` and
`xlsx.note.remove { sheet, address }`. Both resolve an explicit cell, call Univer's native note
commands and history, and mark the existing declarative note journal. Browser open/save now reads
and writes legacy comments plus their VML anchors. Focused and real-host tracers prove set, typed
save, exact OOXML, reopen hydration, native removal, and second save. Note previous/next/show remain
UI navigation, not Agent tools. The product manifest now has ninety-one operations, including
sixty-seven XLSX operations.

R2-90 adds `xlsx.chart.update { chartId, ...finalState }` and
`xlsx.chart.remove { chartId }`. Update bounds title/type/legend/data-label/grouping/axis-title
properties; remove uses the same explicit ID returned by chart insertion. Both call the retained
chart/visual edit refs and native history. Browser save now consumes `chartEdits` and `visualEdits`;
real-host tracers prove session chart update and removal. The product manifest now has ninety-three
operations, including sixty-nine XLSX operations. Palette and series replacement remain separate.

R2-91 adds `xlsx.chart.set_colors { chartId, seriesColors?, pointColors? }` and
`xlsx.chart.set_series { chartId, series }`. Colors are explicit hex arrays bounded to 24 series
or 64 pie/doughnut points; series replacement is bounded to 24 series and 1,000 values each.
Palette, switch-row/column, and Select Data Agent strings now resolve to these final-state
operations. The real-host tracer proves final series caches and colors in saved chart OOXML. The
product manifest now has ninety-five operations, including seventy-one XLSX operations.

R2-92 hydrates browser worksheet drawing/chart relationships after reopen, exposes stable
`file-chart-*` IDs in live context, and extends `xlsx.chart.update` with gridlines, value-axis
bounds, gap width, hole size, pie explosion, and data-label position/format. The real-host tracer
uses the reopened ID for a second edit and verifies the second saved chart part. The product
manifest remains at ninety-five operations, including seventy-one XLSX operations.

R2-93 adds `xlsx.shape.add { sheet, type, anchorCell, fillColor?, text? }`. It uses the bounded
retained gallery enum plus `textbox`, explicit cell placement, and the existing visual Undo/save
route. Browser reopen hydrates shapes and deterministic `file-shape-*` IDs. The product manifest
now contains ninety-six operations, including seventy-two XLSX operations.

R2-94 adds `xlsx.shape.update` and `xlsx.shape.remove` over stable session/reopened IDs. The visual
journal and drawing gateway now preserve file-native shape position, text, fill, and deletion with
renderer Undo and repeated save/reopen evidence. The manifest contains ninety-eight operations,
including seventy-four XLSX operations.

R2-95 hydrates browser image relationships, stable IDs, and package media bytes, then adds
`xlsx.image.move` and `xlsx.image.remove` over the shared visual history/drawing save route. The
manifest contains one hundred operations, including seventy-six XLSX operations.

R2-96 adds `xlsx.defined_name.set` and `xlsx.defined_name.remove`. `set` is an explicit scoped
upsert and accepts `previousName` for one-step rename plus formula update; missing `scopeSheet`
means workbook scope. The retained Name Manager and registry share the same Univer defined-name
action layer, native mutation/history events, declarative save snapshot, and browser XML
save/reopen path. Agent `name-manager-open` is rejected in favor of these exact operations. The
manifest contains one hundred two operations, including seventy-eight XLSX operations.

R2-97 adds `xlsx.range.set_list_validation` with explicit inline values, blank/dropdown final
states, a 10,000-cell target ceiling, and Excel's 255-character inline-source limit. It also adds
`xlsx.range.remove_data_validation` for explicit removal of any rule. Both share the checkbox
tracer's native Univer validation/history and metadata-loading guard, plus declarative browser
save/reopen/removal evidence. Agent `dv-open` now fails closed with exact replacements. The
manifest contains one hundred four operations, including eighty XLSX operations.

R2-98 adds exact `xlsx.range.set_number_between_validation`,
`xlsx.range.set_date_between_validation`, and `xlsx.range.set_custom_formula_validation` tracers.
Numeric bounds must be finite and ascending; dates must be real ascending ISO calendar dates; custom
formulas must be equals-prefixed and at most 8,192 characters. All three reuse the shared native
validation target guard and Univer history/save mapping. The manifest contains one hundred seven
operations, including eighty-three XLSX operations.

R2-99 adds `xlsx.range.set_list_reference_validation` with explicit target and source sheet/range,
blank/dropdown final states, a single-axis source requirement, and a 1,000-source-cell ceiling.
Target and source streaming are checked independently before the native builder/history route runs;
browser save/reopen proves the range formula. The manifest contains one hundred eight operations,
including eighty-four XLSX operations.

R2-100 adds one `xlsx.range.set_comparison_validation` operation for `whole`, `decimal`, `date`,
`time`, and `textLength` with all eight retained Excel comparison operators. Runtime validation
enforces operator arity, type-specific literals, real dates/times, text-length bounds, and ascending
binary operands. It uses the native validation panel's rule model and Undo command. The narrower
R2-98 number/date-between descriptors and helpers are deleted, so five kinds and eight operators
replace them with one descriptor instead of growing a tool matrix. The manifest returns to one
hundred seven operations, including eighty-three XLSX operations.

R2-101 adds `xlsx.range.set_validation_messages` over one existing explicit rule. Nullable prompt
and error fields are required final state, titles are capped at 32 characters, messages at 255,
and `errorStyle: none` requires cleared error text. The operation invokes FDataValidation's native
`setOptions` update/Undo path and the existing declarative save mapping. The manifest contains one
hundred eight operations, including eighty-four XLSX operations.

R2-102 starts the Conditional Formatting migration with
`xlsx.conditional_format.set_comparison` and `xlsx.conditional_format.remove`. Creation uses
`ruleId: null`; updates require a session rule ID published by workbook context, accept only an
existing numeric-comparison rule, and replace its explicit range, operator, operands, visible
format, and `stopIfTrue` state. Both operations require loaded range/rule metadata, use Univer's
native add/set/delete commands and Undo stack, and share the declarative OOXML CF save route. The
browser package host now accepts CF save states; a focused round trip proves comparison creation,
DXF persistence, reopen, and removal. Agent `cf-open` is rejected in favor of explicit operations.
The manifest contains one hundred ten operations, including eighty-six XLSX operations.

R2-103 adds one `xlsx.conditional_format.set_highlight` lifecycle operation for eight retained
predicates: contains/not-contains/starts-with/ends-with text, blank/non-blank, and
duplicate/unique. The explicit nullable `text` field must match predicate semantics; text is bounded
to 255 characters. It reuses R2-102's visible-format, rule-ID, 10,000-cell, streaming, metadata,
native history, context-discovery, and CF/DXF save boundaries. Existing gateway coverage proves all
eight wire shapes without introducing one descriptor per predicate. The manifest contains one
hundred eleven operations, including eighty-seven XLSX operations.

R2-104 adds one `xlsx.conditional_format.set_statistical` lifecycle operation for rank and average
rules. Rank state requires top/bottom, an integer amount (1..1,000, or 1..100 for percent), explicit
percent mode, and `inclusive: null`. Average state requires above/below, explicit inclusive mode,
and null rank/percent fields. This mutually exclusive final-state contract reuses the shared rule
identity, format, target, native history, context, and persistence boundaries. The manifest contains
one hundred twelve operations, including eighty-eight XLSX operations.

R2-105 adds `xlsx.conditional_format.set_formula` for a 2..8,192-character equals-prefixed custom
formula. It creates or updates only formula rules through the shared range, format, session ID,
native Undo, context, and CF/DXF persistence boundaries. The manifest contains one hundred thirteen
operations, including eighty-nine XLSX operations.

R2-106 adds one `xlsx.conditional_format.set_visual` lifecycle operation for color scales, data
bars, and base-OOXML icon sets. Explicit bounded colors and typed thresholds are validated by kind;
data bars allow only the saveable gradient form, and icon sets are restricted to the existing
OOXML whitelist with coherent 3/4/5 threshold counts, value visibility, and reversal. Extended x14
shapes remain fail-closed. The operation uses the shared rule identity, target, native history,
context, and desktop/browser save route. The manifest contains one hundred fourteen operations,
including ninety XLSX operations.

R2-107 completes the saveable highlight audit by extending `set_highlight` with text equal/not-equal
and error/non-error predicates. Reopen hydration translates those file rules into equivalent native
formula rules so their behavior remains visible. The shared UI pre-command/save validator now also
rejects solid data bars and distinct negative colors because they require x14; existing tests already
reject date-occurring, unsupported icon sets/orderings, and equal/not-equal average rules. No new
descriptor is added, so the manifest remains one hundred fourteen operations, including ninety XLSX
operations.

R2-108 adds `xlsx.conditional_format.clear { sheet, scope, range }`. Range scope requires one
bounded A1 target; sheet scope requires `range: null`. Both paths count the affected rules and use
Univer's native range/worksheet clear commands and history. R2-109 adds
`xlsx.conditional_format.set_priority { sheet, ruleId, position }`, replacing relative panel moves
with an explicit one-based final position over the published rule identities. Together they raise
the product manifest to one hundred sixteen operations, including ninety-two XLSX operations.

R2-110 removes the open-ended Agent `ribbon_command` schema, renderer branch, and replacement map.
All already migrated mutations use their typed registry operations; UI-only pickers and view
gestures remain inside the retained Ribbon. The descriptor count is unchanged.

R2-111 through R2-116 establish browser Pivot definition load/save and complete the bounded Pivot
add/refresh/update/member-filter/PivotChart lifecycle before running the first post-Ribbon gap audit.
R2-117 through R2-126 add explicit formula writes, worksheet duplicate/visibility, row/column
visibility, full table lifecycle, bounded text replacement, header/footer, subtotals,
consolidation, and PivotChart persistence. R2-127 through R2-130 close native range movement, row
movement, worksheet tab color, and the browser save/reopen adapter for duplication/visibility.

R2-131 performs the final XLSX retained-command audit. Executable tests map Ribbon/dialog strings
and native grid/table/Pivot/worksheet gestures to exact catalog IDs, and every exact replacement is
required to exist in the format-owned catalog. The generated product Manifest contains 138
operations: 114 XLSX operations (112 Agent-visible and two internal). No unexplained retained XLSX
state-changing command remains; cross-format visual/performance/release gates still keep product
readiness false.

R2-132 establishes the retained DOCX command-parity inventory. It groups visible controls by the
native document transaction they produce, separates renderer-local gestures from document
mutations, and records fourteen open semantic families instead of treating the former structured
command surface as complete. R2-133 closes the first family with `docx.history.undo {}` and
`docx.history.redo {}`. Both use exact empty-input contracts, return an explicit final result, run
through the mounted TipTap history shared with the retained UI, and fail closed when the requested
history entry is unavailable. The generated product Manifest now contains 140 operations: sixteen
DOCX operations (fifteen Agent-visible and one internal) plus the unchanged 114-operation XLSX
catalog.

R2-134 starts the exact DOCX character-format family with
`docx.text.set_character_format { range, format, fields }`. The range is an explicit pair of
revision-scoped ProseMirror positions already exposed by session context. The exact unique field
mask supports final bold, italic, underline, strike, and baseline/superscript/subscript state; its
single transaction preserves unlisted heterogeneous run attributes. The retained Ribbon now uses
the same format-owned helper for non-empty selections, while collapsed-caret stored marks remain a
renderer-local input mode. The generated Manifest contains 141 operations, including seventeen
DOCX operations (sixteen Agent-visible and one internal).

R2-135 deepens the same exact character-format operation with bounded font family, half-point font
size from 1 through 1638pt, `#RRGGBB` color, and the sixteen DOCX named highlight values. Nullable
fields clear their final state. The shared transaction merges each selected text node's existing
`docTextStyle` independently, so unlisted script fonts, character style IDs, spacing, and raw run
metadata do not flatten across a heterogeneous selection. The operation and Manifest counts do not
grow.

R2-136 migrates the retained Clear Formatting action as
`docx.text.clear_character_format { range }`. It reuses the exact revision-scoped range validator,
removes marks only inside that range in one native transaction, and is now the non-empty-selection
route used by the Ribbon. Invalid or textless ranges fail before mutation. The generated Manifest
contains 142 operations, including eighteen DOCX operations (seventeen Agent-visible and one
internal).

R2-164 adds `docx.section.set_columns { sectionIndex, count, spacingTwips }`. Both values are
bounded, aggregate gaps must leave positive text width, and the DOCX engine now writes the explicit
`w:space` instead of its former 425-twip fallback. Registry, Ribbon, native Undo, and save/reopen
share the section journal. The generated Manifest contains 168 operations, including forty-four
DOCX operations (forty-three Agent-visible and one internal).

R2-165 adds `docx.section.set_page_border { sectionIndex, enabled }`. The retained Design-tab
toggle and Agent route share the Undo-owned section journal and existing `w:pgBorders` writer. No
other retained UI mutates `headerDist`, `footerDist`, `vAlign`, `docGrid`, or `textDirection`, so
those parsed fields remain preservation-only. The generated Manifest contains 169 operations,
including forty-five DOCX operations (forty-four Agent-visible and one internal); the retained
page/section family is closed.

R2-166 adds `docx.section.set_different_first_page { sectionIndex, enabled }`. The indexed boolean
is stored on the section anchor for Undo, synchronized with the retained first-page toggle, and
projected to `w:titlePg` for non-final or final sections during save. The generated Manifest
contains 170 operations, including forty-six DOCX operations (forty-five Agent-visible and one
internal); the header/footer family is now active.

R2-167 adds `docx.document.set_different_odd_even_pages { enabled }`. The explicit document-level
final state is stored on the first section anchor for native Undo and projected to
`settings.xml/w:evenAndOddHeaders` during save. Registry execution and the retained Ribbon checkbox
share the same helper; save/reopen preserves the state. The generated Manifest contains 171
operations, including forty-seven DOCX operations (forty-six Agent-visible and one internal).

R2-168 adds `docx.section.set_page_numbering { sectionIndex, format, start }`. Seven retained page
number formats and a nullable bounded start value replace dialog strings and the former
`pgNumEdit`/`pgNumDirtySections` save split. The dialog and registry now share the native section
journal; non-final and final `sectPr` writes both reopen exactly. The generated Manifest contains
172 operations, including forty-eight DOCX operations (forty-seven Agent-visible and one internal).

R2-169 adds `docx.header_footer.set_text { sectionIndex, kind, variant, text }`. The explicit
section/kind/variant target writes a bounded renderer text value through the new Undo-owned
`headerFooterEdits` journal. Retained canvas/Ribbon writers share the same full-value kernel, old
React dirty save channels are removed, and the engine's non-final section writer now preserves
`default|first|even` identities. The generated Manifest contains 173 operations, including
forty-nine DOCX operations (forty-eight Agent-visible and one internal).

R2-170 adds
`docx.header_footer.set_page_number { sectionIndex, kind, variant, enabled, alignment }`. The
explicit final state replaces gallery position strings; PAGE uses the private sentinel, and removal
preserves all non-page content. Ribbon and Agent share the same content journal and persistence
route. The generated Manifest contains 174 operations, including fifty DOCX operations (forty-nine
Agent-visible and one internal).

R2-171 adds `docx.header_footer.set_paragraphs { sectionIndex, kind, variant, paragraphs }`. The
bounded paragraph/segment model covers text, PAGE/NUMPAGES fields, alignment, and retained run
styles without exposing OOXML. It uses the same full-value journal as HeaderFooterArea, validates
aggregate resource size in the renderer, and round-trips rich content. The generated Manifest
contains 175 operations, including fifty-one DOCX operations (fifty Agent-visible and one internal);
the retained header/footer family is closed.

R2-172 adds public `docx.image.insert { path, afterBlockIndex, widthPx, heightPx, alignment }` plus
internal `docx.image.insert_staged`. Broker-owned session staging keeps up to 20 MiB of PNG/JPEG/GIF
bytes outside the public operation envelope; the live DOCX host hydrates them immediately before
dispatch. Renderer byte/magic validation, a stable top-level boundary, bounded 1–4096px axes, and
explicit alignment produce one native image node and Undo entry. Ribbon/dialog/paste and Agent
paths share the node builder, and save/reopen/resave is covered. The generated Manifest contains
177 operations, including fifty-three DOCX operations (fifty-one Agent-visible and two internal).

R2-173 adds public `docx.image.replace { path, imageBlockIndex, widthPx, heightPx }` plus internal
`docx.image.replace_staged`. It reuses Broker staging and renderer byte validation, but targets one
exact native image and declares final geometry. A shared attrs kernel preserves placement/wrap and
other image properties while clearing stale crop windows, then chooses the original-image surgical
patch or unsaved `genImage` route. Ribbon, crop/cutout, and Agent replacement share this kernel;
Undo and save/reopen are covered. The generated Manifest contains 179 operations, including
fifty-five DOCX operations (fifty-two Agent-visible and three internal).

R2-174 adds `docx.image.set_wrap { imageBlockIndex, wrap }`. Its retained finite enum includes
explicit `null` inline state. Inline clears all named/offset position attrs; floating wrap changes
preserve position. ContextMenu, Layout, Picture Format, and Agent routes share one exact-index attrs
kernel and native Undo, and the existing save/reopen image fixture now executes through the
registry. The generated Manifest contains 180 operations, including fifty-six DOCX operations
(fifty-three Agent-visible and three internal).

R2-175 adds `docx.image.set_margin_position { imageBlockIndex, horizontal, vertical }` for the
retained nine margin-relative presets. It derives the exact square wrap, sets both named axes, and
clears free offsets through one registry/UI writer and native Undo. Original-image-only validation
matches the retained gallery and prevents generated images from advertising state their saver
cannot yet encode. Save/reopen/resave is covered. The generated Manifest contains 181 operations,
including fifty-seven DOCX operations (fifty-four Agent-visible and three internal).

R2-176 adds
`docx.image.set_offset_position { imageBlockIndex, wrap, offsetXEmu, offsetYEmu }`. It accepts the
complete retained non-null image-wrap enum and signed 32-bit EMU coordinates, clears named axes,
and supports both original and generated image nodes. Registry execution and retained image drag
share one pure attrs projection; textbox/shape drag stays outside the image operation. Native Undo
and original-image save/reopen are covered. The generated Manifest contains 182 operations,
including fifty-eight DOCX operations (fifty-five Agent-visible and three internal).

R2-177 adds
`docx.image.set_transform { imageBlockIndex, rotationDegrees, flipHorizontal, flipVertical }` as
explicit final state. Rotation is bounded to integer `0..359`; both flip values are required.
Registry and retained Rotate/Flip controls share one canonical attrs projection and native Undo.
The image patch writer now creates a minimal picture transform container when the source omitted
one, so transform edits cannot disappear on save. Original/generated state and original-image
save/reopen are covered. The generated Manifest contains 183 operations, including fifty-nine
DOCX operations (fifty-six Agent-visible and three internal).

R2-178 adds `docx.image.set_crop { imageBlockIndex, left, top, right, bottom }` with bounded source
fractions, positive-area cross-field validation, and all-zero reset. Registry and the retained Crop
dialog share non-destructive image attrs rather than re-encoding bytes. Original-image patches and
new-image generation write `a:srcRect`; stale fill windows are cleared, native Undo is preserved,
and save/reopen is covered. The generated Manifest contains 184 operations, including sixty DOCX
operations (fifty-seven Agent-visible and three internal).

R2-179 adds `docx.image.remove { imageBlockIndex }`. It rejects non-image targets, deletes one
exact native image, and inserts a replacement empty paragraph in the same transaction when the
image was the sole block. Native Undo and save/reopen deletion are covered. The generated Manifest
contains 185 operations, including sixty-one DOCX operations (fifty-eight Agent-visible and three
internal), closing the retained image lifecycle.

R2-180 adds `docx.shape.insert { afterBlockIndex, preset, widthEmu, heightEmu }` for all 104 filled
Gallery presets. Five stroke-only line/connector kinds stay outside the enum for their own exact
operation. Registry, retained Gallery, and draw mode share one renderer-owned node builder; stable
top-level insertion, bounded EMU axes, native Undo, and save/reopen are covered. The generated
Manifest contains 186 operations, including sixty-two DOCX operations (fifty-nine Agent-visible
and three internal).

R2-181 adds `docx.line.insert { afterBlockIndex, kind, widthEmu, heightEmu }` for all five retained
line/connector kinds. Straight kinds require the canonical 114,300 EMU grab height, making the
final state explicit; bent/curved connectors retain both bounded axes. Registry and retained
Gallery/draw mode share one stroke-only builder, native Undo, and save/reopen. The generated
Manifest contains 187 operations, including sixty-three DOCX operations (sixty Agent-visible and
three internal).

R2-182 adds `docx.textbox.insert { afterBlockIndex, widthEmu, heightEmu }`. Stable top-level
insertion and both 9,525–20,000,000 EMU axes are validated before dispatch. Registry and retained
Ribbon share one format-owned textbox node builder, one native Undo entry, and save/reopen
projection. The generated Manifest contains 188 operations, including sixty-four DOCX operations
(sixty-one Agent-visible and three internal).

R2-183 adds bounded `docx.chart.insert` for bar, line, and pie data matrices with explicit final
extent. The renderer validates stable insertion, 1–256 categories, 1–64 series, a 4,096-value
matrix budget, equal category/value dimensions, single-series pie semantics, and bounded finite
numbers. Registry and the retained Chart dialog share the same node builder, native Undo, and
embedded workbook save/reopen route. The generated Manifest contains 189 operations, including
sixty-five DOCX operations (sixty-two Agent-visible and three internal).

R2-184 extends the shared Schema validator with Unicode-aware `minLength`/`maxLength` and adds
`docx.equation.insert` for exact display-block or inline-range placement. Mutually exclusive
nullable coordinates make both forms replayable in one contract; the renderer validates the
same-paragraph inline range and parses bounded LaTeX before mutation. Gallery, modal, and Registry
share LaTeX-to-OMML construction, native Undo, and save/reopen. The generated Manifest contains
190 operations, including sixty-six DOCX operations (sixty-three Agent-visible and three internal).

R2-185 adds aggregate `docx.object.set_size` rather than one tool per drawing kind. One exact
block-indexed kernel serves retained shape, line, textbox/WordArt-like textbox, and chart corner
resizing. Shared bounds preserve chart and straight-line invariants, changed operations start an
independent native Undo group, and save/reopen projection is covered. The generated Manifest
contains 191 operations, including sixty-seven DOCX operations (sixty-four Agent-visible and three
internal).

R2-186 adds aggregate `docx.object.set_offset_position` for all textbox-backed drawings while
leaving images in the closed image family. Explicit finite wrap and signed 32-bit EMU offsets flow
through one kernel shared by shape draw, move handles, and Registry. Numeric-offset OOXML now
preserves left/right square/tight/through direction through `wrapText` without losing wrapPolygon
bytes. Native Undo and negative-offset save/reopen are covered. The generated Manifest contains
192 operations, including sixty-eight DOCX operations (sixty-five Agent-visible and three internal).

R2-187 adds aggregate `docx.object.set_style` rather than separate fill and outline tools. One
masked final-state contract uses an exact block identity, nullable uppercase six-digit colors, and a
unique field mask; lines remain stroke-only. Operation Contract adds `pattern`, and the retained
Shape Format palette and Registry share the same protected-node transaction with native Undo and
save/reopen evidence. The generated Manifest contains 193 operations, including sixty-nine DOCX
operations (sixty-six Agent-visible and three internal).

R2-188 adds aggregate `docx.object.remove` for exact shape, line, textbox/WordArt-like textbox,
chart, diagram, and block-equation identities. Images and generic protected fields fail closed.
Object-mode Backspace/Delete and Registry share one transaction kernel; the last-block invariant,
independent native Undo, and save/reopen are covered. The generated Manifest contains 194
operations, including seventy DOCX operations (sixty-seven Agent-visible and three internal).

R2-189 adds masked `docx.chart.update` over existing title/category/series cache slots. The contract
preserves chart structure and read-only null gaps, while bounded text/value final states flow through
one kernel shared by the protected chart grid and Registry. Native Undo plus original/generated
save/reopen are covered. The generated Manifest contains 195 operations, including seventy-one DOCX
operations (sixty-eight Agent-visible and three internal); chart lifecycle is closed with aggregate
object size/removal.

R2-190 adds `docx.equation.update` with one explicit `latex|tokens` mode discriminator. LaTeX mode
rebuilds an exact block identity or inline atom range; token mode preserves one retained block's
OMML leaf count and is bounded by item, per-item, and aggregate Unicode size. The nullable payload
and target coordinates fail closed when mixed. Registry, EquationModal, and protected token editing
share one close-history transaction kernel, with native Undo and save/reopen evidence for both
representations. The generated Manifest contains 196 operations, including seventy-two DOCX
operations (sixty-nine Agent-visible and three internal).

R2-191 adds aggregate `docx.textbox.set_content` with a stable top-level object identity plus exact
nested textbox index. One bounded rich paragraph/run model covers ordinary textboxes,
textbox-backed shapes/WordArt, and multi-box retained drawings; nullable height captures UI
auto-growth without forcing fixed state. Runtime aggregate budgets and flattened-structure guards
fail before mutation. Registry and nested sub-editors share one batch-capable close-history writer,
with single-step UI/Agent Undo and original OOXML save/reopen evidence. The generated Manifest
contains 197 operations, including seventy-three DOCX operations (seventy Agent-visible and three
internal), closing the retained shape/textbox/equation content row.

R2-192 adds aggregate `docx.text.set_link` with an exact range and nullable href/replacement text.
The discriminated runtime rules cover insertion/replacement, existing-text href updates, and
removal without separate tools. LinkInsertModal and Registry share one close-history writer; native
Undo plus external-relationship save/reopen are covered. The broad whole-block `link` field is
deleted from public `docx.text.set_style`, making the exact route authoritative. The generated
Manifest contains 198 operations, including seventy-four DOCX operations (seventy-one Agent-visible
and three internal).

R2-193 adds final-state `docx.bookmark.set` with a stable top-level block identity, a unique bounded
name, and explicit enabled state. BookmarkModal and Registry share one close-history writer;
idempotent replay, duplicate rejection, native Undo, and OOXML save/reopen are covered. The
generated Manifest contains 199 operations, including seventy-five DOCX operations (seventy-two
Agent-visible and three internal).

R2-194 adds exact-range `docx.cross_reference.insert` with an existing bounded bookmark identity and
explicit cached display text. CrossRefModal and Registry share the same close-history field action;
missing targets fail closed, and native Undo plus REF-field save/reopen are covered. The generated
Manifest contains 200 operations, including seventy-six DOCX operations (seventy-three
Agent-visible and three internal).

R2-195 adds exact-range `docx.field.insert` for the five retained generic instructions and explicit
cached display text. The mounted App/Ribbon and Registry share the same close-history field action;
native Undo and generic-field OOXML save/reopen are covered. The generated Manifest contains 201
operations, including seventy-seven DOCX operations (seventy-four Agent-visible and three internal).

R2-196 adds aggregate `docx.field.update` with 1–1024 exact ranges, bounded original instructions,
and explicit final caches under a 65,536-character aggregate cap. All targets validate before one
descending close-history transaction. App F9/context-menu and Registry share the action;
multi-field Undo and OOXML save/reopen are covered. The generated Manifest contains 202 operations,
including seventy-eight DOCX operations (seventy-five Agent-visible and three internal).

R2-197 routes the retained symbol palette through existing `docx.text.insert`. Its string is now
Unicode-bounded to 1–65,536 characters, and Registry/UI share `insertDocxText` plus one native Undo
boundary. Empty input and symbol save/reopen are covered. No new descriptor is introduced: the
Manifest remains at 202 operations with seventy-eight DOCX operations, closing Priority 9.

R2-198 adds stable-target `docx.toc.refresh` with 1–1024 bounded explicit entries. The shared action
finds the matching original/generated fldChar region, preserves a trailing page break, and replaces
it through one close-history transaction. Retained UI, Registry, Undo/redo, and save/reopen share the
same route. The generated Manifest contains 203 operations, including seventy-nine DOCX operations
(seventy-six Agent-visible and three internal).

R2-199 adds exact-range `docx.note.insert { range, kind, noteId, text }`. The format-owned action
rejects duplicate IDs and cross-block ranges, inserts the reference plus bounded body metadata in
one native-history transaction, and lets the mounted session reconcile note-part membership from
that Undo/Redo-owned atom. The retained dialog and Registry share the action; Broker projection and
save/reopen are covered. The generated Manifest contains 204 operations, including eighty DOCX
operations (seventy-seven Agent-visible and three internal).

R2-200 adds stable-identity `docx.note.update { kind, noteId, text }`. Every matching reference atom
receives the bounded final body in one native-history transaction, while the session records the
pre-migration note body once as the Undo baseline. Retained UI and Registry now share this route;
missing identities, Broker validation, Undo/Redo, and save/reopen are covered. The generated
Manifest contains 205 operations, including eighty-one DOCX operations (seventy-eight Agent-visible
and three internal).

R2-201 adds `docx.note.delete { kind, noteId }`. One stable identity removes all matching reference
atoms and renumbers remaining same-kind peers in one native transaction. The shared note state keeps
the correct pre-migration baseline and deletion intent so Undo/Redo composes with earlier insertion
or update. Retained UI, Registry, Broker validation, and save/reopen are covered. The generated
Manifest contains 206 operations, including eighty-two DOCX operations (seventy-nine Agent-visible
and three internal).

R2-202 adds `docx.source.upsert { source }`. A stable tag selects create/update, bounded modeled
fields form one final Sources array, and the format-owned action writes that array to the first
block's native-history override. Retained UI and Registry commit the same snapshot to customXml;
Broker validation and save/reopen are covered. The generated Manifest contains 207 operations,
including eighty-three DOCX operations (eighty Agent-visible and three internal).

R2-203 adds `docx.citation.insert { range, sourceTag, displayText }`. The exact inline range and
stable existing source tag are validated before one bounded plain-text replacement transaction,
matching retained renderer semantics. UI, Registry, Broker, Undo, and save/reopen share the route.
The generated Manifest contains 208 operations, including eighty-four DOCX operations (eighty-one
Agent-visible and three internal).

R2-204 adds `docx.bibliography.insert { afterBlockIndex, heading, entries }`. The revision-scoped
boundary and bounded explicit text snapshot replace selection-dependent UI insertion; all entry
tags must resolve in the current Sources state. UI, Registry, Broker, one-step native Undo, and
save/reopen share the route. The generated Manifest contains 209 operations, including eighty-five
DOCX operations (eighty-two Agent-visible and three internal).

R2-205 adds `docx.caption.insert { afterBlockIndex, label, number, text }`. The caller supplies the
bounded final SEQ number and text, and one stable top-level boundary replaces selection-dependent
insertion. UI, Registry, Broker, one-step native Undo, and dirty-field save/reopen share the route.
The generated Manifest contains 210 operations, including eighty-six DOCX operations (eighty-three
Agent-visible and three internal).

R2-206 adds `docx.index.mark { range, term }`. One exact same-block inline range anchors a bounded
hidden XE marker, and terms that cannot survive the writer unchanged fail closed. UI, Registry,
Broker, native Undo, and save/reopen share the route. The generated Manifest contains 211
operations, including eighty-seven DOCX operations (eighty-four Agent-visible and three internal).

R2-207 adds `docx.index.insert { afterBlockIndex, label, terms }`. One bounded explicit term snapshot
is normalized/deduplicated/sorted and written as a complete dirty INDEX cache after a stable block
boundary. UI, Registry, Broker, one-step native Undo, and save/reopen share the route. The retained
source UI exposes no deletion mutation, closing the References and Index family. The generated
Manifest contains 212 operations, including eighty-eight DOCX operations (eighty-five Agent-visible
and three internal).

R2-208 adds `docx.comment.add { range, comment }`. One exact non-empty range and explicit bounded
stable metadata drive a single transaction containing both anchor marks and an Undo-owned final
comment snapshot. UI, Registry, Broker, Undo/Redo, and `comments.xml` save/reopen share the route.
The generated Manifest contains 213 operations, including eighty-nine DOCX operations (eighty-six
Agent-visible and three internal).

R2-209 adds `docx.comment.reply { parentId, comment }`. One stable top-level parent supplies all
shared anchors, while explicit reply metadata and the final comments snapshot commit atomically.
UI, Registry, Broker, Undo, and extended-comment save/reopen share the route. The generated Manifest
contains 214 operations, including ninety DOCX operations (eighty-seven Agent-visible and three internal).

R2-210 adds `docx.comment.set_resolved { id, resolved }`. One stable top-level thread and explicit
final boolean atomically update parent/reply metadata in the Undo-owned snapshot. UI, Registry,
Broker, native Undo, and extended-comment save/reopen share the route. The generated Manifest
contains 215 operations, including ninety-one DOCX operations (eighty-eight Agent-visible and three internal).

R2-211 adds `docx.comment.delete { id }`. Top-level deletion cascades direct replies; reply deletion
remains local. Anchor cleanup and the final comments snapshot commit in one native-history
transaction shared by retained UI and Registry, with Broker rejection and save/reopen evidence.
The generated Manifest contains 216 operations, including ninety-two DOCX operations (eighty-nine
Agent-visible and three internal), closing the retained comment lifecycle.

R2-212 adds `docx.revision.set_tracking { enabled }`. Retained Ribbon and Registry share one
explicit mounted-recorder state writer with deterministic no-op reporting. The recorder toggle is
declared non-undoable and does not create a recovery checkpoint; a subsequent edit is still one
native Undo unit and its tracked insertion saves/reopens as `w:ins`. The generated Manifest
contains 217 operations, including ninety-three DOCX operations (ninety Agent-visible and three internal).

R2-213 adds `docx.revision.apply_decision { decision, scope }`. Finite accept/reject and
current/all enums cover all four retained Review actions through one contract and one mounted
native-history transaction. Empty-state failure, selection-resolved current targeting, one-step
Undo, Broker rejection, and final save/reopen are covered. The generated Manifest contains 218
operations, including ninety-four DOCX operations (ninety-one Agent-visible and three internal),
closing the retained revisions family.

R2-214 adds `docx.document.set_protection { enabled, password }`. Bounded input-only credentials
enable passwordless/password-protected read-only state or verify its removal; repeated final state
is a no-op and incorrect passwords fail closed. Retained Review UI and Registry share the resolver
and dirty save journal. The operation is explicitly non-undoable, with recovery and settings-part
enable/remove save/reopen evidence. The generated Manifest contains 219 operations, including
ninety-five DOCX operations (ninety-two Agent-visible and three internal).

R2-215 adds `docx.ink.apply { action, annotation?, ids? }`. One finite action enum covers bounded
vector add, stable-ID delete, and clear; runtime cross-validates action-specific fields and rejects
duplicate/missing identities. Opaque reopened strokes remain deletable, while Agent image/base64
injection is absent. Retained overlay/Ribbon and Registry share the non-undoable ink journal with
dirty recovery and ink-part save/reopen evidence. The generated Manifest contains 220 operations,
including ninety-six DOCX operations (ninety-three Agent-visible and three internal), closing Ink.

R2-216 adds public `docx.document.compare { path }` and internal
`docx.document.compare_staged { blobId, name, size, data }`. The public contract exposes only a
bounded local path to the Broker; the generated Manifest hides the internal bytes contract. Broker
staging feeds the same renderer-owned DOCX parser and deterministic paragraph-diff action used by
the retained Compare UI, then commits one comparison-panel state result without dirty recovery or
document-package mutation. The generated Manifest contains 222 operations, including ninety-eight
DOCX operations (ninety-four Agent-visible and four internal), closing Compare.

R2-217 adds `docx.document.set_design { fields, pageColor?, watermark?, themeFonts?,
themeColors? }`. One unique 1–4-item mask must exactly match the supplied bounded final-state
values, so a complete theme remains one atomic Registry operation rather than UI preset strings or
four public tools. Retained Design UI and Registry share the resolver and commit seam. The
non-undoable document journal stays dirty and recovery-backed, with combined save/reopen evidence.
The generated Manifest contains 223 operations, including ninety-nine DOCX operations
(ninety-five Agent-visible and four internal), closing Document design.

R2-218 adds `docx.cover_page.insert { preset, title, subtitle, author, date, year }`. The twelve
finite presets carry styling only; every visible locale/date-sensitive value is explicit and
bounded for deterministic replay. Retained Ribbon defaults use the same action. Start insertion,
one native Undo transaction, Broker enum rejection, and custom-content save/reopen are covered.
The generated Manifest contains 224 operations, including 100 DOCX operations (ninety-six
Agent-visible and four internal), closing Cover pages.

R2-219 adds `docx.paragraph.set_drop_cap { blockIndex, mode, lines }`. One explicit top-level
paragraph-like identity and finite none/drop/margin final state replace selection-only UI behavior;
runtime requires null lines for none and 2–10 for visible modes. Retained Ribbon and Registry share
one native transaction with no-op, Undo, Broker rejection, and save/reopen evidence. The generated
Manifest contains 225 operations, including 101 DOCX operations (ninety-seven Agent-visible and
four internal), closing Drop caps.

R2-220 adds `docx.wordart.insert { afterBlockIndex, preset, text, widthEmu, heightEmu, drawingId }`.
Finite style, bounded content/geometry, and explicit unique OOXML identity replace private random
Ribbon insertion. Retained UI allocates the lowest available identity and shares the action.
Broker/runtime rejection, one native Undo transaction, and text/readable-color/geometry
save/reopen are covered. The generated Manifest contains 226 operations, including 102 DOCX
operations (ninety-eight Agent-visible and four internal), closing the R2-132 DOCX retained-command
inventory without asserting shared release readiness.

R2-221 adds `markdown.history.undo {}` and `markdown.history.redo {}`. Retained Ribbon and Registry
share one native TipTap history action, while empty-stack attempts fail explicitly. Renderer tests
cover the same stack and MCP tests cover generated discovery plus live-session queue and
acknowledgement. The generated Manifest contains 228 operations, including six Markdown operations
(five Agent-visible and one internal). Block-type parity is the next Markdown slice.

R2-222 adds `markdown.block.set_type { textBlockIndex, type }`. The revision-scoped flattened
text-block address covers paragraph, H1–H6, quote, and code-block final states without relying on
the Agent's ambient selection. Current context exposes the address; retained Ribbon, Slash Menu,
and Registry share the format action. Native Undo, save/reopen, bounds, enum rejection, and real
session acknowledgement are covered. The generated Manifest contains 229 operations, including
seven Markdown operations (six Agent-visible and one internal). Inline marks are next.

R2-223 adds `markdown.text.set_marks { from, to, marks }` as one aggregate final-state operation
instead of five toggles. Its required mask covers bold, italic, strike, inline code, and nullable
link; incompatible code combinations fail explicitly. Retained Ribbon and Registry share the
action. Set/clear, native Undo, save/reopen, runtime and Broker rejection, generated discovery,
and live-session acknowledgement are covered. The generated Manifest contains 230 operations,
including eight Markdown operations (seven Agent-visible and one internal). Lists are next.

R2-224 adds `markdown.list.set_type { textBlockIndex, type }` with explicit
none/bullet/ordered/task final states. It reuses the block address, retains siblings on item exit,
and converges Ribbon, Slash Menu, and Registry on one TipTap action. Undo, task-list save/reopen,
bounds, enum rejection, discovery, and live-session acknowledgement are covered. The generated
Manifest contains 231 operations, including nine Markdown operations (eight Agent-visible and one
internal). Structure insertion is next.

R2-225 adds `markdown.table.insert { position, rows, columns, headerRow }` and
`markdown.divider.insert { position }`. Explicit revision-scoped PM positions keep Ribbon, Slash
Menu, and Registry on shared format actions while bounded dimensions prevent unbounded table
creation. Native Undo, table save/reopen, runtime/Broker failure, discovery, and live-session
acknowledgement are covered. The generated Manifest contains 233 operations, including eleven
Markdown operations (ten Agent-visible and one internal). Staged images are next.

R2-226 adds public `markdown.image.insert` and internal `markdown.image.insert_staged`. The Broker
owns absolute-path, extension, size, and byte staging; the Renderer revalidates hydrated bytes and
embeds a data URL through the action shared with picker and paste/drop. Native Undo, Markdown
reopen, forged-byte rejection, hidden internal discovery, and real no-inline-byte transport are
covered. The generated Manifest contains 235 operations, including thirteen Markdown operations
(eleven Agent-visible and two internal). Frontmatter is next.

R2-227 adds `markdown.frontmatter.set { yaml }` over the retained raw envelope. UI and Registry
share one writer; empty input removes frontmatter while body/BOM/EOL/EOF state remains intact. The
descriptor truthfully declares non-TipTap undoability and relies on dirty/recovery/save-reopen.
The 1 MiB Broker bound, discovery, and live acknowledgement are covered. The generated Manifest
contains 236 operations, including fourteen Markdown operations (twelve Agent-visible and two
internal). Table-relative commands are next.

R2-228 adds `markdown.table.update { position, action, headerRow }`. One finite action enum covers
eight relative table commands; header state is explicit and relationally validated. TableMenu and
Registry share the path. Undo, reopen, non-table/relation failure, Broker rejection, discovery,
and live acknowledgement are covered. The generated Manifest contains 237 operations, including
fifteen Markdown operations (thirteen Agent-visible and two internal). Blocks are next.

R2-229 adds `markdown.block.update { blockIndex, action, afterBlockIndex, content }`. Four bounded
actions cover duplicate/delete/add/move without separate descriptors. Menu, BlockKeymap, plus,
drag-drop, and Registry share the transaction builder; final-block deletion preserves the document
invariant. Undo, ordering, relation failure, Broker rejection, discovery, and live acknowledgement
are covered. The generated Manifest contains 238 operations, including sixteen Markdown
operations (fourteen Agent-visible and two internal). Code language is next.

R2-230 adds `markdown.code_block.set_language { textBlockIndex, language }`. Thirty finite UI
languages, including explicit plaintext/null, replace direct NodeView attribute writes. NodeView
and Registry share one native transaction. Target failure, Undo, reopen, Broker rejection,
discovery, and live acknowledgement are covered. The generated Manifest contains 239 operations,
including seventeen Markdown operations (fifteen Agent-visible and two internal). Retained UI and
direct-edit/export classification is next.

R2-231 adds `markdown.document.save_as {}`. The retained Shift-Save gesture and Registry share the
same forced-picker persistence service, and the operation returns the selected file name or an
explicit cancellation/failure. Generated discovery and live acknowledgement are covered. The
generated Manifest contains 240 operations, including eighteen Markdown operations (sixteen
Agent-visible and two internal). DOCX export is next.

R2-232 through R2-234 add exact retained output/preference operations:
`markdown.document.export_docx`, `markdown.document.open_print_dialog`, and
`markdown.document.set_auto_save`. UI and Registry share the same services; image export is
bounded and byte-validated, popup blocking is explicit, and autosave assigns a persisted final
state. The generated Manifest reaches 243 operations, including twenty-one Markdown operations
(nineteen Agent-visible and two internal).

R2-235 adds bounded `markdown.selection.set { from, to }`, completing arbitrary text targeting
without adding selection-only history or recovery. The generated Manifest contains 244 operations,
including twenty-two Markdown operations (twenty Agent-visible and two internal).

R2-236 and R2-237 close the removed `md-asset` fidelity boundary with a session-owned chunked MCP
asset transport and a directory-authorized standalone browser reader. Both hydrate only display
state, preserve authored paths, enforce PNG/JPEG/GIF and 20 MiB bounds, and feed the shared DOCX
exporter. R2-238 makes the retained-command mapping executable: all 22 descriptors are covered and
`missing` is forbidden. Markdown format-local parity completed here; R6-01 later closed the shared
release gates.

R2-163 adds `docx.section.set_page_size { sectionIndex, widthTwips, heightTwips }` with bounded
actual axes, current-margin fit validation, and deterministic orientation derivation. Ribbon paper
presets and Agent values converge on the full section override journal. The generated Manifest
contains 167 operations, including forty-three DOCX operations (forty-two Agent-visible and one
internal).

R2-137 migrates Change Case as `docx.text.transform_case { range, mode }` with bounded sentence,
lower, upper, and title modes. The helper replaces text through mapped ProseMirror positions, so
Unicode length changes preserve the final selection and each source node's marks. The Ribbon's
private implementation is removed in favor of this shared single-transaction route. The generated
Manifest contains 143 operations, including nineteen DOCX operations (eighteen Agent-visible and
one internal).

R2-138 closes the exact character-format family with
`docx.text.set_character_style { range, styleId }`. The DOCX composition root injects a read-only
predicate over the opened document's style catalog, so arbitrary or paragraph-style IDs fail
before mutation. Application merges the existing run mark; `null` matches the retained Ribbon's
mark-removal behavior. Real style cards and built-in fallback presets now use the shared character
helpers. The generated Manifest contains 144 operations, including twenty DOCX operations
(nineteen Agent-visible and one internal).

R2-139 starts paragraph/list detail with
`docx.paragraph.set_direction { target, direction }`. The target reuses the DOCX block contract;
the executor imports the same pure direction-flip rule as the Ribbon, changing `bidi` and swapping
explicit left/right alignment to preserve logical start/end. All matches commit in one native Undo
transaction. The generated Manifest contains 145 operations, including twenty-one DOCX operations
(twenty Agent-visible and one internal).

R2-140 adds `docx.list.set_level { target, level }`, replacing relative Agent indent gestures with
an explicit absolute `0..8` final level. UI stepping and registry assignment share one normalization
rule; the target-aware executor updates only list items in one transaction. The generated Manifest
contains 146 operations, including twenty-two DOCX operations (twenty-one Agent-visible and one
internal).

R2-141 adds `docx.list.apply_preset { target, levels }` as the first document-definition-producing
DOCX tracer. Its bounded 1–9-level contract creates numbering through the same composition-root
service as the retained custom-list dialog, then binds all eligible target blocks to the returned
`numId` in one native transaction. Invalid or empty presets cannot allocate external numbering
state. The generated Manifest contains 147 operations, including twenty-three DOCX operations
(twenty-two Agent-visible and one internal).

R2-142 adds `docx.list.restart { blockIndex, start }`. It replaces the selection-relative Agent
gesture with a stable top-level block anchor and bounded explicit start value. Registry and retained
context menu call one numbering action that clones the source abstract definition, adds a
level-specific start override, and changes only later same-list items in one native transaction.
The generated Manifest contains 148 operations, including twenty-four DOCX operations
(twenty-three Agent-visible and one internal).

R2-143 adds `docx.list.continue { blockIndex, previousBlockIndex }`. It requires both the current
anchor and the earlier list source instead of guessing from selection context. Registry and UI use
one action that rebinds only current-list items at and after the anchor to the explicit previous
`numId` in one native transaction. The generated Manifest contains 149 operations, including
twenty-five DOCX operations (twenty-four Agent-visible and one internal); the list-numbering family
is closed.

R2-144 deepens `docx.paragraph.set_style` with the ParagraphDialog's exact/at-least raw line-height
state and bounds every retained spacing/indent dimension. UI and Agent continue to mutate the same
mounted paragraph attributes and native history; no duplicate operation is introduced. The
generated Manifest remains at 149 operations, including twenty-five DOCX operations (twenty-four
Agent-visible and one internal).

R2-145 adds bounded `tabStops` final state to the existing paragraph-style operation. Registry and
Ruler share the public-array-to-renderer-attribute adapter; strict ordering, bounded positions, and
finite alignment/leader enums prevent ambiguous ruler gesture replay. The generated Manifest
remains at 149 operations and the retained paragraph-format family is closed.

R2-146 adds `docx.table.insert { afterBlockIndex, rows, columns }` as the table-family tracer.
Stable top-level placement, bounded axes, and a 4096-cell aggregate guard precede creation. Agent
and UI share the format-owned empty-table model and native schema; one insertion is one native Undo
entry. The generated Manifest contains 150 operations, including twenty-six DOCX operations
(twenty-five Agent-visible and one internal).

R2-147 adds `docx.table.delete { tableBlockIndex }` with exact top-level table identity and fail-closed
node-type validation. The native transaction restores with one Undo and preserves a valid document
when the last table is removed. The generated Manifest contains 151 operations, including
twenty-seven DOCX operations (twenty-six Agent-visible and one internal).

R2-148 adds `docx.table.insert_rows { tableBlockIndex, rowIndex, count }` over an explicit row
boundary. It recomputes the native TableMap after every staged addition to preserve rowspan
semantics, but dispatches the accumulated transaction once. The generated Manifest contains 152
operations, including twenty-eight DOCX operations (twenty-seven Agent-visible and one internal).

R2-149 adds `docx.table.delete_rows { tableBlockIndex, rowIndex, count }` with an explicit row
interval. The renderer rejects out-of-range intervals and whole-table deletion before mutation;
the latter remains the responsibility of `docx.table.delete`. Each removal recomputes the native
TableMap for rowspan repair while dispatching one accumulated transaction. The generated Manifest
contains 153 operations, including twenty-nine DOCX operations (twenty-eight Agent-visible and one
internal).

R2-150 adds `docx.table.insert_columns { tableBlockIndex, columnIndex, count }` over an explicit
logical-column boundary. Renderer validation caps the resulting width at 63 and the logical-cell
product at 4096. Recomputing the native TableMap between staged insertions preserves colspan
semantics while one dispatch preserves one-step Undo. The generated Manifest contains 154
operations, including thirty DOCX operations (twenty-nine Agent-visible and one internal).

R2-151 adds `docx.table.delete_columns { tableBlockIndex, columnIndex, count }` with an explicit
logical-column interval. Out-of-range and whole-table outcomes fail before mutation. Each removal
recomputes the native TableMap for colspan repair and the accumulated transaction produces one Undo
entry. The generated Manifest contains 155 operations, including thirty-one DOCX operations
(thirty Agent-visible and one internal).

R2-152 adds `docx.table.merge_cells` over one bounded half-open logical-cell rectangle. Renderer
validation requires an exact `TableMap` rectangle that does not cross existing spans, then maps it
to the native `CellSelection` and retained merge command. The shared row schema accepts an empty
physical row when it is fully covered by rowspan, preventing the prior automatic phantom-cell
repair in both UI and Registry paths. The generated Manifest contains 156 operations, including
thirty-two DOCX operations (thirty-one Agent-visible and one internal).

R2-153 adds `docx.table.split_cell { tableBlockIndex, rowIndex, columnIndex }`. The native TableMap
resolves the logical coordinate to its covering cell, rejects unmerged cells, and supplies the exact
single-cell CellSelection to the retained split command. The original span product is returned and
one native Undo restores it. The generated Manifest contains 157 operations, including
thirty-three DOCX operations (thirty-two Agent-visible and one internal).

R2-154 adds `docx.table.set_cell_format` with an exact logical rectangle, bounded fill/alignment
final state, and a unique field mask. Explicit Agent targets and retained Ribbon selections converge
after target resolution on the same physical-cell attrs kernel and native Undo route. The generated
Manifest contains 158 operations, including thirty-four DOCX operations (thirty-three Agent-visible
and one internal).

R2-155 adds `docx.table.set_cell_borders` with exact rectangle identity, four finite edge policies,
and a bounded nullable border final state. Agent targets and retained Ribbon selections share the
same span-aware edge/write kernel and native Undo route. The generated Manifest contains 159
operations, including thirty-five DOCX operations (thirty-four Agent-visible and one internal).

R2-156 adds `docx.table.set_style` with stable table identity and current-document style
validation. `null` clears the final state; Agent and Ribbon target adapters share the same native
attribute kernel. The generated Manifest contains 160 operations, including thirty-six DOCX
operations (thirty-five Agent-visible and one internal).

R2-157 adds `docx.table.set_row_height` with a bounded physical-row interval and nullable OOXML
twip final state. Agent explicit targets and retained Ribbon selections share the same row-attribute
kernel and native Undo route. The generated Manifest contains 161 operations, including
thirty-seven DOCX operations (thirty-six Agent-visible and one internal).

R2-158 adds `docx.table.set_column_widths` as a bounded complete-grid final state. Agent widths and
Ribbon-derived fitted widths share the same spanning-cell/table-grid write kernel and native Undo
route. The generated Manifest contains 162 operations, including thirty-eight DOCX operations
(thirty-seven Agent-visible and one internal). The retained table family is closed.

R2-159 adds `docx.document.insert_page_break { afterBlockIndex }` at a stable top-level block
boundary. It inserts the retained renderer's native empty `pageBreakBefore` paragraph, shares the
same writer with the Ribbon selection command, rejects invalid boundaries before mutation, and is
one native Undo unit. The generated Manifest contains 163 operations, including thirty-nine DOCX
operations (thirty-eight Agent-visible and one internal). This opens the page/section family.

R2-160 adds `docx.section.insert_break { afterBlockIndex, startType }`. The finite start type is
stored on the native inserted break node for Undo, then projected during save onto the following
section terminator or trailing `sectPr`. Consecutive new breaks, exact-boundary rejection, Ribbon
equivalence, and save/reopen are covered. The generated Manifest contains 164 operations,
including forty DOCX operations (thirty-nine Agent-visible and one internal).

R2-161 adds `docx.section.set_orientation { sectionIndex, orientation }` and establishes the
Undo-owned section settings override journal. The journal stores a complete final `SectionSettings`
snapshot on the indexed section's last visible block; save projects it to the correct non-final or
trailing `sectPr`, while Ribbon state follows Undo/Redo. The generated Manifest contains 165
operations, including forty-one DOCX operations (forty Agent-visible and one internal).

R2-162 adds `docx.section.set_margins { sectionIndex, margins }` with four bounded twip values and
a positive-body geometry check. It reuses the Undo-owned full settings journal now used by every
retained margin preset and custom dialog, and has save/reopen evidence. The generated Manifest
contains 166 operations, including forty-two DOCX operations (forty-one Agent-visible and one
internal).

Deliverables per format:

- descriptors and handlers for the current open, edit, selection, history, and save operations;
- internal-visibility descriptors for staged local-file loading, filtered out of Agent capability discovery;
- exact schemas replacing the current partial `OperationSchema` representation;
- compatibility aliases for current operation names;
- renderer adapters using registry dispatch instead of composition-root operation branches;
- equivalence tests showing unchanged results, revisions, recovery, and save/reopen behavior.

Cross-format deliverables:

- generated manifest compared with the existing `capabilities.ts` output;
- generated installed-Skill reference;
- MCP server reads the generated manifest for discovery and pre-enqueue validation;
- app-only transport tools remain unchanged.
- DOCX polling/acknowledgement converges on the shared host-bridge Module, with font loading supplied as an extension instead of a second state machine;
- PPTX registers one renderer controller/action Interface shared by UI selection and Agent operations;
- PDF publishes the mounted App's actual selection through the same controller used by its command bridge;
- XLSX compatibility dispatch returns an explicit success/failure outcome so unknown or inactive Ribbon strings cannot receive a positive acknowledgement.

Exit gate:

- all current MCP integration and visual tests pass through registry handlers;
- every advertised descriptor has one handler and focused test;
- the central catalog is read-only compatibility data and fails CI if it diverges;
- no new capability family has been added during the migration.

### R3 — Bounded discovery and transaction envelope

Goal: replace shallow discovery and arbitrary operation acceptance.

Status: complete in R6-06 through R6-08. Discovery is bounded, the installed Skill uses exact detail
schemas and the transaction envelope, the live Broker journals idempotent completed/in-flight
replay, and the public legacy input Adapter is removed.

Deliverables:

- `office_get_capabilities` summary/detail views with family filtering, exact lookup, and bounded pagination (complete in R6-06);
- optional `sessionId` availability projection (complete in R6-06);
- `office_execute` transaction envelope with `requestId`, `baseRevision`, and fully qualified ids (complete in R6-07);
- idempotent replay tracking for completed/in-flight requests within the live broker lifetime (complete in R6-07);
- broker-side manifest/schema/context validation before enqueue (complete in R6-07);
- typed transaction errors (complete in R6-07);
- removal of compatibility translation from the old single-operation input, its public aliases,
  and its legacy response branch (complete in R6-08).

Exit gate:

- a schema-invalid operation never reaches the renderer;
- exact replay returns the original result and payload-changing request reuse fails;
- one acknowledged transaction advances revision once;
- multi-operation input is rejected unless the format registry proves one native atomic undo unit;
- the installed Skill uses summary/detail discovery and the transaction form.

### R4 — Remove transitional integration paths

Goal: make the registry the only extension path before broad parity work resumes.

Status: complete through R6-09. Public execution and renderer registry ingress are canonical-only;
only five internal `open_local_file` aliases remain for dedicated staged-file transport.

Deliverables:

- delete hand-written format entries from `apps/mcp-server/src/capabilities.ts`, retaining at most a generated-manifest loader;
- delete direct format operation branches from renderer composition roots after their registry adapters land;
- keep the R2-18 removal of DOCX `batch_update` gated while adding declared atomic grouping;
- replace XLSX `ribbon_command` with semantic typed operations;
- generate protocol and Skill operation references;
- document and test compatibility-alias removal gates.

Exit gate:

- deletion test: removing the registry would force discovery, validation, dispatch, documentation, and test mapping complexity back into multiple callers;
- `rg` finds no open-ended `ribbon_command` execution route;
- no public `commands: array` schema lacks item discrimination;
- no renderer composition root owns the Agent operation inventory;
- new operation pull requests cannot pass CI by editing only MCP server metadata.

### R5 — Format parity slices

Goal: complete each format's retained command inventory through vertical operation-family slices.

Each slice must include, in one change:

1. baseline inventory entries and source routes;
2. exact operation descriptor/schema;
3. context/target projection;
4. renderer handler using the native command seam;
5. user and Agent post-state equivalence;
6. undo/redo evidence when applicable;
7. recovery and save/reopen evidence when persisted state changes;
8. generated manifest, Skill reference, parity report, ledger, and protocol updates.

Format order follows the active renderer roadmap:

#### R5.1 DOCX

- preserve the completed migration of all former structured commands out of `batch_update`;
- inventory and implement text/paragraph formatting, structure, tables, images, references, review, history, layout, and save families;
- keep ProseMirror/Tiptap selection and the existing surgical-save path authoritative.

#### R5.2 XLSX

- replace Ribbon strings with range-format, clipboard, formula, structure, sheet, data, table, pivot, visual, review, view, layout, and save families;
- reuse the retained Univer commands and format-owned browser package gateways;
- reject any operation whose mutation journal cannot yet survive save/reopen.

#### R5.3 PPTX

- add slide, text, shape, picture, table, chart, design, transition, animation, notes, review, master, view, history, and save families;
- reuse the original App/Ribbon/Konva command and browser-presentation seams.

#### R5.4 PDF

- Complete: annotation, drawing, text, image, form, signature, page, properties, history, and save
  families are Registry-owned; navigation state remains correctly view-only.
- Only mutations that the retained App/browser adapter can serialize and reopen are advertised.
  Local page insertion uses an explicit non-undoable persist/reload contract.

#### R5.5 Markdown

- add history, block, mark, list, structure, frontmatter, table, code-block, image, export-setting, and save families;
- route through TipTap commands and the retained Markdown envelope/round-trip implementation.

Exit gate for each format:

- every retained state-changing baseline entry has an executable mapping and evidence;
- no unexplained UI-only mutation remains;
- `ready` changes only after the complete ADR 0003 gate passes.

### R6 — Cross-format release gate and cleanup

Deliverables:

- remove compatibility aliases no longer used by the Skill, tests, or packaged plugin (public
  `office_execute.operation + arguments` Adapter complete in R6-08; all thirty-six public-era
  format aliases complete in R6-09; five internal `open_local_file` transport aliases retained);
- delete transitional operation documentation;
- run generated parity, typecheck, unit, integration, visual, round-trip, performance, package, license, and prohibited-dependency checks;
- confirm all five format manifests and packaged Skill references match;
- update `live-session.md` from transitional to implemented registry contract only after the new envelope is shipped.
- retain R6-03 through R6-09 strict app-only performance traces, bounded discovery,
  revision-guarded replay, transaction-only public execution, and the self-contained XLSX deferred
  module graph, canonical-only public registry execution, and fixed reviewed budgets without adding
  public operation descriptors or removing renderer capability.

Exit gate:

- the generated parity report has zero unexplained retained state-changing gaps;
- the MCP tool list remains bounded and data tools remain UI-free;
- every operation reaches the same mounted renderer state and native undo route as the corresponding UI command;
- the broker still cannot edit when the renderer is offline;
- no independent Runtime, hidden editor, or second document authority exists.

## Pull-request slicing

Implementation should use tracer-bullet vertical slices rather than one repository-wide rewrite.

Recommended sequence:

1. common contract plus generator fixtures;
2. one simple Markdown text operation end to end;
3. migrate the remaining currently advertised operations format by format;
4. complete the switch of MCP discovery/validation to the generated manifest for all formats;
5. land the transaction envelope and Skill migration;
6. remove transitional paths;
7. resume broad format capability families.

Every pull request must leave the root checks green and must not create a second execution path that lacks a named removal gate.

## Test strategy

### Contract tests

- descriptor schema and version compatibility;
- deterministic manifest generation;
- duplicate id/alias rejection;
- exact per-operation JSON Schema validation;
- exact per-operation result-schema validation;
- transaction and error serialization;
- request replay and revision conflict behavior.

### Registry tests

- descriptor/handler one-to-one coverage;
- family and format prefix consistency;
- declared context matches handler requirements;
- compatibility alias resolves to exactly one operation;
- non-atomic combinations are rejected.

### Format integration tests

- same fixture, same selection, UI route and Agent route produce equivalent renderer state;
- one Agent transaction is one undo unit;
- failure does not advance revision;
- dirty/recovery behavior remains correct;
- save/reopen preserves the mutation and untouched format data.

### Generated-artifact tests

- MCP manifest is current;
- Skill operation reference is current;
- baseline mapping cites an existing descriptor and evidence file;
- prohibited entries never appear in the executable projection;
- plugin packaging contains the same manifest version tested in source.

## Risks and controls

| Risk                                                 | Control                                                                                                    |
| ---------------------------------------------------- | ---------------------------------------------------------------------------------------------------------- |
| A new universal Office DSL absorbs format semantics. | Common contract tests reject format-specific descriptor fields; schemas remain in format applications.     |
| Registry and handler drift.                          | One-to-one registry checks and generated manifest checks.                                                  |
| UI and Agent use different selection state.          | Format controller tests require the same selection projection and native action Interface for both routes. |
| Discovery payload grows with hundreds of operations. | Summary/detail queries, family filters, exact schema lookup, pagination.                                   |
| Compatibility layer becomes permanent.               | Every alias has a replacement and removal gate; R4 blocks broad parity work.                               |
| Multi-operation transactions partially apply.        | Default to one operation; allow grouping only with native atomic/undo evidence.                            |
| Server becomes coupled to browser code.              | Build-time serializable manifests; no renderer imports in MCP server.                                      |
| Capability is advertised before save support exists. | Advertisement gate requires save/reopen evidence for persisted mutations.                                  |
| Refactor changes renderer behavior.                  | R2 is behavior-preserving and retains current integration/visual fixtures.                                 |

## Definition of done

The route is complete only when all of the following are true:

- five format-owned executable operation registries are the only Agent-operation extension points;
- MCP discovery, validation, Skill references, and parity reports are generated or checked from those registries;
- the baseline capability inventories are machine-checkable and fully mapped;
- central hand-written catalogs, open string dispatch, and composition-root operation branches are removed;
- every retained state-changing command has UI/Agent equivalence, revision, undo, and required save/reopen evidence;
- public MCP tools remain bounded and do not remount the editor;
- mounted renderers remain the sole document and undo authorities.
