# ADR 0004: Replace transitional command catalogs with format-owned operation registries

- Status: Accepted
- Date: 2026-08-27
- Upstream baseline: `genspark-ai/genoffice@dc4d7e5927864498913b7ba42d0da06cc7cf628e`
- Extends: ADR 0001 and ADR 0003

## Context

ADR 0003 requires a typed MCP route for every retained state-changing community renderer command, but the current implementation is still a narrow migration tracer:

- `apps/mcp-server/src/capabilities.ts` centrally duplicates format operation names and partial schemas;
- `office_execute` accepts an arbitrary operation string and an untyped argument record before the mounted renderer performs format validation;
- format dispatch is embedded in large `App.tsx` or host-adapter condition chains;
- DOCX hides multiple commands behind a shallow `batch_update.commands: array` schema;
- XLSX exposes a transitional `ribbon_command.command: string` route that is neither fully discoverable nor an exact typed capability;
- capability inventories, MCP metadata, renderer handlers, Skill guidance, and parity tests are maintained separately and can drift.

Continuing to append operations to those locations would keep the public MCP tool count stable but move the growth into shallow Modules with poor Locality. It would also make it increasingly difficult to prove that advertised operations are implemented, that retained UI commands are mapped, and that both routes converge on the same renderer state and undo history.

The separate `office-runtime` repository demonstrates useful architectural ideas: a small stable MCP Interface, typed format namespaces, capability discovery, revision-guarded operation envelopes, and generated Skill references. TandemFolio does not adopt that repository's independent Runtime authority, daemon, workers, persistence, HTTP/WebSocket transport, or UI-closed editing model because those choices contradict ADR 0001.

## Decision

### Two registries with different authority

TandemFolio will maintain two related but distinct records.

1. The **baseline capability inventory** records the pinned renderer's retained controls and commands, prohibited product behavior, UI routes, source provenance, and parity evidence. It defines required migration scope but cannot advertise an operation as executable.
2. Each format owns an **executable operation registry** containing only operations that are implemented and tested against that mounted renderer. It is the source of truth for Agent-visible operation identity and behavior.

Every baseline state-changing command must eventually map to one or more executable operations, or to an explicit prohibited classification. Multiple visible controls may map to one semantic operation when the operation has an exact discriminated schema. A generic string dispatcher is not such a mapping.

### Format ownership and the common contract

Format-specific descriptors, schemas, target semantics, handlers, and tests remain inside:

- `apps/docs`;
- `apps/sheets`;
- `apps/slides`;
- `apps/pdf`;
- `apps/markdown`.

A small shared operation-contract Module may define only the cross-format envelope, descriptor fields, standard errors, revision/request identity rules, validation helpers, and generated-manifest format. It may not define a universal paragraph/cell/slide/page DSL or import renderer implementations.

Each format registry has two internal parts behind one Interface:

- a serializable catalog containing the stable id, family, summary, visibility, exact input/output schemas, effects, context requirements, risk, undoability, atomicity, and compatibility aliases;
- an executable Adapter that validates the current renderer context and invokes the renderer's existing command/state/undo route.

The registry must verify a one-to-one relationship between executable descriptors and handlers. The MCP server consumes a generated serializable manifest; it does not import browser renderer code.

### Stable MCP projection

The existing public tool surface remains stable. Format breadth continues to enter through `office_get_capabilities`, `office_get_context`, and `office_execute`, not through one new MCP tool per Ribbon command.

`office_get_capabilities` will evolve from returning one complete hand-written object to a generated projection with bounded discovery:

- summary discovery by format and operation family;
- exact descriptor/schema lookup by operation id;
- active-session availability when a `sessionId` is supplied;
- explicit reasons when connection, selection, document state, or format support makes an operation unavailable.

Release readiness remains separate from operation availability. `ready` may become true only after ADR 0003's complete format gate passes.

### Operation transactions

`office_execute` will evolve to accept a revision-guarded operation transaction with:

- `sessionId`;
- `baseRevision`;
- an idempotent `requestId`;
- one or more fully qualified, typed operations.

One acknowledged transaction advances the revision exactly once. Multiple operations are allowed only when the format registry declares the combination atomic and the mounted renderer can apply it as one native undo unit. Otherwise validation rejects the transaction before dispatch. During migration, the existing single-operation input remains a compatibility Adapter into a one-operation transaction.

The live-session broker validates the generated descriptor before enqueueing. The mounted renderer validates again at the format seam, executes through its native command route, checkpoints recovery when required, and acknowledges the deterministic result. A timeout never proves that a command was not applied; `requestId` and refreshed context govern retry behavior.

### Generated projections and gates

The format registries will generate or validate:

- the MCP capability manifest;
- the installed Skill operation reference;
- per-format executable-operation documentation;
- the mapping between baseline commands and executable operations;
- drift checks for descriptors, handlers, schemas, aliases, and tests.

A registry entry cannot be advertised until its handler, argument validation, context grounding, acknowledgement behavior, user/Agent equivalence test, and required save/reopen evidence pass.

### Replacement, not parallel accumulation

The migration must retire the transitional paths it replaces:

- the hand-written central format catalogs in `apps/mcp-server/src/capabilities.ts`;
- arbitrary operation acceptance in the broker;
- format-level `if`/`switch` operation dispatch embedded in renderer composition roots;
- the open-ended XLSX `ribbon_command` route;
- the shallow public DOCX `batch_update` schema;
- manually duplicated operation tables in the Skill and protocol documentation.

Compatibility aliases may exist only while their callers migrate. Each alias must name its replacement and removal gate. New capability work must enter through a format-owned registry rather than extend a transitional path.

## Consequences

- MCP tool count remains nearly constant while the typed operation catalog grows to cover the pinned renderer baseline.
- Format knowledge gains Locality: one format registry becomes the Interface used by discovery, validation, dispatch, documentation, and tests.
- The MCP server and host bridge remain thin Adapters around the live-session broker instead of becoming a format-neutral document Runtime.
- Capability discovery payloads remain bounded because exact schemas are loaded on demand rather than returned as one ever-growing format union.
- The first implementation work is structural migration of current behavior. Adding broad new capability families before that migration completes is intentionally blocked.
- ADR 0001's mounted-renderer authority, ADR 0002's opaque recovery model, and ADR 0003's full renderer/parity requirement remain unchanged.

## Rejected alternatives

- One MCP tool per renderer command: creates tool explosion and exposes UI placement rather than semantic operations.
- One open-ended `ribbon_command` or catch-all string command: compact but not typed, discoverable, safely targetable, or independently testable.
- One central cross-format registry with renderer semantics: loses format Locality and tends toward a false universal Office DSL.
- Importing renderer handlers into the MCP server: couples Node transport code to browser implementations and risks creating a second execution path.
- Adopting the independent Office Runtime authority model: enables UI-closed editing only by contradicting the accepted live-bound product architecture.
