# GenOffice community upstream synchronization

TandemFolio tracks the public GenOffice community repository through a read-only Git remote named `upstream`. The pinned source baseline remains `dc4d7e5927864498913b7ba42d0da06cc7cf628e`; fetching a newer candidate does not change the product baseline or admit any code into the build.

## Why a remote instead of a submodule

TandemFolio is already a source-preserving fork of the community tree. A second checkout or submodule would duplicate the source, invite accidental imports, and could place the prohibited enterprise `ee/` tree in the workspace.

The TandemFolio workflow therefore uses:

- `upstream.config.json` as the shared remote, branch, baseline, and exclusion policy;
- a local Git remote for commit and tree metadata;
- an explicitly disabled push URL;
- partial fetches with `blob:none`, with no `upstream/genoffice` checkout;
- explicit, path-scoped review and selective porting instead of upstream merges.

## Initialize and refresh

```bash
npm run upstream:setup
npm run upstream:fetch
```

`upstream:setup` is idempotent. It creates `upstream` only when missing and refuses to overwrite a remote with a different URL. `upstream:fetch` refreshes `upstream/main`, prints the number of commits since the pinned baseline, and groups changed community paths by product area. It excludes `ee/**` from inspection.

Fetching is observation only. Do not merge, rebase, pull, or copy the upstream tree wholesale into the TandemFolio branch.

## Review a candidate

```bash
npm run upstream:status
npm run upstream:diff
npm run upstream:diff -- apps/docs packages/docx-engine
```

With no paths, `upstream:diff` prints names and statuses for community changes while excluding `ee/**`. With explicit paths, it prints the patch for those paths and refuses any `ee` path.

Classify candidate changes before porting them:

1. retained DOCX renderer or engine behavior;
2. shared renderer infrastructure used by the current product graph;
3. tests, fixtures, security, dependency, or build changes;
4. future XLSX, PPTX, or PDF migration inputs;
5. excluded AI/model/account, Electron/main/preload/IPC, shell, branding, or telemetry areas;
6. prohibited `ee/` content, which must not be inspected or copied.

Only the first four categories are candidates for selective review. A changed upstream file is never sufficient reason by itself to admit it into the product build.

## Advance the pinned baseline

Move the baseline only in a dedicated reviewed change after selected changes have been ported and verified. Update all of the following together:

- `upstream.config.json`;
- `AGENTS.md` and the baseline metadata in the accepted ADR;
- `docs/migration/provenance.md`;
- `docs/migration/ledger.md` and `docs/migration/roadmap.md`;
- `NOTICE` or file-level attribution when the retained source set changes.

Then run:

```bash
npm run upstream:check
npm run check
npm run smoke:mcp
```

Record each selectively retained, modified, newly admitted, or removed upstream area in the migration ledger and provenance document. Preserve the original Apache-2.0 attribution.

## Safety rules

- `upstream` is read-only source history, not the product's merge parent.
- Never create or use `upstream/genoffice` in this repository.
- Never inspect, copy, diff, or depend on `ee/` or enterprise-only source.
- Product code must not import from remote-tracking storage or a temporary upstream checkout.
- Data and mutation behavior still follow the live-session ADR; upstream synchronization does not change product authority.
