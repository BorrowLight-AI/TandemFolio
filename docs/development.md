# Development, packaging, and operations

## Scope

This runbook covers the current five-format source graph and local Codex plugin workflow. The generated plugin includes Markdown as the fifth UI resource. ADR 0003 additionally requires complete renderer and MCP parity before first-release completion.

## Requirements

- macOS, Linux, or Windows with a modern Chromium-class browser for standalone testing.
- Node.js 22.12 or newer.
- npm 10 or newer.
- Codex CLI/app for local plugin installation.
- Python 3 only when running the plugin validation helper.

No Electron runtime, Genspark account, model key, login service, or external document service is required.

## Install

```bash
npm install --ignore-scripts
```

The explicit root workspace list is a product boundary. Do not replace it with `apps/*` or `packages/*`; doing so would reinstall inactive Electron and AI migration inputs.

## Daily commands

```bash
npm run dev:editor     # standalone DOCX renderer on http://localhost:5173
npm run dev -w @genoffice/sheets # XLSX on :5174 (slides :5175, pdf :5176)
npm run dev -w @genoffice/markdown # Markdown on :5177
npm run operation-manifest:generate # regenerate the real format-owned operation manifest
npm run operation-manifest:check # fail on fixture or product-manifest drift
npm run typecheck      # operation contract, host bridge, five renderers, and MCP server
npm test               # operation contract plus current five-format source suites
npm run build          # five editors + MCP server + self-contained plugin package
npm run smoke:mcp      # stages the plugin like an install, then checks its stdio tool/resource contract
npm run measure:assets # raw/gzip resource measurements and raw-byte budgets
npm run test:visual:host # builds, then checks the real Codex host width matrix in Chromium
npm run test:visual:host:update # deliberately replace inspected host screenshot baselines
npm run release:baseline:capture # build and capture an unapproved local R6 evidence candidate
npm run release:gate   # validate approved evidence and regenerate the ready projection
npm run release:package # fail-closed gate, build, asset measurement, MCP smoke, and licenses
npm run check          # upstream + manifest drift + typecheck + tests + full product build
npm run upstream:fetch # refreshes and summarizes upstream/main without changing the pinned baseline
```

`npm run dev` starts only the MCP server process. For interactive browser work, use `npm run dev:editor`.

The root operation-manifest commands generate and verify `apps/mcp-server/src/generated/operation-manifest.json` from `tools/operation-catalogs.ts`, which aggregates serializable format-owned TypeScript catalogs without importing renderer handlers. The check also verifies the serializable R1 fixture manifest and the 64 KiB per-descriptor ceiling. To inspect operations through MCP, page `office_get_capabilities` summary results (maximum twenty), then request `view: "detail"` for the one canonical id whose schema is needed. Do not expect `{ format }` to return the former complete schema map. Execute new mutations as `{ sessionId, baseRevision, requestId, operations: [{ id, arguments }] }`, using one canonical Agent-visible operation and a fresh caller request id for each intended mutation. Exact timeout recovery must reuse the same request id and structurally identical payload.

To verify the exact installed cache rather than the repository copy, set `TANDEMFOLIO_PLUGIN_ROOT` to the installed version directory before `npm run smoke:mcp`.

Embedded Save does not use a Codex download dialog. The renderer streams its own serialized bytes to
the local MCP server, which atomically writes a Session-bound file. New documents default to
`~/Documents/TandemFolio/<format>-<session-hash>/`; set `TANDEMFOLIO_OUTPUT_DIR` on the MCP
server to use another output root. `TANDEMFOLIO_STATE_DIR` controls recovery checkpoints and the
exact-Session document-path bindings. An explicitly opened local file remains the normal Save target.

## Track community upstream changes

TandemFolio uses a read-only `upstream` Git remote rather than checking out a duplicate submodule. Initialize and refresh it with:

```bash
npm run upstream:setup
npm run upstream:fetch
npm run upstream:status
```

Review changed path names with `npm run upstream:diff`, or pass explicit community paths to inspect their patches, for example `npm run upstream:diff -- apps/docs packages/docx-engine`. The tooling excludes and refuses `ee/`. Fetching a candidate never changes the pinned baseline and upstream commits must not be merged into the TandemFolio branch.

See [`migration/upstream-sync.md`](migration/upstream-sync.md) for classification, selective porting, provenance, and baseline-advance rules.

## Standalone browser behavior

- The File System Access API is used when the browser exposes it.
- A standard `<input type="file">` opens matching DOCX, Markdown, XLSX, PPTX, or PDF files when native file handles are unavailable.
- The File ribbon tab remains visible on macOS because TandemFolio browser/MCP hosts have no native File menu.
- `office_open_local_file` accepts an exact matching absolute DOCX/Markdown/XLSX/PPTX/PDF path and streams it to the renderer in bounded chunks.
- Existing file handles can be overwritten after permission is granted.
- Without a writable handle, save falls back to downloading a format-matching copy.
- PDF Save As/export-image/page-extract/print remain host effects. Mixed-PDF page insertion is
  implemented through the bounded `pdf.page.insert` staged-byte persist/reload route.
- DOCX bundled fonts are loaded lazily by allowlisted face and detected document script; the initial UI resource does not contain font data URLs.
- When embedded in an MCP Apps host, the editor lets the SDK report its initial size, completes one immediate `waitMs: 0` session poll, waits for one stable layout frame, and then requests fullscreen once when the host reports support. It subsequently holds one bounded `waitMs: 10000` poll that Broker enqueue wakes immediately; empty timeouts re-arm, while transport errors alone retry after 500 ms. The ribbon keeps a fullscreen toggle for exiting or manually retrying when the host declines the initial request. Standalone browser mode does not render this host-only action.
- Dirty mounted documents are serialized by their format-owned pipeline and checkpointed locally after
  Agent mutations and at a two-second deduplicated interval for user edits. New sessions default to
  `resume: "none"`; a known document uses `resume: "exact"` with its recorded `sessionId`, while
  `resume: "latest"` is explicit cross-session disaster recovery. Recovery always mounts the matching
  renderer and never enables UI-closed editing.
- Recovery retains one snapshot per format/session, rejects snapshots over 256 MiB, expires after
  seven days, and retires older snapshots/uploads on committed bound Save (not Export Copy).

## Embedded save, resume, and performance troubleshooting

- Expand **文件位置** in the editor to see/copy the committed absolute path, including while waiting
  for a view handoff. Unbound documents explicitly show **尚未保存**. Context also
  returns `session.filePath`. New documents default to
  `~/Documents/TandemFolio/<format>-<session-hash>/<file-name>`; `TANDEMFOLIO_OUTPUT_DIR` overrides the
  root. `office_open_local_file` binds normal Save to that explicit path. Browser File/Open cannot
  discover an absolute source path: it detaches the old target and the next Save creates a new output
  file. Save As rebinds; Export Copy reports its destination without changing the normal Save target.
- Continuing an opened/restored document preserves that file as well as its Session. Do not call
  `pptx.document.create_blank` to prepare a follow-up generation: it deliberately replaces the
  document and detaches the Save target. ADR 0015 rejects resetting named, edited, or history-bearing
  presentations unless `confirmReplace: true` records explicit user replacement intent. A rejection
  leaves the original intact; do not auto-retry with the flag. Apply slide/object edits to continue,
  then use normal Save when requested. Confirmed new saves separately without overwriting the
  prior file; pristine initial blank presentations do not require confirmation.
- A host cold remount or broker restart restores only the recorded Session, preferring its checkpoint
  and otherwise reading its bound file through the native renderer. Keep the original `sessionId`;
  do not use `latest` for a known document. Missing/unreadable sources display an error, not a silently
  empty successful editor. Snapshots are best-effort crash protection, not guaranteed delivery;
  explicitly Save important changes. Native undo history is not persisted across cold mounts.
- A visible replay of the same logical view automatically requests safe handoff from its hidden
  owner. Keep the same Session/view; do not create more editors. If both surfaces wait, or an old
  mount still reports active although the host no longer displays it, click **在此继续编辑** on the
  desired surface. The old owner must still cooperate and checkpoint before authority transfers.
  Automatic waiting stops after 30 seconds; explicit continuation/retry starts another attempt.
  If an old editor disappeared without disconnect, its idle lease can expire after 30 seconds,
  but prepared/uncertain mutations still prevent takeover. A failed checkpoint keeps the original;
  address the reported disk/connection problem before retrying. An uncertain handoff locks the old
  surface; its **重试连接** first confirms ownership, never blindly unlocks it. Yielded views do not
  reclaim authority when they become visible; select **在此继续编辑** to return to one. Restoration
  must finish before editing is enabled. Checkpoints do not imply formal Save or persist native undo.
  Hidden automatic waiting views do not retry; an explicitly selected continuation attempt retains
  its intent through stale visibility hints, still bounded by the 30-second deadline. Healthy hidden
  owners keep one 10-second bounded poll while
  suspending heavy rendering, and yielded owners stop polling. Transient transport errors back off.
- Run `npm run build`, then
  `npx playwright test --config playwright.host.config.ts tests/visual/document-lifecycle.spec.ts tests/visual/session-handoff.spec.ts tests/visual/view-continuation.spec.ts`.
  This suite uses real packaged renderers and isolated stdio brokers to verify each format's Save
  path, exact cold resume after broker restart, same-name browser import/checkpoint/re-save without
  overwriting the previous file, and terminal-error polling across multiple PPTX
  views. Handoff scenarios cover all five native formats, unsaved PPTX round trips, hidden waiting
  views, abandoned owners, failure/explicit retry, path visibility, active-owner user continuation,
  lost prepare/commit responses, yielded views remaining suspended, and PPTX rejection of implicit
  reset followed by edits/saving the original path or explicitly confirmed separate-file creation.
  Tests write only temporary
  fixtures/output and are not full source-current release evidence.

Previous ADR 0013 validation baseline: all fourteen lifecycle/handoff browser scenarios passed, along with
42 Host Bridge tests, six dedicated Broker handoff regressions, typecheck/build/package/MCP smoke,
asset budgets, and license checks. The full repository suite still has three release-evidence failures
(readiness expected true, baseline CLI evidence, missing upstream visual manifest); `release:gate`
reports missing `release/release-evidence.json`. Do not convert these into passing release claims.

These fallbacks are development and recovery behavior; they do not create a second document authority.

ADR 0014 validation (2026-09-03): all **22** packaged-browser lifecycle/handoff/continuation cases,
**44** Host Bridge tests, and **21** focused SessionStore/Broker tests pass. All-workspace typecheck,
five-format build/package, operation-manifest checks, MCP smoke (25 tools), asset budgets, plugin/
skill validation, and license checks pass. The full suite retains the same three release-evidence
failures above. `upstream:check` separately stops because the existing `AGENTS.md` does not record
the configured baseline; it was not changed by this repair. `release:gate` remains fail-closed.
These browser cases use an isolated real Broker and renderers, not a completed native Codex UI
acceptance run. After local plugin reinstall, use a new task to load the new tools/skill; existing
mounted resources do not hot-update. Preserve old unsaved editors and exact document handles.

ADR 0015 validation (2026-09-03): TDD first reproduced a successful implicit reset to Untitled after
real view continuation, then verified rejection without content/path/revision changes. A second
RED/GREEN cycle added typed explicit confirmation and separate-file Save. All **24** packaged-browser
lifecycle/handoff/continuation cases and **177** PPTX tests pass, including continued slide edits
persisting to the original path, clean named/restored blanks, pristine initialization, and preserved
undo/redo. Full-workspace tests report **3,439 passing, one skipped, and the same three release-evidence
failures** listed above. All-workspace typecheck and operation Manifest checks pass. Five-format
build/package, installed MCP smoke, asset budgets, plugin/Skill validation and licenses pass. No
native Codex UI acceptance or formal release certification is claimed; `release:gate` still reports
missing evidence and remains false. User documents were not changed or merged by this repair.

## Codex host width visual regression

Latest release-evidence verification (2026-09-04): the visual manifest and
28 inherited screenshot references are now retained in the checkout. A full 7/7/21
five-format capture produced an approved `release/release-evidence.json` after the
XLSX identity-entry change: XLSX bootstrap p95 is 432.5 ms against
the fixed 500 ms gate; Markdown small-file open p95 is 14.3 ms against 20 ms. The gate
accepts the source-current bundle. See [the verification record](../release/validation.md)
for TDD fixes, actual measurements, test results and the next release steps. Do not
copy `approved`/`ready` flags from historical evidence or relax the fixed ceiling.

`npm run test:visual:host` serves each active packaged renderer in a deterministic Codex host harness. The harness uses the real MCP Apps `AppBridge` and postMessage transport instead of replacing the renderer with a test double. Chromium is fixed to device scale factor 1, Simplified Chinese, light color scheme, and reduced motion.

The committed matrix covers four observable host states for every active format:

| Scenario         | Editor viewport | Required evidence                                                                  |
| ---------------- | --------------: | ---------------------------------------------------------------------------------- |
| Narrow sidebar   |       420 x 900 | editor is nonblank, reports its mounted size, and remains mounted                  |
| Split view       |       720 x 900 | editor remains visible at the normal Codex right-panel width                       |
| First fullscreen |     1332 x 1280 | initial size precedes first poll and fullscreen request; iframe loads exactly once |
| Exit fullscreen  |       720 x 900 | real editor toggle returns inline without remounting the iframe                    |

The 1332-pixel fullscreen width is the measured right-side editor region in the 5120-pixel, 2x Codex
reference capture. All five formats are active in the shared format-driven gate, so CI executes
twenty shared pixel cases. XLSX waits for its real Univer canvas; PPTX waits for its retained stage;
PDF opens a deterministic local fixture and waits for its retained page canvas. PPTX and PDF additionally keep their format-specific
four-scenario suites and interaction/save/reopen tracers.

Update snapshots only after inspecting every changed image and confirming that a renderer or host-contract change is intentional. The CI job installs its pinned Playwright Chromium and runs the same public command.

## Reproducible release evidence

ADR 0005 makes readiness a generated projection of an approved evidence bundle. A normal capture
is deliberately unapproved and may use reduced sample counts for local iteration:

```bash
npm run release:baseline:capture
```

The source fingerprint covers renderer and host sources, dependency locks, packaging and gate
logic, stable plugin inputs, and visual provenance. It intentionally excludes
`plugins/tandemfolio/assets/editors`, which is regenerated by `npm run build` and may differ at the
byte level between build platforms. Those generated bundles remain covered by the packaged-host
visual suite, asset budgets, MCP smoke test, and final distribution checksums.

After inspecting the five pinned-upstream split-view comparisons, capture the formal baseline with
at least seven cold starts, seven opens per fixture size, and twenty-one interactions per format:

```bash
npm run release:baseline:capture -- \
  --approve \
  --output release/release-evidence.json
npm run release:gate
```

The gate rejects missing, malformed, unapproved, undersampled, over-budget, or artifact-hash-mismatched
evidence and writes an all-format `ready: false` projection when it does. It also recomputes the release-source fingerprint, so any release-relevant source,
configuration, visual manifest, or packaged editor change invalidates an older capture. Run the
full release pipeline only after the gate passes:

```bash
npm run release:package
```

Schema-v4 evidence also requires the one-shot XLSX cold-start phase families
`bootstrapMs`, `univerCreateMs`, `worksheetInstallMs`, and `firstCommitMs`, plus bootstrap
subphases `resourceReceiveMs`, `moduleGraphReadyMs`, and `reactMountMs`, from every cold-start
sample. The subphase sum must equal `bootstrapMs` within measurement precision. The first poll must
already observe an active workbook, active worksheet, and Univer canvas; XLSX p95 is fixed at
1,400 ms and bootstrap p95 is independently fixed at 500 ms. Renderer RSS is sampled
asynchronously and serially so the measurement process does not block Playwright's event loop. Do
not replace that sampler with a synchronous high-frequency process probe or derive either fixed
budget from the candidate being approved. Every format is measured in a fresh Chromium process to
prevent cross-format heap and renderer-process history from contaminating cold-start or memory
evidence. Capture marks a format failed before approval when any fixed budget is exceeded;
`release:gate` independently rechecks it.

The committed evidence and screenshots are release inputs. `release/*.candidate.json` remains a
local iteration artifact and is ignored.

## Build outputs

For GitHub Actions publication and Windows/macOS end-user installers, see
[Release distribution](distribution.md). `npm run test:distribution` exercises bundle
integrity and native installer preflight in temporary directories. After approved evidence
is available, `npm run release:distribute` runs the existing release-package chain and a
second read-only source gate before archiving into `out/releases/`. ZIP/tar creation uses
`tar` and `zip` on macOS/Linux; the resulting Node/Web bundle supports both Windows/macOS.
The generated archive has a separate `tandemfolio-releases` catalog and preserves the
checkout's existing `personal` catalog and local manifest cachebuster. `.github/workflows/release.yml`
publishes only after native Windows/macOS extracted-package smoke succeeds.

`npm run build` produces:

| Output                                                                   | Purpose                                     | Tracked |
| ------------------------------------------------------------------------ | ------------------------------------------- | ------- |
| `apps/docs/dist/renderer/`                                               | Intermediate Vite build                     | No      |
| `apps/{markdown,sheets,slides,pdf}/dist/renderer/`                       | Intermediate format builds                  | No      |
| `apps/mcp-server/dist/server.js`                                         | Intermediate bundled MCP server             | No      |
| `plugins/tandemfolio/assets/editor/index.html`                           | Self-contained MCP App HTML/CSS/JS resource | Yes     |
| `plugins/tandemfolio/assets/editors/{markdown,xlsx,pptx,pdf}/index.html` | Self-contained format resources             | Yes     |
| `plugins/tandemfolio/assets/fonts/`                                      | External lazy DOCX font assets              | Yes     |
| `plugins/tandemfolio/dist/server.js`                                     | Self-contained plugin MCP server            | Yes     |
| `plugins/tandemfolio/LICENSE`, `NOTICE`                                  | Distributed attribution                     | Yes     |

The packaging script in `tools/package-plugin.mjs` emits one UI resource per format while copying
fonts separately. DOCX, Markdown, PPTX, and PDF inline their Vite CSS and JavaScript directly. XLSX
preserves Vite dynamic-import boundaries, gzip/base64 embeds every emitted module in the same HTML,
and starts it through an in-memory blob-URL module vault; no sibling hashed chunk or network fetch
is required. It inserts asset contents with replacement callbacks so JavaScript replacement tokens
such as `$&` remain literal, and escapes raw `</script` sequences before direct embedding.
Packaged-resource tests reject external JavaScript/CSS references, enforce per-format raw budgets,
and cap the inflated XLSX entry module at 11,000,000 bytes.

## Product-boundary checks

```bash
npm ls electron electron-updater \
  @genoffice/agent-core @genoffice/ai-provider \
  @genoffice/ai-search @genoffice/project-store @genoffice/electron-utils --all

rg -n '"(@genoffice/(agent-core|ai-provider|ai-search|project-store|electron-utils)|electron|electron-updater)"' \
  package-lock.json

rg -i 'genspark|@genoffice/ai-provider|@genoffice/agent-core' \
  plugins/tandemfolio/assets/editor/index.html plugins/tandemfolio/assets/editors

find . -maxdepth 2 -type d -name ee -print
```

Expected results: an empty `npm ls` tree and no matches from the remaining commands.

Run the production dependency audit separately from development fixtures:

```bash
npm audit --omit=dev --audit-level=high
```

As of 2026-08-28, the production audit reports 45 high-severity transitive findings rooted in the pinned Univer dependency on vulnerable `nanoid`; npm reports no available fix. A normal install reports four additional development-only findings. Treat this audit separately from the prohibited-product dependency search, and do not claim a clean production audit until the pinned renderer dependency chain is deliberately reviewed and updated.

## Validate the plugin

The validator needs PyYAML. Keep it in an ignored local environment:

```bash
python3 -m venv .scratch/plugin-validator
.scratch/plugin-validator/bin/pip install PyYAML
.scratch/plugin-validator/bin/python \
  "$CODEX_HOME/skills/.system/plugin-creator/scripts/validate_plugin.py" \
  plugins/tandemfolio
```

Validate the embedded Skill after editing it:

```bash
python3 "$CODEX_HOME/skills/.system/skill-creator/scripts/quick_validate.py" \
  plugins/tandemfolio/skills/tandemfolio
```

## Install the repository marketplace

This repository uses a non-default marketplace file at `.agents/plugins/marketplace.json`. Confirm its validated name first:

```bash
python3 "$CODEX_HOME/skills/.system/plugin-creator/scripts/read_marketplace_name.py" \
  --marketplace-path .agents/plugins/marketplace.json
```

The current name is `personal`. If this repository marketplace is not already configured:

```bash
codex plugin marketplace add .
codex plugin add tandemfolio@personal
```

If `personal` conflicts with an existing marketplace, stop and resolve the marketplace identity through the plugin scaffold workflow; do not hand-edit installed Codex marketplace state.

Start a new Codex task after installation so the new Skill and MCP tools are discovered.

## Update and reinstall during development

After changing plugin contents:

```bash
npm run check

python3 "$CODEX_HOME/skills/.system/plugin-creator/scripts/read_marketplace_name.py" \
  --marketplace-path .agents/plugins/marketplace.json

python3 "$CODEX_HOME/skills/.system/plugin-creator/scripts/update_plugin_cachebuster.py" \
  plugins/tandemfolio

codex plugin add tandemfolio@personal
```

The cachebuster helper preserves the base version and replaces the prior `+codex.<token>` suffix. Do not repeatedly increment the semantic version just to refresh a local installation. Test the result in a new Codex task.

## Release checklist

1. Update `CONTEXT.md` or add a superseding ADR if authority or product boundaries changed.
2. Update the migration and provenance ledgers.
3. Capture or verify approved R6 evidence, run `npm run release:gate`, then run `npm run check` and
   `npm run smoke:mcp`.
4. Run product-boundary searches and `npm audit --omit=dev`.
5. Validate the plugin and Skill.
6. Inspect packaged asset sizes and investigate unexpected growth.
7. Confirm `LICENSE` and `NOTICE` are present in the plugin.
8. Verify the editor mounts once and mutations do not create another UI resource.
9. Test local and Agent path open, edit, save, reopen, recovery, and an Agent mutation against the same active context for every advertised format.
10. Reconcile the pinned renderer command inventory. Release readiness requires the complete
    retained UI and MCP parity from ADR 0003 plus passing, source-current ADR 0005 evidence.

## Troubleshooting

### `editor_offline`

Call `office_show_editor`, keep the iframe mounted, and wait until `office_get_context` returns `connected: true`. Do not bypass the error with hidden document state.

### `revision_conflict`

Read fresh context and rebuild the operation from the current selection. Never resend the old `baseRevision`.

### `command_in_flight` or `command_timeout`

Do not enqueue a different mutation while an accepted transaction may still be applying. Keep the
mounted editor open. After `command_timeout`, call `office_execute` again with the exact same
`requestId`, `baseRevision`, operation id, and arguments; structural object-key order may differ.
The replay joins the in-flight execution or returns its cached final response without redispatch.
Use a new request id only after fresh context grounds a genuinely new intended mutation. A
`request_reused` response means the payload changed and must not be retried under that id.

### Editor shows the “run npm run build” fallback

The MCP server could not read `assets/editor/index.html` from the packaged plugin or the repository build fallback. Run `npm run build`, confirm the file is included in the plugin source, and run `npm run smoke:mcp`; the smoke test copies the plugin to an install-like temporary layout so repository-relative resource paths cannot pass accidentally.

### Editor HTML loads but scripts do not execute

Confirm the packaged HTML contains inline JavaScript and the CSP permits the generated inline module. Also confirm no hashed sibling JavaScript file remains referenced and run `npm run smoke:mcp`; multiple raw `</script>` tags indicate that asset contents corrupted the enclosing inline module during packaging.

### Codex still exposes old tools or Skill instructions

Run the cachebuster/reinstall loop above and test in a new task. Rebuilding files alone does not refresh an already loaded plugin.

### Browser save downloads instead of overwriting

The browser did not grant or expose a writable `FileSystemFileHandle`. This is expected fallback behavior, not an IPC failure.

### Bundle size grows sharply

Check for reintroduced font imports, AI strings/styles, dynamic chunks, or binary data URLs. Keep renderer, font/resource, and MCP server size changes separate in the investigation.
