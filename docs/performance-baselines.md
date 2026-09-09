# Performance and visual baselines

## Current verification (2026-09-08)

The [source-current verification record](../release/validation.md) and
`release/release-evidence.json` contain the new complete five-format 21/21/21 capture.
Visual differences are 0.1–1.2% under the unchanged 3% limit.
The source-current approved capture passes every fixed gate. XLSX bootstrap p95 is
**450.9 ms < 500 ms**, total XLSX cold start is **799.187 ms < 1,400 ms**, and
Markdown small-file open p95 is **13.3 ms < 20 ms**. The generated readiness
projection is true for all five formats; Windows/macOS release CI remains pending.
The reference PNGs and visual manifest are included in the checkout; no baseline
images were refreshed.

## Historical R6 measurements (not current release authorization)

The following 2026-08-31 values describe the earlier implementation. Its approved
JSON was not present in this checkout, and these historical numbers must not be
presented as current evidence or used to fabricate a replacement approved bundle.

- Captured: 2026-08-31 (Asia/Shanghai)
- Host: Apple M4, 10 logical CPUs, 16 GiB RAM, macOS arm64
- Browser: packaged Playwright Chromium 140.0.7339.16
- Scope: pinned-source packaged-host visual and R6-05 runtime/load/startup release evidence

## Packaged resource budgets

Run `npm run measure:assets` after `npm run build`. The command fails when a raw self-contained HTML resource exceeds its budget.

| Format   |  Raw bytes | Gzip bytes | Raw budget |
| -------- | --------: | ---------: | ---------: |
| DOCX     |  3,519,897 |    990,721 |  3,650,000 |
| Markdown |  2,573,562 |    936,420 |  2,750,000 |
| XLSX     |  8,734,947 |  6,475,283 | 10,000,000 |
| PPTX     |  3,891,323 |  1,146,141 |  4,000,000 |
| PDF      |  6,722,930 |  3,396,388 |  7,000,000 |

The plugin is about 55 MB on disk because allowlisted document/PDF edit fonts remain external lazy
assets (about 20 MB) rather than entering the initial renderer HTML. The bundled MCP server is
1,907,305 bytes raw. Markdown embeds its KaTeX WOFF2 fonts so formulas remain correct inside the
self-contained MCP resource.

The DOCX ceiling was rebaselined on 2026-09-05 for the audited upstream-native pagination, layout,
font and dialog port; it keeps about 4% raw headroom without treating capability removal as an
optimization. XLSX keeps the full permitted pinned community App in one HTML resource. Codex
rejects MCP App HTML above 10,000,000 UTF-8 Blob bytes, so every Vite module, including the initial
entry, is gzip/base64 embedded. The entry inflates to 7,860,540 bytes under the unchanged
11,000,000-byte gate; the shared Office font fallback, locale, package-I/O, operation-registry, and
hyphenation modules inflate only on demand. Both
raw and entry budgets are regression ceilings, not permission to remove renderer capabilities.

The PDF ceiling was recaptured after restoring browser-safe content-stream text/image editing with
the pinned community PDFium implementation. Its WASM is gzip-compressed once inside the
self-contained editor (the uncompressed fallback copy is removed at build time). PDF edit fonts
stay outside the initial HTML and are fetched only when a replacement cannot reuse a PDF standard
font; the CJK, Korean, and Arabic sfnt assets are derived from the already allowlisted OFL subsets.

## Browser mount check

The five standalone Vite entries were mounted in the Codex in-app browser at their fixed development ports. On the captured runs:

- all five document titles and expected empty states appeared;
- DOCX rendered its Ribbon and restored status-bar controls;
- Markdown rendered its retained Ribbon and TipTap surface; text input and bold formatting succeeded with no console errors;
- XLSX mounted the restored Ribbon/dialog shell and real Univer worksheet canvas, including browser and Agent-staged local open, Save copy, and Fullscreen; PPTX mounted all permitted pinned renderer sources through its original App/Ribbon/Konva surface; PDF mounted all permitted pinned renderer files and its retained Ribbon/PDF.js canvas;
- no browser console errors were reported for any empty-state mount;
- dark-mode layouts were visually inspected without clipping or overflow at the default viewport.

## Codex host pixel baselines

All five formats have committed Playwright pixel baselines for their real packaged renderers inside
an MCP Apps `AppBridge` harness. The shared 25-scenario matrix adds a 280 x 900 compact-pane geometry
gate to the 420 x 900 narrow sidebar, 720 x 900 split view, 1332 x 1280 first fullscreen, and return
to 720 x 900 on fullscreen exit. The compact gate requires the document surface to intersect the
viewport and forbids document-body horizontal overflow; Ribbon overflow remains reachable inside
its own scroll container. The PDF
gate opens a deterministic local fixture before asserting its page canvas. Assertions verify that
the renderer reports initial size before polling and requesting fullscreen, the editor is nonblank,
and the iframe loads exactly once across display-mode transitions.

The full real-host suite now lists 138 scenarios. It includes one-load offscreen release/resume
coverage for all five formats and per-version recovery checkpoint coverage for all five formats.
PPTX and PDF additionally retain their format-specific four-width suites and interaction/save/reopen
tracers. R6-01 adds a pinned-source provenance
manifest for the 720 x 900 split-view baselines. Every permitted renderer path is resolved at
`dc4d7e5927864498913b7ba42d0da06cc7cf628e`; no `ee/` path or visual mask is admitted. The approved
packaged-host capture has a `0` pixel-difference ratio for all five formats against those baselines.

## Runtime release baseline

Run `npm run release:baseline:capture` to create an unapproved candidate. An approved release
capture requires all five formats and at least seven cold starts, seven opens for each fixture size,
and twenty-one interactions. `npm run release:gate` validates the approved evidence and generates
the build-time readiness projection. `npm run release:package` is the final package path.

The following values are p95 milliseconds from the approved evidence. Cold start measures
navigation through the first visible, initialized renderer poll. R6-02 open and interaction
samples use the wakeable bounded-poll path; they no longer include the former fixed 500 ms command
cadence. R6-03 additionally records Markdown staged-load phases and fixes its open ceilings at
20 ms / 835.68 ms / 5,000 ms for small / medium / large. R6-04 records the four XLSX startup
phases and fixes XLSX cold-start p95 at 1,400 ms. R6-05 adds three bootstrap subphases and fixes
bootstrap p95 at 500 ms. Each format runs in a fresh Chromium process so
closed-format heap/process history cannot contaminate another format; RSS probes are asynchronous
and serialized so they do not block the Playwright observation loop. ACK columns keep delivery,
hydration, renderer execution, and return transport separate.

| Format   | Cold start | Open small | Open medium | Open large | Interaction |
| -------- | ---------: | ---------: | ----------: | ---------: | ----------: |
| DOCX     |    319.520 |       85.4 |       187.2 |      323.4 |        10.2 |
| Markdown |    270.741 |       13.3 |        61.4 |      379.7 |         2.4 |
| XLSX     |    799.187 |       46.1 |       141.7 |      404.8 |        45.5 |
| PPTX     |    491.776 |      183.4 |       138.4 |      238.4 |        22.6 |
| PDF      |    511.815 |      131.0 |       195.3 |      189.0 |        32.2 |

The Markdown canonical-large p95 remains 98.3% below the R6-02 22,175.5 ms baseline at 385.8 ms.
The source-current schema-v4 phase p95 values are:

| Fixture | Decode | Parse | TipTap state install | React commit |
| ------- | -----: | ----: | -------------------: | -----------: |
| Small   |    0.1 |   4.4 |                  4.2 |          0.6 |
| Medium  |    0.1 |  18.4 |                 12.9 |          0.7 |
| Large   |    0.2 | 128.9 |                 54.7 |          1.7 |

The residual between these four owned phases and total renderer execution includes registry
dispatch and Markdown local-image hydration. The acknowledgement remains after the React commit;
phase data contains durations only, never document text or paths.

The source-current capture retains the fixed 1,400 ms total and 500 ms bootstrap gates. The
seven-sample aggregate phase p95 values are:

| Bootstrap | Univer create | Worksheet install | First commit |
| --------: | ------------: | ----------------: | -----------: |
|     450.9 |           8.1 |              14.3 |         76.1 |

The bootstrap p95 decomposition is:

| Resource receive | Module graph ready | React mount |
| ---------------: | -----------------: | ----------: |
|             28.8 |              418.0 |        10.8 |

The complete split module set stays inside the same HTML. The initial App graph and optional
font, locale, package-I/O, operation-registry, metafile, command-dialog, and hyphenation modules stay
gzip-compressed and retain lazy imports. The residual to host-observed cold start includes outer
host/AppBridge setup. The first
successful poll remains after an active workbook, active worksheet, and canvas; its trace is
retried on transport failure and consumed once. No preset, Registry operation, locale, history
route, or persistence behavior was removed.

ACK p95 decomposition aggregates the three open tiers plus the interaction samples for each
format. `transport` is the residual between total host-observed acknowledgement time and the
measured poll wait, hydration, and renderer execution durations.

| Format   | Poll wait | Hydrate | Renderer execute | ACK transport |
| -------- | --------: | ------: | ---------------: | ------------: |
| DOCX     |       0.1 |     1.4 |            311.5 |          11.1 |
| Markdown |       0.1 |    85.1 |            183.6 |         111.3 |
| XLSX     |       0.1 |    16.7 |            363.6 |          10.4 |
| PPTX     |       0.1 |     5.6 |            206.3 |          18.7 |
| PDF      |       0.1 |    10.6 |            176.4 |           8.6 |

Peak memory is sampled throughout cold start, open, and interaction scenarios. JS heap is the
primary stable browser gate; renderer RSS is the summed Chromium renderer-process diagnostic.

| Format   | Peak JS heap | Peak renderer RSS |
| -------- | -----------: | ----------------: |
| DOCX     |     83.8 MiB |         351.8 MiB |
| Markdown |     55.8 MiB |         365.2 MiB |
| XLSX     |    175.4 MiB |         561.0 MiB |
| PPTX     |    131.1 MiB |         469.8 MiB |
| PDF      |    108.2 MiB |         420.8 MiB |

The exact samples, nearest-rank summaries, reviewed fixed Markdown open/XLSX cold-start/bootstrap
ceilings, derived phase and other regression ceilings, fixture hashes, environment profile, source fingerprint, screenshot
hashes, and artifact paths are stored in
`release/release-evidence.json`. Any release-relevant source change makes that evidence stale and
forces a new reviewed capture before `ready: true` can be regenerated.

## Optimization boundaries

Measure renderer, format engine, lazy font/resource, and MCP transport costs separately. Do not combine them into a single “editor latency” figure. Off-screen complex regions may use `content-visibility` only with an intrinsic-size placeholder and keyboard reachability verification. PDF Canvas output must retain a real DOM text representation for assistive technology and find-in-page behavior.

Retained but offscreen editors keep one bounded long poll for background Agent commands, suspend
periodic recovery serialization, and skip root layout/paint via the shared Host Bridge activity
marker. Browser `contentvisibilityautostatechange` also pauses renderer-heavy work at the pre-render
boundary. DOCX and Markdown release document DOM while retaining TipTap state; XLSX releases Univer
Canvas backing stores while retaining the workbook runtime; PDF cancels page/thumbnail rendering and
releases those canvases while retaining its loaded PDF document; PPTX unmounts its workspace Canvas
nodes while retaining `BrowserPresentation` and undo state. Each renderer resumes without reloading
its iframe, and each exposes a monotonic persisted-state recovery version so unchanged dirty content
is not serialized every two seconds. This releases the dominant renderer allocations but is not a
claim that every iframe allocation is freed; complete release still depends on the host discarding an
iframe.
