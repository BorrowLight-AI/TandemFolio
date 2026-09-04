# Performance and visual baselines

## Current verification (2026-09-04)

The [source-current verification record](../release/validation.md) and
`release/release-evidence.json` contain the new complete five-format 7/7/21 capture.
Visual differences are 0.1–0.5% under the unchanged 3% limit.
The source-current approved capture passes every fixed gate. XLSX bootstrap p95 is
**432.5 ms < 500 ms**, with two preceding seven-sample XLSX candidates at 423.5 ms
and 422.7 ms. Markdown small-file open p95 is **14.3 ms < 20 ms**. The generated
readiness projection is true for all five formats; Windows/macOS release CI remains pending.
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

| Format   | Raw bytes | Gzip bytes | Raw budget |
| -------- | --------: | ---------: | ---------: |
| DOCX     | 3,195,997 |    886,702 |  3,500,000 |
| Markdown | 1,757,105 |    532,321 |  2,500,000 |
| XLSX     | 6,487,913 |  4,794,081 | 21,000,000 |
| PPTX     | 3,373,482 |    988,047 |  4,000,000 |
| PDF      | 6,619,114 |  3,366,252 |  7,000,000 |

The plugin is about 55 MB on disk because allowlisted document/PDF edit fonts remain external lazy
assets (about 20 MB) rather than entering the initial renderer HTML. The bundled MCP server is
1,804,055 bytes raw.

The DOCX ceiling allows completion of typed registry parity without treating capability removal as
an optimization. XLSX keeps the full permitted pinned community App in one HTML resource, but its
Vite modules are individually gzip/base64 embedded: the packaged entry inflates to 10,081,034 bytes
under an 11,000,000-byte gate and optional locale/hyphenation modules inflate only on demand. Both
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
| DOCX     |    328.823 |      131.2 |       182.0 |      335.9 |         5.9 |
| Markdown |    251.242 |       16.7 |        77.1 |      385.8 |         2.0 |
| XLSX     |    728.690 |      105.9 |       114.3 |      368.2 |        33.3 |
| PPTX     |    350.018 |       14.0 |        31.3 |      103.8 |        26.3 |
| PDF      |    437.481 |      130.6 |       132.9 |      150.2 |        32.8 |

The Markdown canonical-large p95 remains 98.3% below the R6-02 22,175.5 ms baseline at 385.8 ms.
The source-current schema-v4 phase p95 values are:

| Fixture | Decode | Parse | TipTap state install | React commit |
| ------- | -----: | ----: | -------------------: | -----------: |
| Small   |    1.2 |   4.8 |                  6.3 |          0.6 |
| Medium  |    0.2 |  22.2 |                 35.2 |          3.2 |
| Large   |    0.2 |  53.3 |                199.2 |          8.4 |

The residual between these four owned phases and total renderer execution includes registry
dispatch and Markdown local-image hydration. The acknowledgement remains after the React commit;
phase data contains durations only, never document text or paths.

R6-05 reduces the approved R6-04 XLSX cold-start p95 from 1,188.7 ms to 728.69 ms, a 38.7%
reduction, and retains the fixed 1,400 ms total gate. Bootstrap falls from 613.5 ms to 469.0 ms,
23.6% lower and below its new fixed 500 ms gate. The seven-sample aggregate phase p95 values are:

| Bootstrap | Univer create | Worksheet install | First commit |
| --------: | ------------: | ----------------: | -----------: |
|     469.0 |          11.6 |              15.7 |         64.8 |

The bootstrap p95 decomposition is:

| Resource receive | Module graph ready | React mount |
| ---------------: | -----------------: | ----------: |
|             23.1 |              411.8 |        35.9 |

The complete split module set stays gzip-compressed inside the same HTML; only the initial graph is
inflated/evaluated before React, while optional locale and hyphenation modules retain their lazy
imports. The residual to host-observed cold start includes outer host/AppBridge setup. The first
successful poll remains after an active workbook, active worksheet, and canvas; its trace is
retried on transport failure and consumed once. No preset, Registry operation, locale, history
route, or persistence behavior was removed.

ACK p95 decomposition aggregates the three open tiers plus the interaction samples for each
format. `transport` is the residual between total host-observed acknowledgement time and the
measured poll wait, hydration, and renderer execution durations.

| Format   | Poll wait | Hydrate | Renderer execute | ACK transport |
| -------- | --------: | ------: | ---------------: | ------------: |
| DOCX     |       0.1 |     1.3 |            315.0 |          12.9 |
| Markdown |       0.0 |    94.4 |            238.9 |          30.3 |
| XLSX     |       0.1 |    13.1 |            345.1 |           6.3 |
| PPTX     |       0.1 |     5.0 |             79.6 |          18.0 |
| PDF      |       0.1 |     9.8 |            125.1 |          11.1 |

Peak memory is sampled throughout cold start, open, and interaction scenarios. JS heap is the
primary stable browser gate; renderer RSS is the summed Chromium renderer-process diagnostic.

| Format   | Peak JS heap | Peak renderer RSS |
| -------- | -----------: | ----------------: |
| DOCX     |     80.9 MiB |         345.1 MiB |
| Markdown |     56.6 MiB |         383.7 MiB |
| XLSX     |    129.0 MiB |         482.4 MiB |
| PPTX     |    109.4 MiB |         378.6 MiB |
| PDF      |    106.1 MiB |         351.7 MiB |

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
