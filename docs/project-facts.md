# Project facts and attribution

This page is the public factual reference for TandemFolio's status, source provenance, and distribution obligations. It complements the file-level [source provenance record](migration/provenance.md); it is not a substitute for that record or for the license text.

## Identity and current status

- **Product:** TandemFolio
- **Form:** a local-first MCP Apps plugin with persistent visual editors and a local MCP server
- **Formats:** DOCX, XLSX, PPTX, PDF, and Markdown
- **Status:** pre-release. Source-current local evidence passes and the generated release-readiness projection is `true` for every format; native release CI and publication remain pending.
- **Distribution license:** [Apache License 2.0](../LICENSE)

The [2026-09-03 full capture](../release/validation.md) now supplies the previously
missing evidence and visual provenance. The 2026-09-04 approved recapture passes:
XLSX bootstrap p95 is 432.5 ms against 500 ms and Markdown small-file open p95 is
14.3 ms against 20 ms. This authorizes packaging but is not a published release.

## Source provenance

TandemFolio is a modified derivative of the Apache-2.0 community edition of GenOffice. The exact upstream repository and pinned commit are recorded in [source provenance](migration/provenance.md). The project retains format renderer structures where a host adapter can replace the former shell integration, while excluding source outside the documented product boundary.

The file-level provenance record is authoritative for:

- the pinned upstream baseline;
- extraction exclusions;
- TandemFolio-owned files; and
- every recorded modified upstream community file.

The [migration ledger](migration/ledger.md) records which upstream areas are retained, adapted, removed, or excluded. It is the operational companion to the provenance table.

## Upstream review boundary

The community source is tracked through a read-only `upstream` remote. Use `npm run upstream:fetch` to inspect candidates, but never merge or rebase `upstream/main` into the TandemFolio branch. Port reviewed changes selectively and advance the pinned baseline only with matching provenance and ledger updates.

Enterprise-only source is outside this project's boundary: do not inspect, copy, or depend on it. The safe, repeatable workflow is documented in [Upstream synchronization](migration/upstream-sync.md).

## Copyright, license, and modification obligations

When distributing or changing upstream-derived code, TandemFolio follows these rules:

1. **Keep original copyright notices.** Do not remove or obscure a copyright header, attribution, or required notice from copied or modified source.
2. **Distribute the Apache-2.0 license.** The repository's [LICENSE](../LICENSE) is included with the source and copied into the packaged plugin at `plugins/tandemfolio/LICENSE`.
3. **Make modifications prominent.** Every newly modified upstream-derived file must carry a clear `Modified by TandemFolio contributors` notice in the file itself, or be named in a file-specific entry in [NOTICE](../NOTICE). Update the file-level provenance record in the same change.
4. **Keep distribution notices intact.** [NOTICE](../NOTICE) and the plugin's `NOTICE` accompany distributed code. The root notice points to the file-level modification record; bundled third-party font and Unicode notices remain separately documented.

These practices are mandatory contribution requirements, not marketing claims. The exact Apache-2.0 terms in [LICENSE](../LICENSE) control if this summary and the license text differ.

## Maintainer workflow for upstream-derived changes

1. Read the relevant architecture decision and source record.
2. Preserve the original header and add a prominent modification notice when changing an upstream-derived file.
3. Update [source provenance](migration/provenance.md) and the [migration ledger](migration/ledger.md).
4. Retain `LICENSE` and `NOTICE` in both source and packaged plugin output.
5. Run the repository license check before publishing.

For selective upstream review and baseline advancement, follow [Upstream synchronization](migration/upstream-sync.md). For all other contributor guidance, see [CONTRIBUTING.md](../CONTRIBUTING.md).
