---
name: tandemfolio
description: Open and operate the currently declared TandemFolio visual editors for local DOCX, Markdown, XLSX, PPTX, or PDF files in Codex.
---

# TandemFolio

Use the live visual editor as the authority for document state.

1. Call `office_get_capabilities` for the format's default bounded summary. Optionally filter by
   `family` and follow `pagination.nextCursor`; each page contains at most twenty schema-free
   operation summaries. `ready` reports the shared ADR 0003 release gate, not whether the editor
   can open a file. When it is `false`, disclose that release validation remains.
2. Resolve the exact editing-session handle before opening anything. For a follow-up edit to a
   document already opened in the current task, reuse its previously returned `sessionId` and call
   `office_get_context` first. Match the user's reference against the recorded format and file name;
   if multiple mounted documents remain plausible, ask which one rather than creating or showing
   another editor.
3. Branch on that exact context for DOCX, Markdown, XLSX, PPTX, and PDF alike:
   - When `connected: true`, continue against that `sessionId`. Do not call
     `office_create_session` and never call a show tool for an already-connected follow-up edit.
   - When the Session exists with `connected: false`, call only its matching show tool and wait for
     `office_get_context` to report `connected: true` again.
   - When the Broker returns `session_not_found` for a known prior `sessionId`, call
     `office_create_session` with the matching format, `resume: "exact"`, and that same `sessionId`.
     Then call the matching show tool once. If no exact recovery exists, stop; never substitute a
     different document.
   - Only for an explicitly new document, or when the current task has no prior document handle,
     call `office_create_session` with `resume: "none"`, retain the returned `sessionId` as that
     document's task handle, and call its matching show tool once: `office_show_editor` (DOCX),
     `office_show_markdown_editor`, `office_show_xlsx_editor`, `office_show_pptx_editor`, or
     `office_show_pdf_editor`.
     Use `resume: "latest"` only when the user explicitly asks to recover the newest unsaved local
     checkpoint without identifying its prior Session. It is cross-session disaster recovery, never
     the default create or follow-up path.
4. To open an explicit local path, call `office_open_local_file` with the exact current revision and an absolute path whose extension matches the session. For user-driven opening, direct the user to the visible Open/File entry.
5. Select only a canonical id returned by summary discovery, then call
   `office_get_capabilities` with `view: "detail"`, that exact `operation`, and the current
   `sessionId`. Use the returned schema and require `availability.available: true`; never infer an
   input shape from the summary or this prose catalog.
6. Before each edit, call `office_get_context` and use its exact `revision` as `baseRevision`.
7. Generate a unique 1–128 character `requestId` for the intended mutation. Call
   `office_execute` with `{ sessionId, baseRevision, requestId, operations: [{ id, arguments }] }`,
   using exactly one canonical operation grounded in the active selection or document context. It
   returns only after renderer acknowledgement; claim success only when `structuredContent.ok` is
   true, `transaction.requestId` matches, `result.revision` advanced by one, and the operation result
   is present.
8. Read context again before issuing a dependent operation.
9. When the task creates a new deliverable and at least one intended content mutation succeeds,
   persist it exactly once after the first complete generation pass. Refresh context, discover the
   exact save descriptor, and execute the matching operation with `{}` and a fresh request id:
   `docx.document.save`, `markdown.document.save`, `xlsx.document.save`,
   `pptx.document.save`, or `pdf.document.save`. Treat the user's request to create or generate the
   deliverable as authorization to start this first save. In Codex, the mounted renderer submits its
   bytes through the Session-bound local persistence protocol; no host download prompt is required.
   Read context again and require a non-null absolute `filePath` before claiming delivery. New files
   land under `TANDEMFOLIO_OUTPUT_DIR`, or `~/Documents/TandemFolio` by default. A standalone
   browser file picker still requires user approval. Cancellation, permission denial, timeout without
   a confirmed replay result, a missing `{ saved: true }` result, or a missing `filePath` is not a
   saved delivery: do not retry under a new request id, do not claim completion, and state what action
   is required. Do not apply this automatic first-save rule to read-only inspection, a document
   opened only for editing, a recovered existing document, or every later follow-up edit unless the
   user also asks to save or has enabled the renderer's AutoSave behavior.

Never retry a new `revision_conflict` using the old revision; refresh context and assign a new
request id to the newly grounded intent. After `command_timeout`, do not create another mutation:
replay the structurally identical envelope with the exact same `requestId`. That replay joins the
in-flight execution or returns its final cached response without dispatching twice. Never retry
`request_reused`; it means the payload changed under an existing id. If
`transaction_not_atomic` is returned, split the work only into separately grounded one-operation
transactions with fresh context and revisions. If an edit returns `editor_offline`, use the exact
same `sessionId`: show that Session again when it still exists, or use `resume: "exact"` after a
Broker restart. Do not create or mutate a hidden second copy and do not fall back to `latest`.

Choose only canonical operations returned by summary discovery and use exact detail discovery as
the schema authority. Internal staged-load ids and compatibility aliases are intentionally absent.
The concise intent catalog is in [references/operations.md](references/operations.md).

Important grounding rules:

- DOCX text operations use the active caret/selection.
- Markdown text operations use the active TipTap caret/selection; context includes the selected text and active block type.
- XLSX mutations require an exact sheet and cell/range from fresh context.
- PPTX text/move operations require selected object ids on the active slide.
- PDF operations use exact saved-annotation, text/image, drawing, form, or page identities and
  bounded PDF-space coordinates from fresh context or preceding results. `pdf.page.insert` accepts
  an absolute local PDF path and is explicitly non-undoable because it persists and reloads the
  active document.
- `save` writes through the format-owned pipeline. A standalone editor writes through its granted
  file handle; a Codex-embedded editor submits the generated bytes and exact file name through the
  lease-checked local Broker protocol, which atomically commits the file and reports its absolute
  path in context. An opened file keeps its exact path; a new file receives a Session-isolated path;
  Save As rebinds future saves while Export Copy does not. The first-generation rule starts this
  operation automatically; standalone destination and permission prompts remain user-controlled.

Local recovery is crash protection, not headless editing. A restored document is editable only after the matching visual renderer reconnects.

This Skill describes the implemented live protocol. All five format-local retained-command producer
baselines and the shared R6-01 packaged-host, pinned-source visual, performance/resource, MCP smoke,
license, and repository gates are closed, so the generated capabilities report `ready: true` for
all five formats. If a package reports false, treat its release evidence as missing or stale rather
than inferring a format-local operation gap. Operations absent from bounded summary discovery are
unavailable to the Agent.
