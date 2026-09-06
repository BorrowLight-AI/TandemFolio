# ADR 0016: Persist Markdown companion assets with the renderer-produced document

- Status: Accepted
- Date: 2026-09-06
- Extends: [ADR 0011](0011-session-bound-local-document-persistence.md) and [ADR 0012](0012-exact-document-resume-and-save-target-lifecycle.md)

## Context

Markdown documents can refer to image files beside the document. The mounted renderer already owns
the authored Markdown, resolves local images for display, and serializes the final text. Saving only
the `.md` bytes loses newly pasted images or produces broken paths after Save As. Letting the broker
parse or rewrite Markdown would create a second document authority and violate the product boundary.

## Decision

Markdown may extend the existing session-bound save upload with bounded opaque companion files and
content-addressed orphan candidates.

- The renderer scans Markdown and raw HTML image destinations outside code ranges. It leaves
  absolute, remote, query-bearing, or traversing sources unchanged, rewrites safe copied images to
  `assets/` paths, and sends the final Markdown bytes plus opaque companion bytes.
- Companion paths are relative PNG, JPEG, or GIF paths. Each file is limited to 20 MiB, a save has at
  most 128 companions, and document plus companions remain inside the existing 256 MiB transaction
  budget.
- The broker validates paths, sizes, leases, offsets, and target containment without parsing Markdown
  or image semantics. It writes companion temporary files first and renames the document last.
- Existing companion files are reused only when their bytes match. A different existing file causes
  a conflict; newly created companions are removed if the document cannot commit.
- Orphan collection is limited to content-addressed paths created by the current live session. The
  broker deletes a candidate only when its on-disk hash still matches its name; missing or modified
  files are preserved and ownership is relinquished.
- Standalone browser mode applies the same order through a user-authorized directory handle. When a
  directory handle is unavailable, the original data URL remains in the Markdown instead of claiming
  that a companion file was saved.
- Manual save, autosave, and typed MCP save share one serialized renderer queue.

## Consequences

Save As keeps local images usable without Electron, IPC, a custom URL protocol, or a second Markdown
runtime. The broker remains an opaque persistence sink. Ownership is intentionally session-scoped;
after a broker restart, unknown historical assets are preserved rather than collected.
