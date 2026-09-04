# ADR 0015: Explicit PPTX document replacement intent

- Status: Accepted
- Date: 2026-09-03
- Extends: ADR 0004, ADR 0010, ADR 0012, ADR 0014

## Context

Exact Session/view continuation can restore the correct file and still be followed by an Agent
calling `pptx.document.create_blank` as a generation setup step. That command intentionally resets
the mounted document and detaches its Save target. Keeping the same Session is therefore not enough
to keep editing the same document: a subsequent Save legitimately creates `Untitled.pptx`.

## Decision

- Follow-up editing/generation in any format preserves the current document and its Save target.
  Reconnection is not authorization to reset, replace, reopen, or create a new document.
- PPTX's format-owned `pptx.document.create_blank` descriptor accepts optional boolean
  `confirmReplace`, defaulting to false. Before invoking the existing native replacement factory,
  the mounted host rejects an unconfirmed replacement if the presentation is named (including
  opened/restored/saved files), dirty, or retains undo/redo history. An untouched initial blank
  presentation remains usable without confirmation.
- Rejection leaves native contents/history, Save binding, recovery and revision unchanged. The
  operation returns an explanatory `execution_failed`; it is not a transient failure and does not
  authorize retry with confirmation enabled.
- The Skill requires explicit user replacement intent before setting `confirmReplace: true`,
  including acknowledgement that current in-editor contents will be discarded. The flag records
  caller intent; it is not independent proof of a human gesture or a new security boundary.
- Confirmed replacement keeps the existing reset semantics: detach the original Save target,
  retire its recovery, install/checkpoint one blank native presentation in the same mounted editor.
  Its next Save creates a separate collision-safe output, never overwriting the prior file.
- The visible File/New gesture remains on its existing format-owned factory. Other formats do not
  currently expose an equivalent whole-document reset operation; no new operations or generic
  dispatch/filter are added. Their continuation intent is clarified in the shared Skill.

## Verification

Real Broker/packaged-renderer tests cover explicit view continuation followed by a rejected reset,
normal slide editing and Save to the original path, plus confirmed replacement/Save preserving the
original bytes. Format-host tests cover pristine initialization, clean restored named blanks,
unnamed edited content, and undo/redo retention. Only temporary test documents are written. Existing
five-format lifecycle/handoff coverage and the fail-closed release gate remain required.
