# ADR 0013: Cooperative handoff of replayed editor views

- Status: Accepted
- Date: 2026-09-03
- Extends: ADR 0006, ADR 0010, ADR 0011
- Supersedes in part: ADR 0012's terminal conflict behavior for replayed mounts of the same view

## Context

The host can retain a hidden renderer while replaying the same show result after a task switch.
The hidden renderer keeps its lease alive, so lease expiry alone never lets the returning view
restore. A conflict also hid the saved path and exposed an editable default blank document.

## Decision

- Session and logical view identities remain immutable. Only a visible replay of the same
  `(sessionId, viewId)` may request cooperative handoff; different views/tasks cannot request it.
- The hidden owner blocks local input and awaits its save/checkpoint work. An app-only prepare
  phase checks that no renderer command or byte upload is pending, then fences new mutations and
  formal saves. The renderer serializes a fresh forced checkpoint through its native format owner.
- Commit requires that newly committed checkpoint. The Broker transfers the lease and marks the
  Session offline until the candidate polls; its native restore command guards further mutations
  until ACK. Old-owner writes are rejected. Opaque recovery bytes never become a headless authority.
- A live visible owner is not forcibly displaced. Visible candidates retry at 1, 2, 4, then at most
  one attempt per 5 seconds; hiding a candidate stops these retries. An abandoned, idle owner may
  still expire after 30 seconds. Queued/active commands and prepared handoffs prevent expiry takeover.
- Checkpoint failure aborts the prepared handoff, discards only its unfinished recovery upload,
  retains the old editor, and stops automatic attempts. Explicit retry from the candidate starts
  a new attempt. If ownership or commit outcome is uncertain, do not overwrite or force takeover.
- A pending/failed restore blocks pointer and keyboard editing of the blank replica. The file
  location is shown independently of ownership: confirmed absolute path, confirming, or unsaved.
  Export Copy continues showing its output without changing the normal Save binding.

## Consequences

All five formats share this transport behavior without remounting a healthy editor on data commands.
Only one renderer accepts editing at a time. A host-created cold view can restore document bytes,
but native undo history is not persisted across that restore. Switching does not trigger formal
Save, change the output path, recover another Session, or resurrect contents absent from both the
exact checkpoint and saved file. An owner lost during a prepared/uncertain mutation remains a
fail-closed recovery case, not permission to steal its lease.

There are still ten public Agent tools and five UI resources. One app-only handoff endpoint is
added (fifteen transport endpoints total). Source-current release evidence remains required.
