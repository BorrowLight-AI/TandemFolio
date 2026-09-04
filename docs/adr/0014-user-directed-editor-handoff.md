# ADR 0014: User-directed continuation of replayed editor views

- Status: Accepted
- Date: 2026-09-03
- Extends: ADR 0010, ADR 0011, ADR 0012
- Supersedes in part: ADR 0013's visibility-only handoff eligibility and unbounded retry duration

## Context

Codex can mount one show result in several native surfaces. A retained surface can still report
intersection/activity when hidden or occluded, leaving every accessible replica waiting indefinitely.
Geometric visibility is useful for rendering work but cannot decide where the user wants to edit.

## Decision

- A waiting replica exposes **在此继续编辑**. This explicit app-only intent is scoped to the same
  immutable Session/view and its requesting mount; it never selects a different document.
- The current owner may cooperate even while reporting active. It blocks local input, drains its
  ongoing work, prepares the existing mutation fence, and commits a fresh native checkpoint before
  transferring the lease. The candidate must restore and acknowledge before becoming editable.
- No caller may force an uncertain/prepared transfer, bypass the owner, or use a stale checkpoint.
  Failure retains the original content and exposes a recoverable error.
- Automatic waiting is capped at 30 seconds per attempt. Expiry stops automatic retries and offers
  explicit actions; it is not authority to steal a lease. A completed outgoing mount stays suspended
  until the user explicitly continues there, preventing visibility-driven reacquisition loops.
  An explicit continuation attempt is not blocked by stale visibility hints; its retries retain
  user intent until restoration, failure, or the deadline, without changing heavy-render scheduling.
- All five formats use this shared transport. Their native serializers, history, Save bindings,
  immutable identity, Agent tool surface and fail-closed release gate remain unchanged.

## Verification boundary

Browser regression tests use real packaged renderers and the real Broker with multiple mounts of
one show result, including an occluded/CSS-hidden owner whose activity remains true. Resizing one
iframe alone is not evidence for native-host view replay or editor handoff.
