/**
 * List numbering actions: allocating numIds, new list definitions, restart /
 * continue numbering. Extracted from App.tsx; the App component passes a
 * NumberingContext built fresh per call so state never goes stale.
 */
import type { Editor } from '@tiptap/core'
import type { CustomNumberingLevel, NumberingDef } from '@genoffice/docx-engine'
import type { DocState, PendingNumbering } from './doc-state'
import { t } from './i18n/locale'

/** The App state the numbering actions need; built fresh per call. */
export interface NumberingContext {
  editor: Editor | null
  doc: DocState | null
  pendingNumbering: PendingNumbering
  setPendingNumbering: (update: (prev: PendingNumbering) => PendingNumbering) => void
  numIdFloorRef: { current: number }
  setStatus: (status: string) => void
}

export interface RestartNumberingInput {
  blockIndex: number
  start: number
}

export interface ContinueNumberingInput {
  blockIndex: number
  previousBlockIndex: number
}

export type NumberingActionResult =
  | { ok: true; changed: number }
  | { ok: false; error: 'invalid_arguments' | 'execution_failed'; message: string }

export function nextNumId(ctx: NumberingContext): string {
  let max = 2 // the blank template occupies 1/2
  for (const id of ctx.doc?.parsed.numbering.keys() ?? [])
    max = Math.max(max, parseInt(id, 10) || 0)
  for (const d of ctx.pendingNumbering.newDefs) max = Math.max(max, parseInt(d.numId, 10) || 0)
  for (const r of ctx.pendingNumbering.restartNums) max = Math.max(max, parseInt(r.numId, 10) || 0)
  max = Math.max(max, ctx.numIdFloorRef.current)
  ctx.numIdFloorRef.current = max + 1
  return String(max + 1)
}

/** Instant display of editor markers: inject pending definitions into listNumbering storage (re-parsing takes over after save) */
export function overlayNumberingDef(ctx: NumberingContext, def: NumberingDef): void {
  if (!ctx.editor) return
  const store = ctx.editor.storage.listNumbering as { defs: Map<string, NumberingDef> }
  store.defs = new Map(store.defs).set(def.numId, def)
}

/** Fallback when a new list can't reuse a numId: adopt an existing same-kind definition, otherwise create a new one */
export function createNumberingDef(ctx: NumberingContext, kind: 'bullet' | 'ordered'): string {
  // brand-new abstractNum + num (blank template style; when the part is missing the engine creates the part too)
  const numId = nextNumId(ctx)
  ctx.setPendingNumbering((p) => ({ ...p, newDefs: [...p.newDefs, { numId, kind }] }))
  overlayNumberingDef(ctx, {
    numId,
    abstractNumId: `pending-${numId}`,
    levels: Object.fromEntries(
      Array.from({ length: 5 }, (_, i) => [
        i,
        {
          numFmt: kind === 'bullet' ? 'bullet' : 'decimal',
          lvlText: kind === 'bullet' ? '' : `%${i + 1}.`,
          start: 1,
          indentLeft: 720 * (i + 1),
          hanging: 360,
        },
      ]),
    ),
    startOverrides: {},
  })
  return numId
}

/** New list definitions with custom levels (bullet library / numbering library / multilevel gallery / define dialog) */
export function createCustomListDef(
  ctx: NumberingContext,
  levels: CustomNumberingLevel[],
): string | null {
  if (!ctx.doc || levels.length === 0) return null
  const kind = levels[0].numFmt === 'bullet' ? ('bullet' as const) : ('ordered' as const)
  const numId = nextNumId(ctx)
  ctx.setPendingNumbering((p) => ({ ...p, newDefs: [...p.newDefs, { numId, kind, levels }] }))
  overlayNumberingDef(ctx, {
    numId,
    abstractNumId: `pending-${numId}`,
    levels: Object.fromEntries(
      levels.map((l, i) => [
        i,
        {
          numFmt: l.numFmt,
          lvlText: l.lvlText,
          start: l.start ?? 1,
          indentLeft: l.indentLeft,
          hanging: l.hanging ?? 360,
        },
      ]),
    ),
    startOverrides: {},
  })
  return numId
}

export function allocateListNumId(
  ctx: NumberingContext,
  kind: 'bullet' | 'ordered',
): string | null {
  if (!ctx.doc) return null
  // the document already has a same-kind numbering definition (even if unused in the body): the new num points at its abstractNum
  const match = [...ctx.doc.parsed.numbering.values()].find(
    (d) => (d.levels[0]?.numFmt === 'bullet') === (kind === 'bullet'),
  )
  if (match) {
    const numId = nextNumId(ctx)
    ctx.setPendingNumbering((p) => ({
      ...p,
      restartNums: [
        ...p.restartNums,
        { numId, abstractNumId: match.abstractNumId, startOverrides: { 0: 1 } },
      ],
    }))
    overlayNumberingDef(ctx, { ...match, numId, startOverrides: { 0: 1 } })
    return numId
  }
  return createNumberingDef(ctx, kind)
}

/** From the cursor's block onward, move list items sharing the numId to a new numId (Word: restart applies to all later items of that list) */
export function rewriteNumIdForward(
  ctx: NumberingContext,
  oldNumId: string,
  newNumId: string,
): void {
  if (!ctx.editor) return
  rewriteNumIdForwardAt(ctx, oldNumId, newNumId, ctx.editor.state.selection.$from.index(0))
}

function rewriteNumIdForwardAt(
  ctx: NumberingContext,
  oldNumId: string,
  newNumId: string,
  startIdx: number,
): number {
  if (!ctx.editor) return 0
  const { state, view } = ctx.editor
  let tr = state.tr
  let changed = 0
  state.doc.forEach((node, offset, idx) => {
    if (idx < startIdx) return
    if (node.type.name === 'docListItem' && node.attrs.numId === oldNumId) {
      tr = tr.setNodeMarkup(offset, undefined, {
        ...node.attrs,
        numId: newNumId,
        externalChanged: true,
      })
      changed++
    }
  })
  if (changed > 0) view.dispatch(tr)
  return changed
}

/** Restart one explicit top-level list block and all later items sharing its list identity. */
export function restartNumberingAt(
  ctx: NumberingContext,
  input: RestartNumberingInput,
): NumberingActionResult {
  if (!ctx.editor || !ctx.doc) {
    return {
      ok: false,
      error: 'execution_failed',
      message: 'DOCX numbering state is unavailable.',
    }
  }
  if (
    !Number.isInteger(input.blockIndex) ||
    input.blockIndex < 0 ||
    !Number.isInteger(input.start) ||
    input.start < 1 ||
    input.start > 1_000_000
  ) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: 'docx.list.restart requires a valid blockIndex and start value.',
    }
  }

  const anchor = ctx.editor.state.doc.maybeChild(input.blockIndex)
  if (!anchor || anchor.type.name !== 'docListItem' || anchor.attrs.numId == null) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: `DOCX block ${input.blockIndex} is not a numbered list item.`,
    }
  }

  const oldNumId = String(anchor.attrs.numId)
  const ilvl = Math.max(0, Math.min(8, Number(anchor.attrs.ilvl) || 0))
  const store = ctx.editor.storage.listNumbering as { defs: Map<string, NumberingDef> }
  const def = store.defs.get(oldNumId) ?? ctx.doc.parsed.numbering.get(oldNumId)
  let numId: string | null
  if (def) {
    numId = nextNumId(ctx)
    ctx.setPendingNumbering((pending) => ({
      ...pending,
      restartNums: [
        ...pending.restartNums,
        {
          numId: numId!,
          abstractNumId: def.abstractNumId,
          startOverrides: { [ilvl]: input.start },
        },
      ],
    }))
    overlayNumberingDef(ctx, {
      ...def,
      numId,
      startOverrides: { [ilvl]: input.start },
    })
  } else {
    const kind = (anchor.attrs.kind as 'bullet' | 'ordered' | null) ?? 'ordered'
    const levelCount = Math.max(5, ilvl + 1)
    numId = createCustomListDef(
      ctx,
      Array.from({ length: levelCount }, (_, level) => ({
        numFmt: kind === 'bullet' ? 'bullet' : 'decimal',
        lvlText: kind === 'bullet' ? '•' : `%${level + 1}.`,
        indentLeft: 720 * (level + 1),
        hanging: 360,
        start: level === ilvl ? input.start : 1,
      })),
    )
  }
  if (!numId) {
    return {
      ok: false,
      error: 'execution_failed',
      message: 'The DOCX numbering definition could not be created.',
    }
  }

  return {
    ok: true,
    changed: rewriteNumIdForwardAt(ctx, oldNumId, numId, input.blockIndex),
  }
}

/** Context menu: restart numbering (new num + startOverride pointing at the same abstractNum;
 *  lists with no definition (numId=0 / CSS counters) get a new definition, naturally starting at 1) */
export function restartNumbering(ctx: NumberingContext): void {
  if (!ctx.editor) return
  const result = restartNumberingAt(ctx, {
    blockIndex: ctx.editor.state.selection.$from.index(0),
    start: 1,
  })
  if (result.ok) ctx.setStatus(t('appNumberingRestarted'))
}

/** Continue one explicit list from a previous explicit list identity. */
export function continueNumberingAt(
  ctx: NumberingContext,
  input: ContinueNumberingInput,
): NumberingActionResult {
  if (!ctx.editor) {
    return {
      ok: false,
      error: 'execution_failed',
      message: 'DOCX numbering state is unavailable.',
    }
  }
  if (
    !Number.isInteger(input.blockIndex) ||
    !Number.isInteger(input.previousBlockIndex) ||
    input.blockIndex < 0 ||
    input.previousBlockIndex < 0 ||
    input.previousBlockIndex >= input.blockIndex
  ) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: 'docx.list.continue requires an earlier previousBlockIndex.',
    }
  }

  const current = ctx.editor.state.doc.maybeChild(input.blockIndex)
  const previous = ctx.editor.state.doc.maybeChild(input.previousBlockIndex)
  if (
    !current ||
    current.type.name !== 'docListItem' ||
    current.attrs.numId == null ||
    !previous ||
    previous.type.name !== 'docListItem' ||
    previous.attrs.numId == null
  ) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: 'docx.list.continue requires two valid list-item blocks.',
    }
  }

  const currentNumId = String(current.attrs.numId)
  const previousNumId = String(previous.attrs.numId)
  if (currentNumId === previousNumId) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: 'The selected DOCX blocks already share one list identity.',
    }
  }
  return {
    ok: true,
    changed: rewriteNumIdForwardAt(ctx, currentNumId, previousNumId, input.blockIndex),
  }
}

/** Context menu: continue numbering (merge back into the previous list's numId; the count continues) */
export function continueNumbering(ctx: NumberingContext): void {
  if (!ctx.editor) return
  const attrs = ctx.editor.getAttributes('docListItem')
  const curNumId = attrs?.numId as string | null
  if (!curNumId) return
  const { state } = ctx.editor
  const startIdx = state.selection.$from.index(0)
  let previousBlockIndex: number | null = null
  state.doc.forEach((node, _offset, idx) => {
    if (idx >= startIdx) return
    if (node.type.name === 'docListItem' && node.attrs.numId && node.attrs.numId !== curNumId) {
      previousBlockIndex = idx
    }
  })
  if (previousBlockIndex == null) return
  const result = continueNumberingAt(ctx, { blockIndex: startIdx, previousBlockIndex })
  if (result.ok) ctx.setStatus(t('appNumberingContinued'))
}
