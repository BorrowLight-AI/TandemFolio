import type { Editor, JSONContent } from '@tiptap/core'
import type { EditorState, Transaction } from '@tiptap/pm/state'
import { closeHistory } from '@tiptap/pm/history'
import type { ChartDisplay, NewChart } from '@genoffice/docx-engine'

export const DEFAULT_DOCX_CHART_WIDTH_PX = 576
export const DEFAULT_DOCX_CHART_HEIGHT_PX = 336

export interface InsertDocxChartInput extends NewChart {
  readonly afterBlockIndex: number
  readonly title: string
  readonly widthPx: number
  readonly heightPx: number
}

export type InsertDocxChartResult =
  | { readonly ok: true; readonly chartBlockIndex: number }
  | {
      readonly ok: false
      readonly error: 'invalid_arguments' | 'execution_failed'
      readonly message: string
    }

export type DocxChartUpdateField = 'title' | 'categories' | 'series'

export interface UpdateDocxChartInput {
  readonly chartBlockIndex: number
  readonly patch: Readonly<{
    title?: string
    categories?: readonly string[]
    series?: ReadonlyArray<{
      readonly name: string | null
      readonly values: readonly (number | null)[]
    }>
  }>
  readonly fields: readonly DocxChartUpdateField[]
}

export type UpdateDocxChartResult =
  | { readonly ok: true; readonly changed: boolean }
  | {
      readonly ok: false
      readonly error: 'invalid_arguments' | 'execution_failed'
      readonly message: string
    }

type DispatchTransaction = (transaction: Transaction) => void

type PositionedChartInput = Omit<InsertDocxChartInput, 'afterBlockIndex'> & {
  readonly position: number
}

function validateChart(input: Omit<InsertDocxChartInput, 'afterBlockIndex'>): string | null {
  if (!Number.isInteger(input.widthPx) || input.widthPx < 120 || input.widthPx > 660) {
    return 'DOCX chart widthPx must be an integer from 120 through 660.'
  }
  if (!Number.isInteger(input.heightPx) || input.heightPx < 80 || input.heightPx > 4096) {
    return 'DOCX chart heightPx must be an integer from 80 through 4096.'
  }
  if (input.categories.length < 1 || input.categories.length > 256) {
    return 'DOCX charts require 1 through 256 categories.'
  }
  if (input.series.length < 1 || input.series.length > 64) {
    return 'DOCX charts require 1 through 64 series.'
  }
  if (input.kind === 'pie' && input.series.length !== 1) {
    return 'DOCX pie charts require exactly one series.'
  }
  if (input.categories.length * input.series.length > 4096) {
    return 'DOCX chart data is limited to 4096 category-series values.'
  }
  for (let index = 0; index < input.series.length; index += 1) {
    const values = input.series[index].values
    if (values.length !== input.categories.length) {
      return `DOCX chart series ${index} has ${values.length} value(s), but ${input.categories.length} categories were supplied.`
    }
    if (
      values.some(
        (value) => value !== null && (!Number.isFinite(value) || value < -1e12 || value > 1e12),
      )
    ) {
      return `DOCX chart series ${index} contains a non-finite or out-of-range value.`
    }
  }
  return null
}

function invalidUpdate(message: string): UpdateDocxChartResult {
  return { ok: false, error: 'invalid_arguments', message }
}

function boundedChartText(value: unknown): value is string {
  return typeof value === 'string' && Array.from(value).length <= 512
}

/** Shared same-shape chart-content writer for the protected grid and Registry. */
export function updateDocxChartAtPosition(
  state: EditorState,
  dispatch: DispatchTransaction,
  position: number,
  input: Pick<UpdateDocxChartInput, 'patch' | 'fields'>,
): UpdateDocxChartResult {
  const node = state.doc.nodeAt(position)
  const current = node?.attrs.chartDisplay as ChartDisplay | null | undefined
  if (!node || node.type.name !== 'docProtected' || !current) {
    return invalidUpdate('The DOCX chart target is not an editable chart block.')
  }
  if (!Array.isArray(input.fields) || input.fields.length < 1 || input.fields.length > 3) {
    return invalidUpdate('DOCX chart fields must contain one through three masked fields.')
  }
  const fields = new Set<DocxChartUpdateField>()
  for (const field of input.fields) {
    if (field !== 'title' && field !== 'categories' && field !== 'series') {
      return invalidUpdate(`Unsupported DOCX chart field: ${String(field)}.`)
    }
    if (fields.has(field)) return invalidUpdate(`DOCX chart field ${field} must not be repeated.`)
    fields.add(field)
  }
  const patchKeys = Object.keys(input.patch)
  if (
    patchKeys.length !== fields.size ||
    patchKeys.some((key) => !fields.has(key as DocxChartUpdateField))
  ) {
    return invalidUpdate('DOCX chart patch must contain exactly the fields named by fields.')
  }

  const next: ChartDisplay = {
    ...current,
    categories: [...current.categories],
    series: current.series.map((series) => ({ ...series, values: [...series.values] })),
  }
  if (fields.has('title')) {
    if (current.title === undefined) {
      return invalidUpdate('The DOCX chart has no editable title slot.')
    }
    if (!boundedChartText(input.patch.title)) {
      return invalidUpdate('DOCX chart title must contain at most 512 characters.')
    }
    next.title = input.patch.title
  }
  if (fields.has('categories')) {
    const categories = input.patch.categories
    if (!Array.isArray(categories) || categories.length !== current.categories.length) {
      return invalidUpdate(
        `DOCX chart categories must preserve the current ${current.categories.length}-item shape.`,
      )
    }
    if (categories.some((category) => !boundedChartText(category))) {
      return invalidUpdate('DOCX chart category labels must contain at most 512 characters each.')
    }
    next.categories = [...categories]
  }
  if (fields.has('series')) {
    const series = input.patch.series
    if (!Array.isArray(series) || series.length !== current.series.length) {
      return invalidUpdate(
        `DOCX chart series must preserve the current ${current.series.length}-series shape.`,
      )
    }
    if (series.length * current.categories.length > 4096) {
      return invalidUpdate('DOCX chart data is limited to 4096 category-series values.')
    }
    const nextSeries: ChartDisplay['series'] = []
    for (let seriesIndex = 0; seriesIndex < series.length; seriesIndex += 1) {
      const source = series[seriesIndex]
      const prior = current.series[seriesIndex]
      if (source.values.length !== prior.values.length) {
        return invalidUpdate(
          `DOCX chart series ${seriesIndex} must preserve its current ${prior.values.length}-value shape.`,
        )
      }
      if (prior.name === undefined && source.name !== null) {
        return invalidUpdate(`DOCX chart series ${seriesIndex} name is read-only.`)
      }
      if (prior.name !== undefined && !boundedChartText(source.name)) {
        return invalidUpdate(
          `DOCX chart series ${seriesIndex} name must contain at most 512 characters.`,
        )
      }
      const values: (number | null)[] = []
      for (let valueIndex = 0; valueIndex < source.values.length; valueIndex += 1) {
        const value = source.values[valueIndex]
        const priorValue = prior.values[valueIndex]
        if (priorValue === null && value !== null) {
          return invalidUpdate(
            `DOCX chart series ${seriesIndex} cache gap at value ${valueIndex} is read-only.`,
          )
        }
        if (priorValue !== null && value === null) {
          return invalidUpdate(
            `DOCX chart series ${seriesIndex} value ${valueIndex} cannot become a cache gap.`,
          )
        }
        if (value !== null && (!Number.isFinite(value) || value < -1e12 || value > 1e12)) {
          return invalidUpdate(
            `DOCX chart series ${seriesIndex} value ${valueIndex} is non-finite or out of range.`,
          )
        }
        values.push(value)
      }
      nextSeries.push({ ...(source.name === null ? {} : { name: source.name }), values })
    }
    next.series = nextSeries
  }

  const changed = JSON.stringify(next) !== JSON.stringify(current)
  if (changed) {
    dispatch(
      closeHistory(
        state.tr.setNodeMarkup(position, undefined, { ...node.attrs, chartDisplay: next }),
      ),
    )
  }
  return { ok: true, changed }
}

/** Resolve one stable top-level block and update its exact chart content. */
export function updateDocxChartAtBlock(
  editor: Editor,
  input: UpdateDocxChartInput,
): UpdateDocxChartResult {
  const { doc } = editor.state
  if (
    !Number.isInteger(input.chartBlockIndex) ||
    input.chartBlockIndex < 0 ||
    input.chartBlockIndex >= doc.childCount
  ) {
    return invalidUpdate(
      `DOCX chart block ${input.chartBlockIndex} is invalid for ${doc.childCount} block(s).`,
    )
  }
  let position = 0
  for (let index = 0; index < input.chartBlockIndex; index += 1) {
    position += doc.child(index).nodeSize
  }
  return updateDocxChartAtPosition(
    editor.state,
    (transaction) => editor.view.dispatch(transaction),
    position,
    input,
  )
}

function chartContent(
  input: Omit<InsertDocxChartInput, 'afterBlockIndex'>,
  label: string,
): JSONContent {
  const spec: NewChart = {
    kind: input.kind,
    title: input.title,
    categories: [...input.categories],
    series: input.series.map((series) => ({
      name: series.name,
      values: [...series.values],
    })),
  }
  const display: ChartDisplay = {
    partPath: '',
    ...spec,
    widthPx: input.widthPx,
    heightPx: input.heightPx,
  }
  return {
    type: 'docProtected',
    attrs: {
      docxIndex: null,
      blockType: 'chart',
      label,
      genChart: spec,
      chartDisplay: display,
    },
  }
}

/** Shared insertion primitive for retained Chart dialog and Registry producers. */
export function insertDocxChartAtPosition(
  editor: Editor,
  input: PositionedChartInput,
  label: string,
): { readonly ok: true } | Exclude<InsertDocxChartResult, { readonly ok: true }> {
  const invalid = validateChart(input)
  if (invalid) return { ok: false, error: 'invalid_arguments', message: invalid }
  const inserted = editor
    .chain()
    .focus()
    .insertContentAt(input.position, chartContent(input, label))
    .run()
  return inserted
    ? { ok: true }
    : {
        ok: false,
        error: 'execution_failed',
        message: 'The mounted DOCX editor rejected chart insertion.',
      }
}

/** Insert a chart after one stable revision-scoped top-level block. */
export function insertDocxChartAfterBlock(
  editor: Editor,
  input: InsertDocxChartInput,
): InsertDocxChartResult {
  const { doc } = editor.state
  if (input.afterBlockIndex < -1 || input.afterBlockIndex >= doc.childCount) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message: `DOCX chart boundary after block ${input.afterBlockIndex} is invalid for ${doc.childCount} block(s).`,
    }
  }
  let position = 0
  for (let index = 0; index <= input.afterBlockIndex; index += 1) {
    position += doc.child(index).nodeSize
  }
  const inserted = insertDocxChartAtPosition(editor, { ...input, position }, 'Chart')
  return inserted.ok ? { ok: true, chartBlockIndex: input.afterBlockIndex + 1 } : inserted
}
