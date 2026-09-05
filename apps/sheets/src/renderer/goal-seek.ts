/**
 * What-If Analysis › Goal Seek: solve one input cell so a formula cell hits a
 * target, Excel-style (secant method, Excel's default limits: 100 iterations,
 * |f - target| ≤ 0.001). Trial guesses are excluded from the edit journal;
 * only the accepted value enters native history, so one Undo restores the
 * original changing cell. Manual calculation mode is supported too.
 */
import { parseAddress } from '../domain/cell-address'
import { calculateNow, isManualCalculation } from './calc-options'
import { t } from './i18n/locale'
import { journalSuppression, type UniverRuntime } from './univer-state'

export interface GoalSeekResult {
  readonly found: boolean
  /// f(x) actually reached (shown next to the target).
  readonly reached: number
  readonly solution: number
  readonly iterations: number
}

const MAX_ITERATIONS = 100
const MAX_CHANGE = 0.001

/// Runs `mutate` and resolves after the recalc it triggers. The listener is
/// attached before the write so the end event cannot be missed; the timeout
/// only covers edits that trigger no recalculation at all.
function mutateAndAwaitRecalc(
  runtime: UniverRuntime,
  mutate: () => void,
  suppressJournal: boolean,
): Promise<void> {
  const formula = runtime.univerAPI.getFormula()
  return new Promise((resolve) => {
    let settled = false
    let timer: ReturnType<typeof setTimeout> | null = null
    const finish = (): void => {
      if (settled) return
      settled = true
      disposable.dispose()
      if (timer) clearTimeout(timer)
      resolve()
    }
    const disposable = formula.calculationEnd(finish)
    timer = setTimeout(finish, 2_000)
    const previousSuppression = journalSuppression.active
    if (suppressJournal) journalSuppression.active = true
    try {
      mutate()
    } finally {
      journalSuppression.active = previousSuppression
    }
    if (isManualCalculation(runtime)) calculateNow(runtime)
  })
}

export async function solveGoalSeek(
  runtime: UniverRuntime,
  input: { readonly setCell: string; readonly toValue: number; readonly byCell: string },
): Promise<GoalSeekResult> {
  const worksheet = runtime.univerAPI.getActiveWorkbook()?.getActiveSheet()
  if (!worksheet) throw new Error(t('appNoActiveSheet'))
  const target = parseAddress(input.setCell.toUpperCase())
  const changing = parseAddress(input.byCell.toUpperCase())
  const targetRange = worksheet.getRange(target.row, target.column, 1, 1)
  const changingRange = worksheet.getRange(changing.row, changing.column, 1, 1)
  if (!targetRange.getFormula()) throw new Error('Goal Seek requires a formula in the target cell.')
  // Excel's constraint too: the changing cell must hold a plain number (or
  // nothing) — a formula or text would be destroyed by the solver's writes.
  const originalRaw = changingRange.getValue()
  if (changingRange.getFormula() || (originalRaw != null && typeof originalRaw !== 'number')) {
    throw new Error('The changing cell must be empty or contain a number.')
  }
  const original = typeof originalRaw === 'number' ? originalRaw : 0
  const write = (value: number | null, suppressJournal: boolean): Promise<void> =>
    mutateAndAwaitRecalc(
      runtime,
      () => (value === null ? changingRange.setValue({ v: null }) : changingRange.setValue(value)),
      suppressJournal,
    )
  const restore = (): Promise<void> => write(originalRaw == null ? null : original, true)
  const commit = async (solution: number): Promise<void> => {
    // Rewind the unjournaled trials first, then apply the accepted solution as
    // the single normal Univer mutation recorded by save and native history.
    await restore()
    await write(solution, false)
  }

  /// NaN when the formula errors for this guess (#DIV/0! mid-solve is
  /// normal); the secant guards treat it as a dead end, not a crash.
  const evaluate = async (x: number): Promise<number> => {
    await write(x, true)
    const value = targetRange.getValue()
    return typeof value === 'number' ? value : Number.NaN
  }

  try {
    let x0 = original
    let f0 = await evaluate(x0)
    if (Number.isNaN(f0)) throw new Error('The target formula must return a number.')
    if (Math.abs(f0 - input.toValue) <= MAX_CHANGE) {
      return { found: true, reached: f0, solution: x0, iterations: 1 }
    }
    // A relative nudge for the second sample; 1 covers the x0 = 0 case.
    let x1 = x0 + Math.max(Math.abs(x0) * 0.01, 1)
    let f1 = await evaluate(x1)
    let iterations = 2
    while (iterations < MAX_ITERATIONS) {
      if (Math.abs(f1 - input.toValue) <= MAX_CHANGE) {
        await commit(x1)
        return { found: true, reached: f1, solution: x1, iterations }
      }
      const slope = f1 - f0
      if (slope === 0 || !Number.isFinite(slope)) break
      const next = x1 - ((f1 - input.toValue) * (x1 - x0)) / slope
      if (!Number.isFinite(next)) break
      x0 = x1
      f0 = f1
      x1 = next
      f1 = await evaluate(x1)
      iterations += 1
    }
    if (Math.abs(f1 - input.toValue) <= MAX_CHANGE) {
      await commit(x1)
      return { found: true, reached: f1, solution: x1, iterations }
    }
    // No convergence: put the original contents back (journaled, undoable).
    await restore()
    return { found: false, reached: f1, solution: x1, iterations }
  } catch (error) {
    // Any abort after the first write must not leave a guess in the grid.
    await restore()
    throw error
  }
}
