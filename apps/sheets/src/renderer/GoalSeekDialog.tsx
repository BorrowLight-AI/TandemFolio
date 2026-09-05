import { useState } from 'react'

import type { GoalSeekResult } from './goal-seek'
import { useI18n } from './i18n/locale'

/// Excel's Goal Seek: set a formula cell to a target value by changing one
/// input cell. The solver writes through Univer so the whole run is undoable.
export function GoalSeekDialog({
  initialSetCell,
  onSolve,
  onClose,
}: {
  readonly initialSetCell: string
  readonly onSolve: (setCell: string, toValue: number, byCell: string) => Promise<GoalSeekResult>
  readonly onClose: () => void
}): React.JSX.Element {
  const { t } = useI18n()
  const [setCell, setSetCell] = useState(initialSetCell)
  const [toValue, setToValue] = useState('')
  const [byCell, setByCell] = useState('')
  const [busy, setBusy] = useState(false)
  const [outcome, setOutcome] = useState<string | null>(null)
  const [error, setError] = useState<string | null>(null)
  const chinese = document.documentElement.lang.toLowerCase().startsWith('zh')
  const copy = chinese
    ? {
        title: '单变量求解',
        setCell: '目标单元格',
        toValue: '目标值',
        byCell: '可变单元格',
        note: '目标单元格必须包含引用可变单元格的公式。',
        needNumber: '目标值必须是数字。',
        solving: '求解中…',
        solve: '求解',
        found: (value: string, reached: string) => `已找到解 ${value}；公式结果 ${reached}。`,
        notFound: (value: string, reached: string) =>
          `未能收敛。最后尝试 ${value}；公式结果 ${reached}。`,
      }
    : {
        title: 'Goal Seek',
        setCell: 'Set cell',
        toValue: 'To value',
        byCell: 'By changing cell',
        note: 'The target cell must contain a formula that refers to the changing cell.',
        needNumber: 'The target value must be a number.',
        solving: 'Solving…',
        solve: 'Solve',
        found: (value: string, reached: string) =>
          `Solution found: ${value}; formula result: ${reached}.`,
        notFound: (value: string, reached: string) =>
          `Could not converge. Last value: ${value}; formula result: ${reached}.`,
      }
  const cellRef = /^[A-Za-z]{1,3}\d{1,7}$/
  const valid = cellRef.test(setCell.trim()) && cellRef.test(byCell.trim()) && toValue.trim() !== ''

  const solve = async (): Promise<void> => {
    const target = Number(toValue)
    if (!Number.isFinite(target)) {
      setError(copy.needNumber)
      return
    }
    setBusy(true)
    setError(null)
    setOutcome(null)
    try {
      const result = await onSolve(setCell.trim(), target, byCell.trim())
      const round = (value: number): string =>
        Number.isFinite(value) ? String(Number(value.toFixed(6))) : '—'
      const value = round(result.solution)
      const reached = round(result.reached)
      setOutcome(result.found ? copy.found(value, reached) : copy.notFound(value, reached))
    } catch (failure) {
      setError(failure instanceof Error ? failure.message : String(failure))
    } finally {
      setBusy(false)
    }
  }

  return (
    <div className="dialog-backdrop" onClick={busy ? undefined : onClose}>
      <div
        className="format-cells-dialog"
        role="dialog"
        aria-label={copy.title}
        onClick={(event) => event.stopPropagation()}
      >
        <header>{copy.title}</header>
        <section className="dialog-body">
          <div className="dialog-grid">
            <label>
              {copy.setCell}
              <input
                autoFocus
                value={setCell}
                placeholder="B5"
                disabled={busy}
                onChange={(event) => setSetCell(event.target.value)}
              />
            </label>
            <label>
              {copy.toValue}
              <input
                value={toValue}
                placeholder="0"
                disabled={busy}
                onChange={(event) => setToValue(event.target.value)}
              />
            </label>
            <label>
              {copy.byCell}
              <input
                value={byCell}
                placeholder="B2"
                disabled={busy}
                onChange={(event) => setByCell(event.target.value)}
              />
            </label>
          </div>
          <p className="dialog-note">{copy.note}</p>
          {outcome && <p className="dialog-note">{outcome}</p>}
          {error && (
            <p className="dialog-note dialog-error" role="alert">
              {error}
            </p>
          )}
        </section>
        <div className="dialog-actions">
          <button className="secondary" disabled={busy} onClick={onClose}>
            {t('appCancel')}
          </button>
          <button className="primary-action" disabled={!valid || busy} onClick={() => void solve()}>
            {busy ? copy.solving : copy.solve}
          </button>
        </div>
      </div>
    </div>
  )
}
