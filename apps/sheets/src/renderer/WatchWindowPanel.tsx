// Modified by TandemFolio contributors: bounded watch-set helpers and browser App wiring.
import { useEffect, useState } from 'react'

import { formatAddress } from '../domain/cell-address'
import { useI18n } from './i18n/locale'

/// One watched cell, addressed by stable ids so sheet renames survive.
export interface WatchCell {
  readonly sheetId: string
  readonly row: number
  readonly column: number
}

export interface WatchRowValue {
  readonly sheetName: string
  readonly value: string
  readonly formula: string
}

export function watchKey(cell: WatchCell): string {
  return `${cell.sheetId}:${cell.row}:${cell.column}`
}

export interface WatchSelection {
  readonly startRow: number
  readonly endRow: number
  readonly startColumn: number
  readonly endColumn: number
}

export const MAX_WATCH_CELLS = 20

export function appendWatchSelection(
  watches: readonly WatchCell[],
  sheetId: string,
  selection: WatchSelection,
): readonly WatchCell[] {
  if (watches.length >= MAX_WATCH_CELLS) return watches
  const seen = new Set(watches.map(watchKey))
  const added: WatchCell[] = []
  outer: for (let row = selection.startRow; row <= selection.endRow; row += 1) {
    for (let column = selection.startColumn; column <= selection.endColumn; column += 1) {
      if (watches.length + added.length >= MAX_WATCH_CELLS) break outer
      const cell = { sheetId, row, column }
      if (seen.has(watchKey(cell))) continue
      seen.add(watchKey(cell))
      added.push(cell)
    }
  }
  return added.length === 0 ? watches : [...watches, ...added]
}

export function removeWatch(
  watches: readonly WatchCell[],
  key: string,
): readonly WatchCell[] {
  return watches.filter((cell) => watchKey(cell) !== key)
}

/// Excel's Watch Window: a floating, non-modal panel pinning cells whose
/// value/formula update live. Values are polled while open — recalc has no
/// single completion event the panel could subscribe to from here, and the
/// handful of watched cells make polling effectively free.
export function WatchWindowPanel({
  watches,
  onResolve,
  onAddSelection,
  onRemove,
  onClose,
}: {
  readonly watches: readonly WatchCell[]
  /// null = the sheet is gone; the row renders struck through.
  readonly onResolve: (cell: WatchCell) => WatchRowValue | null
  readonly onAddSelection: () => void
  readonly onRemove: (key: string) => void
  readonly onClose: () => void
}): React.JSX.Element {
  const { t } = useI18n()
  const [, setTick] = useState(0)
  useEffect(() => {
    const timer = setInterval(() => setTick((value) => value + 1), 800)
    return () => clearInterval(timer)
  }, [])
  return (
    <section className="watch-window" role="dialog" aria-label={t('appWatchWindow')}>
      <header>
        <span className="watch-title">{t('appWatchWindow')}</span>
        <button className="watch-add" onClick={onAddSelection}>
          {t('appWatchAdd')}
        </button>
        <button className="watch-close" data-tip={t('appClose')} onClick={onClose}>
          ✕
        </button>
      </header>
      {watches.length === 0 ? (
        <p className="watch-empty">{t('appWatchEmpty')}</p>
      ) : (
        <table className="watch-table">
          <thead>
            <tr>
              <th>{t('appWatchSheet')}</th>
              <th>{t('appWatchCell')}</th>
              <th>{t('appWatchValue')}</th>
              <th>{t('appWatchFormula')}</th>
              <th />
            </tr>
          </thead>
          <tbody>
            {watches.map((cell) => {
              const resolved = onResolve(cell)
              return (
                <tr key={watchKey(cell)} className={resolved ? '' : 'watch-gone'}>
                  <td>{resolved?.sheetName ?? '—'}</td>
                  <td>{formatAddress(cell.row, cell.column)}</td>
                  <td>{resolved?.value ?? ''}</td>
                  <td className="watch-formula">{resolved?.formula ?? ''}</td>
                  <td>
                    <button
                      className="watch-remove"
                      data-tip={t('appDeleteLabel')}
                      onClick={() => onRemove(watchKey(cell))}
                    >
                      ✕
                    </button>
                  </td>
                </tr>
              )
            })}
          </tbody>
        </table>
      )}
    </section>
  )
}

