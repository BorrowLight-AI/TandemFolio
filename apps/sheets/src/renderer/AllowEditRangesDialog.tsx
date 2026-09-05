import { useState } from 'react'

import { useI18n } from './i18n/locale'

export interface AllowEditRange {
  readonly name: string
  readonly sqref: string
}

const RANGE_REF = /^\$?[A-Za-z]{1,3}\$?\d{1,7}(?::\$?[A-Za-z]{1,3}\$?\d{1,7})?$/

export function AllowEditRangesDialog({
  ranges,
  defaultRef,
  onApply,
  onClose,
}: {
  readonly ranges: readonly AllowEditRange[]
  readonly defaultRef: string
  readonly onApply: (ranges: readonly AllowEditRange[]) => string | null
  readonly onClose: () => void
}): React.JSX.Element {
  const { t } = useI18n()
  const [working, setWorking] = useState<readonly AllowEditRange[]>(ranges)
  const [name, setName] = useState(`Range${ranges.length + 1}`)
  const [reference, setReference] = useState(defaultRef)
  const [error, setError] = useState<string | null>(null)
  const chinese = document.documentElement.lang.toLowerCase().startsWith('zh')
  const copy = chinese
    ? { title: '允许用户编辑区域', hint: '这些区域在工作表受保护时仍可编辑。', name: '名称', ref: '引用位置', add: '添加', remove: '删除', duplicate: '名称必须唯一。', invalid: '请输入有效的单元格或区域引用。' }
    : { title: 'Allow Edit Ranges', hint: 'These ranges remain editable when the sheet is protected.', name: 'Name', ref: 'Refers to cells', add: 'Add', remove: 'Delete', duplicate: 'The name must be unique.', invalid: 'Enter a valid cell or range reference.' }

  const add = (): void => {
    const trimmedName = name.trim()
    const parts = reference.trim().split(/\s+/).filter(Boolean)
    if (!trimmedName || working.some((range) => range.name.toLowerCase() === trimmedName.toLowerCase())) {
      setError(copy.duplicate)
      return
    }
    if (parts.length === 0 || parts.some((part) => !RANGE_REF.test(part))) {
      setError(copy.invalid)
      return
    }
    const next = [
      ...working,
      { name: trimmedName, sqref: parts.join(' ').toUpperCase().replaceAll('$', '') },
    ]
    setWorking(next)
    setName(`Range${next.length + 1}`)
    setReference('')
    setError(null)
  }

  return (
    <div className="dialog-backdrop" onClick={onClose}>
      <div className="format-cells-dialog" role="dialog" aria-label={copy.title} onClick={(event) => event.stopPropagation()}>
        <header>{copy.title}</header>
        <section className="dialog-body">
          <p className="dialog-note">{copy.hint}</p>
          {working.length > 0 && (
            <ul className="allow-edit-list">
              {working.map((range) => (
                <li key={range.name}>
                  <span className="allow-edit-name">{range.name}</span>
                  <span className="allow-edit-ref">{range.sqref}</span>
                  <button className="secondary" onClick={() => setWorking(working.filter((entry) => entry.name !== range.name))}>
                    {copy.remove}
                  </button>
                </li>
              ))}
            </ul>
          )}
          <div className="dialog-grid">
            <label>{copy.name}<input value={name} onChange={(event) => setName(event.target.value)} /></label>
            <label>{copy.ref}<input autoFocus placeholder="A1:B4" value={reference} onChange={(event) => setReference(event.target.value)} onKeyDown={(event) => { if (event.key === 'Enter') add() }} /></label>
          </div>
          <div className="dialog-actions-inline"><button className="secondary" onClick={add}>{copy.add}</button></div>
          {error && <p className="dialog-note dialog-error" role="alert">{error}</p>}
        </section>
        <div className="dialog-actions">
          <button className="secondary" onClick={onClose}>{t('appCancel')}</button>
          <button className="primary-action" onClick={() => { const failed = onApply(working); if (failed) setError(failed); else onClose() }}>{t('appOk')}</button>
        </div>
      </div>
    </div>
  )
}
