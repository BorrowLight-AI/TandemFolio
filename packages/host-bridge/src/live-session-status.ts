/** Host transport status only; never owns document content or editor history. */
export function createLiveSessionStatus(retry: () => void, activate: () => void = retry) {
  const curtain = document.createElement('div')
  curtain.dataset.liveSessionBlocker = ''
  curtain.setAttribute('role', 'status')
  Object.assign(curtain.style, {
    position: 'fixed',
    inset: '0',
    zIndex: '2147483645',
    background: 'Canvas',
    color: 'CanvasText',
    display: 'grid',
    placeItems: 'center',
    padding: '24px',
    font: '14px system-ui',
    textAlign: 'center',
  })
  curtain.hidden = true
  document.body.append(curtain)
  const blockKeys = (event: Event) => {
    if (!curtain.hidden && !(event.target instanceof Node && root.contains(event.target))) {
      event.preventDefault()
      event.stopImmediatePropagation()
    }
  }
  document.addEventListener('keydown', blockKeys, true)
  const root = document.createElement('aside')
  root.dataset.liveSessionStatus = ''
  root.setAttribute('aria-label', '文件与连接状态')
  Object.assign(root.style, {
    position: 'fixed',
    right: '12px',
    bottom: '36px',
    zIndex: '2147483646',
    maxWidth: 'calc(100vw - 24px)',
    padding: '8px 12px',
    border: '1px solid #d1d5db',
    borderRadius: '8px',
    background: 'Canvas',
    color: 'CanvasText',
    boxShadow: '0 2px 10px #0002',
    font: '13px system-ui',
  })
  const error = document.createElement('div')
  error.setAttribute('role', 'status')
  const retryButton = document.createElement('button')
  retryButton.type = 'button'
  retryButton.textContent = '重试连接'
  retryButton.onclick = retry
  const activateButton = document.createElement('button')
  activateButton.type = 'button'
  activateButton.textContent = '在此继续编辑'
  activateButton.onclick = activate
  activateButton.hidden = true
  activateButton.style.marginRight = '8px'
  const details = document.createElement('details')
  const summary = document.createElement('summary')
  summary.textContent = '文件位置 · 正在确认'
  summary.style.cursor = 'pointer'
  const path = document.createElement('input')
  path.readOnly = true
  path.setAttribute('aria-label', '文件绝对路径')
  Object.assign(path.style, {
    width: 'min(480px, calc(100vw - 64px))',
    display: 'block',
    margin: '8px 0',
    padding: '6px',
  })
  const copy = document.createElement('button')
  copy.type = 'button'
  copy.textContent = '复制路径'
  copy.setAttribute('aria-label', '复制文件路径')
  copy.onclick = async () => {
    try {
      if (!navigator.clipboard?.writeText) throw new Error('Clipboard unavailable')
      await navigator.clipboard.writeText(path.value)
      copy.textContent = '已复制'
    } catch {
      path.select()
      copy.textContent = '请按 ⌘C / Ctrl+C 复制'
    }
  }
  details.append(summary, path, copy)
  root.append(error, activateButton, retryButton, details)
  document.body.append(root)
  const sync = () => {
    path.style.display = path.value ? 'block' : 'none'
    copy.hidden = !path.value
    error.hidden = !error.textContent
  }
  retryButton.hidden = true
  sync()
  return {
    activation(available: boolean) {
      activateButton.hidden = !available
    },
    actionsPending(pending: boolean) {
      root.setAttribute('aria-busy', String(pending))
      activateButton.disabled = pending
      retryButton.disabled = pending
    },
    block(message: string | null) {
      curtain.hidden = !message
      curtain.style.display = message ? 'grid' : 'none'
      curtain.textContent = message
      const editor = document.getElementById('root')
      if (editor) editor.inert = Boolean(message)
    },
    path(value: string | null, exported = false) {
      path.value = value ?? ''
      path.title = path.value
      summary.textContent = exported ? '导出副本位置' : value ? '文件位置' : '文件位置 · 尚未保存'
      copy.textContent = '复制路径'
      sync()
    },
    error(message: string | null, canRetry = false) {
      error.textContent = message
      retryButton.hidden = !message || !canRetry
      sync()
    },
    dispose() {
      document.removeEventListener('keydown', blockKeys, true)
      const editor = document.getElementById('root')
      if (editor) editor.inert = false
      curtain.remove()
      root.remove()
    },
  }
}
