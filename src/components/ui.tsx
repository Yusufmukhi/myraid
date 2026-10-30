import { createContext, ReactNode, useCallback, useContext, useEffect, useRef, useState } from 'react'

/* ----------------------------------------------------------------- icons */
const P = (d: string) => (
  <svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round" aria-hidden>
    <path d={d} />
  </svg>
)
export const Icon = {
  edit: P('M4 20h4L19 9l-4-4L4 16v4zM13.5 6.5l4 4'),
  trash: P('M4 7h16M10 11v6M14 11v6M6 7l1 12a2 2 0 002 2h6a2 2 0 002-2l1-12M9 7V4h6v3'),
  plus: P('M12 5v14M5 12h14'),
  save: P('M5 4h11l3 3v13H5zM8 4v5h7V4M8 20v-6h8v6'),
  copy: P('M9 9h10v11H9zM5 15V4h10'),
  check: P('M5 12.5l4.5 4.5L19 7.5'),
  logout: P('M15 4h4a1 1 0 011 1v14a1 1 0 01-1 1h-4M10 8l-4 4 4 4M6 12h10'),
  users: P('M16 19v-1.5a3.5 3.5 0 00-3.5-3.5h-5A3.5 3.5 0 004 17.5V19M10 11a3 3 0 100-6 3 3 0 000 6zM20 19v-1.5a3.5 3.5 0 00-2.5-3.35M15.5 5.2a3 3 0 010 5.6'),
  chevron: P('M7 10l5 5 5-5'),
  chevronDown: P('M6 9l6 6 6-6'),
  chevronRight: P('M9 18l6-6-6-6'),
  chevronLeft: P('M15 18l-6-6 6-6'),
  menu: P('M4 7h16M4 12h16M4 17h16'),
  close: P('M6 6l12 12M18 6L6 18'),
  calc: P('M6 3h12a1 1 0 011 1v16a1 1 0 01-1 1H6a1 1 0 01-1-1V4a1 1 0 011-1zM8 7h8M8 12h2M14 12h2M8 16h2M14 16h2'),
  grid: P('M4 5h16v14H4zM4 10h16M4 15h16M10 5v14M15 5v14'),
  reset: P('M4 12a8 8 0 108-8M4 5v4h4'),
  key: P('M14 10a4 4 0 11-3.2 6.4L4 17v3h3v-2h2v-2h2l1.2-1.2A4 4 0 0114 10zM16 11.5h.01'),
  building: P('M5 21V4h9v17M14 9h5v12M8 8h3M8 12h3M8 16h3M3 21h18'),
  dashboard: P('M3 3h7v7H3zM14 3h7v7h-7zM14 14h7v7h-7zM3 14h7v7H3z'),
  folder: P('M3 7v13a1 1 0 001 1h16a1 1 0 001-1V9a1 1 0 00-1-1h-7.5l-2-2H4a1 1 0 00-1 1z'),
  fileText: P('M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8zM14 2v6h6M16 13H8M16 17H8M10 9H8'),
  settings: P('M12 15a3 3 0 100-6 3 3 0 000 6zM19.4 15a1.65 1.65 0 00.33 1.82l.06.06a2 2 0 010 2.83 2 2 0 01-2.83 0l-.06-.06a1.65 1.65 0 00-1.82-.33 1.65 1.65 0 00-1 1.51V21a2 2 0 01-2 2 2 2 0 01-2-2v-.09A1.65 1.65 0 009 19.4a1.65 1.65 0 00-1.82.33l-.06.06a2 2 0 01-2.83 0 2 2 0 010-2.83l.06-.06a1.65 1.65 0 00.33-1.82 1.65 1.65 0 00-1.51-1H3a2 2 0 01-2-2 2 2 0 012-2h.09A1.65 1.65 0 004.6 9a1.65 1.65 0 00-.33-1.82l-.06-.06a2 2 0 010-2.83 2 2 0 012.83 0l.06.06a1.65 1.65 0 001.82.33H9a1.65 1.65 0 001-1.51V3a2 2 0 012-2 2 2 0 012 2v.09a1.65 1.65 0 001 1.51 1.65 1.65 0 001.82-.33l.06-.06a2 2 0 012.83 0 2 2 0 010 2.83l-.06.06a1.65 1.65 0 00-.33 1.82V9a1.65 1.65 0 001.51 1H21a2 2 0 012 2 2 2 0 01-2 2h-.09a1.65 1.65 0 00-1.51 1z'),
  history: P('M12 8v4l3 3M3.05 11a9 9 0 11.5 4m-.5-4v-4h-4'),
  layers: P('M12 2L2 7l10 5 10-5-10-5zM2 17l10 5 10-5M2 12l10 5 10-5'),
  refresh: P('M21 12a9 9 0 11-2.6-6.4L21 8M21 3v5h-5'),
  search: P('M11 19a8 8 0 100-16 8 8 0 000 16zM21 21l-4.35-4.35'),
  archive: P('M21 8v13H3V8M1 3h22v5H1zM10 12h4'),
  arrowRight: P('M5 12h14M12 5l7 7-7 7'),
  info: P('M12 16v-4M12 8h.01M22 12A10 10 0 112 12a10 10 0 0120 0z'),
  diff: P('M16 3h5v5M4 20L21 3M21 16v5h-5M15 15l6 6M4 4l5 5'),
  shield: P('M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z'),
  lock: P('M12 15v2m-6 4h12a2 2 0 002-2v-6a2 2 0 00-2-2H6a2 2 0 00-2 2v6a2 2 0 002 2zm10-10V7a4 4 0 00-8 0v4h8z'),
  clock: P('M12 8v4l3 3m6-3a9 9 0 11-18 0 9 9 0 0118 0z'),
}

export function Spinner({ size = 16 }: { size?: number }) {
  return <span className="spinner" style={{ width: size, height: size }} aria-label="Loading" />
}

/* ------------------------------------------------------ toasts + dialogs */
interface Toast {
  id: number
  kind: 'ok' | 'err'
  text: string
}
interface ConfirmOpts {
  title: string
  body?: ReactNode
  confirmLabel?: string
  danger?: boolean
}
interface Ctx {
  toast: (text: string, kind?: 'ok' | 'err') => void
  confirm: (o: ConfirmOpts) => Promise<boolean>
}
const UICtx = createContext<Ctx>({ toast: () => {}, confirm: async () => false })
export const useUI = () => useContext(UICtx)

export function UIProvider({ children }: { children: ReactNode }) {
  const [toasts, setToasts] = useState<Toast[]>([])
  const [dlg, setDlg] = useState<(ConfirmOpts & { resolve: (v: boolean) => void }) | null>(null)
  const seq = useRef(0)

  const toast = useCallback((text: string, kind: 'ok' | 'err' = 'ok') => {
    const id = ++seq.current
    setToasts((t) => [...t, { id, kind, text }])
    setTimeout(() => setToasts((t) => t.filter((x) => x.id !== id)), kind === 'err' ? 6000 : 3000)
  }, [])

  const confirm = useCallback(
    (o: ConfirmOpts) => new Promise<boolean>((resolve) => setDlg({ ...o, resolve })),
    [],
  )
  const close = (v: boolean) => {
    dlg?.resolve(v)
    setDlg(null)
  }

  useEffect(() => {
    if (!dlg) return
    const h = (e: KeyboardEvent) => e.key === 'Escape' && close(false)
    window.addEventListener('keydown', h)
    return () => window.removeEventListener('keydown', h)
  })

  return (
    <UICtx.Provider value={{ toast, confirm }}>
      {children}
      <div className="toasts" role="status" aria-live="polite">
        {toasts.map((t) => (
          <div key={t.id} className={`toast ${t.kind}`}>
            {t.kind === 'ok' ? Icon.check : Icon.close}
            <span>{t.text}</span>
          </div>
        ))}
      </div>
      {dlg && (
        <div className="overlay" onMouseDown={() => close(false)}>
          <div className="dialog" role="dialog" aria-modal onMouseDown={(e) => e.stopPropagation()}>
            <h3>{dlg.title}</h3>
            {dlg.body && <p>{dlg.body}</p>}
            <div className="dialog-actions">
              <button className="btn ghost" onClick={() => close(false)}>
                Cancel
              </button>
              <button className={`btn ${dlg.danger ? 'danger' : 'primary'}`} autoFocus onClick={() => close(true)}>
                {dlg.confirmLabel ?? 'Confirm'}
              </button>
            </div>
          </div>
        </div>
      )}
    </UICtx.Provider>
  )
}

/* ----------------------------------------------------------- number input */
export function NumField({
  value,
  onChange,
  step,
  min,
  placeholder,
  className,
}: {
  value: number | null
  onChange: (v: number) => void
  step?: number
  min?: number
  placeholder?: string
  className?: string
}) {
  const shown = value === null ? '' : String(value)
  const [text, setText] = useState(shown)
  const focused = useRef(false)
  useEffect(() => {
    if (!focused.current) setText(shown)
  }, [shown])
  return (
    <input
      className={className}
      type="number"
      inputMode="decimal"
      value={text}
      step={step}
      min={min}
      placeholder={placeholder}
      onFocus={(e) => {
        focused.current = true
        e.target.select()
      }}
      onBlur={() => {
        focused.current = false
        const val = text.trim()
        if (val === '') {
          const fallback = min !== undefined ? min : 0
          setText(String(fallback))
          onChange(fallback)
        } else {
          const n = parseFloat(val)
          if (!Number.isFinite(n)) {
            const fallback = min !== undefined ? min : 0
            setText(String(fallback))
            onChange(fallback)
          } else {
            let clamped = n
            if (min !== undefined && clamped < min) clamped = min
            setText(String(clamped))
            onChange(clamped)
          }
        }
      }}
      onChange={(e) => {
        setText(e.target.value)
        if (e.target.value.trim() === '') return
        const n = parseFloat(e.target.value)
        if (Number.isFinite(n)) {
          let clamped = n
          if (min !== undefined && clamped < min) clamped = min
          onChange(clamped)
        }
      }}
    />
  )
}

export function useCopy() {
  const { toast } = useUI()
  return async (text: string, what = 'Copied') => {
    try {
      await navigator.clipboard.writeText(text)
      toast(what)
    } catch {
      toast('Could not copy – select and copy manually', 'err')
    }
  }
}
