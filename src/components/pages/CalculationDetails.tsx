import { useEffect, useMemo, useRef, useState } from 'react'
import { addr, CellStyle, ErrVal, Overrides, parseAddr, Workbook } from '../../lib/engine'
import { fmtCell } from '../../lib/format'
import { EXTRA_EDITABLE } from '../../lib/inputs'
import { filterCutToCutOverrides } from '../../lib/model'
import { api, Org, Session } from '../../lib/api'
import { Icon, NumField, useUI } from '../ui'

interface Props {
  wb: Workbook
  c2cWb?: Workbook
  overrides?: Overrides
  formulaOverrides?: Record<string, string>
  c2cOverrides?: Overrides
  c2cFormulaOverrides?: Record<string, string>
  setInput: (key: string, value: number | null | undefined, targetWorkbook?: 'actual' | 'c2c') => void
  setFormulaOverride: (key: string, formula: string | undefined, targetWorkbook?: 'actual' | 'c2c') => void
  isAdminOrOwner?: boolean
  session?: Session
  org?: Org
  onSave?: () => void
  hasUnsavedChanges?: boolean
  isSavingProject?: boolean
}

interface Span {
  rs: number
  cs: number
}

function textOn(hex: string): string {
  const [r, g, b] = [0, 2, 4].map((i) => parseInt(hex.slice(i, i + 2), 16) / 255)
  const lum = 0.2126 * r + 0.7152 * g + 0.0722 * b
  return lum > 0.45 ? '#0f172a' : '#ffffff'
}

const DEFAULT_METADATA: Record<string, { desc: string; category: string }> = {
  'Total cal': { desc: 'Summary & Quotation Model', category: 'Summary' },
}

export function CalculationDetails({
  wb,
  c2cWb,
  overrides = {},
  formulaOverrides = {},
  c2cOverrides = {},
  c2cFormulaOverrides = {},
  setInput,
  setFormulaOverride,
  isAdminOrOwner = true,
  session,
  org,
  onSave,
  hasUnsavedChanges,
  isSavingProject,
}: Props) {
  const { toast } = useUI()

  if (!isAdminOrOwner) {
    return (
      <div className="card p-6 text-center" style={{ margin: '32px auto', maxWidth: 600 }}>
        <div style={{ fontSize: '48px', marginBottom: '12px' }}>🔒</div>
        <h3 style={{ fontSize: '18px', fontWeight: 600 }}>Proprietary Worksheet Access Restricted</h3>
        <p className="muted mt-2" style={{ lineHeight: 1.5 }}>
          Engineering worksheets and detailed formulas are restricted to the Organisation Owner to protect confidential calculation IP. Commercial quote summaries and customer documents are available in the quotation viewer.
        </p>
      </div>
    )
  }

  const [workbookType, setWorkbookType] = useState<'actual' | 'c2c'>('actual')
  const [serverWb, setServerWb] = useState<{ actual: Workbook; c2c: Workbook } | null>(null)
  const [serverMeta, setServerMeta] = useState<Record<string, { desc: string; category: string }>>({})
  const [loadingSheets, setLoadingSheets] = useState(false)
  const [fetchError, setFetchError] = useState<string | null>(null)
  const [renderVer, setRenderVer] = useState(0)

  // Fetch full 15 worksheets securely from server on demand for Owner/Admin
  useEffect(() => {
    if (!isAdminOrOwner) return
    let active = true
    setLoadingSheets(true)
    setFetchError(null)

    const effectiveToken = session?.token || 'authenticated-owner'
    const effectiveOrgId = org?.id || 'default-org'

    api.getFullWorksheets(effectiveToken, effectiveOrgId)
      .then((res) => {
        if (!active) return
        const act = new Workbook(res.actual_workbook, res.named_constants)
        const c2c = new Workbook(res.c2c_workbook, res.named_constants)
        act.setOverrides(overrides)
        act.setFormulaOverrides(formulaOverrides)
        c2c.setOverrides({ ...filterCutToCutOverrides(overrides), ...c2cOverrides })
        c2c.setFormulaOverrides({ ...formulaOverrides, ...c2cFormulaOverrides })
        setServerWb({ actual: act, c2c })
        if (res.metadata) setServerMeta(res.metadata)
      })
      .catch((err) => {
        if (!active) return
        console.warn('Could not load full server worksheets:', err)
        setFetchError(err.message || 'Failed to fetch worksheets')
      })
      .finally(() => {
        if (active) setLoadingSheets(false)
      })
    return () => {
      active = false
    }
  }, [isAdminOrOwner, session?.token, org?.id])

  // Synchronize overrides into active server workbooks
  useEffect(() => {
    if (serverWb) {
      serverWb.actual.setOverrides(overrides)
      serverWb.actual.setFormulaOverrides(formulaOverrides)
      serverWb.c2c.setOverrides({ ...filterCutToCutOverrides(overrides), ...c2cOverrides })
      serverWb.c2c.setFormulaOverrides({ ...formulaOverrides, ...c2cFormulaOverrides })
      setRenderVer((v) => v + 1)
    }
  }, [serverWb, overrides, formulaOverrides, c2cOverrides, c2cFormulaOverrides])

  const activeOverrides = workbookType === 'c2c'
    ? { ...filterCutToCutOverrides(overrides), ...c2cOverrides }
    : overrides

  const activeFormulaOverrides = workbookType === 'c2c'
    ? { ...formulaOverrides, ...c2cFormulaOverrides }
    : formulaOverrides

  const activeWb: Workbook = serverWb
    ? (workbookType === 'c2c' ? serverWb.c2c : serverWb.actual)
    : (workbookType === 'c2c' && c2cWb ? c2cWb : wb)

  const [name, setName] = useState('Total cal')
  const [sel, setSel] = useState<{ r: number; c: number } | null>(null)
  const [editFormulaMode, setEditFormulaMode] = useState(false)
  const [formulaInput, setFormulaInput] = useState('')
  const [inlineKey, setInlineKey] = useState<string | null>(null)
  const [inlineVal, setInlineVal] = useState<string>('')

  const sheet = activeWb.sheets.get(name) || activeWb.sheets.values().next().value!
  const tabsRef = useRef<HTMLDivElement>(null)

  // Merged cells
  const { spans, covered } = useMemo(() => {
    const spans = new Map<string, Span>()
    const covered = new Set<string>()
    if (!sheet) return { spans, covered }
    for (const m of sheet.merges ?? []) {
      const [a, b] = m.split(':')
      if (!b) continue
      const A = parseAddr(a)
      const B = parseAddr(b)
      spans.set(`${A.row},${A.col}`, { rs: B.row - A.row + 1, cs: B.col - A.col + 1 })
      for (let r = A.row; r <= B.row; r++)
        for (let c = A.col; c <= B.col; c++) if (r !== A.row || c !== A.col) covered.add(`${r},${c}`)
    }
    return { spans, covered }
  }, [sheet])

  // Column widths from content
  const widths = useMemo(() => {
    const w: number[] = []
    if (!sheet) return w
    for (let c = 0; c < sheet.cols; c++) {
      let max = 3
      for (let r = 0; r < sheet.rows; r++) {
        if (spans.has(`${r},${c}`) || covered.has(`${r},${c}`)) continue
        const raw = sheet.cells[r][c]
        if (raw === null) continue
        const v = activeWb.value(name, r, c)
        const len = fmtCell(v).length
        if (len > max) max = len
      }
      w.push(Math.max(46, Math.min(240, max * 7.4 + 16)))
    }
    return w
  }, [sheet, spans, covered, name, activeWb, renderVer])

  useEffect(() => {
    setSel(null)
    setEditFormulaMode(false)
    setInlineKey(null)
  }, [name, workbookType])

  const selKey = sel ? Workbook.key(name, sel.r, sel.c) : ''
  const selVal = sel ? activeWb.value(name, sel.r, sel.c) : null
  const selRaw = sel ? activeWb.raw(name, sel.r, sel.c) : null
  const isFormula = sel ? typeof selRaw === 'string' && selRaw.startsWith('=') : false
  const isInput = sel ? (activeWb.isInput(name, sel.r, sel.c) || EXTRA_EDITABLE.has(selKey) || (!isFormula && typeof selVal === 'number')) : false
  const isOverriddenInput = sel ? selKey in activeOverrides : false
  const isOverriddenFormula = sel ? selKey in activeFormulaOverrides : false

  const handleStartFormulaEdit = () => {
    if (!sel) return
    setFormulaInput(typeof selRaw === 'string' ? selRaw : '')
    setEditFormulaMode(true)
  }

  const handleSaveFormula = () => {
    if (!sel) return
    const f = formulaInput.trim()
    if (!f.startsWith('=')) {
      toast('Formula must start with =', 'err')
      return
    }
    setFormulaOverride(selKey, f, workbookType)
    setEditFormulaMode(false)
    toast(workbookType === 'c2c' ? 'Cut-to-Cut formula override applied' : 'Project-specific formula override applied')
  }

  const handleResetFormula = () => {
    if (!sel) return
    setFormulaOverride(selKey, undefined, workbookType)
    setEditFormulaMode(false)
    toast('Reverted to template default formula')
  }

  const currentSheetMeta = serverMeta[name] || DEFAULT_METADATA[name] || { desc: name, category: 'Sheet' }

  return (
    <div className="calculation-details-page">
      {/* Streamlined Header */}
      <div className="card row between p-2 px-3 mb-2" style={{ alignItems: 'center' }}>
        <div className="row gap" style={{ alignItems: 'center', flexWrap: 'wrap' }}>
          <span className="chip primary">{currentSheetMeta.category}</span>
          <strong style={{ fontSize: '16px' }}>{name}</strong>
          <span className="tiny muted">{currentSheetMeta.desc}</span>
          {loadingSheets && (
            <span className="tiny muted" style={{ display: 'inline-flex', alignItems: 'center', gap: 4, background: '#f8fafc', padding: '2px 8px', borderRadius: 4, border: '1px solid #e2e8f0' }}>
              ⏳ Syncing all 15 sheets from secure server...
            </span>
          )}
          {fetchError && !serverWb && (
            <span className="tiny warn" style={{ fontSize: '11px' }}>
              (Server sync notice: {fetchError})
            </span>
          )}
          {(c2cWb || serverWb) && (
            <div style={{ marginLeft: 12, display: 'inline-flex', gap: 4, background: '#f1f5f9', padding: '2px 4px', borderRadius: 6, border: '1px solid #cbd5e1' }}>
              <button
                type="button"
                className={`btn xs ${workbookType === 'actual' ? 'primary' : 'ghost'}`}
                onClick={() => setWorkbookType('actual')}
                style={workbookType === 'actual' ? { background: '#0f172a', color: '#fff', fontSize: '11px', fontWeight: 700 } : { fontSize: '11px' }}
                title="Actual Quotation Sheets (₹90 Steel Rate)"
              >
                📊 Actual Quotation
              </button>
              <button
                type="button"
                className={`btn xs ${workbookType === 'c2c' ? 'primary' : 'ghost'}`}
                onClick={() => setWorkbookType('c2c')}
                style={workbookType === 'c2c' ? { background: '#7c3aed', color: '#fff', fontSize: '11px', fontWeight: 700 } : { fontSize: '11px', color: '#7c3aed' }}
                title="Cut-to-Cut Cost Sheets (₹79 Steel Rate)"
              >
                ✂️ Cut-to-Cut Sheet
              </button>
            </div>
          )}
          {workbookType === 'c2c' && (
            <span className="chip purple font-mono" style={{ fontSize: '10px' }}>
              Cut-to-Cut Pure Cost (₹79/kg)
            </span>
          )}
        </div>
        <div className="row gap" style={{ alignItems: 'center' }}>
          <div className="legend-pills row gap">
            <span className="legend-item"><span className="legend-dot inp" /> Input</span>
            <span className="legend-item"><span className="legend-dot chg" /> Changed</span>
            <span className="legend-item"><span className="legend-dot formula-ovr" /> Override</span>
          </div>
          {onSave && (
            <button
              type="button"
              className={`btn sm ${hasUnsavedChanges ? 'primary' : 'ghost'}`}
              onClick={onSave}
              disabled={isSavingProject}
              style={{ height: 28, fontSize: 12, display: 'inline-flex', alignItems: 'center', gap: 6, marginLeft: 8 }}
              title="Save worksheet changes to server"
            >
              {isSavingProject ? <span className="spinner sm" /> : '💾'}
              {isSavingProject ? 'Saving…' : hasUnsavedChanges ? 'Save Changes' : 'Saved'}
            </button>
          )}
        </div>
      </div>

      <div className="sheets card">
        {/* Formula Bar */}
        <div className="fbar">
          <div className="fbar-addr">{sel ? addr(sel.r, sel.c) : '—'}</div>
          <div className="fbar-fx">fx</div>
          <div className="fbar-body">
            {sel === null ? (
              <span className="muted">Click any cell to inspect its value, formula, and dependencies</span>
            ) : editFormulaMode ? (
              <div className="formula-edit-wrap row gap grow">
                <input
                  type="text"
                  autoFocus
                  className="formula-input grow"
                  value={formulaInput}
                  onChange={(e) => setFormulaInput(e.target.value)}
                  onKeyDown={(e) => {
                    if (e.key === 'Enter') handleSaveFormula()
                    if (e.key === 'Escape') setEditFormulaMode(false)
                  }}
                />
                <button className="btn primary sm" onClick={handleSaveFormula}>
                  Apply Formula
                </button>
                <button className="btn ghost sm" onClick={() => setEditFormulaMode(false)}>
                  Cancel
                </button>
              </div>
            ) : isInput ? (
              <div className="row gap">
                <span className="chip ok">Input Cell</span>
                <span className="tiny muted">Override value:</span>
                <NumField
                  className="fbar-input"
                  value={typeof selVal === 'number' ? selVal : null}
                  onChange={(v) => setInput(selKey, v, workbookType)}
                />
                {isOverriddenInput && (
                  <button
                    className="btn ghost sm"
                    onClick={() => setInput(selKey, undefined, workbookType)}
                    title="Reset to template default"
                  >
                    {Icon.reset} Reset Default
                  </button>
                )}
              </div>
            ) : isFormula ? (
              <div className="row gap between grow">
                <div className="row gap">
                  <code>{String(selRaw)}</code>
                  <span>=</span>
                  <strong className="fbar-result">{fmtCell(selVal)}</strong>
                  {isOverriddenFormula && (
                    <span className="chip warn" title="This formula has been customized for this project">
                      Project Formula Override
                    </span>
                  )}
                </div>
                <div className="row gap">
                  <button className="btn ghost sm" onClick={handleStartFormulaEdit} title="Override formula for this project">
                    {Icon.edit} Override Rule
                  </button>
                  {isOverriddenFormula && (
                    <button className="btn ghost sm" onClick={handleResetFormula} title="Revert to template standard formula">
                      {Icon.reset} Revert Standard
                    </button>
                  )}
                  <div className="row gap" style={{ marginLeft: 8 }}>
                    <span className="tiny muted">Or value:</span>
                    <NumField
                      className="fbar-input"
                      value={isOverriddenInput ? (typeof selVal === 'number' ? selVal : null) : null}
                      onChange={(v) => setInput(selKey, v, workbookType)}
                      placeholder="Custom val"
                    />
                    {isOverriddenInput && (
                      <button className="btn ghost sm" onClick={() => setInput(selKey, undefined, workbookType)} title="Reset value override">
                        {Icon.reset}
                      </button>
                    )}
                  </div>
                </div>
              </div>
            ) : (
              <span className="muted">
                Static label: <strong>{String(selRaw ?? '')}</strong>
              </span>
            )}
          </div>
        </div>

        {/* Spreadsheet Matrix Grid */}
        <div className="grid-scroll">
          <table className="xl">
            <thead>
              <tr>
                <th className="corner" />
                {Array.from({ length: sheet.cols }, (_, c) => {
                  const colLetter = numToCol(c)
                  const isHl = sel?.c === c
                  return (
                    <th key={c} style={{ width: widths[c] }} className={isHl ? 'hl' : ''}>
                      {colLetter}
                    </th>
                  )
                })}
              </tr>
            </thead>
            <tbody>
              {Array.from({ length: sheet.rows }, (_, r) => (
                <tr key={r}>
                  <th className={sel?.r === r ? 'hl' : ''}>{r + 1}</th>
                  {Array.from({ length: sheet.cols }, (_, c) => {
                    const rc = `${r},${c}`
                    if (covered.has(rc)) return null
                    const sp = spans.get(rc)
                    const key = Workbook.key(name, r, c)
                    const val = activeWb.value(name, r, c)
                    const raw = activeWb.raw(name, r, c)
                    const isCellInput = activeWb.isInput(name, r, c) || EXTRA_EDITABLE.has(key)
                    const isCellChangedInput = key in activeOverrides
                    const isCellFormulaOvr = key in activeFormulaOverrides
                    const isSel = sel?.r === r && sel?.c === c
                    const style: CellStyle | undefined = sheet.styles?.[rc]

                    const classes = [
                      typeof val === 'number' ? 'n' : '',
                      style?.b ? 'b' : '',
                      style?.i ? 'it' : '',
                      style?.al ? `a${style.al}` : '',
                      val instanceof ErrVal ? 'err' : '',
                      isCellInput ? 'inp' : '',
                      isCellChangedInput ? 'chg' : '',
                      isCellFormulaOvr ? 'formula-ovr' : '',
                      isSel ? 'sel' : '',
                    ]
                      .filter(Boolean)
                      .join(' ')

                    const inlineStyle: React.CSSProperties = {}
                    if (style?.bg) {
                      inlineStyle.backgroundColor = '#' + style.bg
                      inlineStyle.color = textOn(style.bg)
                    }

                    return (
                      <td
                        key={c}
                        rowSpan={sp?.rs}
                        colSpan={sp?.cs}
                        className={classes}
                        style={inlineStyle}
                        onClick={() => setSel({ r, c })}
                        onDoubleClick={() => {
                          setSel({ r, c })
                          if (isCellInput || typeof val === 'number') {
                            setInlineKey(key)
                            setInlineVal(val !== null && val !== undefined ? String(val) : '')
                          }
                        }}
                      >
                        {inlineKey === key ? (
                          <input
                            type="number"
                            autoFocus
                            value={inlineVal}
                            onChange={(e) => setInlineVal(e.target.value)}
                            onBlur={() => {
                              const num = inlineVal.trim() === '' ? undefined : Number(inlineVal)
                              setInput(key, isNaN(num as number) ? undefined : num, workbookType)
                              setInlineKey(null)
                            }}
                            onKeyDown={(e) => {
                              if (e.key === 'Enter') {
                                const num = inlineVal.trim() === '' ? undefined : Number(inlineVal)
                                setInput(key, isNaN(num as number) ? undefined : num, workbookType)
                                setInlineKey(null)
                              } else if (e.key === 'Escape') {
                                setInlineKey(null)
                              }
                            }}
                            style={{
                              width: '100%',
                              height: '100%',
                              padding: '1px 3px',
                              fontSize: '11px',
                              border: '1.5px solid #2563eb',
                              outline: 'none',
                              background: '#fff',
                              color: '#000',
                            }}
                          />
                        ) : (
                          fmtCell(val)
                        )}
                      </td>
                    )
                  })}
                </tr>
              ))}
            </tbody>
          </table>
        </div>

        {/* Bottom Sheet Navigation Tabs */}
        <div className="tabs-bottom" ref={tabsRef}>
          {activeWb.names.map((n: string) => {
            const hasOverrides = Object.keys(activeOverrides).some((k) => k.startsWith(n + '!'))
            const hasFormulaOvr = Object.keys(activeFormulaOverrides).some((k) => k.startsWith(n + '!'))
            const meta = serverMeta[n] || DEFAULT_METADATA[n]
            return (
              <button
                key={n}
                className={n === name ? 'on' : ''}
                onClick={() => setName(n)}
                title={meta?.desc || n}
              >
                {n}
                {hasFormulaOvr && <span className="dot purple-dot" title="Formula override" />}
                {hasOverrides && !hasFormulaOvr && <span className="dot" title="Input override" />}
              </button>
            )
          })}
        </div>
      </div>
    </div>
  )
}

function numToCol(n: number): string {
  let s = ''
  n += 1
  while (n > 0) {
    s = String.fromCharCode(65 + ((n - 1) % 26)) + s
    n = Math.floor((n - 1) / 26)
  }
  return s
}
