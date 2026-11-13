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
  systems?: any[]
  activeSystemIndex?: number
  onSelectSystem?: (idx: number) => void
  overrides?: Overrides
  formulaOverrides?: Record<string, string>
  c2cOverrides?: Overrides
  c2cFormulaOverrides?: Record<string, string>
  setInput: (key: string, value: number | null | undefined, targetWorkbook?: 'actual' | 'c2c') => void
  setFormulaOverride: (key: string, formula: string | undefined, targetWorkbook?: 'actual' | 'c2c') => void
  onSystemInput?: (systemIdx: number, key: string, value: number | null | undefined, targetWorkbook?: 'actual' | 'c2c') => void
  onSystemFormulaOverride?: (systemIdx: number, key: string, formula: string | undefined, targetWorkbook?: 'actual' | 'c2c') => void
  isAdminOrOwner?: boolean
  session?: Session
  org?: Org
  initialWorkbookType?: 'actual' | 'c2c'
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
  systems,
  activeSystemIndex,
  onSelectSystem,
  overrides = {},
  formulaOverrides = {},
  c2cOverrides = {},
  c2cFormulaOverrides = {},
  setInput,
  setFormulaOverride,
  onSystemInput,
  onSystemFormulaOverride,
  isAdminOrOwner = true,
  session,
  org,
  initialWorkbookType,
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

  const [currentSystemIdx, setCurrentSystemIdx] = useState<number>(() => {
    if (typeof activeSystemIndex === 'number' && activeSystemIndex >= 0) return activeSystemIndex
    return 0
  })

  useEffect(() => {
    if (typeof activeSystemIndex === 'number' && activeSystemIndex >= 0) {
      setCurrentSystemIdx(activeSystemIndex)
    }
  }, [activeSystemIndex])

  const selectedSys = systems && systems[currentSystemIdx] ? systems[currentSystemIdx] : null

  const effectiveSysOverrides = useMemo(() => {
    if (currentSystemIdx === 0) {
      return { ...(selectedSys?.overrides || {}), ...overrides }
    }
    if (selectedSys && selectedSys.overrides && Object.keys(selectedSys.overrides).length > 0) {
      return selectedSys.overrides
    }
    return {}
  }, [selectedSys, currentSystemIdx, overrides])

  const effectiveSysFormulaOverrides = useMemo(() => {
    if (currentSystemIdx === 0) {
      return { ...(selectedSys?.formulaOverrides || {}), ...formulaOverrides }
    }
    if (selectedSys && selectedSys.formulaOverrides && Object.keys(selectedSys.formulaOverrides).length > 0) {
      return selectedSys.formulaOverrides
    }
    return {}
  }, [selectedSys, currentSystemIdx, formulaOverrides])

  const effectiveSysC2COverrides = useMemo(() => {
    if (currentSystemIdx === 0) {
      return { ...(selectedSys?.c2c_overrides || {}), ...c2cOverrides }
    }
    if (selectedSys && selectedSys.c2c_overrides) {
      return selectedSys.c2c_overrides
    }
    return {}
  }, [selectedSys, currentSystemIdx, c2cOverrides])

  const effectiveSysC2CFormulaOverrides = useMemo(() => {
    if (currentSystemIdx === 0) {
      return { ...(selectedSys?.c2c_formula_overrides || {}), ...c2cFormulaOverrides }
    }
    if (selectedSys && selectedSys.c2c_formula_overrides) {
      return selectedSys.c2c_formula_overrides
    }
    return {}
  }, [selectedSys, currentSystemIdx, c2cFormulaOverrides])

  const handleApplyInput = (key: string, value: number | null | undefined, targetWorkbook: 'actual' | 'c2c' = 'actual') => {
    const targetWb = targetWorkbook === 'c2c'
      ? (serverWb ? serverWb.c2c : c2cWb)
      : (serverWb ? serverWb.actual : wb)

    if (targetWb) {
      const currentOvr = { ...targetWb.getOverrides() }
      if (value === undefined || value === null) {
        delete currentOvr[key]
      } else {
        currentOvr[key] = value
      }
      targetWb.setOverrides(currentOvr)
      setRenderVer((v) => v + 1)
    }

    if (onSystemInput) {
      onSystemInput(currentSystemIdx, key, value, targetWorkbook)
    } else {
      setInput(key, value, targetWorkbook)
    }
  }

  const handleApplyFormulaOverride = (key: string, formula: string | undefined, targetWorkbook: 'actual' | 'c2c' = 'actual') => {
    const targetWb = targetWorkbook === 'c2c'
      ? (serverWb ? serverWb.c2c : c2cWb)
      : (serverWb ? serverWb.actual : wb)

    if (targetWb) {
      const currentFO = { ...targetWb.getFormulaOverrides() }
      if (formula === undefined || formula === null) {
        delete currentFO[key]
      } else {
        currentFO[key] = formula
      }
      targetWb.setFormulaOverrides(currentFO)
      setRenderVer((v) => v + 1)
    }

    if (onSystemFormulaOverride) {
      onSystemFormulaOverride(currentSystemIdx, key, formula, targetWorkbook)
    } else {
      setFormulaOverride(key, formula, targetWorkbook)
    }
  }

  const handleSwitchSystemTab = (idx: number) => {
    setCurrentSystemIdx(idx)
    onSelectSystem?.(idx)
    setSel(null)
    setEditFormulaMode(false)
    setInlineKey(null)
    const sysName = systems?.[idx]?.name || `SYSTEM ${String.fromCharCode(65 + idx)}`
    toast(`Viewing Excel sheet for ${sysName}`)
  }

  const [workbookType, setWorkbookType] = useState<'actual' | 'c2c'>(initialWorkbookType || 'actual')

  useEffect(() => {
    if (initialWorkbookType) {
      setWorkbookType(initialWorkbookType)
    }
  }, [initialWorkbookType])
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
        act.setOverrides(effectiveSysOverrides)
        act.setFormulaOverrides(effectiveSysFormulaOverrides)
        c2c.setOverrides({ ...filterCutToCutOverrides(effectiveSysOverrides), ...effectiveSysC2COverrides })
        c2c.setFormulaOverrides({ ...effectiveSysFormulaOverrides, ...effectiveSysC2CFormulaOverrides })
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
      serverWb.actual.setOverrides(effectiveSysOverrides)
      serverWb.actual.setFormulaOverrides(effectiveSysFormulaOverrides)
      serverWb.c2c.setOverrides({ ...filterCutToCutOverrides(effectiveSysOverrides), ...effectiveSysC2COverrides })
      serverWb.c2c.setFormulaOverrides({ ...effectiveSysFormulaOverrides, ...effectiveSysC2CFormulaOverrides })
      setRenderVer((v) => v + 1)
    } else {
      wb.setOverrides(effectiveSysOverrides)
      wb.setFormulaOverrides(effectiveSysFormulaOverrides)
      if (c2cWb) {
        c2cWb.setOverrides({ ...filterCutToCutOverrides(effectiveSysOverrides), ...effectiveSysC2COverrides })
        c2cWb.setFormulaOverrides({ ...effectiveSysFormulaOverrides, ...effectiveSysC2CFormulaOverrides })
      }
      setRenderVer((v) => v + 1)
    }
  }, [serverWb, effectiveSysOverrides, effectiveSysFormulaOverrides, effectiveSysC2COverrides, effectiveSysC2CFormulaOverrides, wb, c2cWb])

  const activeOverrides = workbookType === 'c2c'
    ? { ...filterCutToCutOverrides(effectiveSysOverrides), ...effectiveSysC2COverrides }
    : effectiveSysOverrides

  const activeFormulaOverrides = workbookType === 'c2c'
    ? { ...effectiveSysFormulaOverrides, ...effectiveSysC2CFormulaOverrides }
    : effectiveSysFormulaOverrides

  const activeWb: Workbook = serverWb
    ? (workbookType === 'c2c' ? serverWb.c2c : serverWb.actual)
    : (workbookType === 'c2c' && c2cWb ? c2cWb : wb)

  const c2cAnalysis = useMemo(() => {
    const actWbToUse = serverWb ? serverWb.actual : wb
    const c2cWbToUse = serverWb ? serverWb.c2c : (c2cWb || wb)
    if (!actWbToUse || !c2cWbToUse) return { profit: 0, total: 0, ratio: 0 }

    const actualBeforeGst = actWbToUse.num('Total cal', 'N27') || 0
    const c2cExWorks = c2cWbToUse.num('Total cal', 'N22') || 0
    const c2cTransport = c2cWbToUse.num('Total cal', 'N24') || 0
    const c2cAcc = c2cWbToUse.num('Total cal', 'N25') || 0
    const c2cInst = c2cWbToUse.num('Total cal', 'N23') || 0
    const c2cCostBase = c2cExWorks + c2cTransport + c2cAcc + c2cInst
    const profit = actualBeforeGst - c2cCostBase

    const c2cLabour = c2cWbToUse.num('Total cal', 'N50') || 0
    const total = profit + c2cLabour

    const c2cWeight = c2cWbToUse.num('Total cal', 'J44') || c2cWbToUse.num('Total cal', 'J45') || 1
    const ratio = c2cWeight > 0 ? (total / c2cWeight) : 0

    return { profit, total, ratio, costBase: c2cCostBase, labour: c2cLabour, steelWeight: c2cWeight }
  }, [serverWb, wb, c2cWb, effectiveSysOverrides, effectiveSysC2COverrides])

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

  // Fast heuristic column widths without re-evaluating entire AST dependency graph on every stroke
  const widths = useMemo(() => {
    const w: number[] = []
    if (!sheet) return w
    const sampleLimit = Math.min(sheet.rows, 35)
    for (let c = 0; c < sheet.cols; c++) {
      let max = 4
      for (let r = 0; r < sampleLimit; r++) {
        if (spans.has(`${r},${c}`) || covered.has(`${r},${c}`)) continue
        const raw = sheet.cells[r]?.[c]
        if (raw === null || raw === undefined) continue
        const len = typeof raw === 'number' ? 9 : (typeof raw === 'string' ? Math.min(22, raw.length) : 5)
        if (len > max) max = len
      }
      w.push(Math.max(48, Math.min(240, max * 7.4 + 16)))
    }
    return w
  }, [sheet, spans, covered])

  useEffect(() => {
    setSel(null)
    setEditFormulaMode(false)
    setInlineKey(null)
  }, [name, workbookType, currentSystemIdx])

  const selKey = sel ? Workbook.key(name, sel.r, sel.c) : ''
  const selVal = sel ? activeWb.value(name, sel.r, sel.c) : null
  const selRaw = sel ? activeWb.raw(name, sel.r, sel.c) : null
  const isFormula = sel ? typeof selRaw === 'string' && selRaw.startsWith('=') : false
  const isInput = sel ? (activeWb.isInput(name, sel.r, sel.c) || EXTRA_EDITABLE.has(selKey) || (!isFormula && typeof selVal === 'number')) : false
  const isOverriddenInput = sel ? selKey in activeOverrides : false
  const isOverriddenFormula = sel ? selKey in activeFormulaOverrides : false

  const [fbarVal, setFbarVal] = useState<string>('')
  useEffect(() => {
    setFbarVal(selVal !== null && selVal !== undefined && typeof selVal === 'number' ? String(selVal) : '')
  }, [selKey, selVal])

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
    handleApplyFormulaOverride(selKey, f, workbookType)
    setEditFormulaMode(false)
    toast(workbookType === 'c2c' ? 'Cut-to-Cut formula override applied' : 'Project-specific formula override applied')
  }

  const handleResetFormula = () => {
    if (!sel) return
    handleApplyFormulaOverride(selKey, undefined, workbookType)
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

          {/* System Switcher (when multiple systems exist) */}
          {systems && systems.length > 1 && (
            <div style={{ marginLeft: 8, display: 'inline-flex', alignItems: 'center', gap: 3, background: '#f1f5f9', padding: '2px 4px', borderRadius: 6, border: '1px solid #cbd5e1' }}>
              <span style={{ fontSize: '10.5px', fontWeight: 700, color: '#475569', padding: '0 4px', textTransform: 'uppercase', letterSpacing: '0.2px' }}>
                System:
              </span>
              {systems.map((sys: any, idx: number) => {
                const isCurrent = idx === currentSystemIdx
                const sysLabel = sys.name || `SYSTEM ${String.fromCharCode(65 + idx)}`
                return (
                  <button
                    key={idx}
                    type="button"
                    className={`btn xs ${isCurrent ? 'primary' : 'ghost'}`}
                    onClick={() => handleSwitchSystemTab(idx)}
                    style={
                      isCurrent
                        ? { background: '#2563eb', borderColor: '#2563eb', color: '#fff', fontWeight: 700, fontSize: '11px', padding: '3px 8px' }
                        : { fontSize: '11px', padding: '3px 7px', color: '#475569' }
                    }
                    title={`View Excel calculation sheets for ${sysLabel}`}
                  >
                    📦 {sysLabel}
                  </button>
                )
              })}
            </div>
          )}

          {(c2cWb || serverWb) && (
            <div style={{ marginLeft: 8, display: 'inline-flex', gap: 4, background: '#f1f5f9', padding: '2px 4px', borderRadius: 6, border: '1px solid #cbd5e1' }}>
              <button
                type="button"
                className={`btn xs ${workbookType === 'actual' ? 'primary' : 'ghost'}`}
                onClick={() => setWorkbookType('actual')}
                style={workbookType === 'actual' ? { background: '#0f172a', color: '#fff', fontSize: '11px', fontWeight: 700 } : { fontSize: '11px' }}
                title="Actual Quotation Engineering Sheets"
              >
                📊 Actual Quotation
              </button>
              <button
                type="button"
                className={`btn xs ${workbookType === 'c2c' ? 'primary' : 'ghost'}`}
                onClick={() => setWorkbookType('c2c')}
                style={workbookType === 'c2c' ? { background: '#7c3aed', color: '#fff', fontSize: '11px', fontWeight: 700 } : { fontSize: '11px', color: '#7c3aed' }}
                title="Cut-to-Cut Cost Engineering Sheets"
              >
                ✂️ Cut-to-Cut Sheet
              </button>
            </div>
          )}
          {workbookType === 'c2c' && (
            <span className="chip purple font-mono" style={{ fontSize: '10px' }}>
              Cut-to-Cut Cost Baseline
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
                <input
                  type="number"
                  className="fbar-input"
                  style={{ width: '90px', padding: '3px 6px', fontWeight: 600, border: '1.5px solid #cbd5e1', borderRadius: '4px', textAlign: 'center' }}
                  value={fbarVal}
                  onChange={(e) => setFbarVal(e.target.value)}
                  onBlur={() => {
                    const num = fbarVal.trim() === '' ? undefined : Number(fbarVal)
                    handleApplyInput(selKey, isNaN(num as number) ? undefined : num, workbookType)
                  }}
                  onKeyDown={(e) => {
                    if (e.key === 'Enter') {
                      const num = fbarVal.trim() === '' ? undefined : Number(fbarVal)
                      handleApplyInput(selKey, isNaN(num as number) ? undefined : num, workbookType)
                    }
                  }}
                />
                {isOverriddenInput && (
                  <button
                    className="btn ghost sm"
                    onClick={() => handleApplyInput(selKey, undefined, workbookType)}
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
                    <input
                      type="number"
                      className="fbar-input"
                      placeholder="Custom val"
                      style={{ width: '90px', padding: '3px 6px', fontWeight: 600, border: '1.5px solid #cbd5e1', borderRadius: '4px', textAlign: 'center' }}
                      value={fbarVal}
                      onChange={(e) => setFbarVal(e.target.value)}
                      onBlur={() => {
                        const num = fbarVal.trim() === '' ? undefined : Number(fbarVal)
                        handleApplyInput(selKey, isNaN(num as number) ? undefined : num, workbookType)
                      }}
                      onKeyDown={(e) => {
                        if (e.key === 'Enter') {
                          const num = fbarVal.trim() === '' ? undefined : Number(fbarVal)
                          handleApplyInput(selKey, isNaN(num as number) ? undefined : num, workbookType)
                        }
                      }}
                    />
                    {isOverriddenInput && (
                      <button className="btn ghost sm" onClick={() => handleApplyInput(selKey, undefined, workbookType)} title="Reset value override">
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
                    let val = activeWb.value(name, r, c)
                    let raw = activeWb.raw(name, r, c)

                    // Inject 3-line C2C Widget live at E26:G28 on Total cal
                    if (name === 'Total cal') {
                      if (r === 25 && c === 4) { raw = 'C2C PROFIT'; val = 'C2C PROFIT' }
                      if (r === 25 && c === 5) { raw = c2cAnalysis.profit; val = c2cAnalysis.profit }
                      if (r === 25 && c === 6) { raw = 'Actual Before GST - C2C Cost Base'; val = 'Actual Before GST - C2C Cost Base' }

                      if (r === 26 && c === 4) { raw = 'C2C TOTAL'; val = 'C2C TOTAL' }
                      if (r === 26 && c === 5) { raw = c2cAnalysis.total; val = c2cAnalysis.total }
                      if (r === 26 && c === 6) { raw = 'Profit + C2C Labour'; val = 'Profit + C2C Labour' }

                      if (r === 27 && c === 4) { raw = 'C2C RATIO'; val = 'C2C RATIO' }
                      if (r === 27 && c === 5) { raw = Number(c2cAnalysis.ratio.toFixed(2)); val = Number(c2cAnalysis.ratio.toFixed(2)) }
                      if (r === 27 && c === 6) { raw = 'Total / C2C Weight (₹/kg)'; val = 'Total / C2C Weight (₹/kg)' }
                    }
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
                              handleApplyInput(key, isNaN(num as number) ? undefined : num, workbookType)
                              setInlineKey(null)
                            }}
                            onKeyDown={(e) => {
                              if (e.key === 'Enter') {
                                const num = inlineVal.trim() === '' ? undefined : Number(inlineVal)
                                handleApplyInput(key, isNaN(num as number) ? undefined : num, workbookType)
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
