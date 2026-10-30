import { useEffect, useMemo, useState } from 'react'
import data from '../../lib/clientWorkbook.json'
import { api, Org, Session, Template } from '../../lib/api'
import { publishTemplate, rollbackTemplateToVersion } from '../../lib/permissions'
import { Overrides, SheetData, Workbook } from '../../lib/engine'
import { inr, timeAgo } from '../../lib/format'
import { INPUT_GROUPS } from '../../lib/inputs'
import { cleanOverrides, summarize, Summary } from '../../lib/model'
import { Icon, NumField, Spinner, useCopy, useUI } from '../ui'
import { CalculationDetails } from './CalculationDetails'

interface Props {
  session: Session
  org: Org
  onOrgUpdated: () => void
}

type StudioTab = 'sheets' | 'forms' | 'diff'

export function OrganisationSettings({ session, org, onOrgUpdated }: Props) {
  const { toast, confirm } = useUI()
  const copy = useCopy()

  const [templates, setTemplates] = useState<Template[]>([])
  const [loading, setLoading] = useState(true)
  const [publishing, setPublishing] = useState(false)

  // Studio / Editor Mode state
  const [editorMode, setEditorMode] = useState(false)
  const [studioTab, setStudioTab] = useState<StudioTab>('sheets')
  const [draftName, setDraftName] = useState('')
  const [draftDesc, setDraftDesc] = useState('')
  const [draftOverrides, setDraftOverrides] = useState<Overrides>({})
  const [draftFormulaOverrides, setDraftFormulaOverrides] = useState<Record<string, string>>({})

  // Template Baseline: 'actual' (₹90 rate default workbook) or 'c2c' (₹79 rate Cut-to-Cut workbook)
  const [templateBaseline, setTemplateBaseline] = useState<'actual' | 'c2c'>('actual')

  // Dedicated workbook instances for template draft calculation
  const draftWb = useMemo(() => new Workbook(data as unknown as { sheets: SheetData[] }), [])

  const isAdminOrOwner = org.role === 'owner' || org.role === 'admin'

  const loadTemplates = async () => {
    try {
      setLoading(true)
      const list = await api.templateList(session.token, org.id)
      setTemplates(list)
    } catch (e) {
      toast((e as Error).message, 'err')
    } finally {
      setLoading(false)
    }
  }

  useEffect(() => {
    loadTemplates()
  }, [org.id])

  const currentDefaultTemplate = templates.find((t) => t.is_default) || templates[0]
  const nextVersion = (org.current_template_version || 1) + 1

  // Open Template Studio (Supports forking Actual or Cut-to-Cut baseline)
  const handleOpenStudio = (baseTemplate?: Template, forceBaseline?: 'actual' | 'c2c') => {
    const tpl = baseTemplate || currentDefaultTemplate
    const initialInputs = tpl?.definition?.default_inputs || {}
    const initialFormulas = tpl?.definition?.formula_overrides || {}
    const isC2C = forceBaseline === 'c2c' || tpl?.definition?.baseline === 'c2c' || tpl?.name?.toLowerCase().includes('cut-to-cut')

    setTemplateBaseline(isC2C ? 'c2c' : 'actual')

    if (isC2C) {
      setDraftName(baseTemplate ? `${baseTemplate.name} (V${nextVersion})` : `Cut-to-Cut Standards V${nextVersion} (₹79/kg)`)
      setDraftDesc(baseTemplate?.description || `Cut-to-cut internal production standards (Steel Rate ₹79/kg, internal sheared sizes, 0 stiffener) published by ${session.username}`)
      const c2cInputs: Overrides = {
        'Total cal!H3': 79,
        'Total cal!M3': 69,
        'Total cal!J3': 0,
        ...initialInputs,
      }
      setDraftOverrides(c2cInputs)
      setDraftFormulaOverrides({ ...initialFormulas })

      draftWb.setTemplateDefaults(c2cInputs)
      draftWb.setFormulaOverrides(initialFormulas)
      draftWb.setOverrides(c2cInputs)
    } else {
      setDraftName(baseTemplate ? `${baseTemplate.name} (V${nextVersion})` : `Default Production Template V${nextVersion}`)
      setDraftDesc(baseTemplate?.description || `Updated standards & rates across worksheets published by ${session.username}`)
      setDraftOverrides({ ...initialInputs })
      setDraftFormulaOverrides({ ...initialFormulas })

      draftWb.setTemplateDefaults(initialInputs)
      draftWb.setFormulaOverrides(initialFormulas)
      draftWb.setOverrides(initialInputs)
    }

    setStudioTab('sheets')
    setEditorMode(true)
  }

  // Update cell input in template draft
  const handleDraftSetInput = (key: string, value: number | null | undefined) => {
    setDraftOverrides((prev) => {
      const next = { ...prev }
      if (value === undefined || value === null) delete next[key]
      else next[key] = value
      return next
    })
  }

  // Update formula override in template draft
  const handleDraftSetFormulaOverride = (key: string, formula: string | undefined) => {
    setDraftFormulaOverrides((prev) => {
      const next = { ...prev }
      if (formula === undefined || !formula.trim()) delete next[key]
      else next[key] = formula
      return next
    })
  }

  // Live recalculate summary for draft template
  const draftSummary: Summary = useMemo(() => {
    draftWb.setFormulaOverrides(draftFormulaOverrides)
    draftWb.setOverrides(draftOverrides)
    return summarize(draftWb)
  }, [draftOverrides, draftFormulaOverrides, draftWb])

  const modifiedCellsCount = Object.keys(draftOverrides).length + Object.keys(draftFormulaOverrides).length

  // Set any published template version as the active default for future projects
  const handleSetActiveDefault = async (t: Template) => {
    if (!isAdminOrOwner) {
      return toast('Only organisation owners and administrators can change the default template.', 'err')
    }
    const ok = await confirm({
      title: `Set Version ${t.version_num} as Active Default?`,
      body: `Future projects created in ${org.name} will snapshot Version ${t.version_num} (${t.name}) by default. Historical projects will not be affected.`,
      confirmLabel: `Set as Default (v${t.version_num})`,
    })
    if (!ok) return

    setLoading(true)
    try {
      await rollbackTemplateToVersion(session, org, t.id)
      toast(`Version ${t.version_num} is now the active default template for future projects!`)
      await loadTemplates()
      onOrgUpdated()
    } catch (e) {
      toast((e as Error).message, 'err')
    } finally {
      setLoading(false)
    }
  }

  // Publish new template version
  const handlePublishTemplate = async () => {
    if (!isAdminOrOwner) {
      return toast('Only organisation owners and administrators can publish templates.', 'err')
    }
    const nm = draftName.trim()
    if (!nm) return toast('Template name cannot be empty', 'err')

    const steelRate = draftWb.num('Total cal', 'H3') || 90
    const trolleyRate = draftWb.num('Total cal', 'M3') || 90
    const bodyLabourRate = draftWb.num('Total cal', 'B49') || 250

    // Strict input validation to prevent invalid, zero, or negative rates, excessive wastage, or formula crash
    if (steelRate <= 0 || !Number.isFinite(steelRate)) {
      return toast('Steel rate must be a valid positive number.', 'err')
    }
    if (trolleyRate <= 0 || !Number.isFinite(trolleyRate)) {
      return toast('Trolley rate must be a valid positive number.', 'err')
    }
    if (bodyLabourRate <= 0 || !Number.isFinite(bodyLabourRate)) {
      return toast('Body labour rate must be a valid positive number.', 'err')
    }

    // Validate custom cell overrides
    for (const [cell, val] of Object.entries(draftOverrides)) {
      if (val !== null && val !== undefined) {
        if (!Number.isFinite(val)) {
          return toast(`Cell ${cell} contains an invalid number.`, 'err')
        }
        if (val < 0) {
          return toast(`Cell ${cell} cannot have a negative value (${val}).`, 'err')
        }
        if (/wastage|loss/i.test(cell) && val > 100) {
          return toast(`Cell ${cell} wastage percentage cannot exceed 100%.`, 'err')
        }
        if (/density/i.test(cell) && val <= 0) {
          return toast(`Cell ${cell} density factor must be strictly greater than 0.`, 'err')
        }
      }
    }

    const ok = await confirm({
      title: `Publish Template Version ${nextVersion}?`,
      body: `This will become the active default template for all newly created projects in ${org.name}. Previously created projects will NOT be modified.`,
      confirmLabel: `Publish Version ${nextVersion}`,
    })
    if (!ok) return

    setPublishing(true)
    try {
      const definition = {
        version: nextVersion,
        name: nm,
        description: draftDesc,
        baseline: templateBaseline,
        rates: {
          steel: steelRate,
          trolleySteel: trolleyRate,
          bodyLabour: bodyLabourRate,
        },
        default_inputs: {
          ...draftOverrides,
          'Total cal!H3': steelRate,
          'rates!b3': steelRate,
          'Total cal!M3': trolleyRate,
          'rates!b4': trolleyRate,
          'Total cal!B49': bodyLabourRate,
        },
        formula_overrides: draftFormulaOverrides,
      }

      const res = await publishTemplate(session, org, nm, draftDesc, definition)
      toast(`Successfully published Template Version ${res.version_num}!`)
      setEditorMode(false)
      await loadTemplates()
      onOrgUpdated()
    } catch (e) {
      toast((e as Error).message, 'err')
    } finally {
      setPublishing(false)
    }
  }

  // -------------------------------------------------------------------------
  // Render: Template Studio (Full All-15-Sheets Editor)
  // -------------------------------------------------------------------------
  if (editorMode) {
    return (
      <div className="template-studio">
        {/* Big Top Switcher: 1. Actual vs 2. Cut-to-Cut */}
        <div
          style={{
            display: 'grid',
            gridTemplateColumns: '1fr 1fr',
            gap: '12px',
            marginBottom: '14px',
          }}
        >
          <button
            type="button"
            className="btn"
            style={{
              padding: '14px 20px',
              fontSize: '15px',
              fontWeight: 700,
              borderRadius: '10px',
              border: templateBaseline === 'actual' ? '2.5px solid #2563eb' : '1.5px solid #cbd5e1',
              background: templateBaseline === 'actual' ? '#eff6ff' : '#ffffff',
              color: templateBaseline === 'actual' ? '#1e40af' : '#64748b',
              boxShadow: templateBaseline === 'actual' ? '0 4px 12px rgba(37,99,235,0.12)' : 'none',
              display: 'flex',
              alignItems: 'center',
              justifyContent: 'center',
              gap: '10px',
              cursor: 'pointer',
              transition: 'all 0.15s ease',
            }}
            onClick={() => {
              setTemplateBaseline('actual')
              draftWb.setOverrides(draftOverrides)
            }}
          >
            <span style={{ fontSize: '18px' }}>📊</span>
            <span>1. ACTUAL TEMPLATE (₹90/kg)</span>
            {templateBaseline === 'actual' && <span className="chip ok" style={{ fontSize: '11px', padding: '2px 6px' }}>Active</span>}
          </button>
          <button
            type="button"
            className="btn"
            style={{
              padding: '14px 20px',
              fontSize: '15px',
              fontWeight: 700,
              borderRadius: '10px',
              border: templateBaseline === 'c2c' ? '2.5px solid #7c3aed' : '1.5px solid #cbd5e1',
              background: templateBaseline === 'c2c' ? '#faf5ff' : '#ffffff',
              color: templateBaseline === 'c2c' ? '#6d28d9' : '#64748b',
              boxShadow: templateBaseline === 'c2c' ? '0 4px 12px rgba(124,58,237,0.12)' : 'none',
              display: 'flex',
              alignItems: 'center',
              justifyContent: 'center',
              gap: '10px',
              cursor: 'pointer',
              transition: 'all 0.15s ease',
            }}
            onClick={() => {
              setTemplateBaseline('c2c')
              draftWb.setOverrides(draftOverrides)
            }}
          >
            <span style={{ fontSize: '18px' }}>✂️</span>
            <span>2. CUT-TO-CUT TEMPLATE (₹79/kg)</span>
            {templateBaseline === 'c2c' && <span className="chip" style={{ background: '#ede9fe', color: '#6d28d9', fontSize: '11px', padding: '2px 6px' }}>Active</span>}
          </button>
        </div>

        {/* Studio Top Control Bar */}
        <div className="card studio-header mb-3">
          <div className="row between mb-2">
            <button className="btn ghost sm" onClick={() => setEditorMode(false)}>
              ← Back to Organisation Settings
            </button>
            <div className="row gap">
              <span className="badge-lg">Publishing Version {nextVersion}</span>
              <span className="chip" style={{ background: templateBaseline === 'c2c' ? '#ede9fe' : '#dbeafe', color: templateBaseline === 'c2c' ? '#6d28d9' : '#1e40af', fontWeight: 600 }}>
                {templateBaseline === 'c2c' ? '✂️ Cut-to-Cut (₹79 Rate)' : '📊 Actual (₹90 Rate)'}
              </span>
            </div>
            <button
              className="btn primary"
              disabled={publishing || !draftName.trim()}
              onClick={handlePublishTemplate}
            >
              {publishing ? <Spinner size={16} /> : <span style={{ display: 'inline-flex', alignItems: 'center', gap: 6 }}>{Icon.plus} Publish Version {nextVersion}</span>}
            </button>
          </div>

          <div className="row gap mt-2">
            <div className="grow">
              <input
                type="text"
                className="studio-title-input"
                placeholder="Template Version Name…"
                value={draftName}
                onChange={(e) => setDraftName(e.target.value)}
              />
              <input
                type="text"
                className="studio-desc-input"
                placeholder="Release notes & rationale for worksheet/rate changes…"
                value={draftDesc}
                onChange={(e) => setDraftDesc(e.target.value)}
              />
            </div>
            <div className="studio-kpis row gap">
              <div className="studio-kpi">
                <span className="tiny font-bold uppercase muted">Baseline Sales</span>
                <strong>{inr(draftSummary.sales)}</strong>
              </div>
              <div className="studio-kpi">
                <span className="tiny font-bold uppercase muted">Customized Cells</span>
                <span className="chip purple font-bold">{modifiedCellsCount} modified</span>
              </div>
            </div>
          </div>

          {/* Sub Navigation Tabs */}
          <div className="studio-tabs mt-3">
            <button
              className={`studio-tab-btn ${studioTab === 'sheets' ? 'active' : ''}`}
              onClick={() => setStudioTab('sheets')}
            >
              {Icon.grid} All 15 Excel Sheets
            </button>
            <button
              className={`studio-tab-btn ${studioTab === 'forms' ? 'active' : ''}`}
              onClick={() => setStudioTab('forms')}
            >
              {Icon.calc} Standard Parameters &amp; Rates
            </button>
            <button
              className={`studio-tab-btn ${studioTab === 'diff' ? 'active' : ''}`}
              onClick={() => setStudioTab('diff')}
            >
              {Icon.diff} Changes Diff ({modifiedCellsCount})
            </button>
          </div>
        </div>

        {/* Tab 1: All 15 Excel Sheets Grid */}
        {studioTab === 'sheets' && (
          <div className="studio-sheet-view">
            <div className="alert-banner mb-2">
              <div className="row gap">
                {Icon.info}
                <span>
                  <strong>Full 15-Sheet Template Editor:</strong> Click any cell on any of the 15 worksheets below to customize its template default value or formula rule.
                  Values configured here will serve as the default baseline for all new projects.
                </span>
              </div>
            </div>
            <CalculationDetails
              wb={draftWb}
              overrides={draftOverrides}
              formulaOverrides={draftFormulaOverrides}
              setInput={handleDraftSetInput}
              setFormulaOverride={handleDraftSetFormulaOverride}
            />
          </div>
        )}

        {/* Tab 2: Standard Parameters Form */}
        {studioTab === 'forms' && (
          <div className="studio-forms-view card p-4">
            <div className="card-head mb-3">
              <div>
                <h3>Default Parameters &amp; Standard Dimension Baseline</h3>
                <p className="tiny muted">Quick access to primary rates and cabinet dimensions across Total cal</p>
              </div>
              <button
                className="btn ghost sm"
                onClick={() => {
                  setDraftOverrides({})
                  setDraftFormulaOverrides({})
                  draftWb.setTemplateDefaults({})
                  draftWb.setFormulaOverrides({})
                  toast('Reset draft to workbook factory defaults')
                }}
              >
                {Icon.reset} Reset to Factory Defaults
              </button>
            </div>

            <div className="forms-grid">
              {INPUT_GROUPS.map((g) => (
                <section key={g.title} className="card p-3">
                  <header className="mb-2">
                    <h4 className="tiny font-bold uppercase text-blue">{g.title}</h4>
                    {g.note && <span className="muted tiny">{g.note}</span>}
                  </header>
                  <div className="fields">
                    {g.items.map((it) => {
                      const [sheet, a] = [it.key.slice(0, it.key.lastIndexOf('!')), it.key.slice(it.key.lastIndexOf('!') + 1)]
                      const m = /^([A-Z]+)(\d+)$/.exec(a)!
                      let c = 0
                      for (const ch of m[1]) c = c * 26 + ch.charCodeAt(0) - 64
                      const def = draftWb.raw(sheet, parseInt(m[2], 10) - 1, c - 1)
                      const defNum = typeof def === 'number' ? def : null
                      const changed = it.key in draftOverrides
                      const val = changed ? draftOverrides[it.key] : defNum

                      return (
                        <label key={it.key} className={`field ${changed ? 'changed' : ''}`}>
                          <span className="field-label">
                            {it.label}
                            {it.hint && <em>{it.hint}</em>}
                          </span>
                          <span className="field-input">
                            <NumField
                              value={val}
                              step={it.step}
                              min={it.min}
                              onChange={(v) => handleDraftSetInput(it.key, v)}
                            />
                            {it.unit && <span className="unit">{it.unit}</span>}
                            {changed && (
                              <button
                                type="button"
                                className="icon-btn tiny-btn"
                                title="Reset to standard default"
                                onClick={() => handleDraftSetInput(it.key, undefined)}
                              >
                                {Icon.reset}
                              </button>
                            )}
                          </span>
                        </label>
                      )
                    })}
                  </div>
                </section>
              ))}
            </div>
          </div>
        )}

        {/* Tab 3: Changes Diff & Review */}
        {studioTab === 'diff' && (
          <div className="studio-diff-view card p-4">
            <div className="card-head mb-3">
              <div>
                <h3>Template Changes Diff Preview</h3>
                <p className="tiny muted">
                  Review all cell values and formula overrides customized for Template Version {nextVersion}
                </p>
              </div>
              <span className="chip primary">{modifiedCellsCount} Modified Definitions</span>
            </div>

            {modifiedCellsCount === 0 ? (
              <div className="empty p-5">
                <p>No customized cells yet. This template currently matches the standard baseline.</p>
                <p className="tiny muted mt-1">Navigate to <strong>All 15 Excel Sheets</strong> or <strong>Standard Parameters</strong> to customize values.</p>
              </div>
            ) : (
              <table className="tbl">
                <thead>
                  <tr>
                    <th>Worksheet &amp; Cell</th>
                    <th>Type</th>
                    <th>Template Value / Formula</th>
                    <th className="r">Action</th>
                  </tr>
                </thead>
                <tbody>
                  {Object.entries(draftOverrides).map(([k, v]) => (
                    <tr key={k}>
                      <td>
                        <strong>{k}</strong>
                      </td>
                      <td>
                        <span className="chip ok">Input Value</span>
                      </td>
                      <td className="font-bold text-green">{v}</td>
                      <td className="r">
                        <button
                          className="btn ghost sm"
                          onClick={() => handleDraftSetInput(k, undefined)}
                        >
                          {Icon.reset} Revert
                        </button>
                      </td>
                    </tr>
                  ))}
                  {Object.entries(draftFormulaOverrides).map(([k, f]) => (
                    <tr key={k}>
                      <td>
                        <strong>{k}</strong>
                      </td>
                      <td>
                        <span className="chip purple">Formula Rule</span>
                      </td>
                      <td>
                        <code>{f}</code>
                      </td>
                      <td className="r">
                        <button
                          className="btn ghost sm"
                          onClick={() => handleDraftSetFormulaOverride(k, undefined)}
                        >
                          {Icon.reset} Revert
                        </button>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            )}

            <div className="isolation-notice warn mt-4">
              <strong>Enterprise Version Guarantee:</strong>
              <p className="tiny mt-1">
                Publishing Version {nextVersion} will establish these parameters as the default for newly created projects.
                Existing projects in {org.name} will remain untouched at their respective historical versions.
              </p>
            </div>
          </div>
        )}
      </div>
    )
  }

  // -------------------------------------------------------------------------
  // Render: Organisation Settings & Template Version History (Default View)
  // -------------------------------------------------------------------------
  return (
    <div className="org-settings-page">
      <div className="page-header">
        <h2>Organisation Settings &amp; Template Architecture</h2>
        <p className="muted">
          Manage your organisation profile, immutable template versions, and calculation standards across all 15 workbook sheets.
        </p>
      </div>

      {/* Organisation Profile Card */}
      <div className="card settings-card mb-4">
        <header className="card-head">
          <h3>Organisation Profile</h3>
          <span className="chip primary">{org.role.toUpperCase()}</span>
        </header>
        <div className="settings-body">
          <div className="two-col">
            <div className="profile-item">
              <span className="tiny font-bold uppercase muted">Organisation Name</span>
              <strong className="text-lg">{org.name}</strong>
            </div>
          </div>
        </div>
      </div>

      {/* Top 2 Simple Template Cards: 1. Actual vs 2. Cut-to-Cut */}
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(320px, 1fr))', gap: '16px', marginBottom: '24px' }}>
        {/* Card 1: ACTUAL TEMPLATE */}
        <div
          className="card"
          style={{
            border: '2px solid #3b82f6',
            borderRadius: '12px',
            background: 'linear-gradient(180deg, #f0f9ff 0%, #ffffff 100%)',
            padding: '20px',
            boxShadow: '0 4px 12px rgba(59, 130, 246, 0.08)',
          }}
        >
          <div className="row between" style={{ alignItems: 'flex-start', marginBottom: '12px' }}>
            <div>
              <div className="row gap" style={{ alignItems: 'center', marginBottom: '6px' }}>
                <span className="badge-lg" style={{ background: '#dbeafe', color: '#1e40af', fontWeight: 700 }}>
                  📊 1. ACTUAL TEMPLATE
                </span>
                <span className="chip ok" style={{ fontSize: '11px' }}>Customer Baseline</span>
              </div>
              <h3 style={{ margin: '0 0 4px 0', fontSize: '18px', color: '#0f172a' }}>
                {templates.find(t => t.definition?.baseline !== 'c2c' && !t.name.toLowerCase().includes('cut-to-cut'))?.name || 'Default Production Template V1'}
              </h3>
              <p className="tiny muted" style={{ margin: 0 }}>
                Customer-facing quotation baseline with 25% profit margin and standard ₹90/kg steel rate.
              </p>
            </div>
          </div>

          <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '10px', background: '#ffffff', padding: '12px', borderRadius: '8px', border: '1px solid #e2e8f0', marginBottom: '16px' }}>
            <div>
              <div className="tiny muted font-bold uppercase">Steel Rate</div>
              <div style={{ fontSize: '18px', fontWeight: 800, color: '#1e40af' }}>₹90 / kg</div>
            </div>
            <div>
              <div className="tiny muted font-bold uppercase">Trolley Rate</div>
              <div style={{ fontSize: '18px', fontWeight: 800, color: '#1e40af' }}>₹90 / kg</div>
            </div>
          </div>

          {isAdminOrOwner && (
            <button
              className="btn primary"
              style={{ width: '100%', justifyContent: 'center', padding: '10px', fontWeight: 600 }}
              onClick={() => handleOpenStudio(templates.find(t => t.definition?.baseline !== 'c2c' && !t.name.toLowerCase().includes('cut-to-cut')), 'actual')}
            >
              ✏️ Open &amp; Customize Actual Template
            </button>
          )}
        </div>

        {/* Card 2: CUT-TO-CUT TEMPLATE */}
        <div
          className="card"
          style={{
            border: '2px solid #7c3aed',
            borderRadius: '12px',
            background: 'linear-gradient(180deg, #faf5ff 0%, #ffffff 100%)',
            padding: '20px',
            boxShadow: '0 4px 12px rgba(124, 58, 237, 0.08)',
          }}
        >
          <div className="row between" style={{ alignItems: 'flex-start', marginBottom: '12px' }}>
            <div>
              <div className="row gap" style={{ alignItems: 'center', marginBottom: '6px' }}>
                <span className="badge-lg" style={{ background: '#ede9fe', color: '#6d28d9', fontWeight: 700 }}>
                  ✂️ 2. CUT-TO-CUT TEMPLATE
                </span>
                <span className="chip" style={{ background: '#f3e8ff', color: '#7e22ce', borderColor: '#e9d5ff', fontSize: '11px' }}>
                  Internal Cost Baseline
                </span>
              </div>
              <h3 style={{ margin: '0 0 4px 0', fontSize: '18px', color: '#0f172a' }}>
                {templates.find(t => t.definition?.baseline === 'c2c' || t.name.toLowerCase().includes('cut-to-cut'))?.name || 'Cut-to-Cut Standards (₹79/kg)'}
              </h3>
              <p className="tiny muted" style={{ margin: 0 }}>
                Internal confidential costing baseline with minimum material rates (₹79/kg) and sheared sheet sizes.
              </p>
            </div>
          </div>

          <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '10px', background: '#ffffff', padding: '12px', borderRadius: '8px', border: '1px solid #e2e8f0', marginBottom: '16px' }}>
            <div>
              <div className="tiny muted font-bold uppercase">Steel Rate</div>
              <div style={{ fontSize: '18px', fontWeight: 800, color: '#7c3aed' }}>₹79 / kg</div>
            </div>
            <div>
              <div className="tiny muted font-bold uppercase">Trolley Rate</div>
              <div style={{ fontSize: '18px', fontWeight: 800, color: '#7c3aed' }}>₹69 / kg</div>
            </div>
          </div>

          {isAdminOrOwner && (
            <button
              className="btn"
              style={{ width: '100%', justifyContent: 'center', padding: '10px', fontWeight: 600, background: '#7c3aed', color: '#ffffff', borderColor: '#6d28d9' }}
              onClick={() => handleOpenStudio(templates.find(t => t.definition?.baseline === 'c2c' || t.name.toLowerCase().includes('cut-to-cut')), 'c2c')}
            >
              ✂️ Open &amp; Customize Cut-to-Cut Template
            </button>
          )}
        </div>
      </div>

      {/* Template Version History Table */}
      <div className="card">
        <div className="card-head">
          <h3>Template Version History</h3>
          <span className="tiny muted">Audit trail of all published calculation templates</span>
        </div>
        <div className="table-wrap">
          {loading ? (
            <div className="page-loading">
              <Spinner size={24} />
              <p className="muted tiny">Loading template version history…</p>
            </div>
          ) : templates.length === 0 ? (
            <div className="empty">
              <p>No template versions recorded yet</p>
            </div>
          ) : (
            <table className="tbl">
              <thead>
                <tr>
                  <th>Version</th>
                  <th>Template Name &amp; Description</th>
                  <th>Status</th>
                  <th>Customized Overrides</th>
                  <th>Projects Using Version</th>
                  <th>Publisher</th>
                  <th>Published Date</th>
                  {isAdminOrOwner && <th className="r">Actions</th>}
                </tr>
              </thead>
              <tbody>
                {templates.map((t) => {
                  const inputCount = Object.keys(t.definition?.default_inputs || {}).length
                  const formulaCount = Object.keys(t.definition?.formula_overrides || {}).length
                  return (
                    <tr key={t.id} className={t.is_default ? 'highlight-row' : ''}>
                      <td>
                        <span className="version-pill">v{t.version_num}</span>
                      </td>
                      <td>
                        <div className="row gap" style={{ alignItems: 'center' }}>
                          <strong>{t.name}</strong>
                          {(t.definition?.baseline === 'c2c' || t.name.toLowerCase().includes('cut-to-cut')) && (
                            <span className="chip" style={{ background: '#f5f3ff', color: '#6d28d9', borderColor: '#ddd6fe', fontSize: '11px', padding: '2px 6px' }}>
                              ✂️ Cut-to-Cut (₹79)
                            </span>
                          )}
                        </div>
                        {t.description && <div className="tiny muted">{t.description}</div>}
                      </td>
                      <td>
                        {t.is_default ? (
                          <span className="chip ok">Current Default</span>
                        ) : t.status === 'published' ? (
                          <span className="chip primary">Published</span>
                        ) : (
                          <span className="chip warn">Draft</span>
                        )}
                      </td>
                      <td className="tiny">
                        <div>{inputCount} customized input{inputCount !== 1 ? 's' : ''}</div>
                        {formulaCount > 0 && <div className="text-purple">{formulaCount} formula override{formulaCount !== 1 ? 's' : ''}</div>}
                      </td>
                      <td>
                        <span className="chip muted-chip">
                          {t.project_count ?? 0} project{(t.project_count ?? 0) !== 1 ? 's' : ''}
                        </span>
                      </td>
                      <td className="tiny muted">{t.published_by || t.created_by || 'Admin'}</td>
                      <td className="tiny muted">{t.published_at ? timeAgo(t.published_at) : timeAgo(t.created_at)}</td>
                      {isAdminOrOwner && (
                        <td className="r">
                          <div className="row gap" style={{ justifyContent: 'flex-end' }}>
                            {!t.is_default && t.status === 'published' && (
                              <button
                                className="btn primary sm"
                                onClick={() => handleSetActiveDefault(t)}
                                title={`Set Version ${t.version_num} as the default template for future projects`}
                              >
                                {Icon.check} Set as Default
                              </button>
                            )}
                            <button
                              className="btn ghost sm"
                              onClick={() => handleOpenStudio(t, (t.definition?.baseline === 'c2c' || t.name.toLowerCase().includes('cut-to-cut')) ? 'c2c' : 'actual')}
                              title="Create new draft version based on this template"
                            >
                              {Icon.edit} Fork Version
                            </button>
                          </div>
                        </td>
                      )}
                    </tr>
                  )
                })}
              </tbody>
            </table>
          )}
        </div>
      </div>
    </div>
  )
}
