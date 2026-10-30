import { useEffect, useMemo, useState } from 'react'
import data from '../../lib/clientWorkbook.json'
import { api, Org, Session, Template } from '../../lib/api'
import { deleteTemplate, publishTemplate, rollbackTemplateToVersion, updateTemplate } from '../../lib/permissions'
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

export interface DiffEntry {
  key: string
  label?: string
  fromVal: any
  toVal: any
  type: 'input' | 'formula' | 'baseline'
}

export function computeTemplateDiff(baseTpl?: Template | null, targetTpl?: Template | null): DiffEntry[] {
  if (!targetTpl) return []
  const diffs: DiffEntry[] = []

  const baseInputs = baseTpl?.definition?.default_inputs || {}
  const targetInputs = targetTpl?.definition?.default_inputs || {}
  const allInputKeys = Array.from(new Set([...Object.keys(baseInputs), ...Object.keys(targetInputs)]))

  for (const k of allInputKeys) {
    const vFrom = baseInputs[k]
    const vTo = targetInputs[k]
    if (vFrom !== vTo) {
      diffs.push({
        key: k,
        fromVal: vFrom !== undefined && vFrom !== null ? vFrom : '(Standard Default)',
        toVal: vTo !== undefined && vTo !== null ? vTo : '(Standard Default)',
        type: 'input',
      })
    }
  }

  const baseFormulas = baseTpl?.definition?.formula_overrides || {}
  const targetFormulas = targetTpl?.definition?.formula_overrides || {}
  const allFormulaKeys = Array.from(new Set([...Object.keys(baseFormulas), ...Object.keys(targetFormulas)]))

  for (const k of allFormulaKeys) {
    const fFrom = baseFormulas[k]
    const fTo = targetFormulas[k]
    if (fFrom !== fTo) {
      diffs.push({
        key: k,
        fromVal: fFrom || '(Standard Formula)',
        toVal: fTo || '(Standard Formula)',
        type: 'formula',
      })
    }
  }

  const baseType = baseTpl?.definition?.baseline || (baseTpl?.name?.toLowerCase().includes('cut-to-cut') ? 'c2c' : 'actual')
  const targetType = targetTpl?.definition?.baseline || (targetTpl?.name?.toLowerCase().includes('cut-to-cut') ? 'c2c' : 'actual')
  if (baseTpl && baseType !== targetType) {
    diffs.push({
      key: 'Baseline Architecture',
      fromVal: baseType === 'c2c' ? 'Cut-to-Cut Costing' : 'Actual Quotation',
      toVal: targetType === 'c2c' ? 'Cut-to-Cut Costing' : 'Actual Quotation',
      type: 'baseline',
    })
  }

  return diffs
}


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
  const [diffModalTemplate, setDiffModalTemplate] = useState<Template | null>(null)
  const [diffComparisonMode, setDiffComparisonMode] = useState<'prev' | 'initial'>('prev')
  const [editingTemplate, setEditingTemplate] = useState<Template | null>(null)

  // Dedicated workbook instances for template draft calculation (all 15 sheets initialized)
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

  // Open Template Studio (Edit existing version in-place or publish new version)
  const handleOpenStudio = (baseTemplate?: Template, forceBaseline?: 'actual' | 'c2c') => {
    const tpl = baseTemplate || currentDefaultTemplate
    setEditingTemplate(tpl || null)

    const initialInputs = tpl?.definition?.default_inputs || {}
    const initialFormulas = tpl?.definition?.formula_overrides || {}
    const isC2C = forceBaseline === 'c2c' || tpl?.definition?.baseline === 'c2c' || tpl?.name?.toLowerCase().includes('cut-to-cut')

    setTemplateBaseline(isC2C ? 'c2c' : 'actual')

    setDraftName(tpl?.name || (isC2C ? 'Cut-to-Cut Standards' : 'Default Production Template'))
    setDraftDesc(tpl?.description || '')

    if (isC2C) {
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


  // Permanently delete a non-default template version
  const handleDeleteTemplate = async (t: Template) => {
    if (!isAdminOrOwner) {
      return toast('Only organisation owners and administrators can delete templates.', 'err')
    }
    if (t.is_default) {
      return toast('Cannot delete the active default template. Please set another template as default first.', 'err')
    }
    const ok = await confirm({
      title: `Delete Template Version ${t.version_num}?`,
      body: `Are you sure you want to delete "${t.name}" (v${t.version_num})? This action cannot be undone.`,
      confirmLabel: 'Delete Version',
      danger: true,
    })
    if (!ok) return

    setLoading(true)
    try {
      await deleteTemplate(session, org, t.id)
      toast(`Template Version ${t.version_num} has been deleted.`)
      await loadTemplates()
      onOrgUpdated()
    } catch (e) {
      toast((e as Error).message, 'err')
    } finally {
      setLoading(false)
    }
  }

  // Save template: either in-place to current version (v1) or as brand new version (v2)
  const handleSaveTemplate = async (asNewVersion = false) => {
    if (!isAdminOrOwner) {
      return toast('Only organisation owners and administrators can save or publish templates.', 'err')
    }
    const nm = draftName.trim()
    if (!nm) return toast('Template name cannot be empty', 'err')

    const steelRate = draftWb.num('Total cal', 'H3') || 90
    const trolleyRate = draftWb.num('Total cal', 'M3') || 90
    const bodyLabourRate = draftWb.num('Total cal', 'B49') || 250

    if (steelRate <= 0 || !Number.isFinite(steelRate)) {
      return toast('Steel rate must be a valid positive number.', 'err')
    }
    if (trolleyRate <= 0 || !Number.isFinite(trolleyRate)) {
      return toast('Trolley rate must be a valid positive number.', 'err')
    }
    if (bodyLabourRate <= 0 || !Number.isFinite(bodyLabourRate)) {
      return toast('Body labour rate must be a valid positive number.', 'err')
    }

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

    const currentVer = editingTemplate?.version_num || 1
    const targetVer = asNewVersion ? nextVersion : currentVer

    if (asNewVersion) {
      const ok = await confirm({
        title: `Publish as New Template Version ${nextVersion}?`,
        body: `This will create Version ${nextVersion} as a separate published version and set it as default for newly created projects in ${org.name}.`,
        confirmLabel: `Publish Version ${nextVersion}`,
      })
      if (!ok) return
    } else {
      const ok = await confirm({
        title: `Save Changes to Version ${currentVer}?`,
        body: `This will save your updates directly to Version ${currentVer} without creating a new version.`,
        confirmLabel: `Save Version ${currentVer}`,
      })
      if (!ok) return
    }

    setPublishing(true)
    try {
      const definition = {
        version: targetVer,
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

      if (asNewVersion || !editingTemplate?.id) {
        const res = await publishTemplate(session, org, nm, draftDesc, definition)
        toast(`Successfully published new Template Version ${res.version_num}!`)
      } else {
        await updateTemplate(session, org, editingTemplate.id, nm, draftDesc, definition)
        toast(`Successfully saved changes directly to Version ${currentVer}! (No new version created)`)
      }

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
            <span>1. ACTUAL TEMPLATE</span>
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
            <span>2. CUT-TO-CUT TEMPLATE</span>
            {templateBaseline === 'c2c' && <span className="chip" style={{ background: '#ede9fe', color: '#6d28d9', fontSize: '11px', padding: '2px 6px' }}>Active</span>}
          </button>
        </div>

        {/* Studio Top Control Bar */}
        <div className="card studio-header mb-3">
          <div className="row between mb-2">
            <button className="btn ghost sm" onClick={() => setEditorMode(false)}>
              ← Back to Organisation Settings
            </button>
            <div className="row gap" style={{ alignItems: 'center' }}>
              <span className="badge-lg">
                {editingTemplate ? `Editing Version ${editingTemplate.version_num}` : `Publishing Version ${nextVersion}`}
              </span>
              <span className="chip" style={{ background: templateBaseline === 'c2c' ? '#ede9fe' : '#dbeafe', color: templateBaseline === 'c2c' ? '#6d28d9' : '#1e40af', fontWeight: 600 }}>
                {templateBaseline === 'c2c' ? '✂️ Cut-to-Cut Baseline' : '📊 Actual Quotation Baseline'}
              </span>
              {editingTemplate?.is_default && (
                <span className="chip ok" style={{ fontSize: '11px', fontWeight: 600 }}>Active Default</span>
              )}
            </div>
            <div className="row gap" style={{ alignItems: 'center' }}>
              {editingTemplate && (
                <button
                  type="button"
                  className="btn primary"
                  disabled={publishing || !draftName.trim()}
                  onClick={() => handleSaveTemplate(false)}
                  title={`Save changes directly to Version ${editingTemplate.version_num} without creating a new version`}
                  style={{ fontWeight: 700 }}
                >
                  {publishing ? <Spinner size={16} /> : <span>💾 Save to v{editingTemplate.version_num}</span>}
                </button>
              )}
              <button
                type="button"
                className={editingTemplate ? "btn ghost" : "btn primary"}
                disabled={publishing || !draftName.trim()}
                onClick={() => handleSaveTemplate(true)}
                title={`Publish as a brand new Template Version (Version ${nextVersion})`}
                style={editingTemplate ? { borderColor: '#2563eb', color: '#1d4ed8', fontWeight: 600 } : { fontWeight: 700 }}
              >
                {publishing && !editingTemplate ? <Spinner size={16} /> : <span>➕ Publish as New Version (v{nextVersion})</span>}
              </button>
            </div>
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
                              ✂️ Cut-to-Cut
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
                              ✏️ Edit Version
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

      {/* Version Differences & Changelog Modal */}
      {diffModalTemplate && (() => {
        const t = diffModalTemplate
        const prevT = templates.find((p) => p.version_num === t.version_num - 1) || templates.filter(p => p.version_num < t.version_num)[0] || null
        const initialT = templates.find((p) => p.version_num === 1) || templates[templates.length - 1] || null
        const activeBaseT = diffComparisonMode === 'prev' ? prevT : initialT
        const diffs = computeTemplateDiff(activeBaseT, t)

        return (
          <div
            className="modal-backdrop"
            style={{
              position: 'fixed',
              top: 0,
              left: 0,
              right: 0,
              bottom: 0,
              background: 'rgba(15, 23, 42, 0.6)',
              display: 'flex',
              alignItems: 'center',
              justifyContent: 'center',
              zIndex: 9999,
              padding: '16px',
            }}
            onClick={() => setDiffModalTemplate(null)}
          >
            <div
              className="card"
              style={{
                maxWidth: '780px',
                width: '100%',
                maxHeight: '90vh',
                overflowY: 'auto',
                padding: '24px',
                background: '#ffffff',
                borderRadius: '12px',
                boxShadow: '0 20px 25px -5px rgba(0, 0, 0, 0.2)',
              }}
              onClick={(e) => e.stopPropagation()}
            >
              <div className="row between mb-3" style={{ alignItems: 'flex-start' }}>
                <div>
                  <div className="row gap" style={{ alignItems: 'center', marginBottom: '4px' }}>
                    <span className="version-pill" style={{ fontSize: '13px', padding: '3px 8px' }}>v{t.version_num}</span>
                    <h3 style={{ margin: 0, fontSize: '18px' }}>{t.name} — Differences &amp; Changelog</h3>
                  </div>
                  <p className="tiny muted" style={{ margin: 0 }}>
                    Published by {t.published_by || t.created_by || 'Admin'} {t.published_at ? timeAgo(t.published_at) : ''}
                    {t.description ? ` · ${t.description}` : ''}
                  </p>
                </div>
                <button className="icon-btn" onClick={() => setDiffModalTemplate(null)} title="Close">
                  {Icon.close}
                </button>
              </div>

              {/* Comparison Mode Switcher */}
              <div style={{ display: 'flex', gap: '8px', marginBottom: '16px', background: '#f1f5f9', padding: '4px', borderRadius: '8px' }}>
                <button
                  type="button"
                  className={`btn sm ${diffComparisonMode === 'prev' ? 'primary' : 'ghost'}`}
                  style={{ flex: 1, justifyContent: 'center' }}
                  onClick={() => setDiffComparisonMode('prev')}
                >
                  🔄 Differences from Previous Version {prevT ? `(v${prevT.version_num})` : ''}
                </button>
                <button
                  type="button"
                  className={`btn sm ${diffComparisonMode === 'initial' ? 'primary' : 'ghost'}`}
                  style={{ flex: 1, justifyContent: 'center' }}
                  onClick={() => setDiffComparisonMode('initial')}
                >
                  ⭐ Differences from Initial Baseline {initialT ? `(v${initialT.version_num})` : '(v1)'}
                </button>
              </div>

              {/* Difference Summary Header */}
              <div className="card p-3 mb-3" style={{ background: '#f8fafc', border: '1px solid #e2e8f0' }}>
                <div className="row between" style={{ alignItems: 'center' }}>
                  <span className="tiny font-bold uppercase muted">
                    Comparing <strong>v{t.version_num}</strong> with {activeBaseT ? <strong>v{activeBaseT.version_num} ({activeBaseT.name})</strong> : 'Baseline'}
                  </span>
                  <span className="chip primary" style={{ fontWeight: 600 }}>
                    {diffs.length} Change{diffs.length !== 1 ? 's' : ''} Found
                  </span>
                </div>
              </div>

              {/* Diffs Table */}
              {diffs.length === 0 ? (
                <div className="empty p-4 text-center">
                  <p style={{ margin: 0, fontWeight: 600, color: '#64748b' }}>
                    {diffComparisonMode === 'prev' && !prevT
                      ? 'This is the initial template version (Version 1). There is no previous version to compare against.'
                      : 'No parameter or formula changes recorded between these two versions.'}
                  </p>
                </div>
              ) : (
                <div style={{ overflowX: 'auto', border: '1px solid #e2e8f0', borderRadius: '8px' }}>
                  <table className="tbl" style={{ width: '100%', margin: 0, fontSize: '13px' }}>
                    <thead>
                      <tr style={{ background: '#f8fafc' }}>
                        <th>Worksheet / Parameter</th>
                        <th>Type</th>
                        <th>{diffComparisonMode === 'prev' ? `v${prevT?.version_num || 'Previous'}` : `v${initialT?.version_num || '1'} (Initial)`}</th>
                        <th>v{t.version_num} (This Version)</th>
                        <th>Status</th>
                      </tr>
                    </thead>
                    <tbody>
                      {diffs.map((d) => (
                        <tr key={d.key}>
                          <td>
                            <strong>{d.key}</strong>
                          </td>
                          <td>
                            {d.type === 'input' && <span className="chip ok" style={{ fontSize: '11px' }}>Input Parameter</span>}
                            {d.type === 'formula' && <span className="chip purple" style={{ fontSize: '11px' }}>Formula Rule</span>}
                            {d.type === 'baseline' && <span className="chip primary" style={{ fontSize: '11px' }}>Baseline Mode</span>}
                          </td>
                          <td style={{ color: '#64748b' }}>
                            <code>{String(d.fromVal)}</code>
                          </td>
                          <td style={{ fontWeight: 700, color: '#16a34a' }}>
                            <code>{String(d.toVal)}</code>
                          </td>
                          <td>
                            <span className="chip" style={{ background: '#ecfdf5', color: '#047857', borderColor: '#a7f3d0', fontSize: '11px' }}>
                              Modified
                            </span>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              )}

              <div className="row end mt-3">
                <button className="btn ghost sm" onClick={() => setDiffModalTemplate(null)}>
                  Close Differences
                </button>
              </div>
            </div>
          </div>
        )
      })()}
    </div>
  )
}
