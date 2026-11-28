import { useEffect, useMemo, useState } from 'react'
import data from '../../lib/clientWorkbook.json'
import c2cData from '../../lib/clientWorkbook_c2c.json'
import { api, Org, Session, Template } from '../../lib/api'
import { deleteTemplate, publishTemplate, rollbackTemplateToVersion, updateTemplate } from '../../lib/permissions'
import { Overrides, SheetData, Workbook } from '../../lib/engine'
import { inr, timeAgo } from '../../lib/format'
import { INPUT_GROUPS } from '../../lib/inputs'
import { filterCutToCutOverrides, summarize, Summary } from '../../lib/model'
import { Icon, NumField, Spinner, useUI } from '../ui'
import { CalculationDetails } from './CalculationDetails'

interface Props {
  session: Session
  org: Org
  onOrgUpdated: () => void | Promise<void>
}

type StudioTab = 'forms' | 'sheets' | 'diff'

export interface DiffEntry {
  key: string
  label?: string
  fromVal: any
  toVal: any
  type: 'input' | 'formula' | 'baseline'
}

const INPUT_LABEL_MAP: Record<string, string> = (() => {
  const map: Record<string, string> = {}
  for (const grp of INPUT_GROUPS) {
    for (const item of grp.items) {
      map[item.key] = item.unit ? `${item.label} (${item.unit})` : item.label
    }
  }
  return map
})()

const LEGACY_MIRROR_KEYS = new Set(['rates!b3', 'rates!b4'])

function cleanTemplateDefaultInputs(raw?: Record<string, any> | null): Overrides {
  const out: Overrides = {}
  if (!raw || typeof raw !== 'object') return out
  for (const [k, v] of Object.entries(raw)) {
    if (LEGACY_MIRROR_KEYS.has(k)) continue
    if (v === null || typeof v === 'number') out[k] = v
    else if (v !== undefined && !isNaN(Number(v))) out[k] = Number(v)
  }
  return out
}

function getTemplateBaseline(tpl?: Template | null): 'actual' | 'c2c' {
  if (!tpl) return 'actual'
  if (tpl.definition?.baseline === 'c2c') return 'c2c'
  if (tpl.name?.toLowerCase().includes('cut-to-cut')) return 'c2c'
  return 'actual'
}

function readTemplateRates(tpl?: Template | null) {
  const def: any = tpl?.definition || {}
  const inputs = def.default_inputs || {}
  const isC2C = getTemplateBaseline(tpl) === 'c2c'
  const steel = Number(inputs['Total cal!H3'] ?? def.rates?.steel ?? (isC2C ? 79 : 90))
  const trolley = Number(inputs['Total cal!M3'] ?? def.rates?.trolleySteel ?? (isC2C ? 69 : 90))
  const labour = Number(inputs['Total cal!B49'] ?? def.rates?.bodyLabour ?? 250)
  return {
    steel: isNaN(steel) ? (isC2C ? 79 : 90) : steel,
    trolley: isNaN(trolley) ? (isC2C ? 69 : 90) : trolley,
    labour: isNaN(labour) ? 250 : labour,
  }
}

function stripVersionSuffix(name: string): string {
  return name.replace(/(?:\s*\(V\d+\)|\s+v\d+)$/gi, '').trim()
}

export function computeTemplateDiff(baseTpl?: Template | null, targetTpl?: Template | null): DiffEntry[] {
  if (!targetTpl) return []
  const diffs: DiffEntry[] = []

  const baseInputs = cleanTemplateDefaultInputs(baseTpl?.definition?.default_inputs)
  const targetInputs = cleanTemplateDefaultInputs(targetTpl?.definition?.default_inputs)
  const allInputKeys = Array.from(new Set([...Object.keys(baseInputs), ...Object.keys(targetInputs)]))

  for (const k of allInputKeys) {
    const vFrom = baseInputs[k]
    const vTo = targetInputs[k]
    if (vFrom !== vTo) {
      diffs.push({
        key: k,
        label: INPUT_LABEL_MAP[k] || k,
        fromVal: vFrom !== undefined && vFrom !== null ? vFrom : '(Default)',
        toVal: vTo !== undefined && vTo !== null ? vTo : '(Default)',
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
        label: `Formula · ${k}`,
        fromVal: fFrom || '(Default Formula)',
        toVal: fTo || '(Default Formula)',
        type: 'formula',
      })
    }
  }

  const baseType = getTemplateBaseline(baseTpl)
  const targetType = getTemplateBaseline(targetTpl)
  if (baseTpl && baseType !== targetType) {
    diffs.push({
      key: 'Baseline Mode',
      label: 'Baseline Mode',
      fromVal: baseType === 'c2c' ? 'Cut-to-Cut' : 'Standard Actual',
      toVal: targetType === 'c2c' ? 'Cut-to-Cut' : 'Standard Actual',
      type: 'baseline',
    })
  }

  return diffs
}

export function OrganisationSettings({ session, org, onOrgUpdated }: Props) {
  const { toast, confirm } = useUI()

  const [templates, setTemplates] = useState<Template[]>([])
  const [loading, setLoading] = useState(true)
  const [publishing, setPublishing] = useState(false)
  const [worksheetsReady, setWorksheetsReady] = useState(false)

  // Studio / Editor state
  const [editorMode, setEditorMode] = useState(false)
  const [studioTab, setStudioTab] = useState<StudioTab>('forms')
  const [draftName, setDraftName] = useState('')
  const [draftDesc, setDraftDesc] = useState('')
  const [draftOverrides, setDraftOverrides] = useState<Overrides>({})
  const [draftFormulaOverrides, setDraftFormulaOverrides] = useState<Record<string, string>>({})
  const [editingTemplate, setEditingTemplate] = useState<Template | null>(null)

  // Diff modal state
  const [diffModalTemplate, setDiffModalTemplate] = useState<Template | null>(null)
  const [diffComparisonMode, setDiffComparisonMode] = useState<'prev' | 'initial'>('prev')

  const draftActualWb = useMemo(() => new Workbook(data as unknown as { sheets: SheetData[] }), [])
  const draftC2CWb = useMemo(() => new Workbook(c2cData as unknown as { sheets: SheetData[] }), [])

  const isAdminOrOwner = org.role === 'owner' || org.role === 'admin'

  // Load real full 15 worksheets into draftActualWb & draftC2CWb so Actual & C2C previews match Excel
  useEffect(() => {
    if (!isAdminOrOwner) return
    let active = true
    api.getFullWorksheets(session.token, org.id)
      .then((res) => {
        if (!active) return
        if (res.actual_workbook?.sheets) {
          draftActualWb.loadSheets(res.actual_workbook.sheets, res.named_constants)
        }
        if (res.c2c_workbook?.sheets) {
          draftC2CWb.loadSheets(res.c2c_workbook.sheets, res.named_constants)
        }
        setWorksheetsReady(true)
      })
      .catch(() => {
        // Fallback silently if offline
      })
    return () => {
      active = false
    }
  }, [isAdminOrOwner, session.token, org.id, draftActualWb, draftC2CWb])

  const loadTemplates = async () => {
    try {
      setLoading(true)
      const list = await api.templateList(session.token, org.id)
      setTemplates(Array.isArray(list) ? list : [])
    } catch (e) {
      toast((e as Error).message, 'err')
    } finally {
      setLoading(false)
    }
  }

  useEffect(() => {
    loadTemplates()
  }, [org.id])

  const publishedTemplates = useMemo(
    () => templates.filter((t) => t.status === 'published'),
    [templates],
  )
  const currentDefaultTemplate = useMemo(
    () =>
      templates.find((t) => t.is_default) ||
      templates.find((t) => t.version_num === org.current_template_version) ||
      publishedTemplates[0] ||
      templates[0],
    [templates, publishedTemplates, org.current_template_version],
  )

  // Match Postgres template_publish: max(version_num) + 1 across all org templates
  const nextVersion = useMemo(() => {
    const maxV = Math.max(
      org.current_template_version || 1,
      ...templates.map((t) => t.version_num || 0),
    )
    return maxV + 1
  }, [org.current_template_version, templates])

  const handleOpenStudio = (baseTemplate?: Template, createNewMode = false) => {
    const tpl = baseTemplate || currentDefaultTemplate
    const initialInputs = cleanTemplateDefaultInputs(tpl?.definition?.default_inputs)
    const initialFormulas = { ...(tpl?.definition?.formula_overrides || {}) }

    if (tpl) {
      const rates = readTemplateRates(tpl)
      if (initialInputs['Total cal!H3'] === undefined) initialInputs['Total cal!H3'] = rates.steel
      if (initialInputs['Total cal!M3'] === undefined) initialInputs['Total cal!M3'] = rates.trolley
      if (initialInputs['Total cal!B49'] === undefined) initialInputs['Total cal!B49'] = rates.labour
    }

    setEditingTemplate(createNewMode ? null : (tpl || null))

    if (createNewMode) {
      const cleanBase = stripVersionSuffix(tpl?.name || 'Production Standards') || 'Production Standards'
      setDraftName(`${cleanBase} v${nextVersion}`)
      setDraftDesc(tpl?.description || '')
    } else {
      setDraftName(tpl?.name || `Production Standards v${nextVersion}`)
      setDraftDesc(tpl?.description || '')
    }

    setDraftOverrides(initialInputs)
    setDraftFormulaOverrides(initialFormulas)

    draftActualWb.setTemplateDefaults({})
    draftActualWb.setFormulaOverrides(initialFormulas)
    draftActualWb.setOverrides(initialInputs)

    draftC2CWb.setTemplateDefaults({})
    draftC2CWb.setFormulaOverrides(initialFormulas)
    draftC2CWb.setOverrides(filterCutToCutOverrides(initialInputs))

    setStudioTab('forms')
    setEditorMode(true)
  }

  const handleDraftSetInput = (key: string, value: number | null | undefined) => {
    setDraftOverrides((prev) => {
      const next = { ...prev }
      if (value === undefined || value === null) delete next[key]
      else next[key] = value
      return next
    })
  }

  const handleDraftSetFormulaOverride = (key: string, formula: string | undefined) => {
    setDraftFormulaOverrides((prev) => {
      const next = { ...prev }
      if (!formula || formula.trim() === '') delete next[key]
      else next[key] = formula
      return next
    })
  }

  const draftSummary: Summary = useMemo(() => {
    draftActualWb.setTemplateDefaults({})
    draftActualWb.setFormulaOverrides(draftFormulaOverrides)
    draftActualWb.setOverrides(draftOverrides)
    return summarize(draftActualWb)
  }, [draftOverrides, draftFormulaOverrides, draftActualWb, worksheetsReady])

  const draftC2CSummary: Summary = useMemo(() => {
    draftC2CWb.setTemplateDefaults({})
    draftC2CWb.setFormulaOverrides(draftFormulaOverrides)
    draftC2CWb.setOverrides(filterCutToCutOverrides(draftOverrides))
    return summarize(draftC2CWb)
  }, [draftOverrides, draftFormulaOverrides, draftC2CWb, worksheetsReady])

  const buildStudioDefinition = (targetVer: number, nm: string) => {
    draftActualWb.setTemplateDefaults({})
    draftActualWb.setFormulaOverrides(draftFormulaOverrides)
    draftActualWb.setOverrides(draftOverrides)

    const steelRate = Number(
      draftOverrides['Total cal!H3'] ?? (draftActualWb.num('Total cal', 'H3') || 90),
    )
    const trolleyRate = Number(
      draftOverrides['Total cal!M3'] ?? (draftActualWb.num('Total cal', 'M3') || 90),
    )
    const bodyLabourRate = Number(
      draftOverrides['Total cal!B49'] ?? (draftActualWb.num('Total cal', 'B49') || 250),
    )

    const cleanedInputs = cleanTemplateDefaultInputs(draftOverrides)
    cleanedInputs['Total cal!H3'] = steelRate
    cleanedInputs['Total cal!M3'] = trolleyRate
    cleanedInputs['Total cal!B49'] = bodyLabourRate

    return {
      definition: {
        version: targetVer,
        name: nm,
        description: draftDesc.trim(),
        baseline: 'actual',
        rates: {
          steel: steelRate,
          trolleySteel: trolleyRate,
          bodyLabour: bodyLabourRate,
        },
        default_inputs: cleanedInputs,
        formula_overrides: { ...draftFormulaOverrides },
      },
      rates: { steelRate, trolleyRate, bodyLabourRate },
    }
  }

  const canDeleteTemplate = (t: Template): boolean => {
    if (t.status === 'draft') return true
    if (!t.is_default) return true
    // Active default can be deleted as long as at least one other published version remains
    return publishedTemplates.some((other) => other.id !== t.id)
  }

  const handleSetActiveDefault = async (t: Template) => {
    if (t.is_default) return
    const ok = await confirm({
      title: `Activate Template v${t.version_num}?`,
      body: `New quotations will use "${t.name}" (v${t.version_num}) by default. Existing quotations keep their saved version.`,
      confirmLabel: `Set v${t.version_num} as Default`,
    })
    if (!ok) return
    try {
      await rollbackTemplateToVersion(session, org, t.id)
      toast(`Template v${t.version_num} is now active.`)
      await loadTemplates()
      await onOrgUpdated()
    } catch (e) {
      toast((e as Error).message, 'err')
    }
  }

  const handlePublishDraftRow = async (t: Template) => {
    const ok = await confirm({
      title: `Publish Draft "${t.name}"?`,
      body: `Publishes this draft as Template v${nextVersion} and makes it the active default.`,
      confirmLabel: `Publish v${nextVersion}`,
    })
    if (!ok) return
    try {
      const def = {
        ...(t.definition || {}),
        version: nextVersion,
        name: t.name,
        description: t.description || '',
      }
      const res = await publishTemplate(session, org, t.name, t.description || '', def, t.id)
      toast(`Published Template v${res.version_num} as active default.`)
      await loadTemplates()
      await onOrgUpdated()
    } catch (e) {
      toast((e as Error).message, 'err')
    }
  }

  const handleDeleteTemplate = async (t: Template): Promise<boolean> => {
    if (!canDeleteTemplate(t)) {
      toast('Cannot delete the only published template version. Publish a new version first.', 'err')
      return false
    }
    const nextFallback = publishedTemplates.find((other) => other.id !== t.id)
    const extraMsg = t.is_default && nextFallback
      ? ` Active default will automatically switch to v${nextFallback.version_num} (${nextFallback.name}).`
      : ''
    const ok = await confirm({
      title: `Delete Template v${t.version_num}?`,
      body: `Delete "${t.name}" (v${t.version_num}) from template history? Past quotations using v${t.version_num} keep their saved snapshot.${extraMsg}`,
      confirmLabel: 'Delete Version',
      danger: true,
    })
    if (!ok) return false
    try {
      await deleteTemplate(session, org, t.id)
      toast(`Deleted template v${t.version_num}.`)
      setEditorMode(false)
      setEditingTemplate(null)
      await loadTemplates()
      await onOrgUpdated()
      return true
    } catch (e) {
      toast((e as Error).message, 'err')
      return false
    }
  }

  const handleSaveTemplate = async (mode: 'update' | 'publish') => {
    const nm = draftName.trim() || `Production Standards v${nextVersion}`
    const targetVer = mode === 'update' && editingTemplate ? editingTemplate.version_num : nextVersion
    const { definition, rates } = buildStudioDefinition(targetVer, nm)

    if (mode === 'publish') {
      const ok = await confirm({
        title: `Publish Template v${targetVer}?`,
        body: `Sets "${nm}" (Steel ₹${rates.steelRate}/kg · Trolley ₹${rates.trolleyRate}/kg) as active default for new quotations.`,
        confirmLabel: `Publish v${targetVer}`,
      })
      if (!ok) return
    }

    try {
      setPublishing(true)
      if (mode === 'update' && editingTemplate) {
        await updateTemplate(session, org, editingTemplate.id, nm, draftDesc.trim(), definition)
        toast(`Saved changes to v${editingTemplate.version_num}.`)
      } else {
        const res = await publishTemplate(
          session,
          org,
          nm,
          draftDesc.trim(),
          definition,
          editingTemplate?.status === 'draft' ? editingTemplate.id : undefined,
        )
        toast(`Published v${res.version_num} as active default.`)
      }
      setEditorMode(false)
      setEditingTemplate(null)
      await loadTemplates()
      await onOrgUpdated()
    } catch (e) {
      toast((e as Error).message, 'err')
    } finally {
      setPublishing(false)
    }
  }

  const previewDiffEntries = useMemo(() => {
    const baselineTpl = editingTemplate || currentDefaultTemplate
    const pseudoTpl: Template = {
      id: 'draft',
      org_id: org.id,
      version_num: editingTemplate ? editingTemplate.version_num : nextVersion,
      name: draftName,
      description: draftDesc,
      status: 'draft',
      is_default: false,
      definition: {
        version: editingTemplate ? editingTemplate.version_num : nextVersion,
        baseline: 'actual',
        default_inputs: cleanTemplateDefaultInputs(draftOverrides) as any,
        formula_overrides: draftFormulaOverrides,
      },
      created_by: null,
      published_by: null,
      published_at: null,
      created_at: new Date().toISOString(),
      updated_at: new Date().toISOString(),
    } as Template
    return computeTemplateDiff(baselineTpl, pseudoTpl)
  }, [
    editingTemplate,
    currentDefaultTemplate,
    org.id,
    nextVersion,
    draftName,
    draftDesc,
    draftOverrides,
    draftFormulaOverrides,
  ])

  const getPreviousTemplate = (t: Template): Template | undefined => {
    const idx = templates.findIndex((item) => item.id === t.id)
    if (idx === -1 || idx + 1 >= templates.length) return undefined
    return templates[idx + 1]
  }

  const initialV1Template = useMemo(
    () => templates[templates.length - 1],
    [templates],
  )

  const modalDiffEntries = useMemo(() => {
    if (!diffModalTemplate) return []
    const base =
      diffComparisonMode === 'initial'
        ? initialV1Template
        : getPreviousTemplate(diffModalTemplate) || initialV1Template
    if (!base || base.id === diffModalTemplate.id) return []
    return computeTemplateDiff(base, diffModalTemplate)
  }, [diffModalTemplate, diffComparisonMode, templates, initialV1Template])

  if (!isAdminOrOwner) {
    return (
      <div className="org-settings-page">
        <div className="card settings-card">
          <h3>Restricted Access</h3>
          <p className="muted mt-2">
            Only Organisation Owners and Administrators can manage template versions. New quotations automatically use{' '}
            <strong>Template v{org.current_template_version || 1}</strong>.
          </p>
        </div>
      </div>
    )
  }

  // =========================================================================
  // ULTRA-CLEAN MINIMAL TEMPLATE STUDIO VIEW
  // =========================================================================
  if (editorMode) {
    const isEditingExisting = Boolean(editingTemplate && editingTemplate.status === 'published')
    return (
      <div className="template-studio">
        <div className="card tpl-studio-card">
          {/* Compact Top Identity & Action Row */}
          <div className="tpl-studio-top">
            <div className="tpl-studio-identity">
              <button
                className="btn ghost sm"
                onClick={() => {
                  setEditorMode(false)
                  setEditingTemplate(null)
                }}
              >
                ← Back
              </button>
              <span className="version-pill">
                {isEditingExisting ? `v${editingTemplate!.version_num}` : `New v${nextVersion}`}
              </span>
              <div className="tpl-studio-meta">
                <input
                  className="tpl-title-input"
                  value={draftName}
                  onChange={(e) => setDraftName(e.target.value)}
                  placeholder="Template name"
                />
                <input
                  className="tpl-desc-input"
                  value={draftDesc}
                  onChange={(e) => setDraftDesc(e.target.value)}
                  placeholder="Optional release note"
                />
              </div>
            </div>

            <div className="tpl-studio-actions">
              {editingTemplate && canDeleteTemplate(editingTemplate) && (
                <button
                  className="btn ghost sm tpl-del-btn"
                  disabled={publishing}
                  onClick={() => handleDeleteTemplate(editingTemplate)}
                  title={`Delete v${editingTemplate.version_num}`}
                >
                  {Icon.trash} Delete
                </button>
              )}
              {isEditingExisting && (
                <button
                  className="btn ghost sm"
                  disabled={publishing}
                  onClick={() => handleSaveTemplate('update')}
                >
                  Save to v{editingTemplate!.version_num}
                </button>
              )}
              <button
                className="btn primary sm"
                disabled={publishing}
                onClick={() => handleSaveTemplate('publish')}
              >
                {publishing ? 'Publishing…' : `Publish v${nextVersion}`}
              </button>
            </div>
          </div>

          {/* Single Clean Subbar: Tabs + Live Quote Preview */}
          <div className="tpl-studio-subbar">
            <div className="tpl-tabs">
              <button
                className={`tpl-tab-btn ${studioTab === 'forms' ? 'active' : ''}`}
                onClick={() => setStudioTab('forms')}
              >
                Quick Parameters
              </button>
              <button
                className={`tpl-tab-btn ${studioTab === 'sheets' ? 'active' : ''}`}
                onClick={() => setStudioTab('sheets')}
              >
                Workbook &amp; Formulas
              </button>
              <button
                className={`tpl-tab-btn ${studioTab === 'diff' ? 'active' : ''}`}
                onClick={() => setStudioTab('diff')}
              >
                Changes ({previewDiffEntries.length})
              </button>
            </div>

            <div className="tpl-preview-pills">
              <div className="tpl-preview-item">
                <span className="tiny muted">Ex-Works</span>
                <strong className="mono">{worksheetsReady ? inr(draftSummary.subtotal) : '…'}</strong>
              </div>
              <div className="tpl-preview-item">
                <span className="tiny muted">C2C Cost</span>
                <strong className="mono">{worksheetsReady ? inr(draftC2CSummary.subtotal) : '…'}</strong>
              </div>
              <div className="tpl-preview-item">
                <span className="tiny muted">Base Total</span>
                <strong className="mono">{worksheetsReady ? inr(draftSummary.sales) : '…'}</strong>
              </div>
            </div>
          </div>
        </div>

        <div className="studio-content">
          {studioTab === 'forms' && (
            <div className="forms-grid">
              {INPUT_GROUPS.map((group) => (
                <div className="card" key={group.title}>
                  <div className="card-head">
                    <h3>{group.title}</h3>
                    {group.note && <span className="muted tiny">{group.note}</span>}
                  </div>
                  <div className="card-body">
                    {group.items.map((item) => {
                      const val =
                        draftOverrides[item.key] !== undefined
                          ? draftOverrides[item.key]
                          : (() => {
                              const [sh, addr] = item.key.split('!')
                              const numVal = draftActualWb.num(sh, addr)
                              return isNaN(numVal) ? null : numVal
                            })()
                      return (
                        <div className="field" key={item.key}>
                          <label title={item.key}>
                            {item.label}
                            {item.unit ? <span className="unit">{item.unit}</span> : null}
                          </label>
                          <NumField
                            value={val as number | null}
                            step={item.step}
                            min={item.min}
                            onChange={(v) => handleDraftSetInput(item.key, v)}
                          />
                        </div>
                      )
                    })}
                  </div>
                </div>
              ))}
            </div>
          )}

          {studioTab === 'sheets' && (
            <CalculationDetails
              wb={draftActualWb}
              c2cWb={draftC2CWb}
              overrides={draftOverrides}
              formulaOverrides={draftFormulaOverrides}
              c2cOverrides={filterCutToCutOverrides(draftOverrides)}
              c2cFormulaOverrides={draftFormulaOverrides}
              setInput={(k, v, targetWb) => {
                if (targetWb !== 'c2c') handleDraftSetInput(k, v)
              }}
              setFormulaOverride={(k, f, targetWb) => {
                if (targetWb !== 'c2c') handleDraftSetFormulaOverride(k, f)
              }}
              isAdminOrOwner={true}
              session={session}
              org={org}
              initialWorkbookType="actual"
              hideExcelDownload={true}
            />
          )}

          {studioTab === 'diff' && (
            <div className="card">
              <div className="card-head">
                <h3>Pending Changes</h3>
                <span className="muted tiny">
                  Compared against {editingTemplate ? `v${editingTemplate.version_num}` : 'active default'}
                </span>
              </div>
              <div className="card-body">
                {previewDiffEntries.length === 0 ? (
                  <p className="muted tiny">No parameter or formula changes yet.</p>
                ) : (
                  <table className="tpl-diff-table">
                    <thead>
                      <tr>
                        <th>Parameter</th>
                        <th>Previous</th>
                        <th>Updated</th>
                      </tr>
                    </thead>
                    <tbody>
                      {previewDiffEntries.map((d) => (
                        <tr key={d.key}>
                          <td>
                            <div className="font-semibold">{d.label || d.key}</div>
                            {d.label && d.label !== d.key && (
                              <div className="tiny muted font-mono">{d.key}</div>
                            )}
                          </td>
                          <td className="mono muted">{String(d.fromVal)}</td>
                          <td className="mono font-bold text-navy">{String(d.toVal)}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                )}
              </div>
            </div>
          )}
        </div>
      </div>
    )
  }

  // =========================================================================
  // MINIMAL TEMPLATE VERSIONS LIST VIEW
  // =========================================================================
  const activeRates = readTemplateRates(currentDefaultTemplate)
  const activeCustomInputsCount = currentDefaultTemplate
    ? Object.keys(cleanTemplateDefaultInputs(currentDefaultTemplate.definition?.default_inputs)).filter(
        (k) => !['Total cal!H3', 'Total cal!M3', 'Total cal!B49'].includes(k),
      ).length
    : 0
  const activeFormulasCount = Object.keys(
    currentDefaultTemplate?.definition?.formula_overrides || {},
  ).length

  return (
    <div className="org-settings-page">
      {/* Page Header */}
      <div className="page-head">
        <div>
          <h2>Template Versions</h2>
          <p className="muted">
            Standard rates and calculation rules used automatically for new quotations.
          </p>
        </div>
        <div className="row gap-2">
          {currentDefaultTemplate && (
            <button
              className="btn ghost"
              onClick={() => handleOpenStudio(currentDefaultTemplate, false)}
            >
              {Icon.edit} Edit v{currentDefaultTemplate.version_num}
            </button>
          )}
          <button
            className="btn primary"
            onClick={() => handleOpenStudio(currentDefaultTemplate, true)}
          >
            {Icon.plus} New Version
          </button>
        </div>
      </div>

      {/* Single Clean Active Default Card */}
      <div className="card tpl-active-card">
        <div className="tpl-active-head">
          <div className="row gap-2" style={{ flexWrap: 'wrap' }}>
            <span className="tpl-status-pill active">
              Active Default · v{currentDefaultTemplate?.version_num ?? org.current_template_version ?? 1}
            </span>
            <strong style={{ fontSize: 15 }}>
              {currentDefaultTemplate?.name || 'Default Production Template'}
            </strong>
          </div>
          <span className="tiny muted">
            {currentDefaultTemplate?.published_at
              ? `Updated ${timeAgo(currentDefaultTemplate.published_at)}`
              : 'System baseline'}
          </span>
        </div>

        {currentDefaultTemplate?.description && (
          <p className="tiny muted" style={{ margin: '4px 0 0' }}>
            {currentDefaultTemplate.description}
          </p>
        )}

        <div className="tpl-metrics-strip">
          <div className="tpl-metric">
            <span className="tiny muted">Steel Rate</span>
            <strong className="mono">₹{activeRates.steel}/kg</strong>
          </div>
          <div className="tpl-metric">
            <span className="tiny muted">Trolley Rate</span>
            <strong className="mono">₹{activeRates.trolley}/kg</strong>
          </div>
          <div className="tpl-metric">
            <span className="tiny muted">Body Labour</span>
            <strong className="mono">₹{activeRates.labour}</strong>
          </div>
          <div className="tpl-metric">
            <span className="tiny muted">Custom Rules</span>
            <strong className="mono">
              {activeCustomInputsCount + activeFormulasCount}
            </strong>
          </div>
        </div>

        <div className="tpl-active-foot tiny muted">
          Past quotations keep their original template snapshot. Internal Cut-to-Cut costing stays locked at ₹79/kg steel and ₹69/kg trolley.
        </div>
      </div>

      {/* Minimal Version History Table */}
      <div className="card">
        <div className="card-head">
          <h3>Version History ({templates.length})</h3>
          <span className="tiny muted">Switch default or manage published versions</span>
        </div>

        {loading ? (
          <div className="center-pad">
            <Spinner />
          </div>
        ) : templates.length === 0 ? (
          <div className="card-body">
            <p className="muted">No template versions saved yet. Click New Version to create v2.</p>
          </div>
        ) : (
          <div className="tpl-table-wrap">
            <table className="tpl-table">
              <thead>
                <tr>
                  <th>Version</th>
                  <th>Template</th>
                  <th>Default Rates</th>
                  <th>Quotes</th>
                  <th style={{ textAlign: 'right' }}>Actions</th>
                </tr>
              </thead>
              <tbody>
                {templates.map((t) => {
                  const rates = readTemplateRates(t)
                  const inputsMap = cleanTemplateDefaultInputs(t.definition?.default_inputs)
                  const formulasMap = t.definition?.formula_overrides || {}
                  const extraInputs = Object.keys(inputsMap).filter(
                    (k) => !['Total cal!H3', 'Total cal!M3', 'Total cal!B49'].includes(k),
                  ).length
                  const extraRules = extraInputs + Object.keys(formulasMap).length
                  const hasPrev = Boolean(getPreviousTemplate(t))
                  const deletable = canDeleteTemplate(t)

                  return (
                    <tr key={t.id} className={t.is_default ? 'tpl-row-active' : ''}>
                      <td>
                        <div className="row gap-2">
                          <span className="version-pill">v{t.version_num}</span>
                          {t.is_default && <span className="tpl-tag green">Active</span>}
                          {t.status === 'draft' && <span className="tpl-tag amber">Draft</span>}
                        </div>
                      </td>

                      <td>
                        <div className="font-semibold" style={{ color: 'var(--navy-950)' }}>
                          {t.name}
                        </div>
                        {t.description ? (
                          <div className="tiny muted">{t.description}</div>
                        ) : (
                          <div className="tiny muted">
                            {t.published_by ? `By ${t.published_by} · ` : ''}
                            {timeAgo(t.published_at || t.created_at)}
                          </div>
                        )}
                      </td>

                      <td>
                        <div className="tpl-rate-chips">
                          <span className="tpl-rate-chip mono">Steel ₹{rates.steel}</span>
                          <span className="tpl-rate-chip mono">Trolley ₹{rates.trolley}</span>
                          <span className="tpl-rate-chip mono">Labour ₹{rates.labour}</span>
                          {extraRules > 0 && (
                            <span className="tpl-rate-chip accent">+{extraRules} custom</span>
                          )}
                        </div>
                      </td>

                      <td className="tiny muted mono">{t.project_count || 0}</td>

                      <td style={{ textAlign: 'right' }}>
                        <div className="tpl-row-actions">
                          {t.status === 'draft' && (
                            <button
                              className="btn primary sm"
                              onClick={() => handlePublishDraftRow(t)}
                            >
                              Publish
                            </button>
                          )}
                          {!t.is_default && t.status === 'published' && (
                            <button
                              className="btn ghost sm"
                              onClick={() => handleSetActiveDefault(t)}
                            >
                              Set Default
                            </button>
                          )}
                          <button
                            className="btn ghost sm"
                            onClick={() => handleOpenStudio(t, false)}
                            title={`Edit template v${t.version_num}`}
                          >
                            Edit
                          </button>
                          {hasPrev && (
                            <button
                              className="btn ghost sm"
                              onClick={() => {
                                setDiffComparisonMode('prev')
                                setDiffModalTemplate(t)
                              }}
                              title="Compare changes"
                            >
                              Diff
                            </button>
                          )}
                          {deletable && (
                            <button
                              className="btn ghost sm tpl-del-btn"
                              onClick={() => handleDeleteTemplate(t)}
                              title={`Delete v${t.version_num}`}
                            >
                              {Icon.trash} Delete
                            </button>
                          )}
                        </div>
                      </td>
                    </tr>
                  )
                })}
              </tbody>
            </table>
          </div>
        )}
      </div>

      {/* Minimal Compare Diff Modal */}
      {diffModalTemplate && (
        <div className="overlay" onMouseDown={() => setDiffModalTemplate(null)}>
          <div className="dialog" style={{ maxWidth: 640 }} onMouseDown={(e) => e.stopPropagation()}>
            <div className="dialog-header" style={{ justifyContent: 'space-between' }}>
              <div className="row gap-2">
                {Icon.history}
                <h3>Template v{diffModalTemplate.version_num} Changes</h3>
              </div>
              <button
                className="btn ghost sm"
                onClick={() => setDiffModalTemplate(null)}
              >
                {Icon.close}
              </button>
            </div>

            <div className="row gap-2 mt-3" style={{ marginBottom: 12 }}>
              <button
                className={`btn sm ${diffComparisonMode === 'prev' ? 'primary' : 'ghost'}`}
                onClick={() => setDiffComparisonMode('prev')}
              >
                vs Previous Version
              </button>
              <button
                className={`btn sm ${diffComparisonMode === 'initial' ? 'primary' : 'ghost'}`}
                onClick={() => setDiffComparisonMode('initial')}
              >
                vs Initial Baseline
              </button>
            </div>

            {modalDiffEntries.length === 0 ? (
              <p className="muted tiny" style={{ padding: '12px 0' }}>
                No rate or rule differences recorded against the comparison version.
              </p>
            ) : (
              <table className="tpl-diff-table">
                <thead>
                  <tr>
                    <th>Parameter</th>
                    <th>Previous</th>
                    <th>v{diffModalTemplate.version_num} Value</th>
                  </tr>
                </thead>
                <tbody>
                  {modalDiffEntries.map((d) => (
                    <tr key={d.key}>
                      <td>
                        <div className="font-semibold">{d.label || d.key}</div>
                        {d.label && d.label !== d.key && (
                          <div className="tiny muted font-mono">{d.key}</div>
                        )}
                      </td>
                      <td className="mono muted">{String(d.fromVal)}</td>
                      <td className="mono font-bold text-navy">{String(d.toVal)}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            )}

            <div className="dialog-foot mt-4">
              <button className="btn primary sm" onClick={() => setDiffModalTemplate(null)}>
                Close
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  )
}
