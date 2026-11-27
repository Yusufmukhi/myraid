import { CELL_DESCRIPTIONS, cleanDiffText, getInitialCellValue, ValueChange } from './revision'

const CLIENT_FIELD_LABELS: Record<string, string> = {
  qtn_no: 'QTN Number',
  date: 'Quotation Date',
  buyer_name: 'Buyer / Company Name',
  buyer_address: 'Buyer Address',
  buyer_gst: 'Buyer GSTIN',
  delivery: 'Delivery Time',
  buyer_email: 'Buyer Email',
  buyer_contact: 'Buyer Phone',
  kind_att: 'Kind Attention',
  subject: 'PDF Subject',
  selected_contact_id: 'Selected Contact',
  installation_terms_text: 'Installation Term',
  freight_terms: 'Freight Term',
  payment_terms: 'Payment Term',
  unloading_terms: 'Unloading Term',
  packing_terms: 'Packing Term',
  color_terms: 'Powder Coat Color Term',
  validity_terms: 'Validity Term',
  tax_terms: 'Tax Term',
  lodging_terms: 'Lodging & Boarding Term',
  warranty_terms: 'Warranty Term',
  amc_terms: 'AMC Term',
}

const VISIBILITY_TOGGLE_LABELS: Record<string, string> = {
  show_buyer_name: 'Show Buyer Name on PDF',
  show_buyer_address: 'Show Buyer Address on PDF',
  show_buyer_gst: 'Show GSTIN on PDF',
  show_kind_att: 'Show Kind Attn on PDF',
  show_buyer_contact: 'Show Phone on PDF',
  show_buyer_email: 'Show Email on PDF',
  show_subject: 'Show Subject on PDF',
}

function fmtShortVal(v: any): string {
  if (v === null || v === undefined || v === '') return '—'
  const s = String(v).trim().replace(/\s+/g, ' ')
  return s.length > 80 ? `${s.slice(0, 77)}…` : s
}

function fmtCellLabel(key: string): string {
  if (CELL_DESCRIPTIONS[key]) return CELL_DESCRIPTIONS[key]
  if (key.includes('!')) return `Sheet ${key}`
  return key
}

function pushDiff(
  out: ValueChange[],
  key: string,
  label: string,
  oldVal: any,
  newVal: any,
  prefix = '',
) {
  const fullLabel = prefix ? `${prefix} ${label}` : label
  const oStr = fmtShortVal(oldVal)
  const nStr = fmtShortVal(newVal)
  if (oStr === nStr) return
  out.push({
    key,
    label: fullLabel,
    oldVal: oStr,
    newVal: nStr,
    diff: cleanDiffText(`${fullLabel}: ${oStr} → ${nStr}`),
  })
}

export interface MultiSystemDiffInput {
  name?: string
  status?: string
  quotationMode?: string
  profitPct?: number
  discountPct?: number
  includeInstallation?: boolean
  extraAmount?: number
  notes?: string
  description?: string
  reminder?: { date?: string; note?: string; dismissed?: boolean } | null
  total?: number | null
  inputs?: Record<string, any>
  formulaOverrides?: Record<string, string>
  c2cOverrides?: Record<string, any>
  c2cFormulaOverrides?: Record<string, string>
  systems?: Array<{
    id?: string
    name?: string
    quantity?: number
    unit?: string
    bay?: number
    includeInstallation?: boolean
    extraAmount?: number
    overrides?: Record<string, any>
    formulaOverrides?: Record<string, string>
    c2c_overrides?: Record<string, any>
    c2c_formula_overrides?: Record<string, string>
  }>
}

/**
 * Compares saved quotation state against current state across ALL systems,
 * inputs, formulas, C2C overrides, commercial settings, notes, reminders, and totals.
 */
export function detectMultiSystemChanges(
  oldProj: any,
  nextState: MultiSystemDiffInput,
  wb?: any,
): ValueChange[] {
  const changes: ValueChange[] = []
  const oldSum = oldProj?.summary || {}

  // 1. Top-level quotation fields
  if (oldProj?.name && nextState.name && oldProj.name.trim() !== nextState.name.trim()) {
    pushDiff(changes, 'name', 'Quotation Name', oldProj.name, nextState.name)
  }
  if (oldProj?.status && nextState.status && oldProj.status !== nextState.status) {
    pushDiff(changes, 'status', 'Status', String(oldProj.status).toUpperCase(), String(nextState.status).toUpperCase())
  }
  const oldMode = oldSum.quotation_mode || 'set_wise'
  if (nextState.quotationMode && oldMode !== nextState.quotationMode) {
    const modeLabel = (m: string) => (m === 'unit_wise' ? 'Individual Unit-Wise' : m === 'unit_breakdown' ? 'Unit Breakdown' : 'System / Set-Wise')
    pushDiff(changes, 'quotation_mode', 'Quotation Format', modeLabel(oldMode), modeLabel(nextState.quotationMode))
  }

  const oldProfit = oldSum.profit_pct ?? oldProj?.inputs?.['__profit_pct']
  if (nextState.profitPct !== undefined && oldProfit !== undefined && Number(oldProfit) !== Number(nextState.profitPct)) {
    pushDiff(changes, '__profit_pct', 'Profit Margin %', `${oldProfit}%`, `${nextState.profitPct}%`)
  }
  const oldDisc = oldSum.discount_pct ?? oldProj?.inputs?.['__discount_pct'] ?? 0
  if (nextState.discountPct !== undefined && Number(oldDisc) !== Number(nextState.discountPct)) {
    pushDiff(changes, '__discount_pct', 'Discount %', `${oldDisc}%`, `${nextState.discountPct}%`)
  }
  const oldInst = oldSum.include_installation !== undefined
    ? Boolean(oldSum.include_installation)
    : (oldProj?.inputs?.['__include_installation'] !== undefined ? Number(oldProj.inputs['__include_installation']) !== 0 : true)
  if (nextState.includeInstallation !== undefined && oldInst !== nextState.includeInstallation) {
    pushDiff(changes, '__include_installation', 'Installation Included', oldInst ? 'Included' : 'Excluded', nextState.includeInstallation ? 'Included' : 'Excluded')
  }
  const oldExtra = Number(oldSum.extra_amount ?? oldProj?.inputs?.['__extra_amount'] ?? 0)
  if (nextState.extraAmount !== undefined && oldExtra !== Number(nextState.extraAmount)) {
    pushDiff(changes, '__extra_amount', 'Extra Charge / Unit', `₹${oldExtra}`, `₹${nextState.extraAmount}`)
  }

  if (nextState.notes !== undefined && String(oldSum.notes || '') !== String(nextState.notes || '')) {
    pushDiff(changes, 'notes', 'Internal Notes', oldSum.notes || '—', nextState.notes || '—')
  }
  if (nextState.description !== undefined && String(oldSum.description || '') !== String(nextState.description || '')) {
    pushDiff(changes, 'description', 'Scope / Description', oldSum.description || '—', nextState.description || '—')
  }
  const oldRemStr = oldSum.reminder && !oldSum.reminder.dismissed
    ? `${oldSum.reminder.date || ''}${oldSum.reminder.note ? ` (${oldSum.reminder.note})` : ''}`
    : 'None'
  const newRemStr = nextState.reminder && !nextState.reminder.dismissed
    ? `${nextState.reminder.date || ''}${nextState.reminder.note ? ` (${nextState.reminder.note})` : ''}`
    : 'None'
  if (nextState.reminder !== undefined && oldRemStr !== newRemStr) {
    pushDiff(changes, 'reminder', 'Follow-Up Reminder', oldRemStr, newRemStr)
  }

  // 2. Normalize old vs new systems arrays
  const oldSystems: any[] = Array.isArray(oldSum.systems) && oldSum.systems.length > 0
    ? oldSum.systems
    : [{ id: '1', name: 'SYSTEM A', quantity: 1, unit: 'SET', overrides: oldProj?.inputs || {}, formulaOverrides: oldProj?.formula_overrides || {} }]
  const newSystems: any[] = Array.isArray(nextState.systems) && nextState.systems.length > 0
    ? nextState.systems
    : [{ id: '1', name: 'SYSTEM A', quantity: 1, unit: 'SET', overrides: nextState.inputs || {}, formulaOverrides: nextState.formulaOverrides || {} }]

  const isMultiSys = oldSystems.length > 1 || newSystems.length > 1

  const matchOldSys = (ns: any, idx: number) => {
    if (ns.id) {
      const byId = oldSystems.find((os) => os.id && String(os.id) === String(ns.id))
      if (byId) return byId
    }
    return oldSystems[idx] || null
  }

  const matchedOldRefs = new Set<any>()
  newSystems.forEach((ns, idx) => {
    const os = matchOldSys(ns, idx)
    const sysName = String(ns.name || os?.name || `SYSTEM ${String.fromCharCode(65 + idx)}`)
    const sysPrefix = isMultiSys || idx > 0 ? `[${sysName}]` : ''

    if (!os) {
      changes.push({
        key: `sys_add_${idx}`,
        label: 'System Added',
        oldVal: '—',
        newVal: `${sysName} (${ns.quantity || 1} ${ns.unit || 'SET'})`,
        diff: `System Added: ${sysName} (${ns.quantity || 1} ${ns.unit || 'SET'})`,
      })
      return
    }
    matchedOldRefs.add(os)

    if (os.name && ns.name && os.name.trim() !== ns.name.trim()) {
      pushDiff(changes, `sys_${idx}_name`, 'System Name', os.name, ns.name, sysPrefix)
    }
    const oldQty = Number(os.quantity ?? 1)
    const newQty = Number(ns.quantity ?? 1)
    if (oldQty !== newQty) {
      pushDiff(changes, `sys_${idx}_qty`, 'Sets Quantity', oldQty, newQty, sysPrefix)
    }
    if ((os.unit || 'SET') !== (ns.unit || 'SET')) {
      pushDiff(changes, `sys_${idx}_unit`, 'Unit', os.unit || 'SET', ns.unit || 'SET', sysPrefix)
    }
    if (os.bay !== undefined && ns.bay !== undefined && Number(os.bay) !== Number(ns.bay)) {
      pushDiff(changes, `sys_${idx}_bay`, 'Bay Count', os.bay, ns.bay, sysPrefix)
    }
    if (os.includeInstallation !== undefined && ns.includeInstallation !== undefined && Boolean(os.includeInstallation) !== Boolean(ns.includeInstallation)) {
      pushDiff(changes, `sys_${idx}_inst`, 'System Installation', os.includeInstallation ? 'Included' : 'Excluded', ns.includeInstallation ? 'Included' : 'Excluded', sysPrefix)
    }
    if (Number(os.extraAmount || 0) !== Number(ns.extraAmount || 0)) {
      pushDiff(changes, `sys_${idx}_extra`, 'System Extra Charge', `₹${Number(os.extraAmount || 0)}`, `₹${Number(ns.extraAmount || 0)}`, sysPrefix)
    }

    // Per-system cell overrides (always fallback idx===0 to root inputs/formula_overrides if sys empty)
    const oldOv: Record<string, any> = (os.overrides && Object.keys(os.overrides).length > 0)
      ? os.overrides
      : (idx === 0 ? (oldProj?.inputs || {}) : {})
    const newOv: Record<string, any> = (ns.overrides && Object.keys(ns.overrides).length > 0)
      ? ns.overrides
      : (idx === 0 ? (nextState.inputs || {}) : {})

    const allOvKeys = Array.from(new Set([...Object.keys(oldOv), ...Object.keys(newOv)]))
    for (const k of allOvKeys) {
      // Skip quotation-wide __* flags on idx===0 since already emitted above
      if (idx === 0 && (k === '__profit_pct' || k === '__discount_pct' || k === '__include_installation' || k === '__extra_amount')) continue
      const def = getInitialCellValue(k, wb)
      const ov = oldOv[k] !== undefined && oldOv[k] !== null ? oldOv[k] : def
      const nv = newOv[k] !== undefined && newOv[k] !== null ? newOv[k] : def
      if (String(ov) === String(nv)) continue
      pushDiff(changes, `${sysPrefix || 'root'}:${k}`, fmtCellLabel(k), ov, nv, sysPrefix)
    }

    // Per-system formula overrides
    const oldFO: Record<string, string> = os.formulaOverrides || (idx === 0 ? (oldProj?.formula_overrides || {}) : {})
    const newFO: Record<string, string> = ns.formulaOverrides || (idx === 0 ? (nextState.formulaOverrides || {}) : {})
    const allFOKeys = Array.from(new Set([...Object.keys(oldFO), ...Object.keys(newFO)]))
    for (const k of allFOKeys) {
      const ov = oldFO[k] || 'Default Formula'
      const nv = newFO[k] || 'Default Formula'
      if (ov === nv) continue
      pushDiff(changes, `${sysPrefix || 'root'}:fo:${k}`, `Formula (${k})`, ov, nv, sysPrefix)
    }

    // Per-system C2C overrides & formulas
    const oldC2C = os.c2c_overrides || (idx === 0 ? (oldSum.c2c_overrides || {}) : {})
    const newC2C = ns.c2c_overrides || (idx === 0 ? (nextState.c2cOverrides || {}) : {})
    for (const k of Array.from(new Set([...Object.keys(oldC2C), ...Object.keys(newC2C)]))) {
      const ov = oldC2C[k] ?? 'Default'
      const nv = newC2C[k] ?? 'Default'
      if (String(ov) === String(nv)) continue
      pushDiff(changes, `${sysPrefix || 'root'}:c2c:${k}`, `C2C (${fmtCellLabel(k)})`, ov, nv, sysPrefix)
    }

    const oldC2CFO = os.c2c_formula_overrides || (idx === 0 ? (oldSum.c2c_formula_overrides || {}) : {})
    const newC2CFO = ns.c2c_formula_overrides || (idx === 0 ? (nextState.c2cFormulaOverrides || {}) : {})
    for (const k of Array.from(new Set([...Object.keys(oldC2CFO), ...Object.keys(newC2CFO)]))) {
      const ov = oldC2CFO[k] || 'Default Formula'
      const nv = newC2CFO[k] || 'Default Formula'
      if (ov === nv) continue
      pushDiff(changes, `${sysPrefix || 'root'}:c2cfo:${k}`, `C2C Formula (${k})`, ov, nv, sysPrefix)
    }
  })

  // Removed systems
  oldSystems.forEach((os, idx) => {
    if (!matchedOldRefs.has(os)) {
      const sysName = String(os.name || `SYSTEM ${String.fromCharCode(65 + idx)}`)
      changes.push({
        key: `sys_rem_${idx}`,
        label: 'System Removed',
        oldVal: sysName,
        newVal: 'Removed',
        diff: `System Removed: ${sysName}`,
      })
    }
  })

  // Root-level formula / C2C formula diffs not covered by system[0]
  if (nextState.formulaOverrides && !nextState.systems) {
    const oldFO = oldProj?.formula_overrides || {}
    for (const k of Array.from(new Set([...Object.keys(oldFO), ...Object.keys(nextState.formulaOverrides)]))) {
      const ov = oldFO[k] || 'Default'
      const nv = nextState.formulaOverrides[k] || 'Default'
      if (ov !== nv) pushDiff(changes, `fo:${k}`, `Formula (${k})`, ov, nv)
    }
  }

  // 3. Final Quotation Total change
  if (oldProj?.total !== undefined && oldProj?.total !== null && nextState.total !== undefined && nextState.total !== null) {
    if (Math.abs(Number(oldProj.total) - Number(nextState.total)) >= 0.5) {
      const oldT = `₹${Math.round(Number(oldProj.total)).toLocaleString('en-IN')}`
      const newT = `₹${Math.round(Number(nextState.total)).toLocaleString('en-IN')}`
      pushDiff(changes, 'total', 'Quotation Total', oldT, newT)
    }
  }

  return changes
}

export interface PdfEditSnapshot {
  clientInfo: Record<string, any>
  quotationMode: string
  profitPct?: number
  includeInstallation?: boolean
  extraAmount?: number
  includeFullProposal?: boolean
  showBodiesTable?: boolean
  bodyPlacement?: string
  placeBodyOnPage1?: boolean
  tableScale?: number
  tableDensity?: string
  systemsPerPageSetting?: number
  goodsColWidths?: any
  bodyColWidths?: any
  sets?: any[]
  lineItems?: any[]
  bodyItems?: any[]
  termsList?: any[]
  redSpecs?: string[]
  proposalSpecs?: any
}

/**
 * Compares every editable section in the Edit PDF modal so all edits —
 * from tiny custom rows, visibility toggles, and column widths up to rates,
 * systems, bodies table, terms, and 5-page technical proposal text — are logged.
 */
export function detectPdfEditChanges(
  oldSnap: PdfEditSnapshot,
  newSnap: PdfEditSnapshot,
): ValueChange[] {
  const changes: ValueChange[] = []
  const oldCI = oldSnap.clientInfo || {}
  const newCI = newSnap.clientInfo || {}

  // 1. Client Info text & commercial text fields
  for (const [field, label] of Object.entries(CLIENT_FIELD_LABELS)) {
    const ov = oldCI[field] ?? ''
    const nv = newCI[field] ?? ''
    if (String(ov).trim() !== String(nv).trim()) {
      pushDiff(changes, `pdf:${field}`, label, ov, nv, '[Edit PDF]')
    }
  }

  // 2. Visibility checkboxes on PDF
  for (const [field, label] of Object.entries(VISIBILITY_TOGGLE_LABELS)) {
    const ov = oldCI[field] !== false
    const nv = newCI[field] !== false
    if (ov !== nv) {
      pushDiff(changes, `pdf:${field}`, label, ov ? 'Visible' : 'Hidden', nv ? 'Visible' : 'Hidden', '[Edit PDF]')
    }
  }

  // 3. Client Custom Extra Rows
  const oldExtras: any[] = Array.isArray(oldCI.extra_fields) ? oldCI.extra_fields : []
  const newExtras: any[] = Array.isArray(newCI.extra_fields) ? newCI.extra_fields : []
  const extraMax = Math.max(oldExtras.length, newExtras.length)
  for (let i = 0; i < extraMax; i++) {
    const oe = oldExtras[i]
    const ne = newExtras[i]
    if (!oe && ne) {
      pushDiff(changes, `pdf:extra_add_${i}`, 'Custom Row Added', 'None', `${ne.label || 'Row'}: ${ne.value || ''}`, '[Edit PDF]')
    } else if (oe && !ne) {
      pushDiff(changes, `pdf:extra_rem_${i}`, 'Custom Row Removed', `${oe.label || 'Row'}: ${oe.value || ''}`, 'Removed', '[Edit PDF]')
    } else if (oe && ne) {
      if ((oe.label || '') !== (ne.label || '') || (oe.value || '') !== (ne.value || '')) {
        pushDiff(changes, `pdf:extra_edit_${i}`, `Custom Row (${ne.label || oe.label || i + 1})`, oe.value || oe.label, ne.value || ne.label, '[Edit PDF]')
      }
      if ((oe.visible !== false) !== (ne.visible !== false)) {
        pushDiff(changes, `pdf:extra_vis_${i}`, `Custom Row Visible (${ne.label || i + 1})`, oe.visible !== false ? 'Shown' : 'Hidden', ne.visible !== false ? 'Shown' : 'Hidden', '[Edit PDF]')
      }
    }
  }

  // 4. Layout & Commercial Options inside Edit PDF
  if (oldSnap.quotationMode !== newSnap.quotationMode) {
    pushDiff(changes, 'pdf:mode', 'PDF Format Mode', oldSnap.quotationMode, newSnap.quotationMode, '[Edit PDF]')
  }
  if (oldSnap.profitPct !== undefined && newSnap.profitPct !== undefined && Number(oldSnap.profitPct) !== Number(newSnap.profitPct)) {
    pushDiff(changes, 'pdf:profit', 'PDF Profit %', `${oldSnap.profitPct}%`, `${newSnap.profitPct}%`, '[Edit PDF]')
  }
  if (oldSnap.includeInstallation !== undefined && newSnap.includeInstallation !== undefined && oldSnap.includeInstallation !== newSnap.includeInstallation) {
    pushDiff(changes, 'pdf:inst', 'PDF Installation', oldSnap.includeInstallation ? 'Included' : 'Excluded', newSnap.includeInstallation ? 'Included' : 'Excluded', '[Edit PDF]')
  }
  if (oldSnap.extraAmount !== undefined && newSnap.extraAmount !== undefined && Number(oldSnap.extraAmount) !== Number(newSnap.extraAmount)) {
    pushDiff(changes, 'pdf:extra_amt', 'PDF Extra Amount', `₹${oldSnap.extraAmount}`, `₹${newSnap.extraAmount}`, '[Edit PDF]')
  }
  if (oldSnap.includeFullProposal !== undefined && newSnap.includeFullProposal !== undefined && oldSnap.includeFullProposal !== newSnap.includeFullProposal) {
    pushDiff(changes, 'pdf:full_prop', 'Full 5-Page Proposal', oldSnap.includeFullProposal ? '5 Pages' : '2 Pages', newSnap.includeFullProposal ? '5 Pages' : '2 Pages', '[Edit PDF]')
  }
  if (oldSnap.showBodiesTable !== undefined && newSnap.showBodiesTable !== undefined && oldSnap.showBodiesTable !== newSnap.showBodiesTable) {
    pushDiff(changes, 'pdf:show_bodies', 'Installation Bodies Table', oldSnap.showBodiesTable ? 'Shown' : 'Hidden', newSnap.showBodiesTable ? 'Shown' : 'Hidden', '[Edit PDF]')
  }
  if (oldSnap.bodyPlacement !== undefined && newSnap.bodyPlacement !== undefined && oldSnap.bodyPlacement !== newSnap.bodyPlacement) {
    pushDiff(changes, 'pdf:body_place', 'Bodies Table Page', oldSnap.bodyPlacement, newSnap.bodyPlacement, '[Edit PDF]')
  }
  if (oldSnap.tableScale !== undefined && newSnap.tableScale !== undefined && Number(oldSnap.tableScale) !== Number(newSnap.tableScale)) {
    pushDiff(changes, 'pdf:scale', 'PDF Table Scale %', `${oldSnap.tableScale}%`, `${newSnap.tableScale}%`, '[Edit PDF]')
  }
  if (oldSnap.tableDensity !== undefined && newSnap.tableDensity !== undefined && oldSnap.tableDensity !== newSnap.tableDensity) {
    pushDiff(changes, 'pdf:density', 'PDF Table Density', oldSnap.tableDensity, newSnap.tableDensity, '[Edit PDF]')
  }
  if (oldSnap.systemsPerPageSetting !== undefined && newSnap.systemsPerPageSetting !== undefined && Number(oldSnap.systemsPerPageSetting) !== Number(newSnap.systemsPerPageSetting)) {
    pushDiff(changes, 'pdf:sys_per_page', 'Systems Per Page', oldSnap.systemsPerPageSetting, newSnap.systemsPerPageSetting, '[Edit PDF]')
  }
  if (oldSnap.goodsColWidths && newSnap.goodsColWidths && JSON.stringify(oldSnap.goodsColWidths) !== JSON.stringify(newSnap.goodsColWidths)) {
    pushDiff(changes, 'pdf:goods_cols', 'Goods Table Column Widths', JSON.stringify(oldSnap.goodsColWidths), JSON.stringify(newSnap.goodsColWidths), '[Edit PDF]')
  }
  if (oldSnap.bodyColWidths && newSnap.bodyColWidths && JSON.stringify(oldSnap.bodyColWidths) !== JSON.stringify(newSnap.bodyColWidths)) {
    pushDiff(changes, 'pdf:body_cols', 'Bodies Table Column Widths', JSON.stringify(oldSnap.bodyColWidths), JSON.stringify(newSnap.bodyColWidths), '[Edit PDF]')
  }

  // 5. Sets (Multi-system title, quantity, rate, items)
  const oldSets = Array.isArray(oldSnap.sets) ? oldSnap.sets : []
  const newSets = Array.isArray(newSnap.sets) ? newSnap.sets : []
  for (let i = 0; i < Math.max(oldSets.length, newSets.length); i++) {
    const os = oldSets[i]
    const ns = newSets[i]
    const labelName = ns?.system_title || os?.system_title || `System ${i + 1}`
    if (!os && ns) {
      pushDiff(changes, `pdf:set_add_${i}`, 'PDF System Added', 'None', `${labelName} (${ns.set_label || ns.quantity || 1})`, '[Edit PDF]')
    } else if (os && !ns) {
      pushDiff(changes, `pdf:set_rem_${i}`, 'PDF System Removed', labelName, 'Removed', '[Edit PDF]')
    } else if (os && ns) {
      if ((os.system_title || '') !== (ns.system_title || '')) {
        pushDiff(changes, `pdf:set_title_${i}`, `PDF System Header`, os.system_title, ns.system_title, '[Edit PDF]')
      }
      if (Number(os.quantity || 1) !== Number(ns.quantity || 1)) {
        pushDiff(changes, `pdf:set_qty_${i}`, `PDF Set Qty (${labelName})`, os.quantity, ns.quantity, '[Edit PDF]')
      }
      if (Math.round(Number(os.rate || 0)) !== Math.round(Number(ns.rate || 0))) {
        pushDiff(changes, `pdf:set_rate_${i}`, `PDF Set Rate (${labelName})`, `₹${Math.round(Number(os.rate || 0))}`, `₹${Math.round(Number(ns.rate || 0))}`, '[Edit PDF]')
      }
    }
  }

  // 6. Unit-wise line items
  const oldItems = Array.isArray(oldSnap.lineItems) ? oldSnap.lineItems : []
  const newItems = Array.isArray(newSnap.lineItems) ? newSnap.lineItems : []
  for (let i = 0; i < Math.max(oldItems.length, newItems.length); i++) {
    const oi = oldItems[i]
    const ni = newItems[i]
    if (!oi && ni) {
      pushDiff(changes, `pdf:item_add_${i}`, 'PDF Line Item Added', 'None', `${ni.model_code || ni.description || 'Row'}`, '[Edit PDF]')
    } else if (oi && !ni) {
      pushDiff(changes, `pdf:item_rem_${i}`, 'PDF Line Item Removed', `${oi.model_code || oi.description || 'Row'}`, 'Removed', '[Edit PDF]')
    } else if (oi && ni) {
      const mLabel = ni.model_code || oi.model_code || `Row ${i + 1}`
      if ((oi.description || '') !== (ni.description || '')) {
        pushDiff(changes, `pdf:item_desc_${i}`, `Item Description (${mLabel})`, oi.description, ni.description, '[Edit PDF]')
      }
      if ((oi.dimensions || '') !== (ni.dimensions || '')) {
        pushDiff(changes, `pdf:item_dim_${i}`, `Item Dimensions (${mLabel})`, oi.dimensions, ni.dimensions, '[Edit PDF]')
      }
      if (Number(oi.qty || 0) !== Number(ni.qty || 0)) {
        pushDiff(changes, `pdf:item_qty_${i}`, `Item Qty (${mLabel})`, oi.qty, ni.qty, '[Edit PDF]')
      }
      if (Math.round(Number(oi.rate || 0)) !== Math.round(Number(ni.rate || 0))) {
        pushDiff(changes, `pdf:item_rate_${i}`, `Item Rate (${mLabel})`, `₹${Math.round(Number(oi.rate || 0))}`, `₹${Math.round(Number(ni.rate || 0))}`, '[Edit PDF]')
      }
    }
  }

  // 7. Installation Body Table Rows
  const oldBodies = Array.isArray(oldSnap.bodyItems) ? oldSnap.bodyItems : []
  const newBodies = Array.isArray(newSnap.bodyItems) ? newSnap.bodyItems : []
  for (let i = 0; i < Math.max(oldBodies.length, newBodies.length); i++) {
    const ob = oldBodies[i]
    const nb = newBodies[i]
    if (!ob && nb) {
      pushDiff(changes, `pdf:body_add_${i}`, 'Body Row Added', 'None', `${nb.title} (${nb.total_body} bodies)`, '[Edit PDF]')
    } else if (ob && !nb) {
      pushDiff(changes, `pdf:body_rem_${i}`, 'Body Row Removed', `${ob.title} (${ob.total_body} bodies)`, 'Removed', '[Edit PDF]')
    } else if (ob && nb) {
      const bTitle = nb.title || ob.title || `Row ${i + 1}`
      if ((ob.title || '') !== (nb.title || '') || (ob.sub_label || '') !== (nb.sub_label || '')) {
        pushDiff(changes, `pdf:body_lbl_${i}`, `Body Row Label (${bTitle})`, `${ob.title} (${ob.sub_label || ''})`, `${nb.title} (${nb.sub_label || ''})`, '[Edit PDF]')
      }
      if (Number(ob.per_bay_qty || 0) !== Number(nb.per_bay_qty || 0) || Number(ob.qty || 0) !== Number(nb.qty || '') || Number(ob.total_body || 0) !== Number(nb.total_body || 0)) {
        pushDiff(
          changes,
          `pdf:body_qty_${i}`,
          `Body Row (${bTitle})`,
          `${ob.per_bay_qty}×${ob.qty}=${ob.total_body}`,
          `${nb.per_bay_qty}×${nb.qty}=${nb.total_body}`,
          '[Edit PDF]',
        )
      }
    }
  }

  // 8. Terms List
  const oldTerms = Array.isArray(oldSnap.termsList) ? oldSnap.termsList : []
  const newTerms = Array.isArray(newSnap.termsList) ? newSnap.termsList : []
  for (let i = 0; i < Math.max(oldTerms.length, newTerms.length); i++) {
    const ot = oldTerms[i]
    const nt = newTerms[i]
    if (!ot && nt) {
      pushDiff(changes, `pdf:term_add_${i}`, 'Term Added', 'None', `${nt.title}: ${nt.text}`, '[Edit PDF]')
    } else if (ot && !nt) {
      pushDiff(changes, `pdf:term_rem_${i}`, 'Term Removed', `${ot.title}: ${ot.text}`, 'Removed', '[Edit PDF]')
    } else if (ot && nt && ((ot.title || '') !== (nt.title || '') || (ot.text || '') !== (nt.text || ''))) {
      pushDiff(changes, `pdf:term_edit_${i}`, `Term (${nt.title || ot.title})`, ot.text, nt.text, '[Edit PDF]')
    }
  }

  // 9. Red Specs highlights
  const oldReds = Array.isArray(oldSnap.redSpecs) ? oldSnap.redSpecs : []
  const newReds = Array.isArray(newSnap.redSpecs) ? newSnap.redSpecs : []
  if (JSON.stringify(oldReds) !== JSON.stringify(newReds)) {
    for (let i = 0; i < Math.max(oldReds.length, newReds.length); i++) {
      if (oldReds[i] !== newReds[i]) {
        pushDiff(changes, `pdf:red_${i}`, `Highlight Spec #${i + 1}`, oldReds[i] || 'None', newReds[i] || 'Removed', '[Edit PDF]')
      }
    }
  }

  // 10. Proposal Technical Specs (Pages 3, 4, 5)
  const oldProp = oldSnap.proposalSpecs || {}
  const newProp = newSnap.proposalSpecs || {}
  for (const pageKey of ['page3', 'page4', 'page5'] as const) {
    const op = oldProp[pageKey] || {}
    const np = newProp[pageKey] || {}
    const pageLabel = pageKey === 'page3' ? 'Page 3' : pageKey === 'page4' ? 'Page 4' : 'Page 5'
    for (const k of Array.from(new Set([...Object.keys(op), ...Object.keys(np)]))) {
      const ov = typeof op[k] === 'object' ? JSON.stringify(op[k]) : String(op[k] ?? '')
      const nv = typeof np[k] === 'object' ? JSON.stringify(np[k]) : String(np[k] ?? '')
      if (ov !== nv) {
        pushDiff(changes, `pdf:${pageKey}:${k}`, `Proposal ${pageLabel} (${k})`, ov, nv, '[Edit PDF]')
      }
    }
  }

  return changes
}
