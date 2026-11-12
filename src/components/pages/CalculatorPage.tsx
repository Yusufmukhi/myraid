import { useCallback, useEffect, useMemo, useRef, useState } from 'react'

import { api, Org, Project, ProjectStatus, Session, Template } from '../../lib/api'

import { Overrides, Workbook } from '../../lib/engine'

import { count, inr, inr2, kg } from '../../lib/format'

import { cleanOverrides, extractTemplateDefaults, summarize, Summary, filterCutToCutOverrides } from '../../lib/model'

import { exportWorkbookToXLSX } from '../../lib/excelExport'

import { detectValueChanges, parseRevision } from '../../lib/revision'

import { fetchClientIp, getCachedClientIp } from '../../lib/clientInfo'

import { getSavedClients } from '../../lib/clientDirectory'

import { Icon, Spinner, useUI } from '../ui'

import {

  OfficialQuotationDocument,

  QuotationClientInfo,

  QuotationLineItem,

  InstallationBodyItem,

  computeUnitWiseBreakdown,

  UnitWiseBreakdownResult,

} from '../quotation/OfficialQuotationDocument'



interface Props {

  session: Session

  org: Org

  project: Project

  wb: Workbook
  c2cWb?: Workbook

    onProjectUpdated: (p: Project) => void

  onSaveAsCopy: (name: string, payload?: { inputs?: any; formula_overrides?: any; total?: number; summary?: any }) => void

  onNavigateToSheets: (systemIdx?: number) => void

  onSystemChange?: (idx: number) => void

  onDirtyChange?: (dirty: boolean) => void

  saveTriggerRef?: React.MutableRefObject<(() => Promise<boolean>) | null>

}



const snapshot = (name: string, status: string, o: Overrides, fo: Record<string, string>) =>
  JSON.stringify([name.trim(), status, sortObj(o), sortObj(fo)])

function computeProjectSnapshot(
  projName: string,
  projStatus: string,
  projQuotationMode: string,
  systemsList: any[],
  currentActiveIdx: number,
  currentLiveOverrides: Overrides,
  currentLiveFormulaOverrides: Record<string, string>,
  liveExtraAmount: number,
  liveIncludeInstallation: boolean,
  wbInstance?: Workbook,
) {
  const normalizedSystems = (systemsList || []).map((sys, idx) => {
    const isCurrent = idx === currentActiveIdx
    const rawOv = isCurrent ? currentLiveOverrides : (sys.overrides || {})
    const cleanedOv = wbInstance ? cleanOverrides(wbInstance, rawOv) : rawOv
    const rawFo = isCurrent ? currentLiveFormulaOverrides : (sys.formulaOverrides || {})
    const inc = isCurrent ? liveIncludeInstallation : (sys.includeInstallation ?? true)
    const extra = isCurrent ? liveExtraAmount : (sys.extraAmount ?? 0)
    const qty = sys.quantity ?? 1
    const unit = sys.unit ?? 'SET'
    const name = sys.name ?? `SYSTEM ${String.fromCharCode(65 + idx)}`

    const sortedOvKeys = Object.keys(cleanedOv).sort()
    const sortedOv: Record<string, any> = {}
    for (const k of sortedOvKeys) {
      if (cleanedOv[k] !== undefined && cleanedOv[k] !== null && cleanedOv[k] !== '') {
        sortedOv[k] = cleanedOv[k]
      }
    }

    const sortedFoKeys = Object.keys(rawFo).sort()
    const sortedFo: Record<string, any> = {}
    for (const k of sortedFoKeys) {
      if (rawFo[k]) sortedFo[k] = rawFo[k]
    }

    return {
      name,
      qty,
      unit,
      inc: inc ? 1 : 0,
      extra,
      overrides: sortedOv,
      formulaOverrides: sortedFo,
    }
  })

  return JSON.stringify({
    name: (projName || '').trim(),
    status: projStatus,
    mode: projQuotationMode,
    systems: normalizedSystems,
  })
}



function sortObj(o: Record<string, any>) {

  return Object.keys(o)

    .sort()

    .map((k) => [k, o[k]])

}



/**

 * CellInput: maintains a local typing buffer so when the user backspaces,

 * the field stays empty while typing and NEVER snaps back to the template default.

 * Also allows setting and keeping 0.

 */

function CellInput({

  cellKey,

  val,

  step = 1,

  min = 0,

  unit,

  width,

  isOverridden,

  disabled = false,

  onCommit,

}: {

  cellKey: string

  val: number

  step?: number

  min?: number

  unit?: string

  width?: number

  isOverridden: boolean

  disabled?: boolean

  onCommit: (key: string, val: number | null | undefined) => void

}) {

  const [localStr, setLocalStr] = useState<string | null>(null)



  // Use local buffer while focused/typing; otherwise format the numeric val

  const displayVal =

    localStr !== null ? localStr : isNaN(val) || val === null ? '' : String(val)



  return (

    <div

      className={`excel-cell-input-wrapper ${isOverridden ? 'modified' : ''} ${disabled ? 'disabled' : ''}`}

      style={width ? { width } : undefined}

    >

      <input

        type="number"

        className="excel-cell-input"

        disabled={disabled}

        value={displayVal}

        step={step}

        min={min}

        onFocus={() => {

          setLocalStr(isNaN(val) || val === null ? '' : String(val))

        }}

        onChange={(e) => {

          const raw = e.target.value

          setLocalStr(raw)

          if (raw === '') {

            onCommit(cellKey, 0)

          } else {

            const parsed = parseFloat(raw)

            if (!isNaN(parsed)) {

              onCommit(cellKey, parsed)

            }

          }

        }}

        onBlur={() => {

          if (localStr === '' || localStr === null) {

            onCommit(cellKey, 0)

          }

          setLocalStr(null)

        }}

        title={isOverridden ? `Modified (key: ${cellKey})` : cellKey}

      />

      {unit && <span className="excel-cell-unit">{unit}</span>}

    </div>

  )

}



function computeSystemTotalBodies(

  sys: any,

  sysOverrides: Overrides,

  wbInstance: Workbook,

  isCurrent: boolean,

  currentBodies?: number,

): number {

  if (isCurrent && currentBodies !== undefined && currentBodies > 0) {

    return currentBodies

  }

  if (sys?.summary?.bodies && sys.summary.bodies > 0) {

    return sys.summary.bodies

  }

  const dfmuRows = Number(sysOverrides['Total cal!A3'] ?? wbInstance.num('Total cal', 'A3') ?? 0)

  const sfmuRows = Number(sysOverrides['Total cal!A7'] ?? Math.max(0, Math.round(dfmuRows / 2)))

  const sffuRows = Number(sysOverrides['Total cal!A11'] ?? Math.max(0, Math.round(dfmuRows / 2)))

  const dffuRows = Number(sysOverrides['Total cal!A15'] ?? dfmuRows)

  const dfmudRows = Number(sysOverrides['Total cal!A19'] ?? dfmuRows)



  const dfmuQty = Number(sysOverrides['Total cal!G4'] ?? wbInstance.num('Total cal', 'G4') ?? 0)

  const sfmuQty = Number(sysOverrides['Total cal!G8'] ?? wbInstance.num('Total cal', 'G8') ?? 0)

  const sffuQty = Number(sysOverrides['Total cal!G12'] ?? wbInstance.num('Total cal', 'G12') ?? 0)

  const dffuQty = Number(sysOverrides['Total cal!G16'] ?? wbInstance.num('Total cal', 'G16') ?? 0)

  const dfmudQty = Number(sysOverrides['Total cal!G20'] ?? wbInstance.num('Total cal', 'G20') ?? 0)



  const bodies = (dfmuQty * dfmuRows) + (sfmuQty * sfmuRows) + (sffuQty * sffuRows) + (dffuQty * dffuRows) + (dfmudQty * dfmudRows)

  return bodies > 0 ? bodies : 1

}



export function CalculatorPage({

  session,

  org,

  project,

  wb,
  c2cWb,

  

  onProjectUpdated,

  onSaveAsCopy,

  onNavigateToSheets,

  onSystemChange,

  onDirtyChange,

  saveTriggerRef,

}: Props) {

  const { toast, confirm } = useUI()

  const isAdminOrOwner = org.role === 'owner' || org.role === 'admin'
  const activeWb = wb



  const sanitizeOverrides = (raw: Overrides): Overrides => {
    const clean = { ...raw }
    delete clean['Total cal!B7']
    delete clean['Total cal!B11']
    delete clean['Total cal!B15']
    delete clean['Total cal!B19']
    delete clean['Total cal!A7']
    delete clean['Total cal!A11']
    delete clean['Total cal!A15']
    delete clean['Total cal!A19']
    // If project is brand new with no matrix inputs, default Bay and all model QTYs to 0
    if (clean['Total cal!A3'] === undefined && Object.keys(clean).filter((k) => !k.startsWith('__')).length === 0) {
      clean['Total cal!A3'] = 0
      clean['Total cal!G4'] = 0
      clean['Total cal!G8'] = 0
      clean['Total cal!G12'] = 0
      clean['Total cal!G16'] = 0
      clean['Total cal!G20'] = 0
    }
    return clean
  }

  // Project state

  const [name, setName] = useState(project.name)

  const [status, setStatus] = useState<ProjectStatus>(project.status)

  const [projectDescription, setProjectDescription] = useState<string>(project.summary?.description || (project as any).description || '')
  const [projectNotes, setProjectNotes] = useState<string>(project.summary?.notes || '')
  const [reminder, setReminder] = useState<{ date: string; note: string; dismissed?: boolean } | null>(project.summary?.reminder || null)
  const [showReminderModal, setShowReminderModal] = useState(false)
  const [reminderDateTimeInput, setReminderDateTimeInput] = useState<string>(() => {
    if (project.summary?.reminder?.date) {
      const d = new Date(project.summary.reminder.date)
      if (!isNaN(d.getTime())) {
        const pad = (n: number) => String(n).padStart(2, '0')
        return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}T${pad(d.getHours())}:${pad(d.getMinutes())}`
      }
    }
    const d = new Date()
    d.setDate(d.getDate() + 1)
    d.setHours(10, 0, 0, 0) // Default tomorrow 10:00 AM
    const pad = (n: number) => String(n).padStart(2, '0')
    return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}T${pad(d.getHours())}:${pad(d.getMinutes())}`
  })
  const [reminderNoteInput, setReminderNoteInput] = useState<string>(project.summary?.reminder?.note || 'Follow up with client on quotation')

  const handleSaveReminder = async (dateTimeStr: string, noteStr: string) => {
    const rem = {
      date: new Date(dateTimeStr).toISOString(),
      note: noteStr.trim() || 'Follow up on quotation',
      dismissed: false,
    }
    setReminder(rem)
    setShowReminderModal(false)

    const updatedSummary = {
      ...(project.summary || {}),
      reminder: rem,
    }

    try {
      await api.projectUpdate(
        session.token,
        project.id,
        name || project.name,
        status || project.status,
        project.inputs,
        project.formula_overrides,
        project.total,
        updatedSummary,
        'Set follow-up alarm reminder'
      )
      onProjectUpdated({
        ...project,
        summary: updatedSummary,
      })
    } catch (e) {
      console.warn('Failed to sync reminder to database:', e)
    }

    const remDate = new Date(rem.date)
    const dateFormatted = remDate.toLocaleDateString('en-IN', { month: 'short', day: 'numeric', year: 'numeric' })
    const timeFormatted = remDate.toLocaleTimeString('en-IN', { hour: '2-digit', minute: '2-digit' })
    toast(`⏰ Alarm reminder set for ${dateFormatted} at ${timeFormatted}`, 'ok')
  }

  const handleClearReminder = async () => {
    setReminder(null)
    setShowReminderModal(false)

    const updatedSummary = {
      ...(project.summary || {}),
      reminder: null,
    }

    try {
      await api.projectUpdate(
        session.token,
        project.id,
        name || project.name,
        status || project.status,
        project.inputs,
        project.formula_overrides,
        project.total,
        updatedSummary,
        'Cleared follow-up reminder'
      )
      onProjectUpdated({
        ...project,
        summary: updatedSummary,
      })
    } catch (e) {
      console.warn('Failed to clear reminder from database:', e)
    }

    toast('Reminder removed', 'ok')
  }

  const [overrides, setOverrides] = useState<Overrides>(() => sanitizeOverrides(project.inputs ?? {}))

  const [formulaOverrides, setFormulaOverrides] = useState<Record<string, string>>(project.formula_overrides ?? {})

  const [baseline, setBaseline] = useState(() => {
    const initSystems = (project.summary?.quotation_mode === 'unit_wise')
      ? [
          {
            id: '1',
            name: 'SYSTEM A',
            quantity: 1,
            unit: 'SET',
            overrides: cleanOverrides(wb, sanitizeOverrides(project.inputs ?? {})),
            formulaOverrides: project.formula_overrides ?? {},
            includeInstallation: project.inputs?.['__include_installation'] !== 0,
            extraAmount: Number(project.inputs?.['__extra_amount'] ?? project.summary?.extra_amount ?? 0),
          },
        ]
      : (project.summary?.systems && Array.isArray(project.summary.systems) && project.summary.systems.length > 0)
      ? project.summary.systems.map((s: any) => ({
          ...s,
          overrides: cleanOverrides(wb, sanitizeOverrides(s.overrides || {})),
          formulaOverrides: s.formulaOverrides || {},
          includeInstallation: s.includeInstallation ?? (project.inputs?.['__include_installation'] !== 0),
          extraAmount: Number(s.extraAmount ?? project.inputs?.['__extra_amount'] ?? project.summary?.extra_amount ?? 0),
        }))
      : [
          {
            id: '1',
            name: 'SYSTEM A',
            quantity: 1,
            unit: 'SET',
            overrides: cleanOverrides(wb, sanitizeOverrides(project.inputs ?? {})),
            formulaOverrides: project.formula_overrides ?? {},
            includeInstallation: project.inputs?.['__include_installation'] !== 0,
            extraAmount: Number(project.inputs?.['__extra_amount'] ?? project.summary?.extra_amount ?? 0),
          },
        ]

    return computeProjectSnapshot(
      project.name,
      project.status,
      project.summary?.quotation_mode || 'set_wise',
      initSystems,
      0,
      cleanOverrides(wb, sanitizeOverrides(project.inputs ?? {})),
      project.formula_overrides ?? {},
      Number(project.inputs?.['__extra_amount'] ?? project.summary?.extra_amount ?? 0),
      project.inputs?.['__include_installation'] !== 0,
      wb,
    )
  })

  const [showUnsavedLeaveModal, setShowUnsavedLeaveModal] = useState(false)

  const [pendingLeaveAction, setPendingLeaveAction] = useState<(() => void) | null>(null)



  const [saving, setSaving] = useState(false)

  const [showRecalcModal, setShowRecalcModal] = useState(false)

  const [recalcTemplates, setRecalcTemplates] = useState<Template[]>([])

  const [targetTemplateId, setTargetTemplateId] = useState<string>('')

  const revInfo = useMemo(() => parseRevision(name || project.name), [name, project.name])

  const [showReviseModal, setShowReviseModal] = useState(false)

  const [reviseName, setReviseName] = useState(revInfo.nextName)



  useEffect(() => {

    setReviseName(parseRevision(project.name).nextName)

  }, [project.name])



  const [showCopyModal, setShowCopyModal] = useState(false)

  const [copyName, setCopyName] = useState(`${project.name} (Copy)`)

  const [showShareModal, setShowShareModal] = useState(false)

  const [showOfficialQuotation, setShowOfficialQuotation] = useState(false)

  const [shareMembers, setShareMembers] = useState<any[]>([])

  const [loadingMembers, setLoadingMembers] = useState(false)

  const nameInputRef = useRef<HTMLInputElement>(null)
  const bayInputRef = useRef<HTMLInputElement>(null)

  const displayClientName = useMemo(() => {
    const buyer = (project.summary?.client_info?.buyer_name || '').trim()
    if (buyer) return buyer
    const clientId = project.summary?.client_id
    if (clientId) {
      try {
        const allClients = getSavedClients(org.id)
        const matched = allClients.find((c) => c.id === clientId)
        if (matched?.name) return matched.name.trim()
      } catch {}
    }
    const rawInputs = project.inputs as Record<string, any> | undefined
    const inputBuyer = rawInputs?.['buyer_name'] || rawInputs?.['client_name']
    if (typeof inputBuyer === 'string' && inputBuyer.trim()) return inputBuyer.trim()
    return ''
  }, [project.summary, project.inputs, org.id])



  const openShareModal = async () => {

    setShowShareModal(true)

    setLoadingMembers(true)

    try {

      const list = await api.orgMembers(session.token, org.id)

      setShareMembers(Array.isArray(list) ? list : [])

    } catch {

      setShareMembers([])

    } finally {

      setLoadingMembers(false)

    }

  }



  const handleToggleShare = async (targetUsername: string) => {

    const currentShared = Array.isArray(project.summary?.shared_with)

      ? [...project.summary.shared_with]

      : []

    let updated: string[]

    if (currentShared.includes(targetUsername)) {

      updated = currentShared.filter((u) => u !== targetUsername)

    } else {

      updated = [...currentShared, targetUsername]

    }

    try {

      const nextSummary = await api.projectShare(session.token, project.id, updated)

      onProjectUpdated({

        ...project,

        summary: nextSummary,

      })

      toast(

        updated.includes(targetUsername)

          ? `Quotation shared with @${targetUsername}`

          : `Access revoked for @${targetUsername}`,

      )

    } catch (e) {

      toast((e as Error).message, 'err')

    }

  }



  // Commercial profit: client-specific profit margin if client is configured, else input override, else default 25%

  const clientProfitMargin = useMemo(() => {

    const clientName = (project.summary?.client_info?.buyer_name || '').trim().toLowerCase()

    const clientId = project.summary?.client_id

    if (clientId || clientName) {

      const allClients = getSavedClients(org.id)

      const matched = (clientId && allClients.find((c) => c.id === clientId)) ||

        (clientName && allClients.find((c) => c.name.trim().toLowerCase() === clientName))

      if (matched && matched.profit_margin_pct !== undefined) {

        return matched.profit_margin_pct

      }

    }

    return undefined

  }, [project.summary, org.id])



  const sheetProfit = activeWb.num('Total cal', 'C26')
  const profitPct = overrides['Total cal!C26'] !== undefined
    ? Number(overrides['Total cal!C26'])
    : (overrides['__profit_pct'] !== undefined
      ? Number(overrides['__profit_pct'])
      : (sheetProfit > 0 ? sheetProfit : (clientProfitMargin !== undefined ? clientProfitMargin : (project.summary?.profit_pct !== undefined ? Number(project.summary.profit_pct) : 25))))

  const discountPct = overrides['__discount_pct'] !== undefined ? Number(overrides['__discount_pct']) : 0

  const [profitInput, setProfitInput] = useState<string>(() => String(profitPct))
  const [discountInput, setDiscountInput] = useState<string>(() => String(discountPct))

  useEffect(() => {
    setProfitInput(String(profitPct))
  }, [profitPct])

  useEffect(() => {
    setDiscountInput(String(discountPct))
  }, [discountPct])

  const handleProfitInputChange = (val: string) => {
    setProfitInput(val)
    const num = parseFloat(val)
    if (!isNaN(num) && num >= 0 && num <= 200) {
      setProfitPct(num)
    }
  }

  const handleProfitInputBlur = () => {
    const num = parseFloat(profitInput)
    if (isNaN(num) || num < 0) {
      setProfitInput('25')
      setProfitPct(25)
    } else {
      setProfitInput(String(num))
      setProfitPct(num)
    }
  }

  const handleDiscountInputChange = (val: string) => {
    setDiscountInput(val)
    const num = parseFloat(val)
    if (!isNaN(num) && num >= 0 && num <= 100) {
      setDiscountPct(num)
    }
  }

  const handleDiscountInputBlur = () => {
    const num = parseFloat(discountInput)
    if (isNaN(num) || num < 0) {
      setDiscountInput('0')
      setDiscountPct(0)
    } else {
      setDiscountInput(String(num))
      setDiscountPct(num)
    }
  }

  const includeInstallation = overrides['__include_installation'] !== undefined

    ? Boolean(Number(overrides['__include_installation']))

    : true



  // Quotation presentation format: 'set_wise' or 'unit_wise' (selectable directly on calculator page)

  const [quotationMode, setQuotationMode] = useState<'set_wise' | 'unit_wise'>(() => {

    if (project.summary?.quotation_mode === 'unit_wise') return 'unit_wise'

    return 'set_wise'

  })



  // Multi-System Matrices: each system (SYSTEM A, SYSTEM B...) has its own independent matrix, dimensions & inputs
  const [systems, setSystems] = useState<any[]>(() => {
    if (project.summary?.quotation_mode === 'unit_wise') {
      return [
        {
          id: '1',
          name: 'SYSTEM A',
          quantity: project.summary?.systems?.[0]?.quantity ? Math.max(1, Number(project.summary.systems[0].quantity)) : 1,
          unit: 'SET',
          overrides: project.inputs || {},
          formulaOverrides: project.formula_overrides || {},
          includeInstallation: project.inputs?.['__include_installation'] !== 0,
          extraAmount: Number(project.inputs?.['__extra_amount'] ?? project.summary?.extra_amount ?? 0),
        },
      ]
    }
    if (project.summary?.systems && Array.isArray(project.summary.systems) && project.summary.systems.length > 0) {
      return project.summary.systems.map((s: any) => ({
        ...s,
        quantity: (!s.quantity || Number(s.quantity) < 1) ? 1 : Number(s.quantity),
      }))
    }
    return [
      {
        id: '1',
        name: 'SYSTEM A',
        quantity: 1,
        unit: 'SET',
        overrides: project.inputs || {},
        formulaOverrides: project.formula_overrides || {},
        includeInstallation: project.inputs?.['__include_installation'] !== 0,
        extraAmount: Number(project.inputs?.['__extra_amount'] ?? project.summary?.extra_amount ?? 0),
      },
    ]
  })

  const [activeSystemIdx, setActiveSystemIdx] = useState<number>(0)

  // Other Extra Amount in Labour & Transportation (default 0)
  const extraAmount = overrides['__extra_amount'] !== undefined
    ? Number(overrides['__extra_amount'])
    : (systems[activeSystemIdx]?.extraAmount !== undefined
      ? Number(systems[activeSystemIdx].extraAmount)
      : (project.inputs?.['__extra_amount'] !== undefined
        ? Number(project.inputs['__extra_amount'])
        : (project.summary?.extra_amount !== undefined ? Number(project.summary.extra_amount) : 0)))

  const setExtraAmount = (val: number) => {
    const cleanVal = isNaN(val) ? 0 : Math.max(0, val)
    setOverrides((prev) => ({ ...prev, __extra_amount: cleanVal }))
    setSystems((prev) =>
      prev.map((s, idx) =>
        idx === activeSystemIdx
          ? {
              ...s,
              extraAmount: cleanVal,
              overrides: { ...(s.overrides || {}), __extra_amount: cleanVal },
            }
          : s
      )
    )
  }



  const handleSwitchSystem = (targetIdx: number) => {

    if (targetIdx === activeSystemIdx || targetIdx < 0 || targetIdx >= systems.length) return

    const currentFull = { ...overrides }

    const updated = systems.map((sys, idx) =>

      idx === activeSystemIdx

        ? {

            ...sys,

            overrides: currentFull,

            formulaOverrides,

            includeInstallation,

            extraAmount,

            summary,

            breakdown: calcBreakdown,

          }

        : sys

    )

    setSystems(updated)

    setActiveSystemIdx(targetIdx)
    setServerQuote(null)

    const target = updated[targetIdx]

    const targetOverrides = { ...(target.overrides || {}) }

    // Ensure single shared project site charges (Transport, Accommodation, Labour per body, Installation) remain synchronized across all systems
    if (currentFull['Total cal!N24'] !== undefined) targetOverrides['Total cal!N24'] = currentFull['Total cal!N24']
    if (currentFull['Total cal!N25'] !== undefined) targetOverrides['Total cal!N25'] = currentFull['Total cal!N25']
    if (currentFull['Total cal!B49'] !== undefined) targetOverrides['Total cal!B49'] = currentFull['Total cal!B49']
    if (currentFull['__include_installation'] !== undefined) targetOverrides['__include_installation'] = currentFull['__include_installation']

    // Unified Profit % and Discount % apply across all systems
    if (profitPct !== undefined) {
      targetOverrides['__profit_pct'] = profitPct
      targetOverrides['Total cal!C26'] = profitPct
    }
    if (discountPct !== undefined) {
      targetOverrides['__discount_pct'] = discountPct
    }

    // Extra Amount is PER SYSTEM: Load target system's own extra amount
    if (target.extraAmount !== undefined) {
      targetOverrides['__extra_amount'] = Number(target.extraAmount)
    } else if (target.overrides?.['__extra_amount'] !== undefined) {
      targetOverrides['__extra_amount'] = Number(target.overrides['__extra_amount'])
    } else {
      targetOverrides['__extra_amount'] = 0
    }

    setOverrides(targetOverrides)

    if (target.formulaOverrides) setFormulaOverrides(target.formulaOverrides)

    onSystemChange?.(targetIdx)

    toast(`Switched to ${target.name} Matrix`)

  }



  const handleAddSystemMatrix = () => {

    const currentFull = { ...overrides }

    const nextIdx = systems.length

    const nextChar = String.fromCharCode(65 + nextIdx)

    const newSys = {

      id: String(Date.now()),

      name: `SYSTEM ${nextChar}`,

      quantity: 1,

      unit: 'SET',

      overrides: { ...currentFull }, // Clones current matrix settings & shared charges as initial starting point

      formulaOverrides: { ...formulaOverrides },

      includeInstallation,

      extraAmount,

      summary,

      breakdown: calcBreakdown,

    }

    const updated = systems.map((sys, idx) =>

      idx === activeSystemIdx

        ? { ...sys, overrides: currentFull, formulaOverrides, includeInstallation, extraAmount, summary, breakdown: calcBreakdown }

        : sys

    )

    const nextSystems = [...updated, newSys]

    setSystems(nextSystems)

    setActiveSystemIdx(nextIdx)

    toast(`Created independent Matrix for SYSTEM ${nextChar}!`)

  }



  const handleDuplicateSystemMatrix = (idxToDup: number) => {

    const currentFull = { ...overrides }

    const source = systems[idxToDup]

    const nextChar = String.fromCharCode(65 + systems.length)

    const cloned = {

      ...source,

      id: String(Date.now()),

      name: `SYSTEM ${nextChar}`,

      overrides: idxToDup === activeSystemIdx ? { ...currentFull } : { ...(source.overrides || {}) },

    }

    const updated = systems.map((sys, idx) =>

      idx === activeSystemIdx

        ? { ...sys, overrides: currentFull, formulaOverrides, includeInstallation, extraAmount, summary, breakdown: calcBreakdown }

        : sys

    )

    setSystems([...updated, cloned])

    setActiveSystemIdx(systems.length)

    toast(`Duplicated into ${cloned.name} Matrix!`)

  }



  const handleRemoveSystemMatrix = async (idxToRemove: number) => {
    if (systems.length <= 1) return
    const targetSystem = systems[idxToRemove]
    const targetName = targetSystem?.name || `System ${idxToRemove + 1}`
    const ok = await confirm({
      title: `Delete ${targetName}?`,
      body: `Are you sure you want to delete ${targetName}? All inputs, matrices, and settings for this system will be permanently removed.`,
      confirmLabel: 'Delete System',
      danger: true,
    })
    if (!ok) return

    const removedName = systems[idxToRemove].name
    const updated = systems.filter((_, idx) => idx !== idxToRemove)
    setSystems(updated)
    const nextIdx = Math.min(activeSystemIdx, updated.length - 1)
    setActiveSystemIdx(nextIdx)
    setOverrides(updated[nextIdx].overrides || {})
    toast(`Removed ${removedName}`)
  }



  const handleUpdateSystemMeta = (idx: number, patch: Record<string, any>) => {

    setSystems((prev) =>

      prev.map((s, i) => (i === idx ? { ...s, ...patch } : s))

    )

  }



  const handleQuotationModeChange = (mode: 'set_wise' | 'unit_wise') => {
    setQuotationMode(mode)
    const updatedSummary = {
      ...project.summary,
      quotation_mode: mode,
      systems,
    }
    onProjectUpdated({
      ...project,
      summary: updatedSummary,
    })
    toast(mode === 'unit_wise' ? 'Switched to Individual Unit-Wise Mode' : 'Switched to System / Set-Wise Mode')
  }

  const MAX_STAFF_DISCOUNT_PCT = 10
  const discountRequest = project.summary?.discount_request as {
    requested_pct: number
    requested_by: string
    requested_at: string
    status: 'pending' | 'approved' | 'rejected'
    decided_by?: string
    decided_at?: string
  } | undefined

  const [isSendingDiscountRequest, setIsSendingDiscountRequest] = useState(false)



  const setProfitPct = (pct: number) => {
    const cleanPct = isNaN(pct) ? 0 : pct
    setOverrides((prev) => ({ ...prev, __profit_pct: cleanPct, 'Total cal!C26': cleanPct }))
    setSystems((prev) =>
      prev.map((s) => ({
        ...s,
        profitPct: cleanPct,
        overrides: { ...(s.overrides || {}), __profit_pct: cleanPct, 'Total cal!C26': cleanPct },
      }))
    )
  }

  const setDiscountPct = (pct: number) => {
    const cleanPct = isNaN(pct) ? 0 : Math.max(0, pct)
    setOverrides((prev) => ({ ...prev, __discount_pct: cleanPct }))
    setSystems((prev) =>
      prev.map((s) => ({
        ...s,
        discountPct: cleanPct,
        overrides: { ...(s.overrides || {}), __discount_pct: cleanPct },
      }))
    )
  }



  const handleSendDiscountRequest = async (pct: number) => {

    if (pct <= MAX_STAFF_DISCOUNT_PCT) return

    setIsSendingDiscountRequest(true)

    const newRequest = {

      requested_pct: pct,

      requested_by: session.username,

      requested_at: new Date().toISOString(),

      status: 'pending' as const,

    }

    const updatedSummary = {

      ...project.summary,

      discount_request: newRequest,

    }

    try {

      await api.projectUpdate(

        session.token,

        project.id,

        name,

        status,

        cleanOverrides(wb, { ...overrides, __discount_pct: pct }),

        formulaOverrides,

        summary.sales,

        updatedSummary,

        `Staff @${session.username} submitted discount authorization request for ${pct}%`,

      )

      onProjectUpdated({ ...project, summary: updatedSummary })

      toast(`Discount approval request for ${pct}% sent to Owner and Administrators!`)

    } catch (e) {

      toast((e as Error).message, 'err')

    } finally {

      setIsSendingDiscountRequest(false)

    }

  }



  const handleApproveDiscountRequest = async () => {

    if (!discountRequest) return

    const approvedPct = discountRequest.requested_pct

    const updatedRequest = {

      ...discountRequest,

      status: 'approved' as const,

      decided_by: session.username,

      decided_at: new Date().toISOString(),

    }

    const updatedSummary = {

      ...project.summary,

      discount_request: updatedRequest,

    }

    const nextOverrides = { ...overrides, __discount_pct: approvedPct }

    setOverrides(nextOverrides)

    try {

      await api.projectUpdate(

        session.token,

        project.id,

        name,

        status,

        cleanOverrides(wb, nextOverrides),

        formulaOverrides,

        summary.sales,

        updatedSummary,

        `Owner/Admin approved ${approvedPct}% discount requested by @${discountRequest.requested_by}`,

      )

      onProjectUpdated({ ...project, summary: updatedSummary })

      toast(`Approved ${approvedPct}% discount for @${discountRequest.requested_by}`)

    } catch (e) {

      toast((e as Error).message, 'err')

    }

  }



  const handleRejectDiscountRequest = async () => {

    if (!discountRequest) return

    const updatedRequest = {

      ...discountRequest,

      status: 'rejected' as const,

      decided_by: session.username,

      decided_at: new Date().toISOString(),

    }

    const updatedSummary = {

      ...project.summary,

      discount_request: updatedRequest,

    }

    const fallbackPct = Math.min(10, discountPct)

    const nextOverrides = { ...overrides, __discount_pct: fallbackPct }

    setOverrides(nextOverrides)

    try {

      await api.projectUpdate(

        session.token,

        project.id,

        name,

        status,

        cleanOverrides(wb, nextOverrides),

        formulaOverrides,

        summary.sales,

        updatedSummary,

        `Owner/Admin declined ${discountRequest.requested_pct}% discount request by @${discountRequest.requested_by}`,

      )

      onProjectUpdated({ ...project, summary: updatedSummary })

      toast(`Declined ${discountRequest.requested_pct}% discount request`)

    } catch (e) {

      toast((e as Error).message, 'err')

    }

  }



  const setIncludeInstallation = (inc: boolean) => {

    setOverrides((prev) => ({ ...prev, __include_installation: inc ? 1 : 0 }))

    setSystems((prev) =>

      prev.map((s) => ({

        ...s,

        includeInstallation: inc,

        overrides: { ...(s.overrides || {}), __include_installation: inc ? 1 : 0 },

      }))

    )

  }



  const openRecalcModal = async () => {

    setShowRecalcModal(true)

    try {

      const list = await api.templateList(session.token, org.id)

      const published = (Array.isArray(list) ? list : []).filter((t) => t.status === 'published')

      setRecalcTemplates(published)

      const defaultTpl = published.find((t) => t.is_default) || published[0]

      if (defaultTpl) setTargetTemplateId(defaultTpl.id)

    } catch {

      setRecalcTemplates([])

    }

  }



  // Synchronize template defaults to wb whenever project changes

  useEffect(() => {

    let active = true

    const applyTemplate = async () => {

      let defs: Record<string, number | null> = {}

      if (project.template_snapshot?.definition) {

        defs = extractTemplateDefaults(project.template_snapshot.definition)

      } else {

        try {

          const list = await api.templateList(session.token, org.id)

          const matching = list.find((t) => t.id === project.template_id || t.version_num === project.template_version)

          if (matching?.definition) {

            defs = extractTemplateDefaults(matching.definition)

          }

        } catch {

          // ignore

        }

      }

      if (active) {

        wb.setTemplateDefaults(defs)

        if (project.formula_overrides) {

          wb.setFormulaOverrides(project.formula_overrides)

        }

      }

    }

    applyTemplate()

    return () => {

      active = false

    }

  }, [project.id, project.template_id, project.template_version, project.template_snapshot, org.id, session.token, wb])



  const [wbVersion, setWbVersion] = useState(0)

  useEffect(() => {
    if (!isAdminOrOwner) return
    if (wb && wb.sheets.size > 1) return
    let active = true
    api.getFullWorksheets(session.token, org.id)
      .then((res) => {
        if (!active) return
        if (res.actual_workbook?.sheets) {
          wb.loadSheets(res.actual_workbook.sheets, res.named_constants)
          if (c2cWb && res.c2c_workbook?.sheets) {
            c2cWb.loadSheets(res.c2c_workbook.sheets, res.named_constants)
          }
          setWbVersion((v) => v + 1)
        }
      })
      .catch(() => {})
    return () => {
      active = false
    }
  }, [session.token, org.id, wb, c2cWb, isAdminOrOwner])

  // Recalculate summary live whenever inputs, profit %, discount %, or installation toggle change

  const actualSummary: Summary = useMemo(() => {
    wb.setFormulaOverrides(formulaOverrides)
    wb.setOverrides(overrides)
    return summarize(wb, profitPct, discountPct, includeInstallation, extraAmount)
  }, [overrides, formulaOverrides, profitPct, discountPct, includeInstallation, extraAmount, wb, wbVersion])

  const activeSysQty = Math.max(0, Number(systems[activeSystemIdx]?.quantity) || 0)

  // Total Combined Bodies across ALL systems & sets (treats whole project as one job for transportation & accommodation)
  // Formula: sum over each system of (singleSetBodies * sys.quantity)
  const totalCombinedBodies = useMemo(() => {
    return systems.reduce((acc, sys, idx) => {
      const sysOv = idx === activeSystemIdx ? overrides : (sys.overrides || {})
      const bodies = computeSystemTotalBodies(sys, sysOv, wb, idx === activeSystemIdx, actualSummary?.bodies)
      const qty = Math.max(0, Number(sys.quantity) || 0)
      return acc + (bodies * qty)
    }, 0) || 1
  }, [systems, activeSystemIdx, overrides, wb, actualSummary?.bodies])

  // Recalculate Cut-to-Cut summary live whenever inputs change
  const liveC2cSummary: Summary | null = useMemo(() => {
    if (!isAdminOrOwner) return null
    const targetWb = c2cWb || wb
    const activeSysC2C = systems[activeSystemIdx]?.c2c_overrides || {}
    const activeSysC2CFormulas = systems[activeSystemIdx]?.c2c_formula_overrides || {}
    const c2cCustom = { ...(project.summary?.c2c_overrides || {}), ...activeSysC2C }
    const c2cCustomFormulas = { ...(project.summary?.c2c_formula_overrides || {}), ...activeSysC2CFormulas }
    const c2cOv = { ...filterCutToCutOverrides(overrides), ...c2cCustom }
    const labourPerBody = overrides['Total cal!B49'] !== undefined ? Number(overrides['Total cal!B49']) : (wb.num('Total cal', 'B49') || 300)
    c2cOv['Total cal!B49'] = labourPerBody

    const projectTrans = overrides['Total cal!N24'] !== undefined ? Number(overrides['Total cal!N24']) : (wb.num('Total cal', 'N24') || 0)
    const projectAccom = overrides['Total cal!N25'] !== undefined ? Number(overrides['Total cal!N25']) : (wb.num('Total cal', 'N25') || 0)
    const sysBodies = computeSystemTotalBodies(systems[activeSystemIdx], overrides, wb, true, actualSummary?.bodies)
    const shareRatio = totalCombinedBodies > 0 ? (sysBodies / totalCombinedBodies) : 1
    const qtyForDiv = activeSysQty > 0 ? activeSysQty : 1
    c2cOv['Total cal!N24'] = projectTrans * shareRatio
    c2cOv['Total cal!N25'] = projectAccom * shareRatio

    if (overrides['Total cal!N23'] !== undefined) c2cOv['Total cal!N23'] = overrides['Total cal!N23']
    c2cOv['__extra_amount'] = extraAmount
    c2cOv['__include_installation'] = includeInstallation ? 1 : 0
    targetWb.setFormulaOverrides({ ...formulaOverrides, ...c2cCustomFormulas })
    targetWb.setOverrides(c2cOv)
    const res = summarize(targetWb, profitPct, discountPct, includeInstallation, extraAmount)
    if (targetWb === wb) {
      wb.setOverrides(overrides)
    }
    return res
  }, [isAdminOrOwner, overrides, formulaOverrides, profitPct, discountPct, includeInstallation, extraAmount, wb, c2cWb, wbVersion, project.summary?.c2c_overrides, project.summary?.c2c_formula_overrides, activeSysQty, totalCombinedBodies, actualSummary?.bodies, systems, activeSystemIdx])

  const [serverCalculating, setServerCalculating] = useState(false)

  const handleOpenOfficialQuotation = () => {
    if (serverCalculating) {
      toast('Calculating latest rates from server. Please wait a moment before opening PDF…', 'err')
      return
    }
    setShowOfficialQuotation(true)
  }
  const [serverSynced, setServerSynced] = useState(false)
  const [serverQuote, setServerQuote] = useState<any>(() => {
    if (project.summary && project.summary.sales && Array.isArray(project.summary.models)) {
      return { summary: project.summary, computed_total: project.total, c2c_summary: project.summary?.c2cSummary, total_combined_bodies: project.summary?.total_combined_bodies }
    }
    return null
  })



  // Authoritative totals from server calculate function, with profit margin applied on Direct Cost Base
  const summary: Summary = useMemo(() => {
    const isServerQuoteValidForActive = serverQuote && (serverQuote.active_system_idx === undefined || serverQuote.active_system_idx === activeSystemIdx)
    const raw: Summary = (isServerQuoteValidForActive && serverQuote?.summary?.models) ? serverQuote.summary : (actualSummary || (project.summary ?? null))
    if (!raw) return raw

    const projectTrans = (overrides['Total cal!N24'] !== undefined ? Number(overrides['Total cal!N24']) : (wb.num('Total cal', 'N24') || 0))
    const projectAcc = (overrides['Total cal!N25'] !== undefined ? Number(overrides['Total cal!N25']) : (wb.num('Total cal', 'N25') || 0))
    const sysBodies = (raw.bodies && raw.bodies > 0) ? raw.bodies : 36

    // Proportional shared transport and accommodation per set across total combined bodies
    const sysTrans = (totalCombinedBodies > 0 && (activeSysQty > 1 || systems.length > 1))
      ? Math.round(projectTrans * sysBodies / totalCombinedBodies)
      : (raw.transport ?? projectTrans)
    const sysAcc = (totalCombinedBodies > 0 && (activeSysQty > 1 || systems.length > 1))
      ? Math.round(projectAcc * sysBodies / totalCombinedBodies)
      : (raw.accommodation ?? projectAcc)

    // Re-apply profit margin on Direct Cost Base (costBase = Ex-Works + Transport + Accommodation + Installation + Extra)
    const effectiveCostBase = (raw.subtotal || 0) + (raw.installation || 0) + sysTrans + sysAcc + (raw.extraAmount || 0) * (raw.models?.reduce((a: number, m: any) => a + m.qty, 0) || 1)
    const effectiveProfitPct = profitPct !== undefined && !isNaN(profitPct) ? profitPct : (raw.profitPct || 25)
    const effectiveDiscountPct = discountPct !== undefined && !isNaN(discountPct) ? discountPct : (raw.discountPct || 0)

    const profitAmount = effectiveCostBase * (effectiveProfitPct / 100)
    const subtotalWithProfit = effectiveCostBase + profitAmount
    const discountAmount = subtotalWithProfit * (effectiveDiscountPct / 100)
    const beforeGst = subtotalWithProfit - discountAmount
    const gst = beforeGst * 0.18
    const sales = Math.round(beforeGst + gst)

    return {
      ...raw,
      transport: sysTrans,
      accommodation: sysAcc,
      costBase: effectiveCostBase,
      profitPct: effectiveProfitPct,
      profitAmount,
      subtotalWithProfit,
      discountPct: effectiveDiscountPct,
      discountAmount,
      beforeGst,
      gst,
      sales,
    }
  }, [serverQuote, serverCalculating, actualSummary, project.summary, profitPct, discountPct, activeSystemIdx, totalCombinedBodies, activeSysQty, systems.length, overrides])

  const c2cSummary: Summary | null = useMemo(() => {
    if (!isAdminOrOwner) return null
    const raw: Summary | null = serverQuote?.c2c_summary || liveC2cSummary || (project.summary?.c2cSummary ?? null)
    if (!raw) return null

    // Ensure C2C direct costBase strictly includes: Ex-Works + Transport + Accommodation + Installation + Extra Amount
    const projectTrans = (overrides['Total cal!N24'] !== undefined ? Number(overrides['Total cal!N24']) : (wb.num('Total cal', 'N24') || 0))
    const projectAcc = (overrides['Total cal!N25'] !== undefined ? Number(overrides['Total cal!N25']) : (wb.num('Total cal', 'N25') || 0))
    const labourPerBody = (overrides['Total cal!B49'] !== undefined ? Number(overrides['Total cal!B49']) : (wb.num('Total cal', 'B49') || 300))
    const sysBodies = computeSystemTotalBodies(systems[activeSystemIdx], overrides, wb, true, actualSummary?.bodies)
    const shareRatio = totalCombinedBodies > 0 ? (sysBodies / totalCombinedBodies) : 1
    const qtyForDiv = activeSysQty > 0 ? activeSysQty : 1
    const rawT = (raw.transport !== undefined && raw.transport > 0) ? raw.transport : (projectTrans * shareRatio)
    const trans = (activeSysQty > 1 && rawT === projectTrans) ? (rawT / qtyForDiv) : rawT
    const rawA = (raw.accommodation !== undefined && raw.accommodation > 0) ? raw.accommodation : (projectAcc * shareRatio)
    const acc = (activeSysQty > 1 && rawA === projectAcc) ? (rawA / qtyForDiv) : rawA
    const systemBodies = raw.bodies || (raw.models?.reduce((acc: number, m: any) => acc + (m.bodies || 0), 0)) || 51
    const inst = includeInstallation ? (systemBodies * labourPerBody) : 0
    const extraTot = (raw.extraAmount ?? extraAmount ?? 0) * (raw.models?.reduce((a: number, m: any) => a + m.qty, 0) || 1)
    const effC2cCostBase = raw.subtotal + inst + trans + acc + extraTot

    return {
      ...raw,
      transport: trans,
      accommodation: acc,
      installation: inst,
      extraAmount: raw.extraAmount ?? extraAmount ?? 0,
      costBase: effC2cCostBase,
    }
  }, [isAdminOrOwner, serverQuote, serverCalculating, liveC2cSummary, project.summary?.c2cSummary, overrides, includeInstallation, extraAmount, wb])



  // Live Unit-Wise Breakdown: Unit Rate = Base Unit Rate + All Transportation Rate + Extra Amount; Total Amount = Unit Rate * Qty
  // Logistics inside summary (transport, accommodation, installation) are already proportionally scaled for this system's set of bodies
  const calcBreakdown = useMemo(() => {
    return computeUnitWiseBreakdown(
      summary,
      activeWb,
      includeInstallation,
      profitPct,
      discountPct,
      summary?.bodies || undefined,
      extraAmount,
    )
  }, [summary, wb, includeInstallation, profitPct, discountPct, extraAmount, wbVersion])



  const getEffectiveSystems = (liveOverrides: Overrides, liveFO: Record<string, string>) => {
    return systems.map((sys, idx) =>
      idx === activeSystemIdx
        ? {
            ...sys,
            overrides: { ...liveOverrides },
            formulaOverrides: { ...liveFO },
            c2c_overrides: sys.c2c_overrides || (idx === 0 ? project.summary?.c2c_overrides : undefined),
            c2c_formula_overrides: sys.c2c_formula_overrides || (idx === 0 ? project.summary?.c2c_formula_overrides : undefined),
            includeInstallation,
            extraAmount,
          }
        : sys
    )
  }

  // Sync calculation with Server Edge Function — debounced (optimized from 500ms to 180ms for instant feel)
  useEffect(() => {
    let active = true
    const abortController = new AbortController()
    // Responsive debounce: fast timer with AbortController
    const timer = setTimeout(async () => {
      try {
        if (active) setServerCalculating(true)
        const effectiveSystems = getEffectiveSystems(overrides, formulaOverrides)
        const activeSysC2C = systems[activeSystemIdx]?.c2c_overrides || {}
        const activeSysC2CFormulas = systems[activeSystemIdx]?.c2c_formula_overrides || {}
        const mergedC2COv = { ...(project.summary?.c2c_overrides || {}), ...activeSysC2C }
        const mergedC2CFO = { ...(project.summary?.c2c_formula_overrides || {}), ...activeSysC2CFormulas }

        const result = await api.calculateServer(session.token, {
          org_id: org.id,
          inputs: overrides,
          systems: effectiveSystems,
          active_system_idx: activeSystemIdx,
          include_installation: includeInstallation,
          extra_amount: extraAmount,
          profit_pct: profitPct,
          discount_pct: discountPct,
          formula_overrides: formulaOverrides,
          c2c_overrides: Object.keys(mergedC2COv).length > 0 ? mergedC2COv : undefined,
          c2c_formula_overrides: Object.keys(mergedC2CFO).length > 0 ? mergedC2CFO : undefined,
        }, abortController.signal)

        if (active && (result?.summary || result)) {
          const rawSum = result.summary || result
          const effCostBase = rawSum.costBase ?? (rawSum.subtotal + (rawSum.installation || 0) + (rawSum.transport || 0) + (rawSum.accommodation || 0) + (rawSum.extraAmount || 0) * (rawSum.models?.reduce((a: number, m: any) => a + m.qty, 0) || 1))
          const effProfitPct = profitPct !== undefined && !isNaN(profitPct) ? profitPct : (rawSum.profitPct || 25)
          const effDiscountPct = discountPct !== undefined && !isNaN(discountPct) ? discountPct : (rawSum.discountPct || 0)
          const pAmount = effCostBase * (effProfitPct / 100)
          const subWithP = effCostBase + pAmount
          const dAmount = subWithP * (effDiscountPct / 100)
          const bGst = subWithP - dAmount
          const g = bGst * 0.18
          const s = Math.round(bGst + g)

          const normalizedSummary = {
            ...rawSum,
            costBase: effCostBase,
            profitPct: effProfitPct,
            profitAmount: pAmount,
            subtotalWithProfit: subWithP,
            discountPct: effDiscountPct,
            discountAmount: dAmount,
            beforeGst: bGst,
            gst: g,
            sales: s,
          }

          setServerQuote({
            ...result,
            active_system_idx: activeSystemIdx,
            summary: normalizedSummary,
            computed_total: systems.length > 1 ? (result.computed_total ?? s) : s,
          })
          setServerSynced(true)
        }
      } catch (err: any) {
        if (err?.name === 'AbortError' || err?.message?.includes('aborted')) {
          return // Quietly ignored because newer edits took precedence
        }
        // Keep the last server result as requested
        console.warn('Server calculation error, keeping last server result:', err)
      } finally {
        if (active) setServerCalculating(false)
      }
    }, 180)

    return () => {
      active = false
      abortController.abort()
      clearTimeout(timer)
    }
  }, [
    overrides,
    systems,
    activeSystemIdx,
    formulaOverrides,
    profitPct,
    discountPct,
    includeInstallation,
    extraAmount,
    org.id,
    session.token,
    project.summary?.c2c_overrides,
    project.summary?.c2c_formula_overrides,
  ])



  // Compile all systems into quotation_sets for PDF and saving
  const compiledSets = useMemo(() => {
    const liveSystems = getEffectiveSystems(overrides, formulaOverrides)
    const projectTransport = (overrides['Total cal!N24'] !== undefined ? Number(overrides['Total cal!N24']) : (wb.num('Total cal', 'N24') || 0))
    const projectAccommodation = (overrides['Total cal!N25'] !== undefined ? Number(overrides['Total cal!N25']) : (wb.num('Total cal', 'N25') || 0))
    const labourPerBody = (overrides['Total cal!B49'] !== undefined ? Number(overrides['Total cal!B49']) : (wb.num('Total cal', 'B49') || 300))

    return liveSystems.map((sys, idx) => {
      const isCurrent = idx === activeSystemIdx
      const sysOv: Overrides = { ...(sys.overrides || {}) }
      if (overrides['Total cal!B49'] !== undefined) sysOv['Total cal!B49'] = labourPerBody
      const sysExtra = isCurrent
        ? extraAmount
        : (sys.extraAmount !== undefined
          ? Number(sys.extraAmount)
          : (sys.overrides?.['__extra_amount'] !== undefined ? Number(sys.overrides['__extra_amount']) : 0))
      sysOv['__extra_amount'] = sysExtra
      sysOv['__include_installation'] = includeInstallation ? 1 : 0
      delete sysOv['Total cal!B7']
      delete sysOv['Total cal!B11']
      delete sysOv['Total cal!B15']
      delete sysOv['Total cal!B19']

      const sysFO = isCurrent ? formulaOverrides : (sys.formulaOverrides || {})
      const sysIncInst = sys.includeInstallation ?? includeInstallation
      const sysQty = Math.max(1, Number(sys.quantity) || 1)

      const sysBodies = computeSystemTotalBodies(sys, sysOv, wb, isCurrent, summary?.bodies)

      // Proportional shared transport and accommodation per set across total combined bodies
      const sysTransport = totalCombinedBodies > 0
        ? Math.round(projectTransport * sysBodies / totalCombinedBodies)
        : projectTransport
      const sysAccommodation = totalCombinedBodies > 0
        ? Math.round(projectAccommodation * sysBodies / totalCombinedBodies)
        : projectAccommodation

      sysOv['Total cal!N24'] = sysTransport
      sysOv['Total cal!N25'] = sysAccommodation

      let sysSummary: Summary
      let sysBreakdown: UnitWiseBreakdownResult | null = null

      if (isCurrent) {
        sysSummary = summary
        sysBreakdown = calcBreakdown
      } else {
        wb.setFormulaOverrides(sysFO)
        wb.setOverrides(sysOv)
        const baseSum = summarize(wb, profitPct, discountPct, sysIncInst, sysExtra)
        sysSummary = {
          ...baseSum,
          transport: sysTransport,
          accommodation: sysAccommodation,
        }
        sysBreakdown = computeUnitWiseBreakdown(
          sysSummary,
          wb,
          sysIncInst,
          profitPct,
          discountPct,
          sysBodies || undefined,
          sysExtra,
        )
        wb.setFormulaOverrides(formulaOverrides)
        wb.setOverrides(overrides)
      }

      const rate = Math.round(sysBreakdown?.finalBeforeGst || sysSummary?.beforeGst || 0)
      const amount = Math.round(rate * sysQty)
      const preDiscRate = Math.round(sysBreakdown?.totalPriceWithProfit || sysSummary?.subtotalWithProfit || (rate / Math.max(0.01, 1 - (discountPct / 100))))
      const preDiscAmount = Math.round(preDiscRate * sysQty)
      const activeRows = sysBreakdown?.rows && sysBreakdown.rows.length > 0 ? sysBreakdown.rows : []

      return {
        id: sys.id || String(idx + 1),
        system_title: sys.name || ('SYSTEM ' + String.fromCharCode(65 + idx)),
        set_label: (sysQty > 1 ? `${sysQty} ${sys.unit || 'SET'}S` : `1 ${sys.unit || 'SET'}`),
        quantity: sys.quantity || 1,
        unit: sys.unit || 'SET',
        rate,
        amount,
        pre_discount_rate: preDiscRate,
        pre_discount_amount: preDiscAmount,
        summary: sysSummary,
        breakdown: sysBreakdown,
        items: activeRows.map((r: any, rIdx: number) => ({
          id: String(rIdx + 1),
          model_code: r.model_code,
          model_name: r.model_name,
          dimensions: r.dimensions,
          unit_qty: r.qty,
          rate: r.unit_rate_with_profit,
          amount: r.total_price_with_profit,
          final_unit_rate: r.final_unit_rate,
          final_amount: r.final_amount,
        })),
      }
    })
  }, [systems, activeSystemIdx, summary, calcBreakdown, overrides, formulaOverrides, profitPct, discountPct, includeInstallation, extraAmount, totalCombinedBodies, wb, serverQuote])



  const currentSnapshot = useMemo(() => {
    return computeProjectSnapshot(
      name,
      status,
      quotationMode,
      systems,
      activeSystemIdx,
      cleanOverrides(wb, overrides),
      formulaOverrides,
      extraAmount,
      includeInstallation,
      wb,
    )
  }, [name, status, quotationMode, systems, activeSystemIdx, wb, overrides, formulaOverrides, extraAmount, includeInstallation])

  const dirty = currentSnapshot !== baseline ||
    projectNotes !== (project.summary?.notes || '') ||
    projectDescription !== (project.summary?.description || (project as any).description || '') ||
    JSON.stringify(reminder || null) !== JSON.stringify(project.summary?.reminder || null)

  const changedInputsCount = Object.keys(overrides).filter((k) => !k.startsWith('__')).length

  const formulaOverridesCount = Object.keys(formulaOverrides).length



  const setInput = useCallback(

    (key: string, value: number | null | undefined) => {

      setOverrides((prev) => {

        const next = { ...prev }

        if (value === undefined || value === null || isNaN(value)) {

          if (value === 0) {

            next[key] = 0

          } else {

            delete next[key]

          }

        } else {

          next[key] = value

        }

        // If user enters a positive installation amount while turned off, automatically include installation

        if (key === 'Total cal!N23' && typeof value === 'number' && value > 0 && !includeInstallation) {

          next['__include_installation'] = 1

        }

        // Rows for non-master models are strictly linked formulas driven by Bay (A3)
        delete next['Total cal!A7']
        delete next['Total cal!A11']
        delete next['Total cal!A15']
        delete next['Total cal!A19']

        const cleaned = cleanOverrides(wb, next)



        // If this is a shared project charge (Transport, Accommodation, Labour per body), sync across ALL systems

        if (key === 'Total cal!N24' || key === 'Total cal!N25' || key === 'Total cal!B49') {

          setSystems((sysList) =>

            sysList.map((s, idx) => {

              if (idx === activeSystemIdx) return s

              const sOv = { ...(s.overrides || {}) }

              if (cleaned[key] !== undefined) {

                sOv[key] = cleaned[key]

              } else {

                delete sOv[key]

              }

              return { ...s, overrides: sOv }

            })

          )

        }



        return cleaned

      })

    },

    [wb, includeInstallation, activeSystemIdx],

  )



  // Save changes — calls server for authoritative calculation, then persists

  const save = async (saveNote?: string) => {

    const nm = name.trim()

    if (!nm) {

      toast('Project name cannot be empty', 'err')

      nameInputRef.current?.focus()

      return

    }

    setSaving(true)

    const cleaned = cleanOverrides(wb, overrides)



    try {

      // ── Step 1: Call server-side calculation engine ──

      // Build the systems payload for multi-system calculation

      const currentClean = cleanOverrides(wb, overrides)

      const activeSysC2C = systems[activeSystemIdx]?.c2c_overrides || {}
      const activeSysC2CFormulas = systems[activeSystemIdx]?.c2c_formula_overrides || {}
      const mergedC2COv = { ...(project.summary?.c2c_overrides || {}), ...activeSysC2C }
      const mergedC2CFO = { ...(project.summary?.c2c_formula_overrides || {}), ...activeSysC2CFormulas }

      const systemsPayload = systems.map((sys, idx) => ({

        name: sys.name || `SYSTEM ${String.fromCharCode(65 + idx)}`,

        quantity: sys.quantity || 1,

        unit: sys.unit || 'SET',

        overrides: idx === activeSystemIdx ? currentClean : (sys.overrides || {}),

        formulaOverrides: idx === activeSystemIdx ? formulaOverrides : (sys.formulaOverrides || {}),

        c2c_overrides: idx === activeSystemIdx ? mergedC2COv : (sys.c2c_overrides || {}),

        c2c_formula_overrides: idx === activeSystemIdx ? mergedC2CFO : (sys.c2c_formula_overrides || {}),

      }))



      // Get template defaults if available

      const tplDefaults = project.template_snapshot?.definition

        ? extractTemplateDefaults(project.template_snapshot)

        : {}



      let serverSummary = summary

      let serverTotal = systems.length > 1 ? (allSystemsTotals.sales || summary.sales) : summary.sales

      let serverCleanedOverrides = cleaned

      let serverSystems: any[] = []

      let serverTotalBodies = totalCombinedBodies



      try {

        const serverResult = await api.serverCalculate(session.token, {

          org_id: org.id,


          inputs: currentClean,
          active_system_idx: activeSystemIdx,

          formula_overrides: formulaOverrides,

          template_defaults: tplDefaults,

          profit_pct: profitPct,

          discount_pct: discountPct,

          include_installation: includeInstallation,

          extra_amount: extraAmount,

          systems: systemsPayload.length > 0 ? systemsPayload : undefined,

          c2c_overrides: Object.keys(mergedC2COv).length > 0 ? mergedC2COv : undefined,

          c2c_formula_overrides: Object.keys(mergedC2CFO).length > 0 ? mergedC2CFO : undefined,

        }, org.id)



        if (serverResult && serverResult.summary) {
          const rawSum = serverResult.summary
          const effCostBase = rawSum.costBase ?? (rawSum.subtotal + (rawSum.installation || 0) + (rawSum.transport || 0) + (rawSum.accommodation || 0) + (rawSum.extraAmount || 0) * (rawSum.models?.reduce((a: number, m: any) => a + m.qty, 0) || 1))
          const effProfitPct = profitPct !== undefined && !isNaN(profitPct) ? profitPct : (rawSum.profitPct || 25)
          const effDiscountPct = discountPct !== undefined && !isNaN(discountPct) ? discountPct : (rawSum.discountPct || 0)
          const pAmount = effCostBase * (effProfitPct / 100)
          const subWithP = effCostBase + pAmount
          const dAmount = subWithP * (effDiscountPct / 100)
          const bGst = subWithP - dAmount
          const g = bGst * 0.18
          const s = Math.round(bGst + g)

          serverSummary = {
            ...rawSum,
            costBase: effCostBase,
            profitPct: effProfitPct,
            profitAmount: pAmount,
            subtotalWithProfit: subWithP,
            discountPct: effDiscountPct,
            discountAmount: dAmount,
            beforeGst: bGst,
            gst: g,
            sales: s,
          }
          serverTotal = systems.length > 1
            ? (allSystemsTotals.sales || Math.round((serverResult.computed_total ?? s) * 1.18))
            : s
          serverCleanedOverrides = serverResult.cleaned_overrides ?? cleaned
          serverSystems = serverResult.systems || []
          serverTotalBodies = serverResult.total_combined_bodies || totalCombinedBodies
        }

      } catch (calcErr) {

        // If server calculation fails, fall back to local calculation

        console.warn('Server calculation unavailable, using local values:', calcErr)

      }



      // ── Step 2: Detect changes for audit ──

      const detectedChanges = detectValueChanges(

        project.inputs || {},

        serverCleanedOverrides,

        project.status,

        status,

        project.name,

        nm,

        project.total,

        serverTotal,

        wb,

      )

      const formattedDiffs = detectedChanges.map((c) => c.diff)

      const noteToSave = saveNote || (formattedDiffs.length > 0 ? formattedDiffs.join('; ') : 'Saved project')

      const ip = await fetchClientIp().catch(() => getCachedClientIp())



      // ── Step 3: Update systems state ──

      const updatedSystems = systems.map((sys, idx) =>

        idx === activeSystemIdx

          ? {

              ...sys,

              overrides: serverCleanedOverrides,

              formulaOverrides,

              includeInstallation,

              extraAmount,

              summary: serverSummary,

              breakdown: calcBreakdown,

            }

          : sys

      )

      const effectiveSystems = quotationMode === 'unit_wise'

        ? [

            {

              id: '1',

              name: 'SYSTEM A',

              quantity: 1,

              unit: 'SET',

              overrides: serverCleanedOverrides,

              formulaOverrides,

              includeInstallation,

              extraAmount,

              summary: serverSummary,

              breakdown: calcBreakdown,

            }

          ]

        : updatedSystems

      setSystems(effectiveSystems)



      // ── Step 4: Build final summary with server-calculated values ──

      const updatedSummary = {

        ...project.summary,

        ...serverSummary,

        systems: effectiveSystems,

        quotation_sets: compiledSets,

        quotation_mode: quotationMode,

        extra_amount: extraAmount,

        include_installation: includeInstallation,

        total_combined_bodies: quotationMode === 'unit_wise' ? (serverSummary.bodies || 36) * activeSysQty : serverTotalBodies,
        c2c_overrides: systems[activeSystemIdx]?.c2c_overrides || project.summary?.c2c_overrides,
        c2c_formula_overrides: systems[activeSystemIdx]?.c2c_formula_overrides || project.summary?.c2c_formula_overrides,

        notes: projectNotes,
        description: projectDescription,
        reminder,

        profit_pct: profitPct,

        discount_pct: discountPct,

        changes: formattedDiffs,

        ip,

        calculated_on_server: true,

      }



      // ── Step 5: Persist to Supabase ──

      const primaryInputs = effectiveSystems[0]?.overrides || serverCleanedOverrides
      const primaryFO = effectiveSystems[0]?.formulaOverrides || formulaOverrides

      await api.projectUpdate(

        session.token,

        project.id,

        nm,

        status,

        primaryInputs,

        primaryFO,

        serverTotal,

        updatedSummary,

        noteToSave,

      )

      toast(formattedDiffs.length > 0 ? `Saved: ${formattedDiffs.slice(0, 2).join(', ')}${formattedDiffs.length > 2 ? '…' : ''}` : 'Project saved')

      setBaseline(
        computeProjectSnapshot(
          nm,
          status,
          quotationMode,
          effectiveSystems,
          activeSystemIdx,
          effectiveSystems[activeSystemIdx]?.overrides || serverCleanedOverrides,
          formulaOverrides,
          extraAmount,
          includeInstallation,
          wb,
        ),
      )

      const updatedProj: Project = {

        ...project,

        name: nm,

        status,

        inputs: (effectiveSystems[0]?.overrides || serverCleanedOverrides),

        formula_overrides: formulaOverrides,

        total: serverTotal,

        summary: updatedSummary,

        updated_at: new Date().toISOString(),

      }

      onProjectUpdated(updatedProj)

    } catch (e) {

      toast((e as Error).message, 'err')

    } finally {

      setSaving(false)

    }

  }



  // Recalculate using chosen template

  const handleRecalculateWithTarget = async () => {

    try {

      setSaving(true)

      const targetTpl =

        recalcTemplates.find((t) => t.id === targetTemplateId) ||

        recalcTemplates.find((t) => t.is_default) ||

        recalcTemplates[0]

      if (!targetTpl) {

        throw new Error('Please select a target template version.')

      }



      const newDef = targetTpl.definition ?? {}

      wb.setTemplateDefaults(newDef.default_inputs ?? {})

      if (newDef.formula_overrides) {

        wb.setFormulaOverrides({ ...newDef.formula_overrides, ...formulaOverrides })

      }

      const newSummary = summarize(wb, profitPct, discountPct, includeInstallation)



      await api.projectRecalculateTemplate(

        session.token,

        project.id,

        targetTpl.id,

        overrides,

        formulaOverrides,

        newSummary.sales,

        newSummary,

      )



      toast(`Recalculated with Template Version ${targetTpl.version_num}`)

      setShowRecalcModal(false)



      const updated = await api.projectGet(session.token, project.id)

      onProjectUpdated(updated)

    } catch (e) {

      toast((e as Error).message, 'err')

    } finally {

      setSaving(false)

    }

  }



  const handleSaveClientQuotationData = (

    clientInfo: QuotationClientInfo,

    items: QuotationLineItem[],

    bodyItems: InstallationBodyItem[],

    sets?: any[],

    extraSettings?: Record<string, any>

  ) => {

    const nextOverrides = { ...overrides }

    if (extraSettings?.__profit_pct !== undefined) {

      nextOverrides['__profit_pct'] = extraSettings.__profit_pct

    }

    if (extraSettings?.include_installation !== undefined) {

      nextOverrides['__include_installation'] = extraSettings.include_installation ? 1 : 0

    }

    if (extraSettings?.extra_amount !== undefined) {

      nextOverrides['__extra_amount'] = extraSettings.extra_amount

    }

    setOverrides(nextOverrides)

    if (sets && Array.isArray(sets)) {

      setSystems((prev) =>

        prev.map((sys, i) => {

          const matchSet = sets[i]

          if (matchSet) {

            return {

              ...sys,

              name: matchSet.system_title || sys.name,

              quantity: matchSet.quantity || sys.quantity,

              unit: matchSet.unit || sys.unit,

            }

          }

          return sys

        })

      )

    }



    const updatedSummary = {

      ...project.summary,

      client_info: clientInfo,

      quotation_items: items,

      body_items: bodyItems,

      quotation_mode: extraSettings?.quotation_mode || quotationMode,

      extra_amount: extraSettings?.extra_amount !== undefined ? extraSettings.extra_amount : extraAmount,

      ...(sets ? { quotation_sets: sets } : {}),

      ...(extraSettings || {}),

    }



    const updatedProject: Project = {

      ...project,

      inputs: cleanOverrides(wb, nextOverrides),

      summary: updatedSummary,

    }

    onProjectUpdated(updatedProject)

    const finalCleanOverrides = cleanOverrides(wb, nextOverrides)
    const effectiveSets = (sets && Array.isArray(sets)) ? sets : []
    const updatedSystemsList = systems.map((sys, i) => {
      const matchSet = effectiveSets[i]
      const isCur = i === activeSystemIdx
      return {
        ...sys,
        name: matchSet?.system_title || sys.name,
        quantity: matchSet?.quantity || sys.quantity,
        unit: matchSet?.unit || sys.unit,
        overrides: isCur ? finalCleanOverrides : (sys.overrides || {}),
        extraAmount: extraSettings?.extra_amount !== undefined ? extraSettings.extra_amount : extraAmount,
      }
    })

    setBaseline(
      computeProjectSnapshot(
        name,
        status,
        extraSettings?.quotation_mode || quotationMode,
        updatedSystemsList,
        activeSystemIdx,
        finalCleanOverrides,
        formulaOverrides,
        extraSettings?.extra_amount !== undefined ? extraSettings.extra_amount : extraAmount,
        includeInstallation,
        wb,
      ),
    )

    api.projectUpdate(

      session.token,

      project.id,

      name,

      status,

      cleanOverrides(wb, nextOverrides),

      formulaOverrides,

      summary.sales,

      updatedSummary,

      'Updated official quotation client details'

    ).then(() => {

      toast('Quotation client details saved successfully!')

    }).catch((err) => {

      toast((err as Error).message, 'err')

    })

  }



  // Handle Ctrl/Cmd + S and Ctrl/Cmd + P

  useEffect(() => {

    const h = (e: KeyboardEvent) => {

      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 's') {

        e.preventDefault()

        save()

      } else if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'p') {

        e.preventDefault()

        setShowOfficialQuotation(true)

      } else if (e.key === '/') {
        const activeEl = document.activeElement
        const tag = activeEl?.tagName?.toLowerCase()
        const isEditable = tag === 'input' || tag === 'textarea' || (activeEl as HTMLElement)?.isContentEditable
        if (!isEditable) {
          e.preventDefault()
          if (bayInputRef.current) {
            bayInputRef.current.scrollIntoView({ behavior: 'smooth', block: 'center' })
            bayInputRef.current.focus()
            bayInputRef.current.select()
          }
        }
      }

    }

    window.addEventListener('keydown', h)

    return () => window.removeEventListener('keydown', h)

  })



  // Prevent leaving with unsaved work

  useEffect(() => {

    const h = (e: BeforeUnloadEvent) => {

      if (dirty) {

        e.preventDefault()

        e.returnValue = ''

      }

    }

    window.addEventListener('beforeunload', h)

    return () => window.removeEventListener('beforeunload', h)

  }, [dirty])

  useEffect(() => {
    onDirtyChange?.(dirty)
  }, [dirty, onDirtyChange])

  useEffect(() => {
    if (saveTriggerRef) {
      saveTriggerRef.current = async () => {
        try {
          await save()
          return true
        } catch {
          return false
        }
      }
    }
    return () => {
      if (saveTriggerRef) saveTriggerRef.current = null
    }
  }, [saveTriggerRef, save])

  const handleNavigateToSheets = () => {
    if (dirty) {
      toast('Please save your changes first', 'err')
      setPendingLeaveAction(() => () => onNavigateToSheets(activeSystemIdx))
      setShowUnsavedLeaveModal(true)
      return
    }
    onNavigateToSheets(activeSystemIdx)
  }

  const handleDownloadC2CExcel = async () => {
    try {
      toast('Generating Cut-to-Cut Excel workbook…', 'ok')
      let targetWb = c2cWb
      if ((!targetWb || targetWb.sheets.size <= 1) && session?.token && org?.id) {
        try {
          const res = await api.getFullWorksheets(session.token, org.id)
          if (res?.c2c_workbook) {
            targetWb = new Workbook(res.c2c_workbook, res.named_constants)
          }
        } catch {}
      }

      const wbToExport = targetWb || activeWb
      if (wbToExport) {
        // Apply cut-to-cut inputs & formula overrides
        const actInputs = systems[activeSystemIdx]?.overrides || overrides
        const activeSysC2C = systems[activeSystemIdx]?.c2c_overrides || {}
        const activeSysC2CFormulas = systems[activeSystemIdx]?.c2c_formula_overrides || {}
        const mergedC2COv = { ...(project.summary?.c2c_overrides || {}), ...activeSysC2C }
        const mergedC2CFO = { ...(project.summary?.c2c_formula_overrides || {}), ...activeSysC2CFormulas }
        const c2cInputs = { ...filterCutToCutOverrides(actInputs), ...mergedC2COv }
        wbToExport.setOverrides(c2cInputs)
        wbToExport.setFormulaOverrides({ ...formulaOverrides, ...mergedC2CFO })

        const qtnTitle = (project.summary?.client_info?.qtn_no || name || project.name || 'Quotation').replace(/[/\\?%*:|"<>]/g, '_')
        const activeSysName = (systems[activeSystemIdx]?.name || 'System').replace(/\s+/g, '_')
        const fileName = `${qtnTitle}_${activeSysName}_Cut_to_Cut_Analysis.xlsx`

        const c2cData = {
          profit: displayC2CProfit,
          total: displayC2CTotal,
          ratio: displayC2CRatio,
          costBase: c2cSummary?.costBase,
          labour: c2cSummary?.labourSubtotal,
          steelWeight: c2cSummary?.totalWeight,
        }

        exportWorkbookToXLSX(wbToExport, fileName, c2cData)
        toast('Cut-to-Cut Excel spreadsheet downloaded successfully!', 'ok')
      }
    } catch (e: any) {
      toast(`Export failed: ${e?.message || 'Error creating Excel file'}`, 'err')
    }
  }



  const hasNewerTemplate = (org.current_template_version || 1) > project.template_version



  // Helper to read current input value

  const getCellVal = (key: string): number => {
    if (key in overrides && overrides[key] !== null && overrides[key] !== undefined) {
      return Number(overrides[key])
    }
    const [sheet, a] = [key.slice(0, key.lastIndexOf('!')), key.slice(key.lastIndexOf('!') + 1)]
    const val = activeWb.num(sheet, a)
    if (!val || val === 0) {
      if (key === 'Total cal!H3') return 90
      if (key === 'Total cal!M3') return 90
      if (key === 'Total cal!B49') return 300
      if (key === 'Total cal!C26') return 25
    }
    return val
  }



  const isApprovedLocked = status === 'won' || status === 'approved'



  // Dynamic Bay state (drives DFMU/DFFU rows = Bay * 2, and SFMU/SFFU rows = Bay)

  const rawDfmuRows = overrides['Total cal!A3'] !== undefined ? Number(overrides['Total cal!A3']) : activeWb.num('Total cal', 'A3')
  const currentDfmuRows = rawDfmuRows !== undefined && !isNaN(rawDfmuRows) ? rawDfmuRows : 0

  const currentBay = Math.max(0, Math.round(currentDfmuRows / 2))

  const [bayInput, setBayInput] = useState<string>(String(currentBay))



  useEffect(() => {

    setBayInput(String(currentBay))

  }, [currentBay])



  // Render an editable Excel table cell with zero & backspace preservation

  const renderCellInput = (

    key: string,

    step: number = 1,

    min: number = 0,

    unit?: string,

    width?: number,

  ) => {

    const isOverridden = key in overrides && overrides[key] !== null

    const val = getCellVal(key)

    return (

      <CellInput

        key={key}

        cellKey={key}

        val={val}

        step={step}

        min={min}

        unit={unit}

        width={width}

        isOverridden={isOverridden}

        disabled={isApprovedLocked}

        onCommit={setInput}

      />

    )

  }



  // Render a linked / computed cell

  const renderLinkedCell = (val: number | string, formulaBadge: string, hint?: string) => {

    const isNum = typeof val === 'number'

    const displayVal = isNum ? Math.round((val as number) * 100) / 100 : val

    if (!isAdminOrOwner) {

      return (

        <div className="excel-cell-linked plain" title="Calculated specification">

          <span className="excel-linked-val">{displayVal}</span>

        </div>

      )

    }

    return (

      <div className="excel-cell-linked" title={hint ?? `Linked formula: ${formulaBadge}`}>

        <span className="excel-linked-val">{displayVal}</span>

        <span className="excel-linked-badge">{formulaBadge}</span>

      </div>

    )

  }

  // Render a live calculation value with graceful loading state so users don't see stale/hardcoded figures
  const renderLiveVal = (val: number | string | undefined | null, formatter: (n: number) => string = inr2) => {
    if (serverCalculating) {
      return (
        <span
          className="live-calc-pulse"
          title="Calculating live value..."
        >
          —
        </span>
      )
    }
    if (val === undefined || val === null || (typeof val === 'number' && isNaN(val))) {
      return '—'
    }
    return typeof val === 'number' ? formatter(val) : val
  }



  // Small text breakdown for the Body widget: "DFMU: 8 · SFMU: 1 · SFFU: 1"
  const bodyPartsText = (summary?.models || [])
    .filter((m) => m.qty > 0)
    .map((m) => `${m.code}: ${m.qty}`)
    .join(' · ')

  const dfmuModel = (summary?.models || []).find((m) => m.code === 'DFMU')
  const sfmuModel = (summary?.models || []).find((m) => m.code === 'SFMU')
  const sffuModel = (summary?.models || []).find((m) => m.code === 'SFFU')
  const dffuModel = (summary?.models || []).find((m) => m.code === 'DFFU')
  const dfmuDModel = (summary?.models || []).find((m) => m.code === 'DFMU-D')



  const watermarkText = `CONFIDENTIAL · ${org.name.toUpperCase()} · STAFF: ${session.username.toUpperCase()} · REV ${revInfo.currentRev} · ${new Date().toLocaleDateString()}`



  const activeSys = systems[activeSystemIdx]

  const activeSysName = activeSys?.name || `SYSTEM ${String.fromCharCode(65 + activeSystemIdx)}`

  const activeSysUnit = activeSys?.unit || 'SET'

  // Comprehensive Body Calculations for DFMU, SFMU, SFFU, DFFU, DFMU-D
  const bodiesBreakdown = useMemo(() => {
    const dfmuRows = Number(overrides['Total cal!A3'] ?? activeWb.num('Total cal', 'A3') ?? 0)
    const sfmuRows = Number(overrides['Total cal!A7'] ?? Math.max(0, Math.round(dfmuRows / 2)))
    const sffuRows = Number(overrides['Total cal!A11'] ?? Math.max(0, Math.round(dfmuRows / 2)))
    const dffuRows = Number(overrides['Total cal!A15'] ?? dfmuRows)
    const dfmudRows = Number(overrides['Total cal!A19'] ?? dfmuRows)

    const dfmuQty = Number(overrides['Total cal!G4'] ?? activeWb.num('Total cal', 'G4') ?? 0)
    const sfmuQty = Number(overrides['Total cal!G8'] ?? activeWb.num('Total cal', 'G8') ?? 0)
    const sffuQty = Number(overrides['Total cal!G12'] ?? activeWb.num('Total cal', 'G12') ?? 0)
    const dffuQty = Number(overrides['Total cal!G16'] ?? activeWb.num('Total cal', 'G16') ?? 0)
    const dfmudQty = Number(overrides['Total cal!G20'] ?? activeWb.num('Total cal', 'G20') ?? 0)

    const dfmuBodies = Math.round(dfmuQty * dfmuRows)
    const sfmuBodies = Math.round(sfmuQty * sfmuRows)
    const sffuBodies = Math.round(sffuQty * sffuRows)
    const dffuBodies = Math.round(dffuQty * dffuRows)
    const dfmudBodies = Math.round(dfmudQty * dfmudRows)

    const singleSetBodies = (dfmuBodies + sfmuBodies + sffuBodies + dffuBodies + dfmudBodies) || (summary?.bodies || 0)
    const totalBodiesForSets = singleSetBodies * (activeSysQty > 0 ? activeSysQty : 1)

    const activeList: { label: string; count: number; desc: string }[] = []
    if (dfmuQty > 0 || dfmuBodies > 0) activeList.push({ label: 'DFMU', count: dfmuBodies, desc: `${dfmuRows} rows × ${dfmuQty} qty = ${dfmuBodies}` })
    if (sfmuQty > 0 || sfmuBodies > 0) activeList.push({ label: 'SFMU', count: sfmuBodies, desc: `${sfmuRows} rows × ${sfmuQty} qty = ${sfmuBodies}` })
    if (sffuQty > 0 || sffuBodies > 0) activeList.push({ label: 'SFFU', count: sffuBodies, desc: `${sffuRows} rows × ${sffuQty} qty = ${sffuBodies}` })
    if (dffuQty > 0 || dffuBodies > 0) activeList.push({ label: 'DFFU', count: dffuBodies, desc: `${dffuRows} rows × ${dffuQty} qty = ${dffuBodies}` })
    if (dfmudQty > 0 || dfmudBodies > 0) activeList.push({ label: 'DFMU-D', count: dfmudBodies, desc: `${dfmudRows} rows × ${dfmudQty} qty = ${dfmudBodies}` })

    const summaryText = activeList.length > 0 ? activeList.map((m) => `${m.label}: ${m.count}`).join(' · ') : '0 Bodies'
    const tooltipText = activeList.map((m) => `${m.label}: ${m.desc}`).join('\n') + `\nTotal: ${singleSetBodies} bodies / set${activeSysQty > 1 ? ` × ${activeSysQty} sets = ${totalBodiesForSets} bodies` : ''}`

    return {
      dfmuBodies,
      sfmuBodies,
      sffuBodies,
      dffuBodies,
      dfmudBodies,
      singleSetBodies,
      totalBodiesForSets,
      activeList,
      summaryText,
      tooltipText,
    }
  }, [overrides, activeWb, summary?.bodies, activeSysQty])

  const activeSysTotalBodies = bodiesBreakdown.totalBodiesForSets



  // All Systems Combined Metrics across all systems and quantities

  const allSystemsTotals = useMemo(() => {

    let spTotal = 0

    let trolleyTotal = 0

    let totalWeight = 0

    let weightWithMachine = 0

    let labourSubtotal = 0

    let powderTotal = 0

    let beforeGst = 0
    let costBase = 0

    compiledSets.forEach((s) => {
      const q = Math.max(1, Number(s.quantity) || 1)
      const sum = s.summary || summary
      const sysCostBase = s.breakdown?.totalCostBase || sum.costBase || 0

      spTotal += (sum.steel?.spTotal || 0) * q
      trolleyTotal += (sum.trolley?.total || 0) * q
      totalWeight += (sum.totalWeight || 0) * q
      weightWithMachine += (sum.weightWithMachine || 0) * q
      labourSubtotal += (sum.labourSubtotal || 0) * q
      powderTotal += (sum.powderTotal || 0) * q
      costBase += sysCostBase * q
      beforeGst += (s.amount || Math.round((sum.beforeGst || 0) * q))
    })

    const profitAmount = costBase * (profitPct / 100)
    const subtotalWithProfit = costBase + profitAmount
    const discountAmount = subtotalWithProfit * (discountPct / 100)
    const gst = beforeGst * 0.18
    const sales = Math.round(beforeGst + gst)

    return {
      spTotal,
      trolleyTotal,
      totalWeight,
      weightWithMachine,
      labourSubtotal,
      powderTotal,
      costBase,
      profitAmount,
      subtotalWithProfit,
      discountAmount,
      beforeGst,
      gst,
      sales,
    }
  }, [compiledSets, summary, profitPct, discountPct])



  // Cut-to-Cut Combined Metrics across all systems (Admin Only)

  const c2cAllSystemsTotals = useMemo(() => {
    if (!isAdminOrOwner || !c2cSummary) return null

    let spTotal = 0
    let trolleyTotal = 0
    let totalWeight = 0
    let weightWithMachine = 0
    let labourSubtotal = 0
    let powderTotal = 0
    let costBase = 0
    let beforeGst = 0
    let sales = 0

    const projectTransport = overrides['Total cal!N24'] !== undefined ? Number(overrides['Total cal!N24']) : (wb.num('Total cal', 'N24') || 0)
    const projectAccommodation = overrides['Total cal!N25'] !== undefined ? Number(overrides['Total cal!N25']) : (wb.num('Total cal', 'N25') || 0)
    const labourPerBody = overrides['Total cal!B49'] !== undefined ? Number(overrides['Total cal!B49']) : (wb.num('Total cal', 'B49') || 300)

    systems.forEach((sys, idx) => {
      const isCurrent = idx === activeSystemIdx
      const q = Math.max(1, Number(sys.quantity) || 1)

      const targetWb = c2cWb || wb
      const sysCustomC2C = isCurrent
        ? { ...(project.summary?.c2c_overrides || {}), ...(sys.c2c_overrides || {}) }
        : (sys.c2c_overrides || {})
      const sysCustomC2CFO = isCurrent
        ? { ...(project.summary?.c2c_formula_overrides || {}), ...(sys.c2c_formula_overrides || {}) }
        : (sys.c2c_formula_overrides || {})

      const sysOv: Overrides = {
        ...filterCutToCutOverrides(isCurrent ? overrides : (sys.overrides || {})),
        ...sysCustomC2C,
      }
      // In set-wise multi-system mode, Transport and Accommodation are shared project logistics added once at project level
      sysOv['Total cal!N24'] = 0
      sysOv['Total cal!N25'] = 0
      sysOv['Total cal!B49'] = labourPerBody

      const sysExtra = isCurrent
        ? extraAmount
        : (sys.extraAmount !== undefined
          ? Number(sys.extraAmount)
          : (sys.overrides?.['__extra_amount'] !== undefined ? Number(sys.overrides['__extra_amount']) : 0))
      sysOv['__extra_amount'] = sysExtra
      sysOv['__include_installation'] = includeInstallation ? 1 : 0

      targetWb.setFormulaOverrides(isCurrent ? { ...formulaOverrides, ...(project.summary?.c2c_formula_overrides || {}) } : (sys.formulaOverrides || {}))
      targetWb.setOverrides(sysOv)
      const sum = summarize(targetWb, profitPct, discountPct, includeInstallation, sysExtra)

      // C2C Installation = bodies * labour per body rate entered by user
      const sysBodies = computeSystemTotalBodies(sys, sysOv, targetWb, isCurrent, sum.bodies)
      const sysInstallation = includeInstallation ? (sysBodies * labourPerBody) : 0
      const sysExtraTot = sysExtra * (sum.models?.reduce((a, m) => a + m.qty, 0) || 1)
      const sysCostBase = sum.subtotal + sysInstallation + sysExtraTot

      spTotal += (sum.steel?.spTotal || 0) * q
      trolleyTotal += (sum.trolley?.total || 0) * q
      totalWeight += (sum.totalWeight || 0) * q
      weightWithMachine += (sum.weightWithMachine || 0) * q
      labourSubtotal += (sum.labourSubtotal || 0) * q
      powderTotal += (sum.powderTotal || 0) * q
      costBase += sysCostBase * q
      beforeGst += (sum.beforeGst || 0) * q
      sales += (sum.sales || 0) * q
    })

    // Add single shared project Transport and Accommodation once to total combined C2C cost base
    costBase += projectTransport + projectAccommodation

    // Reset targetWb back to current active system state
    const targetWb = c2cWb || wb
    targetWb.setFormulaOverrides(formulaOverrides)
    targetWb.setOverrides(overrides)



    return {

      spTotal,

      trolleyTotal,

      totalWeight,

      weightWithMachine,

      labourSubtotal,

      powderTotal,

      costBase,

      beforeGst,

      sales: Math.round(sales),

    }

  }, [isAdminOrOwner, c2cSummary, systems, activeSystemIdx, overrides, formulaOverrides, profitPct, discountPct, includeInstallation, extraAmount, wb])



  // Cut-to-Cut 3-Line Analysis (Profit, Total, Ratio)

  // Line 1: Profit = Actual Amount with Profit - Cut-to-Cut Pure Cost Base

  // Line 2: Total = Profit + Cut-to-Cut Labour Charges

  // Line 3: Ratio = Total ÷ Cut-to-Cut Steel Weight

  // Single active system C2C metrics
  const singleC2CProfit = c2cSummary
    ? (summary.beforeGst * activeSysQty) - (c2cSummary.costBase * activeSysQty)
    : 0

  const singleC2CTotal = c2cSummary
    ? singleC2CProfit + (c2cSummary.labourSubtotal * activeSysQty)
    : 0

  const singleC2CWeight = (c2cSummary?.totalWeight || 0) * activeSysQty

  const singleC2CRatio = singleC2CWeight > 0 ? singleC2CTotal / singleC2CWeight : 0

  // Combined C2C metrics across all systems and quantities
  const allC2CProfit = (c2cAllSystemsTotals && allSystemsTotals)
    ? (allSystemsTotals.beforeGst - c2cAllSystemsTotals.costBase)
    : singleC2CProfit

  const allC2CTotal = c2cAllSystemsTotals
    ? allC2CProfit + c2cAllSystemsTotals.labourSubtotal
    : singleC2CTotal

  const allC2CWeight = c2cAllSystemsTotals ? c2cAllSystemsTotals.totalWeight : singleC2CWeight

  const allC2CRatio = allC2CWeight > 0 ? allC2CTotal / allC2CWeight : 0

  const isMultiScope = activeSysQty > 1 || systems.length > 1
  const c2cProfit = isMultiScope && allC2CProfit !== null ? allC2CProfit : singleC2CProfit
  const c2cTotal = isMultiScope && allC2CTotal !== null ? allC2CTotal : singleC2CTotal
  const c2cRatio = isMultiScope && allC2CRatio !== null ? allC2CRatio : singleC2CRatio
  const c2cTotalWeight = isMultiScope && allC2CWeight !== null ? allC2CWeight : singleC2CWeight

  // Top widget displays aggregate totals across all systems/sets
  const displayC2CProfit = isMultiScope ? allC2CProfit : singleC2CProfit
  const displayC2CTotal = isMultiScope ? allC2CTotal : singleC2CTotal
  const displayC2CRatio = isMultiScope ? allC2CRatio : singleC2CRatio



  const activeSysSales = Math.round((compiledSets[activeSystemIdx]?.amount || summary.beforeGst * activeSysQty) * 1.18)

  // Top KPI widgets display: Values come from Cut-to-Cut (c2c) for Admin/Owner as requested
  const kpiSummary: Summary = (isAdminOrOwner && c2cSummary) ? c2cSummary : summary
  const kpiAllSystemsTotals = (isAdminOrOwner && c2cAllSystemsTotals) ? c2cAllSystemsTotals : allSystemsTotals
  const kpiActiveSysSales = Math.round((kpiSummary.sales || 0) * activeSysQty)

  const displaySummary: Summary = kpiSummary
  const displayAllSystemsTotals = kpiAllSystemsTotals

  return (

    <div className="calculator-horizontal-page" data-watermark={watermarkText}>

      {/* Quotation Lock Alert Banner */}

      {isApprovedLocked && (

        <div className="quotation-locked-banner">

          <span className="locked-icon">🔒</span>

          <div className="locked-text">

            <strong>Quotation Won & Locked:</strong> This quotation has been marked as Won. Cell dimensions, quantities, profit margins, and discounts are locked as read-only. To modify values, click <b>Revise ({revInfo.nextRev})</b> or set status to Draft.

          </div>

        </div>

      )}



      {/* Pending Discount Approval Request Banner for Owner / Admin */}

      {isAdminOrOwner && discountRequest?.status === 'pending' && (

        <div

          style={{

            margin: '0 0 14px 0',

            padding: '12px 16px',

            background: '#eff6ff',

            border: '1.5px solid #3b82f6',

            borderRadius: '8px',

            display: 'flex',

            alignItems: 'center',

            justifyContent: 'space-between',

            gap: '12px',

            flexWrap: 'wrap',

          }}

        >

          <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>

            <span style={{ fontSize: '20px' }}>📩</span>

            <div>

              <div style={{ fontWeight: 700, color: '#1e40af', fontSize: '13.5px' }}>

                Discount Authorization Request: {discountRequest.requested_pct}%

              </div>

              <div style={{ fontSize: '12.5px', color: '#1e3a8a', marginTop: '2px' }}>

                Staff member <b>@{discountRequest.requested_by}</b> requested a <b>{discountRequest.requested_pct}% discount</b> on this quotation.

              </div>

            </div>

          </div>

          <div style={{ display: 'flex', gap: '8px' }}>

            <button

              type="button"

              className="btn primary sm"

              style={{ background: '#16a34a', borderColor: '#15803d' }}

              onClick={handleApproveDiscountRequest}

            >

              ✓ Approve {discountRequest.requested_pct}%

            </button>

            <button

              type="button"

              className="btn danger sm"

              onClick={handleRejectDiscountRequest}

            >

              ✕ Reject

            </button>

          </div>

        </div>

      )}



      {/* ------------------------------------------------ TOP TIER 1: Clean Project Header & Controls */}
      <div className="card top-project-bar">
        <div className="top-project-main-layout">
          {/* Main Left Content: Part 1 (Upper) & Part 2 (Lower) */}
          <div className="top-project-content-left">
            {/* FIRST PART: Project Identity, Status Dropdown, Version, Saved Status */}
            <div className="top-project-first-part">
              <div className="top-name-wrapper" title="Quotation Name">
                <span className="top-name-icon">{Icon.fileText}</span>
                <input
                  ref={nameInputRef}
                  type="text"
                  className="top-name-input"
                  value={name}
                  onChange={(e) => setName(e.target.value)}
                  placeholder="Project name…"
                  disabled={isApprovedLocked}
                />
              </div>

              {displayClientName && (
                <div className="top-client-name-badge" title={`Client: ${displayClientName}`}>
                  <span className="top-client-icon">{Icon.users}</span>
                  <span className="top-client-text">{displayClientName}</span>
                </div>
              )}

              <div className="top-desc-wrapper" style={{ flex: '1 1 150px', maxWidth: '240px', minWidth: '110px' }}>
                <input
                  type="text"
                  value={projectDescription}
                  onChange={(e) => setProjectDescription(e.target.value)}
                  placeholder="Scope / description…"
                  disabled={isApprovedLocked}
                  title="Project description / scope details (searchable in estimates)"
                  style={{
                    width: '100%',
                    fontSize: '11px',
                    padding: '2px 7px',
                    borderRadius: '5px',
                    border: '1px solid #cbd5e1',
                    background: isApprovedLocked ? '#f8fafc' : '#ffffff',
                    color: '#334155',
                    outline: 'none',
                    boxSizing: 'border-box',
                  }}
                />
              </div>

              <div className="top-pills">
                <select
                  className={`status-select status-${status === 'approved' ? 'won' : status === 'archived' || status === 'lost' ? 'loss' : status}`}
                  value={status === 'draft' ? 'quoted' : status === 'lost' || status === 'archived' ? 'loss' : status === 'approved' ? 'won' : status}
                  onChange={(e) => setStatus(e.target.value as ProjectStatus)}
                  title="Quotation Status"
                >
                  <option value="quoted">Quoted</option>
                  <option value="negotiating">Negotiating</option>
                  <option value="won">Won</option>
                  <option value="loss">Loss</option>
                </select>

                {isAdminOrOwner ? (
                  <button
                    className="chip version-chip"
                    onClick={openRecalcModal}
                    title="Click to switch template version"
                  >
                    {Icon.layers} v{project.template_version}
                  </button>
                ) : (
                  <span className="chip version-chip" title="Template Version">
                    {Icon.layers} v{project.template_version}
                  </span>
                )}

                {isAdminOrOwner && hasNewerTemplate && (
                  <button
                    className="chip warn-btn"
                    onClick={openRecalcModal}
                    title="Newer template version available"
                  >
                    {Icon.refresh} v{org.current_template_version}
                  </button>
                )}

                {dirty ? (
                  <span className="chip warn">● Unsaved</span>
                ) : (
                  <span className="chip ok">{Icon.check} Saved</span>
                )}

                {isAdminOrOwner && formulaOverridesCount > 0 && (
                  <span className="chip purple font-mono">
                    {formulaOverridesCount} Rule{formulaOverridesCount !== 1 ? 's' : ''}
                  </span>
                )}

                {revInfo.currentRev ? (
                  <span className="chip primary font-mono font-bold" title={`Quotation Revision: Rev ${revInfo.currentRev}`}>
                    Rev {revInfo.currentRev}
                  </span>
                ) : null}

                {!isAdminOrOwner && (
                  <span className="chip warning font-mono">Member View</span>
                )}
              </div>
            </div>

            {/* SECOND PART: Sheets, Share, Revise (A), Copy, Format Selector, Official Quotation (PDF), Save */}
            <div className="top-project-second-part">
              <div className="top-toolbar-group">
                {isAdminOrOwner && (
                  <button className="top-toolbar-btn" onClick={handleNavigateToSheets} title="Open 15-Sheet Detail Grid">
                    {Icon.grid} Sheets
                  </button>
                )}
                <button
                  className="top-toolbar-btn"
                  onClick={openShareModal}
                  title="Share this quotation with other staff"
                >
                  {Icon.users} Share
                </button>
                <button
                  className="top-toolbar-btn"
                  onClick={() => {
                    setReviseName(parseRevision(name || project.name).nextName)
                    setShowReviseModal(true)
                  }}
                  title={`Create new revision Rev ${revInfo.nextRev}`}
                >
                  {Icon.refresh} Revise ({revInfo.nextRev})
                </button>
                <button className="top-toolbar-btn" onClick={() => setShowCopyModal(true)} title="Duplicate Project">
                  {Icon.copy} Copy
                </button>
                <button
                  className="top-toolbar-btn"
                  onClick={() => setShowReminderModal(true)}
                  style={reminder && !reminder.dismissed ? { background: '#fef3c7', borderColor: '#f59e0b', color: '#92400e', fontWeight: 700 } : undefined}
                  title={reminder && !reminder.dismissed ? `Alarm Reminder: ${reminder.note || 'Follow up'} (Due: ${new Date(reminder.date).toLocaleDateString()} at ${new Date(reminder.date).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })})` : 'Set a follow-up reminder / alarm for this quotation'}
                >
                  🔔 {reminder && !reminder.dismissed ? `${new Date(reminder.date).toLocaleDateString('en-IN', { month: 'short', day: 'numeric' })}, ${new Date(reminder.date).toLocaleTimeString('en-IN', { hour: '2-digit', minute: '2-digit' })}` : 'Reminder'}
                </button>
              </div>

              <div className="top-format-selector">
                <span className="top-format-label">Format:</span>
                <button
                  type="button"
                  className={`top-format-btn ${quotationMode === 'set_wise' ? 'active' : ''}`}
                  onClick={() => handleQuotationModeChange('set_wise')}
                  title="System / Set-Wise: SYSTEM A, B, C (1 SET rate)"
                >
                  📦 System / Set-Wise
                </button>
                <button
                  type="button"
                  className={`top-format-btn ${quotationMode === 'unit_wise' ? 'active' : ''}`}
                  onClick={() => handleQuotationModeChange('unit_wise')}
                  title="Individual Unit-Wise: Line-by-line model rates (DFMU, SFMU, SFFU) including transportation & extra amount"
                >
                  📊 Individual Unit-Wise
                </button>
              </div>

              <button
                className="top-pdf-btn"
                onClick={handleOpenOfficialQuotation}
                title="View & Print Official Myriad Storage Quotation Document"
              >
                {Icon.fileText} Official Quotation (PDF)
              </button>

              <button className="top-save-btn" disabled={saving} onClick={() => save()}>
                {saving ? <Spinner size={14} /> : Icon.save} Save
              </button>
            </div>
          </div>
        </div>
      </div>

      {/* Prominent Active Reminder / Alarm Banner */}
      {reminder && !reminder.dismissed && (
        <div
          style={{
            margin: '0 20px 14px 20px',
            background: 'linear-gradient(135deg, #fffbeb 0%, #fef3c7 100%)',
            border: '1.5px solid #f59e0b',
            borderRadius: '8px',
            padding: '10px 16px',
            display: 'flex',
            justifyContent: 'space-between',
            alignItems: 'center',
            flexWrap: 'wrap',
            gap: '10px',
            boxShadow: '0 2px 4px rgba(245, 158, 11, 0.1)',
          }}
        >
          <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
            <span style={{ fontSize: '22px' }}>⏰</span>
            <div>
              <div style={{ fontSize: '13px', fontWeight: 800, color: '#92400e' }}>
                Follow-Up Alarm Reminder: <strong>{new Date(reminder.date).toLocaleDateString('en-IN', { month: 'short', day: 'numeric', year: 'numeric' })} at {new Date(reminder.date).toLocaleTimeString('en-IN', { hour: '2-digit', minute: '2-digit' })}</strong>
              </div>
              <div style={{ fontSize: '12px', color: '#78350f', marginTop: '2px' }}>
                📌 <em>{reminder.note || 'Follow up with client on quotation'}</em>
              </div>
            </div>
          </div>
          <div style={{ display: 'flex', gap: '6px', alignItems: 'center' }}>
            <button
              type="button"
              className="btn xs ghost"
              onClick={() => setShowReminderModal(true)}
              style={{ color: '#92400e', borderColor: '#fcd34d', fontWeight: 700, fontSize: '11px', background: '#fff' }}
            >
              ✏️ Edit Alarm
            </button>
            <button
              type="button"
              className="btn xs ghost text-danger"
              onClick={handleClearReminder}
              style={{ borderColor: '#fcd34d', fontSize: '11px', background: '#fff' }}
            >
              ✕ Dismiss
            </button>
          </div>
        </div>
      )}

      {/* ------------------------------------------------ TOP TIER 2: Excel Top Metrics Bar */}

      <div className={`excel-kpi-bar ${!isAdminOrOwner ? 'staff-kpi-bar' : ''} ${serverCalculating ? 'is-calculating' : ''}`}>

        {quotationMode === 'set_wise' && systems.length > 1 ? (

          <>

            {/* Multi-System KPI Cards (SYSTEM A vs All Systems) */}

            <div className="excel-kpi-card" style={{ minWidth: '155px' }} title={bodiesBreakdown.tooltipText}>

              <span className="kpi-label">TOTAL BODIES</span>

              <strong className="kpi-val" style={{ fontSize: '13.5px', color: '#1e40af', lineHeight: 1.25 }}>

                <span style={{ fontSize: '11px', fontWeight: 600, color: '#64748b', marginRight: '4px' }}>

                  {activeSysName}{activeSysQty > 1 ? ` (${activeSysQty} ${activeSysUnit}S)` : ''}:

                </span>

                {bodiesBreakdown.totalBodiesForSets} Bodies
                {activeSysQty > 1 && (
                  <span style={{ fontSize: '10px', fontWeight: 600, color: '#64748b', marginLeft: '4px' }}>
                    ({bodiesBreakdown.singleSetBodies}/set)
                  </span>
                )}

              </strong>

              <span className="kpi-sub" style={{ fontWeight: 700, color: '#047857', fontSize: '10px', marginTop: '2px', display: 'flex', flexDirection: 'column', gap: '1px' }}>

                <span style={{ whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>
                  {bodiesBreakdown.summaryText}
                </span>

                <span style={{ fontSize: '10px', color: '#64748b' }}>
                  All Systems: <strong>{totalCombinedBodies} Bodies</strong>
                </span>

              </span>

            </div>



            <div className="excel-kpi-card">
              <span className="kpi-label">SP TOTAL</span>
              <strong className="kpi-val" style={{ fontSize: '13.5px', color: '#1e40af', lineHeight: 1.25 }}>
                <span style={{ fontSize: '11px', fontWeight: 600, color: '#64748b', marginRight: '4px' }}>
                  {activeSysName}{activeSysQty > 1 ? ` (${activeSysQty} ${activeSysUnit}S)` : ''}:
                </span>
                {serverCalculating ? <span className="live-calc-pulse" style={{ color: '#2563eb' }}>—</span> : kg(displaySummary.steel.spTotal * activeSysQty)}
              </strong>
              <span className="kpi-sub" style={{ fontWeight: 700, color: '#047857', fontSize: '11px', marginTop: '2px' }}>
                <span style={{ fontSize: '10px', color: '#64748b', marginRight: '4px' }}>All Systems:</span>
                {serverCalculating ? <span className="live-calc-pulse" style={{ color: '#047857' }}>—</span> : kg(displayAllSystemsTotals.spTotal)}
              </span>
            </div>

            <div className="excel-kpi-card">
              <span className="kpi-label">TROLLY STEEL</span>
              <strong className="kpi-val" style={{ fontSize: '13.5px', color: '#1e40af', lineHeight: 1.25 }}>
                <span style={{ fontSize: '11px', fontWeight: 600, color: '#64748b', marginRight: '4px' }}>
                  {activeSysName}{activeSysQty > 1 ? ` (${activeSysQty} ${activeSysUnit}S)` : ''}:
                </span>
                {serverCalculating ? <span className="live-calc-pulse" style={{ color: '#2563eb' }}>—</span> : kg(displaySummary.trolley.total * activeSysQty)}
              </strong>
              <span className="kpi-sub" style={{ fontWeight: 700, color: '#047857', fontSize: '11px', marginTop: '2px' }}>
                <span style={{ fontSize: '10px', color: '#64748b', marginRight: '4px' }}>All Systems:</span>
                {serverCalculating ? <span className="live-calc-pulse" style={{ color: '#047857' }}>—</span> : kg(displayAllSystemsTotals.trolleyTotal)}
              </span>
            </div>

            <div className="excel-kpi-card">
              <span className="kpi-label">TOTAL WEIGHT</span>
              <strong className="kpi-val" style={{ fontSize: '13.5px', color: '#1e40af', lineHeight: 1.25 }}>
                <span style={{ fontSize: '11px', fontWeight: 600, color: '#64748b', marginRight: '4px' }}>
                  {activeSysName}{activeSysQty > 1 ? ` (${activeSysQty} ${activeSysUnit}S)` : ''}:
                </span>
                {serverCalculating ? <span className="live-calc-pulse" style={{ color: '#2563eb' }}>—</span> : kg(displaySummary.totalWeight * activeSysQty)}
              </strong>
              <span className="kpi-sub" style={{ fontWeight: 700, color: '#047857', fontSize: '11px', marginTop: '2px' }}>
                <span style={{ fontSize: '10px', color: '#64748b', marginRight: '4px' }}>All Systems:</span>
                {serverCalculating ? <span className="live-calc-pulse" style={{ color: '#047857' }}>—</span> : kg(displayAllSystemsTotals.totalWeight)}
              </span>
            </div>

            <div className="excel-kpi-card highlight-mech">
              <span className="kpi-label">WT WITH MACHINE</span>
              <strong className="kpi-val" style={{ fontSize: '13.5px', color: '#1d4ed8', lineHeight: 1.25 }}>
                <span style={{ fontSize: '11px', fontWeight: 600, color: '#64748b', marginRight: '4px' }}>
                  {activeSysName}{activeSysQty > 1 ? ` (${activeSysQty} ${activeSysUnit}S)` : ''}:
                </span>
                {serverCalculating ? <span className="live-calc-pulse" style={{ color: '#2563eb' }}>—</span> : kg(displaySummary.weightWithMachine * activeSysQty)}
              </strong>
              <span className="kpi-sub" style={{ fontWeight: 700, color: '#047857', fontSize: '11px', marginTop: '2px' }}>
                <span style={{ fontSize: '10px', color: '#64748b', marginRight: '4px' }}>All Systems:</span>
                {serverCalculating ? <span className="live-calc-pulse" style={{ color: '#047857' }}>—</span> : kg(displayAllSystemsTotals.weightWithMachine)}
              </span>
            </div>

            {isAdminOrOwner ? (
              <>
                <div className="excel-kpi-card">
                  <span className="kpi-label">LABOUR</span>
                  <strong className="kpi-val" style={{ fontSize: '13.5px', color: '#1e40af', lineHeight: 1.25 }}>
                    <span style={{ fontSize: '11px', fontWeight: 600, color: '#64748b', marginRight: '4px' }}>
                      {activeSysName}{activeSysQty > 1 ? ` (${activeSysQty} ${activeSysUnit}S)` : ''}:
                    </span>
                    {serverCalculating ? <span className="live-calc-pulse" style={{ color: '#2563eb' }}>—</span> : inr(displaySummary.labourSubtotal * activeSysQty)}
                  </strong>
                  <span className="kpi-sub" style={{ fontWeight: 700, color: '#047857', fontSize: '11px', marginTop: '2px' }}>
                    <span style={{ fontSize: '10px', color: '#64748b', marginRight: '4px' }}>All Systems:</span>
                    {serverCalculating ? <span className="live-calc-pulse" style={{ color: '#047857' }}>—</span> : inr(displayAllSystemsTotals.labourSubtotal)}
                  </span>
                </div>

                <div className="excel-kpi-card">
                  <span className="kpi-label">POWDER COATING</span>
                  <strong className="kpi-val" style={{ fontSize: '13.5px', color: '#1e40af', lineHeight: 1.25 }}>
                    <span style={{ fontSize: '11px', fontWeight: 600, color: '#64748b', marginRight: '4px' }}>
                      {activeSysName}{activeSysQty > 1 ? ` (${activeSysQty} ${activeSysUnit}S)` : ''}:
                    </span>
                    {serverCalculating ? <span className="live-calc-pulse" style={{ color: '#2563eb' }}>—</span> : inr(displaySummary.powderTotal * activeSysQty)}
                  </strong>
                  <span className="kpi-sub" style={{ fontWeight: 700, color: '#047857', fontSize: '11px', marginTop: '2px' }}>
                    <span style={{ fontSize: '10px', color: '#64748b', marginRight: '4px' }}>All Systems:</span>
                    {serverCalculating ? <span className="live-calc-pulse" style={{ color: '#047857' }}>—</span> : inr(displayAllSystemsTotals.powderTotal)}
                  </span>
                </div>



                {isAdminOrOwner ? (
                  <div
                    className="excel-kpi-card highlight-c2c"
                    style={{
                      display: 'flex',
                      flexDirection: 'column',
                      justifyContent: 'center',
                      minWidth: '175px',
                      padding: '4px 10px',
                      gap: '2px',
                    }}
                    title="Cut-to-Cut Calculation (Rate Rs 79/kg)&#10;Line 1: Profit (Actual Before GST - C2C Direct Cost Base)&#10;Line 2: Total (Profit + C2C Labour)&#10;Line 3: Ratio (Total / C2C Weight)&#10;Direct Cost Base includes: Ex-Works + Transport + Accommodation + Installation + Extra"
                  >
                    <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 2 }}>
                      <span style={{ fontSize: '9.5px', fontWeight: 800, color: '#7c3aed', letterSpacing: '0.04em' }}>
                        {systems.length > 1 ? `C2C ANALYSIS (${systems.length} SYS)` : 'C2C ANALYSIS'}
                      </span>
                      {serverCalculating ? (
                        <span style={{ fontSize: '9px', fontWeight: 700, color: '#d97706', display: 'inline-flex', alignItems: 'center', gap: 3 }}>
                          <span className="spinner xs" style={{ width: 8, height: 8 }} /> calculating…
                        </span>
                      ) : (
                        <span style={{ fontSize: '9px', fontWeight: 700, color: '#16a34a' }}>✓ synced</span>
                      )}
                    </div>

                    <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'baseline', gap: '8px', whiteSpace: 'nowrap' }}>
                      <span style={{ fontSize: '10.5px', color: '#475569', fontWeight: 700, textTransform: 'uppercase', letterSpacing: '0.02em' }}>Profit:</span>
                      <span style={{ fontSize: '13px', fontWeight: 800, color: '#15803d', fontFamily: 'monospace' }}>
                        {serverCalculating ? <span style={{ fontSize: '11px', color: '#15803d', fontStyle: 'italic' }}>calculating…</span> : inr(Math.round(displayC2CProfit))}
                      </span>
                    </div>

                    <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'baseline', gap: '8px', whiteSpace: 'nowrap' }}>
                      <span style={{ fontSize: '10.5px', color: '#475569', fontWeight: 700, textTransform: 'uppercase', letterSpacing: '0.02em' }}>Total:</span>
                      <span style={{ fontSize: '13px', fontWeight: 800, color: '#1d4ed8', fontFamily: 'monospace' }}>
                        {serverCalculating ? <span style={{ fontSize: '11px', color: '#1d4ed8', fontStyle: 'italic' }}>calculating…</span> : inr(Math.round(displayC2CTotal))}
                      </span>
                    </div>

                    <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'baseline', gap: '8px', whiteSpace: 'nowrap' }}>
                      <span style={{ fontSize: '10.5px', color: '#475569', fontWeight: 700, textTransform: 'uppercase', letterSpacing: '0.02em' }}>Ratio:</span>
                      <span style={{ fontSize: '13px', fontWeight: 800, color: '#7c3aed', fontFamily: 'monospace' }}>
                        ₹{displayC2CRatio.toFixed(2)}<span style={{ fontSize: '10px', fontWeight: 600, color: '#6b21a8' }}>/kg</span>
                      </span>
                    </div>
                  </div>
                ) : (
                  <div className="excel-kpi-card highlight">
                    <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                      <span className="kpi-label">TOTAL QUOTATION AMOUNT</span>
                      {serverCalculating ? (
                        <span className="kpi-calculating-watermark">
                          <span className="spinner xs" style={{ width: 10, height: 10 }} /> calculating…
                        </span>
                      ) : (
                        <span style={{ fontSize: '10px', color: '#16a34a', fontWeight: 700 }}>✓ Authoritative</span>
                      )}
                    </div>
                    <strong className="kpi-val sales-val" style={{ fontSize: '14.5px', lineHeight: 1.25 }}>
                      <span style={{ fontSize: '11px', fontWeight: 600, color: '#64748b', marginRight: '4px' }}>
                        {activeSysName}{activeSysQty > 1 ? ` (${activeSysQty} ${activeSysUnit}S)` : ''}:
                      </span>
                      {serverCalculating ? <span style={{ color: '#2563eb', fontStyle: 'italic', fontSize: '12px' }}>calculating…</span> : inr(kpiActiveSysSales)}
                    </strong>
                    <span className="kpi-sub" style={{ fontWeight: 700, color: '#047857', fontSize: '11.5px', marginTop: '2px' }}>
                      <span style={{ fontSize: '10px', color: '#64748b', marginRight: '4px' }}>All Systems:</span>
                      {serverCalculating ? <span style={{ color: '#047857', fontStyle: 'italic', fontSize: '11px' }}>calculating…</span> : inr(kpiAllSystemsTotals.sales)}
                    </span>
                  </div>
                )}
              </>
            ) : (
              <div className="excel-kpi-card highlight">
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                  <span className="kpi-label">TOTAL QUOTATION AMOUNT</span>
                  {serverCalculating ? (
                    <span className="kpi-calculating-watermark">
                      <span className="spinner xs" style={{ width: 10, height: 10 }} /> calculating…
                    </span>
                  ) : (
                    <span style={{ fontSize: '10px', color: '#16a34a', fontWeight: 700 }}>✓ Authoritative</span>
                  )}
                </div>
                <strong className="kpi-val sales-val" style={{ fontSize: '14.5px', lineHeight: 1.25 }}>
                  <span style={{ fontSize: '11px', fontWeight: 600, color: '#64748b', marginRight: '4px' }}>
                    {activeSysName}{activeSysQty > 1 ? ` (${activeSysQty} ${activeSysUnit}S)` : ''}:
                  </span>
                  {serverCalculating ? <span style={{ color: '#2563eb', fontStyle: 'italic', fontSize: '12px' }}>calculating…</span> : inr(activeSysSales)}
                </strong>
                <span className="kpi-sub" style={{ fontWeight: 700, color: '#047857', fontSize: '11.5px', marginTop: '2px' }}>
                  <span style={{ fontSize: '10px', color: '#64748b', marginRight: '4px' }}>All Systems:</span>
                  {inr(allSystemsTotals.sales)}
                </span>
              </div>
            )}
          </>
        ) : (
          <>
            {/* Pure Single Calculator Mode (when in individual unit-wise or single system) */}
            <div className="excel-kpi-card" style={{ minWidth: '155px' }} title={bodiesBreakdown.tooltipText}>
              <span className="kpi-label">TOTAL BODIES</span>
              <strong className="kpi-val" style={{ fontSize: '13.5px', color: '#1e40af', lineHeight: 1.25 }}>
                {bodiesBreakdown.totalBodiesForSets} Bodies
                {activeSysQty > 1 && (
                  <span style={{ fontSize: '10px', fontWeight: 600, color: '#64748b', marginLeft: '4px' }}>
                    ({bodiesBreakdown.singleSetBodies}/set)
                  </span>
                )}
              </strong>
              <span className="kpi-sub" style={{ fontWeight: 700, color: '#047857', fontSize: '10px', marginTop: '2px', display: 'flex', flexDirection: 'column', gap: '1px' }}>
                <span style={{ whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>
                  {bodiesBreakdown.summaryText}
                </span>
                <span style={{ fontSize: '10px', color: '#64748b' }}>
                  {count(displaySummary.compartments * (activeSysQty || 1))} Compartments{activeSysQty > 1 ? ` (${activeSysQty} Sets)` : ''}
                </span>
              </span>
            </div>

            <div className="excel-kpi-card">
              <span className="kpi-label">SP TOTAL</span>
              <strong className="kpi-val">{serverCalculating ? <span className="live-calc-pulse" style={{ color: '#2563eb' }}>—</span> : kg(displaySummary.steel.spTotal * activeSysQty)}</strong>
              <span className="kpi-sub">{isAdminOrOwner ? 'C2C Steel Weight' : 'Body Steel Weight'}{activeSysQty > 1 ? ` (${activeSysQty} Sets)` : ''}</span>
            </div>

            <div className="excel-kpi-card">
              <span className="kpi-label">TROLLY STEEL</span>
              <strong className="kpi-val">{serverCalculating ? <span className="live-calc-pulse" style={{ color: '#2563eb' }}>—</span> : kg(displaySummary.trolley.total * activeSysQty)}</strong>
              <span className="kpi-sub">{isAdminOrOwner ? 'C2C Trolley Weight' : 'Trolley Weight'}{activeSysQty > 1 ? ` (${activeSysQty} Sets)` : ''}</span>
            </div>

            <div className="excel-kpi-card">
              <span className="kpi-label">TOTAL WEIGHT</span>
              <strong className="kpi-val">{serverCalculating ? <span className="live-calc-pulse" style={{ color: '#2563eb' }}>—</span> : kg(displaySummary.totalWeight * activeSysQty)}</strong>
              <span className="kpi-sub">{isAdminOrOwner ? 'C2C SP + Trolley' : 'SP + Trolley Steel'}{activeSysQty > 1 ? ` (${activeSysQty} Sets)` : ''}</span>
            </div>

            <div className="excel-kpi-card highlight-mech">
              <span className="kpi-label">WT WITH MACHINE</span>
              <strong className="kpi-val">{serverCalculating ? <span className="live-calc-pulse" style={{ color: '#2563eb' }}>—</span> : kg(displaySummary.weightWithMachine * activeSysQty)}</strong>
              <span className="kpi-sub">{isAdminOrOwner ? 'C2C Total + Mech' : 'Total + Mechanism'}{activeSysQty > 1 ? ` (${activeSysQty} Sets)` : ''}</span>
            </div>

            {isAdminOrOwner ? (
              <>
                <div className="excel-kpi-card">
                  <span className="kpi-label">LABOUR</span>
                  <strong className="kpi-val">{serverCalculating ? <span className="live-calc-pulse" style={{ color: '#2563eb' }}>—</span> : inr(displaySummary.labourSubtotal * activeSysQty)}</strong>
                  <span className="kpi-sub">C2C Total Labour{activeSysQty > 1 ? ` (${activeSysQty} Sets)` : ''}</span>
                </div>

                <div className="excel-kpi-card">
                  <span className="kpi-label">POWDER COATING</span>
                  <strong className="kpi-val">{serverCalculating ? <span className="live-calc-pulse" style={{ color: '#2563eb' }}>—</span> : inr(displaySummary.powderTotal * activeSysQty)}</strong>
                  <span className="kpi-sub">C2C PC Amount{activeSysQty > 1 ? ` (${activeSysQty} Sets)` : ''}</span>
                </div>

                {isAdminOrOwner ? (
                  <div
                    className="excel-kpi-card highlight-c2c"
                    style={{
                      display: 'flex',
                      flexDirection: 'column',
                      justifyContent: 'center',
                      minWidth: '175px',
                      padding: '4px 10px',
                      gap: '2px',
                    }}
                    title="Cut-to-Cut Calculation (Rate Rs 79/kg)&#10;Line 1: Profit (Actual Before GST - C2C Direct Cost Base)&#10;Line 2: Total (Profit + C2C Labour)&#10;Line 3: Ratio (Total / C2C Weight)&#10;Direct Cost Base includes: Ex-Works + Transport + Accommodation + Installation + Extra"
                  >
                    <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 2 }}>
                      <span style={{ fontSize: '9.5px', fontWeight: 800, color: '#7c3aed', letterSpacing: '0.04em' }}>
                        C2C ANALYSIS
                      </span>
                      {serverCalculating ? (
                        <span style={{ fontSize: '9px', fontWeight: 700, color: '#d97706', display: 'inline-flex', alignItems: 'center', gap: 3 }}>
                          <span className="spinner xs" style={{ width: 8, height: 8 }} /> calculating…
                        </span>
                      ) : (
                        <span style={{ fontSize: '9px', fontWeight: 700, color: '#16a34a' }}>✓ synced</span>
                      )}
                    </div>

                    <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'baseline', gap: '8px', whiteSpace: 'nowrap' }}>
                      <span style={{ fontSize: '10.5px', color: '#475569', fontWeight: 700, textTransform: 'uppercase', letterSpacing: '0.02em' }}>Profit:</span>
                      <span style={{ fontSize: '13px', fontWeight: 800, color: '#15803d', fontFamily: 'monospace' }}>
                        {serverCalculating ? <span style={{ fontSize: '11px', color: '#15803d', fontStyle: 'italic' }}>calculating…</span> : inr(Math.round(displayC2CProfit))}
                      </span>
                    </div>

                    <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'baseline', gap: '8px', whiteSpace: 'nowrap' }}>
                      <span style={{ fontSize: '10.5px', color: '#475569', fontWeight: 700, textTransform: 'uppercase', letterSpacing: '0.02em' }}>Total:</span>
                      <span style={{ fontSize: '13px', fontWeight: 800, color: '#1d4ed8', fontFamily: 'monospace' }}>
                        {serverCalculating ? <span style={{ fontSize: '11px', color: '#1d4ed8', fontStyle: 'italic' }}>calculating…</span> : inr(Math.round(displayC2CTotal))}
                      </span>
                    </div>

                    <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'baseline', gap: '8px', whiteSpace: 'nowrap' }}>
                      <span style={{ fontSize: '10.5px', color: '#475569', fontWeight: 700, textTransform: 'uppercase', letterSpacing: '0.02em' }}>Ratio:</span>
                      <span style={{ fontSize: '13px', fontWeight: 800, color: '#7c3aed', fontFamily: 'monospace' }}>
                        ₹{displayC2CRatio.toFixed(2)}<span style={{ fontSize: '10px', fontWeight: 600, color: '#6b21a8' }}>/kg</span>
                      </span>
                    </div>
                  </div>
                ) : (
                  <div className="excel-kpi-card highlight">
                    <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                      <span className="kpi-label">TOTAL QUOTATION AMOUNT</span>
                      {serverCalculating ? (
                        <span className="kpi-calculating-watermark">
                          <span className="spinner xs" style={{ width: 10, height: 10 }} /> calculating…
                        </span>
                      ) : (
                        <span style={{ fontSize: '10px', color: '#16a34a', fontWeight: 700 }}>✓ Authoritative</span>
                      )}
                    </div>
                    <strong className="kpi-val sales-val">
                      {serverCalculating ? <span className="live-calc-pulse" style={{ fontSize: '18px', color: '#2563eb' }}>—</span> : inr(Math.round(displaySummary.sales * activeSysQty))}
                    </strong>
                    <span className="kpi-sub">
                      {serverCalculating ? '—' : `${inr2(summary.beforeGst * activeSysQty)} before GST ${summary.discountPct > 0 ? `(${summary.discountPct}% disc)` : ''}${activeSysQty > 1 ? ` (${activeSysQty} Sets)` : ''}`}
                    </span>
                  </div>
                )}
              </>
            ) : (
              <div className="excel-kpi-card highlight">
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                  <span className="kpi-label">TOTAL QUOTATION AMOUNT</span>
                  {serverCalculating ? (
                    <span className="kpi-calculating-watermark">
                      <span className="spinner xs" style={{ width: 10, height: 10 }} /> calculating…
                    </span>
                  ) : (
                    <span style={{ fontSize: '10px', color: '#16a34a', fontWeight: 700 }}>✓ Authoritative</span>
                  )}
                </div>
                <strong className="kpi-val sales-val">
                  {serverCalculating ? <span className="live-calc-pulse" style={{ fontSize: '18px', color: '#2563eb' }}>—</span> : inr(Math.round(summary.sales * activeSysQty))}
                </strong>
                <span className="kpi-sub">
                  {serverCalculating ? '—' : `${inr2(summary.beforeGst * activeSysQty)} before GST ${summary.discountPct > 0 ? `(${summary.discountPct}% disc)` : ''}${activeSysQty > 1 ? ` (${activeSysQty} Sets)` : ''}`}
                </span>
              </div>
            )}
          </>

        )}

      </div>



      {/* ------------------------------------------------ Main Horizontal Calculation Matrix */}

      <div className="card excel-calc-card">

        <div className="excel-table-header">

          <div className="excel-table-title-group">

            <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
              <h3 className="excel-table-title" style={{ margin: 0 }}>Calculation Matrix</h3>
              {serverCalculating && (
                <span className="chip warn sm" style={{ fontSize: '10.5px', padding: '1px 7px', display: 'inline-flex', alignItems: 'center', gap: 4 }}>
                  <span className="spinner xs" style={{ width: 10, height: 10 }} /> Recalculating totals…
                </span>
              )}
            </div>

            <span className="tiny text-muted">

              {isAdminOrOwner

                ? 'DFMU drives baseline specifications. SFMU and SFFU allow Depth (D) and Quantity overrides.'

                : 'Adjust project dimensions and quantities below. Quotation totals update automatically.'}

            </span>

          </div>

          <div className="excel-table-header-actions" style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>

            {/* Bay Control: Drives all model rows (DFMU = Bay*2, SFMU/SFFU = Bay) */}

            <div

              className="bay-control-badge"

              style={{

                display: 'inline-flex',

                alignItems: 'center',

                gap: '6px',

                background: '#f0f9ff',

                border: '1px solid #bae6fd',

                borderRadius: '6px',

                padding: '3px 8px',

              }}

            >

              <label

                style={{

                  fontSize: '12px',

                  fontWeight: 700,

                  color: '#0369a1',

                  margin: 0,

                  whiteSpace: 'nowrap',

                  cursor: 'default',

                  display: 'flex',

                  alignItems: 'center',

                  gap: '4px',

                }}

              >

                <span>Bay =</span>

              </label>

              <input
                ref={bayInputRef}
                type="number"
                min={1}
                step={1}
                disabled={isApprovedLocked}
                value={bayInput}
                title="Bay count (Press / to jump here)"

                onChange={(e) => {

                  const raw = e.target.value

                  setBayInput(raw)

                  const parsed = parseInt(raw, 10)

                  if (!isNaN(parsed) && parsed > 0) {

                    setInput('Total cal!A3', parsed * 2)

                  }

                }}

                onBlur={() => {

                  const parsed = parseInt(bayInput, 10)

                  if (isNaN(parsed) || parsed < 1) {

                    setBayInput(String(currentBay))

                  }

                }}

                style={{

                  width: '48px',

                  height: '24px',

                  textAlign: 'center',

                  fontWeight: 700,

                  fontSize: '13px',

                  color: '#0c4a6e',

                  border: '1px solid #7dd3fc',

                  borderRadius: '4px',

                  background: '#ffffff',

                  outline: 'none',

                }}

              />

            </div>

          </div>

        </div>



        {/* System Matrix Switcher Tabs (Only visible when in System / Set-Wise mode) */}

        {quotationMode === 'set_wise' && (

          <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', padding: '10px 14px', background: '#f8fafc', borderBottom: '1px solid #e2e8f0', flexWrap: 'wrap', gap: 8 }}>

            <div style={{ display: 'flex', alignItems: 'center', gap: 6, flexWrap: 'wrap' }}>

              <span className="tiny font-bold text-navy" style={{ textTransform: 'uppercase', letterSpacing: '0.05em', marginRight: 4 }}>

                System Matrices ({systems.length}):

              </span>

              {systems.map((sys, idx) => {

                const isActive = idx === activeSystemIdx

                return (

                  <div

                    key={sys.id}

                    onClick={() => handleSwitchSystem(idx)}

                    style={{

                      display: 'inline-flex',

                      alignItems: 'center',

                      gap: 6,

                      padding: '5px 12px',

                      borderRadius: 6,

                      cursor: 'pointer',

                      border: isActive ? '2px solid #2563eb' : '1px solid #cbd5e1',

                      background: isActive ? '#eff6ff' : '#ffffff',

                      color: isActive ? '#1e40af' : '#475569',

                      fontWeight: isActive ? 700 : 500,

                      fontSize: 12,

                      boxShadow: isActive ? '0 1px 3px rgba(37,99,235,0.18)' : 'none',

                      userSelect: 'none',

                    }}

                  >

                    <span>📦 {sys.name}</span>

                    <span style={{ fontSize: 10, background: isActive ? '#dbeafe' : '#f1f5f9', padding: '1px 5px', borderRadius: 4, fontWeight: 600 }}>

                      {sys.quantity || 1} {sys.unit || 'SET'}

                    </span>

                    {systems.length > 1 && (

                      <button

                        type="button"

                        className="icon-btn xs"

                        onClick={(e) => {

                          e.stopPropagation()

                          handleRemoveSystemMatrix(idx)

                        }}

                        style={{ color: '#ef4444', padding: '1px', marginLeft: 2 }}

                        title="Delete this system matrix"

                      >

                        ✕

                      </button>

                    )}

                  </div>

                )

              })}

              <button

                type="button"

                className="btn ghost sm"

                onClick={handleAddSystemMatrix}

                style={{ fontSize: 11, padding: '3px 8px', border: '1px dashed #94a3b8' }}

                title="Add a new system matrix (e.g. SYSTEM B, SYSTEM C...)"

              >

                + Add System Matrix (SYSTEM {String.fromCharCode(65 + systems.length)})

              </button>

            </div>



            <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>

              <label className="tiny muted font-bold">Active System:</label>

              <input

                type="text"

                value={systems[activeSystemIdx]?.name ?? ''}
                placeholder={`SYSTEM ${String.fromCharCode(65 + activeSystemIdx)}`}

                onChange={(e) => handleUpdateSystemMeta(activeSystemIdx, { name: e.target.value })}

                style={{ fontSize: 12, fontWeight: 700, padding: '2px 6px', border: '1px solid #cbd5e1', borderRadius: 4, width: 130 }}

              />

              <label className="tiny muted font-bold">Sets Qty:</label>

              <input

                type="number"

                min={1}

                value={systems[activeSystemIdx]?.quantity ? Math.max(1, Number(systems[activeSystemIdx].quantity)) : 1}

                onChange={(e) => {
                  const val = e.target.value
                  if (val === '') {
                    handleUpdateSystemMeta(activeSystemIdx, { quantity: '' })
                    return
                  }
                  const num = parseInt(val, 10)
                  handleUpdateSystemMeta(activeSystemIdx, { quantity: isNaN(num) ? 1 : Math.max(1, num) })
                }}

                onBlur={() => {
                  const cur = systems[activeSystemIdx]?.quantity
                  if (cur === '' || cur === undefined || isNaN(Number(cur)) || Number(cur) < 1) {
                    handleUpdateSystemMeta(activeSystemIdx, { quantity: 1 })
                  }
                }}

                style={{ fontSize: 12, fontWeight: 700, padding: '2px 6px', border: '1px solid #cbd5e1', borderRadius: 4, width: 50, textAlign: 'center' }}

              />

              <button

                type="button"

                className="btn ghost sm"

                onClick={() => handleDuplicateSystemMatrix(activeSystemIdx)}

                style={{ fontSize: 10, padding: '2px 6px' }}

                title="Duplicate current system matrix to next system"

              >

                Duplicate

              </button>

            </div>

          </div>

        )}



        <div className="excel-table-container">

          <table className="excel-calc-table">

            <thead>

              <tr>

                <th className="col-model">Model</th>

                <th className="col-spec">ROW</th>

                <th className="col-spec">CLOUM</th>

                <th className="col-dim">HT (mm)</th>

                <th className="col-dim">W (mm)</th>

                <th className="col-dim highlight-col">D (mm)</th>

                <th className="col-gauge">Body G</th>

                <th className="col-gauge">Shelf G</th>

                <th className="col-gauge">Stiffener</th>

                <th className="col-spec">Track</th>

                <th className="col-gauge">Trolley G</th>

                {isAdminOrOwner && (

                  <>

                    <th className="col-rate highlight-rate" style={{ background: '#f0fdf4', color: '#166534', minWidth: '85px' }}>

                      Steel Rate (₹/kg)

                    </th>

                    <th className="col-rate highlight-rate" style={{ background: '#f0fdf4', color: '#166534', minWidth: '85px' }}>

                      Trolley Rate (₹/kg)

                    </th>

                  </>

                )}

                <th className="col-qty highlight-qty">QTY</th>

                {isAdminOrOwner && (

                  <>

                    <th className="col-amount text-center" style={{ minWidth: 120, background: '#e0f2fe', color: '#0369a1', fontWeight: 700, textAlign: 'center' }} title="Total Unit Rate = Base Rate + Trans/Inst/Acc + Extra Charge">

                      Total Unit Rate (₹)

                    </th>

                    <th className="col-amount text-center" style={{ minWidth: 120, background: '#f0fdf4', color: '#15803d', fontWeight: 700, textAlign: 'center' }} title="Total Amount = Total Unit Rate × Qty">

                      Total Amount (₹)

                    </th>

                  </>

                )}

              </tr>

            </thead>

            <tbody>

              {/* 1. DFMU (Master Row: All specifications except locked ROW driven by Bay) */}

              <tr className="excel-row master-row">

                <td className="col-model">

                  <div className="model-tag-box">

                    <strong>DFMU</strong>

                    <span className="master-pill">Master</span>

                  </div>

                </td>

                <td className="col-spec">{renderLinkedCell(currentDfmuRows, '=Bay*2', 'DFMU Rows = Bay × 2')}</td>

                <td className="col-spec">{renderCellInput('Total cal!B3', 1, 1)}</td>

                <td className="col-dim">{renderCellInput('Total cal!C3', 10, 0)}</td>

                <td className="col-dim">{renderCellInput('Total cal!D3', 10, 0)}</td>

                <td className="col-dim highlight-col">{renderCellInput('Total cal!E3', 5, 0)}</td>

                <td className="col-gauge">{renderCellInput('Total cal!F3', 0.1, 0)}</td>

                <td className="col-gauge">{renderCellInput('Total cal!G3', 0.1, 0)}</td>

                <td className="col-gauge">{renderCellInput('Total cal!J3', 0.1, 0)}</td>

                <td className="col-spec">{renderCellInput('Total cal!I3', 1, 0)}</td>

                <td className="col-gauge">{renderCellInput('Total cal!L3', 0.1, 0)}</td>

                {isAdminOrOwner && (

                  <>

                    <td className="col-rate highlight-rate">{renderCellInput('Total cal!H3', 1, 0, '₹', 75)}</td>

                    <td className="col-rate highlight-rate">{renderCellInput('Total cal!M3', 1, 0, '₹', 75)}</td>

                  </>

                )}

                <td className="col-qty highlight-qty">{renderCellInput('Total cal!G4', 1, 0)}</td>

                {isAdminOrOwner && (

                  <>

                    <td

                      className="col-amount text-center font-mono font-bold"

                      style={{ background: '#f0f9ff', color: '#0369a1', textAlign: 'center' }}

                      title={`Base: ${inr2(calcBreakdown.rows.find((r) => r.model_code === 'DFMU')?.unit_rate ?? 0)} + Trans/Inst/Acc: ${inr2(calcBreakdown.rows.find((r) => r.model_code === 'DFMU')?.transportation ?? 0)}${extraAmount > 0 ? ` + Extra: ${inr2(extraAmount)}` : ''}`}

                    >

                      {renderLiveVal(calcBreakdown.rows.find((r) => r.model_code === 'DFMU')?.total_unit_rate ?? (dfmuModel?.unitRate ?? activeWb.num('Total cal', 'D4')))}

                    </td>

                    <td className="col-amount text-center font-mono font-bold text-green" style={{ background: '#f0fdf4', textAlign: 'center' }}>

                      {renderLiveVal(calcBreakdown.rows.find((r) => r.model_code === 'DFMU')?.total_price ?? (dfmuModel?.amount ?? activeWb.num('Total cal', 'N4')))}

                    </td>

                  </>

                )}

              </tr>



              {/* 2. SFMU (D & Qty editable; ROW=Bay; all others linked to DFMU) */}

              <tr className="excel-row linked-row">

                <td className="col-model">

                  <div className="model-tag-box">

                    <strong>SFMU</strong>

                    <span className="sub-model-pill">Single Face</span>

                  </div>

                </td>

                <td className="col-spec">{renderLinkedCell(currentBay, '=Bay', 'SFMU Rows = Bay')}</td>

                <td className="col-spec">{renderLinkedCell(activeWb.num('Total cal', 'B3'), '=DFMU')}</td>

                <td className="col-dim">{renderLinkedCell(activeWb.num('Total cal', 'C7'), '=DFMU')}</td>

                <td className="col-dim">{renderLinkedCell(activeWb.num('Total cal', 'D7'), '=DFMU')}</td>

                <td className="col-dim highlight-col editable-cell">{renderCellInput('Total cal!E7', 5, 0)}</td>

                <td className="col-gauge">{renderLinkedCell(activeWb.num('Total cal', 'F7'), '=DFMU')}</td>

                <td className="col-gauge">{renderLinkedCell(activeWb.num('Total cal', 'G7'), '=DFMU')}</td>

                <td className="col-gauge">{renderLinkedCell(activeWb.num('Total cal', 'J7'), '=DFMU')}</td>

                <td className="col-spec">{renderLinkedCell(activeWb.num('Total cal', 'I7'), '=DFMU')}</td>

                <td className="col-gauge">{renderLinkedCell('—', '—')}</td>

                {isAdminOrOwner && (

                  <>

                    <td className="col-rate highlight-rate">{renderLinkedCell(activeWb.num('Total cal', 'H3'), '=DFMU')}</td>

                    <td className="col-rate highlight-rate">{renderLinkedCell('—', '—')}</td>

                  </>

                )}

                <td className="col-qty highlight-qty">{renderCellInput('Total cal!G8', 1, 0)}</td>

                {isAdminOrOwner && (

                  <>

                    <td

                      className="col-amount text-center font-mono font-bold"

                      style={{ background: '#f0f9ff', color: '#0369a1', textAlign: 'center' }}

                      title={`Base: ${inr2(calcBreakdown.rows.find((r) => r.model_code === 'SFMU')?.unit_rate ?? 0)} + Trans/Inst/Acc: ${inr2(calcBreakdown.rows.find((r) => r.model_code === 'SFMU')?.transportation ?? 0)}${extraAmount > 0 ? ` + Extra: ${inr2(extraAmount)}` : ''}`}

                    >

                      {renderLiveVal(calcBreakdown.rows.find((r) => r.model_code === 'SFMU')?.total_unit_rate ?? (sfmuModel?.unitRate ?? activeWb.num('Total cal', 'D8')))}

                    </td>

                    <td className="col-amount text-center font-mono font-bold text-green" style={{ background: '#f0fdf4', textAlign: 'center' }}>

                      {renderLiveVal(calcBreakdown.rows.find((r) => r.model_code === 'SFMU')?.total_price ?? (sfmuModel?.amount ?? activeWb.num('Total cal', 'N8')))}

                    </td>

                  </>

                )}

              </tr>



              {/* 3. SFFU (D & Qty editable; ROW=Bay; all others linked to DFMU) */}

              <tr className="excel-row linked-row">

                <td className="col-model">

                  <div className="model-tag-box">

                    <strong>SFFU</strong>

                    <span className="sub-model-pill">Fixed Unit</span>

                  </div>

                </td>

                <td className="col-spec">{renderLinkedCell(currentBay, '=Bay', 'SFFU Rows = Bay')}</td>

                <td className="col-spec">{renderLinkedCell(activeWb.num('Total cal', 'B3'), '=DFMU')}</td>

                <td className="col-dim">{renderLinkedCell(activeWb.num('Total cal', 'C11'), '=DFMU')}</td>

                <td className="col-dim">{renderLinkedCell(activeWb.num('Total cal', 'D11'), '=DFMU')}</td>

                <td className="col-dim highlight-col editable-cell">{renderCellInput('Total cal!E11', 5, 0)}</td>

                <td className="col-gauge">{renderLinkedCell(activeWb.num('Total cal', 'F11'), '=DFMU')}</td>

                <td className="col-gauge">{renderLinkedCell(activeWb.num('Total cal', 'G11'), '=DFMU')}</td>

                <td className="col-gauge">{renderLinkedCell(activeWb.num('Total cal', 'J11'), '=DFMU')}</td>

                <td className="col-spec">{renderLinkedCell(activeWb.num('Total cal', 'I11'), '=DFMU')}</td>

                <td className="col-gauge">{renderLinkedCell('—', '—')}</td>

                {isAdminOrOwner && (

                  <>

                    <td className="col-rate highlight-rate">{renderLinkedCell(activeWb.num('Total cal', 'H3'), '=DFMU')}</td>

                    <td className="col-rate highlight-rate">{renderLinkedCell('—', '—')}</td>

                  </>

                )}

                <td className="col-qty highlight-qty">{renderCellInput('Total cal!G12', 1, 0)}</td>

                {isAdminOrOwner && (

                  <>

                    <td

                      className="col-amount text-center font-mono font-bold"

                      style={{ background: '#f0f9ff', color: '#0369a1', textAlign: 'center' }}

                      title={`Base: ${inr2(calcBreakdown.rows.find((r) => r.model_code === 'SFFU')?.unit_rate ?? 0)} + Trans/Inst/Acc: ${inr2(calcBreakdown.rows.find((r) => r.model_code === 'SFFU')?.transportation ?? 0)}${extraAmount > 0 ? ` + Extra: ${inr2(extraAmount)}` : ''}`}

                    >

                      {renderLiveVal(calcBreakdown.rows.find((r) => r.model_code === 'SFFU')?.total_unit_rate ?? (sffuModel?.unitRate ?? activeWb.num('Total cal', 'D12')))}

                    </td>

                    <td className="col-amount text-center font-mono font-bold text-green" style={{ background: '#f0fdf4', textAlign: 'center' }}>

                      {renderLiveVal(calcBreakdown.rows.find((r) => r.model_code === 'SFFU')?.total_price ?? (sffuModel?.amount ?? activeWb.num('Total cal', 'N12')))}

                    </td>

                  </>

                )}

              </tr>



              {/* 4. DFFU (Qty editable; ROW=Bay*2; all others linked to DFMU) */}

              <tr className="excel-row linked-row">

                <td className="col-model">

                  <div className="model-tag-box">

                    <strong>DFFU</strong>

                    <span className="sub-model-pill">Fixed Unit</span>

                  </div>

                </td>

                <td className="col-spec">{renderLinkedCell(currentDfmuRows, '=Bay*2', 'DFFU Rows = Bay × 2')}</td>

                <td className="col-spec">{renderLinkedCell(activeWb.num('Total cal', 'B3'), '=DFMU')}</td>

                <td className="col-dim">{renderLinkedCell(activeWb.num('Total cal', 'C15'), '=DFMU')}</td>

                <td className="col-dim">{renderLinkedCell(activeWb.num('Total cal', 'D15'), '=DFMU')}</td>

                <td className="col-dim highlight-col">{renderLinkedCell(activeWb.num('Total cal', 'E15'), '=DFMU')}</td>

                <td className="col-gauge">{renderLinkedCell(activeWb.num('Total cal', 'F15'), '=DFMU')}</td>

                <td className="col-gauge">{renderLinkedCell(activeWb.num('Total cal', 'G15'), '=DFMU')}</td>

                <td className="col-gauge">{renderLinkedCell(activeWb.num('Total cal', 'J15'), '=DFMU')}</td>

                <td className="col-spec">{renderLinkedCell(activeWb.num('Total cal', 'I15'), '=DFMU')}</td>

                <td className="col-gauge">{renderLinkedCell(activeWb.num('Total cal', 'L15'), '=DFMU')}</td>

                {isAdminOrOwner && (

                  <>

                    <td className="col-rate highlight-rate">{renderLinkedCell(activeWb.num('Total cal', 'H3'), '=DFMU')}</td>

                    <td className="col-rate highlight-rate">{renderLinkedCell('—', '—')}</td>

                  </>

                )}

                <td className="col-qty highlight-qty">{renderCellInput('Total cal!G16', 1, 0)}</td>

                {isAdminOrOwner && (

                  <>

                    <td

                      className="col-amount text-center font-mono font-bold"

                      style={{ background: '#f0f9ff', color: '#0369a1', textAlign: 'center' }}

                      title={`Base: ${inr2(calcBreakdown.rows.find((r) => r.model_code === 'DFFU')?.unit_rate ?? 0)} + Trans/Inst/Acc: ${inr2(calcBreakdown.rows.find((r) => r.model_code === 'DFFU')?.transportation ?? 0)}${extraAmount > 0 ? ` + Extra: ${inr2(extraAmount)}` : ''}`}

                    >

                      {renderLiveVal(calcBreakdown.rows.find((r) => r.model_code === 'DFFU')?.total_unit_rate ?? (dffuModel?.unitRate ?? activeWb.num('Total cal', 'D16')))}

                    </td>

                    <td className="col-amount text-center font-mono font-bold text-green" style={{ background: '#f0fdf4', textAlign: 'center' }}>

                      {renderLiveVal(calcBreakdown.rows.find((r) => r.model_code === 'DFFU')?.total_price ?? (dffuModel?.amount ?? activeWb.num('Total cal', 'N16')))}

                    </td>

                  </>

                )}

              </tr>



              {/* 5. DFMU WITH DOOR (Qty editable; ROW=Bay*2; D=DFMU+30; others linked to DFMU) */}

              <tr className="excel-row linked-row">

                <td className="col-model">

                  <div className="model-tag-box">

                    <strong>DFMU-D</strong>

                    <span className="sub-model-pill">With Door</span>

                  </div>

                </td>

                <td className="col-spec">{renderLinkedCell(currentDfmuRows, '=Bay*2', 'DFMU-D Rows = Bay × 2')}</td>

                <td className="col-spec">{renderLinkedCell(activeWb.num('Total cal', 'B3'), '=DFMU')}</td>

                <td className="col-dim">{renderLinkedCell(activeWb.num('Total cal', 'C19'), '=DFMU')}</td>

                <td className="col-dim">{renderLinkedCell(activeWb.num('Total cal', 'D19'), '=DFMU')}</td>

                <td className="col-dim highlight-col">{renderLinkedCell(activeWb.num('Total cal', 'E19'), '=DFMU+30')}</td>

                <td className="col-gauge">{renderLinkedCell(activeWb.num('Total cal', 'F19'), '=DFMU')}</td>

                <td className="col-gauge">{renderLinkedCell(activeWb.num('Total cal', 'G19'), '=DFMU')}</td>

                <td className="col-gauge">{renderLinkedCell(activeWb.num('Total cal', 'J19'), '=DFMU')}</td>

                <td className="col-spec">{renderLinkedCell(activeWb.num('Total cal', 'I19'), '=DFMU')}</td>

                <td className="col-gauge">{renderLinkedCell(activeWb.num('Total cal', 'L19'), '=DFMU')}</td>

                {isAdminOrOwner && (

                  <>

                    <td className="col-rate highlight-rate">{renderLinkedCell(activeWb.num('Total cal', 'H3'), '=DFMU')}</td>

                    <td className="col-rate highlight-rate">{renderLinkedCell(activeWb.num('Total cal', 'M3'), '=DFMU')}</td>

                  </>

                )}

                <td className="col-qty highlight-qty">{renderCellInput('Total cal!G20', 1, 0)}</td>

                {isAdminOrOwner && (

                  <>

                    <td

                      className="col-amount text-center font-mono font-bold"

                      style={{ background: '#f0f9ff', color: '#0369a1', textAlign: 'center' }}

                      title={`Base: ${inr2(calcBreakdown.rows.find((r) => r.model_code === 'DFMU-D' || r.model_code === 'DFFU-D')?.unit_rate ?? 0)} + Trans/Inst/Acc: ${inr2(calcBreakdown.rows.find((r) => r.model_code === 'DFMU-D' || r.model_code === 'DFFU-D')?.transportation ?? 0)}${extraAmount > 0 ? ` + Extra: ${inr2(extraAmount)}` : ''}`}

                    >

                      {renderLiveVal(calcBreakdown.rows.find((r) => r.model_code === 'DFMU-D' || r.model_code === 'DFFU-D')?.total_unit_rate ?? (dfmuDModel?.unitRate ?? activeWb.num('Total cal', 'D20')))}

                    </td>

                    <td className="col-amount text-center font-mono font-bold text-green" style={{ background: '#f0fdf4', textAlign: 'center' }}>

                      {renderLiveVal(calcBreakdown.rows.find((r) => r.model_code === 'DFMU-D' || r.model_code === 'DFFU-D')?.total_price ?? (dfmuDModel?.amount ?? activeWb.num('Total cal', 'N20')))}

                    </td>

                  </>

                )}

              </tr>

            </tbody>

            <tfoot>

              <tr className="excel-foot-row">

                <td colSpan={isAdminOrOwner ? 13 : 11} className="font-bold text-muted text-center" style={{ textAlign: 'center' }}>

                  {isAdminOrOwner ? 'Total Units & Total Value:' : 'Total Units:'}

                </td>

                <td className="col-qty text-center font-bold text-navy font-mono" style={{ fontSize: 13, background: '#fef3c7', textAlign: 'center' }}>

                  {(summary?.models || []).reduce((acc, m) => acc + m.qty, 0)} nos

                </td>

                {isAdminOrOwner && (

                  <>

                    <td className="col-amount text-center text-muted" style={{ textAlign: 'center' }}>—</td>

                    <td className="col-amount text-center font-mono font-bold text-green" style={{ fontSize: 15, background: '#dcfce7', textAlign: 'center' }}>

                      {inr2(calcBreakdown.rows.reduce((sum, r) => sum + r.total_price, 0))}

                    </td>

                  </>

                )}

              </tr>

            </tfoot>

          </table>

        </div>



        {/* Clean All Systems Overview & Comparison Card (Only visible in System / Set-Wise mode) */}

        {quotationMode === 'set_wise' && systems.length > 1 && isAdminOrOwner && (

          <div style={{ marginTop: 12, padding: '12px 14px', background: '#f8fafc', borderRadius: 8, border: '1px solid #e2e8f0' }}>

            <div className="row between mb-2">

              <span className="tiny font-bold text-navy" style={{ textTransform: 'uppercase', letterSpacing: '0.05em', fontSize: 11 }}>

                📦 All Systems Combined Quotation Summary ({systems.length} Systems)

              </span>

              <span className="chip primary sm font-bold" style={{ fontSize: 10 }}>

                Single Shared Transport &amp; Labour across all {totalCombinedBodies} Bodies

              </span>

            </div>

            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(280px, 1fr))', gap: 10 }}>

              {compiledSets.map((s, idx) => (

                <div

                  key={s.id}

                  onClick={() => handleSwitchSystem(idx)}

                  style={{

                    padding: '8px 12px',

                    borderRadius: 6,

                    border: idx === activeSystemIdx ? '2px solid #2563eb' : '1px solid #cbd5e1',

                    background: idx === activeSystemIdx ? '#eff6ff' : '#ffffff',

                    cursor: 'pointer',

                  }}

                >

                  <div className="row between">

                    <strong style={{ fontSize: 12 }}>{s.system_title}</strong>

                    <span className="chip sm" style={{ fontSize: 10 }}>{s.set_label}</span>

                  </div>

                  <div className="row between mt-1">

                    <span className="tiny muted">Rate: {inr2(s.rate)}{s.quantity > 1 ? ' / SET' : ''}</span>

                    <strong className="text-green" style={{ fontSize: 12 }}>Total: {inr2(s.amount)}</strong>

                  </div>

                </div>

              ))}

            </div>

            <div style={{ display: 'flex', flexDirection: 'column', gap: 4, marginTop: 10, paddingTop: 10, borderTop: '1px solid #cbd5e1' }}>

              <div className="row between">

                <span className="tiny font-bold text-muted">Combined All Systems Total (Before GST):</span>

                <strong className="font-mono" style={{ fontSize: 13, color: '#0f172a' }}>

                  {inr2(compiledSets.reduce((sum, s) => sum + s.amount, 0))}

                </strong>

              </div>

              <div className="row between">

                <span className="tiny text-muted">GST @ 18%:</span>

                <span className="font-mono tiny font-bold text-muted">

                  +{inr2(compiledSets.reduce((sum, s) => sum + s.amount, 0) * 0.18)}

                </span>

              </div>

              <div className="row between pt-1" style={{ borderTop: '1px dashed #cbd5e1' }}>

                <strong style={{ fontSize: 13, color: '#166534' }}>Grand Total (Sales with GST):</strong>

                <strong className="font-mono text-green" style={{ fontSize: 15 }}>

                  {inr2(Math.round(compiledSets.reduce((sum, s) => sum + s.amount, 0) * 1.18))}

                </strong>

              </div>

            </div>

          </div>

        )}

      </div>



      {/* ------------------------------------------------ Bottom Split: Charges & Invoice */}

      <div className={`h-bottom-row ${!isAdminOrOwner ? 'staff-bottom-row' : ''}`}>

        {/* Additional Charges & Rates */}

        <div className="card h-charges-card">

          <div className="row between mb-2">

            <h4 className="h-card-subhead" style={{ margin: 0 }}>Labour, Transport &amp; Surcharges</h4>

            {quotationMode === 'set_wise' && systems.length > 1 && (

              <span className="chip primary sm font-bold" style={{ fontSize: 9.5 }}>

                Shared across all {systems.length} systems ({totalCombinedBodies} bodies)

              </span>

            )}

          </div>

          <div className="h-charges-grid">

            <div className="h-charge-item">

              <label>Labour per Body (₹)</label>

              {renderCellInput('Total cal!B49', 10, 0, '₹', 120)}

            </div>

            <div className="h-charge-item">
              <label>Transport / Freight (₹)</label>
              {renderCellInput('Total cal!N24', 100, 0, '₹', 120)}
              {(activeSysQty > 1 || systems.length > 1) && summary?.transport !== undefined && (
                <div style={{ fontSize: '10px', color: '#64748b', marginTop: '2px' }}>
                  ({inr2(summary.transport)}/Set)
                </div>
              )}
            </div>

            <div className="h-charge-item">
              <label>Accommodation (₹)</label>
              {renderCellInput('Total cal!N25', 100, 0, '₹', 120)}
              {(activeSysQty > 1 || systems.length > 1) && summary?.accommodation !== undefined && (
                <div style={{ fontSize: '10px', color: '#64748b', marginTop: '2px' }}>
                  ({inr2(summary.accommodation)}/Set)
                </div>
              )}
            </div>

            <div className="h-charge-item">

              <label>Other Extra Amount {systems.length > 1 ? `(${activeSysName})` : ''} (₹)</label>

              <input

                type="number"

                min={0}

                step={100}

                value={extraAmount === 0 ? '' : extraAmount}

                disabled={isApprovedLocked}

                placeholder="0"

                onChange={(e) => {

                  const val = Math.max(0, parseFloat(e.target.value) || 0)

                  setExtraAmount(val)

                }}

                className="cell-num-input font-bold"

                style={{ width: '100%', borderColor: extraAmount > 0 ? '#2563eb' : undefined }}

              />

            </div>

          </div>



          <div className="h-cost-breakdown-strip">

            <div className="cost-mini-item">

              <div className="row between" style={{ marginBottom: 2 }}>

                <span className="c-lbl">Installation</span>

                <label style={{ display: 'inline-flex', alignItems: 'center', gap: 4, cursor: isApprovedLocked ? 'not-allowed' : 'pointer', margin: 0 }}>

                  <input

                    type="checkbox"

                    checked={includeInstallation}

                    disabled={isApprovedLocked}

                    onChange={(e) => setIncludeInstallation(e.target.checked)}

                    style={{ cursor: isApprovedLocked ? 'not-allowed' : 'pointer', width: 13, height: 13 }}

                  />

                  <span className="tiny font-bold" style={{ fontSize: 9.5, color: includeInstallation ? '#2563eb' : '#94a3b8' }}>

                    {includeInstallation ? 'Including' : 'At your scope'}

                  </span>

                </label>

              </div>

              <strong className="c-val" style={!includeInstallation ? { color: '#94a3b8' } : undefined}>

                {includeInstallation ? inr(summary.installation) : '₹0 (At your scope)'}

              </strong>

            </div>

            <div className="cost-mini-item">

              <span className="c-lbl">Bodies / Compartments</span>

              <strong className="c-val">

                {activeSysQty > 1 ? `${summary.bodies * activeSysQty} (${activeSysQty} Sets)` : (systems.length > 1 ? `${summary.bodies} (Order Total: ${totalCombinedBodies})` : summary.bodies)} / {summary.compartments * activeSysQty}

              </strong>

            </div>

          </div>

          {/* Client Instructions & Project Notes */}
          <div style={{ marginTop: '12px', paddingTop: '10px', borderTop: '1px dashed #cbd5e1' }}>
            <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: '5px' }}>
              <label style={{ fontSize: '11px', fontWeight: 700, color: '#334155', display: 'flex', alignItems: 'center', gap: '5px', margin: 0 }}>
                <span>📝 Client Instructions &amp; Project Notes</span>
              </label>
              <span className="tiny text-muted" style={{ fontSize: '10px' }}>Saved with quotation</span>
            </div>
            <textarea
              value={projectNotes}
              disabled={isApprovedLocked}
              onChange={(e) => setProjectNotes(e.target.value)}
              placeholder="Write client instructions, special requirements, delivery timeline, site contact, RAL shades, or anything discussed with client…"
              rows={3}
              style={{ width: '100%', fontSize: '11.5px', padding: '7px 9px', borderRadius: '6px', border: '1px solid #cbd5e1', background: isApprovedLocked ? '#f8fafc' : '#ffffff', color: '#1e293b', resize: 'vertical', boxSizing: 'border-box', fontFamily: 'inherit', lineHeight: 1.4 }}
            />
          </div>

        </div>



        {/* Quotation & Tax Invoice Breakdown: Admin vs Staff */}

        {isAdminOrOwner ? (

          <div className="card h-summary-card">

            <div className="h-card-header">

              <h4 className="h-card-subhead">Commercial Quotation &amp; Invoice</h4>

              <span className="chip primary sm font-mono">Template v{project.template_version}</span>

            </div>

            <table className="tbl h-tax-tbl">
              <tbody>
                {systems.length > 1 ? (
                  <>
                    {compiledSets.map((s, idx) => (
                      <tr key={s.id || idx} style={{ background: idx === activeSystemIdx ? '#f0f9ff' : undefined }}>
                        <td>
                          <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                            <span style={{ fontWeight: 700, color: '#1e40af', fontSize: '13px' }}>{s.system_title}</span>
                            <span className="chip sm" style={{ fontSize: '10px' }}>{s.quantity} {s.unit || 'SET'}{s.quantity > 1 ? 'S' : ''}</span>
                            {idx === activeSystemIdx && <span className="chip primary xs" style={{ fontSize: '9px' }}>Current System</span>}
                          </div>
                          <div className="tiny muted" style={{ fontSize: '10px', marginTop: '2px' }}>
                            Base Cost: {inr2(s.breakdown?.totalCostBase || s.summary?.costBase || 0)}
                          </div>
                        </td>
                        <td className="r font-mono font-bold" style={{ verticalAlign: 'middle', fontSize: '13px' }}>
                          {inr2(s.amount)}
                        </td>
                      </tr>
                    ))}
                    <tr className="subtotal-row" style={{ background: '#f8fafc', borderTop: '1.5px solid #cbd5e1' }}>
                      <td>
                        <strong>Total Direct Cost Base ({systems.length} Systems)</strong>
                      </td>
                      <td className="r font-mono font-bold">{inr2(allSystemsTotals.costBase)}</td>
                    </tr>
                    <tr className="profit-row">
                      <td className="wide-control-cell">
                        <div className="profit-margin-control">
                          <div className="margin-title">
                            <strong>Profit Margin (All Systems)</strong>
                          </div>
                          <div className="margin-input-wrap">
                            <input
                              type="number"
                              className="margin-num-input"
                              value={profitInput}
                              disabled={isApprovedLocked}
                              min={0}
                              max={200}
                              step={1}
                              onChange={(e) => handleProfitInputChange(e.target.value)}
                              onBlur={handleProfitInputBlur}
                            />
                            <span className="margin-pct-tag">%</span>
                          </div>
                        </div>
                      </td>
                      <td className="r font-mono font-bold text-blue" style={{ verticalAlign: 'middle', fontSize: '13px' }}>
                        +{inr2(allSystemsTotals.profitAmount)}
                      </td>
                    </tr>
                    <tr>
                      <td>
                        <strong>Subtotal with Profit</strong>
                      </td>
                      <td className="r font-mono font-bold">{inr2(allSystemsTotals.subtotalWithProfit)}</td>
                    </tr>
                    <tr className="discount-row">
                      <td className="wide-control-cell">
                        <div className="profit-margin-control">
                          <div className="margin-title">
                            <strong>Discount (All Systems)</strong>
                          </div>
                          <div className="margin-input-wrap discount-border">
                            <input
                              type="number"
                              className="margin-num-input discount-text"
                              value={discountInput}
                              disabled={isApprovedLocked}
                              min={0}
                              max={100}
                              step={1}
                              onChange={(e) => handleDiscountInputChange(e.target.value)}
                              onBlur={handleDiscountInputBlur}
                            />
                            <span className="margin-pct-tag discount-tag">%</span>
                          </div>
                        </div>
                        {discountRequest?.status === 'pending' && (
                          <div style={{ marginTop: '6px', fontSize: '11.5px', color: '#1e40af', display: 'flex', alignItems: 'center', gap: '6px', flexWrap: 'wrap' }}>
                            <span>⏳ Staff @{discountRequest.requested_by} requested {discountRequest.requested_pct}% discount:</span>
                            <button
                              type="button"
                              className="btn primary xs"
                              style={{ padding: '2px 8px', fontSize: '11px', background: '#16a34a', border: 'none' }}
                              onClick={handleApproveDiscountRequest}
                            >
                              Approve
                            </button>
                            <button
                              type="button"
                              className="btn danger xs"
                              style={{ padding: '2px 8px', fontSize: '11px', border: 'none' }}
                              onClick={handleRejectDiscountRequest}
                            >
                              Reject
                            </button>
                          </div>
                        )}
                      </td>
                      <td className="r font-mono font-bold text-danger" style={{ verticalAlign: 'middle', fontSize: '13px' }}>
                        -{inr2(allSystemsTotals.discountAmount)}
                      </td>
                    </tr>
                    <tr className="subtotal-row">
                      <td>
                        <strong>Subtotal Before GST (All Systems)</strong>
                      </td>
                      <td className="r font-mono font-bold">{renderLiveVal(allSystemsTotals.beforeGst)}</td>
                    </tr>
                    <tr>
                      <td>GST (18%)</td>
                      <td className="r font-mono">{renderLiveVal(allSystemsTotals.gst)}</td>
                    </tr>
                    <tr className="final-sales-row" style={{ background: '#ecfdf5', borderTop: '2px solid #10b981' }}>
                      <td>
                        <div style={{ display: 'flex', flexDirection: 'column' }}>
                          <strong style={{ color: '#065f46', fontSize: '14px' }}>Combined Total Quotation Value ({systems.length} Systems)</strong>
                          <span className="tiny muted" style={{ fontSize: '10.5px' }}>Combined Quotation Value (Inc. 18% GST)</span>
                        </div>
                      </td>
                      <td className="r font-mono font-bold text-green" style={{ fontSize: 20, color: '#047857' }}>
                        {renderLiveVal(allSystemsTotals.sales, inr)}
                      </td>
                    </tr>
                  </>
                ) : (
                  <>
                <tr>
                  <td>Ex-Works Subtotal (N22)</td>
                  <td className="r font-mono">{renderLiveVal(summary.subtotal)}</td>
                </tr>

                {/* Packing & Forwarding REMOVED per user request */}

                <tr>

                  <td>Transport / Freight (N24)</td>

                  <td className="r font-mono">{renderLiveVal(summary.transport)}</td>

                </tr>

                <tr>

                  <td>Accommodation (N25)</td>

                  <td className="r font-mono">{renderLiveVal(summary.accommodation ?? summary.accessories)}</td>

                </tr>

                <tr style={!includeInstallation ? { opacity: 0.75, background: '#f8fafc' } : undefined}>

                  <td>

                    <div className="row between" style={{ alignItems: 'center' }}>

                      <label style={{ display: 'inline-flex', alignItems: 'center', gap: 6, cursor: isApprovedLocked ? 'not-allowed' : 'pointer', margin: 0, fontWeight: 600 }}>

                        <input

                          type="checkbox"

                          checked={includeInstallation}

                          disabled={isApprovedLocked}

                          onChange={(e) => setIncludeInstallation(e.target.checked)}

                          style={{ cursor: isApprovedLocked ? 'not-allowed' : 'pointer', width: 14, height: 14 }}

                        />

                        <span>Installation (N23)</span>

                      </label>

                      <span className={`chip sm ${includeInstallation ? 'ok' : 'muted-chip'}`} style={{ fontSize: 9.5 }}>

                        {includeInstallation ? 'Including' : 'At your scope'}

                      </span>

                    </div>

                  </td>

                  <td className="r font-mono">

                    {includeInstallation ? inr2(summary.installation) : <span className="text-muted">₹0.00</span>}

                  </td>

                </tr>

                <tr className="subtotal-row">

                  <td>

                    <strong>Direct Cost Base</strong>

                  </td>

                  <td className="r font-mono font-bold">{inr2(summary.costBase)}</td>

                </tr>

                <tr className="profit-row">

                  <td className="wide-control-cell">

                    <div className="profit-margin-control">

                      <div className="margin-title">

                        <strong>Profit Margin</strong>

                      </div>

                      <div className="margin-input-wrap">

                        <input

                          type="number"

                          className="margin-num-input"

                          value={profitInput}

                          disabled={isApprovedLocked}

                          min={0}

                          max={200}

                          step={1}

                          onChange={(e) => handleProfitInputChange(e.target.value)}

                          onBlur={handleProfitInputBlur}

                        />

                        <span className="margin-pct-tag">%</span>

                      </div>

                    </div>

                  </td>

                  <td className="r font-mono font-bold text-blue" style={{ verticalAlign: 'middle', fontSize: '13px' }}>

                    +{inr2(summary.profitAmount)}

                  </td>

                </tr>

                <tr>

                  <td>

                    <strong>Subtotal with Profit</strong>

                  </td>

                  <td className="r font-mono font-bold">{inr2(summary.subtotalWithProfit)}</td>

                </tr>

                <tr className="discount-row">

                  <td className="wide-control-cell">

                    <div className="profit-margin-control">

                      <div className="margin-title">

                        <strong>Discount</strong>

                      </div>

                      <div className="margin-input-wrap discount-border">

                        <input

                          type="number"

                          className="margin-num-input discount-text"

                          value={discountInput}

                          disabled={isApprovedLocked}

                          min={0}

                          max={100}

                          step={1}

                          onChange={(e) => handleDiscountInputChange(e.target.value)}

                          onBlur={handleDiscountInputBlur}

                        />

                        <span className="margin-pct-tag discount-tag">%</span>

                      </div>

                    </div>

                    {discountRequest?.status === 'pending' && (

                      <div style={{ marginTop: '6px', fontSize: '11.5px', color: '#1e40af', display: 'flex', alignItems: 'center', gap: '6px', flexWrap: 'wrap' }}>

                        <span>⏳ Staff @{discountRequest.requested_by} requested {discountRequest.requested_pct}% discount:</span>

                        <button

                          type="button"

                          className="btn primary xs"

                          style={{ padding: '2px 8px', fontSize: '11px', background: '#16a34a', border: 'none' }}

                          onClick={handleApproveDiscountRequest}

                        >

                          Approve

                        </button>

                        <button

                          type="button"

                          className="btn danger xs"

                          style={{ padding: '2px 8px', fontSize: '11px', border: 'none' }}

                          onClick={handleRejectDiscountRequest}

                        >

                          Reject

                        </button>

                      </div>

                    )}

                  </td>

                  <td className="r font-mono font-bold text-danger" style={{ verticalAlign: 'middle', fontSize: '13px' }}>

                    -{inr2(summary.discountAmount)}

                  </td>

                </tr>

                <tr className="subtotal-row">

                  <td>

                    <strong>Subtotal Before GST</strong>

                  </td>

                  <td className="r font-mono font-bold">{renderLiveVal(summary.beforeGst)}</td>

                </tr>

                <tr>

                  <td>GST (18%)</td>

                  <td className="r font-mono">{renderLiveVal(summary.gst)}</td>

                </tr>

                <tr className="final-sales-row">
                  <td>
                    <strong>Total Quotation Value ({activeSysName})</strong>
                  </td>
                  <td className="r font-mono font-bold text-green" style={{ fontSize: 18 }}>
                    {renderLiveVal(summary.sales, inr)}
                  </td>
                </tr>
                  </>
                )}
              </tbody>

            </table>

          </div>

        ) : (

          /* Staff / Member View: Clean Commercial Summary with Discount % and Total Amount */

          <div className="card h-summary-card">

            <div className="h-card-header">

              <h4 className="h-card-subhead">Quotation Commercials</h4>

              <span className="chip primary sm font-mono">Member View</span>

            </div>

            <table className="tbl h-tax-tbl">
              <tbody>
                {systems.length > 1 ? (
                  <>
                    {compiledSets.map((s, idx) => (
                      <tr key={s.id || idx} style={{ background: idx === activeSystemIdx ? '#f0f9ff' : undefined }}>
                        <td>
                          <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                            <span style={{ fontWeight: 700, color: '#1e40af' }}>{s.system_title}</span>
                            <span className="chip sm" style={{ fontSize: '10px' }}>{s.quantity} {s.unit || 'SET'}{s.quantity > 1 ? 'S' : ''}</span>
                            {idx === activeSystemIdx && <span className="chip primary xs" style={{ fontSize: '9px' }}>Current</span>}
                          </div>
                        </td>
                        <td className="r font-mono font-bold">{inr2(s.amount)}</td>
                      </tr>
                    ))}
                    <tr>
                      <td>Base Quotation (All Systems)</td>
                      <td className="r font-mono font-bold">{inr2(allSystemsTotals.subtotalWithProfit)}</td>
                    </tr>
                    <tr className="discount-row">
                      <td className="wide-control-cell">
                        <div className="profit-margin-control">
                          <div className="margin-title">
                            <strong>Discount % (All Systems)</strong>
                          </div>
                          <div className="margin-input-wrap discount-border">
                            <input
                              type="number"
                              className="margin-num-input discount-text"
                              value={discountInput}
                              disabled={isApprovedLocked}
                              min={0}
                              max={100}
                              step={1}
                              onChange={(e) => handleDiscountInputChange(e.target.value)}
                              onBlur={handleDiscountInputBlur}
                            />
                            <span className="margin-pct-tag discount-tag">%</span>
                          </div>
                        </div>
                      </td>
                      <td className="r font-mono font-bold text-danger" style={{ verticalAlign: 'middle', fontSize: '13px' }}>
                        -{inr2(allSystemsTotals.discountAmount)}
                      </td>
                    </tr>
                    <tr className="subtotal-row">
                      <td>
                        <strong>Subtotal Before GST (All Systems)</strong>
                      </td>
                      <td className="r font-mono font-bold text-navy">{renderLiveVal(allSystemsTotals.beforeGst)}</td>
                    </tr>
                    <tr>
                      <td>GST (18% on Before GST)</td>
                      <td className="r font-mono">{renderLiveVal(allSystemsTotals.gst)}</td>
                    </tr>
                    <tr className="final-sales-row" style={{ background: '#ecfdf5', borderTop: '2px solid #10b981' }}>
                      <td>
                        <strong style={{ fontSize: 14 }}>Combined Total Quotation Amount ({systems.length} Systems)</strong>
                      </td>
                      <td className="r font-mono font-bold text-green" style={{ fontSize: 20, color: '#047857' }}>
                        {renderLiveVal(allSystemsTotals.sales, inr)}
                      </td>
                    </tr>
                  </>
                ) : (
                  <>
                <tr>
                  <td>Base Quotation</td>
                  <td className="r font-mono font-bold">{inr2(summary.subtotalWithProfit)}</td>
                </tr>

                <tr className="discount-row">

                  <td className="wide-control-cell">

                    <div className="profit-margin-control">

                      <div className="margin-title">

                        <strong>Discount %</strong>

                      </div>

                      <div className="margin-input-wrap discount-border">

                        <input

                          type="number"

                          className="margin-num-input discount-text"

                          value={discountInput}

                          disabled={isApprovedLocked}

                          min={0}

                          max={100}

                          step={1}

                          onChange={(e) => handleDiscountInputChange(e.target.value)}

                          onBlur={handleDiscountInputBlur}

                        />

                        <span className="margin-pct-tag discount-tag">%</span>

                      </div>

                    </div>



                    {summary.discountPct > MAX_STAFF_DISCOUNT_PCT && (

                      <div style={{ marginTop: '8px', padding: '9px 12px', background: '#fffbeb', border: '1px solid #fde68a', borderRadius: '6px' }}>

                        {discountRequest?.status === 'pending' && discountRequest.requested_pct === summary.discountPct ? (

                          <div style={{ display: 'flex', alignItems: 'center', gap: '8px', color: '#92400e', fontSize: '12px' }}>

                            <span>⏳</span>

                            <div>

                              <strong>Request Sent ({discountRequest.requested_pct}%)</strong> — Waiting for Owner/Admin approval.

                            </div>

                          </div>

                        ) : discountRequest?.status === 'approved' && discountRequest.requested_pct === summary.discountPct ? (

                          <div style={{ display: 'flex', alignItems: 'center', gap: '8px', color: '#166534', fontSize: '12px' }}>

                            <span>✓</span>

                            <div>

                              <strong>{discountRequest.requested_pct}% Discount Approved</strong> by @{discountRequest.decided_by}

                            </div>

                          </div>

                        ) : discountRequest?.status === 'rejected' && discountRequest.requested_pct === summary.discountPct ? (

                          <div style={{ display: 'flex', alignItems: 'center', gap: '8px', color: '#991b1b', fontSize: '12px' }}>

                            <span>✕</span>

                            <div>

                              <strong>{discountRequest.requested_pct}% Discount Declined</strong> by @{discountRequest.decided_by}.

                            </div>

                          </div>

                        ) : (

                          <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: '8px', flexWrap: 'wrap' }}>

                            <div style={{ fontSize: '12px', color: '#92400e' }}>

                              ⚠️ <b>{summary.discountPct}%</b> exceeds standard 10% limit.

                            </div>

                            <button

                              type="button"

                              className="btn primary sm"

                              style={{ background: '#d97706', borderColor: '#b45309', fontSize: '12px', padding: '4px 10px' }}

                              onClick={() => handleSendDiscountRequest(summary.discountPct)}

                              disabled={isSendingDiscountRequest || isApprovedLocked}

                            >

                              {isSendingDiscountRequest ? 'Sending Request...' : `Send Request for ${summary.discountPct}% Discount`}

                            </button>

                          </div>

                        )}

                      </div>

                    )}

                  </td>

                  <td className="r font-mono font-bold text-danger" style={{ verticalAlign: 'middle', fontSize: '13px' }}>

                    -{inr2(summary.discountAmount)}

                  </td>

                </tr>

                <tr className="subtotal-row">

                  <td>

                    <strong>Subtotal Before GST</strong>

                  </td>

                  <td className="r font-mono font-bold text-navy">{renderLiveVal(summary.beforeGst)}</td>

                </tr>

                <tr>

                  <td>GST (18% on Before GST)</td>

                  <td className="r font-mono">{renderLiveVal(summary.gst)}</td>

                </tr>

                <tr className="final-sales-row">

                  <td>

                    <strong style={{ fontSize: 14 }}>Total Quotation Amount ({activeSysName})</strong>
                  </td>
                  <td className="r font-mono font-bold text-green" style={{ fontSize: 20 }}>
                    {renderLiveVal(summary.sales, inr)}
                  </td>
                </tr>
                  </>
                )}
              </tbody>

            </table>

          </div>

        )}

      </div>



              {/* ------------------------------------------------ Cut-to-Cut Sheet Breakdown for Admin */}

        {isAdminOrOwner && c2cSummary && (

          <div

            className="card c2c-bottom-sheets-card"

            style={{

              marginTop: '16px',

              padding: '16px 20px',

              border: '1.5px solid #d8b4fe',

              background: '#fdf4ff',

              borderRadius: '8px',

            }}

          >

            <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: '14px', flexWrap: 'wrap', gap: '8px' }}>

              <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>

                <span style={{ fontSize: '20px' }}>✂️</span>

                <div>

                  <h4 style={{ margin: 0, fontSize: '14px', fontWeight: 700, color: '#581c87', display: 'flex', alignItems: 'center', gap: 8 }}>

                    Cut-to-Cut Full Sheet Breakdown (All Sheets)
                    {serverCalculating && (
                      <span className="kpi-calculating-watermark">
                        <span className="spinner xs" style={{ width: 10, height: 10 }} /> Recalculating…
                      </span>
                    )}

                  </h4>

                  <span className="tiny muted">

                    Evaluated internal cost sheets driven live by your project matrix inputs.

                  </span>

                </div>

              </div>

              <div style={{ display: 'flex', gap: '8px', flexWrap: 'wrap' }}>
                <button
                  type="button"
                  className="btn primary sm"
                  style={{ background: '#7c3aed', borderColor: '#6d28d9', color: '#fff', fontSize: '12px', display: 'inline-flex', alignItems: 'center', gap: '6px' }}
                  onClick={handleNavigateToSheets}
                >
                  {Icon.grid} Open 15-Sheet Detail Grid
                </button>
                <button
                  type="button"
                  className="btn sm"
                  style={{ background: '#15803d', borderColor: '#166534', color: '#fff', fontSize: '12px', display: 'inline-flex', alignItems: 'center', gap: '6px', fontWeight: 600 }}
                  onClick={handleDownloadC2CExcel}
                  title="Download Cut-to-Cut Excel spreadsheet with all sheets and C2C calculations at E26:G28"
                >
                  📥 Download C2C Excel (.xlsx)
                </button>
              </div>

            </div>



            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(220px, 1fr))', gap: '10px' }}>

              <div style={{ background: '#fff', padding: '10px 14px', borderRadius: '6px', border: '1px solid #e9d5ff' }}>

                <div style={{ fontSize: '11px', color: '#64748b', fontWeight: 600 }}>C2C Ex-Works Total (N22)</div>

                <div style={{ fontSize: '15px', fontWeight: 700, color: '#581c87', fontFamily: 'monospace' }}>{inr2(c2cSummary.subtotal)}</div>

              </div>

              <div style={{ background: '#fff', padding: '10px 14px', borderRadius: '6px', border: '1px solid #e9d5ff' }}>

                <div style={{ fontSize: '11px', color: '#64748b', fontWeight: 600 }}>C2C Steel Weight (J44)</div>

                <div style={{ fontSize: '15px', fontWeight: 700, color: '#581c87', fontFamily: 'monospace' }}>{kg(c2cSummary.totalWeight)}</div>

              </div>

              <div style={{ background: '#fff', padding: '10px 14px', borderRadius: '6px', border: '1px solid #e9d5ff' }}>

                <div style={{ fontSize: '11px', color: '#64748b', fontWeight: 600 }}>C2C Weight + Mach (J45)</div>

                <div style={{ fontSize: '15px', fontWeight: 700, color: '#581c87', fontFamily: 'monospace' }}>{kg(c2cSummary.weightWithMachine)}</div>

              </div>

              <div style={{ background: '#fff', padding: '10px 14px', borderRadius: '6px', border: '1px solid #e9d5ff' }}>

                <div style={{ fontSize: '11px', color: '#64748b', fontWeight: 600 }}>C2C Total Labour (N50)</div>

                <div style={{ fontSize: '15px', fontWeight: 700, color: '#581c87', fontFamily: 'monospace' }}>{inr2(c2cSummary.labourSubtotal)}</div>

              </div>

              <div style={{ background: '#fff', padding: '10px 14px', borderRadius: '6px', border: '1px solid #e9d5ff' }}>

                <div style={{ fontSize: '11px', color: '#64748b', fontWeight: 600 }}>C2C Powder Coating (E39)</div>

                <div style={{ fontSize: '15px', fontWeight: 700, color: '#581c87', fontFamily: 'monospace' }}>{inr2(c2cSummary.powderTotal)}</div>

              </div>

              <div style={{ background: '#fff', padding: '10px 14px', borderRadius: '6px', border: '1px solid #e9d5ff' }}>

                <div style={{ fontSize: '11px', color: '#64748b', fontWeight: 600 }}>C2C Direct Cost Base</div>

                <div style={{ fontSize: '15px', fontWeight: 700, color: '#581c87', fontFamily: 'monospace' }}>{inr2(c2cSummary.costBase)}</div>

                <div style={{ fontSize: '9.5px', color: '#7c3aed', marginTop: '3px' }}>
                  Ex-Works ({inr(c2cSummary.subtotal)}) + Trans ({inr(c2cSummary.transport || 0)}) + Acc ({inr(c2cSummary.accommodation || 0)}) + Inst ({inr(c2cSummary.installation || 0)}){(c2cSummary.extraAmount || 0) > 0 ? ` + Extra (${inr((c2cSummary.extraAmount || 0) * (c2cSummary.models?.reduce((a, m) => a + m.qty, 0) || 1))})` : ''}
                </div>

              </div>

            </div>



            {/* Models Table for C2C */}

            <div style={{ marginTop: '12px', overflowX: 'auto' }}>

              <table className="tbl" style={{ width: '100%', fontSize: '12px', background: '#fff', borderRadius: '6px', border: '1px solid #e9d5ff' }}>

                <thead>

                  <tr style={{ background: '#f5f3ff', color: '#581c87', textAlign: 'left' }}>

                    <th style={{ padding: '8px 10px' }}>Model</th>

                    <th style={{ padding: '8px 10px', textAlign: 'center' }}>Qty</th>

                    <th style={{ padding: '8px 10px', textAlign: 'center' }}>Actual Unit Rate</th>

                    <th style={{ padding: '8px 10px', textAlign: 'center', color: '#7c3aed' }}>C2C Unit Rate</th>

                    <th style={{ padding: '8px 10px', textAlign: 'center' }}>Actual Total</th>

                    <th style={{ padding: '8px 10px', textAlign: 'center', color: '#7c3aed' }}>C2C Total</th>

                    <th style={{ padding: '8px 10px', textAlign: 'center', color: '#15803d' }}>Unit Diff (Profit)</th>

                  </tr>

                </thead>

                <tbody>

                  {(summary?.models || []).filter(m => m.qty > 0).map(m => {
                    const c2cM = (c2cSummary?.models || []).find(x => x.code === m.code)

                    const c2cUnitRate = c2cM?.unitRate || 0

                    const c2cAmount = c2cM?.amount || 0

                    const unitDiff = m.unitRate - c2cUnitRate

                    return (

                      <tr key={m.code} style={{ borderTop: '1px solid #f3e8ff' }}>

                        <td style={{ padding: '8px 10px', fontWeight: 700 }}>{m.code}</td>

                        <td style={{ padding: '8px 10px', textAlign: 'center' }}>{m.qty}</td>

                        <td style={{ padding: '8px 10px', textAlign: 'center', fontFamily: 'monospace' }}>{inr2(m.unitRate)}</td>

                        <td style={{ padding: '8px 10px', textAlign: 'center', fontFamily: 'monospace', color: '#7c3aed', fontWeight: 600 }}>{inr2(c2cUnitRate)}</td>

                        <td style={{ padding: '8px 10px', textAlign: 'center', fontFamily: 'monospace' }}>{inr2(m.amount)}</td>

                        <td style={{ padding: '8px 10px', textAlign: 'center', fontFamily: 'monospace', color: '#7c3aed', fontWeight: 600 }}>{inr2(c2cAmount)}</td>

                        <td style={{ padding: '8px 10px', textAlign: 'center', fontFamily: 'monospace', color: '#15803d', fontWeight: 700 }}>+{inr2(unitDiff * m.qty)}</td>

                      </tr>

                    )

                  })}

                </tbody>

              </table>

            </div>

          </div>

        )}



{/* Quotation Security Audit Strip */}

      <div className="quotation-security-watermark-strip">

        <div className="row gap" style={{ alignItems: 'center' }}>

          <span>🔒 <strong>Enterprise Audit Watermark:</strong> {org.name}</span>

          <span>Staff: <strong>@{session.username}</strong></span>

          <span>Quotation: <strong>{revInfo.currentRev ? `Rev ${revInfo.currentRev}` : (name || project.name)}</strong></span>

          <span>Status: <strong className="font-mono">{status.toUpperCase()}</strong></span>

          <span>Date: <strong>{new Date().toLocaleDateString()}</strong></span>

        </div>

        <button

          className="btn ghost sm"

          onClick={handleOpenOfficialQuotation}

          title="View & Print Official Quotation (Exact Myriad Storage 2-Page Document)"

          style={{ padding: '3px 12px', fontSize: '12px', display: 'inline-flex', alignItems: 'center', gap: '6px' }}

        >

          {Icon.fileText} Official Quotation (Print / PDF)

        </button>

      </div>



      {/* Recalculate with chosen template version modal */}

      {showRecalcModal && (

        <div className="modal-backdrop" onClick={() => setShowRecalcModal(false)}>

          <div className="modal" onClick={(e) => e.stopPropagation()}>

            <div className="modal-header">

              <h3>Switch Template Version</h3>

              <button className="icon-btn" onClick={() => setShowRecalcModal(false)}>

                {Icon.close}

              </button>

            </div>

            <div className="modal-body">

              <p className="tiny text-muted">

                Current project template: <strong>Version {project.template_version}</strong>.

                Select any published template version to recalculate this project:

              </p>



              <div className="field">

                <label>Target Template Version</label>

                <select

                  value={targetTemplateId}

                  onChange={(e) => setTargetTemplateId(e.target.value)}

                >

                  {recalcTemplates.map((t) => (

                    <option key={t.id} value={t.id}>

                      v{t.version_num} {t.is_default ? '(Active Default)' : ''} — {t.description || 'Standard template'}

                    </option>

                  ))}

                </select>

              </div>



              <div className="callout info tiny">

                Existing inputs and custom formula overrides are preserved. Only template baseline values are updated.

              </div>

            </div>

            <div className="modal-actions">

              <button className="btn ghost" onClick={() => setShowRecalcModal(false)}>

                Cancel

              </button>

              <button

                className="btn primary"

                disabled={saving}

                onClick={handleRecalculateWithTarget}

              >

                {saving ? <Spinner size={14} /> : Icon.refresh} Apply &amp; Recalculate

              </button>

            </div>

          </div>

        </div>

      )}



      {/* Revise Quotation Modal */}

      {showReviseModal && (

        <div className="modal-backdrop" onClick={() => setShowReviseModal(false)}>

          <div className="modal" onClick={(e) => e.stopPropagation()}>

            <div className="modal-header">

              <h3>Create Revision {revInfo.nextRev}</h3>

              <button className="icon-btn" onClick={() => setShowReviseModal(false)}>

                {Icon.close}

              </button>

            </div>

            <div className="modal-body">

              <p className="tiny muted" style={{ marginBottom: 14, lineHeight: 1.5 }}>

                Create new revision <strong>Rev {revInfo.nextRev}</strong> from this quotation.

                Current revision <strong>Rev {revInfo.currentRev}</strong> will be preserved untouched.

              </p>

              <div className="field">

                <label>New Revision Name</label>

                <input

                  type="text"

                  value={reviseName}

                  onChange={(e) => setReviseName(e.target.value)}

                  autoFocus

                />

              </div>

            </div>

            <div className="modal-actions">

              <button className="btn ghost" onClick={() => setShowReviseModal(false)}>

                Cancel

              </button>

              <button

                className="btn primary"

                disabled={!reviseName.trim()}

                onClick={() => {

                  if (!reviseName.trim()) return
                  const revTarget = reviseName.trim()
                  setShowReviseModal(false)
                  const liveCleanedOverrides = cleanOverrides(wb, overrides)
                  const liveSystems = getEffectiveSystems(overrides, formulaOverrides)
                  const liveSummary = {
                    ...project.summary,
                    ...summary,
                    systems: liveSystems,
                    quotation_sets: compiledSets,
                    quotation_mode: quotationMode,
                    extra_amount: extraAmount,
                    include_installation: includeInstallation,
                    profit_pct: profitPct,
                    discount_pct: discountPct,
                  }
                  const livePayload = {
                    inputs: liveCleanedOverrides,
                    formula_overrides: formulaOverrides,
                    total: summary.sales,
                    summary: liveSummary,
                  }
                  if (dirty) {
                    save().then(() => onSaveAsCopy(revTarget, livePayload)).catch(() => onSaveAsCopy(revTarget, livePayload))
                  } else {
                    onSaveAsCopy(revTarget, livePayload)
                  }

                }}

              >

                {Icon.refresh} Create Rev {revInfo.nextRev}

              </button>

            </div>

          </div>

        </div>

      )}



      {/* Save as copy modal */}

      {showCopyModal && (

        <div className="modal-backdrop" onClick={() => setShowCopyModal(false)}>

          <div className="modal" onClick={(e) => e.stopPropagation()}>

            <div className="modal-header">

              <h3>Save as Copy</h3>

              <button className="icon-btn" onClick={() => setShowCopyModal(false)}>

                {Icon.close}

              </button>

            </div>

            <div className="modal-body">

              <div className="field">

                <label>New Project Name</label>

                <input

                  type="text"

                  value={copyName}

                  onChange={(e) => setCopyName(e.target.value)}

                  autoFocus

                />

              </div>

            </div>

            <div className="modal-actions">

              <button className="btn ghost" onClick={() => setShowCopyModal(false)}>

                Cancel

              </button>

              <button

                className="btn primary"

                onClick={() => {

                  if (!copyName.trim()) return
                  const copyTarget = copyName.trim()
                  setShowCopyModal(false)
                  const liveCleanedOverrides = cleanOverrides(wb, overrides)
                  const liveSystems = getEffectiveSystems(overrides, formulaOverrides)
                  const liveSummary = {
                    ...project.summary,
                    ...summary,
                    systems: liveSystems,
                    quotation_sets: compiledSets,
                    quotation_mode: quotationMode,
                    extra_amount: extraAmount,
                    include_installation: includeInstallation,
                    profit_pct: profitPct,
                    discount_pct: discountPct,
                  }
                  const livePayload = {
                    inputs: liveCleanedOverrides,
                    formula_overrides: formulaOverrides,
                    total: summary.sales,
                    summary: liveSummary,
                  }
                  if (dirty) {
                    save().then(() => onSaveAsCopy(copyTarget, livePayload)).catch(() => onSaveAsCopy(copyTarget, livePayload))
                  } else {
                    onSaveAsCopy(copyTarget, livePayload)
                  }

                }}

              >

                {Icon.copy} Create Copy

              </button>

            </div>

          </div>

        </div>

      )}



      {/* Quotation Reminder Modal */}
      {showReminderModal && (
        <div className="overlay" onMouseDown={() => setShowReminderModal(false)}>
          <div className="dialog" role="dialog" onMouseDown={(e) => e.stopPropagation()} style={{ maxWidth: 440 }}>
            <div className="dialog-header">
              <div className="avatar md" style={{ background: '#fef3c7', color: '#b45309', fontWeight: 800 }}>🔔</div>
              <div>
                <h3 style={{ margin: 0, fontSize: 16 }}>Follow-up Alarm &amp; Reminder</h3>
                <p className="tiny muted" style={{ margin: 0 }}>Set an exact date &amp; time alarm to follow up on this quotation</p>
              </div>
            </div>
            <div className="dialog-body" style={{ padding: '16px 20px', display: 'flex', flexDirection: 'column', gap: 14 }}>
              <div>
                <label className="tiny font-bold text-muted" style={{ display: 'block', marginBottom: 6 }}>
                  ⏰ QUICK ALARM PRESETS
                </label>
                <div style={{ display: 'flex', gap: 6, flexWrap: 'wrap' }}>
                  {[
                    { label: 'Tomorrow 10:00 AM', days: 1, hour: 10, min: 0 },
                    { label: 'In 2 Days 10:00 AM', days: 2, hour: 10, min: 0 },
                    { label: 'Next Week', days: 7, hour: 10, min: 0 },
                    { label: 'In 1 Hour', hoursOffset: 1 },
                    { label: 'In 3 Hours', hoursOffset: 3 },
                  ].map((p) => (
                    <button
                      key={p.label}
                      type="button"
                      className="btn ghost sm"
                      style={{ fontSize: 11, padding: '3px 8px', fontWeight: 600 }}
                      onClick={() => {
                        const d = new Date()
                        if (p.hoursOffset) {
                          d.setHours(d.getHours() + p.hoursOffset)
                        } else {
                          d.setDate(d.getDate() + (p.days || 1))
                          d.setHours(p.hour ?? 10, p.min ?? 0, 0, 0)
                        }
                        const pad = (n: number) => String(n).padStart(2, '0')
                        setReminderDateTimeInput(`${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}T${pad(d.getHours())}:${pad(d.getMinutes())}`)
                      }}
                    >
                      ⏰ {p.label}
                    </button>
                  ))}
                </div>
              </div>

              <div>
                <label className="tiny font-bold text-muted" style={{ display: 'block', marginBottom: 4 }}>
                  EXACT DATE &amp; TIME (ALARM)
                </label>
                <input
                  type="datetime-local"
                  value={reminderDateTimeInput}
                  onChange={(e) => setReminderDateTimeInput(e.target.value)}
                  style={{ width: '100%', padding: '7px 9px', fontSize: 13, border: '1px solid #cbd5e1', borderRadius: 6, fontWeight: 700 }}
                />
              </div>

              <div>
                <label className="tiny font-bold text-muted" style={{ display: 'block', marginBottom: 4 }}>
                  NOTE / ACTION ITEM
                </label>
                <input
                  type="text"
                  value={reminderNoteInput}
                  onChange={(e) => setReminderNoteInput(e.target.value)}
                  placeholder="e.g. Call client for PO confirmation"
                  style={{ width: '100%', padding: '7px 9px', fontSize: 13, border: '1px solid #cbd5e1', borderRadius: 6 }}
                />
              </div>

              {reminder && !reminder.dismissed && (
                <div style={{ padding: '8px 12px', background: '#ecfdf5', borderRadius: 6, border: '1px solid #a7f3d0', fontSize: 12, color: '#065f46' }}>
                  ⏰ Active Alarm: <strong>{new Date(reminder.date).toLocaleDateString('en-IN', { month: 'short', day: 'numeric', year: 'numeric' })} at {new Date(reminder.date).toLocaleTimeString('en-IN', { hour: '2-digit', minute: '2-digit' })}</strong> — "{reminder.note}"
                </div>
              )}
            </div>
            <div className="dialog-footer" style={{ display: 'flex', justifyContent: 'space-between', padding: '12px 20px', background: '#f8fafc', borderTop: '1px solid #e2e8f0' }}>
              <div>
                {reminder && (
                  <button
                    type="button"
                    className="btn ghost sm text-danger"
                    onClick={handleClearReminder}
                  >
                    Clear Alarm
                  </button>
                )}
              </div>
              <div style={{ display: 'flex', gap: 8 }}>
                <button
                  type="button"
                  className="btn ghost sm"
                  onClick={() => setShowReminderModal(false)}
                >
                  Cancel
                </button>
                <button
                  type="button"
                  className="btn primary sm"
                  onClick={() => handleSaveReminder(reminderDateTimeInput, reminderNoteInput)}
                >
                  Save Alarm Reminder
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* Unsaved Leave Guard Modal */}
      {showUnsavedLeaveModal && (
        <div className="overlay" onMouseDown={() => setShowUnsavedLeaveModal(false)}>
          <div className="dialog" role="dialog" onMouseDown={(e) => e.stopPropagation()} style={{ maxWidth: 440 }}>
            <div className="dialog-header">
              <div className="avatar md" style={{ background: '#fef3c7', color: '#b45309', fontWeight: 800 }}>⚠️</div>
              <div>
                <h3 style={{ margin: 0, fontSize: 16 }}>Unsaved Changes</h3>
                <p className="tiny muted" style={{ margin: 0 }}>Please save your changes before leaving</p>
              </div>
            </div>
            <div className="dialog-body" style={{ padding: '16px 20px', color: '#475569', fontSize: 13, lineHeight: 1.5 }}>
              You have unsaved changes in this calculation. If you leave without saving, any modified inputs or matrices will be lost.
            </div>
            <div className="dialog-footer" style={{ display: 'flex', justifyContent: 'flex-end', gap: 8, padding: '12px 20px', background: '#f8fafc', borderTop: '1px solid #e2e8f0' }}>
              <button
                type="button"
                className="btn ghost sm"
                onClick={() => {
                  setShowUnsavedLeaveModal(false)
                  setPendingLeaveAction(null)
                }}
              >
                Cancel & Stay
              </button>
              <button
                type="button"
                className="btn sm"
                style={{ background: '#fee2e2', color: '#b91c1c', border: '1px solid #fecaca' }}
                onClick={() => {
                  setShowUnsavedLeaveModal(false)
                  const action = pendingLeaveAction
                  setPendingLeaveAction(null)
                  if (action) action()
                }}
              >
                Leave Without Saving
              </button>
              <button
                type="button"
                className="btn primary sm"
                disabled={saving}
                onClick={async () => {
                  try {
                    await save()
                    setShowUnsavedLeaveModal(false)
                    const action = pendingLeaveAction
                    setPendingLeaveAction(null)
                    if (action) action()
                  } catch (e: any) {
                    toast(e?.message || 'Failed to save', 'err')
                  }
                }}
              >
                {saving ? <Spinner size={14} /> : Icon.save} Save & Leave
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Share Quotation Modal */}

      {showShareModal && (

        <div className="overlay" onMouseDown={() => setShowShareModal(false)}>

          <div className="dialog" role="dialog" onMouseDown={(e) => e.stopPropagation()} style={{ maxWidth: 440 }}>

            <div className="dialog-header">

              {Icon.users}

              <div>

                <h3 style={{ margin: 0 }}>Share Quotation</h3>

                <span className="tiny text-muted">“{project.name}”</span>

              </div>

            </div>



            <div className="share-member-list mt-3">

              {loadingMembers ? (

                <div style={{ textAlign: 'center', padding: '16px 0' }}>

                  <Spinner size={20} />

                </div>

              ) : (shareMembers || []).filter((m) => m.username !== session.username && m.role !== 'owner' && m.role !== 'admin').length === 0 ? (

                <div className="empty" style={{ padding: '14px 0' }}>

                  <p className="tiny muted">No other staff members found.</p>

                </div>

              ) : (

                shareMembers

                  .filter((m) => m.username !== session.username && m.role !== 'owner' && m.role !== 'admin')

                  .map((m) => {

                    const currentShared = Array.isArray(project.summary?.shared_with)

                      ? project.summary.shared_with

                      : []

                    const isShared = currentShared.includes(m.username)

                    return (

                      <div

                        key={m.username}

                        className="row between p-2"

                        style={{ borderBottom: '1px solid var(--line)' }}

                      >

                        <div className="row gap">

                          <div className="avatar sm" style={{ background: '#1e3a8a', color: '#fff', fontSize: 11 }}>

                            {m.username[0].toUpperCase()}

                          </div>

                          <div>

                            <strong>@{m.username}</strong>

                            <div className="tiny muted">Staff</div>

                          </div>

                        </div>

                        <button

                          className={`btn sm ${isShared ? 'danger ghost' : 'primary'}`}

                          onClick={() => handleToggleShare(m.username)}

                        >

                          {isShared ? 'Revoke Access' : 'Share Access'}

                        </button>

                      </div>

                    )

                  })

              )}

            </div>



            <div className="dialog-actions mt-4">

              <button className="btn primary" onClick={() => setShowShareModal(false)}>

                Done

              </button>

            </div>

          </div>

        </div>

      )}



      {/* Official 2-Page Quotation Document (Exact Myriad Storage System LLP Format) */}

      {showOfficialQuotation && (() => {
        wb.setFormulaOverrides(formulaOverrides)
        wb.setOverrides(overrides)
        const liveProject: Project = {
          ...project,
          name,
          status,
          inputs: overrides,
          formula_overrides: formulaOverrides,
          summary: {
            ...project.summary,
            ...summary,
            quotation_mode: quotationMode,
            systems,
            quotation_sets: compiledSets,
            include_installation: includeInstallation,
            extra_amount: extraAmount,
            profit_pct: profitPct,
            discount_pct: discountPct,
          },
        }

        return (
          <OfficialQuotationDocument
            project={liveProject}
            summary={summary}
            wb={wb}
            role={org.role}
            isAdminOrOwner={isAdminOrOwner}
            orgId={org.id}
            username={session.username}
            userId={session.user_id}
            initialQuotationMode={quotationMode}
            initialExtraAmount={extraAmount}
            initialSets={compiledSets}
            totalCombinedBodies={totalCombinedBodies}
            isUnsavedWarning={dirty}
            onClose={() => setShowOfficialQuotation(false)}
            onSaveClientInfo={handleSaveClientQuotationData}
          />
        )
      })()}

    </div>

  )

}

