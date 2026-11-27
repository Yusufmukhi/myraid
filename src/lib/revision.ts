import { Project } from './api'

/**
 * Revision utilities for project quotations:
 * Handles versioning workflows like Rev A -> Rev B -> Rev C...
 * Ensures initial quotations are created as Rev A, and revisions create
 * duplicate snapshots with the next sequential revision code.
 */

export interface RevisionInfo {
  baseName: string
  currentRev: string
  nextRev: string
  nextName: string
  hasRevision: boolean
}

export interface ProjectFolder {
  id: string
  folderName: string
  revisions: Project[]
  latestRevision: Project
  latestRevChar: string
  nextRevChar: string
  nextName: string
  totalCount: number
  hasActive: boolean
}

/**
 * Calculates Indian Financial Year string (e.g. "26-27") based on April 1 - March 31.
 * Automatically updates when the year rolls over.
 */
export function getFiscalYear(dateInput?: Date | string): string {
  let d = new Date()
  if (dateInput) {
    if (typeof dateInput === 'string') {
      const trimmed = dateInput.trim()
      // Handle ISO timestamps like "2026-02-15T10:30:00Z" cleanly using local/UTC date portion
      const isoPrefixMatch = trimmed.match(/^(\d{4})-(\d{2})-(\d{2})/)
      if (isoPrefixMatch) {
        d = new Date(
          parseInt(isoPrefixMatch[1], 10),
          parseInt(isoPrefixMatch[2], 10) - 1,
          parseInt(isoPrefixMatch[3], 10),
        )
      } else {
        const parts = trimmed.split(/[-/]/)
        if (parts.length === 3) {
          if (parts[0].length === 4) {
            // YYYY-MM-DD
            d = new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10))
          } else if (parts[2].length === 4) {
            // DD-MM-YYYY
            d = new Date(parseInt(parts[2], 10), parseInt(parts[1], 10) - 1, parseInt(parts[0], 10))
          }
        } else {
          const parsed = new Date(trimmed)
          if (!isNaN(parsed.getTime())) d = parsed
        }
      }
    } else if (!isNaN(dateInput.getTime())) {
      d = dateInput
    }
  }
  const month = d.getMonth() // 0 = Jan, 3 = April
  const year = d.getFullYear()
  const startYr = month >= 3 ? year : year - 1
  const endYr = startYr + 1
  return `${String(startYr).slice(-2)}-${String(endYr).slice(-2)}`
}

/**
 * Formats a date as DD-MM-YYYY for Official PDF Quotation Date.
 */
export function formatPdfDate(dateInput: Date = new Date()): string {
  const pad = (n: number) => String(n).padStart(2, '0')
  return `${pad(dateInput.getDate())}-${pad(dateInput.getMonth() + 1)}-${dateInput.getFullYear()}`
}

/**
 * Extracts explicit FY segment (e.g. "26-27") from an MSSLP QTN string if present.
 */
export function extractFiscalYearFromQtnNo(qtnNo?: string): string | null {
  if (!qtnNo) return null
  const m = qtnNo.trim().match(/^MSSLP\/(\d{2}-\d{2})(?:\/|$)/i)
  return m ? m[1] : null
}

/**
 * Replaces the FY segment inside an MSSLP/{fy}/... quotation number cleanly.
 */
export function syncQtnNoFiscalYear(qtnNo: string, fy: string): string {
  const trimmed = (qtnNo || '').trim()
  if (!trimmed || !/^\d{2}-\d{2}$/.test(fy)) return qtnNo
  const m = trimmed.match(/^(MSSLP\/)\d{2}-\d{2}(\/.*)$/i)
  if (m) {
    return `MSSLP/${fy}${m[2]}`
  }
  return qtnNo
}

/**
 * Resolves the authoritative Financial Year ("25-26", "26-27", etc.) for a quotation project.
 * Prioritizes project creation date (`created_at`) so quotations belong to the FY in which they were issued/revised.
 */
export function getProjectFiscalYear(project: { created_at?: string; summary?: any }): string {
  const explicitSummaryFy = project?.summary?.fy
  if (typeof explicitSummaryFy === 'string' && /^\d{2}-\d{2}$/.test(explicitSummaryFy.trim())) {
    return explicitSummaryFy.trim()
  }
  if (project?.created_at) {
    const parsed = new Date(project.created_at)
    if (!isNaN(parsed.getTime())) {
      return getFiscalYear(project.created_at)
    }
  }
  const qtnFy = extractFiscalYearFromQtnNo(project?.summary?.client_info?.qtn_no)
  if (qtnFy) return qtnFy
  if (project?.summary?.client_info?.date) {
    return getFiscalYear(project.summary.client_info.date)
  }
  return getFiscalYear()
}

/**
 * Returns sorted unique Financial Years present across projects (newest first), always including currentFy.
 */
export function getAvailableFiscalYears(
  projects: Array<{ created_at?: string; summary?: any }>,
  currentFy: string = getFiscalYear(),
): string[] {
  const set = new Set<string>([currentFy])
  for (const p of projects || []) {
    const fy = getProjectFiscalYear(p)
    if (fy) set.add(fy)
  }
  return Array.from(set).sort((a, b) => b.localeCompare(a))
}

/**
 * Checks whether a project matches the active FY multi-select filter (where ['ALL'] allows every FY).
 */
export function matchesFiscalYearFilter(
  project: { created_at?: string; summary?: any },
  selectedFys: string[],
): boolean {
  if (!selectedFys || selectedFys.length === 0 || selectedFys.includes('ALL')) {
    return true
  }
  return selectedFys.includes(getProjectFiscalYear(project))
}

/**
 * Extracts the clean quotation number/code from project name or official MSSLP reference.
 * e.g. "QTY NO 123" -> "123"
 * e.g. "QTY NO 123-A" -> "123-A"
 * e.g. "QTY NO 123 - Rev A" -> "123-A"
 * e.g. "MSSLP/25-26/MT-SL/123-A" -> "123-A"
 */
export function extractQuotationCode(projectName: string): string {
  if (!projectName) return ''
  let cleaned = projectName.trim()

  // Strip MSSLP/{fy}/MT-SL/ or MSSLP/{fy}/ prefix if provided
  cleaned = cleaned.replace(/^MSSLP\/\d{2}-\d{2}\/(?:MT-SL\/)?/i, '').trim()

  // Remove leading "QTY NO", "QTN NO", "QTY NO:", "QTN NO:", "QTY:", "QTN:", "QTY", "QTN"
  cleaned = cleaned.replace(/^(?:QTY|QTN)\s*(?:NO|NUM|NUMBER)?\s*[:\-_/.]*\s*/i, '').trim()

  // Remove leading "MT-SL/" if included
  cleaned = cleaned.replace(/^MT-SL\s*[/:\-_]\s*/i, '').trim()

  // Convert explicit "- Rev A" or "Rev A" or "(Rev A)" to "-A"
  cleaned = cleaned.replace(/[\s\-_(]+(?:rev|revision|\.)[\s\-_]*([A-Za-z0-9.]+)[)\s\-_]*$/i, '-$1')

  // Normalize hyphen spacing before revision letter e.g. "123 - A" -> "123-A"
  cleaned = cleaned.replace(/\s*-\s*([A-Za-z0-9]+)$/, '-$1')

  return cleaned.trim()
}

/**
 * Generates the official Quotation Number matching:
 * MSSLP/{fiscalYear}/MT-SL/{code}
 * e.g. MSSLP/26-27/MT-SL/123
 * e.g. MSSLP/26-27/MT-SL/123-A
 */
export function buildOfficialQtnNo(projectName: string, dateInput?: Date | string): string {
  const trimmed = (projectName || '').trim()
  const fy = getFiscalYear(dateInput)
  const code = extractQuotationCode(trimmed) || '101'
  return `MSSLP/${fy}/MT-SL/${code}`
}

/**
 * Parses project name to extract revision letter and calculate next revision name.
 * Format: QTY NO 123 -> QTY NO 123-A -> QTY NO 123-B ...
 */
export function parseRevision(name: string): RevisionInfo {
  const trimmed = (name || '').trim()

  // 1. Match explicit revision markers like "- Rev A", "Rev. B", "(Rev C)", etc.
  const explicitMatch = trimmed.match(/(?:[\s\-_(]+(?:rev|revision|\.))[\s\-_]*([A-Za-z0-9.]+)[)\s\-_]*$/i)
  if (explicitMatch && explicitMatch.index !== undefined) {
    const rawRev = explicitMatch[1].toUpperCase()
    const baseName = trimmed.slice(0, explicitMatch.index).trim()

    if (/^[A-Z]$/.test(rawRev)) {
      const nextChar = rawRev === 'Z' ? 'ZA' : String.fromCharCode(rawRev.charCodeAt(0) + 1)
      return {
        baseName: baseName || trimmed,
        currentRev: rawRev,
        nextRev: nextChar,
        nextName: `${baseName || trimmed}-${nextChar}`,
        hasRevision: true,
      }
    }

    if (/^\d+$/.test(rawRev)) {
      const nextNum = parseInt(rawRev, 10) + 1
      return {
        baseName: baseName || trimmed,
        currentRev: rawRev,
        nextRev: String(nextNum),
        nextName: `${baseName || trimmed}-${nextNum}`,
        hasRevision: true,
      }
    }

    return {
      baseName: baseName || trimmed,
      currentRev: rawRev,
      nextRev: `${rawRev}.1`,
      nextName: `${baseName || trimmed}-${rawRev}.1`,
      hasRevision: true,
    }
  }

  // 2. Match trailing hyphen letter like "QTY NO 123-A" or "123 - A" or "QTY NO 123-B"
  const trailingHyphenMatch = trimmed.match(/[\s\-_]+([A-Za-z])$/)
  if (trailingHyphenMatch && trailingHyphenMatch.index !== undefined) {
    const revChar = trailingHyphenMatch[1].toUpperCase()
    const baseName = trimmed.slice(0, trailingHyphenMatch.index).trim()
    const nextChar = revChar === 'Z' ? 'ZA' : String.fromCharCode(revChar.charCodeAt(0) + 1)
    return {
      baseName: baseName || trimmed,
      currentRev: revChar,
      nextRev: nextChar,
      nextName: `${baseName || trimmed}-${nextChar}`,
      hasRevision: true,
    }
  }

  // 3. Default when no revision suffix exists yet (e.g. "QTY NO 123")
  const base = trimmed || 'QTN NO'
  return {
    baseName: base,
    currentRev: '',
    nextRev: 'A',
    nextName: `${base}-A`,
    hasRevision: false,
  }
}

/**
 * Formats a newly created project name without forcing - Rev A.
 */
export function formatInitialProjectName(name: string): string {
  return (name || '').trim() || 'QTN NO'
}

/**
 * Groups a flat array of project quotations into Project Folders based on baseName.
 * Inside each folder, revisions are ordered sequentially (Rev A, Rev B, etc.).
 */
export function groupProjectsIntoFolders(
  projects: Project[],
  currentProjectId?: string | null,
): ProjectFolder[] {
  const map = new Map<string, { folderName: string; revisions: Project[] }>()

  for (const p of projects) {
    const revInfo = parseRevision(p.name)
    const base = revInfo.baseName || p.name
    const key = base.toLowerCase().trim()

    if (!map.has(key)) {
      map.set(key, { folderName: base, revisions: [] })
    }
    map.get(key)!.revisions.push(p)
  }

  const folders: ProjectFolder[] = []
  for (const [key, group] of map.entries()) {
    // Sort revisions: A -> B -> C -> ...
    const sorted = [...group.revisions].sort((a, b) => {
      const revA = parseRevision(a.name).currentRev
      const revB = parseRevision(b.name).currentRev
      if (revA !== revB) return revA.localeCompare(revB)
      return new Date(a.created_at).getTime() - new Date(b.created_at).getTime()
    })

    const latest = sorted[sorted.length - 1]
    const latestInfo = parseRevision(latest.name)
    const hasActive = currentProjectId ? sorted.some((r) => r.id === currentProjectId) : false

    folders.push({
      id: key,
      folderName: group.folderName,
      revisions: sorted,
      latestRevision: latest,
      latestRevChar: latestInfo.currentRev,
      nextRevChar: latestInfo.nextRev,
      nextName: latestInfo.nextName,
      totalCount: sorted.length,
      hasActive,
    })
  }

  return folders
}

export const CELL_DESCRIPTIONS: Record<string, string> = {
  'Total cal!A3': 'DFMU ROW',
  'Total cal!B3': 'DFMU COL',
  'Total cal!C3': 'DFMU Height (mm)',
  'Total cal!D3': 'DFMU Width (mm)',
  'Total cal!E3': 'DFMU Depth (mm)',
  'Total cal!F3': 'DFMU Body Gauge',
  'Total cal!G3': 'DFMU Shelf Gauge',
  'Total cal!J3': 'DFMU Stiffener Gauge',
  'Total cal!I3': 'DFMU Track',
  'Total cal!L3': 'DFMU Trolley Gauge',
  'Total cal!G4': 'DFMU Quantity',
  'Total cal!E7': 'SFMU Depth (mm)',
  'Total cal!G8': 'SFMU Quantity',
  'Total cal!E11': 'SFFU Depth (mm)',
  'Total cal!G12': 'SFFU Quantity',
  'Total cal!G16': 'DFFU Quantity',
  'Total cal!G20': 'DFMU-D Quantity',
  'Total cal!B49': 'Labour per Body (₹)',
  'Total cal!N23': 'Installation (₹)',
  'Total cal!N24': 'Transport / Freight (₹)',
  'Total cal!N25': 'Accommodation (₹)',
  '__discount_pct': 'Discount %',
  '__profit_pct': 'Profit Margin %',
  '__include_installation': 'Installation Included',
}

export const CELL_DEFAULT_VALUES: Record<string, any> = {
  'Total cal!A3': 4,
  'Total cal!B3': 5,
  'Total cal!C3': 2200,
  'Total cal!D3': 1830,
  'Total cal!E3': 610,
  'Total cal!F3': 0.8,
  'Total cal!G3': 0.8,
  'Total cal!J3': 0.8,
  'Total cal!I3': 2,
  'Total cal!L3': 2.5,
  'Total cal!G4': 8,
  'Total cal!E7': 475,
  'Total cal!G8': 1,
  'Total cal!E11': 330,
  'Total cal!G12': 1,
  'Total cal!G16': 0,
  'Total cal!G20': 0,
  'Total cal!B49': 250,
  'Total cal!N23': 0,
  'Total cal!N24': 0,
  'Total cal!N25': 0,
  'Rates!B3': 90,
  'Rates!B4': 45,
  'Rates!B5': 100,
  'Rates!B6': 22,
  '__discount_pct': 0,
  '__profit_pct': 25,
  '__include_installation': 1,
}

export const LABEL_DEFAULT_VALUES: Record<string, any> = {
  'dfmu height (mm)': 2200,
  'dfmu height': 2200,
  'height': 2200,
  'total cal!c3': 2200,
  'dfmu width (mm)': 1830,
  'dfmu width': 1830,
  'width': 1830,
  'total cal!d3': 1830,
  'dfmu depth (mm)': 610,
  'dfmu depth': 610,
  'depth': 610,
  'total cal!e3': 610,
  'dfmu row': 4,
  'row': 4,
  'total cal!a3': 4,
  'dfmu col': 5,
  'col': 5,
  'total cal!b3': 5,
  'dfmu body gauge': 0.8,
  'total cal!f3': 0.8,
  'dfmu shelf gauge': 0.8,
  'total cal!g3': 0.8,
  'dfmu stiffener gauge': 0.8,
  'total cal!j3': 0.8,
  'dfmu track': 2,
  'total cal!i3': 2,
  'dfmu trolley gauge': 2.5,
  'total cal!l3': 2.5,
  'dfmu quantity': 8,
  'quantity': 8,
  'qty': 8,
  'total cal!g4': 8,
  'sfmu depth (mm)': 475,
  'total cal!e7': 475,
  'sfmu quantity': 1,
  'total cal!g8': 1,
  'sffu depth (mm)': 330,
  'total cal!e11': 330,
  'sffu quantity': 1,
  'total cal!g12': 1,
  'dffu quantity': 0,
  'total cal!g16': 0,
  'dfmu-d quantity': 0,
  'total cal!g20': 0,
  'labour per body': '₹250',
  'labour per body (₹)': '₹250',
  'total cal!b49': '₹250',
  'installation': 0,
  'installation (₹)': 0,
  'total cal!n23': 0,
  'transport / freight': 0,
  'transport / freight (₹)': 0,
  'transportation': 0,
  'transport': 0,
  'total cal!n24': 0,
  'accommodation': 0,
  'accommodation (₹)': 0,
  'total cal!n25': 0,
  'steel rate (₹/kg)': '₹90',
  'rates!b3': '₹90',
  'powder coating rate (₹/kg)': '₹45',
  'rates!b4': '₹45',
  'trolley rate (₹/kg)': '₹100',
  'rates!b5': '₹100',
  'hardware rate (₹)': '₹22',
  'rates!b6': '₹22',
  'discount %': '0%',
  '__discount_pct': '0%',
  'profit margin %': '25%',
  '__profit_pct': '25%',
  'installation included': 'Included',
  '__include_installation': 'Included',
}

/** Cleans any historical diff text that mistakenly contains 'Default' or wrong initial baselines */
export function cleanDiffText(text: string): string {
  if (!text) return ''

  // Retroactively correct legacy incorrect baseline numbers in historical logs
  let cleaned = text
    .replace(/(Transport\s*\/\s*Freight(?:\s*\(₹\))?:\s*)(?:5000|₹5,000|Default|deafult)(\s*(?:→|->|to)\s*)/gi, '$10$2')
    .replace(/(Transportation(?:\s*\(₹\))?:\s*)(?:5000|₹5,000|Default|deafult)(\s*(?:→|->|to)\s*)/gi, '$10$2')
    .replace(/(Accommodation(?:\s*\(₹\))?:\s*)(?:3000|₹3,000|Default|deafult)(\s*(?:→|->|to)\s*)/gi, '$10$2')
    .replace(/(DFMU Depth \(mm\):\s*)(?:400|Default|deafult)(\s*(?:→|->|to)\s*)/gi, '$1610$2')
    .replace(/(DFMU Width \(mm\):\s*)(?:900|Default|deafult)(\s*(?:→|->|to)\s*)/gi, '$11830$2')
    .replace(/(DFMU Height \(mm\):\s*)(?:2080|Default|deafult)(\s*(?:→|->|to)\s*)/gi, '$12200$2')
    .replace(/(DFMU ROW:\s*)(?:2|Default|deafult)(\s*(?:→|->|to)\s*)/gi, '$14$2')
    .replace(/(DFMU COL:\s*)(?:1|Default|deafult)(\s*(?:→|->|to)\s*)/gi, '$15$2')
    .replace(/(SFMU Depth \(mm\):\s*)(?:400|Default|deafult)(\s*(?:→|->|to)\s*)/gi, '$1475$2')
    .replace(/(SFFU Depth \(mm\):\s*)(?:400|Default|deafult)(\s*(?:→|->|to)\s*)/gi, '$1330$2')

  // Generic replacement for any remaining "Default -> X" or "deafult to X"
  cleaned = cleaned.replace(/^(.*?)(:\s*)?(?:default|deafult)\s*(?:→|->|to)\s*(.*?)$/i, (match, prefix, colon, newVal) => {
    const cleanPrefix = (prefix || '').trim()
    const cleanVal = (newVal || '').trim()
    const key = cleanPrefix.toLowerCase()
    let startingVal = LABEL_DEFAULT_VALUES[key]

    if (startingVal === undefined) {
      if (cleanVal === '130' || cleanVal === '1000' || cleanVal === '800') {
        startingVal = 1830
      } else if (cleanVal === '10' || cleanVal === '450' || cleanVal === '500') {
        startingVal = 610
      } else if (cleanVal === '2100' || cleanVal === '2000' || cleanVal === '1980') {
        startingVal = 2200
      } else {
        startingVal = '0'
      }
    }

    if (cleanPrefix) {
      return `${cleanPrefix}: ${startingVal} → ${cleanVal}`
    }
    return `${startingVal} → ${cleanVal}`
  })

  return cleaned
}

export interface ValueChange {
  key: string
  label: string
  oldVal: any
  newVal: any
  diff: string
}

export function getInitialCellValue(key: string, wb?: any): any {
  if (key === '__profit_pct') return 25
  if (key === '__discount_pct') return 0
  if (key === '__include_installation') return 1

  if (key in CELL_DEFAULT_VALUES && CELL_DEFAULT_VALUES[key] !== undefined) {
    return CELL_DEFAULT_VALUES[key]
  }

  if (wb && key.includes('!')) {
    const p = key.indexOf('!')
    const sheet = key.substring(0, p)
    const a = key.substring(p + 1)
    const m = /^([A-Za-z]+)(\d+)$/.exec(a)
    if (m) {
      let colNum = 0
      const colStr = m[1].toUpperCase()
      for (let i = 0; i < colStr.length; i++) colNum = colNum * 26 + (colStr.charCodeAt(i) - 64)
      const r = parseInt(m[2], 10) - 1
      const c = colNum - 1
      if (typeof wb.defaultInput === 'function') {
        const def = wb.defaultInput(sheet, r, c)
        if (def !== null && def !== undefined) return def
      }
      if (typeof wb.raw === 'function') {
        const raw = wb.raw(sheet, r, c)
        if (typeof raw === 'number') return raw
      }
    }
  }

  const label = CELL_DESCRIPTIONS[key] || key
  const labelKey = label.toLowerCase().trim()
  if (labelKey in LABEL_DEFAULT_VALUES && LABEL_DEFAULT_VALUES[labelKey] !== undefined) {
    return LABEL_DEFAULT_VALUES[labelKey]
  }

  return 0
}

export function detectValueChanges(
  oldInputs: Record<string, any> = {},
  newInputs: Record<string, any> = {},
  oldStatus?: string,
  newStatus?: string,
  oldName?: string,
  newName?: string,
  oldTotal?: number | null,
  newTotal?: number | null,
  wb?: any,
): ValueChange[] {
  const changes: ValueChange[] = []

  // Check name change
  if (oldName && newName && oldName.trim() !== newName.trim()) {
    changes.push({
      key: 'name',
      label: 'Quotation Name',
      oldVal: oldName,
      newVal: newName,
      diff: `Name: “${oldName}” → “${newName}”`,
    })
  }

  // Check status change
  if (oldStatus && newStatus && oldStatus !== newStatus) {
    changes.push({
      key: 'status',
      label: 'Status',
      oldVal: oldStatus,
      newVal: newStatus,
      diff: `Status: ${oldStatus.toUpperCase()} → ${newStatus.toUpperCase()}`,
    })
  }

  // Check inputs differences
  const allKeys = Array.from(new Set([...Object.keys(oldInputs), ...Object.keys(newInputs)]))
  for (const k of allKeys) {
    const def = getInitialCellValue(k, wb)
    const ov = oldInputs[k] !== undefined && oldInputs[k] !== null ? oldInputs[k] : def
    const nv = newInputs[k] !== undefined && newInputs[k] !== null ? newInputs[k] : def
    if (ov === nv) continue

    const label = CELL_DESCRIPTIONS[k] || k
    let formattedOld = ov
    let formattedNew = nv

    if (k === '__include_installation') {
      formattedOld = Boolean(Number(ov)) ? 'Included' : 'Excluded'
      formattedNew = Boolean(Number(nv)) ? 'Included' : 'Excluded'
    } else if (k === '__discount_pct' || k === '__profit_pct') {
      formattedOld = `${ov}%`
      formattedNew = `${nv}%`
    }

    const diff = cleanDiffText(`${label}: ${formattedOld} → ${formattedNew}`)

    changes.push({
      key: k,
      label,
      oldVal: formattedOld,
      newVal: formattedNew,
      diff,
    })
  }

  // Check total value change
  if (oldTotal !== undefined && oldTotal !== null && newTotal !== undefined && newTotal !== null) {
    if (Math.abs(Number(oldTotal) - Number(newTotal)) >= 0.5) {
      changes.push({
        key: 'total',
        label: 'Quotation Total',
        oldVal: Math.round(Number(oldTotal)),
        newVal: Math.round(Number(newTotal)),
        diff: `Quotation Total: ₹${Math.round(Number(oldTotal)).toLocaleString('en-IN')} → ₹${Math.round(Number(newTotal)).toLocaleString('en-IN')}`,
      })
    }
  }

  return changes
}
