/**
 * Centralized Calculation Engine for Bodies, Multi-System Logistics, and Cut-to-Cut Metrics
 *
 * This module is the SINGLE AUTHORITATIVE SOURCE OF TRUTH across:
 * - Web Calculator (Top KPI widgets, commercial breakdown, active system tabs)
 * - Official PDF Quotation (Commercial proposal, annexure installation bodies table, unit rates)
 * - Excel Export & Spreadsheet View
 *
 * Guarantees 100% parity so that total bodies, allocated logistics, and C2C analysis
 * never diverge between Web, PDF, and exported documents.
 */

import { Workbook, Overrides } from './engine'
import { Summary } from './model'

export interface ModelBodyInfo {
  code: string
  label: string
  rowsPerUnit: number
  qty: number
  unitBodies: number // for 1 set = rowsPerUnit * qty
  totalBodies: number // for all sets = unitBodies * setsQty
  desc: string
}

export interface SystemBodiesResult {
  dfmuRows: number
  sfmuRows: number
  sffuRows: number
  dffuRows: number
  dfmudRows: number
  dfmuQty: number
  sfmuQty: number
  sffuQty: number
  dffuQty: number
  dfmudQty: number
  dfmuBodies: number
  sfmuBodies: number
  sffuBodies: number
  dffuBodies: number
  dfmudBodies: number
  singleSetBodies: number
  setsQty: number
  totalBodies: number
  models: ModelBodyInfo[]
  summaryText: string
  tooltipText: string
}

/** Returns true strictly for single-face compactor models (1 body per bay). */
export function isSingleFaceModel(codeOrText?: string): boolean {
  const s = (codeOrText || '').trim().toUpperCase()
  if (!s) return false
  if (s.includes('DFMU-D') || s.includes('DFFU-D') || s.includes('DFMU') || s.includes('DFFU')) return false
  return s.includes('SFMU') || s.includes('SFFU')
}

/** Extracts normalized compactor model code, checking `-D` variants before base codes. */
export function extractBodyModelCode(codeOrText?: string): string {
  const s = (codeOrText || '').trim().toUpperCase()
  if (!s) return ''
  if (s.includes('DFMU-D')) return 'DFMU-D'
  if (s.includes('DFFU-D')) return 'DFFU-D'
  if (s.includes('DFMU')) return 'DFMU'
  if (s.includes('SFMU')) return 'SFMU'
  if (s.includes('SFFU')) return 'SFFU'
  if (s.includes('DFFU')) return 'DFFU'
  return s
}

/**
 * Detects the system's compactor bay count (e.g. 1, 2, 3, 4, 5 BAY) from:
 * 1. Explicit live overrides Total cal!A3 (DFMU rows / 2: 2->1 bay, 4->2, 6->3, 8->4, 10->5)
 * 2. Explicit bay field (sys.bay, sys.bay_count, sys.breakdown?.bay_count, overrides.bayInput)
 * 3. Explicit live overrides width Total cal!D3 (915->1 bay, 1830->2, 2745->3, 3660->4, 4575->5)
 * 4. System title, label, or name (e.g. "2 BAY", "SYSTEM A - 4 BAY", "4-BAY")
 * 5. Breakdown or items dimension width / per-unit rows
 */
export function detectSystemBayCount(sys: any, overrides: Overrides = {}): number | null {
  if (!sys && !overrides) return null

  // Helper regex to extract bay from text string
  const extractBayRegex = (text?: string): number | null => {
    if (!text) return null
    // Matches "2 BAY", "2-BAY", "2BAY", "2 BAYS", "BAY 2", "BAY-2"
    const m = text.match(/(?:^|[^0-9])([1-9]\d?)\s*[-_]?\s*BAYS?(?:$|[^a-zA-Z0-9])/i) ||
              text.match(/\bBAYS?\s*[-_:]?\s*([1-9]\d?)\b/i)
    if (m && m[1]) {
      const val = parseInt(m[1], 10)
      if (val >= 1 && val <= 20) return val
    }
    return null
  }

  // Helper to extract bay from dimension width (standard compactor module width is 915mm per bay)
  const extractBayFromDim = (dim?: string): number | null => {
    if (!dim) return null
    const wMatch = dim.match(/(\d{3,4})\s*(?:mm)?\s*\([Ww]\)/i) ||
                   dim.match(/[Xx]\s*(\d{3,4})\s*(?:mm)?\s*[Xx]/i) ||
                   dim.match(/(\d{3,4})\s*WIDTH/i)
    if (wMatch && wMatch[1]) {
      const w = parseInt(wMatch[1], 10)
      if (w >= 800 && w <= 1100) return 1      // 915 mm
      if (w >= 1600 && w <= 2000) return 2    // 1830 mm
      if (w >= 2500 && w <= 2950) return 3    // 2745 mm
      if (w >= 3400 && w <= 3850) return 4    // 3660 mm
      if (w >= 4300 && w <= 4750) return 5    // 4575 mm
      const approx = Math.round(w / 915)
      if (approx >= 1 && approx <= 20) return approx
    }
    return null
  }

  // 1. Live overrides Total cal!A3 (authoritative row count input — beats stale sys.bay or title)
  const a3 = Number(overrides?.['Total cal!A3'] ?? sys?.overrides?.['Total cal!A3'])
  if (a3 && !isNaN(a3) && a3 > 0) {
    return Math.max(1, Math.round(a3 / 2))
  }

  // 2. Explicit bay field
  const explicitBay = Number(sys?.bay ?? sys?.bay_count ?? sys?.breakdown?.bay_count ?? overrides?.bayInput ?? overrides?.bay ?? overrides?.bay_count ?? sys?.bayInput)
  if (!isNaN(explicitBay) && explicitBay > 0 && explicitBay <= 20) return explicitBay

  // 3. Explicit live overrides width Total cal!D3
  const d3 = Number(overrides?.['Total cal!D3'] ?? sys?.overrides?.['Total cal!D3'])
  if (d3 && !isNaN(d3) && d3 > 0) {
    if (d3 >= 800 && d3 <= 1100) return 1      // 915 mm
    if (d3 >= 1600 && d3 <= 2000) return 2    // 1830 mm
    if (d3 >= 2500 && d3 <= 2950) return 3    // 2745 mm
    if (d3 >= 3400 && d3 <= 3850) return 4    // 3660 mm
    if (d3 >= 4300 && d3 <= 4750) return 5    // 4575 mm
    const approx = Math.round(d3 / 915)
    if (approx >= 1 && approx <= 20) return approx
  }

  // 4. Scan system title, set_label, name, label
  const titleBay = extractBayRegex(sys?.system_title) ||
                   extractBayRegex(sys?.set_label) ||
                   extractBayRegex(sys?.name) ||
                   extractBayRegex(sys?.label)
  if (titleBay) return titleBay

  // 5. Scan items if present
  if (Array.isArray(sys?.items)) {
    for (const it of sys.items) {
      const itBay = extractBayRegex(it.title) ||
                    extractBayRegex(it.model_name) ||
                    extractBayRegex(it.model_code) ||
                    extractBayRegex(it.description)
      if (itBay) return itBay
    }
    for (const it of sys.items) {
      const dimBay = extractBayFromDim(it.dimensions) || extractBayFromDim(it.description)
      if (dimBay) return dimBay
    }
    for (const it of sys.items) {
      const c = (it.model_code || it.model_name || '').toUpperCase()
      if (c.includes('DFMU') && typeof it.rows_per_unit === 'number' && it.rows_per_unit > 0) {
        return Math.max(1, Math.round(it.rows_per_unit / 2))
      }
      if ((c.includes('SFMU') || c.includes('SFFU')) && typeof it.rows_per_unit === 'number' && it.rows_per_unit > 0) {
        return Math.max(1, it.rows_per_unit)
      }
    }
  }

  // 6. Check breakdown rows
  if (Array.isArray(sys?.breakdown?.rows)) {
    for (const r of sys.breakdown.rows) {
      const c = (r.model_code || r.model_name || '').toUpperCase()
      if (c.includes('DFMU') && typeof r.rows_per_unit === 'number' && r.rows_per_unit > 0) {
        return Math.max(1, Math.round(r.rows_per_unit / 2))
      }
      if ((c.includes('SFMU') || c.includes('SFFU')) && typeof r.rows_per_unit === 'number' && r.rows_per_unit > 0) {
        return Math.max(1, r.rows_per_unit)
      }
      const dimBay = extractBayFromDim(r.dimensions)
      if (dimBay) return dimBay
    }
  }

  return null
}

/**
 * Calculates the exact per-model body rows, quantities, bodies count, and breakdown for any system.
 * Authoritative single formula for Total Bodies.
 */
export function calculateSystemBodies(
  sys: any,
  sysOverrides: Overrides = {},
  wbInstance?: Workbook | null,
  isCurrent: boolean = false,
  currentBodies?: number,
): SystemBodiesResult {
  const overrides = sysOverrides || sys?.overrides || {}
  const setsQty = Math.max(1, Number(sys?.quantity) || 1)

  // 1. Detect bay count
  const detectedBay = detectSystemBayCount(sys, overrides)

  // Determine dfmuRows (Master Bay / Row count from Total cal!A3, or saved breakdown, or detected bay)
  let dfmuRows: number
  if (detectedBay && detectedBay > 0) {
    // If the system explicitly defines a bay count (e.g. 2 BAY or width 1830 mm),
    // DFMU rows is strictly 2 rows per bay! (2 BAY -> 4 rows, 3 BAY -> 6 rows, etc.)
    dfmuRows = detectedBay * 2
  } else {
    const breakdownRows = sys?.breakdown?.rows
    const breakdownA3 = breakdownRows?.find((r: any) => r.model_code === 'DFMU')?.rows_per_unit || breakdownRows?.[0]?.rows_per_unit
    const rawA3 = overrides['Total cal!A3'] ?? (isCurrent && wbInstance ? wbInstance.num('Total cal', 'A3') : undefined)
    dfmuRows = Number(rawA3 !== undefined ? rawA3 : (breakdownA3 || (wbInstance ? (wbInstance.num('Total cal', 'A3') || 6) : 6))) || 6
  }

  // Derived row counts (SFMU & SFFU are 1/2 of DFMU; DFFU & DFMU-D match DFMU)
  const getRowsForModel = (code: string, defaultRows: number) => {
    if (detectedBay && detectedBay > 0) {
      return (code === 'SFMU' || code === 'SFFU') ? detectedBay : detectedBay * 2
    }
    const breakdownRows = sys?.breakdown?.rows
    const fromBreakdown = breakdownRows?.find((r: any) => r.model_code === code)?.rows_per_unit
    return fromBreakdown || defaultRows
  }

  const defaultSingleRows = detectedBay ? detectedBay : Math.max(1, Math.round(dfmuRows / 2))
  const sfmuRows = Number(overrides['Total cal!A7'] ?? getRowsForModel('SFMU', defaultSingleRows))
  const sffuRows = Number(overrides['Total cal!A11'] ?? getRowsForModel('SFFU', defaultSingleRows))
  const dffuRows = Number(overrides['Total cal!A15'] ?? getRowsForModel('DFFU', dfmuRows))
  const dfmudRows = Number(overrides['Total cal!A19'] ?? getRowsForModel('DFMU-D', dfmuRows))

  // 2. Model unit quantities from overrides, active workbook, summary models, or sys.items
  const getQ = (cell: string, code: string) => {
    if (overrides[cell] !== undefined && overrides[cell] !== null) return Number(overrides[cell]) || 0
    if (isCurrent && wbInstance) {
      const v = wbInstance.num('Total cal', cell.split('!')[1])
      if (v !== undefined && v !== null && v > 0) return v
    }
    const found = sys?.summary?.models?.find((m: any) => m.code === code) || sys?.models?.find((m: any) => m.code === code)
    if (found && found.qty !== undefined) return Number(found.qty) || 0
    if (Array.isArray(sys?.items)) {
      const item = sys.items.find((it: any) => {
        const c = (it.model_code || it.model_name || '').trim().toUpperCase()
        return c === code || c.startsWith(`${code} `) || c.includes(`(${code})`)
      })
      if (item) {
        const q = item.unit_qty !== undefined ? item.unit_qty : item.quantity
        if (q !== undefined && q !== null) return Number(q) || 0
      }
    }
    return 0
  }

  const dfmuQty = getQ('Total cal!G4', 'DFMU')
  const sfmuQty = getQ('Total cal!G8', 'SFMU')
  const sffuQty = getQ('Total cal!G12', 'SFFU')
  const dffuQty = getQ('Total cal!G16', 'DFFU')
  const dfmudQty = getQ('Total cal!G20', 'DFMU-D')

  const dfmuBodies = Math.round(dfmuQty * dfmuRows)
  const sfmuBodies = Math.round(sfmuQty * sfmuRows)
  const sffuBodies = Math.round(sffuQty * sffuRows)
  const dffuBodies = Math.round(dffuQty * dffuRows)
  const dfmudBodies = Math.round(dfmudQty * dfmudRows)

  const computedBodies = dfmuBodies + sfmuBodies + sffuBodies + dffuBodies + dfmudBodies

  let singleSetBodies = computedBodies
  if (singleSetBodies <= 0) {
    if (isCurrent && currentBodies !== undefined && currentBodies > 0) {
      singleSetBodies = currentBodies
    } else if (sys?.summary?.bodies && sys.summary.bodies > 0) {
      singleSetBodies = sys.summary.bodies
    } else if (wbInstance) {
      const wbBodies = wbInstance.num('Total cal', 'B48')
      if (wbBodies && wbBodies > 0) singleSetBodies = wbBodies
    }
  }
  if (singleSetBodies <= 0) singleSetBodies = 1

  const totalBodies = singleSetBodies * setsQty

  const models: ModelBodyInfo[] = []
  if (dfmuQty > 0 || dfmuBodies > 0) {
    models.push({
      code: 'DFMU',
      label: 'DFMU',
      rowsPerUnit: dfmuRows,
      qty: dfmuQty,
      unitBodies: dfmuBodies,
      totalBodies: dfmuBodies * setsQty,
      desc: `${dfmuRows} rows × ${dfmuQty} qty = ${dfmuBodies}`,
    })
  }
  if (sfmuQty > 0 || sfmuBodies > 0) {
    models.push({
      code: 'SFMU',
      label: 'SFMU',
      rowsPerUnit: sfmuRows,
      qty: sfmuQty,
      unitBodies: sfmuBodies,
      totalBodies: sfmuBodies * setsQty,
      desc: `${sfmuRows} rows × ${sfmuQty} qty = ${sfmuBodies}`,
    })
  }
  if (sffuQty > 0 || sffuBodies > 0) {
    models.push({
      code: 'SFFU',
      label: 'SFFU',
      rowsPerUnit: sffuRows,
      qty: sffuQty,
      unitBodies: sffuBodies,
      totalBodies: sffuBodies * setsQty,
      desc: `${sffuRows} rows × ${sffuQty} qty = ${sffuBodies}`,
    })
  }
  if (dffuQty > 0 || dffuBodies > 0) {
    models.push({
      code: 'DFFU',
      label: 'DFFU',
      rowsPerUnit: dffuRows,
      qty: dffuQty,
      unitBodies: dffuBodies,
      totalBodies: dffuBodies * setsQty,
      desc: `${dffuRows} rows × ${dffuQty} qty = ${dffuBodies}`,
    })
  }
  if (dfmudQty > 0 || dfmudBodies > 0) {
    models.push({
      code: 'DFMU-D',
      label: 'DFMU-D',
      rowsPerUnit: dfmudRows,
      qty: dfmudQty,
      unitBodies: dfmudBodies,
      totalBodies: dfmudBodies * setsQty,
      desc: `${dfmudRows} rows × ${dfmudQty} qty = ${dfmudBodies}`,
    })
  }

  const summaryText = models.length > 0
    ? models.map((m) => `${m.code}: ${m.unitBodies}`).join(' · ')
    : '0 Bodies'

  const tooltipText = models.map((m) => `${m.code}: ${m.desc}`).join('\n') +
    `\nTotal: ${singleSetBodies} bodies / set${setsQty > 1 ? ` × ${setsQty} sets = ${totalBodies} bodies` : ''}`

  return {
    dfmuRows,
    sfmuRows,
    sffuRows,
    dffuRows,
    dfmudRows,
    dfmuQty,
    sfmuQty,
    sffuQty,
    dffuQty,
    dfmudQty,
    dfmuBodies,
    sfmuBodies,
    sffuBodies,
    dffuBodies,
    dfmudBodies,
    singleSetBodies,
    setsQty,
    totalBodies,
    models,
    summaryText,
    tooltipText,
  }
}

/**
 * Calculates total bodies across all systems and quantities.
 * Treats the whole project as one job for transportation and accommodation allocation.
 */
export function calculateProjectCombinedBodies(
  systems: any[],
  activeSystemIdx: number,
  currentOverrides: Overrides,
  wbInstance: Workbook,
  currentBodies?: number,
): {
  totalCombinedBodies: number
  systemBreakdowns: SystemBodiesResult[]
} {
  const breakdowns = (systems || []).map((sys, idx) => {
    const isCurrent = idx === activeSystemIdx
    const ov = isCurrent ? currentOverrides : (sys.overrides || {})
    return calculateSystemBodies(
      sys,
      ov,
      wbInstance,
      isCurrent,
      isCurrent ? currentBodies : undefined,
    )
  })

  const totalCombinedBodies = breakdowns.reduce((acc, b) => acc + b.totalBodies, 0) || 1

  return {
    totalCombinedBodies,
    systemBreakdowns: breakdowns,
  }
}

/**
 * Calculates proportional transport and accommodation allocated per body.
 * Strictly avoids arbitrary 50/50 splits when systems have unequal body counts.
 * Returns both per-set logistics and system total logistics.
 */
export function calculateAllocatedLogistics(
  projectTransport: number,
  projectAccommodation: number,
  systemBodies: number,
  totalCombinedBodies: number,
  isMultiSystemOrSets: boolean = true,
  setsQty: number = 1,
): {
  allocatedTransport: number
  allocatedAccommodation: number
  shareRatio: number
  ratePerBodyTransport: number
  ratePerBodyAccommodation: number
  perSetTransport: number
  perSetAccommodation: number
  systemTotalTransport: number
  systemTotalAccommodation: number
} {
  const cleanTotalBodies = totalCombinedBodies > 0 ? totalCombinedBodies : 1
  const ratePerBodyTransport = projectTransport / cleanTotalBodies
  const ratePerBodyAccommodation = projectAccommodation / cleanTotalBodies

  const shareRatio = (totalCombinedBodies > 0 && isMultiSystemOrSets)
    ? (systemBodies / cleanTotalBodies)
    : 1
  const allocatedTransport = (totalCombinedBodies > 0 && isMultiSystemOrSets)
    ? Math.round(projectTransport * shareRatio)
    : projectTransport
  const allocatedAccommodation = (totalCombinedBodies > 0 && isMultiSystemOrSets)
    ? Math.round(projectAccommodation * shareRatio)
    : projectAccommodation

  const cleanSets = Math.max(1, Number(setsQty) || 1)
  const perSetTransport = allocatedTransport
  const perSetAccommodation = allocatedAccommodation
  const systemTotalTransport = Math.round(allocatedTransport * cleanSets)
  const systemTotalAccommodation = Math.round(allocatedAccommodation * cleanSets)

  return {
    allocatedTransport,
    allocatedAccommodation,
    shareRatio,
    ratePerBodyTransport,
    ratePerBodyAccommodation,
    perSetTransport,
    perSetAccommodation,
    systemTotalTransport,
    systemTotalAccommodation,
  }
}

/**
 * Calculates installation charges individually for a system.
 * System Installation = Total Bodies for System * Labour Charge Rate
 * Total Bodies for System = Single Set Bodies * Sets Quantity
 */
export function calculateSystemInstallation(
  singleSetBodies: number,
  setsQty: number = 1,
  labourPerBody: number = 250,
  includeInstallation: boolean = true,
): {
  unitInstallation: number
  totalInstallation: number
  labourPerBody: number
} {
  if (!includeInstallation) {
    return { unitInstallation: 0, totalInstallation: 0, labourPerBody }
  }
  const cleanLabour = labourPerBody > 0 ? labourPerBody : 250
  const cleanSets = Math.max(1, Number(setsQty) || 1)
  const unitInstallation = Math.round(singleSetBodies * cleanLabour)
  const totalInstallation = Math.round(singleSetBodies * cleanSets * cleanLabour)
  return { unitInstallation, totalInstallation, labourPerBody: cleanLabour }
}

/**
 * Centralized Cut-to-Cut 3-Line Analysis Calculation:
 * Line 1: Profit = Actual Before GST - Cut-to-Cut Cost Base
 * Line 2: Total = Profit + C2C Labour
 * Line 3: Ratio = Total / C2C Steel Weight
 */
export interface C2CAnalysisMetrics {
  profit: number
  total: number
  ratio: number
  costBase: number
  labour: number
  steelWeight: number
}

export function calculateC2CAnalysis(
  subtotalBeforeGst: number,
  costBase: number,
  labourSubtotal: number,
  totalWeight: number,
): C2CAnalysisMetrics {
  const profit = subtotalBeforeGst - costBase
  const total = profit + labourSubtotal
  const ratio = totalWeight > 0 ? Number((total / totalWeight).toFixed(2)) : 0

  return {
    profit,
    total,
    ratio,
    costBase,
    labour: labourSubtotal,
    steelWeight: totalWeight,
  }
}

/**
 * Generates official installation body items for the PDF document directly
 * from the centralized system bodies calculation, guaranteeing 100% parity with web.
 */
export function generateCentralizedBodyItems(
  systems: any[],
  activeSystemIdx: number,
  currentOverrides: Overrides,
  wbInstance: Workbook,
  summary?: Summary,
): any[] {
  const { systemBreakdowns } = calculateProjectCombinedBodies(
    systems,
    activeSystemIdx,
    currentOverrides,
    wbInstance,
    summary?.bodies,
  )

  const allItems: any[] = []
  let counter = 1

  systemBreakdowns.forEach((breakdown, sIdx) => {
    const sys = systems[sIdx] || {}
    const nextChar = String.fromCharCode(65 + sIdx)
    const sysName = (sys.name || sys.system_title || (sys.set_label?.startsWith('SYSTEM') ? sys.set_label : `SYSTEM ${nextChar}`)).trim()
    const sysQty = breakdown.setsQty
    const bayCount = Math.max(1, Math.round(breakdown.dfmuRows / 2))

    breakdown.models.forEach((m) => {
      allItems.push({
        id: `${sys.id || sIdx}-${m.code}-${counter++}`,
        title: `${m.code} ${bayCount} BAY`,
        sub_label: `BODY IN ${m.code}`,
        per_bay_qty: m.rowsPerUnit,
        qty: m.qty,
        total_body: m.totalBodies,
        system_name: sysName,
        system_id: sys.id || String(sIdx),
        set_qty: sysQty,
        unit_body: m.unitBodies,
      })
    })
  })

  return allItems
}
