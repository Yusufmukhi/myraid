import _XLSX from 'xlsx-js-style'
const XLSX: any = (_XLSX as any)?.utils ? _XLSX : ((_XLSX as any)?.default || _XLSX)
import { Workbook, addr } from './engine'
import { getSheetMeta } from './excelStyles'
import { getWidthExtra } from './model'

export interface C2CExportData {
  profit: number
  total: number
  ratio: number
  costBase?: number
  labour?: number
  steelWeight?: number
  subtotalBeforeGst?: number
  setsQty?: number
  sysTransport?: number
  sysAccommodation?: number
}

export interface ExcelExportOptions {
  setsQty?: number
  subtotalBeforeGst?: number
  sysTransport?: number
  sysAccommodation?: number
  modelSummaryStr?: string
  isActual?: boolean
  profitAmount?: number
  profitPct?: number
  discountPct?: number
  includeInstallation?: boolean
  extraAmount?: number
}

export interface MultiSystemExcelEntry {
  name: string
  bay: number
  setsQty: number
  overrides: Record<string, number | null>
  formulaOverrides: Record<string, string>
  profitPct: number
  discountPct: number
  includeInstallation: boolean
  extraAmount: number
  allocatedTransport: number
  allocatedAccommodation: number
  subtotalBeforeGst: number
  preDiscountTotal?: number
  profitAmount?: number
  summary: any
  c2cData?: C2CExportData
}

export interface MultiSystemExcelOptions {
  title?: string
  clientName?: string
  qtnNo?: string
  isActual: boolean
  profitPct: number
  discountPct: number
  includeInstallation: boolean
}

/**
 * Resolves named constants (e.g. CONST_STEEL_FACTOR, CONST_SURFACE_FACTOR) and fixes
 * unquoted identifiers inside an Excel formula outside string literals so Excel never shows #NAME?.
 * Returns null if any CONST_* constant remains unresolved (caller falls back to evaluated numeric value).
 */
export function resolveExcelFormula(
  rawFormula: string,
  namedConstants: Record<string, number>,
): string | null {
  const expr = rawFormula.startsWith('=') ? rawFormula.slice(1) : rawFormula
  const parts = expr.split(/("(?:[^"]|"")*")/g)
  const resolved = parts
    .map((part, idx) => {
      if (idx % 2 === 1) return part // inside quotes
      let seg = part
      for (const [key, val] of Object.entries(namedConstants)) {
        if (val !== undefined && val !== null && !isNaN(Number(val))) {
          const escapedKey = key.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
          seg = seg.replace(new RegExp(`\\b${escapedKey}\\b`, 'g'), String(val))
        }
      }
      // Fix legacy unquoted bare identifiers in workbook formulas that trigger #NAME? in Excel
      seg = seg.replace(/\bNO\b(?!\s*[\(!])/g, '"NO"')
      seg = seg.replace(/\bROW\b(?!\s*[\(!])/g, 'A3')
      return seg
    })
    .join('')

  const stripped = resolved.replace(/"(?:[^"]|"")*"/g, '')
  if (/\bCONST_[A-Za-z0-9_]+\b/.test(stripped)) {
    return null
  }
  return resolved
}

/**
 * Rewrites cross-sheet references when exporting a 3-sheet subset per system
 * (Total cal, SHEARING SIZE, HARDWARE LIST). If the formula references any sheet outside
 * allowedTargetSheets (e.g. DFMU, Sheet1, kg), returns null so the cell preserves its
 * live evaluated value without causing #REF! in Excel.
 */
export function rewriteFormulaForSubset(
  expr: string,
  sheetRenameMap: Record<string, string>,
  allowedTargetSheets: Set<string>,
): string | null {
  const parts = expr.split(/("(?:[^"]|"")*")/g)
  let hasExternalDisallowedRef = false

  const rewritten = parts
    .map((part, idx) => {
      if (idx % 2 === 1) return part // inside quotes
      return part.replace(/(?:'([^']+)'|([A-Za-z0-9_.\-]+))!/g, (_match, quoted, bare) => {
        const refName = quoted || bare
        if (!allowedTargetSheets.has(refName)) {
          hasExternalDisallowedRef = true
          return _match
        }
        const renamed = sheetRenameMap[refName] || refName
        const needsQuotes = /[^A-Za-z0-9_.]/.test(renamed)
        return (needsQuotes ? `'${renamed.replace(/'/g, "''")}'` : renamed) + '!'
      })
    })
    .join('')

  return hasExternalDisallowedRef ? null : rewritten
}

function makeUniqueTabName(prefix: string, suffix: string, usedNames: Set<string>): string {
  const cleanPrefix = (prefix || 'Sys').replace(/[\\/?*[\]:]/g, '_').trim()
  const sep = suffix ? ' - ' : ''
  const maxPrefixLen = Math.max(4, 31 - sep.length - suffix.length)
  let candidate = `${cleanPrefix.slice(0, maxPrefixLen)}${sep}${suffix}`.slice(0, 31)
  let counter = 2
  while (usedNames.has(candidate)) {
    const tag = `_${counter++}`
    const adjLen = Math.max(2, maxPrefixLen - tag.length)
    candidate = `${cleanPrefix.slice(0, adjLen)}${tag}${sep}${suffix}`.slice(0, 31)
  }
  usedNames.add(candidate)
  return candidate
}

const THIN_BORDER = {
  top: { style: 'thin', color: { rgb: 'CBD5E1' } },
  bottom: { style: 'thin', color: { rgb: 'CBD5E1' } },
  left: { style: 'thin', color: { rgb: 'CBD5E1' } },
  right: { style: 'thin', color: { rgb: 'CBD5E1' } },
}

const DARK_BORDER = {
  top: { style: 'thin', color: { rgb: '334155' } },
  bottom: { style: 'thin', color: { rgb: '334155' } },
  left: { style: 'thin', color: { rgb: '334155' } },
  right: { style: 'thin', color: { rgb: '334155' } },
}

/**
 * Temporarily applies scaled unit quantities (G4, G8, G12, G16, G20), profit %, and
 * transport/accommodation onto `wb` so every cross-sheet formula evaluates for `setsQty` sets,
 * invokes `fn`, and guarantees restoring original overrides afterwards.
 */
function withScaledWorkbookOverrides<T>(
  wb: Workbook,
  setsQty: number,
  profitPct: number | undefined,
  allocatedTransport: number | undefined,
  allocatedAccommodation: number | undefined,
  fn: () => T,
): T {
  const origOverrides = wb.getOverrides()
  const origFormulaOverrides = wb.getFormulaOverrides()
  const scaled: Record<string, number | null> = { ...origOverrides }

  if (setsQty > 1) {
    const qtyRows = [3, 7, 11, 15, 19]
    const qtyKeys = ['Total cal!G4', 'Total cal!G8', 'Total cal!G12', 'Total cal!G16', 'Total cal!G20']
    qtyRows.forEach((r, idx) => {
      const baseVal = Number(wb.value('Total cal', r, 6) ?? wb.raw('Total cal', r, 6)) || 0
      scaled[qtyKeys[idx]] = baseVal * setsQty
    })
  }
  if (allocatedTransport !== undefined) {
    scaled['Total cal!N24'] = allocatedTransport
  }
  if (allocatedAccommodation !== undefined) {
    scaled['Total cal!N25'] = allocatedAccommodation
  }
  if (profitPct !== undefined && !isNaN(profitPct)) {
    scaled['Total cal!C26'] = profitPct
  }

  wb.setOverrides(scaled)
  try {
    return fn()
  } finally {
    wb.setOverrides(origOverrides)
    wb.setFormulaOverrides(origFormulaOverrides)
  }
}

interface CrossSheetSubsetConfig {
  sheetRenameMap: Record<string, string>
  allowedTargetSheets: Set<string>
}

function buildStyledWorksheet(
  wb: Workbook,
  sheetName: string,
  c2cData?: C2CExportData,
  options?: ExcelExportOptions,
  crossSheetConfig?: CrossSheetSubsetConfig,
): any | null {
  const sheetData = wb.sheets.get(sheetName)
  if (!sheetData) return null

  const namedConstants = wb.getNamedConstants ? wb.getNamedConstants() : {}
  const setsQty = Math.max(1, Number(options?.setsQty ?? c2cData?.setsQty) || 1)
  const isActualExport = Boolean(options?.isActual || !c2cData)
  const includeInstallation = options?.includeInstallation !== false
  const extraAmount = isNaN(Number(options?.extraAmount)) ? 0 : Math.max(0, Number(options?.extraAmount))
  const discountPct = isNaN(Number(options?.discountPct)) ? 0 : Math.max(0, Number(options?.discountPct))

  const width = Number(wb.value('Total cal', 2, 3)) || 1830
  const extraDfmu = getWidthExtra(width, 'DFMU')
  const extraSfmu = getWidthExtra(width, 'SFMU')
  const extraDfmuD = getWidthExtra(width, 'DFMU-D')

  const dfmuQ = Number(wb.value('Total cal', 3, 6) ?? wb.raw('Total cal', 3, 6)) || 0
  const sfmuQ = Number(wb.value('Total cal', 7, 6) ?? wb.raw('Total cal', 7, 6)) || 0
  const sffuQ = Number(wb.value('Total cal', 11, 6) ?? wb.raw('Total cal', 11, 6)) || 0
  const dffuQ = Number(wb.value('Total cal', 15, 6) ?? wb.raw('Total cal', 15, 6)) || 0
  const dfmudQ = Number(wb.value('Total cal', 19, 6) ?? wb.raw('Total cal', 19, 6)) || 0
  const totalUnitsScaled = dfmuQ + sfmuQ + sffuQ + dffuQ + dfmudQ
  const totalExtraAmountSum = extraAmount * totalUnitsScaled

  const unitModelParts: string[] = []
  const perSetDiv = setsQty > 1 ? setsQty : 1
  if (dfmuQ > 0) unitModelParts.push(`${Math.round(dfmuQ / perSetDiv)} DFMU`)
  if (sffuQ > 0) unitModelParts.push(`${Math.round(sffuQ / perSetDiv)} SFFU`)
  if (sfmuQ > 0) unitModelParts.push(`${Math.round(sfmuQ / perSetDiv)} SFMU`)
  if (dffuQ > 0) unitModelParts.push(`${Math.round(dffuQ / perSetDiv)} DFFU`)
  if (dfmudQ > 0) unitModelParts.push(`${Math.round(dfmudQ / perSetDiv)} DFMU-D`)
  const systemModelSummary = options?.modelSummaryStr || (unitModelParts.length > 0 ? unitModelParts.join(' + ') : '1 SYSTEM')

  // Compute exact financial metrics for Total cal cached values
  const effectiveProfitPct = (options?.profitPct !== undefined && !isNaN(options.profitPct))
    ? options.profitPct
    : (Number(wb.value('Total cal', 25, 2)) || 25)
  const d4Val = (Number(wb.value('Total cal', 3, 3)) || 0) + extraDfmu
  const d8Val = (Number(wb.value('Total cal', 7, 3)) || 0) + extraSfmu
  const d12Val = Number(wb.value('Total cal', 11, 3)) || 0
  const d16Val = Number(wb.value('Total cal', 15, 3)) || 0
  const d20Val = (Number(wb.value('Total cal', 19, 3)) || 0) + extraDfmuD

  const exWorksTotal = dfmuQ * d4Val + sfmuQ * d8Val + sffuQ * d12Val + dffuQ * d16Val + dfmudQ * d20Val
  const instTotal = includeInstallation ? (Number(wb.value('Total cal', 22, 13)) || 0) : 0
  const transTotal = Number(wb.value('Total cal', 23, 13)) || 0
  const accTotal = Number(wb.value('Total cal', 24, 13)) || 0
  const directCostBase = exWorksTotal + instTotal + transTotal + accTotal + totalExtraAmountSum
  const computedProfit = directCostBase * (effectiveProfitPct / 100)
  const subtotalWithProfit = directCostBase + computedProfit
  const discountAmount = subtotalWithProfit * (discountPct / 100)
  const afterDiscount = subtotalWithProfit - discountAmount
  const gstAmount = afterDiscount * 0.18
  const grandTotalSales = afterDiscount + gstAmount

  const normSheet = sheetName.toLowerCase().replace(/[^a-z0-9]/g, '')
  const isShearing = normSheet.includes('shearing')

  const rows = sheetData.rows
  const cols = sheetData.cols
  const targetCols = sheetName === 'Total cal' ? Math.max(cols, 23) : cols
  const aoa: (string | number | boolean | null)[][] = []

  for (let r = 0; r < rows; r++) {
    const rowArr: (string | number | boolean | null)[] = []
    for (let c = 0; c < targetCols; c++) {
      if (sheetName === 'Total cal') {
        // Unit rates with width extra when width >= 3000
        if (c === 3 && r === 3 && extraDfmu > 0) { rowArr.push(d4Val); continue }
        if (c === 3 && r === 7 && extraSfmu > 0) { rowArr.push(d8Val); continue }
        if (c === 3 && r === 19 && extraDfmuD > 0) { rowArr.push(d20Val); continue }

        // N22 (r=21, c=13): Total ex-works
        if (r === 21 && c === 13) { rowArr.push(exWorksTotal); continue }
        // N23 (r=22, c=13): Installation (0 if excluded)
        if (r === 22 && c === 13) { rowArr.push(instTotal); continue }
        // N24 (r=23, c=13): Transport
        if (r === 23 && c === 13) { rowArr.push(transTotal); continue }
        // N25 (r=24, c=13): Accommodation
        if (r === 24 && c === 13) { rowArr.push(accTotal); continue }
        // Extra amount cell (J25 label, K25 value) if extraAmount > 0
        if (r === 24 && c === 9 && totalExtraAmountSum > 0) { rowArr.push('EXTRA AMT'); continue }
        if (r === 24 && c === 10 && totalExtraAmountSum > 0) { rowArr.push(totalExtraAmountSum); continue }

        // C26 (r=25, c=2): Actual Quotation Profit %
        if (r === 25 && c === 2) { rowArr.push(effectiveProfitPct); continue }
        // N26 (r=25, c=13): Profit on SUM(N22:N25) * C26 / 100
        if (r === 25 && c === 13) { rowArr.push(computedProfit); continue }
        // N27 (r=26, c=13): BEFORE GST (Cost Base + Profit)
        if (r === 26 && c === 13) { rowArr.push(subtotalWithProfit); continue }
        // N28 (r=27, c=13): 18% GST
        if (r === 27 && c === 13) { rowArr.push(gstAmount); continue }
        // N29 (r=28, c=13): SALES (Incl. GST)
        if (r === 28 && c === 13) { rowArr.push(grandTotalSales); continue }

        // R5:S9 summary cards:
        // - Actual Export: R5=Subtotal Before GST (=N27), R6=Profit Amount (=N26), R7=Quotation Sales (=N27 / =S11), R8=Profit Margin % (=C26), R9:R13=Discount & GST
        // - C2C Export: R5=Subtotal Before GST (quotation before GST), R6=C2C ANALYSIS ('✓ synced'), R7=Profit: (=S5-SUM(N22:N25)), R8=Total: (=S7+N50), R9=Ratio: (=IF(J44>0,S8/J44,0))
        if (r === 4 && c === 17) { rowArr.push('Subtotal Before GST'); continue }
        if (r === 4 && c === 18) { rowArr.push(c2cData ? (options?.subtotalBeforeGst ?? c2cData.subtotalBeforeGst ?? subtotalWithProfit) : subtotalWithProfit); continue }

        if (isActualExport) {
          if (r === 5 && c === 17) { rowArr.push('Profit Amount'); continue }
          if (r === 5 && c === 18) { rowArr.push(Math.round(computedProfit)); continue }

          if (r === 6 && c === 17) { rowArr.push('Quotation Sales'); continue }
          if (r === 6 && c === 18) { rowArr.push(Math.round(afterDiscount)); continue }

          if (r === 7 && c === 17) { rowArr.push('Profit Margin %'); continue }
          if (r === 7 && c === 18) { rowArr.push(effectiveProfitPct); continue }

          if (discountPct > 0) {
            if (r === 8 && c === 17) { rowArr.push('Discount %'); continue }
            if (r === 8 && c === 18) { rowArr.push(discountPct); continue }
            if (r === 9 && c === 17) { rowArr.push('Discount Amount'); continue }
            if (r === 9 && c === 18) { rowArr.push(discountAmount); continue }
            if (r === 10 && c === 17) { rowArr.push('After Discount'); continue }
            if (r === 10 && c === 18) { rowArr.push(afterDiscount); continue }
            if (r === 11 && c === 17) { rowArr.push('18% GST'); continue }
            if (r === 11 && c === 18) { rowArr.push(gstAmount); continue }
            if (r === 12 && c === 17) { rowArr.push('Grand Total (Incl. GST)'); continue }
            if (r === 12 && c === 18) { rowArr.push(grandTotalSales); continue }
          } else {
            if (r === 8 && c === 17) { rowArr.push('18% GST'); continue }
            if (r === 8 && c === 18) { rowArr.push(gstAmount); continue }
            if (r === 9 && c === 17) { rowArr.push('Grand Total (Incl. GST)'); continue }
            if (r === 9 && c === 18) { rowArr.push(grandTotalSales); continue }
          }
        } else {
          // C2C ANALYSIS block starting from R6
          const c2cBeforeGst = Number(options?.subtotalBeforeGst ?? c2cData?.subtotalBeforeGst ?? subtotalWithProfit)
          const c2cProfitCached = c2cData?.profit !== undefined
            ? Math.round(c2cData.profit)
            : Math.round(c2cBeforeGst - directCostBase)
          const c2cLabourVal = Number(wb.value(sheetName, 49, 13)) || Number(c2cData?.labour) || 0
          const c2cTotalCached = c2cData?.total !== undefined
            ? Math.round(c2cData.total)
            : Math.round(c2cProfitCached + c2cLabourVal)
          const c2cWeightVal = Number(wb.value(sheetName, 43, 9)) || Number(c2cData?.steelWeight) || 0
          const c2cRatioCached = c2cData?.ratio !== undefined
            ? Number(c2cData.ratio.toFixed(2))
            : Number((c2cWeightVal > 0 ? c2cTotalCached / c2cWeightVal : 0).toFixed(2))

          if (r === 5 && c === 17) { rowArr.push('C2C ANALYSIS'); continue }
          if (r === 5 && c === 18) { rowArr.push('✓ synced'); continue }
          if (r === 6 && c === 17) { rowArr.push('Profit:'); continue }
          if (r === 6 && c === 18) { rowArr.push(c2cProfitCached); continue }
          if (r === 7 && c === 17) { rowArr.push('Total:'); continue }
          if (r === 7 && c === 18) { rowArr.push(c2cTotalCached); continue }
          if (r === 8 && c === 17) { rowArr.push('Ratio:'); continue }
          if (r === 8 && c === 18) { rowArr.push(c2cRatioCached); continue }
        }

        // Row 22: O22 (c=14) SET label, P22 (c=15) sets quantity value
        if (r === 21 && c === 14) { rowArr.push('SET'); continue }
        if (r === 21 && c === 15) { rowArr.push(setsQty); continue }
      }

      if (c >= cols) {
        rowArr.push(null)
        continue
      }

      // Shearing size sheet customizations
      if (isShearing) {
        if (r === 5) {
          if (c === 1) { rowArr.push(systemModelSummary); continue }
          if (c === 4) { rowArr.push(setsQty); continue }
          if (c === 5) { rowArr.push(setsQty > 1 ? 'SETS' : 'SET'); continue }
        }
        if (r === 44) {
          if (c === 5) {
            const stiffenerVal = wb.value('Total cal', 2, 9)
            rowArr.push(typeof stiffenerVal === 'number' || typeof stiffenerVal === 'string' ? stiffenerVal : 0)
            continue
          }
          if (c === 7) {
            const stiffenerVal = wb.value('Total cal', 2, 9)
            if (!stiffenerVal || Number(stiffenerVal) === 0) {
              rowArr.push(0)
            } else {
              const h45Val = wb.value(sheetName, 44, 7) ?? wb.value('Sheet1', 81, 8) ?? 0
              rowArr.push(typeof h45Val === 'number' || typeof h45Val === 'string' ? h45Val : 0)
            }
            continue
          }
        }
      }

      const evaluatedVal = wb.value(sheetName, r, c)
      if (typeof evaluatedVal === 'number' || typeof evaluatedVal === 'string' || typeof evaluatedVal === 'boolean') {
        rowArr.push(evaluatedVal)
      } else {
        const raw = wb.raw(sheetName, r, c)
        rowArr.push(typeof raw === 'number' || typeof raw === 'string' ? raw : null)
      }
    }
    aoa.push(rowArr)
  }

  const ws = XLSX.utils.aoa_to_sheet(aoa)

  const applyFormula = (r: number, c: number, expr: string | null) => {
    if (!expr) return
    const ref = addr(r, c)
    if (!ws[ref]) return
    if (crossSheetConfig) {
      const rewritten = rewriteFormulaForSubset(expr, crossSheetConfig.sheetRenameMap, crossSheetConfig.allowedTargetSheets)
      if (rewritten !== null) {
        ws[ref].f = rewritten
      }
    } else {
      ws[ref].f = expr
    }
  }

  // Restore / inject live formulas without leading '=' so Excel evaluates everything dynamically
  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < targetCols; c++) {
      if (sheetName === 'Total cal') {
        // Explicit custom Total cal formulas
        if (c === 3 && r === 3 && extraDfmu > 0) { applyFormula(r, c, `DFMU!E85+${extraDfmu}`); continue }
        if (c === 3 && r === 7 && extraSfmu > 0) { applyFormula(r, c, `SFMU!E88+${extraSfmu}`); continue }
        if (c === 3 && r === 19 && extraDfmuD > 0) { applyFormula(r, c, `'DFMU-D'!E97+${extraDfmuD}`); continue }

        if (c === 6 && [3, 7, 11, 15, 19].includes(r)) continue // scaled quantities are numeric inputs
        if (r === 21 && (c === 14 || c === 15)) continue // O22/P22 SET count
        if (r === 24 && (c === 9 || c === 10) && totalExtraAmountSum > 0) {
          if (c === 10) applyFormula(r, c, `(G4+G8+G12+G16+G20)*${extraAmount}`)
          continue
        }

        // Installation N23: =B50 when included, numeric 0 when excluded
        if (r === 22 && c === 13) {
          if (includeInstallation) applyFormula(r, c, 'B50')
          continue
        }
        if (r === 23 && c === 13) continue // N24 transport numeric input
        if (r === 24 && c === 13) continue // N25 acc numeric input
        if (r === 25 && c === 2) continue // C26 profit % numeric input

        // N26 (Profit): SUM(N22:N25) * C26 / 100 (using quotation profit %, never hardcoded 20%)
        if (r === 25 && c === 13) {
          applyFormula(r, c, totalExtraAmountSum > 0 ? '(SUM(N22:N25)+K25)*C26/100' : 'SUM(N22:N25)*C26/100')
          continue
        }
        // N27 (Before GST / With Profit): SUM(N22:N26)
        if (r === 26 && c === 13) {
          applyFormula(r, c, totalExtraAmountSum > 0 ? 'SUM(N22:N26)+K25' : 'SUM(N22:N26)')
          continue
        }
        // N28 (18% GST) & N29 (Sales)
        if (r === 27 && c === 13) {
          applyFormula(r, c, discountPct > 0 ? 'S11*18%' : 'N27*18%')
          continue
        }
        if (r === 28 && c === 13) {
          applyFormula(r, c, discountPct > 0 ? 'S11+N28' : 'SUM(N27:N28)')
          continue
        }

        // R5:S13 live summary cards
        if (r === 4 && c === 18) {
          // Link S5 directly to N27 when Actual export (never hardcoded!)
          // In C2C export, S5 holds Quotation Subtotal Before GST (referenced by S7 Profit formula)
          if (isActualExport) applyFormula(r, c, 'N27')
          continue
        }
        if (r === 5 && c === 18) {
          if (isActualExport) applyFormula(r, c, 'N26')
          continue
        }
        if (r === 6 && c === 18) {
          if (isActualExport) {
            applyFormula(r, c, discountPct > 0 ? 'S11' : 'N27')
          } else {
            // C2C Profit = Quotation Before GST (S5) - C2C Direct Cost Base (SUM(N22:N25) [+ K25])
            applyFormula(r, c, totalExtraAmountSum > 0 ? 'S5-(SUM(N22:N25)+K25)' : 'S5-SUM(N22:N25)')
          }
          continue
        }
        if (r === 7 && c === 18) {
          if (isActualExport) {
            applyFormula(r, c, 'C26')
          } else {
            // C2C Total = Profit (S7) + C2C Labour Charges (N50)
            applyFormula(r, c, 'S7+N50')
          }
          continue
        }
        if (!isActualExport && r === 8 && c === 18) {
          // C2C Ratio = Total (S8) / Total Steel Weight (J44)
          applyFormula(r, c, 'IF(J44>0,S8/J44,0)')
          continue
        }
        if (isActualExport && (c === 17 || c === 18) && r >= 8 && r <= 12) {
          if (discountPct > 0) {
            if (r === 9 && c === 18) applyFormula(r, c, 'N27*S9/100')
            if (r === 10 && c === 18) applyFormula(r, c, 'N27-S10')
            if (r === 11 && c === 18) applyFormula(r, c, 'S11*0.18')
            if (r === 12 && c === 18) applyFormula(r, c, 'S11+S12')
          } else {
            if (r === 8 && c === 18) applyFormula(r, c, 'N27*0.18')
            if (r === 9 && c === 18) applyFormula(r, c, 'N27+S9')
          }
          continue
        }
        if ((r >= 4 && r <= 12) && (c === 17 || c === 18)) continue
      }

      if (c >= cols) continue

      // Shearing size sheet customizations
      if (isShearing && r === 5 && (c === 1 || c === 4 || c === 5)) continue
      if (isShearing && r === 44 && c === 5) {
        applyFormula(r, c, "'Total cal'!J3")
        continue
      }
      if (isShearing && r === 44 && c === 7) {
        applyFormula(r, c, "IF('Total cal'!J3=0,0,Sheet1!I82)")
        continue
      }

      if (!wb.isOverridden?.(sheetName, r, c)) {
        const raw = wb.raw(sheetName, r, c)
        if (typeof raw === 'string' && raw.startsWith('=')) {
          const resolvedFormula = resolveExcelFormula(raw, namedConstants)
          applyFormula(r, c, resolvedFormula)
        }
      }
    }
  }

  // Load authentic template styling, column widths, and merges
  const meta = getSheetMeta(sheetName)
  if (meta?.cols && meta.cols.length > 0) {
    ws['!cols'] = meta.cols.map((col: any) => ({ ...col }))
  } else {
    const fallbackCols = []
    for (let c = 0; c < Math.min(cols, 20); c++) {
      let maxLen = 8
      for (let r = 0; r < Math.min(rows, 35); r++) {
        const val = aoa[r]?.[c]
        if (val !== null && val !== undefined) {
          const s = String(val)
          if (s.length > 0 && s.length <= 18 && s.length > maxLen) maxLen = s.length
        }
      }
      fallbackCols.push({ wch: Math.min(18, Math.max(10, maxLen + 2)) })
    }
    ws['!cols'] = fallbackCols
  }

  if (sheetName === 'Total cal') {
    const currentCols = ws['!cols'] || []
    while (currentCols.length < 23) currentCols.push({ wch: 10 })
    currentCols[17] = { wch: 20 } // Col R
    currentCols[18] = { wch: 15 } // Col S
    ws['!cols'] = currentCols
  }

  if (meta?.rows && meta.rows.length > 0) {
    ws['!rows'] = meta.rows.map((r: any) => (r ? { ...r } : null))
  }
  if (meta?.merges && meta.merges.length > 0) {
    ws['!merges'] = meta.merges
  }
  if (meta?.styles) {
    for (const [cellRef, style] of Object.entries(meta.styles)) {
      if (ws[cellRef]) ws[cellRef].s = style
    }
  }

  // Style Total cal custom summary cells
  if (sheetName === 'Total cal') {
    const styleLabel = (ref: string) => {
      if (ws[ref]) {
        ws[ref].s = {
          font: { bold: true, color: { rgb: '1E293B' }, sz: 10 },
          fill: { fgColor: { rgb: 'F1F5F9' } },
          alignment: { horizontal: 'right', vertical: 'center' },
          border: THIN_BORDER,
        }
      }
    }
    const styleMoney = (ref: string, bg = 'FFFFFF', fg = '0F172A', fmt?: string) => {
      if (ws[ref]) {
        ws[ref].s = {
          font: { bold: true, color: { rgb: fg }, sz: 11 },
          fill: { fgColor: { rgb: bg } },
          alignment: { horizontal: 'right', vertical: 'center' },
          border: THIN_BORDER,
          numFmt: '"₹"#,##0.00;[Red]-"₹"#,##0.00;"-"',
        }
        if (fmt) ws[ref].s.numFmt = fmt
      }
    }

    styleLabel('R5')
    styleMoney('S5', 'E2EFDA', '0F172A')
    styleLabel('R6')
    if (isActualExport) {
      styleMoney('S6', '92D050', '166534', '"₹"#,##0;[Red]-"₹"#,##0;"-"')
      styleLabel('R7')
      styleMoney('S7', '92D050', '166534', '"₹"#,##0;[Red]-"₹"#,##0;"-"')
      styleLabel('R8')
      styleMoney('S8', 'DBEAFE', '1E3A8A', '0.00"%"')
    } else {
      // C2C ANALYSIS card styling from R6:S9
      if (ws['R6']) {
        ws['R6'].s = {
          font: { bold: true, color: { rgb: '1E3A8A' }, sz: 10 },
          fill: { fgColor: { rgb: 'DBEAFE' } },
          alignment: { horizontal: 'right', vertical: 'center' },
          border: THIN_BORDER,
        }
      }
      if (ws['S6']) {
        ws['S6'].s = {
          font: { bold: true, color: { rgb: '166534' }, sz: 10 },
          fill: { fgColor: { rgb: '92D050' } },
          alignment: { horizontal: 'center', vertical: 'center' },
          border: THIN_BORDER,
        }
      }
      styleLabel('R7')
      styleMoney('S7', 'E2EFDA', '166534', '"₹"#,##0;[Red]-"₹"#,##0;"-"')
      styleLabel('R8')
      styleMoney('S8', 'E2EFDA', '166534', '"₹"#,##0;[Red]-"₹"#,##0;"-"')
      styleLabel('R9')
      styleMoney('S9', 'DBEAFE', '1E3A8A', '"₹"#,##0.00"/kg"')
    }

    if (isActualExport) {
      if (discountPct > 0) {
        styleLabel('R9'); styleMoney('S9', 'FEF9C3', '854D0E', '0.00"%"')
        styleLabel('R10'); styleMoney('S10', 'FEF2F2', 'B91C1C')
        styleLabel('R11'); styleMoney('S11', 'E2EFDA', '0F172A')
        styleLabel('R12'); styleMoney('S12', 'F8FAFC', '334155')
        styleLabel('R13'); styleMoney('S13', 'DBEAFE', '1E3A8A')
      } else {
        styleLabel('R9'); styleMoney('S9', 'F8FAFC', '334155')
        styleLabel('R10'); styleMoney('S10', 'DBEAFE', '1E3A8A')
      }
    }

    if (ws['O22']) {
      ws['O22'].s = {
        font: { bold: true, color: { rgb: '1E3A8A' }, sz: 10 },
        fill: { fgColor: { rgb: 'DBEAFE' } },
        alignment: { horizontal: 'center', vertical: 'center' },
        border: THIN_BORDER,
      }
    }
    if (ws['P22']) {
      ws['P22'].s = {
        font: { bold: true, color: { rgb: '1E3A8A' }, sz: 11 },
        fill: { fgColor: { rgb: 'EFF6FF' } },
        alignment: { horizontal: 'center', vertical: 'center' },
        border: THIN_BORDER,
      }
    }
  }

  if (isShearing) {
    if (ws['E6']) {
      ws['E6'].s = {
        font: { bold: true, color: { rgb: '000000' }, sz: 11 },
        alignment: { horizontal: 'center', vertical: 'center' },
        fill: { fgColor: { rgb: 'FFFF00' } },
        border: { top: { style: 'thin', color: { rgb: 'D9D9D9' } }, bottom: { style: 'thin', color: { rgb: 'D9D9D9' } }, left: { style: 'thin', color: { rgb: 'D9D9D9' } }, right: { style: 'thin', color: { rgb: 'D9D9D9' } } },
      }
    }
    if (ws['F6']) {
      ws['F6'].s = {
        font: { bold: true, color: { rgb: '000000' }, sz: 11 },
        alignment: { horizontal: 'center', vertical: 'center' },
        border: { top: { style: 'thin', color: { rgb: 'D9D9D9' } }, bottom: { style: 'thin', color: { rgb: 'D9D9D9' } }, left: { style: 'thin', color: { rgb: 'D9D9D9' } }, right: { style: 'thin', color: { rgb: 'D9D9D9' } } },
      }
    }
  }

  if (isShearing) {
    ws['!pageSetup'] = { paperSize: 9, orientation: 'portrait' }
    ws['!margins'] = { left: 0.7, right: 0.7, top: 0.2, bottom: 0.2, header: 0.3, footer: 0.3 }
  } else {
    ws['!margins'] = { left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 }
  }

  ws['!views'] = [{ showGridLines: true }]
  return ws
}

/**
 * Exports a single-system multi-sheet Workbook instance to a genuine styled .xlsx Excel file.
 * Preserves all 15 sheets with live formulas, resolved constants (0 #NAME? errors),
 * and full dynamic links (N26 = SUM(N22:N25)*C26/100, S5 = N27).
 */
export function exportWorkbookToXLSX(
  wb: Workbook,
  filename: string,
  c2cData?: C2CExportData,
  options?: ExcelExportOptions,
) {
  const book = XLSX.utils.book_new()
  const setsQty = Math.max(1, Number(options?.setsQty ?? c2cData?.setsQty) || 1)
  const allocatedTransport = options?.sysTransport ?? c2cData?.sysTransport
  const allocatedAccommodation = options?.sysAccommodation ?? c2cData?.sysAccommodation
  const profitPct = options?.profitPct

  withScaledWorkbookOverrides(wb, setsQty, profitPct, allocatedTransport, allocatedAccommodation, () => {
    for (const sheetName of wb.names) {
      const ws = buildStyledWorksheet(wb, sheetName, c2cData, options)
      if (ws) {
        XLSX.utils.book_append_sheet(book, ws, sheetName.slice(0, 31))
      }
    }
  })

  const cleanFileName = filename.endsWith('.xlsx') ? filename : `${filename}.xlsx`
  XLSX.writeFile(book, cleanFileName)
}

/**
 * Exports a multi-system quotation into a single Excel (.xlsx) workbook with:
 * - Sheet 1 ("Total cal"): Complete master calculation across all systems (Profit, Discount,
 *   Transport, Accommodation, Installation, GST) AND per-set + model breakdown for every system.
 * - For each system (SYSTEM A, SYSTEM B, ...): 3 linked sheets:
 *   1. "<System> - Total cal"
 *   2. "<System> - SHEARING SIZE"
 *   3. "<System> - HARDWARE LIST"
 */
export function exportMultiSystemWorkbookToXLSX(
  wb: Workbook,
  systemsData: MultiSystemExcelEntry[],
  filename: string,
  globalOptions: MultiSystemExcelOptions,
) {
  const book = XLSX.utils.book_new()
  const usedTabNames = new Set<string>(['Total cal'])
  const wbSheetKeys = Array.from(wb.sheets.keys())
  const tcSrcName = wbSheetKeys.find((k) => k.toLowerCase() === 'total cal') || 'Total cal'
  const shSrcName = wbSheetKeys.find((k) => /^shearing\b/i.test(k) && /\bsize$/i.test(k))
  const hwSrcName = wbSheetKeys.find((k) => /^hardware\b/i.test(k) && /\blist$/i.test(k))

  // 1. Pre-allocate unique <=31 char tab names for each system's 3 sheets
  const systemSheetsMeta = systemsData.map((sys, idx) => {
    const fallbackName = `SYS ${String.fromCharCode(65 + idx)}`
    const baseLabel = (sys.name || fallbackName).trim()
    const tcTab = makeUniqueTabName(baseLabel, tcSrcName, usedTabNames)
    const shTab = shSrcName ? makeUniqueTabName(baseLabel, shSrcName, usedTabNames) : ''
    const hwTab = hwSrcName ? makeUniqueTabName(baseLabel, hwSrcName, usedTabNames) : ''
    return { sys, idx, tcTab, shTab, hwTab }
  })

  // 2. Build Tab 1: Master "Total cal" Consolidated Summary Sheet
  const quoteRef = (val: string, colCell: string) => {
    const safeTab = `'${val.replace(/'/g, "''")}'`
    return `${safeTab}!${colCell}`
  }

  const masterRows: any[][] = []
  const masterFormulas: Array<{ r: number; c: number; f: string; fmt?: string }> = []
  const masterStyles: Array<{ r: number; c: number; s: any }> = []

  const headerStyle = {
    font: { bold: true, color: { rgb: 'FFFFFF' }, sz: 10 },
    fill: { fgColor: { rgb: '1E3A8A' } },
    alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
    border: DARK_BORDER,
  }
  const subHeaderStyle = {
    font: { bold: true, color: { rgb: '0F172A' }, sz: 10 },
    fill: { fgColor: { rgb: 'DBEAFE' } },
    alignment: { horizontal: 'left', vertical: 'center' },
    border: THIN_BORDER,
  }
  const cellMoneyFmt = '"₹"#,##0.00;[Red]-"₹"#,##0.00;"-"'
  const cellWeightFmt = '#,##0.00 "kg"'
  const cellPctFmt = '0.00"%"'

  const modeLabel = globalOptions.isActual ? 'ACTUAL QUOTATION' : 'CUT-TO-CUT ANALYSIS'
  masterRows.push([
    `${modeLabel} — CONSOLIDATED MASTER CALCULATION`,
    null, null, null, null, null, null, null, null, null, null, null, null, null, null,
  ])
  masterRows.push([
    `Quotation: ${globalOptions.qtnNo || globalOptions.title || 'Project'}`,
    null,
    `Client: ${globalOptions.clientName || '-'}`,
    null, null,
    `Systems: ${systemsData.length}`,
    null,
    `Profit %: ${globalOptions.profitPct}%`,
    null,
    `Discount %: ${globalOptions.discountPct}%`,
    null,
    `Installation: ${globalOptions.includeInstallation ? 'Included' : 'At Buyer Scope'}`,
    null, null, null,
  ])
  masterRows.push([]) // blank row 3

  // Section 1: System-Wise Consolidated Summary Table (row index 3 -> Excel Row 4)
  masterRows.push(['SYSTEM-WISE QUOTATION SUMMARY (ALL SYSTEMS & SETS)'])
  const sysTblHeaderRow = masterRows.length // 0-based
  masterRows.push([
    'Sl',
    'System Name',
    'Bay',
    'Sets',
    'Models Summary',
    'Bodies / Set',
    'Total Bodies',
    'Ex-Works (All Sets)',
    'Installation',
    'Transport',
    'Acc',
    'Extra Amt',
    'Cost Base',
    'Profit %',
    'Profit Amount',
    '1-Set Rate (Pre-Disc)',
    'With Profit',
    'Disc %',
    'Disc Amount',
    '1-Set Rate (Net)',
    'Net Before GST',
    'GST (18%)',
    'Total Incl. GST',
    'SP Steel (kg)',
    'Trolley (kg)',
    'Total Wt (kg)',
    'Wt w/ Mach (kg)',
    'Powder Coat (₹)',
    'Labour (₹)',
  ])
  for (let c = 0; c < 29; c++) {
    masterStyles.push({ r: sysTblHeaderRow, c, s: headerStyle })
  }

  const firstSysRow = masterRows.length + 1 // 1-based Excel row
  let totalCombinedSets = 0
  let totalCombinedBodiesAll = 0

  systemSheetsMeta.forEach(({ sys, idx, tcTab }) => {
    const excelRow = masterRows.length + 1
    const setsQty = Math.max(1, Number(sys.setsQty) || 1)
    const bodiesPerSet = Math.max(1, Number(sys.summary?.bodies) || 1)
    const totalBodiesSys = bodiesPerSet * setsQty
    totalCombinedSets += setsQty
    totalCombinedBodiesAll += totalBodiesSys

    const activeModelsStr = (sys.summary?.models || [])
      .filter((m: any) => m.qty > 0)
      .map((m: any) => `${m.qty} ${m.code}`)
      .join(' + ') || 'Custom'

    const exWorks = (sys.summary?.subtotal || 0) * setsQty
    const inst = (sys.includeInstallation ? (sys.summary?.rawInstallation ?? sys.summary?.installation ?? 0) : 0) * setsQty
    const trans = sys.allocatedTransport || 0
    const acc = sys.allocatedAccommodation || 0
    const totalUnitsOneSet = (sys.summary?.models || []).reduce((a: number, m: any) => a + (Number(m.qty) || 0), 0)
    const extraTotal = (sys.extraAmount || 0) * totalUnitsOneSet * setsQty
    const costBase = exWorks + inst + trans + acc + extraTotal
    const profitAmt = costBase * (sys.profitPct / 100)
    const withProfit = costBase + profitAmt
    const discAmt = withProfit * (sys.discountPct / 100)
    const netBeforeGst = withProfit - discAmt
    const gstAmt = netBeforeGst * 0.18
    const inclGst = netBeforeGst + gstAmt

    const spWt = (sys.summary?.steel?.spTotal || 0) * setsQty
    const trWt = (sys.summary?.trolley?.total || 0) * setsQty
    const totWt = (sys.summary?.totalWeight || 0) * setsQty
    const machWt = (sys.summary?.weightWithMachine || 0) * setsQty
    const pcAmt = (sys.summary?.powderTotal || 0) * setsQty
    const labAmt = (sys.summary?.labourSubtotal || 0) * setsQty

    const rIdx = masterRows.length
    masterRows.push([
      idx + 1,
      sys.name,
      sys.bay,
      setsQty,
      activeModelsStr,
      bodiesPerSet,
      totalBodiesSys,
      exWorks,
      inst,
      trans,
      acc,
      extraTotal,
      costBase,
      sys.profitPct,
      profitAmt,
      withProfit / setsQty,
      withProfit,
      sys.discountPct,
      discAmt,
      netBeforeGst / setsQty,
      netBeforeGst,
      gstAmt,
      inclGst,
      spWt,
      trWt,
      totWt,
      machWt,
      pcAmt,
      labAmt,
    ])

    // Live cross-sheet links to each system's Total cal tab + formula derivations
    masterFormulas.push(
      { r: rIdx, c: 6, f: `F${excelRow}*D${excelRow}` },
      { r: rIdx, c: 7, f: quoteRef(tcTab, 'N22'), fmt: cellMoneyFmt },
      { r: rIdx, c: 8, f: quoteRef(tcTab, 'N23'), fmt: cellMoneyFmt },
      { r: rIdx, c: 9, f: quoteRef(tcTab, 'N24'), fmt: cellMoneyFmt },
      { r: rIdx, c: 10, f: quoteRef(tcTab, 'N25'), fmt: cellMoneyFmt },
      ...(extraTotal > 0 ? [{ r: rIdx, c: 11, f: quoteRef(tcTab, 'K25'), fmt: cellMoneyFmt }] : []),
      { r: rIdx, c: 12, f: `SUM(H${excelRow}:L${excelRow})`, fmt: cellMoneyFmt },
      { r: rIdx, c: 13, f: quoteRef(tcTab, 'C26'), fmt: cellPctFmt },
      { r: rIdx, c: 14, f: quoteRef(tcTab, 'N26'), fmt: cellMoneyFmt },
      { r: rIdx, c: 15, f: `Q${excelRow}/D${excelRow}`, fmt: cellMoneyFmt },
      { r: rIdx, c: 16, f: quoteRef(tcTab, 'N27'), fmt: cellMoneyFmt },
      { r: rIdx, c: 18, f: `Q${excelRow}*(R${excelRow}/100)`, fmt: cellMoneyFmt },
      { r: rIdx, c: 19, f: `U${excelRow}/D${excelRow}`, fmt: cellMoneyFmt },
      { r: rIdx, c: 20, f: `Q${excelRow}-S${excelRow}`, fmt: cellMoneyFmt },
      { r: rIdx, c: 21, f: `U${excelRow}*0.18`, fmt: cellMoneyFmt },
      { r: rIdx, c: 22, f: `U${excelRow}+V${excelRow}`, fmt: cellMoneyFmt },
      { r: rIdx, c: 23, f: quoteRef(tcTab, 'J36'), fmt: cellWeightFmt },
      { r: rIdx, c: 24, f: quoteRef(tcTab, 'J42'), fmt: cellWeightFmt },
      { r: rIdx, c: 25, f: quoteRef(tcTab, 'J44'), fmt: cellWeightFmt },
      { r: rIdx, c: 26, f: quoteRef(tcTab, 'J45'), fmt: cellWeightFmt },
      { r: rIdx, c: 27, f: quoteRef(tcTab, 'E39'), fmt: cellMoneyFmt },
      { r: rIdx, c: 28, f: quoteRef(tcTab, 'N50'), fmt: cellMoneyFmt },
    )
    for (let c = 0; c < 29; c++) {
      const fmt = c === 13 || c === 17 ? cellPctFmt : (c >= 23 && c <= 26 ? cellWeightFmt : (c >= 7 && c <= 28 ? cellMoneyFmt : undefined))
      masterStyles.push({
        r: rIdx,
        c,
        s: {
          border: THIN_BORDER,
          alignment: { horizontal: c === 1 || c === 4 ? 'left' : 'right', vertical: 'center' },
          ...(fmt ? { numFmt: fmt } : {}),
        },
      })
    }
  })

  const lastSysRow = masterRows.length // 1-based Excel row of last system
  const sysTotalRowIdx = masterRows.length
  const sumRange = (colLetter: string) => `SUM(${colLetter}${firstSysRow}:${colLetter}${lastSysRow})`

  masterRows.push([
    '', 'TOTAL (ALL SYSTEMS)', '', totalCombinedSets, '', '', totalCombinedBodiesAll,
    0, 0, 0, 0, 0, 0, globalOptions.profitPct, 0, '', 0, globalOptions.discountPct, 0, '', 0, 0, 0,
    0, 0, 0, 0, 0, 0,
  ])
  const sumCols: Array<[number, string, string?]> = [
    [3, 'D'], [6, 'G'], [7, 'H', cellMoneyFmt], [8, 'I', cellMoneyFmt], [9, 'J', cellMoneyFmt],
    [10, 'K', cellMoneyFmt], [11, 'L', cellMoneyFmt], [12, 'M', cellMoneyFmt], [14, 'O', cellMoneyFmt],
    [16, 'Q', cellMoneyFmt], [18, 'S', cellMoneyFmt], [20, 'U', cellMoneyFmt], [21, 'V', cellMoneyFmt],
    [22, 'W', cellMoneyFmt], [23, 'X', cellWeightFmt], [24, 'Y', cellWeightFmt],
    [25, 'Z', cellWeightFmt], [26, 'AA', cellWeightFmt], [27, 'AB', cellMoneyFmt], [28, 'AC', cellMoneyFmt],
  ]
  sumCols.forEach(([colIdx, colLetter, fmt]) => {
    masterFormulas.push({ r: sysTotalRowIdx, c: colIdx, f: sumRange(colLetter), fmt })
  })
  for (let c = 0; c < 29; c++) {
    masterStyles.push({
      r: sysTotalRowIdx,
      c,
      s: {
        font: { bold: true, color: { rgb: '0F172A' }, sz: 10 },
        fill: { fgColor: { rgb: 'E2EFDA' } },
        border: DARK_BORDER,
        alignment: { horizontal: c === 1 ? 'left' : 'right', vertical: 'center' },
      },
    })
  }

  const totalExcelRow = sysTotalRowIdx + 1
  masterRows.push([]) // blank

  // Section 2: Commercial Grand Totals Block
  masterRows.push(['CONSOLIDATED COMMERCIAL SUMMARY (QUOTATION TOTALS)'])
  const summaryPairs: Array<[string, string, string]> = [
    ['Total Ex-Works (All Systems & Sets)', `H${totalExcelRow}`, cellMoneyFmt],
    ['Total Installation Charges', `I${totalExcelRow}`, cellMoneyFmt],
    ['Total Transportation Charges', `J${totalExcelRow}`, cellMoneyFmt],
    ['Total Accommodation Charges', `K${totalExcelRow}`, cellMoneyFmt],
    ['Total Extra Charges', `L${totalExcelRow}`, cellMoneyFmt],
    ['Total Direct Cost Base', `M${totalExcelRow}`, cellMoneyFmt],
    [`Profit (${globalOptions.profitPct}%)`, `O${totalExcelRow}`, cellMoneyFmt],
    ['Subtotal With Profit (Before Discount)', `Q${totalExcelRow}`, cellMoneyFmt],
    [`Discount (${globalOptions.discountPct}%)`, `S${totalExcelRow}`, cellMoneyFmt],
    ['Net Quotation Total (Before GST)', `U${totalExcelRow}`, cellMoneyFmt],
    ['GST @ 18%', `V${totalExcelRow}`, cellMoneyFmt],
    ['GRAND TOTAL (INCL. 18% GST)', `W${totalExcelRow}`, cellMoneyFmt],
  ]
  summaryPairs.forEach(([label, formulaExpr, fmt], idx) => {
    const rIdx = masterRows.length
    masterRows.push([label, null, 0])
    masterFormulas.push({ r: rIdx, c: 2, f: formulaExpr, fmt })
    const isGrand = idx === summaryPairs.length - 1
    masterStyles.push({
      r: rIdx, c: 0,
      s: {
        font: { bold: true, color: { rgb: isGrand ? 'FFFFFF' : '1E293B' }, sz: 10 },
        fill: { fgColor: { rgb: isGrand ? '1E3A8A' : 'F1F5F9' } },
        border: THIN_BORDER,
      },
    })
    masterStyles.push({
      r: rIdx, c: 2,
      s: {
        font: { bold: true, color: { rgb: isGrand ? '1E3A8A' : '0F172A' }, sz: isGrand ? 11 : 10 },
        fill: { fgColor: { rgb: isGrand ? 'DBEAFE' : 'FFFFFF' } },
        border: THIN_BORDER,
        alignment: { horizontal: 'right', vertical: 'center' },
        numFmt: fmt,
      },
    })
  })

  if (!globalOptions.isActual) {
    masterRows.push([])
    const c2cHeaderRowIdx = masterRows.length
    masterRows.push(['C2C ANALYSIS', null, '✓ synced'])
    masterStyles.push(
      { r: c2cHeaderRowIdx, c: 0, s: { font: { bold: true, color: { rgb: '1E3A8A' }, sz: 10 }, fill: { fgColor: { rgb: 'DBEAFE' } }, border: THIN_BORDER } },
      { r: c2cHeaderRowIdx, c: 2, s: { font: { bold: true, color: { rgb: '166534' }, sz: 10 }, fill: { fgColor: { rgb: '92D050' } }, alignment: { horizontal: 'center', vertical: 'center' }, border: THIN_BORDER } },
    )

    const s5Expr = systemSheetsMeta.map(({ tcTab }) => quoteRef(tcTab, 'S5')).join('+')
    const s5RowIdx = masterRows.length
    const s5ExcelRow = s5RowIdx + 1
    masterRows.push(['Subtotal Before GST (Quotation)', null, 0])
    masterFormulas.push({ r: s5RowIdx, c: 2, f: s5Expr, fmt: cellMoneyFmt })
    masterStyles.push(
      { r: s5RowIdx, c: 0, s: { font: { bold: true, color: { rgb: '1E293B' }, sz: 10 }, fill: { fgColor: { rgb: 'F1F5F9' } }, border: THIN_BORDER } },
      { r: s5RowIdx, c: 2, s: { font: { bold: true, color: { rgb: '0F172A' }, sz: 10 }, fill: { fgColor: { rgb: 'E2EFDA' } }, border: THIN_BORDER, alignment: { horizontal: 'right', vertical: 'center' }, numFmt: cellMoneyFmt } },
    )

    const pRowIdx = masterRows.length
    const pExcelRow = pRowIdx + 1
    masterRows.push(['Profit:', null, 0])
    masterFormulas.push({ r: pRowIdx, c: 2, f: `C${s5ExcelRow}-M${totalExcelRow}`, fmt: '"₹"#,##0;[Red]-"₹"#,##0;"-"' })
    masterStyles.push(
      { r: pRowIdx, c: 0, s: { font: { bold: true, color: { rgb: '1E293B' }, sz: 10 }, fill: { fgColor: { rgb: 'F1F5F9' } }, border: THIN_BORDER } },
      { r: pRowIdx, c: 2, s: { font: { bold: true, color: { rgb: '166534' }, sz: 11 }, fill: { fgColor: { rgb: 'E2EFDA' } }, border: THIN_BORDER, alignment: { horizontal: 'right', vertical: 'center' }, numFmt: '"₹"#,##0;[Red]-"₹"#,##0;"-"' } },
    )

    const tRowIdx = masterRows.length
    const tExcelRow = tRowIdx + 1
    masterRows.push(['Total:', null, 0])
    masterFormulas.push({ r: tRowIdx, c: 2, f: `C${pExcelRow}+AC${totalExcelRow}`, fmt: '"₹"#,##0;[Red]-"₹"#,##0;"-"' })
    masterStyles.push(
      { r: tRowIdx, c: 0, s: { font: { bold: true, color: { rgb: '1E293B' }, sz: 10 }, fill: { fgColor: { rgb: 'F1F5F9' } }, border: THIN_BORDER } },
      { r: tRowIdx, c: 2, s: { font: { bold: true, color: { rgb: '166534' }, sz: 11 }, fill: { fgColor: { rgb: 'E2EFDA' } }, border: THIN_BORDER, alignment: { horizontal: 'right', vertical: 'center' }, numFmt: '"₹"#,##0;[Red]-"₹"#,##0;"-"' } },
    )

    const rRowIdx = masterRows.length
    masterRows.push(['Ratio:', null, 0])
    masterFormulas.push({ r: rRowIdx, c: 2, f: `IF(Z${totalExcelRow}>0,C${tExcelRow}/Z${totalExcelRow},0)`, fmt: '"₹"#,##0.00"/kg"' })
    masterStyles.push(
      { r: rRowIdx, c: 0, s: { font: { bold: true, color: { rgb: '1E293B' }, sz: 10 }, fill: { fgColor: { rgb: 'F1F5F9' } }, border: THIN_BORDER } },
      { r: rRowIdx, c: 2, s: { font: { bold: true, color: { rgb: '1E3A8A' }, sz: 11 }, fill: { fgColor: { rgb: 'DBEAFE' } }, border: THIN_BORDER, alignment: { horizontal: 'right', vertical: 'center' }, numFmt: '"₹"#,##0.00"/kg"' } },
    )
  }

  masterRows.push([]) // blank

  // Section 3: Breakdown For Each System & Each Set (Model-Level Detail)
  masterRows.push(['SET-WISE & MODEL-WISE BREAKDOWN FOR EACH SYSTEM'])
  const bdHeaderIdx = masterRows.length
  masterRows.push([
    'Sl',
    'System / Set',
    'Model Code',
    'Unit Rate (₹)',
    'Width Extra (₹)',
    'Qty / Set',
    'Sets',
    'Total Qty',
    '1-Set Ex-Works (₹)',
    'All-Sets Ex-Works (₹)',
    'Powder Coat (₹)',
    'Cost / Unit Incl. Logistics',
  ])
  for (let c = 0; c < 12; c++) {
    masterStyles.push({ r: bdHeaderIdx, c, s: headerStyle })
  }

  systemSheetsMeta.forEach(({ sys, tcTab }) => {
    const setsQty = Math.max(1, Number(sys.setsQty) || 1)
    const secIdx = masterRows.length
    masterRows.push([`${sys.name} (${sys.bay} BAY) — ${setsQty} ${setsQty > 1 ? 'SETS' : 'SET'} (Tab: ${tcTab})`])
    for (let c = 0; c < 12; c++) masterStyles.push({ r: secIdx, c, s: subHeaderStyle })

    const activeModels = (sys.summary?.models || []).filter((m: any) => Number(m.qty) > 0)
    const modelCellMap: Record<string, string> = { DFMU: 'D4', SFMU: 'D8', SFFU: 'D12', DFFU: 'D16', 'DFMU-D': 'D20' }
    const qtyCellMap: Record<string, string> = { DFMU: 'G4', SFMU: 'G8', SFFU: 'G12', DFFU: 'G16', 'DFMU-D': 'G20' }

    const startModelRow = masterRows.length + 1
    activeModels.forEach((m: any, mIdx: number) => {
      const rIdx = masterRows.length
      const excelR = rIdx + 1
      const qtyPerSet = Number(m.qty) || 0
      const wExtra = Number(m.widthExtra) || 0
      const baseRate = Math.max(0, (Number(m.unitRate) || 0) - wExtra)
      const oneSetEx = (Number(m.unitRate) || 0) * qtyPerSet
      const allSetsEx = oneSetEx * setsQty
      const powderAll = (Number(m.powder) || 0) * setsQty

      masterRows.push([
        mIdx + 1,
        `${setsQty > 1 ? `1 of ${setsQty} Sets` : '1 Set'}`,
        m.code,
        baseRate,
        wExtra,
        qtyPerSet,
        setsQty,
        qtyPerSet * setsQty,
        oneSetEx,
        allSetsEx,
        powderAll,
        baseRate + wExtra + (sys.extraAmount || 0),
      ])

      const rateRef = modelCellMap[m.code]
      const totQtyRef = qtyCellMap[m.code]
      if (rateRef) {
        masterFormulas.push({ r: rIdx, c: 3, f: quoteRef(tcTab, rateRef), fmt: cellMoneyFmt })
      }
      if (totQtyRef) {
        masterFormulas.push({ r: rIdx, c: 7, f: `F${excelR}*G${excelR}` })
      } else {
        masterFormulas.push({ r: rIdx, c: 7, f: `F${excelR}*G${excelR}` })
      }
      masterFormulas.push(
        { r: rIdx, c: 8, f: `(D${excelR}+E${excelR})*F${excelR}`, fmt: cellMoneyFmt },
        { r: rIdx, c: 9, f: `I${excelR}*G${excelR}`, fmt: cellMoneyFmt },
      )
      for (let c = 0; c < 12; c++) {
        masterStyles.push({
          r: rIdx,
          c,
          s: {
            border: THIN_BORDER,
            alignment: { horizontal: c <= 2 ? 'center' : 'right', vertical: 'center' },
            ...(c >= 3 && c !== 5 && c !== 6 && c !== 7 ? { numFmt: cellMoneyFmt } : {}),
          },
        })
      }
    })

    const endModelRow = masterRows.length
    if (endModelRow >= startModelRow) {
      const subIdx = masterRows.length
      masterRows.push(['', `${sys.name} Subtotal (${setsQty} ${setsQty > 1 ? 'Sets' : 'Set'})`, '', '', '', '', setsQty, 0, 0, 0, 0, ''])
      masterFormulas.push(
        { r: subIdx, c: 7, f: `SUM(H${startModelRow}:H${endModelRow})` },
        { r: subIdx, c: 8, f: `SUM(I${startModelRow}:I${endModelRow})`, fmt: cellMoneyFmt },
        { r: subIdx, c: 9, f: `SUM(J${startModelRow}:J${endModelRow})`, fmt: cellMoneyFmt },
        { r: subIdx, c: 10, f: `SUM(K${startModelRow}:K${endModelRow})`, fmt: cellMoneyFmt },
      )
      for (let c = 0; c < 12; c++) {
        masterStyles.push({
          r: subIdx,
          c,
          s: {
            font: { bold: true, color: { rgb: '0F172A' }, sz: 10 },
            fill: { fgColor: { rgb: 'F8FAFC' } },
            border: THIN_BORDER,
          },
        })
      }
    }
  })

  const masterWs = XLSX.utils.aoa_to_sheet(masterRows)
  for (const mf of masterFormulas) {
    const cellRef = addr(mf.r, mf.c)
    if (masterWs[cellRef]) {
      masterWs[cellRef].f = mf.f
      if (mf.fmt) {
        masterWs[cellRef].s = { ...(masterWs[cellRef].s || {}), numFmt: mf.fmt }
      }
    }
  }
  for (const ms of masterStyles) {
    const cellRef = addr(ms.r, ms.c)
    if (masterWs[cellRef]) {
      masterWs[cellRef].s = { ...(masterWs[cellRef].s || {}), ...ms.s }
    }
  }
  masterWs['!cols'] = [
    { wch: 6 }, { wch: 24 }, { wch: 8 }, { wch: 10 }, { wch: 26 },
    { wch: 13 }, { wch: 13 }, { wch: 18 }, { wch: 15 }, { wch: 14 },
    { wch: 13 }, { wch: 13 }, { wch: 16 }, { wch: 10 }, { wch: 16 },
    { wch: 18 }, { wch: 16 }, { wch: 10 }, { wch: 15 }, { wch: 18 },
    { wch: 18 }, { wch: 15 }, { wch: 18 }, { wch: 14 }, { wch: 14 },
    { wch: 14 }, { wch: 15 }, { wch: 15 }, { wch: 14 },
  ]
  masterWs['!views'] = [{ showGridLines: true }]
  masterWs['!pageSetup'] = { paperSize: 9, orientation: 'landscape' }
  XLSX.utils.book_append_sheet(book, masterWs, 'Total cal')

  // 3. Append 3 sheets per system: "<System> - Total cal", "<System> - SHEARING SIZE", "<System> - HARDWARE LIST"
  const origOverrides = wb.getOverrides()
  const origFormulaOverrides = wb.getFormulaOverrides()

  try {
    systemSheetsMeta.forEach(({ sys, tcTab, shTab, hwTab }) => {
      wb.setOverrides(sys.overrides || {})
      wb.setFormulaOverrides(sys.formulaOverrides || {})

      const sheetRenameMap: Record<string, string> = { [tcSrcName]: tcTab }
      const allowedTargetSheets = new Set<string>([tcSrcName])
      if (shSrcName && shTab) {
        sheetRenameMap[shSrcName] = shTab
        allowedTargetSheets.add(shSrcName)
      }
      if (hwSrcName && hwTab) {
        sheetRenameMap[hwSrcName] = hwTab
        allowedTargetSheets.add(hwSrcName)
      }
      const crossSheetConfig: CrossSheetSubsetConfig = {
        sheetRenameMap,
        allowedTargetSheets,
      }

      const sysOpts: ExcelExportOptions = {
        setsQty: sys.setsQty,
        subtotalBeforeGst: sys.subtotalBeforeGst,
        sysTransport: sys.allocatedTransport,
        sysAccommodation: sys.allocatedAccommodation,
        isActual: globalOptions.isActual,
        profitPct: sys.profitPct,
        discountPct: sys.discountPct,
        includeInstallation: sys.includeInstallation,
        extraAmount: sys.extraAmount,
      }

      withScaledWorkbookOverrides(
        wb,
        sys.setsQty,
        sys.profitPct,
        sys.allocatedTransport,
        sys.allocatedAccommodation,
        () => {
          const tcWs = buildStyledWorksheet(wb, tcSrcName, sys.c2cData, sysOpts, crossSheetConfig)
          if (tcWs) XLSX.utils.book_append_sheet(book, tcWs, tcTab)

          if (shSrcName && shTab) {
            const shWs = buildStyledWorksheet(wb, shSrcName, sys.c2cData, sysOpts, crossSheetConfig)
            if (shWs) XLSX.utils.book_append_sheet(book, shWs, shTab)
          }

          if (hwSrcName && hwTab) {
            const hwWs = buildStyledWorksheet(wb, hwSrcName, sys.c2cData, sysOpts, crossSheetConfig)
            if (hwWs) XLSX.utils.book_append_sheet(book, hwWs, hwTab)
          }
        },
      )
    })
  } finally {
    wb.setOverrides(origOverrides)
    wb.setFormulaOverrides(origFormulaOverrides)
  }

  const cleanFileName = filename.endsWith('.xlsx') ? filename : `${filename}.xlsx`
  XLSX.writeFile(book, cleanFileName)
}
