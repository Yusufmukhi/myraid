import { Workbook } from './engine.ts'

/** Reads the "Total cal" sheet (the summary page of the Excel file) into a tidy object. */
export interface Summary {
  models: {
    code: string
    note?: string
    qty: number
    unitRate: number
    widthExtra?: number
    amount: number
    powder: number
  }[]
  subtotal: number // Ex-works
  installation: number
  rawInstallation?: number
  includeInstallation?: boolean
  transport: number
  accommodation: number
  accessories: number // alias for compatibility
  extraAmount?: number
  pf: number
  costBase: number // Ex-Works + Transport + Accommodation + Other Extra Amount + (optional Installation)
  profitPct: number // Default 25%
  profitAmount: number
  subtotalWithProfit: number
  discountPct: number // Default 0%
  discountAmount: number
  beforeGst: number // Subtotal before tax
  gst: number // 18% GST
  sales: number // Final Sales Total
  steel: { sp: number; spWaste: number; spTotal: number; spRate: number; spCost: number; spLabourRate: number; spLabour: number }
  trolley: { kg: number; waste: number; total: number; rate: number; cost: number; labourRate: number; labour: number }
  totalWeight: number
  weightWithMachine: number
  steelCost: number
  steelLabour: number
  fittingLabour: number
  labourSubtotal: number
  clipRate: number
  clipQty: number
  clips: number
  labourTotal: number
  powderTotal: number
  bodies: number
  compartments: number
  perBodyLabour: number
  bodyLabourTotal: number
  door: { perDoor: number; labourPerDoor: number; extraOverDfmu: number }
}

/**
 * Width extra charges (DFMU, SFMU, DFMU-D):
 * 3000 to 4700 mm: ₹2,500
 * 4701 to 6000 mm: ₹4,000
 * else: ₹0
 */
export function getWidthExtra(width: number, modelCode: string): number {
  const code = (modelCode || '').toUpperCase()
  const isApplicable =
    code === 'DFMU' ||
    code === 'SFMU' ||
    code === 'DFMU-D' ||
    code === 'DFFU-D' ||
    code.includes('DFMU') ||
    code.includes('SFMU')
  if (!isApplicable) return 0
  if (width >= 3000 && width <= 4700) {
    return 2500
  }
  if (width >= 4701 && width <= 6000) {
    return 4000
  }
  return 0
}

export function summarize(
  wb: Workbook,
  profitPct?: number,
  discountPct: number = 0,
  includeInstallation: boolean = true,
  extraAmount: number = 0,
): Summary {
  const n = (a: string) => wb.num('Total cal', a)
  const width = n('D3') || 1830
  const extraDfmu = getWidthExtra(width, 'DFMU')
  const extraSfmu = getWidthExtra(width, 'SFMU')
  const extraDfmuD = getWidthExtra(width, 'DFMU-D')

  const models = [
    {
      code: 'DFMU',
      qty: n('G4'),
      unitRate: n('D4') + extraDfmu,
      widthExtra: extraDfmu,
      amount: n('G4') * (n('D4') + extraDfmu),
      powder: n('E34'),
    },
    {
      code: 'SFMU',
      qty: n('G8'),
      unitRate: n('D8') + extraSfmu,
      widthExtra: extraSfmu,
      amount: n('G8') * (n('D8') + extraSfmu),
      powder: n('E35'),
    },
    {
      code: 'SFFU',
      qty: n('G12'),
      unitRate: n('D12'),
      widthExtra: 0,
      amount: n('N12'),
      powder: n('E36'),
    },
    {
      code: 'DFFU',
      qty: n('G16'),
      unitRate: n('D16'),
      widthExtra: 0,
      amount: n('N16'),
      powder: n('E37'),
    },
    {
      code: 'DFMU-D',
      note: 'with door',
      qty: n('G20'),
      unitRate: n('D20') + extraDfmuD,
      widthExtra: extraDfmuD,
      amount: n('G20') * (n('D20') + extraDfmuD),
      powder: n('E38'),
    },
  ]

  const subtotal = models.reduce((acc, m) => acc + m.amount, 0) // Ex-Works
  const rawInstallation = n('N23')
  const installation = includeInstallation ? rawInstallation : 0
  const transport = n('N24')
  const accommodation = n('N25')
  const accessories = accommodation // alias
  const pf = n('N26')

  const safeExtraAmount = isNaN(extraAmount) ? 0 : Math.max(0, extraAmount)
  const totalUnits = models.reduce((acc, m) => acc + m.qty, 0)
  const totalExtraAmountSum = safeExtraAmount * totalUnits

  // Direct Cost Base: Ex-Works + Transport + Accommodation + Other Extra Amount Total + (optional Installation)
  const costBase = subtotal + installation + transport + accommodation + totalExtraAmountSum

  // Profit Margin (Explicit parameter, or cell Total cal!C26, or default 25%)
  // In Excel row 26: N26 = N22 * Profit%, applied to Ex-Works Subtotal
  const sheetProfit = wb.num('Total cal', 'C26')
  const safeProfitPct = (profitPct !== undefined && !isNaN(profitPct))
    ? profitPct
    : (sheetProfit > 0 ? sheetProfit : 25)
  const profitAmount = subtotal * (safeProfitPct / 100)
  const subtotalWithProfit = subtotal + profitAmount

  // Discount Percentage applied on Subtotal with Profit (Before Tax)
  const safeDiscountPct = isNaN(discountPct) ? 0 : discountPct
  const discountAmount = subtotalWithProfit * (safeDiscountPct / 100)
  const beforeGst = (subtotalWithProfit - discountAmount) + installation + transport + accommodation + totalExtraAmountSum

  // 18% GST on Subtotal Before Tax
  const gst = beforeGst * 0.18

  // Final Total Quotation / Sales Value
  const sales = beforeGst + gst

  return {
    models,
    subtotal,
    installation,
    rawInstallation,
    includeInstallation,
    transport,
    accommodation,
    accessories,
    extraAmount: safeExtraAmount,
    pf,
    costBase,
    profitPct: safeProfitPct,
    profitAmount,
    subtotalWithProfit,
    discountPct: safeDiscountPct,
    discountAmount,
    beforeGst,
    gst,
    sales,
    steel: {
      sp: n('J34'),
      spWaste: n('J35'),
      spTotal: n('J36'),
      spRate: n('K36'),
      spCost: n('L36'),
      spLabourRate: n('M36'),
      spLabour: n('N36'),
    },
    trolley: {
      kg: n('J40'),
      waste: n('J41'),
      total: n('J42'),
      rate: n('K42'),
      cost: n('L42'),
      labourRate: n('M42'),
      labour: n('N42'),
    },
    totalWeight: n('J44'),
    weightWithMachine: n('J45'),
    steelCost: n('L44'),
    steelLabour: n('N44'),
    fittingLabour: n('N47'),
    labourSubtotal: n('N50'),
    clipQty: n('L52'),
    clipRate: n('M52'),
    clips: n('N52'),
    labourTotal: n('N54'),
    powderTotal: n('E39'),
    bodies: n('B48'),
    compartments: n('D48'),
    perBodyLabour: n('B49'),
    bodyLabourTotal: n('B50'),
    door: { perDoor: n('E24'), labourPerDoor: n('I24'), extraOverDfmu: n('B20') },
  }
}

/** Extracts template default inputs from a template definition or snapshot. */
export function extractTemplateDefaults(snapshotOrDef: any): Record<string, number | null> {
  if (!snapshotOrDef) return {}
  const def = snapshotOrDef.definition || snapshotOrDef
  const defaults: Record<string, number | null> = {}

  if (def.default_inputs && typeof def.default_inputs === 'object') {
    Object.assign(defaults, def.default_inputs)
  }
  if (def.rates && typeof def.rates === 'object') {
    if (def.rates.steel !== undefined && def.rates.steel !== null) {
      const s = Number(def.rates.steel)
      if (!isNaN(s)) {
        defaults['Total cal!H3'] = s
        defaults['rates!b3'] = s
      }
    }
    if (def.rates.trolleySteel !== undefined && def.rates.trolleySteel !== null) {
      const t = Number(def.rates.trolleySteel)
      if (!isNaN(t)) {
        defaults['Total cal!M3'] = t
        defaults['rates!b4'] = t
      }
    }
    if (def.rates.bodyLabour !== undefined && def.rates.bodyLabour !== null) {
      const l = Number(def.rates.bodyLabour)
      if (!isNaN(l)) defaults['Total cal!B49'] = l
    }
  }
  return defaults
}

/** Overrides that really differ from the workbook defaults (so we only store real changes). */
export function cleanOverrides(wb: Workbook, o: Record<string, number | null>) {
  const out: Record<string, number | null> = {}
  for (const [k, v] of Object.entries(o)) {
    // Preserve custom profit and discount percentages
    if (k.startsWith('__')) {
      if (v !== null && v !== undefined) out[k] = v
      continue
    }
    // Column count for non-master models is strictly driven by DFMU (Total cal!B3)
    if (k === 'Total cal!B7' || k === 'Total cal!B11' || k === 'Total cal!B15' || k === 'Total cal!B19') {
      continue
    }
    const [sheet, a] = splitKey(k)
    const m = /^([A-Z]+)(\d+)$/.exec(a)
    if (!m) continue
    const defNum = wb.defaultInput(sheet, parseInt(m[2], 10) - 1, colNum(m[1]))
    if (v === defNum) continue
    out[k] = v
  }
  return out
}
function colNum(s: string) {
  let n = 0
  for (const ch of s) n = n * 26 + ch.charCodeAt(0) - 64
  return n - 1
}
export function splitKey(k: string): [string, string] {
  const i = k.lastIndexOf('!')
  return [k.slice(0, i), k.slice(i + 1)]
}

/**
 * Filter overrides for Cut-to-Cut workbook.
 * ONLY physical specifications and quantities apply to both workbooks:
 * - Bay (ROW)
 * - Column (CLOUM)
 * - Height (HT)
 * - Width (W)
 * - Depth (D)
 * - Body Gauge (Body G)
 * - Shelf Gauge (Shelf G)
 * - Track
 * - Trolley Gauge (Trolley G)
 * - Quantities (QTY)
 *
 * NOTE: Stiffener (Total cal!J3) and all rate columns (Steel Rate Total cal!H3, Trolley Rate Total cal!M3, etc.)
 * ONLY reflect in the actual workbook and are strictly excluded from Cut-to-Cut so the C2C workbook
 * always retains its internal cut-to-cut rates and stiffener gauge.
 */
const C2C_PHYSICAL_CELLS = new Set([
  // Bay (Rows / 2)
  'Total cal!A3', 'Total cal!A7', 'Total cal!A11', 'Total cal!A15', 'Total cal!A19',
  // Column
  'Total cal!B3', 'Total cal!B7', 'Total cal!B11', 'Total cal!B15', 'Total cal!B19',
  // Height
  'Total cal!C3', 'Total cal!C7', 'Total cal!C11', 'Total cal!C15', 'Total cal!C19',
  // Width
  'Total cal!D3', 'Total cal!D7', 'Total cal!D11', 'Total cal!D15', 'Total cal!D19',
  // Depth
  'Total cal!E3', 'Total cal!E7', 'Total cal!E11', 'Total cal!E15', 'Total cal!E19',
  // Body gauge
  'Total cal!F3', 'Total cal!F7', 'Total cal!F11', 'Total cal!F15', 'Total cal!F19',
  // Shelf gauge
  'Total cal!G3', 'Total cal!G7', 'Total cal!G11', 'Total cal!G15', 'Total cal!G19',
  // Track
  'Total cal!I3', 'Total cal!I7', 'Total cal!I11', 'Total cal!I15', 'Total cal!I19',
  // Trolley gauge
  'Total cal!L3', 'Total cal!L7', 'Total cal!L11', 'Total cal!L15', 'Total cal!L19',
  // Model Quantities
  'Total cal!G4', 'Total cal!G8', 'Total cal!G12', 'Total cal!G16', 'Total cal!G20',
  // Model-level physical specs if passed directly
  'DFMU!A3', 'DFMU!B3', 'DFMU!C3', 'DFMU!D3', 'DFMU!E3', 'DFMU!F3', 'DFMU!G3', 'DFMU!I3', 'DFMU!L3',
  'SFMU!A3', 'SFMU!B3', 'SFMU!C3', 'SFMU!D3', 'SFMU!E3', 'SFMU!F3', 'SFMU!G3', 'SFMU!I3', 'SFMU!L3',
  'SFFU!A3', 'SFFU!B3', 'SFFU!C3', 'SFFU!D3', 'SFFU!E3', 'SFFU!F3', 'SFFU!G3', 'SFFU!I3', 'SFFU!L3',
  'DFFU!A3', 'DFFU!B3', 'DFFU!C3', 'DFFU!D3', 'DFFU!E3', 'DFFU!F3', 'DFFU!G3', 'DFFU!I3', 'DFFU!L3',
  'DFMU-D!A3', 'DFMU-D!B3', 'DFMU-D!C3', 'DFMU-D!D3', 'DFMU-D!E3', 'DFMU-D!F3', 'DFMU-D!G3', 'DFMU-D!I3', 'DFMU-D!L3',
])

export function filterCutToCutOverrides(overrides: Record<string, any>): Record<string, any> {
  const c2cOv: Record<string, any> = {}

  for (const [k, v] of Object.entries(overrides)) {
    const unquoted = k.replace(/'/g, '').replace(/\s+/g, ' ').trim()
    if (C2C_PHYSICAL_CELLS.has(unquoted)) {
      c2cOv[k] = v
    }
  }

  // Ensure Cut-to-Cut base steel rate is strictly preserved at 79
  c2cOv['Total cal!H3'] = 79
  // Ensure Cut-to-Cut trolley rate is strictly preserved at 69 (never overridden by Actual trolley rate)
  c2cOv['Total cal!M3'] = 69
  // Ensure Cut-to-Cut stiffener is strictly preserved at 0
  c2cOv['Total cal!J3'] = 0

  return c2cOv
}

