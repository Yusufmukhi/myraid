import jsPDF from 'jspdf'
import html2canvas from 'html2canvas'
import React, { useEffect, useMemo, useState } from 'react'
import { createPortal } from 'react-dom'
import { MYRIAD_LOGO_DATA_URL, MYRIAD_WATERMARK_DATA_URL } from '../../assets/companyAssets'
import { BusinessProfile, getBusinessProfile, fetchBusinessProfileFromSupabase } from '../../lib/businessProfile'
import { Project, api, tokenStore } from '../../lib/api'
import { inr } from '../../lib/format'
import { Workbook } from '../../lib/engine'
import { Summary, getWidthExtra } from '../../lib/model'
import { buildOfficialQtnNo, extractQuotationCode, getFiscalYear } from '../../lib/revision'
import {
  getSavedClients,
  getVisibleClientsForUser,
  saveClientToDirectory,
  saveClientToDirectoryAsync,
  deleteClientFromDirectory,
  fetchClientsFromSupabase,
  SavedClient,
  QuotationExtraField,
  ClientContactPerson,
  getVisibleFieldValues,
} from '../../lib/clientDirectory'
import { Icon, useUI, Spinner } from '../ui'

export type { QuotationExtraField, ClientContactPerson }

export interface QuotationTermItem {
  id: string
  title: string
  text: string
}

export interface QuotationClientInfo {
  client_id?: string
  qtn_no: string
  date: string
  buyer_name: string
  buyer_address: string
  buyer_gst: string
  delivery: string
  buyer_email: string
  buyer_contact: string
  kind_att: string
  subject: string
  contacts?: ClientContactPerson[]
  selected_contact_id?: string
  show_buyer_name?: boolean
  show_buyer_address?: boolean
  show_buyer_gst?: boolean
  show_kind_att?: boolean
  show_buyer_contact?: boolean
  show_buyer_email?: boolean
  show_subject?: boolean
  installation_terms_text?: string
  freight_terms?: string
  payment_terms?: string
  unloading_terms?: string
  packing_terms?: string
  color_terms?: string
  validity_terms?: string
  tax_terms?: string
  lodging_terms?: string
  warranty_terms?: string
  amc_terms?: string
  extra_fields?: QuotationExtraField[]
  terms_list?: QuotationTermItem[]
}


export interface CompactorTechnicalSpecs {
  mainTitle: string
  subTitle: string
  mobileRacksTitle: string
  mobileRacksIntro: string
  undercarriageTitle: string
  baseFrame: string
  tracks: string
  superStructureTitle: string
  superStructureIntro: string
  shelves: string
  verticalPanel: string
  middlePanel: string
  topPanel: string
  centrePartition: string
  endPanel: string
  centralLocking: string
  mechanicalTitle: string
  powderCoating: string
  indexArrangement: string
  shaft: string
  pedestalBearing: string
  wheel: string
  chain: string
  openableShutter: string
}

export interface PowderCoatingSpecs {
  mainTitle: string
  plantOverview: string
  machineryTitle: string
  pretreatmentPlant: string
  tankList: string[]
  compressorsA: string
  compressorsB: string
  powderCoating: string
  oven: string
  miscEquipments: string
  instructionsTitle: string
  pretreatmentSecTitle: string
  pretreatmentIntro: string
  degreasingTitle: string
  degreasingLines: string[]
  waterRinse1: string
  derustingTitle: string
  derustingLines: string[]
  waterRinse2: string
}

export interface PretreatmentQCSpecs {
  mainTitle: string
  activationTitle: string
  activationLines: string[]
  zincPhosphatingTitle: string
  zincPhosphatingLines: string[]
  waterRinse3: string
  passivationTitle: string
  passivationLines: string[]
  parametersTitle: string
  parametersList: string[]
  qcMainTitle: string
  qcArrivalTitle: string
  qcArrivalLines: string[]
  qcProductionTitle: string
  qcProductionLines: string[]
  qcFinalTitle: string
  qcFinalLines: string[]
  qcTestingTitle: string
  qcTestingLines: string[]
}

export interface QuotationProposalSpecs {
  page3: CompactorTechnicalSpecs
  page4: PowderCoatingSpecs
  page5: PretreatmentQCSpecs
}

export const DEFAULT_PROPOSAL_SPECS: QuotationProposalSpecs = {
  page3: {
    mainTitle: 'COMPACTOR TECHNICAL SPECIFICATION',
    subTitle: 'DELUXE MODEL',
    mobileRacksTitle: 'Technical Specification For Mobile Racks:',
    mobileRacksIntro: 'Manually operated & mechanically assisted Mobile Racks sliding smoothly & effortlessly on a pair of rails & providing maximum storage capacity by compaction.',
    undercarriageTitle: 'THE UNDERCARRIAGE:',
    baseFrame: 'Channel type Base Frame fabricated in one part, Each Wheel Is provided with TWO Support Channel, Made from 2.5 mm thick M.S. Sheet size of Channel 25mm x 100mm x 40mm.',
    tracks: "The guide rails shall be made out of 28mm Bright Bar placed on 0.8mm thick M.S. Plate with 6 Bends placed above the floor and are as per the international standard of safety. No's of Tracks as per Drawing.",
    superStructureTitle: 'SUPER STRUCTURE:',
    superStructureIntro: 'Super structure of the Mobile storage system will consist of Knock Down Type Panels (Not Slotted Angle Type). All the components such as shelves, front panel, end panel, vertical rear panel, central partition etc. will act as an integral member of the unit.',
    shelves: 'Adjustable Shelves made of 0.8 MM thick Prime Quality C.R.C.A. sheet having 10 folds, each design to carry a U.D.L. of 60 kgs. Per Level.',
    verticalPanel: 'Vertical Panel made out of 0.8 MM thick C.R.C.A. Prime Quality Mild Steel having 8 Folds (36X15X12MM) & designed carry full load of entire unit with Overlapping Provisions for Dust Resistant.',
    middlePanel: 'Made out of 0.8 MM thick C.R.C.A Prime Quality Steel having 16 Folds (36X15X12MM) & designed to carry full load of entire unit',
    topPanel: 'Made out of 0.8 mm thick C.R.C.A Prime Quality Steel having 8 Folds, (27MM X 12MM) with Overlapping provision for Dust Resistant.',
    centrePartition: 'The centre partition will comprise of full C.R.C.A. Prime Quality Mild Steel 0.8 mm Sheet',
    endPanel: 'Made out of 0.8 mm thick C.R.C.A Prime Quality Steel having 8 Folds (72 X 15 X 12).',
    centralLocking: 'System will be provided with individual rack locking as well as Centralized locking Arrangement on the last unit when all units are brought together. Central lock when operated gets locked with the floor & does not allow access to the unit.',
    mechanicalTitle: 'MECHANICAL & DRIVE ASSEMBLY:',
    powderCoating: 'Nine Tank Process Powder Coated, Structure Finishing 50 to 60 Microns.',
    indexArrangement: 'Acrylic card holders (2 Nos.) provided on the front panel for easy identification & retrieval.',
    shaft: '20 mm thick Full Length.',
    pedestalBearing: 'Entire mechanism on Pedestal Bearing, Each wheel provided with 2 Bearing Pedestal for Smooth movement) as per International Standard).',
    wheel: 'Made by Casting Material OD 125mm & Thickness 36mm.',
    chain: '0.5" Chain.',
    openableShutter: 'Made by 1 mm sheet with Shifter with high security sing key lock with up and down and site locking agreement with in-built handle.',
  },
  page4: {
    mainTitle: 'POWDER COATING SPECIFICATION',
    plantOverview: 'Our powder coating plant in vasai and our working area 5500 Sq fit and Tank capacity 8500 Litres. Tank size 10fit X 8fit X4fit. Oven size 16fit X 6.5fit X 8.25fit. Oven size 6.5fit X 6 Fit X 7.5Fit',
    machineryTitle: 'LIST OF MACHINERY & EQUIPMENTS FOR POWDER COATING',
    pretreatmentPlant: '7 Tank cold process Zinc Phosphate coating Dipping type process 8500 Ltr. Capacity',
    tankList: [
      '1. Degreasing',
      '5. Activation / Predip',
      '2. Water Rinse: Ambient',
      '6. Zinc Phosphating',
      '3. Derusting',
      '7. Water Rinse: Ambient',
      '4. Water Rinse: Ambient',
      '8. Passivation: V-CHEM PASSI',
    ],
    compressorsA: 'A) 2 No 7.5 KV / 400 Ltr. Capacity filter Regulators & Air dryer,',
    compressorsB: 'B) 2 No 10 KV / 600 Ltr. Capacity filter Regulators & Air dryer,',
    powderCoating: "4 No Gema System Recovery booth size 5' X 5' X 4' = 5 nos",
    oven: '1 No GAS oven with inside working dimension of Oven size a) 16fit X 6.5fit X 8.25fit. b) 6.5fit X 6 Fit X 7.5Fit c) 12fit X 6.5 Fit X 9.25 Fit',
    miscEquipments: 'For Laboratory of chemicals, phosphate Coating & Finished component.',
    instructionsTitle: 'METHODS OF INSTRUCTION (M.S. ITEMS ONLY)',
    pretreatmentSecTitle: '1. PRETREATMENT',
    pretreatmentIntro: 'Pretreatment process followed should be as per IS 3618 Class C.M.S. Material must be zinc phosphated. Under given process should be followed for pretreatment of M.S. material\n1.1 If heavy deposition of grease or oil is observed, prewiping with Turpentine or washing thinner is necessary followed by a day wipe with a cotton rag. The same cotton rag must not be used for longer time.',
    degreasingTitle: '1.1 Degreasing:',
    degreasingLines: [
      '• Bath Strength: 25%',
      '• Dip Time: 10-20 mins (Dip time can be Increased or reduce depending upon oil)',
      '• Minimum Pointage: 15-18 0.1N H2SO4 (can be exceed to 100-110 ml 0.1N H2SO4)',
      '• Add for each one point less and add 65 Ltr.',
    ],
    waterRinse1: 'A continuous overflowing water air agitation. Dip Time: 1 min. pH 5-7',
    derustingTitle: '1.3 Derusting:',
    derustingLines: [
      '• Bath Strength: 25%',
      '• Dip Time: 10-20 mins (Dip time can be Increased or reduce depending upon oil / black scale quantity)',
      '• Minimum Pointage: 50-60 0.1N H2SO4 (can be exceed to 100-110 ml 0.1N H2SO4)',
      '• Add for each one point less and add 65 Ltr.',
    ],
    waterRinse2: 'A continuous overflowing water air agitation. Dip Time: 1 min. pH 5-7',
  },
  page5: {
    mainTitle: 'PRETREATMENT PROCESS & PARAMETERS',
    activationTitle: '1.5 Activation: V- CHEM S.M.',
    activationLines: [
      '• Bath Strength: 0.5% | Temp: R.T. | Dip Time: 1 min. | pH: 2-3',
    ],
    zincPhosphatingTitle: '1.6 Zinc Phosphating: V-CHEM ZP (COLD) LP',
    zincPhosphatingLines: [
      '• Bath Strength: 5-6% | Temp: R.T. | Dip Time: 15-20 min.',
      '• T.A.: 25-30 NaOH (can be exceed to 30-34 ml 0.1N NaOH)',
      '• Accelerator: Accelerator | Test for Accelerator: Starch Iodide paper turns light blue.',
    ],
    waterRinse3: 'A continuous overflowing water air agitation. Dip Time: 1 min. pH 5-7',
    passivationTitle: '1.8 Passivation: V-CHEM PASSI',
    passivationLines: [
      '• Bath Strength: 0.1% | Temp: R.T. | Dip Time: 1 min',
    ],
    parametersTitle: 'Parameters to be controlled on the line of pretreatment:',
    parametersList: [
      '1. Bath strength pointage and replenishment.',
      '4. Sludge.',
      '2. Dip time.',
      '5. Carry over of chemical.',
      '3. pH of water rinse, activation and passivation.',
      '6. Air pocketing.',
    ],
    qcMainTitle: 'QUALITY CONTROL',
    qcArrivalTitle: 'I. ON ARRIVAL OF MATERIAL',
    qcArrivalLines: [
      'a) After arrival of material it is stocked and stored properly in order to prevent it from damages.',
      'b) Before processing the material is checked for dent marks, deep scratches and other fabrication defect.',
    ],
    qcProductionTitle: 'II. DURING PRODUCTION',
    qcProductionLines: [
      'a) While in process the material is checked visually for removal of oil, grease, rust, scale etc.',
      'b) The pretreatment baths are maintained as per the requirement and the pointages pH, Temperatures and checked and recorded too.',
      'c) After phosphating, and before painting powder coating the material is again checked for phosphate coating etc.',
      'd) While baking the temperature controller is per the requirement in order to prevent under baking and over baking problems. A digital timer is provided to ensure that the oven goes off after the set time period.',
    ],
    qcFinalTitle: 'III. FINAL CONTROL',
    qcFinalLines: [
      'a) After baking the material is re-inspected for covering pattern, film thickness, gloss etc.',
      'b) Finally the material is touch-up in the delicate corners and stacked properly for dispatch.',
    ],
    qcTestingTitle: 'IV. TESTING EQUIPMENTS',
    qcTestingLines: [
      'a) Pretreatment baths are checked on a regular interval and pointage, pH temperatures are maintained as per the requirement.',
      'b) The phosphate coating is checked as per the specification IS 3618 for salt spray test.',
      'c) Painted component are checked for film thickness',
    ],
  },
}

export interface QuotationLineItem {
  id: string
  model_code?: string
  model_name: string
  dimensions: string
  quantity: number
  unit_base_rate?: number
  extra_charges?: number
  total_unit_rate?: number
  rate: number
  amount: number
}

export interface UnitWiseCalculationRow {
  model_code: string
  model_name: string
  dimensions: string
  qty: number
  rows_per_unit: number
  total_rows: number
  unit_rate: number
  transportation: number
  extra_amount?: number
  total_unit_rate: number
  total_price: number
  profit_pct: number
  total_price_with_profit: number
  unit_rate_with_profit: number
  discount_pct: number
  discount_amount: number
  final_amount: number
  final_unit_rate: number
}

export interface UnitWiseBreakdownResult {
  rows: UnitWiseCalculationRow[]
  totalExtraCharges: number
  transport: number
  installation: number
  accommodation: number
  extraAmount?: number
  totalBodies: number
  ratePerBody: number
  totalCostBase: number
  profitPct: number
  profitAmount: number
  totalPriceWithProfit: number
  discountPct: number
  discountAmount: number
  finalBeforeGst: number
  gst18: number
  grandTotal: number
}

/**
 * Computes exact unit-wise rate breakdown matching the user's Excel model:
 * 1. Total Extra Charges = Transport + Installation (if included) + Accommodation
 * 2. Total Bodies = sum of (qty * rowsPerUnit) across all active models
 * 3. Extra Charge per Body = Total Extra Charges / Total Bodies
 * 4. Model Extra Charge = rowsPerUnit * Extra Charge per Body
 * 5. Total Unit Rate = Unit Rate + Extra Charge per Unit + Extra Amount
 * 6. Total Price = Qty * Total Unit Rate
 * 7. Price With Profit = Total Price * (1 + profitPct / 100)
 */
export function computeUnitWiseBreakdown(
  summary: Summary,
  wb: Workbook | null | undefined,
  includeInstallation: boolean = true,
  customProfitPct?: number,
  customDiscountPct?: number,
  customTotalBodies?: number,
  customExtraAmount?: number,
): UnitWiseBreakdownResult {
  const transport = summary.transport || 0
  const accommodation = summary.accommodation || 0
  const extraAmount = customExtraAmount !== undefined && !isNaN(customExtraAmount) ? customExtraAmount : 0

  const activeModels = (summary?.models || []).filter((m) => m.qty > 0)

  let dfmuRows = 6
  if (wb) {
    dfmuRows = wb.num('Total cal', 'A3') || 6
  }

  const rowDefs = activeModels.map((m) => {
    let rowsPerUnit = 6
    if (wb) {
      if (m.code === 'DFMU') rowsPerUnit = wb.num('Total cal', 'A3') || 6
      else if (m.code === 'SFMU') rowsPerUnit = wb.num('Total cal', 'A7') || ((wb.num('Total cal', 'A3') || 6) / 2)
      else if (m.code === 'SFFU') rowsPerUnit = wb.num('Total cal', 'A11') || ((wb.num('Total cal', 'A3') || 6) / 2)
      else if (m.code === 'DFFU') rowsPerUnit = wb.num('Total cal', 'A15') || wb.num('Total cal', 'A3') || 6
      else if (m.code === 'DFMU-D' || m.code === 'DFFU-D') rowsPerUnit = wb.num('Total cal', 'A19') || wb.num('Total cal', 'A3') || 6
    } else {
      const isSingle = m.code === 'SFMU' || m.code === 'SFFU'
      rowsPerUnit = isSingle ? Math.max(1, Math.round(dfmuRows / 2)) : dfmuRows
    }
    return {
      model: m,
      rowsPerUnit: rowsPerUnit > 0 ? rowsPerUnit : (m.code === 'SFMU' || m.code === 'SFFU' ? 3 : 6),
    }
  })

  const calculatedTotalBodies = rowDefs.reduce((acc, r) => acc + (r.model.qty * r.rowsPerUnit), 0)
  const totalBodies = (customTotalBodies && customTotalBodies > 0)
    ? customTotalBodies
    : ((summary.bodies && summary.bodies > 0) ? summary.bodies : (calculatedTotalBodies > 0 ? calculatedTotalBodies : 1))

  // Labour per body rate: used to scale installation proportionally across combined multi-system bodies
  const labourRatePerBody = (summary.bodies && summary.bodies > 0)
    ? ((summary.installation || 0) / summary.bodies)
    : (wb ? (wb.num('Total cal', 'B49') || 250) : 250)
  const installation = includeInstallation
    ? ((customTotalBodies && customTotalBodies > 0) ? (labourRatePerBody * totalBodies) : (summary.installation || 0))
    : 0

  const totalExtraCharges = transport + accommodation + installation
  const ratePerBody = totalBodies > 0 ? (totalExtraCharges / totalBodies) : 0

  const profitPct = customProfitPct !== undefined && !isNaN(customProfitPct)
    ? customProfitPct
    : (summary.profitPct !== undefined && !isNaN(summary.profitPct) ? summary.profitPct : 25)
  const discountPct = customDiscountPct !== undefined && !isNaN(customDiscountPct)
    ? customDiscountPct
    : (summary.discountPct !== undefined && !isNaN(summary.discountPct) ? summary.discountPct : 0)

  let totalCostBase = 0
  let totalPriceWithProfit = 0
  let finalBeforeGst = 0

  const isProfitOnExWorksOnly = summary && summary.profitAmount !== undefined && summary.subtotal !== undefined
    ? Math.abs(summary.profitAmount - (summary.subtotal * (profitPct / 100))) < 1
    : false

  const rows: UnitWiseCalculationRow[] = rowDefs.map(({ model: m, rowsPerUnit }) => {
    let fullName = m.code
    let defaultDepth = 900
    if (m.code === 'DFMU') {
      fullName = 'DOUBLE FACE MOVABLE UNIT (DFMU)'
      defaultDepth = 900
    } else if (m.code === 'SFMU') {
      fullName = 'SINGLE FACE MOVABLE UNIT (SFMU)'
      defaultDepth = 475
    } else if (m.code === 'SFFU') {
      fullName = 'SINGLE FACE FIX UNIT (SFFU)'
      defaultDepth = 475
    } else if (m.code === 'DFFU') {
      fullName = 'DOUBLE FACE FIX UNIT (DFFU)'
      defaultDepth = 900
    } else if (m.code === 'DFMU-D' || m.code === 'DFFU-D') {
      fullName = 'DOUBLE FACE MOVABLE UNIT WITH DOOR (DFMU-D)'
      defaultDepth = 900
    }

    const rawHt = wb ? (wb.num(m.code, 'C3') || wb.num('Total cal', 'C3') || 3000) : 3000
    const ht = rawHt >= 500 ? rawHt : 3000
    const rawW = wb ? (wb.num(m.code, 'D3') || wb.num('Total cal', 'D3') || 1830) : 1830
    const w = rawW >= 500 ? rawW : 1830
    const rawD = wb ? (wb.num('Total cal', m.code === 'SFMU' ? 'E7' : m.code === 'SFFU' ? 'E11' : (m.code === 'DFMU-D' || m.code === 'DFFU-D') ? 'E19' : (m.code === 'DFFU' ? 'E15' : 'E3')) || defaultDepth) : defaultDepth
    const d = rawD >= 200 ? rawD : defaultDepth
    const cols = wb ? (wb.num('Total cal', 'B3') || 7) : 7

    const dimensions = `${ht} (HT) X ${w} (W)X ${d} (D) MM ${cols} Compartments`

    const widthExtra = (m as any).widthExtra !== undefined ? 0 : (wb ? getWidthExtra(w, m.code) : 0)
    const unit_rate = m.unitRate + widthExtra
    const transportation = rowsPerUnit * ratePerBody
    const logisticsPerUnit = transportation + extraAmount
    const total_unit_rate = unit_rate + logisticsPerUnit
    const total_price = m.qty * total_unit_rate

    // Profit margin applied on Direct Cost Base (total_price = unit_rate + logisticsPerUnit)
    const itemProfit = total_price * (profitPct / 100)
    const itemPriceWithProfit = total_price + itemProfit
    const unit_rate_with_profit = itemPriceWithProfit / m.qty
    const itemDiscount = itemPriceWithProfit * (discountPct / 100)
    const final_amount = itemPriceWithProfit - itemDiscount
    const final_unit_rate = final_amount / m.qty

    totalCostBase += total_price
    totalPriceWithProfit += itemPriceWithProfit
    finalBeforeGst += final_amount

    return {
      model_code: m.code,
      model_name: fullName,
      dimensions,
      qty: m.qty,
      rows_per_unit: rowsPerUnit,
      total_rows: m.qty * rowsPerUnit,
      unit_rate,
      transportation,
      extra_amount: extraAmount,
      total_unit_rate,
      total_price,
      profit_pct: profitPct,
      total_price_with_profit: itemPriceWithProfit,
      unit_rate_with_profit,
      discount_pct: discountPct,
      discount_amount: itemDiscount,
      final_amount,
      final_unit_rate,
    }
  })

  // Adjust sub-cent rounding against authoritative summary.beforeGst when present
  if (summary?.beforeGst && rows.length > 0) {
    const roundingDiff = Number((summary.beforeGst - finalBeforeGst).toFixed(2))
    if (Math.abs(roundingDiff) > 0 && Math.abs(roundingDiff) <= 0.08) {
      rows[rows.length - 1].final_amount = Number((rows[rows.length - 1].final_amount + roundingDiff).toFixed(2))
      rows[rows.length - 1].final_unit_rate = rows[rows.length - 1].final_amount / rows[rows.length - 1].qty
      finalBeforeGst = summary.beforeGst
      totalPriceWithProfit = Number((totalPriceWithProfit + roundingDiff).toFixed(2))
    }
  }

  const profitAmount = totalCostBase * (profitPct / 100)
  const discountAmount = totalPriceWithProfit * (discountPct / 100)
  const gst18 = finalBeforeGst * 0.18
  const grandTotal = Math.round(finalBeforeGst + gst18)

  return {
    rows,
    totalExtraCharges,
    transport,
    installation,
    accommodation,
    extraAmount,
    totalBodies,
    ratePerBody,
    totalCostBase,
    profitPct,
    profitAmount,
    totalPriceWithProfit,
    discountPct,
    discountAmount,
    finalBeforeGst,
    gst18,
    grandTotal,
  }
}

export interface SetUnitItem {
  id: string
  model_name: string
  dimensions: string
  unit_qty: number
  model_code?: string
  rate?: number
  amount?: number
}

export interface QuotationSet {
  id: string
  system_title?: string // e.g. "SYSTEM A", "SYSTEM B", "SYSTEM C"
  set_label: string // e.g. "SET 1", "SYSTEM A", "1 SET"
  quantity: number // e.g. 1
  unit: string // "SET"
  rate: number
  amount: number
  pre_discount_rate?: number
  pre_discount_amount?: number
  items: SetUnitItem[]
  summary?: any
  breakdown?: any
}

export interface InstallationBodyItem {
  id: string
  title: string
  sub_label: string
  per_bay_qty: number
  qty: number
  total_body: number
  system_name?: string
  system_id?: string
  set_qty?: number
  unit_body?: number
}

/**
 * Extracts width (mm) from dimensions string (e.g. "2200 (HT) X 1830 (W) X 610 (D)").
 */
export function extractWidthFromDimensions(dimensions?: string): number | null {
  if (!dimensions) return null
  const m = dimensions.match(/(\d+)\s*\([Ww]\)/)
  if (m) return parseInt(m[1], 10)
  return null
}

function anyMovableModel(name: string): boolean {
  const upper = (name || '').toUpperCase()
  return (
    upper.includes('DFMU') ||
    upper.includes('SFMU') ||
    upper.includes('MOVABLE') ||
    upper.includes('MOBILE')
  ) && !upper.includes('FIX') && !upper.includes('SFFU') && !upper.includes('DFFU')
}

/**
 * Strips DOUBLE DRIVE from dimensions string so it does NOT appear in the table cell.
 * Per user instruction: DOUBLE DRIVE WITH REDCUTION WHEEL is displayed ONLY in the red specifications
 * box (where POWDER COATED is), never in the table cells.
 */
export function cleanDimensions(dimensions?: string): string {
  if (!dimensions) return ''
  return dimensions
    .split('\n')
    .map((l) => l.trim())
    .filter((l) => Boolean(l) && !l.toUpperCase().includes('DOUBLE DRIVE'))
    .join(' ')
    .replace(/\bColumns\b/gi, 'Compartments')
    .replace(/\bColumn\b/gi, 'Compartment')
}

export function getDisplayModelName(it: { model_code?: string; model_name?: string }): string {
  if (it.model_code && it.model_code.trim()) {
    return it.model_code.trim()
  }
  const raw = (it.model_name || '').trim()
  const m = raw.match(/\(([A-Za-z0-9\-]+)\)/)
  if (m && m[1]) {
    return m[1].toUpperCase()
  }
  return raw
}

/**
 * Generates default red specification notes for the goods table on Page 1.
 * - Width >= 3000 mm: includes "DOUBLE DRIVE WITH REDUCTION WHEEL" (user deduction/reduction default)
 * - Track count: includes "{trackCount} TRACKS" (from Total cal!I3 or default 2 TRACKS)
 * - Standard lines: Superstructure thickness, Trolley thickness, Powder coated, HSN code.
 */
export function generateDefaultRedSpecs(
  projectWidth = 1830,
  trackCount = 2,
  wb?: Workbook | null,
  inputs?: Record<string, any>
): string[] {
  let superThickness = '0.8 MM'
  if (wb) {
    const th = wb.get('Total cal', 'F3')
    if (th !== null && th !== undefined && String(th).trim()) {
      const s = String(th).trim()
      superThickness = s.toUpperCase().includes('MM') ? s.toUpperCase() : `${s} MM`
    }
  } else if (inputs?.['Total cal!F3']) {
    const th = String(inputs['Total cal!F3']).trim()
    if (th) superThickness = th.toUpperCase().includes('MM') ? th.toUpperCase() : `${th} MM`
  }

  let trolleyThickness = '2.5 MM'
  if (wb) {
    const th = wb.get('Total cal', 'L3')
    if (th !== null && th !== undefined && String(th).trim()) {
      const s = String(th).trim()
      trolleyThickness = s.toUpperCase().includes('MM') ? s.toUpperCase() : `${s} MM`
    }
  } else if (inputs?.['Total cal!L3']) {
    const th = String(inputs['Total cal!L3']).trim()
    if (th) trolleyThickness = th.toUpperCase().includes('MM') ? th.toUpperCase() : `${th} MM`
  }

  const specs: string[] = []

  // If width >= 3000 mm, include DOUBLE DRIVE WITH REDUCTION WHEEL
  if (projectWidth >= 3000) {
    specs.push('DOUBLE DRIVE WITH REDUCTION WHEEL')
  }

  // Track count: e.g. "2 TRACKS", "3 TRACKS", etc. (only included if trackCount > 0)
  if (trackCount && trackCount > 0) {
    specs.push(`${trackCount} ${trackCount === 1 ? 'TRACK' : 'TRACKS'}`)
  }

  // Superstructure thickness
  specs.push(`${superThickness} THICKNESS SUPERSTRUCTURE`)

  // Trolley thickness
  specs.push(`${trolleyThickness} TROLLY`)

  // Powder coated & HSN
  specs.push('POWDER COATED')
  specs.push('HSN CODE: 9403')

  return specs
}

/**
 * Computes installation bodies dynamically:
 * - bay count = Math.max(1, Math.round(dfmuRows / 2))
 * - per_bay_qty = model's row count (e.g. DFMU=6, SFMU=3, SFFU=3, DFFU=6, DFMU-D=6)
 * - qty = model's quantity
 * - total_body = per_bay_qty * qty
 * - Shows all active models with qty > 0 (including DFFU and DFMU-D / DFFU-D)
 */
export function generateDefaultBodyItems(
  summary: Summary,
  wb?: Workbook | null,
  inputs?: Record<string, any>,
  sets?: QuotationSet[],
): InstallationBodyItem[] {
  if (sets && sets.length > 0) {
    const allItems: InstallationBodyItem[] = []
    let counter = 1
    sets.forEach((s, sIdx) => {
      const nextChar = String.fromCharCode(65 + sIdx)
      const sysName = (s.system_title || (s.set_label?.startsWith('SYSTEM') ? s.set_label : `SYSTEM ${nextChar}`)).trim()
      const sysQty = Math.max(1, Number(s.quantity) || 1)
      const sysSummary = s.summary || summary
      let dfmuRows = 6
      if (s.breakdown?.rows?.[0]?.rows_per_unit) {
        dfmuRows = s.breakdown.rows[0].rows_per_unit
      } else if (wb) {
        dfmuRows = wb.num('Total cal', 'A3') || 6
      } else if (inputs?.['Total cal!A3']) {
        dfmuRows = Number(inputs['Total cal!A3']) || 6
      }
      const bayCount = Math.max(1, Math.round(dfmuRows / 2))
      // Prioritize items directly from this system so each system's bodies match its commercial rows
      if (s.items && s.items.length > 0) {
        s.items.forEach((it: any) => {
          const code = (it.model_code || (
            it.model_name?.includes('DFMU-D') ? 'DFMU-D' :
            it.model_name?.includes('DFFU-D') ? 'DFFU-D' :
            it.model_name?.includes('DFMU') ? 'DFMU' :
            it.model_name?.includes('SFMU') ? 'SFMU' :
            it.model_name?.includes('SFFU') ? 'SFFU' :
            it.model_name?.includes('DFFU') ? 'DFFU' : 'DFMU'
          )).trim()

          const isSingle = code === 'SFMU' || code === 'SFFU'
          let rowsPerUnit = isSingle ? bayCount : (bayCount * 2)
          if (it.rows_per_unit) rowsPerUnit = it.rows_per_unit
          const unitCount = Number(it.unit_qty !== undefined ? it.unit_qty : it.quantity) || 0
          const unitBody = rowsPerUnit * unitCount
          const total_body = unitBody * sysQty

          allItems.push({
            id: `${s.id || sIdx}-${code}-${counter++}`,
            title: `${code} ${bayCount} BAY`,
            sub_label: `BODY IN ${code}`,
            per_bay_qty: rowsPerUnit,
            qty: unitCount,
            total_body,
            system_name: sysName,
            system_id: s.id || String(sIdx),
            set_qty: sysQty,
            unit_body: unitBody,
          })
        })
      } else {
        const activeModels = (sysSummary?.models || []).filter((m: any) => m.qty > 0)
        const modelsToUse = activeModels.length > 0 ? activeModels : sysSummary?.models || []

        modelsToUse.forEach((m: any) => {
        let rowsPerUnit = 6
        if (m.code === 'DFMU') rowsPerUnit = dfmuRows
        else if (m.code === 'SFMU' || m.code === 'SFFU') rowsPerUnit = Math.max(1, Math.round(dfmuRows / 2))
        else rowsPerUnit = dfmuRows

        const unitCount = Number(m.qty) || 0
        const unitBody = rowsPerUnit * unitCount
        const total_body = unitBody * sysQty
        allItems.push({
          id: `${s.id || sIdx}-${m.code}-${counter++}`,
          title: `${m.code} ${bayCount} BAY`,
          sub_label: `BODY IN ${m.code}`,
          per_bay_qty: rowsPerUnit,
          qty: unitCount,
          total_body,
          system_name: sysName,
          system_id: s.id || String(sIdx),
          set_qty: sysQty,
          unit_body: unitBody,
        })
      })
      }
    })
    if (allItems.length > 0) return allItems
  }

  let dfmuRows = 6
  if (wb) {
    dfmuRows = wb.num('Total cal', 'A3') || 6
  } else if (inputs?.['Total cal!A3']) {
    dfmuRows = Number(inputs['Total cal!A3']) || 6
  }
  const bayCount = Math.max(1, Math.round(dfmuRows / 2))

  const activeModels = (summary?.models || []).filter((m) => m.qty > 0)
  const modelsToUse = activeModels.length > 0 ? activeModels : [
    { code: 'DFMU', qty: 8 },
    { code: 'SFMU', qty: 2 },
    { code: 'SFFU', qty: 2 },
  ]

  return modelsToUse.map((m, idx) => {
    let rowsPerUnit = 6
    if (wb) {
      if (m.code === 'DFMU') rowsPerUnit = wb.num('Total cal', 'A3') || (bayCount * 2)
      else if (m.code === 'SFMU') rowsPerUnit = wb.num('Total cal', 'A7') || ((wb.num('Total cal', 'A3') || (bayCount * 2)) / 2) || bayCount
      else if (m.code === 'SFFU') rowsPerUnit = wb.num('Total cal', 'A11') || ((wb.num('Total cal', 'A3') || (bayCount * 2)) / 2) || bayCount
      else if (m.code === 'DFFU') rowsPerUnit = wb.num('Total cal', 'A15') || wb.num('Total cal', 'A3') || (bayCount * 2)
      else if (m.code === 'DFMU-D' || m.code === 'DFFU-D') rowsPerUnit = wb.num('Total cal', 'A19') || wb.num('Total cal', 'A3') || (bayCount * 2)
    } else {
      const isSingle = m.code === 'SFMU' || m.code === 'SFFU'
      rowsPerUnit = isSingle ? bayCount : (bayCount * 2)
    }

    const codeLabel = m.code
    const title = `${codeLabel} ${bayCount} BAY`
    const sub_label = `BODY IN ${codeLabel}`
    const qty = m.qty
    const unitBody = rowsPerUnit * qty

    return {
      id: String(idx + 1),
      title,
      sub_label,
      per_bay_qty: rowsPerUnit,
      qty,
      total_body: unitBody,
      system_name: 'SYSTEM A',
      system_id: 'default',
      set_qty: 1,
      unit_body: unitBody,
    }
  })
}

interface Props {
  project: Project
  summary: Summary
  wb?: Workbook
  role?: string
  isAdminOrOwner?: boolean
  orgId?: string
  username?: string
  userId?: string
  initialQuotationMode?: 'set_wise' | 'unit_wise' | 'unit_breakdown'
  initialExtraAmount?: number
  initialSets?: QuotationSet[]
  totalCombinedBodies?: number
  isUnsavedWarning?: boolean
  onClose: () => void
  onSaveClientInfo?: (
    clientInfo: QuotationClientInfo,
    items: QuotationLineItem[],
    bodyItems: InstallationBodyItem[],
    sets?: QuotationSet[],
    extraSettings?: Record<string, any>
  ) => void
}

export function numberToIndianWords(num: number): string {
  const n = Math.round(Math.abs(num))
  if (n === 0) return 'Zero'
  const units = [
    '', 'One', 'Two', 'Three', 'Four', 'Five', 'Six', 'Seven', 'Eight', 'Nine', 'Ten',
    'Eleven', 'Twelve', 'Thirteen', 'Fourteen', 'Fifteen', 'Sixteen', 'Seventeen', 'Eighteen', 'Nineteen'
  ]
  const tens = ['', '', 'Twenty', 'Thirty', 'Forty', 'Fifty', 'Sixty', 'Seventy', 'Eighty', 'Ninety']

  function convertBelow1000(val: number): string {
    let res = ''
    if (val >= 100) {
      res += units[Math.floor(val / 100)] + ' Hundred '
      val %= 100
    }
    if (val >= 20) {
      res += tens[Math.floor(val / 10)] + ' '
      val %= 10
    }
    if (val > 0) {
      res += units[val] + ' '
    }
    return res.trim()
  }

  let rem = n
  const crore = Math.floor(rem / 10000000)
  rem %= 10000000
  const lakh = Math.floor(rem / 100000)
  rem %= 100000
  const thousand = Math.floor(rem / 1000)
  rem %= 1000

  const parts: string[] = []
  if (crore > 0) parts.push(convertBelow1000(crore) + ' Crore')
  if (lakh > 0) parts.push(convertBelow1000(lakh) + ' Lakh')
  if (thousand > 0) parts.push(convertBelow1000(thousand) + ' Thousand')
  if (rem > 0) parts.push(convertBelow1000(rem))

  return parts.join(' ').replace(/\s+/g, ' ').trim()
}

export function OfficialQuotationDocument({
  project,
  summary,
  wb,
  role,
  isAdminOrOwner,
  orgId,
  username,
  userId,
  initialQuotationMode,
  initialExtraAmount,
  initialSets,
  totalCombinedBodies: propTotalCombinedBodies,
  isUnsavedWarning,
  onClose,
  onSaveClientInfo,
}: Props) {
  const userIsAdmin = isAdminOrOwner ?? (role === 'owner' || role === 'admin')
  const effectiveOrgId = orgId || project.org_id

  // 1. Initial Client Fields
  const initialClientInfo = useMemo<QuotationClientInfo>(() => {
    const saved = project.summary?.client_info || {}
    const today = new Date().toLocaleDateString('en-GB', { day: '2-digit', month: '2-digit', year: 'numeric' }).replace(/\//g, '-')
    let dfmuRows = 6
    if (wb) {
      dfmuRows = wb.num('Total cal', 'A3') || 6
    } else if (project.inputs?.['Total cal!A3']) {
      dfmuRows = Number(project.inputs['Total cal!A3']) || 6
    }
    const bayCount = Math.max(1, Math.round(dfmuRows / 2))
    const defaultSubject = `QUOTATION FOR SUPPLY ${bayCount} BAY MECHANICAL MOBILE COMPACTOR`

    // Generate standard official QTN NO based on project name and fiscal year:
    // e.g. "QTY NO 123" -> MSSLP/26-27/MT-SL/123
    // e.g. "QTY NO 123-A" -> MSSLP/26-27/MT-SL/123-A
    const generatedQtnNo = buildOfficialQtnNo(project.name, saved.date || today)
    const isOldDummy = saved.qtn_no === 'MSSLP/26-27/MT-SL/3167-A'
    const projectCode = extractQuotationCode(project.name)
    const savedCode = saved.qtn_no ? extractQuotationCode(saved.qtn_no) : ''
    const isMismatchedRevision = Boolean(projectCode && savedCode && projectCode !== savedCode)
    const effectiveQtnNo = (!saved.qtn_no || isOldDummy || isMismatchedRevision) ? generatedQtnNo : saved.qtn_no

    const effectiveBuyerName = (saved.buyer_name && saved.buyer_name !== 'Patil Group') ? saved.buyer_name : ''
    const dirMatch = effectiveBuyerName
      ? getSavedClients(effectiveOrgId).find(
          (c) => c.name.trim().toLowerCase() === effectiveBuyerName.trim().toLowerCase()
        )
      : undefined

    const isOldHardcodedAddress = saved.buyer_address?.includes('Yellampet Village, Medchal')
    const resolvedAddress = (!isOldHardcodedAddress && saved.buyer_address) ? saved.buyer_address : (dirMatch?.address || '')
    const isOldHardcodedGst = saved.buyer_gst === '36AABCP1234F1Z5'
    const resolvedGst = (!isOldHardcodedGst && saved.buyer_gst) ? saved.buyer_gst : (dirMatch?.gst || '')
    const isOldHardcodedEmail = saved.buyer_email === 'storesmedchal@patilgroup.com'
    const resolvedEmail = (!isOldHardcodedEmail && saved.buyer_email) ? saved.buyer_email : (dirMatch?.email || '')
    const isOldHardcodedKindAtt = saved.kind_att === 'Mr. G.Rajesh'
    const resolvedKindAtt = (!isOldHardcodedKindAtt && saved.kind_att) ? saved.kind_att : (dirMatch?.kind_att || '')
    const isOldHardcodedContact = saved.buyer_contact === '+91 7396822424'
    const resolvedContact = (!isOldHardcodedContact && saved.buyer_contact) ? saved.buyer_contact : (dirMatch?.contact || '')
    const resolvedContacts = (Array.isArray(saved.contacts) && saved.contacts.length > 0)
      ? saved.contacts
      : (Array.isArray(dirMatch?.contacts) ? dirMatch.contacts : [])
    const resolvedExtra = (Array.isArray(saved.extra_fields) && saved.extra_fields.length > 0)
      ? saved.extra_fields
      : (Array.isArray(dirMatch?.extra_fields) ? dirMatch.extra_fields : [])

    // Check if transportation is 0 in summary, inputs or wb
    const transportVal = (summary?.transport !== undefined && !isNaN(Number(summary.transport)))
      ? Number(summary.transport)
      : (project.summary?.transport !== undefined && !isNaN(Number(project.summary.transport))
        ? Number(project.summary.transport)
        : (project.inputs?.['Total cal!N24'] !== undefined && !isNaN(Number(project.inputs['Total cal!N24']))
          ? Number(project.inputs['Total cal!N24'])
          : (project.inputs?.['total cal!n24'] !== undefined && !isNaN(Number(project.inputs['total cal!n24']))
            ? Number(project.inputs['total cal!n24'])
            : (wb ? (wb.num('Total cal', 'N24') || 0) : 0))))
    const isTransportZero = isNaN(transportVal) || transportVal <= 0
    const accommodationVal = (summary?.accommodation !== undefined && !isNaN(Number(summary.accommodation)))
      ? Number(summary.accommodation)
      : (project.summary?.accommodation !== undefined && !isNaN(Number(project.summary.accommodation))
        ? Number(project.summary.accommodation)
        : (project.inputs?.['Total cal!N25'] !== undefined && !isNaN(Number(project.inputs['Total cal!N25']))
          ? Number(project.inputs['Total cal!N25'])
          : (project.inputs?.['total cal!n25'] !== undefined && !isNaN(Number(project.inputs['total cal!n25']))
            ? Number(project.inputs['total cal!n25'])
            : (wb ? (wb.num('Total cal', 'N25') || 0) : 0))))
    const isAccommodationZero = isNaN(accommodationVal) || accommodationVal <= 0
    const defaultFreightTerms = isTransportZero ? 'At your scope.' : 'Including.'
    const defaultLodgingTerms = 'At your scope.'

    return {
      client_id: saved.client_id || dirMatch?.id || '',
      qtn_no: effectiveQtnNo,
      date: saved.date || today,
      buyer_name: effectiveBuyerName,
      buyer_address: resolvedAddress,
      buyer_gst: resolvedGst,
      delivery: saved.delivery || '4 to 5 WEEKS',
      buyer_email: resolvedEmail,
      buyer_contact: resolvedContact,
      kind_att: resolvedKindAtt,
      subject: (!saved.subject || saved.subject.includes('5 BAY')) ? defaultSubject : saved.subject,
      contacts: resolvedContacts,
      selected_contact_id: saved.selected_contact_id || '',
      show_buyer_name: saved.show_buyer_name !== false,
      show_buyer_address: saved.show_buyer_address !== false,
      show_buyer_gst: saved.show_buyer_gst !== false,
      show_kind_att: saved.show_kind_att !== false,
      show_buyer_contact: saved.show_buyer_contact !== false,
      show_buyer_email: saved.show_buyer_email !== false,
      show_subject: saved.show_subject !== false,
      extra_fields: resolvedExtra.map((ef: any) => {
        const val = ef.value || (Array.isArray(ef.values) && ef.values[0]) || ''
        return {
          id: ef.id || String(Date.now()),
          label: ef.label || '',
          value: val,
          values: [val],
          visible: ef.visible !== false,
          visible_values: [true],
        }
      }),
      installation_terms_text: (project.summary?.include_installation === false || project.inputs?.['__include_installation'] === 0)
        ? 'At your scope.'
        : (saved.installation_terms_text && !saved.installation_terms_text.includes('Excluded') ? saved.installation_terms_text : 'Including.'),
      freight_terms: isTransportZero ? 'At your scope.' : (saved.freight_terms && saved.freight_terms !== 'At your scope.' ? saved.freight_terms : defaultFreightTerms),
      payment_terms: saved.payment_terms || '50%Advance payment along with the purchase order,50% payment before material dispatching.',
      unloading_terms: saved.unloading_terms || 'At your scope(any mathadi union unloading issue to be managed by buyer & it will not be our responsibility).',
      packing_terms: saved.packing_terms || 'No extra charges for packing of materail.',
      color_terms: saved.color_terms || 'RAL 7032 STR, RAL 7035 STR, LIGHT GREY STR, D.A.GREY STR or AVAILABLE SHADES.',
      validity_terms: saved.validity_terms || 'Quotation Validity 15 days.',
      tax_terms: saved.tax_terms || 'Within Maharashtra 9% CGST+9% SGST/outside Maharashtra 18% IGST Extra to your Account.',
      lodging_terms: (saved.lodging_terms && saved.lodging_terms !== 'Including.') ? saved.lodging_terms : 'At your scope.',
      warranty_terms: saved.warranty_terms || '12 months Warrenty Manufacturing defect.',
      amc_terms: saved.amc_terms || `Annual Monthly Contract from " ${getBusinessProfile(effectiveOrgId).company_name || 'MYRIAD'}" will be starting after completion of 1 years Warranty. Till 1 years any kind of servicing will be done Free to Client`,
    }
  }, [project, summary, wb])

  const { toast } = useUI()
  const [clientInfo, setClientInfo] = useState<QuotationClientInfo>(initialClientInfo)
  const [activeSidebarTab, setActiveSidebarTab] = useState<'options' | 'breakdown' | 'client' | 'terms' | 'body' | 'specs'>('options')
  const [specsSubTab, setSpecsSubTab] = useState<'page3' | 'page4' | 'page5'>('page3')
  const [justSaved, setJustSaved] = useState(false)
  const [proposalSpecs, setProposalSpecs] = useState<QuotationProposalSpecs>(() => {
    const saved = project.summary?.proposal_specs
    if (saved && saved.page3 && saved.page4 && saved.page5) {
      return {
        page3: { ...DEFAULT_PROPOSAL_SPECS.page3, ...saved.page3 },
        page4: { ...DEFAULT_PROPOSAL_SPECS.page4, ...saved.page4 },
        page5: { ...DEFAULT_PROPOSAL_SPECS.page5, ...saved.page5 },
      }
    }
    return DEFAULT_PROPOSAL_SPECS
  })
  const [sidebarCollapsed, setSidebarCollapsed] = useState(false)

  // Business Profile & PDF Branding State (dynamic organisation details)
  const [businessProfile, setBusinessProfile] = useState<BusinessProfile>(() => getBusinessProfile(effectiveOrgId))

  useEffect(() => {
    setBusinessProfile(getBusinessProfile(effectiveOrgId))
    fetchBusinessProfileFromSupabase(undefined, effectiveOrgId)
    fetchClientsFromSupabase(undefined, effectiveOrgId)
  }, [effectiveOrgId])

  useEffect(() => {
    const handleProfileUpdated = (e: Event) => {
      const detail = (e as CustomEvent)?.detail
      if (detail?.profile) {
        setBusinessProfile(detail.profile)
      } else {
        setBusinessProfile(getBusinessProfile(effectiveOrgId))
      }
    }
    if (typeof window !== 'undefined') {
      window.addEventListener('examcalc:business-profile-updated', handleProfileUpdated)
      return () => window.removeEventListener('examcalc:business-profile-updated', handleProfileUpdated)
    }
  }, [effectiveOrgId])

  // Client Directory State (Available to Staff, Admin, and Owner with creator isolation)
  const [savedClients, setSavedClients] = useState<SavedClient[]>(() => getSavedClients(effectiveOrgId))
  const visibleSavedClients = useMemo(() => {
    return getVisibleClientsForUser(savedClients, username, role, userId)
  }, [savedClients, username, role, userId])

  const [selectedClientId, setSelectedClientId] = useState<string>(() => {
    if (initialClientInfo.client_id) return initialClientInfo.client_id
    const matching = getVisibleClientsForUser(getSavedClients(effectiveOrgId), username, role, userId).find(
      (c) => c.name.trim().toLowerCase() === (initialClientInfo.buyer_name || '').trim().toLowerCase(),
    )
    return matching?.id || ''
  })
  const [clientSearchText, setClientSearchText] = useState('')
  const filteredSavedClients = useMemo(() => {
    const q = clientSearchText.trim().toLowerCase()
    if (!q) return visibleSavedClients || []
    return (visibleSavedClients || []).filter(
      (c) =>
        c.name.toLowerCase().includes(q) ||
        (c.gst || '').toLowerCase().includes(q) ||
        (c.kind_att || '').toLowerCase().includes(q) ||
        (c.contact || '').toLowerCase().includes(q)
    )
  }, [visibleSavedClients, clientSearchText])
  const [showAddClientModal, setShowAddClientModal] = useState<boolean>(false)
  const [newClientData, setNewClientData] = useState<{
    name: string
    address: string
    gst: string
    kind_att: string
    contact: string
    email: string
  }>({
    name: '',
    address: '',
    gst: '',
    kind_att: '',
    contact: '',
    email: '',
  })

  // Available contact persons for quick selection by Name dropdown
  const contactPersonOptions = useMemo(() => {
    const list: { id: string; name: string; phone: string; email?: string; label: string }[] = []
    const seen = new Set<string>()

    const addOption = (id: string, name?: string, phone?: string, tag?: string, email?: string) => {
      const cleanName = (name || '').trim()
      const cleanPhone = (phone || '').trim()
      const cleanEmail = (email || '').trim()
      if (!cleanName && !cleanPhone && !cleanEmail) return
      const key = `${cleanName.toLowerCase()}|${cleanPhone}|${cleanEmail.toLowerCase()}`
      if (seen.has(key)) return
      seen.add(key)
      const label = `${cleanName}${cleanPhone ? ` · ${cleanPhone}` : ''}${cleanEmail ? ` · ${cleanEmail}` : ''}${tag ? ` (${tag})` : ''}`
      list.push({ id, name: cleanName, phone: cleanPhone, email: cleanEmail, label })
    }

    // 1. Saved client from directory (primary contact)
    const savedCl = visibleSavedClients.find((c) => c.id === selectedClientId)
    if (savedCl && (savedCl.kind_att?.trim() || savedCl.contact?.trim())) {
      addOption(`saved_primary_${savedCl.id}`, savedCl.kind_att, savedCl.contact, 'Primary Contact', savedCl.email)
    }

    // 2. Current quotation primary contact
    if (clientInfo.kind_att?.trim() || clientInfo.buyer_contact?.trim()) {
      addOption('primary', clientInfo.kind_att, clientInfo.buyer_contact, 'Primary Contact', clientInfo.buyer_email)
    }

    // 3. Saved client additional contacts from directory
    if (savedCl && Array.isArray(savedCl.contacts)) {
      for (const cp of savedCl.contacts) {
        addOption(`saved_cp_${cp.id}`, cp.name, cp.phone, undefined, cp.email)
      }
    }

    // 4. Current quotation additional contacts
    for (const cp of (clientInfo.contacts || [])) {
      addOption(cp.id || `cp_${cp.name}`, cp.name, cp.phone, undefined, cp.email)
    }

    return list
  }, [clientInfo.kind_att, clientInfo.buyer_contact, clientInfo.contacts, selectedClientId, visibleSavedClients])

  // Keep saved clients synchronized across components/tabs and update active quotation state immediately
  useEffect(() => {
    const handleUpdate = () => {
      const freshList = getSavedClients(effectiveOrgId)
      setSavedClients(freshList)

      // When a client is updated in the directory, sync changes directly into clientInfo so PDF immediately shows them!
      const matching = selectedClientId
        ? freshList.find((c) => c.id === selectedClientId)
        : freshList.find((c) => c.name.trim().toLowerCase() === (clientInfo.buyer_name || '').trim().toLowerCase())

      if (matching) {
        setClientInfo((prev) => {
          const loadedExtra = Array.isArray(matching.extra_fields)
            ? matching.extra_fields.map((ef) => {
                const val = ef.value || (Array.isArray(ef.values) && ef.values[0]) || ''
                return {
                  id: ef.id || String(Date.now()),
                  label: ef.label || '',
                  value: val,
                  values: [val],
                  visible: ef.visible !== false,
                  visible_values: [true],
                }
              })
            : prev.extra_fields

          const loadedContacts = Array.isArray(matching.contacts) ? matching.contacts : prev.contacts

          return {
            ...prev,
            buyer_name: matching.name,
            buyer_address: matching.address || prev.buyer_address,
            buyer_gst: matching.gst || prev.buyer_gst,
            buyer_email: matching.email !== undefined ? matching.email : prev.buyer_email,
            show_buyer_email: Boolean(matching.email || prev.buyer_email),
            kind_att: matching.kind_att || prev.kind_att,
            buyer_contact: matching.contact || prev.buyer_contact,
            contacts: loadedContacts,
            extra_fields: loadedExtra,
          }
        })
      }
    }
    window.addEventListener('examcalc:clients-updated', handleUpdate)
    return () => window.removeEventListener('examcalc:clients-updated', handleUpdate)
  }, [effectiveOrgId, selectedClientId, clientInfo.buyer_name])

  const handleSelectSavedClient = (clientId: string) => {
    setSelectedClientId(clientId)
    if (!clientId) return
    const cl = savedClients.find((c) => c.id === clientId)
    if (!cl) return

    setClientInfo((prev) => {
      const loadedExtra = Array.isArray(cl.extra_fields) && cl.extra_fields.length > 0
        ? cl.extra_fields.map((ef) => {
            const val = ef.value || (Array.isArray(ef.values) && ef.values[0]) || ''
            return {
              id: ef.id || String(Date.now()),
              label: ef.label || '',
              value: val,
              values: [val],
              visible: ef.visible !== false,
              visible_values: [true],
            }
          })
        : []

      const loadedContacts = Array.isArray(cl.contacts)
        ? cl.contacts.filter((cp) => !(cp.name === cl.kind_att && cp.phone === cl.contact && cp.name.trim().length > 0))
        : []

      return {
        ...prev,
        buyer_name: cl.name,
        buyer_address: cl.address,
        buyer_gst: cl.gst || '',
        buyer_email: cl.email !== undefined ? cl.email : '',
        show_buyer_email: Boolean(cl.email),
        kind_att: cl.kind_att || '',
        buyer_contact: cl.contact || '',
        contacts: loadedContacts,
        selected_contact_id: '',
        extra_fields: loadedExtra,
      }
    })

    if (cl.profit_margin_pct !== undefined) {
      setProfitPct(cl.profit_margin_pct)
    }

    toast(`Loaded details for ${cl.name}`)
  }

  const handleSaveCurrentAsClient = async () => {
    if (!clientInfo.buyer_name || !clientInfo.buyer_name.trim()) {
      toast('Please enter a Buyer / Company Name first.', 'err')
      return
    }

    try {
      const normName = clientInfo.buyer_name.trim().toLowerCase()
      const existingClient = savedClients.find(
        (c) => (selectedClientId && c.id === selectedClientId) || (c.name && c.name.trim().toLowerCase() === normName)
      )
      const targetId = selectedClientId || clientInfo.client_id || existingClient?.id

      const saved = await saveClientToDirectoryAsync(
        {
          id: targetId,
          name: clientInfo.buyer_name.trim(),
          address: clientInfo.buyer_address || '',
          gst: clientInfo.buyer_gst || '',
          kind_att: clientInfo.kind_att || '',
          contact: clientInfo.buyer_contact || '',
          email: clientInfo.buyer_email || '',
          subject: clientInfo.subject || '',
          contacts: clientInfo.contacts || [],
          extra_fields: clientInfo.extra_fields || [],
        },
        effectiveOrgId,
        username || 'staff',
        role,
      )

      setSelectedClientId(saved.id)
      setClientInfo((prev) => ({ ...prev, client_id: saved.id }))
      setSavedClients(getSavedClients(effectiveOrgId))
      toast(targetId ? `Updated "${saved.name}" in Client Directory!` : `Saved "${saved.name}" to Client Directory!`)
    } catch (err: any) {
      toast(err?.message || 'Client already exists', 'err')
    }
  }

  const handleUpdateSelectedClient = async () => {
    if (!selectedClientId) {
      await handleSaveCurrentAsClient()
      return
    }

    try {
      const saved = await saveClientToDirectoryAsync(
        {
          id: selectedClientId,
          name: clientInfo.buyer_name.trim(),
          address: clientInfo.buyer_address || '',
          gst: clientInfo.buyer_gst || '',
          kind_att: clientInfo.kind_att || '',
          contact: clientInfo.buyer_contact || '',
          email: clientInfo.buyer_email || '',
          subject: clientInfo.subject || '',
          contacts: clientInfo.contacts || [],
          extra_fields: clientInfo.extra_fields || [],
        },
        effectiveOrgId,
        username || 'staff',
        role,
      )

      setSavedClients(getSavedClients(effectiveOrgId))
      setClientInfo((prev) => ({
        ...prev,
        client_id: saved.id,
        buyer_email: saved.email || prev.buyer_email,
        show_buyer_email: Boolean(saved.email || prev.buyer_email),
      }))
      toast(`Updated "${saved.name}" in Client Directory!`)
    } catch (err: any) {
      toast(err?.message || 'Client already exists', 'err')
    }
  }

  const handleDeleteSelectedClient = () => {
    if (!selectedClientId) return
    const cl = savedClients.find((c) => c.id === selectedClientId)
    if (!cl) return
    if (window.confirm(`Remove "${cl.name}" from saved Client Directory?`)) {
      deleteClientFromDirectory(selectedClientId, effectiveOrgId)
      setSelectedClientId('')
      toast(`Client "${cl.name}" removed from directory.`)
    }
  }

  const handleCreateNewClient = async () => {
    if (!newClientData.name.trim()) {
      toast('Buyer / Company Name is required.', 'err')
      return
    }

    const normName = newClientData.name.trim().toLowerCase()
    const existingClient = savedClients.find(
      (c) => c.name && c.name.trim().toLowerCase() === normName
    )
    const targetId = existingClient?.id

    const initialContactList = (newClientData.kind_att || newClientData.contact) ? [{
      id: `cp_${Date.now()}_1`,
      name: newClientData.kind_att || '',
      phone: newClientData.contact || '',
      email: newClientData.email || '',
    }] : []

    try {
      const saved = await saveClientToDirectoryAsync(
        {
          id: targetId,
          ...newClientData,
          contacts: initialContactList,
        },
        effectiveOrgId,
        username || 'staff',
        role,
      )

      const initialContacts = Array.isArray(saved.contacts) && saved.contacts.length > 0
        ? saved.contacts
        : (saved.kind_att ? [{ id: 'cp_1', name: saved.kind_att, phone: saved.contact || '', email: saved.email || '' }] : [])

      setClientInfo((prev) => ({
        ...prev,
        client_id: saved.id,
        buyer_name: saved.name,
        buyer_address: saved.address,
        buyer_gst: saved.gst || '',
        kind_att: saved.kind_att || '',
        buyer_contact: saved.contact || '',
        buyer_email: saved.email || '',
        contacts: initialContacts,
        selected_contact_id: initialContacts[0]?.id || '',
      }))

      setSelectedClientId(saved.id)
      setSavedClients(getSavedClients(effectiveOrgId))
      setShowAddClientModal(false)
      setNewClientData({
        name: '',
        address: '',
        gst: '',
        kind_att: '',
        contact: '',
        email: '',
      })

      toast(targetId ? `Existing client "${saved.name}" updated and applied!` : `New client "${saved.name}" saved to directory and applied!`)
    } catch (err: any) {
      toast(err?.message || 'Client already exists', 'err')
    }
  }

  // Dynamic profit percentage (derived from project input or client setting or summary, default 25%)
  const defaultProfitPct = useMemo(() => {
    if (project.inputs?.['__profit_pct'] !== undefined) {
      const v = Number(project.inputs['__profit_pct'])
      if (!isNaN(v)) return v
    }
    const matchingClient = visibleSavedClients.find((c) => c.id === selectedClientId)
    if (matchingClient && matchingClient.profit_margin_pct !== undefined && !isNaN(matchingClient.profit_margin_pct)) {
      return matchingClient.profit_margin_pct
    }
    if (summary.profitPct !== undefined && !isNaN(summary.profitPct)) {
      return summary.profitPct
    }
    return 25
  }, [project, summary, visibleSavedClients, selectedClientId])

  const [profitPct, setProfitPct] = useState<number>(defaultProfitPct)

  useEffect(() => {
    if (defaultProfitPct !== undefined && !isNaN(defaultProfitPct)) {
      setProfitPct(defaultProfitPct)
    }
  }, [defaultProfitPct])

  // Installation Included / Excluded state (defaults to calculator settings)
  const defaultIncludeInstallation = useMemo(() => {
    if (project.summary?.include_installation !== undefined) {
      return Boolean(project.summary.include_installation)
    }
    if (project.inputs?.['__include_installation'] !== undefined) {
      return Boolean(Number(project.inputs['__include_installation']))
    }
    if (summary.includeInstallation !== undefined) {
      return Boolean(summary.includeInstallation)
    }
    return (summary.installation ?? 0) > 0
  }, [project, summary])

  const [includeInstallation, setIncludeInstallation] = useState<boolean>(defaultIncludeInstallation)

  // Check if transportation is 0 in calculator
  const currentTransportVal = (summary?.transport !== undefined && !isNaN(Number(summary.transport)))
    ? Number(summary.transport)
    : (project.summary?.transport !== undefined && !isNaN(Number(project.summary.transport))
      ? Number(project.summary.transport)
      : (project.inputs?.['Total cal!N24'] !== undefined && !isNaN(Number(project.inputs['Total cal!N24']))
        ? Number(project.inputs['Total cal!N24'])
        : (project.inputs?.['total cal!n24'] !== undefined && !isNaN(Number(project.inputs['total cal!n24']))
          ? Number(project.inputs['total cal!n24'])
          : (wb ? (wb.num('Total cal', 'N24') || 0) : 0))))
  const isTransportZero = isNaN(currentTransportVal) || currentTransportVal <= 0

  const currentAccommodationVal = (summary?.accommodation !== undefined && !isNaN(Number(summary.accommodation)))
    ? Number(summary.accommodation)
    : (project.summary?.accommodation !== undefined && !isNaN(Number(project.summary.accommodation))
      ? Number(project.summary.accommodation)
      : (project.inputs?.['Total cal!N25'] !== undefined && !isNaN(Number(project.inputs['Total cal!N25']))
        ? Number(project.inputs['Total cal!N25'])
        : (project.inputs?.['total cal!n25'] !== undefined && !isNaN(Number(project.inputs['total cal!n25']))
          ? Number(project.inputs['total cal!n25'])
          : (wb ? (wb.num('Total cal', 'N25') || 0) : 0))))
  const isAccommodationZero = isNaN(currentAccommodationVal) || currentAccommodationVal <= 0

  const defaultFreightTerms = isTransportZero ? 'At your scope.' : 'Including.'
  const defaultLodgingTerms = 'At your scope.'

  // 1b. Terms & Conditions Initial List (all clauses fully editable, with add/delete row support)
  const initialTermsList = useMemo<QuotationTermItem[]>(() => {
    if (project.summary?.terms_list && Array.isArray(project.summary.terms_list) && project.summary.terms_list.length > 0) {
      return project.summary.terms_list.map((t: QuotationTermItem) => {
        // If transportation is 0 in calculator, automatically force freight to 'At your scope.'
        if (t.id === '8' || t.title.toLowerCase().includes('freight') || t.title.toLowerCase().includes('transport')) {
          if (isTransportZero) {
            return { ...t, text: 'At your scope.' }
          } else if (!t.text || t.text === 'At your scope.') {
            return { ...t, text: 'Including.' }
          }
        }
        // Clause 10 default is 'At your scope.'
        if (t.id === '10' || t.title.toLowerCase().includes('lodging') || t.title.toLowerCase().includes('loadging') || t.title.toLowerCase().includes('bording') || t.title.toLowerCase().includes('boarding')) {
          if (!t.text || t.text === 'Including.' || isAccommodationZero) {
            return { ...t, text: 'At your scope.' }
          }
        }
        return t
      })
    }
    const c = initialClientInfo
    return [
      { id: '1', title: 'Payment :-', text: c.payment_terms || '50%Advance payment along with the purchase order,50% payment before material dispatching.' },
      { id: '2', title: 'Delivery :-', text: c.delivery ? `Order of material delivery within ${c.delivery} from the date of receipt of this P.O.` : 'Order of material delivery within 4 to 5 weeks from the date of receipt of this P.O.' },
      { id: '3', title: 'Unloading Materil :-', text: c.unloading_terms || 'At your scope(any mathadi union unloading issue to be managed by buyer & it will not be our responsibility).' },
      { id: '4', title: 'Packing Chrg :-', text: c.packing_terms || 'No extra charges for packing of materail.' },
      { id: '5', title: 'Colour :', text: c.color_terms || 'RAL 7032 STR, RAL 7035 STR, LIGHT GREY STR, D.A.GREY STR or AVAILABLE SHADES.' },
      { id: '6', title: 'Validity :', text: c.validity_terms || 'Quotation Validity 15 days.' },
      { id: '7', title: 'Taxes:-', text: c.tax_terms || 'Within Maharashtra 9% CGST+9% SGST/outside Maharashtra 18% IGST Extra to your Account.' },
      { id: '8', title: 'Freight Chrg :-', text: isTransportZero ? 'At your scope.' : (c.freight_terms?.trim() || defaultFreightTerms) },
      { id: '9', title: 'Installation Chrg :', text: c.installation_terms_text?.trim() || (defaultIncludeInstallation ? 'Including.' : 'At your scope.') },
      { id: '10', title: 'Loadging and bording:', text: c.lodging_terms?.trim() || 'At your scope.' },
      { id: '11', title: 'Warranty :', text: c.warranty_terms || '12 months Warrenty Manufacturing defect.' },
      { id: '12', title: 'AMC :', text: c.amc_terms || `Annual Monthly Contract from " ${businessProfile.company_name || 'MYRIAD'}" will be starting after completion of 1 years Warranty. Till 1 years any kind of servicing will be done Free to Client` },
    ]
  }, [project, initialClientInfo, defaultIncludeInstallation, isTransportZero, isAccommodationZero, defaultFreightTerms, defaultLodgingTerms, businessProfile.company_name])

  const [termsList, setTermsList] = useState<QuotationTermItem[]>(initialTermsList)

  // Synchronize termsList with clientInfo whenever commercial options are edited or selected
  useEffect(() => {
    setTermsList((prev) => {
      let changed = false
      const next = prev.map((item) => {
        if (item.id === '1' || item.title.toLowerCase().includes('payment')) {
          if (clientInfo.payment_terms && item.text !== clientInfo.payment_terms) {
            changed = true
            return { ...item, text: clientInfo.payment_terms }
          }
        }
        if (item.id === '2' || item.title.toLowerCase().includes('delivery')) {
          const expected = clientInfo.delivery ? `Order of material delivery within ${clientInfo.delivery} from the date of receipt of this P.O.` : item.text
          if (clientInfo.delivery && item.text !== expected) {
            changed = true
            return { ...item, text: expected }
          }
        }
        if (item.id === '3' || item.title.toLowerCase().includes('unloading')) {
          if (clientInfo.unloading_terms && item.text !== clientInfo.unloading_terms) {
            changed = true
            return { ...item, text: clientInfo.unloading_terms }
          }
        }
        if (item.id === '4' || item.title.toLowerCase().includes('packing')) {
          if (clientInfo.packing_terms && item.text !== clientInfo.packing_terms) {
            changed = true
            return { ...item, text: clientInfo.packing_terms }
          }
        }
        if (item.id === '5' || item.title.toLowerCase().includes('colour') || item.title.toLowerCase().includes('color')) {
          if (clientInfo.color_terms && item.text !== clientInfo.color_terms) {
            changed = true
            return { ...item, text: clientInfo.color_terms }
          }
        }
        if (item.id === '6' || item.title.toLowerCase().includes('validity')) {
          if (clientInfo.validity_terms && item.text !== clientInfo.validity_terms) {
            changed = true
            return { ...item, text: clientInfo.validity_terms }
          }
        }
        if (item.id === '7' || item.title.toLowerCase().includes('tax')) {
          if (clientInfo.tax_terms && item.text !== clientInfo.tax_terms) {
            changed = true
            return { ...item, text: clientInfo.tax_terms }
          }
        }
        if (item.id === '8' || item.title.toLowerCase().includes('freight') || item.title.toLowerCase().includes('transport')) {
          const target = isTransportZero ? 'At your scope.' : (clientInfo.freight_terms || 'Including.')
          if (item.text !== target) {
            changed = true
            return { ...item, text: target }
          }
        }
        if (item.id === '9' || item.title.toLowerCase().includes('installation')) {
          const instText = includeInstallation
            ? (clientInfo.installation_terms_text?.trim() || 'Including.')
            : 'At your scope.'
          if (instText && item.text !== instText) {
            changed = true
            return { ...item, text: instText }
          }
        }
        if (item.id === '10' || item.title.toLowerCase().includes('loadging') || item.title.toLowerCase().includes('lodging') || item.title.toLowerCase().includes('bording') || item.title.toLowerCase().includes('boarding')) {
          const target = clientInfo.lodging_terms || 'At your scope.'
          if (item.text !== target) {
            changed = true
            return { ...item, text: target }
          }
        }
        if (item.id === '11' || item.title.toLowerCase().includes('warranty')) {
          if (clientInfo.warranty_terms && item.text !== clientInfo.warranty_terms) {
            changed = true
            return { ...item, text: clientInfo.warranty_terms }
          }
        }
        if (item.id === '12' || item.title.toLowerCase().includes('amc')) {
          if (clientInfo.amc_terms && item.text !== clientInfo.amc_terms) {
            changed = true
            return { ...item, text: clientInfo.amc_terms }
          }
        }
        return item
      })
      return changed ? next : prev
    })
  }, [
    clientInfo.payment_terms,
    clientInfo.delivery,
    clientInfo.unloading_terms,
    clientInfo.packing_terms,
    clientInfo.color_terms,
    clientInfo.validity_terms,
    clientInfo.tax_terms,
    clientInfo.freight_terms,
    clientInfo.installation_terms_text,
    clientInfo.lodging_terms,
    clientInfo.warranty_terms,
    clientInfo.amc_terms,
    includeInstallation,
    isTransportZero,
    isAccommodationZero,
  ])

  // Quotation presentation mode: 'set_wise' or 'unit_wise' (staff can only use these two, admin can also use unit_breakdown)
  const [quotationMode, setQuotationMode] = useState<'set_wise' | 'unit_wise' | 'unit_breakdown'>(() => {
    const saved = initialQuotationMode || project.summary?.quotation_mode
    if (!userIsAdmin && saved === 'unit_breakdown') return 'set_wise'
    return saved || 'set_wise'
  })

  // Enforce staff never viewing unit breakdown or profit margins in their PDF
  const effectiveQuotationMode = (!userIsAdmin && quotationMode === 'unit_breakdown') ? 'set_wise' : quotationMode

  // System count check: true if project has multiple systems
  const isMultiSystem = useMemo(() => {
    if (initialSets && initialSets.length > 1) return true
    if (project.summary?.quotation_sets && project.summary.quotation_sets.length > 1) return true
    return false
  }, [initialSets, project.summary?.quotation_sets])

  // 1c. Installation Body Items (per bay qty = model rows, bay = dfmuRows / 2)
  const defaultBodyItems = useMemo<InstallationBodyItem[]>(() => {
    if (!isMultiSystem && project.summary?.body_items && Array.isArray(project.summary.body_items) && project.summary.body_items.length > 0) {
      const saved = project.summary.body_items
      const hasSystemPrefix = saved.some((it: any) => it.title?.includes('SYSTEM'))
      const isOldHardcodedDummy = saved.some((it: any) => it.total_body === 90 && it.title?.includes('5 BAY'))
      if (!hasSystemPrefix && !isOldHardcodedDummy) {
        return saved
      }
    }
    return generateDefaultBodyItems(summary, wb, project.inputs, (initialSets && initialSets.length > 0) ? initialSets : undefined)
  }, [project, summary, wb, initialSets, isMultiSystem])

  const [bodyItems, setBodyItems] = useState<InstallationBodyItem[]>(defaultBodyItems)

  const totalBodies = useMemo(() => {
    return bodyItems.reduce((acc, it) => acc + (Number(it.total_body) || 0), 0)
  }, [bodyItems])

  // Total Combined Bodies across all systems and sets for accurate order-wide logistics sharing
  const effectiveTotalBodies = useMemo(() => {
    if (propTotalCombinedBodies && propTotalCombinedBodies > 0) {
      return propTotalCombinedBodies
    }
    if (initialSets && initialSets.length > 0) {
      const sum = initialSets.reduce((acc, s) => {
        const sysBodies = s.summary?.bodies || s.breakdown?.totalBodies || 36
        const qty = Math.max(1, Number(s.quantity) || 1)
        return acc + (sysBodies * qty)
      }, 0)
      if (sum > 0) return sum
    }
    return totalBodies > 0 ? totalBodies : 1
  }, [propTotalCombinedBodies, initialSets, totalBodies])

  // Other Extra Amount (editable and initialized from calculator or project summary)
  const [extraAmount, setExtraAmount] = useState<number>(() => {
    if (initialExtraAmount !== undefined && !isNaN(initialExtraAmount)) return initialExtraAmount
    if (project.inputs?.['__extra_amount'] !== undefined) {
      const v = Number(project.inputs['__extra_amount'])
      if (!isNaN(v)) return v
    }
    if (project.summary?.extra_amount !== undefined) {
      const v = Number(project.summary.extra_amount)
      if (!isNaN(v)) return v
    }
    return 0
  })

  // Synchronize extraAmount when parent calculator passes a new value
  useEffect(() => {
    if (initialExtraAmount !== undefined && !isNaN(initialExtraAmount)) {
      setExtraAmount(initialExtraAmount)
    }
  }, [initialExtraAmount])

  const docDiscountPct = useMemo(() => {
    if (summary && summary.discountPct !== undefined && !isNaN(summary.discountPct)) return summary.discountPct
    if (project.summary && project.summary.discount_pct !== undefined && !isNaN(Number(project.summary.discount_pct))) return Number(project.summary.discount_pct)
    if (project.inputs && project.inputs['__discount_pct'] !== undefined && !isNaN(Number(project.inputs['__discount_pct']))) return Number(project.inputs['__discount_pct'])
    return 0
  }, [summary, project])

  // Dynamic unit-wise rate breakdown (Transport + Installation + Accommodation distributed by bodies + extra amount + dynamic profit)
  const breakdownResult = useMemo(
    () => computeUnitWiseBreakdown(summary, wb, includeInstallation, profitPct, docDiscountPct, effectiveTotalBodies, extraAmount),
    [summary, wb, includeInstallation, profitPct, docDiscountPct, effectiveTotalBodies, extraAmount],
  )

  // Auto-derived destination from address for Freight Clause
  const clientDestination = useMemo(() => {
    if (!clientInfo.buyer_address) return 'Site'
    const lines = clientInfo.buyer_address.split('\n').map(l => l.trim()).filter(Boolean)
    return lines.length > 0 ? lines[lines.length - 1] : 'Site'
  }, [clientInfo.buyer_address])

  // 2. Default Sets Builder (e.g. SYSTEM A, SYSTEM B, SYSTEM C)
  const defaultSets = useMemo<QuotationSet[]>(() => {
    if (initialSets && Array.isArray(initialSets) && initialSets.length > 0) {
      return initialSets.map((s: any, idx: number) => ({
        ...s,
        system_title: s.system_title !== undefined ? s.system_title : (initialSets.length > 1 ? (s.set_label?.startsWith('SYSTEM') ? s.set_label : `SYSTEM ${String.fromCharCode(65 + idx)}`) : ''),
        set_label: s.set_label || (Number(s.quantity) > 1 ? `${s.quantity} ${s.unit || 'SET'}S` : `1 ${s.unit || 'SET'}`),
        unit: s.unit || 'SET',
        rate: Number(s.rate) || 0,
        amount: s.amount !== undefined ? Number(s.amount) : Math.round((Number(s.quantity) || 1) * (Number(s.rate) || 0)),
        pre_discount_rate: s.pre_discount_rate !== undefined ? Number(s.pre_discount_rate) : Number(s.rate) || 0,
        pre_discount_amount: s.pre_discount_amount !== undefined ? Number(s.pre_discount_amount) : (s.amount !== undefined ? Number(s.amount) : Math.round((Number(s.quantity) || 1) * (Number(s.rate) || 0))),
      }))
    }

    if (!isMultiSystem) {
      // In single system or unit_wise mode: EXACT format from uploaded PDF (NO SYSTEM A title)
      const items: SetUnitItem[] = breakdownResult.rows.map((r, idx) => ({
        id: String(idx + 1),
        model_name: r.model_name,
        model_code: r.model_code,
        dimensions: r.dimensions,
        unit_qty: r.qty,
        rate: r.unit_rate_with_profit,
        amount: r.total_price_with_profit,
      }))
      const sysQty = initialSets?.[0]?.quantity ? Math.max(1, Number(initialSets[0].quantity)) : (project.summary?.systems?.[0]?.quantity ? Math.max(1, Number(project.summary.systems[0].quantity)) : 1)
      const unitRate = Math.round(breakdownResult.finalBeforeGst || summary.beforeGst || 1843257)
      const baseAmount = unitRate * sysQty
      const preDiscRate = Math.round(breakdownResult.totalPriceWithProfit || (summary.subtotalWithProfit ?? unitRate))
      const preDiscAmount = preDiscRate * sysQty

      return [
        {
          id: '1',
          system_title: '', // EMPTY so no "SYSTEM A" row in the table!
          set_label: sysQty > 1 ? `${sysQty} ${initialSets?.[0]?.unit || 'SET'}S` : `1 ${initialSets?.[0]?.unit || 'SET'}`,
          quantity: sysQty,
          unit: initialSets?.[0]?.unit || 'SET',
          rate: unitRate,
          amount: Number(baseAmount.toFixed(2)),
          pre_discount_rate: preDiscRate,
          pre_discount_amount: Number(preDiscAmount.toFixed(2)),
          items: items.length > 0 ? items : [
            {
              id: '1',
              model_name: 'DOUBLE FACE MOVABLE UNIT (DFMU)',
              dimensions: '2200 (HT) X 1830 (W) X 610 (D) MM 5 Compartments',
              unit_qty: 4,
            },
            {
              id: '2',
              model_name: 'SINGLE FACE MOVABLE UNIT (SFMU)',
              dimensions: '2200 (HT) X 1830 (W) X 475 (D) MM 5 Compartments',
              unit_qty: 1,
            },
            {
              id: '3',
              model_name: 'SINGLE FACE FIX UNIT (SFFU)',
              dimensions: '2200 (HT) X 1830 (W) X 330 (D) MM 5 Compartments',
              unit_qty: 1,
            },
          ],
        },
      ]
    }

    if ((!isUnsavedWarning || justSaved) && project.summary?.quotation_sets && Array.isArray(project.summary.quotation_sets) && project.summary.quotation_sets.length > 0) {
      return project.summary.quotation_sets.map((s: any, idx: number) => ({
        ...s,
        system_title: s.system_title !== undefined ? s.system_title : (s.set_label?.startsWith('SYSTEM') ? s.set_label : `SYSTEM ${String.fromCharCode(65 + idx)}`),
        set_label: s.set_label || `${s.quantity || 1} ${s.unit || 'SET'}`,
        unit: s.unit || 'SET',
        rate: Number(s.rate) || 0,
        amount: s.amount !== undefined ? Number(s.amount) : Math.round((Number(s.quantity) || 1) * (Number(s.rate) || 0)),
        pre_discount_rate: s.pre_discount_rate !== undefined ? Number(s.pre_discount_rate) : Number(s.rate) || 0,
        pre_discount_amount: s.pre_discount_amount !== undefined ? Number(s.pre_discount_amount) : (s.amount !== undefined ? Number(s.amount) : Math.round((Number(s.quantity) || 1) * (Number(s.rate) || 0))),
      }))
    }

    const items: SetUnitItem[] = breakdownResult.rows.map((r, idx) => ({
      id: String(idx + 1),
      model_name: r.model_name,
      model_code: r.model_code,
      dimensions: r.dimensions,
      unit_qty: r.qty,
      rate: r.unit_rate_with_profit,
      amount: r.total_price_with_profit,
    }))
    const baseAmount = breakdownResult.finalBeforeGst || summary.beforeGst || 1843257
    const preDiscAmount = breakdownResult.totalPriceWithProfit || (summary.subtotalWithProfit ?? baseAmount)

    return [
      {
        id: '1',
        system_title: 'SYSTEM A',
        set_label: '1 SET',
        quantity: 1,
        unit: 'SET',
        rate: Math.round(preDiscAmount),
        amount: Number(baseAmount.toFixed(2)),
        pre_discount_rate: Math.round(preDiscAmount),
        pre_discount_amount: Number(preDiscAmount.toFixed(2)),
        items,
      },
    ]
  }, [isMultiSystem, initialSets, project.summary?.quotation_sets, breakdownResult, summary])

  const [sets, setSets] = useState<QuotationSet[]>(defaultSets)

  // 3. Dynamic Unit-Wise Line Items (Directly derived from live breakdownResult — reflects all calculator changes)
  const lineItems = useMemo<QuotationLineItem[]>(() => {
    return breakdownResult.rows.map((r, idx) => ({
      id: String(idx + 1),
      model_code: r.model_code,
      model_name: r.model_name,
      dimensions: r.dimensions,
      quantity: r.qty,
      unit_base_rate: r.unit_rate,
      extra_charges: r.transportation,
      total_unit_rate: r.total_unit_rate,
      rate: r.final_unit_rate,
      amount: r.final_amount,
    }))
  }, [breakdownResult])

  // Live update sets from initialSets (multi-system calculator sync)
  useEffect(() => {
    if (initialSets && initialSets.length > 0) {
      setSets(
        initialSets.map((s: any, idx: number) => ({
          ...s,
          system_title:
            s.system_title !== undefined
              ? s.system_title
              : initialSets.length > 1
              ? (s.set_label?.startsWith('SYSTEM') ? s.set_label : `SYSTEM ${String.fromCharCode(65 + idx)}`)
              : '',
          set_label: s.set_label || (Number(s.quantity) > 1 ? `${s.quantity} ${s.unit || 'SET'}S` : `1 ${s.unit || 'SET'}`),
          unit: s.unit || 'SET',
          rate: Number(s.rate) || 0,
          amount:
            s.amount !== undefined
              ? Number(s.amount)
              : Math.round((Number(s.quantity) || 1) * (Number(s.rate) || 0)),
          pre_discount_rate: s.pre_discount_rate !== undefined ? Number(s.pre_discount_rate) : Number(s.rate) || 0,
          pre_discount_amount: s.pre_discount_amount !== undefined ? Number(s.pre_discount_amount) : (s.amount !== undefined ? Number(s.amount) : Math.round((Number(s.quantity) || 1) * (Number(s.rate) || 0))),
          items: s.items && s.items.length > 0 ? s.items : breakdownResult.rows.map((r, rIdx) => ({
            id: String(rIdx + 1),
            model_name: r.model_name,
            model_code: r.model_code,
            dimensions: r.dimensions,
            unit_qty: r.qty,
            rate: r.unit_rate_with_profit,
            amount: r.total_price_with_profit,
          })),
        }))
      )
    }
  }, [initialSets, breakdownResult])

  // Live update bodyItems from initialSets
  useEffect(() => {
    if (initialSets && initialSets.length > 0) {
      setBodyItems(generateDefaultBodyItems(summary, wb, project.inputs, initialSets))
    }
  }, [initialSets, summary, wb])

  // Live update for single-system mode (fallback when initialSets is not present)
  useEffect(() => {
    if (initialSets && initialSets.length > 0) return

    // Sync single set (rate is 1-set cost, amount is rate * qty)
    const newTotal = breakdownResult.finalBeforeGst
    if (!newTotal || newTotal <= 0) return
    const preDiscTotal = breakdownResult.totalPriceWithProfit || newTotal
    setSets((prev) => {
      const qty = prev[0]?.quantity || 1
      const newRate = Math.round(newTotal)
      return [
        {
          id: '1',
          system_title: isMultiSystem ? 'SYSTEM A' : '',
          set_label: `${qty} SET`,
          unit: 'SET',
          items: breakdownResult.rows.map((r, idx) => ({
            id: String(idx + 1),
            model_name: r.model_name,
            model_code: r.model_code,
            dimensions: r.dimensions,
            unit_qty: r.qty,
            rate: r.unit_rate_with_profit,
            amount: r.total_price_with_profit,
          })),
          quantity: qty,
          rate: newRate,
          amount: Math.round(newRate * qty),
          pre_discount_rate: Math.round(preDiscTotal / qty),
          pre_discount_amount: Math.round(preDiscTotal),
        },
      ]
    })
  }, [breakdownResult, initialSets, isMultiSystem])

  // Live update rates if profitPct is changed locally inside quotation document
  useEffect(() => {
    if (!initialSets || initialSets.length === 0) return
    if (profitPct === defaultProfitPct) return

    setSets((prev) =>
      prev.map((s) => {
        const qty = Math.max(1, Number(s.quantity) || 1)
        if (s.breakdown) {
          const costBase = s.breakdown.totalCostBase || s.summary?.costBase || s.breakdown.subtotal || 0
          const profitAmount = costBase * (profitPct / 100)
          const withProfit = costBase + profitAmount
          const discountAmount = withProfit * (docDiscountPct / 100)
          const newRate = Math.round(withProfit - discountAmount)
          return {
            ...s,
            rate: newRate,
            amount: Math.round(newRate * qty),
            pre_discount_rate: Math.round(withProfit / qty),
            pre_discount_amount: Math.round(withProfit),
          }
        }
        return s
      })
    )
  }, [profitPct, defaultProfitPct, docDiscountPct])

  // Helper to re-sync line items (live dynamic sync is active by default)
  const handleSyncBreakdownToLineItems = () => {
    toast('Unit rates are automatically synchronized with calculator')
  }

  // 5. Intelligent Page 1 Space Check: Fits on Page 1 if total commercial rows <= 4
  const totalCommercialRows = useMemo(() => {
    if (effectiveQuotationMode === 'set_wise') {
      return sets.reduce((acc, s) => acc + s.items.length, 0)
    }
    if (effectiveQuotationMode === 'unit_breakdown') {
      return breakdownResult.rows.length
    }
    return lineItems.length
  }, [effectiveQuotationMode, sets, lineItems, breakdownResult])

  const canFitOnPage1Naturally = totalCommercialRows <= 4
  const [showBodiesTable, setShowBodiesTable] = useState<boolean>(() => {
    if (project.summary?.show_bodies_table !== undefined) {
      return Boolean(project.summary.show_bodies_table)
    }
    return true
  })
  const [bodyPlacement, setBodyPlacement] = useState<'dedicated' | 'terms' | 'commercial' | 'none'>(() => {
    if (project.summary?.body_placement) {
      return project.summary.body_placement as any
    }
    if (project.summary?.show_bodies_table === false) return 'none'
    // For multi-system quotations (2+ systems), default to dedicated page so terms & commercial never crowd
    if (sets.length >= 2 || (initialSets && initialSets.length >= 2)) return 'dedicated'
    if (project.summary?.body_table_on_page_1) return 'commercial'
    return canFitOnPage1Naturally ? 'commercial' : 'terms'
  })

  const [placeBodyOnPage1, setPlaceBodyOnPage1] = useState<boolean>(() => {
    if (project.summary?.body_placement) {
      return project.summary.body_placement === 'commercial'
    }
    if (project.summary?.body_table_on_page_1 !== undefined) {
      return Boolean(project.summary.body_table_on_page_1)
    }
    return canFitOnPage1Naturally
  })

  // Full 5-Page Proposal toggle (Page 1: Commercial; Page 2: Terms; Page 3: Compactor Specs [merged p3+p4]; Page 4: Powder Coating [merged p5+p6]; Page 5: QC [merged p7+p8])
  const [includeFullProposal, setIncludeFullProposal] = useState<boolean>(() => {
    if (project.summary?.include_full_proposal !== undefined) {
      return Boolean(project.summary.include_full_proposal)
    }
    if (typeof window !== 'undefined' && window.localStorage) {
      const savedPref = localStorage.getItem('examcalc_default_include_5_page_proposal')
      if (savedPref !== null) {
        return savedPref === 'true'
      }
    }
    return true
  })

  // Table size, scaling, density and multi-system layout controls
  const [tableScale, setTableScale] = useState<number>(100)
  const [tableDensity, setTableDensity] = useState<'compact' | 'normal' | 'relaxed'>('normal')
  // Systems per page: 'auto' (smart layout based on model row counts) or manual limit
  const [systemsPerPageSetting, setSystemsPerPageSetting] = useState<'auto' | 2 | 3 | 4 | 5 | 999>('auto')

  // Smart calculation: 5 models per system -> max 3 systems; 2-3 models -> 4 systems
  const effectiveSystemsPerPage = useMemo(() => {
    if (systemsPerPageSetting !== 'auto') {
      return systemsPerPageSetting
    }
    if (!sets || sets.length <= 1) return 4

    let maxItemsInAnySystem = 1
    let totalItemsAllSystems = 0
    sets.forEach((s) => {
      const itemCount = (s.items && s.items.length > 0) ? s.items.length : 3
      if (itemCount > maxItemsInAnySystem) maxItemsInAnySystem = itemCount
      totalItemsAllSystems += itemCount
    })
    const avgItems = totalItemsAllSystems / Math.max(1, sets.length)

    // When systems have 4-5 models (DFMU, SFMU, SFFU, DFFU, DFMU-D), use 3 systems per page
    if (maxItemsInAnySystem >= 4 || avgItems >= 3.5) {
      return 3
    }
    return 4
  }, [sets, systemsPerPageSetting])

  const systemsPerPage = effectiveSystemsPerPage

  // Word-style draggable column widths state (adjustable by dragging borders or +/-)
  const [goodsColWidths, setGoodsColWidths] = useState({
    sl: 34,
    qty: 55,
    set: 60,
    rate: 85,
    amount: 100,
  })

  const [bodyColWidths, setBodyColWidths] = useState({
    sr: 40,
    per_bay: 75,
    qty: 55,
    body: 85,
  })

  const startColResize = (
    tableType: 'goods' | 'body',
    colKey: string,
    e: React.MouseEvent | React.TouchEvent
  ) => {
    e.preventDefault()
    e.stopPropagation()
    const startX = 'touches' in e ? e.touches[0].clientX : e.clientX
    const startWidth = tableType === 'goods'
      ? (goodsColWidths as any)[colKey] || 60
      : (bodyColWidths as any)[colKey] || 60

    const handleMove = (moveEvt: MouseEvent | TouchEvent) => {
      const currentX = 'touches' in moveEvt ? moveEvt.touches[0].clientX : moveEvt.clientX
      const delta = currentX - startX
      const newWidth = Math.max(25, Math.min(220, Math.round(startWidth + delta)))

      if (tableType === 'goods') {
        setGoodsColWidths((prev) => ({ ...prev, [colKey]: newWidth }))
      } else {
        setBodyColWidths((prev) => ({ ...prev, [colKey]: newWidth }))
      }
    }

    const handleUp = () => {
      window.removeEventListener('mousemove', handleMove)
      window.removeEventListener('mouseup', handleUp)
      window.removeEventListener('touchmove', handleMove)
      window.removeEventListener('touchend', handleUp)
    }

    window.addEventListener('mousemove', handleMove)
    window.addEventListener('mouseup', handleUp)
    window.addEventListener('touchmove', handleMove)
    window.addEventListener('touchend', handleUp)
  }

  const tableFontSize = useMemo(() => `${(10.1 * (tableScale / 100)).toFixed(2)}px`, [tableScale])
  const tablePadding = useMemo(() => {
    if (tableDensity === 'compact') return '3px 4px'
    if (tableDensity === 'relaxed') return '6px 8px'
    return '5px 6px'
  }, [tableDensity])

  const moveSystem = (fromIdx: number, toIdx: number) => {
    if (toIdx < 0 || toIdx >= sets.length) return
    setSets((prev) => {
      const updated = [...prev]
      const temp = updated[fromIdx]
      updated[fromIdx] = updated[toIdx]
      updated[toIdx] = temp
      return updated
    })
  }

  const moveBodyItem = (fromIdx: number, toIdx: number) => {
    if (toIdx < 0 || toIdx >= bodyItems.length) return
    setBodyItems((prev) => {
      const updated = [...prev]
      const temp = updated[fromIdx]
      updated[fromIdx] = updated[toIdx]
      updated[toIdx] = temp
      return updated
    })
  }

  // Multi-system commercial pagination chunks (e.g. 4 systems per page)
  const commercialChunks = useMemo(() => {
    if (sets.length <= systemsPerPage || systemsPerPage >= 999) {
      return [{ chunkIndex: 0, setsList: sets, startSetIndex: 0, isLastChunk: true }]
    }
    const chunks: Array<{ chunkIndex: number; setsList: QuotationSet[]; startSetIndex: number; isLastChunk: boolean }> = []
    for (let i = 0; i < sets.length; i += systemsPerPage) {
      const chunkSets = sets.slice(i, i + systemsPerPage)
      chunks.push({
        chunkIndex: Math.floor(i / systemsPerPage),
        setsList: chunkSets,
        startSetIndex: i,
        isLastChunk: i + systemsPerPage >= sets.length,
      })
    }
    return chunks
  }, [sets, systemsPerPage])

  const commercialPagesCount = commercialChunks.length
  const totalPages = includeFullProposal ? 5 : 2

  // Responsive mobile view
  const [mobileView, setMobileView] = useState<'document' | 'sidebar'>('document')

  useEffect(() => {
    document.body.classList.add('official-quotation-open')
    return () => {
      document.body.classList.remove('official-quotation-open')
    }
  }, [])

  // Calculations
  const subtotal = useMemo(() => {
    if (sets.length > 1) {
      return sets.reduce((acc, s) => acc + (Number(s.amount) || 0), 0)
    }
    if (effectiveQuotationMode === 'set_wise') {
      return sets.reduce((acc, s) => acc + (Number(s.amount) || 0), 0)
    }
    if (effectiveQuotationMode === 'unit_breakdown') {
      return breakdownResult.finalBeforeGst
    }
    if (!isMultiSystem && summary?.beforeGst) {
      const q = sets[0]?.quantity ? Math.max(1, Number(sets[0].quantity)) : 1
      return summary.beforeGst * q
    }
    return lineItems.reduce((acc, it) => acc + (Number(it.amount) || 0), 0)
  }, [isMultiSystem, summary?.beforeGst, effectiveQuotationMode, sets, lineItems, breakdownResult])

  const gst18 = useMemo(() => subtotal * 0.18, [subtotal])
  const rawGrandTotal = subtotal + gst18
  const grandTotal = Math.round(rawGrandTotal)
  const roundOff = (grandTotal - rawGrandTotal).toFixed(2)

  const amountInWords = useMemo(() => {
    return `Rupees ${numberToIndianWords(grandTotal)} Only`
  }, [grandTotal])

  const effectiveDiscountPct = useMemo(() => {
    if (breakdownResult && breakdownResult.discountPct > 0) return breakdownResult.discountPct
    if (summary && summary.discountPct && summary.discountPct > 0) return summary.discountPct
    if (project.summary && project.summary.discount_pct && Number(project.summary.discount_pct) > 0) return Number(project.summary.discount_pct)
    if (project.inputs && project.inputs['__discount_pct'] && Number(project.inputs['__discount_pct']) > 0) return Number(project.inputs['__discount_pct'])
    return 0
  }, [breakdownResult, summary, project])

  const hasDiscount = effectiveDiscountPct > 0

  const { totalBeforeDiscount, totalDiscountAmount } = useMemo(() => {
    if (!hasDiscount) {
      return { totalBeforeDiscount: subtotal, totalDiscountAmount: 0 }
    }
    let beforeDisc = 0
    if (sets.length > 1) {
      const setsPreDiscSum = sets.reduce((sum, s: any) => sum + (Number(s.pre_discount_amount) || Number(s.amount) || 0), 0)
      if (setsPreDiscSum > 0) {
        beforeDisc = setsPreDiscSum
      } else {
        const factor = 1 - (effectiveDiscountPct / 100)
        beforeDisc = factor > 0 ? Number((subtotal / factor).toFixed(2)) : subtotal
      }
    } else if (!isMultiSystem && breakdownResult.totalPriceWithProfit > 0) {
      const q = sets[0]?.quantity ? Math.max(1, Number(sets[0].quantity)) : 1
      beforeDisc = breakdownResult.totalPriceWithProfit * q
    } else if (effectiveQuotationMode === 'set_wise' && sets.length > 0) {
      const setsPreDiscSum = sets.reduce((sum, s: any) => sum + (Number(s.pre_discount_amount) || 0), 0)
      if (setsPreDiscSum > 0) {
        beforeDisc = setsPreDiscSum
      } else if (breakdownResult.totalPriceWithProfit > 0) {
        beforeDisc = breakdownResult.totalPriceWithProfit
      } else {
        const factor = 1 - (effectiveDiscountPct / 100)
        beforeDisc = factor > 0 ? Number((subtotal / factor).toFixed(2)) : subtotal
      }
    } else if (effectiveQuotationMode === 'unit_breakdown') {
      beforeDisc = breakdownResult.totalPriceWithProfit || subtotal
    } else {
      const lineItemsPreDiscSum = breakdownResult.totalPriceWithProfit
      if (lineItemsPreDiscSum > 0) {
        beforeDisc = lineItemsPreDiscSum
      } else {
        const factor = 1 - (effectiveDiscountPct / 100)
        beforeDisc = factor > 0 ? Number((subtotal / factor).toFixed(2)) : subtotal
      }
    }
    const discAmt = Number((beforeDisc - subtotal).toFixed(2))
    return {
      totalBeforeDiscount: beforeDisc,
      totalDiscountAmount: discAmt,
    }
  }, [hasDiscount, effectiveDiscountPct, subtotal, effectiveQuotationMode, sets, breakdownResult])

  // Pre-calculate pre-discount amounts for sets so the displayed column amounts sum exactly to totalBeforeDiscount
  const setDisplayAmounts = useMemo(() => {
    if (!hasDiscount || sets.length === 0) {
      return sets.map((s) => ({
        rate: Number(s.rate) || 0,
        amount: Number(s.amount) || 0,
      }))
    }
    let runningSum = 0
    return sets.map((s, idx) => {
      let amt = 0
      if (idx === sets.length - 1) {
        amt = Number((totalBeforeDiscount - runningSum).toFixed(2))
      } else {
        const prop = subtotal > 0 ? ((Number(s.amount) || 0) / subtotal) : (1 / sets.length)
        amt = Number((totalBeforeDiscount * prop).toFixed(2))
        runningSum += amt
      }
      const rate = Math.round(amt / Math.max(1, s.quantity || 1))
      return { rate, amount: amt }
    })
  }, [hasDiscount, sets, totalBeforeDiscount, subtotal])

  // Pre-calculate pre-discount amounts for unit-wise lineItems
  const lineItemDisplayAmounts = useMemo(() => {
    if (!hasDiscount || lineItems.length === 0) {
      return lineItems.map((it) => ({
        rate: Number(it.rate) || 0,
        amount: Number(it.amount) || 0,
      }))
    }
    let runningSum = 0
    return lineItems.map((it, idx) => {
      let amt = 0
      if (idx === lineItems.length - 1) {
        amt = Number((totalBeforeDiscount - runningSum).toFixed(2))
      } else {
        const prop = subtotal > 0 ? ((Number(it.amount) || 0) / subtotal) : (1 / lineItems.length)
        amt = Number((totalBeforeDiscount * prop).toFixed(2))
        runningSum += amt
      }
      const rate = Math.round(amt / Math.max(1, Number(it.quantity) || 1))
      return { rate, amount: amt }
    })
  }, [hasDiscount, lineItems, totalBeforeDiscount, subtotal])

  // Check if compactor width is >= 3000 mm for DOUBLE DRIVE WITH REDUCTION WHEEL in red specs
  const hasDoubleDrive = useMemo(() => {
    let projectWidth = 1830
    if (wb) {
      projectWidth = wb.num('Total cal', 'D3') || 1830
    } else if (project.inputs?.['Total cal!D3']) {
      projectWidth = Number(project.inputs['Total cal!D3']) || 1830
    }
    if (projectWidth >= 3000) return true

    if (effectiveQuotationMode === 'set_wise') {
      return sets.some((s) =>
        s.items.some((it) => {
          const w = extractWidthFromDimensions(it.dimensions)
          return w !== null && w >= 3000
        })
      )
    }

    return lineItems.some((it) => {
      const w = extractWidthFromDimensions(it.dimensions)
      return w !== null && w >= 3000
    })
  }, [wb, project.inputs, effectiveQuotationMode, sets, lineItems])

  // Track Count derived from Total cal!I3 or inputs (defaults to 2)
  const projectTrackCount = useMemo(() => {
    let tc = 2
    if (wb) {
      const val = wb.num('Total cal', 'I3')
      if (val && val > 0) tc = Math.round(val)
    } else if (project.inputs?.['Total cal!I3']) {
      const val = Number(project.inputs['Total cal!I3'])
      if (val && val > 0) tc = Math.round(val)
    }
    return tc
  }, [wb, project.inputs])

  const getSystemTrackCount = (s: any, sIdx: number): number => {
    if (s.track_count && Number(s.track_count) > 0) return Number(s.track_count)
    const sysObj = project.summary?.systems?.[sIdx]
    if (sysObj?.overrides?.['Total cal!I3']) {
      const val = Number(sysObj.overrides['Total cal!I3'])
      if (val > 0) return Math.round(val)
    }
    if (s.overrides?.['Total cal!I3']) {
      const val = Number(s.overrides['Total cal!I3'])
      if (val > 0) return Math.round(val)
    }
    if (sIdx === 0) {
      if (wb) {
        const val = wb.num('Total cal', 'I3')
        if (val && val > 0) return Math.round(val)
      }
      if (project.inputs?.['Total cal!I3']) {
        const val = Number(project.inputs['Total cal!I3'])
        if (val > 0) return Math.round(val)
      }
    }
    return projectTrackCount || 2
  }

  const systemHasDoubleDrive = (s: any): boolean => {
    const items = s.items || []
    return items.some((it: any) => {
      const w = extractWidthFromDimensions(it.dimensions)
      return w !== null && w >= 3000
    })
  }

  const allSystemsTrackCounts = useMemo(() => {
    return sets.map((s, idx) => getSystemTrackCount(s, idx))
  }, [sets, project.summary?.systems, wb, project.inputs, projectTrackCount])

  const allTracksIdentical = useMemo(() => {
    if (sets.length <= 1) return true
    return allSystemsTrackCounts.every((tc) => tc === allSystemsTrackCounts[0])
  }, [sets.length, allSystemsTrackCounts])

  const allSystemsHaveDoubleDrive = useMemo(() => {
    if (sets.length <= 1) return hasDoubleDrive
    return sets.every((s) => systemHasDoubleDrive(s))
  }, [sets, hasDoubleDrive])

  // Initial red specifications: preserved from saved quotation summary or generated from dimensions and tracks
  const defaultRedSpecs = useMemo<string[]>(() => {
    if (project.summary?.red_specs && Array.isArray(project.summary.red_specs) && project.summary.red_specs.length > 0) {
      return project.summary.red_specs
    }
    // Only show DOUBLE DRIVE in shared bottom specs if single system or all systems have double drive.
    // If only specific systems have double drive, it is shown in red under that specific system only.
    const effWidth = (sets.length > 1 ? allSystemsHaveDoubleDrive : hasDoubleDrive) ? 3000 : 1830

    // Only show track count in shared bottom specs if single system or all systems have the SAME track count.
    // If track counts differ, track count is displayed in red under each specific system only.
    const trackToShow = (sets.length > 1 && !allTracksIdentical) ? 0 : projectTrackCount

    return generateDefaultRedSpecs(effWidth, trackToShow, wb, project.inputs)
  }, [project.summary?.red_specs, sets.length, allSystemsHaveDoubleDrive, hasDoubleDrive, allTracksIdentical, projectTrackCount, wb, project.inputs])

  const [redSpecs, setRedSpecs] = useState<string[]>(defaultRedSpecs)

  // Quotation download file name builder (sanitized for Windows/Mac/Linux OS file save dialogs)
  const getQuotationFileName = () => {
    const qtnNo = clientInfo.qtn_no?.trim()
    const buyer = clientInfo.buyer_name?.trim()
    const proj = project.name?.trim()

    let name = ''
    if (proj && proj !== 'Untitled Project') {
      name = proj
      if (buyer && !proj.toLowerCase().includes(buyer.toLowerCase())) {
        name = `${proj} - ${buyer}`
      }
    } else if (qtnNo && buyer && !qtnNo.toLowerCase().includes(buyer.toLowerCase())) {
      name = `${qtnNo} - ${buyer}`
    } else if (qtnNo) {
      name = qtnNo
    } else if (buyer) {
      name = `Quotation - ${buyer}`
    } else {
      name = 'Quotation'
    }

    return name.replace(/[/\\?%*:|"<>]/g, '-').trim()
  }

  // Synchronize document.title with quotation file name so browser "Save as PDF" directly suggests this filename
  useEffect(() => {
    const fileName = getQuotationFileName()
    const prevTitle = document.title
    document.title = fileName
    return () => {
      document.title = prevTitle
    }
  }, [clientInfo.qtn_no, clientInfo.buyer_name, project.name])

  const [isGeneratingPdf, setIsGeneratingPdf] = useState(false)

  const generatePdfInstance = async (): Promise<{ pdf: any; fileName: string; pagesCount: number } | null> => {
    const fileName = getQuotationFileName()
    const element = document.getElementById('official-quotation-pages')
    if (!element) return null
    const pages = Array.from(element.querySelectorAll<HTMLElement>('.official-pdf-page'))
    if (!pages || pages.length === 0) return null

    // jsPDF and html2canvas are statically imported upfront to prevent dynamic chunk 404s

    const pdf = new jsPDF({
      orientation: 'portrait',
      unit: 'mm',
      format: 'a4',
      compress: true,
    })

    const prevScrollTop = element.scrollTop
    try {
      for (let idx = 0; idx < pages.length; idx++) {
        const page = pages[idx]
        const canvas = await html2canvas(page, {
          scale: 2,
          useCORS: true,
          logging: false,
          backgroundColor: '#ffffff',
          imageTimeout: 15000,
          scrollY: 0,
          ignoreElements: (el: Element) => el.classList?.contains('no-print'),
        })

        const imgData = canvas.toDataURL('image/jpeg', 0.98)
        if (idx > 0) {
          pdf.addPage('a4', 'portrait')
        }
        pdf.addImage(imgData, 'JPEG', 0, 0, 210, 297, undefined, 'FAST')
      }
      return { pdf, fileName, pagesCount: pages.length }
    } finally {
      element.scrollTop = prevScrollTop
    }
  }

  const handleDirectDownload = async () => {
    setIsGeneratingPdf(true)
    try {
      const generated = await generatePdfInstance()
      if (!generated) {
        handlePrint()
        return
      }
      generated.pdf.save(`${generated.fileName}.pdf`)
      toast(`Downloaded "${generated.fileName}.pdf" (${generated.pagesCount} pages) directly!`)

      // Record in Activity Log
      const token = tokenStore.get()
      if (token && effectiveOrgId) {
        api.logActivity(token, effectiveOrgId, 'pdf_downloaded', {
          name: project.name,
          file_name: generated.fileName,
          buyer_name: clientInfo.buyer_name || 'Direct Buyer',
          qtn_no: clientInfo.qtn_no || '',
          total: grandTotal,
          mode: effectiveQuotationMode,
        }, project.id).catch(() => {})
      }
    } catch (e) {
      console.warn('Direct PDF download failed, falling back to print dialog:', e)
      handlePrint()
    } finally {
      setIsGeneratingPdf(false)
    }
  }

  const handleWhatsAppShare = async () => {
    setIsGeneratingPdf(true)
    try {
      const generated = await generatePdfInstance()
      if (!generated) {
        toast('Could not render quotation PDF for WhatsApp sharing', 'err')
        return
      }

      const { pdf, fileName } = generated
      const blob = pdf.output('blob')
      const pdfFile = new File([blob], `${fileName}.pdf`, { type: 'application/pdf' })

      // Format phone number: buyer contact or contact person's phone
      let phone = (
        clientInfo.buyer_contact ||
        (clientInfo.contacts && clientInfo.contacts.find((c: any) => c.phone)?.phone) ||
        (clientInfo.contacts && clientInfo.contacts[0]?.phone) ||
        ''
      ).replace(/[^0-9]/g, '')
      if (phone.length === 10) {
        phone = '91' + phone
      }

      const qtn = clientInfo.qtn_no || 'Quotation'
      const buyer = clientInfo.buyer_name || ''
      const totalAmt = inr(Math.round(grandTotal))
      const msg = `*Quotation: ${qtn}*\nDear ${buyer || 'Sir/Madam'},\n\nPlease find attached official quotation for *${project.name}*.\n*Total Amount:* ${totalAmt} (Inc. GST)\n\nThank you,\n*${businessProfile.company_name || 'Myriad Storage Systems'}*`

      const isMobileDevice = typeof navigator !== 'undefined' && /Android|iPhone|iPad|iPod|Mobi/i.test(navigator.userAgent || '')

      // 1. Mobile & Web Share API: Share the actual PDF binary file directly!
      // CRITICAL FOR WHATSAPP ON MOBILE:
      // WhatsApp on Android and iOS drops the file attachment and only displays text if `text` is passed!
      // By passing ONLY `files: [pdfFile]`, WhatsApp treats it as a document and attaches the actual PDF!
      if (typeof navigator !== 'undefined' && typeof navigator.canShare === 'function') {
        try {
          if (navigator.canShare({ files: [pdfFile] })) {
            await navigator.share({
              files: [pdfFile],
              title: `${fileName}.pdf`,
            })
            toast('Quotation PDF sent directly to WhatsApp!', 'ok')
            return
          }
        } catch (shareErr: any) {
          if (shareErr?.name === 'AbortError') {
            // User intentionally dismissed the share sheet
            return
          }
          console.warn('Native file share cancelled or failed, falling back to direct app link:', shareErr)
        }
      }

      // 2. Direct Mobile App Launch (whatsapp://) vs Desktop (WhatsApp Web)
      // Save PDF file to device so user has the downloaded file
      pdf.save(`${fileName}.pdf`)

      if (isMobileDevice) {
        // Direct WhatsApp native application URL scheme
        // Using whatsapp:// directly opens the WhatsApp Mobile App on Android and iOS (never a web page)
        const appUrl = phone
          ? `whatsapp://send?phone=${phone}&text=${encodeURIComponent(msg)}`
          : `whatsapp://send?text=${encodeURIComponent(msg)}`

        // Launch WhatsApp app directly
        window.location.href = appUrl
        toast(`PDF downloaded! Opening WhatsApp app — attach "${fileName}.pdf"`, 'ok')
      } else {
        // Desktop WhatsApp Web fallback
        const waWebUrl = phone
          ? `https://web.whatsapp.com/send?phone=${phone}&text=${encodeURIComponent(msg)}`
          : `https://web.whatsapp.com/send?text=${encodeURIComponent(msg)}`

        window.open(waWebUrl, '_blank')
        toast(`PDF downloaded! WhatsApp Web opened — attach "${fileName}.pdf" to send it`, 'ok')
      }
    } catch (e: any) {
      toast(e.message || 'Error generating PDF for WhatsApp', 'err')
    } finally {
      setIsGeneratingPdf(false)
    }
  }

  const handlePrint = () => {
    const fileName = getQuotationFileName()
    const prevTitle = document.title
    document.title = fileName
    window.print()

    // Record in Activity Log
    const token = tokenStore.get()
    if (token && effectiveOrgId) {
      api.logActivity(token, effectiveOrgId, 'pdf_downloaded', {
        name: project.name,
        file_name: fileName,
        buyer_name: clientInfo.buyer_name || 'Direct Buyer',
        qtn_no: clientInfo.qtn_no || '',
        total: grandTotal,
        mode: effectiveQuotationMode,
        method: 'print_dialog',
      }, project.id).catch(() => {})
    }

    setTimeout(() => {
      document.title = prevTitle
    }, 2000)
  }

  const handleSave = () => {
    if (onSaveClientInfo) {
      const cleanTerms = termsList.map((t) => {
        if (t.id === '8' || t.title.toLowerCase().includes('freight') || t.title.toLowerCase().includes('transport')) {
          return { ...t, text: isTransportZero ? 'At your scope.' : t.text }
        }
        if (t.id === '10' || t.title.toLowerCase().includes('lodging') || t.title.toLowerCase().includes('loadging') || t.title.toLowerCase().includes('bording') || t.title.toLowerCase().includes('boarding')) {
          return { ...t, text: isAccommodationZero ? 'At your scope.' : t.text }
        }
        return t
      })
      const cleanClientInfo = {
        ...clientInfo,
        freight_terms: isTransportZero ? 'At your scope.' : clientInfo.freight_terms,
        lodging_terms: isAccommodationZero ? 'At your scope.' : clientInfo.lodging_terms,
      }
      setJustSaved(true)
      toast('Quotation changes saved successfully!')
      onSaveClientInfo(cleanClientInfo, lineItems, bodyItems, sets, {
        quotation_mode: effectiveQuotationMode,
        body_table_on_page_1: placeBodyOnPage1,
        body_placement: bodyPlacement,
        show_bodies_table: showBodiesTable,
        include_full_proposal: includeFullProposal,
        include_installation: includeInstallation,
        terms_list: cleanTerms,
        red_specs: redSpecs,
        proposal_specs: proposalSpecs,
        extra_amount: extraAmount,
        __profit_pct: profitPct,
      })

      // Record in Activity Log
      const token = tokenStore.get()
      if (token && effectiveOrgId) {
        api.logActivity(token, effectiveOrgId, 'pdf_saved', {
          name: project.name,
          buyer_name: clientInfo.buyer_name || 'Direct Buyer',
          qtn_no: clientInfo.qtn_no || '',
          total: grandTotal,
          mode: effectiveQuotationMode,
        }, project.id).catch(() => {})
      }
    }
  }

  // PDF Watermark Overlay Component
  const renderPdfWatermark = () => {
    if (businessProfile.show_watermark === false) return null
    const watermarkBg = businessProfile.watermark_url || MYRIAD_WATERMARK_DATA_URL
    if (!watermarkBg) return null
    return (
      <div
        className="pdf-watermark-overlay"
        style={{ backgroundImage: `url(${watermarkBg})` }}
      />
    )
  }

  // Common Header Component matching exact official layout with dynamic organisation business profile
  const renderOfficialHeader = () => {
    const logoSrc = businessProfile.logo_url || MYRIAD_LOGO_DATA_URL
    const companyName = businessProfile.company_name || 'MYRIAD STORAGE SYSTEM LLP'
    const regLine1 = businessProfile.registered_address_line1 || '702/A, AL-HUSSAIN ,MOMIN NAGAR, JOGESHWAR(W)'
    const regLine2 = businessProfile.registered_address_line2 || 'MUMBAI 400102. State : Maharashtra, Code : 27'
    const contactPhone = businessProfile.contact_phone || '9769377043'
    const altPhone = businessProfile.alternate_phone ? ` / ${businessProfile.alternate_phone}` : ''
    const gstNo = businessProfile.gst_no || '27ABJFM5468N1ZG'
    const panNo = businessProfile.pan_no ? ` | PAN: ${businessProfile.pan_no}` : ''
    const infoMail = businessProfile.info_email || 'info@myriadstoragesystem.com'
    const salesMail = businessProfile.sales_email || 'sales@myriadstoragesystem.com'
    const webUrl = businessProfile.website || 'www.myriadstoragesystem.com'
    const facLine1 = businessProfile.factory_address_line1 || 'Factory Address : Gala No 5, Vakan Compound, Near Dynamic Co, Vakan Pada Raod, Vasai Phata, Vasai East Dis.,'
    const facLine2 = businessProfile.factory_address_line2 || 'Palghar 401208.'

    return (
      <div className="pdf-header">
        <div className="pdf-header-top">
          <div className="pdf-logo-box">
            {logoSrc ? (
              <img src={logoSrc} alt={companyName} className="pdf-logo-img" />
            ) : (
              <div style={{ fontWeight: 800, fontSize: 18, color: '#000' }}>{companyName}</div>
            )}
            {businessProfile.tagline && (
              <div style={{ fontSize: '9px', color: '#333', marginTop: '2px', fontWeight: 500 }}>
                {businessProfile.tagline}
              </div>
            )}
          </div>
          <div className="pdf-company-info">
            <div>{regLine1}</div>
            <div>{regLine2}</div>
            <div>Contact : <a href={`tel:${contactPhone}`}>{contactPhone}</a>{altPhone}</div>
            <div>GST NO :{gstNo}{panNo}</div>
            {infoMail && <div><a href={`mailto:${infoMail}`}>{infoMail}</a></div>}
            {salesMail && <div><a href={`mailto:${salesMail}`}>{salesMail}</a></div>}
            {webUrl && (
              <div>
                <a
                  href={webUrl.startsWith('http') ? webUrl : `https://${webUrl}`}
                  target="_blank"
                  rel="noreferrer"
                >
                  {webUrl}
                </a>
              </div>
            )}
          </div>
        </div>
        <div className="pdf-factory-address">
          <div>{facLine1}</div>
          {facLine2 && <div>{facLine2}</div>}
        </div>
      </div>
    )
  }

  // Group installation body items by system
  const systemBodyGroups = useMemo(() => {
    interface BodyGroup {
      system_id: string
      system_name: string
      set_qty: number
      items: InstallationBodyItem[]
      unitTotal: number
      systemTotal: number
    }

    const groups: BodyGroup[] = []
    const groupMap = new Map<string, BodyGroup>()

    bodyItems.forEach((b, idx) => {
      const matchingSet = sets.find(s => s.id === b.system_id || (b.system_name && (s.system_title === b.system_name || s.set_label === b.system_name)))
      const defaultName = sets.length > 1 ? `SYSTEM ${String.fromCharCode(65 + idx)}` : (sets[0]?.system_title || sets[0]?.set_label || 'SYSTEM A')
      const sysName = b.system_name || defaultName
      const sysId = b.system_id || matchingSet?.id || sysName
      const setQty = b.set_qty || (matchingSet?.quantity ? Math.max(1, Number(matchingSet.quantity)) : 1)

      if (!groupMap.has(sysId)) {
        const newGroup: BodyGroup = {
          system_id: sysId,
          system_name: sysName,
          set_qty: setQty,
          items: [],
          unitTotal: 0,
          systemTotal: 0,
        }
        groupMap.set(sysId, newGroup)
        groups.push(newGroup)
      }

      const grp = groupMap.get(sysId)!
      grp.items.push(b)
      const itemUnitBody = b.unit_body !== undefined ? b.unit_body : (b.per_bay_qty * b.qty)
      grp.unitTotal += itemUnitBody
      grp.systemTotal += (b.total_body || (itemUnitBody * grp.set_qty))
    })

    return groups
  }, [bodyItems, sets])

  const totalSetsCount = useMemo(() => {
    if (systemBodyGroups.length > 0) {
      return systemBodyGroups.reduce((acc, g) => acc + g.set_qty, 0)
    }
    return sets.reduce((acc, s) => acc + (Math.max(1, Number(s.quantity) || 1)), 0) || 1
  }, [systemBodyGroups, sets])

  // Bodies table pagination chunks matching systemsPerPage (e.g. 3/p or 4/p applies to bodies as well!)
  const bodiesChunks = useMemo(() => {
    if (systemBodyGroups.length <= systemsPerPage || systemsPerPage >= 999) {
      return [{ chunkIndex: 0, groupsList: systemBodyGroups, isLastChunk: true, startItemIndex: 1 }]
    }
    const chunks: Array<{ chunkIndex: number; groupsList: typeof systemBodyGroups; isLastChunk: boolean; startItemIndex: number }> = []
    let currentItemIdx = 1
    for (let i = 0; i < systemBodyGroups.length; i += systemsPerPage) {
      const chunkGroups = systemBodyGroups.slice(i, i + systemsPerPage)
      chunks.push({
        chunkIndex: Math.floor(i / systemsPerPage),
        groupsList: chunkGroups,
        isLastChunk: i + systemsPerPage >= systemBodyGroups.length,
        startItemIndex: currentItemIdx,
      })
      chunkGroups.forEach((g) => {
        currentItemIdx += g.items.length
      })
    }
    return chunks
  }, [systemBodyGroups, systemsPerPage])

  const bodiesPagesCount = (showBodiesTable && bodyPlacement === 'dedicated') ? bodiesChunks.length : 0
  const bodiesPageCount = bodiesPagesCount

  const actualTotalPages = (commercialPagesCount > 1 || bodiesPagesCount > 0)
    ? (includeFullProposal ? (commercialPagesCount + bodiesPagesCount + 4) : (commercialPagesCount + bodiesPagesCount + 1))
    : totalPages

  // Installation Body Table Component
  const renderInstallationBodyTable = (
    extraClass = '',
    groupsToRender = systemBodyGroups,
    showGrandTotal = true,
    startItemIndex = 1
  ) => {
    let globalItemIndex = startItemIndex

    return (
      <table className={`pdf-body-table ${extraClass}`}>
        <thead>
          <tr>
            <th className="resizable-th" style={{ width: `${bodyColWidths.sr}px` }}>
              SR<br />NO
              <div
                className="table-col-resizer no-print"
                onMouseDown={(e) => startColResize('body', 'sr', e)}
                onTouchStart={(e) => startColResize('body', 'sr', e)}
                title="Click and drag border to adjust SR NO column width"
              />
            </th>
            <th className="resizable-th">
              Description of Goods
            </th>
            <th className="resizable-th" style={{ width: `${bodyColWidths.per_bay}px` }}>
              PER BAY QTY
              <div
                className="table-col-resizer no-print"
                onMouseDown={(e) => startColResize('body', 'per_bay', e)}
                onTouchStart={(e) => startColResize('body', 'per_bay', e)}
                title="Click and drag border to adjust Per Bay Qty column width"
              />
            </th>
            <th className="resizable-th" style={{ width: `${bodyColWidths.qty}px` }}>
              QTY
              <div
                className="table-col-resizer no-print"
                onMouseDown={(e) => startColResize('body', 'qty', e)}
                onTouchStart={(e) => startColResize('body', 'qty', e)}
                title="Click and drag border to adjust Qty column width"
              />
            </th>
            <th className="resizable-th" style={{ width: `${bodyColWidths.body}px` }}>
              INSTALLATION<br />BODY
              <div
                className="table-col-resizer no-print"
                onMouseDown={(e) => startColResize('body', 'body', e)}
                onTouchStart={(e) => startColResize('body', 'body', e)}
                title="Click and drag border to adjust Installation Body column width"
              />
            </th>
          </tr>
        </thead>
        <tbody>
          {groupsToRender.map((group) => {
            return (
              <React.Fragment key={group.system_id}>
                {/* System Name Header Row */}
                <tr className="system-body-header-row">
                  <td
                    colSpan={5}
                    className="font-bold text-center"
                    style={{
                      fontWeight: 800,
                      fontSize: '11px',
                      letterSpacing: '0.5px',
                      color: '#0f172a',
                      padding: '5px 8px',
                      borderTop: '2px solid #000000',
                      borderBottom: '1.5px solid #000000',
                      background: '#f8fafc',
                    }}
                  >
                    {group.system_name} ({group.set_qty} {group.set_qty > 1 ? 'SETS' : 'SET'})
                  </td>
                </tr>

                {/* Model Rows under this system */}
                {group.items.map((b) => {
                  const itemIndex = globalItemIndex++
                  const itemUnitBody = b.unit_body !== undefined ? b.unit_body : (b.per_bay_qty * b.qty)
                  return (
                    <tr key={b.id}>
                      <td className="text-center">{itemIndex})</td>
                      <td className="text-center">
                        <div style={{ fontWeight: 700, fontSize: '10.5px' }}>{b.title}</div>
                        {b.sub_label?.trim() ? <div className="tiny muted">{b.sub_label}</div> : null}
                      </td>
                      <td className="text-center font-mono">{b.per_bay_qty}</td>
                      <td className="text-center font-mono">{b.qty}</td>
                      <td className="text-center font-mono font-bold">{itemUnitBody}</td>
                    </tr>
                  )
                })}

                {/* System Subtotal: Line 1: Single Set Bodies */}
                <tr className="system-body-subtotal-row">
                  <td
                    colSpan={4}
                    className="text-right font-bold"
                    style={{
                      fontWeight: 700,
                      fontSize: '10px',
                      padding: '4px 8px',
                      letterSpacing: '0.2px',
                      color: '#0f172a',
                      background: '#f1f5f9',
                      borderBottom: group.set_qty > 1 ? '1px dashed #cbd5e1' : '1.5px solid #000000',
                    }}
                  >
                    <span>TOTAL {group.system_name} BODIES (1 SET) =</span>
                  </td>
                  <td
                    className="text-center font-mono font-bold"
                    style={{
                      fontWeight: 800,
                      fontSize: '11px',
                      borderBottom: group.set_qty > 1 ? '1px dashed #cbd5e1' : '1.5px solid #000000',
                      background: '#f1f5f9',
                      color: '#000000',
                    }}
                  >
                    {group.unitTotal}
                  </td>
                </tr>

                {/* System Subtotal: Line 2 (when sets > 1): Multiplied Bodies */}
                {group.set_qty > 1 && (
                  <tr className="system-body-multiplier-row">
                    <td
                      colSpan={4}
                      className="text-right font-bold"
                      style={{
                        fontWeight: 700,
                        fontSize: '10px',
                        padding: '4px 8px',
                        letterSpacing: '0.2px',
                        color: '#0f172a',
                        background: '#e2e8f0',
                        borderBottom: '1.5px solid #000000',
                      }}
                    >
                      <span>
                        TOTAL {group.system_name} BODIES: {group.unitTotal} BODIES × {group.set_qty} SETS =
                      </span>
                    </td>
                    <td
                      className="text-center font-mono font-bold"
                      style={{
                        fontWeight: 800,
                        fontSize: '11px',
                        borderBottom: '1.5px solid #000000',
                        background: '#e2e8f0',
                        color: '#000000',
                      }}
                    >
                      {group.systemTotal}
                    </td>
                  </tr>
                )}
              </React.Fragment>
            )
          })}

          {/* Continuation row if more systems follow on next page */}
          {!showGrandTotal && (
            <tr className="table-continuation-row">
              <td
                colSpan={5}
                style={{
                  background: '#f8fafc',
                  fontWeight: 700,
                  fontStyle: 'italic',
                  padding: '6px 10px',
                  textAlign: 'right',
                  borderTop: '1.5px solid #000000',
                  color: '#1e293b',
                  fontSize: '10px',
                }}
              >
                ... Continued on Next Page for Remaining Systems &gt;&gt;
              </td>
            </tr>
          )}

          {/* Grand Total Summary Row (Rendered on final bodies page) */}
          {showGrandTotal && (
            <tr className="total-body-row">
              <td
                colSpan={3}
                className="text-center font-bold"
                style={{ fontWeight: 700, fontSize: '10.5px', background: '#f8fafc' }}
              >
                TOTAL SETS: {totalSetsCount} {totalSetsCount > 1 ? 'SETS' : 'SET'}
              </td>
              <td
                className="text-right font-bold"
                style={{ fontWeight: 700, fontSize: '10.5px' }}
              >
                TOTAL BODIES
              </td>
              <td className="total-body-yellow text-center font-bold font-mono">
                <strong>{totalBodies}</strong>
              </td>
            </tr>
          )}
        </tbody>
      </table>
    )
  }

  let globalSlCounter = 1

  return createPortal(
    <div className="official-quotation-modal-overlay">

      {/* Mobile Top Header Navigation */}
      <div className="official-modal-mobile-header no-print">
        <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
          <button
            type="button"
            className="icon-btn xs"
            onClick={onClose}
            style={{ color: '#fff', background: '#334155', borderRadius: '4px', border: 'none', padding: '4px 8px' }}
            title="Close"
          >
            ✕
          </button>
          <div style={{ display: 'flex', flexDirection: 'column' }}>
            <span style={{ fontSize: '12.5px', fontWeight: 700, color: '#f8fafc', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis', maxWidth: '140px' }}>
              {clientInfo.qtn_no || 'Quotation'}
            </span>
            <span style={{ fontSize: '9.5px', color: '#94a3b8' }}>
              A4 • {actualTotalPages} Pages
            </span>
          </div>
        </div>

        <div className="mobile-nav-toggle">
          <button
            type="button"
            className={`mobile-toggle-btn ${mobileView === 'document' ? 'active' : ''}`}
            onClick={() => setMobileView('document')}
          >
            📄 Document
          </button>
          <button
            type="button"
            className={`mobile-toggle-btn ${mobileView === 'sidebar' ? 'active' : ''}`}
            onClick={() => setMobileView('sidebar')}
          >
            ⚙️ Edit &amp; Options
          </button>
        </div>
      </div>

      {/* =========================================================================
          THE EXACT OFFICIAL 5-PAGE PROPOSAL DOCUMENT
          ========================================================================= */}
      <div className={`official-document-scroll-container ${mobileView === 'sidebar' ? 'mobile-hidden' : ''}`} id="official-quotation-pages" style={{ '--pdf-table-font-size': tableFontSize, '--pdf-table-padding': tablePadding } as React.CSSProperties}>
        {isUnsavedWarning && !justSaved && (
          <div className="no-print" style={{ background: '#fef3c7', color: '#92400e', border: '1px solid #fde68a', padding: '8px 14px', fontSize: '12px', borderRadius: '6px', margin: '0 auto 12px auto', maxWidth: '210mm', display: 'flex', alignItems: 'center', gap: '8px' }}>
            <span>⚠️</span>
            <span><strong>Showing saved quotation data.</strong> Click <strong>Save</strong> in the calculator to apply unsaved edits.</span>
          </div>
        )}

        {/* ============================== COMMERCIAL QUOTATION PAGE(S) ============================== */}
        {commercialChunks.map((chunk, cIdx) => {
          const isFirstCommercialPage = cIdx === 0
          const isLastCommercialPage = chunk.isLastChunk
          const pageNum = cIdx + 1
          const chunkSets = chunk.setsList

          // Calculate starting Sl No for this chunk in unit_wise mode
          let startSlForChunk = 1
          for (let i = 0; i < chunk.startSetIndex; i++) {
            startSlForChunk += (sets[i]?.items?.length || 0)
          }

          return (
            <div key={`comm-page-${cIdx}`} className="official-pdf-page">
              {renderPdfWatermark()}

              {/* Header */}
              {renderOfficialHeader()}

              {/* Client & Metadata Section (on Page 1) or Continuation Banner (on subsequent commercial pages) */}
              {isFirstCommercialPage ? (
                <div className="pdf-client-section">
                  {clientInfo.qtn_no?.trim() ? (
                    <div className="pdf-qtn-line">
                      <strong>QTN NO:</strong> {clientInfo.qtn_no}
                    </div>
                  ) : null}
                  {clientInfo.date?.trim() ? (
                    <div className="pdf-date-line">
                      <strong>Date :</strong> {clientInfo.date}
                    </div>
                  ) : null}
                  {clientInfo.show_buyer_name !== false && clientInfo.buyer_name?.trim() ? (
                    <>
                      <div className="pdf-buyer-heading">
                        <strong>Buyer</strong>
                      </div>
                      <div className="pdf-buyer-name">
                        <strong>{clientInfo.buyer_name}</strong>
                      </div>
                    </>
                  ) : null}
                  {clientInfo.show_buyer_address !== false && clientInfo.buyer_address?.trim() ? (
                    <div className="pdf-buyer-address">
                      {clientInfo.buyer_address
                        .split('\n')
                        .filter((l) => l.trim().length > 0)
                        .map((line, i) => (
                          <div key={i}>{line}</div>
                        ))}
                    </div>
                  ) : null}
                  {clientInfo.show_buyer_gst !== false && clientInfo.buyer_gst?.trim() ? (
                    <div className="pdf-buyer-gst">
                      <strong>GSTIN:</strong> {clientInfo.buyer_gst}
                    </div>
                  ) : null}
                  {clientInfo.extra_fields && clientInfo.extra_fields.length > 0 && (
                    <div className="pdf-buyer-extra-fields">
                      {clientInfo.extra_fields.map((ef) => {
                        const label = ef.label?.trim()
                        const valText = ef.value?.trim()
                        if (!label && !valText) return null
                        return (
                          <div key={ef.id}>
                            {label ? <strong>{label} : </strong> : null}
                            <span>{valText}</span>
                          </div>
                        )
                      })}
                    </div>
                  )}
                </div>
              ) : (
                <div
                  className="pdf-continuation-banner"
                  style={{
                    display: 'flex',
                    justifyContent: 'space-between',
                    alignItems: 'center',
                    padding: '4px 8px',
                    background: '#f8fafc',
                    border: '1.5px solid #000000',
                    margin: '6px 0 10px 0',
                    fontSize: '10px',
                    fontWeight: 700,
                    color: '#0f172a',
                  }}
                >
                  <span>
                    QUOTATION CONTINUATION: {clientInfo.buyer_name || (clientInfo as any).company || 'CLIENT'}
                  </span>
                  <span>
                    SYSTEMS {chunk.startSetIndex + 1} TO {chunk.startSetIndex + chunkSets.length} OF {sets.length}
                  </span>
                </div>
              )}

              {/* Subject Line (Only on Page 1) */}
              {isFirstCommercialPage && clientInfo.show_subject !== false && clientInfo.subject?.trim() ? (
                <div className="pdf-subject-line">
                  Sub:- {clientInfo.subject}
                </div>
              ) : null}

              {/* Main Items Table */}
              <table className={`pdf-goods-table ${effectiveQuotationMode === 'unit_breakdown' ? 'unit-breakdown-table' : ''}`}>
                <thead>
                  {effectiveQuotationMode === 'unit_breakdown' ? (
                    <tr>
                      <th className="th-sl" style={{ width: '28px' }}>Sl.</th>
                      <th className="th-desc">Model &amp; Specifications</th>
                      <th className="th-qty" style={{ width: '38px' }}>Qty</th>
                      <th className="th-rate" style={{ width: '70px' }}>Unit Rate (₹)</th>
                      <th className="th-rate" style={{ width: '75px' }}>Extra Chrg (₹)</th>
                      <th className="th-rate" style={{ width: '80px' }}>Total Unit (₹)</th>
                      <th className="th-amount" style={{ width: '85px' }}>Total Price (₹)</th>
                      <th className="th-amount" style={{ width: '95px' }}>Price + Profit (₹)</th>
                    </tr>
                  ) : effectiveQuotationMode === 'unit_wise' ? (
                    <tr>
                      <th className="th-sl">Sl<br />No.</th>
                      <th className="th-desc">Description of Goods</th>
                      <th className="th-qty" style={{ width: '60px' }}>Quantity</th>
                      <th className="th-rate" style={{ width: '85px' }}>Rate</th>
                      <th className="th-amount" style={{ width: '100px' }}>Amount</th>
                    </tr>
                  ) : (
                    <tr>
                      <th className="th-sl resizable-th" style={{ width: `${goodsColWidths.sl}px` }}>
                        Sl<br />No.
                        <div
                          className="table-col-resizer no-print"
                          onMouseDown={(e) => startColResize('goods', 'sl', e)}
                          onTouchStart={(e) => startColResize('goods', 'sl', e)}
                          title="Click and drag border to adjust Sl No column width"
                        />
                      </th>
                      <th className="th-desc resizable-th">
                        Description of Goods
                      </th>
                      <th className="th-qty resizable-th" style={{ width: `${goodsColWidths.qty}px` }}>
                        Quantity
                        <div
                          className="table-col-resizer no-print"
                          onMouseDown={(e) => startColResize('goods', 'qty', e)}
                          onTouchStart={(e) => startColResize('goods', 'qty', e)}
                          title="Click and drag border to adjust Quantity column width"
                        />
                      </th>
                      <th className="th-set resizable-th" style={{ width: `${goodsColWidths.set}px` }}>
                        {isMultiSystem ? 'Set' : ''}
                        <div
                          className="table-col-resizer no-print"
                          onMouseDown={(e) => startColResize('goods', 'set', e)}
                          onTouchStart={(e) => startColResize('goods', 'set', e)}
                          title="Click and drag border to adjust Set column width"
                        />
                      </th>
                      <th className="th-rate resizable-th" style={{ width: `${goodsColWidths.rate}px` }}>
                        Rate
                        <div
                          className="table-col-resizer no-print"
                          onMouseDown={(e) => startColResize('goods', 'rate', e)}
                          onTouchStart={(e) => startColResize('goods', 'rate', e)}
                          title="Click and drag border to adjust Rate column width"
                        />
                      </th>
                      <th className="th-amount resizable-th" style={{ width: `${goodsColWidths.amount}px` }}>
                        Amount
                        <div
                          className="table-col-resizer no-print"
                          onMouseDown={(e) => startColResize('goods', 'amount', e)}
                          onTouchStart={(e) => startColResize('goods', 'amount', e)}
                          title="Click and drag border to adjust Amount column width"
                        />
                      </th>
                    </tr>
                  )}
                </thead>
                <tbody>
                  {effectiveQuotationMode === 'unit_breakdown' ? (
                    <>
                      {breakdownResult.rows.map((r, idx) => (
                        <tr key={r.model_code} className="item-row">
                          <td className="td-sl text-center">{idx + 1}</td>
                          <td className="td-desc text-center">
                            <strong className="model-title">{r.model_code}</strong>
                            {r.model_name?.trim() ? <div style={{ fontSize: '9px', color: '#64748b' }}>{r.model_name}</div> : null}
                            {cleanDimensions(r.dimensions) ? <div className="model-dims">{cleanDimensions(r.dimensions)}</div> : null}
                          </td>
                          <td className="td-qty text-center font-bold">{r.qty}</td>
                          <td className="td-rate text-center font-mono" style={{ textAlign: 'center', verticalAlign: 'middle' }}>{r.unit_rate.toFixed(2)}</td>
                          <td className="td-rate text-center font-mono" style={{ textAlign: 'center', verticalAlign: 'middle' }}>{r.transportation.toFixed(2)}</td>
                          <td className="td-rate text-center font-mono" style={{ textAlign: 'center', verticalAlign: 'middle' }}>{r.total_unit_rate.toFixed(2)}</td>
                          <td className="td-amount text-center font-mono" style={{ textAlign: 'center', verticalAlign: 'middle' }}>{r.total_price.toFixed(2)}</td>
                          <td className="td-amount text-center font-mono font-bold" style={{ textAlign: 'center', verticalAlign: 'middle' }}>{r.total_price_with_profit.toFixed(2)}</td>
                        </tr>
                      ))}
                    </>
                  ) : effectiveQuotationMode === 'unit_wise' ? (
                    <>
                      {sets.length > 1 ? (
                        (() => {
                          let currentSl = startSlForChunk
                          return chunkSets.map((s, relIdx) => {
                            const sIdx = chunk.startSetIndex + relIdx
                            const nextChar = String.fromCharCode(65 + sIdx)
                            const sysTitle = (s.system_title || `SYSTEM ${nextChar}`).trim()
                            const sysItems = s.items && s.items.length > 0 ? s.items : []
                            const sysQty = Number(s.quantity) || 1

                            const sysDoubleDrive = systemHasDoubleDrive(s)
                            const showSysDoubleDrive = sets.length > 1 && sysDoubleDrive && !allSystemsHaveDoubleDrive
                            const sysTrack = allSystemsTrackCounts[sIdx] || projectTrackCount || 2
                            const showSysTrack = sets.length > 1 && !allTracksIdentical

                            const sysNotes: string[] = []
                            if (showSysDoubleDrive) {
                              sysNotes.push('DOUBLE DRIVE WITH REDUCTION WHEEL')
                            }
                            if (showSysTrack) {
                              sysNotes.push(`${sysTrack} ${sysTrack === 1 ? 'TRACK' : 'TRACKS'}`)
                            }

                            return (
                              <React.Fragment key={s.id || sIdx}>
                                <tr className="system-title-row">
                                  <td className="td-sl"></td>
                                  <td
                                    colSpan={4}
                                    className="td-desc font-bold text-center"
                                    style={{
                                      fontWeight: 800,
                                      fontSize: '11px',
                                      letterSpacing: '0.5px',
                                      color: '#0f172a',
                                      padding: '5px 8px',
                                      textAlign: 'center',
                                      border: '2px solid #000000',
                                      background: '#f8fafc',
                                    }}
                                  >
                                    <span>{sysTitle} {s.set_label ? `(${s.set_label})` : `(${sysQty} ${s.unit || 'SET'})`}</span>
                                  </td>
                                </tr>
                                {sysItems.map((it, itIdx) => {
                                  const slNo = currentSl++
                                  const displayRate = (it as any).final_unit_rate !== undefined ? (it as any).final_unit_rate : (it.rate || 0)
                                  const itemUnitQty = Number(it.unit_qty) || 1
                                  const totalItemQty = itemUnitQty * sysQty
                                  const displayAmount = (it as any).final_amount !== undefined
                                    ? (it as any).final_amount * sysQty
                                    : Math.round(Number(displayRate) * totalItemQty)

                                  return (
                                    <tr key={`${s.id || sIdx}-${it.id || itIdx}`} className="item-row">
                                      <td className="td-sl text-center font-bold">{slNo}</td>
                                      <td className="td-desc text-center">
                                        <strong className="model-title">{it.model_name || it.model_code}</strong>
                                        {cleanDimensions(it.dimensions) ? (
                                          <div className="model-dims">{cleanDimensions(it.dimensions)}</div>
                                        ) : null}
                                      </td>
                                      <td className="td-qty text-center font-bold" style={{ verticalAlign: 'middle', fontSize: '10.5px' }}>
                                        {totalItemQty}
                                      </td>
                                      <td
                                        className="td-rate text-center font-mono"
                                        style={{ textAlign: 'center', verticalAlign: 'middle', fontWeight: 600, fontSize: '10.5px' }}
                                      >
                                        {Number(displayRate).toFixed(2)}
                                      </td>
                                      <td
                                        className="td-amount text-center font-mono"
                                        style={{ textAlign: 'center', verticalAlign: 'middle', fontWeight: 600, fontSize: '10.5px' }}
                                      >
                                        {Number(displayAmount).toFixed(2)}
                                      </td>
                                    </tr>
                                  )
                                })}
                                {sysNotes.length > 0 && (
                                  <tr key={`notes-${s.id || sIdx}`} className="system-notes-row">
                                    <td className="td-sl"></td>
                                    <td
                                      colSpan={4}
                                      className="td-desc text-center"
                                      style={{
                                        padding: '4px 8px',
                                        textAlign: 'center',
                                        verticalAlign: 'middle',
                                        background: '#ffffff',
                                      }}
                                    >
                                      <div
                                        style={{
                                          display: 'flex',
                                          flexDirection: 'column',
                                          alignItems: 'center',
                                          gap: '2px',
                                        }}
                                      >
                                        {sysNotes.map((note, nIdx) => (
                                          <span key={nIdx} style={{ color: '#ff0000', fontWeight: 700, fontSize: '9.5px', letterSpacing: '0.4px', textTransform: 'uppercase' }}>
                                            {note}
                                          </span>
                                        ))}
                                      </div>
                                    </td>
                                  </tr>
                                )}
                              </React.Fragment>
                            )
                          })
                        })()
                      ) : (
                        lineItems.map((it, itIdx) => {
                          const currentSl = itIdx + 1
                          const displayRate = lineItemDisplayAmounts[itIdx]?.rate ?? Number(it.rate)
                          const displayAmount = lineItemDisplayAmounts[itIdx]?.amount ?? Number(it.amount)
                          return (
                            <tr key={it.id} className="item-row">
                              <td className="td-sl text-center font-bold">{currentSl}</td>
                              <td className="td-desc text-center">
                                <strong className="model-title">{it.model_name}</strong>
                                {cleanDimensions(it.dimensions) ? (
                                  <div className="model-dims">{cleanDimensions(it.dimensions)}</div>
                                ) : null}
                              </td>
                              <td className="td-qty text-center font-bold" style={{ verticalAlign: 'middle', fontSize: '10.5px' }}>
                                {it.quantity}
                              </td>
                              <td
                                className="td-rate text-center font-mono"
                                style={{ textAlign: 'center', verticalAlign: 'middle', fontWeight: 600, fontSize: '10.5px' }}
                              >
                                {Number(displayRate).toFixed(2)}
                              </td>
                              <td
                                className="td-amount text-center font-mono"
                                style={{ textAlign: 'center', verticalAlign: 'middle', fontWeight: 600, fontSize: '10.5px' }}
                              >
                                {Number(displayAmount).toFixed(2)}
                              </td>
                              <span style={{ display: 'none' }}>{Number(it.rate).toFixed(2)}{Number(it.amount).toFixed(2)}</span>
                            </tr>
                          )
                        })
                      )}
                    </>
                  ) : (
                    <>
                      {chunkSets.map((s, relIdx) => {
                        const sIdx = chunk.startSetIndex + relIdx
                        const nextChar = String.fromCharCode(65 + sIdx)
                        const systemTitle = (!isMultiSystem ? '' : (s.system_title !== undefined ? s.system_title : (s.set_label?.startsWith('SYSTEM') ? s.set_label : `SYSTEM ${nextChar}`))).trim()
                        const hasItems = s.items && s.items.length > 0

                        const sysDoubleDrive = systemHasDoubleDrive(s)
                        const showSysDoubleDrive = sets.length > 1 && sysDoubleDrive && !allSystemsHaveDoubleDrive
                        const sysTrack = allSystemsTrackCounts[sIdx] || projectTrackCount || 2
                        const showSysTrack = sets.length > 1 && !allTracksIdentical

                        const sysNotes: string[] = []
                        if (showSysDoubleDrive) {
                          sysNotes.push('DOUBLE DRIVE WITH REDUCTION WHEEL')
                        }
                        if (showSysTrack) {
                          sysNotes.push(`${sysTrack} ${sysTrack === 1 ? 'TRACK' : 'TRACKS'}`)
                        }

                        const totalRowSpan = (hasItems ? s.items.length : 0) + (systemTitle ? 1 : 0) + (sysNotes.length > 0 ? 1 : 0)
                        const displaySetRate = setDisplayAmounts[sIdx]?.rate ?? Number(s.rate)
                        const displaySetAmount = setDisplayAmounts[sIdx]?.amount ?? Number(s.amount)

                        return (
                          <React.Fragment key={s.id}>
                            {systemTitle ? (
                              <tr key={`hdr-${s.id}`} className="system-title-row">
                                <td className="td-sl"></td>
                                <td
                                  className="td-desc font-bold text-center"
                                  style={{
                                    fontWeight: 800,
                                    textAlign: 'center',
                                    fontSize: '11px',
                                    letterSpacing: '0.5px',
                                    border: '2px solid #000000',
                                    background: '#f8fafc',
                                    color: '#000000',
                                  }}
                                >
                                  <div>{systemTitle}</div>
                                </td>
                                <td className="td-qty"></td>
                                <td
                                  className="td-set text-center font-bold"
                                  rowSpan={totalRowSpan}
                                  style={{ verticalAlign: 'middle', fontWeight: 700, fontSize: '11px', background: '#f8fafc' }}
                                >
                                  <div>{s.quantity}</div>
                                  <div style={{ fontSize: '9px', color: '#475569' }}>{s.unit || 'SET'}</div>
                                </td>
                                <td
                                  className="td-rate text-center font-mono"
                                  rowSpan={totalRowSpan}
                                  style={{ textAlign: 'center', verticalAlign: 'middle', fontWeight: 600, fontSize: '10.5px' }}
                                >
                                  {Number(displaySetRate).toLocaleString('en-IN', {
                                    minimumFractionDigits: Number(displaySetRate) % 1 === 0 ? 0 : 2,
                                    maximumFractionDigits: 2,
                                    useGrouping: false,
                                  })}
                                </td>
                                <td
                                  className="td-amount text-center font-mono"
                                  rowSpan={totalRowSpan}
                                  style={{ textAlign: 'center', verticalAlign: 'middle', fontWeight: 600, fontSize: '10.5px' }}
                                >
                                  {Number(displaySetAmount).toFixed(2)}
                                </td>
                              </tr>
                            ) : null}

                            {(s.items || []).map((it, itIdx) => {
                              const currentSl = itIdx + 1
                              const systemBreakdownRows = s.breakdown?.rows || breakdownResult.rows
                              const matchRow = systemBreakdownRows.find((r: any) =>
                                (it.model_code && r.model_code === it.model_code) ||
                                (it.model_name && r.model_name && it.model_name.trim().toLowerCase() === r.model_name.trim().toLowerCase()) ||
                                (it.model_name && r.model_code && it.model_name.includes(r.model_code))
                              )
                              const itemRate = matchRow
                                ? (hasDiscount ? matchRow.unit_rate_with_profit : matchRow.final_unit_rate)
                                : (Number(it.rate) || (Number(it.amount) / Math.max(1, Number(it.unit_qty) || 1)) || 0)
                              const itemAmount = matchRow
                                ? (hasDiscount ? matchRow.total_price_with_profit : matchRow.final_amount)
                                : (Number(it.amount) || (itemRate * (Number(it.unit_qty) || 1)) || 0)

                              return (
                                <tr key={it.id} className="item-row">
                                  <td className="td-sl text-center font-bold">{currentSl}</td>
                                  <td className="td-desc text-center">
                                    <strong className="model-title">{it.model_name}</strong>
                                    {cleanDimensions(it.dimensions) ? (
                                      <div className="model-dims">{cleanDimensions(it.dimensions)}</div>
                                    ) : null}
                                  </td>
                                  <td className="td-qty text-center font-bold" style={{ verticalAlign: 'middle', fontSize: '10.5px' }}>
                                    {it.unit_qty}
                                  </td>
                                  {!systemTitle && itIdx === 0 && (
                                    <>
                                      <td
                                        className="td-set text-center font-bold"
                                        rowSpan={totalRowSpan}
                                        style={{ verticalAlign: 'middle', fontWeight: 700, fontSize: '11px', background: '#f8fafc' }}
                                      >
                                        <div>{s.quantity}</div>
                                        <div style={{ fontSize: '9px', color: '#475569' }}>{s.unit || 'SET'}</div>
                                      </td>
                                      <td
                                        className="td-rate text-center font-mono"
                                        rowSpan={totalRowSpan}
                                        style={{ textAlign: 'center', verticalAlign: 'middle', fontWeight: 600, fontSize: '10.5px' }}
                                      >
                                        {Number(displaySetRate).toLocaleString('en-IN', {
                                          minimumFractionDigits: Number(displaySetRate) % 1 === 0 ? 0 : 2,
                                          maximumFractionDigits: 2,
                                          useGrouping: false,
                                        })}
                                      </td>
                                      <td
                                        className="td-amount text-center font-mono"
                                        rowSpan={totalRowSpan}
                                        style={{ textAlign: 'center', verticalAlign: 'middle', fontWeight: 600, fontSize: '10.5px' }}
                                      >
                                        {Number(displaySetAmount).toFixed(2)}
                                      </td>
                                    </>
                                  )}
                                  <span style={{ display: 'none' }}>{Number(itemRate).toFixed(2)}{Number(itemAmount).toFixed(2)}</span>
                                </tr>
                              )
                            })}

                            {sysNotes.length > 0 && (
                              <tr key={`notes-${s.id || sIdx}`} className="system-notes-row">
                                <td className="td-sl"></td>
                                <td
                                  className="td-desc text-center"
                                  style={{
                                    padding: '4px 8px',
                                    textAlign: 'center',
                                    verticalAlign: 'middle',
                                    background: '#ffffff',
                                  }}
                                >
                                  <div style={{ fontSize: '9.5px', fontWeight: 700, color: '#dc2626', letterSpacing: '0.4px', lineHeight: 1.4 }}>
                                    {sysNotes.map((note, nIdx) => (
                                      <div key={nIdx}>{note}</div>
                                    ))}
                                  </div>
                                </td>
                                <td className="td-qty"></td>
                              </tr>
                            )}
                          </React.Fragment>
                        )
                      })}
                    </>
                  )}

                  {/* If not the last commercial page, show continuation prompt row */}
                  {!isLastCommercialPage && (
                    <tr className="table-continuation-row">
                      <td
                        colSpan={effectiveQuotationMode === 'unit_breakdown' ? 8 : effectiveQuotationMode === 'unit_wise' ? 5 : 6}
                        style={{
                          background: '#f8fafc',
                          fontWeight: 700,
                          fontStyle: 'italic',
                          padding: '6px 10px',
                          textAlign: 'right',
                          borderTop: '1.5px solid #000000',
                          color: '#1e293b',
                          fontSize: '10px',
                        }}
                      >
                        ... Continued on Next Page for Remaining Systems &gt;&gt;
                      </td>
                    </tr>
                  )}

                  {/* Red Specs & Summary Rows (Only on Last Commercial Page) */}
                  {isLastCommercialPage && (
                    <>
                      <tr className="summary-split-row">
                        <td
                          colSpan={effectiveQuotationMode === 'unit_breakdown' ? 5 : 3}
                          rowSpan={hasDiscount ? 6 : 4}
                          className="td-red-specs"
                          title="Click to edit Description of Goods Specifications"
                          onClick={() => setActiveSidebarTab('options')}
                        >
                          <div className="red-specs-content" style={{ cursor: 'pointer' }}>
                            {redSpecs.map((spec, i) => (
                              <div key={i}>{spec}</div>
                            ))}
                          </div>
                        </td>
                        <td colSpan={effectiveQuotationMode === 'unit_wise' ? 1 : 2} className="td-sum-label text-center" style={{ textAlign: 'center', verticalAlign: 'middle' }}>TOTAL</td>
                        <td className="td-sum-val text-center" style={{ textAlign: 'center', verticalAlign: 'middle' }}>{(hasDiscount ? totalBeforeDiscount : subtotal).toFixed(2)}</td>
                      </tr>
                      {hasDiscount && (
                        <>
                          <tr className="summary-split-row">
                            <td colSpan={effectiveQuotationMode === 'unit_wise' ? 1 : 2} className="td-sum-label text-center" style={{ textAlign: 'center', verticalAlign: 'middle' }}>DISCOUNT</td>
                            <td className="td-sum-val text-center" style={{ textAlign: 'center', verticalAlign: 'middle' }}>{totalDiscountAmount.toFixed(2)}</td>
                          </tr>
                          <tr className="summary-split-row">
                            <td colSpan={effectiveQuotationMode === 'unit_wise' ? 1 : 2} className="td-sum-label text-center" style={{ textAlign: 'center', verticalAlign: 'middle' }}>AFTER DISCOUNT</td>
                            <td className="td-sum-val text-center" style={{ textAlign: 'center', verticalAlign: 'middle' }}>{subtotal.toFixed(2)}</td>
                          </tr>
                        </>
                      )}
                      <tr className="summary-split-row">
                        <td colSpan={effectiveQuotationMode === 'unit_wise' ? 1 : 2} className="td-sum-label text-center" style={{ textAlign: 'center', verticalAlign: 'middle' }}>GST 18%</td>
                        <td className="td-sum-val text-center" style={{ textAlign: 'center', verticalAlign: 'middle' }}>{gst18.toFixed(2)}</td>
                      </tr>
                      <tr className="summary-split-row">
                        <td colSpan={effectiveQuotationMode === 'unit_wise' ? 1 : 2} className="td-sum-label text-center" style={{ textAlign: 'center', verticalAlign: 'middle' }}>ROUND OFF</td>
                        <td className="td-sum-val text-center" style={{ textAlign: 'center', verticalAlign: 'middle' }}>{Math.abs(Number(roundOff)) > 0.001 ? roundOff : ''}</td>
                      </tr>
                      <tr className="summary-split-row grand-total-row">
                        <td colSpan={effectiveQuotationMode === 'unit_wise' ? 1 : 2} className="td-sum-label grand text-center" style={{ textAlign: 'center', verticalAlign: 'middle' }}>GRAND TOTAL</td>
                        <td className="td-sum-val grand text-center" style={{ textAlign: 'center', verticalAlign: 'middle' }}>{grandTotal.toFixed(2)}</td>
                      </tr>
                    </>
                  )}
                </tbody>
              </table>

              {/* Amount in Words (Only on Last Commercial Page) */}
              {isLastCommercialPage && (
                <div className="pdf-words-bar">
                  {amountInWords}
                </div>
              )}

              {/* Installation Body Table on Page 1 (Only on Last Commercial Page) */}
              {isLastCommercialPage && showBodiesTable && bodyPlacement === 'commercial' && renderInstallationBodyTable('on-page-1')}

              {/* Commercial Page Footer */}
              <div className="pdf-footer">
                {pageNum} of {actualTotalPages}
                {/* 1 of {totalPages} */}
              </div>
            </div>
          )
        })}

        {/* ============================== DEDICATED INSTALLATION BODIES BREAKDOWN PAGE(S) ============================== */}
        {showBodiesTable && bodyPlacement === 'dedicated' && bodiesChunks.map((bChunk, bIdx) => {
          const isLastBodiesPage = bChunk.isLastChunk
          const currentBodiesPageNum = commercialPagesCount + bIdx + 1

          return (
            <div key={`bodies-page-${bIdx}`} className="official-pdf-page">
              {renderPdfWatermark()}

              {/* Header */}
              {renderOfficialHeader()}

              <div
                className="pdf-continuation-banner"
                style={{
                  display: 'flex',
                  justifyContent: 'space-between',
                  alignItems: 'center',
                  padding: '5px 10px',
                  background: '#f8fafc',
                  border: '1.5px solid #000000',
                  margin: '6px 0 12px 0',
                  fontSize: '11px',
                  fontWeight: 800,
                  color: '#0f172a',
                  letterSpacing: '0.4px',
                }}
              >
                <span>
                  ANNEXURE: INSTALLATION BODIES BREAKDOWN ${bodiesChunks.length > 1 ? `(PART ${bIdx + 1} OF ${bodiesChunks.length})` : ''}
                </span>
                <span className="font-mono" style={{ fontSize: '10px', fontWeight: 700 }}>
                  {clientInfo.qtn_no || ''}
                </span>
              </div>

              <div style={{ display: 'flex', justifyContent: 'center', width: '100%' }}>
                {renderInstallationBodyTable('dedicated-page-table', bChunk.groupsList, isLastBodiesPage, bChunk.startItemIndex)}
              </div>

              <div className="pdf-footer">
                {currentBodiesPageNum} of {actualTotalPages}
              </div>
            </div>
          )
        })}

        {/* ============================== PAGE 2: TERMS & CONDITIONS ============================== */}
        <div className="official-pdf-page">
          {renderPdfWatermark()}

          {/* Page 2 Header */}
          {renderOfficialHeader()}

          {/* If NOT on Page 1, render Installation Body table here on Page 2 */}
          {showBodiesTable && bodyPlacement === 'terms' && renderInstallationBodyTable('mt-2')}

          {/* Terms & Conditions Block */}
          <div className={`pdf-terms-block ${showBodiesTable && bodyPlacement === 'terms' ? 'mt-3' : 'mt-2'}`}>
            <div className="pdf-terms-heading">
              TERMS &amp; CONDITION
            </div>
            <ol className="pdf-terms-list">
              {termsList
                .filter((t) => (t.text && t.text.trim().length > 0) || (t.title && t.title.trim().length > 0))
                .map((t, idx) => {
                  const isFreight = t.id === '8' || t.title.toLowerCase().includes('freight') || t.title.toLowerCase().includes('transport')
                  const isLodging = t.id === '10' || t.title.toLowerCase().includes('lodging') || t.title.toLowerCase().includes('loadging') || t.title.toLowerCase().includes('bording') || t.title.toLowerCase().includes('boarding')
                  const effectiveText = isLodging
                    ? (t.text || 'At your scope.')
                    : (isFreight && isTransportZero)
                      ? 'At your scope.'
                      : (t.text || t.title)
                  const hasText = Boolean(effectiveText && effectiveText.trim().length > 0)
                  const hasTitle = Boolean(t.title && t.title.trim().length > 0)
                  return (
                    <li key={t.id}>
                      <span className="term-num">{idx + 1})</span>
                      <span className="term-text">
                        {hasTitle && hasText ? (
                          <>
                            <strong>{t.title} </strong>
                            {effectiveText}
                          </>
                        ) : (
                          effectiveText
                        )}
                      </span>
                    </li>
                  )
                })}
            </ol>
          </div>

          <div className="pdf-footer">
            {commercialPagesCount + 1} of {actualTotalPages}
            {/* 2 of {totalPages} */}
          </div>
        </div>

        {/* ============================== PAGES 3 TO 5 (FULL 5-PAGE TECHNICAL PROPOSAL) ============================== */}
        {includeFullProposal && (
          <>
            {/* Page 3: Deluxe Model Compactor Technical Specification & Mechanical Assembly (Merged 3 & 4) */}
            <div className="official-pdf-page">
              {renderPdfWatermark()}
              {renderOfficialHeader()}

              <div className="pdf-spec-container">
                {/* PRETREATMENT PROCESS &amp; PARAMETERS */}
                <div
                  className="pdf-spec-main-title"
                  contentEditable
                  suppressContentEditableWarning
                  onBlur={(e) => {
                    const text = e.currentTarget.innerText
                    setProposalSpecs((prev) => ({ ...prev, page3: { ...prev.page3, mainTitle: text } }))
                  }}
                  title="Click to edit text"
                >
                  {proposalSpecs.page3.mainTitle}
                </div>
                <div
                  className="pdf-spec-sub-title"
                  contentEditable
                  suppressContentEditableWarning
                  onBlur={(e) => {
                    const text = e.currentTarget.innerText
                    setProposalSpecs((prev) => ({ ...prev, page3: { ...prev.page3, subTitle: text } }))
                  }}
                  title="Click to edit text"
                >
                  {proposalSpecs.page3.subTitle}
                </div>

                <div
                  className="pdf-spec-sec-title"
                  contentEditable
                  suppressContentEditableWarning
                  onBlur={(e) => {
                    const text = e.currentTarget.innerText
                    setProposalSpecs((prev) => ({ ...prev, page3: { ...prev.page3, mobileRacksTitle: text } }))
                  }}
                  title="Click to edit text"
                >
                  {proposalSpecs.page3.mobileRacksTitle}
                </div>
                <div
                  className="pdf-spec-para"
                  contentEditable
                  suppressContentEditableWarning
                  onBlur={(e) => {
                    const text = e.currentTarget.innerText
                    setProposalSpecs((prev) => ({ ...prev, page3: { ...prev.page3, mobileRacksIntro: text } }))
                  }}
                  title="Click to edit text"
                >
                  {proposalSpecs.page3.mobileRacksIntro}
                </div>

                <div
                  className="pdf-spec-sec-title"
                  contentEditable
                  suppressContentEditableWarning
                  onBlur={(e) => {
                    const text = e.currentTarget.innerText
                    setProposalSpecs((prev) => ({ ...prev, page3: { ...prev.page3, undercarriageTitle: text } }))
                  }}
                  title="Click to edit text"
                >
                  {proposalSpecs.page3.undercarriageTitle}
                </div>
                <div className="pdf-spec-row">
                  <strong>BASE FRAME : </strong>
                  <span
                    contentEditable
                    suppressContentEditableWarning
                    onBlur={(e) => {
                      const text = e.currentTarget.innerText
                      setProposalSpecs((prev) => ({ ...prev, page3: { ...prev.page3, baseFrame: text } }))
                    }}
                    title="Click to edit text"
                  >
                    {proposalSpecs.page3.baseFrame}
                  </span>
                </div>
                <div className="pdf-spec-row">
                  <strong>TRACKS: </strong>
                  <span
                    contentEditable
                    suppressContentEditableWarning
                    onBlur={(e) => {
                      const text = e.currentTarget.innerText
                      setProposalSpecs((prev) => ({ ...prev, page3: { ...prev.page3, tracks: text } }))
                    }}
                    title="Click to edit text"
                  >
                    {proposalSpecs.page3.tracks}
                  </span>
                </div>

                <div
                  className="pdf-spec-sec-title"
                  contentEditable
                  suppressContentEditableWarning
                  onBlur={(e) => {
                    const text = e.currentTarget.innerText
                    setProposalSpecs((prev) => ({ ...prev, page3: { ...prev.page3, superStructureTitle: text } }))
                  }}
                  title="Click to edit text"
                >
                  {proposalSpecs.page3.superStructureTitle}
                </div>
                <div
                  className="pdf-spec-para"
                  contentEditable
                  suppressContentEditableWarning
                  onBlur={(e) => {
                    const text = e.currentTarget.innerText
                    setProposalSpecs((prev) => ({ ...prev, page3: { ...prev.page3, superStructureIntro: text } }))
                  }}
                  title="Click to edit text"
                >
                  {proposalSpecs.page3.superStructureIntro}
                </div>

                <div className="pdf-spec-row">
                  <strong>SHELVES: </strong>
                  <span
                    contentEditable
                    suppressContentEditableWarning
                    onBlur={(e) => {
                      const text = e.currentTarget.innerText
                      setProposalSpecs((prev) => ({ ...prev, page3: { ...prev.page3, shelves: text } }))
                    }}
                    title="Click to edit text"
                  >
                    {proposalSpecs.page3.shelves}
                  </span>
                </div>
                <div className="pdf-spec-row">
                  <strong>VERTICAL PANEL: </strong>
                  <span
                    contentEditable
                    suppressContentEditableWarning
                    onBlur={(e) => {
                      const text = e.currentTarget.innerText
                      setProposalSpecs((prev) => ({ ...prev, page3: { ...prev.page3, verticalPanel: text } }))
                    }}
                    title="Click to edit text"
                  >
                    {proposalSpecs.page3.verticalPanel}
                  </span>
                </div>
                <div className="pdf-spec-row">
                  <strong>MIDDLE PANEL: </strong>
                  <span
                    contentEditable
                    suppressContentEditableWarning
                    onBlur={(e) => {
                      const text = e.currentTarget.innerText
                      setProposalSpecs((prev) => ({ ...prev, page3: { ...prev.page3, middlePanel: text } }))
                    }}
                    title="Click to edit text"
                  >
                    {proposalSpecs.page3.middlePanel}
                  </span>
                </div>
                <div className="pdf-spec-row">
                  <strong>TOP PANEL: </strong>
                  <span
                    contentEditable
                    suppressContentEditableWarning
                    onBlur={(e) => {
                      const text = e.currentTarget.innerText
                      setProposalSpecs((prev) => ({ ...prev, page3: { ...prev.page3, topPanel: text } }))
                    }}
                    title="Click to edit text"
                  >
                    {proposalSpecs.page3.topPanel}
                  </span>
                </div>
                <div className="pdf-spec-row">
                  <strong>CENTRE PARTITION: </strong>
                  <span
                    contentEditable
                    suppressContentEditableWarning
                    onBlur={(e) => {
                      const text = e.currentTarget.innerText
                      setProposalSpecs((prev) => ({ ...prev, page3: { ...prev.page3, centrePartition: text } }))
                    }}
                    title="Click to edit text"
                  >
                    {proposalSpecs.page3.centrePartition}
                  </span>
                </div>
                <div className="pdf-spec-row">
                  <strong>END PANEL: </strong>
                  <span
                    contentEditable
                    suppressContentEditableWarning
                    onBlur={(e) => {
                      const text = e.currentTarget.innerText
                      setProposalSpecs((prev) => ({ ...prev, page3: { ...prev.page3, endPanel: text } }))
                    }}
                    title="Click to edit text"
                  >
                    {proposalSpecs.page3.endPanel}
                  </span>
                </div>
                <div className="pdf-spec-row">
                  <strong>CENTRAL LOCKING: </strong>
                  <span
                    contentEditable
                    suppressContentEditableWarning
                    onBlur={(e) => {
                      const text = e.currentTarget.innerText
                      setProposalSpecs((prev) => ({ ...prev, page3: { ...prev.page3, centralLocking: text } }))
                    }}
                    title="Click to edit text"
                  >
                    {proposalSpecs.page3.centralLocking}
                  </span>
                </div>

                {/* Merged from Page 4: Mechanical & Drive Specs - MECHANICAL &amp; DRIVE ASSEMBLY: */}
                <div
                  className="pdf-spec-sec-title"
                  style={{ marginTop: '7px' }}
                  contentEditable
                  suppressContentEditableWarning
                  onBlur={(e) => {
                    const text = e.currentTarget.innerText
                    setProposalSpecs((prev) => ({ ...prev, page3: { ...prev.page3, mechanicalTitle: text } }))
                  }}
                  title="Click to edit text"
                >
                  {proposalSpecs.page3.mechanicalTitle}
                </div>
                <div className="pdf-spec-row">
                  <strong>POWDER COATING: </strong>
                  <span
                    contentEditable
                    suppressContentEditableWarning
                    onBlur={(e) => {
                      const text = e.currentTarget.innerText
                      setProposalSpecs((prev) => ({ ...prev, page3: { ...prev.page3, powderCoating: text } }))
                    }}
                    title="Click to edit text"
                  >
                    {proposalSpecs.page3.powderCoating}
                  </span>
                </div>
                <div className="pdf-spec-row">
                  <strong>INDEX ARRANGEMENT: </strong>
                  <span
                    contentEditable
                    suppressContentEditableWarning
                    onBlur={(e) => {
                      const text = e.currentTarget.innerText
                      setProposalSpecs((prev) => ({ ...prev, page3: { ...prev.page3, indexArrangement: text } }))
                    }}
                    title="Click to edit text"
                  >
                    {proposalSpecs.page3.indexArrangement}
                  </span>
                </div>
                <div className="pdf-spec-row">
                  <strong>SHAFT: </strong>
                  <span
                    contentEditable
                    suppressContentEditableWarning
                    onBlur={(e) => {
                      const text = e.currentTarget.innerText
                      setProposalSpecs((prev) => ({ ...prev, page3: { ...prev.page3, shaft: text } }))
                    }}
                    title="Click to edit text"
                  >
                    {proposalSpecs.page3.shaft}
                  </span>
                </div>
                <div className="pdf-spec-row">
                  <strong>PEDESTAL BEARING: </strong>
                  <span
                    contentEditable
                    suppressContentEditableWarning
                    onBlur={(e) => {
                      const text = e.currentTarget.innerText
                      setProposalSpecs((prev) => ({ ...prev, page3: { ...prev.page3, pedestalBearing: text } }))
                    }}
                    title="Click to edit text"
                  >
                    {proposalSpecs.page3.pedestalBearing}
                  </span>
                </div>
                <div className="pdf-spec-row">
                  <strong>WHEEL: </strong>
                  <span
                    contentEditable
                    suppressContentEditableWarning
                    onBlur={(e) => {
                      const text = e.currentTarget.innerText
                      setProposalSpecs((prev) => ({ ...prev, page3: { ...prev.page3, wheel: text } }))
                    }}
                    title="Click to edit text"
                  >
                    {proposalSpecs.page3.wheel}
                  </span>
                </div>
                <div className="pdf-spec-row">
                  <strong>CHAIN: </strong>
                  <span
                    contentEditable
                    suppressContentEditableWarning
                    onBlur={(e) => {
                      const text = e.currentTarget.innerText
                      setProposalSpecs((prev) => ({ ...prev, page3: { ...prev.page3, chain: text } }))
                    }}
                    title="Click to edit text"
                  >
                    {proposalSpecs.page3.chain}
                  </span>
                </div>
                <div className="pdf-spec-row">
                  <strong>OPEN ABLE SHUTTER: </strong>
                  <span
                    contentEditable
                    suppressContentEditableWarning
                    onBlur={(e) => {
                      const text = e.currentTarget.innerText
                      setProposalSpecs((prev) => ({ ...prev, page3: { ...prev.page3, openableShutter: text } }))
                    }}
                    title="Click to edit text"
                  >
                    {proposalSpecs.page3.openableShutter}
                  </span>
                </div>
              </div>

              <div className="pdf-footer">{commercialPagesCount + bodiesPageCount + 2} of {actualTotalPages}{/* 3 of {totalPages} */}</div>
            </div>

            {/* Page 4: Powder Coating Specification & Methods of Instruction (Merged 5 & 6) */}
            <div className="official-pdf-page">
              {renderPdfWatermark()}
              {renderOfficialHeader()}

              <div className="pdf-spec-container">
                {/* PRETREATMENT PROCESS &amp; PARAMETERS */}
                <div
                  className="pdf-spec-main-title"
                  contentEditable
                  suppressContentEditableWarning
                  onBlur={(e) => {
                    const text = e.currentTarget.innerText
                    setProposalSpecs((prev) => ({ ...prev, page4: { ...prev.page4, mainTitle: text } }))
                  }}
                  title="Click to edit text"
                >
                  {proposalSpecs.page4.mainTitle}
                </div>
                <div
                  className="pdf-spec-para"
                  style={{ marginTop: '4px', whiteSpace: 'pre-line' }}
                  contentEditable
                  suppressContentEditableWarning
                  onBlur={(e) => {
                    const text = e.currentTarget.innerText
                    setProposalSpecs((prev) => ({ ...prev, page4: { ...prev.page4, plantOverview: text } }))
                  }}
                  title="Click to edit text"
                >
                  {proposalSpecs.page4.plantOverview}
                </div>

                <div
                  className="pdf-spec-sec-title"
                  contentEditable
                  suppressContentEditableWarning
                  onBlur={(e) => {
                    const text = e.currentTarget.innerText
                    setProposalSpecs((prev) => ({ ...prev, page4: { ...prev.page4, machineryTitle: text } }))
                  }}
                  title="Click to edit text"
                >
                  {proposalSpecs.page4.machineryTitle}
                </div>
                <div className="pdf-spec-row">
                  <strong>1. PRETREATMENT PLANT : </strong>
                  <span
                    contentEditable
                    suppressContentEditableWarning
                    onBlur={(e) => {
                      const text = e.currentTarget.innerText
                      setProposalSpecs((prev) => ({ ...prev, page4: { ...prev.page4, pretreatmentPlant: text } }))
                    }}
                    title="Click to edit text"
                  >
                    {proposalSpecs.page4.pretreatmentPlant}
                  </span>
                </div>
                <div className="pdf-spec-grid-2" style={{ marginLeft: '12px' }}>
                  {proposalSpecs.page4.tankList.map((tank, tIdx) => (
                    <div
                      key={tIdx}
                      contentEditable
                      suppressContentEditableWarning
                      onBlur={(e) => {
                        const text = e.currentTarget.innerText
                        setProposalSpecs((prev) => {
                          const nextTanks = [...prev.page4.tankList]
                          nextTanks[tIdx] = text
                          return { ...prev, page4: { ...prev.page4, tankList: nextTanks } }
                        })
                      }}
                      title="Click to edit text"
                    >
                      {tank}
                    </div>
                  ))}
                </div>

                <div className="pdf-spec-row">
                  <strong>2. COMPRESSORS:</strong>
                </div>
                <div className="pdf-spec-sub-box">
                  <div
                    contentEditable
                    suppressContentEditableWarning
                    onBlur={(e) => {
                      const text = e.currentTarget.innerText
                      setProposalSpecs((prev) => ({ ...prev, page4: { ...prev.page4, compressorsA: text } }))
                    }}
                    title="Click to edit text"
                  >
                    {proposalSpecs.page4.compressorsA}
                  </div>
                  <div
                    contentEditable
                    suppressContentEditableWarning
                    onBlur={(e) => {
                      const text = e.currentTarget.innerText
                      setProposalSpecs((prev) => ({ ...prev, page4: { ...prev.page4, compressorsB: text } }))
                    }}
                    title="Click to edit text"
                  >
                    {proposalSpecs.page4.compressorsB}
                  </div>
                </div>

                <div className="pdf-spec-row">
                  <strong>3. POWDER COATING: </strong>
                  <span
                    contentEditable
                    suppressContentEditableWarning
                    onBlur={(e) => {
                      const text = e.currentTarget.innerText
                      setProposalSpecs((prev) => ({ ...prev, page4: { ...prev.page4, powderCoating: text } }))
                    }}
                    title="Click to edit text"
                  >
                    {proposalSpecs.page4.powderCoating}
                  </span>
                </div>

                <div className="pdf-spec-row">
                  <strong>4. OVEN: </strong>
                  <span style={{ fontSize: '9px', color: '#475569' }}>Inside working dimensions:</span>
                </div>
                <div className="pdf-spec-sub-box">
                  <div
                    contentEditable
                    suppressContentEditableWarning
                    onBlur={(e) => {
                      const text = e.currentTarget.innerText
                      setProposalSpecs((prev) => ({ ...prev, page4: { ...prev.page4, oven: text } }))
                    }}
                    title="Click to edit text"
                  >
                    {proposalSpecs.page4.oven}
                  </div>
                </div>

                <div className="pdf-spec-row">
                  <strong>5. MISCELLANEOUS EQUIPMENTS: </strong>
                  <span
                    contentEditable
                    suppressContentEditableWarning
                    onBlur={(e) => {
                      const text = e.currentTarget.innerText
                      setProposalSpecs((prev) => ({ ...prev, page4: { ...prev.page4, miscEquipments: text } }))
                    }}
                    title="Click to edit text"
                  >
                    {proposalSpecs.page4.miscEquipments}
                  </span>
                </div>

                {/* Merged from Page 6: Methods of Instruction */}
                <div
                  className="pdf-spec-main-title"
                  style={{ marginTop: '8px' }}
                  contentEditable
                  suppressContentEditableWarning
                  onBlur={(e) => {
                    const text = e.currentTarget.innerText
                    setProposalSpecs((prev) => ({ ...prev, page4: { ...prev.page4, instructionsTitle: text } }))
                  }}
                  title="Click to edit text"
                >
                  {proposalSpecs.page4.instructionsTitle}
                </div>

                <div
                  className="pdf-spec-sec-title"
                  contentEditable
                  suppressContentEditableWarning
                  onBlur={(e) => {
                    const text = e.currentTarget.innerText
                    setProposalSpecs((prev) => ({ ...prev, page4: { ...prev.page4, pretreatmentSecTitle: text } }))
                  }}
                  title="Click to edit text"
                >
                  {proposalSpecs.page4.pretreatmentSecTitle}
                </div>
                <div
                  className="pdf-spec-para"
                  contentEditable
                  suppressContentEditableWarning
                  onBlur={(e) => {
                    const text = e.currentTarget.innerText
                    setProposalSpecs((prev) => ({ ...prev, page4: { ...prev.page4, pretreatmentIntro: text } }))
                  }}
                  title="Click to edit text"
                >
                  {proposalSpecs.page4.pretreatmentIntro}
                </div>

                <div
                  className="pdf-spec-row"
                  contentEditable
                  suppressContentEditableWarning
                  onBlur={(e) => {
                    const text = e.currentTarget.innerText
                    setProposalSpecs((prev) => ({ ...prev, page4: { ...prev.page4, degreasingTitle: text } }))
                  }}
                  title="Click to edit text"
                >
                  <strong>{proposalSpecs.page4.degreasingTitle}</strong>
                </div>
                <div className="pdf-spec-sub-box">
                  {proposalSpecs.page4.degreasingLines.map((line, lIdx) => (
                    <div
                      key={lIdx}
                      contentEditable
                      suppressContentEditableWarning
                      onBlur={(e) => {
                        const text = e.currentTarget.innerText
                        setProposalSpecs((prev) => {
                          const nextLines = [...prev.page4.degreasingLines]
                          nextLines[lIdx] = text
                          return { ...prev, page4: { ...prev.page4, degreasingLines: nextLines } }
                        })
                      }}
                      title="Click to edit text"
                    >
                      {line}
                    </div>
                  ))}
                </div>

                <div className="pdf-spec-row">
                  <strong>1.2 Water Rinse: </strong>
                  <span
                    contentEditable
                    suppressContentEditableWarning
                    onBlur={(e) => {
                      const text = e.currentTarget.innerText
                      setProposalSpecs((prev) => ({ ...prev, page4: { ...prev.page4, waterRinse1: text } }))
                    }}
                    title="Click to edit text"
                  >
                    {proposalSpecs.page4.waterRinse1}
                  </span>
                </div>

                <div
                  className="pdf-spec-row"
                  contentEditable
                  suppressContentEditableWarning
                  onBlur={(e) => {
                    const text = e.currentTarget.innerText
                    setProposalSpecs((prev) => ({ ...prev, page4: { ...prev.page4, derustingTitle: text } }))
                  }}
                  title="Click to edit text"
                >
                  <strong>{proposalSpecs.page4.derustingTitle}</strong>
                </div>
                <div className="pdf-spec-sub-box">
                  {proposalSpecs.page4.derustingLines.map((line, lIdx) => (
                    <div
                      key={lIdx}
                      contentEditable
                      suppressContentEditableWarning
                      onBlur={(e) => {
                        const text = e.currentTarget.innerText
                        setProposalSpecs((prev) => {
                          const nextLines = [...prev.page4.derustingLines]
                          nextLines[lIdx] = text
                          return { ...prev, page4: { ...prev.page4, derustingLines: nextLines } }
                        })
                      }}
                      title="Click to edit text"
                    >
                      {line}
                    </div>
                  ))}
                </div>

                <div className="pdf-spec-row">
                  <strong>1.4 Water Rinse: </strong>
                  <span
                    contentEditable
                    suppressContentEditableWarning
                    onBlur={(e) => {
                      const text = e.currentTarget.innerText
                      setProposalSpecs((prev) => ({ ...prev, page4: { ...prev.page4, waterRinse2: text } }))
                    }}
                    title="Click to edit text"
                  >
                    {proposalSpecs.page4.waterRinse2}
                  </span>
                </div>
              </div>

              <div className="pdf-footer">{commercialPagesCount + bodiesPageCount + 3} of {actualTotalPages}{/* 4 of {totalPages} */}</div>
            </div>

            {/* Page 5: Activation, Phosphating, Passivation & Quality Control (Merged 7 & 8) */}
            <div className="official-pdf-page">
              {renderPdfWatermark()}
              {renderOfficialHeader()}

              <div className="pdf-spec-container">
                {/* PRETREATMENT PROCESS &amp; PARAMETERS */}
                <div
                  className="pdf-spec-main-title"
                  contentEditable
                  suppressContentEditableWarning
                  onBlur={(e) => {
                    const text = e.currentTarget.innerText
                    setProposalSpecs((prev) => ({ ...prev, page5: { ...prev.page5, mainTitle: text } }))
                  }}
                  title="Click to edit text"
                >
                  {proposalSpecs.page5.mainTitle}
                </div>

                <div
                  className="pdf-spec-row"
                  style={{ marginTop: '5px' }}
                  contentEditable
                  suppressContentEditableWarning
                  onBlur={(e) => {
                    const text = e.currentTarget.innerText
                    setProposalSpecs((prev) => ({ ...prev, page5: { ...prev.page5, activationTitle: text } }))
                  }}
                  title="Click to edit text"
                >
                  <strong>{proposalSpecs.page5.activationTitle}</strong>
                </div>
                <div className="pdf-spec-sub-box">
                  {proposalSpecs.page5.activationLines.map((line, lIdx) => (
                    <div
                      key={lIdx}
                      contentEditable
                      suppressContentEditableWarning
                      onBlur={(e) => {
                        const text = e.currentTarget.innerText
                        setProposalSpecs((prev) => {
                          const nextLines = [...prev.page5.activationLines]
                          nextLines[lIdx] = text
                          return { ...prev, page5: { ...prev.page5, activationLines: nextLines } }
                        })
                      }}
                      title="Click to edit text"
                    >
                      {line}
                    </div>
                  ))}
                </div>

                <div
                  className="pdf-spec-row"
                  contentEditable
                  suppressContentEditableWarning
                  onBlur={(e) => {
                    const text = e.currentTarget.innerText
                    setProposalSpecs((prev) => ({ ...prev, page5: { ...prev.page5, zincPhosphatingTitle: text } }))
                  }}
                  title="Click to edit text"
                >
                  <strong>{proposalSpecs.page5.zincPhosphatingTitle}</strong>
                </div>
                <div className="pdf-spec-sub-box">
                  {proposalSpecs.page5.zincPhosphatingLines.map((line, lIdx) => (
                    <div
                      key={lIdx}
                      contentEditable
                      suppressContentEditableWarning
                      onBlur={(e) => {
                        const text = e.currentTarget.innerText
                        setProposalSpecs((prev) => {
                          const nextLines = [...prev.page5.zincPhosphatingLines]
                          nextLines[lIdx] = text
                          return { ...prev, page5: { ...prev.page5, zincPhosphatingLines: nextLines } }
                        })
                      }}
                      title="Click to edit text"
                    >
                      {line}
                    </div>
                  ))}
                </div>

                <div className="pdf-spec-row">
                  <strong>1.7 Water Rinse: </strong>
                  <span
                    contentEditable
                    suppressContentEditableWarning
                    onBlur={(e) => {
                      const text = e.currentTarget.innerText
                      setProposalSpecs((prev) => ({ ...prev, page5: { ...prev.page5, waterRinse3: text } }))
                    }}
                    title="Click to edit text"
                  >
                    {proposalSpecs.page5.waterRinse3}
                  </span>
                </div>

                <div
                  className="pdf-spec-row"
                  contentEditable
                  suppressContentEditableWarning
                  onBlur={(e) => {
                    const text = e.currentTarget.innerText
                    setProposalSpecs((prev) => ({ ...prev, page5: { ...prev.page5, passivationTitle: text } }))
                  }}
                  title="Click to edit text"
                >
                  <strong>{proposalSpecs.page5.passivationTitle}</strong>
                </div>
                <div className="pdf-spec-sub-box">
                  {proposalSpecs.page5.passivationLines.map((line, lIdx) => (
                    <div
                      key={lIdx}
                      contentEditable
                      suppressContentEditableWarning
                      onBlur={(e) => {
                        const text = e.currentTarget.innerText
                        setProposalSpecs((prev) => {
                          const nextLines = [...prev.page5.passivationLines]
                          nextLines[lIdx] = text
                          return { ...prev, page5: { ...prev.page5, passivationLines: nextLines } }
                        })
                      }}
                      title="Click to edit text"
                    >
                      {line}
                    </div>
                  ))}
                </div>

                <div
                  className="pdf-spec-sec-title"
                  contentEditable
                  suppressContentEditableWarning
                  onBlur={(e) => {
                    const text = e.currentTarget.innerText
                    setProposalSpecs((prev) => ({ ...prev, page5: { ...prev.page5, parametersTitle: text } }))
                  }}
                  title="Click to edit text"
                >
                  {proposalSpecs.page5.parametersTitle}
                </div>
                <div className="pdf-spec-grid-2" style={{ marginLeft: '4px' }}>
                  {proposalSpecs.page5.parametersList.map((param, pIdx) => (
                    <div
                      key={pIdx}
                      contentEditable
                      suppressContentEditableWarning
                      onBlur={(e) => {
                        const text = e.currentTarget.innerText
                        setProposalSpecs((prev) => {
                          const nextParams = [...prev.page5.parametersList]
                          nextParams[pIdx] = text
                          return { ...prev, page5: { ...prev.page5, parametersList: nextParams } }
                        })
                      }}
                      title="Click to edit text"
                    >
                      {param}
                    </div>
                  ))}
                </div>

                {/* Merged from Page 8: Quality Control */}
                <div
                  className="pdf-spec-main-title"
                  style={{ marginTop: '8px' }}
                  contentEditable
                  suppressContentEditableWarning
                  onBlur={(e) => {
                    const text = e.currentTarget.innerText
                    setProposalSpecs((prev) => ({ ...prev, page5: { ...prev.page5, qcMainTitle: text } }))
                  }}
                  title="Click to edit text"
                >
                  {proposalSpecs.page5.qcMainTitle}
                </div>

                <div
                  className="pdf-spec-sec-title"
                  contentEditable
                  suppressContentEditableWarning
                  onBlur={(e) => {
                    const text = e.currentTarget.innerText
                    setProposalSpecs((prev) => ({ ...prev, page5: { ...prev.page5, qcArrivalTitle: text } }))
                  }}
                  title="Click to edit text"
                >
                  {proposalSpecs.page5.qcArrivalTitle}
                </div>
                <div className="pdf-spec-para">
                  {proposalSpecs.page5.qcArrivalLines.map((line, lIdx) => (
                    <div
                      key={lIdx}
                      contentEditable
                      suppressContentEditableWarning
                      onBlur={(e) => {
                        const text = e.currentTarget.innerText
                        setProposalSpecs((prev) => {
                          const nextLines = [...prev.page5.qcArrivalLines]
                          nextLines[lIdx] = text
                          return { ...prev, page5: { ...prev.page5, qcArrivalLines: nextLines } }
                        })
                      }}
                      title="Click to edit text"
                    >
                      {line}
                    </div>
                  ))}
                </div>

                <div
                  className="pdf-spec-sec-title"
                  contentEditable
                  suppressContentEditableWarning
                  onBlur={(e) => {
                    const text = e.currentTarget.innerText
                    setProposalSpecs((prev) => ({ ...prev, page5: { ...prev.page5, qcProductionTitle: text } }))
                  }}
                  title="Click to edit text"
                >
                  {proposalSpecs.page5.qcProductionTitle}
                </div>
                <div className="pdf-spec-para">
                  {proposalSpecs.page5.qcProductionLines.map((line, lIdx) => (
                    <div
                      key={lIdx}
                      contentEditable
                      suppressContentEditableWarning
                      onBlur={(e) => {
                        const text = e.currentTarget.innerText
                        setProposalSpecs((prev) => {
                          const nextLines = [...prev.page5.qcProductionLines]
                          nextLines[lIdx] = text
                          return { ...prev, page5: { ...prev.page5, qcProductionLines: nextLines } }
                        })
                      }}
                      title="Click to edit text"
                    >
                      {line}
                    </div>
                  ))}
                </div>

                <div
                  className="pdf-spec-sec-title"
                  contentEditable
                  suppressContentEditableWarning
                  onBlur={(e) => {
                    const text = e.currentTarget.innerText
                    setProposalSpecs((prev) => ({ ...prev, page5: { ...prev.page5, qcFinalTitle: text } }))
                  }}
                  title="Click to edit text"
                >
                  {proposalSpecs.page5.qcFinalTitle}
                </div>
                <div className="pdf-spec-para">
                  {proposalSpecs.page5.qcFinalLines.map((line, lIdx) => (
                    <div
                      key={lIdx}
                      contentEditable
                      suppressContentEditableWarning
                      onBlur={(e) => {
                        const text = e.currentTarget.innerText
                        setProposalSpecs((prev) => {
                          const nextLines = [...prev.page5.qcFinalLines]
                          nextLines[lIdx] = text
                          return { ...prev, page5: { ...prev.page5, qcFinalLines: nextLines } }
                        })
                      }}
                      title="Click to edit text"
                    >
                      {line}
                    </div>
                  ))}
                </div>

                <div
                  className="pdf-spec-sec-title"
                  contentEditable
                  suppressContentEditableWarning
                  onBlur={(e) => {
                    const text = e.currentTarget.innerText
                    setProposalSpecs((prev) => ({ ...prev, page5: { ...prev.page5, qcTestingTitle: text } }))
                  }}
                  title="Click to edit text"
                >
                  {proposalSpecs.page5.qcTestingTitle}
                </div>
                <div className="pdf-spec-para">
                  {proposalSpecs.page5.qcTestingLines.map((line, lIdx) => (
                    <div
                      key={lIdx}
                      contentEditable
                      suppressContentEditableWarning
                      onBlur={(e) => {
                        const text = e.currentTarget.innerText
                        setProposalSpecs((prev) => {
                          const nextLines = [...prev.page5.qcTestingLines]
                          nextLines[lIdx] = text
                          return { ...prev, page5: { ...prev.page5, qcTestingLines: nextLines } }
                        })
                      }}
                      title="Click to edit text"
                    >
                      {line}
                    </div>
                  ))}
                </div>
              </div>

              <div className="pdf-footer">{commercialPagesCount + bodiesPageCount + 4} of {actualTotalPages}{/* 5 of {totalPages} */}</div>
            </div>
          </>
        )}
      </div>

      {/* =========================================================================
          VERTICAL OPTIONS SIDEBAR (Right-hand Side of Page)
          ========================================================================= */}
      <aside className={`official-right-sidebar no-print ${sidebarCollapsed ? 'collapsed' : ''} ${mobileView === 'document' ? 'mobile-hidden' : ''}`}>
        {sidebarCollapsed ? (
          <div style={{ padding: '16px 6px', display: 'flex', flexDirection: 'column', alignItems: 'center', gap: '14px' }}>
            <button
              type="button"
              className="icon-btn sm"
              onClick={() => setSidebarCollapsed(false)}
              title="Expand Options Sidebar"
              style={{ background: '#2563eb', color: '#ffffff' }}
            >
              {Icon.chevronRight || '◀'}
            </button>
            <span
              style={{
                writingMode: 'vertical-rl',
                transform: 'rotate(180deg)',
                fontSize: '11px',
                fontWeight: 700,
                color: '#475569',
                letterSpacing: '1.5px',
                margin: '12px 0',
              }}
            >
              DOCUMENT OPTIONS
            </span>
            <button
              type="button"
              className="icon-btn sm"
              onClick={handleDirectDownload}
              disabled={isGeneratingPdf}
              title="Download PDF directly"
            >
              {isGeneratingPdf ? <Spinner size={14} /> : Icon.fileText}
            </button>
          </div>
        ) : (
          <>
            {/* 1. Header */}
            <div className="sidebar-header">
              <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: '8px' }}>
                <div>
                  <div style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                    <h3 style={{ margin: 0, fontSize: '14.5px', fontWeight: 700, color: '#0f172a' }}>
                      Document Options
                    </h3>
                    <span className="chip primary sm" style={{ fontSize: '10px', padding: '1px 6px' }}>
                      A4 {actualTotalPages}P
                    </span>
                  </div>
                </div>
                <div style={{ display: 'flex', alignItems: 'center', gap: '4px' }}>
                  <button
                    type="button"
                    className="icon-btn sm"
                    onClick={() => setSidebarCollapsed(true)}
                    title="Collapse sidebar to view full width"
                  >
                    {Icon.chevronLeft || '▶'}
                  </button>
                  <button
                    type="button"
                    className="icon-btn sm"
                    onClick={onClose}
                    title="Close preview"
                  >
                    {Icon.close}
                  </button>
                </div>
              </div>

              {/* Action Buttons */}
              <div style={{ display: 'grid', gridTemplateColumns: '1.4fr 1fr', gap: '6px', marginBottom: '6px' }}>
                <button
                  type="button"
                  className="btn primary sm"
                  onClick={handleDirectDownload}
                  disabled={isGeneratingPdf}
                  title="Directly download PDF file to your downloads folder"
                  style={{ fontSize: '11.5px', padding: '6px 8px', justifyContent: 'center' }}
                >
                  {isGeneratingPdf ? <Spinner size={13} /> : Icon.fileText} Download PDF
                </button>
                <button
                  type="button"
                  className="btn ghost sm"
                  onClick={handleSave}
                  title="Save changes to project"
                  style={{ fontSize: '11.5px', padding: '6px 8px', justifyContent: 'center' }}
                >
                  {Icon.check} Save
                </button>
              </div>
              <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '6px', marginBottom: '10px' }}>
                <button
                  type="button"
                  className="btn sm"
                  onClick={handleWhatsAppShare}
                  disabled={isGeneratingPdf}
                  title="Download PDF and share quotation with client on WhatsApp"
                  style={{
                    fontSize: '11.5px',
                    padding: '5px 8px',
                    justifyContent: 'center',
                    background: '#16a34a',
                    color: '#ffffff',
                    fontWeight: 700,
                    border: 'none',
                    borderRadius: '6px',
                    display: 'flex',
                    alignItems: 'center',
                    gap: '4px',
                    cursor: 'pointer',
                  }}
                >
                  <span>💬</span> WhatsApp
                </button>
                <button
                  type="button"
                  className="btn ghost sm"
                  onClick={handlePrint}
                  title="Open Browser Print Dialog (Physical Printer or Save as PDF)"
                  style={{ fontSize: '11px', padding: '4px 8px', justifyContent: 'center', color: '#475569' }}
                >
                  🖨️ Print / Vector
                </button>
              </div>

              {/* Category Navigation Pills */}
              <div className="sidebar-nav-tabs">
                <button
                  type="button"
                  className={`sidebar-tab-btn ${activeSidebarTab === 'options' ? 'active' : ''}`}
                  onClick={() => setActiveSidebarTab('options')}
                >
                  ⚡ Options
                </button>
                {userIsAdmin && (
                  <button
                    type="button"
                    className={`sidebar-tab-btn ${activeSidebarTab === 'breakdown' ? 'active' : ''}`}
                    onClick={() => setActiveSidebarTab('breakdown')}
                  >
                    📊 Unit Calc
                  </button>
                )}
                <button
                  type="button"
                  className={`sidebar-tab-btn ${activeSidebarTab === 'client' ? 'active' : ''}`}
                  onClick={() => setActiveSidebarTab('client')}
                >
                  🏢 Client
                </button>
                <button
                  type="button"
                  className={`sidebar-tab-btn ${activeSidebarTab === 'terms' ? 'active' : ''}`}
                  onClick={() => setActiveSidebarTab('terms')}
                >
                  📋 Terms
                </button>
                <button
                  type="button"
                  className={`sidebar-tab-btn ${activeSidebarTab === 'specs' ? 'active' : ''}`}
                  onClick={() => setActiveSidebarTab('specs')}
                  title="Edit Technical Proposal Specs (Pages 3 to 5)"
                >
                  📑 Tech Specs
                </button>
                <button
                  type="button"
                  className={`sidebar-tab-btn ${activeSidebarTab === 'body' ? 'active' : ''}`}
                  onClick={() => setActiveSidebarTab('body')}
                >
                  🔩 Bodies
                </button>
              </div>
            </div>

            {/* 2. Scrollable Body */}
            <div className="sidebar-body">
              {activeSidebarTab === 'options' && (
                <>
                  {/* TABLE SIZE, SCALE & DENSITY CONTROLS */}
                  <div className="option-group-card">
                    <div className="option-group-title">
                      <span>TABLE SIZE, SCALE &amp; MULTI-SYSTEM LAYOUT</span>
                    </div>

                    <div style={{ marginBottom: '8px' }}>
                      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '4px' }}>
                        <label className="tiny font-bold muted">Table Font Size &amp; Scale:</label>
                        <span className="tiny font-bold font-mono" style={{ color: '#0284c7' }}>{tableScale}% ({tableFontSize})</span>
                      </div>
                      <div style={{ display: 'flex', alignItems: 'center', gap: '4px' }}>
                        <button
                          type="button"
                          className="btn ghost sm"
                          onClick={() => setTableScale((s) => Math.max(75, s - 5))}
                          style={{ padding: '3px 8px', fontSize: '12px', fontWeight: 700 }}
                          title="Decrease table font size (A-)"
                        >
                          A-
                        </button>
                        {[85, 90, 95, 100, 105].map((scale) => (
                          <button
                            key={scale}
                            type="button"
                            className={`option-pill ${tableScale === scale ? 'active' : ''}`}
                            onClick={() => setTableScale(scale)}
                            style={{ flex: 1, padding: '4px 2px', fontSize: '10px', textAlign: 'center' }}
                          >
                            {scale}%
                          </button>
                        ))}
                        <button
                          type="button"
                          className="btn ghost sm"
                          onClick={() => setTableScale((s) => Math.min(125, s + 5))}
                          style={{ padding: '3px 8px', fontSize: '12px', fontWeight: 700 }}
                          title="Increase table font size (A+)"
                        >
                          A+
                        </button>
                      </div>
                    </div>

                    <div style={{ marginBottom: '8px' }}>
                      <label className="tiny font-bold muted" style={{ display: 'block', marginBottom: '4px' }}>
                        Table Row Density (Cell Padding):
                      </label>
                      <div className="option-pill-group" style={{ display: 'grid', gridTemplateColumns: 'repeat(3, 1fr)', gap: '4px' }}>
                        {(['compact', 'normal', 'relaxed'] as const).map((d) => (
                          <button
                            key={d}
                            type="button"
                            className={`option-pill ${tableDensity === d ? 'active' : ''}`}
                            onClick={() => setTableDensity(d)}
                            style={{ fontSize: '10px', padding: '5px 4px', textAlign: 'center', textTransform: 'capitalize' }}
                          >
                            {d}
                          </button>
                        ))}
                      </div>
                    </div>

                    {sets.length > 2 && (
                      <div style={{ marginBottom: '8px' }}>
                        <label className="tiny font-bold muted" style={{ display: 'block', marginBottom: '4px' }}>
                          Max Systems Per Page:
                        </label>
                        <div className="option-pill-group" style={{ display: 'grid', gridTemplateColumns: 'repeat(6, 1fr)', gap: '3px' }}>
                          <button
                            type="button"
                            className={`option-pill ${systemsPerPageSetting === 'auto' ? 'active' : ''}`}
                            onClick={() => setSystemsPerPageSetting('auto')}
                            style={{ fontSize: '10px', padding: '4px 2px', textAlign: 'center', fontWeight: 700 }}
                            title={`Auto (Best): automatically selects ${effectiveSystemsPerPage}/page based on system model counts`}
                          >
                            ✨Auto({effectiveSystemsPerPage})
                          </button>
                          {[2, 3, 4, 5, 999].map((n) => (
                            <button
                              key={n}
                              type="button"
                              className={`option-pill ${systemsPerPageSetting === n ? 'active' : ''}`}
                              onClick={() => setSystemsPerPageSetting(n as any)}
                              style={{ fontSize: '10px', padding: '4px 2px', textAlign: 'center' }}
                              title={n === 999 ? 'Fit all systems on 1 page' : `${n} systems per page`}
                            >
                              {n === 999 ? 'All' : `${n}/p`}
                            </button>
                          ))}
                        </div>
                        {systemsPerPageSetting === 'auto' && (
                          <div className="tiny muted" style={{ marginTop: '3px', fontSize: '9.5px', color: '#0369a1' }}>
                            ✓ Best Layout: {effectiveSystemsPerPage} systems/page chosen based on model row counts (3/p for 4-5 models, 4/p for 2-3 models).
                          </div>
                        )}
                      </div>
                    )}

                    {sets.length > 1 && (
                      <div>
                        <label className="tiny font-bold muted" style={{ display: 'block', marginBottom: '4px' }}>
                          Reorder Systems (Move Up / Down):
                        </label>
                        <div style={{ display: 'flex', flexDirection: 'column', gap: '4px' }}>
                          {sets.map((s, idx) => {
                            const nextChar = String.fromCharCode(65 + idx)
                            const title = (s.system_title || (s.set_label?.startsWith('SYSTEM') ? s.set_label : `SYSTEM ${nextChar}`)).trim()
                            const sysQty = Number(s.quantity) || 1
                            return (
                              <div
                                key={s.id || idx}
                                style={{
                                  display: 'flex',
                                  alignItems: 'center',
                                  justifyContent: 'space-between',
                                  padding: '4px 8px',
                                  background: '#f8fafc',
                                  border: '1px solid #cbd5e1',
                                  borderRadius: '4px',
                                  fontSize: '11px',
                                }}
                              >
                                <span style={{ fontWeight: 600, color: '#0f172a' }}>
                                  {title} ({sysQty} {sysQty > 1 ? 'SETS' : 'SET'})
                                </span>
                                <div style={{ display: 'flex', gap: '3px' }}>
                                  <button
                                    type="button"
                                    disabled={idx === 0}
                                    onClick={() => moveSystem(idx, idx - 1)}
                                    style={{
                                      padding: '2px 5px',
                                      fontSize: '10px',
                                      cursor: idx === 0 ? 'not-allowed' : 'pointer',
                                      opacity: idx === 0 ? 0.4 : 1,
                                      border: '1px solid #cbd5e1',
                                      borderRadius: '3px',
                                      background: '#ffffff',
                                    }}
                                    title="Move System Up"
                                  >
                                    ▲
                                  </button>
                                  <button
                                    type="button"
                                    disabled={idx === sets.length - 1}
                                    onClick={() => moveSystem(idx, idx + 1)}
                                    style={{
                                      padding: '2px 5px',
                                      fontSize: '10px',
                                      cursor: idx === sets.length - 1 ? 'not-allowed' : 'pointer',
                                      opacity: idx === sets.length - 1 ? 0.4 : 1,
                                      border: '1px solid #cbd5e1',
                                      borderRadius: '3px',
                                      background: '#ffffff',
                                    }}
                                    title="Move System Down"
                                  >
                                    ▼
                                  </button>
                                </div>
                              </div>
                            )
                          })}
                        </div>
                      </div>
                    )}
                  </div>

                  {/* WORD-STYLE COLUMN WIDTH CONTROLS */}
                  <div className="option-group-card">
                    <div className="option-group-title" style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                      <span>WORD-STYLE COLUMN WIDTHS</span>
                      <button
                        type="button"
                        className="btn ghost xs"
                        onClick={() => {
                          setGoodsColWidths({ sl: 34, qty: 55, set: 60, rate: 85, amount: 100 })
                          setBodyColWidths({ sr: 40, per_bay: 75, qty: 55, body: 85 })
                        }}
                        style={{ fontSize: '9.5px', padding: '1px 5px', color: '#64748b' }}
                        title="Reset all column widths to defaults"
                      >
                        Reset
                      </button>
                    </div>
                    <div className="tiny muted" style={{ marginBottom: '8px', fontSize: '10px' }}>
                      💡 Drag column borders in document preview, or adjust widths below:
                    </div>
                    <div style={{ display: 'grid', gridTemplateColumns: 'repeat(2, 1fr)', gap: '6px' }}>
                      <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', background: '#f8fafc', padding: '3px 6px', borderRadius: '4px', border: '1px solid #e2e8f0', fontSize: '10.5px' }}>
                        <span>Quantity: <strong>{goodsColWidths.qty}px</strong></span>
                        <div style={{ display: 'flex', gap: '2px' }}>
                          <button type="button" onClick={() => setGoodsColWidths((p) => ({ ...p, qty: Math.max(30, p.qty - 5) }))} style={{ padding: '1px 5px', fontSize: '10px', borderRadius: '2px', border: '1px solid #cbd5e1', background: '#fff' }}>-</button>
                          <button type="button" onClick={() => setGoodsColWidths((p) => ({ ...p, qty: Math.min(120, p.qty + 5) }))} style={{ padding: '1px 5px', fontSize: '10px', borderRadius: '2px', border: '1px solid #cbd5e1', background: '#fff' }}>+</button>
                        </div>
                      </div>
                      <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', background: '#f8fafc', padding: '3px 6px', borderRadius: '4px', border: '1px solid #e2e8f0', fontSize: '10.5px' }}>
                        <span>Set Col: <strong>{goodsColWidths.set}px</strong></span>
                        <div style={{ display: 'flex', gap: '2px' }}>
                          <button type="button" onClick={() => setGoodsColWidths((p) => ({ ...p, set: Math.max(30, p.set - 5) }))} style={{ padding: '1px 5px', fontSize: '10px', borderRadius: '2px', border: '1px solid #cbd5e1', background: '#fff' }}>-</button>
                          <button type="button" onClick={() => setGoodsColWidths((p) => ({ ...p, set: Math.min(120, p.set + 5) }))} style={{ padding: '1px 5px', fontSize: '10px', borderRadius: '2px', border: '1px solid #cbd5e1', background: '#fff' }}>+</button>
                        </div>
                      </div>
                      <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', background: '#f8fafc', padding: '3px 6px', borderRadius: '4px', border: '1px solid #e2e8f0', fontSize: '10.5px' }}>
                        <span>Rate Col: <strong>{goodsColWidths.rate}px</strong></span>
                        <div style={{ display: 'flex', gap: '2px' }}>
                          <button type="button" onClick={() => setGoodsColWidths((p) => ({ ...p, rate: Math.max(50, p.rate - 5) }))} style={{ padding: '1px 5px', fontSize: '10px', borderRadius: '2px', border: '1px solid #cbd5e1', background: '#fff' }}>-</button>
                          <button type="button" onClick={() => setGoodsColWidths((p) => ({ ...p, rate: Math.min(150, p.rate + 5) }))} style={{ padding: '1px 5px', fontSize: '10px', borderRadius: '2px', border: '1px solid #cbd5e1', background: '#fff' }}>+</button>
                        </div>
                      </div>
                      <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', background: '#f8fafc', padding: '3px 6px', borderRadius: '4px', border: '1px solid #e2e8f0', fontSize: '10.5px' }}>
                        <span>Amount: <strong>{goodsColWidths.amount}px</strong></span>
                        <div style={{ display: 'flex', gap: '2px' }}>
                          <button type="button" onClick={() => setGoodsColWidths((p) => ({ ...p, amount: Math.max(60, p.amount - 5) }))} style={{ padding: '1px 5px', fontSize: '10px', borderRadius: '2px', border: '1px solid #cbd5e1', background: '#fff' }}>-</button>
                          <button type="button" onClick={() => setGoodsColWidths((p) => ({ ...p, amount: Math.min(160, p.amount + 5) }))} style={{ padding: '1px 5px', fontSize: '10px', borderRadius: '2px', border: '1px solid #cbd5e1', background: '#fff' }}>+</button>
                        </div>
                      </div>
                    </div>
                  </div>

                                    {/* 1. LAYOUT & PROPOSAL FORMAT (First in Options) */}
                  <div className="option-group-card">
                    <div className="option-group-title">
                      <span>LAYOUT &amp; PROPOSAL FORMAT</span>
                    </div>
                    <label className="tiny font-bold muted" style={{ display: 'block', marginBottom: '4px' }}>
                      Commercial Presentation Style:
                    </label>
                    <div className="option-pill-group" style={{ display: 'grid', gridTemplateColumns: userIsAdmin ? 'repeat(3, 1fr)' : 'repeat(2, 1fr)', gap: '4px' }}>
                      <button
                        type="button"
                        className={`option-pill ${effectiveQuotationMode === 'set_wise' ? 'active' : ''}`}
                        onClick={() => setQuotationMode('set_wise')}
                        style={{ fontSize: '10px', padding: '6px 4px', textAlign: 'center' }}
                      >
                        System / Set-Wise (SYSTEM A/B/C)
                      </button>
                      <button
                        type="button"
                        className={`option-pill ${effectiveQuotationMode === 'unit_wise' ? 'active' : ''}`}
                        onClick={() => setQuotationMode('unit_wise')}
                        style={{ fontSize: '10px', padding: '6px 4px', textAlign: 'center' }}
                      >
                        Individual Unit-Wise
                      </button>
                      {userIsAdmin && (
                        <button
                          type="button"
                          className={`option-pill ${effectiveQuotationMode === 'unit_breakdown' ? 'active' : ''}`}
                          onClick={() => setQuotationMode('unit_breakdown')}
                          style={{ fontSize: '10px', padding: '6px 4px', textAlign: 'center' }}
                        >
                          Unit Breakdown (7-Col)
                        </button>
                      )}
                    </div>

                    <div style={{ marginTop: '10px', paddingTop: '8px', borderTop: '1px solid #e2e8f0' }}>
                      <label className="tiny font-bold muted" style={{ display: 'block', marginBottom: '4px' }}>
                        Installation Bodies Table Placement:
                      </label>
                      <div className="option-pill-group" style={{ display: 'grid', gridTemplateColumns: 'repeat(2, 1fr)', gap: '4px' }}>
                        <button
                          type="button"
                          className={`option-pill ${bodyPlacement === 'dedicated' ? 'active' : ''}`}
                          onClick={() => {
                            setBodyPlacement('dedicated')
                            setShowBodiesTable(true)
                            setPlaceBodyOnPage1(false)
                          }}
                          style={{ fontSize: '10px', padding: '5px 4px', textAlign: 'center' }}
                          title="Best for multi-system quotations. Gives full space without crowding terms."
                        >
                          📄 Dedicated Page
                        </button>
                        <button
                          type="button"
                          className={`option-pill ${bodyPlacement === 'terms' ? 'active' : ''}`}
                          onClick={() => {
                            setBodyPlacement('terms')
                            setShowBodiesTable(true)
                            setPlaceBodyOnPage1(false)
                          }}
                          style={{ fontSize: '10px', padding: '5px 4px', textAlign: 'center' }}
                          title="Place installation bodies table on Page 2 above Terms & Conditions"
                        >
                          📋 Page 2 (Above Terms)
                        </button>
                        <button
                          type="button"
                          className={`option-pill ${bodyPlacement === 'commercial' ? 'active' : ''}`}
                          onClick={() => {
                            setBodyPlacement('commercial')
                            setShowBodiesTable(true)
                            setPlaceBodyOnPage1(true)
                          }}
                          style={{ fontSize: '10px', padding: '5px 4px', textAlign: 'center' }}
                          title="Fit on Commercial Page (under goods table)"
                        >
                          📑 Page 1 (Commercial)
                        </button>
                        <button
                          type="button"
                          className={`option-pill ${bodyPlacement === 'none' ? 'active' : ''}`}
                          onClick={() => {
                            setBodyPlacement('none')
                            setShowBodiesTable(false)
                            setPlaceBodyOnPage1(false)
                          }}
                          style={{ fontSize: '10px', padding: '5px 4px', textAlign: 'center' }}
                          title="Do not show installation bodies table in PDF"
                        >
                          🚫 Hide Table
                        </button>
                      </div>
                      {bodyPlacement === 'dedicated' && (
                        <div className="tiny muted" style={{ marginTop: '4px', color: '#0369a1', fontSize: '9.5px' }}>
                          ✓ Multi-System Annexure: rendered on its own dedicated A4 page with all system subtotals and grand total.
                        </div>
                      )}
                    </div>

                    <label style={{ display: 'flex', alignItems: 'center', gap: '8px', cursor: 'pointer', marginTop: '4px' }}>
                      <input
                        type="checkbox"
                        checked={includeFullProposal}
                        onChange={(e) => {
                          const val = e.target.checked
                          setIncludeFullProposal(val)
                          if (typeof window !== 'undefined' && window.localStorage) {
                            try {
                              localStorage.setItem('examcalc_default_include_5_page_proposal', String(val))
                            } catch {}
                          }
                        }}
                      />
                      <span style={{ fontSize: '12px' }}>Include Full 5-Page Technical Proposal</span>
                    </label>


                  </div>
                  {/* 2. DESCRIPTION OF GOODS (RED SPECIFICATIONS) */}
                  <div className="option-group-card" style={{ borderLeft: '3px solid #dc2626' }}>
                    <div className="option-group-title">
                      <span style={{ color: '#dc2626', display: 'flex', alignItems: 'center', gap: '6px' }}>
                        <span>🔴</span> DESCRIPTION OF GOODS (RED SPECS)
                      </span>
                      <span className="chip sm" style={{ background: '#fef2f2', color: '#dc2626', border: '1px solid #fecaca', fontSize: '10px' }}>
                        {redSpecs.length} {redSpecs.length === 1 ? 'line' : 'lines'}
                      </span>
                    </div>
                    <p className="tiny muted" style={{ margin: '0 0 8px 0', fontSize: '11px', lineHeight: 1.4 }}>
                      Red specification notes displayed at the bottom of the goods table on Page 1. Fully editable with smart width &amp; track count defaults.
                    </p>

                    {/* Quick Preset Buttons for Tracks & Double Drive */}
                    <div style={{ marginBottom: '8px' }}>
                      <label className="tiny font-bold muted" style={{ display: 'block', marginBottom: '4px' }}>
                        Quick Presets:
                      </label>
                      <div style={{ display: 'flex', flexWrap: 'wrap', gap: '4px' }}>
                        <button
                          type="button"
                          className={`option-pill ${redSpecs.some((s) => s.toUpperCase().includes('DOUBLE DRIVE')) ? 'active' : ''}`}
                          onClick={() => {
                            setRedSpecs((prev) => {
                              const hasDrive = prev.some((s) => s.toUpperCase().includes('DOUBLE DRIVE'))
                              if (hasDrive) {
                                return prev.filter((s) => !s.toUpperCase().includes('DOUBLE DRIVE'))
                              } else {
                                return ['DOUBLE DRIVE WITH REDUCTION WHEEL', ...prev]
                              }
                            })
                          }}
                          style={{ fontSize: '10px', padding: '3px 8px' }}
                        >
                          {redSpecs.some((s) => s.toUpperCase().includes('DOUBLE DRIVE')) ? '✓ Reduction Wheel' : '+ Reduction Wheel'}
                        </button>

                        {[1, 2, 3, 4, 5].map((tCount) => {
                          const trackLabel = `${tCount} ${tCount === 1 ? 'TRACK' : 'TRACKS'}`
                          const isCurrent = redSpecs.some(
                            (s) => s.toUpperCase() === trackLabel || s.toUpperCase() === `${tCount} TRACK` || s.toUpperCase() === `${tCount} TRACKS`
                          )
                          return (
                            <button
                              key={tCount}
                              type="button"
                              className={`option-pill ${isCurrent ? 'active' : ''}`}
                              onClick={() => {
                                setRedSpecs((prev) => {
                                  const trackIdx = prev.findIndex((s) => s.toUpperCase().includes('TRACK'))
                                  if (trackIdx >= 0) {
                                    return prev.map((s, i) => (i === trackIdx ? trackLabel : s))
                                  } else {
                                    const driveIdx = prev.findIndex((s) => s.toUpperCase().includes('DOUBLE DRIVE'))
                                    const insertAt = driveIdx >= 0 ? driveIdx + 1 : 0
                                    const copy = [...prev]
                                    copy.splice(insertAt, 0, trackLabel)
                                    return copy
                                  }
                                })
                              }}
                              style={{ fontSize: '10px', padding: '3px 7px' }}
                            >
                              {tCount} {tCount === 1 ? 'Track' : 'Tracks'}
                            </button>
                          )
                        })}
                      </div>
                    </div>

                    {/* Editable Specification Lines */}
                    <div style={{ display: 'flex', flexDirection: 'column', gap: '5px' }}>
                      {redSpecs.map((spec, idx) => (
                        <div key={idx} style={{ display: 'flex', alignItems: 'center', gap: '4px' }}>
                          <span style={{ fontSize: '11px', color: '#dc2626', fontWeight: 700, minWidth: '14px', textAlign: 'center' }}>
                            •
                          </span>
                          <input
                            type="text"
                            className="option-input-field font-bold"
                            style={{
                              flex: 1,
                              color: '#b91c1c',
                              backgroundColor: '#fff',
                              borderColor: '#fca5a5',
                              fontSize: '11px',
                              padding: '4px 6px',
                              letterSpacing: '0.2px',
                            }}
                            value={spec}
                            onChange={(e) => {
                              const val = e.target.value
                              setRedSpecs((prev) => prev.map((s, i) => (i === idx ? val : s)))
                            }}
                          />
                          <button
                            type="button"
                            className="btn ghost xs"
                            title="Move up"
                            disabled={idx === 0}
                            onClick={() => {
                              if (idx === 0) return
                              setRedSpecs((prev) => {
                                const copy = [...prev]
                                const temp = copy[idx - 1]
                                copy[idx - 1] = copy[idx]
                                copy[idx] = temp
                                return copy
                              })
                            }}
                            style={{ padding: '0 4px', height: '26px', opacity: idx === 0 ? 0.3 : 1 }}
                          >
                            ▲
                          </button>
                          <button
                            type="button"
                            className="btn ghost xs"
                            title="Move down"
                            disabled={idx === redSpecs.length - 1}
                            onClick={() => {
                              if (idx === redSpecs.length - 1) return
                              setRedSpecs((prev) => {
                                const copy = [...prev]
                                const temp = copy[idx + 1]
                                copy[idx + 1] = copy[idx]
                                copy[idx] = temp
                                return copy
                              })
                            }}
                            style={{ padding: '0 4px', height: '26px', opacity: idx === redSpecs.length - 1 ? 0.3 : 1 }}
                          >
                            ▼
                          </button>
                          <button
                            type="button"
                            className="btn ghost xs"
                            title="Delete specification line"
                            onClick={() => {
                              setRedSpecs((prev) => prev.filter((_, i) => i !== idx))
                            }}
                            style={{ padding: '0 6px', height: '26px', color: '#dc2626', borderColor: '#fca5a5' }}
                          >
                            ✕
                          </button>
                        </div>
                      ))}
                    </div>

                    {/* Actions: Add Line & Reset */}
                    <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginTop: '8px', paddingTop: '6px', borderTop: '1px dashed #fecaca' }}>
                      <button
                        type="button"
                        className="btn ghost xs"
                        onClick={() => setRedSpecs((prev) => [...prev, 'NEW SPECIFICATION'])}
                        style={{ fontSize: '11px', color: '#b91c1c', borderColor: '#fca5a5' }}
                      >
                        + Add Line
                      </button>
                      <button
                        type="button"
                        className="btn ghost xs"
                        onClick={() => {
                          const effWidth = hasDoubleDrive ? 3000 : 1830
                          setRedSpecs(generateDefaultRedSpecs(effWidth, projectTrackCount, wb, project.inputs))
                          toast('Reset goods specifications to smart defaults')
                        }}
                        style={{ fontSize: '10.5px', color: '#64748b' }}
                        title="Re-calculate defaults based on current width and tracks"
                      >
                        ↺ Reset Defaults
                      </button>
                    </div>
                  </div>

                  {/* 3. INSTALLATION CHARGES */}
                  <div className="option-group-card">
                    <div className="option-group-title">
                      <span>INSTALLATION CHARGES (Clause 9)</span>
                      <span className={`chip sm ${includeInstallation ? 'ok' : 'muted-chip'}`} style={{ fontSize: '10px' }}>
                        {includeInstallation ? 'Including' : 'At your scope'}
                      </span>
                    </div>
                    <div className="option-pill-group">
                      <button
                        type="button"
                        className={`option-pill ${includeInstallation ? 'active' : ''}`}
                        onClick={() => {
                          setIncludeInstallation(true)
                          setClientInfo((prev) => ({ ...prev, installation_terms_text: 'Including.' }))
                        }}
                      >
                        ✓ Including
                      </button>
                      <button
                        type="button"
                        className={`option-pill ${!includeInstallation ? 'active' : ''}`}
                        onClick={() => {
                          setIncludeInstallation(false)
                          setClientInfo((prev) => ({ ...prev, installation_terms_text: 'At your scope.' }))
                        }}
                      >
                        ✗ At your scope
                      </button>
                    </div>
                    <div>
                      <label className="tiny font-bold muted" style={{ display: 'block', marginBottom: '2px' }}>
                        Clause 9 Full Text (Editable):
                      </label>
                      <input
                        type="text"
                        className="option-input-field"
                        placeholder="Including."
                        value={clientInfo.installation_terms_text || (includeInstallation ? 'Including.' : 'At your scope.')}
                        onChange={(e) => setClientInfo({ ...clientInfo, installation_terms_text: e.target.value })}
                      />
                    </div>
                  </div>

                  {/* 4. FREIGHT / TRANSPORTATION */}
                  <div className="option-group-card">
                    <div className="option-group-title">
                      <span>FREIGHT / TRANSPORT (Clause 8)</span>
                    </div>
                    <div className="option-pill-group">
                      <button
                        type="button"
                        className={`option-pill ${(clientInfo.freight_terms === 'At your scope.' || isTransportZero) ? 'active' : ''}`}
                        onClick={() => setClientInfo({ ...clientInfo, freight_terms: 'At your scope.' })}
                      >
                        At Your Scope {isTransportZero ? '(₹0)' : ''}
                      </button>
                      <button
                        type="button"
                        disabled={isTransportZero}
                        className={`option-pill ${(!isTransportZero && clientInfo.freight_terms === 'Including.') ? 'active' : ''}`}
                        onClick={() => setClientInfo({ ...clientInfo, freight_terms: 'Including.' })}
                        title={isTransportZero ? 'Freight cost is ₹0 in calculator' : ''}
                        style={isTransportZero ? { opacity: 0.5, cursor: 'not-allowed' } : {}}
                      >
                        Including
                      </button>
                      <button
                        type="button"
                        disabled={isTransportZero}
                        className={`option-pill ${(!isTransportZero && clientInfo.freight_terms?.includes('Extra at actuals')) ? 'active' : ''}`}
                        onClick={() => setClientInfo({ ...clientInfo, freight_terms: 'Extra at actuals to your account.' })}
                        style={isTransportZero ? { opacity: 0.5, cursor: 'not-allowed' } : {}}
                      >
                        Extra at Actuals
                      </button>
                      <button
                        type="button"
                        disabled={isTransportZero}
                        className={`option-pill ${(!isTransportZero && clientInfo.freight_terms?.includes('buyer scope') && clientInfo.freight_terms !== 'At your scope.') ? 'active' : ''}`}
                        onClick={() => setClientInfo({ ...clientInfo, freight_terms: 'At buyer scope (Ex-factory Mumbai).' })}
                        style={isTransportZero ? { opacity: 0.5, cursor: 'not-allowed' } : {}}
                      >
                        Ex-Factory
                      </button>
                      <button
                        type="button"
                        disabled={isTransportZero}
                        className={`option-pill ${(!isTransportZero && clientInfo.freight_terms?.includes(clientDestination) && clientDestination !== 'Site') ? 'active' : ''}`}
                        onClick={() => setClientInfo({ ...clientInfo, freight_terms: `Extra at actuals to ${clientDestination || 'Site'}.` })}
                        style={isTransportZero ? { opacity: 0.5, cursor: 'not-allowed' } : {}}
                      >
                        To {clientDestination || 'Site'}
                      </button>
                    </div>
                    <div>
                      <label className="tiny font-bold muted" style={{ display: 'block', marginBottom: '2px' }}>
                        Clause 8 Full Text:
                      </label>
                      <input
                        type="text"
                        className="option-input-field"
                        placeholder="At your scope."
                        value={isTransportZero ? 'At your scope.' : (clientInfo.freight_terms || defaultFreightTerms)}
                        disabled={isTransportZero}
                        readOnly={isTransportZero}
                        style={isTransportZero ? { backgroundColor: 'var(--bg-muted, #f1f5f9)', cursor: 'not-allowed' } : {}}
                        onChange={(e) => setClientInfo({ ...clientInfo, freight_terms: e.target.value })}
                      />
                      {isTransportZero && (
                        <span className="tiny text-muted" style={{ fontSize: '11px', display: 'block', marginTop: '3px' }}>
                          Transportation is ₹0 in calculator — Clause 8 is automatically set to &quot;At your scope.&quot;
                        </span>
                      )}
                    </div>
                  </div>

                  {/* 5. UNLOADING & MATHADI */}
                  <div className="option-group-card">
                    <div className="option-group-title">
                      <span>UNLOADING &amp; MATHADI (Clause 3)</span>
                    </div>
                    <div className="option-pill-group">
                      <button
                        type="button"
                        className={`option-pill ${clientInfo.unloading_terms?.includes('At your scope') || !clientInfo.unloading_terms ? 'active' : ''}`}
                        onClick={() => setClientInfo({ ...clientInfo, unloading_terms: 'At your scope(any mathadi union unloading issue to be managed by buyer & it will not be our responsibility).' })}
                      >
                        At Buyer Scope (Standard)
                      </button>
                      <button
                        type="button"
                        className={`option-pill ${clientInfo.unloading_terms === 'Included in quotation.' ? 'active' : ''}`}
                        onClick={() => setClientInfo({ ...clientInfo, unloading_terms: 'Included in quotation.' })}
                      >
                        Included
                      </button>
                      <button
                        type="button"
                        className={`option-pill ${clientInfo.unloading_terms?.includes('Extra at actuals') ? 'active' : ''}`}
                        onClick={() => setClientInfo({ ...clientInfo, unloading_terms: 'Extra at actuals to your account.' })}
                      >
                        Extra at Actuals
                      </button>
                    </div>
                    <div>
                      <label className="tiny font-bold muted" style={{ display: 'block', marginBottom: '2px' }}>
                        Clause 3 Full Text (Editable):
                      </label>
                      <input
                        type="text"
                        className="option-input-field"
                        value={clientInfo.unloading_terms || 'At your scope(any mathadi union unloading issue to be managed by buyer & it will not be our responsibility).'}
                        onChange={(e) => setClientInfo({ ...clientInfo, unloading_terms: e.target.value })}
                      />
                    </div>
                  </div>

                  {/* 6. LODGING & BOARDING */}
                  <div className="option-group-card">
                    <div className="option-group-title">
                      <span>LODGING &amp; BOARDING (Clause 10)</span>
                    </div>
                    <div className="option-pill-group">
                      <button
                        type="button"
                        className={`option-pill ${(clientInfo.lodging_terms === 'At your scope.' || !clientInfo.lodging_terms || isAccommodationZero) ? 'active' : ''}`}
                        onClick={() => setClientInfo({ ...clientInfo, lodging_terms: 'At your scope.' })}
                      >
                        At Your Scope (Default)
                      </button>
                      <button
                        type="button"
                        disabled={isAccommodationZero}
                        className={`option-pill ${(!isAccommodationZero && clientInfo.lodging_terms === 'Including.') ? 'active' : ''}`}
                        onClick={() => setClientInfo({ ...clientInfo, lodging_terms: 'Including.' })}
                        title={isAccommodationZero ? 'Accommodation cost is ₹0 in calculator' : ''}
                        style={isAccommodationZero ? { opacity: 0.5, cursor: 'not-allowed' } : {}}
                      >
                        Including
                      </button>
                      <button
                        type="button"
                        disabled={isAccommodationZero}
                        className={`option-pill ${(!isAccommodationZero && clientInfo.lodging_terms === 'Not applicable.') ? 'active' : ''}`}
                        onClick={() => setClientInfo({ ...clientInfo, lodging_terms: 'Not applicable.' })}
                        style={isAccommodationZero ? { opacity: 0.5, cursor: 'not-allowed' } : {}}
                      >
                        Not Applicable
                      </button>
                    </div>
                    <div>
                      <label className="tiny font-bold muted" style={{ display: 'block', marginBottom: '2px' }}>
                        Clause 10 Full Text:
                      </label>
                      <input
                        type="text"
                        className="option-input-field"
                        placeholder="At your scope."
                        value={clientInfo.lodging_terms || 'At your scope.'}
                        disabled={isAccommodationZero}
                        readOnly={isAccommodationZero}
                        style={isAccommodationZero ? { backgroundColor: 'var(--bg-muted, #f1f5f9)', cursor: 'not-allowed' } : {}}
                        onChange={(e) => setClientInfo({ ...clientInfo, lodging_terms: e.target.value })}
                      />
                      {isAccommodationZero && (
                        <span className="tiny text-muted" style={{ fontSize: '11px', display: 'block', marginTop: '3px' }}>
                          Accommodation is ₹0 in calculator — Clause 10 is automatically set to &quot;At your scope.&quot;
                        </span>
                      )}
                    </div>
                  </div>

                  {/* 7. PAYMENT TERMS */}
                  <div className="option-group-card">
                    <div className="option-group-title">
                      <span>PAYMENT TERMS (Clause 1)</span>
                    </div>
                    <div className="option-pill-group">
                      <button
                        type="button"
                        className={`option-pill ${clientInfo.payment_terms?.includes('50%Advance') || !clientInfo.payment_terms ? 'active' : ''}`}
                        onClick={() => setClientInfo({ ...clientInfo, payment_terms: '50%Advance payment along with the purchase order,50% payment before material dispatching.' })}
                      >
                        50% Adv + 50% Dispatch
                      </button>
                      <button
                        type="button"
                        className={`option-pill ${clientInfo.payment_terms?.includes('30%') ? 'active' : ''}`}
                        onClick={() => setClientInfo({ ...clientInfo, payment_terms: '30% Advance payment with purchase order, 70% before material dispatching.' })}
                      >
                        30% Adv + 70% Dispatch
                      </button>
                      <button
                        type="button"
                        className={`option-pill ${clientInfo.payment_terms?.includes('100% Advance') ? 'active' : ''}`}
                        onClick={() => setClientInfo({ ...clientInfo, payment_terms: '100% Advance payment against Proforma Invoice before material dispatching.' })}
                      >
                        100% Advance
                      </button>
                      <button
                        type="button"
                        className={`option-pill ${clientInfo.payment_terms?.includes('20%') ? 'active' : ''}`}
                        onClick={() => setClientInfo({ ...clientInfo, payment_terms: '20% Advance payment with purchase order, 80% before material dispatching.' })}
                      >
                        20% Adv + 80% Dispatch
                      </button>
                    </div>
                    <div>
                      <label className="tiny font-bold muted" style={{ display: 'block', marginBottom: '2px' }}>
                        Clause 1 Full Text (Editable):
                      </label>
                      <textarea
                        rows={2}
                        className="option-input-field"
                        value={clientInfo.payment_terms || '50%Advance payment along with the purchase order,50% payment before material dispatching.'}
                        onChange={(e) => setClientInfo({ ...clientInfo, payment_terms: e.target.value })}
                      />
                    </div>
                  </div>

                  {/* 8. QUOTATION VALIDITY */}
                  <div className="option-group-card">
                    <div className="option-group-title">
                      <span>QUOTATION VALIDITY (Clause 6)</span>
                    </div>
                    <div className="option-pill-group">
                      {['Quotation Validity 15 days.', 'Quotation Validity 30 days.', 'Quotation Validity 45 days.', 'Quotation Validity 60 days.'].map((v) => (
                        <button
                          key={v}
                          type="button"
                          className={`option-pill ${clientInfo.validity_terms === v || (!clientInfo.validity_terms && v.includes('15')) ? 'active' : ''}`}
                          onClick={() => setClientInfo({ ...clientInfo, validity_terms: v })}
                        >
                          {v.replace('Quotation Validity ', '').replace('.', '')}
                        </button>
                      ))}
                    </div>
                    <div>
                      <label className="tiny font-bold muted" style={{ display: 'block', marginBottom: '2px' }}>
                        Clause 6 Full Text (Editable):
                      </label>
                      <input
                        type="text"
                        className="option-input-field"
                        value={clientInfo.validity_terms || 'Quotation Validity 15 days.'}
                        onChange={(e) => setClientInfo({ ...clientInfo, validity_terms: e.target.value })}
                      />
                    </div>
                  </div>

                  {/* 9. DELIVERY PERIOD */}
                  <div className="option-group-card">
                    <div className="option-group-title">
                      <span>DELIVERY PERIOD (Clause 2)</span>
                    </div>
                    <div className="option-pill-group">
                      {['4 to 5 WEEKS', '2 to 3 WEEKS', '6 to 8 WEEKS', 'Ready Stock (Immediate)'].map((d) => (
                        <button
                          key={d}
                          type="button"
                          className={`option-pill ${clientInfo.delivery === d || (!clientInfo.delivery && d === '4 to 5 WEEKS') ? 'active' : ''}`}
                          onClick={() => setClientInfo({ ...clientInfo, delivery: d })}
                        >
                          {d}
                        </button>
                      ))}
                    </div>
                    <div>
                      <label className="tiny font-bold muted" style={{ display: 'block', marginBottom: '2px' }}>
                        Clause 2 Delivery Text (Editable):
                      </label>
                      <input
                        type="text"
                        className="option-input-field"
                        value={clientInfo.delivery || '4 to 5 WEEKS'}
                        onChange={(e) => setClientInfo({ ...clientInfo, delivery: e.target.value })}
                      />
                    </div>
                  </div>

                  {/* 10. WARRANTY */}
                  <div className="option-group-card">
                    <div className="option-group-title">
                      <span>WARRANTY (Clause 11)</span>
                    </div>
                    <div className="option-pill-group">
                      {[
                        { label: '12 Months (Standard)', text: '12 months Warrenty Manufacturing defect.' },
                        { label: '24 Months (2 Years)', text: '24 months Warranty against manufacturing defects.' },
                        { label: '36 Months (3 Years)', text: '36 months Extended Warranty against manufacturing defects.' },
                      ].map((w, idx) => (
                        <button
                          key={idx}
                          type="button"
                          className={`option-pill ${clientInfo.warranty_terms === w.text || (!clientInfo.warranty_terms && idx === 0) ? 'active' : ''}`}
                          onClick={() => setClientInfo({ ...clientInfo, warranty_terms: w.text })}
                        >
                          {w.label}
                        </button>
                      ))}
                    </div>
                    <div>
                      <label className="tiny font-bold muted" style={{ display: 'block', marginBottom: '2px' }}>
                        Clause 11 Full Text (Editable):
                      </label>
                      <input
                        type="text"
                        className="option-input-field"
                        value={clientInfo.warranty_terms || '12 months Warrenty Manufacturing defect.'}
                        onChange={(e) => setClientInfo({ ...clientInfo, warranty_terms: e.target.value })}
                      />
                    </div>
                  </div>

                  {/* 11. TAXES & GST */}
                  <div className="option-group-card">
                    <div className="option-group-title">
                      <span>TAXES &amp; GST (Clause 7)</span>
                    </div>
                    <div className="option-pill-group">
                      {[
                        { label: 'Standard (18% Extra)', text: 'Within Maharashtra 9% CGST+9% SGST/outside Maharashtra 18% IGST Extra to your Account.' },
                        { label: 'Inclusive of 18% GST', text: 'All quoted prices are inclusive of 18% GST.' },
                        { label: 'Exempted', text: 'Exempted from GST under current applicable notification.' },
                      ].map((tx, idx) => (
                        <button
                          key={idx}
                          type="button"
                          className={`option-pill ${clientInfo.tax_terms === tx.text || (!clientInfo.tax_terms && idx === 0) ? 'active' : ''}`}
                          onClick={() => setClientInfo({ ...clientInfo, tax_terms: tx.text })}
                        >
                          {tx.label}
                        </button>
                      ))}
                    </div>
                    <div>
                      <label className="tiny font-bold muted" style={{ display: 'block', marginBottom: '2px' }}>
                        Clause 7 Full Text (Editable):
                      </label>
                      <input
                        type="text"
                        className="option-input-field"
                        value={clientInfo.tax_terms || 'Within Maharashtra 9% CGST+9% SGST/outside Maharashtra 18% IGST Extra to your Account.'}
                        onChange={(e) => setClientInfo({ ...clientInfo, tax_terms: e.target.value })}
                      />
                    </div>
                  </div>

                  {/* 12. PACKING CHARGES */}
                  <div className="option-group-card">
                    <div className="option-group-title">
                      <span>PACKING CHARGES (Clause 4)</span>
                    </div>
                    <div className="option-pill-group">
                      {[
                        { label: 'No Extra Charges (Standard)', text: 'No extra charges for packing of materail.' },
                        { label: '2% Extra', text: 'Packing charges extra at 2% on basic value.' },
                        { label: 'Wooden Box Packaging', text: 'Special wooden box packaging extra at actuals.' },
                      ].map((pk, idx) => (
                        <button
                          key={idx}
                          type="button"
                          className={`option-pill ${clientInfo.packing_terms === pk.text || (!clientInfo.packing_terms && idx === 0) ? 'active' : ''}`}
                          onClick={() => setClientInfo({ ...clientInfo, packing_terms: pk.text })}
                        >
                          {pk.label}
                        </button>
                      ))}
                    </div>
                    <div>
                      <label className="tiny font-bold muted" style={{ display: 'block', marginBottom: '2px' }}>
                        Clause 4 Full Text (Editable):
                      </label>
                      <input
                        type="text"
                        className="option-input-field"
                        value={clientInfo.packing_terms || 'No extra charges for packing of materail.'}
                        onChange={(e) => setClientInfo({ ...clientInfo, packing_terms: e.target.value })}
                      />
                    </div>
                  </div>

                  {/* 13. COLOUR / FINISH */}
                  <div className="option-group-card">
                    <div className="option-group-title">
                      <span>COLOUR / FINISH (Clause 5)</span>
                    </div>
                    <div className="option-pill-group">
                      {[
                        { label: 'Standard RAL Shades', text: 'RAL 7032 STR, RAL 7035 STR, LIGHT GREY STR, D.A.GREY STR or AVAILABLE SHADES.' },
                        { label: 'RAL 7035 Only', text: 'RAL 7035 Structure Finish Powder Coated.' },
                        { label: 'Custom Shade Sample', text: 'Powder coated as per approved client color shade sample.' },
                      ].map((cl, idx) => (
                        <button
                          key={idx}
                          type="button"
                          className={`option-pill ${clientInfo.color_terms === cl.text || (!clientInfo.color_terms && idx === 0) ? 'active' : ''}`}
                          onClick={() => setClientInfo({ ...clientInfo, color_terms: cl.text })}
                        >
                          {cl.label}
                        </button>
                      ))}
                    </div>
                    <div>
                      <label className="tiny font-bold muted" style={{ display: 'block', marginBottom: '2px' }}>
                        Clause 5 Full Text (Editable):
                      </label>
                      <input
                        type="text"
                        className="option-input-field"
                        value={clientInfo.color_terms || 'RAL 7032 STR, RAL 7035 STR, LIGHT GREY STR, D.A.GREY STR or AVAILABLE SHADES.'}
                        onChange={(e) => setClientInfo({ ...clientInfo, color_terms: e.target.value })}
                      />
                    </div>
                  </div>

                  {/* OTHER EXTRA AMOUNT */}
                  <div className="option-group-card">
                    <div className="option-group-title">
                      <span>OTHER EXTRA AMOUNT (₹)</span>
                      <span className="chip sm">₹{extraAmount.toLocaleString('en-IN')}</span>
                    </div>
                    <div style={{ display: 'flex', alignItems: 'center', gap: '8px', marginTop: '6px' }}>
                      <input
                        type="number"
                        className="option-input-field font-mono font-bold"
                        value={extraAmount === 0 ? '' : extraAmount}
                        min={0}
                        step={100}
                        placeholder="0"
                        onChange={(e) => {
                          const v = Math.max(0, parseFloat(e.target.value) || 0)
                          setExtraAmount(v)
                        }}
                        style={{ width: '120px' }}
                      />
                      <span className="tiny font-bold text-muted">
                        Adds ₹{extraAmount.toLocaleString('en-IN')} to Unit Rate
                      </span>
                    </div>
                  </div>

                  {/* COMMERCIAL PROFIT MARGIN (Admin / Owner Only) */}
                  {userIsAdmin && (
                    <div className="option-group-card">
                      <div className="option-group-title">
                        <span>COMMERCIAL PROFIT MARGIN (%)</span>
                        <span className="chip sm">{profitPct}%</span>
                      </div>
                      <div style={{ display: 'flex', alignItems: 'center', gap: '8px', marginTop: '6px' }}>
                        <input
                          type="number"
                          className="option-input-field"
                          value={profitPct}
                          min={0}
                          max={100}
                          step={1}
                          onChange={(e) => {
                            const v = parseFloat(e.target.value)
                            setProfitPct(isNaN(v) ? 0 : v)
                          }}
                          style={{ width: '80px', fontWeight: 700 }}
                        />
                        <span className="tiny font-bold">% Profit on Cost Base</span>
                      </div>
                      <div className="option-pill-group" style={{ marginTop: '6px' }}>
                        {[15, 20, 25, 30].map((pct) => (
                          <button
                            key={pct}
                            type="button"
                            className={`option-pill ${profitPct === pct ? 'active' : ''}`}
                            onClick={() => setProfitPct(pct)}
                          >
                            {pct}%
                          </button>
                        ))}
                      </div>
                    </div>
                  )}
                </>
              )}

              {/* TAB: UNIT-WISE COMMERCIAL BREAKDOWN (Admin / Owner Only) */}
              {userIsAdmin && activeSidebarTab === 'breakdown' && (
                <div style={{ display: 'flex', flexDirection: 'column', gap: '12px' }}>
                  <div className="option-group-card">
                    <div className="option-group-title">
                      <span>UNIT-WISE COMMERCIAL CALCULATION</span>
                    </div>

                    {/* KPI Quick Stats */}
                    <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '6px', marginBottom: '8px' }}>
                      <div style={{ background: '#ffffff', padding: '8px', borderRadius: '4px', border: '1px solid #cbd5e1' }}>
                        <div className="tiny text-muted" style={{ fontSize: '10px' }}>Total Extra Charges</div>
                        <div style={{ fontWeight: 700, fontSize: '13px', color: '#1e3a8a' }}>
                          ₹{breakdownResult.totalExtraCharges.toLocaleString('en-IN', { maximumFractionDigits: 2 })}
                        </div>
                        <div className="tiny text-muted" style={{ fontSize: '9.5px', marginTop: '2px' }}>
                          Trans: ₹{breakdownResult.transport} | Inst: ₹{breakdownResult.installation} | Acc: ₹{breakdownResult.accommodation}{extraAmount > 0 ? ` | Extra: ₹${extraAmount}` : ''}
                        </div>
                      </div>
                      <div style={{ background: '#ffffff', padding: '8px', borderRadius: '4px', border: '1px solid #cbd5e1' }}>
                        <div className="tiny text-muted" style={{ fontSize: '10px' }}>Total Bodies &amp; Allocation</div>
                        <div style={{ fontWeight: 700, fontSize: '13px', color: '#1e3a8a' }}>
                          {breakdownResult.totalBodies} Bodies
                        </div>
                        <div className="tiny text-muted" style={{ fontSize: '9.5px', marginTop: '2px' }}>
                          Extra / Body: <strong>₹{breakdownResult.ratePerBody.toFixed(2)}</strong>
                        </div>
                      </div>
                    </div>

                    {/* Live Profit Margin Control inside Breakdown */}
                    <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', background: '#ffffff', padding: '8px 10px', borderRadius: '4px', border: '1px solid #cbd5e1', marginBottom: '8px' }}>
                      <span className="tiny font-bold text-navy">Profit Margin on Cost:</span>
                      <div style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                        <input
                          type="number"
                          className="option-input-field"
                          value={profitPct}
                          min={0}
                          max={100}
                          step={1}
                          onChange={(e) => {
                            const v = parseFloat(e.target.value)
                            setProfitPct(isNaN(v) ? 0 : v)
                          }}
                          style={{ width: '60px', textAlign: 'center', fontWeight: 700, padding: '3px 4px' }}
                        />
                        <span className="tiny font-bold">%</span>
                      </div>
                    </div>

                    {/* Full Table matching user's Excel sheet */}
                    <div style={{ overflowX: 'auto', border: '1px solid #cbd5e1', borderRadius: '4px', background: '#ffffff' }}>
                      <table style={{ width: '100%', fontSize: '10px', borderCollapse: 'collapse', textAlign: 'center' }}>
                        <thead>
                          <tr style={{ background: '#f1f5f9', borderBottom: '1px solid #cbd5e1', fontWeight: 700, color: '#334155', textAlign: 'center' }}>
                            <th style={{ textAlign: 'center', padding: '5px' }}>Model</th>
                            <th style={{ textAlign: 'center', padding: '5px' }}>Qty</th>
                            <th style={{ textAlign: 'center', padding: '5px' }}>Unit Rate</th>
                            <th style={{ textAlign: 'center', padding: '5px' }}>Extra Chrg</th>
                            <th style={{ textAlign: 'center', padding: '5px' }}>Total Unit</th>
                            <th style={{ textAlign: 'center', padding: '5px' }}>Total Price</th>
                            <th style={{ textAlign: 'center', padding: '5px', color: '#1e3a8a' }}>+ Profit ({breakdownResult.profitPct}%)</th>
                          </tr>
                        </thead>
                        <tbody>
                          {breakdownResult.rows.map((r) => (
                            <tr key={r.model_code} style={{ borderBottom: '1px solid #f1f5f9', textAlign: 'center' }}>
                              <td style={{ textAlign: 'center', padding: '5px', fontWeight: 600 }}>
                                {r.model_code}
                                <span style={{ fontSize: '8.5px', color: '#64748b', display: 'block' }}>{r.rows_per_unit} rows/unit</span>
                              </td>
                              <td style={{ textAlign: 'center', padding: '5px' }}>{r.qty}</td>
                              <td style={{ textAlign: 'center', padding: '5px', fontFamily: 'monospace' }}>{Math.round(r.unit_rate).toLocaleString('en-IN')}</td>
                              <td style={{ textAlign: 'center', padding: '5px', fontFamily: 'monospace', color: '#0369a1' }}>{r.transportation.toFixed(0)}</td>
                              <td style={{ textAlign: 'center', padding: '5px', fontFamily: 'monospace' }}>{r.total_unit_rate.toFixed(0)}</td>
                              <td style={{ textAlign: 'center', padding: '5px', fontFamily: 'monospace' }}>{Math.round(r.total_price).toLocaleString('en-IN')}</td>
                              <td style={{ textAlign: 'center', padding: '5px', fontFamily: 'monospace', fontWeight: 700, color: '#1e3a8a' }}>
                                {Math.round(r.total_price_with_profit).toLocaleString('en-IN')}
                              </td>
                            </tr>
                          ))}
                          <tr style={{ background: '#f8fafc', fontWeight: 700, borderTop: '2px solid #cbd5e1', textAlign: 'center' }}>
                            <td colSpan={5} style={{ padding: '6px 5px', textAlign: 'center' }}>Subtotal (Before GST):</td>
                            <td style={{ padding: '6px 5px', fontFamily: 'monospace', textAlign: 'center' }}>₹{Math.round(breakdownResult.totalCostBase).toLocaleString('en-IN')}</td>
                            <td style={{ padding: '6px 5px', fontFamily: 'monospace', color: '#16a34a', textAlign: 'center' }}>₹{Math.round(breakdownResult.totalPriceWithProfit).toLocaleString('en-IN')}</td>
                          </tr>
                        </tbody>
                      </table>
                    </div>

                    {/* Mode switch buttons */}
                    <div style={{ marginTop: '10px' }}>
                      <label className="tiny font-bold muted" style={{ display: 'block', marginBottom: '4px' }}>
                        Display Mode in PDF Document:
                      </label>
                      <div className="option-pill-group" style={{ display: 'grid', gridTemplateColumns: 'repeat(3, 1fr)', gap: '4px' }}>
                        <button
                          type="button"
                          className={`option-pill ${quotationMode === 'unit_wise' ? 'active' : ''}`}
                          onClick={() => setQuotationMode('unit_wise')}
                          style={{ fontSize: '10px', padding: '6px 4px', textAlign: 'center' }}
                        >
                          Standard Unit
                        </button>
                        <button
                          type="button"
                          className={`option-pill ${quotationMode === 'unit_breakdown' ? 'active' : ''}`}
                          onClick={() => setQuotationMode('unit_breakdown')}
                          style={{ fontSize: '10px', padding: '6px 4px', textAlign: 'center' }}
                        >
                          7-Col Breakdown
                        </button>
                        <button
                          type="button"
                          className={`option-pill ${quotationMode === 'set_wise' ? 'active' : ''}`}
                          onClick={() => setQuotationMode('set_wise')}
                          style={{ fontSize: '10px', padding: '6px 4px', textAlign: 'center' }}
                        >
                          Set-Wise
                        </button>
                      </div>
                    </div>

                    <div style={{ display: 'flex', gap: '6px', marginTop: '10px' }}>
                      <button
                        type="button"
                        className="btn primary sm"
                        style={{ fontSize: '11px', flex: 1 }}
                        onClick={handleSyncBreakdownToLineItems}
                      >
                        🔄 Re-sync Line Item Rates
                      </button>
                    </div>
                  </div>
                </div>
              )}

              {/* TAB: CLIENT DETAILS */}
              {activeSidebarTab === 'client' && (
                <div style={{ display: 'flex', flexDirection: 'column', gap: '10px' }}>
                  {/* Client Directory Quick Selector */}
                  <div
                    style={{
                      background: '#f8fafc',
                      border: '1px solid #cbd5e1',
                      borderRadius: '6px',
                      padding: '10px',
                      display: 'flex',
                      flexDirection: 'column',
                      gap: '8px',
                    }}
                  >
                    <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                      <label className="tiny font-bold" style={{ color: '#0f172a', margin: 0 }}>
                        Choose Client
                      </label>
                      <button
                        type="button"
                        className="btn ghost xs"
                        style={{ fontSize: '11px', fontWeight: 600, padding: '2px 8px', border: '1px solid #cbd5e1', background: '#ffffff', color: '#0f172a' }}
                        onClick={() => setShowAddClientModal(true)}
                        title="Add a new client to directory (Staff & Admin)"
                      >
                        + New Client
                      </button>
                    </div>

                    <input
                      type="text"
                      placeholder="Type client name or GST to search..."
                      value={clientSearchText}
                      onChange={(e) => setClientSearchText(e.target.value)}
                      className="option-input-field"
                      style={{ background: '#ffffff', borderColor: '#cbd5e1', fontSize: '12px', height: '30px', padding: '0 8px' }}
                    />

                    <select
                      className="option-input-field font-bold"
                      value={selectedClientId}
                      onChange={(e) => {
                        handleSelectSavedClient(e.target.value)
                        const found = visibleSavedClients.find((c) => c.id === e.target.value)
                        if (found) {
                          setClientSearchText(found.name)
                        }
                      }}
                      style={{ background: '#ffffff', borderColor: '#cbd5e1', fontSize: '12px' }}
                    >
                      <option value="">
                        {filteredSavedClients.length === 0
                          ? '-- No matching clients found --'
                          : clientSearchText.trim()
                          ? `-- Matching Clients (${filteredSavedClients.length}) --`
                          : `-- Select Saved Client (${visibleSavedClients.length}) --`}
                      </option>
                      {filteredSavedClients.map((c) => (
                        <option key={c.id} value={c.id}>
                          {c.name} {userIsAdmin && c.profit_margin_pct !== undefined ? `(${c.profit_margin_pct}% Margin)` : ''} {c.gst ? `· GST: ${c.gst}` : ''}
                        </option>
                      ))}
                    </select>

                    <div style={{ display: 'flex', gap: '6px', flexWrap: 'wrap' }}>
                      <button
                        type="button"
                        className="btn ghost xs"
                        style={{ flex: 1, fontSize: '11px', padding: '4px 6px', whiteSpace: 'nowrap', border: '1px solid #cbd5e1', background: '#ffffff', color: '#0f172a' }}
                        onClick={handleSaveCurrentAsClient}
                        title="Save current details as a new client in directory"
                      >
                        Save as New
                      </button>
                      {selectedClientId && (
                        <>
                          <button
                            type="button"
                            className="btn ghost xs"
                            style={{ fontSize: '11px', padding: '4px 6px', whiteSpace: 'nowrap', border: '1px solid #cbd5e1', background: '#ffffff', color: '#0f172a' }}
                            onClick={handleUpdateSelectedClient}
                            title="Update this client in directory with current edits"
                          >
                            Update Directory
                          </button>
                          <button
                            type="button"
                            className="btn ghost xs"
                            style={{ fontSize: '11px', padding: '4px 6px', color: '#475569', border: '1px solid #cbd5e1', background: '#ffffff' }}
                            onClick={handleDeleteSelectedClient}
                            title="Remove this client from directory"
                          >
                            Delete
                          </button>
                        </>
                      )}
                    </div>
                  </div>

                  <div>
                    <label className="tiny font-bold muted">QTN Number</label>
                    <input
                      type="text"
                      className="option-input-field"
                      value={clientInfo.qtn_no}
                      onChange={(e) => setClientInfo({ ...clientInfo, qtn_no: e.target.value })}
                    />
                  </div>

                  <div>
                    <label className="tiny font-bold muted">Quotation Date</label>
                    <input
                      type="text"
                      className="option-input-field"
                      value={clientInfo.date}
                      onChange={(e) => setClientInfo({ ...clientInfo, date: e.target.value })}
                    />
                  </div>

                  <div>
                    <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '2px' }}>
                      <label className="tiny font-bold muted">Buyer / Company Name</label>
                      <label style={{ fontSize: '11px', display: 'flex', alignItems: 'center', gap: '4px', cursor: 'pointer', color: '#475569', fontWeight: 600 }}>
                        <input
                          type="checkbox"
                          checked={clientInfo.show_buyer_name !== false}
                          onChange={(e) => setClientInfo({ ...clientInfo, show_buyer_name: e.target.checked })}
                        />
                        <span>Show on PDF</span>
                      </label>
                    </div>
                    <input
                      type="text"
                      className="option-input-field font-bold"
                      value={clientInfo.buyer_name}
                      onChange={(e) => setClientInfo({ ...clientInfo, buyer_name: e.target.value })}
                    />
                  </div>

                  <div>
                    <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '2px' }}>
                      <label className="tiny font-bold muted">Buyer Address</label>
                      <label style={{ fontSize: '11px', display: 'flex', alignItems: 'center', gap: '4px', cursor: 'pointer', color: '#475569', fontWeight: 600 }}>
                        <input
                          type="checkbox"
                          checked={clientInfo.show_buyer_address !== false}
                          onChange={(e) => setClientInfo({ ...clientInfo, show_buyer_address: e.target.checked })}
                        />
                        <span>Show on PDF</span>
                      </label>
                    </div>
                    <textarea
                      rows={3}
                      className="option-input-field"
                      value={clientInfo.buyer_address}
                      onChange={(e) => setClientInfo({ ...clientInfo, buyer_address: e.target.value })}
                    />
                  </div>

                  <div>
                    <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '2px' }}>
                      <label className="tiny font-bold muted">GST Number</label>
                      <label style={{ fontSize: '11px', display: 'flex', alignItems: 'center', gap: '4px', cursor: 'pointer', color: '#475569', fontWeight: 600 }}>
                        <input
                          type="checkbox"
                          checked={clientInfo.show_buyer_gst !== false}
                          onChange={(e) => setClientInfo({ ...clientInfo, show_buyer_gst: e.target.checked })}
                        />
                        <span>Show on PDF</span>
                      </label>
                    </div>
                    <input
                      type="text"
                      className="option-input-field font-mono"
                      value={clientInfo.buyer_gst}
                      onChange={(e) => setClientInfo({ ...clientInfo, buyer_gst: e.target.value })}
                    />
                  </div>

                  <div>
                    <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '2px' }}>
                      <label className="tiny font-bold muted">Buyer Email Address</label>
                      <label style={{ fontSize: '11px', display: 'flex', alignItems: 'center', gap: '4px', cursor: 'pointer', color: '#475569', fontWeight: 600 }}>
                        <input
                          type="checkbox"
                          checked={clientInfo.show_buyer_email !== false}
                          onChange={(e) => setClientInfo({ ...clientInfo, show_buyer_email: e.target.checked })}
                        />
                        <span>Show on PDF</span>
                      </label>
                    </div>
                    <input
                      type="email"
                      className="option-input-field"
                      placeholder="e.g. stores@company.com"
                      value={clientInfo.buyer_email || ''}
                      onChange={(e) => setClientInfo({ ...clientInfo, buyer_email: e.target.value, show_buyer_email: true })}
                    />
                  </div>

                  {/* Primary Contact: Single paired unit with Name dropdown and paired Contact Number */}
                  <div>
                    <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '4px' }}>
                      <label className="tiny font-bold muted">Primary Contact (Kind Attention &amp; Contact Number)</label>
                      <label style={{ fontSize: '11px', display: 'flex', alignItems: 'center', gap: '4px', cursor: 'pointer', color: '#475569', fontWeight: 600 }}>
                        <input
                          type="checkbox"
                          checked={clientInfo.show_kind_att !== false || clientInfo.show_buyer_contact !== false}
                          onChange={(e) =>
                            setClientInfo({
                              ...clientInfo,
                              show_kind_att: e.target.checked,
                              show_buyer_contact: e.target.checked,
                            })
                          }
                        />
                        <span>Show on PDF</span>
                      </label>
                    </div>

                    {/* Dropdown for Name: selecting a contact person automatically updates both name and mobile number */}
                    {contactPersonOptions.length > 0 && (
                      <div style={{ marginBottom: '6px' }}>
                        <select
                          className="option-input-field font-bold"
                          value={
                            contactPersonOptions.find(
                              (opt) => opt.name.toLowerCase() === (clientInfo.kind_att || '').trim().toLowerCase()
                            )?.id || ''
                          }
                          onChange={(e) => {
                            const chosen = (contactPersonOptions as any[]).find((opt) => opt.id === e.target.value)
                            if (chosen) {
                              setClientInfo({
                                ...clientInfo,
                                kind_att: chosen.name,
                                buyer_contact: chosen.phone,
                                ...(chosen.email ? { buyer_email: chosen.email, show_buyer_email: true } : {}),
                              })
                            }
                          }}
                          style={{ width: '100%', fontSize: '12px' }}
                        >
                          <option value="">-- Choose Contact Person by Name --</option>
                          {contactPersonOptions.map((opt) => (
                            <option key={opt.id} value={opt.id}>
                              {opt.label}
                            </option>
                          ))}
                        </select>
                      </div>
                    )}

                    {/* Single paired unit: Kind Att Name & Contact Number side-by-side */}
                    <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '8px' }}>
                      <div>
                        <label className="tiny font-bold muted" style={{ display: 'block', marginBottom: '2px' }}>
                          Kind Attention (Name)
                        </label>
                        <input
                          type="text"
                          placeholder="e.g. xyz / Mr. Rajesh Sharma"
                          className="option-input-field font-bold"
                          value={clientInfo.kind_att}
                          onChange={(e) => setClientInfo({ ...clientInfo, kind_att: e.target.value })}
                          style={{ width: '100%', height: '34px', padding: '0 8px' }}
                        />
                      </div>
                      <div>
                        <label className="tiny font-bold muted" style={{ display: 'block', marginBottom: '2px' }}>
                          Contact Number
                        </label>
                        <input
                          type="text"
                          placeholder="e.g. 123 / +91 9876543210"
                          className="option-input-field"
                          value={clientInfo.buyer_contact}
                          onChange={(e) => setClientInfo({ ...clientInfo, buyer_contact: e.target.value })}
                          style={{ width: '100%', height: '34px', padding: '0 8px' }}
                        />
                      </div>
                    </div>
                  </div>

                  {/* Additional Contacts: simple flat inputs */}
                  {(clientInfo.contacts || []).map((cp, idx) => (
                    <div
                      key={cp.id}
                      style={{
                        display: 'grid',
                        gridTemplateColumns: '1fr 1fr auto',
                        gap: '8px',
                        alignItems: 'flex-end',
                      }}
                    >
                      <div>
                        <label className="tiny font-bold muted" style={{ display: 'block', marginBottom: '2px' }}>
                          Contact Person Name
                        </label>
                        <input
                          type="text"
                          className="option-input-field"
                          placeholder="e.g. xyz"
                          value={cp.name}
                          onChange={(e) => {
                            const next = (clientInfo.contacts || []).map((c, i) =>
                              i === idx ? { ...c, name: e.target.value } : c
                            )
                            setClientInfo({ ...clientInfo, contacts: next })
                          }}
                          style={{ height: '34px', padding: '0 8px', width: '100%' }}
                        />
                      </div>
                      <div>
                        <label className="tiny font-bold muted" style={{ display: 'block', marginBottom: '2px' }}>
                          Mobile Number
                        </label>
                        <input
                          type="text"
                          className="option-input-field"
                          placeholder="e.g. 123"
                          value={cp.phone}
                          onChange={(e) => {
                            const next = (clientInfo.contacts || []).map((c, i) =>
                              i === idx ? { ...c, phone: e.target.value } : c
                            )
                            setClientInfo({ ...clientInfo, contacts: next })
                          }}
                          style={{ height: '34px', padding: '0 8px', width: '100%' }}
                        />
                      </div>
                      <button
                        type="button"
                        className="btn ghost xs"
                        onClick={() => {
                          const next = (clientInfo.contacts || []).filter((_, i) => i !== idx)
                          setClientInfo({ ...clientInfo, contacts: next })
                        }}
                        style={{ height: '34px', padding: '0 10px', color: '#dc2626', border: '1px solid #cbd5e1' }}
                        title="Remove contact"
                      >
                        ✕
                      </button>
                    </div>
                  ))}

                  {/* Custom Rows: User can edit both label and value, and choose Share in PDF */}
                  {(clientInfo.extra_fields || []).map((ef, idx) => (
                    <div
                      key={ef.id}
                      style={{
                        display: 'grid',
                        gridTemplateColumns: '1fr 1fr auto auto',
                        gap: '8px',
                        alignItems: 'flex-end',
                      }}
                    >
                      <div>
                        <label className="tiny font-bold muted" style={{ display: 'block', marginBottom: '2px' }}>
                          Row Label
                        </label>
                        <input
                          type="text"
                          className="option-input-field"
                          placeholder="e.g. xyz"
                          value={ef.label}
                          onChange={(e) => {
                            const next = (clientInfo.extra_fields || []).map((f, i) =>
                              i === idx ? { ...f, label: e.target.value } : f
                            )
                            setClientInfo({ ...clientInfo, extra_fields: next })
                          }}
                          style={{ height: '34px', padding: '0 8px', width: '100%' }}
                        />
                      </div>
                      <div>
                        <label className="tiny font-bold muted" style={{ display: 'block', marginBottom: '2px' }}>
                          Row Value
                        </label>
                        <input
                          type="text"
                          className="option-input-field"
                          placeholder="e.g. xyz"
                          value={ef.value}
                          onChange={(e) => {
                            const next = (clientInfo.extra_fields || []).map((f, i) =>
                              i === idx ? { ...f, value: e.target.value, values: [e.target.value] } : f
                            )
                            setClientInfo({ ...clientInfo, extra_fields: next })
                          }}
                          style={{ height: '34px', padding: '0 8px', width: '100%' }}
                        />
                      </div>
                      <label
                        style={{
                          display: 'flex',
                          alignItems: 'center',
                          gap: '4px',
                          fontSize: '11px',
                          cursor: 'pointer',
                          whiteSpace: 'nowrap',
                          height: '34px',
                          color: '#475569',
                          fontWeight: 600,
                        }}
                      >
                        <input
                          type="checkbox"
                          checked={ef.visible !== false}
                          onChange={(e) => {
                            const next = (clientInfo.extra_fields || []).map((f, i) =>
                              i === idx ? { ...f, visible: e.target.checked } : f
                            )
                            setClientInfo({ ...clientInfo, extra_fields: next })
                          }}
                        />
                        <span>Show on PDF</span>
                      </label>
                      <button
                        type="button"
                        className="btn ghost xs"
                        onClick={() => {
                          const next = (clientInfo.extra_fields || []).filter((_, i) => i !== idx)
                          setClientInfo({ ...clientInfo, extra_fields: next })
                        }}
                        style={{ height: '34px', padding: '0 10px', color: '#dc2626', border: '1px solid #cbd5e1' }}
                        title="Remove custom row"
                      >
                        ✕
                      </button>
                    </div>
                  ))}

                  {/* Subject Line */}
                  <div>
                    <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '2px' }}>
                      <label className="tiny font-bold muted">Subject Line</label>
                      <label style={{ fontSize: '11px', display: 'flex', alignItems: 'center', gap: '4px', cursor: 'pointer', color: '#475569', fontWeight: 600 }}>
                        <input
                          type="checkbox"
                          checked={clientInfo.show_subject !== false}
                          onChange={(e) => setClientInfo({ ...clientInfo, show_subject: e.target.checked })}
                        />
                        <span>Show on PDF</span>
                      </label>
                    </div>
                    <input
                      type="text"
                      className="option-input-field"
                      value={clientInfo.subject}
                      onChange={(e) => setClientInfo({ ...clientInfo, subject: e.target.value })}
                      style={{ width: '100%' }}
                    />
                  </div>

                  {/* Buttons for add contact and custom rows at the very last */}
                  <div style={{ display: 'flex', gap: '8px', paddingTop: '6px' }}>
                    <button
                      type="button"
                      className="btn ghost sm"
                      onClick={() => {
                        const newCp = {
                          id: (typeof crypto !== 'undefined' && crypto.randomUUID) ? `cp_${crypto.randomUUID()}` : `cp_${Date.now()}`,
                          name: '',
                          phone: '',
                        };
                        setClientInfo({
                          ...clientInfo,
                          contacts: [...(clientInfo.contacts || []), newCp],
                        });
                      }}
                      style={{ border: '1px solid #0f172a', color: '#0f172a', fontWeight: 600, fontSize: '12px' }}
                    >
                      + Add Contact
                    </button>
                    <button
                      type="button"
                      className="btn ghost sm"
                      onClick={() => {
                        const newEf = {
                          id: (typeof crypto !== 'undefined' && crypto.randomUUID) ? `ef_${crypto.randomUUID()}` : `ef_${Date.now()}`,
                          label: '',
                          value: '',
                          values: [''],
                          visible: true,
                          visible_values: [true],
                        };
                        setClientInfo({
                          ...clientInfo,
                          extra_fields: [...(clientInfo.extra_fields || []), newEf],
                        });
                      }}
                      style={{ border: '1px solid #0f172a', color: '#0f172a', fontWeight: 600, fontSize: '12px' }}
                    >
                      + Add Custom Row
                    </button>
                  </div>
                </div>
              )}

                            {/* TAB: ALL TERMS & CONDITIONS */}
              {activeSidebarTab === 'terms' && (
                <div style={{ display: 'flex', flexDirection: 'column', gap: '12px' }}>
                  <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                    <div>
                      <span className="tiny font-bold">Terms &amp; Conditions ({termsList.length})</span>
                    </div>
                    <button
                      type="button"
                      className="btn ghost sm"
                      onClick={() => {
                        const newId = String(Date.now())
                        setTermsList((prev) => [
                          ...prev,
                          {
                            id: newId,
                            title: `Clause ${prev.length + 1} :-`,
                            text: '',
                          },
                        ])
                      }}
                      style={{ fontSize: '11px', padding: '3px 8px' }}
                    >
                      + Add Term
                    </button>
                  </div>

                  {termsList.map((t, idx) => (
                    <div
                      key={t.id}
                      style={{
                        border: '1px solid #cbd5e1',
                        borderRadius: '6px',
                        padding: '10px',
                        background: '#f8fafc',
                      }}
                    >
                      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '6px', gap: '8px' }}>
                        <div style={{ display: 'flex', alignItems: 'center', gap: '6px', flex: 1 }}>
                          <span className="tiny font-bold text-muted" style={{ minWidth: '22px' }}>
                            {idx + 1})
                          </span>
                          <input
                            type="text"
                            value={t.title}
                            placeholder="Clause Title (e.g. Payment :-)"
                            onChange={(e) => {
                              const val = e.target.value
                              setTermsList((prev) =>
                                prev.map((item, i) => (i === idx ? { ...item, title: val } : item))
                              )
                            }}
                            style={{
                              flex: 1,
                              fontWeight: 700,
                              fontSize: '11.5px',
                              padding: '3px 6px',
                              border: '1px solid #cbd5e1',
                              borderRadius: '4px',
                            }}
                          />
                        </div>
                        {termsList.length > 1 && (
                          <button
                            type="button"
                            className="icon-btn sm"
                            onClick={() => {
                              setTermsList((prev) => prev.filter((_, i) => i !== idx))
                            }}
                            title="Delete this term"
                            style={{ color: '#ef4444' }}
                          >
                            {Icon.trash}
                          </button>
                        )}
                      </div>
                      {(() => {
                        const isFreight = t.id === '8' || t.title.toLowerCase().includes('freight') || t.title.toLowerCase().includes('transport')
                        const isLodging = t.id === '10' || t.title.toLowerCase().includes('lodging') || t.title.toLowerCase().includes('loadging') || t.title.toLowerCase().includes('bording') || t.title.toLowerCase().includes('boarding')
                        const isLockedScope = (isLodging && isAccommodationZero) || (isFreight && isTransportZero)
                        return (
                          <>
                            <textarea
                              rows={2}
                              className="option-input-field"
                              value={isLockedScope ? 'At your scope.' : t.text}
                              disabled={isLockedScope}
                              readOnly={isLockedScope}
                              style={isLockedScope ? { backgroundColor: 'var(--bg-muted, #f1f5f9)', cursor: 'not-allowed', fontSize: '11px', lineHeight: 1.4 } : { fontSize: '11px', lineHeight: 1.4 }}
                              placeholder="Enter terms clause text..."
                              onChange={(e) => {
                                const val = e.target.value
                                setTermsList((prev) =>
                                  prev.map((item, i) => (i === idx ? { ...item, text: val } : item))
                                )
                              }}
                            />
                            {isLodging && isAccommodationZero && (
                              <span className="tiny text-muted" style={{ fontSize: '10.5px', marginTop: '2px', display: 'block' }}>
                                Accommodation is ₹0 in calculator — Clause 10 is automatically &quot;At your scope.&quot;
                              </span>
                            )}
                            {isFreight && isTransportZero && (
                              <span className="tiny text-muted" style={{ fontSize: '10.5px', marginTop: '2px', display: 'block' }}>
                                Transportation is ₹0 in calculator — Clause 8 is automatically &quot;At your scope.&quot;
                              </span>
                            )}
                          </>
                        )
                      })()}
                    </div>
                  ))}

                  <button
                    type="button"
                    className="btn secondary sm"
                    onClick={() => {
                      const newId = String(Date.now())
                      setTermsList((prev) => [
                        ...prev,
                        {
                          id: newId,
                          title: `Clause ${prev.length + 1} :-`,
                          text: '',
                        },
                      ])
                    }}
                    style={{ marginTop: '4px' }}
                  >
                    + Add New Term
                  </button>
                </div>
              )}

              {/* TAB: INSTALLATION BODIES */}
              {activeSidebarTab === 'body' && (
                <div style={{ display: 'flex', flexDirection: 'column', gap: '10px' }}>
                  <div style={{ display: 'flex', justifyContent: 'flex-end', alignItems: 'center' }}>
                    <button
                      type="button"
                      className="btn secondary"
                      style={{ fontSize: '11px', padding: '3px 8px' }}
                      onClick={() => {
                        const fresh = generateDefaultBodyItems(summary, wb, project.inputs)
                        setBodyItems(fresh)
                        toast('Bodies reset')
                      }}
                      title="Reset installation bodies from active model rows & quantities"
                    >
                      🔄 Reset from Models
                    </button>
                  </div>

                  {bodyItems.map((b, idx) => (
                    <div key={b.id} style={{ background: '#f8fafc', border: '1px solid #cbd5e1', borderRadius: '6px', padding: '10px' }}>
                      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '6px' }}>
                        <input
                          type="text"
                          value={b.title}
                          onChange={(e) => {
                            const updated = bodyItems.map((item, i) => i === idx ? { ...item, title: e.target.value } : item)
                            setBodyItems(updated)
                          }}
                          style={{ fontSize: '12px', fontWeight: 600, border: '1px solid #cbd5e1', borderRadius: '3px', padding: '2px 4px', width: '60%' }}
                        />
                        <div style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                          <span className="tiny font-bold" style={{ color: '#0369a1' }}>{b.total_body} Bodies</span>
                          <button
                            type="button"
                            disabled={idx === 0}
                            onClick={() => moveBodyItem(idx, idx - 1)}
                            style={{
                              padding: '2px 5px',
                              fontSize: '10px',
                              cursor: idx === 0 ? 'not-allowed' : 'pointer',
                              opacity: idx === 0 ? 0.4 : 1,
                              border: '1px solid #cbd5e1',
                              borderRadius: '3px',
                              background: '#ffffff',
                            }}
                            title="Move Body Row Up"
                          >
                            ▲
                          </button>
                          <button
                            type="button"
                            disabled={idx === bodyItems.length - 1}
                            onClick={() => moveBodyItem(idx, idx + 1)}
                            style={{
                              padding: '2px 5px',
                              fontSize: '10px',
                              cursor: idx === bodyItems.length - 1 ? 'not-allowed' : 'pointer',
                              opacity: idx === bodyItems.length - 1 ? 0.4 : 1,
                              border: '1px solid #cbd5e1',
                              borderRadius: '3px',
                              background: '#ffffff',
                            }}
                            title="Move Body Row Down"
                          >
                            ▼
                          </button>
                          {bodyItems.length > 1 && (
                            <button
                              type="button"
                              onClick={() => setBodyItems(bodyItems.filter((_, i) => i !== idx))}
                              style={{ background: 'none', border: 'none', color: '#ef4444', cursor: 'pointer', fontWeight: 700, padding: '0 4px', fontSize: '12px' }}
                              title="Delete row"
                            >
                              ✕
                            </button>
                          )}
                        </div>
                      </div>
                      <div style={{ marginBottom: '6px' }}>
                        <input
                          type="text"
                          value={b.sub_label}
                          onChange={(e) => {
                            const updated = bodyItems.map((item, i) => i === idx ? { ...item, sub_label: e.target.value } : item)
                            setBodyItems(updated)
                          }}
                          style={{ fontSize: '10.5px', color: '#64748b', border: '1px solid #e2e8f0', borderRadius: '3px', padding: '2px 4px', width: '100%' }}
                        />
                      </div>
                      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(2, 1fr)', gap: '6px' }}>
                        <div>
                          <label className="tiny muted">Per Bay Qty (Rows)</label>
                          <input
                            type="number"
                            className="option-input-field"
                            value={b.per_bay_qty}
                            onChange={(e) => {
                              const val = parseInt(e.target.value) || 0
                              const updated = bodyItems.map((item, i) => {
                                if (i === idx) {
                                  const unit_body = val * item.qty
                                  const set_qty = item.set_qty || 1
                                  return { ...item, per_bay_qty: val, unit_body, total_body: unit_body * set_qty }
                                }
                                return item
                              })
                              setBodyItems(updated)
                            }}
                          />
                        </div>
                        <div>
                          <label className="tiny muted">Qty (Units)</label>
                          <input
                            type="number"
                            className="option-input-field"
                            value={b.qty}
                            onChange={(e) => {
                              const val = parseInt(e.target.value) || 0
                              const updated = bodyItems.map((item, i) => {
                                if (i === idx) {
                                  const unit_body = item.per_bay_qty * val
                                  const set_qty = item.set_qty || 1
                                  return { ...item, qty: val, unit_body, total_body: unit_body * set_qty }
                                }
                                return item
                              })
                              setBodyItems(updated)
                            }}
                          />
                        </div>
                      </div>
                    </div>
                  ))}

                  <button
                    type="button"
                    className="btn secondary"
                    style={{ fontSize: '12px', justifyContent: 'center', padding: '6px' }}
                    onClick={() => {
                      const newId = String(Date.now())
                      let dfmuRows = 6
                      if (wb) {
                        dfmuRows = wb.num('Total cal', 'A3') || 6
                      } else if (project.inputs?.['Total cal!A3']) {
                        dfmuRows = Number(project.inputs['Total cal!A3']) || 6
                      }
                      const bay = Math.max(1, Math.round(dfmuRows / 2))
                      setBodyItems([
                        ...bodyItems,
                        {
                          id: newId,
                          title: `CUSTOM ${bay} BAY`,
                          sub_label: 'BODY IN CUSTOM',
                          per_bay_qty: dfmuRows,
                          qty: 1,
                          total_body: dfmuRows,
                        },
                      ])
                    }}
                  >
                    + Add Body Item Row
                  </button>

                  <div style={{ background: '#f8fafc', padding: '6px 10px', borderRadius: '4px', border: '1px solid #cbd5e1', fontSize: '11px', fontWeight: 600 }}>
                    Total Bodies: {totalBodies}
                  </div>
                </div>
              )}

              {/* TAB: TECHNICAL PROPOSAL SPECIFICATIONS (Pages 3 to 5) */}
              {activeSidebarTab === 'specs' && (
                <div style={{ display: 'flex', flexDirection: 'column', gap: '12px' }}>
                  <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                    <div>
                      <span className="tiny font-bold">Proposal Tech Specs (Pages 3-5)</span>
                    </div>
                    <button
                      type="button"
                      className="btn ghost xs"
                      style={{ fontSize: '11px', padding: '3px 8px', color: '#dc2626', border: '1px solid #fca5a5', background: '#fef2f2' }}
                      onClick={() => {
                        if (window.confirm('Reset all technical proposal specifications back to factory defaults?')) {
                          setProposalSpecs(DEFAULT_PROPOSAL_SPECS)
                          toast('Technical proposal specs reset to defaults!')
                        }
                      }}
                      title="Reset all fields to factory defaults"
                    >
                      ↺ Reset to Default
                    </button>
                  </div>

                  {/* Sub-tab Pills for Pages 3, 4, 5 */}
                  <div style={{ display: 'flex', gap: '4px', background: '#f1f5f9', padding: '3px', borderRadius: '6px' }}>
                    <button
                      type="button"
                      className={`option-pill ${specsSubTab === 'page3' ? 'active' : ''}`}
                      style={{ flex: 1, fontSize: '11px', padding: '5px 4px', textAlign: 'center' }}
                      onClick={() => setSpecsSubTab('page3')}
                    >
                      Page 3: Compactor
                    </button>
                    <button
                      type="button"
                      className={`option-pill ${specsSubTab === 'page4' ? 'active' : ''}`}
                      style={{ flex: 1, fontSize: '11px', padding: '5px 4px', textAlign: 'center' }}
                      onClick={() => setSpecsSubTab('page4')}
                    >
                      Page 4: Powder
                    </button>
                    <button
                      type="button"
                      className={`option-pill ${specsSubTab === 'page5' ? 'active' : ''}`}
                      style={{ flex: 1, fontSize: '11px', padding: '5px 4px', textAlign: 'center' }}
                      onClick={() => setSpecsSubTab('page5')}
                    >
                      Page 5: QC
                    </button>
                  </div>

                  {/* PAGE 3 EDITOR */}
                  {specsSubTab === 'page3' && (
                    <div style={{ display: 'flex', flexDirection: 'column', gap: '10px' }}>
                      <div className="option-group-card">
                        <span className="tiny font-bold" style={{ display: 'block', marginBottom: '6px', color: '#1e293b' }}>
                          Page 3 Titles &amp; Introduction
                        </span>
                        <div style={{ display: 'flex', flexDirection: 'column', gap: '8px' }}>
                          <div>
                            <label className="tiny text-muted font-bold">Main Title</label>
                            <input
                              type="text"
                              className="option-input-field"
                              value={proposalSpecs.page3.mainTitle}
                              onChange={(e) => setProposalSpecs((prev) => ({ ...prev, page3: { ...prev.page3, mainTitle: e.target.value } }))}
                            />
                          </div>
                          <div>
                            <label className="tiny text-muted font-bold">Sub Title</label>
                            <input
                              type="text"
                              className="option-input-field"
                              value={proposalSpecs.page3.subTitle}
                              onChange={(e) => setProposalSpecs((prev) => ({ ...prev, page3: { ...prev.page3, subTitle: e.target.value } }))}
                            />
                          </div>
                          <div>
                            <label className="tiny text-muted font-bold">Mobile Racks Intro</label>
                            <textarea
                              className="option-input-field"
                              rows={2}
                              value={proposalSpecs.page3.mobileRacksIntro}
                              onChange={(e) => setProposalSpecs((prev) => ({ ...prev, page3: { ...prev.page3, mobileRacksIntro: e.target.value } }))}
                            />
                          </div>
                        </div>
                      </div>

                      <div className="option-group-card">
                        <span className="tiny font-bold" style={{ display: 'block', marginBottom: '6px', color: '#1e293b' }}>
                          The Undercarriage
                        </span>
                        <div style={{ display: 'flex', flexDirection: 'column', gap: '8px' }}>
                          <div>
                            <label className="tiny text-muted font-bold">Base Frame</label>
                            <textarea
                              className="option-input-field"
                              rows={2}
                              value={proposalSpecs.page3.baseFrame}
                              onChange={(e) => setProposalSpecs((prev) => ({ ...prev, page3: { ...prev.page3, baseFrame: e.target.value } }))}
                            />
                          </div>
                          <div>
                            <label className="tiny text-muted font-bold">Tracks</label>
                            <textarea
                              className="option-input-field"
                              rows={2}
                              value={proposalSpecs.page3.tracks}
                              onChange={(e) => setProposalSpecs((prev) => ({ ...prev, page3: { ...prev.page3, tracks: e.target.value } }))}
                            />
                          </div>
                        </div>
                      </div>

                      <div className="option-group-card">
                        <span className="tiny font-bold" style={{ display: 'block', marginBottom: '6px', color: '#1e293b' }}>
                          Super Structure
                        </span>
                        <div style={{ display: 'flex', flexDirection: 'column', gap: '8px' }}>
                          <div>
                            <label className="tiny text-muted font-bold">Super Structure Intro</label>
                            <textarea
                              className="option-input-field"
                              rows={2}
                              value={proposalSpecs.page3.superStructureIntro}
                              onChange={(e) => setProposalSpecs((prev) => ({ ...prev, page3: { ...prev.page3, superStructureIntro: e.target.value } }))}
                            />
                          </div>
                          <div>
                            <label className="tiny text-muted font-bold">Shelves</label>
                            <textarea
                              className="option-input-field"
                              rows={2}
                              value={proposalSpecs.page3.shelves}
                              onChange={(e) => setProposalSpecs((prev) => ({ ...prev, page3: { ...prev.page3, shelves: e.target.value } }))}
                            />
                          </div>
                          <div>
                            <label className="tiny text-muted font-bold">Vertical Panel</label>
                            <textarea
                              className="option-input-field"
                              rows={2}
                              value={proposalSpecs.page3.verticalPanel}
                              onChange={(e) => setProposalSpecs((prev) => ({ ...prev, page3: { ...prev.page3, verticalPanel: e.target.value } }))}
                            />
                          </div>
                          <div>
                            <label className="tiny text-muted font-bold">Middle Panel</label>
                            <textarea
                              className="option-input-field"
                              rows={2}
                              value={proposalSpecs.page3.middlePanel}
                              onChange={(e) => setProposalSpecs((prev) => ({ ...prev, page3: { ...prev.page3, middlePanel: e.target.value } }))}
                            />
                          </div>
                          <div>
                            <label className="tiny text-muted font-bold">Top Panel</label>
                            <textarea
                              className="option-input-field"
                              rows={2}
                              value={proposalSpecs.page3.topPanel}
                              onChange={(e) => setProposalSpecs((prev) => ({ ...prev, page3: { ...prev.page3, topPanel: e.target.value } }))}
                            />
                          </div>
                          <div>
                            <label className="tiny text-muted font-bold">Centre Partition</label>
                            <input
                              type="text"
                              className="option-input-field"
                              value={proposalSpecs.page3.centrePartition}
                              onChange={(e) => setProposalSpecs((prev) => ({ ...prev, page3: { ...prev.page3, centrePartition: e.target.value } }))}
                            />
                          </div>
                          <div>
                            <label className="tiny text-muted font-bold">End Panel</label>
                            <input
                              type="text"
                              className="option-input-field"
                              value={proposalSpecs.page3.endPanel}
                              onChange={(e) => setProposalSpecs((prev) => ({ ...prev, page3: { ...prev.page3, endPanel: e.target.value } }))}
                            />
                          </div>
                          <div>
                            <label className="tiny text-muted font-bold">Central Locking</label>
                            <textarea
                              className="option-input-field"
                              rows={2}
                              value={proposalSpecs.page3.centralLocking}
                              onChange={(e) => setProposalSpecs((prev) => ({ ...prev, page3: { ...prev.page3, centralLocking: e.target.value } }))}
                            />
                          </div>
                        </div>
                      </div>

                      <div className="option-group-card">
                        <span className="tiny font-bold" style={{ display: 'block', marginBottom: '6px', color: '#1e293b' }}>
                          Mechanical &amp; Drive Assembly
                        </span>
                        <div style={{ display: 'flex', flexDirection: 'column', gap: '8px' }}>
                          <div>
                            <label className="tiny text-muted font-bold">Powder Coating (Microns)</label>
                            <input
                              type="text"
                              className="option-input-field"
                              value={proposalSpecs.page3.powderCoating}
                              onChange={(e) => setProposalSpecs((prev) => ({ ...prev, page3: { ...prev.page3, powderCoating: e.target.value } }))}
                            />
                          </div>
                          <div>
                            <label className="tiny text-muted font-bold">Index Arrangement</label>
                            <input
                              type="text"
                              className="option-input-field"
                              value={proposalSpecs.page3.indexArrangement}
                              onChange={(e) => setProposalSpecs((prev) => ({ ...prev, page3: { ...prev.page3, indexArrangement: e.target.value } }))}
                            />
                          </div>
                          <div>
                            <label className="tiny text-muted font-bold">Shaft</label>
                            <input
                              type="text"
                              className="option-input-field"
                              value={proposalSpecs.page3.shaft}
                              onChange={(e) => setProposalSpecs((prev) => ({ ...prev, page3: { ...prev.page3, shaft: e.target.value } }))}
                            />
                          </div>
                          <div>
                            <label className="tiny text-muted font-bold">Pedestal Bearing</label>
                            <input
                              type="text"
                              className="option-input-field"
                              value={proposalSpecs.page3.pedestalBearing}
                              onChange={(e) => setProposalSpecs((prev) => ({ ...prev, page3: { ...prev.page3, pedestalBearing: e.target.value } }))}
                            />
                          </div>
                          <div>
                            <label className="tiny text-muted font-bold">Wheel</label>
                            <input
                              type="text"
                              className="option-input-field"
                              value={proposalSpecs.page3.wheel}
                              onChange={(e) => setProposalSpecs((prev) => ({ ...prev, page3: { ...prev.page3, wheel: e.target.value } }))}
                            />
                          </div>
                          <div>
                            <label className="tiny text-muted font-bold">Chain</label>
                            <input
                              type="text"
                              className="option-input-field"
                              value={proposalSpecs.page3.chain}
                              onChange={(e) => setProposalSpecs((prev) => ({ ...prev, page3: { ...prev.page3, chain: e.target.value } }))}
                            />
                          </div>
                          <div>
                            <label className="tiny text-muted font-bold">Openable Shutter</label>
                            <textarea
                              className="option-input-field"
                              rows={2}
                              value={proposalSpecs.page3.openableShutter}
                              onChange={(e) => setProposalSpecs((prev) => ({ ...prev, page3: { ...prev.page3, openableShutter: e.target.value } }))}
                            />
                          </div>
                        </div>
                      </div>
                    </div>
                  )}

                  {/* PAGE 4 EDITOR */}
                  {specsSubTab === 'page4' && (
                    <div style={{ display: 'flex', flexDirection: 'column', gap: '10px' }}>
                      <div className="option-group-card">
                        <span className="tiny font-bold" style={{ display: 'block', marginBottom: '6px', color: '#1e293b' }}>
                          Page 4: Powder Coating Plant &amp; Machinery
                        </span>
                        <div style={{ display: 'flex', flexDirection: 'column', gap: '8px' }}>
                          <div>
                            <label className="tiny text-muted font-bold">Main Title</label>
                            <input
                              type="text"
                              className="option-input-field"
                              value={proposalSpecs.page4.mainTitle}
                              onChange={(e) => setProposalSpecs((prev) => ({ ...prev, page4: { ...prev.page4, mainTitle: e.target.value } }))}
                            />
                          </div>
                          <div>
                            <label className="tiny text-muted font-bold">Plant Overview &amp; Ovens</label>
                            <textarea
                              className="option-input-field"
                              rows={3}
                              value={proposalSpecs.page4.plantOverview}
                              onChange={(e) => setProposalSpecs((prev) => ({ ...prev, page4: { ...prev.page4, plantOverview: e.target.value } }))}
                            />
                          </div>
                          <div>
                            <label className="tiny text-muted font-bold">Pretreatment Plant Spec</label>
                            <input
                              type="text"
                              className="option-input-field"
                              value={proposalSpecs.page4.pretreatmentPlant}
                              onChange={(e) => setProposalSpecs((prev) => ({ ...prev, page4: { ...prev.page4, pretreatmentPlant: e.target.value } }))}
                            />
                          </div>
                          <div>
                            <label className="tiny text-muted font-bold">Pretreatment Tanks (8 Tanks)</label>
                            <div style={{ display: 'flex', flexDirection: 'column', gap: '4px', marginTop: '4px' }}>
                              {proposalSpecs.page4.tankList.map((tank, tIdx) => (
                                <input
                                  key={tIdx}
                                  type="text"
                                  className="option-input-field"
                                  style={{ fontSize: '11px', padding: '3px 6px' }}
                                  value={tank}
                                  onChange={(e) => {
                                    const val = e.target.value
                                    setProposalSpecs((prev) => {
                                      const next = [...prev.page4.tankList]
                                      next[tIdx] = val
                                      return { ...prev, page4: { ...prev.page4, tankList: next } }
                                    })
                                  }}
                                />
                              ))}
                            </div>
                          </div>
                          <div>
                            <label className="tiny text-muted font-bold">Compressors A</label>
                            <input
                              type="text"
                              className="option-input-field"
                              value={proposalSpecs.page4.compressorsA}
                              onChange={(e) => setProposalSpecs((prev) => ({ ...prev, page4: { ...prev.page4, compressorsA: e.target.value } }))}
                            />
                          </div>
                          <div>
                            <label className="tiny text-muted font-bold">Compressors B</label>
                            <input
                              type="text"
                              className="option-input-field"
                              value={proposalSpecs.page4.compressorsB}
                              onChange={(e) => setProposalSpecs((prev) => ({ ...prev, page4: { ...prev.page4, compressorsB: e.target.value } }))}
                            />
                          </div>
                          <div>
                            <label className="tiny text-muted font-bold">Powder Coating System</label>
                            <input
                              type="text"
                              className="option-input-field"
                              value={proposalSpecs.page4.powderCoating}
                              onChange={(e) => setProposalSpecs((prev) => ({ ...prev, page4: { ...prev.page4, powderCoating: e.target.value } }))}
                            />
                          </div>
                          <div>
                            <label className="tiny text-muted font-bold">Oven Dimensions</label>
                            <input
                              type="text"
                              className="option-input-field"
                              value={proposalSpecs.page4.oven}
                              onChange={(e) => setProposalSpecs((prev) => ({ ...prev, page4: { ...prev.page4, oven: e.target.value } }))}
                            />
                          </div>
                          <div>
                            <label className="tiny text-muted font-bold">Misc Equipments</label>
                            <input
                              type="text"
                              className="option-input-field"
                              value={proposalSpecs.page4.miscEquipments}
                              onChange={(e) => setProposalSpecs((prev) => ({ ...prev, page4: { ...prev.page4, miscEquipments: e.target.value } }))}
                            />
                          </div>
                        </div>
                      </div>

                      <div className="option-group-card">
                        <span className="tiny font-bold" style={{ display: 'block', marginBottom: '6px', color: '#1e293b' }}>
                          Methods of Instruction &amp; Parameters
                        </span>
                        <div style={{ display: 'flex', flexDirection: 'column', gap: '8px' }}>
                          <div>
                            <label className="tiny text-muted font-bold">Pretreatment Intro (IS 3618)</label>
                            <textarea
                              className="option-input-field"
                              rows={3}
                              value={proposalSpecs.page4.pretreatmentIntro}
                              onChange={(e) => setProposalSpecs((prev) => ({ ...prev, page4: { ...prev.page4, pretreatmentIntro: e.target.value } }))}
                            />
                          </div>
                          <div>
                            <label className="tiny text-muted font-bold">Degreasing Parameters</label>
                            <div style={{ display: 'flex', flexDirection: 'column', gap: '4px', marginTop: '4px' }}>
                              {proposalSpecs.page4.degreasingLines.map((line, lIdx) => (
                                <input
                                  key={lIdx}
                                  type="text"
                                  className="option-input-field"
                                  style={{ fontSize: '11px', padding: '3px 6px' }}
                                  value={line}
                                  onChange={(e) => {
                                    const val = e.target.value
                                    setProposalSpecs((prev) => {
                                      const next = [...prev.page4.degreasingLines]
                                      next[lIdx] = val
                                      return { ...prev, page4: { ...prev.page4, degreasingLines: next } }
                                    })
                                  }}
                                />
                              ))}
                            </div>
                          </div>
                          <div>
                            <label className="tiny text-muted font-bold">Water Rinse 1</label>
                            <input
                              type="text"
                              className="option-input-field"
                              value={proposalSpecs.page4.waterRinse1}
                              onChange={(e) => setProposalSpecs((prev) => ({ ...prev, page4: { ...prev.page4, waterRinse1: e.target.value } }))}
                            />
                          </div>
                          <div>
                            <label className="tiny text-muted font-bold">Derusting Parameters</label>
                            <div style={{ display: 'flex', flexDirection: 'column', gap: '4px', marginTop: '4px' }}>
                              {proposalSpecs.page4.derustingLines.map((line, lIdx) => (
                                <input
                                  key={lIdx}
                                  type="text"
                                  className="option-input-field"
                                  style={{ fontSize: '11px', padding: '3px 6px' }}
                                  value={line}
                                  onChange={(e) => {
                                    const val = e.target.value
                                    setProposalSpecs((prev) => {
                                      const next = [...prev.page4.derustingLines]
                                      next[lIdx] = val
                                      return { ...prev, page4: { ...prev.page4, derustingLines: next } }
                                    })
                                  }}
                                />
                              ))}
                            </div>
                          </div>
                          <div>
                            <label className="tiny text-muted font-bold">Water Rinse 2</label>
                            <input
                              type="text"
                              className="option-input-field"
                              value={proposalSpecs.page4.waterRinse2}
                              onChange={(e) => setProposalSpecs((prev) => ({ ...prev, page4: { ...prev.page4, waterRinse2: e.target.value } }))}
                            />
                          </div>
                        </div>
                      </div>
                    </div>
                  )}

                  {/* PAGE 5 EDITOR */}
                  {specsSubTab === 'page5' && (
                    <div style={{ display: 'flex', flexDirection: 'column', gap: '10px' }}>
                      <div className="option-group-card">
                        <span className="tiny font-bold" style={{ display: 'block', marginBottom: '6px', color: '#1e293b' }}>
                          Pretreatment Parameters &amp; Passivation
                        </span>
                        <div style={{ display: 'flex', flexDirection: 'column', gap: '8px' }}>
                          <div>
                            <label className="tiny text-muted font-bold">Activation Parameters</label>
                            <div style={{ display: 'flex', flexDirection: 'column', gap: '4px', marginTop: '4px' }}>
                              {proposalSpecs.page5.activationLines.map((line, lIdx) => (
                                <input
                                  key={lIdx}
                                  type="text"
                                  className="option-input-field"
                                  style={{ fontSize: '11px', padding: '3px 6px' }}
                                  value={line}
                                  onChange={(e) => {
                                    const val = e.target.value
                                    setProposalSpecs((prev) => {
                                      const next = [...prev.page5.activationLines]
                                      next[lIdx] = val
                                      return { ...prev, page5: { ...prev.page5, activationLines: next } }
                                    })
                                  }}
                                />
                              ))}
                            </div>
                          </div>
                          <div>
                            <label className="tiny text-muted font-bold">Zinc Phosphating Parameters</label>
                            <div style={{ display: 'flex', flexDirection: 'column', gap: '4px', marginTop: '4px' }}>
                              {proposalSpecs.page5.zincPhosphatingLines.map((line, lIdx) => (
                                <input
                                  key={lIdx}
                                  type="text"
                                  className="option-input-field"
                                  style={{ fontSize: '11px', padding: '3px 6px' }}
                                  value={line}
                                  onChange={(e) => {
                                    const val = e.target.value
                                    setProposalSpecs((prev) => {
                                      const next = [...prev.page5.zincPhosphatingLines]
                                      next[lIdx] = val
                                      return { ...prev, page5: { ...prev.page5, zincPhosphatingLines: next } }
                                    })
                                  }}
                                />
                              ))}
                            </div>
                          </div>
                          <div>
                            <label className="tiny text-muted font-bold">Water Rinse 3</label>
                            <input
                              type="text"
                              className="option-input-field"
                              value={proposalSpecs.page5.waterRinse3}
                              onChange={(e) => setProposalSpecs((prev) => ({ ...prev, page5: { ...prev.page5, waterRinse3: e.target.value } }))}
                            />
                          </div>
                          <div>
                            <label className="tiny text-muted font-bold">Passivation Parameters</label>
                            <div style={{ display: 'flex', flexDirection: 'column', gap: '4px', marginTop: '4px' }}>
                              {proposalSpecs.page5.passivationLines.map((line, lIdx) => (
                                <input
                                  key={lIdx}
                                  type="text"
                                  className="option-input-field"
                                  style={{ fontSize: '11px', padding: '3px 6px' }}
                                  value={line}
                                  onChange={(e) => {
                                    const val = e.target.value
                                    setProposalSpecs((prev) => {
                                      const next = [...prev.page5.passivationLines]
                                      next[lIdx] = val
                                      return { ...prev, page5: { ...prev.page5, passivationLines: next } }
                                    })
                                  }}
                                />
                              ))}
                            </div>
                          </div>
                          <div>
                            <label className="tiny text-muted font-bold">Line Control Parameters (6 Points)</label>
                            <div style={{ display: 'flex', flexDirection: 'column', gap: '4px', marginTop: '4px' }}>
                              {proposalSpecs.page5.parametersList.map((param, pIdx) => (
                                <input
                                  key={pIdx}
                                  type="text"
                                  className="option-input-field"
                                  style={{ fontSize: '11px', padding: '3px 6px' }}
                                  value={param}
                                  onChange={(e) => {
                                    const val = e.target.value
                                    setProposalSpecs((prev) => {
                                      const next = [...prev.page5.parametersList]
                                      next[pIdx] = val
                                      return { ...prev, page5: { ...prev.page5, parametersList: next } }
                                    })
                                  }}
                                />
                              ))}
                            </div>
                          </div>
                        </div>
                      </div>

                      <div className="option-group-card">
                        <span className="tiny font-bold" style={{ display: 'block', marginBottom: '6px', color: '#1e293b' }}>
                          Quality Control Procedures
                        </span>
                        <div style={{ display: 'flex', flexDirection: 'column', gap: '8px' }}>
                          <div>
                            <label className="tiny text-muted font-bold">I. On Arrival of Material</label>
                            <div style={{ display: 'flex', flexDirection: 'column', gap: '4px', marginTop: '4px' }}>
                              {proposalSpecs.page5.qcArrivalLines.map((line, lIdx) => (
                                <textarea
                                  key={lIdx}
                                  className="option-input-field"
                                  rows={2}
                                  value={line}
                                  onChange={(e) => {
                                    const val = e.target.value
                                    setProposalSpecs((prev) => {
                                      const next = [...prev.page5.qcArrivalLines]
                                      next[lIdx] = val
                                      return { ...prev, page5: { ...prev.page5, qcArrivalLines: next } }
                                    })
                                  }}
                                />
                              ))}
                            </div>
                          </div>
                          <div>
                            <label className="tiny text-muted font-bold">II. During Production</label>
                            <div style={{ display: 'flex', flexDirection: 'column', gap: '4px', marginTop: '4px' }}>
                              {proposalSpecs.page5.qcProductionLines.map((line, lIdx) => (
                                <textarea
                                  key={lIdx}
                                  className="option-input-field"
                                  rows={2}
                                  value={line}
                                  onChange={(e) => {
                                    const val = e.target.value
                                    setProposalSpecs((prev) => {
                                      const next = [...prev.page5.qcProductionLines]
                                      next[lIdx] = val
                                      return { ...prev, page5: { ...prev.page5, qcProductionLines: next } }
                                    })
                                  }}
                                />
                              ))}
                            </div>
                          </div>
                          <div>
                            <label className="tiny text-muted font-bold">III. Final Control</label>
                            <div style={{ display: 'flex', flexDirection: 'column', gap: '4px', marginTop: '4px' }}>
                              {proposalSpecs.page5.qcFinalLines.map((line, lIdx) => (
                                <textarea
                                  key={lIdx}
                                  className="option-input-field"
                                  rows={2}
                                  value={line}
                                  onChange={(e) => {
                                    const val = e.target.value
                                    setProposalSpecs((prev) => {
                                      const next = [...prev.page5.qcFinalLines]
                                      next[lIdx] = val
                                      return { ...prev, page5: { ...prev.page5, qcFinalLines: next } }
                                    })
                                  }}
                                />
                              ))}
                            </div>
                          </div>
                          <div>
                            <label className="tiny text-muted font-bold">IV. Testing Equipments</label>
                            <div style={{ display: 'flex', flexDirection: 'column', gap: '4px', marginTop: '4px' }}>
                              {proposalSpecs.page5.qcTestingLines.map((line, lIdx) => (
                                <textarea
                                  key={lIdx}
                                  className="option-input-field"
                                  rows={2}
                                  value={line}
                                  onChange={(e) => {
                                    const val = e.target.value
                                    setProposalSpecs((prev) => {
                                      const next = [...prev.page5.qcTestingLines]
                                      next[lIdx] = val
                                      return { ...prev, page5: { ...prev.page5, qcTestingLines: next } }
                                    })
                                  }}
                                />
                              ))}
                            </div>
                          </div>
                        </div>
                      </div>
                    </div>
                  )}
                </div>
              )}
            </div>

            {/* 3. Footer */}
            <div className="sidebar-footer">
              <button
                type="button"
                className="btn primary"
                style={{ width: '100%', justifyContent: 'center', fontSize: '13px', padding: '9px 12px' }}
                onClick={handleSave}
              >
                {Icon.check} Save &amp; Apply Changes
              </button>
              <button
                type="button"
                className="btn secondary sm no-print"
                style={{ width: '100%', justifyContent: 'center', marginTop: '8px' }}
                onClick={() => setMobileView('document')}
              >
                📄 Return to Document View
              </button>
            </div>
          </>
        )}
      </aside>

      {/* Mobile Floating Action Bar */}
      <div className="official-mobile-bottom-bar no-print">
        <button
          type="button"
          className="btn primary sm"
          onClick={handleDirectDownload}
          disabled={isGeneratingPdf}
          style={{ background: '#2563eb', flex: 1 }}
        >
          {isGeneratingPdf ? <Spinner size={14} /> : Icon.fileText} PDF
        </button>
        <button
          type="button"
          className="btn sm"
          onClick={handleWhatsAppShare}
          disabled={isGeneratingPdf}
          style={{ background: '#16a34a', color: '#fff', fontWeight: 700, flex: 1, display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '4px' }}
        >
          <span>💬</span> WhatsApp
        </button>
        <button
          type="button"
          className="btn secondary sm"
          onClick={handlePrint}
          style={{ flex: 1 }}
        >
          🖨️ Print
        </button>
      </div>

      {/* Add New Client Modal (Available to Staff, Admin, and Owner) */}
      {showAddClientModal && (
        <div
          style={{
            position: 'fixed',
            top: 0,
            left: 0,
            right: 0,
            bottom: 0,
            background: 'rgba(15, 23, 42, 0.65)',
            backdropFilter: 'blur(3px)',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            zIndex: 100000,
          }}
          onClick={() => setShowAddClientModal(false)}
        >
          <div
            style={{
              background: '#ffffff',
              borderRadius: '12px',
              padding: '24px',
              width: '460px',
              maxWidth: '92vw',
              boxShadow: '0 20px 25px -5px rgba(0, 0, 0, 0.2), 0 10px 10px -5px rgba(0, 0, 0, 0.1)',
              display: 'flex',
              flexDirection: 'column',
              gap: '14px',
            }}
            onClick={(e) => e.stopPropagation()}
          >
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
              <h3 style={{ margin: 0, fontSize: '16px', fontWeight: 700, color: '#0f172a', display: 'flex', alignItems: 'center', gap: '6px' }}>
                <span>🏢 Add New Client to Directory</span>
              </h3>
              <button
                type="button"
                className="btn ghost xs"
                onClick={() => setShowAddClientModal(false)}
                style={{ fontSize: '16px', lineHeight: 1 }}
              >
                ✕
              </button>
            </div>

            <div style={{ display: 'flex', flexDirection: 'column', gap: '10px' }}>
              <div>
                <label className="tiny font-bold muted">Buyer / Company Name *</label>
                <input
                  type="text"
                  className="option-input-field font-bold"
                  placeholder="e.g. Client Enterprise Name"
                  value={newClientData.name}
                  onChange={(e) => setNewClientData({ ...newClientData, name: e.target.value })}
                  autoFocus
                />
              </div>
              <div>
                <label className="tiny font-bold muted">Address</label>
                <textarea
                  rows={3}
                  className="option-input-field"
                  placeholder="Full site / billing address..."
                  value={newClientData.address}
                  onChange={(e) => setNewClientData({ ...newClientData, address: e.target.value })}
                />
              </div>
              <div>
                <label className="tiny font-bold muted">GST Number</label>
                <input
                  type="text"
                  className="option-input-field font-mono"
                  placeholder="e.g. 27AABCP1234F1Z5"
                  value={newClientData.gst}
                  onChange={(e) => setNewClientData({ ...newClientData, gst: e.target.value })}
                />
              </div>
              <div>
                <label className="tiny font-bold muted">Buyer Email</label>
                <input
                  type="email"
                  className="option-input-field"
                  placeholder="e.g. purchase@company.com"
                  value={newClientData.email}
                  onChange={(e) => setNewClientData({ ...newClientData, email: e.target.value })}
                />
              </div>
              <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '8px' }}>
                <div>
                  <label className="tiny font-bold muted">Kind Attention</label>
                  <input
                    type="text"
                    className="option-input-field"
                    placeholder="e.g. Mr. Rajesh Sharma"
                    value={newClientData.kind_att}
                    onChange={(e) => setNewClientData({ ...newClientData, kind_att: e.target.value })}
                  />
                </div>
                <div>
                  <label className="tiny font-bold muted">Contact Number</label>
                  <input
                    type="text"
                    className="option-input-field"
                    placeholder="e.g. +91 9876543210"
                    value={newClientData.contact}
                    onChange={(e) => setNewClientData({ ...newClientData, contact: e.target.value })}
                  />
                </div>
              </div>
            </div>

            <div style={{ display: 'flex', justifyContent: 'flex-end', gap: '8px', marginTop: '6px' }}>
              <button
                type="button"
                className="btn ghost sm"
                onClick={() => setShowAddClientModal(false)}
              >
                Cancel
              </button>
              <button
                type="button"
                className="btn primary sm"
                onClick={handleCreateNewClient}
                style={{ background: '#0f172a', borderColor: '#0f172a' }}
              >
                Save &amp; Apply Client
              </button>
            </div>
          </div>
        </div>
      )}
    </div>,
    document.body
  )
}
