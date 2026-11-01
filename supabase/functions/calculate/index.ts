// ============================================================================
// Supabase Edge Function: calculate
// Server-Side Quotation Calculation Engine
// All calculations happen HERE — the frontend sends inputs, receives results.
// ============================================================================

import { serve } from 'https://deno.land/std@0.224.0/http/server.ts'
import { createClient } from 'https://esm.sh/@supabase/supabase-js@2.48.1'
import { Workbook, DEFAULT_NAMED_CONSTANTS } from './engine.ts'
import type { SheetData, Overrides } from './engine.ts'
import { summarize, cleanOverrides, filterCutToCutOverrides } from './model.ts'
import type { Summary } from './model.ts'
import workbookData from './workbook.json' with { type: 'json' }
import workbookC2CData from './workbook_c2c.json' with { type: 'json' }

// ------------------------------------------------------------------- CORS
const ALLOWED_ORIGIN_PATTERNS = [
  /^https?:\/\/localhost(:\d+)?$/,
  /^https?:\/\/127\.0\.0\.1(:\d+)?$/,
  /^https:\/\/.*\.vercel\.app$/,
]

function getCorsHeaders(req: Request): Record<string, string> {
  const origin = req.headers.get('Origin') || ''
  const envAllowed = Deno.env.get('ALLOWED_ORIGINS') || ''
  const explicitList = envAllowed.split(',').map((s) => s.trim()).filter(Boolean)

  let allowed = false
  if (origin) {
    for (const pat of ALLOWED_ORIGIN_PATTERNS) {
      if (pat.test(origin)) { allowed = true; break }
    }
    if (!allowed && (explicitList.includes(origin) || origin.endsWith('.vercel.app') || origin === 'https://myraid-khaki.vercel.app')) allowed = true
  }

  const headers: Record<string, string> = {
    'Access-Control-Allow-Headers': 'authorization, x-client-info, apikey, content-type, x-org-id',
    'Access-Control-Allow-Methods': 'POST, OPTIONS',
  }
  if (allowed || origin.endsWith('.vercel.app') || origin.includes('vercel.app')) {
    headers['Access-Control-Allow-Origin'] = origin
    headers['Vary'] = 'Origin'
  }
  return headers
}

function jsonHeaders(cors: Record<string, string>): Record<string, string> {
  return { ...cors, 'Content-Type': 'application/json' }
}

// --------------------------------------------------------- Sliding Rate Limiter
const userRateLimits = new Map<string, { count: number; resetAt: number }>()

function checkRateLimit(userId: string, maxRequests = 60, windowMs = 60000): boolean {
  const now = Date.now()
  const entry = userRateLimits.get(userId)
  if (!entry || now > entry.resetAt) {
    userRateLimits.set(userId, { count: 1, resetAt: now + windowMs })
    return true
  }
  if (entry.count >= maxRequests) return false
  entry.count++
  return true
}

// --------------------------------------------------------- request payload
interface SystemInput {
  name?: string
  quantity?: number
  unit?: string
  overrides?: Overrides
  formulaOverrides?: Record<string, string>
}

interface CalculateRequest {
  action?: string
  org_id?: string
  inputs?: Overrides
  formula_overrides?: Record<string, string>
  template_defaults?: Record<string, number | null>
  profit_pct?: number
  discount_pct?: number
  include_installation?: boolean
  extra_amount?: number
  systems?: SystemInput[]
  c2c_overrides?: Overrides
  c2c_formula_overrides?: Record<string, string>
}

// -------------------------------------------------------- system result
interface SystemResult {
  id: string
  name: string
  quantity: number
  unit: string
  summary: Summary
  bodies: number
  units: number
  rate: number
  amount: number
}

// ------------------------------------------------- body count calculator
function computeBodies(wb: Workbook): number {
  const n = (a: string) => wb.num('Total cal', a)
  const dfmuRows = n('A3') || 4
  const sfmuRows = n('A7') || Math.max(1, Math.round(dfmuRows / 2))
  const sffuRows = n('A11') || Math.max(1, Math.round(dfmuRows / 2))
  const dffuRows = n('A15') || dfmuRows
  const dfmudRows = n('A19') || dfmuRows

  const dfmuQty = n('G4')
  const sfmuQty = n('G8')
  const sffuQty = n('G12')
  const dffuQty = n('G16')
  const dfmudQty = n('G20')

  const bodies = (dfmuQty * dfmuRows) + (sfmuQty * sfmuRows) + (sffuQty * sffuRows) +
                 (dffuQty * dffuRows) + (dfmudQty * dfmudRows)
  return bodies > 0 ? bodies : 1
}

function computeUnits(wb: Workbook): number {
  const n = (a: string) => wb.num('Total cal', a)
  return n('G4') + n('G8') + n('G12') + n('G16') + n('G20')
}

// ---------------------------------------------------- main server handler
serve(async (req: Request) => {
  const corsHeaders = getCorsHeaders(req)

  if (req.method === 'OPTIONS') {
    return new Response('ok', { headers: corsHeaders })
  }

  try {
    const supabaseUrl = Deno.env.get('SUPABASE_URL') || ''
    const supabaseServiceKey = Deno.env.get('SUPABASE_SERVICE_ROLE_KEY') || ''

    if (!supabaseUrl || !supabaseServiceKey) {
      return new Response(JSON.stringify({ error: 'Server configuration error.' }), {
        status: 500, headers: jsonHeaders(corsHeaders),
      })
    }

    // 1. Authenticate via Bearer Token Hash
    const authHeader = req.headers.get('Authorization')
    if (!authHeader) {
      return new Response(JSON.stringify({ error: 'Authentication required.' }), {
        status: 401, headers: jsonHeaders(corsHeaders),
      })
    }

    const token = authHeader.replace(/^Bearer\s+/i, '')
    const supabase = createClient(supabaseUrl, supabaseServiceKey)

    const tokenHash = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(token))
    const hashHex = Array.from(new Uint8Array(tokenHash)).map(b => b.toString(16).padStart(2, '0')).join('')

    // Look up session strictly by token_hash
    const { data: session, error: sessErr } = await supabase
      .from('app_sessions')
      .select('user_id')
      .eq('token_hash', hashHex)
      .gt('expires_at', new Date().toISOString())
      .limit(1)
      .maybeSingle()

    if (sessErr || !session || !session.user_id) {
      return new Response(JSON.stringify({ error: 'Invalid or expired session. Please sign in again.' }), {
        status: 401, headers: jsonHeaders(corsHeaders),
      })
    }

    // Check mandatory password reset policy
    const { data: userProfile } = await supabase
      .from('app_users')
      .select('must_change_password')
      .eq('id', session.user_id)
      .single()

    if (userProfile?.must_change_password) {
      return new Response(JSON.stringify({ error: 'Please set your permanent password before calculating quotations.' }), {
        status: 403, headers: jsonHeaders(corsHeaders),
      })
    }

    // Rate Limiting per authenticated user
    if (!checkRateLimit(session.user_id, 60, 60000)) {
      return new Response(JSON.stringify({ error: 'Too many requests. Please slow down and try again shortly.' }), {
        status: 429, headers: jsonHeaders(corsHeaders),
      })
    }

    // 2. Parse request body
    const body: CalculateRequest = await req.json()
    const {
      inputs = {},
      formula_overrides = {},
      template_defaults = {},
      profit_pct = 25,
      discount_pct = 0,
      include_installation = true,
      extra_amount = 0,
      systems = [],
      c2c_overrides = {},
      c2c_formula_overrides = {},
    } = body

    // 3. Organization Context & Membership Validation
    const effectiveOrgId = body.org_id || req.headers.get('x-org-id')
    if (!effectiveOrgId) {
      return new Response(JSON.stringify({ error: 'org_id required' }), {
        status: 400, headers: jsonHeaders(corsHeaders),
      })
    }

    let userRole = 'member'
    const { data: memberRecord, error: memErr } = await supabase
      .from('org_members')
      .select('role')
      .eq('user_id', session.user_id)
      .eq('org_id', effectiveOrgId)
      .maybeSingle()

    if (memErr || !memberRecord) {
      return new Response(JSON.stringify({ error: 'Access denied: not an active member of this organisation.' }), {
        status: 403, headers: jsonHeaders(corsHeaders),
      })
    }
    userRole = memberRecord.role

    // Owner/Admin on-demand full worksheets fetch (never shipped in static client bundle)
    if (body.action === 'get_worksheets') {
      if (userRole !== 'owner' && userRole !== 'admin') {
        return new Response(JSON.stringify({ error: 'Access denied: Engineering worksheets are restricted to organisation owners and administrators.' }), {
          status: 403, headers: jsonHeaders(corsHeaders),
        })
      }
      return new Response(JSON.stringify({
        actual_workbook: workbookData,
        c2c_workbook: workbookC2CData,
        named_constants: DEFAULT_NAMED_CONSTANTS,
        metadata: {
          'Total cal': { desc: 'Summary & Quotation Model', category: 'Summary' },
          'DFMU': { desc: 'Double Face Movable Unit Matrix', category: 'Model' },
          'SFMU': { desc: 'Single Face Movable Unit Matrix', category: 'Model' },
          'SFFU': { desc: 'Single Face Fix Unit Matrix', category: 'Model' },
          'DFFU': { desc: 'Double Face Fix Unit Matrix', category: 'Model' },
          'DFMU-D': { desc: 'Double Face Movable with Door', category: 'Model' },
          'Sheet1': { desc: 'General Auxiliary Sizing', category: 'Fabrication' },
          'SHEARING SIZE': { desc: 'Sheet Metal Shearing & Blank Sizing', category: 'Fabrication' },
          'HARDWARE LIST': { desc: 'Standard Hardware BOM', category: 'Hardware' },
          'HARDWARE IN FACTORY': { desc: 'In-House Factory Hardware Assembly', category: 'Hardware' },
          'Sheet2': { desc: 'Secondary Hardware Sizing', category: 'Hardware' },
          'kg': { desc: 'Weight & Steel Consumption Sizing', category: 'Weight' },
          'packing list': { desc: 'Dispatch & Packing BOM', category: 'Dispatch' },
          'hardware': { desc: 'Complete Hardware Cost & Qty Matrix', category: 'Hardware' },
          'byfercation': { desc: 'Weight & Cost Bifurcation', category: 'Analysis' },
        },
      }), {
        status: 200, headers: jsonHeaders(corsHeaders),
      })
    }

    // 4. DoS Protection & Payload Bounds Checking
    if (Array.isArray(systems) && systems.length > 20) {
      return new Response(JSON.stringify({ error: 'Maximum 20 systems allowed per quotation.' }), {
        status: 400, headers: jsonHeaders(corsHeaders),
      })
    }

    if (inputs && Object.keys(inputs).length > 200) {
      return new Response(JSON.stringify({ error: 'Too many input fields provided.' }), {
        status: 400, headers: jsonHeaders(corsHeaders),
      })
    }

    if (formula_overrides && Object.keys(formula_overrides).length > 100) {
      return new Response(JSON.stringify({ error: 'Too many formula overrides provided.' }), {
        status: 400, headers: jsonHeaders(corsHeaders),
      })
    }

    if (c2c_overrides && Object.keys(c2c_overrides).length > 200) {
      return new Response(JSON.stringify({ error: 'Too many Cut-to-Cut overrides provided.' }), {
        status: 400, headers: jsonHeaders(corsHeaders),
      })
    }

    if (c2c_formula_overrides && Object.keys(c2c_formula_overrides).length > 100) {
      return new Response(JSON.stringify({ error: 'Too many Cut-to-Cut formula overrides provided.' }), {
        status: 400, headers: jsonHeaders(corsHeaders),
      })
    }

    // Non-admin staff cannot supply formula overrides or Cut-to-Cut overrides
    if (userRole !== 'owner' && userRole !== 'admin' && ((formula_overrides && Object.keys(formula_overrides).length > 0) || (c2c_overrides && Object.keys(c2c_overrides).length > 0))) {
      return new Response(JSON.stringify({ error: 'Staff members cannot alter calculation formulas.' }), {
        status: 403, headers: jsonHeaders(corsHeaders),
      })
    }

    // Validate inputs
    const safeProfitPct = isNaN(profit_pct) ? 25 : profit_pct
    const safeDiscountPct = isNaN(discount_pct) ? 0 : Math.min(Math.max(discount_pct, 0), 99.99)
    const safeExtraAmount = isNaN(extra_amount) ? 0 : Math.max(0, extra_amount)

    // 5. Create Workbook and calculate
    const wb = new Workbook(workbookData as unknown as { sheets: SheetData[] }, DEFAULT_NAMED_CONSTANTS)

    if (Object.keys(template_defaults).length > 0) {
      wb.setTemplateDefaults(template_defaults)
    }
    if (Object.keys(formula_overrides).length > 0) {
      wb.setFormulaOverrides(formula_overrides)
    }
    wb.setOverrides(inputs)

    // 6. Calculate primary summary
    const summary = summarize(wb, safeProfitPct, safeDiscountPct, include_installation, safeExtraAmount)
    const primaryBodies = computeBodies(wb)
    const primaryUnits = computeUnits(wb)
    const cleaned = cleanOverrides(wb, inputs)

    // 7. Handle multi-system calculation
    let systemResults: SystemResult[] = []
    let totalCombinedBodies = primaryBodies
    const activeIdx = typeof (body as any).active_system_idx === 'number' ? (body as any).active_system_idx : 0

    if (systems.length > 1) {
      // First pass: compute total bodies across all systems
      totalCombinedBodies = 0
      for (let idx = 0; idx < systems.length; idx++) {
        const sys = systems[idx]
        const sysOv: Overrides = { ...(sys.overrides || {}) }
        if (idx === activeIdx && inputs && Object.keys(inputs).length > 0) {
          Object.assign(sysOv, inputs)
        }
        // Sync shared project charges
        if (inputs['Total cal!N24'] !== undefined) sysOv['Total cal!N24'] = inputs['Total cal!N24']
        if (inputs['Total cal!N25'] !== undefined) sysOv['Total cal!N25'] = inputs['Total cal!N25']
        if (inputs['Total cal!B49'] !== undefined) sysOv['Total cal!B49'] = inputs['Total cal!B49']

        wb.setFormulaOverrides(idx === activeIdx ? formula_overrides : (sys.formulaOverrides || {}))
        wb.setOverrides(sysOv)
        const sysBodies = computeBodies(wb)
        const sysQty = Math.max(1, sys.quantity || 1)
        totalCombinedBodies += sysBodies * sysQty
      }

      if (totalCombinedBodies <= 0) totalCombinedBodies = 1

      // Second pass: compute each system's summary
      for (let idx = 0; idx < systems.length; idx++) {
        const sys = systems[idx]
        const sysOv: Overrides = { ...(sys.overrides || {}) }
        if (idx === activeIdx && inputs && Object.keys(inputs).length > 0) {
          Object.assign(sysOv, inputs)
        }
        // Sync shared project charges
        if (inputs['Total cal!N24'] !== undefined) sysOv['Total cal!N24'] = inputs['Total cal!N24']
        if (inputs['Total cal!N25'] !== undefined) sysOv['Total cal!N25'] = inputs['Total cal!N25']
        if (inputs['Total cal!B49'] !== undefined) sysOv['Total cal!B49'] = inputs['Total cal!B49']

        wb.setFormulaOverrides(idx === activeIdx ? formula_overrides : (sys.formulaOverrides || {}))
        wb.setOverrides(sysOv)

        const sysSummary = summarize(wb, safeProfitPct, safeDiscountPct, include_installation, safeExtraAmount)
        const sysBodies = computeBodies(wb)
        const sysUnits = computeUnits(wb)
        const sysQty = Math.max(1, sys.quantity || 1)

        const rate = Math.round(sysSummary.beforeGst)
        systemResults.push({
          id: String(idx + 1),
          name: sys.name || `SYSTEM ${String.fromCharCode(65 + idx)}`,
          quantity: sysQty,
          unit: sys.unit || 'SET',
          summary: sysSummary,
          bodies: sysBodies,
          units: sysUnits,
          rate,
          amount: Math.round(sysQty * rate),
        })
      }

      // Reset workbook back to active system
      wb.setFormulaOverrides(formula_overrides)
      wb.setOverrides(inputs)
    } else {
      // Single system
      systemResults = [{
        id: '1',
        name: systems[0]?.name || 'SYSTEM A',
        quantity: Math.max(1, systems[0]?.quantity || 1),
        unit: systems[0]?.unit || 'SET',
        summary,
        bodies: primaryBodies,
        units: primaryUnits,
        rate: Math.round(summary.beforeGst),
        amount: Math.round((systems[0]?.quantity || 1) * Math.round(summary.beforeGst)),
      }]
    }

    // 8. Server-Side Cut-to-Cut Calculation (Admin/Owner Confidential)
    let c2cSummary: Summary | null = null
    let c2cSystemResults: any[] = []
    let c2cAllSystemsTotals: any = null
    let c2cProfit = 0
    let c2cTotal = 0
    let c2cRatio = 0

    if (userRole === 'owner' || userRole === 'admin') {
      try {
        const c2cWb = new Workbook(workbookC2CData as unknown as { sheets: SheetData[] }, DEFAULT_NAMED_CONSTANTS)
        const c2cTplDefs = filterCutToCutOverrides(template_defaults)
        const c2cInputs = { ...filterCutToCutOverrides(inputs), ...c2c_overrides }
        const c2cFormulas = { ...formula_overrides, ...c2c_formula_overrides }

        if (Object.keys(c2cTplDefs).length > 0) {
          c2cWb.setTemplateDefaults(c2cTplDefs)
        }
        if (Object.keys(c2cFormulas).length > 0) {
          c2cWb.setFormulaOverrides(c2cFormulas)
        }
        c2cWb.setOverrides(c2cInputs)

        c2cSummary = summarize(c2cWb, safeProfitPct, safeDiscountPct, include_installation, safeExtraAmount)

        if (systems.length > 1) {
          let spTotal = 0
          let trolleyTotal = 0
          let totalWeight = 0
          let weightWithMachine = 0
          let labourSubtotal = 0
          let powderTotal = 0
          let costBase = 0
          let beforeGst = 0
          let sales = 0

          for (let idx = 0; idx < systems.length; idx++) {
            const sys = systems[idx]
            const isCurrent = idx === activeIdx
            const q = Math.max(1, sys.quantity || 1)
            let sum = c2cSummary
            if (!isCurrent) {
              const sysOv: Overrides = { ...filterCutToCutOverrides(sys.overrides || {}), ...c2c_overrides }
              if (inputs['Total cal!N24'] !== undefined) sysOv['Total cal!N24'] = inputs['Total cal!N24']
              if (inputs['Total cal!N25'] !== undefined) sysOv['Total cal!N25'] = inputs['Total cal!N25']
              if (inputs['Total cal!B49'] !== undefined) sysOv['Total cal!B49'] = inputs['Total cal!B49']
              const sysExtra = sys.extraAmount !== undefined
                ? Number(sys.extraAmount)
                : (sys.overrides?.['__extra_amount'] !== undefined ? Number(sys.overrides['__extra_amount']) : 0)
              sysOv['__extra_amount'] = sysExtra
              sysOv['__include_installation'] = include_installation ? 1 : 0
              c2cWb.setFormulaOverrides({ ...(sys.formulaOverrides || {}), ...c2c_formula_overrides })
              c2cWb.setOverrides(sysOv)
              sum = summarize(c2cWb, safeProfitPct, safeDiscountPct, include_installation, sysExtra)
            }
            spTotal += (sum.steel?.spTotal || 0) * q
            trolleyTotal += (sum.trolley?.total || 0) * q
            totalWeight += (sum.totalWeight || 0) * q
            weightWithMachine += (sum.weightWithMachine || 0) * q
            labourSubtotal += (sum.labourSubtotal || 0) * q
            powderTotal += (sum.powderTotal || 0) * q
            costBase += (sum.costBase || 0) * q
            beforeGst += (sum.beforeGst || 0) * q
            sales += (sum.sales || 0) * q

            c2cSystemResults.push({
              id: String(idx + 1),
              name: sys.name || `SYSTEM ${String.fromCharCode(65 + idx)}`,
              quantity: q,
              summary: sum,
              costBase: sum.costBase * q,
              labourSubtotal: sum.labourSubtotal * q,
              totalWeight: (sum.totalWeight || 0) * q,
            })
          }

          c2cWb.setFormulaOverrides(c2cFormulas)
          c2cWb.setOverrides(c2cInputs)

          c2cAllSystemsTotals = {
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

          const allActualBeforeGst = systemResults.reduce((acc, s) => acc + s.amount, 0)
          c2cProfit = allActualBeforeGst - c2cAllSystemsTotals.costBase
          c2cTotal = c2cProfit + c2cAllSystemsTotals.labourSubtotal
          c2cRatio = c2cAllSystemsTotals.totalWeight > 0 ? c2cTotal / c2cAllSystemsTotals.totalWeight : 0
        } else {
          const activeSysQty = Math.max(1, systems[0]?.quantity || 1)
          c2cProfit = (summary.subtotalWithProfit * activeSysQty) - (c2cSummary.costBase * activeSysQty)
          c2cTotal = c2cProfit + (c2cSummary.labourSubtotal * activeSysQty)
          const c2cWeight = (c2cSummary.totalWeight || 0) * activeSysQty
          c2cRatio = c2cWeight > 0 ? c2cTotal / c2cWeight : 0
        }
      } catch (c2cErr) {
        console.error('Server C2C calculation error:', c2cErr)
      }
    }

    // 9. Return complete verified results
    // If caller is staff (not owner or admin), remove confidential cost and profit details
    const safeSummary: any = { ...summary }
    const safeSystems: SystemResult[] = systemResults.map((sys) => ({
      ...sys,
      summary: { ...sys.summary },
    }))

    if (userRole !== 'owner' && userRole !== 'admin') {
      const stripSensitiveFields = (obj: any) => {
        if (!obj) return
        delete obj.costBase
        delete obj.profitPct
        delete obj.profitAmount
        delete obj.subtotalWithProfit
      }

      stripSensitiveFields(safeSummary)
      for (const sys of safeSystems) {
        stripSensitiveFields(sys.summary)
      }
    }

    const isAuthorizedAdmin = userRole === 'owner' || userRole === 'admin'

    const response = {
      summary: safeSummary,
      systems: safeSystems,
      total_combined_bodies: totalCombinedBodies,
      cleaned_overrides: cleaned,
      computed_total: summary.sales,
      c2c_summary: isAuthorizedAdmin ? c2cSummary : undefined,
      c2c_systems: isAuthorizedAdmin && c2cSystemResults.length > 0 ? c2cSystemResults : undefined,
      c2c_all_systems: isAuthorizedAdmin ? c2cAllSystemsTotals : undefined,
      c2c_profit: isAuthorizedAdmin ? Math.round(c2cProfit) : undefined,
      c2c_total: isAuthorizedAdmin ? Math.round(c2cTotal) : undefined,
      c2c_ratio: isAuthorizedAdmin ? Number(c2cRatio.toFixed(2)) : undefined,
    }

    return new Response(JSON.stringify(response), {
      status: 200, headers: jsonHeaders(corsHeaders),
    })

  } catch (err) {
    console.error('Edge function calculation error:', err)
    return new Response(JSON.stringify({ error: 'Calculation could not be processed. Please verify quotation parameters.' }), {
      status: 500, headers: jsonHeaders(corsHeaders),
    })
  }
})
