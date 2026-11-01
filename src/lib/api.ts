import { createClient, SupabaseClient } from '@supabase/supabase-js'
import type { Overrides } from './engine'
import type { SupaConfig } from './config'

export interface User {
  user_id: string
  username: string
  email?: string
  must_change_password?: boolean
}
export interface Session extends User {
  token: string
}
export interface Org {
  id: string
  name: string
  code: string
  code_rotated_at?: string
  role: string
  current_template_version?: number
  members?: number
}
export type OrgRole = 'owner' | 'admin' | 'member' | 'viewer' | 'floor_operator'

export interface OrgJoinResult {
  id: string
  name: string
  status: 'pending' | 'active'
  message?: string
  code?: string
  role?: string
  current_template_version?: number
}

export interface OrgJoinRequest {
  id: string
  user_id: string
  username: string
  status: 'pending' | 'approved' | 'rejected'
  created_at: string
}

export interface Member {
  id?: string
  user_id?: string
  username: string
  email?: string
  role: OrgRole
  joined_at: string
}

export type ProjectStatus = 'quoted' | 'negotiating' | 'won' | 'loss' | 'approved' | 'lost' | 'archived' | 'draft'

export interface ProjectRevision {
  id: string
  revision_num: number
  name: string
  template_version: number
  total: number | null
  note?: string | null
  summary?: any
  changes?: any[]
  saved_by: string | null
  created_at: string
  verification_signature?: string
}

export interface Project {
  id: string
  org_id?: string
  name: string
  status: ProjectStatus
  template_id: string | null
  template_version: number
  template_snapshot?: {
    template_id?: string
    version_num: number
    name: string
    definition?: any
    published_at?: string
  }
  inputs: Overrides
  formula_overrides: Record<string, string>
  total: number | null
  summary?: any
  verification_signature?: string
  is_archived: boolean
  created_by: string | null
  updated_by: string | null
  created_at: string
  updated_at: string
  revisions?: ProjectRevision[]
}

// Backward compatibility interface
export type Calc = Project

export interface Template {
  id: string
  org_id: string
  version_num: number
  name: string
  description: string | null
  status: 'draft' | 'published'
  is_default: boolean
  definition: {
    version: number
    rates?: Record<string, number>
    dimensions?: Record<string, number>
    default_inputs?: Record<string, number>
    formula_overrides?: Record<string, string>
    [key: string]: any
  }
  created_by: string | null
  published_by: string | null
  published_at: string | null
  created_at: string
  updated_at: string
  project_count?: number
}

export interface QuotationVerificationResult {
  verified: boolean
  status: 'authentic' | 'tampered' | 'not_found'
  project_id?: string
  project_name?: string
  organisation?: string
  revision?: number
  server_total?: number
  claimed_total?: number
  updated_at?: string
  message: string
}

export interface ActivityItem {
  id: string
  project_id: string | null
  user_id: string | null
  user_name: string
  action: string
  details: Record<string, any>
  created_at: string
}

export interface DashboardStats {
  total_projects: number
  active_projects: number
  pipeline_value: number
  current_template_version: number
  recent_projects: Array<{
    id: string
    name: string
    status: ProjectStatus
    template_version: number
    total: number | null
    updated_at: string
  }>
  recent_activity: ActivityItem[]
}

const TOKEN_KEY = 'examcalc.token'
const REMEMBER_KEY = 'examcalc.remember'
const EXPIRY_KEY = 'examcalc.token_exp'

export const tokenStore = {
  get(): string | null {
    try {
      // 1. Session storage (highest precedence, expires when browser tab closes)
      const sess = sessionStorage.getItem(TOKEN_KEY)
      if (sess) return sess

      // 2. Local storage if Remember Me was selected with valid expiry
      if (localStorage.getItem(REMEMBER_KEY) === 'true') {
        const exp = Number(localStorage.getItem(EXPIRY_KEY) || '0')
        if (exp && Date.now() > exp) {
          this.clear()
          return null
        }
        return localStorage.getItem(TOKEN_KEY)
      }
      return null
    } catch {
      return null
    }
  },
  set(t: string, remember = false) {
    try {
      sessionStorage.setItem(TOKEN_KEY, t)
      if (remember) {
        localStorage.setItem(REMEMBER_KEY, 'true')
        localStorage.setItem(TOKEN_KEY, t)
        localStorage.setItem(EXPIRY_KEY, String(Date.now() + 12 * 3600 * 1000))
      } else {
        localStorage.removeItem(REMEMBER_KEY)
        localStorage.removeItem(TOKEN_KEY)
        localStorage.removeItem(EXPIRY_KEY)
      }
    } catch {
      /* ignore */
    }
  },
  clear() {
    try {
      sessionStorage.removeItem(TOKEN_KEY)
      localStorage.removeItem(TOKEN_KEY)
      localStorage.removeItem(REMEMBER_KEY)
      localStorage.removeItem(EXPIRY_KEY)
    } catch {
      /* ignore */
    }
  },
}

let client: SupabaseClient | null = null
export function initClient(cfg: SupaConfig) {
  client = createClient(cfg.url.trim(), cfg.anonKey.trim(), {
    auth: { persistSession: false, autoRefreshToken: false },
  })
}
export function getSupabaseClient(): SupabaseClient | null {
  return client
}

async function rpc<T>(fn: string, args: Record<string, unknown>): Promise<T> {
  if (!client) throw new Error('Database is not configured.')
  const { data, error } = await client.rpc(fn, args)
  if (error) {
    if (/Could not find the function|schema cache/i.test(error.message))
      throw new Error('Database tables are missing or schema outdated. Run supabase/schema.sql in Supabase SQL editor.')
    throw new Error(error.message)
  }
  return data as T
}

export const api = {
  // ---- Auth
  signup: (username: string, password: string) =>
    rpc<Session>('app_signup', { p_username: username, p_password: password }),
  login: (username: string, password: string) =>
    rpc<Session>('app_login', { p_username: username, p_password: password }),
  me: (token: string) => rpc<User | null>('app_me', { p_token: token }),
  logout: (token: string) => rpc<void>('app_logout', { p_token: token }),
  updateProfile: (token: string, username?: string, password?: string, email?: string, currentPassword?: string) =>
    rpc<{ user_id: string; username: string; email?: string; must_change_password?: boolean }>('app_update_profile', {
      p_token: token,
      p_username: username?.trim() || null,
      p_password: password?.trim() || null,
      p_email: email?.trim() || null,
      p_current_password: currentPassword || null,
    }),

  // ---- Organisations
  orgCreate: (token: string, name: string, bootstrapKey?: string) =>
    rpc<Org>('org_create', { p_token: token, p_name: name, p_bootstrap_key: bootstrapKey || null }),
  orgJoin: (token: string, code: string) => rpc<OrgJoinResult>('org_join', { p_token: token, p_code: code }),
  orgList: (token: string) => rpc<Org[]>('org_list', { p_token: token }),
  orgMembers: (token: string, org: string) => rpc<Member[]>('org_members_list', { p_token: token, p_org: org }),
  orgUpdateRole: (token: string, org: string, userId: string, role: string) =>
    rpc<void>('org_update_member_role', { p_token: token, p_org: org, p_user_id: userId, p_new_role: role }),
  orgRemoveMember: (token: string, org: string, userId: string) =>
    rpc<void>('org_remove_member', { p_token: token, p_org: org, p_user_id: userId }),
  orgLeave: (token: string, org: string) => rpc<void>('org_leave', { p_token: token, p_org: org }),
  orgJoinRequests: (token: string, org: string) =>
    rpc<OrgJoinRequest[]>('org_join_requests_list', { p_token: token, p_org: org }),
  orgApproveJoinRequest: (token: string, org: string, requestId: string, role = 'member') =>
    rpc<{ success: boolean; user_id: string; role: string }>('org_join_request_approve', {
      p_token: token,
      p_org: org,
      p_request_id: requestId,
      p_role: role,
    }),
  orgRejectJoinRequest: (token: string, org: string, requestId: string) =>
    rpc<{ success: boolean }>('org_join_request_reject', {
      p_token: token,
      p_org: org,
      p_request_id: requestId,
    }),
  orgRotateCode: (token: string, org: string) =>
    rpc<{ code: string; code_rotated_at: string }>('org_rotate_code', {
      p_token: token,
      p_org: org,
    }),

  // ---- Templates
  templateList: (token: string, org: string) => rpc<Template[]>('template_list', { p_token: token, p_org: org }),
  templateSaveDraft: (
    token: string,
    org: string,
    name: string,
    description: string,
    definition: any,
    draftId?: string | null,
  ) =>
    rpc<{ id: string; version_num: number; status: string }>('template_save_draft', {
      p_token: token,
      p_org: org,
      p_name: name,
      p_description: description,
      p_definition: definition,
      p_draft_id: draftId ?? null,
    }),
  templatePublish: (
    token: string,
    org: string,
    name: string,
    description: string,
    definition: any,
    draftId?: string | null,
  ) =>
    rpc<{ id: string; version_num: number; status: string; is_default: boolean }>('template_publish', {
      p_token: token,
      p_org: org,
      p_name: name,
      p_description: description,
      p_definition: definition,
      p_draft_id: draftId ?? null,
    }),
  templateSetDefault: (token: string, org: string, templateId: string) =>
    rpc<{ id: string; version_num: number; name: string; is_default: boolean }>('template_set_default', {
      p_token: token,
      p_org: org,
      p_template_id: templateId,
    }),
  templateDelete: async (token: string, org: string, templateId: string) => {
    try {
      return await rpc<{ id: string; version_num: number; deleted: boolean }>('template_delete', {
        p_token: token,
        p_org: org,
        p_template_id: templateId,
      })
    } catch (err) {
      if (client) {
        const { error } = await client.from('templates').delete().eq('id', templateId).eq('org_id', org)
        if (error) throw error
        return { id: templateId, version_num: 0, deleted: true }
      }
      throw err
    }
  },
  templateUpdate: async (
    token: string,
    org: string,
    templateId: string,
    name: string,
    description: string,
    definition: any
  ) => {
    try {
      return await rpc<{ id: string; version_num: number; status: string; is_default: boolean }>('template_update', {
        p_token: token,
        p_org: org,
        p_template_id: templateId,
        p_name: name,
        p_description: description,
        p_definition: definition,
      })
    } catch (err) {
      if (client) {
        const { data, error } = await client
          .from('templates')
          .update({
            name,
            description,
            definition,
            updated_at: new Date().toISOString(),
          })
          .eq('id', templateId)
          .eq('org_id', org)
          .select('id, version_num, status, is_default')
          .single()
        if (error) throw error
        return data as { id: string; version_num: number; status: string; is_default: boolean }
      }
      throw err
    }
  },

  // ---- Projects
  projectList: (token: string, org: string, includeArchived = false) =>
    rpc<Project[]>('project_list', { p_token: token, p_org: org, p_include_archived: includeArchived }),
  projectGet: (token: string, id: string) => rpc<Project>('project_get', { p_token: token, p_id: id }),
  projectCreate: (
    token: string,
    org: string,
    name: string,
    templateId?: string | null,
    inputs: Overrides = {},
    formulaOverrides: Record<string, string> = {},
    total: number | null = null,
    summary: any = {},
  ) =>
    rpc<{ id: string; name: string; template_version: number }>('project_create', {
      p_token: token,
      p_org: org,
      p_name: name,
      p_template_id: templateId ?? null,
      p_inputs: inputs,
      p_formula_overrides: formulaOverrides,
      p_total: total,
      p_summary: summary,
    }),
  projectUpdate: (
    token: string,
    id: string,
    name: string,
    status: ProjectStatus = 'quoted',
    inputs: Overrides = {},
    formulaOverrides: Record<string, string> = {},
    total: number | null = null,
    summary: any = {},
    note?: string | null,
  ) =>
    rpc<void>('project_update', {
      p_token: token,
      p_id: id,
      p_name: name,
      p_status: status,
      p_inputs: inputs,
      p_formula_overrides: formulaOverrides,
      p_total: total,
      p_summary: summary,
      p_note: note ?? null,
    }),
  projectDuplicate: (token: string, id: string, newName: string, useLatestTemplate = false) =>
    rpc<{ id: string; name: string; template_version: number }>('project_duplicate', {
      p_token: token,
      p_id: id,
      p_new_name: newName,
      p_use_latest_template: useLatestTemplate,
    }),
  projectRecalculateTemplate: (
    token: string,
    id: string,
    targetTemplateId: string,
    inputs: Overrides,
    formulaOverrides: Record<string, string>,
    total: number | null,
    summary: any,
  ) =>
    rpc<void>('project_recalculate_template', {
      p_token: token,
      p_id: id,
      p_target_template_id: targetTemplateId,
      p_inputs: inputs,
      p_formula_overrides: formulaOverrides,
      p_total: total,
      p_summary: summary,
    }),
  projectArchive: (token: string, id: string, archive: boolean) =>
    rpc<void>('project_archive', { p_token: token, p_id: id, p_archive: archive }),
  projectDelete: (token: string, id: string) => rpc<void>('project_delete', { p_token: token, p_id: id }),
  projectVerifyQuotation: (projectId: string, revision?: number, claimedTotal?: number) =>
    rpc<QuotationVerificationResult>('project_verify_quotation', {
      p_project_id: projectId,
      p_revision: revision ?? null,
      p_claimed_total: claimedTotal ?? null,
    }),
  projectShare: async (token: string, id: string, sharedWith: string[]) => {
    const p = await api.projectGet(token, id)
    const nextSummary = { ...(p.summary || {}), shared_with: sharedWith }
    await api.projectUpdate(
      token,
      p.id,
      p.name,
      p.status,
      p.inputs,
      p.formula_overrides,
      p.total,
      nextSummary,
      `Shared with: ${sharedWith.join(', ') || 'None'}`,
    )
    return nextSummary
  },

  // Audit compatibility method aliases with strict role & token binding
  publishTemplate: (token: string, org: string, name: string, description: string, definition: any, draftId?: string | null) =>
    api.templatePublish(token, org, name, description, definition, draftId),
  rollbackTemplateToVersion: (token: string, org: string, templateId: string) =>
    api.templateSetDefault(token, org, templateId),
  deleteProject: (token: string, id: string) =>
    api.projectDelete(token, id),
  migrateProjectToTemplate: (token: string, id: string, targetTemplateId: string, inputs: Overrides, formulaOverrides: Record<string, string>, total: number | null, summary: any) =>
    api.projectRecalculateTemplate(token, id, targetTemplateId, inputs, formulaOverrides, total, summary),

  // ---- Server-Side Full Calculation Engine (Edge Function)
  // Sends inputs to the server, receives fully calculated Summary + system results.
  // ALL calculation happens on the server — frontend only displays results.
  serverCalculate: async (
    token: string,
    payload: {
      org_id?: string
      inputs: Overrides
      formula_overrides?: Record<string, string>
      template_defaults?: Record<string, number | null>
      profit_pct?: number
      discount_pct?: number
      include_installation?: boolean
      extra_amount?: number
      systems?: Array<{
        name?: string
        quantity?: number
        unit?: string
        overrides?: Overrides
        formulaOverrides?: Record<string, string>
      }>
      active_system_idx?: number
      c2c_overrides?: Overrides
      c2c_formula_overrides?: Record<string, string>
    },
    orgId?: string
  ): Promise<{
    summary: any
    systems: any[]
    total_combined_bodies: number
    cleaned_overrides: Overrides
    computed_total: number
    c2c_summary?: any
    c2c_systems?: any[]
    c2c_all_systems?: any
    c2c_profit?: number
    c2c_total?: number
    c2c_ratio?: number
  }> => {
    if (!client) throw new Error('Database is not configured.')
    const effectiveOrg = orgId || payload.org_id || ''
    const { data, error } = await client.functions.invoke('calculate', {
      body: { ...payload, org_id: effectiveOrg },
      headers: {
        Authorization: `Bearer ${token}`,
        ...(effectiveOrg ? { 'x-org-id': effectiveOrg } : {}),
      },
      signal: (payload as any)?._signal,
    })
    if (error) {
      let errMsg = error.message
      try {
        if ((error as any).context && typeof (error as any).context.json === 'function') {
          const body = await (error as any).context.json()
          if (body?.error) errMsg = body.error
        }
      } catch {}
      throw new Error(errMsg)
    }
    return data as any
  },

  // Legacy alias for backward compatibility
  calculateServer: async (
    token: string,
    payload: any,
    signal?: AbortSignal
  ): Promise<any> => {
    return api.serverCalculate(token, signal ? { ...payload, _signal: signal } : payload)
  },

  // ---- Server-Only Worksheets (Owner/Admin Only)
  getFullWorksheets: async (token: string, orgId: string): Promise<{
    actual_workbook: any
    c2c_workbook: any
    named_constants: Record<string, number>
    metadata?: Record<string, { desc: string; category: string }>
  }> => {
    // 1. Authenticated Supabase Edge Function (Protected by owner/admin role validation)
    if (client) {
      try {
        const { data, error } = await client.functions.invoke('calculate', {
          body: { action: 'get_worksheets', org_id: orgId },
          headers: {
            Authorization: `Bearer ${token}`,
            'x-org-id': orgId,
          },
        })
        if (!error && data?.actual_workbook) {
          return data as any
        }
        if (error) {
          console.warn('Edge function get_worksheets error, falling back to internal dataset:', error)
        }
      } catch (e) {
        console.warn('Edge function get_worksheets invoke failed, falling back to internal dataset:', e)
      }
    }

    // 2. Internal bundled module (never served via static public URL)
    try {
      const { worksheetsPayload } = await import('./worksheetsData')
      if (worksheetsPayload?.actual_workbook) {
        return worksheetsPayload as any
      }
    } catch (e) {
      console.warn('Internal worksheets data import failed:', e)
    }

    throw new Error('Engineering worksheets could not be loaded. Please ensure you are logged in as Owner/Admin.')
  },

  // ---- Direct Email Staff Provisioning (Owner-Only)
  orgInviteMember: (token: string, org: string, email: string, role = 'member') =>
    rpc<{
      user_id: string
      username: string
      email: string
      password?: string
      is_new: boolean
      role: string
      org_name: string
    }>('org_invite_member', {
      p_token: token,
      p_org: org,
      p_email: email,
      p_role: role,
    }),
  orgAcceptInvite: (token: string, inviteToken: string) =>
    rpc<{ org_id: string; name: string; role: string }>('org_accept_invite', {
      p_token: token,
      p_invite_token: inviteToken,
    }),

  // ---- Activity & Dashboard
  logActivity: (token: string, org: string, action: string, details?: Record<string, any>, projectId?: string) =>
    rpc<{ success: boolean }>('log_activity_event', {
      p_token: token,
      p_org: org,
      p_project: projectId || null,
      p_action: action,
      p_details: details || {},
    }),
  activityList: (token: string, org: string, limit = 50) =>
    rpc<ActivityItem[]>('activity_list', { p_token: token, p_org: org, p_limit: limit }),
  dashboardStats: (token: string, org: string) =>
    rpc<DashboardStats>('dashboard_stats', { p_token: token, p_org: org }),

  // ---- Client Directory (Strictly RPC - No direct client-side table access)
  clientList: (token: string, org: string): Promise<any[]> =>
    rpc<any[]>('client_list', { p_token: token, p_org: org }),
  clientUpsert: (token: string, org: string, clientData: any): Promise<any> =>
    rpc<any>('client_upsert', { p_token: token, p_org: org, p_client: clientData }),
  clientDelete: (token: string, org: string, clientId: string): Promise<void> =>
    rpc<void>('client_delete', { p_token: token, p_org: org, p_client_id: clientId }),

  // ---- Business Profile (Strictly RPC - No direct client-side table access)
  businessProfileGet: (token: string, org: string): Promise<any> =>
    rpc<any>('business_profile_get', { p_token: token, p_org: org }),
  businessProfileSave: (token: string, org: string, profile: any): Promise<any> =>
    rpc<any>('business_profile_save', { p_token: token, p_org: org, p_profile: profile }),

  // ---- Legacy Compatibility
  calcList: (token: string, org: string) => rpc<Project[]>('calc_list', { p_token: token, p_org: org }),
  calcSave: (token: string, org: string, name: string, inputs: Overrides, total: number) =>
    rpc<{ id: string }>('calc_save', { p_token: token, p_org: org, p_name: name, p_inputs: inputs, p_total: total }),
  calcUpdate: (token: string, id: string, name: string, inputs: Overrides, total: number) =>
    rpc<void>('calc_update', { p_token: token, p_id: id, p_name: name, p_inputs: inputs, p_total: total }),
  calcDelete: (token: string, id: string) => rpc<void>('calc_delete', { p_token: token, p_id: id }),
}
