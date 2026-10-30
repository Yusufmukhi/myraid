/**
 * Business Profile & Branding Management
 * Allows Owners and Admins to manage and update their organisation's business details,
 * addresses, contact numbers, GST/PAN, emails, website, logo, and PDF quotation header branding.
 */

import { MYRIAD_LOGO_DATA_URL, MYRIAD_WATERMARK_DATA_URL } from '../assets/companyAssets'
import { api, tokenStore } from './api'

export interface BusinessProfile {
  company_name: string
  tagline?: string
  registered_address_line1: string
  registered_address_line2: string
  contact_phone: string
  alternate_phone?: string
  gst_no: string
  pan_no?: string
  info_email: string
  sales_email?: string
  website: string
  factory_address_line1: string
  factory_address_line2: string
  signatory_title?: string
  logo_url?: string
  watermark_url?: string
  show_watermark?: boolean
  updated_at?: string
  updated_by?: string
}

export const DEFAULT_BUSINESS_PROFILE: BusinessProfile = {
  company_name: 'MYRIAD STORAGE SYSTEM LLP',
  tagline: 'Manufacturer of Mobile Compactor & Storage Systems',
  registered_address_line1: '702/A, AL-HUSSAIN, MOMIN NAGAR, JOGESHWARI (W)',
  registered_address_line2: 'MUMBAI - 400102. State : Maharashtra, Code : 27',
  contact_phone: '9769377043',
  alternate_phone: '',
  gst_no: '27ABJFM5468N1ZG',
  pan_no: 'ABJFM5468N',
  info_email: 'info@myriadstoragesystem.com',
  sales_email: 'sales@myriadstoragesystem.com',
  website: 'www.myriadstoragesystem.com',
  factory_address_line1: 'Factory Address : Gala No 5, Vakan Compound, Near Dynamic Co, Vakan Pada Raod, Vasai Phata, Vasai East Dis.,',
  factory_address_line2: 'Palghar 401208.',
  signatory_title: 'Authorized Signatory',
  logo_url: MYRIAD_LOGO_DATA_URL,
  watermark_url: MYRIAD_WATERMARK_DATA_URL,
  show_watermark: true,
}

const STORAGE_PREFIX = 'examcalc_business_profile'

// In-memory fallback for Node/tests
const memoryStore = new Map<string, string>()

function getStoreItem(key: string): string | null {
  if (typeof window !== 'undefined' && window.localStorage) {
    try {
      const val = localStorage.getItem(key)
      if (val !== null) return val
    } catch {
      /* ignore */
    }
  }
  return memoryStore.get(key) || null
}

function setStoreItem(key: string, value: string): void {
  if (typeof window !== 'undefined' && window.localStorage) {
    try {
      localStorage.setItem(key, value)
    } catch {
      /* ignore */
    }
  }
  memoryStore.set(key, value)
}

function removeStoreItem(key: string): void {
  if (typeof window !== 'undefined' && window.localStorage) {
    try {
      localStorage.removeItem(key)
    } catch {
      /* ignore */
    }
  }
  memoryStore.delete(key)
}

function dispatchProfileUpdated(profile: BusinessProfile): void {
  if (typeof window !== 'undefined' && typeof window.dispatchEvent === 'function') {
    try {
      window.dispatchEvent(new CustomEvent('examcalc:business-profile-updated', { detail: { profile } }))
    } catch {
      /* ignore */
    }
  }
}

function getStorageKey(orgId?: string): string {
  return orgId ? `${STORAGE_PREFIX}_${orgId}` : STORAGE_PREFIX
}

/**
 * Retrieves the current business profile for the specified organisation.
 * Falls back to default Myriad template details if not customized.
 */
export function getBusinessProfile(orgId?: string): BusinessProfile {
  const key = getStorageKey(orgId)
  try {
    const raw = getStoreItem(key) || (!orgId ? getStoreItem(STORAGE_PREFIX) : null)
    if (!raw) {
      return { ...DEFAULT_BUSINESS_PROFILE }
    }
    const parsed = JSON.parse(raw)
    return {
      ...DEFAULT_BUSINESS_PROFILE,
      ...parsed,
    }
  } catch (err) {
    console.error('Failed to read business profile:', err)
    return { ...DEFAULT_BUSINESS_PROFILE }
  }
}

/**
 * Saves or updates business profile fields.
 * Only owners and admins can call this from the UI.
 */
export function saveBusinessProfile(
  patch: Partial<BusinessProfile>,
  orgId?: string,
  updatedBy?: string
): BusinessProfile {
  const current = getBusinessProfile(orgId)
  const updated: BusinessProfile = {
    ...current,
    ...patch,
    updated_at: new Date().toISOString(),
    updated_by: updatedBy || current.updated_by || 'Admin',
  }

  const key = getStorageKey(orgId)
  setStoreItem(key, JSON.stringify(updated))

  dispatchProfileUpdated(updated)

  // Synchronize to Supabase cloud in background
  const activeToken = tokenStore.get()
  if (activeToken && orgId) {
    api.businessProfileSave(activeToken, orgId, updated).catch((err) => {
      console.warn('Supabase business profile sync error:', err)
    })
  }

  return updated
}

/**
 * Resets the business profile to default Myriad values.
 */
export function resetBusinessProfile(orgId?: string, resetBy?: string): BusinessProfile {
  const key = getStorageKey(orgId)
  removeStoreItem(key)
  const resetProfile = {
    ...DEFAULT_BUSINESS_PROFILE,
    updated_at: new Date().toISOString(),
    updated_by: resetBy || 'Admin',
  }
  setStoreItem(key, JSON.stringify(resetProfile))
  dispatchProfileUpdated(resetProfile)

  // Synchronize reset to Supabase cloud in background
  const activeToken = tokenStore.get()
  if (activeToken && orgId) {
    api.businessProfileSave(activeToken, orgId, resetProfile).catch((err) => {
      console.warn('Supabase business profile reset error:', err)
    })
  }

  return resetProfile
}

/**
 * Fetches latest business profile from Supabase cloud database and updates local cache.
 */
export async function fetchBusinessProfileFromSupabase(token?: string, orgId?: string): Promise<BusinessProfile> {
  const activeToken = token || tokenStore.get()
  if (!activeToken || !orgId) return getBusinessProfile(orgId)

  try {
    const remote = await api.businessProfileGet(activeToken, orgId)
    if (remote && remote.company_name) {
      const merged: BusinessProfile = {
        ...DEFAULT_BUSINESS_PROFILE,
        ...remote,
      }
      const key = getStorageKey(orgId)
      setStoreItem(key, JSON.stringify(merged))
      dispatchProfileUpdated(merged)
      return merged
    }
  } catch (err) {
    console.warn('Failed to fetch business profile from Supabase:', err)
  }
  return getBusinessProfile(orgId)
}
