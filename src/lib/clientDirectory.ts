/**
 * Client Directory Management
 * Allows Staff, Members, Admins, and Owners to save, manage, and pick reusable client information.
 * Fully synchronized with Supabase database cloud.
 */

import { api, tokenStore } from './api'

export interface ClientContactPerson {
  id: string
  name: string
  phone: string
  email?: string
  designation?: string
}

export interface QuotationExtraField {
  id: string
  label: string
  value: string
  values?: string[]
  visible?: boolean
  visible_values?: boolean[]
}

/**
 * Returns all active visible values for an extra field row.
 */
export function getVisibleFieldValues(field: QuotationExtraField): string[] {
  if (field.visible === false) return []
  const valuesWithContent = Array.isArray(field.values) ? field.values.filter((v) => v && v.trim()) : []
  const all = valuesWithContent.length > 0
    ? valuesWithContent
    : (field.value?.trim() ? [field.value.trim()] : [])

  const visList = field.visible_values || []
  return all.filter((val, idx) => {
    if (!val || !val.trim()) return false
    return visList[idx] !== false
  })
}

export interface SavedClient {
  id: string
  name: string
  address: string
  gst?: string
  delivery?: string
  kind_att?: string
  contact?: string
  email?: string
  subject?: string
  contacts?: ClientContactPerson[]
  extra_fields?: QuotationExtraField[]
  profit_margin_pct?: number // Custom profit margin % set by Owner/Admin (hidden from staff)
  shared_with?: string[] // Staff usernames or ['all'] that owner has shared client with
  created_by?: string
  created_by_role?: string
  created_at?: string
  updated_at?: string
}

const STORAGE_PREFIX = 'examcalc_saved_clients'

export const DEFAULT_CLIENTS: SavedClient[] = []

// In-memory fallback for Node test environments or when storage is unavailable
const memoryStorage = new Map<string, string>()

function getStoreItem(key: string): string | null {
  if (typeof window !== 'undefined' && window.localStorage) {
    try {
      const val = localStorage.getItem(key)
      if (val !== null) return val
    } catch {
      /* ignore */
    }
  }
  return memoryStorage.get(key) || null
}

function setStoreItem(key: string, value: string): void {
  if (typeof window !== 'undefined' && window.localStorage) {
    try {
      localStorage.setItem(key, value)
    } catch {
      /* ignore */
    }
  }
  memoryStorage.set(key, value)
}

function dispatchClientsUpdated(list: SavedClient[]): void {
  if (typeof window !== 'undefined' && typeof window.dispatchEvent === 'function') {
    try {
      window.dispatchEvent(new CustomEvent('examcalc:clients-updated', { detail: { clients: list } }))
    } catch {
      /* ignore */
    }
  }
}

function getStorageKey(orgId?: string): string {
  return orgId ? `${STORAGE_PREFIX}_${orgId}` : STORAGE_PREFIX
}

/**
 * Retrieves all saved clients for the given organisation.
 * Automatically purges any legacy hardcoded demo clients.
 */
export function getSavedClients(orgId?: string): SavedClient[] {
  const key = getStorageKey(orgId)
  try {
    const raw = getStoreItem(key) || (orgId ? getStoreItem(STORAGE_PREFIX) : null)
    if (!raw) {
      setStoreItem(key, JSON.stringify([]))
      return []
    }
    const parsed = JSON.parse(raw)
    if (Array.isArray(parsed)) {
      // Purge any legacy demo clients or auto-harvested phantom clients
      const cleaned = parsed.filter(
        (c) =>
          c &&
          !c.id?.startsWith('client_harvest_') &&
          c.created_by !== 'project' &&
          c.id !== 'client_patil_group' &&
          c.id !== 'client_tata_projects' &&
          c.id !== 'client_l_and_t' &&
          c.created_by !== 'system' &&
          c.name?.trim().toLowerCase() !== 'patil group' &&
          c.name?.trim().toLowerCase() !== 'tata projects limited' &&
          c.name?.trim().toLowerCase() !== 'larsen & toubro limited' &&
          c.name?.trim().toLowerCase() !== 'unnamed client'
      )
      if (cleaned.length !== parsed.length) {
        setStoreItem(key, JSON.stringify(cleaned))
        if (orgId) setStoreItem(STORAGE_PREFIX, JSON.stringify(cleaned))
      }
      return cleaned
    }
    return []
  } catch (err) {
    console.error('Failed to read saved clients:', err)
    return []
  }
}

/**
 * Saves a new client or updates an existing client.
 * Accessible to Staff, Admin, and Owner.
 */
export function saveClientToDirectory(
  clientData: Omit<SavedClient, 'id'> & { id?: string },
  orgId?: string,
  username?: string,
  role?: string,
): SavedClient {
  const list = [...getSavedClients(orgId)]
  const key = getStorageKey(orgId)
  const now = new Date().toISOString()
  const authorRole = role || (username === 'owner' ? 'owner' : username === 'admin' ? 'admin' : 'staff')
  const normName = (clientData.name || '').trim().toLowerCase()
  const normGst = (clientData.gst || '').trim().toLowerCase()

  // Strict collision check: If client with same Name or GST already exists and is owned by another staff member
  const existingDup = list.find((c) => {
    if (clientData.id && c.id === clientData.id) return false
    const matchName = normName && c.name && c.name.trim().toLowerCase() === normName
    const matchGst = normGst && c.gst && c.gst.trim().toLowerCase() === normGst
    return Boolean(matchName || matchGst)
  })

  if (existingDup && authorRole !== 'owner' && authorRole !== 'admin') {
    const isOwnerOrCreator = existingDup.created_by === username || existingDup.created_by === 'user'
    const isShared = Array.isArray(existingDup.shared_with) && (existingDup.shared_with.includes(username || '') || existingDup.shared_with.includes('all'))
    if (!isOwnerOrCreator && !isShared) {
      throw new Error('Client already exists')
    }
  }

  let savedItem: SavedClient

  if (clientData.id) {
    // Update existing
    const idx = list.findIndex((c) => c.id === clientData.id)
    if (idx >= 0) {
      const target = list[idx]
      if (authorRole !== 'owner' && authorRole !== 'admin') {
        const isOwnerOrCreator = target.created_by === username || target.created_by === 'user'
        const isShared = Array.isArray(target.shared_with) && (target.shared_with.includes(username || '') || target.shared_with.includes('all'))
        if (!isOwnerOrCreator && !isShared) {
          throw new Error('Client already exists')
        }
      }
      savedItem = {
        ...target,
        ...clientData,
        id: clientData.id,
        updated_at: now,
      }
      list[idx] = savedItem
    } else {
      savedItem = {
        ...clientData,
        id: clientData.id,
        created_by: username || 'user',
        created_by_role: authorRole,
        created_at: now,
        updated_at: now,
      }
      list.push(savedItem)
    }
  } else {
    // Check if duplicate name already exists (case-insensitive)
    const existingIdx = list.findIndex(
      (c) => c.name.trim().toLowerCase() === normName,
    )
    if (existingIdx >= 0) {
      const target = list[existingIdx]
      if (authorRole !== 'owner' && authorRole !== 'admin') {
        const isOwnerOrCreator = target.created_by === username || target.created_by === 'user'
        const isShared = Array.isArray(target.shared_with) && (target.shared_with.includes(username || '') || target.shared_with.includes('all'))
        if (!isOwnerOrCreator && !isShared) {
          throw new Error('Client already exists')
        }
      }
      savedItem = {
        ...target,
        ...clientData,
        id: list[existingIdx].id,
        updated_at: now,
      }
      list[existingIdx] = savedItem
    } else {
      const newId = (typeof crypto !== 'undefined' && crypto.randomUUID)
        ? `client_${crypto.randomUUID()}`
        : `client_${Date.now()}_${Math.random().toString(36).slice(2, 7)}`
      savedItem = {
        ...clientData,
        id: newId,
        created_by: username || 'user',
        created_by_role: authorRole,
        created_at: now,
        updated_at: now,
      }
      list.unshift(savedItem)
    }
  }

  try {
    const serialized = JSON.stringify(list)
    setStoreItem(key, serialized)
    if (orgId) {
      setStoreItem(STORAGE_PREFIX, serialized)
    }
    dispatchClientsUpdated(list)

    // Synchronize to Supabase cloud in background
    const activeToken = tokenStore.get()
    if (activeToken && orgId) {
      api.clientUpsert(activeToken, orgId, savedItem).catch((err) => {
        console.warn('Supabase client sync error:', err)
      })
    }
  } catch (err) {
    console.error('Failed to save client to directory:', err)
  }

  return savedItem
}

/**
 * Async version of saveClientToDirectory that strictly validates against the Supabase server RPC first.
 */
export async function saveClientToDirectoryAsync(
  clientData: Omit<SavedClient, 'id'> & { id?: string },
  orgId?: string,
  username?: string,
  role?: string,
): Promise<SavedClient> {
  const activeToken = tokenStore.get()
  const authorRole = role || (username === 'owner' ? 'owner' : username === 'admin' ? 'admin' : 'staff')
  const normName = (clientData.name || '').trim().toLowerCase()
  const normGst = (clientData.gst || '').trim().toLowerCase()

  // 1. Local pre-validation
  const list = [...getSavedClients(orgId)]
  const existingDup = list.find((c) => {
    if (clientData.id && c.id === clientData.id) return false
    const matchName = normName && c.name && c.name.trim().toLowerCase() === normName
    const matchGst = normGst && c.gst && c.gst.trim().toLowerCase() === normGst
    return Boolean(matchName || matchGst)
  })

  if (existingDup && authorRole !== 'owner' && authorRole !== 'admin') {
    const isOwnerOrCreator = existingDup.created_by === username || existingDup.created_by === 'user'
    const isShared = Array.isArray(existingDup.shared_with) && (existingDup.shared_with.includes(username || '') || existingDup.shared_with.includes('all'))
    if (!isOwnerOrCreator && !isShared) {
      throw new Error('Client already exists')
    }
  }

  // Reuse existing ID if client with same name or GST already exists to prevent duplicates
  const clientDataToSave = {
    ...clientData,
    id: clientData.id || existingDup?.id,
  }

  // 2. Server-side validation via RPC if online
  if (activeToken && orgId) {
    try {
      const serverRes = await api.clientUpsert(activeToken, orgId, clientDataToSave)
      if (serverRes && typeof serverRes === 'object' && serverRes.id) {
        clientDataToSave.id = serverRes.id
      }
    } catch (err: any) {
      if (/Client already exists/i.test(err?.message || '')) {
        throw new Error('Client already exists')
      }
      throw err
    }
  }

  // 3. Save to local storage
  return saveClientToDirectory(clientDataToSave, orgId, username, role)
}

/**
 * Fetches latest clients from Supabase cloud database and updates local cache.
 */
export async function fetchClientsFromSupabase(token?: string, orgId?: string): Promise<SavedClient[]> {
  const activeToken = token || tokenStore.get()
  if (!activeToken || !orgId) return getSavedClients(orgId)

  try {
    const list = await api.clientList(activeToken, orgId)
    if (Array.isArray(list)) {
      const cleaned = list.filter(
        (c) =>
          c &&
          !c.id?.startsWith('client_harvest_') &&
          c.created_by !== 'project' &&
          c.id !== 'client_patil_group' &&
          c.id !== 'client_tata_projects' &&
          c.id !== 'client_l_and_t' &&
          c.created_by !== 'system' &&
          c.name?.trim().toLowerCase() !== 'patil group' &&
          c.name?.trim().toLowerCase() !== 'tata projects limited' &&
          c.name?.trim().toLowerCase() !== 'larsen & toubro limited' &&
          c.name?.trim().toLowerCase() !== 'unnamed client'
      )
      const key = getStorageKey(orgId)
      const serialized = JSON.stringify(cleaned)
      setStoreItem(key, serialized)
      if (orgId) setStoreItem(STORAGE_PREFIX, serialized)
      dispatchClientsUpdated(cleaned)
      return cleaned
    }
  } catch (err) {
    console.warn('Failed to fetch clients from Supabase:', err)
  }
  return getSavedClients(orgId)
}

/**
 * Determines whether a user with given username, role, and userId can view a client.
 * Rules:
 * - Owner and Admins can view ALL clients.
 * - Staff members can view:
 *   1) Clients created by themselves (c.created_by === username or userId).
 *   2) Clients shared with them by owner/admin (c.shared_with includes username, userId, or 'all').
 *   Clients not created by or shared with that staff member are strictly hidden.
 */
export function canUserViewClient(client: SavedClient, username?: string, role?: string, userId?: string): boolean {
  const normRole = (role || '').toLowerCase()
  const isOwnerOrAdmin = normRole === 'owner' || normRole === 'admin' || username === 'owner' || username === 'admin'
  if (isOwnerOrAdmin) return true

  if (!username && !userId) return false
  const userClean = (username || '').toLowerCase()
  const idClean = (userId || '').toLowerCase()

  // 1. Staff can view if they created the client
  const createdBy = (client.created_by || '').toLowerCase()
  if (userClean && createdBy === userClean) return true
  if (idClean && createdBy === idClean) return true

  // 2. Staff can view if the owner/admin shared access with them
  if (client.shared_with && Array.isArray(client.shared_with)) {
    return client.shared_with.some((u) => {
      const uClean = (u || '').trim().toLowerCase()
      return (
        uClean === 'all' ||
        uClean === '*' ||
        (userClean && uClean === userClean) ||
        (idClean && uClean === idClean)
      )
    })
  }

  return false
}

/**
 * Returns the client's configured profit margin percentage if set by owner.
 * Fallback to default 25%.
 */
export function getClientProfitMargin(client?: SavedClient | null): number {
  if (client && client.profit_margin_pct !== undefined && typeof client.profit_margin_pct === 'number' && !isNaN(client.profit_margin_pct)) {
    return client.profit_margin_pct
  }
  return 25
}

/**
 * Filters the list of clients according to user role & creator isolation.
 */
export function getVisibleClientsForUser(clients: SavedClient[], username?: string, role?: string, userId?: string): SavedClient[] {
  return (clients || []).filter((c) => canUserViewClient(c, username, role, userId))
}

/**
 * Deletes a client from the directory.
 */
export function deleteClientFromDirectory(clientId: string, orgId?: string): boolean {
  const list = getSavedClients(orgId) || []
  const key = getStorageKey(orgId)
  const next = list.filter((c) => c.id !== clientId)

  try {
    const serialized = JSON.stringify(next)
    setStoreItem(key, serialized)
    if (orgId) {
      setStoreItem(STORAGE_PREFIX, serialized)
    }
    dispatchClientsUpdated(next)

    // Synchronize deletion to Supabase cloud in background
    const activeToken = tokenStore.get()
    if (activeToken && orgId) {
      api.clientDelete(activeToken, orgId, clientId).catch((err) => {
        console.warn('Supabase client delete error:', err)
      })
    }

    return true
  } catch (err) {
    console.error('Failed to delete client:', err)
    return false
  }
}

/**
 * Harvests clients from existing projects so historical clients are automatically saved in the directory.
 */
export function harvestClientsFromProjects(_projects: any[], _orgId?: string): number {
  return 0
}
