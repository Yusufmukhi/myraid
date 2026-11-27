import { Fragment, useEffect, useMemo, useRef, useState } from 'react'
import { api, Org, Project, ProjectStatus, Session, tokenStore } from '../../lib/api'
import { deleteProject } from '../../lib/permissions'
import { inr, timeAgo } from '../../lib/format'
import { getSavedClients } from '../../lib/clientDirectory'
import {
  buildOfficialQtnNo,
  formatPdfDate,
  getAvailableFiscalYears,
  getFiscalYear,
  getProjectFiscalYear,
  groupProjectsIntoFolders,
  matchesFiscalYearFilter,
  parseRevision,
  ProjectFolder,
} from '../../lib/revision'
import { Icon, Spinner, useUI } from '../ui'

interface Props {
  session: Session
  org: Org
  currentProjectId: string | null
  onOpenProject: (projectId: string) => void
  onNewProject: () => void
  onProjectsChanged?: () => void
  initialScope?: 'all' | 'mine' | 'shared'
  onSessionChange?: (s: Session) => void
}

type SortField = 'updated_at' | 'name' | 'total' | 'created_at'

export function ProjectsPage({ session, org, currentProjectId, onOpenProject, onNewProject, onProjectsChanged, initialScope, onSessionChange }: Props) {
  const { toast, confirm } = useUI()
  const isAdminOrOwner = org.role === 'owner' || org.role === 'admin'
  const [projects, setProjects] = useState<Project[]>([])
  const [loading, setLoading] = useState(true)
  const [search, setSearch] = useState('')
  type SearchField = 'all' | 'qtn' | 'client' | 'address' | 'contact' | 'amount' | 'description' | 'specs'
  const [searchType, setSearchType] = useState<SearchField>('all')
  const savedClients = useMemo(() => {
    try {
      return getSavedClients(org.id)
    } catch {
      return []
    }
  }, [org.id])
  const [statusFilter, setStatusFilter] = useState<string>('all')
  const [scopeFilter, setScopeFilter] = useState<'all' | 'mine' | 'shared' | 'reminders'>(initialScope || 'all')
  const [showArchived, setShowArchived] = useState(false)
  const currentFy = useMemo(() => getFiscalYear(), [])
  const [selectedFys, setSelectedFys] = useState<string[]>(() => [getFiscalYear()])
  const [fyDropdownOpen, setFyDropdownOpen] = useState(false)
  const fyDropdownRef = useRef<HTMLDivElement>(null)

  useEffect(() => {
    if (!fyDropdownOpen) return
    const handleOutsideClick = (e: MouseEvent) => {
      if (fyDropdownRef.current && !fyDropdownRef.current.contains(e.target as Node)) {
        setFyDropdownOpen(false)
      }
    }
    document.addEventListener('mousedown', handleOutsideClick)
    return () => document.removeEventListener('mousedown', handleOutsideClick)
  }, [fyDropdownOpen])

  const { availableFys, fyCounts } = useMemo(() => {
    const counts: Record<string, number> = {}
    for (const p of projects || []) {
      const fy = getProjectFiscalYear(p)
      counts[fy] = (counts[fy] || 0) + 1
    }
    const list = getAvailableFiscalYears(projects || [], currentFy)
    return { availableFys: list, fyCounts: counts }
  }, [projects, currentFy])

  const toggleAllFys = () => {
    setSelectedFys((prev) => (prev.includes('ALL') ? [currentFy] : ['ALL']))
  }

  const toggleSingleFy = (fy: string) => {
    setSelectedFys((prev) => {
      if (prev.includes('ALL')) {
        return [fy]
      }
      if (prev.includes(fy)) {
        const next = prev.filter((item) => item !== fy)
        return next.length > 0 ? next : ['ALL']
      }
      return [...prev, fy]
    })
  }

  const fyButtonLabel = useMemo(() => {
    if (selectedFys.includes('ALL')) return 'All FY'
    if (selectedFys.length === 1) {
      const only = selectedFys[0]
      return only === currentFy ? `FY ${only} (Current)` : `FY ${only}`
    }
    return `FY (${selectedFys.length} selected)`
  }, [selectedFys, currentFy])
  const [sortBy, setSortBy] = useState<SortField>('updated_at')
  const [sortAsc, setSortAsc] = useState(false)
  const [viewMode, setViewMode] = useState<'folders' | 'flat' | 'pipeline'>('folders')
  // By default, all folders are closed (collapsed) until clicked
  const [openFolders, setOpenFolders] = useState<Record<string, boolean>>({})

  useEffect(() => {
    if (initialScope) {
      setScopeFilter(initialScope)
    }
  }, [initialScope])

  const toggleFolder = (folderId: string) => {
    setOpenFolders((prev) => ({
      ...prev,
      [folderId]: !prev[folderId],
    }))
  }

  const expandAllFolders = () => {
    const map: Record<string, boolean> = {}
    for (const f of projectFolders) {
      map[f.id] = true
    }
    setOpenFolders(map)
  }

  const collapseAllFolders = () => {
    setOpenFolders({})
  }

  // Modals state
  const [reviseProj, setReviseProj] = useState<Project | null>(null)
  const [revName, setRevName] = useState('')
  const [renameProj, setRenameProj] = useState<Project | null>(null)
  const [newName, setNewName] = useState('')
  const [duplicateProj, setDuplicateProj] = useState<Project | null>(null)
  const [dupName, setDupName] = useState('')
  const [dupUseLatest, setDupUseLatest] = useState(false)
  const [shareProj, setShareProj] = useState<Project | null>(null)
  const [shareMembers, setShareMembers] = useState<any[]>([])
  const [loadingMembers, setLoadingMembers] = useState(false)

  // Batch selection and deletion state
  const [selectedIds, setSelectedIds] = useState<Set<string>>(new Set())
  const [singleDeleteProj, setSingleDeleteProj] = useState<Project | null>(null)
  const [confirmNameInput, setConfirmNameInput] = useState('')
  const [deletingSingle, setDeletingSingle] = useState(false)

  const [showBatchDeleteModal, setShowBatchDeleteModal] = useState(false)
  const [batchDeletePassword, setBatchDeletePassword] = useState('')
  const [batchDeleteError, setBatchDeleteError] = useState('')
  const [deletingBatch, setDeletingBatch] = useState(false)

  const openShareModal = async (p: Project) => {
    setShareProj(p)
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
    if (!shareProj) return
    const currentShared = Array.isArray(shareProj.summary?.shared_with)
      ? [...shareProj.summary.shared_with]
      : []
    let updated: string[]
    if (currentShared.includes(targetUsername)) {
      updated = currentShared.filter((u) => u !== targetUsername)
    } else {
      updated = [...currentShared, targetUsername]
    }
    try {
      await api.projectShare(session.token, shareProj.id, updated)
      const nextSummary = { ...(shareProj.summary || {}), shared_with: updated }
      setShareProj({ ...shareProj, summary: nextSummary })
      toast(
        updated.includes(targetUsername)
          ? `Quotation shared with @${targetUsername}`
          : `Access revoked for @${targetUsername}`,
      )
      loadProjects()
      onProjectsChanged?.()
    } catch (e) {
      toast((e as Error).message, 'err')
    }
  }

  const loadProjects = async () => {
    try {
      setLoading(true)
      const data = await api.projectList(session.token, org.id, true)
      setProjects(data)
    } catch (e) {
      toast((e as Error).message, 'err')
    } finally {
      setLoading(false)
    }
  }

  useEffect(() => {
    loadProjects()
  }, [org.id])

  // Count projects by scope (My Quotations vs Shared with Me)
  const { myProjectsCount, sharedProjectsCount } = useMemo(() => {
    let mine = 0
    let shared = 0
    for (const p of projects) {
      if (!showArchived && p.is_archived) continue
      const sharedList = (p.summary?.shared_with || (p as any).shared_with) as string[] | undefined
      const isSharedWithMe =
        Array.isArray(sharedList) &&
        (sharedList.includes(session.username) || (session.user_id ? sharedList.includes(session.user_id) : false)) &&
        p.created_by !== session.username &&
        p.created_by !== session.user_id

      if (isSharedWithMe) {
        if (p.status?.toLowerCase() !== 'approved') {
          shared++
        }
      } else if (p.created_by === session.username || p.created_by === session.user_id) {
        mine++
      }
    }
    return { myProjectsCount: mine, sharedProjectsCount: shared }
  }, [projects, showArchived, session.username, session.user_id])

  const remindersCount = useMemo(() => {
    return projects.filter((p) => p.summary?.reminder && !p.summary.reminder.dismissed).length
  }, [projects])

  const filteredProjects = useMemo(() => {
    return (projects || [])
      .filter((p) => {
        const sharedList = (p.summary?.shared_with || (p as any).shared_with) as string[] | undefined
        const isSharedWithMe =
          Array.isArray(sharedList) &&
          (sharedList.includes(session.username) || (session.user_id ? sharedList.includes(session.user_id) : false)) &&
          p.created_by !== session.username &&
          p.created_by !== session.user_id

        const isMine =
          p.created_by === session.user_id ||
          p.created_by === session.username

        // Staff can only see their own projects and active/quoted projects shared with them (approved quotations removed from shared view)
        if (!isAdminOrOwner) {
          if (!isMine && !isSharedWithMe) return false
          if (isSharedWithMe && p.status?.toLowerCase() === 'approved') return false
        }

        // Financial Year filter (default: current FY only; user can select 'ALL' or multiple FYs)
        if (!matchesFiscalYearFilter(p, selectedFys)) return false

        // Scope filter
        if (scopeFilter === 'mine' && !isMine) return false
        if (scopeFilter === 'shared' && !isSharedWithMe) return false
        if (scopeFilter === 'reminders' && (!p.summary?.reminder || p.summary.reminder.dismissed)) return false

        if (!showArchived && p.is_archived && statusFilter !== 'lost' && statusFilter !== 'loss') return false
        const isLoss = p.is_archived || p.status === 'loss' || p.status === 'archived' || p.status === 'lost'
        const isWon = !p.is_archived && (p.status === 'won' || p.status === 'approved')
        const isNegotiating = p.status === 'negotiating'
        const isQuoted = p.status === 'quoted' || (!isLoss && !isWon && !isNegotiating)

        // Pipeline board view shows all non-archived projects across stages
        if (viewMode === 'pipeline') {
          if (!showArchived && p.is_archived) return false
          return true
        }
        if (statusFilter === 'all') {
          // When actively searching, search across ALL projects regardless of status!
          // Only filter out won/loss if NOT actively searching (clean pipeline view)
          if (!search.trim() && (isWon || isLoss)) return false
        } else {
          if ((statusFilter === 'won' || statusFilter === 'approved') && !isWon) return false
          if ((statusFilter === 'loss' || statusFilter === 'lost') && !isLoss) return false
          if (statusFilter === 'negotiating' && !isNegotiating) return false
          if (statusFilter === 'quoted' && !isQuoted) return false
        }
        if (search.trim()) {
          const q = search.toLowerCase().trim()
          const clientInfo = p.summary?.client_info
          const clientName = (clientInfo?.buyer_name || '').toLowerCase()
          const qtnNo = (clientInfo?.qtn_no || '').toLowerCase()
          const desc = ((p.summary?.description || (p as any).description || '') + ' ' + (p.summary?.notes || '')).toLowerCase()

          // Matched client directory entry if saved
          const dirClient = (p.summary?.client_id ? savedClients.find((c) => c.id === p.summary.client_id) : null)
            || (clientName ? savedClients.find((c) => c.name.toLowerCase() === clientName) : null)

          // 1. Compile all address strings associated with this quotation
          const addressStrings = [
            clientInfo?.buyer_address,
            clientInfo?.buyer_city,
            clientInfo?.buyer_state,
            clientInfo?.buyer_pin,
            (clientInfo as any)?.address,
            (clientInfo as any)?.city,
            (clientInfo as any)?.state,
            (clientInfo as any)?.pin,
            (clientInfo as any)?.pincode,
            (clientInfo as any)?.site_address,
            (clientInfo as any)?.delivery_address,
            dirClient?.address,
            (dirClient as any)?.city,
            (dirClient as any)?.state,
            (dirClient as any)?.pincode,
            (p.inputs as any)?.['buyer_address'],
            (p.inputs as any)?.['client_address'],
            (p.inputs as any)?.['address'],
            (p.inputs as any)?.['site_address'],
            (p.inputs as any)?.['delivery_address'],
          ].filter(Boolean).map((a) => String(a).toLowerCase())

          const matchesAddress = addressStrings.some((addr) => addr.includes(q))

          // 2. Compile contact strings (phone, email, GSTIN)
          const contactStrings = [
            clientInfo?.buyer_contact,
            clientInfo?.buyer_email,
            clientInfo?.buyer_gst,
            dirClient?.contact,
            dirClient?.email,
            dirClient?.gst,
            (p.inputs as any)?.['buyer_gst'],
            (p.inputs as any)?.['gst'],
            (p.inputs as any)?.['buyer_contact'],
            (p.inputs as any)?.['buyer_email'],
            ...(clientInfo?.contacts || []).map((c: any) => c.phone || ''),
            ...(clientInfo?.contacts || []).map((c: any) => c.email || ''),
            ...(dirClient?.contacts || []).map((c: any) => c.phone || ''),
            ...(dirClient?.contacts || []).map((c: any) => c.email || ''),
          ].filter(Boolean).map((c) => String(c).toLowerCase())

          const matchesContact = contactStrings.some((c) => c.includes(q))

          // 3. Compile numeric / currency amount strings
          const tot = Number(p.total) || 0
          const beforeGst = Number(p.summary?.beforeGst) || 0
          const sales = Number(p.summary?.sales) || 0
          const grandTotal = Number(p.summary?.grandTotal) || 0
          const matchesAmount = [
            tot.toString(),
            Math.round(tot).toString(),
            tot.toLocaleString('en-IN').toLowerCase(),
            Math.round(tot).toLocaleString('en-IN').toLowerCase(),
            beforeGst.toString(),
            Math.round(beforeGst).toString(),
            sales.toString(),
            Math.round(sales).toString(),
            grandTotal.toString(),
            Math.round(grandTotal).toString(),
          ].some((amt) => amt.includes(q))

          // 4. Compile technical specs, models, bay, bodies, sets strings
          const systemsList = (p.summary?.systems || []) as any[]
          const bodies = p.summary?.bodies ?? p.summary?.total_combined_bodies
          const compartments = p.summary?.compartments
          const modelsList = (p.summary?.models || []) as any[]
          const rawInputs = (p.inputs || {}) as Record<string, any>
          const dfmuRows = Number(rawInputs['Total cal!A3'] ?? 0)
          const bayCount = Math.max(0, Math.round(dfmuRows / 2))

          const specsStrings = [
            'dfmu', 'sfmu', 'sffu', 'dffu', 'dfmu-d',
            'mechanical', 'double drive', 'reduction wheel',
            bayCount > 0 ? `${bayCount} bay` : '',
            bayCount > 0 ? `bay ${bayCount}` : '',
            bodies ? `${bodies} bodies` : '',
            bodies ? `${bodies} body` : '',
            compartments ? `${compartments} compartments` : '',
            ...systemsList.map((s) => s.name || ''),
            ...systemsList.map((s) => s.quantity ? `${s.quantity} sets` : ''),
            ...modelsList.map((m) => m.model_code || ''),
            ...modelsList.map((m) => m.name || ''),
            ...modelsList.map((m) => m.dimensions || ''),
          ].filter(Boolean).map((s) => String(s).toLowerCase())

          const matchesSpecs = specsStrings.some((s) => s.includes(q))

          // ── Field-Specific Searches ──
          if (searchType === 'qtn') {
            if (p.name.toLowerCase().includes(q)) return true
            if (qtnNo.includes(q)) return true
            if (p.id && p.id.toLowerCase().includes(q)) return true
            if (p.summary?.qtn_code && p.summary.qtn_code.toLowerCase().includes(q)) return true
            return false
          }

          if (searchType === 'client') {
            if (clientName.includes(q)) return true
            if (clientInfo?.kind_att && clientInfo.kind_att.toLowerCase().includes(q)) return true
            if (dirClient && dirClient.name.toLowerCase().includes(q)) return true
            if ((clientInfo?.contacts || []).some((c: any) => (c.name || '').toLowerCase().includes(q))) return true
            if ((dirClient?.contacts || []).some((c: any) => (c.name || '').toLowerCase().includes(q))) return true
            if ((p.inputs as any)?.['buyer_name']?.toLowerCase().includes(q)) return true
            if ((p.inputs as any)?.['client_name']?.toLowerCase().includes(q)) return true
            return false
          }

          if (searchType === 'address') {
            return matchesAddress
          }

          if (searchType === 'contact') {
            return matchesContact
          }

          if (searchType === 'amount') {
            return matchesAmount
          }

          if (searchType === 'description') {
            if (desc.includes(q)) return true
            if (Array.isArray(clientInfo?.custom_rows)) {
              for (const row of clientInfo.custom_rows) {
                if ((row.label || '').toLowerCase().includes(q)) return true
                if ((row.value || '').toLowerCase().includes(q)) return true
              }
            }
            return false
          }

          if (searchType === 'specs') {
            return matchesSpecs
          }

          // ── searchType === 'all' (Default: Searches EVERYTHING across all quotations) ──
          if (p.name.toLowerCase().includes(q)) return true
          if (qtnNo.includes(q)) return true
          if (clientName.includes(q)) return true
          if (matchesAddress) return true
          if (matchesContact) return true
          if (matchesAmount) return true
          if (matchesSpecs) return true
          if (desc.includes(q)) return true
          if (p.created_by && p.created_by.toLowerCase().includes(q)) return true
          if (p.id && p.id.toLowerCase().includes(q)) return true
          if (p.status && p.status.toLowerCase().includes(q)) return true

          // Client contacts names and designations
          if (clientInfo) {
            if (clientInfo.kind_att && clientInfo.kind_att.toLowerCase().includes(q)) return true
            if (Array.isArray(clientInfo.contacts)) {
              for (const c of clientInfo.contacts) {
                if ((c.name || '').toLowerCase().includes(q)) return true
                if ((c.designation || '').toLowerCase().includes(q)) return true
              }
            }
            if (Array.isArray(clientInfo.custom_rows)) {
              for (const row of clientInfo.custom_rows) {
                if ((row.label || '').toLowerCase().includes(q)) return true
                if ((row.value || '').toLowerCase().includes(q)) return true
              }
            }
          }

          // Directory client matches
          if (dirClient) {
            if (dirClient.name.toLowerCase().includes(q)) return true
            if (Array.isArray(dirClient.contacts)) {
              for (const c of dirClient.contacts) {
                if ((c.name || '').toLowerCase().includes(q)) return true
                if ((c.designation || '').toLowerCase().includes(q)) return true
              }
            }
          }

          // Summary and inputs JSON fallback for custom specifications
          if (p.summary) {
            try {
              const summaryStr = JSON.stringify(p.summary).toLowerCase()
              if (summaryStr.includes(q)) return true
            } catch {}
          }

          if (p.inputs) {
            try {
              const inputsStr = JSON.stringify(p.inputs).toLowerCase()
              if (inputsStr.includes(q)) return true
            } catch {}
          }

          return false
        }
        return true
      })
      .sort((a, b) => {
        let cmp = 0
        if (sortBy === 'name') cmp = a.name.localeCompare(b.name)
        else if (sortBy === 'total') cmp = (Number(a.total) || 0) - (Number(b.total) || 0)
        else if (sortBy === 'created_at') cmp = new Date(a.created_at).getTime() - new Date(b.created_at).getTime()
        else cmp = new Date(a.updated_at).getTime() - new Date(b.updated_at).getTime()
        return sortAsc ? cmp : -cmp
      })
  }, [projects, isAdminOrOwner, session.username, session.user_id, scopeFilter, selectedFys, search, searchType, statusFilter, showArchived, sortBy, sortAsc, savedClients, viewMode])

  // Group filtered projects into Project Folders
  const projectFolders = useMemo(() => {
    const folders = groupProjectsIntoFolders(filteredProjects, currentProjectId)
    return folders.sort((a, b) => {
      let cmp = 0
      if (sortBy === 'name') cmp = a.folderName.localeCompare(b.folderName)
      else if (sortBy === 'total') cmp = (Number(a.latestRevision.total) || 0) - (Number(b.latestRevision.total) || 0)
      else if (sortBy === 'created_at') cmp = new Date(a.revisions[0].created_at).getTime() - new Date(b.revisions[0].created_at).getTime()
      else cmp = new Date(a.latestRevision.updated_at).getTime() - new Date(b.latestRevision.updated_at).getTime()
      return sortAsc ? cmp : -cmp
    })
  }, [filteredProjects, currentProjectId, sortBy, sortAsc])
  // Sales Pipeline Kanban Board States & Data
  const [dragOverColumn, setDragOverColumn] = useState<string | null>(null)
  const [draggedProjectId, setDraggedProjectId] = useState<string | null>(null)

  const PIPELINE_COLUMNS: { id: ProjectStatus; title: string; color: string; icon: string; bg: string; border: string }[] = [
    { id: 'draft', title: 'Draft / Preparation', color: '#475569', icon: '📝', bg: '#f8fafc', border: '#cbd5e1' },
    { id: 'quoted', title: 'Quotation Sent', color: '#1d4ed8', icon: '📤', bg: '#eff6ff', border: '#93c5fd' },
    { id: 'negotiating', title: 'Negotiating / Review', color: '#b45309', icon: '🤝', bg: '#fffbeb', border: '#fcd34d' },
    { id: 'won', title: 'Won / PO Received', color: '#15803d', icon: '🎉', bg: '#f0fdf4', border: '#86efac' },
    { id: 'loss', title: 'Lost / Closed', color: '#b91c1c', icon: '✕', bg: '#fef2f2', border: '#fca5a5' },
  ]

  const pipelineData = useMemo(() => {
    const groups: Record<string, Project[]> = {
      draft: [],
      quoted: [],
      negotiating: [],
      won: [],
      loss: [],
    }
    for (const p of filteredProjects) {
      if (p.is_archived || p.status === 'loss' || p.status === 'lost' || p.status === 'archived') {
        groups.loss.push(p)
      } else if (p.status === 'won' || p.status === 'approved') {
        groups.won.push(p)
      } else if (p.status === 'negotiating') {
        groups.negotiating.push(p)
      } else if (p.status === 'draft') {
        groups.draft.push(p)
      } else {
        groups.quoted.push(p)
      }
    }
    const totals = {
      draft: groups.draft.reduce((sum, p) => sum + (Number(p.total) || 0), 0),
      quoted: groups.quoted.reduce((sum, p) => sum + (Number(p.total) || 0), 0),
      negotiating: groups.negotiating.reduce((sum, p) => sum + (Number(p.total) || 0), 0),
      won: groups.won.reduce((sum, p) => sum + (Number(p.total) || 0), 0),
      loss: groups.loss.reduce((sum, p) => sum + (Number(p.total) || 0), 0),
    }
    const activePipelineTotal = totals.draft + totals.quoted + totals.negotiating
    return { groups, totals, activePipelineTotal }
  }, [filteredProjects])

  const handleUpdateProjectStatus = async (projectId: string, newStatus: ProjectStatus) => {
    const target = projects.find((p) => p.id === projectId)
    if (!target) return
    if (target.status === newStatus) return
    try {
      await api.projectUpdate(
        session.token,
        target.id,
        target.name,
        newStatus,
        target.inputs,
        target.formula_overrides,
        target.total,
        target.summary,
        `Quotation status moved to ${newStatus}`
      )
      toast(`Quotation moved to ${newStatus.toUpperCase()}`)
      loadProjects()
      onProjectsChanged?.()
    } catch (e) {
      toast((e as Error).message, 'err')
    }
  }

  // Selection helpers
  const allVisibleSelected = useMemo(() => {
    if (filteredProjects.length === 0) return false
    return filteredProjects.every((p) => selectedIds.has(p.id))
  }, [filteredProjects, selectedIds])

  const someVisibleSelected = useMemo(() => {
    return filteredProjects.some((p) => selectedIds.has(p.id))
  }, [filteredProjects, selectedIds])

  const handleToggleSelectAll = () => {
    if (allVisibleSelected) {
      setSelectedIds((prev) => {
        const next = new Set(prev)
        for (const p of filteredProjects) {
          next.delete(p.id)
        }
        return next
      })
    } else {
      setSelectedIds((prev) => {
        const next = new Set(prev)
        for (const p of filteredProjects) {
          next.add(p.id)
        }
        return next
      })
    }
  }

  const handleToggleSelectOne = (id: string, checked: boolean) => {
    setSelectedIds((prev) => {
      const next = new Set(prev)
      if (checked) {
        next.add(id)
      } else {
        next.delete(id)
      }
      return next
    })
  }

  const handleToggleFolderSelect = (folder: ProjectFolder, checked: boolean) => {
    setSelectedIds((prev) => {
      const next = new Set(prev)
      for (const rev of folder.revisions) {
        if (checked) {
          next.add(rev.id)
        } else {
          next.delete(rev.id)
        }
      }
      return next
    })
  }

  // Actions
  const handleRevise = async () => {
    if (!reviseProj) return
    const nm = revName.trim()
    if (!nm) return toast('Revision name cannot be empty', 'err')
    try {
      const res = await api.projectDuplicate(session.token, reviseProj.id, nm, false)
      const revNow = new Date()
      const revSummary = {
        ...(reviseProj.summary || {}),
        fy: getFiscalYear(revNow),
        client_info: {
          ...(reviseProj.summary?.client_info || {}),
          qtn_no: buildOfficialQtnNo(res.name || nm, revNow),
          date: formatPdfDate(revNow),
        },
      }
      try {
        await api.projectUpdate(
          session.token,
          res.id,
          res.name,
          'quoted',
          reviseProj.inputs,
          reviseProj.formula_overrides,
          reviseProj.total,
          revSummary,
          `Revision ${res.name} created - Status set to Quoted`,
        )
      } catch (err) {
        await api.projectUpdate(
          session.token,
          res.id,
          res.name,
          'quoted',
          reviseProj.inputs,
          reviseProj.formula_overrides,
          reviseProj.total,
          revSummary,
          `Revision ${res.name} created`,
        )
      }
      toast(`Created revision “${res.name}” (Status: Quoted)`)
      setReviseProj(null)
      await loadProjects()
      onProjectsChanged?.()
      onOpenProject(res.id)
    } catch (e) {
      toast((e as Error).message, 'err')
    }
  }

  const handleRename = async () => {
    if (!renameProj) return
    const nm = newName.trim()
    if (!nm) return toast('Project name cannot be empty', 'err')
    try {
      await api.projectUpdate(session.token, renameProj.id, nm, renameProj.status, renameProj.inputs, renameProj.formula_overrides, renameProj.total, renameProj.summary, 'Renamed project')
      toast('Project renamed')
      setRenameProj(null)
      loadProjects()
      onProjectsChanged?.()
    } catch (e) {
      toast((e as Error).message, 'err')
    }
  }

  const handleDuplicate = async () => {
    if (!duplicateProj) return
    const nm = dupName.trim()
    if (!nm) return toast('Duplicate project name cannot be empty', 'err')
    try {
      const res = await api.projectDuplicate(session.token, duplicateProj.id, nm, dupUseLatest)
      const dupNow = new Date()
      const dupSummary = {
        ...(duplicateProj.summary || {}),
        fy: getFiscalYear(dupNow),
        client_info: {
          ...(duplicateProj.summary?.client_info || {}),
          qtn_no: buildOfficialQtnNo(res.name || nm, dupNow),
          date: formatPdfDate(dupNow),
        },
      }
      try {
        await api.projectUpdate(
          session.token,
          res.id,
          nm,
          'quoted',
          duplicateProj.inputs,
          duplicateProj.formula_overrides,
          duplicateProj.total,
          dupSummary,
          `Quotation ${nm} created as copy - Status set to Quoted`,
        )
      } catch {}
      toast(`Project copied successfully (Status: Quoted)`)
      setDuplicateProj(null)
      await loadProjects()
      onProjectsChanged?.()
      onOpenProject(res.id)
    } catch (e) {
      toast((e as Error).message, 'err')
    }
  }

  const handleArchive = async (p: Project) => {
    if (!isAdminOrOwner) {
      toast('Staff members cannot archive or restore projects.', 'err')
      return
    }
    const toArchive = !p.is_archived
    const ok = await confirm({
      title: toArchive ? `Archive “${p.name}”?` : `Unarchive “${p.name}”?`,
      body: toArchive
        ? 'Archived projects are hidden from normal views and dashboard statistics but their historical calculations are preserved.'
        : 'The project will be restored to active projects.',
      confirmLabel: toArchive ? 'Archive' : 'Restore',
    })
    if (!ok) return
    try {
      await api.projectArchive(session.token, p.id, toArchive)
      toast(toArchive ? 'Project archived' : 'Project restored')
      loadProjects()
      onProjectsChanged?.()
    } catch (e) {
      toast((e as Error).message, 'err')
    }
  }

  const handleDelete = (p: Project) => {
    if (!isAdminOrOwner) {
      toast('Staff members cannot delete projects or quotations.', 'err')
      return
    }
    setSingleDeleteProj(p)
    setConfirmNameInput('')
  }

  const handleConfirmSingleDelete = async () => {
    if (!singleDeleteProj) return
    if (confirmNameInput.trim() !== singleDeleteProj.name.trim()) {
      toast('Quotation name does not match', 'err')
      return
    }
    setDeletingSingle(true)
    try {
      await deleteProject(session, org, singleDeleteProj.id)
      toast(`Deleted quotation “${singleDeleteProj.name}”`)
      setSelectedIds((prev) => {
        const next = new Set(prev)
        next.delete(singleDeleteProj.id)
        return next
      })
      setSingleDeleteProj(null)
      setConfirmNameInput('')
      await loadProjects()
      onProjectsChanged?.()
    } catch (e) {
      toast((e as Error).message, 'err')
    } finally {
      setDeletingSingle(false)
    }
  }

  const handleTriggerBatchDelete = () => {
    if (!isAdminOrOwner) {
      toast('Staff members cannot delete projects or quotations.', 'err')
      return
    }
    if (selectedIds.size === 0) return
    if (selectedIds.size === 1) {
      const singleId = Array.from(selectedIds)[0]
      const target = projects.find((p) => p.id === singleId)
      if (target) {
        handleDelete(target)
      }
      return
    }
    setShowBatchDeleteModal(true)
    setBatchDeletePassword('')
    setBatchDeleteError('')
  }

  const handleConfirmBatchDelete = async () => {
    if (!batchDeletePassword) return
    setDeletingBatch(true)
    setBatchDeleteError('')
    try {
      let freshSession: Session
      try {
        freshSession = await api.login(session.username, batchDeletePassword)
      } catch {
        setBatchDeleteError('Incorrect password. Please enter your valid account password.')
        setDeletingBatch(false)
        return
      }

      // Synchronize fresh session immediately across storage, state and local session ref
      tokenStore.set(freshSession.token)
      session.token = freshSession.token
      onSessionChange?.(freshSession)

      const idsToDelete = Array.from(selectedIds)
      for (const id of idsToDelete) {
        await deleteProject(freshSession, org, id)
      }

      toast(`Successfully deleted ${idsToDelete.length} quotations`)
      setSelectedIds(new Set())
      setShowBatchDeleteModal(false)
      setBatchDeletePassword('')
      await loadProjects()
      onProjectsChanged?.()
    } catch (e) {
      toast((e as Error).message, 'err')
    } finally {
      setDeletingBatch(false)
    }
  }

  // Render individual quotation revision file row
  const renderProjectRow = (p: Project, isSubRow = false) => {
    const isSelected = p.id === currentProjectId
    const hasNewerTemplate = (org.current_template_version || 1) > p.template_version
    const rev = parseRevision(p.name)

    const buyer = p.summary?.client_info?.buyer_name || ''
    const isWon = !p.is_archived && (p.status === 'won' || p.status === 'approved')
    const isLoss = p.is_archived || p.status === 'loss' || p.status === 'archived' || p.status === 'lost'
    const isNegotiating = p.status === 'negotiating'
    const statusLabel = isWon ? 'Won' : isLoss ? 'Loss' : isNegotiating ? 'Negotiating' : 'Quoted'
    const statusClass = isWon ? 'won' : isLoss ? 'loss' : isNegotiating ? 'negotiating' : 'quoted'

    return (
      <tr
        key={p.id}
        className={`${isSubRow ? 'file-subrow' : ''} ${isSelected ? 'selected-row' : ''} ${p.is_archived ? 'dim' : ''} ${selectedIds.has(p.id) ? 'batch-selected-row' : ''}`}
      >
        <td style={{ width: '40px', textAlign: 'center' }} onClick={(e) => e.stopPropagation()}>
          <input
            type="checkbox"
            aria-label={`Select quotation ${p.name}`}
            checked={selectedIds.has(p.id)}
            onChange={(e) => handleToggleSelectOne(p.id, e.target.checked)}
          />
        </td>
        <td className={isSubRow ? 'file-subrow-cell' : ''}>
          {isSubRow && (
            <>
              <div className="file-tree-guide" />
              <div className="file-tree-branch" />
            </>
          )}
          <div style={{ display: 'flex', flexDirection: 'column', gap: 2 }}>
            <div className="row gap" style={{ flexWrap: 'wrap' }}>
              {isSubRow && <span className="file-icon" title="Quotation File">{Icon.fileText}</span>}
              <button className="proj-title-btn" onClick={() => onOpenProject(p.id)}>
                <strong>{p.name}</strong>
              </button>
              {rev.currentRev ? (
                <span className="chip primary sm font-mono font-bold" title={`Revision ${rev.currentRev}`}>
                  Rev {rev.currentRev}
                </span>
              ) : null}
              {isSelected && <span className="chip ok">Current</span>}
              {p.is_archived && <span className="chip muted-chip">Archived</span>}
              {p.summary?.reminder && !p.summary.reminder.dismissed && (
                <span
                  className="chip warning sm"
                  style={{ fontSize: '10px', padding: '1px 5px', gap: '3px', fontWeight: 600 }}
                  title={`Alarm Reminder: ${p.summary.reminder.note || 'Follow up'} (Due: ${new Date(p.summary.reminder.date).toLocaleDateString()} at ${new Date(p.summary.reminder.date).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })})`}
                >
                  🔔 {new Date(p.summary.reminder.date).toLocaleDateString('en-IN', { month: 'short', day: 'numeric' })}, {new Date(p.summary.reminder.date).toLocaleTimeString('en-IN', { hour: '2-digit', minute: '2-digit' })}
                </span>
              )}
              {(() => {
                const shared = (p.summary?.shared_with || (p as any).shared_with) as string[] | undefined
                const isSharedToMe =
                  Array.isArray(shared) &&
                  (shared.includes(session.username) || (session.user_id && shared.includes(session.user_id))) &&
                  p.created_by !== session.username
                const count = Array.isArray(shared) ? shared.length : 0
                if (isSharedToMe) {
                  return (
                    <span className="chip purple sm" title={`Shared with you by ${p.created_by || 'teammate'}`}>
                      Shared with you
                    </span>
                  )
                }
                if (count > 0 && Array.isArray(shared)) {
                  return (
                    <span className="chip sm font-mono" title={`Shared with: ${shared.join(', ')}`}>
                      Shared ({count})
                    </span>
                  )
                }
                return null
              })()}
            </div>
            {Boolean(p.summary?.description || (p as any).description) && (
              <div style={{ fontSize: '11px', color: '#64748b', fontStyle: 'italic', maxWidth: '380px', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                {p.summary?.description || (p as any).description}
              </div>
            )}
          </div>
        </td>
        <td>
          {(() => {
              const clientInfo = p.summary?.client_info
            const clientName = clientInfo?.buyer_name || ''
            const qtnNo = clientInfo?.qtn_no
            const rowFy = getProjectFiscalYear(p)
            const city = clientInfo?.buyer_city || ''
            const address = clientInfo?.buyer_address || ''
            if (!clientName) {
              return (
                <div style={{ display: 'flex', gap: 6, flexWrap: 'wrap', alignItems: 'center' }}>
                  <span className="muted tiny">—</span>
                  <span className="chip sm font-mono" style={{ fontSize: '10px', background: '#f1f5f9', color: '#475569' }}>FY {rowFy}</span>
                </div>
              )
            }
            return (
              <div style={{ display: 'flex', flexDirection: 'column', gap: 2 }}>
                <div style={{ display: 'flex', gap: 6, alignItems: 'center', flexWrap: 'wrap' }}>
                  <strong style={{ color: '#0f172a', fontSize: '12px' }}>{clientName}</strong>
                  <span className="chip sm font-mono" style={{ fontSize: '10px', background: rowFy === currentFy ? '#eff6ff' : '#f1f5f9', color: rowFy === currentFy ? '#1d4ed8' : '#475569' }}>
                    FY {rowFy}
                  </span>
                </div>
                <div style={{ display: 'flex', gap: 6, flexWrap: 'wrap', alignItems: 'center' }}>
                  {qtnNo && <span className="tiny font-mono text-muted" style={{ fontSize: '10.5px' }}>{qtnNo}</span>}
                  {(city || address) && (
                    <span className="tiny text-muted" style={{ fontSize: '10.5px', color: '#64748b', display: 'inline-flex', alignItems: 'center', gap: '2px' }} title={address}>
                      📍 {city || (address.length > 25 ? address.slice(0, 25) + '…' : address)}
                    </span>
                  )}
                </div>
              </div>
            )
          })()}
        </td>
        <td>
          <span className={`status-pill status-${statusClass}`}>{statusLabel}</span>
          {p.summary?.discount_request?.status === 'pending' && (
            <span
              className="chip warning sm"
              style={{ marginLeft: '4px', fontSize: '11px', padding: '2px 6px' }}
              title={`Staff @${p.summary.discount_request.requested_by} requested ${p.summary.discount_request.requested_pct}% discount`}
            >
              ⏳ {p.summary.discount_request.requested_pct}% Disc Req
            </span>
          )}
        </td>
        <td>
          <div className="template-badge-wrap">
            <span className="version-pill">v{p.template_version}</span>
            {hasNewerTemplate && (
              <span
                className="upgrade-hint"
                title={`New default template v${org.current_template_version} available`}
              >
                ↑ v{org.current_template_version} avail
              </span>
            )}
          </div>
        </td>
        <td className="r font-bold text-green">
          {p.total ? inr(Number(p.total)) : '—'}
        </td>
        <td className="muted tiny">
          {p.created_by || 'Unknown'}
        </td>
        <td className="muted tiny">{timeAgo(p.updated_at)}</td>
        <td className="r">
          <div className="row-actions">
            <button
              className="btn primary sm"
              onClick={() => onOpenProject(p.id)}
              title="Open Calculator"
            >
              Open
            </button>
            <button
              className="btn ghost sm"
              title={`Revise quotation (Create Rev ${parseRevision(p.name).nextRev})`}
              onClick={() => {
                const info = parseRevision(p.name)
                setReviseProj(p)
                setRevName(info.nextName)
              }}
            >
              {Icon.refresh} Revise
            </button>
            <button
              className="icon-btn sm"
              title="Share quotation with team members"
              onClick={() => openShareModal(p)}
            >
              {Icon.users}
            </button>
            <button
              className="icon-btn sm"
              title="Rename"
              onClick={() => {
                setRenameProj(p)
                setNewName(p.name)
              }}
            >
              {Icon.edit}
            </button>
            <button
              className="icon-btn sm"
              title="Duplicate snapshot"
              onClick={() => {
                setDuplicateProj(p)
                setDupName(`${p.name} (Copy)`)
                setDupUseLatest(false)
              }}
            >
              {Icon.copy}
            </button>
            {isAdminOrOwner && (
              <button
                className="icon-btn sm"
                title={p.is_archived ? 'Unarchive' : 'Archive'}
                onClick={() => handleArchive(p)}
              >
                {Icon.archive}
              </button>
            )}
            {isAdminOrOwner && (
              <button
                className="icon-btn sm del"
                title="Delete permanently"
                onClick={() => handleDelete(p)}
              >
                {Icon.trash}
              </button>
            )}
          </div>
        </td>
      </tr>
    )
  }

  return (
    <div className="projects-page">
      {/* Page Header with Folders / Files View Toggle */}
      <div className="row between mb-2" style={{ flexWrap: 'wrap', gap: 10 }}>
        <div className="row gap" style={{ flexWrap: 'wrap' }}>
          <h2 style={{ fontSize: '18px' }}>
            {isAdminOrOwner
              ? scopeFilter === 'shared'
                ? 'Shared Quotations'
                : scopeFilter === 'mine'
                  ? 'My Projects & Quotations'
                  : 'Projects & Quotations'
              : scopeFilter === 'shared'
                ? 'Quotations Shared with Me'
                : scopeFilter === 'mine'
                  ? 'My Quotations'
                  : 'My Projects & Shared Quotations'}
          </h2>
          <div className="view-mode-tabs">
            <button
              className={`view-mode-btn ${viewMode === 'folders' ? 'active' : ''}`}
              onClick={() => setViewMode('folders')}
              title="Group quotations into Project Folders"
            >
              <span className="folder-icon">{Icon.folder}</span> Folders ({projectFolders.length})
            </button>
            <button
              className={`view-mode-btn ${viewMode === 'flat' ? 'active' : ''}`}
              onClick={() => setViewMode('flat')}
              title="Show all individual revision files"
            >
              <span className="file-icon">{Icon.fileText}</span> All Files ({filteredProjects.length})
            </button>
            <button
              className={'view-mode-btn ' + (viewMode === 'pipeline' ? 'active' : '')}
              onClick={() => setViewMode('pipeline')}
              title="Sales Pipeline Funnel Kanban Board"
            >
              <span style={{ fontSize: 13 }}>📋</span> Pipeline ({filteredProjects.length})
            </button>
          </div>
          {viewMode === 'folders' && projectFolders.length > 0 && (
            <div className="row gap" style={{ marginLeft: 4 }}>
              <button className="link-inline tiny" onClick={expandAllFolders}>Expand All</button>
              <span className="tiny text-muted">·</span>
              <button className="link-inline tiny" onClick={collapseAllFolders}>Collapse All</button>
            </div>
          )}
        </div>
        <button className="btn primary sm" onClick={onNewProject}>
          {Icon.plus} New Project
        </button>
      </div>

      {/* Filter and Search Bar */}
      <div className="card filter-bar">
        <div className="search-wrap" style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
          <select
            value={searchType}
            onChange={(e) => setSearchType(e.target.value as any)}
            className="search-type-select"
            title="Search all fields or choose a specific quotation field"
            style={{
              height: '32px',
              border: '1px solid #cbd5e1',
              borderRadius: '6px',
              background: '#f8fafc',
              fontSize: '11.5px',
              fontWeight: 600,
              color: '#334155',
              padding: '0 8px',
              cursor: 'pointer',
              outline: 'none',
              flexShrink: 0,
            }}
          >
            <option value="all">🔍 All Fields (Search Everything)</option>
            <option value="qtn">📄 Quotation # / Project Name</option>
            <option value="client">🏢 Client &amp; Company Name</option>
            <option value="address">📍 Address, City &amp; State</option>
            <option value="contact">📞 Phone, Email &amp; GST</option>
            <option value="amount">💰 Amount &amp; Total Value (₹)</option>
            <option value="description">📝 Description &amp; Client Notes</option>
            <option value="specs">⚙️ Specs, Models, Bay &amp; Sets</option>
          </select>
          <span className="search-icon">{Icon.search}</span>
          <input
            type="text"
            className="search-input"
            placeholder={
              searchType === 'qtn'
                ? `Search by quotation # (e.g. MSSLP/${currentFy}/...), project name…`
                : searchType === 'client'
                ? 'Search by client, company name, kind attn, person…'
                : searchType === 'address'
                ? 'Search by address, city, state, pin, site location…'
                : searchType === 'contact'
                ? 'Search by mobile number, email address, GSTIN…'
                : searchType === 'amount'
                ? 'Search by quotation amount, sales total (e.g. 150000, ₹1,50,000)…'
                : searchType === 'description'
                ? 'Search by project description, scope, client notes…'
                : searchType === 'specs'
                ? 'Search by model (DFMU, SFMU, SFFU), bay count, dimensions…'
                : 'Search anything across all quotations (client, quotation #, address, specs, notes, amount)…'
            }
            value={search}
            onChange={(e) => setSearch(e.target.value)}
          />
          {search && (
            <button className="icon-btn sm" onClick={() => setSearch('')}>
              {Icon.close}
            </button>
          )}
        </div>

        <div className="filter-controls">
          {/* Financial Year (FY) Multi-Select Filter Dropdown — Defaults to Current FY */}
          <div ref={fyDropdownRef} style={{ position: 'relative' }}>
            <button
              type="button"
              className="btn ghost sm"
              onClick={() => setFyDropdownOpen((o) => !o)}
              title="Filter quotations by Financial Year (default: Current FY). Select All FY or multiple years."
              style={{
                height: '30px',
                border: selectedFys.includes('ALL') || selectedFys.length > 1 || selectedFys[0] !== currentFy
                  ? '1.5px solid #2563eb'
                  : '1px solid #cbd5e1',
                background: selectedFys.includes('ALL') || selectedFys.length > 1 || selectedFys[0] !== currentFy
                  ? '#eff6ff'
                  : '#ffffff',
                color: '#1e293b',
                fontWeight: 600,
                fontSize: '12px',
                display: 'inline-flex',
                alignItems: 'center',
                gap: '6px',
                padding: '0 10px',
                borderRadius: '6px',
              }}
            >
              <span>📅 {fyButtonLabel}</span>
              <span style={{ fontSize: '10px', opacity: 0.7 }}>{fyDropdownOpen ? '▲' : '▼'}</span>
            </button>

            {fyDropdownOpen && (
              <div
                style={{
                  position: 'absolute',
                  top: 'calc(100% + 4px)',
                  left: 0,
                  zIndex: 60,
                  minWidth: '220px',
                  background: '#ffffff',
                  border: '1px solid #cbd5e1',
                  borderRadius: '8px',
                  boxShadow: '0 10px 25px rgba(15, 23, 42, 0.14)',
                  padding: '6px',
                }}
              >
                <div style={{ fontSize: '10.5px', fontWeight: 700, color: '#64748b', padding: '4px 8px', textTransform: 'uppercase', letterSpacing: '0.4px' }}>
                  Financial Year Filter
                </div>

                <label
                  style={{
                    display: 'flex',
                    alignItems: 'center',
                    justifyContent: 'space-between',
                    gap: '8px',
                    padding: '6px 8px',
                    borderRadius: '6px',
                    cursor: 'pointer',
                    fontSize: '12px',
                    fontWeight: 600,
                    background: selectedFys.includes('ALL') ? '#eff6ff' : 'transparent',
                    color: selectedFys.includes('ALL') ? '#1d4ed8' : '#0f172a',
                  }}
                >
                  <span style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                    <input
                      type="checkbox"
                      checked={selectedFys.includes('ALL')}
                      onChange={toggleAllFys}
                    />
                    <span>All FY</span>
                  </span>
                  <span className="tiny muted">{projects.length}</span>
                </label>

                <div style={{ height: '1px', background: '#e2e8f0', margin: '4px 0' }} />

                <div style={{ maxHeight: '190px', overflowY: 'auto', display: 'flex', flexDirection: 'column', gap: '2px' }}>
                  {availableFys.map((fy) => {
                    const checked = !selectedFys.includes('ALL') && selectedFys.includes(fy)
                    const isCurr = fy === currentFy
                    return (
                      <label
                        key={fy}
                        style={{
                          display: 'flex',
                          alignItems: 'center',
                          justifyContent: 'space-between',
                          gap: '8px',
                          padding: '6px 8px',
                          borderRadius: '6px',
                          cursor: 'pointer',
                          fontSize: '12px',
                          fontWeight: isCurr ? 700 : 500,
                          background: checked ? '#f0fdf4' : 'transparent',
                          color: '#0f172a',
                        }}
                      >
                        <span style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                          <input
                            type="checkbox"
                            checked={checked}
                            onChange={() => toggleSingleFy(fy)}
                          />
                          <span>FY {fy}{isCurr ? ' (Current)' : ''}</span>
                        </span>
                        <span className="tiny muted">{fyCounts[fy] || 0}</span>
                      </label>
                    )
                  })}
                </div>

                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginTop: '6px', paddingTop: '6px', borderTop: '1px solid #e2e8f0' }}>
                  <button
                    type="button"
                    className="link-inline tiny"
                    onClick={() => {
                      setSelectedFys([currentFy])
                      setFyDropdownOpen(false)
                    }}
                  >
                    Reset to Current FY
                  </button>
                  <button
                    type="button"
                    className="btn ghost sm"
                    style={{ padding: '2px 8px', fontSize: '11px' }}
                    onClick={() => setFyDropdownOpen(false)}
                  >
                    Done
                  </button>
                </div>
              </div>
            )}
          </div>

          {/* Scope Filter Tabs: All, My Quotations, Shared with Me */}
          <div className="status-tabs scope-filter-tabs">
            <button
              className={`status-tab ${scopeFilter === 'all' ? 'active' : ''}`}
              onClick={() => setScopeFilter('all')}
              title="Show all accessible quotations"
            >
              All {isAdminOrOwner ? `(${projects.length})` : `(${myProjectsCount + sharedProjectsCount})`}
            </button>
            <button
              className={`status-tab ${scopeFilter === 'mine' ? 'active' : ''}`}
              onClick={() => setScopeFilter('mine')}
              title="Show quotations created by me"
            >
              My Quotations ({myProjectsCount})
            </button>
            <button
              className={`status-tab ${scopeFilter === 'shared' ? 'active' : ''}`}
              onClick={() => setScopeFilter('shared')}
              title="Show quotations shared with me by other staff"
            >
              🤝 Shared with Me {sharedProjectsCount > 0 ? `(${sharedProjectsCount})` : ''}
            </button>
            {remindersCount > 0 && (
              <button
                className={`status-tab ${scopeFilter === 'reminders' ? 'active' : ''}`}
                onClick={() => setScopeFilter('reminders')}
                style={{ color: scopeFilter === 'reminders' ? '#fff' : '#b45309', background: scopeFilter === 'reminders' ? '#f59e0b' : '#fffbeb', borderColor: '#fcd34d', fontWeight: 700 }}
                title="Show quotations with active follow-up alarms & reminders"
              >
                🔔 Reminders ({remindersCount})
              </button>
            )}
          </div>

          <div className="status-tabs">
            {[
              { id: 'all', label: 'ACTIVE (QUOTED & NEGOTIATING)' },
              { id: 'quoted', label: 'QUOTED' },
              { id: 'negotiating', label: 'NEGOTIATING' },
              { id: 'won', label: 'WON' },
              { id: 'loss', label: 'LOSS' },
            ].map(({ id, label }) => (
              <button
                key={id}
                className={`status-tab ${statusFilter === id ? 'active' : ''}`}
                onClick={() => setStatusFilter(id)}
              >
                {label}
              </button>
            ))}
          </div>

          <label className="checkbox-label">
            <input
              type="checkbox"
              checked={showArchived}
              onChange={(e) => setShowArchived(e.target.checked)}
            />
            <span>Show Archived</span>
          </label>

          <div className="sort-group">
            <span className="tiny muted">Sort:</span>
            <select
              value={sortBy}
              onChange={(e) => setSortBy(e.target.value as SortField)}
              className="select-input"
            >
              <option value="updated_at">Recently Updated</option>
              <option value="name">Project Name</option>
              <option value="total">Quotation Amount</option>
              <option value="created_at">Date Created</option>
            </select>
            <button
              className="icon-btn sm"
              title={sortAsc ? 'Ascending' : 'Descending'}
              onClick={() => setSortAsc(!sortAsc)}
            >
              {sortAsc ? '↑' : '↓'}
            </button>
          </div>
        </div>
      </div>

      {/* Batch Actions Bar */}
      {selectedIds.size > 0 && (
        <div className="batch-actions-bar">
          <div className="batch-actions-left">
            <span className="batch-count-badge">
              {selectedIds.size} quotation{selectedIds.size > 1 ? 's' : ''} selected
            </span>
            <button
              type="button"
              className="btn ghost sm"
              onClick={() => setSelectedIds(new Set())}
            >
              Clear Selection
            </button>
          </div>
          <div className="batch-actions-right">
            {isAdminOrOwner && (
              <button
                type="button"
                className="btn danger sm"
                onClick={handleTriggerBatchDelete}
                style={{ display: 'flex', alignItems: 'center', gap: '6px' }}
              >
                {Icon.trash} Delete Selected ({selectedIds.size})
              </button>
            )}
          </div>
        </div>
      )}

      {/* Projects Table */}
      <div className="card table-card">
        {loading ? (
          <div className="page-loading">
            <Spinner size={28} />
            <p className="muted">Loading projects…</p>
          </div>
        ) : filteredProjects.length === 0 ? (
          <div className="empty">
            <p>
              {scopeFilter === 'shared'
                ? 'No quotations have been shared with you yet'
                : scopeFilter === 'mine'
                  ? 'You have not created any quotations yet'
                  : 'No projects match your current filters'}
            </p>
            {search || statusFilter !== 'all' || showArchived ? (
              <button
                className="btn ghost sm"
                onClick={() => {
                  setSearch('')
                  setStatusFilter('all')
                  setShowArchived(false)
                }}
              >
                Reset Filters
              </button>
            ) : (
              <button className="btn primary sm" onClick={onNewProject}>
                {Icon.plus} Create First Project
              </button>
            )}
          </div>
        ) : viewMode === 'pipeline' ? (
          <div>
            {/* Pipeline Funnel Header */}
            <div
              style={{
                padding: '12px 16px',
                background: 'linear-gradient(135deg, #f8fafc 0%, #f1f5f9 100%)',
                borderBottom: '1px solid #e2e8f0',
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'space-between',
                flexWrap: 'wrap',
                gap: '12px',
              }}
            >
              <div style={{ display: 'flex', alignItems: 'center', gap: '16px', flexWrap: 'wrap' }}>
                <div>
                  <span style={{ fontSize: '10.5px', fontWeight: 700, color: '#64748b', textTransform: 'uppercase', letterSpacing: '0.04em' }}>
                    Active Pipeline Value
                  </span>
                  <div style={{ fontSize: '18px', fontWeight: 800, color: '#1e40af' }}>
                    {inr(Math.round(pipelineData.activePipelineTotal))}
                  </div>
                </div>
                <div style={{ width: '1px', height: '28px', background: '#cbd5e1' }} />
                <div>
                  <span style={{ fontSize: '10.5px', fontWeight: 700, color: '#64748b', textTransform: 'uppercase', letterSpacing: '0.04em' }}>
                    Won Deals Total
                  </span>
                  <div style={{ fontSize: '18px', fontWeight: 800, color: '#16a34a' }}>
                    {inr(Math.round(pipelineData.totals.won))}
                  </div>
                </div>
                <div style={{ width: '1px', height: '28px', background: '#cbd5e1' }} />
                <div>
                  <span style={{ fontSize: '10.5px', fontWeight: 700, color: '#64748b', textTransform: 'uppercase', letterSpacing: '0.04em' }}>
                    Total Opportunities
                  </span>
                  <div style={{ fontSize: '18px', fontWeight: 800, color: '#0f172a' }}>
                    {filteredProjects.length} Deals
                  </div>
                </div>
              </div>
              <div className="tiny text-muted" style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
                <span>💡 Drag &amp; drop cards between columns or use the stage selector to advance deals</span>
              </div>
            </div>

            {/* 5-Column Pipeline Kanban */}
            <div className="pipeline-board-container">
              {PIPELINE_COLUMNS.map((col) => {
                const colProjects = pipelineData.groups[col.id] || []
                const colTotal = pipelineData.totals[col.id as keyof typeof pipelineData.totals] || 0
                const isDragOver = dragOverColumn === col.id

                return (
                  <div
                    key={col.id}
                    className={'pipeline-column ' + (isDragOver ? 'drag-over' : '')}
                    onDragOver={(e) => {
                      e.preventDefault()
                      e.dataTransfer.dropEffect = 'move'
                      if (dragOverColumn !== col.id) setDragOverColumn(col.id)
                    }}
                    onDragLeave={(e) => {
                      if (dragOverColumn === col.id) setDragOverColumn(null)
                    }}
                    onDrop={(e) => {
                      e.preventDefault()
                      setDragOverColumn(null)
                      const pId = e.dataTransfer.getData('text/plain') || draggedProjectId
                      if (pId) handleUpdateProjectStatus(pId, col.id)
                      setDraggedProjectId(null)
                    }}
                  >
                    {/* Column Header */}
                    <div className="pipeline-column-header">
                      <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
                        <div style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                          <span>{col.icon}</span>
                          <strong style={{ fontSize: '12.5px', color: col.color }}>{col.title}</strong>
                        </div>
                        <span
                          className="chip sm font-bold"
                          style={{
                            fontSize: '10px',
                            background: col.bg,
                            color: col.color,
                            borderColor: col.border,
                          }}
                        >
                          {colProjects.length}
                        </span>
                      </div>
                      <div style={{ fontSize: '11px', fontWeight: 700, color: '#047857' }}>
                        {inr(Math.round(colTotal))}
                      </div>
                    </div>

                    {/* Column Cards */}
                    <div className="pipeline-cards-list">
                      {colProjects.length === 0 ? (
                        <div
                          className="tiny muted text-center"
                          style={{ padding: '30px 10px', border: '1.5px dashed #e2e8f0', borderRadius: '6px' }}
                        >
                          No quotations in this stage
                        </div>
                      ) : (
                        colProjects.map((p) => {
                          const buyer = p.summary?.client_info?.buyer_name || p.summary?.client_info?.name || ''
                          const city = p.summary?.client_info?.city || p.summary?.client_info?.state || ''
                          const bodies = p.summary?.total_combined_bodies || p.summary?.bodies
                          const rev = parseRevision(p.name).currentRev

                          return (
                            <div
                              key={p.id}
                              className="pipeline-card"
                              draggable={true}
                              onDragStart={(e) => {
                                e.dataTransfer.setData('text/plain', p.id)
                                setDraggedProjectId(p.id)
                              }}
                              onDragEnd={() => setDraggedProjectId(null)}
                              onClick={() => onOpenProject(p.id)}
                              title="Click to open quotation in Calculator (Drag to move stage)"
                            >
                              {/* Card Header: QTN Name & Stage Selector */}
                              <div style={{ display: 'flex', alignItems: 'flex-start', justifyContent: 'space-between', gap: 6 }}>
                                <div>
                                  <strong style={{ fontSize: '13px', color: '#1e40af', display: 'flex', alignItems: 'center', gap: 5 }}>
                                    📄 {p.name}
                                  </strong>
                                  {rev && (
                                    <span className="version-pill" style={{ fontSize: 9.5, padding: '0 4px', marginTop: 2 }}>
                                      Rev {rev}
                                    </span>
                                  )}
                                </div>
                                <select
                                  value={col.id}
                                  onChange={(e) => handleUpdateProjectStatus(p.id, e.target.value as ProjectStatus)}
                                  onClick={(e) => e.stopPropagation()}
                                  className="pipeline-status-select"
                                  title="Change deal stage"
                                >
                                  <option value="draft">📝 Draft</option>
                                  <option value="quoted">📤 Quoted</option>
                                  <option value="negotiating">🤝 Negotiating</option>
                                  <option value="won">🎉 Won</option>
                                  <option value="loss">✕ Lost</option>
                                </select>
                              </div>

                              {/* Client & Specs */}
                              <div style={{ fontSize: '11.5px', color: '#334155' }}>
                                <div style={{ fontWeight: 600, color: '#0f172a' }}>
                                  {buyer || '(Client not specified)'}
                                </div>
                                {city && <div className="tiny muted">{city}</div>}
                                {bodies ? (
                                  <div className="tiny" style={{ color: '#0369a1', fontWeight: 600, marginTop: 2 }}>
                                    📦 {bodies} Bodies
                                  </div>
                                ) : null}
                              </div>

                              {/* Quotation Amount & Badges */}
                              <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', borderTop: '1px dashed #e2e8f0', paddingTop: 6 }}>
                                <strong style={{ fontSize: '13.5px', color: '#16a34a' }}>
                                  {p.total ? inr(Number(p.total)) : '—'}
                                </strong>
                                <div className="tiny text-muted">
                                  @{p.created_by || 'staff'}
                                </div>
                              </div>

                              {/* Active Reminders / Badges */}
                              {p.summary?.reminder && !p.summary.reminder.dismissed && (
                                <div
                                  className="chip warning sm"
                                  style={{ fontSize: 9.5, padding: '1px 5px', display: 'flex', alignItems: 'center', gap: 3 }}
                                >
                                  🔔 {new Date(p.summary.reminder.date).toLocaleDateString('en-IN', { month: 'short', day: 'numeric' })}, {new Date(p.summary.reminder.date).toLocaleTimeString('en-IN', { hour: '2-digit', minute: '2-digit' })}
                                </div>
                              )}
                            </div>
                          )
                        })
                      )}
                    </div>
                  </div>
                )
              })}
            </div>
          </div>
        ) : (
          <div className="table-scroll-wrap">
            <table className="tbl full-table">
            <thead>
              <tr>
                <th style={{ width: '40px', textAlign: 'center' }}>
                  <input
                    type="checkbox"
                    aria-label="Select all visible quotations"
                    checked={allVisibleSelected}
                    ref={(el) => {
                      if (el) el.indeterminate = someVisibleSelected && !allVisibleSelected
                    }}
                    onChange={handleToggleSelectAll}
                  />
                </th>
                <th>{viewMode === 'folders' ? 'Project Folder / Revision File' : 'Project Name'}</th>
                <th>Client</th>
                <th>Status</th>
                <th>Template Version</th>
                <th className="r">Quotation Total</th>
                <th>Owner / Creator</th>
                <th>Last Updated</th>
                <th className="r">Actions</th>
              </tr>
            </thead>
            <tbody>
              {viewMode === 'folders' ? (
                projectFolders.map((folder) => {
                  const isOpen = Boolean(openFolders[folder.id])
                  const latest = folder.latestRevision
                  const hasNewer = (org.current_template_version || 1) > latest.template_version
                  const folderAllSelected = folder.revisions.length > 0 && folder.revisions.every((r) => selectedIds.has(r.id))
                  const folderSomeSelected = folder.revisions.some((r) => selectedIds.has(r.id))

                  return (
                    <Fragment key={folder.id}>
                      {/* Project Folder Header Row (Click anywhere to open/close) */}
                      <tr
                        className={`folder-header-row ${folder.hasActive ? 'selected-row' : ''}`}
                        onClick={() => toggleFolder(folder.id)}
                        style={{ cursor: 'pointer' }}
                      >
                        <td style={{ width: '40px', textAlign: 'center' }} onClick={(e) => e.stopPropagation()}>
                          <input
                            type="checkbox"
                            aria-label={`Select folder ${folder.folderName}`}
                            checked={folderAllSelected}
                            ref={(el) => {
                              if (el) el.indeterminate = folderSomeSelected && !folderAllSelected
                            }}
                            onChange={(e) => handleToggleFolderSelect(folder, e.target.checked)}
                          />
                        </td>
                        <td>
                          <div className="row gap" style={{ flexWrap: 'wrap' }}>
                            <button
                              className="folder-toggle-btn"
                              onClick={(e) => {
                                e.stopPropagation()
                                toggleFolder(folder.id)
                              }}
                              title={isOpen ? 'Click to close folder' : 'Click to open folder'}
                            >
                              <span className="folder-icon">{Icon.folder}</span>
                              <strong>{folder.folderName}</strong>
                              <span style={{ color: 'var(--muted)', display: 'inline-flex' }}>
                                {isOpen ? Icon.chevronDown : Icon.chevronRight}
                              </span>
                            </button>
                            <span className="folder-count-badge">
                              {folder.totalCount} File{folder.totalCount > 1 ? 's' : ''} ({folder.revisions.map((r) => parseRevision(r.name).currentRev).join(', ')})
                            </span>
                            {folder.hasActive && <span className="chip ok">Active Folder</span>}
                            {folder.revisions.some((r) => {
                              const shared = (r.summary?.shared_with || (r as any).shared_with) as string[] | undefined
                              return (
                                Array.isArray(shared) &&
                                (shared.includes(session.username) || (session.user_id && shared.includes(session.user_id))) &&
                                r.created_by !== session.username
                              )
                            }) && (
                              <span className="chip purple sm" title="Contains quotations shared with you">Shared with you</span>
                            )}
                          </div>
                        </td>
                        <td>
                          {(() => {
                            const clientInfo = latest.summary?.client_info || folder.revisions.find((r) => r.summary?.client_info?.buyer_name)?.summary?.client_info
                            const folderClient = clientInfo?.buyer_name || ''
                            const qtnNo = clientInfo?.qtn_no
                            const city = clientInfo?.buyer_city || ''
                            const address = clientInfo?.buyer_address || ''
                            if (!folderClient) return <span className="muted tiny">—</span>
                            return (
                              <div style={{ display: 'flex', flexDirection: 'column', gap: 2 }}>
                                <strong style={{ color: '#0f172a', fontSize: '12px' }}>{folderClient}</strong>
                                <div style={{ display: 'flex', gap: 6, flexWrap: 'wrap', alignItems: 'center' }}>
                                  {qtnNo && <span className="tiny font-mono text-muted" style={{ fontSize: '10.5px' }}>{qtnNo}</span>}
                                  {(city || address) && (
                                    <span className="tiny text-muted" style={{ fontSize: '10.5px', color: '#64748b', display: 'inline-flex', alignItems: 'center', gap: '2px' }} title={address}>
                                      📍 {city || (address.length > 25 ? address.slice(0, 25) + '…' : address)}
                                    </span>
                                  )}
                                </div>
                              </div>
                            )
                          })()}
                        </td>
                        <td>
                          {(() => {
                            const isWon = !latest.is_archived && (latest.status === 'won' || latest.status === 'approved')
                            const isLoss = latest.is_archived || latest.status === 'loss' || latest.status === 'archived' || latest.status === 'lost'
                            const isNeg = latest.status === 'negotiating'
                            const label = isWon ? 'Won' : isLoss ? 'Loss' : isNeg ? 'Negotiating' : 'Quoted'
                            const cls = isWon ? 'won' : isLoss ? 'loss' : isNeg ? 'negotiating' : 'quoted'
                            return <span className={`status-pill status-${cls}`}>{label}</span>
                          })()}
                          {latest.summary?.discount_request?.status === 'pending' && (
                            <span
                              className="chip warning sm"
                              style={{ marginLeft: '4px', fontSize: '11px', padding: '2px 6px' }}
                              title={`Staff @${latest.summary.discount_request.requested_by} requested ${latest.summary.discount_request.requested_pct}% discount`}
                            >
                              ⏳ {latest.summary.discount_request.requested_pct}% Disc Req
                            </span>
                          )}
                          {latest.summary?.reminder && !latest.summary.reminder.dismissed && (
                            <span
                              className="chip warning sm"
                              style={{ marginLeft: '4px', fontSize: '10px', padding: '1px 5px', gap: '3px', fontWeight: 600 }}
                              title={`Alarm Reminder: ${latest.summary.reminder.note || 'Follow up'} (Due: ${new Date(latest.summary.reminder.date).toLocaleDateString()} at ${new Date(latest.summary.reminder.date).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })})`}
                            >
                              🔔 {new Date(latest.summary.reminder.date).toLocaleDateString('en-IN', { month: 'short', day: 'numeric' })}, {new Date(latest.summary.reminder.date).toLocaleTimeString('en-IN', { hour: '2-digit', minute: '2-digit' })}
                            </span>
                          )}
                        </td>
                        <td>
                          <div className="template-badge-wrap">
                            <span className="version-pill">v{latest.template_version}</span>
                            {hasNewer && (
                              <span className="upgrade-hint" title={`Template v${org.current_template_version} available`}>
                                ↑ v{org.current_template_version}
                              </span>
                            )}
                          </div>
                        </td>
                        <td className="r font-bold text-green">
                          {latest.total ? inr(Number(latest.total)) : '—'}
                        </td>
                        <td className="muted tiny">
                          {latest.created_by || 'Unknown'}
                        </td>
                        <td className="muted tiny">{timeAgo(latest.updated_at)}</td>
                        <td className="r">
                          <div className="folder-actions" onClick={(e) => e.stopPropagation()}>
                            <button
                              className="btn primary sm folder-open-btn"
                              onClick={() => onOpenProject(latest.id)}
                              title={`Open latest revision (${folder.latestRevChar})`}
                            >
                              Open Latest ({folder.latestRevChar})
                            </button>
                            <button
                              className="btn ghost sm"
                              title={`Revise to Rev ${folder.nextRevChar}`}
                              onClick={() => {
                                setReviseProj(latest)
                                setRevName(folder.nextName)
                              }}
                            >
                              {Icon.refresh} Revise ({folder.nextRevChar})
                            </button>
                          </div>
                        </td>
                      </tr>

                      {/* Contained Revision Files (Only opened when folder is clicked!) */}
                      {isOpen && folder.revisions.map((file) => renderProjectRow(file, true))}
                    </Fragment>
                  )
                })
              ) : (
                filteredProjects.map((p) => renderProjectRow(p, false))
              )}
            </tbody>
          </table>
          </div>
        )}
      </div>

      {/* Revise Quotation Dialog */}
      {reviseProj && (
        <div className="overlay" onMouseDown={() => setReviseProj(null)}>
          <div className="dialog" role="dialog" onMouseDown={(e) => e.stopPropagation()}>
            <div className="dialog-header">
              {Icon.refresh}
              <div>
                <h3 style={{ margin: 0 }}>Create Revision {parseRevision(reviseProj.name).nextRev}</h3>
                <span className="tiny text-muted">From “{reviseProj.name}”</span>
              </div>
            </div>
            <p className="tiny muted mt-2" style={{ lineHeight: 1.5 }}>
              This will create a new quotation snapshot <strong>Rev {parseRevision(reviseProj.name).nextRev}</strong> with all current inputs, rates, profit margins, and discounts intact.
              The prior quotation (<strong>Rev {parseRevision(reviseProj.name).currentRev}</strong>) will remain preserved in your projects list.
            </p>
            <div className="form mt-3">
              <label>
                <span>Revision Quotation Name</span>
                <input
                  type="text"
                  autoFocus
                  className="modal-input"
                  value={revName}
                  onChange={(e) => setRevName(e.target.value)}
                  onKeyDown={(e) => e.key === 'Enter' && handleRevise()}
                />
              </label>
            </div>
            <div className="dialog-actions mt-4">
              <button className="btn ghost" onClick={() => setReviseProj(null)}>
                Cancel
              </button>
              <button className="btn primary" onClick={handleRevise} disabled={!revName.trim()}>
                {Icon.refresh} Create Rev {parseRevision(reviseProj.name).nextRev}
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Rename Dialog */}
      {renameProj && (
        <div className="overlay" onMouseDown={() => setRenameProj(null)}>
          <div className="dialog" role="dialog" onMouseDown={(e) => e.stopPropagation()}>
            <h3>Rename Project</h3>
            <p className="tiny muted">Enter a new title for this project</p>
            <input
              type="text"
              autoFocus
              className="modal-input"
              value={newName}
              onChange={(e) => setNewName(e.target.value)}
              onKeyDown={(e) => e.key === 'Enter' && handleRename()}
            />
            <div className="dialog-actions mt-4">
              <button className="btn ghost" onClick={() => setRenameProj(null)}>
                Cancel
              </button>
              <button className="btn primary" onClick={handleRename}>
                Save Name
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Duplicate Dialog */}
      {duplicateProj && (
        <div className="overlay" onMouseDown={() => setDuplicateProj(null)}>
          <div className="dialog" role="dialog" onMouseDown={(e) => e.stopPropagation()}>
            <h3>Duplicate Project</h3>
            <p className="tiny muted">
              Creates an independent copy of <strong>{duplicateProj.name}</strong>. Changes to the copy will never affect the original project.
            </p>
            <div className="form">
              <label>
                <span>New Project Name</span>
                <input
                  type="text"
                  autoFocus
                  value={dupName}
                  onChange={(e) => setDupName(e.target.value)}
                  onKeyDown={(e) => e.key === 'Enter' && handleDuplicate()}
                />
              </label>

              <div className="duplicate-template-choice">
                <span className="tiny font-bold text-muted">Template Version for Duplicate:</span>
                <label className="radio-label">
                  <input
                    type="radio"
                    name="dup_version"
                    checked={!dupUseLatest}
                    onChange={() => setDupUseLatest(false)}
                  />
                  <div>
                    <strong>Retain original Template (v{duplicateProj.template_version})</strong>
                    <span className="tiny muted block">Keep identical formula and rule definitions as the source</span>
                  </div>
                </label>
                <label className="radio-label">
                  <input
                    type="radio"
                    name="dup_version"
                    checked={dupUseLatest}
                    onChange={() => setDupUseLatest(true)}
                  />
                  <div>
                    <strong>Apply Organisation Default (v{org.current_template_version || 1})</strong>
                    <span className="tiny muted block">Snapshots the currently published default template rules</span>
                  </div>
                </label>
              </div>
            </div>

            <div className="dialog-actions mt-4">
              <button className="btn ghost" onClick={() => setDuplicateProj(null)}>
                Cancel
              </button>
              <button className="btn primary" onClick={handleDuplicate}>
                Create Duplicate
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Share Quotation Dialog */}
      {shareProj && (
        <div className="overlay" onMouseDown={() => setShareProj(null)}>
          <div className="dialog" role="dialog" onMouseDown={(e) => e.stopPropagation()} style={{ maxWidth: 440 }}>
            <div className="dialog-header">
              {Icon.users}
              <div>
                <h3 style={{ margin: 0 }}>Share Quotation</h3>
                <span className="tiny text-muted">“{shareProj.name}”</span>
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
                    const currentShared = Array.isArray(shareProj.summary?.shared_with)
                      ? shareProj.summary.shared_with
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
              <button className="btn primary" onClick={() => setShareProj(null)}>
                Done
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Single Quotation Permanent Deletion Dialog (Requires Exact Name) */}
      {singleDeleteProj && (
        <div className="overlay" onMouseDown={() => !deletingSingle && setSingleDeleteProj(null)}>
          <div className="dialog danger-dialog" role="dialog" onMouseDown={(e) => e.stopPropagation()} style={{ maxWidth: 440 }}>
            <div className="dialog-header">
              <span style={{ color: '#dc2626', display: 'flex', alignItems: 'center' }}>{Icon.trash}</span>
              <div>
                <h3 style={{ margin: 0, color: '#dc2626' }}>Delete Quotation</h3>
                <span className="tiny text-muted">{singleDeleteProj.name}</span>
              </div>
            </div>
            <p className="tiny muted mt-2" style={{ lineHeight: 1.5 }}>
              This will permanently delete quotation <strong>“{singleDeleteProj.name}”</strong> and its calculation history. This action <strong>cannot be undone</strong>.
            </p>
            <div className="form mt-3">
              <label>
                <span style={{ fontSize: '12px' }}>
                  To confirm, type <strong>{singleDeleteProj.name}</strong> in the box below:
                </span>
                <div className="confirm-name-target">{singleDeleteProj.name}</div>
                <input
                  type="text"
                  autoFocus
                  className="modal-input"
                  placeholder="Type quotation name to confirm"
                  value={confirmNameInput}
                  onChange={(e) => setConfirmNameInput(e.target.value)}
                  onKeyDown={(e) => {
                    if (e.key === 'Enter' && confirmNameInput.trim() === singleDeleteProj.name.trim()) {
                      handleConfirmSingleDelete()
                    }
                  }}
                />
              </label>
            </div>
            <div className="dialog-actions mt-4">
              <button
                type="button"
                className="btn ghost sm"
                onClick={() => setSingleDeleteProj(null)}
                disabled={deletingSingle}
              >
                Cancel
              </button>
              <button
                type="button"
                className="btn danger sm"
                disabled={confirmNameInput.trim() !== singleDeleteProj.name.trim() || deletingSingle}
                onClick={handleConfirmSingleDelete}
                style={{ display: 'flex', alignItems: 'center', gap: '6px' }}
              >
                {deletingSingle ? <Spinner size={14} /> : Icon.trash} Delete Permanently
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Multiple Quotations Batch Deletion Dialog (Requires Account Password) */}
      {showBatchDeleteModal && (
        <div className="overlay" onMouseDown={() => !deletingBatch && setShowBatchDeleteModal(false)}>
          <div className="dialog danger-dialog" role="dialog" onMouseDown={(e) => e.stopPropagation()} style={{ maxWidth: 460 }}>
            <div className="dialog-header">
              <span style={{ color: '#dc2626', display: 'flex', alignItems: 'center' }}>{Icon.trash}</span>
              <div>
                <h3 style={{ margin: 0, color: '#dc2626' }}>Delete {selectedIds.size} Quotations</h3>
                <span className="tiny text-muted">Batch Permanent Deletion</span>
              </div>
            </div>
            <p className="tiny muted mt-2" style={{ lineHeight: 1.5 }}>
              You are about to permanently delete <strong>{selectedIds.size}</strong> quotations and their entire calculation histories. This action <strong>cannot be undone</strong>.
            </p>
            <div
              className="batch-delete-preview-box mt-2"
              style={{
                maxHeight: '120px',
                overflowY: 'auto',
                background: '#f8fafc',
                border: '1px solid #e2e8f0',
                borderRadius: '6px',
                padding: '8px 12px',
              }}
            >
              <ul style={{ margin: 0, paddingLeft: '18px', fontSize: '12px', color: '#475569' }}>
                {Array.from(selectedIds).map((id) => {
                  const p = projects.find((proj) => proj.id === id)
                  const buyer = p?.summary?.client_info?.buyer_name
                  return (
                    <li key={id} style={{ marginBottom: '2px' }}>
                      <strong>{p?.name || id}</strong>
                      {buyer ? <span className="tiny muted"> — {buyer}</span> : null}
                    </li>
                  )
                })}
              </ul>
            </div>
            <div className="form mt-3">
              <label>
                <span style={{ fontSize: '12px', fontWeight: 600 }}>Enter your account password to confirm deletion:</span>
                <input
                  type="password"
                  autoFocus
                  className="modal-input"
                  placeholder="Enter your account password"
                  value={batchDeletePassword}
                  onChange={(e) => {
                    setBatchDeletePassword(e.target.value)
                    if (batchDeleteError) setBatchDeleteError('')
                  }}
                  onKeyDown={(e) => {
                    if (e.key === 'Enter' && batchDeletePassword && !deletingBatch) {
                      handleConfirmBatchDelete()
                    }
                  }}
                />
              </label>
              {batchDeleteError && (
                <div className="tiny mt-1" style={{ color: '#dc2626', fontWeight: 600 }}>
                  {batchDeleteError}
                </div>
              )}
            </div>
            <div className="dialog-actions mt-4">
              <button
                type="button"
                className="btn ghost sm"
                onClick={() => setShowBatchDeleteModal(false)}
                disabled={deletingBatch}
              >
                Cancel
              </button>
              <button
                type="button"
                className="btn danger sm"
                disabled={!batchDeletePassword || deletingBatch}
                onClick={handleConfirmBatchDelete}
                style={{ display: 'flex', alignItems: 'center', gap: '6px' }}
              >
                {deletingBatch ? <Spinner size={14} /> : Icon.trash} Confirm &amp; Delete All ({selectedIds.size})
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  )
}
