import React, { useEffect, useMemo, useState } from 'react'
import { Member, Org, Project, Session, api } from '../../lib/api'
import {
  SavedClient,
  fetchClientsFromSupabase,
  getSavedClients,
} from '../../lib/clientDirectory'
import { inr, timeAgo } from '../../lib/format'
import { Icon, Spinner, useUI } from '../ui'

interface Props {
  session: Session
  org: Org
  onOpenProject: (projectId: string) => void
  onCreateQuotationWithClient?: (client: SavedClient) => void
  onNavigate?: (tab: string) => void
}

interface StaffQuotationCount {
  staffUsername: string
  displayName: string
  role: string
  count: number
  totalValue: number
}

interface ClientGroup {
  clientId: string
  clientName: string
  address: string
  gst: string
  contactPerson: string
  phone: string
  email: string
  rawClient?: SavedClient
  projects: Project[]
  staffList: StaffQuotationCount[]
  totalQuotes: number
  activeQuotes: number
  wonQuotes: number
  lostQuotes: number
  totalValue: number
  activeValue: number
  wonValue: number
}

function getPageNumbers(current: number, total: number): (number | string)[] {
  if (total <= 7) {
    return Array.from({ length: total }, (_, i) => i + 1)
  }
  if (current <= 4) {
    return [1, 2, 3, 4, 5, '...', total]
  }
  if (current >= total - 3) {
    return [1, '...', total - 4, total - 3, total - 2, total - 1, total]
  }
  return [1, '...', current - 1, current, current + 1, '...', total]
}

export function ClientRelationsPage({
  session,
  org,
  onOpenProject,
  onCreateQuotationWithClient,
  onNavigate,
}: Props) {
  const { toast } = useUI()
  const [loading, setLoading] = useState(true)
  const [projects, setProjects] = useState<Project[]>([])
  const [members, setMembers] = useState<Member[]>([])
  const [savedClients, setSavedClients] = useState<SavedClient[]>(() => getSavedClients(org.id))

  // Filters & Pagination
  const [search, setSearch] = useState('')
  const [staffFilter, setStaffFilter] = useState<string>('all')
  const [statusFilter, setStatusFilter] = useState<'all' | 'active' | 'won' | 'lost'>('all')
  const [sortBy, setSortBy] = useState<'quotes_desc' | 'value_desc' | 'name_asc'>('quotes_desc')
  const [expandedClients, setExpandedClients] = useState<Set<string>>(new Set())
  const [currentPage, setCurrentPage] = useState(1)
  const PAGE_SIZE = 12

  const loadData = async () => {
    try {
      setLoading(true)
      const [projList, memberList] = await Promise.all([
        api.projectList(session.token, org.id, true),
        api.orgMembers(session.token, org.id).catch(() => []),
        fetchClientsFromSupabase(session.token, org.id).catch(() => []),
      ])
      setProjects(Array.isArray(projList) ? projList : [])
      setMembers(Array.isArray(memberList) ? memberList : [])
      setSavedClients(getSavedClients(org.id))
    } catch (e) {
      toast((e as Error).message, 'err')
    } finally {
      setLoading(false)
    }
  }

  useEffect(() => {
    loadData()
  }, [org.id])

  // Reset pagination on filter change
  useEffect(() => {
    setCurrentPage(1)
  }, [search, staffFilter, statusFilter, sortBy])

  // Member map for quick staff lookup
  const memberMap = useMemo(() => {
    const map = new Map<string, Member>()
    for (const m of members) {
      map.set(m.username.toLowerCase(), m)
      const extraName = (m as any).name
      if (extraName) map.set(String(extraName).toLowerCase(), m)
    }
    return map
  }, [members])

  // Group quotations by client
  const clientGroups = useMemo(() => {
    const groups = new Map<string, ClientGroup>()

    // First register all saved clients from directory
    for (const sc of savedClients) {
      const key = sc.name.trim().toLowerCase()
      if (!groups.has(key)) {
        groups.set(key, {
          clientId: sc.id,
          clientName: sc.name.trim(),
          address: sc.address || '',
          gst: sc.gst || '',
          contactPerson: sc.kind_att || sc.contacts?.[0]?.name || '',
          phone: sc.contact || sc.contacts?.[0]?.phone || '',
          email: sc.email || sc.contacts?.[0]?.email || '',
          rawClient: sc,
          projects: [],
          staffList: [],
          totalQuotes: 0,
          activeQuotes: 0,
          wonQuotes: 0,
          lostQuotes: 0,
          totalValue: 0,
          activeValue: 0,
          wonValue: 0,
        })
      }
    }

    // Now map every project to its client
    for (const p of projects) {
      const clientInfo = p.summary?.client_info
      const buyerName =
        clientInfo?.buyer_name ||
        p.inputs?.['buyer_name'] ||
        p.inputs?.['client_name'] ||
        'Unassigned / Walk-in Client'

      const key = buyerName.trim().toLowerCase()
      let group = groups.get(key)
      if (!group) {
        group = {
          clientId: p.summary?.client_id || `temp-${key}`,
          clientName: buyerName.trim(),
          address: clientInfo?.buyer_address || '',
          gst: clientInfo?.buyer_gst || '',
          contactPerson: clientInfo?.kind_att || '',
          phone: clientInfo?.buyer_contact || '',
          email: clientInfo?.buyer_email || '',
          projects: [],
          staffList: [],
          totalQuotes: 0,
          activeQuotes: 0,
          wonQuotes: 0,
          lostQuotes: 0,
          totalValue: 0,
          activeValue: 0,
          wonValue: 0,
        }
        groups.set(key, group)
      } else {
        // Fill in metadata if missing
        if (!group.address && clientInfo?.buyer_address) group.address = clientInfo.buyer_address
        if (!group.gst && clientInfo?.buyer_gst) group.gst = clientInfo.buyer_gst
        if (!group.contactPerson && clientInfo?.kind_att) group.contactPerson = clientInfo.kind_att
        if (!group.phone && clientInfo?.buyer_contact) group.phone = clientInfo.buyer_contact
        if (!group.email && clientInfo?.buyer_email) group.email = clientInfo.buyer_email
      }

      group.projects.push(p)
      const val = Number(p.total) || 0
      group.totalQuotes++
      group.totalValue += val

      const isLoss = p.is_archived || p.status === 'loss' || p.status === 'archived' || p.status === 'lost'
      const isWon = !p.is_archived && (p.status === 'won' || p.status === 'approved')

      if (isWon) {
        group.wonQuotes++
        group.wonValue += val
      } else if (isLoss) {
        group.lostQuotes++
      } else {
        group.activeQuotes++
        group.activeValue += val
      }
    }

    // For each client group, compute staff breakdown
    groups.forEach((group) => {
      const staffMap = new Map<string, StaffQuotationCount>()
      for (const p of group.projects) {
        const creatorKey = p.created_by || 'Unassigned'
        const matchedMember = memberMap.get(creatorKey.toLowerCase())
        const staffKey = matchedMember?.username || creatorKey
        const displayName = (matchedMember as any)?.name || staffKey
        const role = matchedMember?.role || 'member'

        let entry = staffMap.get(staffKey)
        if (!entry) {
          entry = {
            staffUsername: staffKey,
            displayName,
            role,
            count: 0,
            totalValue: 0,
          }
          staffMap.set(staffKey, entry)
        }
        entry.count++
        entry.totalValue += Number(p.total) || 0
      }

      group.staffList = Array.from(staffMap.values()).sort((a, b) => b.count - a.count)
    })

    return Array.from(groups.values())
  }, [projects, savedClients, memberMap])

  // Filtered and Sorted Client Groups
  const filteredGroups = useMemo(() => {
    return clientGroups.filter((g) => {
      // Staff filter
      if (staffFilter !== 'all') {
        const hasStaff = g.staffList.some(
          (s) => s.staffUsername.toLowerCase() === staffFilter.toLowerCase()
        )
        if (!hasStaff) return false
      }

      // Status filter
      if (statusFilter === 'active' && g.activeQuotes === 0) return false
      if (statusFilter === 'won' && g.wonQuotes === 0) return false
      if (statusFilter === 'lost' && g.lostQuotes === 0) return false

      // Search filter
      if (search.trim()) {
        const q = search.toLowerCase().trim()
        const matchName = g.clientName.toLowerCase().includes(q)
        const matchGst = g.gst.toLowerCase().includes(q)
        const matchContact = g.contactPerson.toLowerCase().includes(q)
        const matchPhone = g.phone.toLowerCase().includes(q)
        const matchStaff = g.staffList.some(
          (s) =>
            s.staffUsername.toLowerCase().includes(q) ||
            s.displayName.toLowerCase().includes(q)
        )
        const matchProject = g.projects.some((p) => p.name.toLowerCase().includes(q))
        return matchName || matchGst || matchContact || matchPhone || matchStaff || matchProject
      }

      return true
    }).sort((a, b) => {
      if (sortBy === 'quotes_desc') return b.totalQuotes - a.totalQuotes
      if (sortBy === 'value_desc') return b.totalValue - a.totalValue
      return a.clientName.localeCompare(b.clientName)
    })
  }, [clientGroups, staffFilter, statusFilter, search, sortBy])

  // Pagination calculations
  const totalPages = Math.max(1, Math.ceil(filteredGroups.length / PAGE_SIZE))
  const activePage = Math.min(currentPage, totalPages)
  const paginatedGroups = useMemo(() => {
    const start = (activePage - 1) * PAGE_SIZE
    return filteredGroups.slice(start, start + PAGE_SIZE)
  }, [filteredGroups, activePage])

  // Summary Metrics
  const summaryMetrics = useMemo(() => {
    let totalClientsWithQuotes = 0
    let totalQuotes = 0
    let totalPipeline = 0
    let totalWon = 0
    const activeStaffSet = new Set<string>()

    for (const g of clientGroups) {
      if (g.totalQuotes > 0) totalClientsWithQuotes++
      totalQuotes += g.totalQuotes
      totalPipeline += g.activeValue
      totalWon += g.wonValue
      for (const s of g.staffList) {
        if (s.staffUsername !== 'Unassigned') {
          activeStaffSet.add(s.staffUsername)
        }
      }
    }

    return {
      totalClients: clientGroups.length,
      totalClientsWithQuotes,
      totalQuotes,
      totalPipeline,
      totalWon,
      activeStaffCount: activeStaffSet.size,
    }
  }, [clientGroups])

  // Distinct staff members across all clients for dropdown
  const allStaffOptions = useMemo(() => {
    const map = new Map<string, string>()
    for (const m of members) {
      map.set(m.username, (m as any)?.name || m.username)
    }
    for (const g of clientGroups) {
      for (const s of g.staffList) {
        if (!map.has(s.staffUsername)) {
          map.set(s.staffUsername, s.displayName)
        }
      }
    }
    return Array.from(map.entries()).sort((a, b) => a[1].localeCompare(b[1]))
  }, [members, clientGroups])

  const toggleClientExpand = (key: string) => {
    setExpandedClients((prev) => {
      const next = new Set(prev)
      if (next.has(key)) next.delete(key)
      else next.add(key)
      return next
    })
  }

  const handleExpandAll = () => {
    if (expandedClients.size === filteredGroups.length) {
      setExpandedClients(new Set())
    } else {
      setExpandedClients(new Set(filteredGroups.map((g) => g.clientName.toLowerCase())))
    }
  }

  return (
    <div className="client-relations-page">
      {/* Top Header Card */}
      <div className="cr-header-card">
        <div>
          <div className="cr-header-title-row">
            <span className="cr-header-icon">👥</span>
            <div>
              <h2 className="cr-header-title">Client &amp; Staff Relations</h2>
              <p className="cr-header-subtitle">
                Account distribution matrix: manage client relationships, track quotation assignments across team members, and view account pipelines.
              </p>
            </div>
          </div>
        </div>

        <div className="row gap" style={{ alignItems: 'center' }}>
          <button
            type="button"
            className="btn ghost sm"
            onClick={handleExpandAll}
            title={expandedClients.size === filteredGroups.length ? 'Collapse All Accounts' : 'Expand All Accounts'}
          >
            {expandedClients.size === filteredGroups.length ? '▴ Collapse All' : '▾ Expand All'}
          </button>
          {onNavigate && (
            <button
              type="button"
              className="btn ghost sm"
              onClick={() => onNavigate('clients')}
              title="Open Client Directory"
            >
              {Icon.building} Client Directory
            </button>
          )}
          <button type="button" className="btn ghost sm" onClick={loadData} title="Refresh relations matrix">
            {Icon.refresh} Refresh
          </button>
        </div>
      </div>

      {/* KPI Overview Cards */}
      <div className="cr-kpi-grid">
        <div className="cr-kpi-card">
          <div className="cr-kpi-top">
            <span className="cr-kpi-label">Total Accounts</span>
            <div className="cr-kpi-icon-wrap" style={{ background: '#f1f5f9', color: '#0f172a' }}>
              🏢
            </div>
          </div>
          <div className="cr-kpi-value">{summaryMetrics.totalClients}</div>
          <div className="cr-kpi-desc">
            {summaryMetrics.totalClientsWithQuotes} clients with quotation history
          </div>
        </div>

        <div className="cr-kpi-card">
          <div className="cr-kpi-top">
            <span className="cr-kpi-label">Total Quotations</span>
            <div className="cr-kpi-icon-wrap" style={{ background: '#eff6ff', color: '#2563eb' }}>
              📑
            </div>
          </div>
          <div className="cr-kpi-value" style={{ color: '#2563eb' }}>
            {summaryMetrics.totalQuotes}
          </div>
          <div className="cr-kpi-desc">
            Handled across {summaryMetrics.activeStaffCount} active staff members
          </div>
        </div>

        <div className="cr-kpi-card">
          <div className="cr-kpi-top">
            <span className="cr-kpi-label">Active Pipeline</span>
            <div className="cr-kpi-icon-wrap" style={{ background: '#fef3c7', color: '#d97706' }}>
              💼
            </div>
          </div>
          <div className="cr-kpi-value" style={{ color: '#b45309' }}>
            {inr(summaryMetrics.totalPipeline)}
          </div>
          <div className="cr-kpi-desc">Quoted &amp; negotiating quotation value</div>
        </div>

        <div className="cr-kpi-card">
          <div className="cr-kpi-top">
            <span className="cr-kpi-label">Won Revenue</span>
            <div className="cr-kpi-icon-wrap" style={{ background: '#ecfdf5', color: '#059669' }}>
              🏆
            </div>
          </div>
          <div className="cr-kpi-value" style={{ color: '#059669' }}>
            {inr(summaryMetrics.totalWon)}
          </div>
          <div className="cr-kpi-desc">Successfully converted business</div>
        </div>
      </div>

      {/* Filter and Search Bar */}
      <div className="cr-filter-card">
        <div className="cr-search-box">
          <span className="cr-search-icon">🔍</span>
          <input
            type="text"
            className="cr-search-input"
            placeholder="Search by client name, quotation title, contact, GST, or handling staff…"
            value={search}
            onChange={(e) => setSearch(e.target.value)}
          />
          {search && (
            <button type="button" className="cr-search-clear" onClick={() => setSearch('')}>
              ✕
            </button>
          )}
        </div>

        <div className="cr-filter-group">
          {/* Staff Filter */}
          <select
            className="cr-select"
            value={staffFilter}
            onChange={(e) => setStaffFilter(e.target.value)}
            title="Filter by Handling Staff"
          >
            <option value="all">All Handling Staff ({allStaffOptions.length})</option>
            {allStaffOptions.map(([username, dispName]) => (
              <option key={username} value={username}>
                👤 {dispName} ({username})
              </option>
            ))}
          </select>

          {/* Status Filter */}
          <select
            className="cr-select"
            value={statusFilter}
            onChange={(e) => setStatusFilter(e.target.value as any)}
            title="Filter by Quotation Status"
          >
            <option value="all">All Statuses</option>
            <option value="active">Active (Quoted &amp; Negotiating)</option>
            <option value="won">Won Only</option>
            <option value="lost">Lost / Archived Only</option>
          </select>

          {/* Sort Filter */}
          <select
            className="cr-select"
            value={sortBy}
            onChange={(e) => setSortBy(e.target.value as any)}
            title="Sort Client Accounts"
          >
            <option value="quotes_desc">Sort: Most Quotes First</option>
            <option value="value_desc">Sort: Highest Total Value</option>
            <option value="name_asc">Sort: Client Name (A-Z)</option>
          </select>
        </div>
      </div>

      {/* Main Content: Client List with Handled Quotations */}
      {loading ? (
        <div className="card" style={{ padding: '60px 20px', textAlign: 'center', background: '#ffffff', borderRadius: '12px' }}>
          <Spinner size={34} />
          <p className="muted mt-3" style={{ fontSize: '14px' }}>Loading client relationships &amp; quotations matrix…</p>
        </div>
      ) : filteredGroups.length === 0 ? (
        <div className="card" style={{ padding: '60px 20px', textAlign: 'center', background: '#ffffff', borderRadius: '12px', border: '1px dashed #cbd5e1' }}>
          <span style={{ fontSize: '32px' }}>🔍</span>
          <h3 style={{ margin: '12px 0 6px 0', fontSize: '16px', color: '#0f172a' }}>No Client Accounts Found</h3>
          <p className="muted" style={{ margin: 0, fontSize: '13px' }}>No client accounts match your current search and filter criteria.</p>
          <button
            type="button"
            className="btn ghost sm mt-3"
            onClick={() => {
              setSearch('')
              setStaffFilter('all')
              setStatusFilter('all')
            }}
          >
            Reset Filters
          </button>
        </div>
      ) : (
        <div className="cr-client-list">
          {paginatedGroups.map((group) => {
            const clientKey = group.clientName.toLowerCase()
            const isExpanded = expandedClients.has(clientKey)

            return (
              <div
                key={clientKey}
                className={`cr-client-card ${isExpanded ? 'is-expanded' : ''}`}
              >
                {/* Main Client Summary Row */}
                <div
                  className="cr-client-header"
                  onClick={() => toggleClientExpand(clientKey)}
                  role="button"
                  tabIndex={0}
                  onKeyDown={(e) => {
                    if (e.key === 'Enter' || e.key === ' ') {
                      e.preventDefault()
                      toggleClientExpand(clientKey)
                    }
                  }}
                >
                  {/* Left: Avatar + Details + Staff Chips */}
                  <div style={{ display: 'flex', gap: '14px', alignItems: 'flex-start', flex: 1 }}>
                    <div className="cr-client-avatar">
                      🏢
                    </div>

                    <div className="cr-client-info">
                      <div className="cr-client-name-row">
                        <span className="cr-client-title">{group.clientName}</span>
                        {group.gst && (
                          <span className="cr-gst-tag">
                            GST: {group.gst}
                          </span>
                        )}
                      </div>

                      {/* Contact metadata */}
                      <div className="cr-meta-chips">
                        {group.contactPerson && (
                          <span className="cr-meta-item">
                            <span>👤</span> <strong>Attn:</strong> {group.contactPerson}
                          </span>
                        )}
                        {group.phone && (
                          <span className="cr-meta-item">
                            <span>📞</span> {group.phone}
                          </span>
                        )}
                        {group.email && (
                          <span className="cr-meta-item">
                            <span>✉️</span> {group.email}
                          </span>
                        )}
                        {group.address && (
                          <span className="cr-meta-item" title={group.address}>
                            <span>📍</span> {group.address.slice(0, 40)}{group.address.length > 40 ? '…' : ''}
                          </span>
                        )}
                      </div>

                      {/* Staff Handling Chips */}
                      <div className="cr-staff-badges-wrap">
                        <span className="cr-staff-label">Handled By:</span>
                        {group.staffList.length === 0 ? (
                          <span className="tiny muted">No quotations assigned yet</span>
                        ) : (
                          group.staffList.map((st) => (
                            <span
                              key={st.staffUsername}
                              className="cr-staff-badge"
                              title={`${st.displayName} (@${st.staffUsername}) handles ${st.count} quotation(s) totalling ${inr(st.totalValue)}`}
                            >
                              <span className="cr-staff-avatar-mini">
                                {st.displayName.charAt(0).toUpperCase()}
                              </span>
                              <span>{st.displayName}</span>
                              <span className="cr-staff-count">{st.count}</span>
                            </span>
                          ))
                        )}
                      </div>
                    </div>
                  </div>

                  {/* Right: Metrics + Action Buttons */}
                  <div className="cr-client-actions">
                    <div className="cr-quote-stat-pill">
                      <div className="cr-quote-count-bold">
                        {group.totalQuotes} Quotation{group.totalQuotes !== 1 ? 's' : ''}
                      </div>
                      <div className="cr-quote-value-bold">
                        {inr(group.totalValue)} total
                      </div>
                      {group.wonValue > 0 && (
                        <div className="cr-quote-won-tiny">
                          🏆 {inr(group.wonValue)} won ({group.wonQuotes})
                        </div>
                      )}
                    </div>

                    {group.rawClient && onCreateQuotationWithClient && (
                      <button
                        type="button"
                        className="btn ghost sm"
                        onClick={(e) => {
                          e.stopPropagation()
                          onCreateQuotationWithClient(group.rawClient!)
                        }}
                        title="Create new quotation for this client"
                        style={{ padding: '6px 12px', fontSize: '12px', fontWeight: 600 }}
                      >
                        + New Quote
                      </button>
                    )}

                    <button
                      type="button"
                      className="cr-btn-toggle"
                      onClick={(e) => {
                        e.stopPropagation()
                        toggleClientExpand(clientKey)
                      }}
                      title={isExpanded ? 'Collapse quotations' : 'Expand quotations'}
                    >
                      <span>{isExpanded ? 'Hide' : 'View'} ({group.projects.length})</span>
                      <span style={{ fontSize: '11px' }}>{isExpanded ? '▴' : '▾'}</span>
                    </button>
                  </div>
                </div>

                {/* Expanded Drawer: All Quotations for this Client */}
                {isExpanded && (
                  <div className="cr-drawer">
                    <div className="cr-drawer-header">
                      <div>
                        <strong style={{ fontSize: '13.5px', color: '#0f172a' }}>
                          Quotations for {group.clientName} ({group.projects.length})
                        </strong>
                        <div className="tiny muted" style={{ marginTop: '2px' }}>
                          Manage handled estimates, check staff ownership, and jump directly to the calculator.
                        </div>
                      </div>
                    </div>

                    {group.projects.length === 0 ? (
                      <div className="empty" style={{ padding: '20px', background: '#ffffff', borderRadius: '8px', border: '1px solid #e2e8f0', textAlign: 'center' }}>
                        <p className="tiny muted" style={{ margin: 0 }}>
                          No quotations created yet for this client account.
                        </p>
                      </div>
                    ) : (
                      <div className="cr-table-wrap">
                        <table className="cr-table">
                          <thead>
                            <tr>
                              <th style={{ width: '32%' }}>Quotation / Project</th>
                              <th style={{ width: '24%' }}>Handled / Created By</th>
                              <th style={{ width: '14%' }}>Status</th>
                              <th style={{ width: '16%', textAlign: 'right' }}>Total Value</th>
                              <th style={{ width: '14%', textAlign: 'right' }}>Action</th>
                            </tr>
                          </thead>
                          <tbody>
                            {group.projects.map((proj) => {
                              const creatorKey = proj.created_by || 'Unassigned'
                              const matchedMember = memberMap.get(creatorKey.toLowerCase())
                              const staffName = (matchedMember as any)?.name || creatorKey
                              const staffRole = matchedMember?.role || 'member'

                              const isLoss = proj.is_archived || proj.status === 'loss' || proj.status === 'archived' || proj.status === 'lost'
                              const isWon = !proj.is_archived && (proj.status === 'won' || proj.status === 'approved')
                              const isNegotiating = proj.status === 'negotiating'

                              const statusClass = isWon ? 'cr-badge-won' : isLoss ? 'cr-badge-loss' : isNegotiating ? 'cr-badge-negotiating' : 'cr-badge-quoted'
                              const statusLabel = isWon ? 'WON' : isLoss ? 'LOST' : isNegotiating ? 'NEGOTIATING' : 'QUOTED'

                              return (
                                <tr key={proj.id}>
                                  {/* Quotation Name */}
                                  <td>
                                    <div style={{ fontWeight: 700, color: '#0f172a', fontSize: '13px' }}>
                                      {proj.name}
                                    </div>
                                    <div className="tiny muted font-mono" style={{ marginTop: '2px' }}>
                                      ID: {proj.id.slice(0, 8)}… • {timeAgo(proj.updated_at || proj.created_at)}
                                    </div>
                                  </td>

                                  {/* Handled By Staff */}
                                  <td>
                                    <div className="row gap" style={{ alignItems: 'center' }}>
                                      <div
                                        style={{
                                          width: '26px',
                                          height: '26px',
                                          borderRadius: '50%',
                                          background: '#e0e7ff',
                                          color: '#4338ca',
                                          display: 'flex',
                                          alignItems: 'center',
                                          justifyContent: 'center',
                                          fontWeight: 700,
                                          fontSize: '11px',
                                          flexShrink: 0,
                                        }}
                                      >
                                        {staffName.charAt(0).toUpperCase()}
                                      </div>
                                      <div>
                                        <div style={{ fontWeight: 600, color: '#1e293b' }}>
                                          {staffName}
                                        </div>
                                        <div className="tiny muted font-mono">
                                          @{creatorKey} • <span style={{ textTransform: 'capitalize' }}>{staffRole}</span>
                                        </div>
                                      </div>
                                    </div>
                                  </td>

                                  {/* Status */}
                                  <td>
                                    <span className={`cr-quote-badge ${statusClass}`}>
                                      {isWon ? '✓ ' : isNegotiating ? '⏳ ' : isLoss ? '✕ ' : '• '}
                                      {statusLabel}
                                    </span>
                                  </td>

                                  {/* Total Value */}
                                  <td style={{ textAlign: 'right' }}>
                                    <strong style={{ color: '#0f172a', fontFamily: 'monospace', fontSize: '13px' }}>
                                      {inr(Number(proj.total) || 0)}
                                    </strong>
                                  </td>

                                  {/* Action */}
                                  <td style={{ textAlign: 'right' }}>
                                    <button
                                      type="button"
                                      className="cr-btn-open"
                                      onClick={() => onOpenProject(proj.id)}
                                      title="Open this quotation in the calculator"
                                    >
                                      <span>Open Quote</span>
                                      <span>→</span>
                                    </button>
                                  </td>
                                </tr>
                              )
                            })}
                          </tbody>
                        </table>
                      </div>
                    )}
                  </div>
                )}
              </div>
            )
          })}

          {/* Google-Style Pagination Footer */}
          {filteredGroups.length > 0 && (
            <div className="google-pagination-container mt-2">
              <div className="pagination-info">
                Showing <strong>{(activePage - 1) * PAGE_SIZE + 1}</strong> – <strong>{Math.min(activePage * PAGE_SIZE, filteredGroups.length)}</strong> of <strong>{filteredGroups.length}</strong> client accounts
              </div>

              {totalPages > 1 && (
                <div className="google-pagination-controls">
                  <button
                    type="button"
                    className="pagination-nav-btn"
                    disabled={activePage <= 1}
                    onClick={() => {
                      setCurrentPage((p) => Math.max(1, p - 1))
                      window.scrollTo({ top: 0, behavior: 'smooth' })
                    }}
                    title="Previous Page"
                  >
                    ‹ Previous
                  </button>

                  <div className="pagination-numbers">
                    {getPageNumbers(activePage, totalPages).map((p, idx) => {
                      if (p === '...') {
                        return (
                          <span key={`dots-${idx}`} className="pagination-ellipsis">
                            …
                          </span>
                        )
                      }
                      const pageNum = Number(p)
                      const isActive = pageNum === activePage
                      return (
                        <button
                          key={pageNum}
                          type="button"
                          className={`pagination-number-btn ${isActive ? 'active' : ''}`}
                          onClick={() => {
                            setCurrentPage(pageNum)
                            window.scrollTo({ top: 0, behavior: 'smooth' })
                          }}
                        >
                          {pageNum}
                        </button>
                      )
                    })}
                  </div>

                  <button
                    type="button"
                    className="pagination-nav-btn"
                    disabled={activePage >= totalPages}
                    onClick={() => {
                      setCurrentPage((p) => Math.min(totalPages, p + 1))
                      window.scrollTo({ top: 0, behavior: 'smooth' })
                    }}
                    title="Next Page"
                  >
                    Next ›
                  </button>
                </div>
              )}
            </div>
          )}
        </div>
      )}
    </div>
  )
}
