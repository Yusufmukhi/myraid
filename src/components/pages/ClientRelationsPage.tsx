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

  // Filters
  const [search, setSearch] = useState('')
  const [staffFilter, setStaffFilter] = useState<string>('all')
  const [statusFilter, setStatusFilter] = useState<'all' | 'active' | 'won' | 'lost'>('all')
  const [sortBy, setSortBy] = useState<'quotes_desc' | 'value_desc' | 'name_asc'>('quotes_desc')
  const [expandedClients, setExpandedClients] = useState<Set<string>>(new Set())

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
      {/* Top Header */}
      <div className="activity-top-header mb-3">
        <div>
          <h2 style={{ fontSize: '20px', margin: 0, fontWeight: 700 }}>
            Client &amp; Staff Relations
          </h2>
          <p className="tiny muted" style={{ margin: '3px 0 0 0' }}>
            Account distribution matrix: manage client relationships, view quotation assignments, and track multi-staff interactions.
          </p>
        </div>
        <div className="row gap" style={{ alignItems: 'center' }}>
          <button
            className="btn ghost sm"
            onClick={handleExpandAll}
            title={expandedClients.size === filteredGroups.length ? 'Collapse All Accounts' : 'Expand All Accounts'}
          >
            {expandedClients.size === filteredGroups.length ? '▴ Collapse All' : '▾ Expand All'}
          </button>
          {onNavigate && (
            <button
              className="btn ghost sm"
              onClick={() => onNavigate('clients')}
              title="Open Client Directory"
            >
              {Icon.building} Client Directory
            </button>
          )}
          <button className="btn ghost sm" onClick={loadData} title="Refresh relations matrix">
            {Icon.refresh} Refresh
          </button>
        </div>
      </div>

      {/* KPI Overview Cards */}
      <div className="grid-cards mb-3" style={{ gridTemplateColumns: 'repeat(auto-fit, minmax(200px, 1fr))' }}>
        <div className="card stat-card">
          <span className="tiny text-muted uppercase font-bold">Total Clients Tracked</span>
          <div className="stat-value" style={{ fontSize: '24px', fontWeight: 700, color: '#1e293b' }}>
            {summaryMetrics.totalClients}
          </div>
          <span className="tiny muted">{summaryMetrics.totalClientsWithQuotes} clients with active quotations</span>
        </div>

        <div className="card stat-card">
          <span className="tiny text-muted uppercase font-bold">Total Quotations</span>
          <div className="stat-value" style={{ fontSize: '24px', fontWeight: 700, color: '#2563eb' }}>
            {summaryMetrics.totalQuotes}
          </div>
          <span className="tiny muted">Handled across {summaryMetrics.activeStaffCount} staff members</span>
        </div>

        <div className="card stat-card">
          <span className="tiny text-muted uppercase font-bold">Active Pipeline (₹)</span>
          <div className="stat-value" style={{ fontSize: '24px', fontWeight: 700, color: '#f59e0b' }}>
            {inr(summaryMetrics.totalPipeline)}
          </div>
          <span className="tiny muted">Quoted &amp; negotiating value</span>
        </div>

        <div className="card stat-card">
          <span className="tiny text-muted uppercase font-bold">Won Revenue (₹)</span>
          <div className="stat-value" style={{ fontSize: '24px', fontWeight: 700, color: '#10b981' }}>
            {inr(summaryMetrics.totalWon)}
          </div>
          <span className="tiny muted">Successfully closed business</span>
        </div>
      </div>

      {/* Filter and Search Controls */}
      <div className="card filter-bar mb-3">
        <div className="search-wrap">
          <span className="search-icon">{Icon.search}</span>
          <input
            type="text"
            className="search-input"
            placeholder="Search by client name, quotation title, contact, GST, or handling staff…"
            value={search}
            onChange={(e) => setSearch(e.target.value)}
          />
          {search && (
            <button className="search-clear" onClick={() => setSearch('')}>
              ✕
            </button>
          )}
        </div>

        <div className="filter-selects">
          {/* Staff Filter */}
          <select
            className="filter-select"
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

          {/* Quotation Status Filter */}
          <select
            className="filter-select"
            value={statusFilter}
            onChange={(e) => setStatusFilter(e.target.value as any)}
            title="Filter by Quotation Status"
          >
            <option value="all">All Statuses</option>
            <option value="active">Active (Quoted &amp; Negotiating)</option>
            <option value="won">Won Only</option>
            <option value="lost">Lost / Archived Only</option>
          </select>

          {/* Sort By */}
          <select
            className="filter-select"
            value={sortBy}
            onChange={(e) => setSortBy(e.target.value as any)}
            title="Sort Client Accounts"
          >
            <option value="quotes_desc">Sort: Most Quotations First</option>
            <option value="value_desc">Sort: Highest Total Value</option>
            <option value="name_asc">Sort: Client Name (A-Z)</option>
          </select>
        </div>
      </div>

      {/* Main Content: Client List with Handled Quotations */}
      {loading ? (
        <div className="page-loading card" style={{ padding: '40px', textAlign: 'center' }}>
          <Spinner size={32} />
          <p className="muted mt-2">Loading client relationships &amp; quotations…</p>
        </div>
      ) : filteredGroups.length === 0 ? (
        <div className="card empty" style={{ padding: '40px', textAlign: 'center' }}>
          <p>No client relations match your criteria.</p>
          <button className="btn ghost sm mt-2" onClick={() => { setSearch(''); setStaffFilter('all'); setStatusFilter('all'); }}>
            Reset Filters
          </button>
        </div>
      ) : (
        <div className="client-relations-list" style={{ display: 'flex', flexDirection: 'column', gap: '12px' }}>
          {filteredGroups.map((group) => {
            const clientKey = group.clientName.toLowerCase()
            const isExpanded = expandedClients.has(clientKey)

            return (
              <div
                key={clientKey}
                className={`card client-relation-card ${isExpanded ? 'is-expanded' : ''}`}
                style={{
                  border: isExpanded ? '1px solid #93c5fd' : '1px solid #e2e8f0',
                  boxShadow: isExpanded ? '0 4px 12px rgba(37,99,235,0.08)' : undefined,
                  transition: 'all 0.2s ease',
                  padding: '16px',
                }}
              >
                {/* Client Main Summary Row */}
                <div
                  className="row between"
                  style={{
                    alignItems: 'flex-start',
                    cursor: 'pointer',
                    userSelect: 'none',
                    gap: '16px',
                  }}
                  onClick={() => toggleClientExpand(clientKey)}
                >
                  {/* Left: Client Details & Staff Tags */}
                  <div style={{ flex: 1, minWidth: '260px' }}>
                    <div className="row gap" style={{ alignItems: 'center', marginBottom: '6px' }}>
                      <span style={{ fontSize: '18px' }}>🏢</span>
                      <strong style={{ fontSize: '16px', color: '#0f172a' }}>
                        {group.clientName}
                      </strong>
                      {group.gst && (
                        <span className="badge" style={{ background: '#f1f5f9', color: '#475569', fontSize: '11px', fontWeight: 600 }}>
                          GST: {group.gst}
                        </span>
                      )}
                    </div>

                    {/* Metadata line */}
                    <div className="row gap" style={{ flexWrap: 'wrap', fontSize: '12px', color: '#64748b', marginBottom: '8px' }}>
                      {group.contactPerson && (
                        <span>
                          <strong>Attn:</strong> {group.contactPerson}
                        </span>
                      )}
                      {group.phone && (
                        <span>
                          <strong>Phone:</strong> {group.phone}
                        </span>
                      )}
                      {group.email && (
                        <span>
                          <strong>Email:</strong> {group.email}
                        </span>
                      )}
                      {group.address && (
                        <span title={group.address}>
                          <strong>Address:</strong> {group.address.slice(0, 45)}{group.address.length > 45 ? '…' : ''}
                        </span>
                      )}
                    </div>

                    {/* Handling Staff Badges */}
                    <div className="row gap" style={{ alignItems: 'center', flexWrap: 'wrap', marginTop: '6px' }}>
                      <span className="tiny font-bold uppercase" style={{ color: '#64748b', marginRight: '2px' }}>
                        Handling Staff:
                      </span>
                      {group.staffList.length === 0 ? (
                        <span className="tiny muted">No quotations assigned yet</span>
                      ) : (
                        group.staffList.map((st) => (
                          <span
                            key={st.staffUsername}
                            className="badge"
                            style={{
                              background: '#eff6ff',
                              color: '#1d4ed8',
                              border: '1px solid #bfdbfe',
                              fontSize: '11px',
                              padding: '2px 8px',
                              borderRadius: '12px',
                              fontWeight: 600,
                            }}
                            title={`${st.displayName} handles ${st.count} quotation(s) totalling ${inr(st.totalValue)}`}
                          >
                            👤 {st.displayName} (<strong>{st.count} quote{st.count > 1 ? 's' : ''}</strong>)
                          </span>
                        ))
                      )}
                    </div>
                  </div>

                  {/* Right: Metrics & Expand Toggle */}
                  <div className="row gap" style={{ alignItems: 'center', minWidth: '280px', justifyContent: 'flex-end' }}>
                    <div style={{ textAlign: 'right', marginRight: '8px' }}>
                      <div style={{ fontSize: '15px', fontWeight: 700, color: '#1e293b' }}>
                        {group.totalQuotes} Quotation{group.totalQuotes !== 1 ? 's' : ''}
                      </div>
                      <div className="tiny" style={{ color: '#059669', fontWeight: 600 }}>
                        {inr(group.totalValue)} total
                      </div>
                      {group.wonValue > 0 && (
                        <div className="tiny muted">
                          {inr(group.wonValue)} won ({group.wonQuotes})
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
                        style={{ padding: '4px 10px', fontSize: '12px' }}
                      >
                        + New Quote
                      </button>
                    )}

                    <button
                      type="button"
                      className="btn ghost sm"
                      onClick={(e) => {
                        e.stopPropagation()
                        toggleClientExpand(clientKey)
                      }}
                      style={{ padding: '4px 8px', minWidth: '32px' }}
                      title={isExpanded ? 'Collapse' : 'Expand'}
                    >
                      {isExpanded ? '▴' : '▾'}
                    </button>
                  </div>
                </div>

                {/* Expanded Table: Quotations for this Client */}
                {isExpanded && (
                  <div
                    className="client-quotations-drawer mt-3 pt-3"
                    style={{
                      borderTop: '1px solid #e2e8f0',
                      animation: 'fadeIn 0.15s ease',
                    }}
                  >
                    <div className="row between mb-2" style={{ alignItems: 'center' }}>
                      <strong style={{ fontSize: '13px', color: '#334155' }}>
                        Quotations for {group.clientName} ({group.projects.length})
                      </strong>
                      <span className="tiny muted">
                        Click on any quotation to open directly in the calculator
                      </span>
                    </div>

                    {group.projects.length === 0 ? (
                      <div className="empty" style={{ padding: '16px', background: '#f8fafc', borderRadius: '6px' }}>
                        <p className="tiny muted" style={{ margin: 0 }}>
                          No quotations created yet for this client account.
                        </p>
                      </div>
                    ) : (
                      <div className="table-responsive" style={{ overflowX: 'auto' }}>
                        <table className="table" style={{ width: '100%', fontSize: '12.5px', margin: 0 }}>
                          <thead>
                            <tr style={{ background: '#f8fafc', color: '#64748b', textAlign: 'left' }}>
                              <th style={{ padding: '8px 10px', width: '30%' }}>Quotation / Project</th>
                              <th style={{ padding: '8px 10px', width: '22%' }}>Handled / Created By</th>
                              <th style={{ padding: '8px 10px', width: '15%' }}>Status</th>
                              <th style={{ padding: '8px 10px', width: '15%', textAlign: 'right' }}>Total Value</th>
                              <th style={{ padding: '8px 10px', width: '18%', textAlign: 'right' }}>Action</th>
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
                              const statusLabel = isWon ? 'WON' : isLoss ? 'LOST' : isNegotiating ? 'NEGOTIATING' : 'QUOTED'
                              const statusBg = isWon ? '#dcfce7' : isLoss ? '#f1f5f9' : isNegotiating ? '#fef3c7' : '#dbeafe'
                              const statusColor = isWon ? '#15803d' : isLoss ? '#64748b' : isNegotiating ? '#b45309' : '#1d4ed8'

                              return (
                                <tr
                                  key={proj.id}
                                  style={{
                                    borderBottom: '1px solid #f1f5f9',
                                    transition: 'background 0.15s ease',
                                  }}
                                  className="client-project-row"
                                >
                                  {/* Quotation Name */}
                                  <td style={{ padding: '8px 10px', verticalAlign: 'middle' }}>
                                    <div style={{ fontWeight: 600, color: '#1e293b' }}>
                                      {proj.name}
                                    </div>
                                    <div className="tiny muted font-mono">
                                      ID: {proj.id.slice(0, 8)}… • {timeAgo(proj.updated_at || proj.created_at)}
                                    </div>
                                  </td>

                                  {/* Handled By Staff */}
                                  <td style={{ padding: '8px 10px', verticalAlign: 'middle' }}>
                                    <div className="row gap" style={{ alignItems: 'center' }}>
                                      <div
                                        style={{
                                          width: '24px',
                                          height: '24px',
                                          borderRadius: '50%',
                                          background: '#e0e7ff',
                                          color: '#4338ca',
                                          display: 'flex',
                                          alignItems: 'center',
                                          justifyContent: 'center',
                                          fontWeight: 700,
                                          fontSize: '11px',
                                        }}
                                      >
                                        {staffName.charAt(0).toUpperCase()}
                                      </div>
                                      <div>
                                        <div style={{ fontWeight: 600, color: '#334155' }}>
                                          {staffName}
                                        </div>
                                        <span className="tiny muted font-mono">
                                          @{creatorKey} • {staffRole}
                                        </span>
                                      </div>
                                    </div>
                                  </td>

                                  {/* Status */}
                                  <td style={{ padding: '8px 10px', verticalAlign: 'middle' }}>
                                    <span
                                      className="badge"
                                      style={{
                                        background: statusBg,
                                        color: statusColor,
                                        fontWeight: 700,
                                        fontSize: '11px',
                                        padding: '2px 8px',
                                        borderRadius: '6px',
                                      }}
                                    >
                                      {statusLabel}
                                    </span>
                                  </td>

                                  {/* Total Value */}
                                  <td style={{ padding: '8px 10px', verticalAlign: 'middle', textAlign: 'right' }}>
                                    <strong style={{ color: '#0f172a', fontFamily: 'monospace' }}>
                                      {inr(Number(proj.total) || 0)}
                                    </strong>
                                  </td>

                                  {/* Action */}
                                  <td style={{ padding: '8px 10px', verticalAlign: 'middle', textAlign: 'right' }}>
                                    <button
                                      type="button"
                                      className="btn primary sm"
                                      style={{ padding: '4px 10px', fontSize: '11.5px', whiteSpace: 'nowrap' }}
                                      onClick={() => onOpenProject(proj.id)}
                                      title="Open this quotation in the calculator"
                                    >
                                      {Icon.fileText} Open Quote →
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
        </div>
      )}
    </div>
  )
}
