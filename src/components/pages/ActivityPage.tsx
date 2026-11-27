import { useEffect, useMemo, useState } from 'react'
import { ActivityItem, api, Member, Org, Session } from '../../lib/api'
import { inr, timeAgo } from '../../lib/format'
import { cleanDiffText } from '../../lib/revision'
import { fetchClientIp, formatAuditDevice, formatAuditNetwork, getCachedClientIp } from '../../lib/clientInfo'
import { Icon, Spinner, useUI } from '../ui'

interface Props {
  session: Session
  org: Org
  onOpenProject?: (projectId: string) => void
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

export function ActivityPage({ session, org, onOpenProject }: Props) {
  const { toast } = useUI()
  const [logs, setLogs] = useState<ActivityItem[]>([])
  const [orgMembers, setOrgMembers] = useState<Member[]>([])
  const [loading, setLoading] = useState(true)
  const [filterAction, setFilterAction] = useState<string>('all')
  const [selectedStaff, setSelectedStaff] = useState<string>('all')
  const [dateRange, setDateRange] = useState<'all' | 'today' | '7days' | '30days' | '90days'>('all')
  const [search, setSearch] = useState('')
  const [projectMap, setProjectMap] = useState<Map<string, any>>(new Map())
  const [projectRevisionsMap, setProjectRevisionsMap] = useState<Map<string, any[]>>(new Map())
  const [clientIp, setClientIp] = useState<string>(getCachedClientIp())
  const [expandedLogIds, setExpandedLogIds] = useState<Set<string>>(new Set())
  const [modalLog, setModalLog] = useState<ActivityItem | null>(null)
  const [currentPage, setCurrentPage] = useState(1)
  const [summaryTimeframe, setSummaryTimeframe] = useState<'all' | 'today' | 'week' | 'month'>('all')
  const [staffSummaryOpen, setStaffSummaryOpen] = useState(false)
  const [filtersOpen, setFiltersOpen] = useState(false)
  const PAGE_SIZE = 20

  const loadActivity = async () => {
    try {
      setLoading(true)
      const [data, projs, members] = await Promise.all([
        api.activityList(session.token, org.id, 1000),
        api.projectList(session.token, org.id, true).catch(() => []),
        api.orgMembers(session.token, org.id).catch(() => []),
      ])
      setLogs(Array.isArray(data) ? data : [])
      setOrgMembers(Array.isArray(members) ? members : [])
      const pMap = new Map<string, any>()
      for (const p of projs) {
        pMap.set(p.id, p)
        if (p.name) pMap.set(p.name.toLowerCase().trim(), p)
      }

      // Fetch revisions for projects referenced in the audit logs (optimized slice for fast loading)
      const uniqueProjectIds = Array.from(
        new Set(data.map((l) => l.project_id).filter(Boolean))
      ) as string[]

      const revMap = new Map<string, any[]>()
      if (uniqueProjectIds.length > 0) {
        try {
          const fullProjs = await Promise.all(
            uniqueProjectIds.slice(0, 12).map((pid) => api.projectGet(session.token, pid).catch(() => null))
          )
          for (const fp of fullProjs) {
            if (fp && fp.id) {
              pMap.set(fp.id, fp)
              if (fp.name) pMap.set(fp.name.toLowerCase().trim(), fp)
              if (Array.isArray(fp.revisions)) {
                revMap.set(fp.id, fp.revisions)
              }
            }
          }
        } catch {
          // Gracefully continue with available data
        }
      }

      setProjectMap(pMap)
      setProjectRevisionsMap(revMap)
    } catch (e) {
      toast((e as Error).message, 'err')
    } finally {
      setLoading(false)
    }
  }

  useEffect(() => {
    loadActivity()
    fetchClientIp().then(setClientIp).catch(() => {})
  }, [org.id])

  const ownerUsername = (org.role === 'owner' || org.role === 'admin') ? session.username : undefined

  // Identify Admin & Owner usernames to strictly exclude them from Staff activity & performance monitoring
  const adminAndOwnerUsernames = useMemo(() => {
    const set = new Set<string>()
    for (const m of orgMembers) {
      if (m.role === 'owner' || m.role === 'admin') {
        if (m.username) set.add(m.username.toLowerCase().trim())
      }
    }
    // Also protect current session user if they hold admin or owner role
    if (org.role === 'owner' || org.role === 'admin') {
      if (session.username) set.add(session.username.toLowerCase().trim())
    }
    return set
  }, [orgMembers, org.role, session.username])

  // Staff members list strictly excludes Admin and Owner accounts
  const staffMembers = useMemo(() => {
    const set = new Set<string>()
    // 1. Registered org members with staff roles (member or viewer)
    for (const m of orgMembers) {
      if (m.role !== 'owner' && m.role !== 'admin' && m.username) {
        set.add(m.username.trim())
      }
    }
    // 2. Also check log actors, strictly excluding admin and owner
    for (const l of logs || []) {
      const actor = (l.user_name || (l as any).username || l.details?.username || l.details?.user_name || '').trim()
      if (actor && !adminAndOwnerUsernames.has(actor.toLowerCase())) {
        set.add(actor)
      }
    }
    return Array.from(set).sort((a, b) => a.localeCompare(b))
  }, [orgMembers, logs, adminAndOwnerUsernames])

  const securityAlertsCount = useMemo(() => {
    return (logs || []).filter((l) => checkSecurityIncident(l, ownerUsername).isIncident).length
  }, [logs, ownerUsername])

  const staffActivitySummary = useMemo(() => {
    const now = new Date()
    const startOfToday = new Date(now.getFullYear(), now.getMonth(), now.getDate()).getTime()
    const sevenDaysAgo = now.getTime() - 7 * 24 * 60 * 60 * 1000
    const thirtyDaysAgo = now.getTime() - 30 * 24 * 60 * 60 * 1000

    const summaryMap = new Map<string, {
      username: string
      totalActions: number
      created: number
      updated: number
      shared: number
      lastAction: string
      lastActionTime: string
    }>()

    for (const member of staffMembers) {
      summaryMap.set(member, {
        username: member,
        totalActions: 0,
        created: 0,
        updated: 0,
        shared: 0,
        lastAction: 'No activity in period',
        lastActionTime: '',
      })
    }

    for (const log of logs || []) {
      const logTime = new Date(log.created_at).getTime()
      if (summaryTimeframe === 'today' && logTime < startOfToday) continue
      if (summaryTimeframe === 'week' && logTime < sevenDaysAgo) continue
      if (summaryTimeframe === 'month' && logTime < thirtyDaysAgo) continue

      const actor = (log.user_name || (log as any).username || log.details?.username || log.details?.user_name || '').trim()
      if (!actor || adminAndOwnerUsernames.has(actor.toLowerCase())) continue

      let entry = summaryMap.get(actor)
      if (!entry) {
        entry = {
          username: actor,
          totalActions: 0,
          created: 0,
          updated: 0,
          shared: 0,
          lastAction: '',
          lastActionTime: '',
        }
        summaryMap.set(actor, entry)
      }

      entry.totalActions++
      const action = (log.action || '').toLowerCase()
      if (action.includes('create') || action.includes('new')) {
        entry.created++
      } else if (action.includes('update') || action.includes('save') || action.includes('calc') || action.includes('revise')) {
        entry.updated++
      } else if (action.includes('share') || action.includes('pdf')) {
        entry.shared++
      }

      if (!entry.lastActionTime || logTime > new Date(entry.lastActionTime).getTime()) {
        entry.lastActionTime = log.created_at
        entry.lastAction = log.action
      }
    }

    return Array.from(summaryMap.values()).sort((a, b) => b.totalActions - a.totalActions)
  }, [logs, staffMembers, summaryTimeframe, adminAndOwnerUsernames])

  const filteredLogs = useMemo(() => {
    const now = new Date().getTime()
    const startOfToday = new Date(new Date().getFullYear(), new Date().getMonth(), new Date().getDate()).getTime()
    const sevenDaysAgo = now - 7 * 24 * 60 * 60 * 1000
    const thirtyDaysAgo = now - 30 * 24 * 60 * 60 * 1000
    const ninetyDaysAgo = now - 90 * 24 * 60 * 60 * 1000

    return (logs || []).filter((item) => {
      // Date range filter
      if (dateRange !== 'all') {
        const itemTime = new Date(item.created_at).getTime()
        if (dateRange === 'today' && itemTime < startOfToday) return false
        if (dateRange === '7days' && itemTime < sevenDaysAgo) return false
        if (dateRange === '30days' && itemTime < thirtyDaysAgo) return false
        if (dateRange === '90days' && itemTime < ninetyDaysAgo) return false
      }

      if (filterAction === 'security_alerts') {
        if (!checkSecurityIncident(item, ownerUsername).isIncident) return false
      } else if (filterAction !== 'all' && item.action !== filterAction) {
        return false
      }

      if (selectedStaff !== 'all') {
        const actor = (item.user_name || (item as any).username || item.details?.username || item.details?.user_name || '').trim().toLowerCase()
        if (selectedStaff === '__staff_only__') {
          if (adminAndOwnerUsernames.has(actor)) return false
        } else if (selectedStaff === '__leadership__') {
          if (!adminAndOwnerUsernames.has(actor)) return false
        } else {
          if (actor !== selectedStaff.toLowerCase()) return false
        }
      }

      if (search.trim()) {
        const q = search.toLowerCase()
        const userMatch = item.user_name?.toLowerCase().includes(q)
        const nameMatch = item.details?.name?.toLowerCase().includes(q)
        const actionMatch = item.action?.toLowerCase().includes(q)
        return userMatch || nameMatch || actionMatch
      }
      return true
    })
  }, [logs, dateRange, filterAction, selectedStaff, search, ownerUsername, adminAndOwnerUsernames])

  // Reset to page 1 whenever filters, date range, or search change
  useEffect(() => {
    setCurrentPage(1)
  }, [filterAction, selectedStaff, search, dateRange])

  const totalPages = Math.max(1, Math.ceil(filteredLogs.length / PAGE_SIZE))
  const activePage = Math.min(currentPage, totalPages)

  const paginatedLogs = useMemo(() => {
    const start = (activePage - 1) * PAGE_SIZE
    return filteredLogs.slice(start, start + PAGE_SIZE)
  }, [filteredLogs, activePage])

  const groupedPaginatedLogs = useMemo(() => {
    const groups: Array<{ dateKey: string; dateTitle: string; items: Array<{ item: ActivityItem; flatIdx: number }> }> = []
    const now = new Date()
    const todayStr = now.toDateString()
    const yest = new Date(now)
    yest.setDate(yest.getDate() - 1)
    const yestStr = yest.toDateString()

    paginatedLogs.forEach((item, flatIdx) => {
      const d = new Date(item.created_at)
      const dKey = isNaN(d.getTime()) ? 'Unknown' : d.toDateString()
      let title = dKey
      if (!isNaN(d.getTime())) {
        const niceDate = d.toLocaleDateString('en-IN', { day: '2-digit', month: 'short', year: 'numeric' })
        if (dKey === todayStr) title = `Today · ${niceDate}`
        else if (dKey === yestStr) title = `Yesterday · ${niceDate}`
        else title = d.toLocaleDateString('en-IN', { weekday: 'short', day: '2-digit', month: 'short', year: 'numeric' })
      }
      let grp = groups.find((g) => g.dateKey === dKey)
      if (!grp) {
        grp = { dateKey: dKey, dateTitle: title, items: [] }
        groups.push(grp)
      }
      grp.items.push({ item, flatIdx })
    })
    return groups
  }, [paginatedLogs])

  const toggleExpand = (id: string) => {
    setExpandedLogIds((prev) => {
      const next = new Set(prev)
      if (next.has(id)) {
        next.delete(id)
      } else {
        next.add(id)
      }
      return next
    })
  }

  const handleExpandAll = () => {
    if (expandedLogIds.size === paginatedLogs.length) {
      setExpandedLogIds(new Set())
    } else {
      setExpandedLogIds(new Set(paginatedLogs.map((l) => l.id)))
    }
  }

  const activityFilterCount =
    (dateRange !== 'all' ? 1 : 0) +
    (selectedStaff !== 'all' ? 1 : 0) +
    (filterAction !== 'all' ? 1 : 0)

  return (
    <div className="activity-page minimal-activity-page">
      {/* Minimal Top Header */}
      <div className="row between" style={{ flexWrap: 'wrap', gap: 8, alignItems: 'center' }}>
        <div className="row gap" style={{ alignItems: 'baseline', gap: 10 }}>
          <h2 style={{ fontSize: '16px', margin: 0, fontWeight: 700, color: '#0f172a' }}>Activity Log</h2>
          <span className="tiny muted">{filteredLogs.length} events</span>
        </div>
        <div className="row gap" style={{ alignItems: 'center', gap: 6, flexWrap: 'wrap' }}>
          <button
            type="button"
            className="btn ghost sm minimal-pill-btn"
            onClick={() => setStaffSummaryOpen((o) => !o)}
            style={{
              borderColor: staffSummaryOpen ? '#2563eb' : '#e2e8f0',
              background: staffSummaryOpen ? '#eff6ff' : '#ffffff',
              color: staffSummaryOpen ? '#1d4ed8' : '#334155',
            }}
          >
            <span>👥 Staff Summary</span>
            <span style={{ fontSize: '9px', opacity: 0.7 }}>{staffSummaryOpen ? '▲' : '▼'}</span>
          </button>
          <button
            className="btn ghost sm minimal-pill-btn"
            onClick={handleExpandAll}
            title={expandedLogIds.size === paginatedLogs.length ? 'Collapse all log details' : 'Expand all log details'}
          >
            {expandedLogIds.size === paginatedLogs.length && paginatedLogs.length > 0 ? '▴ Collapse All' : '▾ Expand All'}
          </button>
          <button className="btn ghost sm minimal-pill-btn" onClick={loadActivity} title="Refresh activity log">
            {Icon.refresh} Refresh
          </button>
        </div>
      </div>

      {/* Security Incident Banner (Shows ONLY when real incidents exist) */}
      {securityAlertsCount > 0 && (
        <div className="card" style={{ background: '#fef2f2', border: '1px solid #fecaca', borderRadius: 8, padding: '8px 12px', display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: 8 }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: 8, fontSize: '12px', color: '#991b1b' }}>
            <span>🛡️</span>
            <strong>{securityAlertsCount} Security Event{securityAlertsCount > 1 ? 's' : ''}</strong>
            <span className="tiny" style={{ color: '#7f1d1d' }}>· Access requests, deletions, or high discounts</span>
          </div>
          <button
            className={`btn sm ${filterAction === 'security_alerts' ? 'primary' : 'ghost'}`}
            onClick={() => setFilterAction(filterAction === 'security_alerts' ? 'all' : 'security_alerts')}
            style={filterAction === 'security_alerts' ? { height: 26, fontSize: '11.5px' } : { height: 26, fontSize: '11.5px', color: '#991b1b', borderColor: '#fca5a5' }}
          >
            {filterAction === 'security_alerts' ? 'Show All' : `Filter (${securityAlertsCount})`}
          </button>
        </div>
      )}

      {/* Collapsible Staff Activity Summary (Hidden by default; opens on Staff Summary button click) */}
      {staffSummaryOpen && (
        <div
          className="card"
          style={{
            background: '#ffffff',
            border: '1px solid #e2e8f0',
            borderRadius: '8px',
            padding: '12px 14px',
          }}
        >
          <div
            style={{
              display: 'flex',
              alignItems: 'center',
              justifyContent: 'space-between',
              marginBottom: '10px',
              flexWrap: 'wrap',
              gap: '8px',
            }}
          >
            <strong style={{ fontSize: '13px', color: '#0f172a' }}>📈 Staff Activity Summary</strong>
            <div className="status-tabs" style={{ margin: 0 }}>
              <button className={'status-tab ' + (summaryTimeframe === 'all' ? 'active' : '')} onClick={() => setSummaryTimeframe('all')} style={{ fontSize: '11px', padding: '2px 8px' }}>All</button>
              <button className={'status-tab ' + (summaryTimeframe === 'today' ? 'active' : '')} onClick={() => setSummaryTimeframe('today')} style={{ fontSize: '11px', padding: '2px 8px' }}>Today</button>
              <button className={'status-tab ' + (summaryTimeframe === 'week' ? 'active' : '')} onClick={() => setSummaryTimeframe('week')} style={{ fontSize: '11px', padding: '2px 8px' }}>7d</button>
              <button className={'status-tab ' + (summaryTimeframe === 'month' ? 'active' : '')} onClick={() => setSummaryTimeframe('month')} style={{ fontSize: '11px', padding: '2px 8px' }}>30d</button>
            </div>
          </div>

          {staffActivitySummary.length === 0 ? (
            <div className="tiny muted text-center" style={{ padding: '8px 0' }}>
              No staff activity recorded for this time period.
            </div>
          ) : (
            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(200px, 1fr))', gap: '8px' }}>
              {staffActivitySummary.map((staff) => {
                const isSelected = selectedStaff === staff.username
                return (
                  <div
                    key={staff.username}
                    onClick={() => setSelectedStaff(isSelected ? 'all' : staff.username)}
                    style={{
                      background: isSelected ? '#eff6ff' : '#f8fafc',
                      border: isSelected ? '1.5px solid #2563eb' : '1px solid #e2e8f0',
                      borderRadius: '7px',
                      padding: '8px 10px',
                      cursor: 'pointer',
                    }}
                    title={`Click to filter by @${staff.username}`}
                  >
                    <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: '4px' }}>
                      <strong style={{ fontSize: '12.5px', color: '#1e293b' }}>@{staff.username}</strong>
                      <span className="chip sm font-bold" style={{ fontSize: '10px', background: staff.totalActions > 0 ? '#dbeafe' : '#f1f5f9', color: staff.totalActions > 0 ? '#1d4ed8' : '#64748b' }}>
                        {staff.totalActions}
                      </span>
                    </div>
                    <div style={{ display: 'flex', gap: '10px', fontSize: '11px', color: '#475569' }}>
                      <span>New: <strong>{staff.created}</strong></span>
                      <span>Saved: <strong>{staff.updated}</strong></span>
                      <span>Shared: <strong>{staff.shared}</strong></span>
                    </div>
                  </div>
                )
              })}
            </div>
          )}
        </div>
      )}

      {/* Slim Search Toolbar + Collapsible Filters Box */}
      <div className="card minimal-toolbar-card">
        <div className="minimal-toolbar-row">
          <div className="search-wrap minimal-search-wrap">
            <span className="search-icon">{Icon.search}</span>
            <input
              type="text"
              className="search-input"
              placeholder="Search user, quotation name, or event…"
              value={search}
              onChange={(e) => setSearch(e.target.value)}
            />
            {search && (
              <button className="icon-btn sm" onClick={() => setSearch('')} title="Clear search">
                {Icon.close}
              </button>
            )}
          </div>

          <button
            type="button"
            className="btn ghost sm minimal-pill-btn"
            onClick={() => setFiltersOpen((o) => !o)}
            style={{
              borderColor: filtersOpen || activityFilterCount > 0 ? '#2563eb' : '#e2e8f0',
              background: filtersOpen || activityFilterCount > 0 ? '#eff6ff' : '#ffffff',
              color: filtersOpen || activityFilterCount > 0 ? '#1d4ed8' : '#334155',
            }}
          >
            <span>🎛 Filters</span>
            {activityFilterCount > 0 && <span className="minimal-count-dot">{activityFilterCount}</span>}
          </button>
        </div>

        {filtersOpen && (
          <div className="minimal-filters-box">
            <div className="minimal-filter-group">
              <span className="minimal-filter-label">Date Range</span>
              <select className="select-input compact" value={dateRange} onChange={(e) => setDateRange(e.target.value as any)}>
                <option value="all">All Time</option>
                <option value="today">Today</option>
                <option value="7days">Last 7 Days</option>
                <option value="30days">Last 30 Days</option>
                <option value="90days">Last 90 Days</option>
              </select>
            </div>

            <div className="minimal-filter-group">
              <span className="minimal-filter-label">User</span>
              <select className="select-input compact" value={selectedStaff} onChange={(e) => setSelectedStaff(e.target.value)}>
                <option value="all">All Users ({logs.length})</option>
                {staffMembers.length > 0 && <option value="__staff_only__">Staff Only</option>}
                {staffMembers.map((staff) => (
                  <option key={staff} value={staff}>@{staff}</option>
                ))}
                <option value="__leadership__">Admins &amp; Owners</option>
              </select>
            </div>

            <div className="minimal-filter-group">
              <span className="minimal-filter-label">Event Type</span>
              <select className="select-input compact" value={filterAction} onChange={(e) => setFilterAction(e.target.value)}>
                <option value="all">All Events</option>
                {securityAlertsCount > 0 && <option value="security_alerts">Security Incidents ({securityAlertsCount})</option>}
                <option value="project_created">Created</option>
                <option value="project_updated">Saved / Updated</option>
                <option value="pdf_downloaded">PDF Downloaded</option>
                <option value="project_duplicated">Revised / Copied</option>
                <option value="project_recalculated">Recalculated</option>
                <option value="template_published">Template Published</option>
                <option value="project_archived">Archived</option>
                <option value="project_deleted">Deleted</option>
                <option value="member_joined">Member Joined</option>
                <option value="member_role_changed">Role Changed</option>
                <option value="join_requested">Join Requested</option>
                <option value="member_rejected">Request Rejected</option>
                <option value="join_code_rotated">Code Rotated</option>
              </select>
            </div>

            {activityFilterCount > 0 && (
              <button
                type="button"
                className="link-inline tiny"
                style={{ marginLeft: 'auto' }}
                onClick={() => {
                  setDateRange('all')
                  setSelectedStaff('all')
                  setFilterAction('all')
                }}
              >
                Reset
              </button>
            )}
          </div>
        )}
      </div>

      {/* Activity Timeline List (Grouped by Day) */}
      <div className="card activity-audit-card clean-audit-shell">
        {loading ? (
          <div className="page-loading">
            <Spinner size={28} />
            <p className="muted">Loading activity log…</p>
          </div>
        ) : filteredLogs.length === 0 ? (
          <div className="empty">
            <p>No activity records match your filters</p>
          </div>
        ) : (
          <div className="clean-audit-feed">
            {groupedPaginatedLogs.map((group) => (
              <div key={group.dateKey} className="clean-day-group">
                <div className="clean-day-header">
                  <span className="clean-day-title">{group.dateTitle}</span>
                  <span className="clean-day-count">{group.items.length} event{group.items.length > 1 ? 's' : ''}</span>
                </div>

                {group.items.map(({ item, flatIdx }) => {
                  const { badgeClass, label } = getActionMeta(item.action)
                  const shortTag = getActionShortTag(item.action)
                  const proj = item.project_id
                    ? projectMap.get(item.project_id)
                    : item.details?.name
                    ? projectMap.get(item.details.name.toLowerCase().trim())
                    : null

                  const rawIp = item.details?.ip || proj?.summary?.ip || proj?.summary?.client_info?.ip || clientIp
                  const netInfo = formatAuditNetwork(rawIp)
                  const rawDevice = item.details?.device || proj?.summary?.device || proj?.summary?.client_info?.device
                  const devInfo = formatAuditDevice(rawDevice)

                  const extracted = extractLogChanges(
                    item,
                    flatIdx,
                    paginatedLogs,
                    projectMap,
                    projectRevisionsMap
                  )
                  const isExpanded = expandedLogIds.has(item.id)
                  const changeCount = extracted?.changes?.length || 0
                  const plainSummaryText = changeCount > 0
                    ? extracted.changes
                        .slice(0, 2)
                        .map((c: any) => cleanDiffText(typeof c === 'string' ? c : (c?.diff || `${c?.label || ''}: ${c?.oldVal || ''} → ${c?.newVal || ''}`)))
                        .join(' · ')
                    : ''
                  const exactClockTime = new Date(item.created_at).toLocaleTimeString('en-IN', {
                    hour: '2-digit',
                    minute: '2-digit',
                  })

                  return (
                    <div
                      key={item.id}
                      className={`clean-audit-row-wrap ${isExpanded ? 'is-open' : ''}`}
                    >
                      <div
                        className="clean-audit-row"
                        onClick={() => toggleExpand(item.id)}
                        role="button"
                        tabIndex={0}
                        onKeyDown={(e) => {
                          if (e.key === 'Enter' || e.key === ' ') {
                            e.preventDefault()
                            toggleExpand(item.id)
                          }
                        }}
                      >
                        {/* 1. Time Column */}
                        <div className="clean-time-col" title={new Date(item.created_at).toLocaleString('en-IN', { dateStyle: 'full', timeStyle: 'medium' })}>
                          <span className="clean-clock">{exactClockTime}</span>
                          <span className="clean-ago">{timeAgo(item.created_at)}</span>
                        </div>

                        {/* 2. Status Icon + Category Tag */}
                        <div className="clean-type-col">
                          <span className={`audit-badge clean-icon-badge ${badgeClass}`}>{getActionIcon(item.action)}</span>
                          <span className={`clean-event-tag ${badgeClass}`}>{shortTag}</span>
                        </div>

                        {/* 3. Plain-English Event & Changes Preview */}
                        <div className="clean-main-col">
                          <div className="clean-headline">
                            <strong className="clean-actor">{item.user_name}</strong>
                            <span className="clean-verb">{label}</span>
                            {item.details?.client_name && (
                              <strong className="clean-target-client">“{item.details.client_name}”</strong>
                            )}
                            {item.details?.name && !item.details?.client_name && (
                              <button
                                type="button"
                                className="clean-target-link"
                                onClick={(e) => {
                                  if (item.project_id && onOpenProject) {
                                    e.stopPropagation()
                                    onOpenProject(item.project_id)
                                  }
                                }}
                                title={item.project_id ? 'Open quotation in calculator' : undefined}
                              >
                                {item.details.name}
                              </button>
                            )}
                            {changeCount > 0 && (
                              <span className="clean-change-count">
                                {changeCount} change{changeCount > 1 ? 's' : ''}
                              </span>
                            )}
                            {item.action === 'client_conflict' && item.details?.existing_owner && (
                              <span className="clean-warn-note">
                                (Handled by @{item.details.existing_owner}{item.details?.gst ? ` · GST ${item.details.gst}` : ''})
                              </span>
                            )}
                          </div>

                          {plainSummaryText && (
                            <div className="clean-preview-chips">
                              {extracted.changes.slice(0, 3).map((ch: any, cIdx: number) => {
                                const parsed = parseDiffItem(ch)
                                return (
                                  <span key={cIdx} className="clean-preview-chip">
                                    {parsed.scope && <strong className="clean-chip-scope">{parsed.scope}</strong>}
                                    <span>{parsed.label || parsed.rawText}</span>
                                    {parsed.newVal !== undefined && (
                                      <strong className="clean-chip-val">→ {parsed.newVal}</strong>
                                    )}
                                  </span>
                                )
                              })}
                              {changeCount > 3 && (
                                <span className="clean-more-chip">+{changeCount - 3} more</span>
                              )}
                            </div>
                          )}
                        </div>

                        {/* 4. Right Amount & Expand Button */}
                        <div className="clean-right-col" onClick={(e) => e.stopPropagation()}>
                          <div className="clean-right-meta">
                            {item.details?.total !== undefined && Number(item.details.total) > 0 && (
                              <strong className="clean-amount">{inr(Number(item.details.total))}</strong>
                            )}
                            <span className="clean-device-ip" title={`${devInfo.label} · ${netInfo.tooltip}`}>
                              {devInfo.label} · {netInfo.ip || netInfo.label}
                            </span>
                          </div>
                          <button
                            className={`btn sm ${isExpanded ? 'primary' : 'ghost'} minimal-pill-btn clean-toggle-btn`}
                            onClick={() => toggleExpand(item.id)}
                            aria-expanded={isExpanded}
                          >
                            {isExpanded ? 'Hide ▴' : 'Details ▾'}
                          </button>
                        </div>
                      </div>

                      {/* Expanded Clean Structured Change Table & Metadata Footer */}
                      {isExpanded && (
                        <div className="clean-drawer">
                          <div className="clean-diff-table-wrap">
                            <div className="clean-diff-table-head">
                              <span>{extracted.sectionTitle} ({changeCount} item{changeCount === 1 ? '' : 's'})</span>
                              {item.details?.buyer_name && (
                                <span className="tiny muted">Buyer: <strong>{item.details.buyer_name}</strong></span>
                              )}
                            </div>
                            <table className="clean-diff-table">
                              <thead>
                                <tr>
                                  <th style={{ width: '110px' }}>Section</th>
                                  <th style={{ width: '28%' }}>Field</th>
                                  <th>Before</th>
                                  <th style={{ width: '24px' }}></th>
                                  <th>After</th>
                                </tr>
                              </thead>
                              <tbody>
                                {extracted.changes.map((ch: any, idx: number) => (
                                  <DiffItemRenderer key={idx} item={ch} />
                                ))}
                              </tbody>
                            </table>
                          </div>

                          <div className="clean-meta-strip">
                            <div className="clean-meta-items">
                              <span><strong>Actor:</strong> {item.user_name} ({item.user_name.toLowerCase() === session.username.toLowerCase() ? org.role : 'Staff'})</span>
                              <span><strong>Time:</strong> {new Date(item.created_at).toLocaleString('en-IN', { dateStyle: 'medium', timeStyle: 'short' })}</span>
                              <span><strong>Device:</strong> {devInfo.icon} {devInfo.label}</span>
                              <span><strong>Network:</strong> {netInfo.ip} ({netInfo.label})</span>
                              <span><strong>Event:</strong> <code>{item.action}</code></span>
                            </div>
                            <div className="clean-meta-actions">
                              {item.project_id && onOpenProject && (
                                <button
                                  type="button"
                                  className="btn primary sm"
                                  style={{ height: '26px', padding: '0 10px', fontSize: '11.5px' }}
                                  onClick={() => onOpenProject(item.project_id!)}
                                >
                                  Open Quotation →
                                </button>
                              )}
                              <button
                                type="button"
                                className="btn ghost sm"
                                style={{ height: '26px', padding: '0 9px', fontSize: '11.5px' }}
                                onClick={() => setModalLog(item)}
                                title="Inspect raw JSON event payload"
                              >
                                🔍 Full Inspector
                              </button>
                            </div>
                          </div>
                        </div>
                      )}
                    </div>
                  )
                })}
              </div>
            ))}
          </div>
        )}

        {/* Google-Style Pagination Footer (20 items per page) */}
        {!loading && filteredLogs.length > 0 && (
          <div className="google-pagination-container">
            <div className="pagination-info">
              Showing <strong>{(activePage - 1) * PAGE_SIZE + 1}</strong> – <strong>{Math.min(activePage * PAGE_SIZE, filteredLogs.length)}</strong> of <strong>{filteredLogs.length}</strong> audit records
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

      {/* Modal Inspector for Deep Audit Inspection */}
      {modalLog && (
        <div className="modal-backdrop" onClick={() => setModalLog(null)}>
          <div className="modal" onClick={(e) => e.stopPropagation()} style={{ maxWidth: '640px' }}>
            <div className="modal-header">
              <h3>Audit Record Inspector: {modalLog.user_name} ({modalLog.action})</h3>
              <button className="icon-btn" onClick={() => setModalLog(null)}>
                {Icon.close}
              </button>
            </div>
            <div className="audit-modal-body">
              <div className="row between mb-2">
                <div>
                  <strong>Event ID:</strong> <span className="font-mono">{modalLog.id}</span>
                </div>
                <span className="tiny muted">
                  {new Date(modalLog.created_at).toLocaleString('en-IN', {
                    dateStyle: 'full',
                    timeStyle: 'medium',
                  })}
                </span>
              </div>

              <h4 style={{ fontSize: '13px', marginTop: '12px', marginBottom: '6px' }}>Detailed Log Parameters:</h4>
              <pre
                style={{
                  background: '#0f172a',
                  color: '#38bdf8',
                  padding: '12px',
                  borderRadius: '8px',
                  fontSize: '12px',
                  overflowX: 'auto',
                  maxHeight: '300px',
                }}
              >
                {JSON.stringify(modalLog, null, 2)}
              </pre>

              <div className="row end gap mt-3">
                <button className="btn ghost sm" onClick={() => setModalLog(null)}>
                  Close
                </button>
                {modalLog.project_id && onOpenProject && (
                  <button
                    className="btn primary sm"
                    onClick={() => {
                      onOpenProject(modalLog.project_id!)
                      setModalLog(null)
                    }}
                  >
                    Open Quotation
                  </button>
                )}
              </div>
            </div>
          </div>
        </div>
      )}
    </div>
  )
}

function getActionIcon(action: string) {
  switch (action) {
    case 'project_created':
      return Icon.plus
    case 'pdf_downloaded':
    case 'pdf_saved':
      return Icon.fileText
    case 'project_updated':
      return Icon.save
    case 'project_duplicated':
      return Icon.copy
    case 'project_recalculated':
      return Icon.refresh
    case 'template_published':
      return Icon.layers
    case 'project_archived':
      return Icon.archive
    case 'project_deleted':
      return Icon.trash
    case 'member_joined':
    case 'member_role_changed':
      return Icon.users
    case 'join_requested':
      return Icon.clock
    case 'member_rejected':
      return Icon.close
    case 'join_code_rotated':
      return Icon.shield
    default:
      return Icon.info
  }
}

function getActionShortTag(action: string): string {
  switch (action) {
    case 'project_created': return 'Created'
    case 'pdf_downloaded': return 'PDF Download'
    case 'pdf_saved': return 'Edit PDF'
    case 'project_updated': return 'Saved'
    case 'project_duplicated': return 'Revised'
    case 'project_recalculated': return 'Recalc'
    case 'template_published': return 'Template'
    case 'project_archived': return 'Archived'
    case 'project_deleted': return 'Deleted'
    case 'member_joined': return 'Joined'
    case 'member_role_changed': return 'Role'
    case 'join_requested': return 'Access Req'
    case 'member_rejected': return 'Rejected'
    case 'join_code_rotated': return 'Security'
    case 'client_conflict': return 'Conflict'
    default: return action.replace(/_/g, ' ')
  }
}

function getActionMeta(action: string): { badgeClass: string; label: string } {
  switch (action) {
    case 'project_created':
      return { badgeClass: 'badge-green', label: 'created project' }
    case 'pdf_downloaded':
      return { badgeClass: 'badge-emerald', label: 'downloaded PDF quotation' }
    case 'pdf_saved':
      return { badgeClass: 'badge-blue', label: 'saved quotation details' }
    case 'project_updated':
      return { badgeClass: 'badge-blue', label: 'updated project calculations' }
    case 'project_duplicated':
      return { badgeClass: 'badge-purple', label: 'duplicated snapshot of' }
    case 'project_recalculated':
      return { badgeClass: 'badge-amber', label: 'recalculated project template' }
    case 'template_published':
      return { badgeClass: 'badge-indigo', label: 'published default template' }
    case 'project_archived':
      return { badgeClass: 'badge-muted', label: 'archived project' }
    case 'project_deleted':
      return { badgeClass: 'badge-red', label: 'deleted project' }
    case 'member_joined':
      return { badgeClass: 'badge-green', label: 'joined organisation' }
    case 'member_role_changed':
      return { badgeClass: 'badge-blue', label: 'updated role' }
    case 'join_requested':
      return { badgeClass: 'badge-amber', label: 'requested to join organisation' }
    case 'member_rejected':
      return { badgeClass: 'badge-red', label: 'join request rejected' }
    case 'join_code_rotated':
      return { badgeClass: 'badge-purple', label: 'regenerated organisation join code' }
    case 'client_conflict':
      return { badgeClass: 'badge-amber', label: 'attempted to register an existing client' }
    default:
      return { badgeClass: 'badge-muted', label: action.replace(/_/g, ' ') }
  }
}

function checkSecurityIncident(
  item: ActivityItem,
  ownerUsername?: string
): { isIncident: boolean; isAuthorizedOwner?: boolean; reason?: string } {
  const isOwnerAction = Boolean(ownerUsername && item.user_name.toLowerCase() === ownerUsername.toLowerCase())

  if (item.action === 'join_requested') {
    return { isIncident: true, reason: 'Pending access request from external user' }
  }
  if (item.action === 'member_rejected') {
    return { isIncident: true, reason: 'Join request was rejected by admin' }
  }

  if (item.action === 'project_deleted') {
    return { isIncident: !isOwnerAction, reason: 'Quotation permanently deleted' }
  }
  if (item.action === 'client_conflict') {
    return {
      isIncident: true,
      reason: `Staff collision: attempted to add client already assigned to @${item.details?.existing_owner || 'another staff'}`,
    }
  }

  const note = String(item.details?.note || '')
  const changes = Array.isArray(item.details?.changes) ? item.details.changes.join(' ') : ''
  const combined = `${note} ${changes}`
  if (/discount.*[2-9][0-9]%/i.test(combined) || combined.includes('→ 0') || (item.details?.total !== undefined && Number(item.details.total) === 0)) {
    if (!isOwnerAction) {
      return { isIncident: true, reason: 'Excessive staff discount (>20%) or price zero-out' }
    }
  }

  return { isIncident: false }
}

interface LogExtractedChanges {
  sectionTitle: string
  changes: any[]
}

function extractLogChanges(
  item: ActivityItem,
  itemIdx: number,
  allLogs: ActivityItem[],
  projectMap: Map<string, any>,
  projectRevisionsMap: Map<string, any[]>
): LogExtractedChanges {
  const proj = item.project_id
    ? projectMap.get(item.project_id)
    : item.details?.name
    ? projectMap.get(item.details.name.toLowerCase().trim())
    : null

  // 1. Direct changes array in item.details
  if (Array.isArray(item.details?.changes) && item.details.changes.length > 0) {
    return {
      sectionTitle: 'Modified Values',
      changes: item.details.changes,
    }
  }

  // 2. Parsable diffs in item.details.note
  const note = item.details?.note
  if (note && note !== 'Project update' && note !== 'Saved project') {
    const parts = note.split(/[;\n]/).map((s: string) => s.trim()).filter(Boolean)
    if (parts.length > 0 && parts.some((p: string) => p.includes('→') || p.includes('->') || p.includes(':'))) {
      return {
        sectionTitle: 'Modified Values',
        changes: parts,
      }
    }
  }

  // 3. For project_updated, check project_revisions map
  if (item.action === 'project_updated') {
    const revs = item.project_id ? projectRevisionsMap.get(item.project_id) : null
    if (Array.isArray(revs) && revs.length > 0) {
      const itemTotal = item.details?.total !== undefined ? Number(item.details.total) : null
      const itemTime = new Date(item.created_at).getTime()

      let matchedRev = revs.find((r) => {
        if (itemTotal !== null && r.total !== null && Math.abs(Number(r.total) - itemTotal) < 1.0) {
          return true
        }
        return false
      })

      if (!matchedRev) {
        matchedRev = revs.find((r) => {
          const revTime = new Date(r.created_at).getTime()
          return Math.abs(revTime - itemTime) < 25000
        })
      }

      if (matchedRev) {
        const revChanges = matchedRev.changes || matchedRev.summary?.changes
        if (Array.isArray(revChanges) && revChanges.length > 0) {
          return {
            sectionTitle: 'Modified Values',
            changes: revChanges,
          }
        }
        if (matchedRev.note && matchedRev.note !== 'Project update' && matchedRev.note !== 'Saved project') {
          const parts = matchedRev.note.split(/[;\n]/).map((s: string) => s.trim()).filter(Boolean)
          if (parts.length > 0 && parts.some((p: string) => p.includes('→') || p.includes('->') || p.includes(':'))) {
            return {
              sectionTitle: 'Modified Values',
              changes: parts,
            }
          }
        }
      }
    }

    // 4. Check if current project summary changes apply
    if (proj?.summary?.changes && Array.isArray(proj.summary.changes) && proj.summary.changes.length > 0) {
      const isLatest =
        !item.details?.total ||
        !proj?.total ||
        Math.abs(Number(item.details.total) - Number(proj.total)) < 1.0
      if (isLatest) {
        return {
          sectionTitle: 'Modified Values',
          changes: proj.summary.changes,
        }
      }
    }

    // 5. Compare with adjacent previous log of the same project
    const prevLog = allLogs.slice(itemIdx + 1).find((l) => {
      if (item.project_id && l.project_id === item.project_id) return true
      if (item.details?.name && l.details?.name && item.details.name === l.details.name) return true
      return false
    })

    const synthChanges: string[] = []
    if (item.details?.old_version && item.details?.new_version && item.details.old_version !== item.details.new_version) {
      synthChanges.push(`Template Engine: v${item.details.old_version} → v${item.details.new_version}`)
    }
    if (
      prevLog?.details?.total !== undefined &&
      item.details?.total !== undefined &&
      Math.abs(Number(prevLog.details.total) - Number(item.details.total)) >= 1.0
    ) {
      synthChanges.push(
        `Quotation Total: ₹${Math.round(Number(prevLog.details.total)).toLocaleString('en-IN')} → ₹${Math.round(
          Number(item.details.total)
        ).toLocaleString('en-IN')}`
      )
    }
    if (prevLog?.details?.status && item.details?.status && prevLog.details.status !== item.details.status) {
      synthChanges.push(`Status: ${String(prevLog.details.status).toUpperCase()} → ${String(item.details.status).toUpperCase()}`)
    }
    if (prevLog?.details?.name && item.details?.name && prevLog.details.name !== item.details.name) {
      synthChanges.push(`Quotation Name: “${prevLog.details.name}” → “${item.details.name}”`)
    }

    if (synthChanges.length > 0) {
      return {
        sectionTitle: 'Modified Values',
        changes: synthChanges,
      }
    }

    const displayTotal = item.details?.total ? inr(Number(item.details.total)) : proj?.total ? inr(Number(proj.total)) : null
    return {
      sectionTitle: 'Quotation State',
      changes: [
        displayTotal
          ? `Quotation Verified & Saved (Total: ${displayTotal})`
          : 'Quotation parameters verified and saved',
      ],
    }
  }

  // 6. Non-update Actions (project_created, duplicated, recalculated, member actions, etc.)
  if (item.action === 'project_created') {
    const list: string[] = []
    if (item.details?.total) list.push(`Quotation Total: ${inr(Number(item.details.total))}`)
    if (item.details?.template_version) list.push(`Template Baseline: v${item.details.template_version}`)
    if (item.details?.status) list.push(`Status: ${String(item.details.status).toUpperCase()}`)
    return {
      sectionTitle: 'Initial Creation Details',
      changes: list.length > 0 ? list : ['Initial quotation created and calculation baseline initialized'],
    }
  }

  if (item.action === 'project_duplicated') {
    const list: string[] = []
    if (item.details?.source_name) list.push(`Cloned From: “${item.details.source_name}”`)
    if (item.details?.name) list.push(`New Snapshot: “${item.details.name}”`)
    if (item.details?.total) list.push(`Quotation Total: ${inr(Number(item.details.total))}`)
    return {
      sectionTitle: 'Duplicate Snapshot Details',
      changes: list.length > 0 ? list : ['Quotation duplicated from previous version snapshot'],
    }
  }

  if (item.action === 'project_recalculated') {
    const list: string[] = []
    if (item.details?.old_version && item.details?.new_version) {
      list.push(`Template Upgrade: v${item.details.old_version} → v${item.details.new_version}`)
    }
    if (item.details?.total) list.push(`Recalculated Total: ${inr(Number(item.details.total))}`)
    return {
      sectionTitle: 'Recalculation Details',
      changes: list.length > 0 ? list : ['Formulas recalculated against current pricing template'],
    }
  }

  if (item.action === 'template_published') {
    return {
      sectionTitle: 'Template Publication',
      changes: [
        `Template Version v${item.details?.version_num || item.details?.new_version || '1'} Published`,
        'Active as default calculation model for new quotations',
      ],
    }
  }

  if (item.action === 'member_joined') {
    return {
      sectionTitle: 'Membership Access',
      changes: [
        'Team member onboarded to organization',
        `Assigned Role: ${String(item.details?.role || 'member').toUpperCase()}`,
      ],
    }
  }

  if (item.action === 'member_role_changed') {
    return {
      sectionTitle: 'Role Update',
      changes: [
        `Permission Level Updated to: ${String(item.details?.new_role || 'staff').toUpperCase()}`,
      ],
    }
  }

  if (item.action === 'join_requested') {
    return {
      sectionTitle: 'Access Review',
      changes: ['External user submitted organization join code for administrative review'],
    }
  }

  if (item.action === 'member_rejected') {
    return {
      sectionTitle: 'Request Outcome',
      changes: ['Organization join request was rejected by administrator'],
    }
  }

  if (item.action === 'join_code_rotated') {
    return {
      sectionTitle: 'Security Key Rotation',
      changes: ['Organization invite code invalidated; new secret key generated'],
    }
  }

  if (item.action === 'pdf_downloaded' || item.action === 'pdf_saved') {
    const list: string[] = []
    if (item.details?.file_name) list.push(`PDF File: “${item.details.file_name}.pdf”`)
    if (item.details?.buyer_name) list.push(`Client / Buyer: ${item.details.buyer_name}`)
    if (item.details?.qtn_no) list.push(`Quotation No: ${item.details.qtn_no}`)
    if (item.details?.total) list.push(`Quotation Total: ₹${Math.round(Number(item.details.total)).toLocaleString('en-IN')}`)
    if (item.details?.mode) list.push(`Quotation Layout: ${String(item.details.mode).replace('_', ' ').toUpperCase()}`)
    if (item.details?.method === 'print_dialog') list.push('Method: Browser Print / Save as PDF')
    return {
      sectionTitle: item.action === 'pdf_downloaded' ? 'PDF Generation & Download' : 'Quotation Details Saved',
      changes: list.length > 0 ? list : ['Quotation PDF event recorded'],
    }
  }

  if (item.action === 'project_deleted') {
    return {
      sectionTitle: 'Quotation Removal',
      changes: [`Quotation “${item.details?.name || 'Project'}” permanently removed from active records`],
    }
  }

  if (item.action === 'project_archived') {
    return {
      sectionTitle: 'Archive Storage',
      changes: ['Quotation moved into archived records'],
    }
  }

  return {
    sectionTitle: 'Event Details',
    changes: [item.details?.note || `Event recorded: ${item.action.replace(/_/g, ' ')}`],
  }
}

function parseDiffItem(raw: any) {
  if (!raw) return { rawText: '' }

  let label = ''
  let oldVal = ''
  let newVal = ''
  let isTotal = false
  let rawText = ''

  if (typeof raw === 'object' && raw.label && raw.oldVal !== undefined && raw.newVal !== undefined) {
    label = String(raw.label)
    oldVal = String(raw.oldVal)
    newVal = String(raw.newVal)
    isTotal = raw.key === 'total' || /total/i.test(label)
  } else {
    rawText = typeof raw === 'string' ? raw : raw.diff || String(raw)
    rawText = cleanDiffText(rawText).replace(/^[✎•\-\*\s]+/, '').trim()

    // Match "Label: OldVal → NewVal" or "Label: OldVal -> NewVal"
    const m = rawText.match(/^([^:]+):\s*(.*?)\s*(?:→|->)\s*(.*?)$/)
    if (m) {
      label = m[1].trim()
      oldVal = m[2].trim()
      newVal = m[3].trim()
      isTotal = /total/i.test(label)
    } else {
      // Match "OldVal → NewVal"
      const m2 = rawText.match(/^(.*?)\s*(?:→|->)\s*(.*?)$/)
      if (m2) {
        label = 'Value'
        oldVal = m2[1].trim()
        newVal = m2[2].trim()
      }
    }
  }

  if (oldVal && newVal) {
    let scope = ''
    let cleanLabel = label
    const scopeMatch = label.match(/^\[([^\]]+)\]\s*(.+)$/)
    if (scopeMatch) {
      scope = scopeMatch[1].trim()
      cleanLabel = scopeMatch[2].trim()
    }
    let delta = undefined
    if (isTotal || /total/i.test(cleanLabel) || (oldVal.startsWith('₹') && newVal.startsWith('₹'))) {
      const oNum = Number(oldVal.replace(/[^0-9.-]+/g, ''))
      const nNum = Number(newVal.replace(/[^0-9.-]+/g, ''))
      if (!isNaN(oNum) && !isNaN(nNum) && oNum !== nNum) {
        const diffNum = nNum - oNum
        delta = {
          diffNum,
          formatted: (diffNum > 0 ? '+₹' : '-₹') + Math.abs(Math.round(diffNum)).toLocaleString('en-IN'),
          isIncrease: diffNum > 0,
        }
      }
    }
    return { scope, label: cleanLabel, oldVal, newVal, isTotal, delta }
  }

  return { scope: '', rawText: rawText || String(raw) }
}

function DiffItemRenderer({ item }: { item: any }) {
  const parsed = parseDiffItem(item)

  if (parsed.oldVal !== undefined && parsed.newVal !== undefined) {
    return (
      <tr className={parsed.isTotal ? 'clean-diff-total-row' : ''}>
        <td>
          <span className={`clean-scope-pill ${parsed.scope === 'Edit PDF' ? 'pdf' : parsed.scope ? 'sys' : 'gen'}`}>
            {parsed.scope || 'General'}
          </span>
        </td>
        <td className="clean-diff-field">{parsed.label}</td>
        <td>
          <span className="clean-val-old" title="Previous value">{parsed.oldVal}</span>
        </td>
        <td className="clean-arrow-cell">→</td>
        <td>
          <span className="clean-val-new" title="New value">{parsed.newVal}</span>
          {parsed.delta && (
            <span className={`clean-delta-badge ${parsed.delta.isIncrease ? 'pos' : 'neg'}`}>
              {parsed.delta.formatted}
            </span>
          )}
        </td>
      </tr>
    )
  }

  return (
    <tr>
      <td>
        <span className="clean-scope-pill gen">Info</span>
      </td>
      <td colSpan={4} style={{ color: '#334155', fontSize: '12px' }}>
        {parsed.rawText}
      </td>
    </tr>
  )
}
