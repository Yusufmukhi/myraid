import React, { useEffect, useMemo, useState } from 'react'
import { ActivityItem, api, Member, Org, Project, Session } from '../../lib/api'
import { inr, timeAgo } from '../../lib/format'
import { Icon, Spinner, useUI } from '../ui'

interface Props {
  session: Session
  org: Org
  onOpenProject: (projectId: string) => void
  onNavigate?: (tab: string) => void
}

type TimeRange = 'overall' | 'year' | 'month'

interface StaffMetrics {
  userIdOrName: string
  displayName: string
  role: string
  joinedAt?: string
  totalCount: number
  totalValue: number
  approvedCount: number
  approvedValue: number
  approvedPct: number
  lostCount: number
  lostValue: number
  lostPct: number
  negotiatingCount: number
  negotiatingValue: number
  negotiatingPct: number
  quotedCount: number
  quotedValue: number
  projects: Project[]
}

export function StaffPerformancePage({ session, org, onOpenProject, onNavigate }: Props) {
  const { toast } = useUI()
  const role = org?.role || 'member'
  const isAdminOrOwner = role === 'owner' || role === 'admin'

  const [loading, setLoading] = useState(true)
  const [projects, setProjects] = useState<Project[]>([])
  const [members, setMembers] = useState<Member[]>([])
  const [activities, setActivities] = useState<ActivityItem[]>([])
  const [timeRange, setTimeRange] = useState<TimeRange>('overall')
  const [selectedYear, setSelectedYear] = useState<number>(() => new Date().getFullYear())
  const [selectedMonth, setSelectedMonth] = useState<string>(() => new Date().toISOString().slice(0, 7))
  const [searchStaff, setSearchStaff] = useState('')
  const [sortBy, setSortBy] = useState<'quotations' | 'won_value' | 'win_rate' | 'name'>('quotations')

  // Selected staff for dedicated full-page view
  const [selectedStaffUsername, setSelectedStaffUsername] = useState<string | null>(null)
  const [staffTab, setStaffTab] = useState<'quotations' | 'activity'>('quotations')
  const [staffQuotationSearch, setStaffQuotationSearch] = useState('')
  const [staffStatusFilter, setStaffStatusFilter] = useState<'all' | 'quoted' | 'negotiating' | 'won' | 'loss' | 'approved' | 'lost'>('all')

  // Available years and months for selection
  const availableYears = useMemo(() => {
    const current = new Date().getFullYear()
    return [current, current - 1, current - 2, current - 3]
  }, [])

  const availableMonths = useMemo(() => {
    const list: Array<{ value: string; label: string }> = []
    const now = new Date()
    for (let i = 0; i < 12; i++) {
      const d = new Date(now.getFullYear(), now.getMonth() - i, 1)
      const val = d.toISOString().slice(0, 7)
      const label = d.toLocaleDateString('en-US', { month: 'short', year: 'numeric' })
      list.push({ value: val, label })
    }
    return list
  }, [])

  const loadData = async () => {
    try {
      setLoading(true)
      const [projList, memberList, actList] = await Promise.all([
        api.projectList(session.token, org.id, true),
        api.orgMembers(session.token, org.id),
        api.activityList(session.token, org.id, 200).catch(() => []),
      ])
      setProjects(projList || [])
      setMembers(memberList || [])
      setActivities(actList || [])
    } catch (e) {
      toast((e as Error).message, 'err')
    } finally {
      setLoading(false)
    }
  }

  useEffect(() => {
    loadData()
  }, [org.id])

  // Filter projects by the selected date range
  const filteredProjects = useMemo(() => {
    if (timeRange === 'overall') return projects

    return (projects || []).filter((p) => {
      const d = p.created_at ? new Date(p.created_at) : null
      if (!d || isNaN(d.getTime())) return true

      if (timeRange === 'year') {
        return d.getFullYear() === selectedYear
      }
      if (timeRange === 'month') {
        return d.toISOString().slice(0, 7) === selectedMonth
      }
      return true
    })
  }, [projects, timeRange, selectedYear, selectedMonth])

  // Map of userId or username to Member details
  const memberMap = useMemo(() => {
    const map = new Map<string, Member>()
    for (const m of members) {
      if (m.username) map.set(m.username.toLowerCase(), m)
      if (m.id) map.set(m.id, m)
    }
    return map
  }, [members])

  // Group and compute staff performance metrics
  const staffList = useMemo<StaffMetrics[]>(() => {
    const grouped = new Map<string, Project[]>()

    // Ensure every known staff member (excluding admin and owner) has an entry
    for (const m of members) {
      if (m.role === 'owner' || m.role === 'admin') continue
      const key = m.username
      if (!grouped.has(key)) {
        grouped.set(key, [])
      }
    }

    // Distribute filtered projects to creators (excluding admin and owner creators)
    for (const p of filteredProjects) {
      const creatorKey = p.created_by || 'Unknown'
      const matchedMember = memberMap.get(creatorKey.toLowerCase()) || memberMap.get(creatorKey)
      if (matchedMember?.role === 'owner' || matchedMember?.role === 'admin') continue
      const effectiveKey = matchedMember?.username || creatorKey

      const existing = grouped.get(effectiveKey) || []
      existing.push(p)
      grouped.set(effectiveKey, existing)
    }

    const result: StaffMetrics[] = []

    grouped.forEach((projArr, key) => {
      const matchedMember = memberMap.get(key.toLowerCase())
      const roleStr = matchedMember?.role || 'member'
      const joinedDate = matchedMember?.joined_at

      let approvedCount = 0
      let approvedValue = 0
      let lostCount = 0
      let lostValue = 0
      let negotiatingCount = 0
      let negotiatingValue = 0
      let quotedCount = 0
      let quotedValue = 0
      let totalValue = 0

      for (const p of projArr) {
        const val = Number(p.total) || 0
        totalValue += val
        const isLoss = p.is_archived || p.status === 'loss' || p.status === 'archived' || p.status === 'lost'
        const isWon = !p.is_archived && (p.status === 'won' || p.status === 'approved')
        const isNegotiating = p.status === 'negotiating'

        if (isWon) {
          approvedCount++
          approvedValue += val
        } else if (isLoss) {
          lostCount++
          lostValue += val
        } else if (isNegotiating) {
          negotiatingCount++
          negotiatingValue += val
        } else {
          quotedCount++
          quotedValue += val
        }
      }

      const totalCount = projArr.length
      const approvedPct = totalCount > 0 ? (approvedCount / totalCount) * 100 : 0
      const lostPct = totalCount > 0 ? (lostCount / totalCount) * 100 : 0
      const negotiatingPct = totalCount > 0 ? (negotiatingCount / totalCount) * 100 : 0

      result.push({
        userIdOrName: key,
        displayName: key,
        role: roleStr,
        joinedAt: joinedDate,
        totalCount,
        totalValue,
        approvedCount,
        approvedValue,
        approvedPct,
        lostCount,
        lostValue,
        lostPct,
        negotiatingCount,
        negotiatingValue,
        negotiatingPct,
        quotedCount,
        quotedValue,
        projects: projArr,
      })
    })

    // Filter by search query
    const q = searchStaff.trim().toLowerCase()
    const filtered = q
      ? result.filter(
          (s) =>
            s.displayName.toLowerCase().includes(q) ||
            s.role.toLowerCase().includes(q)
        )
      : result

    // Sort staff
    return filtered.sort((a, b) => {
      if (sortBy === 'quotations') return b.totalCount - a.totalCount
      if (sortBy === 'won_value') return b.approvedValue - a.approvedValue
      if (sortBy === 'win_rate') return b.approvedPct - a.approvedPct
      if (sortBy === 'name') return a.displayName.localeCompare(b.displayName)
      return 0
    })
  }, [members, filteredProjects, memberMap, searchStaff, sortBy])

  // Selected individual staff metrics for dedicated full page
  const selectedStaff = useMemo(() => {
    if (!selectedStaffUsername) return null
    return staffList.find((s) => s.displayName.toLowerCase() === selectedStaffUsername.toLowerCase()) || null
  }, [selectedStaffUsername, staffList])

  // Filter activities for the selected staff member
  const staffActivities = useMemo(() => {
    if (!selectedStaff) return []
    const targetName = selectedStaff.displayName.toLowerCase()
    const targetMember = memberMap.get(targetName)
    const targetId = targetMember?.id

    return (activities || []).filter((a) => {
      const uName = (a.user_name || '').toLowerCase()
      const uId = a.user_id
      return (
        uName === targetName ||
        (targetId && uId === targetId) ||
        (a.details && typeof a.details === 'object' && a.details.user === selectedStaff.displayName)
      )
    })
  }, [selectedStaff, activities, memberMap])

  // Overall totals across all staff
  const overallMetrics = useMemo(() => {
    let totalQuotations = 0
    let totalQuotedValue = 0
    let totalWonQuotations = 0
    let totalWonValue = 0
    let totalLostQuotations = 0
    let totalLostValue = 0
    let totalActiveQuoted = 0
    let totalActiveQuotedValue = 0

    for (const p of filteredProjects) {
      const val = Number(p.total) || 0
      totalQuotations++
      totalQuotedValue += val

      const isLoss = p.is_archived || p.status === 'archived' || p.status === 'lost'
      const isWon = !p.is_archived && p.status === 'approved'

      if (isWon) {
        totalWonQuotations++
        totalWonValue += val
      } else if (isLoss) {
        totalLostQuotations++
        totalLostValue += val
      } else {
        totalActiveQuoted++
        totalActiveQuotedValue += val
      }
    }

    const overallWinRate = totalQuotations > 0 ? ((totalWonQuotations / totalQuotations) * 100).toFixed(1) : '0'
    const overallLossRate = totalQuotations > 0 ? ((totalLostQuotations / totalQuotations) * 100).toFixed(1) : '0'

    return {
      totalQuotations,
      totalQuotedValue,
      totalWonQuotations,
      totalWonValue,
      totalLostQuotations,
      totalLostValue,
      totalActiveQuoted,
      totalActiveQuotedValue,
      overallWinRate,
      overallLossRate,
    }
  }, [filteredProjects])

  if (!isAdminOrOwner) {
    return (
      <div className="page-container" style={{ padding: '32px 24px', maxWidth: 800, margin: '0 auto' }}>
        <div className="card" style={{ padding: 32, textAlign: 'center', border: '1px solid #e5e7eb' }}>
          <div style={{ fontSize: 36, marginBottom: 12 }}>🔒</div>
          <h2 style={{ fontSize: 20, fontWeight: 700, marginBottom: 8, color: '#111827' }}>
            Access Restricted
          </h2>
          <p style={{ color: '#4b5563', fontSize: 14, maxWidth: 480, margin: '0 auto' }}>
            Only Organisation <strong>Owners</strong> and <strong>Administrators</strong> are authorized to view team staff quotations, values, and won/loss analytics.
          </p>
        </div>
      </div>
    )
  }

  if (loading) {
    return (
      <div className="page-loading" style={{ minHeight: 400, display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
        <Spinner size={32} />
      </div>
    )
  }

  // =========================================================================
  // VIEW 1: DEDICATED FULL PAGE FOR INDIVIDUAL STAFF MEMBER
  // =========================================================================
  if (selectedStaff) {
    const st = selectedStaff
    const rawProjects = st.projects
    const filteredStaffProjects = (rawProjects || []).filter((p) => {
      const q = staffQuotationSearch.trim().toLowerCase()
      const isLoss = p.is_archived || p.status === 'loss' || p.status === 'archived' || p.status === 'lost'
      const isWon = !p.is_archived && (p.status === 'won' || p.status === 'approved')
      const isNegotiating = p.status === 'negotiating'
      const isQuoted = p.status === 'quoted' || (!isLoss && !isWon && !isNegotiating)

      if ((staffStatusFilter === 'won' || staffStatusFilter === 'approved') && !isWon) return false
      if ((staffStatusFilter === 'loss' || staffStatusFilter === 'lost') && !isLoss) return false
      if (staffStatusFilter === 'negotiating' && !isNegotiating) return false
      if (staffStatusFilter === 'quoted' && !isQuoted) return false

      if (q) {
        const nameMatch = p.name.toLowerCase().includes(q)
        const buyerMatch = (p.summary?.client_info?.buyer_name || '').toLowerCase().includes(q)
        return nameMatch || buyerMatch
      }
      return true
    })

    return (
      <div className="page-container" style={{ padding: '20px 24px', maxWidth: 1320, margin: '0 auto' }}>
        {/* Navigation & Header Bar */}
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 18, flexWrap: 'wrap', gap: 12 }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
            <button
              type="button"
              className="btn ghost sm"
              onClick={() => setSelectedStaffUsername(null)}
              style={{ fontWeight: 700, display: 'inline-flex', alignItems: 'center', gap: 6, fontSize: 12, border: '1px solid #cbd5e1' }}
            >
              ← Back to All Staff
            </button>
            <span style={{ color: '#94a3b8' }}>/</span>
            <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
              <strong style={{ fontSize: 16, color: '#0f172a' }}>{st.displayName}</strong>
              <span className={`status-pill status-${st.role}`} style={{ fontSize: 10, padding: '1px 6px' }}>{st.role}</span>
            </div>
          </div>

          {/* Time Range Selector for Staff Member */}
          <div style={{ display: 'flex', alignItems: 'center', gap: 6, background: '#f1f5f9', padding: 3, borderRadius: 6 }}>
            <button
              type="button"
              className={`btn sm ${timeRange === 'overall' ? 'primary' : 'ghost'}`}
              onClick={() => setTimeRange('overall')}
              style={{ padding: '3px 10px', fontSize: 11, fontWeight: 700 }}
            >
              Overall (All Time)
            </button>
            <button
              type="button"
              className={`btn sm ${timeRange === 'year' ? 'primary' : 'ghost'}`}
              onClick={() => setTimeRange('year')}
              style={{ padding: '3px 10px', fontSize: 11, fontWeight: 700 }}
            >
              By Year
            </button>
            <button
              type="button"
              className={`btn sm ${timeRange === 'month' ? 'primary' : 'ghost'}`}
              onClick={() => setTimeRange('month')}
              style={{ padding: '3px 10px', fontSize: 11, fontWeight: 700 }}
            >
              By Month
            </button>

            {timeRange === 'year' && (
              <select
                className="select sm"
                value={selectedYear}
                onChange={(e) => setSelectedYear(Number(e.target.value))}
                style={{ fontSize: 11.5, padding: '2px 6px', fontWeight: 600, borderRadius: 4 }}
              >
                {availableYears.map((y) => (
                  <option key={y} value={y}>Year {y}</option>
                ))}
              </select>
            )}

            {timeRange === 'month' && (
              <select
                className="select sm"
                value={selectedMonth}
                onChange={(e) => setSelectedMonth(e.target.value)}
                style={{ fontSize: 11.5, padding: '2px 6px', fontWeight: 600, borderRadius: 4 }}
              >
                {availableMonths.map((m) => (
                  <option key={m.value} value={m.value}>{m.label}</option>
                ))}
              </select>
            )}
          </div>
        </div>

        {/* 2-Column Responsive Dashboard Layout */}
        <div className="staff-detail-grid">
          {/* Left Column: Staff Performance Summary Card */}
          <div className="staff-sidebar-card">
            {/* Avatar & Profile */}
            <div style={{ display: 'flex', alignItems: 'center', gap: 12, marginBottom: 16 }}>
              <div
                style={{
                  width: 48,
                  height: 48,
                  borderRadius: '50%',
                  background: '#0f172a',
                  color: '#ffffff',
                  display: 'flex',
                  alignItems: 'center',
                  justifyContent: 'center',
                  fontWeight: 800,
                  fontSize: 18,
                  flexShrink: 0,
                  boxShadow: '0 2px 5px rgba(0,0,0,0.1)',
                }}
              >
                {(st.displayName[0] || 'U').toUpperCase()}
              </div>
              <div style={{ minWidth: 0 }}>
                <div style={{ fontWeight: 800, fontSize: 16, color: '#0f172a', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                  {st.displayName}
                </div>
                <div style={{ fontSize: 11.5, color: '#64748b', marginTop: 2 }}>
                  {org.name}
                </div>
              </div>
            </div>

            {/* Quick Rates (Win Rate vs Loss Rate) */}
            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 8, padding: '10px 12px', background: '#f8fafc', borderRadius: 8, border: '1px solid #f1f5f9', marginBottom: 14 }}>
              <div>
                <div style={{ fontSize: 10, fontWeight: 700, color: '#166534', textTransform: 'uppercase', letterSpacing: '0.04em' }}>
                  Win Rate
                </div>
                <div style={{ fontSize: 18, fontWeight: 800, color: '#15803d', marginTop: 2 }}>
                  +{st.approvedPct.toFixed(1)}%
                </div>
              </div>
              <div style={{ borderLeft: '1px solid #e2e8f0', paddingLeft: 10 }}>
                <div style={{ fontSize: 10, fontWeight: 700, color: '#991b1b', textTransform: 'uppercase', letterSpacing: '0.04em' }}>
                  Loss Rate
                </div>
                <div style={{ fontSize: 18, fontWeight: 800, color: '#dc2626', marginTop: 2 }}>
                  {st.lostPct.toFixed(1)}%
                </div>
              </div>
            </div>

            {/* Visual Conversion Progress Bar */}
            {st.totalCount > 0 && (
              <div style={{ marginBottom: 16 }}>
                <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: 10.5, fontWeight: 600, color: '#64748b', marginBottom: 5 }}>
                  <span>Conversion Ratio</span>
                  <span>Won vs Loss</span>
                </div>
                <div style={{ height: 8, width: '100%', background: '#f1f5f9', borderRadius: 999, overflow: 'hidden', display: 'flex' }}>
                  <div
                    style={{ width: `${st.approvedPct}%`, background: '#10b981', height: '100%' }}
                    title={`Won: ${st.approvedCount} (${st.approvedPct.toFixed(1)}%)`}
                  />
                  <div
                    style={{ width: `${st.lostPct}%`, background: '#ef4444', height: '100%' }}
                    title={`Loss: ${st.lostCount} (${st.lostPct.toFixed(1)}%)`}
                  />
                  <div
                    style={{
                      width: `${Math.max(0, 100 - st.approvedPct - st.lostPct)}%`,
                      background: '#f59e0b',
                      height: '100%',
                    }}
                    title={`Quoted: ${st.quotedCount}`}
                  />
                </div>
              </div>
            )}

            {/* Compact Executive KPI Metric List */}
            <div style={{ fontSize: 11, fontWeight: 700, textTransform: 'uppercase', color: '#64748b', letterSpacing: '0.05em', marginBottom: 8 }}>
              Performance Metrics
            </div>
            <div className="staff-kpi-compact">
              <div className="staff-kpi-item">
                <div>
                  <div style={{ fontSize: 11, fontWeight: 700, color: '#475569', textTransform: 'uppercase' }}>Total Quotations</div>
                  <div style={{ fontSize: 11, color: '#64748b', marginTop: 1 }}>{inr(st.totalValue)}</div>
                </div>
                <strong style={{ fontSize: 20, color: '#0f172a' }}>{st.totalCount}</strong>
              </div>

              <div className="staff-kpi-item won-accent">
                <div>
                  <div style={{ fontSize: 11, fontWeight: 700, color: '#166534', textTransform: 'uppercase' }}>Won</div>
                  <div style={{ fontSize: 11, color: '#166534', marginTop: 1 }}>Won: {inr(st.approvedValue)}</div>
                </div>
                <div style={{ textAlign: 'right' }}>
                  <strong style={{ fontSize: 20, color: '#15803d' }}>{st.approvedCount}</strong>
                  <div style={{ fontSize: 10, fontWeight: 700, color: '#166534' }}>{st.approvedPct.toFixed(1)}%</div>
                </div>
              </div>

              <div className="staff-kpi-item loss-accent">
                <div>
                  <div style={{ fontSize: 11, fontWeight: 700, color: '#991b1b', textTransform: 'uppercase' }}>Loss</div>
                  <div style={{ fontSize: 11, color: '#991b1b', marginTop: 1 }}>Lost: {inr(st.lostValue)}</div>
                </div>
                <div style={{ textAlign: 'right' }}>
                  <strong style={{ fontSize: 20, color: '#dc2626' }}>{st.lostCount}</strong>
                  <div style={{ fontSize: 10, fontWeight: 700, color: '#dc2626' }}>{st.lostPct.toFixed(1)}%</div>
                </div>
              </div>

              <div className="staff-kpi-item negotiating-accent">
                <div>
                  <div style={{ fontSize: 11, fontWeight: 700, color: '#1d4ed8', textTransform: 'uppercase' }}>Negotiating</div>
                  <div style={{ fontSize: 11, color: '#1d4ed8', marginTop: 1 }}>Negotiating: {inr(st.negotiatingValue)}</div>
                </div>
                <div style={{ textAlign: 'right' }}>
                  <strong style={{ fontSize: 20, color: '#2563eb' }}>{st.negotiatingCount}</strong>
                  <div style={{ fontSize: 10, fontWeight: 700, color: '#1d4ed8' }}>{st.negotiatingPct.toFixed(1)}%</div>
                </div>
              </div>

              <div className="staff-kpi-item quoted-accent">
                <div>
                  <div style={{ fontSize: 11, fontWeight: 700, color: '#b45309', textTransform: 'uppercase' }}>Quoted</div>
                  <div style={{ fontSize: 11, color: '#b45309', marginTop: 1 }}>Quoted: {inr(st.quotedValue)}</div>
                </div>
                <strong style={{ fontSize: 20, color: '#d97706' }}>{st.quotedCount}</strong>
              </div>
            </div>

            {st.joinedAt && (
              <div style={{ fontSize: 11, color: '#94a3b8', marginTop: 16, textAlign: 'center', borderTop: '1px solid #f1f5f9', paddingTop: 12 }}>
                Joined team on {new Date(st.joinedAt).toLocaleDateString()}
              </div>
            )}
          </div>

          {/* Right Column: Work Area (Quotations & Activity Log) */}
          <div style={{ display: 'flex', flexDirection: 'column', gap: 14 }}>
            {/* Top Workspace Tab Selector & Quotations Search */}
            <div
              className="card p-2"
              style={{
                display: 'flex',
                justifyContent: 'space-between',
                alignItems: 'center',
                flexWrap: 'wrap',
                gap: 10,
                background: '#ffffff',
                border: '1px solid #e2e8f0',
                borderRadius: 10,
              }}
            >
              <div style={{ display: 'flex', gap: 4, background: '#f1f5f9', padding: 3, borderRadius: 6 }}>
                <button
                  type="button"
                  className={`btn sm ${staffTab === 'quotations' ? 'primary' : 'ghost'}`}
                  onClick={() => setStaffTab('quotations')}
                  style={{ padding: '4px 14px', fontSize: 12, fontWeight: 700 }}
                >
                  📋 Quotations ({rawProjects.length})
                </button>
                <button
                  type="button"
                  className={`btn sm ${staffTab === 'activity' ? 'primary' : 'ghost'}`}
                  onClick={() => setStaffTab('activity')}
                  style={{ padding: '4px 14px', fontSize: 12, fontWeight: 700 }}
                >
                  🕒 Activity &amp; Audit Log ({staffActivities.length})
                </button>
              </div>

              {staffTab === 'quotations' && (
                <div style={{ display: 'flex', gap: 8, alignItems: 'center', flexWrap: 'wrap' }}>
                  <div style={{ display: 'flex', gap: 3, background: '#f1f5f9', padding: 2, borderRadius: 4 }}>
                    <button
                      type="button"
                      className={`btn sm ${staffStatusFilter === 'all' ? 'primary' : 'ghost'}`}
                      onClick={() => setStaffStatusFilter('all')}
                      style={{ padding: '2px 8px', fontSize: 11, fontWeight: 600 }}
                    >
                      All ({rawProjects.length})
                    </button>
                    <button
                      type="button"
                      className={`btn sm ${staffStatusFilter === 'quoted' ? 'primary' : 'ghost'}`}
                      onClick={() => setStaffStatusFilter('quoted')}
                      style={{ padding: '2px 8px', fontSize: 11, fontWeight: 600 }}
                    >
                      Quoted ({st.quotedCount})
                    </button>
                    <button
                      type="button"
                      className={`btn sm ${staffStatusFilter === 'negotiating' ? 'primary' : 'ghost'}`}
                      onClick={() => setStaffStatusFilter('negotiating')}
                      style={{ padding: '2px 8px', fontSize: 11, fontWeight: 600 }}
                    >
                      Negotiating ({st.negotiatingCount})
                    </button>
                    <button
                      type="button"
                      className={`btn sm ${staffStatusFilter === 'won' || staffStatusFilter === 'approved' ? 'primary' : 'ghost'}`}
                      onClick={() => setStaffStatusFilter('won')}
                      style={{ padding: '2px 8px', fontSize: 11, fontWeight: 600 }}
                    >
                      Won ({st.approvedCount})
                    </button>
                    <button
                      type="button"
                      className={`btn sm ${staffStatusFilter === 'loss' || staffStatusFilter === 'lost' ? 'primary' : 'ghost'}`}
                      onClick={() => setStaffStatusFilter('loss')}
                      style={{ padding: '2px 8px', fontSize: 11, fontWeight: 600 }}
                    >
                      Loss ({st.lostCount})
                    </button>
                  </div>

                  <div style={{ position: 'relative', width: 180 }}>
                    <input
                      type="text"
                      className="input sm"
                      placeholder="Search quotations…"
                      value={staffQuotationSearch}
                      onChange={(e) => setStaffQuotationSearch(e.target.value)}
                      style={{ width: '100%', paddingLeft: 24, fontSize: 11.5, boxSizing: 'border-box', height: 28 }}
                    />
                    <span style={{ position: 'absolute', left: 7, top: 6, color: '#9ca3af', fontSize: 11 }}>
                      🔍
                    </span>
                  </div>
                </div>
              )}
            </div>

            {/* Tab 1: Staff Quotations */}
            {staffTab === 'quotations' && (
              <div>
                {filteredStaffProjects.length === 0 ? (
                  <div className="card" style={{ padding: 40, textAlign: 'center', color: '#6b7280', border: '1px dashed #cbd5e1' }}>
                    No quotations match the filter for {st.displayName}.
                  </div>
                ) : (
                  <div className="card" style={{ border: '1px solid #e5e7eb', overflow: 'hidden', borderRadius: 10 }}>
                    <table className="tbl" style={{ width: '100%', margin: 0 }}>
                      <thead>
                        <tr style={{ background: '#f8fafc' }}>
                          <th style={{ textAlign: 'left', padding: '10px 14px' }}>Quotation &amp; Client</th>
                          <th style={{ textAlign: 'right', padding: '10px 14px' }}>Total Amount</th>
                          <th style={{ textAlign: 'center', padding: '10px 14px' }}>Status</th>
                          <th style={{ textAlign: 'center', padding: '10px 14px' }}>Created Date</th>
                          <th style={{ textAlign: 'right', padding: '10px 14px' }}>Action</th>
                        </tr>
                      </thead>
                      <tbody>
                        {filteredStaffProjects.map((p) => {
                          const buyer = p.summary?.client_info?.buyer_name || ''
                          const isLoss = p.is_archived || p.status === 'archived' || p.status === 'lost'
                          const isWon = !p.is_archived && p.status === 'approved'
                          const statusLabel = isWon ? 'APPROVED (WON)' : isLoss ? 'ARCHIVED (LOSS)' : 'QUOTED'
                          const statusClass = isWon ? 'approved' : isLoss ? 'lost' : 'quoted'

                          return (
                            <tr key={p.id} style={{ borderBottom: '1px solid #f1f5f9' }}>
                              <td style={{ padding: '10px 14px' }}>
                                <div style={{ display: 'flex', flexDirection: 'column', gap: 2 }}>
                                  <button
                                    type="button"
                                    className="link-inline proj-title-btn"
                                    onClick={() => onOpenProject(p.id)}
                                    style={{ fontWeight: 700, textAlign: 'left', color: '#0f172a', fontSize: 13.5 }}
                                  >
                                    {p.name}
                                  </button>
                                  {buyer && (
                                    <div style={{ fontSize: 11.5, color: '#64748b' }}>
                                      Client: <strong style={{ color: '#334155' }}>{buyer}</strong>
                                    </div>
                                  )}
                                </div>
                              </td>
                              <td style={{ padding: '10px 14px', textAlign: 'right', fontWeight: 800, fontSize: 14 }}>
                                {inr(Number(p.total) || 0)}
                              </td>
                              <td style={{ padding: '10px 14px', textAlign: 'center' }}>
                                <span className={`status-pill status-${statusClass}`}>
                                  {statusLabel}
                                </span>
                              </td>
                              <td style={{ padding: '10px 14px', textAlign: 'center', fontSize: 12, color: '#6b7280' }}>
                                {p.created_at ? new Date(p.created_at).toLocaleDateString() : '—'}
                              </td>
                              <td style={{ padding: '10px 14px', textAlign: 'right' }}>
                                <button
                                  type="button"
                                  className="btn ghost sm"
                                  onClick={() => onOpenProject(p.id)}
                                  style={{ padding: '3px 10px', fontSize: 11, fontWeight: 700 }}
                                >
                                  Open →
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

            {/* Tab 2: Staff Activity & Audit Log */}
            {staffTab === 'activity' && (
              <div>
                {staffActivities.length === 0 ? (
                  <div className="card" style={{ padding: 40, textAlign: 'center', color: '#6b7280', border: '1px dashed #cbd5e1' }}>
                    No recorded activity actions found for {st.displayName}.
                  </div>
                ) : (
                  <div className="card" style={{ border: '1px solid #e5e7eb', overflow: 'hidden', borderRadius: 10 }}>
                    <table className="tbl" style={{ width: '100%', margin: 0 }}>
                      <thead>
                        <tr style={{ background: '#f8fafc' }}>
                          <th style={{ textAlign: 'left', padding: '10px 14px' }}>Time</th>
                          <th style={{ textAlign: 'left', padding: '10px 14px' }}>Action</th>
                          <th style={{ textAlign: 'left', padding: '10px 14px' }}>Details / Quotation</th>
                          <th style={{ textAlign: 'right', padding: '10px 14px' }}>Action</th>
                        </tr>
                      </thead>
                      <tbody>
                        {staffActivities.map((act) => {
                          const actionReadable = act.action.replace(/_/g, ' ').toUpperCase()
                          const projName = act.details?.name || act.details?.project_name || act.details?.from_name || ''
                          const totalAmt = act.details?.total ? inr(act.details.total) : null

                          return (
                            <tr key={act.id} style={{ borderBottom: '1px solid #f1f5f9' }}>
                              <td style={{ padding: '10px 14px', whiteSpace: 'nowrap', fontSize: 12, color: '#6b7280' }}>
                                <span style={{ fontWeight: 600, color: '#111827', display: 'block' }}>
                                  {timeAgo(act.created_at)}
                                </span>
                                <span style={{ fontSize: 11 }}>
                                  {new Date(act.created_at).toLocaleString()}
                                </span>
                              </td>
                              <td style={{ padding: '10px 14px' }}>
                                <span
                                  style={{
                                    fontSize: 10,
                                    fontWeight: 700,
                                    padding: '2px 6px',
                                    borderRadius: 4,
                                    background: '#f1f5f9',
                                    color: '#334155',
                                    border: '1px solid #cbd5e1',
                                    textTransform: 'uppercase',
                                  }}
                                >
                                  {actionReadable}
                                </span>
                              </td>
                              <td style={{ padding: '10px 14px', fontSize: 13, color: '#334155' }}>
                                {projName && (
                                  <strong style={{ color: '#111827', marginRight: 6 }}>
                                    {projName}
                                  </strong>
                                )}
                                {totalAmt && (
                                  <span style={{ color: '#047857', fontWeight: 600, marginRight: 6 }}>
                                    ({totalAmt})
                                  </span>
                                )}
                                {act.details?.note && (
                                  <span style={{ color: '#64748b', fontSize: 12 }}>
                                    “{act.details.note}”
                                  </span>
                                )}
                              </td>
                              <td style={{ padding: '10px 14px', textAlign: 'right' }}>
                                {act.project_id && (
                                  <button
                                    type="button"
                                    className="btn ghost sm"
                                    onClick={() => onOpenProject(act.project_id!)}
                                    style={{ padding: '2px 8px', fontSize: 11, fontWeight: 600 }}
                                  >
                                    View Quotation →
                                  </button>
                                )}
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
        </div>
      </div>
    )
  }

  // =========================================================================
  // VIEW 2: OVERVIEW OF ALL STAFF MEMBERS
  // =========================================================================
  return (
    <div className="page-container" style={{ padding: '24px 32px', maxWidth: 1200, margin: '0 auto' }}>
      {/* Top Header */}
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', marginBottom: 20, gap: 16, flexWrap: 'wrap' }}>
        <div>
          <div style={{ display: 'flex', alignItems: 'center', gap: 10, marginBottom: 4 }}>
            <h1 style={{ fontSize: 24, fontWeight: 700, margin: 0, color: '#111827' }}>
              Staff Quotations &amp; Analytics
            </h1>
            <span
              style={{
                fontSize: 11,
                fontWeight: 700,
                padding: '2px 8px',
                borderRadius: 4,
                border: '1px solid #111827',
                background: '#f9fafb',
                color: '#111827',
                textTransform: 'uppercase',
              }}
            >
              {role.toUpperCase()} ONLY
            </span>
          </div>
          <p style={{ margin: 0, fontSize: 13, color: '#4b5563' }}>
            Click on any staff member to open their complete dedicated profile, quotations list, and audit activity trail.
          </p>
        </div>

        {/* Date Range Selector Pills */}
        <div style={{ display: 'flex', alignItems: 'center', gap: 8, background: '#f3f4f6', padding: 4, borderRadius: 6, border: '1px solid #e5e7eb' }}>
          <button
            type="button"
            className={`btn sm ${timeRange === 'overall' ? 'primary' : 'ghost'}`}
            onClick={() => setTimeRange('overall')}
            style={{ fontWeight: 600, padding: '4px 12px' }}
          >
            Overall (All Time)
          </button>
          <button
            type="button"
            className={`btn sm ${timeRange === 'year' ? 'primary' : 'ghost'}`}
            onClick={() => setTimeRange('year')}
            style={{ fontWeight: 600, padding: '4px 12px' }}
          >
            By Year
          </button>
          <button
            type="button"
            className={`btn sm ${timeRange === 'month' ? 'primary' : 'ghost'}`}
            onClick={() => setTimeRange('month')}
            style={{ fontWeight: 600, padding: '4px 12px' }}
          >
            By Month
          </button>

          {/* Year selector dropdown */}
          {timeRange === 'year' && (
            <select
              className="select sm"
              value={selectedYear}
              onChange={(e) => setSelectedYear(Number(e.target.value))}
              style={{ marginLeft: 6, padding: '3px 8px', fontSize: 12, fontWeight: 600 }}
            >
              {availableYears.map((y) => (
                <option key={y} value={y}>
                  Year {y}
                </option>
              ))}
            </select>
          )}

          {/* Month selector dropdown */}
          {timeRange === 'month' && (
            <select
              className="select sm"
              value={selectedMonth}
              onChange={(e) => setSelectedMonth(e.target.value)}
              style={{ marginLeft: 6, padding: '3px 8px', fontSize: 12, fontWeight: 600 }}
            >
              {availableMonths.map((m) => (
                <option key={m.value} value={m.value}>
                  {m.label}
                </option>
              ))}
            </select>
          )}
        </div>
      </div>

      {/* High Level KPI Overview Cards */}
      <div className="kpis">
        <div className="kpi primary">
          <span>Total Quotations</span>
          <strong>{overallMetrics.totalQuotations}</strong>
          <small>Total Pipeline: {inr(overallMetrics.totalQuotedValue)}</small>
        </div>

        <div className="kpi kpi-won">
          <span style={{ color: '#166534' }}>Won</span>
          <strong style={{ color: '#15803d' }}>
            {overallMetrics.totalWonQuotations}
            <span className="metric-pill won">+{overallMetrics.overallWinRate}%</span>
          </strong>
          <small style={{ color: '#166534' }}>Won: {inr(overallMetrics.totalWonValue)}</small>
        </div>

        <div className="kpi kpi-loss">
          <span style={{ color: '#991b1b' }}>Loss</span>
          <strong style={{ color: '#dc2626' }}>
            {overallMetrics.totalLostQuotations}
            <span className="metric-pill loss">{overallMetrics.overallLossRate}%</span>
          </strong>
          <small style={{ color: '#991b1b' }}>Lost: {inr(overallMetrics.totalLostValue)}</small>
        </div>

        <div className="kpi kpi-quoted">
          <span style={{ color: '#b45309' }}>Quoted / In Progress</span>
          <strong style={{ color: '#d97706' }}>
            {overallMetrics.totalActiveQuoted}
          </strong>
          <small style={{ color: '#b45309' }}>Quoted: {inr(overallMetrics.totalActiveQuotedValue)}</small>
        </div>
      </div>

      {/* Conversion Rate Visual Bar */}
      {overallMetrics.totalQuotations > 0 && (
        <div className="card" style={{ padding: '14px 18px', marginBottom: 24, border: '1px solid #e5e7eb' }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 8 }}>
            <span style={{ fontSize: 12, fontWeight: 700, textTransform: 'uppercase', color: '#374151' }}>
              Quotation Conversion Overview ({timeRange.toUpperCase()})
            </span>
            <span style={{ fontSize: 12, color: '#4b5563' }}>
              Win Rate: <strong>{overallMetrics.overallWinRate}%</strong> | Loss Rate: <strong>{overallMetrics.overallLossRate}%</strong> ({overallMetrics.totalActiveQuoted} Quoted)
            </span>
          </div>

          <div style={{ height: 12, width: '100%', background: '#f3f4f6', borderRadius: 999, overflow: 'hidden', display: 'flex' }}>
            <div
              style={{
                width: `${overallMetrics.overallWinRate}%`,
                background: '#10b981',
                height: '100%',
                transition: 'width 0.3s ease',
              }}
              title={`Won: ${overallMetrics.totalWonQuotations} (${overallMetrics.overallWinRate}%)`}
            />
            <div
              style={{
                width: `${overallMetrics.overallLossRate}%`,
                background: '#ef4444',
                height: '100%',
                transition: 'width 0.3s ease',
              }}
              title={`Loss: ${overallMetrics.totalLostQuotations} (${overallMetrics.overallLossRate}%)`}
            />
            <div
              style={{
                width: `${Math.max(0, 100 - Number(overallMetrics.overallWinRate) - Number(overallMetrics.overallLossRate))}%`,
                background: '#f59e0b',
                height: '100%',
                transition: 'width 0.3s ease',
              }}
              title={`Quoted: ${overallMetrics.totalActiveQuoted}`}
            />
          </div>

          <div style={{ display: 'flex', gap: 20, marginTop: 10, fontSize: 12, color: '#4b5563' }}>
            <span style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
              <span style={{ width: 10, height: 10, borderRadius: 2, background: '#10b981', display: 'inline-block' }} />
              Won: {overallMetrics.totalWonQuotations} ({overallMetrics.overallWinRate}%)
            </span>
            <span style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
              <span style={{ width: 10, height: 10, borderRadius: 2, background: '#ef4444', display: 'inline-block' }} />
              Losses: {overallMetrics.totalLostQuotations} ({overallMetrics.overallLossRate}%)
            </span>
            <span style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
              <span style={{ width: 10, height: 10, borderRadius: 2, background: '#f59e0b', display: 'inline-block' }} />
              Quoted / In Progress: {overallMetrics.totalActiveQuoted}
            </span>
          </div>
        </div>
      )}

      {/* Staff Table Filter & Controls */}
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 16, gap: 12, flexWrap: 'wrap' }}>
        <div style={{ position: 'relative', width: 280 }}>
          <input
            type="text"
            className="input"
            placeholder="Search staff member…"
            value={searchStaff}
            onChange={(e) => setSearchStaff(e.target.value)}
            style={{ width: '100%', paddingLeft: 30, boxSizing: 'border-box' }}
          />
          <span style={{ position: 'absolute', left: 10, top: 8, color: '#9ca3af' }}>
            🔍
          </span>
        </div>

        <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
          <span style={{ fontSize: 12, fontWeight: 600, color: '#6b7280' }}>Sort by:</span>
          <select
            className="select sm"
            value={sortBy}
            onChange={(e) => setSortBy(e.target.value as any)}
            style={{ padding: '4px 8px', fontSize: 12 }}
          >
            <option value="quotations">Most Quotations</option>
            <option value="won_value">Highest Won Value</option>
            <option value="win_rate">Highest Win Rate %</option>
            <option value="name">Staff Name (A-Z)</option>
          </select>
        </div>
      </div>

      {/* Staff Performance Cards List */}
      <div style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
        {staffList.length === 0 ? (
          <div className="card" style={{ padding: 40, textAlign: 'center', color: '#6b7280' }}>
            No staff quotations found for the selected time range.
          </div>
        ) : (
          staffList.map((st) => {
            return (
              <div
                key={st.displayName}
                className="card"
                style={{
                  border: '1px solid #e5e7eb',
                  background: '#ffffff',
                  borderRadius: 6,
                  overflow: 'hidden',
                  transition: 'box-shadow 0.2s',
                  cursor: 'pointer',
                }}
                onClick={() => setSelectedStaffUsername(st.displayName)}
              >
                {/* Staff Summary Header Row */}
                <div
                  style={{
                    padding: '16px 20px',
                    display: 'grid',
                    gridTemplateColumns: 'minmax(180px, 2fr) minmax(100px, 1fr) minmax(140px, 1.5fr) minmax(140px, 1.5fr) minmax(120px, 1fr) 140px',
                    alignItems: 'center',
                    gap: 16,
                  }}
                >
                  {/* Staff Identity */}
                  <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
                    <div
                      style={{
                        width: 36,
                        height: 36,
                        borderRadius: '50%',
                        background: '#1e3a8a',
                        color: '#ffffff',
                        display: 'flex',
                        alignItems: 'center',
                        justifyContent: 'center',
                        fontWeight: 700,
                        fontSize: 14,
                        flexShrink: 0,
                      }}
                    >
                      {(st.displayName[0] || 'U').toUpperCase()}
                    </div>
                    <div>
                      <div style={{ fontWeight: 700, fontSize: 14, color: '#111827' }}>
                        {st.displayName}
                      </div>
                      <span
                        style={{
                          fontSize: 10,
                          fontWeight: 700,
                          padding: '1px 6px',
                          borderRadius: 4,
                          background: '#f1f5f9',
                          color: '#475569',
                          textTransform: 'uppercase',
                        }}
                      >
                        {st.role}
                      </span>
                    </div>
                  </div>

                  {/* Total Quotations */}
                  <div>
                    <div style={{ fontSize: 11, color: '#6b7280', textTransform: 'uppercase', fontWeight: 600 }}>
                      Quotations
                    </div>
                    <div style={{ fontSize: 16, fontWeight: 800, color: '#111827' }}>
                      {st.totalCount}
                    </div>
                    <small style={{ fontSize: 11, color: '#6b7280' }}>
                      {inr(st.totalValue)}
                    </small>
                  </div>

                  {/* Won / Approved */}
                  <div>
                    <div style={{ fontSize: 11, color: '#166534', textTransform: 'uppercase', fontWeight: 600 }}>
                      Won
                    </div>
                    <div style={{ fontSize: 16, fontWeight: 800, color: '#15803d', display: 'flex', alignItems: 'center', gap: 6 }}>
                      {st.approvedCount}
                      <span className="metric-pill won">
                        +{st.approvedPct.toFixed(1)}%
                      </span>
                    </div>
                    <small style={{ fontSize: 11, color: '#166534' }}>
                      {inr(st.approvedValue)}
                    </small>
                  </div>

                  {/* Losses / Lost */}
                  <div>
                    <div style={{ fontSize: 11, color: '#991b1b', textTransform: 'uppercase', fontWeight: 600 }}>
                      Loss
                    </div>
                    <div style={{ fontSize: 16, fontWeight: 800, color: '#dc2626', display: 'flex', alignItems: 'center', gap: 6 }}>
                      {st.lostCount}
                      <span className="metric-pill loss">
                        {st.lostPct.toFixed(1)}%
                      </span>
                    </div>
                    <small style={{ fontSize: 11, color: '#991b1b' }}>
                      {inr(st.lostValue)}
                    </small>
                  </div>

                  {/* Quoted / Pending (NO percentage) */}
                  <div>
                    <div style={{ fontSize: 11, color: '#b45309', textTransform: 'uppercase', fontWeight: 600 }}>
                      Quoted
                    </div>
                    <div style={{ fontSize: 16, fontWeight: 800, color: '#d97706' }}>
                      {st.quotedCount}
                    </div>
                    <small style={{ fontSize: 11, color: '#b45309' }}>
                      {inr(st.quotedValue)}
                    </small>
                  </div>

                  {/* Open Staff Page Button */}
                  <div style={{ textAlign: 'right' }}>
                    <button
                      type="button"
                      className="btn primary sm"
                      style={{ fontSize: 11, fontWeight: 700, padding: '4px 10px' }}
                      onClick={(e) => {
                        e.stopPropagation()
                        setSelectedStaffUsername(st.displayName)
                      }}
                    >
                      View Staff Page →
                    </button>
                  </div>
                </div>

                {/* Progress breakdown bar for this staff */}
                <div style={{ height: 4, width: '100%', background: '#f3f4f6', display: 'flex' }}>
                  <div style={{ width: `${st.approvedPct}%`, background: '#10b981', height: '100%' }} />
                  <div style={{ width: `${st.lostPct}%`, background: '#ef4444', height: '100%' }} />
                  <div style={{ width: `${Math.max(0, 100 - st.approvedPct - st.lostPct)}%`, background: '#f59e0b', height: '100%' }} />
                </div>
              </div>
            )
          })
        )}
      </div>
    </div>
  )
}
