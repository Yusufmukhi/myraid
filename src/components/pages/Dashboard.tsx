import { useEffect, useMemo, useState } from 'react'
import { api, DashboardStats, Org, Project, Session } from '../../lib/api'
import { inr, timeAgo } from '../../lib/format'
import { parseRevision } from '../../lib/revision'
import { Icon, Spinner, useUI } from '../ui'

interface Props {
  session: Session
  org: Org
  onOpenProject: (projectId: string) => void
  onNewProject: () => void
  onNavigate: (tab: string) => void
}

export function Dashboard({ session, org, onOpenProject, onNewProject, onNavigate }: Props) {
  const { toast } = useUI()
  const isAdminOrOwner = org.role === 'owner' || org.role === 'admin'
  const [stats, setStats] = useState<DashboardStats | null>(null)
  const [myProjects, setMyProjects] = useState<Project[]>([])
  const [loading, setLoading] = useState(true)

  // Date range filtering for Owner & Admin (collapsed behind Filter button for minimal UI)
  type TimeRange = 'overall' | 'year' | 'month'
  const [timeRange, setTimeRange] = useState<TimeRange>('overall')
  const [analyticsFilterOpen, setAnalyticsFilterOpen] = useState(false)
  const [showCode, setShowCode] = useState(false)
  const [selectedYear, setSelectedYear] = useState<number>(() => new Date().getFullYear())
  const [selectedMonth, setSelectedMonth] = useState<string>(() => new Date().toISOString().slice(0, 7))

  const availableYears = useMemo(() => {
    const cur = new Date().getFullYear()
    return [cur, cur - 1, cur - 2, cur - 3]
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

  const loadStats = async () => {
    try {
      setLoading(true)
      const [data, projList] = await Promise.all([
        api.dashboardStats(session.token, org.id),
        api.projectList(session.token, org.id, true),
      ])
      setStats(data)
      const safeProjList = Array.isArray(projList) ? projList : []
      const userProjects = !isAdminOrOwner
        ? safeProjList.filter((p) => {
            const shared = (p.summary?.shared_with || (p as any).shared_with) as string[] | undefined
            const isShared =
              Array.isArray(shared) &&
              (shared.includes(session.username) || (session.user_id && shared.includes(session.user_id)))
            return (
              p.created_by === session.user_id ||
              p.created_by === session.username ||
              isShared
            )
          })
        : safeProjList
      setMyProjects(userProjects)
    } catch (e) {
      toast((e as Error).message, 'err')
    } finally {
      setLoading(false)
    }
  }

  useEffect(() => {
    loadStats()
  }, [org.id])

  const { createdCount, activeCount, sharedCount } = useMemo(() => {
    let created = 0
    let active = 0
    let shared = 0

    for (const p of myProjects) {
      const sharedList = (p.summary?.shared_with || (p as any).shared_with) as string[] | undefined
      const isSharedWithMe =
        Array.isArray(sharedList) &&
        (sharedList.includes(session.username) || (session.user_id ? sharedList.includes(session.user_id) : false)) &&
        p.created_by !== session.username &&
        p.created_by !== session.user_id

      const isCreatedByMe =
        p.created_by === session.username ||
        p.created_by === session.user_id

      if (isCreatedByMe) {
        created++
      }
      if (isSharedWithMe) {
        shared++
      }
      if (!p.is_archived) {
        active++
      }
    }

    return { createdCount: created, activeCount: active, sharedCount: shared }
  }, [myProjects, session.username, session.user_id])

  const staffStats = useMemo(() => {
    const createdByMe = (myProjects || []).filter(
      (p) => p.created_by === session.username || p.created_by === session.user_id
    )
    const myTotalVal = createdByMe.reduce((sum, p) => sum + (Number(p.total) || 0), 0)
    const wonCount = createdByMe.filter(
      (p) => !p.is_archived && (p.status === 'won' || p.status === 'approved')
    ).length
    const lossCount = createdByMe.filter(
      (p) => p.is_archived || p.status === 'loss' || p.status === 'archived' || p.status === 'lost'
    ).length
    const activeCount = createdByMe.filter(
      (p) => !p.is_archived && p.status !== 'won' && p.status !== 'approved' && p.status !== 'loss' && p.status !== 'archived' && p.status !== 'lost'
    ).length
    const totalDecided = wonCount + lossCount
    const winRate = totalDecided > 0 ? Math.round((wonCount / totalDecided) * 100) : 0
    const myReminders = createdByMe.filter((p) => p.summary?.reminder && !p.summary.reminder.dismissed)
    return {
      count: createdByMe.length,
      totalVal: myTotalVal,
      wonCount,
      lossCount,
      activeCount,
      winRate,
      remindersCount: myReminders.length,
    }
  }, [myProjects, session.username, session.user_id])

  const filteredDashboardProjects = useMemo(() => {
    if (!isAdminOrOwner || timeRange === 'overall') return myProjects

    return (myProjects || []).filter((p) => {
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
  }, [myProjects, isAdminOrOwner, timeRange, selectedYear, selectedMonth])

  const conversionStats = useMemo(() => {
    const total = filteredDashboardProjects.length
    let won = 0
    let lost = 0
    let negotiating = 0
    let quoted = 0
    let totalVal = 0
    let wonVal = 0
    let lostVal = 0
    let negotiatingVal = 0
    let quotedVal = 0

    for (const p of filteredDashboardProjects) {
      const val = Number(p.total) || 0
      totalVal += val
      const isLoss = p.is_archived || p.status === 'loss' || p.status === 'archived' || p.status === 'lost'
      const isWon = !p.is_archived && (p.status === 'won' || p.status === 'approved')
      const isNegotiating = p.status === 'negotiating'

      if (isWon) {
        won++
        wonVal += val
      } else if (isLoss) {
        lost++
        lostVal += val
      } else if (isNegotiating) {
        negotiating++
        negotiatingVal += val
      } else {
        quoted++
        quotedVal += val
      }
    }

    const decidedCount = won + lost
    const decidedVal = wonVal + lostVal
    const fmtPct = (num: number, den: number) => {
      if (den <= 0) return '0'
      const raw = (num / den) * 100
      return Number.isInteger(Math.round(raw * 10) / 10)
        ? String(Math.round(raw))
        : raw.toFixed(1)
    }
    const winPct = fmtPct(won, decidedCount)
    const lossPct = fmtPct(lost, decidedCount)
    const negotiatingPct = total > 0 ? ((negotiating / total) * 100).toFixed(1) : '0'

    return {
      total,
      won,
      lost,
      decidedCount,
      decidedVal,
      negotiating,
      quoted,
      approved: won,
      totalVal,
      wonVal,
      lostVal,
      negotiatingVal,
      quotedVal,
      approvedVal: wonVal,
      wonPct: winPct,
      approvedPct: winPct,
      winPct,
      lossPct,
      lostPct: lossPct,
      negotiatingPct,
    }
  }, [filteredDashboardProjects])

  if (loading && !stats) {
    return (
      <div className="page-loading">
        <Spinner size={24} />
      </div>
    )
  }

  const s = stats ?? {
    total_projects: 0,
    active_projects: 0,
    pipeline_value: 0,
    current_template_version: org.current_template_version || 1,
    recent_projects: [],
    recent_activity: [],
  }

  // Display projects: for staff, show accessible projects (created + shared); for admin, show recent org projects
  const displayedProjects = !isAdminOrOwner
    ? myProjects.slice(0, 5)
    : s.recent_projects

  const rangeLabel = timeRange === 'overall' ? 'All Time' : timeRange === 'year' ? `Year ${selectedYear}` : selectedMonth

  return (
    <div className="dashboard-page minimal-dashboard">
      {/* Compact Minimal Header Bar */}
      <div className="card minimal-toolbar-card">
        <div className="minimal-toolbar-row" style={{ justifyContent: 'space-between' }}>
          <div className="row gap" style={{ flexWrap: 'wrap', alignItems: 'center', gap: 8 }}>
            <h2 style={{ fontSize: '16px', fontWeight: 700, margin: 0, color: '#0f172a' }}>{org.name}</h2>
            {isAdminOrOwner && (
              <>
                <span
                  className="chip primary sm font-mono cursor-pointer"
                  onClick={() => setShowCode(!showCode)}
                  title="Click to reveal or hide organization join code"
                  style={{ fontSize: '11px' }}
                >
                  Code: {showCode ? org.code : '••••••••'}
                </span>
                <span className="version-pill" style={{ fontSize: '11px' }}>Default: v{s.current_template_version}</span>
              </>
            )}
          </div>

          <div className="row gap" style={{ flexWrap: 'wrap', alignItems: 'center', gap: 6 }}>
            {isAdminOrOwner && (
              <>
                <button
                  type="button"
                  className="btn ghost sm minimal-pill-btn"
                  onClick={() => setAnalyticsFilterOpen((o) => !o)}
                  title="Filter dashboard KPIs by Overall, Year, or Month"
                  style={{
                    borderColor: analyticsFilterOpen || timeRange !== 'overall' ? '#2563eb' : '#e2e8f0',
                    background: analyticsFilterOpen || timeRange !== 'overall' ? '#eff6ff' : '#ffffff',
                    color: analyticsFilterOpen || timeRange !== 'overall' ? '#1d4ed8' : '#334155',
                  }}
                >
                  <span>📅 {rangeLabel}</span>
                  <span style={{ fontSize: '9px', opacity: 0.7 }}>{analyticsFilterOpen ? '▲' : '▼'}</span>
                </button>
                <button
                  type="button"
                  className="btn ghost sm minimal-pill-btn"
                  onClick={() => onNavigate('staff_analytics')}
                >
                  👥 Staff Stats
                </button>
              </>
            )}
            <button className="btn ghost sm minimal-pill-btn" onClick={() => onNavigate('projects')}>
              {Icon.folder} {isAdminOrOwner ? 'Projects' : 'My Projects'}
            </button>
            <button className="btn primary sm" onClick={onNewProject} style={{ height: '32px', padding: '0 12px', fontSize: '12.5px' }}>
              {Icon.plus} New Project
            </button>
          </div>
        </div>

        {/* Collapsible Timeframe Box (Opens on Button Click) */}
        {isAdminOrOwner && analyticsFilterOpen && (
          <div className="minimal-filters-box" style={{ alignItems: 'center', justifyContent: 'space-between' }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: 10, flexWrap: 'wrap' }}>
              <span className="minimal-filter-label">Period</span>
              <div className="status-tabs">
                <button
                  type="button"
                  className={`status-tab ${timeRange === 'overall' ? 'active' : ''}`}
                  onClick={() => setTimeRange('overall')}
                >
                  Overall (All Time)
                </button>
                <button
                  type="button"
                  className={`status-tab ${timeRange === 'year' ? 'active' : ''}`}
                  onClick={() => setTimeRange('year')}
                >
                  By Year
                </button>
                <button
                  type="button"
                  className={`status-tab ${timeRange === 'month' ? 'active' : ''}`}
                  onClick={() => setTimeRange('month')}
                >
                  By Month
                </button>
              </div>

              {timeRange === 'year' && (
                <select
                  className="select-input compact"
                  value={selectedYear}
                  onChange={(e) => setSelectedYear(Number(e.target.value))}
                >
                  {availableYears.map((y) => (
                    <option key={y} value={y}>Year {y}</option>
                  ))}
                </select>
              )}

              {timeRange === 'month' && (
                <select
                  className="select-input compact"
                  value={selectedMonth}
                  onChange={(e) => setSelectedMonth(e.target.value)}
                >
                  {availableMonths.map((m) => (
                    <option key={m.value} value={m.value}>{m.label}</option>
                  ))}
                </select>
              )}
            </div>

            <button
              type="button"
              className="link-inline tiny"
              onClick={() => setAnalyticsFilterOpen(false)}
            >
              Done
            </button>
          </div>
        )}
      </div>

      {/* Compact Primary KPI Row */}
      <div className="kpis minimal-kpis">
        {isAdminOrOwner ? (
          <>
            <div className="kpi primary">
              <span>Total Quotations</span>
              <strong>{conversionStats.total}</strong>
              <small>Pipeline: {inr(conversionStats.totalVal)}</small>
            </div>
            <div className="kpi kpi-won">
              <span style={{ color: '#166534' }}>Won</span>
              <strong style={{ color: '#15803d' }}>
                {conversionStats.won}
                {conversionStats.decidedCount > 0 && (
                  <span className="metric-pill won">{conversionStats.winPct}% Win</span>
                )}
              </strong>
              <small style={{ color: '#166534' }}>{inr(conversionStats.wonVal)}</small>
            </div>
            <div className="kpi kpi-negotiating">
              <span style={{ color: '#1e40af' }}>Negotiating</span>
              <strong style={{ color: '#2563eb' }}>{conversionStats.negotiating}</strong>
              <small style={{ color: '#1e40af' }}>{inr(conversionStats.negotiatingVal)}</small>
            </div>
            <div className="kpi kpi-quoted">
              <span style={{ color: '#b45309' }}>Quoted</span>
              <strong style={{ color: '#d97706' }}>{conversionStats.quoted}</strong>
              <small style={{ color: '#b45309' }}>{inr(conversionStats.quotedVal)}</small>
            </div>
            <div className="kpi kpi-loss">
              <span style={{ color: '#991b1b' }}>Loss</span>
              <strong style={{ color: '#dc2626' }}>
                {conversionStats.lost}
                {conversionStats.decidedCount > 0 && (
                  <span className="metric-pill loss">{conversionStats.lossPct}% Loss</span>
                )}
              </strong>
              <small style={{ color: '#991b1b' }}>{inr(conversionStats.lostVal)}</small>
            </div>
          </>
        ) : (
          <>
            <div className="kpi primary">
              <span>Active Quotations</span>
              <strong>{activeCount}</strong>
              <small>Accessible Quotations</small>
            </div>
            <div className="kpi">
              <span>Quotations Created</span>
              <strong>{createdCount}</strong>
              <small>Created by You</small>
            </div>
            <div className="kpi">
              <span>Shared with Me</span>
              <strong style={{ color: '#7c3aed' }}>{sharedCount}</strong>
              <small>Shared Collaborations</small>
            </div>
          </>
        )}
      </div>

      {/* Staff Personal Sales Performance Dashboard */}
      {!isAdminOrOwner && (
        <div
          className="card staff-dashboard-card mb-3"
          style={{
            background: 'linear-gradient(135deg, #f8fafc 0%, #f1f5f9 100%)',
            border: '1px solid #cbd5e1',
            borderRadius: '10px',
            padding: '16px 20px',
            boxShadow: '0 2px 4px rgba(0,0,0,0.03)',
          }}
        >
          <div
            style={{
              display: 'flex',
              alignItems: 'center',
              justifyContent: 'space-between',
              marginBottom: '14px',
              flexWrap: 'wrap',
              gap: '8px',
            }}
          >
            <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
              <span style={{ fontSize: '22px' }}>📊</span>
              <div>
                <h3 style={{ margin: 0, fontSize: '16px', fontWeight: 800, color: '#0f172a' }}>
                  My Performance Dashboard
                </h3>
                <span className="tiny text-muted">Personal sales metrics & conversion for @{session.username}</span>
              </div>
            </div>
            <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
              <span
                className="chip primary sm font-bold"
                style={{ fontSize: '11px', background: '#e0f2fe', color: '#0369a1', border: '1px solid #bae6fd' }}
              >
                👤 Staff Member
              </span>
              <button
                type="button"
                className="btn ghost sm"
                onClick={() => onNavigate('projects')}
                style={{ fontSize: '11.5px', padding: '3px 10px', fontWeight: 600 }}
              >
                Go to Projects →
              </button>
            </div>
          </div>

          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(150px, 1fr))', gap: '12px' }}>
            {/* My Quotations */}
            <div
              className="kpi-mini-card"
              style={{
                background: '#ffffff',
                border: '1px solid #e2e8f0',
                borderRadius: '8px',
                padding: '12px 14px',
                cursor: 'pointer',
                transition: 'transform 0.1s ease',
              }}
              onClick={() => onNavigate('projects')}
              title="Click to view all my quotations in projects tab"
            >
              <span style={{ fontSize: '10.5px', fontWeight: 700, color: '#64748b', textTransform: 'uppercase', letterSpacing: '0.04em' }}>
                My Quotations
              </span>
              <div style={{ fontSize: '22px', fontWeight: 800, color: '#1e40af', marginTop: '2px' }}>
                {staffStats.count}
              </div>
              <span style={{ fontSize: '10px', color: '#0284c7', fontWeight: 600 }}>
                {staffStats.activeCount} currently active
              </span>
            </div>

            {/* Total Quotation Value */}
            <div
              className="kpi-mini-card"
              style={{
                background: '#ffffff',
                border: '1px solid #e2e8f0',
                borderRadius: '8px',
                padding: '12px 14px',
              }}
              title="Total combined quotation value created by me"
            >
              <span style={{ fontSize: '10.5px', fontWeight: 700, color: '#64748b', textTransform: 'uppercase', letterSpacing: '0.04em' }}>
                Total Quoted Value
              </span>
              <div style={{ fontSize: '19px', fontWeight: 800, color: '#16a34a', marginTop: '2px' }}>
                {inr(Math.round(staffStats.totalVal))}
              </div>
              <span style={{ fontSize: '10px', color: '#64748b' }}>
                across {staffStats.count} quotation{staffStats.count === 1 ? '' : 's'}
              </span>
            </div>

            {/* Won / Loss Ratio */}
            <div
              className="kpi-mini-card"
              style={{
                background: '#ffffff',
                border: '1px solid #e2e8f0',
                borderRadius: '8px',
                padding: '12px 14px',
              }}
              title="Win/Loss ratio of closed quotations"
            >
              <span style={{ fontSize: '10.5px', fontWeight: 700, color: '#64748b', textTransform: 'uppercase', letterSpacing: '0.04em' }}>
                Won / Loss Ratio
              </span>
              <div style={{ fontSize: '18px', fontWeight: 800, color: '#0f172a', marginTop: '2px', display: 'flex', alignItems: 'baseline', gap: '6px' }}>
                <span style={{ color: '#16a34a' }}>{staffStats.wonCount}W</span>
                <span style={{ color: '#cbd5e1' }}>/</span>
                <span style={{ color: '#dc2626' }}>{staffStats.lossCount}L</span>
                {staffStats.wonCount + staffStats.lossCount > 0 && (
                  <span style={{ fontSize: '11.5px', color: '#2563eb', fontWeight: 700 }}>
                    ({staffStats.winRate}%)
                  </span>
                )}
              </div>
              <span style={{ fontSize: '10px', color: '#64748b' }}>
                {staffStats.wonCount + staffStats.lossCount === 0 ? 'No closed quotations yet' : (staffStats.wonCount + staffStats.lossCount) + ' decided'}
              </span>
            </div>

            {/* Pending Follow-Ups */}
            <div
              className="kpi-mini-card"
              style={{
                background: staffStats.remindersCount > 0 ? '#fffbeb' : '#ffffff',
                border: '1px solid ' + (staffStats.remindersCount > 0 ? '#fde68a' : '#e2e8f0'),
                borderRadius: '8px',
                padding: '12px 14px',
                cursor: 'pointer',
              }}
              onClick={() => onNavigate('projects')}
              title="Click to view quotations with active reminders in projects tab"
            >
              <span style={{ fontSize: '10.5px', fontWeight: 700, color: staffStats.remindersCount > 0 ? '#b45309' : '#64748b', textTransform: 'uppercase', letterSpacing: '0.04em' }}>
                Pending Follow-Ups
              </span>
              <div style={{ fontSize: '20px', fontWeight: 800, color: staffStats.remindersCount > 0 ? '#d97706' : '#64748b', marginTop: '2px' }}>
                🔔 {staffStats.remindersCount}
              </div>
              <span style={{ fontSize: '10px', color: staffStats.remindersCount > 0 ? '#b45309' : '#94a3b8', fontWeight: 600 }}>
                {staffStats.remindersCount > 0 ? 'Active alarm reminders' : 'No pending alarms'}
              </span>
            </div>
          </div>
        </div>
      )}

      {/* Quotation Conversion Summary — Strictly Won vs Loss */}
      {isAdminOrOwner && (conversionStats.decidedCount > 0 || conversionStats.total > 0) && (
        <div className="card" style={{ background: '#ffffff', border: '1px solid #e2e8f0', borderRadius: 10, padding: '12px 14px' }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 8, flexWrap: 'wrap', gap: 6 }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
              <span style={{ fontSize: 11, fontWeight: 700, textTransform: 'uppercase', color: '#475569', letterSpacing: '0.04em' }}>
                Quotation Conversion (Won vs Loss)
              </span>
              {conversionStats.decidedCount > 0 && (
                <span className="tiny muted">({conversionStats.decidedCount} closed)</span>
              )}
            </div>
            <span style={{ fontSize: 11.5, color: '#64748b', fontWeight: 600 }}>
              {rangeLabel}
            </span>
          </div>

          <div style={{ height: 8, width: '100%', background: '#f1f5f9', borderRadius: 999, overflow: 'hidden', display: 'flex' }}>
            {conversionStats.decidedCount > 0 && (
              <>
                <div
                  style={{ width: `${conversionStats.winPct}%`, background: '#10b981', height: '100%' }}
                  title={`Win: ${conversionStats.won} (${conversionStats.winPct}%)`}
                />
                <div
                  style={{ width: `${conversionStats.lossPct}%`, background: '#ef4444', height: '100%' }}
                  title={`Loss: ${conversionStats.lost} (${conversionStats.lossPct}%)`}
                />
              </>
            )}
          </div>

          <div style={{ display: 'flex', gap: 18, marginTop: 8, fontSize: 12, color: '#334155', fontWeight: 500, flexWrap: 'wrap', alignItems: 'center' }}>
            {conversionStats.decidedCount === 0 ? (
              <span className="tiny muted">No Won or Loss quotations recorded yet in {rangeLabel}.</span>
            ) : (
              <>
                <span style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
                  <span style={{ width: 8, height: 8, borderRadius: '50%', background: '#10b981', display: 'inline-block' }} />
                  Win: <strong style={{ color: '#15803d' }}>{conversionStats.won} ({conversionStats.winPct}%)</strong>
                  <span className="tiny muted">· {inr(conversionStats.wonVal)}</span>
                </span>
                <span style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
                  <span style={{ width: 8, height: 8, borderRadius: '50%', background: '#ef4444', display: 'inline-block' }} />
                  Loss: <strong style={{ color: '#dc2626' }}>{conversionStats.lost} ({conversionStats.lossPct}%)</strong>
                  <span className="tiny muted">· {inr(conversionStats.lostVal)}</span>
                </span>
              </>
            )}
          </div>
        </div>
      )}

      {/* Two Column Grid for Admin; Single Column for Member */}
      <div className={isAdminOrOwner ? 'dash-grid' : 'dash-grid member-grid'}>
        {/* Recent Projects */}
        <div className="card dash-card" style={!isAdminOrOwner ? { gridColumn: '1 / -1' } : undefined}>
          <div className="card-head">
            <h3>{isAdminOrOwner ? 'Recent Quotations' : 'My Recent Quotations'}</h3>
            <button className="btn ghost sm" onClick={() => onNavigate('projects')}>
              View All
            </button>
          </div>
          <div className="dash-card-body">
            {displayedProjects.length === 0 ? (
              <div className="empty">
                <button className="btn primary sm" onClick={onNewProject}>
                  {Icon.plus} Create First Project
                </button>
              </div>
            ) : (
              <table className="tbl">
                <thead>
                  <tr>
                    <th>Quotation &amp; Client</th>
                    <th>Status</th>
                    <th>Template</th>
                    <th className="r">Amount</th>
                    <th className="r">Action</th>
                  </tr>
                </thead>
                <tbody>
                  {displayedProjects.map((p) => {
                    const anyP = p as any
                    const buyer = anyP.summary?.client_info?.buyer_name || ''
                    const isWon = !anyP.is_archived && (p.status === 'won' || p.status === 'approved')
                    const isLoss = anyP.is_archived || p.status === 'loss' || p.status === 'archived' || p.status === 'lost'
                    const isNegotiating = p.status === 'negotiating'
                    const statusLabel = isWon ? 'Won' : isLoss ? 'Loss' : isNegotiating ? 'Negotiating' : 'Quoted'
                    const statusClass = isWon ? 'won' : isLoss ? 'loss' : isNegotiating ? 'negotiating' : 'quoted'

                    return (
                      <tr key={p.id}>
                        <td>
                          <div style={{ display: 'flex', flexDirection: 'column', gap: 2 }}>
                            <div className="row gap" style={{ flexWrap: 'wrap' }}>
                              <button className="proj-title-btn" onClick={() => onOpenProject(p.id)}>
                                <strong>{p.name}</strong>
                              </button>
                              {(() => {
                                const rev = parseRevision(p.name)
                                if (!rev.currentRev) return null
                                return (
                                  <span className="chip primary sm font-mono font-bold" title={`Revision ${rev.currentRev}`}>
                                    Rev {rev.currentRev}
                                  </span>
                                )
                              })()}
                              {(() => {
                                const anyP = p as any
                                const shared = anyP.summary?.shared_with as string[] | undefined
                                const isSharedToMe =
                                  Array.isArray(shared) &&
                                  (shared.includes(session.username) || (session.user_id && shared.includes(session.user_id))) &&
                                  anyP.created_by !== session.username
                                if (isSharedToMe) {
                                  return (
                                    <span className="chip purple sm">
                                      Shared
                                    </span>
                                  )
                                }
                                return null
                              })()}
                            </div>
                            {buyer && (
                              <div style={{ fontSize: 11.5, color: '#64748b' }}>
                                Client: <strong style={{ color: '#334155' }}>{buyer}</strong>
                              </div>
                            )}
                          </div>
                        </td>
                        <td>
                          <span className={`status-pill status-${statusClass}`}>{statusLabel}</span>
                        </td>
                        <td>
                          <span className="version-pill">v{p.template_version}</span>
                        </td>
                        <td className="r font-bold text-green">
                          {p.total ? inr(Number(p.total)) : '—'}
                        </td>
                        <td className="r">
                          <button className="btn ghost sm" onClick={() => onOpenProject(p.id)} style={{ fontWeight: 600 }}>
                            Open →
                          </button>
                        </td>
                      </tr>
                    )
                  })}
                </tbody>
              </table>
            )}
          </div>
        </div>

        {/* Recent Activity (Admin / Owner ONLY) */}
        {isAdminOrOwner && (
          <div className="card dash-card">
            <div className="card-head">
              <h3>Activity</h3>
              <button className="btn ghost sm" onClick={() => onNavigate('activity')}>
                Full Log
              </button>
            </div>
            <div className="dash-card-body">
              {s.recent_activity.length === 0 ? (
                <div className="empty">
                  <p className="tiny muted">No recent activity recorded</p>
                </div>
              ) : (
                <div className="activity-timeline minimal-activity-feed">
                  {s.recent_activity.slice(0, 7).map((act) => (
                    <div key={act.id} className="activity-item minimal-activity-item">
                      <div className="act-bullet" />
                      <div className="act-content">
                        <div className="act-header" style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'baseline', gap: 8 }}>
                          <span style={{ fontSize: '12.5px', color: '#0f172a' }}>
                            <strong>{act.user_name}</strong>{' '}
                            <span style={{ color: '#475569' }}>{formatAction(act.action)}</span>
                            {act.details?.name ? (
                              <> <strong style={{ color: '#1d4ed8' }}>{act.details.name}</strong></>
                            ) : null}
                          </span>
                          <span className="act-time muted tiny" style={{ whiteSpace: 'nowrap' }}>{timeAgo(act.created_at)}</span>
                        </div>
                        {(act.details?.total || act.details?.buyer_name || act.details?.template_version) && (
                          <div className="tiny muted" style={{ marginTop: 2 }}>
                            {[
                              act.details?.buyer_name,
                              act.details?.total ? inr(Number(act.details.total)) : null,
                              act.details?.template_version ? `v${act.details.template_version}` : null,
                            ].filter(Boolean).join(' · ')}
                          </div>
                        )}
                      </div>
                    </div>
                  ))}
                </div>
              )}
            </div>
          </div>
        )}
      </div>
    </div>
  )
}

function formatAction(action: string) {
  switch (action) {
    case 'project.created':
    case 'project_created':
      return 'created'
    case 'project.updated':
    case 'project_updated':
      return 'updated'
    case 'project.recalculated':
    case 'project_recalculated':
      return 'recalculated'
    case 'project.duplicated':
    case 'project_duplicated':
      return 'revised / copied'
    case 'project.archived':
    case 'project_archived':
      return 'archived'
    case 'project.restored':
    case 'project_restored':
      return 'restored'
    case 'project_deleted':
      return 'deleted'
    case 'pdf_downloaded':
      return 'downloaded PDF'
    case 'pdf_saved':
      return 'saved PDF details'
    case 'template.published':
    case 'template_published':
      return 'published template'
    case 'template.set_default':
      return 'set default template'
    case 'member_joined':
      return 'joined team'
    case 'member_role_changed':
      return 'changed role'
    default:
      return action.replace(/[_.]/g, ' ')
  }
}
