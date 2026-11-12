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

  // Date range filtering for Owner & Admin
  type TimeRange = 'overall' | 'year' | 'month'
  const [timeRange, setTimeRange] = useState<TimeRange>('overall')
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

    const wonPct = total > 0 ? ((won / total) * 100).toFixed(1) : '0'
    const lostPct = total > 0 ? ((lost / total) * 100).toFixed(1) : '0'
    const negotiatingPct = total > 0 ? ((negotiating / total) * 100).toFixed(1) : '0'

    return {
      total,
      won,
      lost,
      negotiating,
      quoted,
      approved: won,
      totalVal,
      wonVal,
      lostVal,
      negotiatingVal,
      quotedVal,
      approvedVal: wonVal,
      wonPct,
      approvedPct: wonPct,
      lostPct,
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

  return (
    <div className="dashboard-page">
      {/* Streamlined Header Bar */}
      <div className="card row between p-3 mb-2">
        <div className="row gap">
          <h2 style={{ fontSize: '18px' }}>{org.name}</h2>
          {isAdminOrOwner && (
            <>
              <span
                className="chip primary font-mono cursor-pointer"
                onClick={() => setShowCode(!showCode)}
                title="Click to reveal or hide organization join code"
              >
                Code: {showCode ? org.code : '••••••••'}
              </span>
              <span className="version-pill">Default: v{s.current_template_version}</span>
            </>
          )}
        </div>
        <div className="row gap">
          <button className="btn primary sm" onClick={onNewProject}>
            {Icon.plus} New Project
          </button>
          <button className="btn ghost sm" onClick={() => onNavigate('projects')}>
            {Icon.folder} {isAdminOrOwner ? 'All Projects' : 'My Projects'}
          </button>
        </div>
      </div>

      {/* Owner & Admin Time Range Controls */}
      {isAdminOrOwner && (
        <div
          className="card p-2 mb-3"
          style={{
            display: 'flex',
            justifyContent: 'space-between',
            alignItems: 'center',
            flexWrap: 'wrap',
            gap: 12,
            background: '#ffffff',
            border: '1px solid #e2e8f0',
            borderRadius: 10,
            padding: '10px 16px',
          }}
        >
          <div style={{ display: 'flex', alignItems: 'center', gap: 10, flexWrap: 'wrap' }}>
            <span style={{ fontSize: 11, fontWeight: 700, textTransform: 'uppercase', color: '#64748b', letterSpacing: '0.05em' }}>
              Quotation Analytics:
            </span>
            <div style={{ display: 'flex', gap: 4, background: '#f1f5f9', padding: 3, borderRadius: 6 }}>
              <button
                type="button"
                className={`btn sm ${timeRange === 'overall' ? 'primary' : 'ghost'}`}
                onClick={() => setTimeRange('overall')}
                style={{ padding: '3px 12px', fontSize: 11.5, fontWeight: 700 }}
              >
                Overall (All Time)
              </button>
              <button
                type="button"
                className={`btn sm ${timeRange === 'year' ? 'primary' : 'ghost'}`}
                onClick={() => setTimeRange('year')}
                style={{ padding: '3px 12px', fontSize: 11.5, fontWeight: 700 }}
              >
                By Year
              </button>
              <button
                type="button"
                className={`btn sm ${timeRange === 'month' ? 'primary' : 'ghost'}`}
                onClick={() => setTimeRange('month')}
                style={{ padding: '3px 12px', fontSize: 11.5, fontWeight: 700 }}
              >
                By Month
              </button>
            </div>

            {timeRange === 'year' && (
              <select
                className="select sm"
                value={selectedYear}
                onChange={(e) => setSelectedYear(Number(e.target.value))}
                style={{ fontSize: 12, padding: '3px 8px', fontWeight: 600, borderRadius: 6 }}
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
                style={{ fontSize: 12, padding: '3px 8px', fontWeight: 600, borderRadius: 6 }}
              >
                {availableMonths.map((m) => (
                  <option key={m.value} value={m.value}>{m.label}</option>
                ))}
              </select>
            )}
          </div>

          <button
            type="button"
            className="btn ghost sm"
            onClick={() => onNavigate('staff_analytics')}
            style={{ fontWeight: 700, fontSize: 12, border: '1px solid #cbd5e1', padding: '4px 12px', borderRadius: 6 }}
          >
            👥 Staff Performance Breakdown →
          </button>
        </div>
      )}

      {/* Primary KPI Row */}
      <div className="kpis">
        {isAdminOrOwner ? (
          <>
            <div className="kpi primary">
              <span>Total Quotations</span>
              <strong>{conversionStats.total}</strong>
              <small>Total Pipeline: {inr(conversionStats.totalVal)}</small>
            </div>
            <div className="kpi kpi-won">
              <span style={{ color: '#166534' }}>Won</span>
              <strong style={{ color: '#15803d' }}>
                {conversionStats.won}
                <span className="metric-pill won">+{conversionStats.wonPct}%</span>
              </strong>
              <small style={{ color: '#166534' }}>Won: {inr(conversionStats.wonVal)}</small>
            </div>
            <div className="kpi kpi-negotiating">
              <span style={{ color: '#1e40af' }}>Negotiating</span>
              <strong style={{ color: '#2563eb' }}>
                {conversionStats.negotiating}
                <span className="metric-pill negotiating">{conversionStats.negotiatingPct}%</span>
              </strong>
              <small style={{ color: '#1e40af' }}>Negotiating: {inr(conversionStats.negotiatingVal)}</small>
            </div>
            <div className="kpi kpi-quoted">
              <span style={{ color: '#b45309' }}>Quoted</span>
              <strong style={{ color: '#d97706' }}>
                {conversionStats.quoted}
              </strong>
              <small style={{ color: '#b45309' }}>Quoted: {inr(conversionStats.quotedVal)}</small>
            </div>
            <div className="kpi kpi-loss">
              <span style={{ color: '#991b1b' }}>Loss</span>
              <strong style={{ color: '#dc2626' }}>
                {conversionStats.lost}
                <span className="metric-pill loss">{conversionStats.lostPct}%</span>
              </strong>
              <small style={{ color: '#991b1b' }}>Lost: {inr(conversionStats.lostVal)}</small>
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

      {/* Visual Conversion Progress Bar for Admin & Owner */}
      {isAdminOrOwner && conversionStats.total > 0 && (
        <div className="card p-3 mb-3" style={{ background: '#ffffff', border: '1px solid #e2e8f0', borderRadius: 10 }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 8 }}>
            <span style={{ fontSize: 11, fontWeight: 700, textTransform: 'uppercase', color: '#475569', letterSpacing: '0.04em' }}>
              Quotation Conversion Overview
            </span>
            <span style={{ fontSize: 12, color: '#64748b', fontWeight: 600 }}>
              {timeRange === 'overall' ? 'All Time' : timeRange === 'year' ? `Year ${selectedYear}` : selectedMonth}
            </span>
          </div>
          <div style={{ height: 8, width: '100%', background: '#f1f5f9', borderRadius: 999, overflow: 'hidden', display: 'flex' }}>
            <div
              style={{ width: `${conversionStats.wonPct}%`, background: '#10b981', height: '100%' }}
              title={`Won: ${conversionStats.won} (${conversionStats.wonPct}%)`}
            />
            <div
              style={{ width: `${conversionStats.negotiatingPct}%`, background: '#3b82f6', height: '100%' }}
              title={`Negotiating: ${conversionStats.negotiating} (${conversionStats.negotiatingPct}%)`}
            />
            <div
              style={{ width: `${conversionStats.lostPct}%`, background: '#ef4444', height: '100%' }}
              title={`Loss: ${conversionStats.lost} (${conversionStats.lostPct}%)`}
            />
            <div
              style={{
                width: `${Math.max(0, 100 - Number(conversionStats.wonPct) - Number(conversionStats.negotiatingPct) - Number(conversionStats.lostPct))}%`,
                background: '#f59e0b',
                height: '100%',
              }}
              title={`Quoted: ${conversionStats.quoted}`}
            />
          </div>
          <div style={{ display: 'flex', gap: 16, marginTop: 8, fontSize: 11.5, color: '#475569', fontWeight: 500, flexWrap: 'wrap' }}>
            <span style={{ display: 'flex', alignItems: 'center', gap: 5 }}>
              <span style={{ width: 8, height: 8, borderRadius: '50%', background: '#10b981', display: 'inline-block' }} />
              Won: <strong>{conversionStats.won} ({conversionStats.wonPct}%)</strong>
            </span>
            <span style={{ display: 'flex', alignItems: 'center', gap: 5 }}>
              <span style={{ width: 8, height: 8, borderRadius: '50%', background: '#3b82f6', display: 'inline-block' }} />
              Negotiating: <strong>{conversionStats.negotiating} ({conversionStats.negotiatingPct}%)</strong>
            </span>
            <span style={{ display: 'flex', alignItems: 'center', gap: 5 }}>
              <span style={{ width: 8, height: 8, borderRadius: '50%', background: '#f59e0b', display: 'inline-block' }} />
              Quoted: <strong>{conversionStats.quoted}</strong>
            </span>
            <span style={{ display: 'flex', alignItems: 'center', gap: 5 }}>
              <span style={{ width: 8, height: 8, borderRadius: '50%', background: '#ef4444', display: 'inline-block' }} />
              Loss: <strong>{conversionStats.lost} ({conversionStats.lostPct}%)</strong>
            </span>
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
                <div className="activity-timeline">
                  {s.recent_activity.slice(0, 6).map((act) => (
                    <div key={act.id} className="activity-item">
                      <div className="act-bullet" />
                      <div className="act-content">
                        <div className="act-header">
                          <span className="act-user">{act.user_name}</span>
                          <span className="act-action">{formatAction(act.action)}</span>
                          <span className="act-time muted tiny">{timeAgo(act.created_at)}</span>
                        </div>
                        {act.details?.name && (
                          <div className="act-detail">
                            {act.details.name} {act.details.template_version ? `(v${act.details.template_version})` : ''}
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
      return 'created project'
    case 'project.updated':
      return 'saved project'
    case 'project.recalculated':
      return 'recalculated project'
    case 'project.duplicated':
      return 'duplicated project'
    case 'project.archived':
      return 'archived project'
    case 'project.restored':
      return 'restored project'
    case 'template.published':
      return 'published new template'
    case 'template.set_default':
      return 'switched default template'
    case 'member_joined':
      return 'joined organisation'
    default:
      return action.replace(/_/g, ' ')
  }
}
