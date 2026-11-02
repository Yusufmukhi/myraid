import { useEffect, useMemo, useState } from 'react'
import { ActivityItem, api, Org, Session } from '../../lib/api'
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
  const [loading, setLoading] = useState(true)
  const [filterAction, setFilterAction] = useState<string>('all')
  const [selectedStaff, setSelectedStaff] = useState<string>('all')
  const [search, setSearch] = useState('')
  const [projectMap, setProjectMap] = useState<Map<string, any>>(new Map())
  const [projectRevisionsMap, setProjectRevisionsMap] = useState<Map<string, any[]>>(new Map())
  const [clientIp, setClientIp] = useState<string>(getCachedClientIp())
  const [expandedLogIds, setExpandedLogIds] = useState<Set<string>>(new Set())
  const [modalLog, setModalLog] = useState<ActivityItem | null>(null)
  const [currentPage, setCurrentPage] = useState(1)
  const PAGE_SIZE = 20

  const loadActivity = async () => {
    try {
      setLoading(true)
      const [data, projs] = await Promise.all([
        api.activityList(session.token, org.id, 100),
        api.projectList(session.token, org.id, true).catch(() => []),
      ])
      setLogs(Array.isArray(data) ? data : [])
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

  const ownerUsername = org.role === 'owner' ? session.username : undefined

  const staffMembers = useMemo(() => {
    const set = new Set<string>()
    for (const l of logs || []) {
      const actor = l.user_name || (l as any).username || l.details?.username || l.details?.user_name
      if (actor && typeof actor === 'string' && actor.trim()) {
        set.add(actor.trim())
      }
    }
    return Array.from(set).sort((a, b) => a.localeCompare(b))
  }, [logs])

  const securityAlertsCount = useMemo(() => {
    return (logs || []).filter((l) => checkSecurityIncident(l, ownerUsername).isIncident).length
  }, [logs, ownerUsername])

  const filteredLogs = useMemo(() => {
    return (logs || []).filter((item) => {
      if (filterAction === 'security_alerts') {
        if (!checkSecurityIncident(item, ownerUsername).isIncident) return false
      } else if (filterAction !== 'all' && item.action !== filterAction) {
        return false
      }
      if (selectedStaff !== 'all') {
        const actor = item.user_name || (item as any).username || item.details?.username || item.details?.user_name
        if (actor !== selectedStaff) return false
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
  }, [logs, filterAction, selectedStaff, search, ownerUsername])

  // Reset to page 1 whenever filters or search change
  useEffect(() => {
    setCurrentPage(1)
  }, [filterAction, selectedStaff, search])

  const totalPages = Math.max(1, Math.ceil(filteredLogs.length / PAGE_SIZE))
  const activePage = Math.min(currentPage, totalPages)

  const paginatedLogs = useMemo(() => {
    const start = (activePage - 1) * PAGE_SIZE
    return filteredLogs.slice(start, start + PAGE_SIZE)
  }, [filteredLogs, activePage])

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

  return (
    <div className="activity-page">
      {/* Top Header */}
      <div className="activity-top-header mb-3">
        <div>
          <h2 style={{ fontSize: '20px', margin: 0, fontWeight: 700 }}>Organisation Audit Trail</h2>
          <p className="tiny muted" style={{ margin: '3px 0 0 0' }}>
            Permanent revision history, system access events, and field-level modification records ({filteredLogs.length} events)
          </p>
        </div>
        <div className="row gap" style={{ alignItems: 'center' }}>
          <button
            className="btn ghost sm"
            onClick={handleExpandAll}
            title={expandedLogIds.size === filteredLogs.length ? 'Collapse all log details' : 'Expand all log details'}
          >
            {expandedLogIds.size === filteredLogs.length ? '▴ Collapse All' : '▾ Expand All'}
          </button>
          <button className="btn ghost sm" onClick={loadActivity} title="Refresh audit log">
            {Icon.refresh} Refresh
          </button>
        </div>
      </div>

      {/* Security Incident Banner (Shows ONLY when real incidents exist) */}
      {securityAlertsCount > 0 && (
        <div className="security-incidents-banner mb-3">
          <div className="row between" style={{ alignItems: 'center' }}>
            <div className="row gap" style={{ alignItems: 'center' }}>
              <span style={{ fontSize: '22px' }}>🛡️</span>
              <div>
                <strong style={{ fontSize: '14px', color: '#991b1b' }}>
                  Security Incident Monitor: {securityAlertsCount} Event{securityAlertsCount > 1 ? 's' : ''} Flagged
                </strong>
                <div className="tiny" style={{ color: '#7f1d1d' }}>
                  Flagging external access requests, project deletions, and rogue price revisions.
                </div>
              </div>
            </div>
            <button
              className={`btn sm ${filterAction === 'security_alerts' ? 'primary' : 'ghost'}`}
              onClick={() => setFilterAction(filterAction === 'security_alerts' ? 'all' : 'security_alerts')}
              style={filterAction === 'security_alerts' ? undefined : { color: '#991b1b', borderColor: '#fca5a5' }}
            >
              {filterAction === 'security_alerts' ? 'Show All Activity' : `Filter Incidents (${securityAlertsCount})`}
            </button>
          </div>
        </div>
      )}

      {/* Filter and Search Bar */}
      <div className="card filter-bar mb-3">
        <div className="search-wrap">
          <span className="search-icon">{Icon.search}</span>
          <input
            type="text"
            className="search-input"
            placeholder="Search by actor, quotation name, or event…"
            value={search}
            onChange={(e) => setSearch(e.target.value)}
          />
          {search && (
            <button className="icon-btn sm" onClick={() => setSearch('')}>
              {Icon.close}
            </button>
          )}
        </div>

        <div className="filter-controls" style={{ display: 'flex', gap: '8px', flexWrap: 'wrap' }}>
          {staffMembers.length > 0 && (
            <select
              className="select-input"
              value={selectedStaff}
              onChange={(e) => setSelectedStaff(e.target.value)}
              title="Filter activity by staff member"
              style={{ minWidth: '150px' }}
            >
              <option value="all">👤 All Staff ({staffMembers.length})</option>
              {staffMembers.map((staff) => (
                <option key={staff} value={staff}>
                  👤 {staff}
                </option>
              ))}
            </select>
          )}

          <select
            className="select-input"
            value={filterAction}
            onChange={(e) => setFilterAction(e.target.value)}
          >
            <option value="all">All Event Types</option>
            {securityAlertsCount > 0 && (
              <option value="security_alerts">🛡️ Security Incidents Only ({securityAlertsCount})</option>
            )}
            <option value="project_created">Project Created</option>
            <option value="pdf_downloaded">📄 PDF Downloads &amp; Saves</option>
            <option value="project_updated">Project Updated</option>
            <option value="project_duplicated">Project Duplicated</option>
            <option value="project_recalculated">Project Recalculated</option>
            <option value="template_published">Template Published</option>
            <option value="project_archived">Project Archived</option>
            <option value="project_deleted">Project Deleted</option>
            <option value="member_joined">Member Joined</option>
            <option value="member_role_changed">Role Changed</option>
            <option value="join_requested">Join Requested</option>
            <option value="member_rejected">Request Rejected</option>
            <option value="join_code_rotated">Join Code Rotated</option>
          </select>
        </div>
      </div>

      {/* Activity Timeline List */}
      <div className="card activity-audit-card">
        {loading ? (
          <div className="page-loading">
            <Spinner size={28} />
            <p className="muted">Loading audit records…</p>
          </div>
        ) : filteredLogs.length === 0 ? (
          <div className="empty">
            <p>No activity records match your criteria</p>
          </div>
        ) : (
          <div className="activity-audit-list">
            {paginatedLogs.map((item, index) => {
              const { badgeClass, label } = getActionMeta(item.action)
              const proj = item.project_id
                ? projectMap.get(item.project_id)
                : item.details?.name
                ? projectMap.get(item.details.name.toLowerCase().trim())
                : null

              // Network & IP resolution
              const rawIp = item.details?.ip || proj?.summary?.ip || proj?.summary?.client_info?.ip || clientIp
              const netInfo = formatAuditNetwork(rawIp)

              // Device resolution
              const rawDevice = item.details?.device || proj?.summary?.device || proj?.summary?.client_info?.device
              const devInfo = formatAuditDevice(rawDevice)

              // Extract changes
              const extracted = extractLogChanges(
                item,
                index,
                paginatedLogs,
                projectMap,
                projectRevisionsMap
              )
              const hasChanges = Boolean(extracted && extracted.changes && extracted.changes.length > 0)
              const isExpanded = expandedLogIds.has(item.id)

              return (
                <div
                  key={item.id}
                  className={`audit-row-container ${isExpanded ? 'is-expanded' : ''}`}
                >
                  {/* Clean, Non-Cluttered Primary Summary Row */}
                  <div
                    className="audit-summary-row"
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
                    {/* Status Icon */}
                    <div className="audit-icon-col">
                      <span className={`audit-badge ${badgeClass}`}>{getActionIcon(item.action)}</span>
                    </div>

                    {/* Middle Content */}
                    <div className="audit-middle-col">
                      <div className="audit-actor-line">
                        <strong className="audit-user">{item.user_name}</strong>
                        <span className="audit-action-label">{label}</span>
                        {item.details?.client_name && (
                          <span className="audit-project-pill" style={{ color: '#d97706', borderColor: '#fde68a', background: '#fffbeb' }}>
                            “{item.details.client_name}”
                          </span>
                        )}
                        {item.details?.name && !item.details?.client_name && (
                          <span
                            className="audit-project-pill"
                            onClick={(e) => {
                              if (item.project_id && onOpenProject) {
                                e.stopPropagation()
                                onOpenProject(item.project_id)
                              }
                            }}
                            title={item.project_id ? 'Click to open project in calculator' : undefined}
                          >
                            “{item.details.name}”
                          </span>
                        )}
                      </div>

                      {/* Clean Sub-Metadata Badges */}
                      <div className="audit-meta-pills">
                        {item.action === 'client_conflict' && (
                          <>
                            {item.details?.existing_owner && (
                              <span className="chip warn sm" style={{ fontWeight: 600 }}>
                                Handled by: @{item.details.existing_owner}
                              </span>
                            )}
                            {item.details?.gst && (
                              <span className="version-pill">
                                GST: {item.details.gst}
                              </span>
                            )}
                          </>
                        )}
                        {item.details?.template_version && (
                          <span className="version-pill">Template v{item.details.template_version}</span>
                        )}
                        {item.details?.total !== undefined && Number(item.details.total) > 0 && (
                          <span className="audit-total-pill">Total: {inr(Number(item.details.total))}</span>
                        )}
                        {item.details?.old_version && item.details?.new_version && (
                          <span className="chip warn sm">
                            v{item.details.old_version} → v{item.details.new_version}
                          </span>
                        )}
                        {item.details?.new_role && (
                          <span className="chip primary sm">Role: {item.details.new_role}</span>
                        )}

                        {/* Interactive Changes Counter Pill */}
                        {hasChanges && (
                          <span
                            className={`audit-changes-tag ${isExpanded ? 'active' : ''}`}
                            title="Click to view full modification details"
                          >
                            📝 {extracted.changes.length} change{extracted.changes.length > 1 ? 's' : ''}{' '}
                            <span className="tag-arrow">{isExpanded ? '▴' : '▾'}</span>
                          </span>
                        )}

                        {/* Device Badge */}
                        <span className="audit-device-pill" title={`Device: ${devInfo.label}`}>
                          <span>{devInfo.icon}</span> {devInfo.label}
                        </span>
                      </div>
                    </div>

                    {/* Right Meta & Expand Action */}
                    <div className="audit-right-col" onClick={(e) => e.stopPropagation()}>
                      {/* Network & IP pill */}
                      <span
                        className={`audit-network-pill ${netInfo.networkType}`}
                        title={netInfo.tooltip}
                      >
                        <span>{netInfo.icon}</span> {netInfo.label}
                      </span>

                      {/* Timestamp */}
                      <span
                        className="audit-time-label"
                        title={new Date(item.created_at).toLocaleString('en-IN', {
                          dateStyle: 'full',
                          timeStyle: 'medium',
                        })}
                      >
                        {timeAgo(item.created_at)}
                      </span>

                      {/* Toggle button */}
                      <button
                        className={`btn sm ${isExpanded ? 'primary' : 'ghost'} audit-expand-btn`}
                        onClick={() => toggleExpand(item.id)}
                        aria-expanded={isExpanded}
                      >
                        {isExpanded ? 'Hide Details ▴' : 'View Changes ▾'}
                      </button>
                    </div>
                  </div>

                  {/* =========================================================================
                      DETAILED MODIFIED CHANGES & AUDIT CONTEXT (OPENS ON CLICK)
                      ========================================================================= */}
                  {isExpanded && (
                    <div className="audit-details-accordion">
                      <div className="audit-drawer-inner">
                        {/* 1. What Changes Have Been Made (Field-by-Field Diff) */}
                        <div className="audit-section-box">
                          <div className="audit-section-head">
                            <div className="row gap" style={{ alignItems: 'center' }}>
                              <span style={{ fontSize: '15px' }}>📝</span>
                              <strong style={{ fontSize: '13px' }}>
                                What has been changed ({extracted.changes.length} recorded update{extracted.changes.length > 1 ? 's' : ''})
                              </strong>
                            </div>
                            <span className="tiny muted">{extracted.sectionTitle}</span>
                          </div>

                          <div className="audit-diff-grid">
                            {extracted.changes.map((ch: any, idx: number) => (
                              <DiffItemRenderer key={idx} item={ch} />
                            ))}
                          </div>
                        </div>

                        {/* 2. And Other Changes Also (Comprehensive Event & Technical Context) */}
                        <div className="audit-section-box mt-2">
                          <div className="audit-section-head">
                            <div className="row gap" style={{ alignItems: 'center' }}>
                              <span style={{ fontSize: '15px' }}>ℹ️</span>
                              <strong style={{ fontSize: '13px' }}>
                                Other changes & technical audit details
                              </strong>
                            </div>
                            <button
                              className="btn ghost sm py-0 px-2 tiny"
                              onClick={() => setModalLog(item)}
                              title="Inspect full raw JSON & snapshot in modal"
                            >
                              🔍 Full Inspector
                            </button>
                          </div>

                          <div className="audit-context-grid">
                            <div className="context-card">
                              <span className="context-label">ACTOR & PERMISSIONS</span>
                              <div className="context-value">
                                <strong>{item.user_name}</strong>
                                <span className="chip sm ml-2">
                                  {item.user_name.toLowerCase() === session.username.toLowerCase()
                                    ? org.role
                                    : 'Staff'}
                                </span>
                              </div>
                            </div>

                            <div className="context-card">
                              <span className="context-label">AUDITED NETWORK & IP</span>
                              <div className="context-value">
                                <span className="font-mono">{netInfo.ip}</span>
                                <span className="tiny muted ml-2">({netInfo.label})</span>
                              </div>
                            </div>

                            <div className="context-card">
                              <span className="context-label">EXACT TIMESTAMP</span>
                              <div className="context-value font-mono">
                                {new Date(item.created_at).toLocaleString('en-IN', {
                                  dateStyle: 'medium',
                                  timeStyle: 'medium',
                                })}
                              </div>
                            </div>

                            <div className="context-card">
                              <span className="context-label">DEVICE & OS PLATFORM</span>
                              <div className="context-value">
                                <span>{devInfo.icon}</span> {devInfo.label}
                              </div>
                            </div>

                            <div className="context-card">
                              <span className="context-label">EVENT TYPE & ACTION</span>
                              <div className="context-value font-mono tiny">
                                {item.action}
                              </div>
                            </div>

                            <div className="context-card">
                              <span className="context-label">TARGET RESOURCE</span>
                              <div className="context-value">
                                {item.details?.name ? (
                                  <span>“{item.details.name}”</span>
                                ) : (
                                  <span className="tiny muted">Organisation Level Event</span>
                                )}
                              </div>
                            </div>
                          </div>

                          {/* Action Footer */}
                          {item.project_id && onOpenProject && (
                            <div className="row end gap mt-3">
                              <button
                                className="btn primary sm"
                                onClick={() => onOpenProject(item.project_id!)}
                              >
                                {Icon.fileText} Open This Quotation in Calculator →
                              </button>
                            </div>
                          )}
                        </div>
                      </div>
                    </div>
                  )}
                </div>
              )
            })}
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
    let delta = undefined
    if (isTotal || /total/i.test(label) || (oldVal.startsWith('₹') && newVal.startsWith('₹'))) {
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
    return { label, oldVal, newVal, isTotal, delta }
  }

  return { rawText: rawText || String(raw) }
}

function DiffItemRenderer({ item }: { item: any }) {
  const parsed = parseDiffItem(item)

  if (parsed.oldVal !== undefined && parsed.newVal !== undefined) {
    return (
      <div className={`audit-diff-card ${parsed.isTotal ? 'diff-total' : ''}`}>
        <div className="diff-field">{parsed.label}</div>
        <div className="diff-values">
          <span className={`diff-val-old ${parsed.isTotal ? 'strike' : ''}`} title="Previous value">
            {parsed.oldVal}
          </span>
          <span className="diff-arrow" aria-hidden="true">
            →
          </span>
          <span className={`diff-val-new ${parsed.isTotal ? 'highlight' : ''}`} title="Updated value">
            {parsed.newVal}
          </span>
          {parsed.delta && (
            <span
              className={`diff-delta ${parsed.delta.isIncrease ? 'pos' : 'neg'}`}
              title={parsed.delta.isIncrease ? 'Price Increase' : 'Price Decrease'}
            >
              {parsed.delta.formatted}
            </span>
          )}
        </div>
      </div>
    )
  }

  return (
    <div className="audit-diff-card diff-action-info">
      <span className="diff-bullet">✓</span>
      <span className="diff-text">{parsed.rawText}</span>
    </div>
  )
}
