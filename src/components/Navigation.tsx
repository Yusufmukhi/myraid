import { Org } from '../lib/api'
import { Icon } from './ui'

export type NavTab = 'dashboard' | 'projects' | 'shared_projects' | 'clients' | 'calculator' | 'details' | 'approvals' | 'settings' | 'members' | 'activity' | 'business' | 'staff_analytics'

interface Props {
  currentTab: NavTab
  onSelectTab: (tab: NavTab) => void
  org: Org
  currentProjectName?: string | null
  projectCount?: number
  sharedCount?: number
  clientCount?: number
  hasUnsavedChanges?: boolean
  pendingApprovalsCount?: number
  onOpenCreateModal?: () => void
}

export function Navigation({
  currentTab,
  onSelectTab,
  org,
  currentProjectName,
  projectCount,
  sharedCount,
  clientCount,
  hasUnsavedChanges,
  pendingApprovalsCount,
  onOpenCreateModal,
}: Props) {
  const role = org?.role || 'member'
  const isAdminOrOwner = role === 'owner' || role === 'admin'
  const displayRole = (role === 'floor_operator' ? 'member' : role).toUpperCase()

  return (
    <nav className="nav-sidebar">
      <div className="nav-org-pill">
        <span className="tiny font-bold uppercase text-muted">Active Org</span>
        <strong className="nav-org-name">{org?.name || 'Organisation'}</strong>
        <span className="nav-org-role">{displayRole}</span>
      </div>

      <div className="nav-group">
        <span className="nav-heading">Workspace</span>
        <button
          className={`nav-item ${currentTab === 'dashboard' ? 'active' : ''}`}
          onClick={() => onSelectTab('dashboard')}
        >
          <span className="nav-icon">{Icon.dashboard}</span>
          <span className="nav-text">Dashboard</span>
        </button>

        <button
          className={`nav-item ${currentTab === 'projects' ? 'active' : ''}`}
          onClick={() => onSelectTab('projects')}
        >
          <span className="nav-icon">{Icon.folder}</span>
          <span className="nav-text">{isAdminOrOwner ? 'Projects & Estimates' : 'My Quotations'}</span>
          {typeof projectCount === 'number' && <span className="nav-badge">{projectCount}</span>}
        </button>

        {!isAdminOrOwner && (
          <button
            className={`nav-item ${currentTab === 'shared_projects' ? 'active' : ''}`}
            onClick={() => onSelectTab('shared_projects')}
          >
            <span className="nav-icon">{Icon.users}</span>
            <span className="nav-text">Shared Quotations</span>
            {typeof sharedCount === 'number' && (
              <span className="nav-badge" style={{ background: '#7c3aed', color: '#fff' }}>
                {sharedCount}
              </span>
            )}
          </button>
        )}

        <button
          className={`nav-item ${currentTab === 'clients' ? 'active' : ''}`}
          onClick={() => onSelectTab('clients')}
          title="Manage client directory, addresses, and GST numbers"
        >
          <span className="nav-icon">{Icon.building}</span>
          <span className="nav-text">Client Directory</span>
          {typeof clientCount === 'number' && <span className="nav-badge">{clientCount}</span>}
        </button>
      </div>

      <div className="nav-group">
        <span className="nav-heading">Active Project</span>
        <button
          className={`nav-item ${currentTab === 'calculator' ? 'active' : ''}`}
          onClick={() => {
            onSelectTab('calculator')
            if (!currentProjectName && onOpenCreateModal) {
              onOpenCreateModal()
            }
          }}
        >
          <span className="nav-icon">{Icon.calc}</span>
          <div className="nav-stacked">
            <span className="nav-text">Calculator</span>
            <span className="nav-subtext">
              {currentProjectName || '+ Create Project'}
            </span>
          </div>
          {hasUnsavedChanges && <span className="dot amber-dot" title="Unsaved changes" />}
        </button>

        {isAdminOrOwner && (
          <button
            className={`nav-item ${currentTab === 'details' ? 'active' : ''}`}
            onClick={() => onSelectTab('details')}
          >
            <span className="nav-icon">{Icon.grid}</span>
            <span className="nav-text">Calculation Details</span>
            <span className="nav-badge-pill">15 sheets</span>
          </button>
        )}
      </div>

      {isAdminOrOwner && (
        <div className="nav-group">
          <span className="nav-heading">Organisation &amp; Governance</span>
          <button
            className={`nav-item ${currentTab === 'business' ? 'active' : ''}`}
            onClick={() => onSelectTab('business')}
            title="Edit company profile, addresses, tax IDs, and PDF quotation header"
          >
            <span className="nav-icon">{Icon.building}</span>
            <span className="nav-text">Business Profile</span>
          </button>

          <button
            className={`nav-item ${currentTab === 'staff_analytics' ? 'active' : ''}`}
            onClick={() => onSelectTab('staff_analytics')}
            title="View each staff member's quotations, values, and won/lost analytics"
          >
            <span className="nav-icon">{Icon.users}</span>
            <span className="nav-text">Staff Analytics</span>
          </button>

          <button
            className={`nav-item ${currentTab === 'approvals' ? 'active' : ''}`}
            onClick={() => onSelectTab('approvals')}
          >
            <span className="nav-icon">{Icon.check}</span>
            <span className="nav-text">Approvals</span>
            {typeof pendingApprovalsCount === 'number' && pendingApprovalsCount > 0 && (
              <span
                className="nav-badge"
                style={{
                  background: '#f59e0b',
                  color: '#fff',
                  borderRadius: '10px',
                  padding: '1px 6px',
                  fontSize: '11px',
                  fontWeight: 800,
                }}
              >
                {pendingApprovalsCount}
              </span>
            )}
          </button>

          <button
            className={`nav-item ${currentTab === 'settings' ? 'active' : ''}`}
            onClick={() => onSelectTab('settings')}
          >
            <span className="nav-icon">{Icon.layers}</span>
            <span className="nav-text">Template Versioning</span>
            <span className="nav-version-tag">v{org.current_template_version || 1}</span>
          </button>

          <button
            className={`nav-item ${currentTab === 'members' ? 'active' : ''}`}
            onClick={() => onSelectTab('members')}
          >
            <span className="nav-icon">{Icon.users}</span>
            <span className="nav-text">Members &amp; Access</span>
          </button>

          <button
            className={`nav-item ${currentTab === 'activity' ? 'active' : ''}`}
            onClick={() => onSelectTab('activity')}
          >
            <span className="nav-icon">{Icon.history}</span>
            <span className="nav-text">Activity &amp; Audit Log</span>
          </button>
        </div>
      )}
    </nav>
  )
}
