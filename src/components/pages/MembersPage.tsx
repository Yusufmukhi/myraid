import { useEffect, useState, useMemo } from 'react'
import { api, Member, Org, Session } from '../../lib/api'
import { timeAgo } from '../../lib/format'
import { Icon, Spinner, useCopy, useUI } from '../ui'

interface Props {
  session: Session
  org: Org
  onOrgLeave: () => void
}

export function MembersPage({ session, org, onOrgLeave }: Props) {
  const { toast, confirm } = useUI()
  const copy = useCopy()

  const [members, setMembers] = useState<Member[]>([])
  const [loading, setLoading] = useState(true)
  const [inviteEmail, setInviteEmail] = useState('')
  const [inviteRole, setInviteRole] = useState<'member' | 'admin' | 'viewer'>('member')
  const [inviting, setInviting] = useState(false)
  const [searchQuery, setSearchQuery] = useState('')
  const [roleFilter, setRoleFilter] = useState<'all' | 'owner' | 'admin' | 'member' | 'viewer'>('all')
  const [createdUser, setCreatedUser] = useState<{
    email: string
    username: string
    password?: string
    is_new: boolean
    role: string
    org_name: string
  } | null>(null)

  const isOwner = org.role === 'owner'
  const isAdminOrOwner = org.role === 'owner' || org.role === 'admin'

  const loadMembers = async () => {
    try {
      setLoading(true)
      const data = await api.orgMembers(session.token, org.id)
      setMembers(data)
    } catch (e) {
      toast((e as Error).message, 'err')
    } finally {
      setLoading(false)
    }
  }

  useEffect(() => {
    loadMembers()
  }, [org.id])

  const counts = useMemo(() => {
    const owner = members.filter((m) => m.role === 'owner').length
    const admin = members.filter((m) => m.role === 'admin').length
    const staff = members.filter((m) => m.role === 'member').length
    const viewer = members.filter((m) => m.role === 'viewer' || (m.role as string) === 'floor_operator').length
    return { total: members.length, owner, admin, staff, viewer }
  }, [members])

  const filteredMembers = useMemo(() => {
    return members.filter((m) => {
      const q = searchQuery.trim().toLowerCase()
      const matchesSearch =
        !q ||
        m.username.toLowerCase().includes(q) ||
        (m.email && m.email.toLowerCase().includes(q))
      if (!matchesSearch) return false

      if (roleFilter === 'all') return true
      if (roleFilter === 'owner') return m.role === 'owner'
      if (roleFilter === 'admin') return m.role === 'admin'
      if (roleFilter === 'member') return m.role === 'member'
      if (roleFilter === 'viewer') return m.role === 'viewer' || (m.role as string) === 'floor_operator'
      return true
    })
  }, [members, searchQuery, roleFilter])

  const handleRoleChange = async (member: Member, newRole: string) => {
    if (!isAdminOrOwner) return toast('Admin privileges required', 'err')
    if (member.role === 'owner') return toast('Cannot change owner role', 'err')
    try {
      if (member.id) await api.orgUpdateRole(session.token, org.id, member.id, newRole)
      const roleLabel =
        newRole === 'viewer' || newRole === 'floor_operator'
          ? 'Floor Operator'
          : newRole === 'admin'
          ? 'Administrator'
          : 'Staff Member'
      toast(`Updated ${member.username} role to ${roleLabel}`)
      loadMembers()
    } catch (e) {
      toast((e as Error).message, 'err')
    }
  }

  const handleRemoveMember = async (member: Member) => {
    if (!isAdminOrOwner) return toast('Admin privileges required', 'err')
    if (member.role === 'owner') return toast('Cannot remove owner', 'err')
    const ok = await confirm({
      title: `Remove ${member.username}?`,
      body: `Revoke access for ${member.username} from ${org.name}? They will no longer be able to access this organisation's quotations.`,
      confirmLabel: 'Remove Member',
      danger: true,
    })
    if (!ok) return
    try {
      if (member.id) await api.orgRemoveMember(session.token, org.id, member.id)
      toast(`Removed ${member.username}`)
      loadMembers()
    } catch (e) {
      toast((e as Error).message, 'err')
    }
  }

  const handleSendInvite = async () => {
    if (!isOwner) return toast('Only the organisation owner can provision new staff accounts.', 'err')
    if (!inviteEmail || !inviteEmail.includes('@')) return toast('Enter a valid email address.', 'err')
    setInviting(true)
    const targetEmail = inviteEmail.trim()
    try {
      const res = await api.orgInviteMember(session.token, org.id, targetEmail, inviteRole)
      setCreatedUser(res)
      setInviteEmail('')
      loadMembers()
      toast(res.is_new ? `Staff account created for ${res.email}!` : `Existing user ${res.email} added to organisation!`)
    } catch (e) {
      toast((e as Error).message, 'err')
    } finally {
      setInviting(false)
    }
  }

  const getLoginMessage = () => {
    if (!createdUser) return ''
    const loginUrl = window.location.origin
    return `Hello,\n\nYou have been added to ${createdUser.org_name || org.name} on the Apollo Storage Quotation Portal.\n\nPortal URL: ${loginUrl}\nLogin Email: ${createdUser.email}\nUsername: ${createdUser.username}\n${createdUser.password ? `Temporary Password: ${createdUser.password}\n\nPlease sign in and change your username and password anytime in My Profile.` : 'You can sign in using your existing account credentials.'}\n\nBest regards,\n${session.username} (${org.name})`
  }

  const handleCopyCredentials = () => {
    copy(getLoginMessage(), 'Staff login details copied to clipboard!')
  }

  const getMailtoUrl = () => {
    if (!createdUser) return '#'
    const subject = encodeURIComponent(`Your Login for ${createdUser.org_name || org.name} Quotation Portal`)
    const body = encodeURIComponent(getLoginMessage())
    return `mailto:${createdUser.email}?subject=${subject}&body=${body}`
  }

  const handleLeave = async () => {
    if (isOwner) return toast('Owners cannot leave without transferring ownership', 'err')
    const ok = await confirm({
      title: `Leave ${org.name}?`,
      body: 'You will need an invitation from the owner to access this organisation again.',
      confirmLabel: 'Leave Organisation',
      danger: true,
    })
    if (!ok) return
    try {
      await api.orgLeave(session.token, org.id)
      onOrgLeave()
    } catch (e) {
      toast((e as Error).message, 'err')
    }
  }

  const getAvatarBg = (username: string, role: string) => {
    if (role === 'owner') return '#7c3aed'
    if (role === 'admin') return '#2563eb'
    if (role === 'viewer' || (role as string) === 'floor_operator') return '#d97706'
    return '#059669'
  }

  return (
    <div className="members-page">
      {/* Streamlined Top Header */}
      <div className="card p-3 mb-2" style={{ borderRadius: '10px', boxShadow: '0 1px 4px rgba(0, 0, 0, 0.04)' }}>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: '12px' }}>
          <div>
            <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
              <h2 style={{ fontSize: '18px', fontWeight: 800, margin: 0, color: '#0f172a' }}>
                👥 Team Members &amp; Access Management
              </h2>
              <span className="chip primary font-bold" style={{ textTransform: 'uppercase' }}>
                {org.name}
              </span>
            </div>
            <div className="tiny muted mt-1" style={{ fontSize: '11.5px', color: '#64748b' }}>
              Manage active user accounts, grant role permissions, and provision secure credentials for {org.name}.
            </div>
          </div>

          {!isOwner && (
            <button className="btn ghost danger sm" onClick={handleLeave} title="Leave this organisation">
              {Icon.logout} Leave Org
            </button>
          )}
        </div>
      </div>

      {/* KPI Stats Overview Cards */}
      <div className="members-stats-grid">
        <div className="members-stat-card">
          <div className="members-stat-icon" style={{ background: '#f1f5f9', color: '#334155' }}>
            👥
          </div>
          <div className="members-stat-info">
            <span className="members-stat-value">{counts.total}</span>
            <span className="members-stat-label">Total Members</span>
          </div>
        </div>

        <div className="members-stat-card">
          <div className="members-stat-icon" style={{ background: '#fef3c7', color: '#b45309' }}>
            👑
          </div>
          <div className="members-stat-info">
            <span className="members-stat-value">{counts.owner}</span>
            <span className="members-stat-label">Owner</span>
          </div>
        </div>

        <div className="members-stat-card">
          <div className="members-stat-icon" style={{ background: '#eff6ff', color: '#1d4ed8' }}>
            🛡️
          </div>
          <div className="members-stat-info">
            <span className="members-stat-value">{counts.admin}</span>
            <span className="members-stat-label">Administrators</span>
          </div>
        </div>

        <div className="members-stat-card">
          <div className="members-stat-icon" style={{ background: '#f0fdf4', color: '#15803d' }}>
            💼
          </div>
          <div className="members-stat-info">
            <span className="members-stat-value">{counts.staff}</span>
            <span className="members-stat-label">Staff Members</span>
          </div>
        </div>

        <div className="members-stat-card">
          <div className="members-stat-icon" style={{ background: '#fffbeb', color: '#d97706' }}>
            🏭
          </div>
          <div className="members-stat-info">
            <span className="members-stat-value">{counts.viewer}</span>
            <span className="members-stat-label">Floor Operators</span>
          </div>
        </div>
      </div>

      {/* Role Permissions Reference Guide */}
      <div className="role-guide-card">
        <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', flexWrap: 'wrap', gap: '8px' }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
            <span style={{ fontSize: '15px' }}>🛡️</span>
            <strong style={{ fontSize: '12.5px', color: '#0f172a' }}>Role Permissions &amp; Access Levels</strong>
          </div>
          <span className="tiny muted">Assigned roles control sheets, pricing visibility, and editing scope</span>
        </div>
        <div className="role-guide-grid">
          <div className="role-guide-item">
            <div className="role-guide-title">
              <span>👑</span>
              <span>Owner</span>
            </div>
            <div className="role-guide-desc">
              Full authority. Manages organization settings, billing, confidential profit margins, and provisions staff accounts.
            </div>
          </div>
          <div className="role-guide-item">
            <div className="role-guide-title">
              <span>🛡️</span>
              <span>Administrator</span>
            </div>
            <div className="role-guide-desc">
              Access to all 15 calculation sheets, custom formula overrides, template switching, and team role updates.
            </div>
          </div>
          <div className="role-guide-item">
            <div className="role-guide-title">
              <span>💼</span>
              <span>Staff Member</span>
            </div>
            <div className="role-guide-desc">
              Standard commercial access. Can create quotations, manage assigned client accounts, and export official proposal PDFs.
            </div>
          </div>
          <div className="role-guide-item">
            <div className="role-guide-title">
              <span>🏭</span>
              <span>Floor Operator</span>
            </div>
            <div className="role-guide-desc">
              Production shop floor access. Read-only specifications, body quantities, and dimensions with pricing hidden.
            </div>
          </div>
        </div>
      </div>

      {/* Owner-Only Email Provisioning Box */}
      {isOwner && (
        <div className="card p-3" style={{ borderLeft: '4px solid #2563eb', borderRadius: '10px', boxShadow: '0 1px 4px rgba(0, 0, 0, 0.04)' }}>
          <div className="row between border-b pb-2 mb-3">
            <div>
              <strong style={{ fontSize: '14.5px', color: '#0f172a', display: 'flex', alignItems: 'center', gap: '6px' }}>
                <span>✉️</span> Provision New Staff Account
              </strong>
              <div className="tiny muted mt-1" style={{ fontSize: '11.5px', color: '#64748b' }}>
                Directly create a staff account for {org.name}. A temporary password will be generated for immediate sign-in.
              </div>
            </div>
          </div>

          <div style={{ display: 'flex', gap: '10px', flexWrap: 'wrap', alignItems: 'center' }}>
            <div style={{ display: 'flex', alignItems: 'center', background: '#f8fafc', border: '1px solid #cbd5e1', borderRadius: '6px', padding: '0 10px', height: '36px', gap: '8px', flex: 1, minWidth: '260px' }}>
              <span style={{ color: '#64748b', fontSize: '14px' }}>✉️</span>
              <input
                type="email"
                placeholder="colleague@company.com"
                value={inviteEmail}
                onChange={(e) => setInviteEmail(e.target.value)}
                style={{ border: 'none', background: 'transparent', outline: 'none', width: '100%', fontSize: '13px', color: '#0f172a' }}
              />
            </div>

            <select
              value={inviteRole}
              onChange={(e) => setInviteRole(e.target.value as any)}
              className="role-select-custom"
              style={{ height: '36px', minWidth: '160px' }}
              title="Select Role for new account"
            >
              <option value="member">💼 Staff Member</option>
              <option value="admin">🛡️ Administrator</option>
              <option value="viewer">🏭 Floor Operator</option>
            </select>

            <button
              className="btn primary sm"
              disabled={inviting || !inviteEmail.includes('@')}
              onClick={handleSendInvite}
              style={{ height: '36px', padding: '0 16px', fontWeight: 700, whiteSpace: 'nowrap' }}
            >
              {inviting ? <Spinner size={13} /> : 'Add Member & Generate Credentials'}
            </button>
          </div>

          {/* Provisioned Credentials Success Card */}
          {createdUser && (
            <div className="mt-3 p-3" style={{ background: '#f0fdf4', borderRadius: '8px', border: '1px solid #bbf7d0' }}>
              <div className="row between mb-2">
                <div className="row gap" style={{ alignItems: 'center' }}>
                  <span style={{ fontSize: '18px' }}>🎉</span>
                  <strong style={{ color: '#166534', fontSize: '14px' }}>
                    {createdUser.is_new ? 'New Staff Account Provisioned!' : 'Staff Member Added!'}
                  </strong>
                </div>
                <button className="btn ghost sm" onClick={() => setCreatedUser(null)}>
                  ✕ Dismiss
                </button>
              </div>

              <div
                style={{
                  display: 'grid',
                  gridTemplateColumns: 'repeat(auto-fit, minmax(200px, 1fr))',
                  gap: '8px',
                  background: '#ffffff',
                  padding: '12px',
                  borderRadius: '6px',
                  border: '1px solid #dcfce7',
                }}
              >
                <div>
                  <div className="tiny muted font-bold" style={{ fontSize: '10px', textTransform: 'uppercase' }}>Login Email:</div>
                  <div className="font-mono text-sm" style={{ fontWeight: 600, color: '#0f172a' }}>
                    {createdUser.email}
                  </div>
                </div>
                <div>
                  <div className="tiny muted font-bold" style={{ fontSize: '10px', textTransform: 'uppercase' }}>Username:</div>
                  <div className="font-mono text-sm" style={{ fontWeight: 600, color: '#0f172a' }}>
                    {createdUser.username}
                  </div>
                </div>
                <div>
                  <div className="tiny muted font-bold" style={{ fontSize: '10px', textTransform: 'uppercase' }}>Temporary Password:</div>
                  <div className="font-mono text-sm" style={{ fontWeight: 700, color: '#2563eb' }}>
                    {createdUser.password || 'Existing Account Password'}
                  </div>
                </div>
              </div>

              <div className="row gap mt-3">
                <button className="btn primary sm" onClick={handleCopyCredentials} style={{ fontWeight: 600 }}>
                  {Icon.copy} Copy Login Details
                </button>
                <a className="btn ghost sm" href={getMailtoUrl()} target="_blank" rel="noopener noreferrer">
                  📬 Open in Mail App
                </a>
              </div>

              <div className="tiny muted mt-2" style={{ color: '#15803d' }}>
                ✓ The staff member can now sign in directly at the login screen using either their Email or Username with the password above.
              </div>
            </div>
          )}
        </div>
      )}

      {/* Members List Table Card */}
      <div className="card" style={{ borderRadius: '10px', overflow: 'hidden', boxShadow: '0 1px 4px rgba(0, 0, 0, 0.04)' }}>
        <div className="members-table-toolbar">
          <div style={{ display: 'flex', alignItems: 'center', gap: '10px', flexWrap: 'wrap' }}>
            <div className="members-search-wrapper">
              <span style={{ color: '#64748b', fontSize: '13px' }}>🔍</span>
              <input
                type="text"
                className="members-search-input"
                placeholder="Search by name or email…"
                value={searchQuery}
                onChange={(e) => setSearchQuery(e.target.value)}
              />
              {searchQuery && (
                <button
                  type="button"
                  onClick={() => setSearchQuery('')}
                  style={{ background: 'transparent', border: 'none', cursor: 'pointer', color: '#94a3b8', fontSize: '12px' }}
                >
                  ✕
                </button>
              )}
            </div>

            <div style={{ display: 'inline-flex', background: '#f1f5f9', padding: '2px', borderRadius: '6px', gap: '2px' }}>
              {(['all', 'admin', 'member', 'viewer'] as const).map((r) => {
                const isActive = roleFilter === r
                const label =
                  r === 'all' ? `All (${counts.total})` :
                  r === 'admin' ? `Admin (${counts.admin})` :
                  r === 'member' ? `Staff (${counts.staff})` :
                  `Operator (${counts.viewer})`
                return (
                  <button
                    key={r}
                    type="button"
                    onClick={() => setRoleFilter(r)}
                    style={{
                      border: 'none',
                      background: isActive ? '#ffffff' : 'transparent',
                      color: isActive ? '#0f172a' : '#64748b',
                      fontWeight: isActive ? 700 : 500,
                      fontSize: '11px',
                      padding: '4px 8px',
                      borderRadius: '4px',
                      cursor: 'pointer',
                      boxShadow: isActive ? '0 1px 2px rgba(0,0,0,0.06)' : 'none',
                    }}
                  >
                    {label}
                  </button>
                )
              })}
            </div>
          </div>

          <span className="tiny muted font-mono" style={{ fontSize: '11px' }}>
            Showing {filteredMembers.length} of {members.length} member{members.length !== 1 ? 's' : ''}
          </span>
        </div>

        <div className="table-wrap">
          {loading ? (
            <div className="page-loading" style={{ padding: '40px', textAlign: 'center' }}>
              <Spinner size={24} />
            </div>
          ) : filteredMembers.length === 0 ? (
            <div className="empty" style={{ padding: '36px', textAlign: 'center' }}>
              <p className="tiny muted">No members matching current search/filter.</p>
            </div>
          ) : (
            <table className="tbl" style={{ width: '100%', borderCollapse: 'collapse' }}>
              <thead>
                <tr style={{ background: '#f8fafc', borderBottom: '1px solid #e2e8f0', color: '#475569', fontSize: '11px', textTransform: 'uppercase', letterSpacing: '0.4px' }}>
                  <th style={{ padding: '10px 16px', textAlign: 'left' }}>Team Member</th>
                  <th style={{ padding: '10px 16px', textAlign: 'left' }}>Role &amp; Permissions</th>
                  <th style={{ padding: '10px 16px', textAlign: 'left' }}>Joined</th>
                  {isAdminOrOwner && <th style={{ padding: '10px 16px', textAlign: 'right' }}>Actions</th>}
                </tr>
              </thead>
              <tbody>
                {filteredMembers.map((m) => {
                  const isCurrentSessionUser = m.username === session.username
                  const avatarBg = getAvatarBg(m.username, m.role)
                  return (
                    <tr key={m.username} style={{ borderBottom: '1px solid #f1f5f9' }}>
                      <td style={{ padding: '12px 16px' }}>
                        <div className="row gap" style={{ alignItems: 'center' }}>
                          <div className="member-avatar" style={{ background: avatarBg }}>
                            {(m.username[0] || 'U').toUpperCase()}
                          </div>
                          <div>
                            <div className="row gap" style={{ alignItems: 'center' }}>
                              <strong style={{ fontSize: '13px', color: '#0f172a' }}>{m.username}</strong>
                              {isCurrentSessionUser && (
                                <span className="chip ok ml-2" style={{ fontSize: '10px', padding: '1px 6px' }}>
                                  You
                                </span>
                              )}
                            </div>
                            {m.email && (
                              <div className="tiny muted" style={{ fontSize: '11.5px', color: '#64748b' }}>
                                {m.email}
                              </div>
                            )}
                          </div>
                        </div>
                      </td>

                      <td style={{ padding: '12px 16px' }}>
                        {m.role === 'owner' ? (
                          <span className="role-badge-owner">👑 OWNER</span>
                        ) : isAdminOrOwner ? (
                          <select
                            className="role-select-custom"
                            value={m.role === 'floor_operator' ? 'viewer' : m.role}
                            onChange={(e) => handleRoleChange(m, e.target.value)}
                            title={`Change role for ${m.username}`}
                          >
                            <option value="admin">🛡️ Administrator</option>
                            <option value="member">💼 Staff Member</option>
                            <option value="viewer">🏭 Floor Operator</option>
                          </select>
                        ) : (
                          <span
                            style={
                              m.role === 'admin'
                                ? { background: '#eff6ff', color: '#1d4ed8', border: '1px solid #bfdbfe', fontWeight: 700, fontSize: '11px', padding: '3px 8px', borderRadius: '4px', display: 'inline-flex', alignItems: 'center', gap: '4px' }
                                : m.role === 'viewer' || (m.role as string) === 'floor_operator'
                                ? { background: '#fffbeb', color: '#b45309', border: '1px solid #fde68a', fontWeight: 700, fontSize: '11px', padding: '3px 8px', borderRadius: '4px', display: 'inline-flex', alignItems: 'center', gap: '4px' }
                                : { background: '#f0fdf4', color: '#15803d', border: '1px solid #bbf7d0', fontWeight: 700, fontSize: '11px', padding: '3px 8px', borderRadius: '4px', display: 'inline-flex', alignItems: 'center', gap: '4px' }
                            }
                          >
                            {m.role === 'admin' ? '🛡️ ADMIN' : m.role === 'viewer' || (m.role as string) === 'floor_operator' ? '🏭 FLOOR OPERATOR' : '💼 STAFF MEMBER'}
                          </span>
                        )}
                      </td>

                      <td style={{ padding: '12px 16px', fontSize: '12px', color: '#64748b' }}>
                        {timeAgo(m.joined_at)}
                      </td>

                      {isAdminOrOwner && (
                        <td style={{ padding: '12px 16px', textAlign: 'right' }}>
                          {m.role !== 'owner' && !isCurrentSessionUser && (
                            <button
                              className="btn ghost sm danger"
                              onClick={() => handleRemoveMember(m)}
                              title={`Remove ${m.username} from organisation`}
                              style={{ padding: '4px 8px', height: '28px' }}
                            >
                              {Icon.trash}
                            </button>
                          )}
                        </td>
                      )}
                    </tr>
                  )
                })}
              </tbody>
            </table>
          )}
        </div>
      </div>
    </div>
  )
}
