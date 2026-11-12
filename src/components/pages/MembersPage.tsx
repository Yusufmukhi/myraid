function generateRandom10CharPassword(): string {
  const upper = 'ABCDEFGHJKLMNPQRSTUVWXYZ'
  const lower = 'abcdefghjkmnpqrstuvwxyz'
  const digits = '23456789'
  const allChars = upper + lower + digits

  const chars = [
    upper[Math.floor(Math.random() * upper.length)],
    upper[Math.floor(Math.random() * upper.length)],
    lower[Math.floor(Math.random() * lower.length)],
    lower[Math.floor(Math.random() * lower.length)],
    digits[Math.floor(Math.random() * digits.length)],
    digits[Math.floor(Math.random() * digits.length)],
  ]

  while (chars.length < 10) {
    chars.push(allChars[Math.floor(Math.random() * allChars.length)])
  }

  for (let i = chars.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1))
    ;[chars[i], chars[j]] = [chars[j], chars[i]]
  }

  return chars.join('')
}

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

  // Password change state
  const [changePasswordMember, setChangePasswordMember] = useState<Member | null>(null)
  const [newMemberPassword, setNewMemberPassword] = useState('')
  const [savingPassword, setSavingPassword] = useState(false)

  // Member deletion with admin password confirmation state
  const [deleteMemberTarget, setDeleteMemberTarget] = useState<Member | null>(null)
  const [adminConfirmPassword, setAdminConfirmPassword] = useState('')
  const [deletingMember, setDeletingMember] = useState(false)

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
          ? 'Admin'
          : 'Staff Member'
      toast(`Updated ${member.username} role to ${roleLabel}`)
      loadMembers()
    } catch (e) {
      toast((e as Error).message, 'err')
    }
  }

  // Deletion with required Admin/Owner password verification
  const handleConfirmDeleteMember = async () => {
    if (!deleteMemberTarget?.id) return
    if (!adminConfirmPassword.trim()) {
      return toast('Please enter your password to confirm removal.', 'err')
    }

    setDeletingMember(true)
    try {
      // 1. Re-authenticate admin/owner with their password
      try {
        await api.login(session.username, adminConfirmPassword.trim())
      } catch {
        toast('Incorrect password. Member was not removed.', 'err')
        setDeletingMember(false)
        return
      }

      // 2. Verified! Remove member from organisation
      await api.orgRemoveMember(session.token, org.id, deleteMemberTarget.id)
      toast(`Removed ${deleteMemberTarget.username} from ${org.name}`)
      setDeleteMemberTarget(null)
      setAdminConfirmPassword('')
      loadMembers()
    } catch (e) {
      toast((e as Error).message, 'err')
    } finally {
      setDeletingMember(false)
    }
  }

  // Password change for a member by Admin/Owner
  const handleSaveMemberPassword = async () => {
    if (!changePasswordMember?.id) return
    const pwd = newMemberPassword.trim()
    if (pwd.length < 10) {
      return toast('Password must be at least 10 characters long.', 'err')
    }

    setSavingPassword(true)
    try {
      await api.orgAdminChangeMemberPassword(session.token, org.id, changePasswordMember.id, pwd)
      toast(`Password updated for ${changePasswordMember.username}!`)
      copy(
        `Login credentials for ${changePasswordMember.username}:\nPortal: ${window.location.origin}\nUsername: ${changePasswordMember.username}\nNew Password: ${pwd}`,
        `New password copied to clipboard!`
      )
      setChangePasswordMember(null)
      setNewMemberPassword('')
    } catch (e) {
      toast((e as Error).message, 'err')
    } finally {
      setSavingPassword(false)
    }
  }

  // Provisioning new account (Owners and Admins have same access)
  const handleSendInvite = async () => {
    if (!isAdminOrOwner) return toast('Admin privileges required to provision accounts.', 'err')
    if (!inviteEmail || !inviteEmail.includes('@')) return toast('Enter a valid email address.', 'err')
    setInviting(true)
    const targetEmail = inviteEmail.trim()
    try {
      const res = await api.orgInviteMember(session.token, org.id, targetEmail, inviteRole)
      let finalPassword = res.password

      if (res.is_new && res.password) {
        const clean10CharPassword = generateRandom10CharPassword()
        try {
          const tempSession = await api.login(res.username, res.password)
          if (tempSession && tempSession.token) {
            await api.updateProfile(tempSession.token, undefined, clean10CharPassword)
            await api.logout(tempSession.token)
            finalPassword = clean10CharPassword
          }
        } catch (pwErr) {
          console.warn('Could not auto-update to clean 10-char password:', pwErr)
        }
      }

      setCreatedUser({ ...res, password: finalPassword })
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
    return `Hello,\n\nYou have been added to ${createdUser.org_name || org.name} Quotation Portal.\n\nPortal URL: ${loginUrl}\nLogin Email: ${createdUser.email}\nUsername: ${createdUser.username}\n${createdUser.password ? `Temporary Password: ${createdUser.password}\n\nPlease sign in and update your profile anytime.` : 'You can sign in using your existing account credentials.'}\n\nBest regards,\n${session.username} (${org.name})`
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
      body: 'You will need an invitation to access this organisation again.',
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
          <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
            <h2 style={{ fontSize: '18px', fontWeight: 800, margin: 0, color: '#0f172a' }}>
              👥 Team Members &amp; Access
            </h2>
            <span className="chip primary font-bold" style={{ textTransform: 'uppercase' }}>
              {org.name}
            </span>
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
            <span className="members-stat-label">Admins</span>
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

      {/* Provisioning Box (Available to both Owner & Admin) */}
      {isAdminOrOwner && (
        <div className="card p-3" style={{ borderLeft: '4px solid #2563eb', borderRadius: '10px', boxShadow: '0 1px 4px rgba(0, 0, 0, 0.04)' }}>
          <div className="border-b pb-2 mb-3">
            <strong style={{ fontSize: '14px', color: '#0f172a', display: 'flex', alignItems: 'center', gap: '6px' }}>
              <span>✉️</span> Add Team Member
            </strong>
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
              <option value="admin">🛡️ Admin</option>
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
                    {createdUser.is_new ? 'New Account Provisioned!' : 'Team Member Added!'}
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
                            <option value="admin">🛡️ Admin</option>
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
                          <div style={{ display: 'inline-flex', alignItems: 'center', gap: '6px' }}>
                            {/* Change Password button for Admin/Owner */}
                            {m.id && (
                              <button
                                className="btn ghost sm"
                                onClick={() => {
                                  setChangePasswordMember(m)
                                  setNewMemberPassword(generateRandom10CharPassword())
                                }}
                                title={`Change password for ${m.username}`}
                                style={{ padding: '4px 8px', height: '28px', color: '#2563eb', border: '1px solid #cbd5e1' }}
                              >
                                🔑 Password
                              </button>
                            )}

                            {/* Delete Member button (requires current admin password) */}
                            {m.role !== 'owner' && !isCurrentSessionUser && (
                              <button
                                className="btn ghost sm danger"
                                onClick={() => {
                                  setDeleteMemberTarget(m)
                                  setAdminConfirmPassword('')
                                }}
                                title={`Remove ${m.username} from organisation`}
                                style={{ padding: '4px 8px', height: '28px' }}
                              >
                                {Icon.trash}
                              </button>
                            )}
                          </div>
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

      {/* Change Member Password Modal */}
      {changePasswordMember && (
        <div
          style={{
            position: 'fixed',
            top: 0,
            left: 0,
            right: 0,
            bottom: 0,
            background: 'rgba(15, 23, 42, 0.6)',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            zIndex: 100000,
            padding: '16px',
          }}
          onClick={() => !savingPassword && setChangePasswordMember(null)}
        >
          <div
            style={{
              background: '#ffffff',
              borderRadius: '10px',
              border: '1px solid #cbd5e1',
              padding: '24px',
              width: '440px',
              maxWidth: '96vw',
              boxShadow: '0 20px 25px -5px rgba(0, 0, 0, 0.2)',
            }}
            onClick={(e) => e.stopPropagation()}
          >
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '16px' }}>
              <h3 style={{ margin: 0, fontSize: '16px', fontWeight: 800, color: '#0f172a', display: 'flex', alignItems: 'center', gap: '8px' }}>
                <span>🔑</span> Change Password
              </h3>
              <button
                type="button"
                className="icon-btn sm"
                onClick={() => setChangePasswordMember(null)}
                disabled={savingPassword}
                style={{ fontSize: '16px', lineHeight: 1 }}
              >
                ✕
              </button>
            </div>

            <div style={{ background: '#f8fafc', padding: '10px 12px', borderRadius: '6px', marginBottom: '14px', border: '1px solid #e2e8f0' }}>
              <div style={{ fontSize: '13px', fontWeight: 700, color: '#0f172a' }}>
                {changePasswordMember.username}
              </div>
              {changePasswordMember.email && (
                <div style={{ fontSize: '11.5px', color: '#64748b' }}>
                  {changePasswordMember.email}
                </div>
              )}
            </div>

            <div style={{ marginBottom: '16px' }}>
              <label className="tiny font-bold muted" style={{ display: 'block', marginBottom: '4px' }}>
                New Password (minimum 10 characters)
              </label>
              <div style={{ display: 'flex', gap: '6px' }}>
                <input
                  type="text"
                  className="option-input-field font-mono"
                  value={newMemberPassword}
                  onChange={(e) => setNewMemberPassword(e.target.value)}
                  placeholder="Enter new password…"
                  style={{ flex: 1, height: '36px', padding: '0 10px', fontSize: '13px' }}
                  autoFocus
                />
                <button
                  type="button"
                  className="btn ghost sm"
                  onClick={() => setNewMemberPassword(generateRandom10CharPassword())}
                  title="Generate 10-char random secure password"
                  style={{ fontSize: '11px', whiteSpace: 'nowrap' }}
                >
                  🎲 Auto Generate
                </button>
              </div>
              {newMemberPassword.length > 0 && newMemberPassword.length < 10 && (
                <div style={{ fontSize: '11px', color: '#dc2626', marginTop: '4px' }}>
                  Password must be at least 10 characters ({newMemberPassword.length}/10)
                </div>
              )}
            </div>

            <div style={{ display: 'flex', justifyContent: 'flex-end', gap: '8px' }}>
              <button
                type="button"
                className="btn ghost sm"
                onClick={() => setChangePasswordMember(null)}
                disabled={savingPassword}
              >
                Cancel
              </button>
              <button
                type="button"
                className="btn primary sm"
                onClick={handleSaveMemberPassword}
                disabled={savingPassword || newMemberPassword.trim().length < 10}
                style={{ fontWeight: 700 }}
              >
                {savingPassword ? <Spinner size={13} /> : 'Save New Password'}
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Security Confirm Delete Member Modal (Password Verification) */}
      {deleteMemberTarget && (
        <div
          style={{
            position: 'fixed',
            top: 0,
            left: 0,
            right: 0,
            bottom: 0,
            background: 'rgba(15, 23, 42, 0.6)',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            zIndex: 100000,
            padding: '16px',
          }}
          onClick={() => !deletingMember && setDeleteMemberTarget(null)}
        >
          <div
            style={{
              background: '#ffffff',
              borderRadius: '10px',
              border: '1px solid #cbd5e1',
              padding: '24px',
              width: '440px',
              maxWidth: '96vw',
              boxShadow: '0 20px 25px -5px rgba(0, 0, 0, 0.2)',
            }}
            onClick={(e) => e.stopPropagation()}
          >
            <div style={{ display: 'flex', alignItems: 'center', gap: '8px', color: '#dc2626', marginBottom: '12px' }}>
              <span style={{ fontSize: '20px' }}>⚠️</span>
              <h3 style={{ margin: 0, fontSize: '16px', fontWeight: 800, color: '#dc2626' }}>
                Confirm Member Removal
              </h3>
            </div>

            <p style={{ fontSize: '13px', color: '#334155', margin: '0 0 14px 0', lineHeight: 1.5 }}>
              You are removing <strong>{deleteMemberTarget.username}</strong> from <strong>{org.name}</strong>.
              They will immediately lose access to this organisation's projects and quotations.
            </p>

            <div style={{ marginBottom: '16px' }}>
              <label className="tiny font-bold muted" style={{ display: 'block', marginBottom: '6px' }}>
                Enter YOUR password to confirm removal:
              </label>
              <input
                type="password"
                className="option-input-field"
                value={adminConfirmPassword}
                onChange={(e) => setAdminConfirmPassword(e.target.value)}
                placeholder="Your account password…"
                style={{ width: '100%', height: '36px', padding: '0 10px', fontSize: '13px' }}
                autoFocus
                onKeyDown={(e) => {
                  if (e.key === 'Enter' && adminConfirmPassword.trim()) {
                    handleConfirmDeleteMember()
                  }
                }}
              />
            </div>

            <div style={{ display: 'flex', justifyContent: 'flex-end', gap: '8px' }}>
              <button
                type="button"
                className="btn ghost sm"
                onClick={() => {
                  setDeleteMemberTarget(null)
                  setAdminConfirmPassword('')
                }}
                disabled={deletingMember}
              >
                Cancel
              </button>
              <button
                type="button"
                className="btn danger sm"
                onClick={handleConfirmDeleteMember}
                disabled={deletingMember || !adminConfirmPassword.trim()}
                style={{ fontWeight: 700 }}
              >
                {deletingMember ? <Spinner size={13} /> : 'Verify Password & Remove'}
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  )
}
