import { useEffect, useState } from 'react'
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

  const handleRoleChange = async (member: Member, newRole: string) => {
    if (!isAdminOrOwner) return toast('Admin privileges required', 'err')
    if (member.role === 'owner') return toast('Cannot change owner role', 'err')
    try {
      if (member.id) await api.orgUpdateRole(session.token, org.id, member.id, newRole)
      const roleLabel = newRole === 'viewer' || newRole === 'floor_operator' ? 'Floor Operator' : newRole
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
      body: `Revoke access for ${member.username} from this organisation?`,
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
    return `Hello,\n\nYou have been added to ${createdUser.org_name || org.name} on the Apollo Quotation Portal.\n\nPortal URL: ${loginUrl}\nLogin Email: ${createdUser.email}\nUsername: ${createdUser.username}\n${createdUser.password ? `Temporary Password: ${createdUser.password}\n\nPlease sign in and change your username and password anytime in My Profile.` : 'You can sign in using your existing account credentials.'}\n\nBest regards,\n${session.username} (${org.name})`
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
      confirmLabel: 'Leave',
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

  return (
    <div className="members-page">
      {/* Streamlined Header */}
      <div className="card row between p-3 mb-3">
        <div>
          <div className="row gap" style={{ alignItems: 'center' }}>
            <h2 style={{ fontSize: '18px', margin: 0 }}>Team Members ({members.length})</h2>
            <span className="chip primary">{org.name}</span>
          </div>
          <div className="tiny muted mt-1">
            🔒 Direct staff provisioning: staff accounts are created by the Owner and join this organisation directly.
          </div>
        </div>

        {!isOwner && (
          <button className="btn ghost danger sm" onClick={handleLeave}>
            {Icon.logout} Leave Org
          </button>
        )}
      </div>

      {/* Owner-Only Email Provisioning Box */}
      {isOwner && (
        <div className="card mb-3 p-3" style={{ borderLeft: '4px solid var(--blue, #2563eb)' }}>
          <div className="row between border-b pb-2 mb-3">
            <div>
              <strong style={{ fontSize: '15px' }}>✉️ Add Staff Member by Email</strong>
              <div className="tiny muted">
                Enter staff member's email. An account with a random username and secure temporary password will be created, and they will directly join {org.name}.
              </div>
            </div>
          </div>

          <div className="row gap" style={{ flexWrap: 'wrap' }}>
            <input
              type="email"
              placeholder="colleague@company.com"
              value={inviteEmail}
              onChange={(e) => setInviteEmail(e.target.value)}
              style={{ minWidth: 260 }}
            />
            <select
              value={inviteRole}
              onChange={(e) => setInviteRole(e.target.value as any)}
              className="role-select"
            >
              <option value="member">Staff Member</option>
              <option value="viewer">Floor Operator</option>
              <option value="admin">Administrator</option>
            </select>
            <button
              className="btn primary sm"
              disabled={inviting || !inviteEmail.includes('@')}
              onClick={handleSendInvite}
            >
              {inviting ? <Spinner size={13} /> : 'Add Member & Generate Login'}
            </button>
          </div>

          {/* Provisioned Credentials Success Card */}
          {createdUser && (
            <div className="mt-3 p-3" style={{ background: '#f0fdf4', borderRadius: 8, border: '1px solid #bbf7d0' }}>
              <div className="row between mb-2">
                <div className="row gap" style={{ alignItems: 'center' }}>
                  <span style={{ fontSize: '16px' }}>🎉</span>
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
                  <div className="tiny muted font-bold">Email (Login):</div>
                  <div className="font-mono text-sm" style={{ fontWeight: 600, color: '#0f172a' }}>
                    {createdUser.email}
                  </div>
                </div>
                <div>
                  <div className="tiny muted font-bold">Username:</div>
                  <div className="font-mono text-sm" style={{ fontWeight: 600, color: '#0f172a' }}>
                    {createdUser.username}
                  </div>
                </div>
                <div>
                  <div className="tiny muted font-bold">Temporary Password:</div>
                  <div className="font-mono text-sm" style={{ fontWeight: 700, color: '#2563eb' }}>
                    {createdUser.password || 'Existing Account Password'}
                  </div>
                </div>
              </div>

              <div className="row gap mt-3">
                <button className="btn primary sm" onClick={handleCopyCredentials}>
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

      {/* Members List Table */}
      <div className="card">
        <div className="table-wrap">
          {loading ? (
            <div className="page-loading">
              <Spinner size={24} />
            </div>
          ) : members.length === 0 ? (
            <div className="empty">
              <p className="tiny muted">No members found</p>
            </div>
          ) : (
            <table className="tbl">
              <thead>
                <tr>
                  <th>Member</th>
                  <th>Role</th>
                  <th>Joined</th>
                  {isAdminOrOwner && <th className="r">Actions</th>}
                </tr>
              </thead>
              <tbody>
                {members.map((m) => {
                  const isCurrentSessionUser = m.username === session.username
                  return (
                    <tr key={m.username}>
                      <td>
                        <div className="row gap" style={{ alignItems: 'center' }}>
                          <div className="avatar">{(m.username[0] || 'U').toUpperCase()}</div>
                          <div>
                            <div className="row gap" style={{ alignItems: 'center' }}>
                              <strong>{m.username}</strong>
                              {isCurrentSessionUser && <span className="chip ok ml-2">You</span>}
                            </div>
                            {m.email && <div className="tiny muted" style={{ fontSize: '11px' }}>{m.email}</div>}
                          </div>
                        </div>
                      </td>
                      <td>
                        {isAdminOrOwner && m.role !== 'owner' ? (
                          <select
                            className="role-select"
                            value={m.role}
                            onChange={(e) => handleRoleChange(m, e.target.value)}
                          >
                            <option value="admin">Admin</option>
                            <option value="member">Staff Member</option>
                            <option value="viewer">Floor Operator</option>
                          </select>
                        ) : (
                          <span className={`status-pill status-${m.role}`}>
                            {m.role === 'viewer' || (m.role as string) === 'floor_operator' ? 'FLOOR OPERATOR' : m.role.toUpperCase()}
                          </span>
                        )}
                      </td>
                      <td className="tiny muted">{timeAgo(m.joined_at)}</td>
                      {isAdminOrOwner && (
                        <td className="r">
                          {m.role !== 'owner' && !isCurrentSessionUser && (
                            <button
                              className="btn ghost sm danger"
                              onClick={() => handleRemoveMember(m)}
                              title="Remove member"
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

