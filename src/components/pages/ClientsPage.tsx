import React, { useEffect, useMemo, useState } from 'react'
import { Member, Org, Session, api } from '../../lib/api'
import {
  countUniqueSharedStaff,
  deleteClientFromDirectory,
  fetchClientsFromSupabase,
  getSavedClients,
  getVisibleClientsForUser,
  saveClientToDirectory,
  SavedClient,
  ClientContactPerson,
  QuotationExtraField,
} from '../../lib/clientDirectory'
import { Icon, useCopy, useUI } from '../ui'

interface Props {
  session: Session
  org: Org
  onCreateQuotationWithClient?: (client: SavedClient) => void
  onNavigate?: (tab: string) => void
}

const CLIENTS_PER_PAGE = 9

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

function getInitials(name: string): string {
  const words = (name || '')
    .trim()
    .split(/\s+/)
    .filter(Boolean)
  if (words.length === 0) return 'CL'
  if (words.length === 1) return words[0].slice(0, 2).toUpperCase()
  return (words[0][0] + words[1][0]).toUpperCase()
}

export function ClientsPage({ session, org, onCreateQuotationWithClient, onNavigate }: Props) {
  const { toast, confirm } = useUI()
  const copy = useCopy()

  const isOwnerOrAdmin =
    org.role === 'owner' ||
    org.role === 'admin' ||
    session.username === 'owner' ||
    session.username === 'admin'

  const [clients, setClients] = useState<SavedClient[]>(() => getSavedClients(org.id))
  const [searchQuery, setSearchQuery] = useState('')
  const [techFilter, setTechFilter] = useState<'all' | 'tech_yes' | 'tech_no' | 'shared'>('all')
  const [currentPage, setCurrentPage] = useState(1)
  const [editingClient, setEditingClient] = useState<SavedClient | null>(null)
  const [isCreatingNew, setIsCreatingNew] = useState(false)
  const [orgStaffMembers, setOrgStaffMembers] = useState<Member[]>([])
  const [customShareStaffInput, setCustomShareStaffInput] = useState('')

  useEffect(() => {
    if (org.role !== 'viewer' && session.token && org.id) {
      api
        .orgMembers(session.token, org.id)
        .then((members) => {
          if (Array.isArray(members)) {
            setOrgStaffMembers(
              members.filter((m) => m.username !== session.username && m.role !== 'viewer'),
            )
          }
        })
        .catch(() => {})
    }
  }, [org.role, session.token, org.id, session.username])

  const canManageClientShare =
    isOwnerOrAdmin ||
    isCreatingNew ||
    Boolean(
      editingClient &&
        ((session.username && editingClient.created_by === session.username) ||
          (session.user_id && editingClient.created_by === session.user_id)),
    )

  const [formData, setFormData] = useState<{
    name: string
    address: string
    gst: string
    email: string
    kind_att: string
    contact: string
    contacts: ClientContactPerson[]
    extra_fields: QuotationExtraField[]
    include_full_proposal: boolean
    profit_margin_pct: number
    shared_with: string[]
  }>({
    name: '',
    address: '',
    gst: '',
    email: '',
    kind_att: '',
    contact: '',
    contacts: [],
    extra_fields: [],
    include_full_proposal: false,
    profit_margin_pct: 25,
    shared_with: [],
  })

  useEffect(() => {
    fetchClientsFromSupabase(session.token, org.id)
    const handleUpdate = () => {
      setClients(getSavedClients(org.id))
    }
    window.addEventListener('examcalc:clients-updated', handleUpdate)
    return () => window.removeEventListener('examcalc:clients-updated', handleUpdate)
  }, [session.token, org.id])

  const userVisibleClients = useMemo(() => {
    const visible = getVisibleClientsForUser(clients, session.username, org.role, session.user_id) || []
    return [...visible].sort((a, b) =>
      (a.name || '').localeCompare(b.name || '', undefined, { sensitivity: 'base' }),
    )
  }, [clients, session.username, org.role, session.user_id])

  const filteredClients = useMemo(() => {
    const q = searchQuery.trim().toLowerCase()
    return userVisibleClients.filter((c) => {
      if (techFilter === 'tech_yes' && c.include_full_proposal !== true) return false
      if (techFilter === 'tech_no' && c.include_full_proposal === true) return false
      if (techFilter === 'shared') {
        const isShared = Array.isArray(c.shared_with) && c.shared_with.length > 0
        if (!isShared) return false
      }
      if (!q) return true
      return (
        (c.name || '').toLowerCase().includes(q) ||
        (c.gst || '').toLowerCase().includes(q) ||
        (c.kind_att || '').toLowerCase().includes(q) ||
        (c.address || '').toLowerCase().includes(q) ||
        (c.contact || '').toLowerCase().includes(q) ||
        (c.email || '').toLowerCase().includes(q) ||
        (c.contacts || []).some(
          (cp) => cp.name.toLowerCase().includes(q) || cp.phone.toLowerCase().includes(q),
        ) ||
        (c.extra_fields || []).some(
          (ef) =>
            (ef.label || '').toLowerCase().includes(q) ||
            (ef.value || '').toLowerCase().includes(q),
        )
      )
    })
  }, [userVisibleClients, searchQuery, techFilter])

  // Reset to page 1 when search or filter changes
  useEffect(() => {
    setCurrentPage(1)
  }, [searchQuery, techFilter])

  const totalPages = Math.max(1, Math.ceil(filteredClients.length / CLIENTS_PER_PAGE))
  const safePage = Math.min(currentPage, totalPages)
  const paginatedClients = useMemo(() => {
    const start = (safePage - 1) * CLIENTS_PER_PAGE
    return filteredClients.slice(start, start + CLIENTS_PER_PAGE)
  }, [filteredClients, safePage])

  const techYesCount = useMemo(
    () => userVisibleClients.filter((c) => c.include_full_proposal === true).length,
    [userVisibleClients],
  )

  const handleOpenClient = (client: SavedClient) => {
    setIsCreatingNew(false)
    setEditingClient(client)

    const rawContacts = Array.isArray(client.contacts) ? client.contacts : []
    const cleanContacts = rawContacts.filter(
      (cp) => !(cp.name === client.kind_att && cp.phone === client.contact && cp.name.trim().length > 0),
    )

    const rawExtra = Array.isArray(client.extra_fields) ? client.extra_fields : []

    setFormData({
      name: client.name || '',
      address: client.address || '',
      gst: client.gst || '',
      email: client.email || '',
      kind_att: client.kind_att || '',
      contact: client.contact || '',
      contacts: cleanContacts,
      extra_fields: rawExtra.map((ef) => {
        const val = ef.value || (Array.isArray(ef.values) && ef.values[0]) || ''
        return {
          id:
            ef.id ||
            (typeof crypto !== 'undefined' && crypto.randomUUID
              ? `ef_${crypto.randomUUID()}`
              : `ef_${Date.now()}`),
          label: ef.label || '',
          value: val,
          values: [val],
          visible: ef.visible !== false,
          visible_values: [ef.visible !== false],
        }
      }),
      include_full_proposal: client.include_full_proposal === true,
      profit_margin_pct: client.profit_margin_pct !== undefined ? client.profit_margin_pct : 25,
      shared_with: Array.isArray(client.shared_with) ? client.shared_with : [],
    })
  }

  const handleStartCreateNew = () => {
    setIsCreatingNew(true)
    setEditingClient(null)
    setFormData({
      name: '',
      address: '',
      gst: '',
      email: '',
      kind_att: '',
      contact: '',
      contacts: [],
      extra_fields: [],
      include_full_proposal: false,
      profit_margin_pct: 25,
      shared_with: [],
    })
    setCustomShareStaffInput('')
  }

  const handleCloseModal = () => {
    setEditingClient(null)
    setIsCreatingNew(false)
  }

  const handleAddContact = () => {
    const newCp: ClientContactPerson = {
      id:
        typeof crypto !== 'undefined' && crypto.randomUUID
          ? `cp_${crypto.randomUUID()}`
          : `cp_${Date.now()}`,
      name: '',
      phone: '',
      visible: true,
    }
    setFormData((prev) => ({
      ...prev,
      contacts: [...prev.contacts, newCp],
    }))
  }

  const handleRemoveContact = (idx: number) => {
    setFormData((prev) => ({
      ...prev,
      contacts: prev.contacts.filter((_, i) => i !== idx),
    }))
  }

  const handleUpdateContact = (idx: number, field: 'name' | 'phone', val: string) => {
    setFormData((prev) => ({
      ...prev,
      contacts: prev.contacts.map((cp, i) => (i === idx ? { ...cp, [field]: val } : cp)),
    }))
  }

  const handleAddCustomRow = () => {
    const newEf: QuotationExtraField = {
      id:
        typeof crypto !== 'undefined' && crypto.randomUUID
          ? `ef_${crypto.randomUUID()}`
          : `ef_${Date.now()}`,
      label: '',
      value: '',
      values: [''],
      visible: true,
      visible_values: [true],
    }
    setFormData((prev) => ({
      ...prev,
      extra_fields: [...prev.extra_fields, newEf],
    }))
  }

  const handleRemoveCustomRow = (idx: number) => {
    setFormData((prev) => ({
      ...prev,
      extra_fields: prev.extra_fields.filter((_, i) => i !== idx),
    }))
  }

  const handleUpdateCustomRow = (idx: number, field: 'label' | 'value' | 'visible', val: any) => {
    setFormData((prev) => ({
      ...prev,
      extra_fields: prev.extra_fields.map((ef, i) => {
        if (i !== idx) return ef
        if (field === 'label') return { ...ef, label: String(val) }
        if (field === 'value') return { ...ef, value: String(val), values: [String(val)] }
        if (field === 'visible') return { ...ef, visible: Boolean(val), visible_values: [Boolean(val)] }
        return ef
      }),
    }))
  }

  const handleSave = () => {
    const cleanName = formData.name.trim()
    if (!cleanName) {
      toast('Please enter a Buyer / Company Name.', 'err')
      return
    }

    const payload = {
      name: cleanName,
      address: formData.address || '',
      gst: formData.gst || '',
      email: formData.email || '',
      kind_att: formData.kind_att || '',
      contact: formData.contact || '',
      contacts: formData.contacts.filter((cp) => cp.name.trim() || cp.phone.trim()),
      extra_fields: formData.extra_fields
        .filter((ef) => ef.label.trim() || ef.value.trim())
        .map((ef) => ({
          ...ef,
          value: ef.value.trim(),
          values: [ef.value.trim()],
          visible: ef.visible !== false,
          visible_values: [ef.visible !== false],
        })),
      include_full_proposal: Boolean(formData.include_full_proposal),
      profit_margin_pct: isOwnerOrAdmin
        ? formData.profit_margin_pct !== undefined
          ? formData.profit_margin_pct
          : 25
        : editingClient?.profit_margin_pct !== undefined
        ? editingClient.profit_margin_pct
        : 25,
      shared_with: canManageClientShare
        ? Array.isArray(formData.shared_with)
          ? formData.shared_with
          : []
        : editingClient?.shared_with || [],
    }

    if (isCreatingNew) {
      const saved = saveClientToDirectory(payload, org.id, session.username, org.role)
      toast(`Client "${saved.name}" created!`)
      handleCloseModal()
    } else if (editingClient) {
      const saved = saveClientToDirectory(
        {
          ...payload,
          id: editingClient.id,
        },
        org.id,
        session.username,
        org.role,
      )
      toast(`Client "${saved.name}" updated!`)
      handleCloseModal()
    }
  }

  const handleDelete = async () => {
    if (!editingClient) return
    const ok = await confirm({
      title: `Delete Client "${editingClient.name}"?`,
      body: 'This client profile will be removed from your saved directory. Existing quotations will keep their saved details.',
      confirmLabel: 'Delete Client',
      danger: true,
    })
    if (!ok) return

    deleteClientFromDirectory(editingClient.id, org.id)
    toast(`Client "${editingClient.name}" removed from directory.`)
    handleCloseModal()
  }

  const handleCopyDetails = () => {
    if (!formData.name) return
    const contactLines = formData.contacts.map((cp) => `Contact: ${cp.name} (${cp.phone})`)
    const customLines = formData.extra_fields.map((ef) => `${ef.label}: ${ef.value}`)
    const text = [
      formData.name,
      formData.address,
      formData.gst ? `GST: ${formData.gst}` : null,
      formData.email ? `Email: ${formData.email}` : null,
      formData.kind_att ? `Kind Att: ${formData.kind_att}` : null,
      formData.contact ? `Contact: ${formData.contact}` : null,
      `Technical Pages (3–5): ${formData.include_full_proposal ? 'Yes' : 'No'}`,
      ...contactLines,
      ...customLines,
    ]
      .filter(Boolean)
      .join('\n')

    copy(text, 'Client details copied to clipboard!')
  }

  const isModalOpen = isCreatingNew || editingClient !== null
  const startItemNum = filteredClients.length === 0 ? 0 : (safePage - 1) * CLIENTS_PER_PAGE + 1
  const endItemNum = Math.min(safePage * CLIENTS_PER_PAGE, filteredClients.length)

  return (
    <div className="clients-page" style={{ padding: '16px 20px', maxWidth: '1320px', margin: '0 auto' }}>
      {/* Clean Minimal Header */}
      <div
        style={{
          display: 'flex',
          justifyContent: 'space-between',
          alignItems: 'center',
          flexWrap: 'wrap',
          gap: '12px',
          marginBottom: '14px',
        }}
      >
        <div>
          <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
            <h2 style={{ margin: 0, fontSize: '18px', fontWeight: 800, color: '#0f172a', letterSpacing: '-0.01em' }}>
              Client Directory
            </h2>
            <span
              style={{
                background: '#f1f5f9',
                color: '#334155',
                padding: '2px 8px',
                borderRadius: '999px',
                fontSize: '11.5px',
                fontWeight: 700,
                border: '1px solid #e2e8f0',
              }}
            >
              {userVisibleClients.length}
            </span>
          </div>
          <p style={{ margin: '2px 0 0 0', fontSize: '12.5px', color: '#64748b' }}>
            Manage buyer profiles, contacts, and default PDF technical pages preference.
          </p>
        </div>

        <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
          {onNavigate && (
            <button
              type="button"
              className="btn ghost sm"
              onClick={() => onNavigate('client_relations')}
              style={{ display: 'inline-flex', alignItems: 'center', gap: '6px', fontWeight: 600 }}
              title="View Client and Staff Relations Matrix"
            >
              {Icon.users} Relations
            </button>
          )}
          <button
            type="button"
            className="btn primary sm"
            onClick={handleStartCreateNew}
            style={{
              background: '#0f172a',
              borderColor: '#0f172a',
              color: '#ffffff',
              display: 'inline-flex',
              alignItems: 'center',
              gap: '6px',
              fontWeight: 700,
              padding: '7px 14px',
            }}
          >
            {Icon.plus} New Client
          </button>
        </div>
      </div>

      {/* Minimal Search + Filter Toolbar */}
      <div
        className="card"
        style={{
          padding: '10px 14px',
          marginBottom: '16px',
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'space-between',
          flexWrap: 'wrap',
          gap: '10px',
          border: '1px solid #e2e8f0',
        }}
      >
        <div style={{ position: 'relative', flex: '1 1 280px', maxWidth: '420px' }}>
          <input
            type="text"
            placeholder="Search by company, GST, contact, phone, email…"
            value={searchQuery}
            onChange={(e) => setSearchQuery(e.target.value)}
            style={{
              width: '100%',
              padding: '7px 28px 7px 30px',
              fontSize: '12.5px',
              borderRadius: '6px',
              border: '1px solid #cbd5e1',
              outline: 'none',
              background: '#f8fafc',
              color: '#0f172a',
            }}
          />
          <span
            style={{
              position: 'absolute',
              left: '9px',
              top: '50%',
              transform: 'translateY(-50%)',
              fontSize: '12px',
              color: '#94a3b8',
              pointerEvents: 'none',
            }}
          >
            {Icon.search}
          </span>
          {searchQuery && (
            <button
              type="button"
              onClick={() => setSearchQuery('')}
              style={{
                position: 'absolute',
                right: '8px',
                top: '50%',
                transform: 'translateY(-50%)',
                background: 'none',
                border: 'none',
                fontSize: '12px',
                color: '#94a3b8',
                cursor: 'pointer',
              }}
              title="Clear search"
            >
              ✕
            </button>
          )}
        </div>

        <div style={{ display: 'flex', alignItems: 'center', gap: '6px', flexWrap: 'wrap' }}>
          {(
            [
              { id: 'all', label: `All (${userVisibleClients.length})` },
              { id: 'tech_yes', label: `Tech Pages: Yes (${techYesCount})` },
              { id: 'tech_no', label: `Tech Pages: No (${userVisibleClients.length - techYesCount})` },
              ...(isOwnerOrAdmin ? [{ id: 'shared' as const, label: 'Shared' }] : []),
            ] as const
          ).map((f) => {
            const active = techFilter === f.id
            return (
              <button
                key={f.id}
                type="button"
                onClick={() => setTechFilter(f.id)}
                style={{
                  padding: '5px 10px',
                  borderRadius: '999px',
                  fontSize: '11.5px',
                  fontWeight: 600,
                  cursor: 'pointer',
                  border: active ? '1px solid #0f172a' : '1px solid #e2e8f0',
                  background: active ? '#0f172a' : '#ffffff',
                  color: active ? '#ffffff' : '#475569',
                  transition: 'all 0.12s ease',
                }}
              >
                {f.label}
              </button>
            )
          })}
        </div>
      </div>

      {/* Empty State or Clean 3x3 Client Grid */}
      {paginatedClients.length === 0 ? (
        <div
          className="card"
          style={{
            textAlign: 'center',
            padding: '44px 16px',
            color: '#64748b',
            border: '1px dashed #cbd5e1',
          }}
        >
          <p style={{ margin: '0 0 6px 0', fontWeight: 700, color: '#0f172a', fontSize: '14.5px' }}>
            No client records found
          </p>
          <span className="tiny" style={{ display: 'block', marginBottom: '14px' }}>
            {userVisibleClients.length === 0
              ? 'You have not created any clients yet. Add your first client to prefill quotations.'
              : 'No clients match your current search or filter.'}
          </span>
          <button
            type="button"
            className="btn ghost sm"
            onClick={userVisibleClients.length === 0 ? handleStartCreateNew : () => { setSearchQuery(''); setTechFilter('all') }}
            style={{ border: '1px solid #0f172a', color: '#0f172a', fontWeight: 600 }}
          >
            {userVisibleClients.length === 0 ? '+ New Client' : 'Reset Filters'}
          </button>
        </div>
      ) : (
        <>
          <div className="client-cards-grid">
            {paginatedClients.map((client) => {
              const addressFirst = client.address ? client.address.split('\n')[0] : 'No address provided'
              const extraContactsCount = (client.contacts || []).filter(
                (cp) => !(cp.name === client.kind_att && cp.phone === client.contact),
              ).length
              const customRowsCount = (client.extra_fields || []).length
              const techEnabled = client.include_full_proposal === true

              return (
                <div
                  key={client.id}
                  onClick={() => handleOpenClient(client)}
                  style={{
                    background: '#ffffff',
                    border: '1px solid #e2e8f0',
                    borderRadius: '10px',
                    padding: '14px 16px',
                    cursor: 'pointer',
                    transition: 'border-color 0.15s ease, box-shadow 0.15s ease, transform 0.15s ease',
                    display: 'flex',
                    flexDirection: 'column',
                    gap: '8px',
                  }}
                  onMouseEnter={(e) => {
                    e.currentTarget.style.borderColor = '#94a3b8'
                    e.currentTarget.style.boxShadow = '0 4px 14px rgba(15, 23, 42, 0.06)'
                  }}
                  onMouseLeave={(e) => {
                    e.currentTarget.style.borderColor = '#e2e8f0'
                    e.currentTarget.style.boxShadow = 'none'
                  }}
                >
                  {/* Top Row: Monogram + Name + Compact Badges */}
                  <div style={{ display: 'flex', alignItems: 'flex-start', justifyContent: 'space-between', gap: '10px' }}>
                    <div style={{ display: 'flex', alignItems: 'center', gap: '10px', minWidth: 0 }}>
                      <div
                        style={{
                          width: '34px',
                          height: '34px',
                          borderRadius: '8px',
                          background: '#f1f5f9',
                          border: '1px solid #e2e8f0',
                          color: '#0f172a',
                          fontSize: '12px',
                          fontWeight: 800,
                          display: 'grid',
                          placeItems: 'center',
                          flexShrink: 0,
                        }}
                      >
                        {getInitials(client.name)}
                      </div>
                      <div style={{ minWidth: 0 }}>
                        <div
                          style={{
                            fontSize: '14px',
                            color: '#0f172a',
                            fontWeight: 700,
                            lineHeight: 1.25,
                            overflow: 'hidden',
                            textOverflow: 'ellipsis',
                            whiteSpace: 'nowrap',
                          }}
                          title={client.name}
                        >
                          {client.name}
                        </div>
                        {client.gst ? (
                          <div style={{ fontSize: '10.5px', fontFamily: 'monospace', color: '#64748b', marginTop: '1px' }}>
                            GST: {client.gst}
                          </div>
                        ) : (
                          <div style={{ fontSize: '10.5px', color: '#94a3b8', marginTop: '1px' }}>No GSTIN</div>
                        )}
                      </div>
                    </div>

                    <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'flex-end', gap: '4px', flexShrink: 0 }}>
                      <span
                        title="Default Technical Pages (Pages 3–5) in PDF"
                        style={{
                          fontSize: '10px',
                          fontWeight: 700,
                          padding: '2px 7px',
                          borderRadius: '999px',
                          background: techEnabled ? '#eff6ff' : '#f8fafc',
                          color: techEnabled ? '#1d4ed8' : '#64748b',
                          border: techEnabled ? '1px solid #bfdbfe' : '1px solid #e2e8f0',
                          whiteSpace: 'nowrap',
                        }}
                      >
                        Tech Pages: {techEnabled ? 'Yes' : 'No'}
                      </span>
                      {isOwnerOrAdmin && (
                        <span
                          title="Client Profit Margin % (hidden from staff)"
                          style={{
                            fontSize: '10px',
                            fontWeight: 700,
                            padding: '1px 7px',
                            borderRadius: '999px',
                            background: '#f0fdf4',
                            color: '#166534',
                            border: '1px solid #bbf7d0',
                            whiteSpace: 'nowrap',
                          }}
                        >
                          {client.profit_margin_pct !== undefined ? `${client.profit_margin_pct}% Margin` : '25% Margin'}
                        </span>
                      )}
                    </div>
                  </div>

                  {/* Address & Contact Summary */}
                  <div
                    style={{
                      fontSize: '12px',
                      color: '#64748b',
                      whiteSpace: 'nowrap',
                      overflow: 'hidden',
                      textOverflow: 'ellipsis',
                    }}
                    title={client.address || ''}
                  >
                    {addressFirst}
                  </div>

                  <div
                    style={{
                      borderTop: '1px solid #f1f5f9',
                      paddingTop: '8px',
                      display: 'flex',
                      flexDirection: 'column',
                      gap: '2px',
                    }}
                  >
                    <div style={{ fontSize: '12px', color: '#0f172a', fontWeight: 600 }}>
                      {client.kind_att || client.contact ? (
                        <>
                          {client.kind_att || 'Primary Contact'}
                          {client.contact ? (
                            <span style={{ color: '#475569', fontWeight: 500 }}> · {client.contact}</span>
                          ) : null}
                        </>
                      ) : (
                        <span style={{ color: '#94a3b8', fontWeight: 400, fontSize: '11.5px' }}>
                          No primary contact
                        </span>
                      )}
                    </div>
                    {client.email && (
                      <div
                        style={{
                          fontSize: '11px',
                          color: '#64748b',
                          whiteSpace: 'nowrap',
                          overflow: 'hidden',
                          textOverflow: 'ellipsis',
                        }}
                      >
                        {client.email}
                      </div>
                    )}
                  </div>

                  {/* Footer Meta & Quick Quote */}
                  <div
                    style={{
                      display: 'flex',
                      justifyContent: 'space-between',
                      alignItems: 'center',
                      marginTop: 'auto',
                      paddingTop: '6px',
                      borderTop: '1px solid #f8fafc',
                      gap: '6px',
                      flexWrap: 'wrap',
                    }}
                  >
                    <div style={{ display: 'flex', alignItems: 'center', gap: '6px', flexWrap: 'wrap' }}>
                      <span style={{ fontSize: '11px', color: '#64748b' }}>
                        {extraContactsCount > 0 || customRowsCount > 0
                          ? [
                              extraContactsCount > 0
                                ? `+${extraContactsCount} contact${extraContactsCount > 1 ? 's' : ''}`
                                : null,
                              customRowsCount > 0
                                ? `${customRowsCount} row${customRowsCount > 1 ? 's' : ''}`
                                : null,
                            ]
                              .filter(Boolean)
                              .join(' · ')
                          : 'Standard'}
                      </span>
                      {(isOwnerOrAdmin || (client.shared_with && client.shared_with.length > 0)) && (
                        <span
                          style={{
                            fontSize: '10px',
                            color:
                              client.shared_with?.includes('all') ||
                              (client.shared_with && client.shared_with.length > 0)
                                ? '#1d4ed8'
                                : '#64748b',
                            background: '#f8fafc',
                            padding: '1px 6px',
                            borderRadius: '4px',
                            border: '1px solid #e2e8f0',
                          }}
                        >
                          {client.shared_with?.includes('all')
                            ? 'All Staff'
                            : client.shared_with && client.shared_with.length > 0
                            ? `Shared (${countUniqueSharedStaff(client.shared_with)} Staff)`
                            : 'Private'}
                        </span>
                      )}
                    </div>

                    {onCreateQuotationWithClient && (
                      <button
                        type="button"
                        className="btn ghost xs"
                        onClick={(e) => {
                          e.stopPropagation()
                          onCreateQuotationWithClient(client)
                        }}
                        style={{
                          fontSize: '11px',
                          fontWeight: 600,
                          color: '#0f172a',
                          border: '1px solid #cbd5e1',
                          padding: '2px 8px',
                        }}
                        title={`Create new quotation for ${client.name}`}
                      >
                        + Quote
                      </button>
                    )}
                  </div>
                </div>
              )
            })}
          </div>

          {/* Clean Numbered Pagination Bar */}
          <div
            style={{
              display: 'flex',
              justifyContent: 'space-between',
              alignItems: 'center',
              flexWrap: 'wrap',
              gap: '10px',
              marginTop: '16px',
              padding: '10px 14px',
              background: '#ffffff',
              border: '1px solid #e2e8f0',
              borderRadius: '8px',
            }}
          >
            <span style={{ fontSize: '12px', color: '#64748b' }}>
              Showing <strong style={{ color: '#0f172a' }}>{startItemNum}–{endItemNum}</strong> of{' '}
              <strong style={{ color: '#0f172a' }}>{filteredClients.length}</strong> client
              {filteredClients.length === 1 ? '' : 's'}
            </span>

            {totalPages > 1 && (
              <div style={{ display: 'flex', alignItems: 'center', gap: '4px', flexWrap: 'wrap' }}>
                <button
                  type="button"
                  className="btn ghost xs"
                  disabled={safePage <= 1}
                  onClick={() => setCurrentPage((p) => Math.max(1, p - 1))}
                  style={{ fontSize: '11.5px', padding: '4px 9px' }}
                >
                  ‹ Prev
                </button>
                {getPageNumbers(safePage, totalPages).map((pg, idx) =>
                  typeof pg === 'number' ? (
                    <button
                      key={`pg-${pg}`}
                      type="button"
                      onClick={() => setCurrentPage(pg)}
                      style={{
                        minWidth: '28px',
                        height: '28px',
                        padding: '0 6px',
                        borderRadius: '6px',
                        fontSize: '12px',
                        fontWeight: safePage === pg ? 700 : 600,
                        cursor: 'pointer',
                        border: safePage === pg ? '1px solid #0f172a' : '1px solid #e2e8f0',
                        background: safePage === pg ? '#0f172a' : '#ffffff',
                        color: safePage === pg ? '#ffffff' : '#334155',
                      }}
                    >
                      {pg}
                    </button>
                  ) : (
                    <span key={`ell-${idx}`} style={{ padding: '0 4px', color: '#94a3b8', fontSize: '12px' }}>
                      …
                    </span>
                  ),
                )}
                <button
                  type="button"
                  className="btn ghost xs"
                  disabled={safePage >= totalPages}
                  onClick={() => setCurrentPage((p) => Math.min(totalPages, p + 1))}
                  style={{ fontSize: '11.5px', padding: '4px 9px' }}
                >
                  Next ›
                </button>
              </div>
            )}
          </div>
        </>
      )}

      {/* Client Form Modal (zIndex cleanly below .overlay ConfirmDialog at 200000) */}
      {isModalOpen && (
        <div
          style={{
            position: 'fixed',
            inset: 0,
            background: 'rgba(15, 23, 42, 0.55)',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            zIndex: 10000,
            padding: '16px',
            backdropFilter: 'blur(2px)',
          }}
          onClick={handleCloseModal}
        >
          <div
            style={{
              background: '#ffffff',
              borderRadius: '10px',
              border: '1px solid #cbd5e1',
              padding: '20px 22px',
              width: '640px',
              maxWidth: '96vw',
              maxHeight: '92vh',
              overflowY: 'auto',
              boxShadow: '0 20px 25px -5px rgba(0, 0, 0, 0.2)',
              display: 'flex',
              flexDirection: 'column',
              gap: '12px',
            }}
            onClick={(e) => e.stopPropagation()}
          >
            {/* Header */}
            <div
              style={{
                display: 'flex',
                justifyContent: 'space-between',
                alignItems: 'center',
                borderBottom: '1px solid #e2e8f0',
                paddingBottom: '10px',
              }}
            >
              <h3 style={{ margin: 0, fontSize: '16px', fontWeight: 800, color: '#0f172a' }}>
                {isCreatingNew ? 'New Client' : `Edit Client: ${formData.name || ''}`}
              </h3>
              <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                {!isCreatingNew && (
                  <button
                    type="button"
                    className="btn ghost xs"
                    onClick={handleCopyDetails}
                    title="Copy details"
                    style={{ fontSize: '11px', padding: '4px 8px' }}
                  >
                    {Icon.copy} Copy Info
                  </button>
                )}
                <button
                  type="button"
                  className="icon-btn sm"
                  onClick={handleCloseModal}
                  style={{ fontSize: '16px', lineHeight: 1 }}
                >
                  ✕
                </button>
              </div>
            </div>

            {/* Company Name & Address */}
            <div>
              <label className="tiny font-bold muted" style={{ display: 'block', marginBottom: '3px' }}>
                Buyer / Company Name *
              </label>
              <input
                type="text"
                placeholder="e.g. Acme Corporation"
                className="option-input-field font-bold"
                value={formData.name}
                onChange={(e) => setFormData({ ...formData, name: e.target.value })}
                style={{ height: '35px', padding: '0 10px', width: '100%' }}
                autoFocus={isCreatingNew}
              />
            </div>

            <div>
              <label className="tiny font-bold muted" style={{ display: 'block', marginBottom: '3px' }}>
                Site / Billing Address
              </label>
              <textarea
                rows={2}
                placeholder="Address..."
                className="option-input-field"
                value={formData.address}
                onChange={(e) => setFormData({ ...formData, address: e.target.value })}
                style={{ fontSize: '12px', padding: '6px 10px', lineHeight: '1.4', width: '100%' }}
              />
            </div>

            {/* GST & Email in clean 2-col row */}
            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '10px' }}>
              <div>
                <label className="tiny font-bold muted" style={{ display: 'block', marginBottom: '3px' }}>
                  GST Number
                </label>
                <input
                  type="text"
                  placeholder="e.g. 36AABCP1234F1Z5"
                  className="option-input-field font-mono"
                  value={formData.gst}
                  onChange={(e) => setFormData({ ...formData, gst: e.target.value })}
                  style={{ height: '34px', padding: '0 10px', width: '100%' }}
                />
              </div>
              <div>
                <label className="tiny font-bold muted" style={{ display: 'block', marginBottom: '3px' }}>
                  Buyer Email
                </label>
                <input
                  type="email"
                  placeholder="e.g. stores@company.com"
                  className="option-input-field"
                  value={formData.email}
                  onChange={(e) => setFormData({ ...formData, email: e.target.value })}
                  style={{ height: '34px', padding: '0 10px', width: '100%' }}
                />
              </div>
            </div>

            {/* Default PDF Technical Pages (Pages 3-5) Yes / No */}
            <div
              style={{
                background: '#f8fafc',
                border: '1px solid #cbd5e1',
                borderRadius: '8px',
                padding: '10px 12px',
                display: 'flex',
                justifyContent: 'space-between',
                alignItems: 'center',
                gap: '12px',
                flexWrap: 'wrap',
              }}
            >
              <div>
                <div style={{ fontSize: '12.5px', fontWeight: 700, color: '#0f172a' }}>
                  Technical Pages in PDF (Pages 3–5)
                </div>
                <div style={{ fontSize: '11px', color: '#64748b', marginTop: '1px' }}>
                  If Yes, PDF shows technical specification pages (3–5) by default. If No, only Commercial &amp; Terms (1–2) show until toggled manually.
                </div>
              </div>
              <div style={{ display: 'inline-flex', background: '#e2e8f0', padding: '3px', borderRadius: '7px', gap: '3px' }}>
                <button
                  type="button"
                  onClick={() => setFormData({ ...formData, include_full_proposal: false })}
                  style={{
                    padding: '5px 14px',
                    borderRadius: '5px',
                    border: 'none',
                    fontSize: '12px',
                    fontWeight: 700,
                    cursor: 'pointer',
                    background: !formData.include_full_proposal ? '#ffffff' : 'transparent',
                    color: !formData.include_full_proposal ? '#0f172a' : '#475569',
                    boxShadow: !formData.include_full_proposal ? '0 1px 2px rgba(0,0,0,0.08)' : 'none',
                  }}
                >
                  No
                </button>
                <button
                  type="button"
                  onClick={() => setFormData({ ...formData, include_full_proposal: true })}
                  style={{
                    padding: '5px 14px',
                    borderRadius: '5px',
                    border: 'none',
                    fontSize: '12px',
                    fontWeight: 700,
                    cursor: 'pointer',
                    background: formData.include_full_proposal ? '#1d4ed8' : 'transparent',
                    color: formData.include_full_proposal ? '#ffffff' : '#475569',
                    boxShadow: formData.include_full_proposal ? '0 1px 2px rgba(29,78,216,0.25)' : 'none',
                  }}
                >
                  Yes
                </button>
              </div>
            </div>

            {/* Primary Contact */}
            <div>
              <label className="tiny font-bold muted" style={{ display: 'block', marginBottom: '3px' }}>
                Primary Contact (Name &amp; Mobile Number)
              </label>
              <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '10px' }}>
                <input
                  type="text"
                  placeholder="Person Name (e.g. Mr. Rajesh Sharma)"
                  className="option-input-field font-bold"
                  value={formData.kind_att}
                  onChange={(e) => setFormData({ ...formData, kind_att: e.target.value })}
                  style={{ height: '34px', padding: '0 10px', width: '100%' }}
                />
                <input
                  type="text"
                  placeholder="Mobile Number (e.g. +91 9876543210)"
                  className="option-input-field"
                  value={formData.contact}
                  onChange={(e) => setFormData({ ...formData, contact: e.target.value })}
                  style={{ height: '34px', padding: '0 10px', width: '100%' }}
                />
              </div>
            </div>

            {/* Additional Contacts */}
            {formData.contacts.map((cp, idx) => (
              <div
                key={cp.id}
                style={{
                  display: 'grid',
                  gridTemplateColumns: '1fr 1fr auto auto',
                  gap: '8px',
                  alignItems: 'flex-end',
                }}
              >
                <div>
                  <label className="tiny font-bold muted" style={{ display: 'block', marginBottom: '3px' }}>
                    Kind Attention (Name)
                  </label>
                  <input
                    type="text"
                    placeholder="e.g. A Dileep"
                    className="option-input-field font-bold"
                    value={cp.name}
                    onChange={(e) => handleUpdateContact(idx, 'name', e.target.value)}
                    style={{ height: '34px', padding: '0 8px', width: '100%' }}
                  />
                </div>
                <div>
                  <label className="tiny font-bold muted" style={{ display: 'block', marginBottom: '3px' }}>
                    Contact Number
                  </label>
                  <input
                    type="text"
                    placeholder="e.g. 98492 61819"
                    className="option-input-field"
                    value={cp.phone}
                    onChange={(e) => handleUpdateContact(idx, 'phone', e.target.value)}
                    style={{ height: '34px', padding: '0 8px', width: '100%' }}
                  />
                </div>
                <label
                  style={{
                    display: 'flex',
                    alignItems: 'center',
                    gap: '4px',
                    fontSize: '11px',
                    cursor: 'pointer',
                    whiteSpace: 'nowrap',
                    height: '34px',
                    color: '#475569',
                    fontWeight: 600,
                  }}
                >
                  <input
                    type="checkbox"
                    checked={cp.visible !== false}
                    onChange={(e) => {
                      const next = formData.contacts.map((c, i) =>
                        i === idx ? { ...c, visible: e.target.checked } : c,
                      )
                      setFormData({ ...formData, contacts: next })
                    }}
                  />
                  <span>Show on PDF</span>
                </label>
                <button
                  type="button"
                  className="btn ghost xs"
                  onClick={() => handleRemoveContact(idx)}
                  style={{ height: '34px', padding: '0 10px', color: '#dc2626', border: '1px solid #cbd5e1' }}
                  title="Remove contact"
                >
                  ✕
                </button>
              </div>
            ))}

            {/* Custom Rows */}
            {formData.extra_fields.map((ef, idx) => (
              <div
                key={ef.id}
                style={{
                  display: 'grid',
                  gridTemplateColumns: '1fr 1fr auto auto',
                  gap: '8px',
                  alignItems: 'flex-end',
                }}
              >
                <div>
                  <label className="tiny font-bold muted" style={{ display: 'block', marginBottom: '3px' }}>
                    Row Label
                  </label>
                  <input
                    type="text"
                    placeholder="e.g. PO Ref"
                    className="option-input-field"
                    value={ef.label}
                    onChange={(e) => handleUpdateCustomRow(idx, 'label', e.target.value)}
                    style={{ height: '34px', padding: '0 8px', width: '100%' }}
                  />
                </div>
                <div>
                  <label className="tiny font-bold muted" style={{ display: 'block', marginBottom: '3px' }}>
                    Row Value
                  </label>
                  <input
                    type="text"
                    placeholder="e.g. Value"
                    className="option-input-field"
                    value={ef.value}
                    onChange={(e) => handleUpdateCustomRow(idx, 'value', e.target.value)}
                    style={{ height: '34px', padding: '0 8px', width: '100%' }}
                  />
                </div>
                <label
                  style={{
                    display: 'flex',
                    alignItems: 'center',
                    gap: '4px',
                    fontSize: '11px',
                    cursor: 'pointer',
                    whiteSpace: 'nowrap',
                    height: '34px',
                    color: '#475569',
                    fontWeight: 600,
                  }}
                >
                  <input
                    type="checkbox"
                    checked={ef.visible !== false}
                    onChange={(e) => handleUpdateCustomRow(idx, 'visible', e.target.checked)}
                  />
                  <span>Show on PDF</span>
                </label>
                <button
                  type="button"
                  className="btn ghost xs"
                  onClick={() => handleRemoveCustomRow(idx)}
                  style={{ height: '34px', padding: '0 10px', color: '#dc2626', border: '1px solid #cbd5e1' }}
                  title="Remove custom row"
                >
                  ✕
                </button>
              </div>
            ))}

            {/* Buttons for add contact & custom rows */}
            <div style={{ display: 'flex', gap: '8px' }}>
              <button
                type="button"
                className="btn ghost sm"
                onClick={handleAddContact}
                style={{ border: '1px solid #cbd5e1', color: '#0f172a', fontWeight: 600, fontSize: '12px' }}
              >
                + Add Contact
              </button>
              <button
                type="button"
                className="btn ghost sm"
                onClick={handleAddCustomRow}
                style={{ border: '1px solid #cbd5e1', color: '#0f172a', fontWeight: 600, fontSize: '12px' }}
              >
                + Add Custom Row
              </button>
            </div>

            {/* Owner & Admin Confidential Profit Controls */}
            {isOwnerOrAdmin && (
              <div style={{ display: 'flex', flexDirection: 'column', gap: '10px', marginTop: '4px' }}>
                <div style={{ background: '#f0fdf4', border: '1px solid #bbf7d0', borderRadius: '8px', padding: '10px 14px' }}>
                  <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', gap: '10px' }}>
                    <div>
                      <label className="tiny font-bold" style={{ color: '#166534', display: 'block', margin: 0, fontSize: '12px' }}>
                        Client Profit Margin (%)
                      </label>
                      <span className="tiny" style={{ color: '#15803d', display: 'block', marginTop: '2px', fontSize: '11px' }}>
                        Default 25%. Future quotations for this client use this value. Strictly hidden from staff.
                      </span>
                    </div>
                    <div style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                      <input
                        type="number"
                        min={0}
                        max={200}
                        step={1}
                        className="option-input-field font-mono font-bold"
                        value={formData.profit_margin_pct}
                        onChange={(e) =>
                          setFormData({
                            ...formData,
                            profit_margin_pct: Math.max(0, parseFloat(e.target.value) || 0),
                          })
                        }
                        style={{
                          width: '72px',
                          height: '34px',
                          textAlign: 'right',
                          color: '#166534',
                          background: '#ffffff',
                          borderColor: '#86efac',
                          fontSize: '14px',
                        }}
                      />
                      <span style={{ fontWeight: 700, color: '#166534', fontSize: '13px' }}>%</span>
                    </div>
                  </div>
                </div>
              </div>
            )}

            {/* Share Client & Auto-Share Quotations with Staff */}
            {canManageClientShare && (
              <div style={{ background: '#f8fafc', border: '1px solid #e2e8f0', borderRadius: '8px', padding: '10px 14px', marginTop: '4px' }}>
                <label className="tiny font-bold" style={{ color: '#0f172a', display: 'block', marginBottom: '2px', fontSize: '12px' }}>
                  Share Client &amp; Quotations with Staff
                </label>
                <span className="tiny muted" style={{ display: 'block', marginBottom: '8px', fontSize: '11px' }}>
                  Selected staff members can view this client and will automatically receive any quotation created for this client.
                </span>

                <div style={{ display: 'flex', flexDirection: 'column', gap: '8px' }}>
                  <label style={{ display: 'flex', alignItems: 'center', gap: '8px', cursor: 'pointer', fontSize: '12px', fontWeight: 600, color: '#1e293b' }}>
                    <input
                      type="checkbox"
                      checked={formData.shared_with.includes('all')}
                      onChange={(e) => {
                        setFormData({
                          ...formData,
                          shared_with: e.target.checked ? ['all'] : [],
                        })
                      }}
                    />
                    <span>Share with All Staff Members</span>
                  </label>

                  {!formData.shared_with.includes('all') && (
                    <div>
                      {orgStaffMembers.length > 0 ? (
                        <div style={{ display: 'flex', flexWrap: 'wrap', gap: '6px' }}>
                          {orgStaffMembers.map((m) => {
                            const isSelected =
                              formData.shared_with.includes(m.username) ||
                              Boolean(m.user_id && formData.shared_with.includes(m.user_id))
                            return (
                              <button
                                key={m.username}
                                type="button"
                                onClick={() => {
                                  const next = isSelected
                                    ? formData.shared_with.filter(
                                        (u) => u !== m.username && (!m.user_id || u !== m.user_id),
                                      )
                                    : [
                                        ...formData.shared_with,
                                        m.username,
                                        ...(m.user_id ? [m.user_id] : []),
                                      ]
                                  setFormData({ ...formData, shared_with: next })
                                }}
                                style={{
                                  padding: '4px 10px',
                                  borderRadius: '16px',
                                  fontSize: '11px',
                                  fontWeight: 600,
                                  cursor: 'pointer',
                                  border: isSelected ? '1px solid #2563eb' : '1px solid #cbd5e1',
                                  background: isSelected ? '#eff6ff' : '#ffffff',
                                  color: isSelected ? '#1d4ed8' : '#475569',
                                  display: 'inline-flex',
                                  alignItems: 'center',
                                  gap: '4px',
                                }}
                              >
                                <span>{isSelected ? '✓' : '+'}</span>
                                <span>@{m.username}</span>
                                <span style={{ fontSize: '9px', opacity: 0.7 }}>({m.role})</span>
                              </button>
                            )
                          })}
                        </div>
                      ) : (
                        <div style={{ display: 'flex', gap: '8px', alignItems: 'center' }}>
                          <input
                            type="text"
                            placeholder="Enter staff username to share (e.g. staff1)"
                            className="option-input-field"
                            value={customShareStaffInput}
                            onChange={(e) => setCustomShareStaffInput(e.target.value)}
                            style={{ height: '32px', fontSize: '12px', flex: 1 }}
                          />
                          <button
                            type="button"
                            className="btn ghost sm"
                            onClick={() => {
                              const val = customShareStaffInput.trim()
                              if (val && !formData.shared_with.includes(val)) {
                                setFormData({ ...formData, shared_with: [...formData.shared_with, val] })
                                setCustomShareStaffInput('')
                              }
                            }}
                            style={{ height: '32px', fontSize: '11px' }}
                          >
                            + Add Staff
                          </button>
                        </div>
                      )}
                    </div>
                  )}
                </div>
              </div>
            )}

            {/* Modal Bottom Actions */}
            <div
              style={{
                display: 'flex',
                justifyContent: 'space-between',
                alignItems: 'center',
                borderTop: '1px solid #e2e8f0',
                paddingTop: '12px',
                marginTop: '4px',
                gap: '10px',
              }}
            >
              <div>
                {!isCreatingNew && editingClient && (
                  <button
                    type="button"
                    className="btn ghost xs"
                    onClick={handleDelete}
                    style={{ color: '#dc2626', border: '1px solid #fecaca', background: '#fef2f2' }}
                  >
                    {Icon.trash} Delete Client
                  </button>
                )}
              </div>

              <div style={{ display: 'flex', gap: '8px', alignItems: 'center' }}>
                {!isCreatingNew && editingClient && onCreateQuotationWithClient && (
                  <button
                    type="button"
                    className="btn ghost sm"
                    onClick={() => {
                      handleCloseModal()
                      onCreateQuotationWithClient(editingClient)
                    }}
                    style={{ border: '1px solid #0f172a', color: '#0f172a', fontWeight: 600 }}
                  >
                    Start Quotation
                  </button>
                )}

                <button type="button" className="btn ghost sm" onClick={handleCloseModal}>
                  Cancel
                </button>

                <button
                  type="button"
                  className="btn primary sm"
                  onClick={handleSave}
                  style={{
                    background: '#0f172a',
                    borderColor: '#0f172a',
                    color: '#ffffff',
                    fontWeight: 700,
                    padding: '6px 16px',
                  }}
                >
                  {isCreatingNew ? 'Save Client' : 'Save Changes'}
                </button>
              </div>
            </div>
          </div>
        </div>
      )}
    </div>
  )
}
