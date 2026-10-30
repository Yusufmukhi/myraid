import React, { useEffect, useMemo, useState } from 'react'
import { Member, Org, Session, api } from '../../lib/api'
import {
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
}

export function ClientsPage({ session, org, onCreateQuotationWithClient }: Props) {
  const { toast, confirm } = useUI()
  const copy = useCopy()

  const isOwnerOrAdmin = org.role === 'owner' || org.role === 'admin' || session.username === 'owner' || session.username === 'admin'

  const [clients, setClients] = useState<SavedClient[]>(() => getSavedClients(org.id))
  const [searchQuery, setSearchQuery] = useState('')
  const [editingClient, setEditingClient] = useState<SavedClient | null>(null)
  const [isCreatingNew, setIsCreatingNew] = useState(false)
  const [orgStaffMembers, setOrgStaffMembers] = useState<Member[]>([])
  const [customShareStaffInput, setCustomShareStaffInput] = useState('')

  // Load org members for Owner/Admin so they can pick staff to share with
  useEffect(() => {
    if (isOwnerOrAdmin && session.token && org.id) {
      api.orgMembers(session.token, org.id)
        .then((members) => {
          if (Array.isArray(members)) {
            const staff = members.filter((m) => m.username !== session.username)
            setOrgStaffMembers(staff)
          }
        })
        .catch(() => {})
    }
  }, [isOwnerOrAdmin, session.token, org.id, session.username])

  // Simple flat form state: label up, value down
  const [formData, setFormData] = useState<{
    name: string
    address: string
    gst: string
    email: string
    kind_att: string
    contact: string
    contacts: ClientContactPerson[]
    extra_fields: QuotationExtraField[]
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
    profit_margin_pct: 25,
    shared_with: [],
  })

  // Synchronize client list across tabs/windows and fetch from Supabase
  useEffect(() => {
    fetchClientsFromSupabase(session.token, org.id)
    const handleUpdate = () => {
      const updated = getSavedClients(org.id)
      setClients(updated)
    }
    window.addEventListener('examcalc:clients-updated', handleUpdate)
    return () => window.removeEventListener('examcalc:clients-updated', handleUpdate)
  }, [session.token, org.id])

  // Filter clients based on user role & ownership:
  // - Owner and Admins see ALL clients
  // - Staff only sees clients created by themselves
  const userVisibleClients = useMemo(() => {
    return getVisibleClientsForUser(clients, session.username, org.role, session.user_id)
  }, [clients, session.username, org.role, session.user_id])

  // Search filter across visible clients
  const filteredClients = useMemo(() => {
    const q = searchQuery.trim().toLowerCase()
    if (!q) return userVisibleClients || []
    return (userVisibleClients || []).filter(
      (c) =>
        (c.name || '').toLowerCase().includes(q) ||
        (c.gst || '').toLowerCase().includes(q) ||
        (c.kind_att || '').toLowerCase().includes(q) ||
        (c.address || '').toLowerCase().includes(q) ||
        (c.contact || '').toLowerCase().includes(q) ||
        (c.email || '').toLowerCase().includes(q) ||
        (c.contacts || []).some((cp) => cp.name.toLowerCase().includes(q) || cp.phone.toLowerCase().includes(q)) ||
        (c.extra_fields || []).some((ef) => (ef.label || '').toLowerCase().includes(q) || (ef.value || '').toLowerCase().includes(q)),
    )
  }, [userVisibleClients, searchQuery])

  // Open existing client
  const handleOpenClient = (client: SavedClient) => {
    setIsCreatingNew(false)
    setEditingClient(client)

    // Filter out contacts duplicating primary contact
    const rawContacts = Array.isArray(client.contacts) ? client.contacts : []
    const cleanContacts = rawContacts.filter(
      (cp) => !(cp.name === client.kind_att && cp.phone === client.contact && cp.name.trim().length > 0)
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
          id: ef.id || ((typeof crypto !== 'undefined' && crypto.randomUUID) ? `ef_${crypto.randomUUID()}` : `ef_${Date.now()}`),
          label: ef.label || '',
          value: val,
          values: [val],
          visible: ef.visible !== false,
          visible_values: [true],
        }
      }),
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
      id: (typeof crypto !== 'undefined' && crypto.randomUUID) ? `cp_${crypto.randomUUID()}` : `cp_${Date.now()}`,
      name: '',
      phone: '',
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
      id: (typeof crypto !== 'undefined' && crypto.randomUUID) ? `ef_${crypto.randomUUID()}` : `ef_${Date.now()}`,
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
          visible: true,
          visible_values: [true],
        })),
      profit_margin_pct: isOwnerOrAdmin
        ? (formData.profit_margin_pct !== undefined ? formData.profit_margin_pct : 25)
        : (editingClient?.profit_margin_pct !== undefined ? editingClient.profit_margin_pct : 25),
      shared_with: isOwnerOrAdmin
        ? (Array.isArray(formData.shared_with) ? formData.shared_with : [])
        : (editingClient?.shared_with || []),
    }

    if (isCreatingNew) {
      const saved = saveClientToDirectory(
        payload,
        org.id,
        session.username,
        org.role,
      )
      toast(`Client "${saved.name}" created successfully!`)
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
      toast(`Client "${saved.name}" updated successfully!`)
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
      ...contactLines,
      ...customLines,
    ]
      .filter(Boolean)
      .join('\n')

    copy(text, 'Client details copied to clipboard!')
  }

  const isModalOpen = isCreatingNew || editingClient !== null

  return (
    <div className="clients-page" style={{ padding: '20px 24px', maxWidth: '1400px', margin: '0 auto' }}>
      {/* Top Header Card */}
      <div
        className="card"
        style={{
          display: 'flex',
          justifyContent: 'space-between',
          alignItems: 'center',
          flexWrap: 'wrap',
          gap: '14px',
          padding: '16px 20px',
          marginBottom: '20px',
          borderLeft: '4px solid #0f172a',
        }}
      >
        <div>
          <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
            <span style={{ fontSize: '18px' }}>🏢</span>
            <h2 style={{ margin: 0, fontSize: '18px', fontWeight: 800, color: '#0f172a' }}>
              Client Directory
            </h2>
            <span
              style={{
                background: '#f1f5f9',
                color: '#0f172a',
                padding: '2px 8px',
                borderRadius: '6px',
                fontSize: '12px',
                fontWeight: 600,
                border: '1px solid #cbd5e1',
              }}
            >
              {userVisibleClients.length} {userVisibleClients.length === 1 ? 'Client' : 'Clients'}
            </span>
          </div>
          <p style={{ margin: '4px 0 0 0', fontSize: '13px', color: '#64748b' }}>
            Click on any client cell to view, edit, or launch quotations.
          </p>
        </div>

        <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
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
            }}
          >
            {Icon.plus} + Add New Client
          </button>
        </div>
      </div>

      {/* Real-time Search Filter Bar */}
      <div style={{ position: 'relative', marginBottom: '16px', maxWidth: '460px' }}>
        <input
          type="text"
          placeholder="Search client cells by company name, GST, contact person, or phone…"
          value={searchQuery}
          onChange={(e) => setSearchQuery(e.target.value)}
          style={{
            width: '100%',
            padding: '8px 12px 8px 32px',
            fontSize: '13px',
            borderRadius: '6px',
            border: '1px solid #cbd5e1',
            outline: 'none',
            background: '#ffffff',
          }}
        />
        <span
          style={{
            position: 'absolute',
            left: '10px',
            top: '50%',
            transform: 'translateY(-50%)',
            fontSize: '13px',
            color: '#94a3b8',
            pointerEvents: 'none',
          }}
        >
          🔍
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
          >
            ✕
          </button>
        )}
      </div>

      {/* Grid of Client Cells: 3*3 layout */}
      {filteredClients.length === 0 ? (
        <div
          className="card"
          style={{
            textAlign: 'center',
            padding: '48px 16px',
            color: '#64748b',
            border: '1px dashed #cbd5e1',
          }}
        >
          <div style={{ fontSize: '28px', marginBottom: '8px' }}>🏢</div>
          <p style={{ margin: '0 0 6px 0', fontWeight: 700, color: '#0f172a', fontSize: '15px' }}>
            No client records found
          </p>
          <span className="tiny" style={{ display: 'block', marginBottom: '14px' }}>
            {userVisibleClients.length === 0
              ? 'You have not created any clients yet. Click below to add your first client.'
              : 'No clients match your search query.'}
          </span>
          <button
            type="button"
            className="btn ghost sm"
            onClick={handleStartCreateNew}
            style={{ border: '1px solid #0f172a', color: '#0f172a' }}
          >
            + Add New Client
          </button>
        </div>
      ) : (
        <div
          className="client-cards-grid"
          style={{
            display: 'grid',
            gridTemplateColumns: 'repeat(3, minmax(0, 1fr))',
            gap: '16px',
          }}
        >
          {filteredClients.map((client) => {
            const addressFirst = client.address ? client.address.split('\n')[0] : 'No address provided'
            const extraContactsCount = (client.contacts || []).filter(
              (cp) => !(cp.name === client.kind_att && cp.phone === client.contact)
            ).length
            const customRowsCount = (client.extra_fields || []).length

            return (
              <div
                key={client.id}
                onClick={() => handleOpenClient(client)}
                style={{
                  background: '#ffffff',
                  border: '1px solid #cbd5e1',
                  borderRadius: '8px',
                  padding: '16px',
                  cursor: 'pointer',
                  transition: 'border-color 0.15s ease, box-shadow 0.15s ease',
                  display: 'flex',
                  flexDirection: 'column',
                  gap: '8px',
                }}
                onMouseEnter={(e) => {
                  e.currentTarget.style.borderColor = '#0f172a'
                  e.currentTarget.style.boxShadow = '0 2px 8px rgba(0,0,0,0.06)'
                }}
                onMouseLeave={(e) => {
                  e.currentTarget.style.borderColor = '#cbd5e1'
                  e.currentTarget.style.boxShadow = 'none'
                }}
              >
                {/* Cell Header: Name & GST & Margin (Margin is strictly Owner/Admin only) */}
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', gap: '8px' }}>
                  <strong style={{ fontSize: '15px', color: '#0f172a', fontWeight: 800, lineHeight: 1.2 }}>
                    {client.name}
                  </strong>
                  <div style={{ display: 'flex', alignItems: 'center', gap: '4px', flexWrap: 'wrap', justifyContent: 'flex-end' }}>
                    {isOwnerOrAdmin && (
                      <span
                        title="Client Profit Margin % (strictly hidden from staff)"
                        style={{
                          fontSize: '10px',
                          fontWeight: 700,
                          padding: '2px 6px',
                          background: '#f0fdf4',
                          color: '#166534',
                          borderRadius: '4px',
                          border: '1px solid #bbf7d0',
                          whiteSpace: 'nowrap',
                        }}
                      >
                        {client.profit_margin_pct !== undefined ? `${client.profit_margin_pct}% Margin` : '25% Margin'}
                      </span>
                    )}
                    {client.gst && (
                      <span
                        style={{
                          fontSize: '10px',
                          fontFamily: 'monospace',
                          fontWeight: 700,
                          padding: '2px 6px',
                          background: '#f1f5f9',
                          color: '#475569',
                          borderRadius: '4px',
                          border: '1px solid #cbd5e1',
                          whiteSpace: 'nowrap',
                        }}
                      >
                        {client.gst}
                      </span>
                    )}
                  </div>
                </div>

                {/* Address snippet */}
                <span
                  style={{
                    fontSize: '12px',
                    color: '#64748b',
                    whiteSpace: 'nowrap',
                    overflow: 'hidden',
                    textOverflow: 'ellipsis',
                  }}
                >
                  📍 {addressFirst}
                </span>

                {/* Company Email */}
                {client.email && (
                  <div style={{ fontSize: '11px', color: '#64748b', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>
                    ✉️ {client.email}
                  </div>
                )}

                {/* Primary Contact details */}
                <div style={{ borderTop: '1px solid #f1f5f9', paddingTop: '8px', marginTop: '4px', display: 'flex', flexDirection: 'column', gap: '3px' }}>
                  {(client.kind_att || client.contact) ? (
                    <div style={{ fontSize: '12px', color: '#0f172a', fontWeight: 600 }}>
                      👤 {client.kind_att || 'Contact'} {client.contact ? `· ${client.contact}` : ''}
                    </div>
                  ) : (
                    <div style={{ fontSize: '11px', color: '#94a3b8' }}>
                      No primary contact specified
                    </div>
                  )}
                </div>

                {/* Cell Footer: Extra items & Creator */}
                <div
                  style={{
                    display: 'flex',
                    justifyContent: 'space-between',
                    alignItems: 'center',
                    marginTop: 'auto',
                    paddingTop: '6px',
                    fontSize: '11px',
                    color: '#94a3b8',
                    flexWrap: 'wrap',
                    gap: '4px',
                  }}
                >
                  <span style={{ fontSize: '11px', color: '#64748b' }}>
                    {extraContactsCount > 0 && `+${extraContactsCount} contact${extraContactsCount > 1 ? 's' : ''}`}
                    {extraContactsCount > 0 && customRowsCount > 0 && ' · '}
                    {customRowsCount > 0 && `${customRowsCount} custom row${customRowsCount > 1 ? 's' : ''}`}
                    {extraContactsCount === 0 && customRowsCount === 0 && 'Standard profile'}
                  </span>
                  <div style={{ display: 'flex', alignItems: 'center', gap: '4px' }}>
                    {isOwnerOrAdmin && (
                      <span
                        title="Staff sharing status"
                        style={{
                          fontSize: '10px',
                          color: client.shared_with?.includes('all') || (client.shared_with && client.shared_with.length > 0) ? '#1d4ed8' : '#64748b',
                          background: client.shared_with?.includes('all') || (client.shared_with && client.shared_with.length > 0) ? '#eff6ff' : '#f8fafc',
                          padding: '1px 5px',
                          borderRadius: '3px',
                          border: client.shared_with?.includes('all') || (client.shared_with && client.shared_with.length > 0) ? '1px solid #bfdbfe' : '1px solid #e2e8f0',
                        }}
                      >
                        {client.shared_with?.includes('all')
                          ? '👥 All Staff'
                          : client.shared_with && client.shared_with.length > 0
                          ? `👥 ${client.shared_with.length} Staff`
                          : '🔒 Private'}
                      </span>
                    )}
                    {!isOwnerOrAdmin && client.created_by !== session.username && (
                      <span
                        style={{
                          fontSize: '10px',
                          color: '#1d4ed8',
                          background: '#eff6ff',
                          padding: '1px 5px',
                          borderRadius: '3px',
                          border: '1px solid #bfdbfe',
                        }}
                      >
                        Shared by Owner
                      </span>
                    )}
                    <span style={{ fontSize: '10px', color: '#475569', background: '#f8fafc', padding: '1px 5px', borderRadius: '3px', border: '1px solid #e2e8f0' }}>
                      {client.created_by === session.username ? 'Created by you' : (client.created_by || 'Directory')}
                    </span>
                  </div>
                </div>
              </div>
            )
          })}
        </div>
      )}

      {/* Simple Flat Client Form Modal (No widgets, label above, input below) */}
      {isModalOpen && (
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
          onClick={handleCloseModal}
        >
          <div
            style={{
              background: '#ffffff',
              borderRadius: '8px',
              border: '1px solid #cbd5e1',
              padding: '24px',
              width: '640px',
              maxWidth: '96vw',
              maxHeight: '92vh',
              overflowY: 'auto',
              boxShadow: '0 20px 25px -5px rgba(0, 0, 0, 0.2)',
              display: 'flex',
              flexDirection: 'column',
              gap: '14px',
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
                {isCreatingNew ? 'Add Client' : `Edit Client: ${formData.name || ''}`}
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

            {/* Simple Form Fields: label on top, value below */}
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
                style={{ height: '36px', padding: '0 10px', width: '100%' }}
                autoFocus={isCreatingNew}
              />
            </div>

            <div>
              <label className="tiny font-bold muted" style={{ display: 'block', marginBottom: '3px' }}>
                Site / Billing Address
              </label>
              <textarea
                rows={3}
                placeholder="Address..."
                className="option-input-field"
                value={formData.address}
                onChange={(e) => setFormData({ ...formData, address: e.target.value })}
                style={{ fontSize: '12px', padding: '6px 10px', lineHeight: '1.4', width: '100%' }}
              />
            </div>

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
                style={{ height: '36px', padding: '0 10px', width: '100%' }}
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
                style={{ height: '36px', padding: '0 10px', width: '100%' }}
              />
            </div>

            <div>
              <label className="tiny font-bold muted" style={{ display: 'block', marginBottom: '3px' }}>
                Primary Contact (Name &amp; Mobile Number)
              </label>
              <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '8px' }}>
                <input
                  type="text"
                  placeholder="Person Name (e.g. Mr. Rajesh Sharma)"
                  className="option-input-field font-bold"
                  value={formData.kind_att}
                  onChange={(e) => setFormData({ ...formData, kind_att: e.target.value })}
                  style={{ height: '36px', padding: '0 10px', width: '100%' }}
                />
                <input
                  type="text"
                  placeholder="Mobile Number (e.g. +91 9876543210)"
                  className="option-input-field"
                  value={formData.contact}
                  onChange={(e) => setFormData({ ...formData, contact: e.target.value })}
                  style={{ height: '36px', padding: '0 10px', width: '100%' }}
                />
              </div>
            </div>

            {/* Additional Contacts: simple inputs, label up, value down */}
            {formData.contacts.map((cp, idx) => (
              <div
                key={cp.id}
                style={{
                  display: 'grid',
                  gridTemplateColumns: '1fr 1fr auto',
                  gap: '8px',
                  alignItems: 'flex-end',
                }}
              >
                <div>
                  <label className="tiny font-bold muted" style={{ display: 'block', marginBottom: '3px' }}>
                    Contact Person Name
                  </label>
                  <input
                    type="text"
                    placeholder="e.g. xyz"
                    className="option-input-field"
                    value={cp.name}
                    onChange={(e) => handleUpdateContact(idx, 'name', e.target.value)}
                    style={{ height: '34px', padding: '0 8px', width: '100%' }}
                  />
                </div>
                <div>
                  <label className="tiny font-bold muted" style={{ display: 'block', marginBottom: '3px' }}>
                    Mobile Number
                  </label>
                  <input
                    type="text"
                    placeholder="e.g. 123"
                    className="option-input-field"
                    value={cp.phone}
                    onChange={(e) => handleUpdateContact(idx, 'phone', e.target.value)}
                    style={{ height: '34px', padding: '0 8px', width: '100%' }}
                  />
                </div>
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

            {/* Custom Rows: user can edit both label and value, and choose Share in PDF */}
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
                    placeholder="e.g. xyz"
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
                    placeholder="e.g. xyz"
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

            {/* Buttons for add contact and custom rows at the very last */}
            <div style={{ display: 'flex', gap: '8px', paddingTop: '4px' }}>
              <button
                type="button"
                className="btn ghost sm"
                onClick={handleAddContact}
                style={{ border: '1px solid #0f172a', color: '#0f172a', fontWeight: 600, fontSize: '12px' }}
              >
                + Add Contact
              </button>
              <button
                type="button"
                className="btn ghost sm"
                onClick={handleAddCustomRow}
                style={{ border: '1px solid #0f172a', color: '#0f172a', fontWeight: 600, fontSize: '12px' }}
              >
                + Add Custom Row
              </button>
            </div>

            {/* Owner & Admin Confidential Controls: Profit Margin & Staff Access Sharing */}
            {isOwnerOrAdmin && (
              <div style={{ display: 'flex', flexDirection: 'column', gap: '10px', marginTop: '6px' }}>
                {/* 1. Custom Profit Margin (%) */}
                <div style={{ background: '#f0fdf4', border: '1px solid #bbf7d0', borderRadius: '6px', padding: '12px 14px' }}>
                  <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', gap: '10px' }}>
                    <div>
                      <label className="tiny font-bold" style={{ color: '#166534', display: 'flex', alignItems: 'center', gap: '6px', margin: 0, fontSize: '12px' }}>
                        <span>💰</span> Client Profit Margin (%)
                      </label>
                      <span className="tiny" style={{ color: '#15803d', display: 'block', marginTop: '2px', fontSize: '11px' }}>
                        Sets the profit margin for this client (default 25%). Future quotations for this client will automatically use this value. <strong>Strictly hidden from staff.</strong>
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
                        onChange={(e) => setFormData({ ...formData, profit_margin_pct: Math.max(0, parseFloat(e.target.value) || 0) })}
                        style={{ width: '75px', height: '34px', textAlign: 'right', color: '#166534', background: '#ffffff', borderColor: '#86efac', fontSize: '14px' }}
                      />
                      <span style={{ fontWeight: 700, color: '#166534', fontSize: '13px' }}>%</span>
                    </div>
                  </div>
                </div>

                {/* 2. Share Client Access with Staff */}
                <div style={{ background: '#f8fafc', border: '1px solid #e2e8f0', borderRadius: '6px', padding: '12px 14px' }}>
                  <label className="tiny font-bold" style={{ color: '#0f172a', display: 'flex', alignItems: 'center', gap: '6px', marginBottom: '4px', fontSize: '12px' }}>
                    <span>👥</span> Share Client Access with Staff
                  </label>
                  <span className="tiny muted" style={{ display: 'block', marginBottom: '10px', fontSize: '11px' }}>
                    Choose which staff members can view this client and create quotations. (Staff cannot see profit margin).
                  </span>

                  <div style={{ display: 'flex', flexDirection: 'column', gap: '8px' }}>
                    <label style={{ display: 'flex', alignItems: 'center', gap: '8px', cursor: 'pointer', fontSize: '12px', fontWeight: 600, color: '#1e293b' }}>
                      <input
                        type="checkbox"
                        checked={formData.shared_with.includes('all')}
                        onChange={(e) => {
                          if (e.target.checked) {
                            setFormData({ ...formData, shared_with: ['all'] })
                          } else {
                            setFormData({ ...formData, shared_with: [] })
                          }
                        }}
                      />
                      <span>Share with All Staff Members</span>
                    </label>

                    {!formData.shared_with.includes('all') && (
                      <div style={{ marginTop: '4px' }}>
                        <span className="tiny font-bold muted" style={{ display: 'block', marginBottom: '6px' }}>
                          Or select specific staff members:
                        </span>
                        {orgStaffMembers.length > 0 ? (
                          <div style={{ display: 'flex', flexWrap: 'wrap', gap: '6px' }}>
                            {orgStaffMembers.map((m) => {
                              const isSelected = formData.shared_with.includes(m.username) || (m.user_id && formData.shared_with.includes(m.user_id))
                              return (
                                <button
                                  key={m.username}
                                  type="button"
                                  onClick={() => {
                                    const next = isSelected
                                      ? formData.shared_with.filter((u) => u !== m.username && (!m.user_id || u !== m.user_id))
                                      : [...formData.shared_with, m.username, ...(m.user_id ? [m.user_id] : [])]
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
                        {formData.shared_with.length > 0 && !formData.shared_with.includes('all') && (
                          <div className="row gap mt-2" style={{ flexWrap: 'wrap' }}>
                            {formData.shared_with.map((u) => (
                              <span
                                key={u}
                                className="chip primary sm font-mono"
                                style={{ display: 'inline-flex', alignItems: 'center', gap: '4px', fontSize: '11px' }}
                              >
                                @{u}
                                <span
                                  onClick={() => setFormData({ ...formData, shared_with: formData.shared_with.filter((x) => x !== u) })}
                                  style={{ cursor: 'pointer', fontWeight: 800, marginLeft: '2px' }}
                                >
                                  ✕
                                </span>
                              </span>
                            ))}
                          </div>
                        )}
                      </div>
                    )}
                  </div>
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
                paddingTop: '14px',
                marginTop: '8px',
                gap: '10px',
              }}
            >
              <div>
                {!isCreatingNew && editingClient && (
                  <button
                    type="button"
                    className="btn ghost xs"
                    onClick={handleDelete}
                    style={{ color: '#dc2626', border: '1px solid #cbd5e1' }}
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

                <button
                  type="button"
                  className="btn ghost sm"
                  onClick={handleCloseModal}
                >
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
