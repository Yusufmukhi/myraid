import { useMemo, useState } from 'react'
import { api, Org, Project, Session } from '../../lib/api'
import { inr } from '../../lib/format'
import { cleanOverrides, summarize } from '../../lib/model'
import { Workbook } from '../../lib/engine'
import { Icon, Spinner, useUI } from '../ui'

interface Props {
  session: Session
  org: Org
  projects: Project[]
  wb: Workbook
  onOpenProject: (projectId: string) => void
  onProjectsChanged: () => void
  onOrgUpdated?: () => void
}

interface DiscountRequestItem {
  project: Project
  requested_pct: number
  requested_by: string
  requested_at: string
  status: 'pending' | 'approved' | 'rejected'
  decided_by?: string
  decided_at?: string
}

export function ApprovalsPage({
  session,
  org,
  projects,
  wb,
  onOpenProject,
  onProjectsChanged,
}: Props) {
  const { toast, confirm } = useUI()
  const [actionBusyId, setActionBusyId] = useState<string | null>(null)
  const [discountFilter, setDiscountFilter] = useState<'all' | 'pending' | 'approved' | 'rejected'>('pending')

  // Aggregate discount requests across all projects
  const discountRequests: DiscountRequestItem[] = useMemo(() => {
    const list: DiscountRequestItem[] = []
    for (const p of projects) {
      const req = p.summary?.discount_request
      if (req && req.requested_pct) {
        list.push({
          project: p,
          requested_pct: req.requested_pct,
          requested_by: req.requested_by || 'Staff',
          requested_at: req.requested_at || p.updated_at,
          status: req.status || 'pending',
          decided_by: req.decided_by,
          decided_at: req.decided_at,
        })
      }
    }
    return list.sort((a, b) => new Date(b.requested_at).getTime() - new Date(a.requested_at).getTime())
  }, [projects])

  const pendingDiscountCount = discountRequests.filter((d) => d.status === 'pending').length

  // Handlers for Discount Requests
  const handleApproveDiscount = async (item: DiscountRequestItem) => {
    setActionBusyId(item.project.id)
    try {
      const p = item.project
      const updatedReq = {
        ...p.summary.discount_request,
        status: 'approved' as const,
        decided_by: session.username,
        decided_at: new Date().toISOString(),
      }
      const updatedSystems = Array.isArray(p.summary?.systems)
        ? p.summary.systems.map((sys: any) => ({
            ...sys,
            discountPct: item.requested_pct,
            overrides: { ...(sys.overrides || {}), __discount_pct: item.requested_pct },
          }))
        : p.summary?.systems
      const updatedSummary = {
        ...p.summary,
        discount_request: updatedReq,
        discount_pct: item.requested_pct,
        systems: updatedSystems,
      }
      const nextOverrides: Record<string, number | null> = {
        ...(p.inputs || {}),
        __discount_pct: item.requested_pct,
      }
      const profitVal = nextOverrides['__profit_pct'] !== undefined && nextOverrides['__profit_pct'] !== null
        ? Number(nextOverrides['__profit_pct'])
        : 25
      const s = summarize(wb, profitVal, item.requested_pct)
      const prevDiscount = Number(p.summary?.discount_pct || p.inputs?.['__discount_pct'] || 0)
      const finalSales = Array.isArray(updatedSystems) && updatedSystems.length > 1 && p.total
        ? Math.round((p.total / Math.max(0.01, 1 - (prevDiscount / 100))) * (1 - (item.requested_pct / 100)))
        : s.sales
      await api.projectUpdate(
        session.token,
        p.id,
        p.name,
        p.status,
        cleanOverrides(wb, nextOverrides),
        p.formula_overrides,
        finalSales,
        updatedSummary,
        `Owner/Admin authorized ${item.requested_pct}% discount requested by @${item.requested_by}`,
      )
      toast(`Approved ${item.requested_pct}% discount for "${p.name}".`)
      onProjectsChanged()
    } catch (e) {
      toast((e as Error).message, 'err')
    } finally {
      setActionBusyId(null)
    }
  }

  const handleRejectDiscount = async (item: DiscountRequestItem) => {
    const ok = await confirm({
      title: 'Decline Discount Request',
      body: `Decline the ${item.requested_pct}% discount requested by @${item.requested_by} for "${item.project.name}"? The discount will revert to default threshold.`,
      confirmLabel: 'Decline Discount',
      danger: true,
    })
    if (!ok) return

    setActionBusyId(item.project.id)
    try {
      const p = item.project
      const updatedReq = {
        ...p.summary.discount_request,
        status: 'rejected' as const,
        decided_by: session.username,
        decided_at: new Date().toISOString(),
      }
      const fallbackPct = Math.min(10, Number(p.inputs?.['__discount_pct'] || 0))
      const updatedSystems = Array.isArray(p.summary?.systems)
        ? p.summary.systems.map((sys: any) => ({
            ...sys,
            discountPct: fallbackPct,
            overrides: { ...(sys.overrides || {}), __discount_pct: fallbackPct },
          }))
        : p.summary?.systems
      const updatedSummary = {
        ...p.summary,
        discount_request: updatedReq,
        discount_pct: fallbackPct,
        systems: updatedSystems,
      }
      const nextOverrides: Record<string, number | null> = {
        ...(p.inputs || {}),
        __discount_pct: fallbackPct,
      }
      const profitVal = nextOverrides['__profit_pct'] !== undefined && nextOverrides['__profit_pct'] !== null
        ? Number(nextOverrides['__profit_pct'])
        : 25
      const s = summarize(wb, profitVal, fallbackPct)
      const prevDiscount = Number(p.summary?.discount_pct || p.inputs?.['__discount_pct'] || 0)
      const finalSales = Array.isArray(updatedSystems) && updatedSystems.length > 1 && p.total
        ? Math.round((p.total / Math.max(0.01, 1 - (prevDiscount / 100))) * (1 - (fallbackPct / 100)))
        : s.sales
      await api.projectUpdate(
        session.token,
        p.id,
        p.name,
        p.status,
        cleanOverrides(wb, nextOverrides),
        p.formula_overrides,
        finalSales,
        updatedSummary,
        `Owner/Admin declined ${item.requested_pct}% discount requested by @${item.requested_by}`,
      )
      toast(`Declined ${item.requested_pct}% discount request for "${p.name}".`)
      onProjectsChanged()
    } catch (e) {
      toast((e as Error).message, 'err')
    } finally {
      setActionBusyId(null)
    }
  }

  // Filtered list
  const filteredDiscounts = discountRequests.filter((d) =>
    discountFilter === 'all' ? true : d.status === discountFilter
  )

  return (
    <div className="approvals-page" style={{ padding: '24px 32px', maxWidth: '1200px', margin: '0 auto' }}>
      {/* Page Header */}
      <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: '24px', flexWrap: 'wrap', gap: '16px' }}>
        <div>
          <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
            <h2 style={{ margin: 0, fontSize: '24px', fontWeight: 800 }}>Quotation Approvals</h2>
            <span className="chip primary font-mono">{org.name}</span>
          </div>
          <p className="tiny muted" style={{ margin: '4px 0 0 0', fontSize: '13px' }}>
            Review and authorize staff quotation discount requests and special pricing.
          </p>
        </div>
      </div>

      {/* Header Info Banner */}
      <div
        style={{
          padding: '12px 16px',
          background: '#fffbeb',
          border: '1px solid #fde68a',
          borderRadius: '8px',
          marginBottom: '20px',
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'space-between',
          gap: '12px',
          flexWrap: 'wrap',
        }}
      >
        <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
          <span style={{ fontSize: '18px' }}>💰</span>
          <div style={{ fontSize: '12.5px', color: '#92400e' }}>
            <strong>Discount Guardrail Threshold:</strong> Non-admin staff cannot apply discounts above <b>10%</b> without explicit Owner / Admin approval. Special requests appear here.
          </div>
        </div>
        <button className="btn ghost sm" onClick={onProjectsChanged}>
          {Icon.refresh} Refresh Quotations
        </button>
      </div>

      {/* Filter Pills */}
      <div style={{ display: 'flex', gap: '8px', marginBottom: '16px' }}>
        {(['pending', 'all', 'approved', 'rejected'] as const).map((st) => (
          <button
            key={st}
            className={`chip ${discountFilter === st ? 'primary' : ''}`}
            onClick={() => setDiscountFilter(st)}
            style={{ textTransform: 'capitalize', cursor: 'pointer' }}
          >
            {st} {st === 'pending' ? `(${pendingDiscountCount})` : ''}
          </button>
        ))}
      </div>

      {/* Discounts Table */}
      {filteredDiscounts.length === 0 ? (
        <div
          className="card"
          style={{
            padding: '40px',
            textAlign: 'center',
            color: '#64748b',
            background: '#f8fafc',
            border: '1px dashed #cbd5e1',
          }}
        >
          <div style={{ fontSize: '32px', marginBottom: '8px' }}>🏷️</div>
          <h4 style={{ margin: 0, fontSize: '16px', color: '#334155' }}>
            No {discountFilter !== 'all' ? discountFilter : ''} discount requests
          </h4>
          <p className="tiny muted" style={{ marginTop: '4px' }}>
            {discountFilter === 'pending'
              ? 'All staff discount authorization requests have been handled.'
              : 'Whenever a staff member enters an 11% or higher discount, it appears here for approval.'}
          </p>
        </div>
      ) : (
        <div className="card" style={{ padding: 0, overflow: 'hidden' }}>
          <table className="tbl" style={{ width: '100%', margin: 0 }}>
            <thead>
              <tr>
                <th>Quotation</th>
                <th>Requested By</th>
                <th className="r">Discount %</th>
                <th className="r">Total Value</th>
                <th>Date</th>
                <th>Status</th>
                <th className="r">Actions</th>
              </tr>
            </thead>
            <tbody>
              {filteredDiscounts.map((item) => {
                const isBusy = actionBusyId === item.project.id
                return (
                  <tr key={item.project.id}>
                    <td>
                      <div>
                        <div
                          style={{ fontWeight: 700, fontSize: '13.5px', color: '#2563eb', cursor: 'pointer' }}
                          onClick={() => onOpenProject(item.project.id)}
                          title="Open quotation in calculator"
                        >
                          {item.project.name}
                        </div>
                        <span className="tiny muted">v{item.project.template_version}</span>
                      </div>
                    </td>
                    <td style={{ verticalAlign: 'middle' }}>
                      <span style={{ fontWeight: 600 }}>@{item.requested_by}</span>
                    </td>
                    <td className="r font-mono" style={{ verticalAlign: 'middle' }}>
                      <span
                        style={{
                          background: '#fef3c7',
                          color: '#92400e',
                          padding: '2px 8px',
                          borderRadius: '4px',
                          fontWeight: 800,
                          fontSize: '13px',
                        }}
                      >
                        {item.requested_pct}%
                      </span>
                    </td>
                    <td className="r font-mono font-bold text-green" style={{ verticalAlign: 'middle', fontSize: '14px' }}>
                      {item.project.total ? inr(Number(item.project.total)) : '—'}
                    </td>
                    <td className="tiny muted" style={{ verticalAlign: 'middle' }}>
                      {new Date(item.requested_at).toLocaleDateString('en-IN', {
                        month: 'short',
                        day: 'numeric',
                        hour: '2-digit',
                        minute: '2-digit',
                      })}
                    </td>
                    <td style={{ verticalAlign: 'middle' }}>
                      <span
                        className="status-pill"
                        style={{
                          background:
                            item.status === 'approved'
                              ? '#dcfce7'
                              : item.status === 'rejected'
                              ? '#fee2e2'
                              : '#fef3c7',
                          color:
                            item.status === 'approved'
                              ? '#15803d'
                              : item.status === 'rejected'
                              ? '#b91c1c'
                              : '#b45309',
                          fontWeight: 700,
                          fontSize: '11px',
                          textTransform: 'uppercase',
                        }}
                      >
                        {item.status}
                      </span>
                    </td>
                    <td className="r" style={{ verticalAlign: 'middle' }}>
                      {item.status === 'pending' ? (
                        <div style={{ display: 'inline-flex', gap: '6px' }}>
                          <button
                            className="btn primary sm"
                            style={{ background: '#16a34a', borderColor: '#15803d', fontSize: '12px', padding: '4px 10px' }}
                            onClick={() => handleApproveDiscount(item)}
                            disabled={isBusy}
                          >
                            {isBusy ? <Spinner /> : <span style={{ display: 'inline-flex', alignItems: 'center', gap: 4 }}>{Icon.check} Authorize</span>}
                          </button>
                          <button
                            className="btn danger sm"
                            style={{ fontSize: '12px', padding: '4px 10px' }}
                            onClick={() => handleRejectDiscount(item)}
                            disabled={isBusy}
                          >
                            Decline
                          </button>
                        </div>
                      ) : (
                        <span className="tiny muted">
                          {item.decided_by ? `By @${item.decided_by}` : 'Decision logged'}
                        </span>
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
  )
}
