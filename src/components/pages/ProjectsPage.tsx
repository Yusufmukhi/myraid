import { Fragment, useEffect, useMemo, useState } from 'react'
import { api, Org, Project, ProjectStatus, Session } from '../../lib/api'
import { deleteProject } from '../../lib/permissions'
import { inr, timeAgo } from '../../lib/format'
import { groupProjectsIntoFolders, parseRevision, ProjectFolder } from '../../lib/revision'
import { Icon, Spinner, useUI } from '../ui'

interface Props {
  session: Session
  org: Org
  currentProjectId: string | null
  onOpenProject: (projectId: string) => void
  onNewProject: () => void
  onProjectsChanged?: () => void
  initialScope?: 'all' | 'mine' | 'shared'
}

type SortField = 'updated_at' | 'name' | 'total' | 'created_at'

export function ProjectsPage({ session, org, currentProjectId, onOpenProject, onNewProject, onProjectsChanged, initialScope }: Props) {
  const { toast, confirm } = useUI()
  const isAdminOrOwner = org.role === 'owner' || org.role === 'admin'
  const [projects, setProjects] = useState<Project[]>([])
  const [loading, setLoading] = useState(true)
  const [search, setSearch] = useState('')
  const [statusFilter, setStatusFilter] = useState<string>('all')
  const [scopeFilter, setScopeFilter] = useState<'all' | 'mine' | 'shared'>(initialScope || 'all')
  const [showArchived, setShowArchived] = useState(false)
  const [sortBy, setSortBy] = useState<SortField>('updated_at')
  const [sortAsc, setSortAsc] = useState(false)
  const [viewMode, setViewMode] = useState<'folders' | 'flat'>('folders')
  // By default, all folders are closed (collapsed) until clicked
  const [openFolders, setOpenFolders] = useState<Record<string, boolean>>({})

  useEffect(() => {
    if (initialScope) {
      setScopeFilter(initialScope)
    }
  }, [initialScope])

  const toggleFolder = (folderId: string) => {
    setOpenFolders((prev) => ({
      ...prev,
      [folderId]: !prev[folderId],
    }))
  }

  const expandAllFolders = () => {
    const map: Record<string, boolean> = {}
    for (const f of projectFolders) {
      map[f.id] = true
    }
    setOpenFolders(map)
  }

  const collapseAllFolders = () => {
    setOpenFolders({})
  }

  // Modals state
  const [reviseProj, setReviseProj] = useState<Project | null>(null)
  const [revName, setRevName] = useState('')
  const [renameProj, setRenameProj] = useState<Project | null>(null)
  const [newName, setNewName] = useState('')
  const [duplicateProj, setDuplicateProj] = useState<Project | null>(null)
  const [dupName, setDupName] = useState('')
  const [dupUseLatest, setDupUseLatest] = useState(false)
  const [shareProj, setShareProj] = useState<Project | null>(null)
  const [shareMembers, setShareMembers] = useState<any[]>([])
  const [loadingMembers, setLoadingMembers] = useState(false)

  // Batch selection and deletion state
  const [selectedIds, setSelectedIds] = useState<Set<string>>(new Set())
  const [singleDeleteProj, setSingleDeleteProj] = useState<Project | null>(null)
  const [confirmNameInput, setConfirmNameInput] = useState('')
  const [deletingSingle, setDeletingSingle] = useState(false)

  const [showBatchDeleteModal, setShowBatchDeleteModal] = useState(false)
  const [batchDeletePassword, setBatchDeletePassword] = useState('')
  const [batchDeleteError, setBatchDeleteError] = useState('')
  const [deletingBatch, setDeletingBatch] = useState(false)

  const openShareModal = async (p: Project) => {
    setShareProj(p)
    setLoadingMembers(true)
    try {
      const list = await api.orgMembers(session.token, org.id)
      setShareMembers(Array.isArray(list) ? list : [])
    } catch {
      setShareMembers([])
    } finally {
      setLoadingMembers(false)
    }
  }

  const handleToggleShare = async (targetUsername: string) => {
    if (!shareProj) return
    const currentShared = Array.isArray(shareProj.summary?.shared_with)
      ? [...shareProj.summary.shared_with]
      : []
    let updated: string[]
    if (currentShared.includes(targetUsername)) {
      updated = currentShared.filter((u) => u !== targetUsername)
    } else {
      updated = [...currentShared, targetUsername]
    }
    try {
      await api.projectShare(session.token, shareProj.id, updated)
      const nextSummary = { ...(shareProj.summary || {}), shared_with: updated }
      setShareProj({ ...shareProj, summary: nextSummary })
      toast(
        updated.includes(targetUsername)
          ? `Quotation shared with @${targetUsername}`
          : `Access revoked for @${targetUsername}`,
      )
      loadProjects()
      onProjectsChanged?.()
    } catch (e) {
      toast((e as Error).message, 'err')
    }
  }

  const loadProjects = async () => {
    try {
      setLoading(true)
      const data = await api.projectList(session.token, org.id, true)
      setProjects(data)
    } catch (e) {
      toast((e as Error).message, 'err')
    } finally {
      setLoading(false)
    }
  }

  useEffect(() => {
    loadProjects()
  }, [org.id])

  // Count projects by scope (My Quotations vs Shared with Me)
  const { myProjectsCount, sharedProjectsCount } = useMemo(() => {
    let mine = 0
    let shared = 0
    for (const p of projects) {
      if (!showArchived && p.is_archived) continue
      const sharedList = (p.summary?.shared_with || (p as any).shared_with) as string[] | undefined
      const isSharedWithMe =
        Array.isArray(sharedList) &&
        (sharedList.includes(session.username) || (session.user_id ? sharedList.includes(session.user_id) : false)) &&
        p.created_by !== session.username &&
        p.created_by !== session.user_id

      if (isSharedWithMe) {
        if (p.status?.toLowerCase() !== 'approved') {
          shared++
        }
      } else if (p.created_by === session.username || p.created_by === session.user_id) {
        mine++
      }
    }
    return { myProjectsCount: mine, sharedProjectsCount: shared }
  }, [projects, showArchived, session.username, session.user_id])

  const filteredProjects = useMemo(() => {
    return (projects || [])
      .filter((p) => {
        const sharedList = (p.summary?.shared_with || (p as any).shared_with) as string[] | undefined
        const isSharedWithMe =
          Array.isArray(sharedList) &&
          (sharedList.includes(session.username) || (session.user_id ? sharedList.includes(session.user_id) : false)) &&
          p.created_by !== session.username &&
          p.created_by !== session.user_id

        const isMine =
          p.created_by === session.user_id ||
          p.created_by === session.username

        // Staff can only see their own projects and active/quoted projects shared with them (approved quotations removed from shared view)
        if (!isAdminOrOwner) {
          if (!isMine && !isSharedWithMe) return false
          if (isSharedWithMe && p.status?.toLowerCase() === 'approved') return false
        }

        // Scope filter
        if (scopeFilter === 'mine' && !isMine) return false
        if (scopeFilter === 'shared' && !isSharedWithMe) return false

        if (!showArchived && p.is_archived && statusFilter !== 'lost' && statusFilter !== 'loss') return false
        const isLoss = p.is_archived || p.status === 'loss' || p.status === 'archived' || p.status === 'lost'
        const isWon = !p.is_archived && (p.status === 'won' || p.status === 'approved')
        const isNegotiating = p.status === 'negotiating'
        const isQuoted = p.status === 'quoted' || (!isLoss && !isWon && !isNegotiating)

        if (statusFilter === 'all') {
          // Default ALL view only shows active pipeline (Quoted & Negotiating); excludes Won and Loss
          if (isWon || isLoss) return false
        } else {
          if ((statusFilter === 'won' || statusFilter === 'approved') && !isWon) return false
          if ((statusFilter === 'loss' || statusFilter === 'lost') && !isLoss) return false
          if (statusFilter === 'negotiating' && !isNegotiating) return false
          if (statusFilter === 'quoted' && !isQuoted) return false
        }
        if (search.trim()) {
          const q = search.toLowerCase().trim()
          // 1. Direct project name, creator, id, status
          if (p.name.toLowerCase().includes(q)) return true
          if (p.created_by && p.created_by.toLowerCase().includes(q)) return true
          if (p.id && p.id.toLowerCase().includes(q)) return true
          if (p.status && p.status.toLowerCase().includes(q)) return true

          // 2. Numeric totals & formatted currency amounts
          const tot = Number(p.total) || 0
          if (tot.toString().includes(q)) return true
          if (Math.round(tot).toString().includes(q)) return true
          if (tot.toLocaleString('en-IN').toLowerCase().includes(q)) return true

          // 3. Client info, address, city, state, pin, GST, email, contacts, custom rows
          const clientInfo = p.summary?.client_info
          if (clientInfo) {
            if ((clientInfo.buyer_name || '').toLowerCase().includes(q)) return true
            if ((clientInfo.buyer_address || '').toLowerCase().includes(q)) return true
            if ((clientInfo.buyer_city || '').toLowerCase().includes(q)) return true
            if ((clientInfo.buyer_state || '').toLowerCase().includes(q)) return true
            if ((clientInfo.buyer_pin || '').toLowerCase().includes(q)) return true
            if ((clientInfo.buyer_gst || '').toLowerCase().includes(q)) return true
            if ((clientInfo.buyer_contact || '').toLowerCase().includes(q)) return true
            if ((clientInfo.buyer_email || '').toLowerCase().includes(q)) return true
            if (Array.isArray(clientInfo.contacts)) {
              for (const c of clientInfo.contacts) {
                if ((c.name || '').toLowerCase().includes(q)) return true
                if ((c.phone || '').toLowerCase().includes(q)) return true
                if ((c.designation || '').toLowerCase().includes(q)) return true
                if ((c.email || '').toLowerCase().includes(q)) return true
              }
            }
            if (Array.isArray(clientInfo.custom_rows)) {
              for (const row of clientInfo.custom_rows) {
                if ((row.label || '').toLowerCase().includes(q)) return true
                if ((row.value || '').toLowerCase().includes(q)) return true
              }
            }
          }

          // 4. Quotation fields, systems, model codes, and all calculated values in summary
          if (p.summary) {
            try {
              const summaryStr = JSON.stringify(p.summary).toLowerCase()
              if (summaryStr.includes(q)) return true
            } catch {}
          }

          // 5. Input specifications
          if (p.inputs) {
            try {
              const inputsStr = JSON.stringify(p.inputs).toLowerCase()
              if (inputsStr.includes(q)) return true
            } catch {}
          }

          return false
        }
        return true
      })
      .sort((a, b) => {
        let cmp = 0
        if (sortBy === 'name') cmp = a.name.localeCompare(b.name)
        else if (sortBy === 'total') cmp = (Number(a.total) || 0) - (Number(b.total) || 0)
        else if (sortBy === 'created_at') cmp = new Date(a.created_at).getTime() - new Date(b.created_at).getTime()
        else cmp = new Date(a.updated_at).getTime() - new Date(b.updated_at).getTime()
        return sortAsc ? cmp : -cmp
      })
  }, [projects, isAdminOrOwner, session.username, session.user_id, scopeFilter, search, statusFilter, showArchived, sortBy, sortAsc])

  // Group filtered projects into Project Folders
  const projectFolders = useMemo(() => {
    const folders = groupProjectsIntoFolders(filteredProjects, currentProjectId)
    return folders.sort((a, b) => {
      let cmp = 0
      if (sortBy === 'name') cmp = a.folderName.localeCompare(b.folderName)
      else if (sortBy === 'total') cmp = (Number(a.latestRevision.total) || 0) - (Number(b.latestRevision.total) || 0)
      else if (sortBy === 'created_at') cmp = new Date(a.revisions[0].created_at).getTime() - new Date(b.revisions[0].created_at).getTime()
      else cmp = new Date(a.latestRevision.updated_at).getTime() - new Date(b.latestRevision.updated_at).getTime()
      return sortAsc ? cmp : -cmp
    })
  }, [filteredProjects, currentProjectId, sortBy, sortAsc])

  // Selection helpers
  const allVisibleSelected = useMemo(() => {
    if (filteredProjects.length === 0) return false
    return filteredProjects.every((p) => selectedIds.has(p.id))
  }, [filteredProjects, selectedIds])

  const someVisibleSelected = useMemo(() => {
    return filteredProjects.some((p) => selectedIds.has(p.id))
  }, [filteredProjects, selectedIds])

  const handleToggleSelectAll = () => {
    if (allVisibleSelected) {
      setSelectedIds((prev) => {
        const next = new Set(prev)
        for (const p of filteredProjects) {
          next.delete(p.id)
        }
        return next
      })
    } else {
      setSelectedIds((prev) => {
        const next = new Set(prev)
        for (const p of filteredProjects) {
          next.add(p.id)
        }
        return next
      })
    }
  }

  const handleToggleSelectOne = (id: string, checked: boolean) => {
    setSelectedIds((prev) => {
      const next = new Set(prev)
      if (checked) {
        next.add(id)
      } else {
        next.delete(id)
      }
      return next
    })
  }

  const handleToggleFolderSelect = (folder: ProjectFolder, checked: boolean) => {
    setSelectedIds((prev) => {
      const next = new Set(prev)
      for (const rev of folder.revisions) {
        if (checked) {
          next.add(rev.id)
        } else {
          next.delete(rev.id)
        }
      }
      return next
    })
  }

  // Actions
  const handleRevise = async () => {
    if (!reviseProj) return
    const nm = revName.trim()
    if (!nm) return toast('Revision name cannot be empty', 'err')
    try {
      const res = await api.projectDuplicate(session.token, reviseProj.id, nm, false)
      try {
        await api.projectUpdate(
          session.token,
          res.id,
          res.name,
          'quoted',
          reviseProj.inputs,
          reviseProj.formula_overrides,
          reviseProj.total,
          reviseProj.summary,
          `Revision ${res.name} created - Status set to Quoted`,
        )
      } catch (err) {
        await api.projectUpdate(
          session.token,
          res.id,
          res.name,
          'quoted',
          reviseProj.inputs,
          reviseProj.formula_overrides,
          reviseProj.total,
          reviseProj.summary,
          `Revision ${res.name} created`,
        )
      }
      toast(`Created revision “${res.name}” (Status: Quoted)`)
      setReviseProj(null)
      await loadProjects()
      onProjectsChanged?.()
      onOpenProject(res.id)
    } catch (e) {
      toast((e as Error).message, 'err')
    }
  }

  const handleRename = async () => {
    if (!renameProj) return
    const nm = newName.trim()
    if (!nm) return toast('Project name cannot be empty', 'err')
    try {
      await api.projectUpdate(session.token, renameProj.id, nm, renameProj.status, renameProj.inputs, renameProj.formula_overrides, renameProj.total, renameProj.summary, 'Renamed project')
      toast('Project renamed')
      setRenameProj(null)
      loadProjects()
      onProjectsChanged?.()
    } catch (e) {
      toast((e as Error).message, 'err')
    }
  }

  const handleDuplicate = async () => {
    if (!duplicateProj) return
    const nm = dupName.trim()
    if (!nm) return toast('Duplicate project name cannot be empty', 'err')
    try {
      const res = await api.projectDuplicate(session.token, duplicateProj.id, nm, dupUseLatest)
      try {
        await api.projectUpdate(
          session.token,
          res.id,
          nm,
          'quoted',
          duplicateProj.inputs,
          duplicateProj.formula_overrides,
          duplicateProj.total,
          duplicateProj.summary,
          `Quotation ${nm} created as copy - Status set to Quoted`,
        )
      } catch {}
      toast(`Project copied successfully (Status: Quoted)`)
      setDuplicateProj(null)
      await loadProjects()
      onProjectsChanged?.()
      onOpenProject(res.id)
    } catch (e) {
      toast((e as Error).message, 'err')
    }
  }

  const handleArchive = async (p: Project) => {
    if (!isAdminOrOwner) {
      toast('Staff members cannot archive or restore projects.', 'err')
      return
    }
    const toArchive = !p.is_archived
    const ok = await confirm({
      title: toArchive ? `Archive “${p.name}”?` : `Unarchive “${p.name}”?`,
      body: toArchive
        ? 'Archived projects are hidden from normal views and dashboard statistics but their historical calculations are preserved.'
        : 'The project will be restored to active projects.',
      confirmLabel: toArchive ? 'Archive' : 'Restore',
    })
    if (!ok) return
    try {
      await api.projectArchive(session.token, p.id, toArchive)
      toast(toArchive ? 'Project archived' : 'Project restored')
      loadProjects()
      onProjectsChanged?.()
    } catch (e) {
      toast((e as Error).message, 'err')
    }
  }

  const handleDelete = (p: Project) => {
    if (!isAdminOrOwner) {
      toast('Staff members cannot delete projects or quotations.', 'err')
      return
    }
    setSingleDeleteProj(p)
    setConfirmNameInput('')
  }

  const handleConfirmSingleDelete = async () => {
    if (!singleDeleteProj) return
    if (confirmNameInput.trim() !== singleDeleteProj.name.trim()) {
      toast('Quotation name does not match', 'err')
      return
    }
    setDeletingSingle(true)
    try {
      await deleteProject(session, org, singleDeleteProj.id)
      toast(`Deleted quotation “${singleDeleteProj.name}”`)
      setSelectedIds((prev) => {
        const next = new Set(prev)
        next.delete(singleDeleteProj.id)
        return next
      })
      setSingleDeleteProj(null)
      setConfirmNameInput('')
      await loadProjects()
      onProjectsChanged?.()
    } catch (e) {
      toast((e as Error).message, 'err')
    } finally {
      setDeletingSingle(false)
    }
  }

  const handleTriggerBatchDelete = () => {
    if (!isAdminOrOwner) {
      toast('Staff members cannot delete projects or quotations.', 'err')
      return
    }
    if (selectedIds.size === 0) return
    if (selectedIds.size === 1) {
      const singleId = Array.from(selectedIds)[0]
      const target = projects.find((p) => p.id === singleId)
      if (target) {
        handleDelete(target)
      }
      return
    }
    setShowBatchDeleteModal(true)
    setBatchDeletePassword('')
    setBatchDeleteError('')
  }

  const handleConfirmBatchDelete = async () => {
    if (!batchDeletePassword) return
    setDeletingBatch(true)
    setBatchDeleteError('')
    try {
      try {
        await api.login(session.username, batchDeletePassword)
      } catch {
        setBatchDeleteError('Incorrect password. Please enter your valid account password.')
        setDeletingBatch(false)
        return
      }

      const idsToDelete = Array.from(selectedIds)
      for (const id of idsToDelete) {
        await deleteProject(session, org, id)
      }

      toast(`Successfully deleted ${idsToDelete.length} quotations`)
      setSelectedIds(new Set())
      setShowBatchDeleteModal(false)
      setBatchDeletePassword('')
      await loadProjects()
      onProjectsChanged?.()
    } catch (e) {
      toast((e as Error).message, 'err')
    } finally {
      setDeletingBatch(false)
    }
  }

  // Render individual quotation revision file row
  const renderProjectRow = (p: Project, isSubRow = false) => {
    const isSelected = p.id === currentProjectId
    const hasNewerTemplate = (org.current_template_version || 1) > p.template_version
    const rev = parseRevision(p.name)

    const buyer = p.summary?.client_info?.buyer_name || ''
    const isWon = !p.is_archived && (p.status === 'won' || p.status === 'approved')
    const isLoss = p.is_archived || p.status === 'loss' || p.status === 'archived' || p.status === 'lost'
    const isNegotiating = p.status === 'negotiating'
    const statusLabel = isWon ? 'Won' : isLoss ? 'Loss' : isNegotiating ? 'Negotiating' : 'Quoted'
    const statusClass = isWon ? 'won' : isLoss ? 'loss' : isNegotiating ? 'negotiating' : 'quoted'

    return (
      <tr
        key={p.id}
        className={`${isSubRow ? 'file-subrow' : ''} ${isSelected ? 'selected-row' : ''} ${p.is_archived ? 'dim' : ''} ${selectedIds.has(p.id) ? 'batch-selected-row' : ''}`}
      >
        <td style={{ width: '40px', textAlign: 'center' }} onClick={(e) => e.stopPropagation()}>
          <input
            type="checkbox"
            aria-label={`Select quotation ${p.name}`}
            checked={selectedIds.has(p.id)}
            onChange={(e) => handleToggleSelectOne(p.id, e.target.checked)}
          />
        </td>
        <td className={isSubRow ? 'file-subrow-cell' : ''}>
          {isSubRow && (
            <>
              <div className="file-tree-guide" />
              <div className="file-tree-branch" />
            </>
          )}
          <div style={{ display: 'flex', flexDirection: 'column', gap: 2 }}>
            <div className="row gap" style={{ flexWrap: 'wrap' }}>
              {isSubRow && <span className="file-icon" title="Quotation File">{Icon.fileText}</span>}
              <button className="proj-title-btn" onClick={() => onOpenProject(p.id)}>
                <strong>{p.name}</strong>
              </button>
              {rev.currentRev ? (
                <span className="chip primary sm font-mono font-bold" title={`Revision ${rev.currentRev}`}>
                  Rev {rev.currentRev}
                </span>
              ) : null}
              {isSelected && <span className="chip ok">Current</span>}
              {p.is_archived && <span className="chip muted-chip">Archived</span>}
              {(() => {
                const shared = (p.summary?.shared_with || (p as any).shared_with) as string[] | undefined
                const isSharedToMe =
                  Array.isArray(shared) &&
                  (shared.includes(session.username) || (session.user_id && shared.includes(session.user_id))) &&
                  p.created_by !== session.username
                const count = Array.isArray(shared) ? shared.length : 0
                if (isSharedToMe) {
                  return (
                    <span className="chip purple sm" title={`Shared with you by ${p.created_by || 'teammate'}`}>
                      Shared with you
                    </span>
                  )
                }
                if (count > 0 && Array.isArray(shared)) {
                  return (
                    <span className="chip sm font-mono" title={`Shared with: ${shared.join(', ')}`}>
                      Shared ({count})
                    </span>
                  )
                }
                return null
              })()}
            </div>
          </div>
        </td>
        <td>
          {(() => {
            const clientName = p.summary?.client_info?.buyer_name || ''
            const qtnNo = p.summary?.client_info?.qtn_no
            if (!clientName) return <span className="muted tiny">—</span>
            return (
              <div style={{ display: 'flex', flexDirection: 'column', gap: 2 }}>
                <strong style={{ color: '#0f172a', fontSize: '12px' }}>{clientName}</strong>
                {qtnNo && <span className="tiny font-mono text-muted" style={{ fontSize: '10.5px' }}>{qtnNo}</span>}
              </div>
            )
          })()}
        </td>
        <td>
          <span className={`status-pill status-${statusClass}`}>{statusLabel}</span>
          {p.summary?.discount_request?.status === 'pending' && (
            <span
              className="chip warning sm"
              style={{ marginLeft: '4px', fontSize: '11px', padding: '2px 6px' }}
              title={`Staff @${p.summary.discount_request.requested_by} requested ${p.summary.discount_request.requested_pct}% discount`}
            >
              ⏳ {p.summary.discount_request.requested_pct}% Disc Req
            </span>
          )}
        </td>
        <td>
          <div className="template-badge-wrap">
            <span className="version-pill">v{p.template_version}</span>
            {hasNewerTemplate && (
              <span
                className="upgrade-hint"
                title={`New default template v${org.current_template_version} available`}
              >
                ↑ v{org.current_template_version} avail
              </span>
            )}
          </div>
        </td>
        <td className="r font-bold text-green">
          {p.total ? inr(Number(p.total)) : '—'}
        </td>
        <td className="muted tiny">
          {p.created_by || 'Unknown'}
        </td>
        <td className="muted tiny">{timeAgo(p.updated_at)}</td>
        <td className="r">
          <div className="row-actions">
            <button
              className="btn primary sm"
              onClick={() => onOpenProject(p.id)}
              title="Open Calculator"
            >
              Open
            </button>
            <button
              className="btn ghost sm"
              title={`Revise quotation (Create Rev ${parseRevision(p.name).nextRev})`}
              onClick={() => {
                const info = parseRevision(p.name)
                setReviseProj(p)
                setRevName(info.nextName)
              }}
            >
              {Icon.refresh} Revise
            </button>
            <button
              className="icon-btn sm"
              title="Share quotation with team members"
              onClick={() => openShareModal(p)}
            >
              {Icon.users}
            </button>
            <button
              className="icon-btn sm"
              title="Rename"
              onClick={() => {
                setRenameProj(p)
                setNewName(p.name)
              }}
            >
              {Icon.edit}
            </button>
            <button
              className="icon-btn sm"
              title="Duplicate snapshot"
              onClick={() => {
                setDuplicateProj(p)
                setDupName(`${p.name} (Copy)`)
                setDupUseLatest(false)
              }}
            >
              {Icon.copy}
            </button>
            <button
              className="icon-btn sm"
              title={p.is_archived ? 'Unarchive' : 'Archive'}
              onClick={() => handleArchive(p)}
            >
              {Icon.archive}
            </button>
            {isAdminOrOwner && (
              <button
                className="icon-btn sm del"
                title="Delete permanently"
                onClick={() => handleDelete(p)}
              >
                {Icon.trash}
              </button>
            )}
          </div>
        </td>
      </tr>
    )
  }

  return (
    <div className="projects-page">
      {/* Page Header with Folders / Files View Toggle */}
      <div className="row between mb-2" style={{ flexWrap: 'wrap', gap: 10 }}>
        <div className="row gap" style={{ flexWrap: 'wrap' }}>
          <h2 style={{ fontSize: '18px' }}>
            {isAdminOrOwner
              ? scopeFilter === 'shared'
                ? 'Shared Quotations'
                : scopeFilter === 'mine'
                  ? 'My Projects & Quotations'
                  : 'Projects & Quotations'
              : scopeFilter === 'shared'
                ? 'Quotations Shared with Me'
                : scopeFilter === 'mine'
                  ? 'My Quotations'
                  : 'My Projects & Shared Quotations'}
          </h2>
          <div className="view-mode-tabs">
            <button
              className={`view-mode-btn ${viewMode === 'folders' ? 'active' : ''}`}
              onClick={() => setViewMode('folders')}
              title="Group quotations into Project Folders"
            >
              <span className="folder-icon">{Icon.folder}</span> Folders ({projectFolders.length})
            </button>
            <button
              className={`view-mode-btn ${viewMode === 'flat' ? 'active' : ''}`}
              onClick={() => setViewMode('flat')}
              title="Show all individual revision files"
            >
              <span className="file-icon">{Icon.fileText}</span> All Files ({filteredProjects.length})
            </button>
          </div>
          {viewMode === 'folders' && projectFolders.length > 0 && (
            <div className="row gap" style={{ marginLeft: 4 }}>
              <button className="link-inline tiny" onClick={expandAllFolders}>Expand All</button>
              <span className="tiny text-muted">·</span>
              <button className="link-inline tiny" onClick={collapseAllFolders}>Collapse All</button>
            </div>
          )}
        </div>
        <button className="btn primary sm" onClick={onNewProject}>
          {Icon.plus} New Project
        </button>
      </div>

      {/* Filter and Search Bar */}
      <div className="card filter-bar">
        <div className="search-wrap">
          <span className="search-icon">{Icon.search}</span>
          <input
            type="text"
            className="search-input"
            placeholder="Search estimates by client, city, address, quotation #, amount, specs…"
            value={search}
            onChange={(e) => setSearch(e.target.value)}
          />
          {search && (
            <button className="icon-btn sm" onClick={() => setSearch('')}>
              {Icon.close}
            </button>
          )}
        </div>

        <div className="filter-controls">
          {/* Scope Filter Tabs: All, My Quotations, Shared with Me */}
          <div className="status-tabs scope-filter-tabs">
            <button
              className={`status-tab ${scopeFilter === 'all' ? 'active' : ''}`}
              onClick={() => setScopeFilter('all')}
              title="Show all accessible quotations"
            >
              All {isAdminOrOwner ? `(${projects.length})` : `(${myProjectsCount + sharedProjectsCount})`}
            </button>
            <button
              className={`status-tab ${scopeFilter === 'mine' ? 'active' : ''}`}
              onClick={() => setScopeFilter('mine')}
              title="Show quotations created by me"
            >
              My Quotations ({myProjectsCount})
            </button>
            <button
              className={`status-tab ${scopeFilter === 'shared' ? 'active' : ''}`}
              onClick={() => setScopeFilter('shared')}
              title="Show quotations shared with me by other staff"
            >
              🤝 Shared with Me {sharedProjectsCount > 0 ? `(${sharedProjectsCount})` : ''}
            </button>
          </div>

          <div className="status-tabs">
            {[
              { id: 'all', label: 'ACTIVE (QUOTED & NEGOTIATING)' },
              { id: 'quoted', label: 'QUOTED' },
              { id: 'negotiating', label: 'NEGOTIATING' },
              { id: 'won', label: 'WON' },
              { id: 'loss', label: 'LOSS' },
            ].map(({ id, label }) => (
              <button
                key={id}
                className={`status-tab ${statusFilter === id ? 'active' : ''}`}
                onClick={() => setStatusFilter(id)}
              >
                {label}
              </button>
            ))}
          </div>

          <label className="checkbox-label">
            <input
              type="checkbox"
              checked={showArchived}
              onChange={(e) => setShowArchived(e.target.checked)}
            />
            <span>Show Archived</span>
          </label>

          <div className="sort-group">
            <span className="tiny muted">Sort:</span>
            <select
              value={sortBy}
              onChange={(e) => setSortBy(e.target.value as SortField)}
              className="select-input"
            >
              <option value="updated_at">Recently Updated</option>
              <option value="name">Project Name</option>
              <option value="total">Quotation Amount</option>
              <option value="created_at">Date Created</option>
            </select>
            <button
              className="icon-btn sm"
              title={sortAsc ? 'Ascending' : 'Descending'}
              onClick={() => setSortAsc(!sortAsc)}
            >
              {sortAsc ? '↑' : '↓'}
            </button>
          </div>
        </div>
      </div>

      {/* Batch Actions Bar */}
      {selectedIds.size > 0 && (
        <div className="batch-actions-bar">
          <div className="batch-actions-left">
            <span className="batch-count-badge">
              {selectedIds.size} quotation{selectedIds.size > 1 ? 's' : ''} selected
            </span>
            <button
              type="button"
              className="btn ghost sm"
              onClick={() => setSelectedIds(new Set())}
            >
              Clear Selection
            </button>
          </div>
          <div className="batch-actions-right">
            {isAdminOrOwner && (
              <button
                type="button"
                className="btn danger sm"
                onClick={handleTriggerBatchDelete}
                style={{ display: 'flex', alignItems: 'center', gap: '6px' }}
              >
                {Icon.trash} Delete Selected ({selectedIds.size})
              </button>
            )}
          </div>
        </div>
      )}

      {/* Projects Table */}
      <div className="card table-card">
        {loading ? (
          <div className="page-loading">
            <Spinner size={28} />
            <p className="muted">Loading projects…</p>
          </div>
        ) : filteredProjects.length === 0 ? (
          <div className="empty">
            <p>
              {scopeFilter === 'shared'
                ? 'No quotations have been shared with you yet'
                : scopeFilter === 'mine'
                  ? 'You have not created any quotations yet'
                  : 'No projects match your current filters'}
            </p>
            {search || statusFilter !== 'all' || showArchived ? (
              <button
                className="btn ghost sm"
                onClick={() => {
                  setSearch('')
                  setStatusFilter('all')
                  setShowArchived(false)
                }}
              >
                Reset Filters
              </button>
            ) : (
              <button className="btn primary sm" onClick={onNewProject}>
                {Icon.plus} Create First Project
              </button>
            )}
          </div>
        ) : (
          <div className="table-scroll-wrap">
            <table className="tbl full-table">
            <thead>
              <tr>
                <th style={{ width: '40px', textAlign: 'center' }}>
                  <input
                    type="checkbox"
                    aria-label="Select all visible quotations"
                    checked={allVisibleSelected}
                    ref={(el) => {
                      if (el) el.indeterminate = someVisibleSelected && !allVisibleSelected
                    }}
                    onChange={handleToggleSelectAll}
                  />
                </th>
                <th>{viewMode === 'folders' ? 'Project Folder / Revision File' : 'Project Name'}</th>
                <th>Client</th>
                <th>Status</th>
                <th>Template Version</th>
                <th className="r">Quotation Total</th>
                <th>Owner / Creator</th>
                <th>Last Updated</th>
                <th className="r">Actions</th>
              </tr>
            </thead>
            <tbody>
              {viewMode === 'folders' ? (
                projectFolders.map((folder) => {
                  const isOpen = Boolean(openFolders[folder.id])
                  const latest = folder.latestRevision
                  const hasNewer = (org.current_template_version || 1) > latest.template_version
                  const folderAllSelected = folder.revisions.length > 0 && folder.revisions.every((r) => selectedIds.has(r.id))
                  const folderSomeSelected = folder.revisions.some((r) => selectedIds.has(r.id))

                  return (
                    <Fragment key={folder.id}>
                      {/* Project Folder Header Row (Click anywhere to open/close) */}
                      <tr
                        className={`folder-header-row ${folder.hasActive ? 'selected-row' : ''}`}
                        onClick={() => toggleFolder(folder.id)}
                        style={{ cursor: 'pointer' }}
                      >
                        <td style={{ width: '40px', textAlign: 'center' }} onClick={(e) => e.stopPropagation()}>
                          <input
                            type="checkbox"
                            aria-label={`Select folder ${folder.folderName}`}
                            checked={folderAllSelected}
                            ref={(el) => {
                              if (el) el.indeterminate = folderSomeSelected && !folderAllSelected
                            }}
                            onChange={(e) => handleToggleFolderSelect(folder, e.target.checked)}
                          />
                        </td>
                        <td>
                          <div className="row gap" style={{ flexWrap: 'wrap' }}>
                            <button
                              className="folder-toggle-btn"
                              onClick={(e) => {
                                e.stopPropagation()
                                toggleFolder(folder.id)
                              }}
                              title={isOpen ? 'Click to close folder' : 'Click to open folder'}
                            >
                              <span className="folder-icon">{Icon.folder}</span>
                              <strong>{folder.folderName}</strong>
                              <span style={{ color: 'var(--muted)', display: 'inline-flex' }}>
                                {isOpen ? Icon.chevronDown : Icon.chevronRight}
                              </span>
                            </button>
                            <span className="folder-count-badge">
                              {folder.totalCount} File{folder.totalCount > 1 ? 's' : ''} ({folder.revisions.map((r) => parseRevision(r.name).currentRev).join(', ')})
                            </span>
                            {folder.hasActive && <span className="chip ok">Active Folder</span>}
                            {folder.revisions.some((r) => {
                              const shared = (r.summary?.shared_with || (r as any).shared_with) as string[] | undefined
                              return (
                                Array.isArray(shared) &&
                                (shared.includes(session.username) || (session.user_id && shared.includes(session.user_id))) &&
                                r.created_by !== session.username
                              )
                            }) && (
                              <span className="chip purple sm" title="Contains quotations shared with you">Shared with you</span>
                            )}
                          </div>
                        </td>
                        <td>
                          {(() => {
                            const folderClient = latest.summary?.client_info?.buyer_name || folder.revisions.find((r) => r.summary?.client_info?.buyer_name)?.summary?.client_info?.buyer_name || ''
                            const qtnNo = latest.summary?.client_info?.qtn_no
                            if (!folderClient) return <span className="muted tiny">—</span>
                            return (
                              <div style={{ display: 'flex', flexDirection: 'column', gap: 2 }}>
                                <strong style={{ color: '#0f172a', fontSize: '12px' }}>{folderClient}</strong>
                                {qtnNo && <span className="tiny font-mono text-muted" style={{ fontSize: '10.5px' }}>{qtnNo}</span>}
                              </div>
                            )
                          })()}
                        </td>
                        <td>
                          {(() => {
                            const isWon = !latest.is_archived && (latest.status === 'won' || latest.status === 'approved')
                            const isLoss = latest.is_archived || latest.status === 'loss' || latest.status === 'archived' || latest.status === 'lost'
                            const isNeg = latest.status === 'negotiating'
                            const label = isWon ? 'Won' : isLoss ? 'Loss' : isNeg ? 'Negotiating' : 'Quoted'
                            const cls = isWon ? 'won' : isLoss ? 'loss' : isNeg ? 'negotiating' : 'quoted'
                            return <span className={`status-pill status-${cls}`}>{label}</span>
                          })()}
                          {latest.summary?.discount_request?.status === 'pending' && (
                            <span
                              className="chip warning sm"
                              style={{ marginLeft: '4px', fontSize: '11px', padding: '2px 6px' }}
                              title={`Staff @${latest.summary.discount_request.requested_by} requested ${latest.summary.discount_request.requested_pct}% discount`}
                            >
                              ⏳ {latest.summary.discount_request.requested_pct}% Disc Req
                            </span>
                          )}
                        </td>
                        <td>
                          <div className="template-badge-wrap">
                            <span className="version-pill">v{latest.template_version}</span>
                            {hasNewer && (
                              <span className="upgrade-hint" title={`Template v${org.current_template_version} available`}>
                                ↑ v{org.current_template_version}
                              </span>
                            )}
                          </div>
                        </td>
                        <td className="r font-bold text-green">
                          {latest.total ? inr(Number(latest.total)) : '—'}
                        </td>
                        <td className="muted tiny">
                          {latest.created_by || 'Unknown'}
                        </td>
                        <td className="muted tiny">{timeAgo(latest.updated_at)}</td>
                        <td className="r">
                          <div className="folder-actions" onClick={(e) => e.stopPropagation()}>
                            <button
                              className="btn primary sm folder-open-btn"
                              onClick={() => onOpenProject(latest.id)}
                              title={`Open latest revision (${folder.latestRevChar})`}
                            >
                              Open Latest ({folder.latestRevChar})
                            </button>
                            <button
                              className="btn ghost sm"
                              title={`Revise to Rev ${folder.nextRevChar}`}
                              onClick={() => {
                                setReviseProj(latest)
                                setRevName(folder.nextName)
                              }}
                            >
                              {Icon.refresh} Revise ({folder.nextRevChar})
                            </button>
                          </div>
                        </td>
                      </tr>

                      {/* Contained Revision Files (Only opened when folder is clicked!) */}
                      {isOpen && folder.revisions.map((file) => renderProjectRow(file, true))}
                    </Fragment>
                  )
                })
              ) : (
                filteredProjects.map((p) => renderProjectRow(p, false))
              )}
            </tbody>
          </table>
          </div>
        )}
      </div>

      {/* Revise Quotation Dialog */}
      {reviseProj && (
        <div className="overlay" onMouseDown={() => setReviseProj(null)}>
          <div className="dialog" role="dialog" onMouseDown={(e) => e.stopPropagation()}>
            <div className="dialog-header">
              {Icon.refresh}
              <div>
                <h3 style={{ margin: 0 }}>Create Revision {parseRevision(reviseProj.name).nextRev}</h3>
                <span className="tiny text-muted">From “{reviseProj.name}”</span>
              </div>
            </div>
            <p className="tiny muted mt-2" style={{ lineHeight: 1.5 }}>
              This will create a new quotation snapshot <strong>Rev {parseRevision(reviseProj.name).nextRev}</strong> with all current inputs, rates, profit margins, and discounts intact.
              The prior quotation (<strong>Rev {parseRevision(reviseProj.name).currentRev}</strong>) will remain preserved in your projects list.
            </p>
            <div className="form mt-3">
              <label>
                <span>Revision Quotation Name</span>
                <input
                  type="text"
                  autoFocus
                  className="modal-input"
                  value={revName}
                  onChange={(e) => setRevName(e.target.value)}
                  onKeyDown={(e) => e.key === 'Enter' && handleRevise()}
                />
              </label>
            </div>
            <div className="dialog-actions mt-4">
              <button className="btn ghost" onClick={() => setReviseProj(null)}>
                Cancel
              </button>
              <button className="btn primary" onClick={handleRevise} disabled={!revName.trim()}>
                {Icon.refresh} Create Rev {parseRevision(reviseProj.name).nextRev}
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Rename Dialog */}
      {renameProj && (
        <div className="overlay" onMouseDown={() => setRenameProj(null)}>
          <div className="dialog" role="dialog" onMouseDown={(e) => e.stopPropagation()}>
            <h3>Rename Project</h3>
            <p className="tiny muted">Enter a new title for this project</p>
            <input
              type="text"
              autoFocus
              className="modal-input"
              value={newName}
              onChange={(e) => setNewName(e.target.value)}
              onKeyDown={(e) => e.key === 'Enter' && handleRename()}
            />
            <div className="dialog-actions mt-4">
              <button className="btn ghost" onClick={() => setRenameProj(null)}>
                Cancel
              </button>
              <button className="btn primary" onClick={handleRename}>
                Save Name
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Duplicate Dialog */}
      {duplicateProj && (
        <div className="overlay" onMouseDown={() => setDuplicateProj(null)}>
          <div className="dialog" role="dialog" onMouseDown={(e) => e.stopPropagation()}>
            <h3>Duplicate Project</h3>
            <p className="tiny muted">
              Creates an independent copy of <strong>{duplicateProj.name}</strong>. Changes to the copy will never affect the original project.
            </p>
            <div className="form">
              <label>
                <span>New Project Name</span>
                <input
                  type="text"
                  autoFocus
                  value={dupName}
                  onChange={(e) => setDupName(e.target.value)}
                  onKeyDown={(e) => e.key === 'Enter' && handleDuplicate()}
                />
              </label>

              <div className="duplicate-template-choice">
                <span className="tiny font-bold text-muted">Template Version for Duplicate:</span>
                <label className="radio-label">
                  <input
                    type="radio"
                    name="dup_version"
                    checked={!dupUseLatest}
                    onChange={() => setDupUseLatest(false)}
                  />
                  <div>
                    <strong>Retain original Template (v{duplicateProj.template_version})</strong>
                    <span className="tiny muted block">Keep identical formula and rule definitions as the source</span>
                  </div>
                </label>
                <label className="radio-label">
                  <input
                    type="radio"
                    name="dup_version"
                    checked={dupUseLatest}
                    onChange={() => setDupUseLatest(true)}
                  />
                  <div>
                    <strong>Apply Organisation Default (v{org.current_template_version || 1})</strong>
                    <span className="tiny muted block">Snapshots the currently published default template rules</span>
                  </div>
                </label>
              </div>
            </div>

            <div className="dialog-actions mt-4">
              <button className="btn ghost" onClick={() => setDuplicateProj(null)}>
                Cancel
              </button>
              <button className="btn primary" onClick={handleDuplicate}>
                Create Duplicate
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Share Quotation Dialog */}
      {shareProj && (
        <div className="overlay" onMouseDown={() => setShareProj(null)}>
          <div className="dialog" role="dialog" onMouseDown={(e) => e.stopPropagation()} style={{ maxWidth: 440 }}>
            <div className="dialog-header">
              {Icon.users}
              <div>
                <h3 style={{ margin: 0 }}>Share Quotation</h3>
                <span className="tiny text-muted">“{shareProj.name}”</span>
              </div>
            </div>

            <div className="share-member-list mt-3">
              {loadingMembers ? (
                <div style={{ textAlign: 'center', padding: '16px 0' }}>
                  <Spinner size={20} />
                </div>
              ) : (shareMembers || []).filter((m) => m.username !== session.username && m.role !== 'owner' && m.role !== 'admin').length === 0 ? (
                <div className="empty" style={{ padding: '14px 0' }}>
                  <p className="tiny muted">No other staff members found.</p>
                </div>
              ) : (
                shareMembers
                  .filter((m) => m.username !== session.username && m.role !== 'owner' && m.role !== 'admin')
                  .map((m) => {
                    const currentShared = Array.isArray(shareProj.summary?.shared_with)
                      ? shareProj.summary.shared_with
                      : []
                    const isShared = currentShared.includes(m.username)
                    return (
                      <div
                        key={m.username}
                        className="row between p-2"
                        style={{ borderBottom: '1px solid var(--line)' }}
                      >
                        <div className="row gap">
                          <div className="avatar sm" style={{ background: '#1e3a8a', color: '#fff', fontSize: 11 }}>
                            {m.username[0].toUpperCase()}
                          </div>
                          <div>
                            <strong>@{m.username}</strong>
                            <div className="tiny muted">Staff</div>
                          </div>
                        </div>
                        <button
                          className={`btn sm ${isShared ? 'danger ghost' : 'primary'}`}
                          onClick={() => handleToggleShare(m.username)}
                        >
                          {isShared ? 'Revoke Access' : 'Share Access'}
                        </button>
                      </div>
                    )
                  })
              )}
            </div>

            <div className="dialog-actions mt-4">
              <button className="btn primary" onClick={() => setShareProj(null)}>
                Done
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Single Quotation Permanent Deletion Dialog (Requires Exact Name) */}
      {singleDeleteProj && (
        <div className="overlay" onMouseDown={() => !deletingSingle && setSingleDeleteProj(null)}>
          <div className="dialog danger-dialog" role="dialog" onMouseDown={(e) => e.stopPropagation()} style={{ maxWidth: 440 }}>
            <div className="dialog-header">
              <span style={{ color: '#dc2626', display: 'flex', alignItems: 'center' }}>{Icon.trash}</span>
              <div>
                <h3 style={{ margin: 0, color: '#dc2626' }}>Delete Quotation</h3>
                <span className="tiny text-muted">{singleDeleteProj.name}</span>
              </div>
            </div>
            <p className="tiny muted mt-2" style={{ lineHeight: 1.5 }}>
              This will permanently delete quotation <strong>“{singleDeleteProj.name}”</strong> and its calculation history. This action <strong>cannot be undone</strong>.
            </p>
            <div className="form mt-3">
              <label>
                <span style={{ fontSize: '12px' }}>
                  To confirm, type <strong>{singleDeleteProj.name}</strong> in the box below:
                </span>
                <div className="confirm-name-target">{singleDeleteProj.name}</div>
                <input
                  type="text"
                  autoFocus
                  className="modal-input"
                  placeholder="Type quotation name to confirm"
                  value={confirmNameInput}
                  onChange={(e) => setConfirmNameInput(e.target.value)}
                  onKeyDown={(e) => {
                    if (e.key === 'Enter' && confirmNameInput.trim() === singleDeleteProj.name.trim()) {
                      handleConfirmSingleDelete()
                    }
                  }}
                />
              </label>
            </div>
            <div className="dialog-actions mt-4">
              <button
                type="button"
                className="btn ghost sm"
                onClick={() => setSingleDeleteProj(null)}
                disabled={deletingSingle}
              >
                Cancel
              </button>
              <button
                type="button"
                className="btn danger sm"
                disabled={confirmNameInput.trim() !== singleDeleteProj.name.trim() || deletingSingle}
                onClick={handleConfirmSingleDelete}
                style={{ display: 'flex', alignItems: 'center', gap: '6px' }}
              >
                {deletingSingle ? <Spinner size={14} /> : Icon.trash} Delete Permanently
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Multiple Quotations Batch Deletion Dialog (Requires Account Password) */}
      {showBatchDeleteModal && (
        <div className="overlay" onMouseDown={() => !deletingBatch && setShowBatchDeleteModal(false)}>
          <div className="dialog danger-dialog" role="dialog" onMouseDown={(e) => e.stopPropagation()} style={{ maxWidth: 460 }}>
            <div className="dialog-header">
              <span style={{ color: '#dc2626', display: 'flex', alignItems: 'center' }}>{Icon.trash}</span>
              <div>
                <h3 style={{ margin: 0, color: '#dc2626' }}>Delete {selectedIds.size} Quotations</h3>
                <span className="tiny text-muted">Batch Permanent Deletion</span>
              </div>
            </div>
            <p className="tiny muted mt-2" style={{ lineHeight: 1.5 }}>
              You are about to permanently delete <strong>{selectedIds.size}</strong> quotations and their entire calculation histories. This action <strong>cannot be undone</strong>.
            </p>
            <div
              className="batch-delete-preview-box mt-2"
              style={{
                maxHeight: '120px',
                overflowY: 'auto',
                background: '#f8fafc',
                border: '1px solid #e2e8f0',
                borderRadius: '6px',
                padding: '8px 12px',
              }}
            >
              <ul style={{ margin: 0, paddingLeft: '18px', fontSize: '12px', color: '#475569' }}>
                {Array.from(selectedIds).map((id) => {
                  const p = projects.find((proj) => proj.id === id)
                  const buyer = p?.summary?.client_info?.buyer_name
                  return (
                    <li key={id} style={{ marginBottom: '2px' }}>
                      <strong>{p?.name || id}</strong>
                      {buyer ? <span className="tiny muted"> — {buyer}</span> : null}
                    </li>
                  )
                })}
              </ul>
            </div>
            <div className="form mt-3">
              <label>
                <span style={{ fontSize: '12px', fontWeight: 600 }}>Enter your account password to confirm deletion:</span>
                <input
                  type="password"
                  autoFocus
                  className="modal-input"
                  placeholder="Enter your account password"
                  value={batchDeletePassword}
                  onChange={(e) => {
                    setBatchDeletePassword(e.target.value)
                    if (batchDeleteError) setBatchDeleteError('')
                  }}
                  onKeyDown={(e) => {
                    if (e.key === 'Enter' && batchDeletePassword && !deletingBatch) {
                      handleConfirmBatchDelete()
                    }
                  }}
                />
              </label>
              {batchDeleteError && (
                <div className="tiny mt-1" style={{ color: '#dc2626', fontWeight: 600 }}>
                  {batchDeleteError}
                </div>
              )}
            </div>
            <div className="dialog-actions mt-4">
              <button
                type="button"
                className="btn ghost sm"
                onClick={() => setShowBatchDeleteModal(false)}
                disabled={deletingBatch}
              >
                Cancel
              </button>
              <button
                type="button"
                className="btn danger sm"
                disabled={!batchDeletePassword || deletingBatch}
                onClick={handleConfirmBatchDelete}
                style={{ display: 'flex', alignItems: 'center', gap: '6px' }}
              >
                {deletingBatch ? <Spinner size={14} /> : Icon.trash} Confirm &amp; Delete All ({selectedIds.size})
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  )
}
