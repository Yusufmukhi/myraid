import { useCallback, useEffect, useMemo, useRef, useState } from 'react'
import data from '../lib/clientWorkbook.json'
import c2cData from '../lib/clientWorkbook_c2c.json'
import { api, Member, Org, Project, Session, Template } from '../lib/api'
import { DEFAULT_NAMED_CONSTANTS, Overrides, SheetData, Workbook } from '../lib/engine'
import { inr, timeAgo } from '../lib/format'
import { cleanOverrides, extractTemplateDefaults, summarize, filterCutToCutOverrides } from '../lib/model'
import { formatInitialProjectName, parseRevision } from '../lib/revision'
import { getSavedClients } from '../lib/clientDirectory'
import { Navigation, NavTab } from './Navigation'
import { ActivityPage } from './pages/ActivityPage'
import { ApprovalsPage } from './pages/ApprovalsPage'
import { CalculationDetails } from './pages/CalculationDetails'
import { CalculatorPage } from './pages/CalculatorPage'
import { Dashboard } from './pages/Dashboard'
import { MembersPage } from './pages/MembersPage'
import { OrganisationSettings } from './pages/OrganisationSettings'
import { ProjectsPage } from './pages/ProjectsPage'
import { ClientsPage } from './pages/ClientsPage'
import { ClientRelationsPage } from './pages/ClientRelationsPage'
import { BusinessProfilePage } from './pages/BusinessProfilePage'
import { StaffPerformancePage } from './pages/StaffPerformancePage'
import { Logo } from './Screens'
import { Icon, Spinner, useCopy, useUI } from './ui'

// Single base workbook definition loaded into engine
const baseWb = new Workbook(data as unknown as { sheets: SheetData[] }, DEFAULT_NAMED_CONSTANTS)
const c2cWb = new Workbook(c2cData as unknown as { sheets: SheetData[] }, DEFAULT_NAMED_CONSTANTS)

interface Props {
  session: Session
  org: Org
  orgs: Org[]
  onSwitchOrg: (o: Org) => void
  onAddOrg: () => void
  onOrgsChanged: () => void
  onLogout: () => void
}

export function Workspace({ session, org, orgs, onSwitchOrg, onAddOrg, onOrgsChanged, onLogout }: Props) {
  const { toast, confirm } = useUI()
  const copy = useCopy()
  const isAdminOrOwner = org.role === 'owner' || org.role === 'admin'

  const [tab, setTab] = useState<NavTab>('dashboard')
  const [calculatorDirty, setCalculatorDirty] = useState(false)
  const calculatorSaveRef = useRef<(() => Promise<boolean>) | null>(null)
  const [showWorkspaceUnsavedModal, setShowWorkspaceUnsavedModal] = useState(false)
  const [pendingTab, setPendingTab] = useState<NavTab | null>(null)
  const [worksheetsLoaded, setWorksheetsLoaded] = useState(false)

  // Dynamically load full 15 worksheets into baseWb & c2cWb on mount (Owners/Admins only)
  useEffect(() => {
    if (!isAdminOrOwner) return
    let active = true
    api.getFullWorksheets(session.token, org.id)
      .then((res) => {
        if (!active) return
        if (res.actual_workbook?.sheets) {
          baseWb.loadSheets(res.actual_workbook.sheets, res.named_constants)
        }
        if (res.c2c_workbook?.sheets) {
          c2cWb.loadSheets(res.c2c_workbook.sheets, res.named_constants)
        }
        setWorksheetsLoaded(true)
      })
      .catch((err) => {
        console.warn('Could not load dynamic worksheets for workspace:', err)
      })
    return () => {
      active = false
    }
  }, [session.token, org.id, isAdminOrOwner])

  const handleTabSwitch = useCallback((newTab: NavTab) => {
    if (tab === 'calculator' && calculatorDirty && newTab !== 'calculator') {
      toast('Please save your changes first', 'err')
      setPendingTab(newTab)
      setShowWorkspaceUnsavedModal(true)
      return
    }
    setTab(newTab)
    setNavOpen(false)
  }, [tab, calculatorDirty, toast])

  const [projects, setProjects] = useState<Project[]>([])
  const [currentProject, setCurrentProject] = useState<Project | null>(null)
  const [loading, setLoading] = useState(true)
  const [navOpen, setNavOpen] = useState(false)
  const [orgMenu, setOrgMenu] = useState(false)
  const [members, setMembers] = useState<Member[] | null>(null)

  // Profile Modal state
  const [showProfileModal, setShowProfileModal] = useState(Boolean(session.must_change_password))
  const [profileUsername, setProfileUsername] = useState(session.username)
  const [currentPassword, setCurrentPassword] = useState('')
  const [newPassword, setNewPassword] = useState('')
  const [confirmPassword, setConfirmPassword] = useState('')
  const [updatingProfile, setUpdatingProfile] = useState(false)

  // Create Project Modal state
  const [showCreateModal, setShowCreateModal] = useState(false)
  const [newProjectName, setNewProjectName] = useState('QTY NO ')
  const [availableTemplates, setAvailableTemplates] = useState<Template[]>([])
  const [selectedTemplateId, setSelectedTemplateId] = useState<string>('')
  const [selectedClientForNewProj, setSelectedClientForNewProj] = useState<string>('')
  const [clientSearchTermForNewProj, setClientSearchTermForNewProj] = useState<string>('')
  const [clientDropdownOpen, setClientDropdownOpen] = useState(false)
  const [creating, setCreating] = useState(false)
  const savedClients = useMemo(() => getSavedClients(org.id), [org.id, showCreateModal])

  const filteredClientsForNewProj = useMemo(() => {
    const q = clientSearchTermForNewProj.trim().toLowerCase()
    if (!q) return savedClients || []
    return (savedClients || []).filter(
      (c) =>
        c.name.toLowerCase().includes(q) ||
        (c.gst || '').toLowerCase().includes(q) ||
        (c.kind_att || '').toLowerCase().includes(q) ||
        (c.contact || '').toLowerCase().includes(q)
    )
  }, [savedClients, clientSearchTermForNewProj])

  // Protect restricted tabs from non-admins
  useEffect(() => {
    if (!isAdminOrOwner && ['details', 'settings', 'members', 'activity', 'approvals', 'business', 'staff_analytics'].includes(tab)) {
      setTab('dashboard')
    }
  }, [isAdminOrOwner, tab])

  const openCreateModal = async () => {
    setShowCreateModal(true)
    setNewProjectName('QTY NO ')
    setSelectedTemplateId('')
    setSelectedClientForNewProj('')
    setClientSearchTermForNewProj('')
    setClientDropdownOpen(false)
    try {
      const list = await api.templateList(session.token, org.id)
      setAvailableTemplates(list)
    } catch {
      setAvailableTemplates([])
    }
  }

  // Load all projects for current organisation
  const loadProjects = useCallback(async (selectId?: string) => {
    try {
      setLoading(true)
      const res = await api.projectList(session.token, org.id, true)
      const list = Array.isArray(res) ? res : []
      setProjects(list)

      if (selectId) {
        const found = list.find((p) => p.id === selectId)
        if (found) setCurrentProject(found)
      } else if (!currentProject) {
        const userProjects = !isAdminOrOwner
          ? list.filter((p) => {
              const shared = (p.summary?.shared_with || (p as any).shared_with) as string[] | undefined
              const isShared = Array.isArray(shared) && (shared.includes(session.username) || (session.user_id && shared.includes(session.user_id)))
              const isMine = p.created_by === session.user_id || p.created_by === session.username
              if (isMine) return true
              if (isShared) return p.status?.toLowerCase() !== 'approved'
              return false
            })
          : list
        if (userProjects.length > 0) {
          const active = userProjects.find((p) => !p.is_archived) || userProjects[0]
          setCurrentProject(active)
        }
      }
    } catch (e) {
      toast((e as Error).message, 'err')
    } finally {
      setLoading(false)
    }
  }, [session.token, session.user_id, session.username, org.id, currentProject, isAdminOrOwner, toast])

  useEffect(() => {
    loadProjects()
  }, [org.id])

  const loadMembers = async () => {
    setMembers(null)
    try {
      setMembers(await api.orgMembers(session.token, org.id))
    } catch {
      setMembers([])
    }
  }

  // Open an existing project and switch to calculator
  const handleOpenProject = async (projectId: string) => {
    try {
      const p = await api.projectGet(session.token, projectId)
      setCurrentProject(p)
      setTab('calculator')
      setNavOpen(false)
    } catch (e) {
      toast((e as Error).message, 'err')
    }
  }

  // Create project flow
  const handleCreateProject = async () => {
    const raw = newProjectName.trim()
    if (!raw) return toast('Please enter a project name', 'err')
    const nm = formatInitialProjectName(raw)
    setCreating(true)
    try {
      let initialSummary: any = {}
      const initialOverrides: Record<string, any> = {}
      const targetClientId = selectedClientForNewProj || savedClients.find((c) => c.name.trim().toLowerCase() === clientSearchTermForNewProj.trim().toLowerCase())?.id
      if (targetClientId) {
        const cl = savedClients.find((c) => c.id === targetClientId)
        if (cl) {
          const clientMargin = cl.profit_margin_pct !== undefined ? cl.profit_margin_pct : 25
          initialSummary = {
            client_id: cl.id,
            profit_pct: clientMargin,
            client_info: {
              buyer_name: cl.name,
              buyer_address: cl.address,
              buyer_gst: cl.gst || '',
              buyer_email: cl.email || '',
              kind_att: cl.kind_att || '',
              buyer_contact: cl.contact || '',
              contacts: cl.contacts || [],
              extra_fields: cl.extra_fields || [],
            },
          }
          initialOverrides['__profit_pct'] = clientMargin
        }
      } else if (clientSearchTermForNewProj.trim()) {
        const matched = savedClients.find((c) => c.name.trim().toLowerCase() === clientSearchTermForNewProj.trim().toLowerCase())
        const clientMargin = matched && matched.profit_margin_pct !== undefined ? matched.profit_margin_pct : 25
        initialSummary = {
          ...(matched ? { client_id: matched.id, profit_pct: clientMargin } : {}),
          client_info: {
            buyer_name: clientSearchTermForNewProj.trim(),
          },
        }
        if (matched && matched.profit_margin_pct !== undefined) {
          initialOverrides['__profit_pct'] = clientMargin
        }
      }
      const res = await api.projectCreate(
        session.token,
        org.id,
        nm,
        selectedTemplateId || null,
        initialOverrides,
        {},
        null,
        initialSummary,
      )
      toast(`Created quotation “${res.name}”`)
      setShowCreateModal(false)
      setNewProjectName('QTY NO ')
      setSelectedTemplateId('')
      setSelectedClientForNewProj('')
      const fullProj = await api.projectGet(session.token, res.id)
      await loadProjects(res.id)
      setCurrentProject(fullProj)
      setTab('calculator')
    } catch (e) {
      toast((e as Error).message, 'err')
    } finally {
      setCreating(false)
    }
  }

  // Synchronize template defaults to baseWb whenever currentProject changes
  useEffect(() => {
    if (!currentProject) return
    let active = true
    const applyTpl = async () => {
      let defs: Record<string, number | null> = {}
      if (currentProject.template_snapshot?.definition) {
        defs = extractTemplateDefaults(currentProject.template_snapshot.definition)
      } else {
        try {
          const list = await api.templateList(session.token, org.id)
          const matching = list.find((t) => t.id === currentProject.template_id || t.version_num === currentProject.template_version)
          if (matching?.definition) {
            defs = extractTemplateDefaults(matching.definition)
          }
        } catch {
          // ignore
        }
      }
      if (active) {
        baseWb.setTemplateDefaults(defs)
        baseWb.setFormulaOverrides(currentProject.formula_overrides || {})
        baseWb.setOverrides(currentProject.inputs || {})
        c2cWb.setTemplateDefaults(filterCutToCutOverrides(defs))
        c2cWb.setFormulaOverrides({ ...(currentProject.formula_overrides || {}), ...(currentProject.summary?.c2c_formula_overrides || {}) })
        c2cWb.setOverrides({ ...filterCutToCutOverrides(currentProject.inputs || {}), ...(currentProject.summary?.c2c_overrides || {}) })
      }
    }
    applyTpl()
    return () => {
      active = false
    }
  }, [currentProject?.id, currentProject?.template_id, currentProject?.template_version, currentProject?.template_snapshot, org.id, session.token])

  // Save copy / revision flow from calculator
  const handleSaveAsCopy = async (copyName: string) => {
    if (!currentProject) return
    try {
      const res = await api.projectDuplicate(session.token, currentProject.id, copyName, false)
      // When revised from calculator, automatically set status to 'negotiating'
      try {
        await api.projectUpdate(
          session.token,
          res.id,
          res.name,
          'negotiating',
          currentProject.inputs,
          currentProject.formula_overrides,
          currentProject.total,
          currentProject.summary,
          `Revision ${res.name} created - Status set to Negotiating`,
        )
      } catch (err) {
        // Safe fallback if database check constraint is pending migration
        await api.projectUpdate(
          session.token,
          res.id,
          res.name,
          'quoted',
          currentProject.inputs,
          currentProject.formula_overrides,
          currentProject.total,
          currentProject.summary,
          `Revision ${res.name} created`,
        )
      }
      toast(`Quotation revision “${res.name}” created (Status: Negotiating)!`)
      await loadProjects(res.id)
      const newProj = await api.projectGet(session.token, res.id)
      setCurrentProject(newProj)
      setTab('calculator')
    } catch (e) {
      toast((e as Error).message, 'err')
    }
  }

  // Profile update handler
  const handleUpdateProfile = async (e: React.FormEvent) => {
    e.preventDefault()
    const trimmedUser = profileUsername.trim()
    if (!trimmedUser || trimmedUser.length < 3) {
      return toast('Username must be at least 3 characters', 'err')
    }
    if (newPassword && newPassword.length < 10) {
      return toast('New password must be at least 10 characters long', 'err')
    }
    if (session.must_change_password && !newPassword) {
      return toast('You must set a new password of at least 10 characters', 'err')
    }
    if (newPassword && !session.must_change_password && !currentPassword) {
      return toast('Current password is required to change your password', 'err')
    }
    if (newPassword && newPassword !== confirmPassword) {
      return toast('Passwords do not match', 'err')
    }

    setUpdatingProfile(true)
    try {
      const res = await api.updateProfile(
        session.token,
        trimmedUser !== session.username ? trimmedUser : undefined,
        newPassword || undefined,
        undefined,
        currentPassword || undefined,
      )
      if (res && res.username) {
        session.username = res.username
      }
      session.must_change_password = false
      toast('Security credentials updated successfully!')
      setShowProfileModal(false)
      setCurrentPassword('')
      setNewPassword('')
      setConfirmPassword('')
    } catch (err) {
      toast((err as Error).message, 'err')
    } finally {
      setUpdatingProfile(false)
    }
  }

  // Projects created by the current user (only ones created by user, excluding shared)
  const myCreatedProjects = useMemo(() => {
    return (projects || []).filter((p) => {
      if (p.is_archived) return false
      if (isAdminOrOwner) return true
      return p.created_by === session.user_id || p.created_by === session.username
    })
  }, [projects, isAdminOrOwner, session.user_id, session.username])

  // Projects shared with the current user by other staff or owner
  const mySharedProjects = useMemo(() => {
    if (isAdminOrOwner) return []
    return (projects || []).filter((p) => {
      if (p.is_archived) return false
      // If approved, automatically remove from the shared side
      if (p.status?.toLowerCase() === 'approved') return false
      const shared = (p.summary?.shared_with || (p as any).shared_with) as string[] | undefined
      const isShared = Array.isArray(shared) && (shared.includes(session.username) || (session.user_id && shared.includes(session.user_id)))
      const isCreatedByOther = p.created_by !== session.username && p.created_by !== session.user_id
      return isShared && isCreatedByOther
    })
  }, [projects, isAdminOrOwner, session.user_id, session.username])

  // Filter projects visible to current user (including shared projects)
  const myVisibleProjects = useMemo(() => {
    return (projects || []).filter((p) => {
      if (p.is_archived) return false
      if (isAdminOrOwner) return true
      const shared = (p.summary?.shared_with || (p as any).shared_with) as string[] | undefined
      const isShared = Array.isArray(shared) && (shared.includes(session.username) || (session.user_id && shared.includes(session.user_id)))
      const isMine = p.created_by === session.user_id || p.created_by === session.username
      if (isMine) return true
      if (isShared) {
        // If approved, automatically remove from the shared side
        return p.status?.toLowerCase() !== 'approved'
      }
      return false
    })
  }, [projects, isAdminOrOwner, session.user_id, session.username])

  const pendingDiscountCount = useMemo(() => {
    return (projects || []).filter((p) => p.summary?.discount_request?.status === 'pending').length
  }, [projects])

  const totalPendingApprovals = pendingDiscountCount

  const [hasUnsavedChanges, setHasUnsavedChanges] = useState(false)
  const [isSavingProject, setIsSavingProject] = useState(false)

  // Centralized project save that first calls server calculate and saves its summary and total
  const saveProjectToServer = useCallback(
    async (proj = currentProject, showToast = false) => {
      if (!proj || !session?.token) return

      // Read required quotation settings from project summary or inputs (supporting snake_case, camelCase & defaults)
      const summary = proj.summary || {}

      const incSetting =
        summary.include_installation !== undefined
          ? Boolean(summary.include_installation)
          : summary.includeInstallation !== undefined
          ? Boolean(summary.includeInstallation)
          : proj.inputs?.['__include_installation'] !== undefined
          ? Number(proj.inputs['__include_installation']) === 1
          : true

      const extraSetting =
        summary.extra_amount !== undefined
          ? Number(summary.extra_amount)
          : summary.extraAmount !== undefined
          ? Number(summary.extraAmount)
          : proj.inputs?.['__extra_amount'] !== undefined
          ? Number(proj.inputs['__extra_amount'])
          : 0

      const profitSetting =
        summary.profit_pct !== undefined
          ? Number(summary.profit_pct)
          : summary.profitPct !== undefined
          ? Number(summary.profitPct)
          : proj.inputs?.['Total cal!C26'] !== undefined
          ? Number(proj.inputs['Total cal!C26'])
          : proj.inputs?.['__profit_pct'] !== undefined
          ? Number(proj.inputs['__profit_pct'])
          : 25

      const discountSetting =
        summary.discount_pct !== undefined
          ? Number(summary.discount_pct)
          : summary.discountPct !== undefined
          ? Number(summary.discountPct)
          : proj.inputs?.['__discount_pct'] !== undefined
          ? Number(proj.inputs['__discount_pct'])
          : 0

      const cleanSystems = Array.isArray(summary.systems) && summary.systems.length > 0
        ? summary.systems.map((sys: any, idx: number) => ({
            name: String(sys.name || `SYSTEM ${String.fromCharCode(65 + idx)}`),
            quantity: Math.max(1, Number(sys.quantity) || 1),
            unit: String(sys.unit || 'SET'),
            overrides: (sys.overrides && Object.keys(sys.overrides).length > 0) ? sys.overrides : (idx === 0 ? (proj.inputs || {}) : {}),
            formulaOverrides: (sys.formulaOverrides && Object.keys(sys.formulaOverrides).length > 0) ? sys.formulaOverrides : (idx === 0 ? (proj.formula_overrides || {}) : {}),
          }))
        : undefined

      setIsSavingProject(true)
      try {
        const calcRes = await api.serverCalculate(
          session.token,
          {
            org_id: org.id,
            inputs: cleanOverrides(baseWb, proj.inputs || {}),
            formula_overrides: proj.formula_overrides,
            profit_pct: profitSetting,
            discount_pct: discountSetting,
            include_installation: incSetting,
            extra_amount: extraSetting,
            systems: cleanSystems,
            c2c_overrides: summary.c2c_overrides || undefined,
            c2c_formula_overrides: summary.c2c_formula_overrides || undefined,
          },
          org.id,
        )

        if (!calcRes || !calcRes.summary) {
          toast('Server calculation failed, not saved: empty calculation result', 'err')
          return
        }

        const serverSummary = {
          ...summary,
          ...calcRes.summary,
          c2c_overrides: summary.c2c_overrides || undefined,
          c2c_formula_overrides: summary.c2c_formula_overrides || undefined,
          c2cSummary: calcRes.c2c_summary || summary.c2cSummary,
          c2c_summary: calcRes.c2c_summary || summary.c2c_summary,
          c2cAllSystems: calcRes.c2c_all_systems || summary.c2cAllSystems,
          c2c_all_systems: calcRes.c2c_all_systems || summary.c2c_all_systems,
          c2cProfit: calcRes.c2c_profit !== undefined ? calcRes.c2c_profit : summary.c2cProfit,
          c2c_profit: calcRes.c2c_profit !== undefined ? calcRes.c2c_profit : summary.c2c_profit,
          c2cTotal: calcRes.c2c_total !== undefined ? calcRes.c2c_total : summary.c2cTotal,
          c2c_total: calcRes.c2c_total !== undefined ? calcRes.c2c_total : summary.c2c_total,
          c2cRatio: calcRes.c2c_ratio !== undefined ? calcRes.c2c_ratio : summary.c2cRatio,
          c2c_ratio: calcRes.c2c_ratio !== undefined ? calcRes.c2c_ratio : summary.c2c_ratio,
          include_installation: incSetting,
          includeInstallation: incSetting,
          extra_amount: extraSetting,
          extraAmount: extraSetting,
          profit_pct: profitSetting,
          profitPct: profitSetting,
          discount_pct: discountSetting,
          discountPct: discountSetting,
          systems: summary.systems || undefined,
          calculated_on_server: true,
        }
        const serverTotal = calcRes.computed_total ?? calcRes.summary.sales

        await api.projectUpdate(
          session.token,
          proj.id,
          proj.name,
          proj.status,
          proj.inputs,
          proj.formula_overrides,
          serverTotal,
          serverSummary,
        )

        setCurrentProject((prev) => (prev ? { ...prev, total: serverTotal, summary: serverSummary } : null))
        setHasUnsavedChanges(false)
        if (showToast) toast('Saved & synced with server')
      } catch (calcErr: any) {
        console.error('Server calculation failed:', calcErr)
        toast(`Server calculation failed: ${calcErr?.message || 'not saved'}`, 'err')
        return
      } finally {
        setIsSavingProject(false)
      }
    },
    [currentProject, session?.token, org.id, baseWb, toast],
  )

  // 3-second idle auto-save
  useEffect(() => {
    if (!hasUnsavedChanges || !currentProject || !session?.token) return
    const timer = setTimeout(() => {
      saveProjectToServer(currentProject, false)
    }, 3000)
    return () => clearTimeout(timer)
  }, [hasUnsavedChanges, currentProject, saveProjectToServer, session?.token])

  // Formula override handler for CalculationDetails page (screen state only, no immediate API call)
  const handleSetFormulaOverride = (key: string, formula: string | undefined, targetWorkbook: 'actual' | 'c2c' = 'actual') => {
    if (!currentProject) return

    if (targetWorkbook === 'c2c') {
      const currentC2CFormulas = { ...(currentProject.summary?.c2c_formula_overrides || {}) }
      if (formula === undefined) {
        delete currentC2CFormulas[key]
      } else {
        currentC2CFormulas[key] = formula
      }
      const effectiveFormulas = { ...(currentProject.formula_overrides || {}), ...currentC2CFormulas }
      c2cWb.setFormulaOverrides(effectiveFormulas)
      const newC2CSummary = summarize(c2cWb)

      const updated: Project = {
        ...currentProject,
        summary: {
          ...(currentProject.summary || {}),
          c2c_formula_overrides: currentC2CFormulas,
          c2c_summary: newC2CSummary,
          c2cSummary: newC2CSummary,
        },
      }
      setCurrentProject(updated)
      setHasUnsavedChanges(true)
      return
    }

    const nextOverrides = { ...currentProject.formula_overrides }
    if (formula === undefined) {
      delete nextOverrides[key]
    } else {
      nextOverrides[key] = formula
    }
    baseWb.setFormulaOverrides(nextOverrides)
    c2cWb.setFormulaOverrides({ ...nextOverrides, ...(currentProject.summary?.c2c_formula_overrides || {}) })
    const newSummary = summarize(baseWb)
    const newC2CSummary = summarize(c2cWb)

    const updated: Project = {
      ...currentProject,
      formula_overrides: nextOverrides,
      total: newSummary.sales,
      summary: {
        ...(currentProject.summary || {}),
        ...newSummary,
        c2c_summary: newC2CSummary,
        c2cSummary: newC2CSummary,
      },
    }
    setCurrentProject(updated)
    setHasUnsavedChanges(true)
  }

  // Input override handler for CalculationDetails page (screen state only, no immediate API call)
  const handleSetInput = (key: string, value: number | null | undefined, targetWorkbook: 'actual' | 'c2c' = 'actual') => {
    if (!currentProject) return

    if (targetWorkbook === 'c2c') {
      const currentC2C = { ...(currentProject.summary?.c2c_overrides || {}) }
      if (value === undefined) {
        delete currentC2C[key]
      } else {
        currentC2C[key] = value
      }

      const effectiveC2C = { ...filterCutToCutOverrides(currentProject.inputs || {}), ...currentC2C }
      c2cWb.setOverrides(effectiveC2C)
      const newC2CSummary = summarize(c2cWb)

      const updated: Project = {
        ...currentProject,
        summary: {
          ...(currentProject.summary || {}),
          c2c_overrides: currentC2C,
          c2c_summary: newC2CSummary,
          c2cSummary: newC2CSummary,
        },
      }
      setCurrentProject(updated)
      setHasUnsavedChanges(true)
      return
    }

    const nextInputs = { ...currentProject.inputs }
    if (value === undefined) {
      delete nextInputs[key]
    } else {
      nextInputs[key] = value
    }
    const cleaned = cleanOverrides(baseWb, nextInputs)
    baseWb.setOverrides(cleaned)
    const effectiveC2C = { ...filterCutToCutOverrides(cleaned), ...(currentProject.summary?.c2c_overrides || {}) }
    c2cWb.setOverrides(effectiveC2C)
    const newSummary = summarize(baseWb)
    const newC2CSummary = summarize(c2cWb)

    const updated: Project = {
      ...currentProject,
      inputs: cleaned,
      total: newSummary.sales,
      summary: {
        ...(currentProject.summary || {}),
        ...newSummary,
        c2c_summary: newC2CSummary,
        c2cSummary: newC2CSummary,
      },
    }
    setCurrentProject(updated)
    setHasUnsavedChanges(true)
  }

  return (
    <div className="app-shell">
      {/* ---------------------------------------------------- Global Top Bar */}
      <header className="topbar">
        <button
          className="icon-btn only-mobile"
          onClick={() => setNavOpen((v) => !v)}
          aria-label="Toggle navigation"
        >
          {Icon.menu}
        </button>
        <Logo light />

        <div className="topbar-current-tab">
          <span className="topbar-tab-title">
            {tab === 'dashboard' && 'Dashboard Overview'}
            {tab === 'projects' && 'Projects & Quotations'}
            {tab === 'clients' && 'Client Directory'}
            {tab === 'client_relations' && 'Client & Staff Relations'}
            {tab === 'business' && 'Business Profile & PDF Branding'}
            {tab === 'staff_analytics' && 'Staff Quotation & Performance Analytics'}
            {tab === 'calculator' && (currentProject ? `Calculator — ${currentProject.name}` : 'Calculator')}
            {tab === 'details' && (currentProject ? `Calculation Details — ${currentProject.name}` : 'Calculation Details')}
            {tab === 'approvals' && 'Requests & Approvals Hub'}
            {tab === 'settings' && 'Organisation Settings & Versioning'}
            {tab === 'members' && 'Team Members & Permissions'}
            {tab === 'activity' && 'Audit Trail & Event History'}
          </span>
        </div>

        {currentProject && (
          <div className="row gap" style={{ alignItems: 'center', marginLeft: 16 }}>
            {hasUnsavedChanges ? (
              <button
                type="button"
                className="btn primary sm"
                onClick={() => saveProjectToServer(currentProject, true)}
                disabled={isSavingProject}
                style={{ height: 28, fontSize: 12, display: 'inline-flex', alignItems: 'center', gap: 6, background: '#2563eb' }}
                title="Save changes to server (or wait 3s for idle auto-save)"
              >
                {isSavingProject ? <Spinner size={12} /> : '💾'}
                {isSavingProject ? 'Saving…' : 'Save'}
              </button>
            ) : (
              <span className="tiny font-mono" style={{ color: '#10b981', background: '#ecfdf5', padding: '2px 8px', borderRadius: 4, border: '1px solid #a7f3d0', fontSize: 11 }}>
                ✓ Saved
              </span>
            )}
          </div>
        )}

        <div className="grow" />

        {/* User Profile Quick Access */}
        <button
          className="btn ghost sm"
          onClick={() => {
            setProfileUsername(session.username)
            setNewPassword('')
            setConfirmPassword('')
            setShowProfileModal(true)
          }}
          title="My Profile & Password"
          style={{ gap: 7, height: 32, padding: '0 10px', marginRight: 4 }}
        >
          <div className="avatar sm" style={{ width: 22, height: 22, fontSize: 11, background: '#1e3a8a', color: '#fff' }}>
            {(session?.username?.[0] || 'U').toUpperCase()}
          </div>
          <span style={{ fontSize: 13, fontWeight: 600 }}>{session.username || 'User'}</span>
        </button>

        {/* Organisation Switcher Dropdown */}
        <div className="org-menu">
          <button
            className="org-btn"
            onClick={() => {
              setOrgMenu((v) => !v)
              if (!orgMenu) loadMembers()
            }}
          >
            {Icon.building}
            <span className="org-name">{org.name}</span>
            <span className="org-role-tag">{org.role === 'floor_operator' ? 'MEMBER' : org.role}</span>
            {Icon.chevron}
          </button>
          {orgMenu && (
            <>
              <div className="menu-backdrop" onClick={() => setOrgMenu(false)} />
              <div className="menu card">
                <div className="menu-sec">
                  <span className="label">Switch Organisation</span>
                  {orgs.map((o) => (
                    <button
                      key={o.id}
                      className={`menu-item ${o.id === org.id ? 'font-bold text-blue' : ''}`}
                      onClick={() => {
                        onSwitchOrg(o)
                        setOrgMenu(false)
                      }}
                    >
                      {Icon.building}
                      <span className="grow">{o.name}</span>
                      {o.id === org.id && <span className="chip ok sm">Active</span>}
                    </button>
                  ))}
                  {isAdminOrOwner && (
                    <button
                      className="menu-item"
                      onClick={() => {
                        onAddOrg()
                        setOrgMenu(false)
                      }}
                    >
                      {Icon.plus} Create another organisation…
                    </button>
                  )}
                </div>

                <div className="menu-sec">
                  <span className="label">Account</span>
                  <button
                    className="menu-item"
                    onClick={() => {
                      setProfileUsername(session?.username || '')
                      setNewPassword('')
                      setConfirmPassword('')
                      setShowProfileModal(true)
                      setOrgMenu(false)
                    }}
                    style={{ width: '100%', textAlign: 'left', display: 'flex', alignItems: 'center', gap: 10 }}
                  >
                    <div className="avatar sm">{(session?.username?.[0] || 'U').toUpperCase()}</div>
                    <div style={{ display: 'flex', flexDirection: 'column' }}>
                      <span style={{ fontWeight: 700 }}>{session.username || 'User'}</span>
                      <span className="tiny text-blue" style={{ fontSize: 11 }}>Edit Profile &amp; Password</span>
                    </div>
                  </button>
                  <button className="menu-item danger" onClick={() => onLogout()}>
                    {Icon.logout} Sign out
                  </button>
                </div>
              </div>
            </>
          )}
        </div>
      </header>

      {/* ---------------------------------------------------- App Workspace Body */}
      <div className="body">
        {/* Navigation Sidebar */}
        <div className={`sidebar-container ${navOpen ? 'open' : ''}`}>
          <Navigation
            currentTab={tab}
            onSelectTab={handleTabSwitch}
            org={org}
            currentProjectName={currentProject ? currentProject.name : null}
            projectCount={isAdminOrOwner ? myVisibleProjects.length : myCreatedProjects.length}
            sharedCount={mySharedProjects.length}
            clientCount={savedClients.length}
            hasUnsavedChanges={hasUnsavedChanges || calculatorDirty}
            pendingApprovalsCount={totalPendingApprovals}
            onOpenCreateModal={openCreateModal}
          />
        </div>

        {/* Mobile Scrim */}
        {navOpen && <div className="scrim" onClick={() => setNavOpen(false)} />}

        {/* Main Content Area */}
        <main className="main-content">
          {tab === 'dashboard' && (
            <Dashboard
              session={session}
              org={org}
              onOpenProject={handleOpenProject}
              onNewProject={openCreateModal}
              onNavigate={(t) => handleTabSwitch(t as NavTab)}
            />
          )}

          {(tab === 'projects' || tab === 'shared_projects') && (
            <ProjectsPage
              session={session}
              org={org}
              currentProjectId={currentProject?.id ?? null}
              onOpenProject={handleOpenProject}
              onNewProject={openCreateModal}
              onProjectsChanged={() => loadProjects()}
              initialScope={tab === 'shared_projects' ? 'shared' : 'all'}
            />
          )}

          {tab === 'clients' && (
            <ClientsPage
              session={session}
              org={org}
              onCreateQuotationWithClient={(client) => {
                setSelectedClientForNewProj(client.id)
                setClientSearchTermForNewProj(client.name)
                setClientDropdownOpen(false)
                setNewProjectName('QTY NO ')
                setShowCreateModal(true)
              }}
              onNavigate={(t) => setTab(t as NavTab)}
            />
          )}

          {tab === 'client_relations' && (
            <ClientRelationsPage
              session={session}
              org={org}
              onOpenProject={handleOpenProject}
              onCreateQuotationWithClient={(client) => {
                setSelectedClientForNewProj(client.id)
                setClientSearchTermForNewProj(client.name)
                setClientDropdownOpen(false)
                setNewProjectName('QTY NO ')
                setShowCreateModal(true)
              }}
              onNavigate={(t) => setTab(t as NavTab)}
            />
          )}

          {tab === 'calculator' && (
            currentProject ? (
              <CalculatorPage
                key={`${currentProject.id}-${worksheetsLoaded ? 'full' : 'basic'}`}
                session={session}
                org={org}
                project={currentProject}
                wb={baseWb}
                c2cWb={c2cWb}
                onProjectUpdated={(p) => {
                  setCurrentProject(p)
                  loadProjects()
                }}
                onSaveAsCopy={handleSaveAsCopy}
                onNavigateToSheets={() => handleTabSwitch('details')}
                onDirtyChange={(d) => setCalculatorDirty(d)}
                saveTriggerRef={calculatorSaveRef}
              />
            ) : (
              <div style={{ padding: '60px 20px', display: 'flex', justifyContent: 'center' }}>
                <div
                  className="card"
                  style={{
                    maxWidth: '460px',
                    width: '100%',
                    padding: '36px 28px',
                    textAlign: 'center',
                    display: 'flex',
                    flexDirection: 'column',
                    alignItems: 'center',
                    gap: '12px',
                    background: '#ffffff',
                    borderRadius: '8px',
                    boxShadow: '0 1px 3px rgba(0,0,0,0.1)',
                  }}
                >
                  <div
                    style={{
                      width: '56px',
                      height: '56px',
                      borderRadius: '50%',
                      background: '#f1f5f9',
                      display: 'flex',
                      alignItems: 'center',
                      justifyContent: 'center',
                      color: '#475569',
                      fontSize: '24px',
                    }}
                  >
                    {Icon.calc}
                  </div>
                  <h3 style={{ margin: 0, fontSize: '18px', fontWeight: 700, color: '#0f172a' }}>
                    No Quotation Project Selected
                  </h3>
                  <p style={{ margin: 0, fontSize: '13px', color: '#64748b', lineHeight: 1.5 }}>
                    You haven't created or opened any quotation projects yet. Click below to create your project and start calculating.
                  </p>
                  <button
                    type="button"
                    className="btn primary"
                    style={{ marginTop: '8px', padding: '10px 20px', fontSize: '13px', fontWeight: 600, display: 'inline-flex', alignItems: 'center', gap: '8px' }}
                    onClick={openCreateModal}
                  >
                    {Icon.plus} Create a Project
                  </button>
                </div>
              </div>
            )
          )}

          {tab === 'business' && isAdminOrOwner && (
            <BusinessProfilePage
              session={session}
              org={org}
            />
          )}

          {tab === 'staff_analytics' && isAdminOrOwner && (
            <StaffPerformancePage
              session={session}
              org={org}
              onOpenProject={handleOpenProject}
              onNavigate={(t) => handleTabSwitch(t as NavTab)}
            />
          )}

          {tab === 'details' && isAdminOrOwner && (
            currentProject ? (
              <CalculationDetails
                wb={baseWb}
                c2cWb={c2cWb}
                overrides={currentProject.inputs}
                formulaOverrides={currentProject.formula_overrides}
                c2cOverrides={currentProject.summary?.c2c_overrides}
                c2cFormulaOverrides={currentProject.summary?.c2c_formula_overrides}
                setInput={handleSetInput}
                setFormulaOverride={handleSetFormulaOverride}
                isAdminOrOwner={isAdminOrOwner}
                session={session}
                org={org}
                onSave={() => saveProjectToServer(currentProject, true)}
                hasUnsavedChanges={hasUnsavedChanges}
                isSavingProject={isSavingProject}
              />
            ) : (
              <div style={{ padding: '40px 20px', textAlign: 'center', color: '#64748b' }}>
                Please select or create a project first to view calculation sheets.
              </div>
            )
          )}

          {tab === 'approvals' && isAdminOrOwner && (
            <ApprovalsPage
              session={session}
              org={org}
              projects={projects}
              wb={baseWb}
              onOpenProject={handleOpenProject}
              onProjectsChanged={() => loadProjects()}
              onOrgUpdated={() => onOrgsChanged()}
            />
          )}

          {tab === 'settings' && isAdminOrOwner && (
            <OrganisationSettings
              session={session}
              org={org}
              onOrgUpdated={() => {
                onOrgsChanged()
                loadProjects()
              }}
            />
          )}

          {tab === 'members' && (
            <MembersPage
              session={session}
              org={org}
              onOrgLeave={() => onOrgsChanged()}
            />
          )}

          {tab === 'activity' && (
            <ActivityPage
              session={session}
              org={org}
              onOpenProject={handleOpenProject}
            />
          )}
        </main>
      </div>

      {/* ---------------------------------------------------- Create Project Modal */}
      {showCreateModal && (
        <div className="overlay" onMouseDown={() => setShowCreateModal(false)}>
          <div className="dialog" role="dialog" onMouseDown={(e) => e.stopPropagation()}>
            <div className="dialog-header">
              {Icon.plus}
              <h3>Create New Project</h3>
            </div>
            <div className="form mt-3">
              <label>
                <span>Project Name</span>
                <input
                  type="text"
                  autoFocus
                  placeholder="e.g. QTY NO 102"
                  value={newProjectName}
                  onChange={(e) => setNewProjectName(e.target.value)}
                  onKeyDown={(e) => e.key === 'Enter' && handleCreateProject()}
                />
              </label>

              <div style={{ position: 'relative' }}>
                <label style={{ display: 'block', marginBottom: '4px' }}>
                  <span>Client / Buyer (Optional)</span>
                </label>
                <div style={{ position: 'relative', display: 'flex', alignItems: 'center' }}>
                  <input
                    type="text"
                    placeholder="Type to search or choose client..."
                    value={clientSearchTermForNewProj}
                    onFocus={() => setClientDropdownOpen(true)}
                    onChange={(e) => {
                      const val = e.target.value
                      setClientSearchTermForNewProj(val)
                      setClientDropdownOpen(true)
                      const matched = savedClients.find(
                        (c) => c.name.trim().toLowerCase() === val.trim().toLowerCase()
                      )
                      setSelectedClientForNewProj(matched ? matched.id : '')
                    }}
                    style={{
                      width: '100%',
                      height: '38px',
                      padding: '0 36px 0 10px',
                      border: '1px solid #cbd5e1',
                      borderRadius: '6px',
                      background: '#ffffff',
                      fontSize: '13px',
                      fontWeight: 500,
                    }}
                  />
                  <button
                    type="button"
                    tabIndex={-1}
                    onClick={() => setClientDropdownOpen((prev) => !prev)}
                    style={{
                      position: 'absolute',
                      right: '8px',
                      background: 'transparent',
                      border: 'none',
                      cursor: 'pointer',
                      padding: '4px 6px',
                      color: '#64748b',
                      display: 'flex',
                      alignItems: 'center',
                      justifyContent: 'center',
                    }}
                    title="Toggle client list"
                  >
                    <span style={{ fontSize: '10px' }}>{clientDropdownOpen ? '▲' : '▼'}</span>
                  </button>
                </div>

                {/* Dropdown list that opens on focus or click or typing */}
                {clientDropdownOpen && (
                  <>
                    <div
                      style={{ position: 'fixed', inset: 0, zIndex: 99 }}
                      onClick={() => setClientDropdownOpen(false)}
                    />
                    <div
                      style={{
                        position: 'absolute',
                        top: '100%',
                        left: 0,
                        right: 0,
                        zIndex: 100,
                        background: '#ffffff',
                        border: '1px solid #cbd5e1',
                        borderRadius: '6px',
                        boxShadow: '0 4px 14px rgba(0,0,0,0.12)',
                        maxHeight: '220px',
                        overflowY: 'auto',
                        marginTop: '4px',
                      }}
                    >
                      <div
                        onClick={() => {
                          setSelectedClientForNewProj('')
                          setClientSearchTermForNewProj('')
                          setClientDropdownOpen(false)
                        }}
                        style={{
                          padding: '8px 12px',
                          fontSize: '12px',
                          cursor: 'pointer',
                          borderBottom: '1px solid #f1f5f9',
                          color: '#64748b',
                          fontStyle: 'italic',
                        }}
                        onMouseEnter={(e) => (e.currentTarget.style.background = '#f8fafc')}
                        onMouseLeave={(e) => (e.currentTarget.style.background = 'transparent')}
                      >
                        -- None (Leave empty / enter later) --
                      </div>
                      {filteredClientsForNewProj.length === 0 ? (
                        <div style={{ padding: '10px 12px', fontSize: '12px', color: '#94a3b8' }}>
                          No clients found matching "{clientSearchTermForNewProj}"
                        </div>
                      ) : (
                        filteredClientsForNewProj.map((c) => {
                          const isSelected = c.id === selectedClientForNewProj
                          return (
                            <div
                              key={c.id}
                              onClick={() => {
                                setSelectedClientForNewProj(c.id)
                                setClientSearchTermForNewProj(c.name)
                                setClientDropdownOpen(false)
                              }}
                              style={{
                                padding: '8px 12px',
                                cursor: 'pointer',
                                borderBottom: '1px solid #f8fafc',
                                background: isSelected ? '#f1f5f9' : 'transparent',
                              }}
                              onMouseEnter={(e) => (e.currentTarget.style.background = '#f1f5f9')}
                              onMouseLeave={(e) =>
                                (e.currentTarget.style.background = isSelected ? '#f1f5f9' : 'transparent')
                              }
                            >
                              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                                <div style={{ fontWeight: 600, fontSize: '13px', color: '#0f172a' }}>{c.name}</div>
                                {isAdminOrOwner && c.profit_margin_pct !== undefined && (
                                  <span style={{ fontSize: '10px', background: '#f0fdf4', color: '#166534', border: '1px solid #bbf7d0', padding: '1px 5px', borderRadius: '3px', fontWeight: 700 }}>
                                    {c.profit_margin_pct}% Margin
                                  </span>
                                )}
                              </div>
                              {(c.gst || c.kind_att || c.contact) && (
                                <div style={{ fontSize: '11px', color: '#64748b', marginTop: '2px' }}>
                                  {c.gst ? `GST: ${c.gst}` : ''}
                                  {c.gst && (c.kind_att || c.contact) ? ' · ' : ''}
                                  {c.kind_att ? `Att: ${c.kind_att}` : ''}
                                  {c.contact ? ` (${c.contact})` : ''}
                                </div>
                              )}
                            </div>
                          )
                        })
                      )}
                    </div>
                  </>
                )}
              </div>
            </div>
            <div className="dialog-actions mt-4">
              <button className="btn ghost" onClick={() => setShowCreateModal(false)}>
                Cancel
              </button>
              <button
                className="btn primary"
                disabled={creating || !newProjectName.trim()}
                onClick={handleCreateProject}
              >
                {creating ? <Spinner size={16} /> : 'Create Project'}
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ---------------------------------------------------- My Profile Modal */}
      {showProfileModal && (
        <div className="overlay" onMouseDown={() => { if (!session.must_change_password) setShowProfileModal(false) }}>
          <div className="dialog" role="dialog" onMouseDown={(e) => e.stopPropagation()} style={{ maxWidth: 460 }}>
            <div className="dialog-header">
              <div className="avatar md" style={{ background: '#1e3a8a', color: '#fff', fontWeight: 800 }}>
                {(session?.username?.[0] || 'U').toUpperCase()}
              </div>
              <div style={{ marginLeft: 6 }}>
                <h3 style={{ margin: 0 }}>{session.must_change_password ? 'Security: Set Permanent Password' : 'My Profile & Security'}</h3>
              </div>
            </div>

            {session.must_change_password && (
              <div style={{ padding: '10px 12px', background: '#fef2f2', border: '1px solid #fecaca', borderRadius: 6, marginTop: 12 }}>
                <div className="font-bold text-sm" style={{ color: '#b91c1c' }}>Temporary Password Detected</div>
                <div className="tiny" style={{ color: '#991b1b', marginTop: 2 }}>
                  For account security, you must set a permanent password (minimum 10 characters) before continuing.
                </div>
              </div>
            )}

            <form onSubmit={handleUpdateProfile} className="form mt-3">
              {session?.email && (
                <div style={{ padding: '8px 10px', background: '#f8fafc', borderRadius: 6, border: '1px solid #e2e8f0', marginBottom: 12 }}>
                  <div className="tiny muted font-bold">Login Email:</div>
                  <div className="font-mono text-sm" style={{ color: '#0f172a', fontWeight: 600 }}>{session.email}</div>
                </div>
              )}

              <label>
                <span>Username</span>
                <input
                  type="text"
                  value={profileUsername}
                  onChange={(e) => setProfileUsername(e.target.value)}
                  placeholder="Enter your username"
                  autoFocus={!session.must_change_password}
                  required
                />
              </label>

              {newPassword.length > 0 && !session.must_change_password && (
                <label>
                  <span>Current Password <strong style={{ color: '#dc2626' }}>*</strong></span>
                  <input
                    type="password"
                    value={currentPassword}
                    onChange={(e) => setCurrentPassword(e.target.value)}
                    placeholder="Enter current password to verify"
                    autoComplete="current-password"
                    required
                  />
                </label>
              )}

              <label>
                <span>{session.must_change_password ? 'New Permanent Password (min. 10 characters)' : 'New Password (min. 10 characters)'}</span>
                <input
                  type="password"
                  value={newPassword}
                  onChange={(e) => setNewPassword(e.target.value)}
                  placeholder={session.must_change_password ? 'Enter minimum 10 characters' : 'Leave empty to keep existing password'}
                  autoComplete="new-password"
                  autoFocus={Boolean(session.must_change_password)}
                  required={Boolean(session.must_change_password)}
                />
              </label>

              {newPassword.length > 0 && (
                <label>
                  <span>Confirm New Password</span>
                  <input
                    type="password"
                    value={confirmPassword}
                    onChange={(e) => setConfirmPassword(e.target.value)}
                    placeholder="Re-type new password"
                    autoComplete="new-password"
                    required
                  />
                </label>
              )}

              <div className="dialog-actions mt-3">
                {!session.must_change_password && (
                  <button type="button" className="btn ghost" onClick={() => setShowProfileModal(false)}>
                    Cancel
                  </button>
                )}
                <button type="submit" className="btn primary" disabled={updatingProfile} style={session.must_change_password ? { width: '100%' } : undefined}>
                  {updatingProfile ? <Spinner size={16} /> : (session.must_change_password ? 'Set Password & Enter' : 'Save Profile')}
                </button>
              </div>
            </form>
          </div>
        </div>
      )}
          {showWorkspaceUnsavedModal && (
        <div className="overlay" onMouseDown={() => setShowWorkspaceUnsavedModal(false)}>
          <div className="dialog" role="dialog" onMouseDown={(e) => e.stopPropagation()} style={{ maxWidth: 440 }}>
            <div className="dialog-header">
              <div className="avatar md" style={{ background: '#fef3c7', color: '#b45309', fontWeight: 800 }}>⚠️</div>
              <div>
                <h3 style={{ margin: 0, fontSize: 16 }}>Unsaved Changes</h3>
                <p className="tiny muted" style={{ margin: 0 }}>Please save your changes before leaving</p>
              </div>
            </div>
            <div className="dialog-body" style={{ padding: '16px 20px', color: '#475569', fontSize: 13, lineHeight: 1.5 }}>
              You have unsaved changes in this calculation. If you leave without saving, any modified inputs or matrices will be lost.
            </div>
            <div className="dialog-footer" style={{ display: 'flex', justifyContent: 'flex-end', gap: 8, padding: '12px 20px', background: '#f8fafc', borderTop: '1px solid #e2e8f0' }}>
              <button
                type="button"
                className="btn ghost sm"
                onClick={() => {
                  setShowWorkspaceUnsavedModal(false)
                  setPendingTab(null)
                }}
              >
                Cancel & Stay
              </button>
              <button
                type="button"
                className="btn sm"
                style={{ background: '#fee2e2', color: '#b91c1c', border: '1px solid #fecaca' }}
                onClick={() => {
                  setCalculatorDirty(false)
                  setShowWorkspaceUnsavedModal(false)
                  if (pendingTab) {
                    setTab(pendingTab)
                    setPendingTab(null)
                    setNavOpen(false)
                  }
                }}
              >
                Leave Without Saving
              </button>
              <button
                type="button"
                className="btn primary sm"
                onClick={async () => {
                  if (calculatorSaveRef.current) {
                    const ok = await calculatorSaveRef.current()
                    if (ok) {
                      setCalculatorDirty(false)
                      setShowWorkspaceUnsavedModal(false)
                      if (pendingTab) {
                        setTab(pendingTab)
                        setPendingTab(null)
                        setNavOpen(false)
                      }
                    }
                  } else {
                    setCalculatorDirty(false)
                    setShowWorkspaceUnsavedModal(false)
                    if (pendingTab) {
                      setTab(pendingTab)
                      setPendingTab(null)
                      setNavOpen(false)
                    }
                  }
                }}
              >
                💾 Save & Leave
              </button>
            </div>
          </div>
        </div>
      )}
</div>
  )
}
