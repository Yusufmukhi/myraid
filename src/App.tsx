import { useCallback, useEffect, useState } from 'react'
import { api, initClient, Org, Session, tokenStore } from './lib/api'
import { loadConfig, SupaConfig } from './lib/config'
import { AuthScreen, BootScreen, FirstTimePasswordChangeScreen, OrgGate, SetupScreen } from './components/Screens'
import { Workspace } from './components/Workspace'

const ORG_KEY = 'examcalc.org'
const ORGS_CACHE_KEY = 'examcalc.orgs_cache'

const getSavedOrg = () => {
  try {
    return localStorage.getItem(ORG_KEY)
  } catch {
    return null
  }
}
const setSavedOrg = (id: string) => {
  try {
    localStorage.setItem(ORG_KEY, id)
  } catch {
    /* ignore */
  }
}

const getCachedOrgs = (): Org[] | null => {
  try {
    const raw = localStorage.getItem(ORGS_CACHE_KEY)
    return raw ? JSON.parse(raw) : null
  } catch {
    return null
  }
}
const setCachedOrgs = (list: Org[]) => {
  try {
    localStorage.setItem(ORGS_CACHE_KEY, JSON.stringify(list))
  } catch {
    /* ignore */
  }
}

export default function App() {
  const [cfg, setCfg] = useState<SupaConfig | null>(() => {
    const c = loadConfig()
    if (c) initClient(c)
    return c
  })
  const [booting, setBooting] = useState(!!cfg)
  const [session, setSession] = useState<Session | null>(null)
  const [orgs, setOrgs] = useState<Org[] | null>(() => getCachedOrgs())
  const [orgId, setOrgId] = useState<string | null>(getSavedOrg())
  const [addingOrg, setAddingOrg] = useState(false)
  const [bootError, setBootError] = useState('')

  // restore a previous session
  useEffect(() => {
    if (!cfg) return
    let alive = true
    ;(async () => {
      const token = tokenStore.get()
      if (!token) return setBooting(false)
      try {
        const me = await api.me(token)
        if (!alive) return
        if (me) setSession({ ...me, token })
        else tokenStore.clear()
      } catch (e) {
        if (alive) setBootError((e as Error).message)
      } finally {
        if (alive) setBooting(false)
      }
    })()
    return () => {
      alive = false
    }
  }, [cfg])

  const loadOrgs = useCallback(async (s: Session) => {
    try {
      const list = await api.orgList(s.token)
      if (Array.isArray(list) && list.length > 0) {
        setOrgs(list)
        setCachedOrgs(list)
      } else if (Array.isArray(list) && list.length === 0) {
        // Only set empty if server actually returned an empty array AND no cached orgs exist
        const cached = getCachedOrgs()
        if (!cached || cached.length === 0) {
          setOrgs([])
        }
      }
    } catch (e) {
      const msg = (e as Error).message
      if (/permanent password|must_change_password|security policy/i.test(msg)) {
        if (s) {
          setSession((prev) => (prev ? { ...prev, must_change_password: true } : null))
        }
        return
      }
      if (/expired|invalid token|jwt/i.test(msg)) {
        tokenStore.clear()
        setSession(null)
        setOrgs(null)
        try { localStorage.removeItem(ORGS_CACHE_KEY) } catch {}
        return
      }
      // If a transient network glitch or sleep/wake happens, preserve existing loaded orgs or cached orgs
      const cached = getCachedOrgs()
      if (cached && cached.length > 0) {
        setOrgs((prev) => (prev && prev.length > 0 ? prev : cached))
      }
    }
  }, [])

  useEffect(() => {
    if (session && !session.must_change_password) loadOrgs(session)
    else if (!session) setOrgs(null)
  }, [session, loadOrgs])

  // Automatically re-fetch organisations when the user returns to the tab or network reconnects
  useEffect(() => {
    if (!session || session.must_change_password) return
    const handleRecheck = () => {
      loadOrgs(session)
    }
    window.addEventListener('focus', handleRecheck)
    window.addEventListener('online', handleRecheck)
    return () => {
      window.removeEventListener('focus', handleRecheck)
      window.removeEventListener('online', handleRecheck)
    }
  }, [session, loadOrgs])

  const logout = async (reason?: string) => {
    if (session) api.logout(session.token).catch(() => {})
    tokenStore.clear()
    try { localStorage.removeItem(ORGS_CACHE_KEY) } catch {}
    setSession(null)
    setOrgs(null)
    // Only display an alert if reason is an explicit non-empty string message (not a DOM ClickEvent object)
    if (typeof reason === 'string' && reason.trim().length > 0) {
      setTimeout(() => alert(reason), 100)
    }
  }

  // 2-Hour Inactivity Auto-Lock for factory floor / workstation safety
  useEffect(() => {
    if (!session) return

    let timer: any = null
    const INACTIVITY_MS = 2 * 60 * 60 * 1000 // 2 hours

    const resetTimer = () => {
      if (timer) clearTimeout(timer)
      timer = setTimeout(() => {
        logout('Session locked due to 2 hours of inactivity to protect sensitive quotation data.')
      }, INACTIVITY_MS)
    }

    resetTimer()
    const events = ['mousedown', 'mousemove', 'keydown', 'scroll', 'touchstart']
    const handler = () => resetTimer()
    for (const evt of events) {
      window.addEventListener(evt, handler, { passive: true })
    }

    return () => {
      if (timer) clearTimeout(timer)
      for (const evt of events) {
        window.removeEventListener(evt, handler)
      }
    }
  }, [session])

  const pickOrg = (o: Org) => {
    setSavedOrg(o.id)
    setOrgId(o.id)
    setAddingOrg(false)
    if (session) loadOrgs(session)
  }

  if (!cfg)
    return (
      <SetupScreen
        onDone={(c) => {
          initClient(c)
          setCfg(c)
          setBooting(false)
        }}
      />
    )
  if (booting) return <BootScreen />
  if (bootError)
    return (
      <div className="boot">
        <div className="card auth-card">
          <h2>Cannot reach the database</h2>
          <p className="muted">{bootError}</p>
          <button className="btn primary" onClick={() => location.reload()}>
            Retry
          </button>
        </div>
      </div>
    )
  if (!session)
    return (
      <AuthScreen
        onAuth={(s) => {
          tokenStore.set(s.token)
          setSession(s)
        }}
      />
    )

  if (session.must_change_password) {
    return (
      <FirstTimePasswordChangeScreen
        session={session}
        onPasswordChanged={(updated) => {
          setSession(updated)
          loadOrgs(updated)
        }}
        onLogout={logout}
      />
    )
  }

  if (orgs === null) return <BootScreen text="Loading your organisations…" />
  if (orgs.length === 0 || addingOrg)
    return <OrgGate session={session} onJoined={pickOrg} onCancel={orgs.length ? () => setAddingOrg(false) : undefined} />

  const org = (orgs && orgs.length > 0) ? (orgs.find((o) => o.id === orgId) ?? orgs[0]) : null
  if (!org) {
    return <OrgGate session={session} onJoined={pickOrg} />
  }
  return (
    <Workspace
      key={org.id}
      session={session}
      org={org}
      orgs={orgs}
      onSwitchOrg={pickOrg}
      onAddOrg={() => setAddingOrg(true)}
      onOrgsChanged={() => loadOrgs(session)}
      onLogout={logout}
      onSessionChange={setSession}
    />
  )
}
