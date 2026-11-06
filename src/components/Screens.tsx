import { FormEvent, useState } from 'react'
import { api, Org, Session, initClient } from '../lib/api'
import { APP_NAME, clearConfig, hasEnvConfig, saveConfig, SupaConfig } from '../lib/config'
import { Icon, Spinner, useCopy, useUI } from './ui'

export function Logo({ light }: { light?: boolean }) {
  return (
    <span className={`logo ${light ? 'light' : ''}`}>
      <svg viewBox="0 0 32 32" width="30" height="30" aria-hidden>
        <rect width="32" height="32" rx="8" fill={light ? '#ffffff' : '#1e3a8a'} />
        <rect x="7" y="7" width="18" height="6" rx="1.5" fill={light ? '#1e3a8a' : '#fff'} />
        <rect x="7" y="16" width="7" height="9" rx="1.5" fill={light ? '#60a5fa' : '#93c5fd'} />
        <rect x="17" y="16" width="8" height="9" rx="1.5" fill={light ? '#60a5fa' : '#93c5fd'} />
      </svg>
      <span>{APP_NAME}</span>
    </span>
  )
}

function Shell({ children, wide }: { children: React.ReactNode; wide?: boolean }) {
  return (
    <div className="auth-shell">
      <div className="auth-brand">
        <Logo light />
        <div>
          <h1>Estimate storage units in seconds.</h1>
          <p>
            Every formula from your Excel costing workbook – body, shelves, trolley, hardware, powder coating, labour and GST –
            calculated live, saved to your team.
          </p>
        </div>
        <ul>
          <li>Same results as the Excel sheets</li>
          <li>Save, edit and rename estimates</li>
          <li>Collaborate seamlessly with your team</li>
        </ul>
      </div>
      <div className="auth-panel">
        <div className={`card auth-card ${wide ? 'wide' : ''}`}>{children}</div>
      </div>
    </div>
  )
}

/* ------------------------------------------------------- Supabase setup */
export function SetupScreen({ onDone }: { onDone: (c: SupaConfig) => void }) {
  const [url, setUrl] = useState('')
  const [key, setKey] = useState('')
  const [err, setErr] = useState('')
  const submit = (e: FormEvent) => {
    e.preventDefault()
    if (!/^https:\/\/.+\.supabase\.(co|in)\/?$/.test(url.trim())) {
      setErr('That does not look like a Supabase project URL (https://xxxx.supabase.co).')
      return
    }
    if (key.trim().length < 20) {
      setErr('Paste the anon / public key from Project Settings → API.')
      return
    }
    const c = { url: url.trim().replace(/\/$/, ''), anonKey: key.trim() }
    saveConfig(c)
    onDone(c)
  }
  return (
    <Shell>
      <h2>Connect your database</h2>
      <p className="muted">
        One-time setup. In Supabase open <b>Project Settings → API</b> and copy the Project URL and the <b>anon public</b> key. Then run{' '}
        <code>supabase/schema.sql</code> in the SQL editor.
      </p>
      <form onSubmit={submit} className="form">
        <label>
          Project URL
          <input value={url} onChange={(e) => setUrl(e.target.value)} placeholder="https://abcdxyz.supabase.co" autoFocus />
        </label>
        <label>
          Anon public key
          <input value={key} onChange={(e) => setKey(e.target.value)} placeholder="eyJhbGciOi…" />
        </label>
        {err && <div className="form-error">{err}</div>}
        <button className="btn primary lg">Connect</button>
      </form>
    </Shell>
  )
}

/* ---------------------------------------------------------------- Auth */
function evaluatePasswordStrength(pass: string) {
  const hasMinLength = pass.length >= 10
  const hasNumber = /[0-9]/.test(pass)
  const hasUpper = /[A-Z]/.test(pass)
  const hasLower = /[a-z]/.test(pass)
  const hasSymbol = /[^A-Za-z0-9]/.test(pass)

  let score = 0
  if (hasMinLength) score++
  if (hasNumber) score++
  if (hasUpper && hasLower) score++
  if (hasSymbol) score++
  if (pass.length >= 12) score++

  let label = 'Too Weak'
  let color = '#ef4444' // red
  let pct = 20
  if (score <= 1) {
    label = 'Weak'
    color = '#ef4444'
    pct = 25
  } else if (score === 2) {
    label = 'Fair'
    color = '#f97316'
    pct = 50
  } else if (score === 3) {
    label = 'Good'
    color = '#eab308'
    pct = 75
  } else if (score >= 4) {
    label = 'Strong'
    color = '#10b981'
    pct = 100
  }

  return {
    score,
    label,
    color,
    pct,
    hasMinLength,
    hasNumber,
    hasUpper,
    hasLower,
    hasSymbol,
  }
}

export function AuthScreen({ onAuth }: { onAuth: (s: Session) => void }) {
  const [username, setUsername] = useState('')
  const [password, setPassword] = useState('')
  const [showPassword, setShowPassword] = useState(false)
  const [rememberMe, setRememberMe] = useState(true)
  const [busy, setBusy] = useState(false)
  const [err, setErr] = useState('')

  const submit = async (e: FormEvent) => {
    e.preventDefault()
    setErr('')
    setBusy(true)
    try {
      const s = await api.login(username, password)
      try {
        const me = await api.me(s.token)
        if (me?.email) s.email = me.email
      } catch {}
      if (!s.email && username.includes('@')) {
        s.email = username.trim()
      }
      // Save token according to remember me preference
      try {
        const { tokenStore } = await import('../lib/api')
        tokenStore.set(s.token, rememberMe)
      } catch {
        /* ignore */
      }
      onAuth(s)
    } catch (ex) {
      setErr((ex as Error).message)
    } finally {
      setBusy(false)
    }
  }

  // Detect brute-force lockout warning
  const isLockoutWarning = err.includes('attempt(s) remaining')
  const isAccountLocked = err.includes('temporarily locked') || err.includes('Account locked')

  return (
    <Shell>
      <h2>Sign in to your organisation</h2>
      <p className="tiny muted mb-3" style={{ textAlign: 'center' }}>
        Enter your registered email address or username to access your workspace.
      </p>
      <form onSubmit={submit} className="form">
        <label>
          Email or Username
          <input
            value={username}
            onChange={(e) => setUsername(e.target.value)}
            autoComplete="username"
            placeholder="colleague@company.com or username"
            autoFocus
            required
          />
        </label>
        <label>
          Password
          <div style={{ position: 'relative', display: 'flex', alignItems: 'center' }}>
            <input
              type={showPassword ? 'text' : 'password'}
              value={password}
              onChange={(e) => setPassword(e.target.value)}
              autoComplete="current-password"
              placeholder="Enter your password"
              style={{ width: '100%', paddingRight: '40px' }}
              required
            />
            <button
              type="button"
              className="icon-btn sm"
              onClick={() => setShowPassword(!showPassword)}
              style={{
                position: 'absolute',
                right: '8px',
                background: 'transparent',
                border: 'none',
                cursor: 'pointer',
                fontSize: '15px',
                padding: '4px',
              }}
              title={showPassword ? 'Hide password' : 'Show password'}
            >
              {showPassword ? '🙈' : '👁️'}
            </button>
          </div>
        </label>

        {/* Real-time Password Strength Meter (Shown on Signup or when typing) */}
        {/* Remember Me Option */}
        <div style={{ marginTop: '4px', marginBottom: '6px' }}>
          <label style={{ display: 'inline-flex', alignItems: 'center', gap: '8px', cursor: 'pointer', margin: 0, fontSize: '12.5px' }}>
            <input
              type="checkbox"
              checked={rememberMe}
              onChange={(e) => setRememberMe(e.target.checked)}
              style={{ width: 'auto', margin: 0 }}
            />
            <span>Remember this device</span>
          </label>
        </div>

        {/* Dynamic Brute Force / Lockout Warnings */}
        {err && (
          <div
            className={`form-error ${isAccountLocked ? 'locked-box' : isLockoutWarning ? 'warning-box' : ''}`}
            style={{
              padding: '10px 12px',
              borderRadius: '6px',
              background: isAccountLocked ? '#fef2f2' : isLockoutWarning ? '#fffbeb' : '#fef2f2',
              border: `1px solid ${isAccountLocked ? '#fca5a5' : isLockoutWarning ? '#fde68a' : '#fecaca'}`,
              color: isAccountLocked ? '#991b1b' : isLockoutWarning ? '#92400e' : '#b91c1c',
              fontSize: '12.5px',
              lineHeight: '1.4',
            }}
          >
            {isAccountLocked ? '🔒 ' : isLockoutWarning ? '⚠️ ' : ''}
            {err}
          </div>
        )}

        <button className="btn primary lg" disabled={busy}>
          {busy ? <Spinner /> : 'Sign in'}
        </button>
      </form>
      {!hasEnvConfig() && (
        <button
          className="link tiny mt-3"
          onClick={() => {
            clearConfig()
            location.reload()
          }}
        >
          Change database connection
        </button>
      )}
    </Shell>
  )
}

/* ----------------------------------------------- Create organisation */
export function OrgGate({
  session,
  onJoined,
  onCancel,
}: {
  session: Session
  onJoined: (o: Org) => void
  onCancel?: () => void
}) {
  const [name, setName] = useState('')
  const [busy, setBusy] = useState(false)
  const [err, setErr] = useState('')
  const [created, setCreated] = useState<Org | null>(null)

  const create = async (e: FormEvent) => {
    e.preventDefault()
    if (!name.trim()) return
    setErr('')
    setBusy(true)
    try {
      setCreated(await api.orgCreate(session.token, name.trim()))
    } catch (ex) {
      setErr((ex as Error).message)
    } finally {
      setBusy(false)
    }
  }

  if (created) {
    return (
      <Shell>
        <div className="success-badge">{Icon.check}</div>
        <h2>{created.name} is ready</h2>
        <p className="muted">
          Your organisation workspace has been created. You can now configure projects, manage pricing, and invite team members directly by email from the Team Members page.
        </p>
        <button className="btn primary lg" onClick={() => onJoined(created)}>
          Open workspace
        </button>
      </Shell>
    )
  }

  return (
    <Shell>
      <div className="row between">
        <h2>Create an organisation</h2>
        {onCancel ? (
          <button className="btn ghost sm" onClick={onCancel}>
            Back
          </button>
        ) : (
          <button className="btn ghost sm" onClick={() => location.reload()} title="Reload and open your existing workspace">
            Reload Workspace
          </button>
        )}
      </div>
      <p className="muted">
        Signed in as <b>{session.username}</b>. Set up your company workspace to manage projects, cutting formulas, and team members.
      </p>
      <form onSubmit={create} className="form">
        <label>
          Organisation Name
          <input
            value={name}
            onChange={(e) => setName(e.target.value)}
            placeholder="e.g. Apollo Storage Systems"
            autoFocus
          />
        </label>
        {err && <div className="form-error">{err}</div>}
        <button className="btn primary lg" disabled={busy || !name.trim()}>
          {busy ? <Spinner /> : 'Create organisation'}
        </button>
      </form>
    </Shell>
  )
}

export function BootScreen({ text = 'Loading…' }: { text?: string }) {
  return (
    <div className="boot">
      <Logo />
      <div className="row gap">
        <Spinner size={18} />
        <span className="muted">{text}</span>
      </div>
    </div>
  )
}

export { initClient }


export function FirstTimePasswordChangeScreen({
  session,
  onPasswordChanged,
  onLogout,
}: {
  session: Session
  onPasswordChanged: (updated: Session) => void
  onLogout: () => void
}) {
  const [username, setUsername] = useState(session.username || '')
  const [email, setEmail] = useState(session.email || '')
  const [newPassword, setNewPassword] = useState('')
  const [confirmPassword, setConfirmPassword] = useState('')
  const [showPassword, setShowPassword] = useState(false)
  const [busy, setBusy] = useState(false)
  const [err, setErr] = useState('')

  const submit = async (e: FormEvent) => {
    e.preventDefault()
    setErr('')

    if (!newPassword || newPassword.length < 10) {
      setErr('Password must be at least 10 characters long.')
      return
    }
    const hasUpper = /[A-Z]/.test(newPassword)
    const hasLower = /[a-z]/.test(newPassword)
    const hasDigit = /[0-9]/.test(newPassword)
    if (!hasUpper || !hasLower || !hasDigit) {
      setErr('Password must contain a mix of uppercase letters, lowercase letters, and numbers.')
      return
    }
    if (newPassword !== confirmPassword) {
      setErr('Passwords do not match.')
      return
    }

    setBusy(true)
    try {
      const res = await api.updateProfile(
        session.token,
        username.trim() !== session.username ? username.trim() : undefined,
        newPassword,
        email.trim() || undefined
      )
      const updated: Session = {
        ...session,
        username: res?.username || username.trim() || session.username,
        email: res?.email || email.trim() || session.email,
        must_change_password: false,
      }
      onPasswordChanged(updated)
    } catch (ex: any) {
      setErr(ex?.message || 'Failed to update password')
    } finally {
      setBusy(false)
    }
  }

  return (
    <Shell>
      <div style={{ textAlign: 'center', marginBottom: '16px' }}>
        <div style={{ fontSize: '32px', marginBottom: '8px' }}>🔐</div>
        <h2>Set Your Permanent Password</h2>
        <p className="tiny muted" style={{ maxWidth: '340px', margin: '4px auto 0', lineHeight: 1.4 }}>
          You have signed in with temporary credentials. For portal security, please choose a permanent password (min. 10 mixed characters) before continuing.
        </p>
      </div>

      <form onSubmit={submit} className="form">
        <label>
          Username
          <input
            value={username}
            onChange={(e) => setUsername(e.target.value)}
            required
            autoComplete="username"
          />
        </label>

        <label>
          Email Address
          <input
            type="email"
            value={email}
            onChange={(e) => setEmail(e.target.value)}
            placeholder="colleague@company.com"
          />
        </label>

        <label>
          New Permanent Password (min. 10 mixed characters)
          <div style={{ position: 'relative', display: 'flex', alignItems: 'center' }}>
            <input
              type={showPassword ? 'text' : 'password'}
              value={newPassword}
              onChange={(e) => setNewPassword(e.target.value)}
              placeholder="e.g. K9mP2xW7qR"
              autoFocus
              required
              autoComplete="new-password"
              style={{ width: '100%', paddingRight: '40px' }}
            />
            <button
              type="button"
              className="icon-btn sm"
              onClick={() => setShowPassword(!showPassword)}
              style={{
                position: 'absolute',
                right: '8px',
                background: 'transparent',
                border: 'none',
                cursor: 'pointer',
                fontSize: '15px',
                padding: '4px',
              }}
              title={showPassword ? 'Hide password' : 'Show password'}
            >
              {showPassword ? '🙈' : '👁️'}
            </button>
          </div>
          <span className="tiny muted" style={{ fontSize: '11px', marginTop: '2px', display: 'block' }}>
            Must contain at least 10 characters with uppercase, lowercase, and numbers.
          </span>
        </label>

        <label>
          Confirm Permanent Password
          <input
            type="password"
            value={confirmPassword}
            onChange={(e) => setConfirmPassword(e.target.value)}
            placeholder="Re-enter your new password"
            required
            autoComplete="new-password"
          />
        </label>

        {err && <div className="banner danger" style={{ marginTop: '8px' }}>{err}</div>}

        <button type="submit" className="btn primary block mt-3" disabled={busy} style={{ width: '100%', height: '38px', fontWeight: 700 }}>
          {busy ? <Spinner size={16} /> : 'Set Permanent Password & Enter Portal'}
        </button>

        <button
          type="button"
          className="btn ghost block mt-2"
          onClick={onLogout}
          style={{ width: '100%', fontSize: '12px', color: '#64748b' }}
        >
          Sign out and return later
        </button>
      </form>
    </Shell>
  )
}
