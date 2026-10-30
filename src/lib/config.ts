export const APP_NAME = 'Exam Calc'

const KEY = 'examcalc.supabase'

export interface SupaConfig {
  url: string
  anonKey: string
}

/** Config comes from .env (VITE_SUPABASE_URL / VITE_SUPABASE_ANON_KEY) or from the in-app setup screen. */
export function loadConfig(): SupaConfig | null {
  const envUrl = import.meta.env.VITE_SUPABASE_URL as string | undefined
  const envKey = import.meta.env.VITE_SUPABASE_ANON_KEY as string | undefined
  if (envUrl && envKey) return { url: envUrl, anonKey: envKey }
  try {
    const raw = localStorage.getItem(KEY)
    if (raw) {
      const c = JSON.parse(raw)
      if (c?.url && c?.anonKey) return c
    }
  } catch {
    /* storage unavailable */
  }
  return null
}

export function saveConfig(c: SupaConfig) {
  try {
    localStorage.setItem(KEY, JSON.stringify(c))
  } catch {
    /* ignore */
  }
}

export function clearConfig() {
  try {
    localStorage.removeItem(KEY)
  } catch {
    /* ignore */
  }
}

export function hasEnvConfig() {
  return !!(import.meta.env.VITE_SUPABASE_URL && import.meta.env.VITE_SUPABASE_ANON_KEY)
}
