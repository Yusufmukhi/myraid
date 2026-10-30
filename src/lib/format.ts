import { ErrVal, Val } from './engine'

const IN = 'en-IN'

export const inr = (n: number) =>
  '₹' + (Number.isFinite(n) ? n : 0).toLocaleString(IN, { maximumFractionDigits: 0 })

export const inr2 = (n: number) =>
  '₹' + (Number.isFinite(n) ? n : 0).toLocaleString(IN, { minimumFractionDigits: 2, maximumFractionDigits: 2 })

export const kg = (n: number) =>
  (Number.isFinite(n) ? n : 0).toLocaleString(IN, { minimumFractionDigits: 1, maximumFractionDigits: 1 }) + ' kg'

export const count = (n: number) => (Number.isFinite(n) ? n : 0).toLocaleString(IN, { maximumFractionDigits: 2 })

/** Excel-like "General" display for grid cells. */
export function fmtCell(v: Val): string {
  if (v === null || v === '') return ''
  if (v instanceof ErrVal) return v.code
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE'
  if (typeof v === 'string') return v
  if (!Number.isFinite(v)) return String(v)
  if (Number.isInteger(v)) return v.toLocaleString(IN, { useGrouping: Math.abs(v) >= 100000 })
  const abs = Math.abs(v)
  const d = abs >= 100 ? 2 : abs >= 1 ? 3 : 4
  return v.toLocaleString(IN, { maximumFractionDigits: d, useGrouping: abs >= 100000 })
}

export function timeAgo(iso: string): string {
  const t = new Date(iso).getTime()
  if (Number.isNaN(t)) return ''
  const s = Math.max(0, (Date.now() - t) / 1000)
  if (s < 60) return 'just now'
  if (s < 3600) return `${Math.floor(s / 60)} min ago`
  if (s < 86400) return `${Math.floor(s / 3600)} h ago`
  if (s < 86400 * 7) return `${Math.floor(s / 86400)} d ago`
  return new Date(iso).toLocaleDateString(IN, { day: 'numeric', month: 'short', year: 'numeric' })
}
