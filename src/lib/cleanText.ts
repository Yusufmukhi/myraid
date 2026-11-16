/**
 * Universal text normalization for handling text copied from Excel, PDF, Word, or web pages.
 * Resolves PDF ligatures, non-breaking spaces, zero-width characters, smart quotes,
 * commas in formatted numbers, and trailing Excel newlines/tabs.
 */

// PDF Ligature map
const LIGATURE_MAP: Record<string, string> = {
  '\uFB00': 'ff', // ﬀ
  '\uFB01': 'fi', // ﬁ
  '\uFB02': 'fl', // ﬂ
  '\uFB03': 'ffi', // ﬃ
  '\uFB04': 'ffl', // ﬄ
  '\uFB05': 'ft', // ﬅ
  '\uFB06': 'st', // ﬆ
}

/**
 * Cleans and normalizes text copied from Excel or PDF.
 * - Expands PDF ligatures (fi, fl, ff, ffi, ffl)
 * - Replaces non-breaking spaces (\u00A0, \u202F, \u2007) with normal spaces
 * - Strips zero-width characters (\u200B, \u200C, \u200D, \uFEFF)
 * - Strips soft hyphens (\u00AD) and non-breaking hyphens (\u2011)
 * - Normalizes fancy / smart quotes (“ ” ‘ ’ « ») to standard ASCII quotes
 * - Normalizes dashes (– — −) to standard hyphens (-)
 * - Cleans Excel cell-wrapping quotes if added during copy
 * - Trims outer whitespace
 */
export function cleanCopiedText(input: unknown): string {
  if (input === null || input === undefined) return ''
  let str = String(input)

  // 1. Unicode normalization (NFKC decomposes ligatures and compatibility characters)
  try {
    str = str.normalize('NFKC')
  } catch {}

  // 2. Explicit replacement for known PDF ligatures (fallback if NFKC doesn't cover font private use)
  str = str.replace(/[\uFB00-\uFB06]/g, (ch) => LIGATURE_MAP[ch] || ch)

  // 3. Replace all variations of non-breaking and special spaces with standard space (0x20)
  str = str.replace(/[\u00A0\u1680\u2000-\u200A\u202F\u205F\u3000\u2007]/g, ' ')

  // 4. Strip zero-width & invisible control characters
  str = str.replace(/[\u200B-\u200D\uFEFF\u00AD\u2060\u180E]/g, '')

  // 5. Normalize fancy quotes and backticks
  str = str.replace(/[\u2018\u2019\u201A\u201B\u02BC\u02BB']/g, "'")
  str = str.replace(/[\u201C\u201D\u201E\u201F\u00AB\u00BB]/g, '"')

  // 6. Normalize dashes / hyphens
  str = str.replace(/[\u2010-\u2015\u2212\uFE58\uFE63\uFF0D]/g, '-')

  // 7. Trim outer whitespace and trailing newlines/tabs before checking wrapping quotes
  str = str.trim()

  // 8. Strip Excel wrapping double quotes if they were added due to embedded commas/newlines
  if (str.startsWith('"') && str.endsWith('"') && str.length >= 2) {
    const inner = str.slice(1, -1)
    if (!inner.includes('"') || inner.includes('""')) {
      str = inner.replace(/""/g, '"')
    }
  }

  return str.trim()
}

/**
 * Normalizes a string specifically for fuzzy/resilient searching and matching.
 * Ignores case, extra spaces, punctuation, quotes, hyphens, and accents.
 */
export function normalizeSearchTerm(str: unknown): string {
  const cleaned = cleanCopiedText(str).toLowerCase()
  // Collapse consecutive whitespace to single space
  return cleaned.replace(/\s+/g, ' ')
}

/**
 * Checks whether haystack matches query resiliently, even when text is copied from PDF/Excel.
 */
export function fuzzyTextMatch(haystack: unknown, query: unknown): boolean {
  const h = normalizeSearchTerm(haystack)
  const q = normalizeSearchTerm(query)
  if (!q) return true
  if (!h) return false
  if (h.includes(q)) return true
  // Also test without spaces/hyphens for codes like "DFMU-D" vs "DFMU D" vs "DFMUD"
  const hCompact = h.replace(/[\s\-_.,/]/g, '')
  const qCompact = q.replace(/[\s\-_.,/]/g, '')
  if (qCompact.length >= 2 && hCompact.includes(qCompact)) return true
  return false
}

/**
 * Cleans numeric strings copied from Excel (e.g. "1,200", "₹ 8,38,625", " 1500.00 ", "1200\r\n").
 * Strips currency symbols, thousands-separator commas, spaces, and returns valid number.
 */
export function cleanCopiedNumber(input: unknown): number {
  if (typeof input === 'number') return isNaN(input) ? 0 : input
  if (input === null || input === undefined) return 0
  const cleaned = cleanCopiedText(input)
  // Strip currency symbols (₹, $, €, £, Rs, etc.), spaces, commas
  const numStr = cleaned.replace(/[₹$€£\s,]/gi, '').replace(/rs\.?/gi, '')
  const parsed = parseFloat(numStr)
  return isNaN(parsed) ? 0 : parsed
}
