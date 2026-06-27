// Pure helpers for turning a flexible "unit specification" string into a
// concrete list of unit numbers. Safe to import from both client (live
// preview) and server (bulk insert) — no DB or server-only imports here.

export const MAX_UNITS_PER_SPEC = 1000

/**
 * Expand a flexible unit spec into a deduped, ordered list of unit numbers.
 *
 * Supported syntax (comma- or newline-separated tokens):
 *   "1-20"              -> 1, 2, ... , 20
 *   "101-110, 201-210"  -> 101..110, 201..210
 *   "A, B, C"           -> A, B, C
 *   "1A, 1B, PH1"       -> literal tokens kept as-is
 *   "A1-A5"             -> A1, A2, A3, A4, A5   (matching prefixes expand)
 *   "001-010"           -> 001, 002, ... , 010  (zero-padding preserved)
 */
export function expandUnitSpec(spec: string): string[] {
  if (!spec || typeof spec !== 'string') return []

  const out: string[] = []
  const seen = new Set<string>()
  const push = (value: string) => {
    const t = value.trim()
    if (!t) return
    const key = t.toLowerCase()
    if (seen.has(key)) return
    seen.add(key)
    out.push(t)
  }

  const tokens = spec
    .split(/[\n,]+/)
    .map(s => s.trim())
    .filter(Boolean)

  for (const token of tokens) {
    if (out.length >= MAX_UNITS_PER_SPEC) break

    // Range: optional matching alpha prefix + number on both sides (e.g. 1-20, A1-A5, 101-110)
    const range = token.match(/^([A-Za-z]*)(\d+)\s*-\s*([A-Za-z]*)(\d+)$/)
    if (range) {
      const [, p1, n1, p2, n2] = range
      if (p1.toLowerCase() === p2.toLowerCase()) {
        const start = parseInt(n1, 10)
        const end = parseInt(n2, 10)
        // preserve zero padding when the start is padded (e.g. 001-010)
        const pad = n1.length > 1 && n1.startsWith('0') ? n1.length : 0
        const step = start <= end ? 1 : -1
        for (let i = start; step > 0 ? i <= end : i >= end; i += step) {
          const num = pad ? String(Math.abs(i)).padStart(pad, '0') : String(i)
          push(`${p1}${num}`)
          if (out.length >= MAX_UNITS_PER_SPEC) break
        }
        continue
      }
    }

    push(token)
  }

  return out
}

/** Short human summary of what a spec will create, for live previews. */
export function summarizeUnits(units: string[]): string {
  if (units.length === 0) return 'No units'
  const preview = units.slice(0, 6).join(', ')
  const more = units.length > 6 ? `, +${units.length - 6} more` : ''
  return `${units.length} unit${units.length === 1 ? '' : 's'}: ${preview}${more}`
}
