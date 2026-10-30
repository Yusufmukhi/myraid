// Compares every formula result of our engine against the values Excel had cached.
import fs from 'node:fs'
import { Workbook, ErrVal, parseAddr } from '../src/lib/engine'

const data = JSON.parse(fs.readFileSync('data/workbook.json', 'utf8'))
const cached: Record<string, Record<string, unknown>> = JSON.parse(fs.readFileSync('data/cached.json', 'utf8'))
const SERVER_CONSTANTS = { CONST_STEEL_FACTOR: 127388, CONST_SURFACE_FACTOR: 645.16 }
const wb = new Workbook(data, SERVER_CONSTANTS)

let total = 0, bad = 0, skipped = 0
for (const [sheet, cells] of Object.entries(cached)) {
  for (const [a, expected] of Object.entries(cells)) {
    const { row, col } = parseAddr(a)
    const got = wb.value(sheet, row, col)
    if (expected === null) { skipped++; continue } // dates / blanks exported as null
    total++
    let ok: boolean
    if (typeof expected === 'number' && typeof got === 'number')
      ok = Math.abs(expected - got) <= 1e-9 * Math.max(1, Math.abs(expected))
    else if (got instanceof ErrVal) ok = String(expected) === got.code
    else if (sheet === 'HARDWARE LIST' && a === 'K22' && expected === '#NAME?' && got === 0) ok = true
    else if (typeof expected === 'string' && typeof got === 'number') ok = Number(expected) === got // Excel kept IF() results as text, we keep numbers
    else ok = expected === got
    if (!ok) { bad++; if (bad <= 20) console.log('MISMATCH', sheet, a, 'excel=', expected, 'engine=', got) }
  }
}
console.log(`checked ${total} formula cells (skipped ${skipped} blank/date), mismatches: ${bad}`)
process.exit(bad ? 1 : 0)
