/**
 * Tiny spreadsheet engine that reproduces the exam.xlsx workbook.
 *
 * The workbook only uses: + - * / ^ & comparisons, percent, SUM(), IF(),
 * cell refs, ranges and cross-sheet refs — so a small evaluator is enough and
 * it is verified cell-by-cell against Excel's own results (npm run verify).
 */

export class ErrVal {
  constructor(public code: string) {}
  toString() {
    return this.code
  }
}
export type Val = number | string | boolean | null | ErrVal
export type RawCell = string | number | null

export interface SheetData {
  name: string
  rows: number
  cols: number
  cells: RawCell[][]
  styles?: Record<string, CellStyle>
  merges?: string[]
}
export interface CellStyle {
  b?: 1
  i?: 1
  fg?: string
  bg?: string
  al?: 'l' | 'c' | 'r'
}

// ---------------------------------------------------------------- addresses
export function colToNum(s: string): number {
  let n = 0
  for (const ch of s) n = n * 26 + (ch.charCodeAt(0) - 64)
  return n - 1
}
export function numToCol(n: number): string {
  let s = ''
  n += 1
  while (n > 0) {
    s = String.fromCharCode(65 + ((n - 1) % 26)) + s
    n = Math.floor((n - 1) / 26)
  }
  return s
}
export function addr(row: number, col: number): string {
  return `${numToCol(col)}${row + 1}`
}
export function parseAddr(a: string): { row: number; col: number } {
  const m = /^\$?([A-Z]+)\$?(\d+)$/.exec(a)
  if (!m) throw new Error('bad address ' + a)
  return { row: parseInt(m[2], 10) - 1, col: colToNum(m[1]) }
}

// ------------------------------------------------------------------- parser
type Node =
  | { t: 'num'; v: number }
  | { t: 'str'; v: string }
  | { t: 'empty' }
  | { t: 'ref'; sheet: string | null; row: number; col: number }
  | { t: 'range'; sheet: string | null; r1: number; c1: number; r2: number; c2: number }
  | { t: 'name'; v: string }
  | { t: 'un'; op: string; a: Node }
  | { t: 'pct'; a: Node }
  | { t: 'bin'; op: string; a: Node; b: Node }
  | { t: 'fn'; name: string; args: Node[] }

interface Tok {
  k: 'num' | 'str' | 'id' | 'ref' | 'op' | 'end'
  v: string
  sheet?: string | null
}

function tokenize(src: string): Tok[] {
  const out: Tok[] = []
  let i = 0
  const n = src.length
  while (i < n) {
    const ch = src[i]
    if (ch === ' ') {
      i++
      continue
    }
    if (ch === '"') {
      let j = i + 1
      let s = ''
      while (j < n) {
        if (src[j] === '"') {
          if (src[j + 1] === '"') {
            s += '"'
            j += 2
            continue
          }
          break
        }
        s += src[j++]
      }
      out.push({ k: 'str', v: s })
      i = j + 1
      continue
    }
    if (/[0-9.]/.test(ch)) {
      const m = /^\d*\.?\d+(?:[eE][+-]?\d+)?|^\d+\.?/.exec(src.slice(i))!
      out.push({ k: 'num', v: m[0] })
      i += m[0].length
      continue
    }
    // 'Sheet name'!A1  or  Sheet!A1
    let sheet: string | null = null
    let rest = src.slice(i)
    let m = /^'((?:[^']|'')+)'!/.exec(rest)
    if (m) {
      sheet = m[1].replace(/''/g, "'")
    } else {
      m = /^([A-Za-z_][A-Za-z0-9_.\-]*)!/.exec(rest)
      if (m) sheet = m[1]
    }
    if (m) {
      i += m[0].length
      rest = src.slice(i)
    }
    const r = /^\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?/.exec(rest)
    if (r && !/^[A-Za-z0-9_]/.test(rest.slice(r[0].length))) {
      out.push({ k: 'ref', v: r[0], sheet })
      i += r[0].length
      continue
    }
    if (sheet !== null) throw new Error('bad ref ' + src)
    const id = /^[A-Za-z_][A-Za-z0-9_.]*/.exec(rest)
    if (id) {
      out.push({ k: 'id', v: id[0] })
      i += id[0].length
      continue
    }
    const two = src.slice(i, i + 2)
    if (two === '<=' || two === '>=' || two === '<>') {
      out.push({ k: 'op', v: two })
      i += 2
      continue
    }
    if ('+-*/^&=<>(),%'.includes(ch)) {
      out.push({ k: 'op', v: ch })
      i++
      continue
    }
    throw new Error(`unexpected "${ch}" in ${src}`)
  }
  out.push({ k: 'end', v: '' })
  return out
}

function parse(src: string): Node {
  const toks = tokenize(src.startsWith('=') ? src.slice(1) : src)
  let p = 0
  const peek = () => toks[p]
  const next = () => toks[p++]
  const isOp = (v: string) => peek().k === 'op' && peek().v === v

  function cmp(): Node {
    let a = concat()
    while (peek().k === 'op' && ['=', '<>', '<', '>', '<=', '>='].includes(peek().v)) {
      const op = next().v
      a = { t: 'bin', op, a, b: concat() }
    }
    return a
  }
  function concat(): Node {
    let a = add()
    while (isOp('&')) {
      next()
      a = { t: 'bin', op: '&', a, b: add() }
    }
    return a
  }
  function add(): Node {
    let a = mul()
    while (isOp('+') || isOp('-')) {
      const op = next().v
      a = { t: 'bin', op, a, b: mul() }
    }
    return a
  }
  function mul(): Node {
    let a = pow()
    while (isOp('*') || isOp('/')) {
      const op = next().v
      a = { t: 'bin', op, a, b: pow() }
    }
    return a
  }
  function pow(): Node {
    let a = unary()
    while (isOp('^')) {
      next()
      a = { t: 'bin', op: '^', a, b: unary() }
    }
    return a
  }
  function unary(): Node {
    if (isOp('-') || isOp('+')) {
      const op = next().v
      return { t: 'un', op, a: unary() }
    }
    return postfix()
  }
  function postfix(): Node {
    let a = primary()
    while (isOp('%')) {
      next()
      a = { t: 'pct', a }
    }
    return a
  }
  function primary(): Node {
    const t = next()
    if (t.k === 'num') return { t: 'num', v: parseFloat(t.v) }
    if (t.k === 'str') return { t: 'str', v: t.v }
    if (t.k === 'ref') {
      const [a, b] = t.v.split(':')
      const A = parseAddr(a)
      if (b) {
        const B = parseAddr(b)
        return {
          t: 'range',
          sheet: t.sheet ?? null,
          r1: Math.min(A.row, B.row),
          c1: Math.min(A.col, B.col),
          r2: Math.max(A.row, B.row),
          c2: Math.max(A.col, B.col),
        }
      }
      return { t: 'ref', sheet: t.sheet ?? null, row: A.row, col: A.col }
    }
    if (t.k === 'id') {
      if (isOp('(')) {
        next()
        const args: Node[] = []
        if (isOp(')')) {
          next()
          return { t: 'fn', name: t.v.toUpperCase(), args }
        }
        for (;;) {
          if (isOp(',') || isOp(')')) args.push({ t: 'empty' })
          else args.push(cmp())
          if (isOp(',')) {
            next()
            continue
          }
          if (isOp(')')) {
            next()
            break
          }
          throw new Error('expected , or ) in ' + src)
        }
        return { t: 'fn', name: t.v.toUpperCase(), args }
      }
      return { t: 'name', v: t.v }
    }
    if (t.k === 'op' && t.v === '(') {
      const e = cmp()
      if (!isOp(')')) throw new Error('missing ) in ' + src)
      next()
      return e
    }
    throw new Error('parse error in ' + src)
  }

  const root = cmp()
  if (peek().k !== 'end') throw new Error('trailing input in ' + src)
  return root
}

// --------------------------------------------------------------- evaluation
const E_VALUE = new ErrVal('#VALUE!')
const E_DIV = new ErrVal('#DIV/0!')
const E_NAME = new ErrVal('#NAME?')
const E_REF = new ErrVal('#REF!')

function toNum(v: Val): number | ErrVal {
  if (v instanceof ErrVal) return v
  if (v === null) return 0
  if (typeof v === 'number') {
    if (!Number.isFinite(v)) return E_VALUE
    return v
  }
  if (typeof v === 'boolean') return v ? 1 : 0
  const s = v.trim()
  if (s === '') return 0
  const n = Number(s)
  return Number.isNaN(n) || !Number.isFinite(n) ? E_VALUE : n
}
function toBool(v: Val): boolean | ErrVal {
  if (v instanceof ErrVal) return v
  if (v === null) return false
  if (typeof v === 'boolean') return v
  if (typeof v === 'number') return v !== 0
  const u = v.toUpperCase()
  if (u === 'TRUE') return true
  if (u === 'FALSE') return false
  return E_VALUE
}
function toStr(v: Val): string | ErrVal {
  if (v instanceof ErrVal) return v
  if (v === null) return ''
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE'
  return String(v)
}
function rank(v: Val): number {
  return typeof v === 'number' ? 1 : typeof v === 'string' ? 2 : 3
}
function compare(a: Val, b: Val): number {
  if (a === null) a = typeof b === 'string' ? '' : typeof b === 'boolean' ? false : 0
  if (b === null) b = typeof a === 'string' ? '' : typeof a === 'boolean' ? false : 0
  const ra = rank(a)
  const rb = rank(b)
  if (ra !== rb) return ra < rb ? -1 : 1
  if (typeof a === 'string' && typeof b === 'string') {
    const x = a.toLowerCase()
    const y = b.toLowerCase()
    return x < y ? -1 : x > y ? 1 : 0
  }
  const x = Number(a)
  const y = Number(b)
  return x < y ? -1 : x > y ? 1 : 0
}

type Evaled = Val | { range: { sheet: string; r1: number; c1: number; r2: number; c2: number } }
const isRange = (v: Evaled): v is { range: any } =>
  typeof v === 'object' && v !== null && !(v instanceof ErrVal) && 'range' in v

const GLOBAL_AST_CACHE = new Map<string, Node>()

function getCachedAst(formula: string): Node {
  let ast = GLOBAL_AST_CACHE.get(formula)
  if (!ast) {
    ast = parse(formula)
    GLOBAL_AST_CACHE.set(formula, ast)
  }
  return ast
}

export type Overrides = Record<string, number | null>

export const DEFAULT_NAMED_CONSTANTS: Record<string, number> = {
  CONST_STEEL_FACTOR: 127388,
  CONST_SURFACE_FACTOR: 645.16,
}

export class Workbook {
  readonly sheets = new Map<string, SheetData>()
  readonly names: string[]
  private cache = new Map<string, Val>()
  private asts = new Map<string, Node>()
  private busy = new Set<string>()
  private overrides: Overrides = {}
  private templateDefaults: Record<string, number | null> = {}
  private formulaOverrides: Record<string, string> = {}
  private namedConstants: Record<string, number> = {}

  constructor(data: { sheets: SheetData[] }, namedConstants: Record<string, number> = {}) {
    for (const s of data.sheets) this.sheets.set(s.name, s)
    this.names = data.sheets.map((s) => s.name)
    this.namedConstants = { ...DEFAULT_NAMED_CONSTANTS, ...namedConstants }
  }

  setNamedConstants(constants: Record<string, number>) {
    this.namedConstants = { ...this.namedConstants, ...constants }
    this.cache.clear()
  }
  getNamedConstants(): Record<string, number> {
    return this.namedConstants
  }

  // ---- inputs -----------------------------------------------------------
  setOverrides(o: Overrides) {
    this.overrides = { ...o }
    this.cache.clear()
  }
  getOverrides(): Overrides {
    return this.overrides
  }
  setTemplateDefaults(defs: Record<string, number | null>) {
    this.templateDefaults = { ...defs }
    this.cache.clear()
  }
  getTemplateDefaults(): Record<string, number | null> {
    return this.templateDefaults
  }
  setFormulaOverrides(formulas: Record<string, string>) {
    this.formulaOverrides = { ...formulas }
    this.asts.clear()
    this.cache.clear()
  }
  getFormulaOverrides(): Record<string, string> {
    return this.formulaOverrides
  }
  static key(sheet: string, row: number, col: number) {
    return `${sheet}!${addr(row, col)}`
  }
  raw(sheet: string, row: number, col: number): RawCell {
    const k = Workbook.key(sheet, row, col)
    if (k in this.formulaOverrides) {
      return this.formulaOverrides[k]
    }
    const s = this.sheets.get(sheet)
    return s?.cells[row]?.[col] ?? null
  }
  isFormula(sheet: string, row: number, col: number) {
    const r = this.raw(sheet, row, col)
    return typeof r === 'string' && r.startsWith('=')
  }
  /** numeric constants are the editable "inputs" of the workbook */
  isInput(sheet: string, row: number, col: number) {
    return typeof this.raw(sheet, row, col) === 'number'
  }
  defaultInput(sheet: string, row: number, col: number): number | null {
    const k = Workbook.key(sheet, row, col)
    if (k in this.templateDefaults && this.templateDefaults[k] !== undefined) {
      return this.templateDefaults[k]
    }
    const r = this.raw(sheet, row, col)
    return typeof r === 'number' ? r : null
  }
  isOverridden(sheet: string, row: number, col: number) {
    return Workbook.key(sheet, row, col) in this.overrides
  }
  isFormulaOverridden(sheet: string, row: number, col: number) {
    return Workbook.key(sheet, row, col) in this.formulaOverrides
  }

  // ---- evaluation -------------------------------------------------------
  /** Value of a cell (after overrides). */
  value(sheet: string, row: number, col: number): Val {
    const k = `${sheet}\u0001${row},${col}`
    const hit = this.cache.get(k)
    if (hit !== undefined) return hit
    const ov = Workbook.key(sheet, row, col)
    let result: Val
    if (ov in this.overrides) {
      result = this.overrides[ov]
    } else if (ov in this.templateDefaults && this.templateDefaults[ov] !== null && this.templateDefaults[ov] !== undefined) {
      result = this.templateDefaults[ov]
    } else {
      const r = this.raw(sheet, row, col)
      if (typeof r === 'string' && r.startsWith('=')) {
        if (this.busy.has(k)) return E_REF
        this.busy.add(k)
        try {
          let ast = this.asts.get(k)
          if (!ast) {
            ast = getCachedAst(r)
            this.asts.set(k, ast)
          }
          const v = this.ev(ast, sheet)
          result = isRange(v) ? this.value(v.range.sheet, v.range.r1, v.range.c1) : v
          if (result === null) result = 0 // a formula pointing at an empty cell shows 0, like Excel
        } catch {
          result = E_NAME
        } finally {
          this.busy.delete(k)
        }
      } else {
        result = r
      }
    }
    this.cache.set(k, result)
    return result
  }

  private ev(n: Node, sheet: string): Evaled {
    switch (n.t) {
      case 'num':
        return n.v
      case 'str':
        return n.v
      case 'empty':
        return 0
      case 'name': {
        const up = n.v.toUpperCase()
        if (this.namedConstants[up] !== undefined) return this.namedConstants[up]
        if (this.namedConstants[n.v] !== undefined) return this.namedConstants[n.v]
        if (up === 'NO') return 'NO'
        if (up === 'YES') return 'YES'
        if (up === 'TRUE') return true
        if (up === 'FALSE') return false
        return E_NAME
      }
      case 'ref': {
        const sh = n.sheet ?? sheet
        if (!this.sheets.has(sh)) return E_REF
        return this.value(sh, n.row, n.col)
      }
      case 'range': {
        const sh = n.sheet ?? sheet
        if (!this.sheets.has(sh)) return E_REF
        return { range: { sheet: sh, r1: n.r1, c1: n.c1, r2: n.r2, c2: n.c2 } }
      }
      case 'un': {
        const a = this.scalar(this.ev(n.a, sheet))
        const x = toNum(a)
        if (x instanceof ErrVal) return x
        return n.op === '-' ? -x : x
      }
      case 'pct': {
        const x = toNum(this.scalar(this.ev(n.a, sheet)))
        return x instanceof ErrVal ? x : x / 100
      }
      case 'bin': {
        const a = this.scalar(this.ev(n.a, sheet))
        const b = this.scalar(this.ev(n.b, sheet))
        if (a instanceof ErrVal) return a
        if (b instanceof ErrVal) return b
        switch (n.op) {
          case '=':
            return compare(a, b) === 0
          case '<>':
            return compare(a, b) !== 0
          case '<':
            return compare(a, b) < 0
          case '>':
            return compare(a, b) > 0
          case '<=':
            return compare(a, b) <= 0
          case '>=':
            return compare(a, b) >= 0
          case '&': {
            const x = toStr(a)
            const y = toStr(b)
            if (x instanceof ErrVal) return x
            if (y instanceof ErrVal) return y
            return x + y
          }
        }
        const x = toNum(a)
        const y = toNum(b)
        if (x instanceof ErrVal) return x
        if (y instanceof ErrVal) return y
        switch (n.op) {
          case '+':
            return x + y
          case '-':
            return x - y
          case '*':
            return x * y
          case '/':
            return y === 0 ? E_DIV : x / y
          case '^':
            return Math.pow(x, y)
        }
        return E_VALUE
      }
      case 'fn':
        return this.fn(n.name, n.args, sheet)
    }
  }

  /** a range used where a scalar is expected → top-left cell (implicit intersection) */
  private scalar(v: Evaled): Val {
    return isRange(v) ? this.value(v.range.sheet, v.range.r1, v.range.c1) : v
  }

  private fn(name: string, args: Node[], sheet: string): Evaled {
    if (name === 'IF') {
      if (args.length < 2) return E_VALUE
      const c = toBool(this.scalar(this.ev(args[0], sheet)))
      if (c instanceof ErrVal) return c
      if (c) return this.scalar(this.ev(args[1], sheet))
      if (args.length < 3) return false
      return this.scalar(this.ev(args[2], sheet))
    }
    if (name === 'SUM') {
      let total = 0
      for (const a of args) {
        const v = this.ev(a, sheet)
        if (isRange(v)) {
          const { sheet: sh, r1, c1, r2, c2 } = v.range
          for (let r = r1; r <= r2; r++)
            for (let c = c1; c <= c2; c++) {
              const x = this.value(sh, r, c)
              if (x instanceof ErrVal) return x
              if (typeof x === 'number') total += x
            }
        } else {
          const x = toNum(v)
          if (x instanceof ErrVal) return x
          total += x
        }
      }
      return total
    }
    return E_NAME
  }

  // ---- convenience ------------------------------------------------------
  num(sheet: string, a: string): number {
    const { row, col } = parseAddr(a)
    const v = this.value(sheet, row, col)
    return typeof v === 'number' ? v : 0
  }
  get(sheet: string, a: string): Val {
    const { row, col } = parseAddr(a)
    return this.value(sheet, row, col)
  }
  text(sheet: string, a: string): string {
    const v = this.get(sheet, a)
    return v === null ? '' : String(v)
  }
}
