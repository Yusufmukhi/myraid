import { useEffect, useMemo, useRef, useState } from 'react'
import { addr, CellStyle, ErrVal, Overrides, parseAddr, Workbook } from '../lib/engine'
import { fmtCell } from '../lib/format'
import { EXTRA_EDITABLE } from '../lib/inputs'
import { Icon, NumField } from './ui'

interface Props {
  wb: Workbook
  overrides: Overrides
  setInput: (key: string, value: number | null | undefined) => void
}

interface Span {
  rs: number
  cs: number
}

function textOn(hex: string): string {
  const [r, g, b] = [0, 2, 4].map((i) => parseInt(hex.slice(i, i + 2), 16) / 255)
  const lum = 0.2126 * r + 0.7152 * g + 0.0722 * b
  return lum > 0.45 ? '#0f172a' : '#ffffff'
}

const SHEET_NOTES: Record<string, string> = {
  'Total cal': 'Summary & quotation',
}

export function SheetViewer({ wb, overrides, setInput }: Props) {
  const [name, setName] = useState('Total cal')
  const [sel, setSel] = useState<{ r: number; c: number } | null>(null)
  const sheet = wb.sheets.get(name)!
  const tabsRef = useRef<HTMLDivElement>(null)

  // merged cells → spans + covered set
  const { spans, covered } = useMemo(() => {
    const spans = new Map<string, Span>()
    const covered = new Set<string>()
    for (const m of sheet.merges ?? []) {
      const [a, b] = m.split(':')
      if (!b) continue
      const A = parseAddr(a)
      const B = parseAddr(b)
      spans.set(`${A.row},${A.col}`, { rs: B.row - A.row + 1, cs: B.col - A.col + 1 })
      for (let r = A.row; r <= B.row; r++)
        for (let c = A.col; c <= B.col; c++) if (r !== A.row || c !== A.col) covered.add(`${r},${c}`)
    }
    return { spans, covered }
  }, [sheet])

  // column widths from content
  const widths = useMemo(() => {
    const w: number[] = []
    for (let c = 0; c < sheet.cols; c++) {
      let max = 3
      for (let r = 0; r < sheet.rows; r++) {
        if (spans.has(`${r},${c}`) || covered.has(`${r},${c}`)) continue
        const raw = sheet.cells[r][c]
        if (raw === null) continue
        const v = wb.value(name, r, c)
        const len = fmtCell(v).length
        if (len > max) max = len
      }
      w.push(Math.max(46, Math.min(230, max * 7.4 + 16)))
    }
    return w
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [sheet, spans, covered])

  useEffect(() => setSel(null), [name])

  const changedCount = useMemo(
    () => Object.keys(overrides).filter((k) => k.startsWith(name + '!')).length,
    [overrides, name],
  )

  const selVal = sel ? wb.value(name, sel.r, sel.c) : null
  const selRaw = sel ? wb.raw(name, sel.r, sel.c) : null
  const selKey = sel ? Workbook.key(name, sel.r, sel.c) : ''
  const selEditable = sel ? wb.isInput(name, sel.r, sel.c) || EXTRA_EDITABLE.has(selKey) : false

  const isEditable = (r: number, c: number) =>
    wb.isInput(name, r, c) || EXTRA_EDITABLE.has(Workbook.key(name, r, c))

  return (
    <div className="sheets card">
      {/* formula bar */}
      <div className="fbar">
        <div className="fbar-addr">{sel ? addr(sel.r, sel.c) : '—'}</div>
        <div className="fbar-fx">fx</div>
        <div className="fbar-body">
          {sel === null ? (
            <span className="muted">
              Select a cell. <b className="inputdot" /> Blue cells are inputs you can change – everything else is calculated.
            </span>
          ) : selEditable ? (
            <>
              <NumField
                className="fbar-input"
                value={selKey in overrides ? overrides[selKey] : (selRaw as number | null)}
                onChange={(v) => setInput(selKey, v)}
              />
              {selKey in overrides && (
                <button className="btn ghost sm" onClick={() => setInput(selKey, undefined)}>
                  {Icon.reset} Reset to {String(selRaw ?? 'blank')}
                </button>
              )}
            </>
          ) : wb.isFormula(name, sel.r, sel.c) ? (
            <>
              <code>{String(selRaw)}</code>
              <span className="fbar-result">= {fmtCell(selVal)}</span>
            </>
          ) : (
            <span>{fmtCell(selVal)}</span>
          )}
        </div>
        {changedCount > 0 && <span className="chip warn">{changedCount} changed on this sheet</span>}
      </div>

      {/* grid */}
      <div className="grid-scroll">
        <table className="xl">
          <colgroup>
            <col style={{ width: 40 }} />
            {widths.map((w, i) => (
              <col key={i} style={{ width: w }} />
            ))}
          </colgroup>
          <thead>
            <tr>
              <th className="corner" />
              {widths.map((_, c) => (
                <th key={c} className={sel?.c === c ? 'hl' : ''}>
                  {addr(0, c).replace(/\d+/, '')}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {sheet.cells.map((row, r) => (
              <tr key={r}>
                <th className={sel?.r === r ? 'hl' : ''}>{r + 1}</th>
                {row.map((raw, c) => {
                  const k = `${r},${c}`
                  if (covered.has(k)) return null
                  const span = spans.get(k)
                  const st: CellStyle | undefined = sheet.styles?.[addr(r, c)]
                  const edit = isEditable(r, c)
                  const key = Workbook.key(name, r, c)
                  const changed = key in overrides
                  const v = raw === null && !changed ? null : wb.value(name, r, c)
                  const text = fmtCell(v)
                  const isNum = typeof v === 'number'
                  const cls = [
                    isNum ? 'n' : '',
                    v instanceof ErrVal ? 'err' : '',
                    edit ? 'inp' : '',
                    changed ? 'chg' : '',
                    sel?.r === r && sel?.c === c ? 'sel' : '',
                    st?.b ? 'b' : '',
                    st?.i ? 'it' : '',
                    st?.al === 'c' ? 'ac' : st?.al === 'r' ? 'ar' : st?.al === 'l' ? 'al' : '',
                  ]
                    .filter(Boolean)
                    .join(' ')
                  const style: React.CSSProperties = {}
                  if (st?.bg && !edit) {
                    style.background = '#' + st.bg
                    style.color = textOn(st.bg) // always readable on the Excel fill colour
                  } else if (st?.fg && !edit) style.color = '#' + st.fg
                  return (
                    <td
                      key={c}
                      className={cls}
                      style={style}
                      rowSpan={span?.rs}
                      colSpan={span?.cs}
                      title={wb.isFormula(name, r, c) ? String(raw) : undefined}
                      onClick={() => setSel({ r, c })}
                    >
                      {text}
                    </td>
                  )
                })}
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      {/* sheet tabs */}
      <div className="tabs-bottom" ref={tabsRef}>
        {wb.names.map((n) => (
          <button key={n} className={n === name ? 'on' : ''} onClick={() => setName(n)} title={SHEET_NOTES[n]}>
            {n}
            {Object.keys(overrides).some((k) => k.startsWith(n + '!')) && <i className="dot" />}
          </button>
        ))}
      </div>
    </div>
  )
}
