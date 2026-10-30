import { INPUT_GROUPS } from '../lib/inputs'
import { Overrides, Workbook } from '../lib/engine'
import { count, inr, inr2, kg } from '../lib/format'
import { Summary } from '../lib/model'
import { Icon, NumField } from './ui'

interface Props {
  wb: Workbook
  overrides: Overrides
  summary: Summary
  setInput: (key: string, value: number | null | undefined) => void
}

function defaultOf(wb: Workbook, key: string): number | null {
  const [sheet, a] = [key.slice(0, key.lastIndexOf('!')), key.slice(key.lastIndexOf('!') + 1)]
  const m = /^([A-Z]+)(\d+)$/.exec(a)!
  let c = 0
  for (const ch of m[1]) c = c * 26 + ch.charCodeAt(0) - 64
  return wb.defaultInput(sheet, parseInt(m[2], 10) - 1, c - 1)
}

export function Calculator({ wb, overrides, summary: s, setInput }: Props) {
  return (
    <div className="calc-grid">
      {/* ------------------------------------------------------- inputs */}
      <aside className="inputs">
        {INPUT_GROUPS.map((g) => (
          <section key={g.title} className="card input-group">
            <header>
              <h3>{g.title}</h3>
              {g.note && <span className="muted tiny">{g.note}</span>}
            </header>
            <div className="fields">
              {g.items.map((it) => {
                const def = defaultOf(wb, it.key)
                const changed = it.key in overrides
                const val = changed ? overrides[it.key] : def
                return (
                  <label key={it.key} className={`field ${changed ? 'changed' : ''}`}>
                    <span className="field-label">
                      {it.label}
                      {it.hint && <em>{it.hint}</em>}
                    </span>
                    <span className="field-input">
                      <NumField
                        value={val}
                        step={it.step}
                        min={it.min}
                        placeholder={def === null ? '0' : undefined}
                        onChange={(v) => setInput(it.key, v)}
                      />
                      {it.unit && <span className="unit">{it.unit}</span>}
                      {changed && (
                        <button
                          type="button"
                          className="icon-btn tiny-btn"
                          title={`Reset to ${def ?? 'blank'}`}
                          onClick={() => setInput(it.key, undefined)}
                        >
                          {Icon.reset}
                        </button>
                      )}
                    </span>
                  </label>
                )
              })}
            </div>
          </section>
        ))}
      </aside>

      {/* ------------------------------------------------------ results */}
      <div className="results">
        <div className="kpis">
          <div className="kpi primary">
            <span>Total (incl. 18% GST)</span>
            <strong>{inr(s.sales)}</strong>
            <small>{inr2(s.beforeGst)} before GST</small>
          </div>
          <div className="kpi">
            <span>Steel weight</span>
            <strong>{kg(s.totalWeight)}</strong>
            <small>{kg(s.weightWithMachine)} with machine parts</small>
          </div>
          <div className="kpi">
            <span>Bodies</span>
            <strong>{count(s.bodies)}</strong>
            <small>{count(s.compartments)} compartments</small>
          </div>
          <div className="kpi">
            <span>Powder coating</span>
            <strong>{inr(s.powderTotal)}</strong>
            <small>included in unit rates</small>
          </div>
        </div>

        <section className="card">
          <header className="card-head">
            <h3>Quotation</h3>
            <span className="muted tiny">Total cal sheet</span>
          </header>
          <div className="table-wrap">
            <table className="tbl">
              <thead>
                <tr>
                  <th>Item</th>
                  <th className="r">Qty</th>
                  <th className="r">Rate / unit</th>
                  <th className="r">Amount</th>
                </tr>
              </thead>
              <tbody>
                {s.models.map((m) => (
                  <tr key={m.code} className={m.qty === 0 ? 'dim' : ''}>
                    <td>
                      <b>{m.code}</b> {m.note && <span className="muted">{m.note}</span>}
                    </td>
                    <td className="r">{count(m.qty)}</td>
                    <td className="r">{inr2(m.unitRate)}</td>
                    <td className="r">{inr2(m.amount)}</td>
                  </tr>
                ))}
                <tr className="sub">
                  <td colSpan={3}>Total</td>
                  <td className="r">{inr2(s.subtotal)}</td>
                </tr>
                <tr>
                  <td colSpan={3}>Installation</td>
                  <td className="r">{inr2(s.installation)}</td>
                </tr>
                <tr>
                  <td colSpan={3}>Transport</td>
                  <td className="r">{inr2(s.transport)}</td>
                </tr>
                <tr>
                  <td colSpan={3}>Accommodation</td>
                  <td className="r">{inr2(s.accommodation ?? s.accessories)}</td>
                </tr>
                <tr>
                  <td colSpan={3}>PF (20% of total)</td>
                  <td className="r">{inr2(s.pf)}</td>
                </tr>
                <tr className="sub">
                  <td colSpan={3}>Before GST</td>
                  <td className="r">{inr2(s.beforeGst)}</td>
                </tr>
                <tr>
                  <td colSpan={3}>GST 18%</td>
                  <td className="r">{inr2(s.gst)}</td>
                </tr>
                <tr className="grand">
                  <td colSpan={3}>Sales (grand total)</td>
                  <td className="r">{inr2(s.sales)}</td>
                </tr>
              </tbody>
            </table>
          </div>
        </section>

        <div className="cards-3">
          <section className="card">
            <header className="card-head">
              <h3>Steel</h3>
            </header>
            <dl className="kv">
              <dt>Body &amp; shelves</dt>
              <dd>{kg(s.steel.sp)}</dd>
              <dt>Wastage 7%</dt>
              <dd>{kg(s.steel.spWaste)}</dd>
              <dt className="strong">Total body steel</dt>
              <dd className="strong">{kg(s.steel.spTotal)}</dd>
              <dt>Cost @ {inr(s.steel.spRate)}/kg</dt>
              <dd>{inr(s.steel.spCost)}</dd>
              <dt>Trolley steel</dt>
              <dd>{kg(s.trolley.kg)}</dd>
              <dt>Wastage 7%</dt>
              <dd>{kg(s.trolley.waste)}</dd>
              <dt className="strong">Total trolley steel</dt>
              <dd className="strong">{kg(s.trolley.total)}</dd>
              <dt>Cost @ {inr(s.trolley.rate)}/kg</dt>
              <dd>{inr(s.trolley.cost)}</dd>
              <dt className="strong">Total steel cost</dt>
              <dd className="strong">{inr(s.steelCost)}</dd>
            </dl>
          </section>

          <section className="card">
            <header className="card-head">
              <h3>Labour</h3>
            </header>
            <dl className="kv">
              <dt>Body steel @ {inr(s.steel.spLabourRate)}/kg</dt>
              <dd>{inr(s.steel.spLabour)}</dd>
              <dt>Trolley steel @ {inr(s.trolley.labourRate)}/kg</dt>
              <dd>{inr(s.trolley.labour)}</dd>
              <dt>Steel &amp; trolley labour</dt>
              <dd>{inr(s.steelLabour)}</dd>
              <dt>Trolley fitting charges</dt>
              <dd>{inr(s.fittingLabour)}</dd>
              <dt className="strong">Labour subtotal</dt>
              <dd className="strong">{inr(s.labourSubtotal)}</dd>
              <dt>
                Clips ({count(s.clipQty)} × {inr(s.clipRate)})
              </dt>
              <dd>{inr(s.clips)}</dd>
              <dt className="strong">Total labour</dt>
              <dd className="strong">{inr(s.labourTotal)}</dd>
              <dt>
                Per body ({count(s.bodies)} × {inr(s.perBodyLabour)})
              </dt>
              <dd>{inr(s.bodyLabourTotal)}</dd>
            </dl>
          </section>

          <section className="card">
            <header className="card-head">
              <h3>Powder coating</h3>
            </header>
            <dl className="kv">
              {s.models.map((m) => (
                <div className="kv-row" key={m.code}>
                  <dt>
                    {m.code} × {count(m.qty)}
                  </dt>
                  <dd>{inr(m.powder)}</dd>
                </div>
              ))}
              <dt className="strong">Total</dt>
              <dd className="strong">{inr(s.powderTotal)}</dd>
            </dl>
            <header className="card-head sub-head">
              <h3>Door</h3>
            </header>
            <dl className="kv">
              <dt>Per door</dt>
              <dd>{inr2(s.door.perDoor)}</dd>
              <dt>Door labour</dt>
              <dd>{inr2(s.door.labourPerDoor)}</dd>
            </dl>
          </section>
        </div>
      </div>
    </div>
  )
}
