/** The inputs that live on the "Total cal" sheet – shown as a friendly form. */
export interface InputDef {
  key: string // "Sheet!A1"
  label: string
  unit?: string
  hint?: string
  step?: number
  min?: number
}
export interface InputGroup {
  title: string
  note?: string
  items: InputDef[]
}

const T = 'Total cal!'

export const INPUT_GROUPS: InputGroup[] = [
  {
    title: 'Cabinet size',
    note: 'Main dimensions used by every sheet',
    items: [
      { key: T + 'A3', label: 'Rows', hint: 'ROW – up to 14', step: 2, min: 2 },
      { key: T + 'B3', label: 'Columns', hint: 'CLOUM', step: 1, min: 1 },
      { key: T + 'C3', label: 'Height', unit: 'mm', hint: 'HT', step: 10, min: 0 },
      { key: T + 'D3', label: 'Width', unit: 'mm', hint: 'W', step: 10, min: 0 },
      { key: T + 'E3', label: 'Depth (DFMU / DFFU)', unit: 'mm', hint: 'D', step: 5, min: 0 },
      { key: T + 'E7', label: 'Depth (SFMU)', unit: 'mm', step: 5, min: 0 },
      { key: T + 'E11', label: 'Depth (SFFU)', unit: 'mm', step: 5, min: 0 },
      { key: T + 'I3', label: 'Track', hint: 'TRACK', step: 1, min: 0 },
    ],
  },
  {
    title: 'Sheet thickness',
    items: [
      { key: T + 'F3', label: 'Body', unit: 'mm', hint: 'body G', step: 0.1, min: 0 },
      { key: T + 'G3', label: 'Shelf', unit: 'mm', hint: 'shelf G', step: 0.1, min: 0 },
      { key: T + 'J3', label: 'Shelf stiffener', unit: 'mm', step: 0.1, min: 0 },
      { key: T + 'L3', label: 'Trolley', unit: 'mm', hint: 'trolly G', step: 0.1, min: 0 },
    ],
  },
  {
    title: 'Rates',
    items: [
      { key: T + 'H3', label: 'Steel rate', unit: '₹/kg', hint: 'RATE', step: 1, min: 0 },
      { key: T + 'M3', label: 'Trolley steel rate', unit: '₹/kg', step: 1, min: 0 },
      { key: T + 'B49', label: 'Labour per body', unit: '₹', step: 10, min: 0 },
    ],
  },
  {
    title: 'Quantities',
    note: 'How many of each unit are in this order',
    items: [
      { key: T + 'G4', label: 'DFMU', unit: 'nos', step: 1, min: 0 },
      { key: T + 'G8', label: 'SFMU', unit: 'nos', step: 1, min: 0 },
      { key: T + 'G12', label: 'SFFU', unit: 'nos', step: 1, min: 0 },
      { key: T + 'G16', label: 'DFFU', unit: 'nos', step: 1, min: 0 },
      { key: T + 'G20', label: 'DFMU with door', unit: 'nos', step: 1, min: 0 },
    ],
  },
  {
    title: 'Extra charges',
    note: 'Added before PF and GST',
    items: [
      { key: T + 'N24', label: 'Transport', unit: '₹', step: 100, min: 0 },
      { key: T + 'N25', label: 'Accessories', unit: '₹', step: 100, min: 0 },
    ],
  },
]

/** Empty cells that are meant to be typed into (they are blank in the original workbook). */
export const EXTRA_EDITABLE = new Set<string>([T + 'N24', T + 'N25', T + 'C26'])
