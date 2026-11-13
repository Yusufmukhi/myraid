import _XLSX from 'xlsx-js-style'
const XLSX: any = (_XLSX as any)?.utils ? _XLSX : ((_XLSX as any)?.default || _XLSX)
import { Workbook, addr } from './engine'
import { getSheetMeta, CellStyle } from './excelStyles'

export interface C2CExportData {
  profit: number
  total: number
  ratio: number
  costBase?: number
  labour?: number
  steelWeight?: number
  subtotalBeforeGst?: number
  setsQty?: number
}

export interface ExcelExportOptions {
  setsQty?: number
  subtotalBeforeGst?: number
}

/**
 * Exports a full multi-sheet Workbook instance to a genuine styled .xlsx Excel spreadsheet file.
 * Preserves the exact color scheme, column widths, and merged cells from the original exam.xlsx file.
 * - Yellow (FFFF00): Specifications, dimensions, thicknesses, and price inputs
 * - Green (92D050): Calculated unit rates, steel weights, and totals
 * - Red (FF0000): Total amounts and special sheet constraints
 * - Blue (00B0F0): Ex-Works subtotal and section headers
 * - Displays tables with native gridlines visible.
 * - Places Subtotal Before GST at Total cal!S5 and C2C Profit at Total cal!S6.
 * - Multiplies model unit quantities by setsQty on Total cal (G4, G8, G12, G16, G20).
 * - Writes SET label in O22 and sets quantity value in P22.
 * - Corrects Shearing size row 45 stiffener gauge reference to Total cal!J3 and zeros H45 if stiffener is 0.
 */
export function exportWorkbookToXLSX(
  wb: Workbook,
  filename: string,
  c2cData?: C2CExportData,
  options?: ExcelExportOptions
) {
  const book = XLSX.utils.book_new()
  const setsQty = Math.max(1, Number(options?.setsQty ?? c2cData?.setsQty) || 1)
  const subtotalBeforeGst = options?.subtotalBeforeGst ?? c2cData?.subtotalBeforeGst ?? 0

  for (const sheetName of wb.names) {
    const sheetData = wb.sheets.get(sheetName)
    if (!sheetData) continue

    const normSheet = sheetName.toLowerCase().replace(/[^a-z0-9]/g, '')
    const isShearing = normSheet.includes('shearing')

    const rows = sheetData.rows
    const cols = sheetData.cols
    const aoa: (string | number | boolean | null)[][] = []

    for (let r = 0; r < rows; r++) {
      const rowArr: (string | number | boolean | null)[] = []
      for (let c = 0; c < cols; c++) {
        // Total cal sheet customizations
        if (sheetName === 'Total cal') {
          // Model quantities: multiply unit quantity by setsQty when exporting
          // G4 (DFMU: r=3), G8 (SFMU: r=7), G12 (SFFU: r=11), G16 (DFFU: r=15), G20 (DFMU-D: r=19)
          if (c === 6 && [3, 7, 11, 15, 19].includes(r)) {
            const rawVal = wb.value(sheetName, r, c) ?? wb.raw(sheetName, r, c)
            const numVal = typeof rawVal === 'number' ? rawVal : Number(rawVal) || 0
            rowArr.push(setsQty > 1 ? numVal * setsQty : numVal)
            continue
          }

          // Row 5: R5 (c=17) Subtotal Before GST label, S5 (c=18) value
          if (r === 4 && c === 17) { rowArr.push('Subtotal Before GST'); continue }
          if (r === 4 && c === 18) { rowArr.push(subtotalBeforeGst || Number(wb.value(sheetName, 26, 13)) || 0); continue }

          // Row 6: R6 (c=17) C2C Profit label, S6 (c=18) value
          if (r === 5 && c === 17) { rowArr.push('C2C Profit'); continue }
          if (r === 5 && c === 18) { rowArr.push(c2cData ? Math.round(c2cData.profit) : 0); continue }

          // Row 22: O22 (c=14) SET label, P22 (c=15) sets quantity value
          if (r === 21 && c === 14) { rowArr.push('SET'); continue }
          if (r === 21 && c === 15) { rowArr.push(setsQty); continue }
        }

        // Shearing size sheet row 45 (r=44): stiffener gauge and quantity corrections
        if (isShearing && r === 44) {
          if (c === 5) {
            // F45: shelf stiffener gauge from Total cal!J3 (r=2, c=9)
            const stiffenerVal = wb.value('Total cal', 2, 9)
            const numOrStr = typeof stiffenerVal === 'number' || typeof stiffenerVal === 'string' ? stiffenerVal : 0
            rowArr.push(numOrStr)
            continue
          }
          if (c === 7) {
            // H45: if stiffener is 0, show 0, else show Sheet1!I82
            const stiffenerVal = wb.value('Total cal', 2, 9)
            if (!stiffenerVal || Number(stiffenerVal) === 0) {
              rowArr.push(0)
            } else {
              const h45Val = wb.value(sheetName, 44, 7) ?? wb.value('Sheet1', 81, 8) ?? 0
              const numOrStr = typeof h45Val === 'number' || typeof h45Val === 'string' ? h45Val : 0
              rowArr.push(numOrStr)
            }
            continue
          }
        }

        const evaluatedVal = wb.value(sheetName, r, c)
        if (typeof evaluatedVal === 'number') {
          rowArr.push(evaluatedVal)
        } else if (typeof evaluatedVal === 'string') {
          rowArr.push(evaluatedVal)
        } else if (typeof evaluatedVal === 'boolean') {
          rowArr.push(evaluatedVal)
        } else {
          const raw = wb.raw(sheetName, r, c)
          if (typeof raw === 'number' || typeof raw === 'string') {
            rowArr.push(raw)
          } else {
            rowArr.push(null)
          }
        }
      }
      aoa.push(rowArr)
    }

    const ws = XLSX.utils.aoa_to_sheet(aoa)

    // Restore formulas without the leading '=' so Excel evaluates them natively
    for (let r = 0; r < rows; r++) {
      for (let c = 0; c < cols; c++) {
        // Do not overwrite our special cells on Total cal
        if (sheetName === 'Total cal') {
          if (c === 6 && [3, 7, 11, 15, 19].includes(r)) continue
          if ((r === 4 || r === 5) && (c === 17 || c === 18)) continue
          if (r === 21 && (c === 14 || c === 15)) continue
        }

        // Shearing size sheet row 45 dynamic formula corrections
        if (isShearing && r === 44 && c === 5) {
          const cellRef = addr(r, c)
          if (ws[cellRef]) {
            ws[cellRef].f = "'Total cal'!J3"
          }
          continue
        }
        if (isShearing && r === 44 && c === 7) {
          const cellRef = addr(r, c)
          if (ws[cellRef]) {
            ws[cellRef].f = "IF('Total cal'!J3=0,0,Sheet1!I82)"
          }
          continue
        }

        const raw = wb.raw(sheetName, r, c)
        if (typeof raw === 'string' && raw.startsWith('=')) {
          const cellRef = addr(r, c)
          if (ws[cellRef]) {
            ws[cellRef].f = raw.slice(1)
          }
        }
      }
    }

    // Load authentic template styling, column widths, and merges from original exam.xlsx
    const meta = getSheetMeta(sheetName)

    // Apply exact column widths (preventing giant / stretched cells)
    if (meta?.cols && meta.cols.length > 0) {
      ws['!cols'] = meta.cols.map((col: any) => ({ ...col }))
    } else {
      // Natural compact fallback: clamp between 10 and 18 chars, ignoring long multi-column titles
      const fallbackCols = []
      for (let c = 0; c < Math.min(cols, 20); c++) {
        let maxLen = 8
        for (let r = 0; r < Math.min(rows, 35); r++) {
          const val = aoa[r]?.[c]
          if (val !== null && val !== undefined) {
            const s = String(val)
            if (s.length > 0 && s.length <= 18) {
              if (s.length > maxLen) maxLen = s.length
            }
          }
        }
        fallbackCols.push({ wch: Math.min(18, Math.max(10, maxLen + 2)) })
      }
      ws['!cols'] = fallbackCols
    }

    // Ensure Total cal columns include O, P, R, S with clean readable widths
    if (sheetName === 'Total cal') {
      const currentCols = ws['!cols'] || []
      while (currentCols.length < 19) {
        currentCols.push({ wch: 12 })
      }
      currentCols[14] = { wch: 9 }   // Col O: "SET"
      currentCols[15] = { wch: 9 }   // Col P: Sets count
      currentCols[17] = { wch: 22 }  // Col R: Labels
      currentCols[18] = { wch: 18 }  // Col S: Values
      ws['!cols'] = currentCols
    }

    // Apply merged cells if present
    if (meta?.merges && meta.merges.length > 0) {
      ws['!merges'] = meta.merges
    }

    // Apply authentic colors & styles (Yellow FFFF00, Green 92D050, Red FF0000, Blue 00B0F0)
    if (meta?.styles) {
      for (const [cellRef, style] of Object.entries(meta.styles)) {
        if (ws[cellRef]) {
          ws[cellRef].s = style
        }
      }
    }

    // Style Total cal custom cells (R5, S5, R6, S6, O22, P22)
    if (sheetName === 'Total cal') {
      const thinBorder = {
        top: { style: 'thin', color: { rgb: 'CBD5E1' } },
        bottom: { style: 'thin', color: { rgb: 'CBD5E1' } },
        left: { style: 'thin', color: { rgb: 'CBD5E1' } },
        right: { style: 'thin', color: { rgb: 'CBD5E1' } },
      }

      // R5: "Subtotal Before GST" label
      if (ws['R5']) {
        ws['R5'].s = {
          font: { bold: true, color: { rgb: '1E293B' }, sz: 10 },
          fill: { fgColor: { rgb: 'F1F5F9' } },
          alignment: { horizontal: 'right', vertical: 'center' },
          border: thinBorder,
        }
      }

      // S5: Sale value is Subtotal Before GST
      if (ws['S5']) {
        ws['S5'].s = {
          font: { bold: true, color: { rgb: '0F172A' }, sz: 11 },
          fill: { fgColor: { rgb: 'E2EFDA' } }, // Clean soft green
          alignment: { horizontal: 'right', vertical: 'center' },
          border: thinBorder,
          numFmt: '₹#,##0.00;[Red]-₹#,##0.00;"-"',
        }
      }

      // R6: "C2C Profit" label
      if (ws['R6']) {
        ws['R6'].s = {
          font: { bold: true, color: { rgb: '1E293B' }, sz: 10 },
          fill: { fgColor: { rgb: 'F1F5F9' } },
          alignment: { horizontal: 'right', vertical: 'center' },
          border: thinBorder,
        }
      }

      // S6: C2C Profit value
      if (ws['S6']) {
        ws['S6'].s = {
          font: { bold: true, color: { rgb: '166534' }, sz: 11 },
          fill: { fgColor: { rgb: '92D050' } }, // Classic Excel totals green
          alignment: { horizontal: 'right', vertical: 'center' },
          border: thinBorder,
          numFmt: '₹#,##0;[Red]-₹#,##0;"-"',
        }
      }

      // O22: "SET" label
      if (ws['O22']) {
        ws['O22'].s = {
          font: { bold: true, color: { rgb: '1E3A8A' }, sz: 10 },
          fill: { fgColor: { rgb: 'DBEAFE' } },
          alignment: { horizontal: 'center', vertical: 'center' },
          border: thinBorder,
        }
      }

      // P22: Sets Quantity Value
      if (ws['P22']) {
        ws['P22'].s = {
          font: { bold: true, color: { rgb: '1E3A8A' }, sz: 11 },
          fill: { fgColor: { rgb: 'EFF6FF' } },
          alignment: { horizontal: 'center', vertical: 'center' },
          border: thinBorder,
        }
      }
    }

    // Enable native Excel grid lines across all sheets so exported tables display clearly
    ws['!views'] = [{ showGridLines: true }]

    XLSX.utils.book_append_sheet(book, ws, sheetName.slice(0, 31))
  }

  const cleanFileName = filename.endsWith('.xlsx') ? filename : `${filename}.xlsx`
  XLSX.writeFile(book, cleanFileName)
}
