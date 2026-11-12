import * as XLSX from 'xlsx'
import { Workbook, addr, numToCol } from './engine'

export interface C2CExportData {
  profit: number
  total: number
  ratio: number
  costBase?: number
  labour?: number
  steelWeight?: number
}

/**
 * Exports a full multi-sheet Workbook instance to a genuine .xlsx Excel spreadsheet file.
 * Automatically injects the 3-line Cut-to-Cut widget calculations into "Total cal" at E26:G28.
 */
export function exportWorkbookToXLSX(
  wb: Workbook,
  filename: string,
  c2cData?: C2CExportData
) {
  const book = XLSX.utils.book_new()

  for (const sheetName of wb.names) {
    const sheetData = wb.sheets.get(sheetName)
    if (!sheetData) continue

    const rows = sheetData.rows
    const cols = sheetData.cols
    const aoa: (string | number | boolean | null)[][] = []

    for (let r = 0; r < rows; r++) {
      const rowArr: (string | number | boolean | null)[] = []
      for (let c = 0; c < cols; c++) {
        // Inject 3-line C2C Widget at E26:G28 on "Total cal" sheet
        if (sheetName === 'Total cal' && c2cData) {
          // Row 26 in Excel is r = 25 (0-indexed)
          if (r === 25 && c === 4) { rowArr.push('C2C PROFIT'); continue }
          if (r === 25 && c === 5) { rowArr.push(Math.round(c2cData.profit)); continue }
          if (r === 25 && c === 6) { rowArr.push('Actual Before GST - C2C Cost Base'); continue }

          // Row 27 in Excel is r = 26
          if (r === 26 && c === 4) { rowArr.push('C2C TOTAL'); continue }
          if (r === 26 && c === 5) { rowArr.push(Math.round(c2cData.total)); continue }
          if (r === 26 && c === 6) { rowArr.push('Profit + C2C Labour'); continue }

          // Row 28 in Excel is r = 27
          if (r === 27 && c === 4) { rowArr.push('C2C RATIO'); continue }
          if (r === 27 && c === 5) { rowArr.push(Number(c2cData.ratio.toFixed(2))); continue }
          if (r === 27 && c === 6) { rowArr.push('Total / C2C Weight (Rs/kg)'); continue }
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
        // Do not overwrite our injected C2C cells
        if (sheetName === 'Total cal' && c2cData && r >= 25 && r <= 27 && c >= 4 && c <= 6) {
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

    // Set readable column widths
    const colWidths = []
    for (let c = 0; c < cols; c++) {
      let maxLen = 6
      for (let r = 0; r < Math.min(rows, 60); r++) {
        const val = aoa[r]?.[c]
        if (val !== null && val !== undefined) {
          const s = String(val)
          if (s.length > maxLen) maxLen = Math.min(40, s.length)
        }
      }
      colWidths.push({ wch: maxLen + 3 })
    }
    ws['!cols'] = colWidths

    XLSX.utils.book_append_sheet(book, ws, sheetName.slice(0, 31))
  }

  const cleanFileName = filename.endsWith('.xlsx') ? filename : `${filename}.xlsx`
  XLSX.writeFile(book, cleanFileName)
}
