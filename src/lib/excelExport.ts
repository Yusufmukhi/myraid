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
}

/**
 * Exports a full multi-sheet Workbook instance to a genuine styled .xlsx Excel spreadsheet file.
 * Preserves the exact color scheme, column widths, and merged cells from the original exam.xlsx file.
 * - Yellow (FFFF00): Specifications, dimensions, thicknesses, and price inputs
 * - Green (92D050): Calculated unit rates, steel weights, and totals
 * - Red (FF0000): Total amounts and special sheet constraints
 * - Blue (00B0F0): Ex-Works subtotal and section headers
 * - Automatically injects and styles the 3-line Cut-to-Cut widget calculations into "Total cal" at E26:G28.
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

    // Load authentic template styling, column widths, and merges from original exam.xlsx
    const meta = getSheetMeta(sheetName)

    // Apply exact column widths (preventing giant / stretched cells)
    if (meta?.cols && meta.cols.length > 0) {
      ws['!cols'] = meta.cols
    } else {
      // Natural compact fallback: clamp between 10 and 18 chars, ignoring long multi-column titles
      const fallbackCols = []
      for (let c = 0; c < Math.min(cols, 20); c++) {
        let maxLen = 8
        for (let r = 0; r < Math.min(rows, 35); r++) {
          const val = aoa[r]?.[c]
          if (val !== null && val !== undefined) {
            const s = String(val)
            // Ignore long paragraphs or banner labels so they don't stretch single columns
            if (s.length > 0 && s.length <= 18) {
              if (s.length > maxLen) maxLen = s.length
            }
          }
        }
        fallbackCols.push({ wch: Math.min(18, Math.max(10, maxLen + 2)) })
      }
      ws['!cols'] = fallbackCols
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

    // Highlight and style injected Cut-to-Cut 3-Line Widget on "Total cal" at E26:G28
    if (sheetName === 'Total cal' && c2cData) {
      const widgetBorder = {
        top: { style: 'thin', color: { rgb: '000000' } },
        bottom: { style: 'thin', color: { rgb: '000000' } },
        left: { style: 'thin', color: { rgb: '000000' } },
        right: { style: 'thin', color: { rgb: '000000' } },
      }

      const labelStyle: CellStyle = {
        fill: { fgColor: { rgb: '00B0F0' } }, // Classic Blue matching N22
        font: { bold: true, color: { rgb: 'FFFFFF' }, sz: 10 },
        alignment: { horizontal: 'center', vertical: 'center' },
        border: widgetBorder,
      }

      const valueStyle: CellStyle = {
        fill: { fgColor: { rgb: '92D050' } }, // Classic Green matching D4 / totals
        font: { bold: true, color: { rgb: '000000' }, sz: 11 },
        alignment: { horizontal: 'right', vertical: 'center' },
        border: widgetBorder,
      }

      const descStyle: CellStyle = {
        fill: { fgColor: { rgb: 'F2F2F2' } },
        font: { color: { rgb: '334155' }, sz: 9 },
        alignment: { horizontal: 'left', vertical: 'center' },
        border: {
          top: { style: 'thin', color: { rgb: 'CBD5E1' } },
          bottom: { style: 'thin', color: { rgb: 'CBD5E1' } },
          left: { style: 'thin', color: { rgb: 'CBD5E1' } },
          right: { style: 'thin', color: { rgb: 'CBD5E1' } },
        },
      }

      // E26:G26 (C2C PROFIT)
      if (ws['E26']) ws['E26'].s = labelStyle
      if (ws['F26']) ws['F26'].s = valueStyle
      if (ws['G26']) ws['G26'].s = descStyle

      // E27:G27 (C2C TOTAL)
      if (ws['E27']) ws['E27'].s = labelStyle
      if (ws['F27']) ws['F27'].s = valueStyle
      if (ws['G27']) ws['G27'].s = descStyle

      // E28:G28 (C2C RATIO)
      if (ws['E28']) ws['E28'].s = labelStyle
      if (ws['F28']) ws['F28'].s = valueStyle
      if (ws['G28']) ws['G28'].s = descStyle
    }

    XLSX.utils.book_append_sheet(book, ws, sheetName.slice(0, 31))
  }

  const cleanFileName = filename.endsWith('.xlsx') ? filename : `${filename}.xlsx`
  XLSX.writeFile(book, cleanFileName)
}
