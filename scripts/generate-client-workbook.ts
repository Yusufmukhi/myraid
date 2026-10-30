import fs from 'fs'
import { Workbook } from '../src/lib/engine'
import { summarize } from '../src/lib/model'

const SERVER_CONSTANTS = {
  CONST_STEEL_FACTOR: 127388,
  CONST_SURFACE_FACTOR: 645.16,
}

function processWorkbook(inputPath: string, outputPath: string) {
  const rawWb = JSON.parse(fs.readFileSync(inputPath, 'utf8'))
  const fullWb = new Workbook(rawWb, SERVER_CONSTANTS)
  const expectedSum = summarize(fullWb, 20)

  const tc = JSON.parse(JSON.stringify(rawWb.sheets.find((s: any) => s.name === 'Total cal')))

  for (let r = 0; r < tc.cells.length; r++) {
    const row = tc.cells[r]
    if (!row) continue
    for (let c = 0; c < row.length; c++) {
      const val = row[c]
      if (typeof val === 'string') {
        // Row 7 (SFMU) and Row 11 (SFFU) master links to Row 3 (DFMU)
        if (r === 6 || r === 10) {
          if (val.includes('DFMU!A3')) {
            row[c] = '=(A3/2)'
            continue
          }
          if (val.includes('DFMU!')) {
            row[c] = '=' + val.replace(/[=()]/g, '').replace('DFMU!', '')
            continue
          }
        }

        // External sheet references (cutting/fabrication sheets) -> evaluate to base value
        if (val.includes('!')) {
          const evalVal = fullWb.value('Total cal', r, c)
          row[c] = typeof evalVal === 'number' ? Math.round(evalVal * 1000000) / 1000000 : evalVal
        }
      }
    }
  }

  const clientWb = {
    sheets: [tc]
  }

  fs.writeFileSync(outputPath, JSON.stringify(clientWb, null, 2))

  const testWb = new Workbook(clientWb)
  const gotSum = summarize(testWb, 20)

  console.log(`Generated ${outputPath}:`)
  console.log('  Expected Sales:', expectedSum.sales, 'Got Sales:', gotSum.sales)
  console.log('  Diff:', Math.abs(expectedSum.sales - gotSum.sales))
  if (Math.abs(expectedSum.sales - gotSum.sales) > 1e-4) {
    throw new Error(`Mismatch in ${outputPath}! Expected ${expectedSum.sales}, got ${gotSum.sales}`)
  }
}

processWorkbook('data/workbook.json', 'src/lib/clientWorkbook.json')
processWorkbook('data/workbook_c2c.json', 'src/lib/clientWorkbook_c2c.json')
console.log('✓ Successfully generated both client workbooks!')
