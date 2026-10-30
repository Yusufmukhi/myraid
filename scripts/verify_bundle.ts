import fs from 'node:fs'
import path from 'node:path'

const distAssets = fs.readdirSync('dist/assets')
const jsFiles = distAssets.filter((f) => f.endsWith('.js'))

const searchPatterns = [
  'SHEARING SIZE',
  '=C3-160',
  '127388',
  '645.16',
  'MTI3Mzg4',
  'HARDWARE LIST',
  '/api/worksheets',
  '/api/calculate',
  '160+5',
]

let leaks = 0
for (const f of jsFiles) {
  const content = fs.readFileSync(path.join('dist/assets', f), 'utf8')
  for (const p of searchPatterns) {
    if (content.includes(p)) {
      console.log(`SECURITY LEAK: Found "${p}" in ${f}`)
      leaks++
    }
  }
}

if (leaks === 0) {
  console.log('✓ 100% SECURE: Zero formulas, constants, fabrication sheets, or local server endpoints found in browser bundle!')
} else {
  console.error(`✗ FAILED: Found ${leaks} leaks in browser bundle!`)
  process.exit(1)
}
