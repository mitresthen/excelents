import { describe, expect, it } from 'vitest'
import { referencesExcelJs, scanSource } from './scan'

const EXCELJS_FILE = `import ExcelJS from 'exceljs'

const wb = new ExcelJS.Workbook()
const ws = wb.addWorksheet('Data', { views: [] })
ws.getCell('A1').value = 7
ws.getCell(2, 1).value = 8
ws.getCell('A1').font = { bold: true, color: { argb: 'FFFF0000' } }
ws.getCell('A1').numFmt = '0.00%'
ws.mergeCells('A1:B2')
ws.autoFilter = 'A1:C1'
ws.views = [{ state: 'frozen', ySplit: 1 }]
ws.getCell('B2').dataValidation = { type: 'list', formulae: ['"a,b"'] }
ws.getCell('C1').note = 'a comment'
ws.pageSetup = { orientation: 'landscape' }
const buf = await wb.xlsx.writeBuffer()
await wb.xlsx.load(buf)
`

describe('scanSource', () => {
  const findings = scanSource(EXCELJS_FILE, 'src/report.ts')

  it('labels findings with file and 1-based line numbers', () => {
    const imp = findings.find((f) => f.ruleId === 'import')
    expect(imp).toMatchObject({ file: 'src/report.ts', line: 1, category: 'map' })
  })

  it('classifies mechanical renames as map', () => {
    const ids = findings.filter((f) => f.category === 'map').map((f) => f.ruleId)
    expect(ids).toContain('new-workbook')
    expect(ids).toContain('add-worksheet')
    expect(ids).toContain('merge-cells')
    expect(ids).toContain('font-assign')
    expect(ids).toContain('num-fmt')
    expect(ids).toContain('auto-filter')
    expect(ids).toContain('xlsx-write-buffer')
    expect(ids).toContain('views-frozen')
  })

  it('only flags getCell with a string address', () => {
    const cellFindings = findings.filter((f) => f.ruleId === 'get-cell-address')
    // lines 5, 7, 8, 12, 13 use string addresses; line 6 (numeric) must not match
    expect(cellFindings.map((f) => f.line)).not.toContain(6)
    expect(cellFindings.length).toBeGreaterThan(0)
  })

  it('classifies load and per-cell validation as restructure', () => {
    const ids = findings.filter((f) => f.category === 'restructure').map((f) => f.ruleId)
    expect(ids).toContain('xlsx-load')
    expect(ids).toContain('data-validation')
  })

  it('classifies notes and pageSetup as blocked', () => {
    const ids = findings.filter((f) => f.category === 'blocked').map((f) => f.ruleId)
    expect(ids).toContain('cell-notes')
    expect(ids).toContain('page-setup')
  })

  it('suppresses the generic views advice when the line is a frozen view', () => {
    const viewLines = findings.filter((f) => f.ruleId === 'views-assign')
    expect(viewLines).toEqual([])
    expect(findings.some((f) => f.ruleId === 'views-frozen')).toBe(true)
  })

  it('points every finding at a guide section', () => {
    for (const f of findings) expect(f.guideSection).toMatch(/^[a-z-]+$/)
  })
})

describe('context gating', () => {
  it('skips generic member patterns in files that never mention exceljs', () => {
    const source = `el.style.font = 'bold'\nrow.hidden = true\nconfig.views = []\n`
    expect(scanSource(source, 'ui.ts')).toEqual([])
  })

  it('still flags distinctive exceljs APIs in helper files without an import', () => {
    const source = `export function fill(ws) {\n  ws.addWorksheet('x')\n  ws.mergeCells('A1:B1')\n}\n`
    const ids = scanSource(source, 'helper.ts').map((f) => f.ruleId)
    expect(ids).toEqual(['add-worksheet', 'merge-cells'])
  })

  it('returns nothing for clean files', () => {
    expect(scanSource(`export const x = 1\n`, 'clean.ts')).toEqual([])
  })
})

describe('referencesExcelJs', () => {
  it('detects imports, requires, and the global name', () => {
    expect(referencesExcelJs(`import x from 'exceljs'`)).toBe(true)
    expect(referencesExcelJs(`const x = require('exceljs')`)).toBe(true)
    expect(referencesExcelJs(`new ExcelJS.Workbook()`)).toBe(true)
    expect(referencesExcelJs(`const sheet = 'excel'`)).toBe(false)
  })
})
