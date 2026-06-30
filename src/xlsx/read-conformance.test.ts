import { expect, test } from 'vitest'
import { listFixtures, parseFixture } from '../../test/conformance/harness'
import { readXlsx } from './read'

// missing-bits.xlsx is an intentionally-incomplete ZIP (no [Content_Types].xml, no
// workbook part) used to test missing-part handling; our reader reads it gracefully
// (no throw) but recovers no sheets. Excluded from the "has sheets" assertion only.
const KNOWN_INCOMPLETE = new Set(['missing-bits.xlsx'])

test('readXlsx parses every fixture without throwing', async () => {
  const fixtures = listFixtures()
  expect(fixtures.length).toBeGreaterThanOrEqual(30)
  for (const path of fixtures) {
    const name = path.split('/').pop()!
    const bytes = await parseFixture(path)
    const wb = await readXlsx(bytes) // must not throw, even on the incomplete fixture
    if (!KNOWN_INCOMPLETE.has(name)) {
      expect(wb.sheets.length).toBeGreaterThanOrEqual(1)
    }
  }
})

test('recovered primitive cell values match exceljs for a representative fixture', async () => {
  const ExcelJS = (await import('exceljs')).default
  const path = listFixtures().sort()[0]!
  const bytes = await parseFixture(path)

  const ours = await readXlsx(bytes)
  const oracle = new ExcelJS.Workbook()
  // @ts-expect-error -- @types/node Buffer<ArrayBuffer> vs exceljs decl
  await oracle.xlsx.load(Buffer.from(bytes))

  const ourSheet = ours.sheets[0]!
  const oracleSheet = oracle.worksheets[0]!
  let compared = 0
  oracleSheet.eachRow((row, r) => {
    row.eachCell((cell, c) => {
      const v = cell.value
      if (typeof v === 'string' || typeof v === 'number' || typeof v === 'boolean') {
        expect(ourSheet.getCell(r, c).value).toBe(v)
        compared++
      }
    })
  })
  expect(compared).toBeGreaterThan(0) // non-vacuous
})

test('a formula cell that also carries a hyperlink keeps its formula', async () => {
  // formulas.xlsx A1 has a formula AND a <hyperlink> annotation. The cell value must
  // remain the formula (our model cannot hold formula+hyperlink at once) — it must NOT
  // be clobbered into a blank { text: '', hyperlink } wrapper.
  const path = listFixtures().find((p) => p.endsWith('formulas.xlsx'))!
  const wb = await readXlsx(await parseFixture(path))
  const a1 = wb.sheets[0]!.cell('A1').value
  if (a1 === null || typeof a1 !== 'object' || !('formula' in a1)) {
    throw new Error(`expected A1 to stay a formula, got ${JSON.stringify(a1)}`)
  }
  expect(a1.formula).toContain('CONCAT')
})
