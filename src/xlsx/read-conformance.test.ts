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
}, 30000) // sweeps every fixture incl. the ~15 MB huge.xlsx — exceeds the 5s default on CI runners

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

const byName = async (suffix: string) =>
  readXlsx(await parseFixture(listFixtures().find((p) => p.endsWith(suffix))!))

test('the SP-5 readers fire on real fixtures (not just synthetic round-trips)', async () => {
  // Defined names: test-issue-877.xlsx carries a set of workbook-scoped names.
  const named = await byName('test-issue-877.xlsx')
  expect(named.definedNames.length).toBe(15)
  expect(named.definedNames.every((n) => n.name !== '' && n.formula !== '')).toBe(true)

  // Data validation: test-pr-1204.xlsx carries one validation.
  const validated = await byName('test-pr-1204.xlsx')
  const dvCount = validated.sheets.reduce((n, s) => n + s.dataValidations.length, 0)
  expect(dvCount).toBeGreaterThanOrEqual(1)
  const dv = validated.sheets.flatMap((s) => s.dataValidations)[0]!
  expect(dv.sqref).not.toBe('')
  expect(dv.formula1.length).toBeGreaterThan(0)

  // Tables: test-issue-1669.xlsx carries two tables, each with a name and columns.
  const tabled = await byName('test-issue-1669.xlsx')
  const tables = tabled.sheets.flatMap((s) => s.tables)
  expect(tables.length).toBe(2)
  expect(tables.every((t) => t.name !== '' && t.ref !== '' && t.columns.length > 0)).toBe(true)
})

test('the 1904 date-system fixture reads dates on the 1904 epoch', async () => {
  // 1904.xlsx sets <workbookPr date1904="1"/>; B4 is serial 0 = 1904-01-01
  // (confirmed against the exceljs oracle). On the 1900 epoch it would read
  // as 1900-01-01 — four years off.
  const wb = await byName('1904.xlsx')
  expect(wb.sheets[0]!.cell('B4').value).toEqual(new Date(Date.UTC(1904, 0, 1)))
})

test('sheet-scoped defined names in a real fixture keep their localSheetId', async () => {
  // test-issue-877.xlsx has 4 names with a localSheetId (3 sharing a label with a global
  // twin). They must stay scoped, not collapse to colliding global names.
  const wb = await byName('test-issue-877.xlsx')
  const scoped = wb.definedNames.filter((n) => n.localSheetId !== undefined)
  expect(scoped.length).toBe(4)
})

test('every fixture still parses and exposes definedNames as an array', async () => {
  for (const path of listFixtures()) {
    if (KNOWN_INCOMPLETE.has(path.split('/').pop()!)) continue
    const wb = await readXlsx(await parseFixture(path))
    expect(Array.isArray(wb.definedNames)).toBe(true)
  }
}, 30000) // sweeps every fixture incl. the ~15 MB huge.xlsx — exceeds the 5s default on CI runners
