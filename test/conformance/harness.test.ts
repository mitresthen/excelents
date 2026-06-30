import { expect, test } from 'vitest'
import { listFixtures, parseFixture, withOracle } from './harness'

test('finds the carried-over xlsx fixtures', () => {
  const fixtures = listFixtures()
  expect(fixtures.length).toBeGreaterThanOrEqual(30)
})

test('loads a fixture as bytes', async () => {
  const [first] = listFixtures()
  const bytes = await parseFixture(first!)
  // xlsx is a zip: first two bytes are 'PK'
  expect(bytes[0]).toBe(0x50)
  expect(bytes[1]).toBe(0x4b)
})

test('the exceljs oracle loads as a devDependency', async () => {
  const ExcelJS = (await import('exceljs')).default
  expect(typeof ExcelJS.Workbook).toBe('function')
})

test('withOracle runs build fn for both oracle and excelents', async () => {
  const result = await withOracle(async (lib) => lib)
  expect(result.oracle).toBe('oracle')
  expect(result.excelents).toBe('excelents')
})
