import { expect, test } from 'vitest'
import { canonicalizeXml, listFixtures, parseFixture, withOracle } from './harness'

test('canonicalizeXml drops whitespace-only text runs between elements', () => {
  const pretty = '<root>\n  <a/>\n  <b/>\n</root>'
  const compact = '<root><a/><b/></root>'
  // A pretty-printed producer and a compact one must canonicalize identically.
  expect(canonicalizeXml(pretty)).toBe(canonicalizeXml(compact))
})

test('canonicalizeXml preserves text runs that contain non-whitespace content', () => {
  expect(canonicalizeXml('<t>hello</t>')).toContain('hello')
  // Content with internal whitespace is kept verbatim.
  expect(canonicalizeXml('<t>a b</t>')).toContain('a b')
})

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
