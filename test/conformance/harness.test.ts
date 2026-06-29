import { expect, test } from 'vitest'
import { listFixtures, NotImplemented, parseFixture, unzipToParts } from './harness'

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

test('zip/xml-dependent harness fns throw NotImplemented until SP-1', () => {
  expect(() => unzipToParts(new Uint8Array([0]))).toThrow(NotImplemented)
})
