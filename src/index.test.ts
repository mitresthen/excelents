import { readFileSync } from 'node:fs'
import { expect, test } from 'vitest'
import {
  createWorkbook,
  type DataValidation,
  readXlsx,
  type TableDefinition,
  version,
  writeXlsx,
} from './index'

test('the package exports a version string kept in sync with package.json', () => {
  expect(typeof version).toBe('string')
  const raw = readFileSync(new URL('../package.json', import.meta.url), 'utf-8')
  const match = /"version":\s*"([^"]+)"/.exec(raw)
  expect(match?.[1]).toBe(version)
})

test('the package exposes createWorkbook building a usable workbook', () => {
  const wb = createWorkbook()
  const ws = wb.addSheet('Sheet1')
  ws.cell('A1').value = 'Hello'
  expect(ws.cell('A1').value).toBe('Hello')
  expect(wb.sheets.map((s) => s.name)).toEqual(['Sheet1'])
})

test('the package exposes writeXlsx producing ZIP-magic bytes', async () => {
  const wb = createWorkbook()
  wb.addSheet('Sheet1').cell('A1').value = 'Hello'
  const bytes = await writeXlsx(wb)
  // .xlsx is a ZIP: the local-file-header signature is 'PK\x03\x04'.
  expect(bytes[0]).toBe(0x50)
  expect(bytes[1]).toBe(0x4b)
  expect(bytes[2]).toBe(0x03)
  expect(bytes[3]).toBe(0x04)
})

test('the package exposes readXlsx round-tripping writeXlsx', async () => {
  const wb = createWorkbook()
  wb.addSheet('S').cell('A1').value = 'hello'
  const restored = await readXlsx(await writeXlsx(wb))
  expect(restored.sheets[0]?.cell('A1').value).toBe('hello')
})

test('the package exposes the SP-5 feature surface (names, validation, tables)', async () => {
  const wb = createWorkbook()
  const ws = wb.addSheet('S')
  wb.defineName('R', 'S!$A$1')
  const dv: DataValidation = { sqref: 'A1', type: 'list', formula1: '"a,b"' }
  ws.addDataValidation(dv)
  const table: TableDefinition = { name: 'T', ref: 'A1:B2', columns: ['a', 'b'] }
  ws.addTable(table)

  const r = await readXlsx(await writeXlsx(wb))
  expect(r.definedNames).toEqual([{ name: 'R', formula: 'S!$A$1' }])
  expect(r.sheets[0]?.dataValidations[0]).toEqual(dv)
  expect(r.sheets[0]?.tables[0]).toEqual(table)
})
