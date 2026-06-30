import { Readable } from 'node:stream'
import { expect, test } from 'vitest'
import { readCsv, writeCsv } from '../csv'
import { createWorkbook } from '../model/workbook'

test('writeCsv -> readCsv round-trips values with conservative inference', () => {
  const wb = createWorkbook()
  const ws = wb.addSheet('S')
  ws.cell('A1').value = 'text'
  ws.cell('B1').value = 42
  ws.cell('C1').value = true
  ws.cell('A2').value = 'has,comma'
  ws.cell('B2').value = '00210' // must come back as a string, not 210
  const r = readCsv(writeCsv(wb)).sheets[0]!
  expect(r.cell('A1').value).toBe('text')
  expect(r.cell('B1').value).toBe(42)
  expect(r.cell('C1').value).toBe(true)
  expect(r.cell('A2').value).toBe('has,comma')
  expect(r.cell('B2').value).toBe('00210')
})

test('exceljs reads our CSV; we read exceljs CSV (cross-conformance)', async () => {
  const wb = createWorkbook()
  const ws = wb.addSheet('S')
  ws.cell('A1').value = 'a,b'
  ws.cell('B1').value = 7
  const ours = writeCsv(wb)

  const ExcelJS = (await import('exceljs')).default
  const oracle = new ExcelJS.Workbook()
  const oraSheet = await oracle.csv.read(Readable.from(ours))
  expect(oraSheet.getCell('A1').value).toBe('a,b')
  expect(oraSheet.getCell('B1').value).toBe(7)

  const theirs = new TextDecoder().decode(await oracle.csv.writeBuffer())
  const back = readCsv(theirs).sheets[0]!
  expect(back.cell('A1').value).toBe('a,b')
  expect(back.cell('B1').value).toBe(7)
})
