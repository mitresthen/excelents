import { expect, test } from 'vitest'
import { readableToBytes } from '../io/streams'
import { readXlsx } from './read'
import { createXlsxStreamWriter, writeXlsxStream } from './stream-writer'

const ROWS = [
  ['name', 'qty', 'flag', 'when'],
  ['wid,get', 42, true, new Date(Date.UTC(2020, 0, 15))],
  ['  pad  ', -3.5, false, null],
]

test('writeXlsxStream output round-trips through readXlsx', async () => {
  const bytes = await readableToBytes(writeXlsxStream(ROWS, { sheet: 'Data' }))
  const ws = (await readXlsx(bytes)).sheets[0]!
  expect(ws.name).toBe('Data')
  expect(ws.cell('A1').value).toBe('name')
  expect(ws.cell('B2').value).toBe(42)
  expect(ws.cell('C2').value).toBe(true)
  expect(ws.cell('A2').value).toBe('wid,get')
  expect(ws.cell('A3').value).toBe('  pad  ') // surrounding whitespace preserved
  expect(ws.cell('B3').value).toBe(-3.5)
  const when = ws.cell('D2').value
  expect(when).toBeInstanceOf(Date)
  if (when instanceof Date) expect(when.getUTCFullYear()).toBe(2020)
})

test('exceljs reads the streamed workbook', async () => {
  const bytes = await readableToBytes(writeXlsxStream(ROWS, { sheet: 'Data' }))
  const ExcelJS = (await import('exceljs')).default
  const oracle = new ExcelJS.Workbook()
  // @ts-expect-error -- @types/node Buffer<ArrayBuffer> vs exceljs decl
  await oracle.xlsx.load(Buffer.from(bytes))
  const ws = oracle.getWorksheet('Data')!
  expect(ws.getCell('A1').value).toBe('name')
  expect(ws.getCell('B2').value).toBe(42)
})

test('the builder API (addRow/close) produces the same result', async () => {
  const w = createXlsxStreamWriter({ sheet: 'Data' })
  const collect = readableToBytes(w.readable)
  for (const row of ROWS) await w.addRow(row)
  await w.close()
  const ws = (await readXlsx(await collect)).sheets[0]!
  expect(ws.cell('A1').value).toBe('name')
  expect(ws.cell('B2').value).toBe(42)
  expect(ws.cell('A2').value).toBe('wid,get')
})

test('an empty stream still produces a valid (empty) workbook', async () => {
  const bytes = await readableToBytes(writeXlsxStream([], { sheet: 'S' }))
  const wb = await readXlsx(bytes)
  expect(wb.sheets[0]!.name).toBe('S')
})
