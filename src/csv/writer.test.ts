import { expect, test } from 'vitest'
import { createWorkbook } from '../model/workbook'
import { writeCsv } from './writer'

const BOM = String.fromCharCode(0xfeff)

test('renders a simple grid with bare numbers and booleans', () => {
  const wb = createWorkbook()
  const ws = wb.addSheet('S')
  ws.cell('A1').value = 'name'
  ws.cell('B1').value = 42
  ws.cell('C1').value = true
  ws.cell('A2').value = 'x'
  expect(writeCsv(wb)).toBe('name,42,true\nx,,')
})

test('quotes fields containing the delimiter, quotes, or newlines (RFC 4180)', () => {
  const wb = createWorkbook()
  const ws = wb.addSheet('S')
  ws.cell('A1').value = 'has,comma'
  ws.cell('B1').value = 'has"quote'
  ws.cell('C1').value = 'has\nnl'
  expect(writeCsv(ws)).toBe('"has,comma","has""quote","has\nnl"')
})

test('honors delimiter, quoteAll, rowDelimiter, and bom options', () => {
  const wb = createWorkbook()
  const ws = wb.addSheet('S')
  ws.cell('A1').value = 'a'
  ws.cell('B1').value = 'b'
  ws.cell('A2').value = 'c'
  ws.cell('B2').value = 'd'
  expect(writeCsv(ws, { delimiter: ';', quoteAll: true, rowDelimiter: '\r\n', bom: true })).toBe(
    `${BOM}"a";"b"\r\n"c";"d"`,
  )
})

test('renders formula results, rich text, and hyperlinks as text', () => {
  const wb = createWorkbook()
  const ws = wb.addSheet('S')
  ws.cell('A1').value = { formula: 'SUM(B1:B2)', result: 7 }
  ws.cell('B1').value = { richText: [{ text: 'Hel' }, { text: 'lo' }] }
  ws.cell('C1').value = { text: 'site', hyperlink: 'https://x' }
  expect(writeCsv(ws)).toBe('7,Hello,site')
})
