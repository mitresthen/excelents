import { expect, test } from 'vitest'
import { createWorkbook } from '../model/workbook'
import { readXlsx } from './read'
import { writeXlsx } from './write'

test('round-trips string, number, and boolean cell values', async () => {
  const wb = createWorkbook()
  const ws = wb.addSheet('S')
  ws.cell('A1').value = 'apple'
  ws.cell('A2').value = 'apple' // dedup -> same shared string
  ws.cell('B1').value = 42
  ws.cell('B2').value = 1234.5
  ws.cell('C1').value = true
  ws.cell('C2').value = false
  const r = await readXlsx(await writeXlsx(wb))
  const s = r.sheets[0]!
  expect(s.cell('A1').value).toBe('apple')
  expect(s.cell('A2').value).toBe('apple')
  expect(s.cell('B1').value).toBe(42)
  expect(s.cell('B2').value).toBe(1234.5)
  expect(s.cell('C1').value).toBe(true)
  expect(s.cell('C2').value).toBe(false)
})

test('round-trips applied cell styles and Date values', async () => {
  const wb = createWorkbook()
  const ws = wb.addSheet('S')
  const bold = ws.cell('A1')
  bold.value = 'x'
  bold.style = { font: { bold: true }, numberFormat: '0.00' }
  const when = new Date(Date.UTC(2026, 5, 30, 12, 0, 0))
  ws.cell('B1').value = when
  const r = await readXlsx(await writeXlsx(wb))
  const s = r.sheets[0]!
  expect(s.cell('A1').style.font?.bold).toBe(true)
  expect(s.cell('A1').style.numberFormat).toBe('0.00')
  const bVal = s.cell('B1').value
  expect(bVal).toBeInstanceOf(Date)
  if (!(bVal instanceof Date)) throw new Error('expected a Date')
  expect(bVal.getTime()).toBe(when.getTime())
})

test('round-trips formula, richText, merge, column width, and row height', async () => {
  const wb = createWorkbook()
  const ws = wb.addSheet('S')
  ws.cell('A1').value = { formula: '1+2', result: 3 }
  ws.cell('A2').value = { richText: [{ text: 'Hi ' }, { text: 'bold', font: { bold: true } }] }
  ws.cell('A3').value = 'm'
  ws.merge('A3:B4')
  ws.column(1).width = 20
  ws.getRow(1).height = 30
  const r = await readXlsx(await writeXlsx(wb))
  const s = r.sheets[0]!
  expect(s.cell('A1').value).toEqual({ formula: '1+2', result: 3 })
  const rich = s.cell('A2').value
  if (rich === null || typeof rich !== 'object' || !('richText' in rich)) {
    throw new Error('expected richText value')
  }
  expect(rich.richText.map((run) => run.text).join('')).toBe('Hi bold')
  expect(s.merges).toContain('A3:B4')
  expect(s.column(1).width).toBe(20)
  expect(s.getRow(1).height).toBe(30)
})

test('round-trips a hyperlink as text + target', async () => {
  const wb = createWorkbook()
  const ws = wb.addSheet('S')
  ws.cell('A1').value = { text: 'site', hyperlink: 'https://example.com/' }
  const r = await readXlsx(await writeXlsx(wb))
  expect(r.sheets[0]!.cell('A1').value).toEqual({ text: 'site', hyperlink: 'https://example.com/' })
})
