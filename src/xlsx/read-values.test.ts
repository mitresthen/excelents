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
