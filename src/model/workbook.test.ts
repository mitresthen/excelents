import { expect, test } from 'vitest'
import { createWorkbook } from './workbook'

test('createWorkbook starts with no sheets', () => {
  expect(createWorkbook().sheets).toEqual([])
})

test('addSheet adds a named sheet and getSheet finds it', () => {
  const wb = createWorkbook()
  const ws = wb.addSheet('Data')
  expect(ws.name).toBe('Data')
  expect(wb.getSheet('Data')).toBe(ws)
  expect(wb.sheets.map((s) => s.name)).toEqual(['Data'])
})

test('removeSheet removes by name', () => {
  const wb = createWorkbook()
  wb.addSheet('A')
  wb.addSheet('B')
  wb.removeSheet('A')
  expect(wb.sheets.map((s) => s.name)).toEqual(['B'])
})

test('end-to-end: build a workbook and read values back', () => {
  const wb = createWorkbook()
  const ws = wb.addSheet('Sheet1')
  ws.cell('A1').value = 'Hello'
  ws.addRow(['world', 42])
  expect(ws.cell('A1').value).toBe('Hello')
  expect(ws.cell('A2').value).toBe('world')
  expect(ws.cell('B2').value).toBe(42)
})
