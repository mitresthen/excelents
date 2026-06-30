import { expect, test } from 'vitest'
import { createWorkbook } from './index'

test('the package exposes createWorkbook building a usable workbook', () => {
  const wb = createWorkbook()
  const ws = wb.addSheet('Sheet1')
  ws.cell('A1').value = 'Hello'
  expect(ws.cell('A1').value).toBe('Hello')
  expect(wb.sheets.map((s) => s.name)).toEqual(['Sheet1'])
})
