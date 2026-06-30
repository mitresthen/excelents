import { expect, test } from 'vitest'
import { createWorkbook, readXlsx, writeXlsx } from './index'

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
