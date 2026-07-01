import { expect, test } from 'vitest'
import { createWorkbook } from './workbook'

test('addImage registers media with sequential ids; base64 decodes, bytes pass through', () => {
  const wb = createWorkbook()
  const id0 = wb.addImage({ data: 'aGk=', extension: 'png' }) // "hi"
  const id1 = wb.addImage({ data: new Uint8Array([1, 2, 3]), extension: 'jpeg' })
  expect([id0, id1]).toEqual([0, 1])
  expect(wb.media.length).toBe(2)
  expect(Array.from(wb.media[0]!.data)).toEqual([0x68, 0x69])
  expect(wb.media[0]!.extension).toBe('png')
  expect(Array.from(wb.media[1]!.data)).toEqual([1, 2, 3])
})

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
