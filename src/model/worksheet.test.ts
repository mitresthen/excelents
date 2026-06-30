import { expect, test } from 'vitest'
import { Worksheet } from './worksheet'

test('cell(ref) addresses the same cell as getCell(row, col)', () => {
  const ws = new Worksheet('Sheet1')
  ws.cell('B2').value = 'hi'
  expect(ws.getCell(2, 2).value).toBe('hi')
  expect(ws.cell('B2')).toBe(ws.getCell(2, 2))
})

test('addRow appends after the last populated row', () => {
  const ws = new Worksheet('s')
  ws.cell('A1').value = 'header'
  const row = ws.addRow(['a', 1, true])
  expect(row.number).toBe(2)
  expect(ws.cell('A2').value).toBe('a')
  expect(ws.cell('B2').value).toBe(1)
  expect(ws.cell('C2').value).toBe(true)
})

test('merge records ranges', () => {
  const ws = new Worksheet('s')
  ws.merge('A1:B2')
  expect(ws.merges).toEqual(['A1:B2'])
})

test('dimensions spans the populated cells', () => {
  const ws = new Worksheet('s')
  ws.cell('B2').value = 'x'
  ws.cell('D5').value = 'y'
  expect(ws.dimensions).toEqual({ top: 2, left: 2, bottom: 5, right: 4 })
})

test('dimensions is undefined when empty', () => {
  expect(new Worksheet('s').dimensions).toBeUndefined()
})

test('rows lists only populated rows ascending', () => {
  const ws = new Worksheet('s')
  ws.cell('A3').value = 'c'
  ws.cell('A1').value = 'a'
  ws.getRow(2) // touched but empty — excluded
  expect(ws.rows.map((r) => r.number)).toEqual([1, 3])
})
