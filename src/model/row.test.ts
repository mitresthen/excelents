import { expect, test } from 'vitest'
import { Row } from './row'

test('getCell lazily creates and caches a cell at (row, col)', () => {
  const row = new Row(5)
  const a = row.getCell(1)
  expect(a.address).toBe('A5')
  expect(row.getCell(1)).toBe(a) // cached (same instance)
})

test('cells lists populated cells in ascending column order', () => {
  const row = new Row(1)
  row.getCell(3).value = 'c'
  row.getCell(1).value = 'a'
  expect(row.cells.map((c) => c.address)).toEqual(['A1', 'C1'])
})

test('carries an optional height', () => {
  const row = new Row(2)
  row.height = 18
  expect(row.height).toBe(18)
})
