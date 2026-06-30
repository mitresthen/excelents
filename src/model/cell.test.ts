import { expect, test } from 'vitest'
import { Cell } from './cell'

test('a new cell is null at its address', () => {
  const c = new Cell(1, 1)
  expect(c.address).toBe('A1')
  expect(c.value).toBeNull()
  expect(c.type).toBe('null')
})

test('detects scalar types from the value', () => {
  const c = new Cell(2, 3) // C2
  expect(c.address).toBe('C2')
  c.value = 'hello'
  expect(c.type).toBe('string')
  c.value = 42
  expect(c.type).toBe('number')
  c.value = true
  expect(c.type).toBe('boolean')
  c.value = new Date(Date.UTC(2024, 0, 1))
  expect(c.type).toBe('date')
})

test('detects formula, richText, and hyperlink values', () => {
  const c = new Cell(1, 1)
  c.value = { formula: 'A1+B1', result: 7 }
  expect(c.type).toBe('formula')
  c.value = { richText: [{ text: 'a' }, { text: 'b', font: { bold: true } }] }
  expect(c.type).toBe('richText')
  c.value = { text: 'site', hyperlink: 'https://example.com' }
  expect(c.type).toBe('hyperlink')
})

test('carries a mutable style', () => {
  const c = new Cell(1, 1)
  c.style = { font: { bold: true }, numberFormat: '0.00' }
  expect(c.style.numberFormat).toBe('0.00')
})
