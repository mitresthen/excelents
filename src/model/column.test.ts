import { expect, test } from 'vitest'
import { Column } from './column'

test('exposes its letter, key, and width', () => {
  const col = new Column(27)
  expect(col.letter).toBe('AA')
  col.key = 'id'
  col.width = 12
  expect(col.key).toBe('id')
  expect(col.width).toBe(12)
})
