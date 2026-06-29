import { expect, test } from 'vitest'
import { SharedStrings } from './shared-strings'

test('add interns strings and returns stable indices', () => {
  const sst = new SharedStrings()
  expect(sst.add('hello')).toBe(0)
  expect(sst.add('world')).toBe(1)
  expect(sst.add('hello')).toBe(0) // duplicate returns existing index
})

test('count tracks total references; uniqueCount tracks distinct strings', () => {
  const sst = new SharedStrings()
  sst.add('a')
  sst.add('b')
  sst.add('a')
  expect(sst.count).toBe(3)
  expect(sst.uniqueCount).toBe(2)
})

test('getIndex and getString round-trip', () => {
  const sst = new SharedStrings()
  sst.add('x')
  sst.add('y')
  expect(sst.getIndex('y')).toBe(1)
  expect(sst.getString(0)).toBe('x')
  expect(sst.getIndex('missing')).toBeUndefined()
  expect(sst.getString(99)).toBeUndefined()
})

test('values lists distinct strings in insertion order', () => {
  const sst = new SharedStrings()
  sst.add('first')
  sst.add('second')
  sst.add('first')
  expect(sst.values).toEqual(['first', 'second'])
})
