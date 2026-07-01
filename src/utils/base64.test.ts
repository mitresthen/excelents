import { expect, test } from 'vitest'
import { base64ToBytes } from './base64'

test('base64ToBytes decodes ASCII text', () => {
  // "hi" -> "aGk="
  expect(Array.from(base64ToBytes('aGk='))).toEqual([0x68, 0x69])
})

test('base64ToBytes preserves binary (non-ASCII) bytes', () => {
  // bytes [0x00, 0xff, 0x10] -> "AP8Q"
  expect(Array.from(base64ToBytes('AP8Q'))).toEqual([0x00, 0xff, 0x10])
})

test('base64ToBytes decodes an empty string to empty bytes', () => {
  expect(base64ToBytes('').length).toBe(0)
})
