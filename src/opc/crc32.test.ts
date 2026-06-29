import { expect, test } from 'vitest'
import { crc32 } from './crc32'

const enc = (s: string): Uint8Array => new TextEncoder().encode(s)

test('crc32 of empty input is 0', () => {
  expect(crc32(new Uint8Array(0))).toBe(0)
})

test('crc32 matches known vectors', () => {
  // Well-known CRC-32 test vectors.
  expect(crc32(enc('123456789'))).toBe(0xcbf43926)
  expect(crc32(enc('The quick brown fox jumps over the lazy dog'))).toBe(0x414fa339)
})

test('crc32 returns an unsigned 32-bit value', () => {
  const c = crc32(enc('hello world'))
  expect(c).toBeGreaterThanOrEqual(0)
  expect(c).toBeLessThanOrEqual(0xffffffff)
  expect(Number.isInteger(c)).toBe(true)
})
