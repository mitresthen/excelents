import { expect, test } from 'vitest'
import { colToNumber, decodeAddress, encodeAddress, numberToCol } from './address'

test('colToNumber maps letters to 1-based numbers', () => {
  expect(colToNumber('A')).toBe(1)
  expect(colToNumber('Z')).toBe(26)
  expect(colToNumber('AA')).toBe(27)
  expect(colToNumber('XFD')).toBe(16384) // Excel max column
})

test('numberToCol is the inverse of colToNumber', () => {
  expect(numberToCol(1)).toBe('A')
  expect(numberToCol(26)).toBe('Z')
  expect(numberToCol(27)).toBe('AA')
  expect(numberToCol(16384)).toBe('XFD')
  for (const n of [1, 2, 26, 27, 52, 53, 702, 703, 16384]) {
    expect(colToNumber(numberToCol(n))).toBe(n)
  }
})

test('decodeAddress parses cell addresses, tolerating absolute markers', () => {
  expect(decodeAddress('A1')).toEqual({ row: 1, col: 1 })
  expect(decodeAddress('B2')).toEqual({ row: 2, col: 2 })
  expect(decodeAddress('$A$1')).toEqual({ row: 1, col: 1 })
  expect(decodeAddress('AA100')).toEqual({ row: 100, col: 27 })
})

test('decodeAddress throws on malformed input', () => {
  expect(() => decodeAddress('1A')).toThrow()
  expect(() => decodeAddress('A')).toThrow()
  expect(() => decodeAddress('')).toThrow()
})

test('encodeAddress is the inverse of decodeAddress', () => {
  expect(encodeAddress(1, 1)).toBe('A1')
  expect(encodeAddress(100, 27)).toBe('AA100')
  for (const [r, c] of [
    [1, 1],
    [10, 26],
    [5, 27],
    [100, 16384],
  ] as const) {
    const a = encodeAddress(r, c)
    expect(decodeAddress(a)).toEqual({ row: r, col: c })
  }
})
