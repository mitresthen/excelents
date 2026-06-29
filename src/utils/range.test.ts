import { expect, test } from 'vitest'
import { decodeRange, encodeRange } from './range'

test('decodeRange parses a two-cell range', () => {
  expect(decodeRange('A1:B2')).toEqual({ top: 1, left: 1, bottom: 2, right: 2 })
  expect(decodeRange('B2:D10')).toEqual({ top: 2, left: 2, bottom: 10, right: 4 })
})

test('decodeRange treats a single cell as a 1x1 box', () => {
  expect(decodeRange('C3')).toEqual({ top: 3, left: 3, bottom: 3, right: 3 })
})

test('decodeRange normalizes reversed corners', () => {
  expect(decodeRange('B2:A1')).toEqual({ top: 1, left: 1, bottom: 2, right: 2 })
})

test('encodeRange is the inverse of decodeRange', () => {
  expect(encodeRange({ top: 1, left: 1, bottom: 2, right: 2 })).toBe('A1:B2')
  expect(encodeRange({ top: 3, left: 3, bottom: 3, right: 3 })).toBe('C3')
  for (const s of ['A1:B2', 'B2:D10', 'C3']) {
    expect(encodeRange(decodeRange(s))).toBe(s)
  }
})
