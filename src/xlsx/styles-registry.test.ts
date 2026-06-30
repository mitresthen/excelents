import { expect, test } from 'vitest'
import { StyleRegistry } from './styles-writer'

test('empty style maps to xf 0 without creating a duplicate default', () => {
  const r = new StyleRegistry()
  expect(r.xfIndexFor({})).toBe(0)
  expect(r.xfIndexFor({})).toBe(0)
  // An empty alignment object is meaningless and must not split the default xf.
  expect(r.xfIndexFor({ alignment: {} })).toBe(0)
  expect(r.xfs.length).toBe(1)
})

test('identical styles dedup; distinct facets allocate new xfs', () => {
  const r = new StyleRegistry()
  const bold = r.xfIndexFor({ font: { bold: true } })
  expect(r.xfIndexFor({ font: { bold: true } })).toBe(bold)
  const boldFmt = r.xfIndexFor({ font: { bold: true }, numberFormat: '0.00' })
  expect(boldFmt).not.toBe(bold)
})

test('alignment is part of the xf identity', () => {
  const r = new StyleRegistry()
  const plain = r.xfIndexFor({ font: { bold: true } })
  const aligned = r.xfIndexFor({ font: { bold: true }, alignment: { horizontal: 'center' } })
  expect(aligned).not.toBe(plain)
  // Same alignment + same font dedups.
  expect(r.xfIndexFor({ font: { bold: true }, alignment: { horizontal: 'center' } })).toBe(aligned)
})

test('alignment dedup key is insensitive to facet insertion order', () => {
  const r = new StyleRegistry()
  const a = r.xfIndexFor({ alignment: { horizontal: 'center', vertical: 'middle' } })
  const b = r.xfIndexFor({ alignment: { vertical: 'middle', horizontal: 'center' } })
  expect(b).toBe(a)
})

test('mandatory fills none@0 and gray125@1 are always present', () => {
  const r = new StyleRegistry()
  expect(r.fills.length).toBeGreaterThanOrEqual(2)
  expect(r.fills[0]).toMatchObject({ pattern: 'none' })
  expect(r.fills[1]).toMatchObject({ pattern: 'gray125' })
})
