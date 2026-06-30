import { expect, test } from 'vitest'
import { type CellStyle, mergeStyles } from './style'

test('mergeStyles overlays patch facets over base', () => {
  const base: CellStyle = { font: { bold: true }, numberFormat: '0.00' }
  const patch: CellStyle = {
    font: { italic: true },
    fill: { type: 'pattern', pattern: 'solid', fgColor: 'FFFF0000' },
  }
  const merged = mergeStyles(base, patch)
  // patch's font replaces base's font (facet-level merge, not deep)
  expect(merged.font).toEqual({ italic: true })
  expect(merged.fill).toEqual({ type: 'pattern', pattern: 'solid', fgColor: 'FFFF0000' })
  expect(merged.numberFormat).toBe('0.00')
})

test('mergeStyles does not mutate its inputs', () => {
  const base: CellStyle = { font: { bold: true } }
  mergeStyles(base, { numberFormat: '0' })
  expect(base).toEqual({ font: { bold: true } })
})

test('mergeStyles of empties is empty', () => {
  expect(mergeStyles({}, {})).toEqual({})
})
