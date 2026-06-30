import { expect, test } from 'vitest'
import { readStyles } from './styles-reader'
import { StyleRegistry, writeStylesXml } from './styles-writer'

test('reads back cellXfs into CellStyles round-tripped from the writer', () => {
  const reg = new StyleRegistry()
  const xfBold = reg.xfIndexFor({ font: { bold: true } })
  const xfFmt = reg.xfIndexFor({ numberFormat: '0.00' })
  const xfAlign = reg.xfIndexFor({ alignment: { horizontal: 'center', vertical: 'middle' } })
  const xfFill = reg.xfIndexFor({
    fill: { type: 'pattern', pattern: 'solid', fgColor: 'FFFF0000' },
  })
  const xfBorder = reg.xfIndexFor({ border: { top: { style: 'thin', color: 'FF000000' } } })

  const parsed = readStyles(writeStylesXml(reg))
  expect(parsed.cellStyles[0]).toEqual({}) // default
  expect(parsed.cellStyles[xfBold]?.font?.bold).toBe(true)
  expect(parsed.cellStyles[xfFmt]?.numberFormat).toBe('0.00')
  expect(parsed.cellStyles[xfAlign]?.alignment).toEqual({
    horizontal: 'center',
    vertical: 'middle',
  })
  expect(parsed.cellStyles[xfFill]?.fill).toEqual({
    type: 'pattern',
    pattern: 'solid',
    fgColor: 'FFFF0000',
  })
  expect(parsed.cellStyles[xfBorder]?.border?.top).toEqual({ style: 'thin', color: 'FF000000' })
})

test('reads a custom number-format code via the numFmts table', () => {
  const reg = new StyleRegistry()
  const xf = reg.xfIndexFor({ numberFormat: 'yyyy-mm-dd' })
  const parsed = readStyles(writeStylesXml(reg))
  expect(parsed.cellStyles[xf]?.numberFormat).toBe('yyyy-mm-dd')
})
