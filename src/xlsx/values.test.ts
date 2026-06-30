import { expect, test } from 'vitest'
import { createWorkbook } from '../model/workbook'
import { writeXlsx } from './write'

async function readBack(bytes: Uint8Array, sheet = 'S') {
  const ExcelJS = (await import('exceljs')).default
  const oracle = new ExcelJS.Workbook()
  // @ts-expect-error -- @types/node 22 Buffer<ArrayBuffer> vs non-generic Buffer in exceljs decl
  await oracle.xlsx.load(Buffer.from(bytes))
  return oracle.getWorksheet(sheet)
}

test('a Date value round-trips through exceljs as a Date', async () => {
  const wb = createWorkbook()
  const ws = wb.addSheet('S')
  const when = new Date(Date.UTC(2026, 5, 30, 12, 0, 0))
  ws.cell('A1').value = when
  const bytes = await writeXlsx(wb)

  const value = (await readBack(bytes))?.getCell('A1')?.value
  expect(value).toBeInstanceOf(Date)
  expect(value).toEqual(when)
})

test('a non-finite number is written as an error cell, not invalid XML', async () => {
  const wb = createWorkbook()
  const ws = wb.addSheet('S')
  ws.cell('A1').value = NaN
  ws.cell('A2').value = Infinity
  const bytes = await writeXlsx(wb)

  const sheet = await readBack(bytes)
  for (const ref of ['A1', 'A2']) {
    const value = sheet?.getCell(ref).value
    if (value === null || typeof value !== 'object' || !('error' in value)) {
      throw new Error(`expected an error cell value at ${ref}`)
    }
    expect(value.error).toBe('#NUM!')
  }
})

test('a formula with a numeric result round-trips through exceljs', async () => {
  const wb = createWorkbook()
  const ws = wb.addSheet('S')
  ws.cell('A1').value = { formula: '1+2', result: 3 }
  const bytes = await writeXlsx(wb)

  const cell = (await readBack(bytes))?.getCell('A1')
  expect(cell?.formula).toBe('1+2')
  expect(cell?.result).toBe(3)
})

test('a formula with a string result round-trips through exceljs', async () => {
  const wb = createWorkbook()
  const ws = wb.addSheet('S')
  ws.cell('A1').value = { formula: 'CONCAT("a","b")', result: 'ab' }
  const bytes = await writeXlsx(wb)

  const cell = (await readBack(bytes))?.getCell('A1')
  expect(cell?.formula).toBe('CONCAT("a","b")')
  expect(cell?.result).toBe('ab')
})

test('a merged range round-trips through exceljs', async () => {
  const wb = createWorkbook()
  const ws = wb.addSheet('S')
  ws.cell('A1').value = 'merged'
  ws.merge('A1:B2')
  const bytes = await writeXlsx(wb)

  const sheet = await readBack(bytes)
  expect(sheet?.getCell('A1').isMerged).toBe(true)
  expect(sheet?.getCell('B2').isMerged).toBe(true)
  expect(sheet?.getCell('B2').master.address).toBe('A1')
})

test('column width and row height round-trip through exceljs', async () => {
  const wb = createWorkbook()
  const ws = wb.addSheet('S')
  ws.cell('A1').value = 'x'
  ws.column(1).width = 20
  ws.getRow(1).height = 30
  const bytes = await writeXlsx(wb)

  const sheet = await readBack(bytes)
  expect(sheet?.getColumn(1).width).toBe(20)
  expect(sheet?.getRow(1).height).toBe(30)
})

test('rich text with mixed run formatting round-trips through exceljs', async () => {
  const wb = createWorkbook()
  const ws = wb.addSheet('S')
  ws.cell('A1').value = {
    richText: [{ text: 'Hello ' }, { text: 'World', font: { bold: true } }],
  }
  const bytes = await writeXlsx(wb)

  const value = (await readBack(bytes))?.getCell('A1')?.value
  if (value === null || typeof value !== 'object' || !('richText' in value)) {
    throw new Error('expected a rich-text cell value')
  }
  const runs = value.richText
  expect(runs.map((r) => r.text).join('')).toBe('Hello World')
  expect(runs).toHaveLength(2)
  expect(runs[1]?.font?.bold).toBe(true)
})

test('a hyperlink round-trips through exceljs as cell.hyperlink + cell.text', async () => {
  const wb = createWorkbook()
  const ws = wb.addSheet('S')
  ws.cell('A1').value = { text: 'Anthropic', hyperlink: 'https://www.anthropic.com/' }
  const bytes = await writeXlsx(wb)

  const cell = (await readBack(bytes))?.getCell('A1')
  expect(cell?.hyperlink).toBe('https://www.anthropic.com/')
  expect(cell?.text).toBe('Anthropic')
})

test('two hyperlinks on one sheet get distinct relationship ids', async () => {
  const wb = createWorkbook()
  const ws = wb.addSheet('S')
  ws.cell('A1').value = { text: 'a', hyperlink: 'https://a.example/' }
  ws.cell('A2').value = { text: 'b', hyperlink: 'https://b.example/' }
  const bytes = await writeXlsx(wb)

  const sheet = await readBack(bytes)
  expect(sheet?.getCell('A1').hyperlink).toBe('https://a.example/')
  expect(sheet?.getCell('A2').hyperlink).toBe('https://b.example/')
})
