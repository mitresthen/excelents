import { expect, test } from 'vitest'
import { createWorkbook } from '../model/workbook'
import { writeXlsx } from './write'

test('a bold cell with a number format round-trips through exceljs', async () => {
  const wb = createWorkbook()
  const ws = wb.addSheet('S')
  const c = ws.cell('A1')
  c.value = 1234.5
  c.style = { font: { bold: true }, numberFormat: '0.00' }
  const bytes = await writeXlsx(wb)

  const ExcelJS = (await import('exceljs')).default
  const oracle = new ExcelJS.Workbook()
  // @ts-expect-error -- @types/node 22 Buffer<ArrayBuffer> vs non-generic Buffer in exceljs decl
  await oracle.xlsx.load(Buffer.from(bytes))
  const cell = oracle.getWorksheet('S')?.getCell('A1')
  expect(cell?.value).toBe(1234.5)
  expect(cell?.font?.bold).toBe(true)
  expect(cell?.numFmt).toBe('0.00')
})

test('cell alignment round-trips through exceljs', async () => {
  const wb = createWorkbook()
  const ws = wb.addSheet('S')
  const c = ws.cell('A1')
  c.value = 'x'
  c.style = { alignment: { horizontal: 'center', vertical: 'middle', wrapText: true } }
  const bytes = await writeXlsx(wb)

  const ExcelJS = (await import('exceljs')).default
  const oracle = new ExcelJS.Workbook()
  // @ts-expect-error -- @types/node 22 Buffer<ArrayBuffer> vs non-generic Buffer in exceljs decl
  await oracle.xlsx.load(Buffer.from(bytes))
  const cell = oracle.getWorksheet('S')?.getCell('A1')
  // exceljs maps OOXML vertical="center" back to 'middle', matching our model.
  expect(cell?.alignment?.horizontal).toBe('center')
  expect(cell?.alignment?.vertical).toBe('middle')
  expect(cell?.alignment?.wrapText).toBe(true)
})

test('alignment participates in xf dedup distinctly from an unaligned style', async () => {
  const wb = createWorkbook()
  const ws = wb.addSheet('S')
  ws.cell('A1').value = 'plain'
  const aligned = ws.cell('A2')
  aligned.value = 'aligned'
  aligned.style = { alignment: { horizontal: 'right' } }
  const bytes = await writeXlsx(wb)

  const ExcelJS = (await import('exceljs')).default
  const oracle = new ExcelJS.Workbook()
  // @ts-expect-error -- @types/node 22 Buffer<ArrayBuffer> vs non-generic Buffer in exceljs decl
  await oracle.xlsx.load(Buffer.from(bytes))
  const sheet = oracle.getWorksheet('S')
  expect(sheet?.getCell('A1').alignment).toBeUndefined()
  expect(sheet?.getCell('A2').alignment?.horizontal).toBe('right')
})
