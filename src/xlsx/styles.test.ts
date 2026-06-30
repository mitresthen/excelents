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
