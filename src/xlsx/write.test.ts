import { expect, test } from 'vitest'
import { createWorkbook } from '../model/workbook'
import { writeXlsx } from './write'

test('writeXlsx produces a zip exceljs can open and read back values', async () => {
  const wb = createWorkbook()
  const ws = wb.addSheet('Sheet1')
  ws.cell('A1').value = 'Hello'
  ws.cell('B1').value = 42
  ws.cell('A2').value = true

  const bytes = await writeXlsx(wb)

  const ExcelJS = (await import('exceljs')).default
  const oracle = new ExcelJS.Workbook()
  // @ts-expect-error -- @types/node 22 Buffer<ArrayBuffer> vs non-generic Buffer in exceljs decl
  await oracle.xlsx.load(Buffer.from(bytes))
  const sheet = oracle.getWorksheet('Sheet1')
  expect(sheet?.getCell('A1').value).toBe('Hello')
  expect(sheet?.getCell('B1').value).toBe(42)
  expect(sheet?.getCell('A2').value).toBe(true)
})

test('writeXlsx round-trips through our own OPC reader', async () => {
  const wb = createWorkbook()
  wb.addSheet('S').cell('A1').value = 'x'
  const { OpcPackage } = await import('../opc/package')
  const pkg = await OpcPackage.read(await writeXlsx(wb))
  expect(pkg.partNames()).toContain('xl/workbook.xml')
  expect(pkg.partNames()).toContain('xl/worksheets/sheet1.xml')
})
