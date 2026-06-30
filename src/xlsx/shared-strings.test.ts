import { expect, test } from 'vitest'
import { createWorkbook } from '../model/workbook'
import { writeXlsx } from './write'

test('string cells round-trip via shared strings (exceljs reads them)', async () => {
  const wb = createWorkbook()
  const ws = wb.addSheet('S')
  ws.cell('A1').value = 'apple'
  ws.cell('A2').value = 'banana'
  ws.cell('A3').value = 'apple' // duplicate → same shared index
  const bytes = await writeXlsx(wb)

  const ExcelJS = (await import('exceljs')).default
  const oracle = new ExcelJS.Workbook()
  // @ts-expect-error -- @types/node 22 Buffer<ArrayBuffer> vs non-generic Buffer in exceljs decl
  await oracle.xlsx.load(Buffer.from(bytes))
  const sheet = oracle.getWorksheet('S')
  expect(sheet?.getCell('A1').value).toBe('apple')
  expect(sheet?.getCell('A2').value).toBe('banana')
  expect(sheet?.getCell('A3').value).toBe('apple')

  // the part exists and dedups (2 unique strings)
  const { OpcPackage } = await import('../opc/package')
  const pkg = await OpcPackage.read(bytes)
  const sst = new TextDecoder().decode(pkg.getPart('xl/sharedStrings.xml'))
  expect(sst).toContain('uniqueCount="2"')
})
