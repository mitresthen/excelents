import { expect, test } from 'vitest'
import { createWorkbook } from '../model/workbook'
import { readXlsx } from './read'
import { writeXlsx } from './write'

test('a table round-trips through write+read', async () => {
  const wb = createWorkbook()
  const ws = wb.addSheet('S')
  ws.cell('A1').value = 'Col1'
  ws.cell('B1').value = 'Col2'
  ws.cell('C1').value = 'Col3'
  ws.addTable({ name: 'Table1', ref: 'A1:C2', columns: ['Col1', 'Col2', 'Col3'] })
  const r = await readXlsx(await writeXlsx(wb))
  expect(r.sheets[0]!.tables).toEqual([
    { name: 'Table1', ref: 'A1:C2', columns: ['Col1', 'Col2', 'Col3'] },
  ])
})

test('a styled table round-trips its style name', async () => {
  const wb = createWorkbook()
  const ws = wb.addSheet('S')
  ws.addTable({
    name: 'T2',
    ref: 'A1:B3',
    columns: ['H1', 'H2'],
    styleName: 'TableStyleLight1',
  })
  const r = await readXlsx(await writeXlsx(wb))
  expect(r.sheets[0]!.tables[0]).toEqual({
    name: 'T2',
    ref: 'A1:B3',
    columns: ['H1', 'H2'],
    styleName: 'TableStyleLight1',
  })
})

test('exceljs reads our table back', async () => {
  const wb = createWorkbook()
  const ws = wb.addSheet('S')
  ws.cell('A1').value = 'H1'
  ws.cell('B1').value = 'H2'
  ws.addTable({ name: 'T', ref: 'A1:B2', columns: ['H1', 'H2'] })
  const bytes = await writeXlsx(wb)
  const ExcelJS = (await import('exceljs')).default
  const oracle = new ExcelJS.Workbook()
  // @ts-expect-error -- @types/node 22 Buffer<ArrayBuffer> vs non-generic Buffer in exceljs decl
  await oracle.xlsx.load(Buffer.from(bytes))
  const t = oracle.getWorksheet('S')?.getTable('T')
  expect(t).toBeDefined()
})
