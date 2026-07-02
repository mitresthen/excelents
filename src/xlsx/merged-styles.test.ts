import { expect, test } from 'vitest'
import { createWorkbook } from '../model/workbook'
import { readXlsx } from './read'
import { writeXlsx } from './write'

const BLUE = { type: 'pattern', pattern: 'solid', fgColor: 'FF005BA1' } as const

test('a fill set on the master after merging paints the whole merged range', async () => {
  const wb = createWorkbook()
  const ws = wb.addSheet('S')
  ws.merge('A1:C2')
  const master = ws.cell('A1')
  master.value = 'Title'
  master.style.fill = { ...BLUE }
  const bytes = await writeXlsx(wb)

  const ExcelJS = (await import('exceljs')).default
  const oracle = new ExcelJS.Workbook()
  // @ts-expect-error -- @types/node 22 Buffer<ArrayBuffer> vs non-generic Buffer in exceljs decl
  await oracle.xlsx.load(Buffer.from(bytes))
  const sheet = oracle.getWorksheet('S')
  for (const ref of ['A1', 'B1', 'C1', 'A2', 'B2', 'C2']) {
    const fill = sheet?.getCell(ref).fill
    expect(fill, `${ref} carries the merged fill`).toMatchObject({
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF005BA1' },
    })
  }
  expect(sheet?.getCell('A1').value).toBe('Title')
})

test('merged cells share the master style object, before or after styling', () => {
  const wb = createWorkbook()
  const ws = wb.addSheet('S')
  ws.cell('A1').style.font = { bold: true } // styled BEFORE the merge
  ws.merge('A1:B1')
  expect(ws.cell('B1').style).toBe(ws.cell('A1').style)
  expect(ws.cell('B1').style.font?.bold).toBe(true)

  ws.merge('A3:B3') // styled AFTER the merge
  ws.cell('A3').style.fill = { ...BLUE }
  expect(ws.cell('B3').style.fill).toEqual(BLUE)
})

test('merging materializes its rows, so addRow lands below the merged block', () => {
  const wb = createWorkbook()
  const ws = wb.addSheet('S')
  ws.cell('A1').value = 'Title'
  ws.merge('A2:F2') // spacer row with no values, exceljs-style
  const header = ws.addRow(['a', 'b'])
  expect(header.number).toBe(3)
})

test('a styled blank cell is written and read back by exceljs', async () => {
  const wb = createWorkbook()
  const ws = wb.addSheet('S')
  ws.cell('A1').value = 'x'
  ws.cell('A2').style.fill = { ...BLUE } // no value, only a fill
  const bytes = await writeXlsx(wb)

  const ExcelJS = (await import('exceljs')).default
  const oracle = new ExcelJS.Workbook()
  // @ts-expect-error -- @types/node 22 Buffer<ArrayBuffer> vs non-generic Buffer in exceljs decl
  await oracle.xlsx.load(Buffer.from(bytes))
  const cell = oracle.getWorksheet('S')?.getCell('A2')
  expect(cell?.value).toBeNull()
  expect(cell?.fill).toMatchObject({ fgColor: { argb: 'FF005BA1' } })
})

test('a styled blank cell round-trips through readXlsx', async () => {
  const wb = createWorkbook()
  const ws = wb.addSheet('S')
  ws.cell('A1').value = 'x'
  ws.cell('A2').style.fill = { ...BLUE }
  const bytes = await writeXlsx(wb)

  const restored = await readXlsx(bytes)
  const cell = restored.sheets[0]?.cell('A2')
  expect(cell?.value).toBeNull()
  expect(cell?.style.fill).toEqual(BLUE)
})
