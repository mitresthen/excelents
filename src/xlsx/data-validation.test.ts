import { expect, test } from 'vitest'
import { createWorkbook } from '../model/workbook'
import { readXlsx } from './read'
import { writeXlsx } from './write'

test('a list data validation round-trips through write+read', async () => {
  const wb = createWorkbook()
  const ws = wb.addSheet('S')
  ws.cell('A1').value = 'pick'
  ws.addDataValidation({
    sqref: 'A1:A10',
    type: 'list',
    formula1: '"apple,banana,cherry"',
    allowBlank: true,
  })
  const r = await readXlsx(await writeXlsx(wb))
  expect(r.sheets[0]!.dataValidations).toEqual([
    { sqref: 'A1:A10', type: 'list', formula1: '"apple,banana,cherry"', allowBlank: true },
  ])
})

test('a numeric between validation round-trips with operator and two formulas', async () => {
  const wb = createWorkbook()
  const ws = wb.addSheet('S')
  ws.addDataValidation({
    sqref: 'B1:B5',
    type: 'whole',
    operator: 'between',
    formula1: '1',
    formula2: '100',
  })
  const r = await readXlsx(await writeXlsx(wb))
  expect(r.sheets[0]!.dataValidations[0]).toMatchObject({
    sqref: 'B1:B5',
    type: 'whole',
    operator: 'between',
    formula1: '1',
    formula2: '100',
  })
})

test('showDropDown:false round-trips (OOXML stores the inverted flag)', async () => {
  const wb = createWorkbook()
  const ws = wb.addSheet('S')
  ws.addDataValidation({
    sqref: 'A1',
    type: 'list',
    formula1: '"x,y"',
    showDropDown: false,
  })
  const r = await readXlsx(await writeXlsx(wb))
  expect(r.sheets[0]!.dataValidations[0]?.showDropDown).toBe(false)
})
