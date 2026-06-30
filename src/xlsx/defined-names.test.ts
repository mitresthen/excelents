import { expect, test } from 'vitest'
import { createWorkbook } from '../model/workbook'
import { readXlsx } from './read'
import { writeXlsx } from './write'

test('a workbook records defined names', () => {
  const wb = createWorkbook()
  wb.addSheet('Sheet1')
  wb.defineName('MyRange', 'Sheet1!$A$1:$B$2')
  expect(wb.definedNames).toEqual([{ name: 'MyRange', formula: 'Sheet1!$A$1:$B$2' }])
})

test('defined names round-trip through write+read', async () => {
  const wb = createWorkbook()
  wb.addSheet('Sheet1')
  wb.defineName('MyRange', 'Sheet1!$A$1:$B$2')
  wb.defineName('Tax', '0.2')
  const restored = await readXlsx(await writeXlsx(wb))
  expect(restored.definedNames).toEqual([
    { name: 'MyRange', formula: 'Sheet1!$A$1:$B$2' },
    { name: 'Tax', formula: '0.2' },
  ])
})

test('a sheet-scoped defined name round-trips its localSheetId (no collision with a global twin)', async () => {
  // A scoped name and a global name may share a label; both must survive distinctly, else
  // re-serializing yields two colliding workbook-global names that Excel rejects.
  const wb = createWorkbook()
  wb.addSheet('S')
  wb.defineName('Leiter', 'S!$A$1', 0)
  wb.defineName('Leiter', 'S!$B$1')
  const restored = await readXlsx(await writeXlsx(wb))
  expect(restored.definedNames).toEqual([
    { name: 'Leiter', formula: 'S!$A$1', localSheetId: 0 },
    { name: 'Leiter', formula: 'S!$B$1' },
  ])
})
