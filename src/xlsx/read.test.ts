import { expect, test } from 'vitest'
import { createWorkbook } from '../model/workbook'
import { readXlsx } from './read'
import { writeXlsx } from './write'

test('readXlsx recovers sheet names round-tripped from writeXlsx', async () => {
  const wb = createWorkbook()
  wb.addSheet('Alpha')
  wb.addSheet('Beta')
  const restored = await readXlsx(await writeXlsx(wb))
  expect(restored.sheets.map((s) => s.name)).toEqual(['Alpha', 'Beta'])
})
