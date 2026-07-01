import { expect, test } from 'vitest'
import { createWorkbook } from '../model/workbook'
import { writeXlsx } from './write'

// A 1x1 transparent PNG (base64) — also exercises the base64 decode path in addImage.
const PNG_1X1 =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=='

async function loadOracle(bytes: Uint8Array): Promise<import('exceljs').Workbook> {
  const ExcelJS = (await import('exceljs')).default
  const oracle = new ExcelJS.Workbook()
  // @ts-expect-error -- @types/node 22 Buffer<ArrayBuffer> vs non-generic Buffer in exceljs decl
  await oracle.xlsx.load(Buffer.from(bytes))
  return oracle
}

test('autoFilter round-trips through exceljs', async () => {
  const wb = createWorkbook()
  const ws = wb.addSheet('S')
  ws.cell('A9').value = 'H'
  ws.setAutoFilter('A9:F123')
  const oracle = await loadOracle(await writeXlsx(wb))
  expect(oracle.getWorksheet('S')?.autoFilter).toBe('A9:F123')
})

test('frozen header rows round-trip through exceljs', async () => {
  const wb = createWorkbook()
  const ws = wb.addSheet('S')
  ws.cell('A1').value = 'H'
  ws.freeze({ rows: 9 })
  const oracle = await loadOracle(await writeXlsx(wb))
  // exceljs's WorksheetView union doesn't surface ySplit on the narrowed type.
  const view = oracle.getWorksheet('S')?.views?.[0] as
    | { state?: string; ySplit?: number }
    | undefined
  expect(view?.state).toBe('frozen')
  expect(view?.ySplit).toBe(9)
})

test('alignment indent round-trips through exceljs', async () => {
  const wb = createWorkbook()
  const ws = wb.addSheet('S')
  const c = ws.cell('A1')
  c.value = 'x'
  c.style = { alignment: { horizontal: 'right', indent: 2 } }
  const oracle = await loadOracle(await writeXlsx(wb))
  expect(oracle.getWorksheet('S')?.getCell('A1').alignment?.indent).toBe(2)
})

test('embedded image round-trips through exceljs (anchor, size, media)', async () => {
  const wb = createWorkbook()
  const ws = wb.addSheet('S')
  ws.cell('A1').value = 'x'
  const id = wb.addImage({ data: PNG_1X1, extension: 'png' })
  ws.placeImage(id, { tl: 'F1', size: { width: 180, height: 101 } })

  const oracle = await loadOracle(await writeXlsx(wb))
  const images = oracle.getWorksheet('S')?.getImages() ?? []
  expect(images.length).toBe(1)
  // F1 -> col 5, row 0 (0-indexed)
  expect(images[0]?.range.tl.nativeCol).toBe(5)
  expect(images[0]?.range.tl.nativeRow).toBe(0)
  // @ts-expect-error -- exceljs's ImageRange type omits `ext`, present at runtime for a oneCellAnchor.
  expect(images[0]?.range.ext).toEqual({ width: 180, height: 101 })
  expect(oracle.model.media.length).toBe(1)
})

test('all four features coexist in one workbook and open cleanly', async () => {
  const wb = createWorkbook()
  const ws = wb.addSheet('Report')
  ws.cell('A9').value = 'Header'
  ws.cell('A10').value = 'data'
  ws.cell('A9').style = { alignment: { horizontal: 'right', indent: 1 } }
  ws.setAutoFilter('A9:F10')
  ws.freeze({ rows: 9 })
  const id = wb.addImage({ data: PNG_1X1, extension: 'png' })
  ws.placeImage(id, { tl: 'F1', size: { width: 180, height: 101 } })

  const oracle = await loadOracle(await writeXlsx(wb))
  const sheet = oracle.getWorksheet('Report')
  expect(sheet?.autoFilter).toBe('A9:F10')
  expect(sheet?.views?.[0]?.state).toBe('frozen')
  expect(sheet?.getCell('A9').alignment?.indent).toBe(1)
  expect(sheet?.getImages().length).toBe(1)
})
