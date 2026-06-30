import { expect, test } from 'vitest'
import {
  compareWriteParts,
  extractSheetContent,
  unzipToParts,
} from '../../test/conformance/harness'
import { createWorkbook } from '../model/workbook'
import { writeXlsx } from './write'

/**
 * Build the same representative workbook with both producers: a few strings, numbers,
 * a bold cell, a number-format cell, and a merge. Returns serialized .xlsx bytes.
 */
async function build(lib: 'oracle' | 'excelents'): Promise<Uint8Array> {
  if (lib === 'excelents') {
    const wb = createWorkbook()
    const ws = wb.addSheet('S')
    ws.cell('A1').value = 'apple'
    ws.cell('A2').value = 'banana'
    ws.cell('B1').value = 42
    ws.cell('B2').value = 1234.5
    const bold = ws.cell('C1')
    bold.value = 'bold'
    bold.style = { font: { bold: true } }
    const fmt = ws.cell('C2')
    fmt.value = 9.99
    fmt.style = { numberFormat: '0.00' }
    ws.cell('A4').value = 'merged'
    ws.merge('A4:B5')
    return writeXlsx(wb)
  }
  const ExcelJS = (await import('exceljs')).default
  const wb = new ExcelJS.Workbook()
  const ws = wb.addWorksheet('S')
  ws.getCell('A1').value = 'apple'
  ws.getCell('A2').value = 'banana'
  ws.getCell('B1').value = 42
  ws.getCell('B2').value = 1234.5
  ws.getCell('C1').value = 'bold'
  ws.getCell('C1').font = { bold: true }
  ws.getCell('C2').value = 9.99
  ws.getCell('C2').numFmt = '0.00'
  ws.getCell('A4').value = 'merged'
  ws.mergeCells('A4:B5')
  const buf = await wb.xlsx.writeBuffer()
  return new Uint8Array(buf as ArrayBuffer)
}

test('worksheet cell content matches exceljs for a representative workbook', async () => {
  const ours = await extractSheetContent(await build('excelents'))
  const theirs = await extractSheetContent(await build('oracle'))
  // Guard against a vacuous pass ({} deep-equals {}): the representative sheet has
  // exactly 7 non-empty cells (A1,A2,B1,B2,C1,C2,A4). If extraction silently
  // collapsed, this fails instead of trivially matching.
  expect(Object.keys(ours.cells)).toHaveLength(7)
  // Semantic equivalence: same resolved cell values at the same refs (shared-string
  // index ordering and incidental styling differences are normalized away).
  expect(ours.cells).toEqual(theirs.cells)
})

test('merge ranges match exceljs', async () => {
  const ours = await extractSheetContent(await build('excelents'))
  const theirs = await extractSheetContent(await build('oracle'))
  expect(ours.merges).toEqual(theirs.merges)
})

test('both producers emit the worksheet and sharedStrings parts', async () => {
  // Assert presence DIRECTLY (not inferred from an empty diff, which would pass
  // vacuously if neither side emitted the part).
  const ours = Object.keys(await unzipToParts(await build('excelents')))
  const theirs = Object.keys(await unzipToParts(await build('oracle')))
  for (const part of ['xl/worksheets/sheet1.xml', 'xl/sharedStrings.xml']) {
    expect(ours).toContain(part)
    expect(theirs).toContain(part)
  }
  // Exercise the comparator over the owned sharedStrings part: both emit it, so any
  // diff is 'content differs' (incidental index ordering), never 'missing on …'.
  const diffs = await compareWriteParts(build, (p) => p === 'xl/sharedStrings.xml')
  for (const d of diffs) {
    expect(d.detail).toBe('content differs')
  }
})
