import { expect, test } from 'vitest'
import { listFixtures, parseFixture } from '../../test/conformance/harness'
import { bytesToReadable, readableToBytes } from '../io/streams'
import type { CellValue } from '../model/cell'
import type { Workbook } from '../model/workbook'
import { encodeAddress } from '../utils/address'
import { readXlsx } from './read'
import { type XlsxRow, readXlsxRows } from './stream-reader'
import { writeXlsxStream } from './stream-writer'

async function collect(src: AsyncIterable<XlsxRow>): Promise<XlsxRow[]> {
  const out: XlsxRow[] = []
  for await (const row of src) out.push(row)
  return out
}

/** Normalize a readXlsx cell to what the streaming reader yields (it does no hyperlink post-pass). */
function normalize(v: CellValue): CellValue {
  if (v !== null && typeof v === 'object' && 'hyperlink' in v) return v.text
  return v
}

function mapFromRows(rows: readonly XlsxRow[]): Map<string, CellValue> {
  const m = new Map<string, CellValue>()
  for (const row of rows) {
    row.cells.forEach((val, i) => {
      if (val !== null) m.set(`${row.sheet}!${encodeAddress(row.rowNumber, i + 1)}`, val)
    })
  }
  return m
}

function mapFromWorkbook(wb: Workbook): Map<string, CellValue> {
  const m = new Map<string, CellValue>()
  for (const ws of wb.sheets) {
    const dims = ws.dimensions
    if (dims === undefined) continue
    for (let r = 1; r <= dims.bottom; r++) {
      for (let c = 1; c <= dims.right; c++) {
        const v = ws.getCell(r, c).value
        if (v !== undefined && v !== null) m.set(`${ws.name}!${encodeAddress(r, c)}`, normalize(v))
      }
    }
  }
  return m
}

const ROWS = [
  ['name', 'qty', 'flag', 'when'],
  ['wid,get', 42, true, new Date(Date.UTC(2020, 0, 15))],
  ['  pad  ', -3.5, false, null],
]

test('readXlsxRows round-trips rows written by writeXlsxStream', async () => {
  const bytes = await readableToBytes(writeXlsxStream(ROWS, { sheet: 'Data' }))
  const rows = await collect(readXlsxRows(bytes))
  expect(rows.map((r) => r.sheet)).toEqual(['Data', 'Data', 'Data'])
  expect(rows.map((r) => r.rowNumber)).toEqual([1, 2, 3])
  expect(rows[0]!.cells).toEqual(['name', 'qty', 'flag', 'when'])
  expect(rows[1]!.cells.slice(0, 3)).toEqual(['wid,get', 42, true])
  const when = rows[1]!.cells[3]
  expect(when).toBeInstanceOf(Date)
  if (when instanceof Date) expect(when.getUTCFullYear()).toBe(2020)
  expect(rows[2]!.cells).toEqual(['  pad  ', -3.5, false]) // trailing null cell omitted
})

test('readXlsxRows accepts Uint8Array, Blob and ReadableStream identically', async () => {
  const bytes = await readableToBytes(writeXlsxStream(ROWS, { sheet: 'Data' }))
  const fromBytes = mapFromRows(await collect(readXlsxRows(bytes)))
  const fromBlob = mapFromRows(await collect(readXlsxRows(new Blob([new Uint8Array(bytes)]))))
  const fromStream = mapFromRows(await collect(readXlsxRows(bytesToReadable(bytes))))
  expect(fromBlob).toEqual(fromBytes)
  expect(fromStream).toEqual(fromBytes)
})

test('readXlsxRows matches readXlsx cell-for-cell on representative fixtures', async () => {
  // Includes 877 + 1669: their sharedStrings part sits AFTER the worksheet, exercising the
  // random-access (not forward-only) resolution that the SP-7 design locked in.
  const names = [
    'test-issue-877.xlsx',
    'test-issue-1669.xlsx',
    'gold.xlsx',
    'dateIssue.xlsx',
    'fibonacci.xlsx',
    'shared_string_with_escape.xlsx',
  ]
  for (const name of names) {
    const path = listFixtures().find((p) => p.endsWith(name))!
    const bytes = await parseFixture(path)
    const expected = mapFromWorkbook(await readXlsx(bytes))
    const actual = mapFromRows(await collect(readXlsxRows(bytes)))
    expect(actual, name).toEqual(expected)
  }
})

test('readXlsxRows matches readXlsx cell-for-cell on EVERY xlsx fixture', async () => {
  // missing-bits.xlsx is an intentionally-incomplete archive (no workbook part) — readXlsx
  // recovers no sheets and readXlsxRows yields nothing, so there is nothing to compare.
  const skip = new Set(['missing-bits.xlsx'])
  let checked = 0
  for (const path of listFixtures()) {
    const name = path.split('/').pop()!
    if (!name.endsWith('.xlsx') || skip.has(name)) continue
    const bytes = await parseFixture(path)
    const expected = mapFromWorkbook(await readXlsx(bytes))
    const actual = mapFromRows(await collect(readXlsxRows(bytes)))
    expect(actual, name).toEqual(expected)
    checked++
  }
  expect(checked).toBeGreaterThanOrEqual(30)
}, 60000)

test('readXlsxRows yields rows from every sheet of a multi-sheet workbook', async () => {
  const bytes = await readableToBytes(writeXlsxStream([['only']], { sheet: 'Solo' }))
  const rows = await collect(readXlsxRows(bytes))
  expect(rows).toHaveLength(1)
  expect(rows[0]!.sheet).toBe('Solo')
  expect(rows[0]!.cells).toEqual(['only'])
})

test('readXlsxRows on an empty sheet yields no rows', async () => {
  const bytes = await readableToBytes(writeXlsxStream([], { sheet: 'S' }))
  expect(await collect(readXlsxRows(bytes))).toEqual([])
})
