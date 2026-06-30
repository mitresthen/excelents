import { expect, test } from 'vitest'
import { readableToBytes } from './io/streams'
import type { CellValue } from './model/cell'
import * as streamEntry from './stream'
import {
  createXlsxStreamWriter,
  readCsvRows,
  readXlsxRows,
  writeCsvStream,
  writeXlsxStream,
} from './stream'

test('the ./stream entry exposes the full streaming surface', () => {
  expect(typeof streamEntry.writeXlsxStream).toBe('function')
  expect(typeof streamEntry.createXlsxStreamWriter).toBe('function')
  expect(typeof streamEntry.readXlsxRows).toBe('function')
  expect(typeof streamEntry.writeCsvStream).toBe('function')
  expect(typeof streamEntry.readCsvRows).toBe('function')
})

test('xlsx write→read round-trips through the streaming surface', async () => {
  const rows: CellValue[][] = [
    ['a', 1, true],
    ['b', 2.5, false],
  ]
  const bytes = await readableToBytes(writeXlsxStream(rows, { sheet: 'S' }))
  const out: CellValue[][] = []
  for await (const row of readXlsxRows(bytes)) out.push([...row.cells])
  expect(out).toEqual(rows)
})

test('the builder writer is wired through ./stream and reads back', async () => {
  const writer = createXlsxStreamWriter({ sheet: 'B' })
  const collect = readableToBytes(writer.readable)
  await writer.addRow(['only'])
  await writer.close()
  const rows = []
  for await (const row of readXlsxRows(await collect)) rows.push(row)
  expect(rows[0]!.sheet).toBe('B')
  expect(rows[0]!.cells).toEqual(['only'])
})

test('csv write→read round-trips through the streaming surface', async () => {
  async function* source(): AsyncGenerator<CellValue[]> {
    yield ['x', 1]
    yield ['y', 2]
  }
  const bytes = await readableToBytes(writeCsvStream(source()))
  const out: (readonly CellValue[])[] = []
  for await (const row of readCsvRows(new TextDecoder().decode(bytes))) out.push(row)
  expect(out).toEqual([
    ['x', 1],
    ['y', 2],
  ])
})
