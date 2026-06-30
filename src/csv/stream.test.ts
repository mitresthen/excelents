import { expect, test } from 'vitest'
import { readableToBytes } from '../io/streams'
import type { CellValue } from '../model/cell'
import { readCsv } from './reader'
import { readCsvRows, writeCsvStream } from './stream'
import { writeCsv } from './writer'
import { createWorkbook } from '../model/workbook'

const dec = (bytes: Uint8Array): string =>
  new TextDecoder('utf-8', { ignoreBOM: true }).decode(bytes)

async function* asyncRows(
  rows: ReadonlyArray<readonly CellValue[]>,
): AsyncGenerator<readonly CellValue[]> {
  for (const r of rows) yield r
}

async function collectRows(
  src: AsyncIterable<readonly CellValue[]>,
): Promise<Array<readonly CellValue[]>> {
  const out: Array<readonly CellValue[]> = []
  for await (const r of src) out.push(r)
  return out
}

test('writeCsvStream renders rows as RFC 4180 with quoting', async () => {
  const rows: CellValue[][] = [
    ['name', 'qty', 'note'],
    ['wid,get', 42, 'has "quote"'],
    ['multi\nline', -3.5, true],
  ]
  const text = dec(await readableToBytes(writeCsvStream(asyncRows(rows))))
  expect(text).toBe('name,qty,note\n"wid,get",42,"has ""quote"""\n"multi\nline",-3.5,true')
})

test('writeCsvStream output is byte-identical to writeCsv for the same grid', async () => {
  const rows: CellValue[][] = [
    ['a', 'b', 'c'],
    ['1', '2', '3'],
    ['x,y', 'p"q', 'z'],
  ]
  const wb = createWorkbook()
  const ws = wb.addSheet('S')
  rows.forEach((row, r) => row.forEach((v, c) => (ws.getCell(r + 1, c + 1).value = v)))
  const streamed = dec(await readableToBytes(writeCsvStream(asyncRows(rows))))
  expect(streamed).toBe(writeCsv(wb))
})

test('writeCsvStream honours delimiter, rowDelimiter and bom options', async () => {
  const rows: CellValue[][] = [
    ['a', 'b'],
    ['c', 'd'],
  ]
  const bytes = await readableToBytes(
    writeCsvStream(asyncRows(rows), { delimiter: ';', rowDelimiter: '\r\n', bom: true }),
  )
  expect([...bytes.slice(0, 3)]).toEqual([0xef, 0xbb, 0xbf]) // UTF-8 BOM
  expect(dec(bytes)).toBe('﻿a;b\r\nc;d')
})

test('writeCsvStream accepts a plain (sync) iterable too', async () => {
  const text = dec(await readableToBytes(writeCsvStream([['a', 1]])))
  expect(text).toBe('a,1')
})

test('readCsvRows yields the same rows readCsv produces (incl. inference)', async () => {
  const text = 'name,qty,flag\nfoo,42,true\n"bar,baz",00210,false'
  const viaRows = await collectRows(readCsvRows(text))
  const wb = readCsv(text)
  const ws = wb.sheets[0]!
  const viaModel = [1, 2, 3].map((r) => [1, 2, 3].map((c) => ws.getCell(r, c).value ?? null))
  // readCsvRows yields null for empty cells too, so align the comparison shape.
  expect(viaRows.map((r) => r.map((v) => v ?? null))).toEqual(viaModel)
})

test('readCsvRows round-trips writeCsvStream', async () => {
  const rows: CellValue[][] = [
    ['greeting', 'count'],
    ['hello, world', 7],
    ['line1\nline2', 0],
    ['tab\tsep', -1.5],
  ]
  const bytes = await readableToBytes(writeCsvStream(asyncRows(rows)))
  const back = await collectRows(readCsvRows(dec(bytes)))
  expect(back).toEqual([
    ['greeting', 'count'],
    ['hello, world', 7],
    ['line1\nline2', 0],
    ['tab\tsep', -1.5],
  ])
})

test('readCsvRows matches readCsv across chunk boundaries (every split point)', async () => {
  const text = '"a,b",1\r\n"c""d",true\nlast,00\r"q\nr",2'
  const reference = await collectRows(readCsvRows(text))
  const enc = new TextEncoder().encode(text)
  for (let split = 0; split <= enc.length; split++) {
    const stream = new ReadableStream<Uint8Array>({
      start(controller) {
        controller.enqueue(enc.slice(0, split))
        controller.enqueue(enc.slice(split))
        controller.close()
      },
    })
    const chunked = await collectRows(readCsvRows(stream))
    expect(chunked, `split at ${split}`).toEqual(reference)
  }
})

test('readCsvRows handles a multi-character delimiter split across chunks', async () => {
  const text = 'a||b||c\n1||2||3'
  const reference = await collectRows(readCsvRows(text, { delimiter: '||' }))
  expect(reference).toEqual([
    ['a', 'b', 'c'],
    [1, 2, 3],
  ])
  const enc = new TextEncoder().encode(text)
  for (let split = 0; split <= enc.length; split++) {
    const stream = new ReadableStream<Uint8Array>({
      start(controller) {
        controller.enqueue(enc.slice(0, split))
        controller.enqueue(enc.slice(split))
        controller.close()
      },
    })
    const chunked = await collectRows(readCsvRows(stream, { delimiter: '||' }))
    expect(chunked, `split at ${split}`).toEqual(reference)
  }
})

test('readCsvRows decodes multi-byte UTF-8 split across chunk boundaries', async () => {
  const text = 'café,Ω\nنعم,1'
  const enc = new TextEncoder().encode(text)
  for (let split = 0; split <= enc.length; split++) {
    const stream = new ReadableStream<Uint8Array>({
      start(controller) {
        controller.enqueue(enc.slice(0, split))
        controller.enqueue(enc.slice(split))
        controller.close()
      },
    })
    expect(await collectRows(readCsvRows(stream)), `split at ${split}`).toEqual([
      ['café', 'Ω'],
      ['نعم', 1],
    ])
  }
})

test('readCsvRows on empty input yields no rows', async () => {
  expect(await collectRows(readCsvRows(''))).toEqual([])
})
