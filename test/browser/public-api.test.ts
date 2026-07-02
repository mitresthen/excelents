/**
 * Browser-mode smoke suite: exercises the public API in real Chromium, where the
 * platform primitives (CompressionStream deflate-raw, Web Streams, Blob) come from
 * the browser, not Node. Mirrors the README examples rather than re-testing codec
 * internals — those are covered by the node-project unit + conformance suites.
 */
import { describe, expect, it } from 'vitest'
import { createWorkbook, readXlsx, writeXlsx } from '../../src/index'
import { readCsv, writeCsv } from '../../src/csv'
import {
  createXlsxStreamWriter,
  readCsvRows,
  readXlsxRows,
  writeCsvStream,
  writeXlsxStream,
} from '../../src/stream'

// 1x1 red PNG
const PNG_1X1 =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=='

function buildWorkbook() {
  const wb = createWorkbook()
  const ws = wb.addSheet('Report')
  ws.cell('A1').value = 'Hello'
  ws.cell('B1').value = 42
  ws.cell('A2').value = new Date(Date.UTC(2026, 0, 1))
  ws.cell('A3').value = { formula: 'SUM(B1:B9)', result: 42 }
  ws.cell('A4').value = { text: 'docs', hyperlink: 'https://example.com' }
  ws.cell('A5').value = { richText: [{ text: 'bold', font: { bold: true } }, { text: ' plain' }] }
  ws.cell('A1').style = {
    font: { bold: true, color: 'FFFFFFFF' },
    fill: { type: 'pattern', pattern: 'solid', fgColor: 'FF4472C4' },
    alignment: { horizontal: 'right', indent: 1 },
    numberFormat: '#,##0.00',
  }
  ws.merge('A7:C7')
  ws.column(1).width = 24
  ws.getRow(1).height = 30
  ws.freeze({ rows: 1 })
  ws.setAutoFilter('A1:B1')
  const logo = wb.addImage({ data: PNG_1X1, extension: 'png' })
  ws.placeImage(logo, { tl: 'B3', size: { width: 18, height: 18 } })

  const data = wb.addSheet('Data')
  data.addRow(['Region', 'Qty'])
  data.addRow(['North', 1])
  data.addRow(['South', 2])
  data.addTable({ name: 'Sales', ref: 'A1:B3', columns: ['Region', 'Qty'] })
  data.addDataValidation({ sqref: 'B2:B100', type: 'list', formula1: '"1,2,3"' })
  wb.defineName('TaxRate', 'Report!$B$1')
  return wb
}

describe('xlsx in the browser', () => {
  it('round-trips a featureful workbook via writeXlsx/readXlsx', async () => {
    const source = buildWorkbook()
    // write-side-only features live on the model (serialized, but not parsed back)
    expect(source.getSheet('Report')?.frozen).toEqual({ rows: 1, cols: 0 })
    expect(source.getSheet('Report')?.autoFilter).toBe('A1:B1')
    expect(source.getSheet('Report')?.images).toHaveLength(1)

    const bytes = await writeXlsx(source)
    const wb = await readXlsx(bytes)

    const ws = wb.getSheet('Report')
    expect(ws?.cell('A1').value).toBe('Hello')
    expect(ws?.cell('A2').type).toBe('date')
    expect(ws?.cell('A3').value).toEqual({ formula: 'SUM(B1:B9)', result: 42 })
    expect(ws?.cell('A1').style.fill?.fgColor).toBe('FF4472C4')
    expect(ws?.merges).toContain('A7:C7')

    const data = wb.getSheet('Data')
    expect(data?.tables).toHaveLength(1)
    expect(data?.dataValidations).toHaveLength(1)
    expect(wb.definedNames.some((d) => d.name === 'TaxRate')).toBe(true)
  })

  it('streams rows out and back in, including from a Blob', async () => {
    async function* rows() {
      yield ['name', 'qty']
      yield ['widget', 42]
    }
    const writer = createXlsxStreamWriter({ sheet: 'Data' })
    const collecting = (async () => {
      const chunks: Uint8Array[] = []
      for await (const chunk of writer.readable) chunks.push(chunk)
      return chunks
    })()
    await writer.addRow(['name', 'qty'])
    await writer.addRow(['widget', 42])
    await writer.close()
    const chunks = await collecting
    const bytes = new Uint8Array(chunks.reduce((n, c) => n + c.length, 0))
    let offset = 0
    for (const chunk of chunks) {
      bytes.set(chunk, offset)
      offset += chunk.length
    }
    const blob = new Blob([bytes])

    const seen: unknown[][] = []
    for await (const row of readXlsxRows(blob)) seen.push([...row.cells])
    expect(seen).toEqual([
      ['name', 'qty'],
      ['widget', 42],
    ])

    // functional form consumed as a plain ReadableStream
    const stream = writeXlsxStream(rows())
    expect(stream).toBeInstanceOf(ReadableStream)
    const roundTripped = []
    for await (const row of readXlsxRows(stream)) roundTripped.push(row.cells[0])
    expect(roundTripped).toEqual(['name', 'widget'])
  })
})

describe('csv in the browser', () => {
  it('round-trips buffered and streaming CSV', async () => {
    const wb = readCsv('name,qty\nwidget,42')
    expect(wb.sheets[0]?.cell('B2').value).toBe(42)
    expect(writeCsv(wb)).toBe('name,qty\nwidget,42')

    async function* rows() {
      yield ['name', 'qty']
      yield ['widget', 42]
    }
    const out: unknown[][] = []
    for await (const row of readCsvRows(writeCsvStream(rows()))) out.push([...row])
    expect(out).toEqual([
      ['name', 'qty'],
      ['widget', 42],
    ])
  })
})
