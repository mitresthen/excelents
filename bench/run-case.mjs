/**
 * Runs a single benchmark case in this process and prints JSON to stdout:
 *   { ms: <wall time>, peakRssMb: <max RSS observed> }
 * Invoked by bench.mjs in a fresh child per case so cases don't share heap state.
 *
 * Usage: node bench/run-case.mjs <case> <rows> [inputFile]
 */
import { readFile } from 'node:fs/promises'
import { Writable } from 'node:stream'

const [, , caseName, rowsArg, inputFile] = process.argv
const ROWS = Number(rowsArg)
const COLS = 10

let peakRss = process.memoryUsage.rss()
const sampler = setInterval(() => {
  const rss = process.memoryUsage.rss()
  if (rss > peakRss) peakRss = rss
}, 10)
sampler.unref()

function* rowValues() {
  for (let r = 0; r < ROWS; r++) {
    const row = new Array(COLS)
    for (let c = 0; c < COLS; c++) {
      row[c] = c % 2 === 0 ? `cell ${r}:${c}` : r * COLS + c
    }
    yield row
  }
}

const devNull = () =>
  new Writable({
    write(_chunk, _enc, cb) {
      cb()
    },
  })

async function drain(readable) {
  const reader = readable.getReader()
  for (;;) {
    const { done } = await reader.read()
    if (done) return
  }
}

const cases = {
  async 'excelents-write'() {
    const { createWorkbook, writeXlsx } = await import('../dist/index.js')
    const wb = createWorkbook()
    const ws = wb.addSheet('Data')
    for (const row of rowValues()) ws.addRow(row)
    await writeXlsx(wb)
  },
  async 'exceljs-write'() {
    const { default: ExcelJS } = await import('exceljs')
    const wb = new ExcelJS.Workbook()
    const ws = wb.addWorksheet('Data')
    for (const row of rowValues()) ws.addRow(row)
    await wb.xlsx.writeBuffer()
  },
  async 'excelents-read'() {
    const { readXlsx } = await import('../dist/index.js')
    const bytes = new Uint8Array(await readFile(inputFile))
    await readXlsx(bytes)
  },
  async 'exceljs-read'() {
    const { default: ExcelJS } = await import('exceljs')
    const wb = new ExcelJS.Workbook()
    await wb.xlsx.load(await readFile(inputFile))
  },
  async 'excelents-stream-write'() {
    const { writeXlsxStream } = await import('../dist/stream.js')
    await drain(writeXlsxStream(rowValues(), { sheet: 'Data' }))
  },
  async 'exceljs-stream-write'() {
    const { default: ExcelJS } = await import('exceljs')
    const wb = new ExcelJS.stream.xlsx.WorkbookWriter({ stream: devNull() })
    const ws = wb.addWorksheet('Data')
    for (const row of rowValues()) ws.addRow(row).commit()
    ws.commit()
    await wb.commit()
  },
  async 'excelents-stream-read'() {
    const { readXlsxRows } = await import('../dist/stream.js')
    const bytes = new Uint8Array(await readFile(inputFile))
    let n = 0
    for await (const row of readXlsxRows(bytes)) n = row.rowNumber
    if (n !== ROWS) throw new Error(`expected ${ROWS} rows, saw ${n}`)
  },
  async 'exceljs-stream-read'() {
    const { default: ExcelJS } = await import('exceljs')
    const reader = new ExcelJS.stream.xlsx.WorkbookReader(inputFile)
    let n = 0
    for await (const worksheet of reader) {
      for await (const row of worksheet) n = row.number
    }
    if (n !== ROWS) throw new Error(`expected ${ROWS} rows, saw ${n}`)
  },
}

const run = cases[caseName]
if (!run) {
  console.error(`unknown case: ${caseName}`)
  process.exit(1)
}

const start = performance.now()
await run()
const ms = performance.now() - start
const rss = process.memoryUsage.rss()
if (rss > peakRss) peakRss = rss
console.log(JSON.stringify({ ms: Math.round(ms), peakRssMb: Math.round(peakRss / 1024 / 1024) }))
