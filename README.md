# excelents

Universal, tree-shakeable, TypeScript-first spreadsheet (xlsx/csv) library. Zero runtime
dependencies; built on web-standard APIs (`CompressionStream`, `ReadableStream`, `TextDecoder`),
so it runs the same in Node, browsers, and edge runtimes.

## Install

```sh
npm install excelents
```

## Quick start

```ts
import { createWorkbook, writeXlsx, readXlsx } from 'excelents'

const wb = createWorkbook()
const ws = wb.addSheet('Sheet1')
ws.cell('A1').value = 'Hello'
ws.cell('B1').value = 42

const bytes: Uint8Array = await writeXlsx(wb) // a .xlsx file

const restored = await readXlsx(bytes)
restored.sheets[0]?.cell('A1').value // 'Hello'
```

### CSV

```ts
import { writeCsv, readCsv } from 'excelents/csv'

const csv = writeCsv(wb) // RFC 4180 string (pass a Workbook or a Worksheet)
const parsed = readCsv('name,qty\nwidget,42') // -> Workbook
```

## Streaming (`excelents/stream`)

For large spreadsheets, the `./stream` entry produces and consumes rows incrementally, so memory
stays **bounded regardless of file size**. It's a separate, tree-shakeable entry that does not pull
in the buffered codecs.

### Write

Functional form — `rows` is any (async) iterable of cell-value arrays:

```ts
import { writeXlsxStream } from 'excelents/stream'

async function* rows() {
  yield ['name', 'qty']
  yield ['widget', 42]
}

const stream: ReadableStream<Uint8Array> = writeXlsxStream(rows(), { sheet: 'Data' })
// pipe `stream` to a file or HTTP response
```

Builder form — push rows imperatively; awaiting `addRow` applies backpressure:

```ts
import { createXlsxStreamWriter } from 'excelents/stream'

const writer = createXlsxStreamWriter({ sheet: 'Data' })
// writer.readable: ReadableStream<Uint8Array>
await writer.addRow(['name', 'qty'])
await writer.addRow(['widget', 42])
await writer.close()
```

### Read

`readXlsxRows` async-iterates worksheet rows without building the whole workbook. The source may be
a `Uint8Array`, a `Blob`, or a `ReadableStream<Uint8Array>`:

```ts
import { readXlsxRows } from 'excelents/stream'

for await (const { sheet, rowNumber, cells } of readXlsxRows(bytes)) {
  // cells: CellValue[] indexed by column (index 0 = column A)
  console.log(sheet, rowNumber, cells)
}
```

### Streaming CSV

```ts
import { writeCsvStream, readCsvRows } from 'excelents/stream'

const csvStream = writeCsvStream(rows()) // ReadableStream<Uint8Array>

for await (const row of readCsvRows(csvStream)) {
  // row: CellValue[]  (readCsvRows also accepts a plain string)
}
```

## Entry points

| Import | Contents |
| --- | --- |
| `excelents` | `createWorkbook`, `writeXlsx`, `readXlsx`, the model classes, and `version` |
| `excelents/csv` | `writeCsv`, `readCsv` |
| `excelents/stream` | `writeXlsxStream`, `createXlsxStreamWriter`, `readXlsxRows`, `writeCsvStream`, `readCsvRows` |
| `excelents/node` | Node filesystem helpers |

## License

MIT
