# excelents

> Universal, tree-shakeable, TypeScript-first spreadsheets — `.xlsx` and `.csv`, read, write, and stream. Zero runtime dependencies.

[![CI](https://github.com/mitresthen/excelents/actions/workflows/ci.yml/badge.svg)](https://github.com/mitresthen/excelents/actions/workflows/ci.yml)
[![npm version](https://img.shields.io/npm/v/%40mitresthen%2Fexcelents)](https://www.npmjs.com/package/@mitresthen/excelents)
[![node](https://img.shields.io/node/v/%40mitresthen%2Fexcelents)](#runtime-support)
[![license](https://img.shields.io/npm/l/%40mitresthen%2Fexcelents)](./LICENSE)

`excelents` is a ground-up rewrite of [ExcelJS](https://github.com/exceljs/exceljs) for the modern
JavaScript ecosystem. The ZIP container, XML engine, and CSV codec are all implemented in-house on
web-standard APIs (`CompressionStream`, Web Streams, `TextEncoder`/`TextDecoder`), so the same
code runs unmodified in Node, browsers, Deno, Bun, and edge runtimes — with **no runtime
dependencies at all**.

- **Tiny.** The entire library — reader, writer, streaming, CSV — is ~31 KB gzipped. ExcelJS's
  minified browser bundle is ~252 KB gzipped. Subpath entries and `sideEffects: false` mean a
  write-only app never bundles the reader.
- **Universal.** The core contains zero `node:` imports (enforced by a purity test in CI).
  Node-specific filesystem helpers live behind the separate `./node` entry.
- **TypeScript-first.** Written in TypeScript; every shipped type is generated from the source,
  never hand-maintained. Checked with [publint](https://publint.dev) and
  [Are the Types Wrong](https://arethetypeswrong.github.io) in CI.
- **Streaming.** The `./stream` entry reads and writes xlsx/CSV row-by-row over Web Streams with
  backpressure, so memory stays bounded regardless of file size.
- **Verified against ExcelJS.** Every serialize/parse path is validated by a conformance harness
  that XML-diffs our output against `exceljs@4.4.0` and round-trips 36 real-world fixture files.

## Contents

- [Install](#install)
- [Quick start](#quick-start)
- [Feature tour](#feature-tour)
- [Streaming](#streaming-mitresthenexcelentsstream)
- [Node filesystem helpers](#node-filesystem-helpers-mitresthenexcelentsnode)
- [Entry points](#entry-points)
- [Performance](#performance)
- [Runtime support](#runtime-support)
- [Coming from ExcelJS?](#coming-from-exceljs)
- [FAQ](#faq)
- [Roadmap](#roadmap)
- [Contributing](#contributing)
- [License](#license)

## Install

```sh
npm install @mitresthen/excelents
```

Requirements: the package is **ESM-only** and targets **Node ≥ 24** or any runtime with
`CompressionStream`/`DecompressionStream` support (all evergreen browsers, Deno, Bun, edge
runtimes). See [Runtime support](#runtime-support).

## Quick start

```ts
import { createWorkbook, writeXlsx, readXlsx } from '@mitresthen/excelents'

const wb = createWorkbook()
const ws = wb.addSheet('Sheet1')
ws.cell('A1').value = 'Hello'
ws.cell('B1').value = 42

const bytes: Uint8Array = await writeXlsx(wb) // a complete .xlsx file

const restored = await readXlsx(bytes)
restored.sheets[0]?.cell('A1').value // 'Hello'
```

In the browser, turn the bytes into a download:

```ts
const blob = new Blob([await writeXlsx(wb)], {
  type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
})
// hand `URL.createObjectURL(blob)` to an <a download> ...
```

## Feature tour

### Cell values

A cell's `value` is a plain discriminated union — `string`, `number`, `boolean`, `Date`, `null`,
or one of three object shapes:

```ts
ws.cell('A1').value = new Date('2026-01-01')                    // date (serial-date encoded)
ws.cell('A2').value = { formula: 'SUM(B1:B9)', result: 42 }     // formula + cached result
ws.cell('A3').value = { text: 'docs', hyperlink: 'https://example.com' }
ws.cell('A4').value = {
  richText: [{ text: 'bold', font: { bold: true } }, { text: ' plain' }],
}
```

`readXlsx` restores the same shapes, and `cell.type` tells you which one you got
(`'string' | 'number' | 'boolean' | 'date' | 'formula' | 'richText' | 'hyperlink' | 'null'`).

### Styling

Styles are one plain object per cell — fonts, fills, borders, alignment, and number formats:

```ts
ws.cell('A1').style = {
  font: { name: 'Calibri', size: 12, bold: true, color: 'FFFFFFFF' },
  fill: { type: 'pattern', pattern: 'solid', fgColor: 'FF4472C4' },
  border: { bottom: { style: 'thin', color: 'FF000000' } },
  alignment: { horizontal: 'right', indent: 1, wrapText: true },
  numberFormat: '#,##0.00',
}
```

### Layout: merges, sizing, freezing, filters

```ts
ws.merge('A1:C1')               // merged cells
ws.column(1).width = 24         // column width (column A)
ws.getRow(1).height = 30        // row height
ws.freeze({ rows: 1, cols: 1 }) // frozen panes
ws.setAutoFilter('A1:F1')       // autoFilter over the header range
```

Note: `freeze`, `setAutoFilter`, alignment `indent`, and images are **write-side** features — they
serialize into the file, but `readXlsx` does not currently parse them back (see
[Roadmap](#roadmap)). Merges, sizing, values, and the rest of the styles round-trip fully.

### Images

Embed PNG/JPEG/GIF images (write-side), anchored to a cell and sized in pixels:

```ts
const logo = wb.addImage({ data: pngBytesOrBase64, extension: 'png' }) // -> image id
ws.placeImage(logo, { tl: 'F1', size: { width: 180, height: 101 } })
```

### Tables, data validation & defined names

```ts
ws.addTable({
  name: 'Sales',
  ref: 'A1:C4',
  columns: ['Region', 'Q1', 'Q2'],
  styleName: 'TableStyleMedium2',
})

ws.addDataValidation({
  sqref: 'B2:B100',
  type: 'list',
  formula1: '"Small,Medium,Large"', // in-cell dropdown
})

wb.defineName('TaxRate', 'Sheet1!$B$1')
```

All three survive a `writeXlsx` → `readXlsx` round-trip.

### CSV

RFC 4180 reader and writer in a dedicated ~0.1 KB entry:

```ts
import { writeCsv, readCsv } from '@mitresthen/excelents/csv'

const csv = writeCsv(wb, { bom: true })        // pass a Workbook or a Worksheet
const parsed = readCsv('name,qty\nwidget,42')  // -> Workbook (numbers/booleans inferred)
```

Both take options for delimiters, quoting, BOM, sheet selection, and inference
(`CsvWriteOptions` / `CsvReadOptions`).

## Streaming (`@mitresthen/excelents/stream`)

For large spreadsheets, the `./stream` entry produces and consumes rows incrementally, so memory
stays **bounded regardless of file size**. It's a separate, tree-shakeable entry that does not
pull in the buffered codecs.

### Write

Functional form — `rows` is any (async) iterable of cell-value arrays:

```ts
import { writeXlsxStream } from '@mitresthen/excelents/stream'

async function* rows() {
  yield ['name', 'qty']
  yield ['widget', 42]
}

const stream: ReadableStream<Uint8Array> = writeXlsxStream(rows(), { sheet: 'Data' })
// pipe `stream` to a file or HTTP response
```

Builder form — push rows imperatively; awaiting `addRow` applies backpressure:

```ts
import { createXlsxStreamWriter } from '@mitresthen/excelents/stream'

const writer = createXlsxStreamWriter({ sheet: 'Data' })
// writer.readable: ReadableStream<Uint8Array>
await writer.addRow(['name', 'qty'])
await writer.addRow(['widget', 42])
await writer.close()
```

### Read

`readXlsxRows` async-iterates worksheet rows without building the whole workbook. The source may
be a `Uint8Array`, a `Blob`, or a `ReadableStream<Uint8Array>`:

```ts
import { readXlsxRows } from '@mitresthen/excelents/stream'

for await (const { sheet, rowNumber, cells } of readXlsxRows(bytes)) {
  // cells: CellValue[] indexed by column (index 0 = column A)
  console.log(sheet, rowNumber, cells)
}
```

### Streaming CSV

```ts
import { writeCsvStream, readCsvRows } from '@mitresthen/excelents/stream'

const csvStream = writeCsvStream(rows()) // ReadableStream<Uint8Array>

for await (const row of readCsvRows(csvStream)) {
  // row: CellValue[]  (readCsvRows also accepts a plain string)
}
```

## Node filesystem helpers (`@mitresthen/excelents/node`)

The core never touches the filesystem. In Node, use the adapter to read/write files and to bridge
Node streams to Web Streams:

```ts
import { createWorkbook, writeXlsx, readXlsx } from '@mitresthen/excelents'
import { writeXlsxStream } from '@mitresthen/excelents/stream'
import { nodeFileSystem } from '@mitresthen/excelents/node'

// buffered
await nodeFileSystem.writeFile('report.xlsx', await writeXlsx(wb))
const wb2 = await readXlsx(await nodeFileSystem.readFile('report.xlsx'))

// streaming, bounded memory
await writeXlsxStream(rows()).pipeTo(nodeFileSystem.createWritable('big.xlsx'))
```

## Entry points

| Import | Contents | Size (gzip) |
| --- | --- | --- |
| `@mitresthen/excelents` | `createWorkbook`, `writeXlsx`, `readXlsx`, the model classes, all public types, `version` | ~8 KB + shared chunks |
| `@mitresthen/excelents/csv` | `writeCsv`, `readCsv` | ~0.1 KB + shared chunks |
| `@mitresthen/excelents/stream` | `writeXlsxStream`, `createXlsxStreamWriter`, `readXlsxRows`, `writeCsvStream`, `readCsvRows` | ~7 KB |
| `@mitresthen/excelents/node` | `nodeFileSystem` (Node FS ↔ Web Streams adapter) | ~0.3 KB |

Everything is named exports with `sideEffects: false`, so bundlers drop whatever you don't use —
per-entry gzip budgets are asserted in CI on every commit.

## Performance

100,000 rows × 10 columns (5 strings + 5 numbers per row), Apple M4 Max, Node 24. Each case runs
in a fresh process; peak RSS sampled in-process. Reproduce with `pnpm build && pnpm bench`.

| Scenario | excelents | exceljs | excelents peak RSS | exceljs peak RSS |
| --- | --- | --- | --- | --- |
| buffered write | 7.3 s | 3.3 s | 991 MB | 1457 MB |
| buffered read | 0.8 s | 1.5 s | 605 MB | 793 MB |
| streaming write | 1.6 s | 0.7 s | 124 MB | 161 MB |
| streaming read | 0.8 s | 1.4 s | 225 MB | 358 MB |

Honest summary: reads are ~2× faster and every path uses meaningfully less memory, but ExcelJS
currently wins on raw write throughput — write-path optimization is on the [roadmap](#roadmap).
For large exports the streaming writer's bounded memory (124 MB vs ~1 GB buffered) matters more
than the seconds.

## Runtime support

| Runtime | Status |
| --- | --- |
| Node.js ≥ 24 | ✅ Tested in CI |
| Evergreen browsers | ✅ Tested in CI (Vitest browser mode, real Chromium) |
| Deno / Bun / edge (Workers, etc.) | ✅ Expected to work — the core uses only web-standard APIs |

The only platform capabilities required are `CompressionStream`/`DecompressionStream` with
`deflate-raw`, Web Streams, and `TextEncoder`/`TextDecoder`.

The package ships **ESM only** (since v0.2.0). If you're on CommonJS, use dynamic `import()`.

## Coming from ExcelJS?

`excelents` began as a fork of `exceljs@4.4.0` (unmaintained upstream since 2023) and was
rewritten from scratch: new architecture, new public API, no CommonJS, no dependencies. It is
**not a drop-in replacement**, but the model maps closely and correctness is continuously verified
against ExcelJS itself.

### What's supported

Cell values (strings, numbers, booleans, dates, formulas with cached results, rich text,
hyperlinks) · styling (fonts, fills, borders, alignment, number formats) · merged cells · row
heights & column widths · tables · data validation · defined names · frozen panes, autoFilter,
alignment indent & embedded images (write-side) · 1900/1904 date systems · CSV · streaming
read/write.

### What was dropped from the fork

These ExcelJS features were **deliberately cut** and are not planned:

| Dropped | Notes |
| --- | --- |
| Cell comments / notes | Rarely round-trips cleanly; cut to keep the model small |
| Pivot tables | ExcelJS support was partial/experimental to begin with |
| Encryption / password protection | Out of scope |
| Page setup, print options, headers/footers | Not carried over — open an issue if you need this |
| Worksheet/workbook protection | Not carried over |
| CommonJS build | ESM-only since v0.2.0 |

Still on the roadmap (in scope, not yet shipped): **conditional formatting**, and parsing the
write-side features (images, frozen panes, autoFilter, alignment indent) back out of workbooks
on read.

### API mapping

| ExcelJS | excelents |
| --- | --- |
| `new Excel.Workbook()` | `createWorkbook()` |
| `wb.addWorksheet('Data')` | `wb.addSheet('Data')` |
| `ws.getCell('A1').value = 1` | `ws.cell('A1').value = 1` |
| `ws.getCell('A1').font = {...}` | `ws.cell('A1').style = { font: {...} }` |
| `cell.numFmt = '0.00%'` | `cell.style = { numberFormat: '0.00%' }` |
| `ws.mergeCells('A1:B2')` | `ws.merge('A1:B2')` |
| `ws.getColumn(1).width = 20` | `ws.column(1).width = 20` |
| `ws.autoFilter = 'A1:C1'` | `ws.setAutoFilter('A1:C1')` |
| `ws.views = [{ state: 'frozen', ySplit: 1 }]` | `ws.freeze({ rows: 1 })` |
| `wb.xlsx.writeBuffer()` | `writeXlsx(wb)` |
| `wb.xlsx.load(buffer)` | `readXlsx(bytes)` |
| `wb.xlsx.writeFile(path)` | `nodeFileSystem.writeFile(path, await writeXlsx(wb))` |
| `wb.csv.writeBuffer()` | `writeCsv(wb)` |
| `new Excel.stream.xlsx.WorkbookWriter(...)` | `createXlsxStreamWriter(...)` / `writeXlsxStream(...)` |

The general pattern: workbooks are built through a small model API, and serialization is **free
functions** (`writeXlsx`, `readXlsx`, `writeCsv`, …) instead of methods on the workbook — that's
what makes the unused codecs tree-shake away.

## FAQ

**Does it evaluate formulas?**
No — like ExcelJS, it stores formula text plus an optional cached `result`. Excel recalculates on
open.

**Why Node ≥ 24?**
That's what CI tests against. The core is pure web-standard code, so earlier Node versions with
`deflate-raw` support will likely work, but they're not verified.

**How do I know the output is correct?**
Every codec path runs through a conformance harness in CI: workbooks are built with both
`excelents` and `exceljs`, unzipped, canonicalized, and XML-diffed part by part; the read path
parses 36 real-world fixture workbooks with both libraries and compares the resulting models.

**Buffered or streaming — which do I pick?**
`writeXlsx`/`readXlsx` build the whole workbook in memory and support every feature (styles,
images, tables, …). The `./stream` entry handles arbitrarily large files with bounded memory, but
works row-by-row with plain values. Reports → buffered; exports/imports of millions of rows →
streaming.

**Is the API stable?**
It's pre-1.0: the API is small and settling, but minor versions may still contain breaking
changes. Pin a version and read release notes when bumping.

## Roadmap

- Conditional formatting (write + read)
- Write-path performance (currently slower than ExcelJS — see [Performance](#performance))
- Reading embedded images back out of workbooks
- An ExcelJS migration guide covering every mapping
- v1.0: API freeze + packaging polish

## Contributing

```sh
git clone https://github.com/mitresthen/excelents
cd excelents
pnpm install
pnpm test        # vitest (node + browser projects), incl. the conformance suite
```

| Script | What it does |
| --- | --- |
| `pnpm build` | Build `dist/` with tsdown (Rolldown) |
| `pnpm typecheck` | `tsc --noEmit` |
| `pnpm lint` / `pnpm lint:fix` | oxlint (type-aware) |
| `pnpm format` / `pnpm format:fix` | oxfmt |
| `pnpm test` | Vitest (node project), including the ExcelJS conformance oracle |
| `pnpm test:browser` | The public-API suite in real Chromium (needs `pnpm exec playwright install chromium` once) |
| `pnpm size` | Assert per-entry gzip budgets (`size-budget.json`) |
| `pnpm smoke` | Pack the tarball and import it from a scratch project |

Design documents for each subsystem live in [`docs/superpowers`](./docs/superpowers). Releases are
automated via npm trusted publishing — see [RELEASING.md](./RELEASING.md).

## License

[MIT](./LICENSE)

`excelents` descends from [ExcelJS](https://github.com/exceljs/exceljs) by Guyon Roche and
contributors — the fixture suite and the conformance oracle stand on their work. 🙏
