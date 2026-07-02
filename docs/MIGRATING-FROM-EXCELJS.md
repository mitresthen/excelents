# Migrating from ExcelJS

`excelents` is a ground-up rewrite, not a fork that kept the API. This guide maps every ExcelJS
concept to its excelents equivalent, lists what has no equivalent, and calls out behavioral
differences that can bite silently.

## Sizing your migration

Before reading further, let the scanner inventory your codebase:

```sh
npx @mitresthen/excelents .          # human-readable report
npx @mitresthen/excelents . --json   # machine-readable, for tooling/agents
```

It scans for ExcelJS usage and buckets every call site: **auto-mappable** (mechanical renames,
listed with their new name), **needs restructuring** (same capability, different shape — each
points at a section of this guide), and **no equivalent** (dropped features — decide on these
first). It never modifies files. The scanner is also importable — `scanSource` from
`@mitresthen/excelents/migrate` — if you're building tooling on top.

**Migrating with a coding agent?** Point it at this guide plus the `--json` scanner output; the
combination is designed to give an agent everything it needs — the inventory of sites and the
exact mapping for each.

## The three big conceptual changes

**1. Serialization is free functions, not workbook methods.**
ExcelJS hangs codecs off the workbook (`wb.xlsx.writeBuffer()`, `wb.csv.readFile()`); excelents
exposes them as importable functions (`writeXlsx(wb)`, `readCsv(text)`) so bundlers can
tree-shake codecs you don't use. The workbook is a plain data model.

**2. One `style` object per cell, plain values.**
ExcelJS spreads styling across cell properties (`cell.font`, `cell.fill`, `cell.numFmt`, …) and
wraps colors in objects (`{ argb: 'FF0000FF' }`). excelents has a single `cell.style` object with
plain ARGB strings.

**3. ESM-only, zero dependencies, universal.**
There is no CommonJS build (use dynamic `import()` from CJS), no polyfills, and no Node-specific
code in the core — filesystem access moved to the explicit `@mitresthen/excelents/node` entry.
Requires Node ≥ 24 in Node environments.

## Package & imports

```ts
// ExcelJS
const ExcelJS = require('exceljs')
const wb = new ExcelJS.Workbook()

// excelents
import { createWorkbook, writeXlsx, readXlsx } from '@mitresthen/excelents'
import { writeCsv, readCsv } from '@mitresthen/excelents/csv'
import { writeXlsxStream, readXlsxRows } from '@mitresthen/excelents/stream'
import { nodeFileSystem } from '@mitresthen/excelents/node'
const wb = createWorkbook()
```

## Workbook & worksheets

| ExcelJS | excelents |
| --- | --- |
| `new ExcelJS.Workbook()` | `createWorkbook()` |
| `wb.addWorksheet('Data')` | `wb.addSheet('Data')` |
| `wb.getWorksheet('Data')` | `wb.getSheet('Data')` |
| `wb.removeWorksheet(id)` | `wb.removeSheet(name)` |
| `wb.worksheets` / `wb.eachSheet(cb)` | `wb.sheets` (a plain readonly array — iterate it) |
| `wb.creator`, `wb.created`, … | Not modeled (document properties are not written) |

## Reading & writing files

| ExcelJS | excelents |
| --- | --- |
| `await wb.xlsx.writeBuffer()` | `await writeXlsx(wb)` → `Uint8Array` |
| `await wb.xlsx.load(buffer)` | `await readXlsx(bytes)` → new `Workbook` |
| `await wb.xlsx.writeFile(path)` | `await nodeFileSystem.writeFile(path, await writeXlsx(wb))` |
| `await wb.xlsx.readFile(path)` | `await readXlsx(await nodeFileSystem.readFile(path))` |
| `wb.xlsx.write(stream)` | `writeXlsxStream(...)` piped to a `WritableStream` (see [Streaming](#streaming)) |

Note `readXlsx` **returns** a workbook rather than mutating an existing one.

## Cells & values

Addressing is 1-based in both libraries.

| ExcelJS | excelents |
| --- | --- |
| `ws.getCell('A1')` | `ws.cell('A1')` |
| `ws.getCell(row, col)` | `ws.getCell(row, col)` |
| `ws.getRow(1)` / `ws.getColumn(1)` | `ws.getRow(1)` / `ws.column(1)` |
| `ws.addRow([...])` | `ws.addRow([...])` |
| `ws.addRow({ id: 1, name: 'x' })` (object by column key) | Not supported — pass arrays |
| `ws.spliceRows`, `ws.insertRow`, `row.splice` | Not supported — build rows in the right order |
| `cell.value = 7` / `'text'` / `new Date()` / `true` / `null` | Same |
| `cell.value = { formula: 'A1+A2', result: 7 }` | Same shape |
| `cell.value = { richText: [{ text: 'a', font: {...} }] }` | Same shape |
| `cell.value = { text: 'link', hyperlink: 'https://…' }` | Same shape |
| `cell.value = { error: '#N/A' }` | Not modeled — error cells read back as empty |
| `cell.value = { sharedFormula: 'A1' }` | Not modeled — write full formulas per cell |
| `cell.type` (numeric `ValueType` enum) | `cell.type` (string union: `'string' \| 'number' \| …`) |
| `cell.note = '...'` (comments) | Dropped |

## Styles

Everything lives on `cell.style`; colors are plain 8-digit ARGB strings.

```ts
// ExcelJS
cell.font = { name: 'Calibri', bold: true, color: { argb: 'FFFF0000' } }
cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4472C4' } }
cell.border = { bottom: { style: 'thin', color: { argb: 'FF000000' } } }
cell.alignment = { horizontal: 'right', wrapText: true }
cell.numFmt = '#,##0.00'

// excelents
cell.style = {
  font: { name: 'Calibri', bold: true, color: 'FFFF0000' },
  fill: { type: 'pattern', pattern: 'solid', fgColor: 'FF4472C4' },
  border: { bottom: { style: 'thin', color: 'FF000000' } },
  alignment: { horizontal: 'right', wrapText: true },
  numberFormat: '#,##0.00',
}
```

The style model is a deliberate subset of what ExcelJS accepted:

- **Font:** `name`, `size`, `bold`, `italic`, `underline` (boolean), `color`. No `strike`,
  `vertAlign`, `outline`, underline styles, or theme fonts.
- **Fill:** pattern fills only, `solid` or `none`, with `fgColor`. No gradients, no `bgColor`.
- **Border edges:** `top`/`bottom`/`left`/`right` with styles
  `thin | medium | thick | dashed | dotted | double`. No `hair`, `mediumDashed`, diagonals, etc.
- **Alignment:** `horizontal: left | center | right`, `vertical: top | middle | bottom`,
  `wrapText`, `indent`. No `justify`, `fill`, `distributed`, `textRotation`, `shrinkToFit`.
- **Colors:** ARGB strings only. No `{ theme, tint }` or indexed colors — resolve them to ARGB.
- **Number formats:** any format string, plus the OOXML built-in table on read.

Row/column styles (`row.font = …`, `column.style = …`) are not modeled — style cells directly.

## Rows, columns & layout

| ExcelJS | excelents |
| --- | --- |
| `ws.getColumn(1).width = 20` | `ws.column(1).width = 20` |
| `ws.getRow(1).height = 30` | `ws.getRow(1).height = 30` |
| `ws.columns = [{ header, key, width }]` | Not supported — no column keys; set widths directly |
| `row.hidden = true`, `column.hidden = true` | Not modeled |
| `ws.mergeCells('A1:B2')` | `ws.merge('A1:B2')` |
| `ws.unMergeCells(...)` | Not supported |
| `ws.views = [{ state: 'frozen', xSplit: 1, ySplit: 2 }]` | `ws.freeze({ cols: 1, rows: 2 })` |
| `ws.autoFilter = 'A1:C1'` | `ws.setAutoFilter('A1:C1')` |

`freeze`, `setAutoFilter`, alignment `indent`, and images are currently **write-side only**: they
serialize correctly, but `readXlsx` does not parse them back yet.

## Images

```ts
// ExcelJS
const id = wb.addImage({ base64: png, extension: 'png' })
ws.addImage(id, { tl: { col: 5, row: 0 }, ext: { width: 180, height: 101 } })

// excelents — anchor is a cell ref string, size in pixels
const id = wb.addImage({ data: png, extension: 'png' }) // Uint8Array or base64 string
ws.placeImage(id, { tl: 'F1', size: { width: 180, height: 101 } })
```

Supported extensions: `png`, `jpeg`, `gif`. Two-cell (`tl`/`br` range) anchors and
`editAs` are not modeled.

## Tables, data validation, defined names

```ts
// ExcelJS tables carry their own row data:
ws.addTable({ name: 'T', ref: 'A1', headerRow: true,
  columns: [{ name: 'Region' }, { name: 'Qty' }], rows: [['North', 1]] })

// excelents tables describe a range — you write the cells yourself:
ws.addRow(['Region', 'Qty'])
ws.addRow(['North', 1])
ws.addTable({ name: 'T', ref: 'A1:B2', columns: ['Region', 'Qty'] })
```

```ts
// ExcelJS attaches validation per cell:
ws.getCell('B2').dataValidation = { type: 'list', formulae: ['"a,b,c"'] }

// excelents attaches it to a range (sqref), like the file format does:
ws.addDataValidation({ sqref: 'B2:B100', type: 'list', formula1: '"a,b,c"' })
```

```ts
// ExcelJS
wb.definedNames.add('Sheet1!$A$1', 'TaxRate')

// excelents
wb.defineName('TaxRate', 'Sheet1!$A$1')
```

Total rows, calculated table columns, and table themes beyond `styleName` are not modeled.

## CSV

excelents CSV works on strings, not files or Node streams — pair it with `nodeFileSystem` (or any
IO) yourself:

| ExcelJS | excelents |
| --- | --- |
| `await wb.csv.writeBuffer()` | `writeCsv(wb)` → `string` |
| `await wb.csv.writeFile(path)` | `nodeFileSystem.writeFile(path, new TextEncoder().encode(writeCsv(wb)))` |
| `await wb.csv.readFile(path)` (uses fast-csv options) | `readCsv(text, { delimiter, parseNumbers, parseBooleans })` |
| streaming CSV via fast-csv | `writeCsvStream(rows)` / `readCsvRows(source)` from `./stream` |

`writeCsv` emits RFC 4180 output with no trailing newline; pass `{ bom: true }` for an
Excel-friendly UTF-8 BOM.

## Streaming

| ExcelJS | excelents |
| --- | --- |
| `new ExcelJS.stream.xlsx.WorkbookWriter({ filename })` | `writeXlsxStream(rows).pipeTo(nodeFileSystem.createWritable(path))` |
| `new ExcelJS.stream.xlsx.WorkbookWriter({ stream })` | `writeXlsxStream(rows)` / `createXlsxStreamWriter()` → Web `ReadableStream` |
| `row.commit()` / `ws.commit()` / `wb.commit()` | `await writer.addRow([...])` … `await writer.close()` (backpressure via `await`) |
| `new ExcelJS.stream.xlsx.WorkbookReader(path)` | `readXlsxRows(bytes \| Blob \| ReadableStream)` |
| `{ useStyles: true }` on the stream writer | Not supported — the streaming writer emits plain values on a single sheet |
| multiple sheets per streamed workbook | Not supported yet — one sheet per stream |

## Dropped entirely (no equivalent)

| Feature | Suggestion |
| --- | --- |
| Cell comments / notes | none — deliberately out of scope |
| Pivot tables | none — was experimental in ExcelJS too |
| Encrypted / password-protected workbooks | none — out of scope |
| Page setup, print areas, headers/footers | keep using ExcelJS if this is load-bearing, or open an issue |
| Worksheet/workbook protection | same as above |
| Conditional formatting | on the excelents roadmap |
| Document properties (creator, modified, …) | not written; most consumers never notice |
| Sheet/tab colors, hidden sheets, outlines | not modeled |

## Behavioral gotchas

- **`readXlsx` never throws on unknown parts** — it simply ignores features it doesn't model. A
  read-modify-write cycle **does not preserve** unmodeled content (comments, pivot caches, …).
  ExcelJS partially preserved some of these.
- **Formula cells:** neither library evaluates formulas. excelents writes the cached `result` if
  you provide one; omit it and Excel shows the value after recalculation.
- **Error cells** (`#N/A`, `#DIV/0!`) come back as empty cells, not error values.
- **Dates** are honored under both date systems on read (`date1904` workbooks included); writes
  always emit the standard 1900 system. Round-tripping a 1904 workbook keeps every date the same
  instant while normalizing the file to 1900.
- **Number inference in CSV:** `readCsv` only converts strings that round-trip losslessly
  (`parseNumbers: false` to disable), so ID-like strings such as `007` stay strings.
- **Streaming rows are arrays indexed from column A** (`cells[0]` = A), while the model API is
  1-based (`getCell(1, 1)` = A1).
