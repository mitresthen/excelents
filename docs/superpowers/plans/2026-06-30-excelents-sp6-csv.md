# SP-6: CSV (read + write) Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** A tree-shakeable `./csv` entry exposing `writeCsv` and `readCsv` — RFC 4180-compliant CSV serialize/parse over the existing model, pulling only the model (never the xlsx codecs).

**Architecture:** `src/csv.ts` re-exports a writer (`src/csv/writer.ts`) and reader (`src/csv/reader.ts`) plus option types (`src/csv/options.ts`). The writer renders a worksheet's sparse grid to a rectangular CSV string with RFC-4180 quoting; the reader is a single-pass character state machine that yields rows of string fields, then applies *conservative* type inference into a one-sheet `Workbook`. Both are `node:`-free and depend only on `src/model/*`.

**Tech Stack:** TypeScript 7 (isolatedDeclarations), the SP-2 model, Vitest 4 + `exceljs@4.4.0` oracle. Zero runtime deps.

## Global Constraints

- **Zero runtime dependencies.** No new `package.json#dependencies`.
- **`node:`-free core.** No `node:` imports anywhere under `src/csv/` or `src/csv.ts` (enforced by `test/core-purity.test.ts`). Tests may use `node:` (the oracle needs `node:stream`).
- **`isolatedDeclarations: true`.** Explicit return types on every exported function.
- **TDD, frequent commits.** Red → green. Run gates chained with `&&` before each commit so a failing gate blocks it.
- **Tree-shaking is a hard requirement.** `dist/csv.js` must NOT contain any reference to the xlsx writer/reader or OPC/zip code. Verified in Task 3.
- **Validation:** primary is `readCsv(writeCsv(wb))` round-trip; supplement with exceljs cross-checks (our output parses in exceljs; exceljs's output parses in us) and the real fixture `test/fixtures/test-issue-991.csv`.

## Oracle behavior (exceljs 4.4.0, measured)

- **Write:** a field is quoted iff it contains the delimiter, a `"`, CR, or LF; embedded `"` is doubled; numbers/booleans render bare (`42`, `3.14`, `true`); row delimiter is `\n`; **no** trailing newline.
- **Read:** numeric strings → numbers, `true`/`false` → booleans, embedded newlines survive inside quotes, everything else stays a string.
- **Trap to avoid (the fixture's point):** `test-issue-991.csv` holds `2019-11-04`, `11-04-2019`, `2019-11-04T10:17:55`, `00210PRG1`, `1234-5thisisnotadate`. None must become a Date, and `00210PRG1` / leading-zero strings must NOT become numbers. Hence: **dates off by default**, and number inference must be **round-trip-safe** (`String(Number(s)) === s`), so `00210` stays a string.

---

## File Structure

- `src/csv/options.ts` — **create**: `CsvWriteOptions`, `CsvReadOptions`.
- `src/csv/writer.ts` — **create**: `writeCsv` + private quoting/rendering.
- `src/csv/reader.ts` — **create**: `readCsv` + private parser + inference.
- `src/csv.ts` — **modify** (replace placeholder): re-export the public surface.
- Tests: `src/csv/writer.test.ts`, `src/csv/reader.test.ts`, `src/csv/csv.test.ts`.

**Shared option types:**

```typescript
// src/csv/options.ts
export interface CsvWriteOptions {
  /** Field delimiter (default ','). */
  readonly delimiter?: string
  /** Row delimiter (default '\n'). */
  readonly rowDelimiter?: string
  /** Quote every field, not just those that require it (default false). */
  readonly quoteAll?: boolean
  /** Prepend a UTF-8 BOM (U+FEFF) for Excel compatibility (default false). */
  readonly bom?: boolean
  /** When given a Workbook, which sheet to emit (name or 0-based index; default first). */
  readonly sheet?: string | number
}

export interface CsvReadOptions {
  /** Field delimiter (default ','). */
  readonly delimiter?: string
  /** Name for the single produced worksheet (default 'Sheet1'). */
  readonly sheetName?: string
  /** Infer round-trip-safe numbers (default true). */
  readonly parseNumbers?: boolean
  /** Infer 'true'/'false' as booleans (default true). */
  readonly parseBooleans?: boolean
}
```

---

### Task 1: CSV writer

**Files:**
- Create: `src/csv/options.ts`, `src/csv/writer.ts`
- Test: `src/csv/writer.test.ts`

**Interfaces:**
- Consumes: `Workbook`, `Worksheet` from `../model/*`; `CsvWriteOptions`.
- Produces: `writeCsv(source: Workbook | Worksheet, options?: CsvWriteOptions): string`.

Render a worksheet's grid from `(1,1)` to `(maxRow, maxCol)` where `maxRow`/`maxCol` come from `ws.dimensions` (undefined dimensions → empty string). Each cell value renders by type; a `null`/absent cell → empty field. RFC-4180 quote when the rendered field contains the delimiter, `"`, CR, or LF (or always when `quoteAll`), doubling embedded `"`.

Value rendering (`CellValue` union): `string`→as-is; `number`→`String(n)` when finite else `''`; `boolean`→`'true'`/`'false'`; `Date`→`toISOString()`; `null`→`''`; `FormulaValue`→render its `result` as a primitive (string/number/boolean/Date) or `''`; `RichTextValue`→concatenated `run.text`; `HyperlinkValue`→its `text`.

- [ ] **Step 1: Write the failing test**

```typescript
// src/csv/writer.test.ts
import { expect, test } from 'vitest'
import { createWorkbook } from '../model/workbook'
import { writeCsv } from './writer'

test('renders a simple grid with bare numbers and booleans', () => {
  const wb = createWorkbook()
  const ws = wb.addSheet('S')
  ws.cell('A1').value = 'name'
  ws.cell('B1').value = 42
  ws.cell('C1').value = true
  ws.cell('A2').value = 'x'
  expect(writeCsv(wb)).toBe('name,42,true\nx,,')
})

test('quotes fields containing the delimiter, quotes, or newlines (RFC 4180)', () => {
  const wb = createWorkbook()
  const ws = wb.addSheet('S')
  ws.cell('A1').value = 'has,comma'
  ws.cell('B1').value = 'has"quote'
  ws.cell('C1').value = 'has\nnl'
  expect(writeCsv(ws)).toBe('"has,comma","has""quote","has\nnl"')
})

test('honors delimiter, quoteAll, rowDelimiter, and bom options', () => {
  const wb = createWorkbook()
  const ws = wb.addSheet('S')
  ws.cell('A1').value = 'a'
  ws.cell('B1').value = 'b'
  ws.cell('A2').value = 'c'
  ws.cell('B2').value = 'd'
  expect(writeCsv(ws, { delimiter: ';', quoteAll: true, rowDelimiter: '\r\n', bom: true })).toBe(
    '﻿"a";"b"\r\n"c";"d"',
  )
})
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx vitest run src/csv/writer.test.ts`
Expected: FAIL — `writeCsv` not found.

- [ ] **Step 3: Implement the writer**

```typescript
// src/csv/options.ts  (full content under "Shared option types" above)
```
```typescript
// src/csv/writer.ts
import type { Cell, CellValue } from '../model/cell'
import { Workbook } from '../model/workbook'
import type { Worksheet } from '../model/worksheet'
import type { CsvWriteOptions } from './options'

function resolveSheet(source: Workbook | Worksheet, sheet: string | number | undefined): Worksheet | undefined {
  if (!(source instanceof Workbook)) return source
  if (typeof sheet === 'string') return source.getSheet(sheet)
  return source.sheets[typeof sheet === 'number' ? sheet : 0]
}

function renderValue(v: CellValue): string {
  if (v === null) return ''
  if (typeof v === 'string') return v
  if (typeof v === 'number') return Number.isFinite(v) ? String(v) : ''
  if (typeof v === 'boolean') return v ? 'true' : 'false'
  if (v instanceof Date) return v.toISOString()
  if ('richText' in v) return v.richText.map((r) => r.text).join('')
  if ('hyperlink' in v) return v.text
  if ('formula' in v) {
    const r = v.result
    return r === undefined ? '' : renderValue(r)
  }
  return ''
}

function quoteField(field: string, delimiter: string, quoteAll: boolean): string {
  const mustQuote =
    quoteAll ||
    field.includes(delimiter) ||
    field.includes('"') ||
    field.includes('\n') ||
    field.includes('\r')
  return mustQuote ? `"${field.replaceAll('"', '""')}"` : field
}

export function writeCsv(source: Workbook | Worksheet, options: CsvWriteOptions = {}): string {
  const { delimiter = ',', rowDelimiter = '\n', quoteAll = false, bom = false } = options
  const ws = resolveSheet(source, options.sheet)
  const dims = ws?.dimensions
  if (ws === undefined || dims === undefined) return bom ? '﻿' : ''

  const lines: string[] = []
  for (let r = 1; r <= dims.bottom; r++) {
    const fields: string[] = []
    for (let c = 1; c <= dims.right; c++) {
      const cell: Cell = ws.getCell(r, c)
      fields.push(quoteField(renderValue(cell.value), delimiter, quoteAll))
    }
    lines.push(fields.join(delimiter))
  }
  return (bom ? '﻿' : '') + lines.join(rowDelimiter)
}
```

Note on `renderValue`: confirm the `CellValue` union shape against `src/model/cell.ts` before writing — use the exact discriminants (`'richText'`, `'hyperlink'`, `'formula'`, `result`). If `getCell` auto-creates empty cells (it does), that's fine — an absent cell renders `''`. The `dims.right`/`dims.bottom` come from `RangeBox` (`{ top, left, bottom, right }`).

- [ ] **Step 4: Run test to verify it passes**

Run: `npx vitest run src/csv/writer.test.ts`
Expected: PASS (3 tests).

- [ ] **Step 5: Run gates and commit**

Run: `npx vitest run && npm run -s typecheck && npm run -s lint && npm run -s format`

```bash
git add src/csv/options.ts src/csv/writer.ts src/csv/writer.test.ts
git commit -m "feat(csv): RFC 4180 writer (writeCsv) over the model"
```

---

### Task 2: CSV reader

**Files:**
- Create: `src/csv/reader.ts`
- Test: `src/csv/reader.test.ts`

**Interfaces:**
- Consumes: `createWorkbook`/`Workbook`; `CsvReadOptions`; the real fixture.
- Produces: `readCsv(text: string, options?: CsvReadOptions): Workbook`.

A single-pass character state machine producing `string[][]` (rows × fields), then inference into one sheet. Parser rules (RFC 4180): a field is quoted iff it starts with `"`; inside a quoted field, `""` is a literal `"` and the field ends at the closing `"`; unquoted fields end at the delimiter or row end; rows end at `\n` (a preceding `\r` is stripped — handle `\r\n` and bare `\n`); a trailing newline does not create a spurious empty row, but a genuinely empty line between rows does. A leading BOM (`﻿`) is stripped.

Inference per field: empty string → leave the cell unset (no value); else if `parseNumbers` and `String(Number(s)) === s` and `s.trim() !== ''` → `Number(s)`; else if `parseBooleans` and `s === 'true'`/`'false'` → boolean; else → the string verbatim. (Round-trip-safe number test keeps `00210`, `1e3`-vs-`1000`, and `1234-5...` as strings.)

- [ ] **Step 1: Write the failing test**

```typescript
// src/csv/reader.test.ts
import { expect, test } from 'vitest'
import { listFixtures, parseFixture } from '../../test/conformance/harness'
import { readCsv } from './reader'

test('parses a grid with quoted fields, doubled quotes, and embedded newlines', () => {
  const wb = readCsv('name,42,true\n"has,comma","has""quote","line\nbreak"')
  const ws = wb.sheets[0]!
  expect(ws.cell('A1').value).toBe('name')
  expect(ws.cell('B1').value).toBe(42)
  expect(ws.cell('C1').value).toBe(true)
  expect(ws.cell('A2').value).toBe('has,comma')
  expect(ws.cell('B2').value).toBe('has"quote')
  expect(ws.cell('C2').value).toBe('line\nbreak')
})

test('keeps leading-zero and date-like strings as strings (no over-eager inference)', async () => {
  // test-issue-991.csv: none of these may become a Date or number.
  const path = listFixtures('.csv').find((p) => p.endsWith('test-issue-991.csv'))!
  const text = new TextDecoder().decode(await parseFixture(path))
  const ws = readCsv(text).sheets[0]!
  expect(ws.cell('A1').value).toBe('2019-11-04')
  expect(ws.cell('A2').value).toBe('11-04-2019')
  expect(ws.cell('A3').value).toBe('2019-11-04T10:17:55')
  expect(ws.cell('A4').value).toBe('00210PRG1')
  expect(ws.cell('A5').value).toBe('1234-5thisisnotadate')
})

test('strips a leading BOM and handles CRLF rows', () => {
  const ws = readCsv('﻿a,b\r\nc,d').sheets[0]!
  expect(ws.cell('A1').value).toBe('a')
  expect(ws.cell('A2').value).toBe('c')
})
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx vitest run src/csv/reader.test.ts`
Expected: FAIL — `readCsv` not found.

- [ ] **Step 3: Implement the parser + inference**

```typescript
// src/csv/reader.ts
import type { CellValue } from '../model/cell'
import { createWorkbook, type Workbook } from '../model/workbook'
import type { CsvReadOptions } from './options'

/** Split CSV text into rows of raw string fields (RFC 4180). */
function parseRows(text: string, delimiter: string): string[][] {
  const rows: string[][] = []
  let field = ''
  let row: string[] = []
  let inQuotes = false
  let started = false // whether the current row has any content/fields yet
  const pushField = (): void => {
    row.push(field)
    field = ''
    started = true
  }
  const pushRow = (): void => {
    pushField()
    rows.push(row)
    row = []
    started = false
  }
  for (let i = 0; i < text.length; i++) {
    const ch = text[i]!
    if (inQuotes) {
      if (ch === '"') {
        if (text[i + 1] === '"') {
          field += '"'
          i++
        } else inQuotes = false
      } else field += ch
      continue
    }
    if (ch === '"' && field === '') inQuotes = true
    else if (ch === delimiter) pushField()
    else if (ch === '\n') pushRow()
    else if (ch === '\r') {
      /* swallow; the following \n (or row end) terminates the row */
    } else field += ch
  }
  // Flush a final row unless the text ended exactly on a row terminator.
  if (started || field !== '' || row.length > 0) pushRow()
  return rows
}

function infer(s: string, opts: Required<Pick<CsvReadOptions, 'parseNumbers' | 'parseBooleans'>>): CellValue {
  if (s === '') return null
  if (opts.parseNumbers && s.trim() !== '' && String(Number(s)) === s) return Number(s)
  if (opts.parseBooleans && (s === 'true' || s === 'false')) return s === 'true'
  return s
}

export function readCsv(text: string, options: CsvReadOptions = {}): Workbook {
  const { delimiter = ',', sheetName = 'Sheet1', parseNumbers = true, parseBooleans = true } = options
  const clean = text.startsWith('﻿') ? text.slice(1) : text
  const rows = parseRows(clean, delimiter)
  const wb = createWorkbook()
  const ws = wb.addSheet(sheetName)
  rows.forEach((fields, r) => {
    fields.forEach((raw, c) => {
      const value = infer(raw, { parseNumbers, parseBooleans })
      if (value !== null) ws.getCell(r + 1, c + 1).value = value
    })
  })
  return wb
}
```

- [ ] **Step 4: Run test to verify it passes**

Run: `npx vitest run src/csv/reader.test.ts`
Expected: PASS (3 tests). If the final-row flush double-counts or drops a row, adjust the flush condition — assert exact `ws.dimensions` in a scratch check.

- [ ] **Step 5: Run gates and commit**

Run: `npx vitest run && npm run -s typecheck && npm run -s lint && npm run -s format`

```bash
git add src/csv/reader.ts src/csv/reader.test.ts
git commit -m "feat(csv): RFC 4180 reader (readCsv) with conservative inference"
```

---

### Task 3: Public API, conformance, size + tree-shake

**Files:**
- Modify: `src/csv.ts` (replace placeholder), `size-budget.json` (if needed)
- Test: `src/csv/csv.test.ts`

**Interfaces:**
- Produces (from `./csv`): `writeCsv`, `readCsv`, `CsvWriteOptions`, `CsvReadOptions`.

- [ ] **Step 1: Replace the placeholder entry**

```typescript
// src/csv.ts
export { writeCsv } from './csv/writer'
export { readCsv } from './csv/reader'
export type { CsvWriteOptions, CsvReadOptions } from './csv/options'
```

- [ ] **Step 2: Write the conformance + round-trip test**

```typescript
// src/csv/csv.test.ts
import { Readable } from 'node:stream'
import { expect, test } from 'vitest'
import { createWorkbook } from '../model/workbook'
import { readCsv, writeCsv } from '../csv'

test('writeCsv -> readCsv round-trips values with conservative inference', () => {
  const wb = createWorkbook()
  const ws = wb.addSheet('S')
  ws.cell('A1').value = 'text'
  ws.cell('B1').value = 42
  ws.cell('C1').value = true
  ws.cell('A2').value = 'has,comma'
  ws.cell('B2').value = '00210' // must come back as a string, not 210
  const r = readCsv(writeCsv(wb)).sheets[0]!
  expect(r.cell('A1').value).toBe('text')
  expect(r.cell('B1').value).toBe(42)
  expect(r.cell('C1').value).toBe(true)
  expect(r.cell('A2').value).toBe('has,comma')
  expect(r.cell('B2').value).toBe('00210')
})

test('exceljs reads our CSV; we read exceljs CSV (cross-conformance)', async () => {
  const wb = createWorkbook()
  const ws = wb.addSheet('S')
  ws.cell('A1').value = 'a,b'
  ws.cell('B1').value = 7
  const ours = writeCsv(wb)

  const ExcelJS = (await import('exceljs')).default
  const oracle = new ExcelJS.Workbook()
  const oraSheet = await oracle.csv.read(Readable.from(ours))
  expect(oraSheet.getCell('A1').value).toBe('a,b')
  expect(oraSheet.getCell('B1').value).toBe(7)

  const theirs = (await oracle.csv.writeBuffer()).toString()
  const back = readCsv(theirs).sheets[0]!
  expect(back.cell('A1').value).toBe('a,b')
  expect(back.cell('B1').value).toBe(7)
})
```

- [ ] **Step 3: Build and verify tree-shaking + size**

Run: `npm run -s build`
Then assert `dist/csv.js` has zero xlsx/opc references:

```bash
# Expected: no matches (csv bundle must not pull the xlsx codecs or zip/opc code)
grep -Ec "writeXlsx|readXlsx|OpcPackage|sharedStrings|deflate|crc32" dist/csv.js || echo "clean: 0 xlsx refs"
```

Run: `npm run -s size`. If `csv.js` exceeds its 2048-byte budget (the model core travels with it), raise `size-budget.json#csv.js` to the measured gzip size + ~15% and note the measured value in the commit. Do not inflate it preemptively.

- [ ] **Step 4: Run all gates and commit**

Run: `npx vitest run && npm run -s typecheck && npm run -s lint && npm run -s format && npm run -s build && npm run -s size`

```bash
git add src/csv.ts src/csv/csv.test.ts size-budget.json
git commit -m "feat(csv): export writeCsv/readCsv from ./csv; conformance + size"
```

---

## SP-6 Acceptance Criteria

1. `writeCsv(Workbook | Worksheet, options?)` emits RFC-4180 CSV matching exceljs's quoting, with `delimiter`/`rowDelimiter`/`quoteAll`/`bom`/`sheet` options.
2. `readCsv(text, options?)` parses quoted fields, doubled quotes, embedded newlines, CRLF, and a BOM, into a one-sheet `Workbook`.
3. Inference is conservative: round-trip-safe numbers only (so `00210` stays a string), booleans optional, **no** date parsing — `test-issue-991.csv` round-trips as all-strings.
4. `writeCsv`→`readCsv` round-trips; exceljs and excelents read each other's CSV.
5. `dist/csv.js` contains zero xlsx/opc references and stays within (a measured, documented) size budget; all gates green; zero runtime deps.

## Self-Review

**1. Spec coverage** (architecture §SP-6 "CSV"): serialize → Task 1; parse → Task 2; entry/export + conformance + tree-shake → Task 3. The fixture trap (dates/leading-zeros) is covered by Task 2's fixture test and Task 3's round-trip. ✓

**2. Placeholder scan:** all three tasks carry complete code. One verify-against-source note (the `CellValue` union discriminants in `renderValue`) is called out explicitly in Task 1 Step 3 — the implementer must confirm `src/model/cell.ts` shapes, not guess. No "TBD"/"handle errors" placeholders.

**3. Type consistency:** `CsvWriteOptions`/`CsvReadOptions` are defined once in `options.ts` and imported by writer/reader/entry. `writeCsv(source: Workbook | Worksheet, options?)` and `readCsv(text, options?): Workbook` signatures are stable across Tasks 1-3 and the acceptance criteria. `renderValue`/`infer`/`parseRows`/`quoteField`/`resolveSheet` are private helpers; only `writeCsv`/`readCsv` + the two option types are exported. The `RangeBox` field names (`bottom`/`right`) used in the writer match `src/utils/range.ts`.

**Risk note:** The reader's final-row flush (Task 2) is the fiddliest bit — trailing newline vs genuine empty last row. Task 2 Step 4 calls for a scratch `ws.dimensions` check if the row count looks off. The size budget (Task 3) may need a documented bump since the model core ships inside `csv.js`; that's expected, not a regression.
