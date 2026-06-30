# SP-3 — xlsx write (baseline) Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Serialize the object model to a valid `.xlsx` file via a `writeXlsx(wb)` free function — workbook/worksheet/sharedStrings/styles parts assembled into an OPC package — validated by having the exceljs oracle read our output back and recover values + styles.

**Architecture:** A new `src/xlsx/` layer that turns the SP-2 model into OOXML parts using the SP-1c `XmlWriter`, assembles them with the SP-1d `OpcPackage`, and emits bytes. Serialization is a **free function** (`writeXlsx`) so a read-only consumer tree-shakes it away. Correctness is oracle-driven: each task asserts that exceljs (`ExcelJS.xlsx.load`) reads our output and recovers the data.

**Tech Stack:** TypeScript 7, Vitest 4, oxlint type-aware, oxfmt. Pure JS, zero runtime deps. `exceljs@4.4.0` devDep as the read-back oracle.

## Global Constraints

- **Web-standard core only:** `src/xlsx/*.ts` must import NO `node:` builtin (core-purity guard). Tests may use `node:`/import exceljs.
- **Zero runtime dependencies.** No third-party imports in `src/`.
- **`isolatedDeclarations`:** explicit return types on every export. **Named exports only.** **`sideEffects: false`.**
- All gates green: typecheck, lint (0/0), format, test, build, size.
- TDD: failing test first (RED), minimal implementation (GREEN), commit. Report RED/GREEN evidence.
- **Validation strategy:** the primary conformance check is **exceljs reads our output** — `const oracle = new ExcelJS.Workbook(); await oracle.xlsx.load(bytes)` then assert recovered cell values/styles. Supplement with XML-part assertions where exact structure matters (Task 5 adds the XML-diff comparator).
- **Scope IN:** values (string/number/boolean/date/formula/hyperlink/richText), base styling (font/fill/border/alignment/number-format), shared strings, merges, dimensions, column widths/row heights. **Scope OUT:** images, data validation, tables, conditional formatting, defined names (SP-5); reading (SP-4); streaming (SP-7).
- OOXML namespaces: main = `http://schemas.openxmlformats.org/spreadsheetml/2006/main`; relationships = `http://schemas.openxmlformats.org/officeDocument/2006/relationships`; package-rels = `http://schemas.openxmlformats.org/package/2006/relationships`.
- Branch: `excelents-rewrite` (continues from SP-2).

---

### Task 1: Minimal opening xlsx — workbook + worksheet writers + `writeXlsx`

**Files:**
- Create: `src/xlsx/content-types.ts` (content-type constants), `src/xlsx/worksheet-writer.ts`, `src/xlsx/workbook-writer.ts`, `src/xlsx/write.ts`
- Modify: `src/model/worksheet.ts` (add a `get rows(): readonly Row[]` accessor — populated rows ascending — needed to enumerate content for serialization)
- Test: `src/xlsx/write.test.ts`, plus a model test addition in `src/model/worksheet.test.ts`

**Interfaces:**
- Consumes: `Workbook`/`Worksheet`/`Row`/`Cell` from `../model/*`; `XmlWriter` from `../xml/writer`; `OpcPackage` from `../opc/package`; `encodeRange` from `../utils/range`.
- Produces:
  - `src/model/worksheet.ts`: `get rows(): readonly Row[]` (rows that have at least one populated cell, ascending by number).
  - `content-types.ts`: `CT` const map (workbook/worksheet/styles/sharedStrings/rels content-type strings).
  - `writeWorksheetXml(ws: Worksheet): string`
  - `writeWorkbookXml(wb: Workbook): string` and `writeWorkbookRelsXml(wb: Workbook): string`
  - `writeRootRelsXml(): string`
  - `writeXlsx(wb: Workbook): Promise<Uint8Array>` (assembles the OPC package, returns bytes)

- [ ] **Step 1: Write the failing test** (`src/xlsx/write.test.ts`)

```ts
import { expect, test } from 'vitest'
import { createWorkbook } from '../model/workbook'
import { writeXlsx } from './write'

test('writeXlsx produces a zip exceljs can open and read back values', async () => {
  const wb = createWorkbook()
  const ws = wb.addSheet('Sheet1')
  ws.cell('A1').value = 'Hello'
  ws.cell('B1').value = 42
  ws.cell('A2').value = true

  const bytes = await writeXlsx(wb)

  const ExcelJS = (await import('exceljs')).default
  const oracle = new ExcelJS.Workbook()
  await oracle.xlsx.load(Buffer.from(bytes))
  const sheet = oracle.getWorksheet('Sheet1')
  expect(sheet?.getCell('A1').value).toBe('Hello')
  expect(sheet?.getCell('B1').value).toBe(42)
  expect(sheet?.getCell('A2').value).toBe(true)
})

test('writeXlsx round-trips through our own OPC reader', async () => {
  const wb = createWorkbook()
  wb.addSheet('S').cell('A1').value = 'x'
  const { OpcPackage } = await import('../opc/package')
  const pkg = await OpcPackage.read(await writeXlsx(wb))
  expect(pkg.partNames()).toContain('xl/workbook.xml')
  expect(pkg.partNames()).toContain('xl/worksheets/sheet1.xml')
})
```

- [ ] **Step 2: Run the test to verify it fails**

Run: `pnpm exec vitest run src/xlsx/write.test.ts`
Expected: FAIL — cannot resolve `./write`.

- [ ] **Step 3: Add `get rows()` to `src/model/worksheet.ts`**

Add this getter to the `Worksheet` class (it exposes populated rows for serialization):

```ts
  get rows(): readonly Row[] {
    return [...this.rowsMap.keys()]
      .sort((a, b) => a - b)
      .map((n) => this.rowsMap.get(n)!)
      .filter((row) => row.cells.length > 0)
  }
```

(Adapt `this.rowsMap` to the actual private field name used for the rows `Map` — read the file. Import `Row` as a type if not already.) Add a model test in `src/model/worksheet.test.ts`:

```ts
test('rows lists only populated rows ascending', () => {
  const ws = new Worksheet('s')
  ws.cell('A3').value = 'c'
  ws.cell('A1').value = 'a'
  ws.getRow(2) // touched but empty — excluded
  expect(ws.rows.map((r) => r.number)).toEqual([1, 3])
})
```

- [ ] **Step 4: Write `src/xlsx/content-types.ts`**

```ts
/** OOXML content types for the parts we emit. */
export const CT = {
  workbook: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
  worksheet: 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml',
  styles: 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml',
  sharedStrings: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml',
  rels: 'application/vnd.openxmlformats-package.relationships+xml',
} as const
```

- [ ] **Step 5: Write `src/xlsx/worksheet-writer.ts`**

```ts
import type { Cell } from '../model/cell'
import type { Worksheet } from '../model/worksheet'
import { encodeRange } from '../utils/range'
import { XmlWriter } from '../xml/writer'

const MAIN_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'

function writeCellValue(w: XmlWriter, cell: Cell): void {
  const v = cell.value
  if (v === null) return
  if (typeof v === 'string') {
    w.open('c', { r: cell.address, t: 'inlineStr' }).open('is').open('t').text(v)
    w.close('t').close('is').close('c')
    return
  }
  if (typeof v === 'number') {
    w.open('c', { r: cell.address }).open('v').text(String(v)).close('v').close('c')
    return
  }
  if (typeof v === 'boolean') {
    w.open('c', { r: cell.address, t: 'b' }).open('v').text(v ? '1' : '0').close('v').close('c')
    return
  }
  // Other types (date/formula/hyperlink/richText) land in Task 4; for the baseline,
  // serialize their string form so the file stays valid.
  w.open('c', { r: cell.address, t: 'inlineStr' }).open('is').open('t').text(String(v))
  w.close('t').close('is').close('c')
}

/** Serialize one worksheet to `xl/worksheets/sheetN.xml`. */
export function writeWorksheetXml(ws: Worksheet): string {
  const w = new XmlWriter().declaration().open('worksheet', { xmlns: MAIN_NS })
  const dims = ws.dimensions
  if (dims !== undefined) w.leaf('dimension', { ref: encodeRange(dims) })
  w.open('sheetData')
  for (const row of ws.rows) {
    w.open('row', { r: row.number })
    for (const cell of row.cells) writeCellValue(w, cell)
    w.close('row')
  }
  w.close('sheetData').close('worksheet')
  return w.toString()
}
```

- [ ] **Step 6: Write `src/xlsx/workbook-writer.ts`**

```ts
import type { Workbook } from '../model/workbook'
import { XmlWriter } from '../xml/writer'

const MAIN_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
const REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const PKG_REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'
const OFFICE_DOC = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument'
const WORKSHEET_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet'

/** `xl/workbook.xml` — the sheet list. */
export function writeWorkbookXml(wb: Workbook): string {
  const w = new XmlWriter().declaration().open('workbook', { xmlns: MAIN_NS, 'xmlns:r': REL_NS })
  w.open('sheets')
  wb.sheets.forEach((sheet, i) => {
    w.leaf('sheet', { name: sheet.name, sheetId: i + 1, 'r:id': `rId${i + 1}` })
  })
  w.close('sheets').close('workbook')
  return w.toString()
}

/** `xl/_rels/workbook.xml.rels` — workbook → each worksheet. */
export function writeWorkbookRelsXml(wb: Workbook): string {
  const w = new XmlWriter().declaration().open('Relationships', { xmlns: PKG_REL_NS })
  wb.sheets.forEach((_sheet, i) => {
    w.leaf('Relationship', {
      Id: `rId${i + 1}`,
      Type: WORKSHEET_REL,
      Target: `worksheets/sheet${i + 1}.xml`,
    })
  })
  w.close('Relationships')
  return w.toString()
}

/** `_rels/.rels` — package → workbook. */
export function writeRootRelsXml(): string {
  return new XmlWriter()
    .declaration()
    .open('Relationships', { xmlns: PKG_REL_NS })
    .leaf('Relationship', { Id: 'rId1', Type: OFFICE_DOC, Target: 'xl/workbook.xml' })
    .close('Relationships')
    .toString()
}
```

- [ ] **Step 7: Write `src/xlsx/write.ts`**

```ts
import type { Workbook } from '../model/workbook'
import { OpcPackage } from '../opc/package'
import { CT } from './content-types'
import { writeRootRelsXml, writeWorkbookRelsXml, writeWorkbookXml } from './workbook-writer'
import { writeWorksheetXml } from './worksheet-writer'

const enc = (s: string): Uint8Array => new TextEncoder().encode(s)

/** Serialize a workbook to `.xlsx` bytes. */
export async function writeXlsx(wb: Workbook): Promise<Uint8Array> {
  const pkg = OpcPackage.empty()
  pkg.setPart('_rels/.rels', CT.rels, enc(writeRootRelsXml()))
  pkg.setPart('xl/workbook.xml', CT.workbook, enc(writeWorkbookXml(wb)))
  pkg.setPart('xl/_rels/workbook.xml.rels', CT.rels, enc(writeWorkbookRelsXml(wb)))
  wb.sheets.forEach((sheet, i) => {
    pkg.setPart(`xl/worksheets/sheet${i + 1}.xml`, CT.worksheet, enc(writeWorksheetXml(sheet)))
  })
  return pkg.toBytes()
}
```

- [ ] **Step 8: Run the tests to verify they pass**

Run: `pnpm exec vitest run src/xlsx/write.test.ts src/model/worksheet.test.ts`
Expected: PASS. If exceljs fails to load the file, the most likely causes are: a missing/incorrect content type, a malformed relationship target, or an unescaped value — inspect the exact exceljs error and fix the offending part. (Tip: write the bytes to a temp `.xlsx` and `unzip -l` it, or use our `OpcPackage.read` to dump part contents.)

- [ ] **Step 9: Final gates + commit**

Run: `pnpm typecheck && pnpm lint && pnpm format && pnpm test && pnpm build && pnpm size`
Expected: all green. (`writeXlsx` is not yet exported from `src/index.ts` — that's Task 5 — so `dist` entry sizes are unchanged.)

```bash
git add src/xlsx/ src/model/worksheet.ts src/model/worksheet.test.ts
git commit -m "feat(xlsx): minimal writeXlsx — opening xlsx exceljs can read (inline strings/numbers/bools)"
```

---

### Task 2: Shared strings

**Files:**
- Create: `src/xlsx/shared-strings-writer.ts`
- Modify: `src/xlsx/worksheet-writer.ts` (route string cells through the SST), `src/xlsx/write.ts` (emit `xl/sharedStrings.xml` + its relationship)
- Test: `src/xlsx/shared-strings.test.ts`

**Interfaces:**
- Consumes: `SharedStrings` from `../utils/shared-strings`; `XmlWriter`.
- Produces:
  - `writeSharedStringsXml(sst: SharedStrings): string` — `<sst count uniqueCount><si><t>...</t></si>...</sst>`.
  - `writeWorksheetXml(ws, sst)` updated to take a `SharedStrings`, emit string cells as `<c t="s"><v>index</v></c>`, and intern each string.

- [ ] **Step 1: Write the failing test** (`src/xlsx/shared-strings.test.ts`)

```ts
import { expect, test } from 'vitest'
import { createWorkbook } from '../model/workbook'
import { writeXlsx } from './write'

test('string cells round-trip via shared strings (exceljs reads them)', async () => {
  const wb = createWorkbook()
  const ws = wb.addSheet('S')
  ws.cell('A1').value = 'apple'
  ws.cell('A2').value = 'banana'
  ws.cell('A3').value = 'apple' // duplicate → same shared index
  const bytes = await writeXlsx(wb)

  const ExcelJS = (await import('exceljs')).default
  const oracle = new ExcelJS.Workbook()
  await oracle.xlsx.load(Buffer.from(bytes))
  const sheet = oracle.getWorksheet('S')
  expect(sheet?.getCell('A1').value).toBe('apple')
  expect(sheet?.getCell('A2').value).toBe('banana')
  expect(sheet?.getCell('A3').value).toBe('apple')

  // the part exists and dedups (2 unique strings)
  const { OpcPackage } = await import('../opc/package')
  const pkg = await OpcPackage.read(bytes)
  const sst = new TextDecoder().decode(pkg.getPart('xl/sharedStrings.xml')!)
  expect(sst).toContain('uniqueCount="2"')
})
```

- [ ] **Step 2-4: RED → implement → GREEN**

Run the test (FAIL: no sharedStrings part yet). Then:

Write `src/xlsx/shared-strings-writer.ts`:

```ts
import type { SharedStrings } from '../utils/shared-strings'
import { XmlWriter } from '../xml/writer'

const MAIN_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'

/** Serialize the shared-string table to `xl/sharedStrings.xml`. */
export function writeSharedStringsXml(sst: SharedStrings): string {
  const w = new XmlWriter()
    .declaration()
    .open('sst', { xmlns: MAIN_NS, count: sst.count, uniqueCount: sst.uniqueCount })
  for (const value of sst.values) w.open('si').open('t').text(value).close('t').close('si')
  return w.close('sst').toString()
}
```

Update `writeWorksheetXml(ws, sst)` so a string cell interns via `sst.add(value)` and emits `w.open('c', { r: address, t: 's' }).open('v').text(String(index)).close('v').close('c')`. Update `src/xlsx/write.ts` to: create a `SharedStrings`, pass it to each `writeWorksheetXml`, then emit `xl/sharedStrings.xml` (content type `CT.sharedStrings`) and add a `sharedStrings` relationship to `xl/_rels/workbook.xml.rels` (Type `.../sharedStrings`, Target `sharedStrings.xml`) — only when `sst.uniqueCount > 0`. Note: the SST must be fully populated (all sheets serialized) before `writeSharedStringsXml` is called, and the workbook-rels must include the sharedStrings + (later) styles relationships — be careful to keep relationship Ids consistent. The simplest ordering: serialize all worksheets first (interning strings), then emit sharedStrings, then build workbook rels listing worksheets + sharedStrings.

- [ ] **Step 5: GREEN + gates + commit**

Run: `pnpm exec vitest run src/xlsx/ && pnpm typecheck && pnpm lint && pnpm test`
Expected: PASS; exceljs reads the shared strings; `uniqueCount="2"`.

```bash
git add src/xlsx/
git commit -m "feat(xlsx): shared strings (sst dedup; string cells via t=s)"
```

---

### Task 3: Styles

**Files:**
- Create: `src/xlsx/styles-writer.ts`
- Modify: `src/xlsx/worksheet-writer.ts` (emit `s="xfIndex"` on styled cells), `src/xlsx/write.ts` (emit `xl/styles.xml` + relationship)
- Test: `src/xlsx/styles.test.ts`

**Interfaces:**
- Produces:
  - `class StyleRegistry` — interns `CellStyle`s into the OOXML style tables (numFmts, fonts, fills, borders, cellXfs) and returns an `xf` index per cell; `xfIndexFor(style: CellStyle): number`.
  - `writeStylesXml(registry: StyleRegistry): string` — `<styleSheet><numFmts/><fonts/><fills/><borders/><cellXfs/></styleSheet>`.

**Notes:** OOXML styles require: index 0 of fonts/fills/borders/cellXfs is the default; `fills` must contain at least the two mandatory default fills (`none` at 0, `gray125` at 1) — Excel rejects files otherwise. Custom number formats use ids >= 164. Build the registry by collecting every cell's `style`, interning each facet (font/fill/border/numFmt) into its table, and creating one `cellXf` per unique combination. The empty style maps to xf 0 (the default).

- [ ] **Step 1: Write the failing test** (`src/xlsx/styles.test.ts`)

```ts
import { expect, test } from 'vitest'
import { createWorkbook } from '../model/workbook'
import { writeXlsx } from './write'

test('a bold cell with a number format round-trips through exceljs', async () => {
  const wb = createWorkbook()
  const ws = wb.addSheet('S')
  const c = ws.cell('A1')
  c.value = 1234.5
  c.style = { font: { bold: true }, numberFormat: '0.00' }
  const bytes = await writeXlsx(wb)

  const ExcelJS = (await import('exceljs')).default
  const oracle = new ExcelJS.Workbook()
  await oracle.xlsx.load(Buffer.from(bytes))
  const cell = oracle.getWorksheet('S')?.getCell('A1')
  expect(cell?.value).toBe(1234.5)
  expect(cell?.font?.bold).toBe(true)
  expect(cell?.numFmt).toBe('0.00')
})
```

- [ ] **Step 2-5: RED → implement `StyleRegistry` + `writeStylesXml` → wire into write.ts → GREEN → commit**

Implement `src/xlsx/styles-writer.ts` with the `StyleRegistry` (intern fonts/fills/borders/numFmts; map each `CellStyle` to a `cellXf` index; index 0 = default empty style; fills index 0 = none, 1 = gray125). `writeStylesXml` emits the `<styleSheet>` with `numFmts`, `fonts`, `fills`, `borders`, `cellXfs` in that order (OOXML schema order matters). Update `writeWorksheetXml` to look up `registry.xfIndexFor(cell.style)` and add `s="<index>"` to the `<c>` attributes when the index is non-zero. Update `write.ts` to build the registry (one pass over all cells), emit `xl/styles.xml` (content type `CT.styles`) + a `styles` relationship. Validate via the exceljs read-back above.

```bash
git add src/xlsx/
git commit -m "feat(xlsx): styles.xml (fonts/fills/borders/numFmts/cellXfs); cell style refs"
```

---

### Task 4: Full value encoding + merges + dimensions + sizing

**Files:**
- Modify: `src/xlsx/worksheet-writer.ts` (date/formula/hyperlink/richText encoding; `<mergeCells>`; `<cols>` widths; row `ht`)
- Test: `src/xlsx/values.test.ts`

**Interfaces:**
- `writeWorksheetXml` handles every `CellValue`:
  - date → numeric serial (`dateToSerial`) + a date number-format applied via the style registry (so exceljs reads it back as a Date).
  - formula → `<c><f>EXPR</f><v>RESULT</v></c>`.
  - hyperlink → the cell text + a `<hyperlinks>` entry (relationship-based) OR the simpler `HYPERLINK()` formula form — choose the form exceljs reads back; assert it.
  - richText → `<c t="s">` pointing at a rich shared string (`<si>` with multiple `<r><rPr>…</rPr><t>…</t></r>` runs).
  - merges → `<mergeCells count><mergeCell ref="A1:B2"/></mergeCells>` after `</sheetData>`.
  - column widths → `<cols><col min max width customWidth="1"/></cols>` before `<sheetData>`; row heights → `<row r ht customHeight="1">`.

- [ ] **Steps:** TDD each value form against an exceljs read-back assertion (date → `instanceof Date`; formula → `cell.formula`/`cell.result`; hyperlink → `cell.hyperlink`; merges → `sheet.getCell('A1').isMerged`). Element ordering within `<worksheet>` matters (cols → sheetData → mergeCells) — follow the OOXML schema sequence. Commit when exceljs reads every form back.

```bash
git commit -m "feat(xlsx): full value encoding (date/formula/hyperlink/richText), merges, dimensions, sizing"
```

---

### Task 5: Oracle XML-diff comparator + public `writeXlsx` + conformance sweep

**Files:**
- Modify: `test/conformance/harness.ts` (`canonicalizeXml` whitespace-normalization; add `compareWriteParts(build)` helper), `src/index.ts` (export `writeXlsx`), `src/index.test.ts`
- Test: `src/xlsx/conformance.test.ts`

**Interfaces:**
- `canonicalizeXml` updated to drop **whitespace-only text runs between elements** (insignificant formatting whitespace) while preserving non-whitespace text — so comparing two producers' compact XML doesn't false-diff on indentation. Document the one edge case (a string that is itself only whitespace inside `<t xml:space="preserve">`) as a known limitation.
- `compareWriteParts(build: (lib: 'oracle' | 'excelents') => Promise<Uint8Array>)` — builds the same workbook with exceljs and excelents, unzips+canonicalizes both, returns the `diffParts` result for the parts we own (filter to `xl/worksheets/*`, `xl/sharedStrings.xml`, etc.).
- `src/index.ts` adds `export { writeXlsx } from './xlsx/write'` (the public serialization free function).

- [ ] **Steps:** Update `canonicalizeXml` (+ a unit test that whitespace-only text is dropped but content text preserved). Add `compareWriteParts`. Write `src/xlsx/conformance.test.ts` that builds a representative workbook (a few strings, numbers, a bold/number-format cell, a merge) with both libraries and asserts the **worksheet `sheetData` cell structure matches** after canonicalization (or, where exceljs differs in incidental attributes, assert the cell `<c>`/`<v>` content matches — be pragmatic: the goal is semantic equivalence, not byte-identity, since exceljs adds metadata parts we don't). Export `writeXlsx`, update `index.test.ts` to call `writeXlsx(createWorkbook())` and assert it returns bytes starting with the ZIP magic `PK`. Raise `index.js` size budget if needed. Final full-gate run.

```bash
git add src/xlsx/ src/index.ts src/index.test.ts test/conformance/harness.ts size-budget.json
git commit -m "feat(xlsx): public writeXlsx + oracle XML-diff conformance comparator"
```

---

## SP-3 Acceptance Criteria

1. `src/xlsx/` contains the writers (`worksheet`/`workbook`/`shared-strings`/`styles`), `content-types.ts`, and `write.ts`.
2. `writeXlsx(wb)` produces a `.xlsx` that **exceljs opens and reads back** — values (string/number/bool/date/formula/hyperlink/richText), bold + number-format styling, shared strings, merges.
3. `writeXlsx` is exported from `src/index.ts` (the public serialization free function); it tree-shakes away for read-only consumers (verify via bundle inspection).
4. An oracle XML-diff comparator validates the worksheet structure against exceljs for a representative workbook.
5. All `src/xlsx/*` are `node:`-free and dependency-free; all gates green (typecheck, lint 0/0, format, test, build, size).

---

## Self-Review

**1. Spec coverage** (architecture "SP-3 — xlsx write (baseline)"):
- Object model → valid `.xlsx` with base styling → Tasks 1-4. ✓
- Validated via the oracle → exceljs read-back every task + XML-diff comparator (Task 5). ✓
- `writeXlsx` free function (treeshakeable serialization) → Tasks 1 + 5. ✓
- Shared strings, styles, value encoding, merges, dimensions, sizing → Tasks 2-4. ✓
- Whitespace-normalization decision for the comparator (flagged by the SP-1 review) → Task 5. ✓

**2. Placeholder scan:** Tasks 1-3 have complete code; Tasks 4-5 specify exact OOXML forms + the exceljs read-back assertions per value type rather than full code for every branch (the value-encoding branches are mechanical extensions of Task 1's `writeCellValue`, each gated by a concrete read-back test). The executor writes each branch test-first. This is a deliberate density choice for the most mechanical/repetitive task, not a placeholder — each form names its exact XML and its exact assertion.

**3. Type consistency:** `writeWorksheetXml` gains a `SharedStrings` parameter in Task 2 and a `StyleRegistry` in Task 3 — the executor threads these through `write.ts` consistently (the plan calls this out at each step). `OpcPackage.setPart`/`toBytes`, `XmlWriter`, `encodeRange`, `dateToSerial`, `SharedStrings` are all SP-1/SP-2 APIs used as-defined. `writeXlsx(wb: Workbook): Promise<Uint8Array>` is stable across tasks.

Note for the executor: the relationship-Id bookkeeping in `xl/_rels/workbook.xml.rels` (worksheets, then sharedStrings, then styles) is the easiest thing to get wrong — keep a single ordered list of relationships and assign Ids from it. The exceljs read-back is the safety net: if a relationship Id or Target is off, exceljs throws on load.
