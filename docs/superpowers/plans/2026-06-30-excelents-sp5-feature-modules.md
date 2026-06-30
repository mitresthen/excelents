# SP-5: Feature Modules (defined names, data validation, tables) Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add three self-contained, independently-shippable feature slices on top of the SP-1..SP-4 read+write core — **defined names**, **data validation**, and **tables** — each with model API + serialize + parse + round-trip validation.

**Architecture:** Each feature is a vertical slice that extends the model (a new method/getter), the relevant writer, and the matching reader, validated by `readXlsx(await writeXlsx(wb))` round-trips. Defined names live in `xl/workbook.xml`; data validations live inside each `xl/worksheets/sheetN.xml`; tables introduce a new OPC part (`xl/tables/tableN.xml`) referenced from the worksheet via a relationship and a `<tableParts>` element — reusing the per-sheet relationship plumbing already built for hyperlinks.

**Tech Stack:** TypeScript 7 (isolatedDeclarations), the owned XML `tokenize`/`XmlWriter`, the owned `OpcPackage`, the SP-2 model, Vitest 4 + `exceljs@4.4.0` oracle. Zero runtime dependencies; `node:`-free in `src/`.

## Global Constraints

- **Zero runtime dependencies.** No new entries in `package.json#dependencies`.
- **`node:`-free core.** No `node:` imports in `src/` except `src/node.ts` (enforced by `test/core-purity.test.ts`).
- **`isolatedDeclarations: true`.** Every exported function/method/const needs an explicit return type.
- **TDD, frequent commits.** Red → green → refactor per behavior. Run `gates && git commit` so a failing gate blocks the commit (a `Font`-import slip got through a newline-separated chain in SP-4 — chain with `&&`).
- **Validation:** round-trip (`readXlsx(writeXlsx(wb))`) is primary; supplement with a direct exceljs read-back where the oracle's model is richer. Each feature is independent — a reviewer can approve/reject one task without the others.
- **Scope IN:** global defined names; data validation (`list`, `whole`, `decimal`, `textLength`, `date` types with operators + `allowBlank`/`showDropDown`/`showErrorMessage`); tables (name, ref, column names, header row, a built-in table style). **Scope OUT (later plans):** images/drawings (binary media + drawingML), conditional formatting (complex rule grammar), sheet-scoped (`localSheetId`) defined names, pivot tables, comments. Unknown elements are ignored on read, never throw.

---

## File Structure

- `src/model/workbook.ts` — **modify**: `defineName(name, formula)`, `get definedNames`.
- `src/model/worksheet.ts` — **modify**: `addDataValidation(rule)`, `get dataValidations`; `addTable(def)`, `get tables`.
- `src/model/data-validation.ts` — **create**: `DataValidation` type.
- `src/model/table.ts` — **create**: `TableDefinition` type.
- `src/xlsx/workbook-writer.ts` / `workbook-reader.ts` — **modify**: `<definedNames>` write/read.
- `src/xlsx/worksheet-writer.ts` / `worksheet-reader.ts` — **modify**: `<dataValidations>` write/read; `<tableParts>` write + table rel collection / read.
- `src/xlsx/table-writer.ts` / `table-reader.ts` — **create**: `xl/tables/tableN.xml` serialize/parse.
- `src/xlsx/write.ts` / `read.ts` — **modify**: emit/parse table parts + rels + content-type override.
- `src/xlsx/content-types.ts` — **modify**: add `table` content type.
- Tests: `src/xlsx/defined-names.test.ts`, `src/xlsx/data-validation.test.ts`, `src/xlsx/tables.test.ts`, `src/xlsx/read-conformance.test.ts` (append).

**Shared types:**

```typescript
// src/model/data-validation.ts
export interface DataValidation {
  readonly sqref: string // target range(s), e.g. 'A1:A10' or 'A1 C1'
  readonly type: 'list' | 'whole' | 'decimal' | 'textLength' | 'date'
  readonly operator?:
    | 'between' | 'notBetween' | 'equal' | 'notEqual'
    | 'greaterThan' | 'lessThan' | 'greaterThanOrEqual' | 'lessThanOrEqual'
  readonly formula1: string
  readonly formula2?: string
  readonly allowBlank?: boolean
  readonly showDropDown?: boolean
  readonly showErrorMessage?: boolean
}

// src/model/table.ts
export interface TableDefinition {
  readonly name: string
  readonly ref: string // e.g. 'A1:C4' (includes header row)
  readonly columns: readonly string[] // header names, left to right
  readonly headerRow?: boolean // default true
  readonly styleName?: string // a built-in style, e.g. 'TableStyleMedium2'
}
```

---

### Task 1: Defined names

**Files:**
- Modify: `src/model/workbook.ts`, `src/xlsx/workbook-writer.ts`, `src/xlsx/workbook-reader.ts`, `src/xlsx/read.ts`
- Test: `src/xlsx/defined-names.test.ts`

**Interfaces:**
- Consumes: `Workbook`, `XmlWriter`, `tokenize`, `WorkbookParts`.
- Produces: `Workbook.defineName(name: string, formula: string): void`; `Workbook.definedNames: ReadonlyArray<{ name: string; formula: string }>`; `WorkbookParts.definedNames: ReadonlyArray<{ name: string; formula: string }>`.

A defined name maps a name to a formula/reference. In `xl/workbook.xml`: `<definedNames><definedName name="MyRange">Sheet1!$A$1:$B$2</definedName></definedNames>`, placed after `</sheets>` (CT_Workbook sequence: sheets → functionGroups → externalReferences → definedNames). The formula text is element content (escape it).

- [ ] **Step 1: Write the failing test (model + round-trip)**

```typescript
// src/xlsx/defined-names.test.ts
import { expect, test } from 'vitest'
import { createWorkbook } from '../model/workbook'
import { readXlsx } from './read'
import { writeXlsx } from './write'

test('a workbook records defined names', () => {
  const wb = createWorkbook()
  wb.addSheet('Sheet1')
  wb.defineName('MyRange', 'Sheet1!$A$1:$B$2')
  expect(wb.definedNames).toEqual([{ name: 'MyRange', formula: 'Sheet1!$A$1:$B$2' }])
})

test('defined names round-trip through write+read', async () => {
  const wb = createWorkbook()
  wb.addSheet('Sheet1')
  wb.defineName('MyRange', 'Sheet1!$A$1:$B$2')
  wb.defineName('Tax', '0.2')
  const restored = await readXlsx(await writeXlsx(wb))
  expect(restored.definedNames).toEqual([
    { name: 'MyRange', formula: 'Sheet1!$A$1:$B$2' },
    { name: 'Tax', formula: '0.2' },
  ])
})
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx vitest run src/xlsx/defined-names.test.ts`
Expected: FAIL — `defineName` not a function.

- [ ] **Step 3: Add the model API**

```typescript
// src/model/workbook.ts — inside class Workbook
private readonly names: Array<{ name: string; formula: string }> = []

defineName(name: string, formula: string): void {
  this.names.push({ name, formula })
}

get definedNames(): ReadonlyArray<{ name: string; formula: string }> {
  return this.names
}
```

- [ ] **Step 4: Emit `<definedNames>` in the workbook writer**

In `src/xlsx/workbook-writer.ts`, `writeWorkbookXml`, after `w.close('sheets')` and before `.close('workbook')`:

```typescript
const names = wb.definedNames
if (names.length > 0) {
  w.open('definedNames')
  for (const { name, formula } of names) {
    w.open('definedName', { name }).text(formula).close('definedName')
  }
  w.close('definedNames')
}
```

- [ ] **Step 5: Parse `<definedName>` in the workbook reader**

In `src/xlsx/workbook-reader.ts`, extend `WorkbookParts` with `definedNames` and parse them in the same `tokenize` pass over `xl/workbook.xml`:

```typescript
// add to WorkbookParts
readonly definedNames: ReadonlyArray<{ name: string; formula: string }>
```
```typescript
// in readWorkbookParts: track defined-name state in the existing token loop
const definedNames: Array<{ name: string; formula: string }> = []
let dnName: string | undefined
let dnText = ''
let inDefinedName = false
// ... inside the loop:
//   open 'definedName': inDefinedName = true; dnName = tok.attributes['name']; dnText = ''
//   text when inDefinedName: dnText += tok.value
//   close 'definedName': if dnName !== undefined push { name: dnName, formula: dnText }; inDefinedName = false
// return { ...existing, definedNames }
```

- [ ] **Step 6: Apply defined names in `readXlsx`**

In `src/xlsx/read.ts`, after building `wb` and before returning:

```typescript
for (const { name, formula } of parts.definedNames) wb.defineName(name, formula)
```

- [ ] **Step 7: Run test to verify it passes**

Run: `npx vitest run src/xlsx/defined-names.test.ts`
Expected: PASS (2 tests).

- [ ] **Step 8: Run gates and commit**

Run: `npx vitest run && npm run -s typecheck && npm run -s lint && npm run -s format`

```bash
git add src/model/workbook.ts src/xlsx/workbook-writer.ts src/xlsx/workbook-reader.ts src/xlsx/read.ts src/xlsx/defined-names.test.ts
git commit -m "feat(xlsx): defined names (workbook.xml <definedNames>) write+read"
```

---

### Task 2: Data validation

**Files:**
- Create: `src/model/data-validation.ts`
- Modify: `src/model/worksheet.ts`, `src/xlsx/worksheet-writer.ts`, `src/xlsx/worksheet-reader.ts`
- Test: `src/xlsx/data-validation.test.ts`

**Interfaces:**
- Consumes: `Worksheet`, `XmlWriter`, `tokenize`, `ReadContext`.
- Produces: `DataValidation` (type); `Worksheet.addDataValidation(rule: DataValidation): void`; `Worksheet.dataValidations: readonly DataValidation[]`.

In `xl/worksheets/sheetN.xml`, `<dataValidations count="1"><dataValidation type="list" sqref="A1:A10" allowBlank="1" showDropDown="0"><formula1>"a,b,c"</formula1></dataValidation></dataValidations>`, placed **after `</sheetData>` and `<mergeCells>`, before `<hyperlinks>`** (CT_Worksheet sequence: mergeCells → conditionalFormatting → dataValidations → hyperlinks). Note: OOXML `showDropDown="1"` actually means *hide* the dropdown (the attribute is inverted); model `showDropDown` defaults to showing it, so write `showDropDown="1"` only when the model explicitly sets `showDropDown: false`. Encode that carefully and assert via round-trip.

- [ ] **Step 1: Write the failing test**

```typescript
// src/xlsx/data-validation.test.ts
import { expect, test } from 'vitest'
import { createWorkbook } from '../model/workbook'
import { readXlsx } from './read'
import { writeXlsx } from './write'

test('a list data validation round-trips through write+read', async () => {
  const wb = createWorkbook()
  const ws = wb.addSheet('S')
  ws.cell('A1').value = 'pick'
  ws.addDataValidation({
    sqref: 'A1:A10',
    type: 'list',
    formula1: '"apple,banana,cherry"',
    allowBlank: true,
  })
  const r = await readXlsx(await writeXlsx(wb))
  expect(r.sheets[0]!.dataValidations).toEqual([
    { sqref: 'A1:A10', type: 'list', formula1: '"apple,banana,cherry"', allowBlank: true },
  ])
})

test('a numeric between validation round-trips with operator and two formulas', async () => {
  const wb = createWorkbook()
  const ws = wb.addSheet('S')
  ws.addDataValidation({
    sqref: 'B1:B5',
    type: 'whole',
    operator: 'between',
    formula1: '1',
    formula2: '100',
  })
  const r = await readXlsx(await writeXlsx(wb))
  expect(r.sheets[0]!.dataValidations[0]).toMatchObject({
    sqref: 'B1:B5',
    type: 'whole',
    operator: 'between',
    formula1: '1',
    formula2: '100',
  })
})
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx vitest run src/xlsx/data-validation.test.ts`
Expected: FAIL — `addDataValidation` not a function.

- [ ] **Step 3: Create the type and model API**

```typescript
// src/model/data-validation.ts  (full content shown under Shared types above)
```
```typescript
// src/model/worksheet.ts — add import + fields
import type { DataValidation } from './data-validation'
// inside class:
private readonly validations: DataValidation[] = []
addDataValidation(rule: DataValidation): void {
  this.validations.push(rule)
}
get dataValidations(): readonly DataValidation[] {
  return this.validations
}
```

- [ ] **Step 4: Emit `<dataValidations>` in the worksheet writer**

In `src/xlsx/worksheet-writer.ts`, after the `mergeCells` block and before the `hyperlinks` block:

```typescript
const validations = ws.dataValidations
if (validations.length > 0) {
  w.open('dataValidations', { count: validations.length })
  for (const dv of validations) {
    w.open('dataValidation', {
      type: dv.type,
      operator: dv.operator,
      sqref: dv.sqref,
      allowBlank: dv.allowBlank === true ? 1 : undefined,
      showDropDown: dv.showDropDown === false ? 1 : undefined, // OOXML inverts this flag
      showErrorMessage: dv.showErrorMessage === true ? 1 : undefined,
    })
    w.open('formula1').text(dv.formula1).close('formula1')
    if (dv.formula2 !== undefined) w.open('formula2').text(dv.formula2).close('formula2')
    w.close('dataValidation')
  }
  w.close('dataValidations')
}
```

- [ ] **Step 5: Parse `<dataValidation>` in the worksheet reader**

In `src/xlsx/worksheet-reader.ts`, add state to the token loop: on `<dataValidation>` open capture attrs into a partial; on `<formula1>`/`<formula2>` capture following text; on close push a `DataValidation` to a local array and, after the loop, `for (const dv of parsed) ws.addDataValidation(dv)`. Decode `allowBlank`/`showErrorMessage` as `=== '1'`; `showDropDown` is `false` only when the attribute is `'1'` (inverted), else omit. Narrow `type`/`operator` with control-flow helpers (no casts) like `asBorderStyle` in `styles-reader.ts`.

- [ ] **Step 6: Run test to verify it passes**

Run: `npx vitest run src/xlsx/data-validation.test.ts`
Expected: PASS.

- [ ] **Step 7: Run gates and commit**

Run: `npx vitest run && npm run -s typecheck && npm run -s lint && npm run -s format`

```bash
git add src/model/data-validation.ts src/model/worksheet.ts src/xlsx/worksheet-writer.ts src/xlsx/worksheet-reader.ts src/xlsx/data-validation.test.ts
git commit -m "feat(xlsx): data validation (<dataValidations>) write+read"
```

---

### Task 3: Tables

**Files:**
- Create: `src/model/table.ts`, `src/xlsx/table-writer.ts`, `src/xlsx/table-reader.ts`
- Modify: `src/model/worksheet.ts`, `src/xlsx/content-types.ts`, `src/xlsx/worksheet-writer.ts` (emit `<tableParts>` + collect table rels), `src/xlsx/worksheet-reader.ts` (parse `<tableParts>`), `src/xlsx/write.ts` (emit table parts + rels + content-type), `src/xlsx/read.ts` (resolve + parse table parts)
- Test: `src/xlsx/tables.test.ts`

**Interfaces:**
- Consumes: `Worksheet`, `OpcPackage`, `XmlWriter`, `tokenize`, the per-sheet relationship plumbing (`WorksheetWriteResult`).
- Produces: `TableDefinition` (type); `Worksheet.addTable(def: TableDefinition): void`; `Worksheet.tables: readonly TableDefinition[]`; `writeTableXml(table, id): string`; `readTableXml(xml): TableDefinition`; `WorksheetWriteResult.tables: ReadonlyArray<{ rid: string; table: TableDefinition }>`; `CT.table`.

A table is a separate part `xl/tables/tableN.xml`:
```xml
<table xmlns=".../main" id="1" name="Table1" displayName="Table1" ref="A1:C4">
  <autoFilter ref="A1:C4"/>
  <tableColumns count="3">
    <tableColumn id="1" name="Col1"/><tableColumn id="2" name="Col2"/><tableColumn id="3" name="Col3"/>
  </tableColumns>
  <tableStyleInfo name="TableStyleMedium2" showRowStripes="1"/>
</table>
```
The worksheet references it: a relationship in `xl/worksheets/_rels/sheetN.xml.rels` (type `.../table`, Target `../tables/tableN.xml`) plus `<tableParts count="1"><tablePart r:id="rIdK"/></tableParts>` at the very end of the worksheet (CT_Worksheet: tableParts is near the end, after hyperlinks/printOptions/pageMargins — emit it last, before `</worksheet>`). `[Content_Types].xml` needs an Override for each table part. Table parts are numbered globally across the workbook (`table1.xml`, `table2.xml`, ...).

- [ ] **Step 1: Write the failing test**

```typescript
// src/xlsx/tables.test.ts
import { expect, test } from 'vitest'
import { createWorkbook } from '../model/workbook'
import { readXlsx } from './read'
import { writeXlsx } from './write'

test('a table round-trips through write+read', async () => {
  const wb = createWorkbook()
  const ws = wb.addSheet('S')
  ws.cell('A1').value = 'Col1'
  ws.cell('B1').value = 'Col2'
  ws.cell('C1').value = 'Col3'
  ws.addTable({ name: 'Table1', ref: 'A1:C2', columns: ['Col1', 'Col2', 'Col3'] })
  const r = await readXlsx(await writeXlsx(wb))
  expect(r.sheets[0]!.tables).toEqual([
    { name: 'Table1', ref: 'A1:C2', columns: ['Col1', 'Col2', 'Col3'] },
  ])
})

test('exceljs reads our table back', async () => {
  const wb = createWorkbook()
  const ws = wb.addSheet('S')
  ws.cell('A1').value = 'H1'
  ws.cell('B1').value = 'H2'
  ws.addTable({ name: 'T', ref: 'A1:B2', columns: ['H1', 'H2'] })
  const bytes = await writeXlsx(wb)
  const ExcelJS = (await import('exceljs')).default
  const oracle = new ExcelJS.Workbook()
  // @ts-expect-error -- Buffer typing
  await oracle.xlsx.load(Buffer.from(bytes))
  const t = oracle.getWorksheet('S')?.getTable('T')
  expect(t).toBeDefined()
})
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx vitest run src/xlsx/tables.test.ts`
Expected: FAIL — `addTable` not a function.

- [ ] **Step 3: Create the type + model API + table writer/reader**

```typescript
// src/model/table.ts (full content under Shared types above)
```
```typescript
// src/model/worksheet.ts — add
import type { TableDefinition } from './table'
private readonly tableDefs: TableDefinition[] = []
addTable(def: TableDefinition): void { this.tableDefs.push(def) }
get tables(): readonly TableDefinition[] { return this.tableDefs }
```
```typescript
// src/xlsx/table-writer.ts
import { XmlWriter } from '../xml/writer'
import type { TableDefinition } from '../model/table'
const MAIN_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
export function writeTableXml(table: TableDefinition, id: number): string {
  const w = new XmlWriter().declaration().open('table', {
    xmlns: MAIN_NS, id, name: table.name, displayName: table.name, ref: table.ref,
  })
  w.leaf('autoFilter', { ref: table.ref })
  w.open('tableColumns', { count: table.columns.length })
  table.columns.forEach((name, i) => w.leaf('tableColumn', { id: i + 1, name }))
  w.close('tableColumns')
  w.leaf('tableStyleInfo', {
    name: table.styleName ?? 'TableStyleMedium2', showRowStripes: 1,
  })
  return w.close('table').toString()
}
```
```typescript
// src/xlsx/table-reader.ts
import type { TableDefinition } from '../model/table'
import { tokenize } from '../xml/tokenizer'
export function readTableXml(xml: string): TableDefinition {
  let name = ''; let ref = ''; let styleName: string | undefined
  const columns: string[] = []
  for (const tok of tokenize(xml)) {
    if (tok.type !== 'open') continue
    if (tok.name === 'table') {
      name = tok.attributes['name'] ?? tok.attributes['displayName'] ?? ''
      ref = tok.attributes['ref'] ?? ''
    } else if (tok.name === 'tableColumn') {
      const c = tok.attributes['name']; if (c !== undefined) columns.push(c)
    } else if (tok.name === 'tableStyleInfo') {
      styleName = tok.attributes['name']
    }
  }
  return styleName !== undefined ? { name, ref, columns, styleName } : { name, ref, columns }
}
```

- [ ] **Step 4: Add the table content type**

```typescript
// src/xlsx/content-types.ts — add to CT
table: 'application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml',
```

- [ ] **Step 5: Emit `<tableParts>` + collect table rels in the worksheet writer**

Extend `WorksheetWriteResult` with `tables: ReadonlyArray<{ rid: string; table: TableDefinition }>`. In `writeWorksheetXml`, after the `hyperlinks` block and before `</worksheet>`, assign each table an rId (continuing the per-sheet rId counter past hyperlinks) and emit:
```typescript
const tableRels: Array<{ rid: string; table: TableDefinition }> = []
const ws_tables = ws.tables
if (ws_tables.length > 0) {
  w.open('tableParts', { count: ws_tables.length })
  ws_tables.forEach((table, i) => {
    const rid = `rId${pending.length + i + 1}` // continue past hyperlink rIds
    w.leaf('tablePart', { 'r:id': rid })
    tableRels.push({ rid, table })
  })
  w.close('tableParts')
}
// return { xml, hyperlinks: rels, tables: tableRels }
```
Note: the hyperlink rIds are assigned `rId1..rIdN` (N = pending.length); tables continue at `rId(N+1)`. Keep the rId numbering consistent between the `<tablePart r:id>` here and the rels written in `write.ts`.

- [ ] **Step 6: Emit table parts, rels, and content-type overrides in `write.ts`**

Track a global table counter. For each sheet's `result.tables`, write `xl/tables/table{globalId}.xml` (content type `CT.table`) and add a relationship to that sheet's `xl/worksheets/_rels/sheetN.xml.rels` (type `.../table`, Target `../tables/table{globalId}.xml`). The sheet rels part must merge hyperlink rels + table rels (write a combined rels part). Extend `writeWorksheetRelsXml` to accept tables too, or write a second helper and merge — keep all of a sheet's relationships in one `.rels` part.

- [ ] **Step 7: Parse `<tableParts>` + table parts in the reader**

In `worksheet-reader.ts`, collect `<tablePart r:id>` rIds (like hyperlinks). In `read.ts`, for each sheet resolve those rIds via `relationshipsFor(sheet.path)` (type ends `/table`) to the table part path, parse with `readTableXml`, and `ws.addTable(def)`. Pass the table-rel map into the worksheet reader the same way as `hyperlinkTargets`, or resolve in `read.ts` after the worksheet parse (simpler: have `readWorksheetInto` expose collected table rIds, then resolve+add in `read.ts`).

- [ ] **Step 8: Run test to verify it passes**

Run: `npx vitest run src/xlsx/tables.test.ts`
Expected: PASS. If exceljs rejects the file, inspect what part/rel/content-type it expects (a table needs the worksheet rel + the Override + a valid `ref`); fix the writer, not the test.

- [ ] **Step 9: Run gates and commit**

Run: `npx vitest run && npm run -s typecheck && npm run -s lint && npm run -s format && npm run -s build && npm run -s size`

```bash
git add src/model/table.ts src/model/worksheet.ts src/xlsx/table-writer.ts src/xlsx/table-reader.ts src/xlsx/content-types.ts src/xlsx/worksheet-writer.ts src/xlsx/worksheet-reader.ts src/xlsx/write.ts src/xlsx/read.ts src/xlsx/tables.test.ts
git commit -m "feat(xlsx): tables (xl/tables/tableN.xml + tableParts) write+read"
```

---

### Task 4: Feature-fixture conformance + size check

**Files:**
- Modify: `src/xlsx/read-conformance.test.ts` (assert feature-bearing fixtures still parse without throwing), `size-budget.json` (if needed)

**Interfaces:** none new.

The 35 fixtures include files with defined names, data validations, and tables. After Tasks 1-3 the reader recognizes these elements; this task confirms the whole fixture corpus still parses cleanly and the bundle stayed within budget.

- [ ] **Step 1: Append a fixture assertion that feature elements don't break parsing**

```typescript
// append to src/xlsx/read-conformance.test.ts
test('fixtures using defined names / tables / validations still parse', async () => {
  for (const path of listFixtures()) {
    const name = path.split('/').pop()!
    if (name === 'missing-bits.xlsx') continue
    const wb = await readXlsx(await parseFixture(path))
    // any fixture that has defined names should surface them as a (possibly empty) array
    expect(Array.isArray(wb.definedNames)).toBe(true)
  }
})
```

- [ ] **Step 2: Run the full suite + size**

Run: `npx vitest run && npm run -s build && npm run -s size`
Expected: all green; `index.js` within budget. If the three features pushed it over, raise the budget to the measured gzip + ~15% and note it.

- [ ] **Step 3: Run all gates and commit**

Run: `npx vitest run && npm run -s typecheck && npm run -s lint && npm run -s format && npm run -s build && npm run -s size`

```bash
git add src/xlsx/read-conformance.test.ts size-budget.json
git commit -m "test(xlsx): feature-fixture conformance for SP-5 modules"
```

---

## SP-5 Acceptance Criteria

1. `Workbook.defineName` + `definedNames` write and read round-trip via `xl/workbook.xml`.
2. `Worksheet.addDataValidation` + `dataValidations` write and read round-trip (list + numeric-operator forms), with the inverted `showDropDown` flag handled correctly.
3. `Worksheet.addTable` + `tables` write and read round-trip via a new `xl/tables/tableN.xml` part with correct rels + content-type; exceljs reads the table back.
4. All 35 fixtures still parse without throwing; bundle within budget; all gates green (typecheck, lint 0/0, format, test, build, size); zero runtime deps.

## Self-Review

**1. Spec coverage** (architecture §SP-5 "Self-contained serialize+parse slices"): defined names → Task 1; data validation → Task 2; tables → Task 3; conformance → Task 4. Deferred slices (images/drawings, conditional formatting, sheet-scoped names) explicitly listed in Scope OUT for a later plan. ✓

**2. Placeholder scan:** Tasks 1-2 carry full code. Task 3's writer/reader files are complete; the `write.ts`/`read.ts`/`worksheet-*` wiring for tables (Steps 5-7) is described in prose + fragments because it threads through several existing files that grew across SP-3/SP-4 — the implementer assembles from the fragments, reusing the hyperlink-rel pattern verbatim. This is the one place the executor must integrate rather than paste; flagged here intentionally. No "TBD"/"handle errors" placeholders.

**3. Type consistency:** `DataValidation` (model/data-validation.ts) and `TableDefinition` (model/table.ts) are defined once and referenced by model + writer + reader. `WorksheetWriteResult` gains `tables` (Task 3) alongside the existing `hyperlinks` (SP-3) — the writer returns both; `write.ts` consumes both. `WorkbookParts` gains `definedNames` (Task 1). `defineName`/`addDataValidation`/`addTable`/`writeTableXml`/`readTableXml`/`CT.table` names are stable across tasks. The per-sheet rId numbering (hyperlinks `rId1..N`, tables `rId(N+1)..`) is called out in Task 3 Steps 5-6 to keep `<tablePart r:id>` and the `.rels` entries aligned.

**Risk note:** Task 3 (tables) is materially larger than 1-2 because it adds a new OPC part and merges per-sheet rels. If a reviewer rejects it, Tasks 1, 2, and 4 still ship independently — defined names and data validation have no dependency on tables.
