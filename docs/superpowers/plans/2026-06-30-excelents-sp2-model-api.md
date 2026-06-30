# SP-2 — Object model + public API Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Implement the in-memory object model (Workbook / Worksheet / Row / Column / Cell / styles + cell value types) and the redesigned TS-first public API (`createWorkbook` + types), with NO serialization (codecs are SP-3/SP-4).

**Architecture:** A pure, web-standard model under `src/model/`. Lightweight construction via `createWorkbook()`; cells mutated by property setters (`ws.cell('A1').value = 'Hello'`). Serialization is deliberately absent — it lands later as free functions (`writeXlsx`/`readXlsx`) so an app that only writes never pulls the reader. The model consumes only `src/utils/` (address/range); it stays `node:`-free and dependency-free.

**Tech Stack:** TypeScript 7, Vitest 4, oxlint type-aware, oxfmt. Pure JS, zero deps.

## Global Constraints

- **Web-standard core only:** `src/model/*.ts` must import NO `node:` builtin (core-purity guard).
- **Zero runtime dependencies.** No third-party imports in `src/`.
- **`isolatedDeclarations`:** explicit return types on every exported binding (incl. class getters/methods).
- **Named exports only**; no default export; no barrel file inside `model/` (import modules directly), but `src/index.ts` IS the public entry and re-exports the public surface.
- **`sideEffects: false`** stays true.
- All gates green: `pnpm typecheck`, `pnpm lint` (0/0), `pnpm format`, `pnpm test`, `pnpm build`, `pnpm size`.
- TDD: failing test first (RED), minimal implementation (GREEN), commit. Report RED/GREEN evidence.
- **Scope IN:** core model + base styling (font/fill/border/alignment/number-format) + rich-text & hyperlink cell values. **Scope OUT (deferred to SP-5):** images, data validation, tables, conditional formatting, defined names — their model lands with their serialization.
- Public API shape (confirmed): `createWorkbook()`, `wb.addSheet(name)`, `ws.cell('A1').value = ...`, property-setter mutation. Serialization free functions come in SP-3/SP-4.
- Branch: `excelents-rewrite` (continues from SP-1).

---

### Task 1: Styling types (`src/model/style.ts`)

**Files:**
- Create: `src/model/style.ts`
- Test: `src/model/style.test.ts`

**Interfaces:**
- Produces: `Font`, `Fill`, `BorderEdge`, `Borders`, `Alignment`, `CellStyle` interfaces, and `mergeStyles(base: CellStyle, patch: CellStyle): CellStyle` (shallow-merges top-level style facets; `patch` wins).

- [ ] **Step 1: Write the failing test** (`src/model/style.test.ts`)

```ts
import { expect, test } from 'vitest'
import { type CellStyle, mergeStyles } from './style'

test('mergeStyles overlays patch facets over base', () => {
  const base: CellStyle = { font: { bold: true }, numberFormat: '0.00' }
  const patch: CellStyle = { font: { italic: true }, fill: { type: 'pattern', pattern: 'solid', fgColor: 'FFFF0000' } }
  const merged = mergeStyles(base, patch)
  // patch's font replaces base's font (facet-level merge, not deep)
  expect(merged.font).toEqual({ italic: true })
  expect(merged.fill).toEqual({ type: 'pattern', pattern: 'solid', fgColor: 'FFFF0000' })
  expect(merged.numberFormat).toBe('0.00')
})

test('mergeStyles does not mutate its inputs', () => {
  const base: CellStyle = { font: { bold: true } }
  mergeStyles(base, { numberFormat: '0' })
  expect(base).toEqual({ font: { bold: true } })
})

test('mergeStyles of empties is empty', () => {
  expect(mergeStyles({}, {})).toEqual({})
})
```

- [ ] **Step 2: Run the test to verify it fails**

Run: `pnpm exec vitest run src/model/style.test.ts`
Expected: FAIL — cannot resolve `./style`.

- [ ] **Step 3: Write `src/model/style.ts`**

```ts
/** Font facet. `color` is ARGB hex, e.g. `'FF0000FF'`. */
export interface Font {
  name?: string
  size?: number
  bold?: boolean
  italic?: boolean
  underline?: boolean
  color?: string
}

/** Fill facet (pattern fills only for now). */
export interface Fill {
  type: 'pattern'
  pattern: 'solid' | 'none'
  fgColor?: string
}

export interface BorderEdge {
  style?: 'thin' | 'medium' | 'thick' | 'dashed' | 'dotted' | 'double'
  color?: string
}

export interface Borders {
  top?: BorderEdge
  bottom?: BorderEdge
  left?: BorderEdge
  right?: BorderEdge
}

export interface Alignment {
  horizontal?: 'left' | 'center' | 'right'
  vertical?: 'top' | 'middle' | 'bottom'
  wrapText?: boolean
}

/** A cell's visual style. All facets optional; `numberFormat` is a format code. */
export interface CellStyle {
  font?: Font
  fill?: Fill
  border?: Borders
  alignment?: Alignment
  numberFormat?: string
}

/** Overlay `patch`'s facets over `base` (facet-level, not deep). Inputs are not mutated. */
export function mergeStyles(base: CellStyle, patch: CellStyle): CellStyle {
  return { ...base, ...patch }
}
```

- [ ] **Step 4: Run the test to verify it passes**

Run: `pnpm exec vitest run src/model/style.test.ts`
Expected: PASS (3 tests).

- [ ] **Step 5: Commit**

```bash
git add src/model/style.ts src/model/style.test.ts
git commit -m "feat(model): cell styling types + mergeStyles"
```

---

### Task 2: Cell value types + Cell (`src/model/cell.ts`)

**Files:**
- Create: `src/model/cell.ts`
- Test: `src/model/cell.test.ts`

**Interfaces:**
- Consumes: `CellStyle`, `Font` from `./style`; `encodeAddress` from `../utils/address`.
- Produces:
  - `interface RichTextRun { text: string; font?: Font }`
  - `interface FormulaValue { formula: string; result?: string | number | boolean | Date | null }`
  - `interface RichTextValue { richText: RichTextRun[] }`
  - `interface HyperlinkValue { text: string; hyperlink: string }`
  - `type CellValue = null | string | number | boolean | Date | FormulaValue | RichTextValue | HyperlinkValue`
  - `type CellType = 'null' | 'string' | 'number' | 'boolean' | 'date' | 'formula' | 'richText' | 'hyperlink'`
  - `class Cell` with `readonly row`, `readonly col`, `value: CellValue` (default null), `style: CellStyle` (default {}), `get address(): string`, `get type(): CellType`.

- [ ] **Step 1: Write the failing test** (`src/model/cell.test.ts`)

```ts
import { expect, test } from 'vitest'
import { Cell } from './cell'

test('a new cell is null at its address', () => {
  const c = new Cell(1, 1)
  expect(c.address).toBe('A1')
  expect(c.value).toBeNull()
  expect(c.type).toBe('null')
})

test('detects scalar types from the value', () => {
  const c = new Cell(2, 3) // C2
  expect(c.address).toBe('C2')
  c.value = 'hello'
  expect(c.type).toBe('string')
  c.value = 42
  expect(c.type).toBe('number')
  c.value = true
  expect(c.type).toBe('boolean')
  c.value = new Date(Date.UTC(2024, 0, 1))
  expect(c.type).toBe('date')
})

test('detects formula, richText, and hyperlink values', () => {
  const c = new Cell(1, 1)
  c.value = { formula: 'A1+B1', result: 7 }
  expect(c.type).toBe('formula')
  c.value = { richText: [{ text: 'a' }, { text: 'b', font: { bold: true } }] }
  expect(c.type).toBe('richText')
  c.value = { text: 'site', hyperlink: 'https://example.com' }
  expect(c.type).toBe('hyperlink')
})

test('carries a mutable style', () => {
  const c = new Cell(1, 1)
  c.style = { font: { bold: true }, numberFormat: '0.00' }
  expect(c.style.numberFormat).toBe('0.00')
})
```

- [ ] **Step 2: Run the test to verify it fails**

Run: `pnpm exec vitest run src/model/cell.test.ts`
Expected: FAIL — cannot resolve `./cell`.

- [ ] **Step 3: Write `src/model/cell.ts`**

```ts
import { encodeAddress } from '../utils/address'
import type { CellStyle, Font } from './style'

export interface RichTextRun {
  text: string
  font?: Font
}

export interface FormulaValue {
  formula: string
  result?: string | number | boolean | Date | null
}

export interface RichTextValue {
  richText: RichTextRun[]
}

export interface HyperlinkValue {
  text: string
  hyperlink: string
}

export type CellValue =
  | null
  | string
  | number
  | boolean
  | Date
  | FormulaValue
  | RichTextValue
  | HyperlinkValue

export type CellType =
  | 'null'
  | 'string'
  | 'number'
  | 'boolean'
  | 'date'
  | 'formula'
  | 'richText'
  | 'hyperlink'

function detectType(value: CellValue): CellType {
  if (value === null) return 'null'
  if (typeof value === 'string') return 'string'
  if (typeof value === 'number') return 'number'
  if (typeof value === 'boolean') return 'boolean'
  if (value instanceof Date) return 'date'
  if ('formula' in value) return 'formula'
  if ('richText' in value) return 'richText'
  return 'hyperlink'
}

/** A single spreadsheet cell: a typed value plus an optional style. */
export class Cell {
  value: CellValue = null
  style: CellStyle = {}

  constructor(
    readonly row: number,
    readonly col: number,
  ) {}

  get address(): string {
    return encodeAddress(this.row, this.col)
  }

  get type(): CellType {
    return detectType(this.value)
  }
}
```

- [ ] **Step 4: Run the test to verify it passes**

Run: `pnpm exec vitest run src/model/cell.test.ts`
Expected: PASS (4 tests).

- [ ] **Step 5: Commit**

```bash
git add src/model/cell.ts src/model/cell.test.ts
git commit -m "feat(model): cell value types + Cell"
```

---

### Task 3: Row & Column (`src/model/row.ts`, `src/model/column.ts`)

**Files:**
- Create: `src/model/row.ts`, `src/model/column.ts`
- Test: `src/model/row.test.ts`, `src/model/column.test.ts`

**Interfaces:**
- Consumes: `Cell` from `./cell`; `numberToCol` from `../utils/address`.
- Produces:
  - `class Row` with `readonly number`, `height?: number`, `getCell(col: number): Cell` (lazy-creates and caches), `get cells(): Cell[]` (populated cells, ascending column).
  - `class Column` with `readonly number`, `key?: string`, `width?: number`, `get letter(): string`.

- [ ] **Step 1: Write the failing tests**

`src/model/row.test.ts`:
```ts
import { expect, test } from 'vitest'
import { Row } from './row'

test('getCell lazily creates and caches a cell at (row, col)', () => {
  const row = new Row(5)
  const a = row.getCell(1)
  expect(a.address).toBe('A5')
  expect(row.getCell(1)).toBe(a) // cached (same instance)
})

test('cells lists populated cells in ascending column order', () => {
  const row = new Row(1)
  row.getCell(3).value = 'c'
  row.getCell(1).value = 'a'
  expect(row.cells.map((c) => c.address)).toEqual(['A1', 'C1'])
})

test('carries an optional height', () => {
  const row = new Row(2)
  row.height = 18
  expect(row.height).toBe(18)
})
```

`src/model/column.test.ts`:
```ts
import { expect, test } from 'vitest'
import { Column } from './column'

test('exposes its letter, key, and width', () => {
  const col = new Column(27)
  expect(col.letter).toBe('AA')
  col.key = 'id'
  col.width = 12
  expect(col.key).toBe('id')
  expect(col.width).toBe(12)
})
```

- [ ] **Step 2: Run the tests to verify they fail**

Run: `pnpm exec vitest run src/model/row.test.ts src/model/column.test.ts`
Expected: FAIL — cannot resolve `./row` / `./column`.

- [ ] **Step 3: Write `src/model/column.ts`**

```ts
import { numberToCol } from '../utils/address'

/** Column metadata (1-based). */
export class Column {
  key?: string
  width?: number

  constructor(readonly number: number) {}

  get letter(): string {
    return numberToCol(this.number)
  }
}
```

- [ ] **Step 4: Write `src/model/row.ts`**

```ts
import { Cell } from './cell'

/** A worksheet row (1-based) owning its cells sparsely. */
export class Row {
  height?: number
  private readonly cellsByCol = new Map<number, Cell>()

  constructor(readonly number: number) {}

  getCell(col: number): Cell {
    let cell = this.cellsByCol.get(col)
    if (cell === undefined) {
      cell = new Cell(this.number, col)
      this.cellsByCol.set(col, cell)
    }
    return cell
  }

  get cells(): Cell[] {
    return [...this.cellsByCol.keys()].sort((a, b) => a - b).map((col) => this.cellsByCol.get(col)!)
  }
}
```

- [ ] **Step 5: Run the tests to verify they pass**

Run: `pnpm exec vitest run src/model/row.test.ts src/model/column.test.ts`
Expected: PASS (4 tests). If lint flags `this.cellsByCol.get(col)!` in `cells`, restructure to map entries directly (`[...this.cellsByCol.entries()].sort(...).map(([, c]) => c)`) to avoid the assertion.

- [ ] **Step 6: Commit**

```bash
git add src/model/row.ts src/model/row.test.ts src/model/column.ts src/model/column.test.ts
git commit -m "feat(model): Row (owns cells) + Column"
```

---

### Task 4: Worksheet (`src/model/worksheet.ts`)

**Files:**
- Create: `src/model/worksheet.ts`
- Test: `src/model/worksheet.test.ts`

**Interfaces:**
- Consumes: `Row` from `./row`; `Column` from `./column`; `Cell` from `./cell`; `CellValue` from `./cell`; `decodeAddress` from `../utils/address`; `decodeRange`, `encodeRange`, `type RangeBox` from `../utils/range`.
- Produces: `class Worksheet` with `name: string`, `getRow(n): Row` (lazy), `cell(ref: string): Cell`, `getCell(row, col): Cell`, `addRow(values: CellValue[]): Row` (appends after the last populated row), `column(n: number): Column` (lazy), `merge(range: string): void`, `get merges(): readonly string[]`, `get rowCount(): number`, `get dimensions(): RangeBox | undefined`.

- [ ] **Step 1: Write the failing test** (`src/model/worksheet.test.ts`)

```ts
import { expect, test } from 'vitest'
import { Worksheet } from './worksheet'

test('cell(ref) addresses the same cell as getCell(row, col)', () => {
  const ws = new Worksheet('Sheet1')
  ws.cell('B2').value = 'hi'
  expect(ws.getCell(2, 2).value).toBe('hi')
  expect(ws.cell('B2')).toBe(ws.getCell(2, 2))
})

test('addRow appends after the last populated row', () => {
  const ws = new Worksheet('s')
  ws.cell('A1').value = 'header'
  const row = ws.addRow(['a', 1, true])
  expect(row.number).toBe(2)
  expect(ws.cell('A2').value).toBe('a')
  expect(ws.cell('B2').value).toBe(1)
  expect(ws.cell('C2').value).toBe(true)
})

test('merge records ranges', () => {
  const ws = new Worksheet('s')
  ws.merge('A1:B2')
  expect(ws.merges).toEqual(['A1:B2'])
})

test('dimensions spans the populated cells', () => {
  const ws = new Worksheet('s')
  ws.cell('B2').value = 'x'
  ws.cell('D5').value = 'y'
  expect(ws.dimensions).toEqual({ top: 2, left: 2, bottom: 5, right: 4 })
})

test('dimensions is undefined when empty', () => {
  expect(new Worksheet('s').dimensions).toBeUndefined()
})
```

- [ ] **Step 2: Run the test to verify it fails**

Run: `pnpm exec vitest run src/model/worksheet.test.ts`
Expected: FAIL — cannot resolve `./worksheet`.

- [ ] **Step 3: Write `src/model/worksheet.ts`**

```ts
import { decodeAddress } from '../utils/address'
import type { RangeBox } from '../utils/range'
import type { Cell, CellValue } from './cell'
import { Column } from './column'
import { Row } from './row'

/** A worksheet: a sparse grid of cells, plus columns, merges, and a name. */
export class Worksheet {
  private readonly rows = new Map<number, Row>()
  private readonly cols = new Map<number, Column>()
  private readonly mergeRanges: string[] = []

  constructor(public name: string) {}

  getRow(n: number): Row {
    let row = this.rows.get(n)
    if (row === undefined) {
      row = new Row(n)
      this.rows.set(n, row)
    }
    return row
  }

  getCell(row: number, col: number): Cell {
    return this.getRow(row).getCell(col)
  }

  cell(ref: string): Cell {
    const { row, col } = decodeAddress(ref)
    return this.getCell(row, col)
  }

  addRow(values: CellValue[]): Row {
    const row = this.getRow(this.rowCount + 1)
    values.forEach((value, i) => {
      row.getCell(i + 1).value = value
    })
    return row
  }

  column(n: number): Column {
    let col = this.cols.get(n)
    if (col === undefined) {
      col = new Column(n)
      this.cols.set(n, col)
    }
    return col
  }

  merge(range: string): void {
    this.mergeRanges.push(range)
  }

  get merges(): readonly string[] {
    return this.mergeRanges
  }

  get rowCount(): number {
    let max = 0
    for (const n of this.rows.keys()) {
      if (n > max) max = n
    }
    return max
  }

  get dimensions(): RangeBox | undefined {
    let top = Infinity
    let left = Infinity
    let bottom = 0
    let right = 0
    let any = false
    for (const row of this.rows.values()) {
      for (const cell of row.cells) {
        any = true
        if (cell.row < top) top = cell.row
        if (cell.row > bottom) bottom = cell.row
        if (cell.col < left) left = cell.col
        if (cell.col > right) right = cell.col
      }
    }
    return any ? { top, left, bottom, right } : undefined
  }
}
```

- [ ] **Step 4: Run the test to verify it passes**

Run: `pnpm exec vitest run src/model/worksheet.test.ts`
Expected: PASS (5 tests). Note `rowCount` counts the max populated row number; `addRow` after a single `A1` write lands on row 2 — matching the test.

- [ ] **Step 5: Commit**

```bash
git add src/model/worksheet.ts src/model/worksheet.test.ts
git commit -m "feat(model): Worksheet (sparse cells, addRow, merges, dimensions)"
```

---

### Task 5: Workbook + public API (`src/model/workbook.ts`, `src/index.ts`)

**Files:**
- Create: `src/model/workbook.ts`, `src/model/workbook.test.ts`
- Modify: `src/index.ts` (export the public surface), `src/index.test.ts` (update the version-stub test to the real API), `size-budget.json` (raise the `index.js` budget for the now-real model)

**Interfaces:**
- Consumes: `Worksheet` from `./worksheet`.
- Produces:
  - `class Workbook` with `addSheet(name: string): Worksheet`, `getSheet(name: string): Worksheet | undefined`, `removeSheet(name: string): void`, `get sheets(): readonly Worksheet[]`.
  - `createWorkbook(): Workbook` factory.
  - `src/index.ts` public exports: `createWorkbook`, and the public types/classes (`Workbook`, `Worksheet`, `Row`, `Column`, `Cell`, `CellValue`, `CellType`, `CellStyle`, `Font`, `Fill`, `Borders`, `BorderEdge`, `Alignment`, `RichTextRun`, `FormulaValue`, `RichTextValue`, `HyperlinkValue`).

- [ ] **Step 1: Write the failing test** (`src/model/workbook.test.ts`)

```ts
import { expect, test } from 'vitest'
import { createWorkbook } from './workbook'

test('createWorkbook starts with no sheets', () => {
  expect(createWorkbook().sheets).toEqual([])
})

test('addSheet adds a named sheet and getSheet finds it', () => {
  const wb = createWorkbook()
  const ws = wb.addSheet('Data')
  expect(ws.name).toBe('Data')
  expect(wb.getSheet('Data')).toBe(ws)
  expect(wb.sheets.map((s) => s.name)).toEqual(['Data'])
})

test('removeSheet removes by name', () => {
  const wb = createWorkbook()
  wb.addSheet('A')
  wb.addSheet('B')
  wb.removeSheet('A')
  expect(wb.sheets.map((s) => s.name)).toEqual(['B'])
})

test('end-to-end: build a workbook and read values back', () => {
  const wb = createWorkbook()
  const ws = wb.addSheet('Sheet1')
  ws.cell('A1').value = 'Hello'
  ws.addRow(['world', 42])
  expect(ws.cell('A1').value).toBe('Hello')
  expect(ws.cell('A2').value).toBe('world')
  expect(ws.cell('B2').value).toBe(42)
})
```

- [ ] **Step 2: Run the test to verify it fails**

Run: `pnpm exec vitest run src/model/workbook.test.ts`
Expected: FAIL — cannot resolve `./workbook`.

- [ ] **Step 3: Write `src/model/workbook.ts`**

```ts
import { Worksheet } from './worksheet'

/** A workbook: an ordered collection of named worksheets. */
export class Workbook {
  private readonly worksheets: Worksheet[] = []

  addSheet(name: string): Worksheet {
    const ws = new Worksheet(name)
    this.worksheets.push(ws)
    return ws
  }

  getSheet(name: string): Worksheet | undefined {
    return this.worksheets.find((ws) => ws.name === name)
  }

  removeSheet(name: string): void {
    const i = this.worksheets.findIndex((ws) => ws.name === name)
    if (i !== -1) this.worksheets.splice(i, 1)
  }

  get sheets(): readonly Worksheet[] {
    return this.worksheets
  }
}

/** Create an empty workbook. */
export function createWorkbook(): Workbook {
  return new Workbook()
}
```

- [ ] **Step 4: Run the test to verify it passes**

Run: `pnpm exec vitest run src/model/workbook.test.ts`
Expected: PASS (4 tests).

- [ ] **Step 5: Wire the public API in `src/index.ts`** (replace the version stub)

```ts
export { createWorkbook, Workbook } from './model/workbook'
export { Worksheet } from './model/worksheet'
export { Row } from './model/row'
export { Column } from './model/column'
export { Cell } from './model/cell'
export type {
  CellValue,
  CellType,
  RichTextRun,
  FormulaValue,
  RichTextValue,
  HyperlinkValue,
} from './model/cell'
export type {
  CellStyle,
  Font,
  Fill,
  Borders,
  BorderEdge,
  Alignment,
} from './model/style'
```

- [ ] **Step 6: Update `src/index.test.ts`** to test the real API instead of the version stub

```ts
import { expect, test } from 'vitest'
import { createWorkbook } from './index'

test('the package exposes createWorkbook building a usable workbook', () => {
  const wb = createWorkbook()
  const ws = wb.addSheet('Sheet1')
  ws.cell('A1').value = 'Hello'
  expect(ws.cell('A1').value).toBe('Hello')
  expect(wb.sheets.map((s) => s.name)).toEqual(['Sheet1'])
})
```

- [ ] **Step 7: Raise the `index.js` size budget**

Build first to measure: `pnpm build && pnpm size`. The `.` entry now contains the real model. In `size-budget.json`, set `index.js` to a value comfortably above the measured gzip size (round up to the next ~2 KB; e.g. if it measures 3.1 KB, set 6144). Leave `csv.js`/`node.js` unchanged. Note the measured size in your report.

- [ ] **Step 8: Final full verification + commit**

Run: `pnpm typecheck && pnpm lint && pnpm format && pnpm test && pnpm build && pnpm size`
Expected: all green; the public `.` entry exports `createWorkbook` (verified by the updated `index.test.ts`); `pnpm size` passes against the raised budget.

```bash
git add src/model/workbook.ts src/model/workbook.test.ts src/index.ts src/index.test.ts size-budget.json
git commit -m "feat(model): Workbook + public API (createWorkbook); wire src/index.ts"
```

---

## SP-2 Acceptance Criteria

1. `src/model/` contains `style.ts`, `cell.ts`, `row.ts`, `column.ts`, `worksheet.ts`, `workbook.ts`, each with colocated tests.
2. All model modules are `node:`-free and dependency-free (core-purity guard passes).
3. `src/index.ts` exports `createWorkbook` + the public model classes/types; the end-to-end test builds a workbook and reads values back.
4. No serialization exists yet (deferred to SP-3/SP-4) — the model is pure in-memory.
5. All gates green: typecheck, lint (0/0), format, test, build, size (with the raised `index.js` budget).

---

## Self-Review

**1. Spec coverage** (SP-1 foundation spec "Unit: Object model + public API", and the architecture's API direction):
- In-memory model (Workbook/Worksheet/Row/Column/Cell/Range/styles) → Tasks 1-5. ✓ (Range reuses `utils/range`'s `RangeBox` rather than a new class — YAGNI.)
- Cell value types incl. rich text + hyperlinks (in-scope advanced) → Task 2. ✓
- Base styling (font/fill/border/alignment/number-format) → Task 1. ✓
- Redesigned public API: `createWorkbook` + property-setter mutation → Tasks 5 + 2. ✓
- No serialization (free functions deferred) → enforced by scope. ✓
- Out-of-scope advanced features (images/validation/tables/CF/defined-names) deferred to SP-5. ✓

**2. Placeholder scan:** No TBD/"add later". Every step has complete code, exact commands, expected results.

**3. Type consistency:** `CellStyle`/`Font` (Task 1) consumed by `Cell` (Task 2). `Cell`/`CellValue` (Task 2) consumed by `Row` (Task 3) and `Worksheet` (Task 4). `Row`/`Column` (Task 3) consumed by `Worksheet` (Task 4). `Worksheet` (Task 4) consumed by `Workbook` (Task 5). `RangeBox` (from `utils/range`) used in `Worksheet.dimensions`. The `src/index.ts` exports (Task 5) name exactly the types defined in Tasks 1-2 and classes in Tasks 2-5. `createWorkbook`/`addSheet`/`cell().value=` match the confirmed public API.

Note for the executor: this is the first task that makes `src/index.ts` non-trivial, so the `.` bundle grows — the size-budget bump in Task 5 Step 7 is expected, not a regression. Keep the model `node:`-free (the core-purity guard covers `src/model/`).
