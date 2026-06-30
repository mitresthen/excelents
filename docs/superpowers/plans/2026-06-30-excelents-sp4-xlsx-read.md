# SP-4: xlsx Read (baseline) Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Parse a `.xlsx` byte array into the object model via a `readXlsx(bytes)` free function — the inverse of SP-3's `writeXlsx` — recovering values, styling, shared strings, merges, and sizing, validated by round-tripping SP-3 output and by parsing the 35 real fixtures against the exceljs oracle.

**Architecture:** A thin orchestrator `readXlsx(bytes): Promise<Workbook>` opens the OPC package (`OpcPackage.read`), resolves relationships to locate the workbook/worksheet/sharedStrings/styles parts, and delegates to focused part readers (`workbook-reader`, `shared-strings-reader`, `styles-reader`, `worksheet-reader`) that tokenize each XML part and build up the model with the SP-2 model API. Each reader is the structural inverse of its SP-3 writer counterpart.

**Tech Stack:** TypeScript 7 (isolatedDeclarations), the owned XML `tokenize`, the owned `OpcPackage`/`readZip`, the SP-2 model (`createWorkbook`/`Worksheet`/`Cell`), Vitest 4 + `exceljs@4.4.0` oracle. Zero runtime dependencies; `node:`-free in `src/`.

## Global Constraints

- **Zero runtime dependencies.** No new entries in `package.json#dependencies`. (copied from architecture spec)
- **`node:`-free core.** No `node:` imports anywhere in `src/` except `src/node.ts` (enforced by `test/core-purity.test.ts`). Readers use `TextDecoder`, the owned tokenizer, and the model only.
- **`isolatedDeclarations: true`.** Every exported function/const needs an explicit return type.
- **Web-standard only.** No Node builtins; tokenize + model + OPC are already web-standard.
- **TDD, frequent commits.** Red → green → refactor per behavior; commit when a value form round-trips.
- **Validation strategy:** two oracles. (a) **Round-trip:** `readXlsx(await writeXlsx(wb))` recovers the model we wrote. (b) **Fixture parity:** parse a real `.xlsx` fixture with both exceljs and `readXlsx` and compare recovered cell values. Round-trip is primary for new value forms; fixture parity is the realism gate in Task 5.
- **Scope IN:** values (string/number/boolean/date/formula/hyperlink/richText), base styling (font/fill/border/alignment/number-format), shared strings, merges, dimensions, column widths/row heights. **Scope OUT:** images, data validation, tables, conditional formatting, defined names (SP-5); streaming (SP-7); encryption (later). These are skipped gracefully (unknown parts/elements ignored, never throw).

---

## File Structure

- `src/xlsx/read.ts` — `readXlsx(bytes): Promise<Workbook>` orchestrator (inverse of `write.ts`).
- `src/xlsx/workbook-reader.ts` — `readWorkbookParts(pkg): WorkbookParts` — locate workbook part via root rels, parse `xl/workbook.xml` sheet list + `xl/_rels/workbook.xml.rels` to resolve sheet/sharedStrings/styles targets.
- `src/xlsx/shared-strings-reader.ts` — `readSharedStrings(xml): SharedStringValue[]` (inverse of `shared-strings-writer.ts`).
- `src/xlsx/styles-reader.ts` — `readStyles(xml): ParsedStyles` — `{ cellStyles: CellStyle[]; numFmtById: Map<number,string> }` (inverse of `styles-writer.ts` / `StyleRegistry`).
- `src/xlsx/worksheet-reader.ts` — `readWorksheetInto(ws, xml, ctx)` — populate a `Worksheet` from one `sheetN.xml` (inverse of `worksheet-writer.ts`).
- `src/opc/package.ts` — **modify**: add `relationshipsFor(partName): OpcRelationship[]` (part-scoped rels + relative-target resolution; flagged in the SP-1 forward-notes).
- `src/utils/number-format.ts` — **modify**: add `isDateFormat(code): boolean` (decide whether a numeric cell's format means it is a date).
- `src/index.ts` — **modify**: `export { readXlsx } from './xlsx/read'`.
- Tests: `src/opc/relationships.test.ts`, `src/xlsx/read.test.ts`, `src/xlsx/shared-strings-reader.test.ts`, `src/xlsx/styles-reader.test.ts`, `src/xlsx/read-values.test.ts`, `src/xlsx/read-conformance.test.ts`, plus additions to `src/index.test.ts` and `src/utils/number-format.test.ts`.

**Shared types** (declared in the files that own them; consumed across tasks):

```typescript
// src/xlsx/shared-strings-reader.ts
import type { RichTextRun } from '../model/cell'
export type SharedStringValue =
  | { readonly kind: 'plain'; readonly text: string }
  | { readonly kind: 'rich'; readonly runs: RichTextRun[] }

// src/xlsx/workbook-reader.ts
export interface WorkbookParts {
  /** Worksheet parts in workbook order: display name + absolute part path. */
  readonly sheets: ReadonlyArray<{ name: string; path: string }>
  /** Absolute path to sharedStrings.xml, or undefined if absent. */
  readonly sharedStringsPath: string | undefined
  /** Absolute path to styles.xml, or undefined if absent. */
  readonly stylesPath: string | undefined
}

// src/xlsx/styles-reader.ts
export interface ParsedStyles {
  /** CellStyle for each cellXfs index. Index 0 is the default empty style. */
  readonly cellStyles: CellStyle[]
  /** numFmtId -> format code, for every xf's numFmtId (builtin + custom). */
  readonly numFmtById: Map<number, string>
}
```

---

### Task 1: OPC part-scoped relationships + `readXlsx` skeleton (sheet list)

**Files:**
- Modify: `src/opc/package.ts` (add `relationshipsFor`)
- Create: `src/xlsx/workbook-reader.ts`
- Create: `src/xlsx/read.ts`
- Test: `src/opc/relationships.test.ts`, `src/xlsx/read.test.ts`

**Interfaces:**
- Consumes: `OpcPackage.read(bytes)`, `OpcPackage.getPart(name)`, `OpcPackage.rootRelationships()`, `OpcRelationship { id, type, target }`, `tokenize`, `createWorkbook()`, `Workbook.addSheet(name)`, `writeXlsx` (test only).
- Produces: `OpcPackage.relationshipsFor(partName: string): OpcRelationship[]`; `readWorkbookParts(pkg: OpcPackage): WorkbookParts`; `readXlsx(bytes: Uint8Array): Promise<Workbook>`.

The OPC tweak: `rootRelationships()` parses only `_rels/.rels` with raw targets. `relationshipsFor('xl/workbook.xml')` must read `xl/_rels/workbook.xml.rels` and resolve each `target` **relative to the part's directory** (`xl/`), yielding absolute part paths (`xl/worksheets/sheet1.xml`). Targets may be already-absolute (`/xl/...`) — strip the leading slash. External targets (`TargetMode="External"`, e.g. hyperlinks) are returned verbatim; callers decide.

- [ ] **Step 1: Write the failing test for part-scoped relationship resolution**

```typescript
// src/opc/relationships.test.ts
import { expect, test } from 'vitest'
import { OpcPackage } from './package'

const enc = (s: string): Uint8Array => new TextEncoder().encode(s)

const REL = (id: string, type: string, target: string, mode?: string): string =>
  `<Relationship Id="${id}" Type="${type}" Target="${target}"${mode ? ` TargetMode="${mode}"` : ''}/>`

async function pkgWith(parts: Record<string, string>): Promise<OpcPackage> {
  const pkg = OpcPackage.empty()
  for (const [name, body] of Object.entries(parts)) pkg.setPart(name, 'application/xml', enc(body))
  // round-trip through bytes so we exercise the real read path
  return OpcPackage.read(await pkg.toBytes())
}

test('relationshipsFor resolves targets relative to the part directory', async () => {
  const pkg = await pkgWith({
    'xl/workbook.xml': '<workbook/>',
    'xl/_rels/workbook.xml.rels':
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      REL('rId1', 'http://x/worksheet', 'worksheets/sheet1.xml') +
      REL('rId2', 'http://x/styles', '/xl/styles.xml') +
      REL('rId3', 'http://x/hyperlink', 'https://example.com/', 'External') +
      `</Relationships>`,
  })
  const rels = pkg.relationshipsFor('xl/workbook.xml')
  expect(rels.find((r) => r.id === 'rId1')?.target).toBe('xl/worksheets/sheet1.xml')
  expect(rels.find((r) => r.id === 'rId2')?.target).toBe('xl/styles.xml')
  // External targets are returned verbatim.
  expect(rels.find((r) => r.id === 'rId3')?.target).toBe('https://example.com/')
})

test('relationshipsFor returns [] when the part has no rels', async () => {
  const pkg = await pkgWith({ 'xl/workbook.xml': '<workbook/>' })
  expect(pkg.relationshipsFor('xl/workbook.xml')).toEqual([])
})
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx vitest run src/opc/relationships.test.ts`
Expected: FAIL — `relationshipsFor` is not a function.

- [ ] **Step 3: Implement `relationshipsFor` + a `TargetMode` field**

In `src/opc/package.ts`, extend `OpcRelationship` with an optional `targetMode` and add the method. Add a helper that resolves a relative target against a part's directory.

```typescript
// add to OpcRelationship
export interface OpcRelationship {
  id: string
  type: string
  target: string
  targetMode?: string
}

// add inside class OpcPackage
/** Directory of a part path (`xl/worksheets/sheet1.xml` -> `xl/worksheets`), '' for root. */
private static dirOf(partName: string): string {
  const slash = partName.lastIndexOf('/')
  return slash === -1 ? '' : partName.slice(0, slash)
}

/** Resolve a relationship Target against a base directory into an absolute part path. */
private static resolveTarget(baseDir: string, target: string): string {
  if (target.startsWith('/')) return target.slice(1) // already package-absolute
  if (baseDir === '') return target
  return `${baseDir}/${target}`
}

/** Relationships declared for a specific part (`<dir>/_rels/<file>.rels`), targets resolved absolute. */
relationshipsFor(partName: string): OpcRelationship[] {
  const dir = OpcPackage.dirOf(partName)
  const base = partName.slice(dir === '' ? 0 : dir.length + 1)
  const relsPath = dir === '' ? `_rels/${base}.rels` : `${dir}/_rels/${base}.rels`
  const bytes = this.parts.get(relsPath)
  if (bytes === undefined) return []
  const rels = this.parseRelationships(new TextDecoder().decode(bytes))
  return rels.map((r) =>
    r.targetMode === 'External'
      ? r
      : { ...r, target: OpcPackage.resolveTarget(dir, r.target) },
  )
}
```

Update `parseRelationships` to capture `TargetMode`:

```typescript
private parseRelationships(xml: string): OpcRelationship[] {
  const out: OpcRelationship[] = []
  for (const tok of tokenize(xml)) {
    if (tok.type === 'open' && tok.name === 'Relationship') {
      const id = tok.attributes['Id']
      const type = tok.attributes['Type']
      const target = tok.attributes['Target']
      if (id !== undefined && type !== undefined && target !== undefined) {
        const targetMode = tok.attributes['TargetMode']
        out.push(targetMode === undefined ? { id, type, target } : { id, type, target, targetMode })
      }
    }
  }
  return out
}
```

- [ ] **Step 4: Run test to verify it passes**

Run: `npx vitest run src/opc/relationships.test.ts`
Expected: PASS (2 tests).

- [ ] **Step 5: Write the failing test for `readXlsx` recovering sheet names**

```typescript
// src/xlsx/read.test.ts
import { expect, test } from 'vitest'
import { createWorkbook } from '../model/workbook'
import { readXlsx } from './read'
import { writeXlsx } from './write'

test('readXlsx recovers sheet names round-tripped from writeXlsx', async () => {
  const wb = createWorkbook()
  wb.addSheet('Alpha')
  wb.addSheet('Beta')
  const restored = await readXlsx(await writeXlsx(wb))
  expect(restored.sheets.map((s) => s.name)).toEqual(['Alpha', 'Beta'])
})
```

- [ ] **Step 6: Run test to verify it fails**

Run: `npx vitest run src/xlsx/read.test.ts`
Expected: FAIL — `readXlsx` not defined.

- [ ] **Step 7: Implement `readWorkbookParts` and the `readXlsx` skeleton**

`xl/workbook.xml` has `<sheets><sheet name="..." sheetId="N" r:id="rIdK"/></sheets>`. The workbook part itself is found via the root relationship whose target ends in `workbook.xml` (type `.../officeDocument`). Map each sheet's `r:id` through `relationshipsFor(workbookPath)` to the worksheet part path. sharedStrings/styles are the workbook rels whose type ends in `/sharedStrings` resp. `/styles`.

```typescript
// src/xlsx/workbook-reader.ts
import type { OpcPackage } from '../opc/package'
import { tokenize } from '../xml/tokenizer'

export interface WorkbookParts {
  readonly sheets: ReadonlyArray<{ name: string; path: string }>
  readonly sharedStringsPath: string | undefined
  readonly stylesPath: string | undefined
}

const OFFICE_DOC = '/officeDocument'
const SHARED_STRINGS = '/sharedStrings'
const STYLES = '/styles'

function workbookPartPath(pkg: OpcPackage): string {
  const rel = pkg.rootRelationships().find((r) => r.type.endsWith(OFFICE_DOC))
  // root rels targets are raw and package-absolute-ish (`xl/workbook.xml`); strip a leading slash.
  const target = rel?.target ?? 'xl/workbook.xml'
  return target.replace(/^\//, '')
}

export function readWorkbookParts(pkg: OpcPackage): WorkbookParts {
  const workbookPath = workbookPartPath(pkg)
  const rels = pkg.relationshipsFor(workbookPath)
  const byId = new Map(rels.map((r) => [r.id, r]))

  const sheets: Array<{ name: string; path: string }> = []
  const xml = new TextDecoder().decode(pkg.getPart(workbookPath) ?? new Uint8Array())
  for (const tok of tokenize(xml)) {
    if (tok.type === 'open' && tok.name === 'sheet') {
      const name = tok.attributes['name']
      const rid = tok.attributes['r:id']
      const target = rid !== undefined ? byId.get(rid)?.target : undefined
      if (name !== undefined && target !== undefined) sheets.push({ name, path: target })
    }
  }
  const sharedStringsPath = rels.find((r) => r.type.endsWith(SHARED_STRINGS))?.target
  const stylesPath = rels.find((r) => r.type.endsWith(STYLES))?.target
  return { sheets, sharedStringsPath, stylesPath }
}
```

```typescript
// src/xlsx/read.ts
import type { Workbook } from '../model/workbook'
import { createWorkbook } from '../model/workbook'
import { OpcPackage } from '../opc/package'
import { readWorkbookParts } from './workbook-reader'

/** Parse `.xlsx` bytes into a Workbook. */
export async function readXlsx(bytes: Uint8Array): Promise<Workbook> {
  const pkg = await OpcPackage.read(bytes)
  const parts = readWorkbookParts(pkg)
  const wb = createWorkbook()
  for (const sheet of parts.sheets) wb.addSheet(sheet.name)
  return wb
}
```

- [ ] **Step 8: Run test to verify it passes**

Run: `npx vitest run src/xlsx/read.test.ts`
Expected: PASS.

- [ ] **Step 9: Run full gates**

Run: `npx vitest run && npm run -s typecheck && npm run -s lint && npm run -s format`
Expected: all green, no warnings.

- [ ] **Step 10: Commit**

```bash
git add src/opc/package.ts src/opc/relationships.test.ts src/xlsx/workbook-reader.ts src/xlsx/read.ts src/xlsx/read.test.ts
git commit -m "feat(xlsx): readXlsx skeleton — OPC part-scoped rels + sheet list"
```

---

### Task 2: sharedStrings reader + worksheet cell values (string/number/bool/inlineStr)

**Files:**
- Create: `src/xlsx/shared-strings-reader.ts`
- Create: `src/xlsx/worksheet-reader.ts`
- Modify: `src/xlsx/read.ts` (wire sharedStrings + worksheet parsing)
- Test: `src/xlsx/shared-strings-reader.test.ts`, `src/xlsx/read-values.test.ts`

**Interfaces:**
- Consumes: `tokenize`, `Worksheet.getCell(row,col)` / `Worksheet.cell(ref)`, `decodeAddress` from `../utils/address`, `WorkbookParts`.
- Produces: `readSharedStrings(xml: string): SharedStringValue[]`; `SharedStringValue`; `readWorksheetInto(ws: Worksheet, xml: string, ctx: ReadContext): void`; `ReadContext` (declared here, grows in later tasks).

```typescript
// src/xlsx/worksheet-reader.ts — ReadContext (extended in Tasks 3-4)
export interface ReadContext {
  readonly sharedStrings: SharedStringValue[]
}
```

The sharedStrings reader inverts `shared-strings-writer.ts`: each `<si>` is either `<t>text</t>` (plain) or a run of `<r>[<rPr>…</rPr>]<t>…</t></r>` (rich). The worksheet reader walks `<sheetData>`: for each `<c r=.. t=..>` read `<v>` (and `<is><t>` for inline). Decode by `t`: `s` → shared string index → value; `str` → string literal; `b` → boolean; `inlineStr` → inline `<is><t>`; absent → number.

- [ ] **Step 1: Write the failing test for the sharedStrings reader**

```typescript
// src/xlsx/shared-strings-reader.test.ts
import { expect, test } from 'vitest'
import { readSharedStrings } from './shared-strings-reader'
import { SharedStrings } from '../utils/shared-strings'
import { writeSharedStringsXml } from './shared-strings-writer'

test('reads plain and rich shared strings round-tripped from the writer', () => {
  const sst = new SharedStrings()
  sst.add('apple')
  sst.addRich([{ text: 'Hello ' }, { text: 'World', font: { bold: true } }])
  const values = readSharedStrings(writeSharedStringsXml(sst))

  expect(values[0]).toEqual({ kind: 'plain', text: 'apple' })
  expect(values[1]?.kind).toBe('rich')
  if (values[1]?.kind === 'rich') {
    expect(values[1].runs.map((r) => r.text).join('')).toBe('Hello World')
    expect(values[1].runs[1]?.font?.bold).toBe(true)
  }
})

test('preserves significant whitespace in plain strings', () => {
  const sst = new SharedStrings()
  sst.add('  leading')
  const values = readSharedStrings(writeSharedStringsXml(sst))
  expect(values[0]).toEqual({ kind: 'plain', text: '  leading' })
})
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx vitest run src/xlsx/shared-strings-reader.test.ts`
Expected: FAIL — `readSharedStrings` not defined.

- [ ] **Step 3: Implement the sharedStrings reader**

Run properties invert `writeRun`: `<b/>`→bold, `<i/>`→italic, `<u/>`→underline, `<sz val>`→size, `<color rgb>`→color, `<rFont val>`→name.

```typescript
// src/xlsx/shared-strings-reader.ts
import type { Font, RichTextRun } from '../model/cell'
import { tokenize } from '../xml/tokenizer'

export type SharedStringValue =
  | { readonly kind: 'plain'; readonly text: string }
  | { readonly kind: 'rich'; readonly runs: RichTextRun[] }

export function readSharedStrings(xml: string): SharedStringValue[] {
  const out: SharedStringValue[] = []
  let runs: RichTextRun[] = []
  let plain = ''
  let sawRun = false
  let inText = false
  let text = ''
  let font: Font | undefined
  let inFont = false

  for (const tok of tokenize(xml)) {
    if (tok.type === 'open') {
      if (tok.name === 'si') {
        runs = []
        plain = ''
        sawRun = false
      } else if (tok.name === 'r') {
        sawRun = true
        font = undefined
        text = ''
      } else if (tok.name === 'rPr') {
        inFont = true
        font = {}
      } else if (tok.name === 't') {
        inText = true
      } else if (inFont && font !== undefined) {
        if (tok.name === 'b') font.bold = true
        else if (tok.name === 'i') font.italic = true
        else if (tok.name === 'u') font.underline = true
        else if (tok.name === 'sz') font.size = Number(tok.attributes['val'])
        else if (tok.name === 'color') font.color = tok.attributes['rgb']
        else if (tok.name === 'rFont') font.name = tok.attributes['val']
      }
    } else if (tok.type === 'close') {
      if (tok.name === 't') inText = false
      else if (tok.name === 'rPr') inFont = false
      else if (tok.name === 'r') {
        runs.push(font !== undefined && Object.keys(font).length > 0 ? { text, font } : { text })
      } else if (tok.name === 'si') {
        out.push(sawRun ? { kind: 'rich', runs } : { kind: 'plain', text: plain })
      }
    } else if (tok.type === 'text' && inText) {
      if (sawRun) text += tok.value
      else plain += tok.value
    }
  }
  return out
}
```

- [ ] **Step 4: Run test to verify it passes**

Run: `npx vitest run src/xlsx/shared-strings-reader.test.ts`
Expected: PASS (2 tests).

- [ ] **Step 5: Write the failing test for reading string/number/bool cells**

```typescript
// src/xlsx/read-values.test.ts
import { expect, test } from 'vitest'
import { createWorkbook } from '../model/workbook'
import { readXlsx } from './read'
import { writeXlsx } from './write'

test('round-trips string, number, and boolean cell values', async () => {
  const wb = createWorkbook()
  const ws = wb.addSheet('S')
  ws.cell('A1').value = 'apple'
  ws.cell('A2').value = 'apple' // dedup -> same shared string
  ws.cell('B1').value = 42
  ws.cell('B2').value = 1234.5
  ws.cell('C1').value = true
  ws.cell('C2').value = false
  const r = await readXlsx(await writeXlsx(wb))
  const s = r.sheets[0]!
  expect(s.cell('A1').value).toBe('apple')
  expect(s.cell('A2').value).toBe('apple')
  expect(s.cell('B1').value).toBe(42)
  expect(s.cell('B2').value).toBe(1234.5)
  expect(s.cell('C1').value).toBe(true)
  expect(s.cell('C2').value).toBe(false)
})
```

- [ ] **Step 6: Run test to verify it fails**

Run: `npx vitest run src/xlsx/read-values.test.ts`
Expected: FAIL — cells are empty (worksheet not parsed yet).

- [ ] **Step 7: Implement the worksheet reader (values) + wire into `readXlsx`**

```typescript
// src/xlsx/worksheet-reader.ts
import type { CellValue } from '../model/cell'
import type { Worksheet } from '../model/worksheet'
import { decodeAddress } from '../utils/address'
import { tokenize } from '../xml/tokenizer'
import type { SharedStringValue } from './shared-strings-reader'

export interface ReadContext {
  readonly sharedStrings: SharedStringValue[]
}

function sharedToValue(v: SharedStringValue | undefined): CellValue {
  if (v === undefined) return ''
  if (v.kind === 'plain') return v.text
  return { richText: v.runs }
}

export function readWorksheetInto(ws: Worksheet, xml: string, ctx: ReadContext): void {
  let ref: string | undefined
  let type = 'n'
  let v: string | undefined
  let isV = false
  let inlineText: string | undefined
  let inIs = false
  let inT = false

  const finalize = (): void => {
    if (ref === undefined) return
    const cell = ws.cell(ref)
    if (type === 's' && v !== undefined) {
      cell.value = sharedToValue(ctx.sharedStrings[Number(v)])
    } else if (type === 'inlineStr' && inlineText !== undefined) {
      cell.value = inlineText
    } else if (type === 'str' && v !== undefined) {
      cell.value = v
    } else if (type === 'b' && v !== undefined) {
      cell.value = v === '1'
    } else if (v !== undefined) {
      cell.value = Number(v)
    }
  }

  for (const tok of tokenize(xml)) {
    if (tok.type === 'open') {
      if (tok.name === 'c') {
        finalize()
        ref = tok.attributes['r']
        type = tok.attributes['t'] ?? 'n'
        v = undefined
        inlineText = undefined
        isV = false
        inIs = false
        if (tok.selfClosing) ref = undefined
      } else if (tok.name === 'v') isV = true
      else if (tok.name === 'is') inIs = true
      else if (tok.name === 't') inT = true
    } else if (tok.type === 'close') {
      if (tok.name === 'v') isV = false
      else if (tok.name === 'is') inIs = false
      else if (tok.name === 't') inT = false
      else if (tok.name === 'c') {
        finalize()
        ref = undefined
      }
    } else if (tok.type === 'text') {
      if (isV) v = (v ?? '') + tok.value
      else if (inIs && inT) inlineText = (inlineText ?? '') + tok.value
    }
  }
  finalize()
}
```

Wire into `read.ts`:

```typescript
// src/xlsx/read.ts
import type { Workbook } from '../model/workbook'
import { createWorkbook } from '../model/workbook'
import { OpcPackage } from '../opc/package'
import { readSharedStrings, type SharedStringValue } from './shared-strings-reader'
import { readWorkbookParts } from './workbook-reader'
import { readWorksheetInto } from './worksheet-reader'

const decode = (bytes: Uint8Array | undefined): string =>
  bytes === undefined ? '' : new TextDecoder().decode(bytes)

export async function readXlsx(bytes: Uint8Array): Promise<Workbook> {
  const pkg = await OpcPackage.read(bytes)
  const parts = readWorkbookParts(pkg)

  const sharedStrings: SharedStringValue[] =
    parts.sharedStringsPath !== undefined
      ? readSharedStrings(decode(pkg.getPart(parts.sharedStringsPath)))
      : []
  const ctx = { sharedStrings }

  const wb = createWorkbook()
  for (const sheet of parts.sheets) {
    const ws = wb.addSheet(sheet.name)
    readWorksheetInto(ws, decode(pkg.getPart(sheet.path)), ctx)
  }
  return wb
}
```

- [ ] **Step 8: Run test to verify it passes**

Run: `npx vitest run src/xlsx/read-values.test.ts`
Expected: PASS.

- [ ] **Step 9: Run full gates, then commit**

Run: `npx vitest run && npm run -s typecheck && npm run -s lint && npm run -s format`

```bash
git add src/xlsx/shared-strings-reader.ts src/xlsx/shared-strings-reader.test.ts src/xlsx/worksheet-reader.ts src/xlsx/read.ts src/xlsx/read-values.test.ts
git commit -m "feat(xlsx): read sharedStrings + string/number/bool/inline cell values"
```

---

### Task 3: styles reader → CellStyle resolution + date detection

**Files:**
- Create: `src/xlsx/styles-reader.ts`
- Modify: `src/utils/number-format.ts` (add `isDateFormat`)
- Modify: `src/xlsx/worksheet-reader.ts` (apply style index; numeric+date-format → Date), `src/xlsx/read.ts` (parse styles, pass into ctx)
- Test: `src/xlsx/styles-reader.test.ts`, `src/utils/number-format.test.ts` (append), `src/xlsx/read-values.test.ts` (append)

**Interfaces:**
- Consumes: `tokenize`, `builtinFormatCode` from `../utils/number-format`, `CellStyle`/`Font`/`Fill`/`Borders`/`BorderEdge`/`Alignment` from `../model/style`, `serialToDate` from `../utils/date`.
- Produces: `readStyles(xml: string): ParsedStyles`; `ParsedStyles`; `isDateFormat(code: string): boolean`; `ReadContext` gains `cellStyles: CellStyle[]` and `numFmtById: Map<number,string>`.

`styles-reader` inverts `styles-writer`: collect `<numFmts>` (id→code), `<fonts>`, `<fills>`, `<borders>`, then `<cellXfs>` where each `<xf numFmtId fontId fillId borderId>` (+ optional `<alignment>`) maps back to a `CellStyle` referencing those tables. Builtin numFmt ids resolve via `builtinFormatCode`; the default style (xf 0) is `{}`.

`isDateFormat`: builtin date ids 14–22 and 45–47 are dates; a custom code is a date if it contains an unescaped date/time token (`y m d h s`) outside quotes, brackets, and color/condition sections. The vertical alignment maps OOXML `center` → model `middle` (inverse of the writer).

- [ ] **Step 1: Write the failing test for `isDateFormat`**

```typescript
// append to src/utils/number-format.test.ts
import { isDateFormat } from './number-format'

test('isDateFormat recognizes date/time codes and rejects numeric ones', () => {
  expect(isDateFormat('mm-dd-yy')).toBe(true)
  expect(isDateFormat('yyyy-mm-dd')).toBe(true)
  expect(isDateFormat('h:mm:ss')).toBe(true)
  expect(isDateFormat('0.00')).toBe(false)
  expect(isDateFormat('#,##0')).toBe(false)
  // 'd' inside a quoted literal or color section is not a date token
  expect(isDateFormat('"day"0')).toBe(false)
  expect(isDateFormat('[Red]0.0')).toBe(false)
})
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx vitest run src/utils/number-format.test.ts`
Expected: FAIL — `isDateFormat` not exported.

- [ ] **Step 3: Implement `isDateFormat`**

```typescript
// add to src/utils/number-format.ts
/** True when a number-format code renders a date/time (so a numeric cell should be read as a Date). */
export function isDateFormat(code: string): boolean {
  let inQuote = false
  let inBracket = false
  for (let i = 0; i < code.length; i++) {
    const ch = code[i]!
    if (inQuote) {
      if (ch === '"') inQuote = false
      continue
    }
    if (inBracket) {
      if (ch === ']') inBracket = false
      continue
    }
    if (ch === '"') inQuote = true
    else if (ch === '[') inBracket = true
    else if (ch === '\\') i++ // escaped next char is a literal
    else if ('ymdhsYMDHS'.includes(ch)) return true
  }
  return false
}
```

- [ ] **Step 4: Run test to verify it passes**

Run: `npx vitest run src/utils/number-format.test.ts`
Expected: PASS.

- [ ] **Step 5: Write the failing test for the styles reader**

```typescript
// src/xlsx/styles-reader.test.ts
import { expect, test } from 'vitest'
import { StyleRegistry, writeStylesXml } from './styles-writer'
import { readStyles } from './styles-reader'

test('reads back cellXfs into CellStyles round-tripped from the writer', () => {
  const reg = new StyleRegistry()
  const xfBold = reg.xfIndexFor({ font: { bold: true } })
  const xfFmt = reg.xfIndexFor({ numberFormat: '0.00' })
  const xfAlign = reg.xfIndexFor({ alignment: { horizontal: 'center', vertical: 'middle' } })
  const xfFill = reg.xfIndexFor({ fill: { type: 'pattern', pattern: 'solid', fgColor: 'FFFF0000' } })

  const parsed = readStyles(writeStylesXml(reg))
  expect(parsed.cellStyles[0]).toEqual({}) // default
  expect(parsed.cellStyles[xfBold]?.font?.bold).toBe(true)
  expect(parsed.cellStyles[xfFmt]?.numberFormat).toBe('0.00')
  expect(parsed.cellStyles[xfAlign]?.alignment).toEqual({ horizontal: 'center', vertical: 'middle' })
  expect(parsed.cellStyles[xfFill]?.fill).toEqual({ type: 'pattern', pattern: 'solid', fgColor: 'FFFF0000' })
})
```

- [ ] **Step 6: Run test to verify it fails**

Run: `npx vitest run src/xlsx/styles-reader.test.ts`
Expected: FAIL — `readStyles` not defined.

- [ ] **Step 7: Implement the styles reader**

Parse in document order: `<numFmts>` → map; `<fonts>`/`<fills>`/`<borders>` → arrays; `<cellXfs>` → for each `<xf>` build a `CellStyle` from referenced entries + inline `<alignment>`. Track which table we're inside via the container open/close. The vertical map inverts the writer (`center`→`middle`).

```typescript
// src/xlsx/styles-reader.ts
import type { Alignment, BorderEdge, Borders, CellStyle, Fill, Font } from '../model/style'
import { builtinFormatCode } from '../utils/number-format'
import { tokenize } from '../xml/tokenizer'

export interface ParsedStyles {
  readonly cellStyles: CellStyle[]
  readonly numFmtById: Map<number, string>
}

const VERTICAL_FROM_OOXML: Record<string, Alignment['vertical']> = {
  top: 'top',
  center: 'middle',
  bottom: 'bottom',
}

export function readStyles(xml: string): ParsedStyles {
  const numFmtById = new Map<number, string>()
  const fonts: Font[] = []
  const fills: Array<Fill | undefined> = []
  const borders: Array<Borders | undefined> = []
  const cellStyles: CellStyle[] = []

  type Section = 'numFmts' | 'fonts' | 'fills' | 'borders' | 'cellXfs' | null
  let section: Section = null
  let font: Font | undefined
  let fill: Fill | undefined
  let border: Borders | undefined
  let edge: keyof Borders | undefined
  let xf: CellStyle | undefined
  let inEdgeColor = false

  for (const tok of tokenize(xml)) {
    if (tok.type === 'open' || tok.type === 'close') {
      const open = tok.type === 'open'
      switch (tok.name) {
        case 'numFmts': case 'fonts': case 'fills': case 'borders': case 'cellXfs':
          section = open ? (tok.name as Section) : null
          continue
        case 'numFmt':
          if (open) {
            const id = Number(tok.attributes['numFmtId'])
            const code = tok.attributes['formatCode']
            if (!Number.isNaN(id) && code !== undefined) numFmtById.set(id, code)
          }
          continue
      }

      if (section === 'fonts') {
        if (tok.name === 'font') { if (open) font = {}; else if (font) { fonts.push(font); font = undefined } }
        else if (open && font !== undefined) {
          if (tok.name === 'b') font.bold = true
          else if (tok.name === 'i') font.italic = true
          else if (tok.name === 'u') font.underline = true
          else if (tok.name === 'sz') font.size = Number(tok.attributes['val'])
          else if (tok.name === 'color') font.color = tok.attributes['rgb']
          else if (tok.name === 'name') font.name = tok.attributes['val']
        }
      } else if (section === 'fills') {
        if (tok.name === 'fill') { if (open) fill = undefined; else fills.push(fill) }
        else if (open && tok.name === 'patternFill') {
          const pt = tok.attributes['patternType']
          fill = pt === 'solid' ? { type: 'pattern', pattern: 'solid' } : undefined
        } else if (open && tok.name === 'fgColor' && fill !== undefined) {
          fill = { ...fill, fgColor: tok.attributes['rgb'] }
        }
      } else if (section === 'borders') {
        if (tok.name === 'border') { if (open) border = {}; else { borders.push(border); border = undefined } }
        else if (border !== undefined && (tok.name === 'left' || tok.name === 'right' || tok.name === 'top' || tok.name === 'bottom')) {
          if (open) {
            const style = tok.attributes['style']
            if (style !== undefined) { edge = tok.name; border[edge] = { style } as BorderEdge }
            if (tok.selfClosing) edge = undefined
          } else edge = undefined
        } else if (open && tok.name === 'color' && edge !== undefined && border !== undefined) {
          const e = border[edge]
          if (e !== undefined) border[edge] = { ...e, color: tok.attributes['rgb'] }
        }
      } else if (section === 'cellXfs') {
        if (tok.name === 'xf') {
          if (open) {
            xf = {}
            const numFmtId = Number(tok.attributes['numFmtId'] ?? '0')
            const fontId = Number(tok.attributes['fontId'] ?? '0')
            const fillId = Number(tok.attributes['fillId'] ?? '0')
            const borderId = Number(tok.attributes['borderId'] ?? '0')
            const code = numFmtById.get(numFmtId) ?? builtinFormatCode(numFmtId)
            if (numFmtId !== 0 && code !== undefined) xf.numberFormat = code
            if (fontId !== 0 && fonts[fontId] !== undefined) xf.font = fonts[fontId]
            if (fillId > 1 && fills[fillId] !== undefined) xf.fill = fills[fillId]
            if (borderId !== 0 && borders[borderId] !== undefined && Object.keys(borders[borderId]!).length > 0)
              xf.border = borders[borderId]
            if (tok.selfClosing) { cellStyles.push(xf); xf = undefined }
          } else if (xf !== undefined) { cellStyles.push(xf); xf = undefined }
        } else if (open && tok.name === 'alignment' && xf !== undefined) {
          const a: Alignment = {}
          const h = tok.attributes['horizontal']
          const vv = tok.attributes['vertical']
          if (h === 'left' || h === 'center' || h === 'right') a.horizontal = h
          if (vv !== undefined && VERTICAL_FROM_OOXML[vv] !== undefined) a.vertical = VERTICAL_FROM_OOXML[vv]
          if (tok.attributes['wrapText'] === '1') a.wrapText = true
          if (Object.keys(a).length > 0) xf.alignment = a
        }
      }
    }
  }
  return { cellStyles, numFmtById }
}
```

- [ ] **Step 8: Run test to verify it passes**

Run: `npx vitest run src/xlsx/styles-reader.test.ts`
Expected: PASS. (If a facet mismatches, fix the reader, not the test.)

- [ ] **Step 9: Write the failing test for applied styles + date cells round-tripping**

```typescript
// append to src/xlsx/read-values.test.ts
test('round-trips applied cell styles and Date values', async () => {
  const wb = createWorkbook()
  const ws = wb.addSheet('S')
  const bold = ws.cell('A1'); bold.value = 'x'; bold.style = { font: { bold: true } }
  const when = new Date(Date.UTC(2026, 5, 30, 12, 0, 0))
  ws.cell('B1').value = when
  const r = await readXlsx(await writeXlsx(wb))
  const s = r.sheets[0]!
  expect(s.cell('A1').style.font?.bold).toBe(true)
  expect(s.cell('B1').value).toBeInstanceOf(Date)
  expect((s.cell('B1').value as Date).getTime()).toBe(when.getTime())
})
```

- [ ] **Step 10: Run test to verify it fails**

Run: `npx vitest run src/xlsx/read-values.test.ts`
Expected: FAIL — style not applied; B1 read as a number, not a Date.

- [ ] **Step 11: Wire styles + date detection into the worksheet reader and `read.ts`**

Extend `ReadContext` and apply: a cell's `s` attribute indexes `cellStyles`; a numeric cell whose style's `numberFormat` is a date format becomes a `Date` via `serialToDate`.

```typescript
// src/xlsx/worksheet-reader.ts — extend ReadContext
import type { CellStyle } from '../model/style'
import { serialToDate } from '../utils/date'
import { isDateFormat } from '../utils/number-format'

export interface ReadContext {
  readonly sharedStrings: SharedStringValue[]
  readonly cellStyles: CellStyle[]
}
```

In the `<c>` open branch, capture `s`:
```typescript
// add alongside ref/type capture
sIndex = tok.attributes['s'] !== undefined ? Number(tok.attributes['s']) : undefined
```
(declare `let sIndex: number | undefined` with the other state). In `finalize`, after computing the value, apply style and date detection:
```typescript
const style = sIndex !== undefined ? ctx.cellStyles[sIndex] : undefined
if (style !== undefined) cell.style = style
// numeric cell with a date format -> Date
if ((type === 'n') && v !== undefined && style?.numberFormat !== undefined && isDateFormat(style.numberFormat)) {
  cell.value = serialToDate(Number(v))
}
```
Reset `sIndex = undefined` on each `<c>` open and after finalize.

Wire styles in `read.ts`:
```typescript
import { readStyles } from './styles-reader'
// ...
const styles = parts.stylesPath !== undefined ? readStyles(decode(pkg.getPart(parts.stylesPath))) : { cellStyles: [], numFmtById: new Map() }
const ctx = { sharedStrings, cellStyles: styles.cellStyles }
```

- [ ] **Step 12: Run test to verify it passes**

Run: `npx vitest run src/xlsx/read-values.test.ts`
Expected: PASS.

- [ ] **Step 13: Run full gates, then commit**

Run: `npx vitest run && npm run -s typecheck && npm run -s lint && npm run -s format`

```bash
git add src/xlsx/styles-reader.ts src/xlsx/styles-reader.test.ts src/utils/number-format.ts src/utils/number-format.test.ts src/xlsx/worksheet-reader.ts src/xlsx/read.ts src/xlsx/read-values.test.ts
git commit -m "feat(xlsx): read styles.xml -> CellStyle + date detection"
```

---

### Task 4: formula, richText, hyperlinks, merges, columns, row heights

**Files:**
- Modify: `src/xlsx/worksheet-reader.ts` (formula `<f>`, hyperlinks via sheet rels, `<mergeCells>`, `<cols>`, row `ht`), `src/xlsx/read.ts` (pass each sheet's hyperlink rels into ctx)
- Test: `src/xlsx/read-values.test.ts` (append)

**Interfaces:**
- Consumes: `OpcPackage.relationshipsFor`, `Worksheet.merge`, `Worksheet.column`, `Worksheet.getRow`, `FormulaValue`/`HyperlinkValue` from `../model/cell`.
- Produces: `ReadContext` gains `hyperlinks: Map<string, string>` (cell ref → external target). `readWorksheetInto` now also reads merges/cols/heights/formulas/hyperlinks.

richText already returns from `sharedToValue` (Task 2) — a `t="s"` cell pointing at a rich `<si>` becomes `{ richText }`. This task adds: `<f>` (formula; result from `<v>` typed by `t`), `<hyperlinks>` (ref → r:id → external target, becomes `{ text, hyperlink }`), `<mergeCells>`, `<cols>`, and row `ht`.

- [ ] **Step 1: Write the failing test for formula/richText/merge/cols/height**

```typescript
// append to src/xlsx/read-values.test.ts
test('round-trips formula, richText, merge, column width, and row height', async () => {
  const wb = createWorkbook()
  const ws = wb.addSheet('S')
  ws.cell('A1').value = { formula: '1+2', result: 3 }
  ws.cell('A2').value = { richText: [{ text: 'Hi ' }, { text: 'bold', font: { bold: true } }] }
  ws.cell('A3').value = 'm'
  ws.merge('A3:B4')
  ws.column(1).width = 20
  ws.getRow(1).height = 30
  const r = await readXlsx(await writeXlsx(wb))
  const s = r.sheets[0]!
  expect(s.cell('A1').value).toEqual({ formula: '1+2', result: 3 })
  const rich = s.cell('A2').value
  expect(rich && typeof rich === 'object' && 'richText' in rich).toBe(true)
  expect(s.merges).toContain('A3:B4')
  expect(s.column(1).width).toBe(20)
  expect(s.getRow(1).height).toBe(30)
})

test('round-trips a hyperlink as text + target', async () => {
  const wb = createWorkbook()
  const ws = wb.addSheet('S')
  ws.cell('A1').value = { text: 'site', hyperlink: 'https://example.com/' }
  const r = await readXlsx(await writeXlsx(wb))
  expect(r.sheets[0]!.cell('A1').value).toEqual({ text: 'site', hyperlink: 'https://example.com/' })
})
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx vitest run src/xlsx/read-values.test.ts`
Expected: FAIL — formula/merge/width/height/hyperlink not recovered.

- [ ] **Step 3: Implement formula/merge/cols/height/hyperlink in the worksheet reader**

Add formula state (`<f>` text + the `<v>` typed result), and post-`sheetData` handling for `<mergeCell ref>`, `<col min max width>`, row `ht`, and `<hyperlink ref r:id>`. Hyperlink targets come from `ctx.hyperlinks` (ref → URL); when a cell ref is in that map, wrap its text value as `{ text, hyperlink }`.

Key additions to `readWorksheetInto`:
- Track `let f: string | undefined` and `let inF = false`; accumulate `<f>` text; in `finalize`, if `f !== undefined` build `{ formula: f, result: typedResult }` where `typedResult` decodes `v` by `type` (`str`→string, `b`→boolean, else number; omit if `v` undefined).
- On `<row r ht>` open with `ht`, set `ws.getRow(Number(r)).height = Number(ht)`.
- On `<col min max width>` open, for `n` in `[min,max]` set `ws.column(n).width = Number(width)`.
- On `<mergeCell ref>` open, `ws.merge(ref)`.
- On `<hyperlink ref r:id>` open, `const url = ctx.hyperlinks.get(ref); ` then re-wrap: store pending hyperlinks and apply after value finalize. Simplest: after the full parse, for each `[ref, url]` in `ctx.hyperlinks`, if the cell's value is a string, replace with `{ text: value, hyperlink: url }`.

```typescript
// formula finalize branch (inside finalize(), before the numeric fallback)
if (f !== undefined) {
  let result: FormulaValue['result']
  if (type === 'str') result = v
  else if (type === 'b') result = v === '1'
  else if (v !== undefined) result = Number(v)
  cell.value = result === undefined ? { formula: f } : { formula: f, result }
  return
}
```

```typescript
// after the token loop in readWorksheetInto, apply hyperlinks:
for (const [hRef, url] of ctx.hyperlinks) {
  const cell = ws.cell(hRef)
  if (typeof cell.value === 'string') cell.value = { text: cell.value, hyperlink: url }
}
```

Add the merge/col/row handling in the open-tag switch:
```typescript
else if (tok.name === 'mergeCell' && tok.attributes['ref'] !== undefined) ws.merge(tok.attributes['ref'])
else if (tok.name === 'col') {
  const min = Number(tok.attributes['min']); const max = Number(tok.attributes['max'])
  const width = tok.attributes['width']
  if (width !== undefined) for (let n = min; n <= max; n++) ws.column(n).width = Number(width)
}
else if (tok.name === 'row' && tok.attributes['ht'] !== undefined) {
  ws.getRow(Number(tok.attributes['r'])).height = Number(tok.attributes['ht'])
}
else if (tok.name === 'f') inF = true
```
(and `else if (tok.name === 'f') inF = false` in the close branch; accumulate `f` in the text branch when `inF`).

Extend `ReadContext` with `hyperlinks: Map<string, string>` and build it in `read.ts` per sheet:
```typescript
// src/xlsx/read.ts — per sheet
const rels = pkg.relationshipsFor(sheet.path)
const hyperlinks = new Map<string, string>()
// worksheet <hyperlink ref r:id> needs ref->rid; collect rids->target here, ref->rid in the reader.
```
Because `<hyperlink>` lives in the worksheet XML, the cleanest split is: the worksheet reader collects `ref -> rid` while parsing, and resolves `rid -> target` from a `relsById` map passed in. So set `ReadContext.hyperlinkTargets: Map<string,string>` (rid → external URL) and have the reader map ref→rid→URL.

Final `ReadContext`:
```typescript
export interface ReadContext {
  readonly sharedStrings: SharedStringValue[]
  readonly cellStyles: CellStyle[]
  readonly hyperlinkTargets: Map<string, string> // rId -> external URL
}
```
In the reader, on `<hyperlink ref r:id>` open, record `pendingHyperlinks.push({ ref, rid })`; after the loop, for each, `const url = ctx.hyperlinkTargets.get(rid)` and re-wrap the cell value. Build `hyperlinkTargets` in `read.ts` from `pkg.relationshipsFor(sheet.path)` filtering `targetMode === 'External'` (or type ending `/hyperlink`).

- [ ] **Step 4: Run test to verify it passes**

Run: `npx vitest run src/xlsx/read-values.test.ts`
Expected: PASS (all value-form round-trips).

- [ ] **Step 5: Run full gates, then commit**

Run: `npx vitest run && npm run -s typecheck && npm run -s lint && npm run -s format`

```bash
git add src/xlsx/worksheet-reader.ts src/xlsx/read.ts src/xlsx/read-values.test.ts
git commit -m "feat(xlsx): read formula/richText/hyperlink + merges + cols/row sizing"
```

---

### Task 5: fixture conformance sweep + public `readXlsx` export

**Files:**
- Modify: `src/index.ts` (export `readXlsx`), `src/index.test.ts` (append), `size-budget.json` (raise `index.js` if needed)
- Create: `src/xlsx/read-conformance.test.ts`
- Use: `test/conformance/harness.ts` (`listFixtures`, `parseFixture`, `withOracle`)

**Interfaces:**
- Consumes: `readXlsx`, `listFixtures`, `parseFixture`, exceljs oracle.
- Produces: `readXlsx` exported from the package root.

The realism gate: parse each fixture with both exceljs and `readXlsx`, and compare recovered cell **values** for a representative sheet. Be pragmatic — exceljs's value model differs in incidental ways (rich text objects, hyperlink wrappers, formula shapes); compare primitive cells (string/number/boolean) and Date timestamps, and assert our reader does not throw on any fixture and recovers a non-trivial number of cells.

- [ ] **Step 1: Write the failing test exporting `readXlsx`**

```typescript
// append to src/index.test.ts
import { readXlsx, writeXlsx, createWorkbook } from './index'

test('the package exposes readXlsx round-tripping writeXlsx', async () => {
  const wb = createWorkbook()
  wb.addSheet('S').cell('A1').value = 'hello'
  const restored = await readXlsx(await writeXlsx(wb))
  expect(restored.sheets[0]?.cell('A1').value).toBe('hello')
})
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx vitest run src/index.test.ts`
Expected: FAIL — `readXlsx` not exported.

- [ ] **Step 3: Export `readXlsx`**

```typescript
// src/index.ts — add near writeXlsx
export { readXlsx } from './xlsx/read'
```

- [ ] **Step 4: Run test to verify it passes; check size budget**

Run: `npx vitest run src/index.test.ts && npm run -s build && npm run -s size`
Expected: PASS; size within budget. If `index.js` exceeds the budget, raise it in `size-budget.json` to the measured gzip value + ~15% headroom.

- [ ] **Step 5: Write the fixture-parity conformance test**

```typescript
// src/xlsx/read-conformance.test.ts
import { expect, test } from 'vitest'
import { listFixtures, parseFixture } from '../../test/conformance/harness'
import { readXlsx } from './read'

test('readXlsx parses every fixture without throwing', async () => {
  const fixtures = listFixtures()
  expect(fixtures.length).toBeGreaterThanOrEqual(30)
  for (const path of fixtures) {
    const bytes = await parseFixture(path)
    const wb = await readXlsx(bytes)
    expect(wb.sheets.length).toBeGreaterThanOrEqual(1)
  }
})

test('recovered primitive cell values match exceljs for a representative fixture', async () => {
  const ExcelJS = (await import('exceljs')).default
  // pick a small, simple fixture deterministically (first by name)
  const path = listFixtures().sort()[0]!
  const bytes = await parseFixture(path)

  const ours = await readXlsx(bytes)
  const oracle = new ExcelJS.Workbook()
  // @ts-expect-error -- Buffer typing
  await oracle.xlsx.load(Buffer.from(bytes))

  const ourSheet = ours.sheets[0]!
  const oracleSheet = oracle.worksheets[0]!
  let compared = 0
  oracleSheet.eachRow((row, r) => {
    row.eachCell((cell, c) => {
      const v = cell.value
      if (typeof v === 'string' || typeof v === 'number' || typeof v === 'boolean') {
        expect(ourSheet.getCell(r, c).value).toBe(v)
        compared++
      }
    })
  })
  expect(compared).toBeGreaterThan(0) // non-vacuous
})
```

- [ ] **Step 6: Run the conformance test; fix reader bugs it surfaces**

Run: `npx vitest run src/xlsx/read-conformance.test.ts`
Expected: PASS. Real fixtures exercise edge cases the round-trip tests miss (inline strings, unusual numFmts, sheets without sharedStrings/styles). If a fixture throws or a value mismatches, fix the reader (e.g. guard missing parts, handle `t="inlineStr"`), not the test. If a specific fixture is structurally unsupported (note which), document and skip it explicitly with a comment — never silently.

- [ ] **Step 7: Run full gates**

Run: `npx vitest run && npm run -s typecheck && npm run -s lint && npm run -s format && npm run -s build && npm run -s size`
Expected: all green, 0 warnings.

- [ ] **Step 8: Commit**

```bash
git add src/index.ts src/index.test.ts src/xlsx/read-conformance.test.ts size-budget.json
git commit -m "feat(xlsx): public readXlsx + fixture-parity conformance sweep"
```

---

## SP-4 Acceptance Criteria

1. `readXlsx(bytes): Promise<Workbook>` parses `.xlsx` into the SP-2 model.
2. **Round-trip:** `readXlsx(await writeXlsx(wb))` recovers values (string/number/bool/date/formula/hyperlink/richText), base styling (font/fill/border/alignment/number-format), shared strings, merges, dimensions, and column/row sizing.
3. **Fixture parity:** every one of the ≥30 fixtures parses without throwing; recovered primitive cell values match exceljs for a representative fixture (non-vacuously).
4. `readXlsx` is exported from `src/index.ts` and tree-shakes away for write-only / model-only consumers.
5. All `src/xlsx/*` and `src/opc/*` changes are `node:`-free and dependency-free; all gates green (typecheck, lint 0/0, format, test, build, size).

## Self-Review

**1. Spec coverage** (architecture §SP-4 "Parse .xlsx → object model; round-trips with SP-3"):
- OPC open + relationship resolution → Task 1. ✓
- Sheet list + worksheet location → Task 1. ✓
- Shared strings (plain + rich) → Task 2. ✓
- Cell values string/number/bool/inline → Task 2; date → Task 3; formula/richText/hyperlink → Task 4. ✓
- Styles (font/fill/border/alignment/numFmt) → Task 3. ✓
- Merges, columns, row heights → Task 4. ✓
- Round-trip oracle + fixture parity → every task (round-trip) + Task 5 (fixtures). ✓
- Export + tree-shake → Task 5. ✓

**2. Placeholder scan:** Task 4's hyperlink/formula wiring is described in prose + code fragments rather than one complete file rewrite (the file is large and grows across Tasks 2–4); the implementer assembles from the fragments. All other steps carry complete code. No "TBD"/"handle edge cases" placeholders.

**3. Type consistency:** `ReadContext` is introduced in Task 2 (`sharedStrings`) and explicitly extended in Task 3 (`+cellStyles`) and Task 4 (`+hyperlinkTargets`) — each extension is called out so the implementer updates the single interface, not divergent copies. `SharedStringValue` (Task 2), `WorkbookParts` (Task 1), `ParsedStyles` (Task 3) are referenced consistently. `relationshipsFor`/`readWorkbookParts`/`readSharedStrings`/`readStyles`/`readWorksheetInto`/`readXlsx`/`isDateFormat` names are stable across tasks. Model methods (`cell`, `getCell`, `merge`, `column`, `getRow`, `addSheet`, `sheets`) match SP-2.

**Note on `ReadContext` growth:** because the interface gains fields across tasks, the worksheet reader's `read.ts` caller must construct the full object each task. Each task's wiring step shows the current shape; the implementer keeps `read.ts`'s `ctx` literal in sync.
