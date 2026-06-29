# SP-1d — `opc/` (ZIP container + OPC package) Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Implement the owned OPC (Open Packaging Conventions) layer of the `excelents` substrate: CRC-32, a standard-ZIP reader and writer (over the native DEFLATE codec), and an OPC `Package` model (content-types + relationships). The capstone wires the conformance harness's `unzipToParts`/`diffParts` to real implementations and validates them against all 35 real `.xlsx` fixtures — completing SP-1.

**Architecture:** Five focused modules under `src/opc/`. `crc32.ts` is a pure table-based CRC-32. `zip-reader.ts`/`zip-writer.ts` parse and produce standard (32-bit) ZIP archives using `DataView` for little-endian fields and the SP-1b `Codec` for DEFLATE. `package.ts` layers OPC semantics (`[Content_Types].xml`, `_rels/.rels`) over the zip using the SP-1c XML tokenizer/writer. The capstone replaces the `NotImplemented` stubs in `test/conformance/harness.ts`.

**Tech Stack:** TypeScript 7, Vitest 4, oxlint type-aware, oxfmt. Pure web-standard JS, zero deps.

## Global Constraints

- **Web-standard core only:** `src/opc/*.ts` must import NO `node:` builtin (core-purity guard).
- **Zero runtime dependencies.** No third-party imports in `src/`.
- **`isolatedDeclarations`:** explicit return types on every exported binding.
- **Named exports only**; no default export; no barrel file.
- **`sideEffects: false`** stays true.
- All gates green: `pnpm typecheck`, `pnpm lint` (0/0), `pnpm format`, `pnpm test`, `pnpm build`, `pnpm size`.
- TDD: failing test first (RED), minimal implementation (GREEN), commit. Report RED/GREEN evidence.
- **Scope:** standard ZIP (32-bit sizes/offsets), compression methods STORE (0) and DEFLATE (8). ZIP64 (>4 GB / >65535 entries) is OUT of scope — deferred to SP-7; no fixture needs it. All multi-byte ZIP fields are little-endian.
- Consumes SP-1b (`../io/codec` `nativeCodec`/`Codec`, `../io/streams` `concatUint8`) and SP-1c (`../xml/tokenizer`, `../xml/writer`).
- Branch: `excelents-rewrite` (continues from SP-1c).

---

### Task 1: CRC-32 (`src/opc/crc32.ts`)

**Files:**
- Create: `src/opc/crc32.ts`
- Test: `src/opc/crc32.test.ts`

**Interfaces:**
- Produces: `crc32(data: Uint8Array): number` — the standard ISO-HDLC CRC-32 (polynomial 0xEDB88320), as an unsigned 32-bit number.

- [ ] **Step 1: Write the failing test** (`src/opc/crc32.test.ts`)

```ts
import { expect, test } from 'vitest'
import { crc32 } from './crc32'

const enc = (s: string): Uint8Array => new TextEncoder().encode(s)

test('crc32 of empty input is 0', () => {
  expect(crc32(new Uint8Array(0))).toBe(0)
})

test('crc32 matches known vectors', () => {
  // Well-known CRC-32 test vectors.
  expect(crc32(enc('123456789'))).toBe(0xcbf43926)
  expect(crc32(enc('The quick brown fox jumps over the lazy dog'))).toBe(0x414fa339)
})

test('crc32 returns an unsigned 32-bit value', () => {
  const c = crc32(enc('hello world'))
  expect(c).toBeGreaterThanOrEqual(0)
  expect(c).toBeLessThanOrEqual(0xffffffff)
  expect(Number.isInteger(c)).toBe(true)
})
```

- [ ] **Step 2: Run the test to verify it fails**

Run: `pnpm exec vitest run src/opc/crc32.test.ts`
Expected: FAIL — cannot resolve `./crc32`.

- [ ] **Step 3: Write `src/opc/crc32.ts`**

```ts
function buildTable(): Uint32Array {
  const table = new Uint32Array(256)
  for (let n = 0; n < 256; n++) {
    let c = n
    for (let k = 0; k < 8; k++) {
      c = (c & 1) === 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1
    }
    table[n] = c >>> 0
  }
  return table
}

const CRC_TABLE = buildTable()

/** Standard ISO-HDLC CRC-32 (polynomial 0xEDB88320), unsigned 32-bit. */
export function crc32(data: Uint8Array): number {
  let c = 0xffffffff
  for (let i = 0; i < data.length; i++) {
    c = CRC_TABLE[(c ^ data[i]!) & 0xff]! ^ (c >>> 8)
  }
  return (c ^ 0xffffffff) >>> 0
}
```

- [ ] **Step 4: Run the test to verify it passes**

Run: `pnpm exec vitest run src/opc/crc32.test.ts`
Expected: PASS (3 tests). If oxlint flags the module-level `const CRC_TABLE = buildTable()` as a side effect under `sideEffects:false`, convert to lazy init (a `let table` populated on first `crc32` call) — keep `crc32`'s signature identical.

- [ ] **Step 5: Commit**

```bash
git add src/opc/crc32.ts src/opc/crc32.test.ts
git commit -m "feat(opc): CRC-32 (ISO-HDLC)"
```

---

### Task 2: ZIP reader (`src/opc/zip-reader.ts`)

**Files:**
- Create: `src/opc/zip-reader.ts`
- Test: `src/opc/zip-reader.test.ts`

**Interfaces:**
- Consumes: `Codec`, `nativeCodec` from `../io/codec`.
- Produces: `readZip(bytes: Uint8Array, codec?: Codec): Promise<Map<string, Uint8Array>>` — maps each entry's name to its decompressed bytes. Throws on a non-zip input.

**Notes:** Find the End-Of-Central-Directory record by scanning backward for signature `0x06054b50`. Read the central-directory entry count and offset, walk each central-directory record (sig `0x02014b50`), and for each, read the file's data by jumping to its local-header offset (the local header's own name/extra lengths determine where data begins). DEFLATE (method 8) → `codec.inflateRaw`; STORE (method 0) → raw copy.

- [ ] **Step 1: Write the failing test** (`src/opc/zip-reader.test.ts`)

```ts
import { expect, test } from 'vitest'
import { readZip } from './zip-reader'

test('throws on non-zip input', async () => {
  await expect(readZip(new Uint8Array([1, 2, 3, 4]))).rejects.toThrow()
})

test('reads a real xlsx fixture and finds the core OPC parts', async () => {
  const { readFile } = await import('node:fs/promises')
  const url = new URL('../../test/fixtures/1904.xlsx', import.meta.url)
  const bytes = new Uint8Array(await readFile(url))
  const parts = await readZip(bytes)
  expect(parts.has('[Content_Types].xml')).toBe(true)
  expect(parts.has('xl/workbook.xml')).toBe(true)
  // the workbook part is real XML
  expect(new TextDecoder().decode(parts.get('xl/workbook.xml')!)).toContain('<workbook')
})
```

- [ ] **Step 2: Run the test to verify it fails**

Run: `pnpm exec vitest run src/opc/zip-reader.test.ts`
Expected: FAIL — cannot resolve `./zip-reader`.

- [ ] **Step 3: Write `src/opc/zip-reader.ts`**

```ts
import { type Codec, nativeCodec } from '../io/codec'

const EOCD_SIG = 0x06054b50
const CD_SIG = 0x02014b50

/** Read a standard ZIP archive into a map of entry name -> decompressed bytes. */
export async function readZip(
  bytes: Uint8Array,
  codec: Codec = nativeCodec,
): Promise<Map<string, Uint8Array>> {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  let eocd = -1
  for (let i = bytes.length - 22; i >= 0; i--) {
    if (view.getUint32(i, true) === EOCD_SIG) {
      eocd = i
      break
    }
  }
  if (eocd === -1) throw new Error('Not a ZIP archive: no end-of-central-directory record')

  const count = view.getUint16(eocd + 10, true)
  let p = view.getUint32(eocd + 16, true)
  const decoder = new TextDecoder()
  const out = new Map<string, Uint8Array>()

  for (let n = 0; n < count; n++) {
    if (view.getUint32(p, true) !== CD_SIG) {
      throw new Error('Corrupt ZIP: bad central-directory signature')
    }
    const method = view.getUint16(p + 10, true)
    const compSize = view.getUint32(p + 20, true)
    const nameLen = view.getUint16(p + 28, true)
    const extraLen = view.getUint16(p + 30, true)
    const commentLen = view.getUint16(p + 32, true)
    const localOffset = view.getUint32(p + 42, true)
    const name = decoder.decode(bytes.subarray(p + 46, p + 46 + nameLen))

    const localNameLen = view.getUint16(localOffset + 26, true)
    const localExtraLen = view.getUint16(localOffset + 28, true)
    const dataStart = localOffset + 30 + localNameLen + localExtraLen
    const body = bytes.subarray(dataStart, dataStart + compSize)
    const data = method === 8 ? await codec.inflateRaw(body) : body.slice()
    out.set(name, data)

    p += 46 + nameLen + extraLen + commentLen
  }
  return out
}
```

- [ ] **Step 4: Run the test to verify it passes**

Run: `pnpm exec vitest run src/opc/zip-reader.test.ts`
Expected: PASS (2 tests) — including reading the real `1904.xlsx` fixture.

- [ ] **Step 5: Commit**

```bash
git add src/opc/zip-reader.ts src/opc/zip-reader.test.ts
git commit -m "feat(opc): standard ZIP reader (DEFLATE via native codec)"
```

---

### Task 3: ZIP writer (`src/opc/zip-writer.ts`)

**Files:**
- Create: `src/opc/zip-writer.ts`
- Test: `src/opc/zip-writer.test.ts`

**Interfaces:**
- Consumes: `Codec`, `nativeCodec` from `../io/codec`; `concatUint8` from `../io/streams`; `crc32` from `./crc32`; `readZip` from `./zip-reader` (round-trip test).
- Produces:
  - `interface ZipEntry { name: string; data: Uint8Array }`
  - `writeZip(entries: ZipEntry[], codec?: Codec): Promise<Uint8Array>` — a standard ZIP archive. Each entry is DEFLATE'd, or STORE'd if deflate doesn't shrink it.

- [ ] **Step 1: Write the failing test** (`src/opc/zip-writer.test.ts`)

```ts
import { expect, test } from 'vitest'
import { readZip } from './zip-reader'
import { writeZip } from './zip-writer'

const enc = (s: string): Uint8Array => new TextEncoder().encode(s)
const dec = (b: Uint8Array): string => new TextDecoder().decode(b)

test('writeZip then readZip round-trips multiple entries', async () => {
  const archive = await writeZip([
    { name: '[Content_Types].xml', data: enc('<Types/>') },
    { name: 'xl/workbook.xml', data: enc('<workbook>' + 'x'.repeat(5000) + '</workbook>') },
    { name: 'empty.bin', data: new Uint8Array(0) },
  ])
  const parts = await readZip(archive)
  expect(dec(parts.get('[Content_Types].xml')!)).toBe('<Types/>')
  expect(dec(parts.get('xl/workbook.xml')!)).toBe('<workbook>' + 'x'.repeat(5000) + '</workbook>')
  expect(parts.get('empty.bin')!.length).toBe(0)
})

test('produces a valid EOCD signature', async () => {
  const archive = await writeZip([{ name: 'a.txt', data: enc('hello') }])
  const view = new DataView(archive.buffer, archive.byteOffset, archive.byteLength)
  // EOCD is the last 22 bytes (no comment)
  expect(view.getUint32(archive.length - 22, true)).toBe(0x06054b50)
})

test('round-trips binary data with a correct CRC (no corruption)', async () => {
  const data = new Uint8Array(512)
  for (let i = 0; i < data.length; i++) data[i] = (i * 7) & 0xff
  const parts = await readZip(await writeZip([{ name: 'b.bin', data }]))
  expect([...parts.get('b.bin')!]).toEqual([...data])
})
```

- [ ] **Step 2: Run the test to verify it fails**

Run: `pnpm exec vitest run src/opc/zip-writer.test.ts`
Expected: FAIL — cannot resolve `./zip-writer`.

- [ ] **Step 3: Write `src/opc/zip-writer.ts`**

```ts
import { type Codec, nativeCodec } from '../io/codec'
import { concatUint8 } from '../io/streams'
import { crc32 } from './crc32'

export interface ZipEntry {
  name: string
  data: Uint8Array
}

const DOS_DATE = 0x0021 // 1980-01-01 (minimum DOS date)
const UTF8_FLAG = 0x0800 // bit 11: filename is UTF-8

/** Produce a standard ZIP archive; each entry DEFLATE'd, or STORE'd if smaller. */
export async function writeZip(entries: ZipEntry[], codec: Codec = nativeCodec): Promise<Uint8Array> {
  const encoder = new TextEncoder()
  const locals: Uint8Array[] = []
  const centrals: Uint8Array[] = []
  let offset = 0

  for (const entry of entries) {
    const nameBytes = encoder.encode(entry.name)
    const uncompSize = entry.data.length
    const crc = crc32(entry.data)
    const deflated = await codec.deflateRaw(entry.data)
    const store = deflated.length >= uncompSize
    const method = store ? 0 : 8
    const body = store ? entry.data : deflated
    const compSize = body.length

    const local = new Uint8Array(30 + nameBytes.length)
    const lv = new DataView(local.buffer)
    lv.setUint32(0, 0x04034b50, true)
    lv.setUint16(4, 20, true)
    lv.setUint16(6, UTF8_FLAG, true)
    lv.setUint16(8, method, true)
    lv.setUint16(10, 0, true)
    lv.setUint16(12, DOS_DATE, true)
    lv.setUint32(14, crc, true)
    lv.setUint32(18, compSize, true)
    lv.setUint32(22, uncompSize, true)
    lv.setUint16(26, nameBytes.length, true)
    lv.setUint16(28, 0, true)
    local.set(nameBytes, 30)
    locals.push(local, body)

    const central = new Uint8Array(46 + nameBytes.length)
    const cv = new DataView(central.buffer)
    cv.setUint32(0, 0x02014b50, true)
    cv.setUint16(4, 20, true)
    cv.setUint16(6, 20, true)
    cv.setUint16(8, UTF8_FLAG, true)
    cv.setUint16(10, method, true)
    cv.setUint16(12, 0, true)
    cv.setUint16(14, DOS_DATE, true)
    cv.setUint32(16, crc, true)
    cv.setUint32(20, compSize, true)
    cv.setUint32(24, uncompSize, true)
    cv.setUint16(28, nameBytes.length, true)
    cv.setUint16(30, 0, true)
    cv.setUint16(32, 0, true)
    cv.setUint16(34, 0, true)
    cv.setUint16(36, 0, true)
    cv.setUint32(38, 0, true)
    cv.setUint32(42, offset, true)
    central.set(nameBytes, 46)
    centrals.push(central)

    offset += local.length + body.length
  }

  const cdOffset = offset
  let cdSize = 0
  for (const c of centrals) cdSize += c.length

  const eocd = new Uint8Array(22)
  const ev = new DataView(eocd.buffer)
  ev.setUint32(0, 0x06054b50, true)
  ev.setUint16(4, 0, true)
  ev.setUint16(6, 0, true)
  ev.setUint16(8, entries.length, true)
  ev.setUint16(10, entries.length, true)
  ev.setUint32(12, cdSize, true)
  ev.setUint32(16, cdOffset, true)
  ev.setUint16(20, 0, true)

  return concatUint8([...locals, ...centrals, eocd])
}
```

- [ ] **Step 4: Run the test to verify it passes**

Run: `pnpm exec vitest run src/opc/zip-writer.test.ts`
Expected: PASS (3 tests).

- [ ] **Step 5: Commit**

```bash
git add src/opc/zip-writer.ts src/opc/zip-writer.test.ts
git commit -m "feat(opc): standard ZIP writer (CRC-32 + DEFLATE/STORE)"
```

---

### Task 4: OPC package model (`src/opc/package.ts`)

**Files:**
- Create: `src/opc/package.ts`
- Test: `src/opc/package.test.ts`

**Interfaces:**
- Consumes: `readZip`/`writeZip`/`ZipEntry`; `tokenize` from `../xml/tokenizer`; `XmlWriter` from `../xml/writer`.
- Produces:
  - `interface OpcRelationship { id: string; type: string; target: string }`
  - `class OpcPackage`:
    - `static read(bytes: Uint8Array): Promise<OpcPackage>`
    - `getPart(name: string): Uint8Array | undefined`
    - `partNames(): string[]`
    - `contentTypeOf(partName: string): string | undefined` — resolves via Override (by part name) then Default (by extension)
    - `rootRelationships(): OpcRelationship[]` — parsed from `_rels/.rels`
    - `setPart(name: string, contentType: string, data: Uint8Array): void` — adds/replaces a part + registers an Override content type
    - `toBytes(): Promise<Uint8Array>` — regenerates `[Content_Types].xml` from the registered defaults/overrides and writes all parts as a ZIP

**Notes:** `[Content_Types].xml` has `<Default Extension="xml" ContentType="..."/>` and `<Override PartName="/xl/workbook.xml" ContentType="..."/>` children of `<Types>`. Part names in Override are absolute (leading `/`); zip entry names are relative (no leading `/`). Normalize consistently (store parts by their zip name without leading `/`; Override PartName uses `/` + zip name).

- [ ] **Step 1: Write the failing test** (`src/opc/package.test.ts`)

```ts
import { expect, test } from 'vitest'
import { OpcPackage } from './package'

const enc = (s: string): Uint8Array => new TextEncoder().encode(s)
const dec = (b: Uint8Array): string => new TextDecoder().decode(b)

test('reads content types and root relationships from a real fixture', async () => {
  const { readFile } = await import('node:fs/promises')
  const bytes = new Uint8Array(await readFile(new URL('../../test/fixtures/1904.xlsx', import.meta.url)))
  const pkg = await OpcPackage.read(bytes)
  expect(pkg.partNames()).toContain('xl/workbook.xml')
  expect(pkg.contentTypeOf('xl/workbook.xml')).toContain('spreadsheetml')
  // the root rels point at the workbook
  const officeDoc = pkg.rootRelationships().find((r) => r.type.includes('officeDocument'))
  expect(officeDoc?.target).toContain('workbook.xml')
})

test('builds a package and round-trips through read', async () => {
  const pkg = await OpcPackage.read(await buildMinimal())
  expect(pkg.contentTypeOf('xl/workbook.xml')).toBe('application/test+xml')
  expect(dec(pkg.getPart('xl/workbook.xml')!)).toBe('<workbook/>')
})

async function buildMinimal(): Promise<Uint8Array> {
  // Construct a minimal package via setPart + toBytes, then return its bytes.
  const { OpcPackage: P } = await import('./package')
  const pkg = P.empty()
  pkg.setPart('xl/workbook.xml', 'application/test+xml', enc('<workbook/>'))
  return pkg.toBytes()
}
```

- [ ] **Step 2: Run the test to verify it fails**

Run: `pnpm exec vitest run src/opc/package.test.ts`
Expected: FAIL — cannot resolve `./package`.

- [ ] **Step 3: Write `src/opc/package.ts`**

```ts
import { XmlWriter } from '../xml/writer'
import { tokenize } from '../xml/tokenizer'
import { readZip } from './zip-reader'
import { type ZipEntry, writeZip } from './zip-writer'

export interface OpcRelationship {
  id: string
  type: string
  target: string
}

const CONTENT_TYPES = '[Content_Types].xml'
const ROOT_RELS = '_rels/.rels'
const TYPES_NS = 'http://schemas.openxmlformats.org/package/2006/content-types'
const RELS_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'

function extensionOf(name: string): string {
  const dot = name.lastIndexOf('.')
  return dot === -1 ? '' : name.slice(dot + 1).toLowerCase()
}

/** An Open Packaging Conventions container over a standard ZIP. */
export class OpcPackage {
  private readonly parts = new Map<string, Uint8Array>()
  private readonly defaults = new Map<string, string>() // extension -> content type
  private readonly overrides = new Map<string, string>() // part name (no leading /) -> content type

  static empty(): OpcPackage {
    const pkg = new OpcPackage()
    pkg.defaults.set('rels', 'application/vnd.openxmlformats-package.relationships+xml')
    pkg.defaults.set('xml', 'application/xml')
    return pkg
  }

  static async read(bytes: Uint8Array): Promise<OpcPackage> {
    const entries = await readZip(bytes)
    const pkg = new OpcPackage()
    for (const [name, data] of entries) {
      if (name === CONTENT_TYPES) continue
      pkg.parts.set(name, data)
    }
    const ct = entries.get(CONTENT_TYPES)
    if (ct !== undefined) pkg.parseContentTypes(new TextDecoder().decode(ct))
    return pkg
  }

  private parseContentTypes(xml: string): void {
    for (const tok of tokenize(xml)) {
      if (tok.type !== 'open') continue
      if (tok.name === 'Default') {
        const ext = tok.attributes['Extension']
        const type = tok.attributes['ContentType']
        if (ext !== undefined && type !== undefined) this.defaults.set(ext.toLowerCase(), type)
      } else if (tok.name === 'Override') {
        const part = tok.attributes['PartName']
        const type = tok.attributes['ContentType']
        if (part !== undefined && type !== undefined) {
          this.overrides.set(part.replace(/^\//, ''), type)
        }
      }
    }
  }

  getPart(name: string): Uint8Array | undefined {
    return this.parts.get(name)
  }

  partNames(): string[] {
    return [...this.parts.keys()]
  }

  contentTypeOf(partName: string): string | undefined {
    return this.overrides.get(partName) ?? this.defaults.get(extensionOf(partName))
  }

  rootRelationships(): OpcRelationship[] {
    const rels = this.parts.get(ROOT_RELS)
    if (rels === undefined) return []
    return this.parseRelationships(new TextDecoder().decode(rels))
  }

  private parseRelationships(xml: string): OpcRelationship[] {
    const out: OpcRelationship[] = []
    for (const tok of tokenize(xml)) {
      if (tok.type === 'open' && tok.name === 'Relationship') {
        const id = tok.attributes['Id']
        const type = tok.attributes['Type']
        const target = tok.attributes['Target']
        if (id !== undefined && type !== undefined && target !== undefined) {
          out.push({ id, type, target })
        }
      }
    }
    return out
  }

  setPart(name: string, contentType: string, data: Uint8Array): void {
    this.parts.set(name, data)
    this.overrides.set(name, contentType)
  }

  private buildContentTypes(): string {
    const w = new XmlWriter().declaration().open('Types', { xmlns: TYPES_NS })
    for (const [ext, type] of this.defaults) w.leaf('Default', { Extension: ext, ContentType: type })
    for (const [part, type] of this.overrides) {
      w.leaf('Override', { PartName: `/${part}`, ContentType: type })
    }
    return w.close('Types').toString()
  }

  async toBytes(): Promise<Uint8Array> {
    const encoder = new TextEncoder()
    const entries: ZipEntry[] = [
      { name: CONTENT_TYPES, data: encoder.encode(this.buildContentTypes()) },
    ]
    for (const [name, data] of this.parts) entries.push({ name, data })
    return writeZip(entries)
  }
}
```

Note: `RELS_NS` is referenced for documentation of the relationships namespace; if oxlint flags it as unused, remove the constant.

- [ ] **Step 4: Run the test to verify it passes**

Run: `pnpm exec vitest run src/opc/package.test.ts`
Expected: PASS (2 tests) — reads a real fixture's content types + relationships, and round-trips a built package.

- [ ] **Step 5: Commit**

```bash
git add src/opc/package.ts src/opc/package.test.ts
git commit -m "feat(opc): OPC package model (content-types + relationships)"
```

---

### Task 5: Wire the conformance harness + validate all fixtures (capstone)

**Files:**
- Modify: `test/conformance/harness.ts` (replace the `unzipToParts`/`diffParts` `NotImplemented` stubs with real implementations; the `NotImplemented` class may be removed if unused)
- Create: `test/conformance/harness-fixtures.test.ts`
- Modify: `test/conformance/harness.test.ts` (drop the now-obsolete "stubs throw NotImplemented" assertions; keep the others)

**Interfaces:**
- Consumes: `readZip` from `../../src/opc/zip-reader`; `tokenize` from `../../src/xml/tokenizer`; `XmlWriter` from `../../src/xml/writer`.
- Produces (real implementations):
  - `unzipToParts(bytes: Uint8Array): Promise<Record<string, string>>` — unzips and, for each `.xml`/`.rels` part, canonicalizes the XML (attributes sorted) to a stable string; non-XML parts are recorded as a `bytes:<length>` marker.
  - `diffParts(a, b): { part: string; detail: string }[]` — structural diff over the canonical part maps (missing parts on either side, and content differences).
  - `canonicalizeXml(xml: string): string` — tokenize, sort each element's attributes, re-emit compact via `XmlWriter`.

- [ ] **Step 1: Write the failing test** (`test/conformance/harness-fixtures.test.ts`)

```ts
import { expect, test } from 'vitest'
import { canonicalizeXml, diffParts, listFixtures, parseFixture, unzipToParts } from './harness'

test('canonicalizeXml sorts attributes so order does not matter', () => {
  expect(canonicalizeXml('<c r="A1" s="2" t="n"/>')).toBe(canonicalizeXml('<c t="n" s="2" r="A1"/>'))
  expect(canonicalizeXml('<a x="1"/>')).not.toBe(canonicalizeXml('<a x="2"/>'))
})

test('unzipToParts canonicalizes every xlsx fixture and finds [Content_Types].xml', async () => {
  for (const path of listFixtures()) {
    const parts = await unzipToParts(await parseFixture(path))
    expect(Object.keys(parts)).toContain('[Content_Types].xml')
  }
})

test('diffParts reports no differences for identical part maps and flags real ones', async () => {
  const parts = await unzipToParts(await parseFixture(listFixtures()[0]!))
  expect(diffParts(parts, parts)).toEqual([])
  const mutated = { ...parts, 'extra.xml': '<x/>' }
  expect(diffParts(parts, mutated).length).toBeGreaterThan(0)
})
```

- [ ] **Step 2: Run the test to verify it fails**

Run: `pnpm exec vitest run test/conformance/harness-fixtures.test.ts`
Expected: FAIL — `canonicalizeXml`/`unzipToParts` not implemented (still stubs).

- [ ] **Step 3: Replace the stubs in `test/conformance/harness.ts`**

Replace the `unzipToParts` and `diffParts` stub bodies (and remove the `NotImplemented` import/usage if it becomes unused) with:

```ts
import { readZip } from '../../src/opc/zip-reader'
import { tokenize } from '../../src/xml/tokenizer'
import { XmlWriter } from '../../src/xml/writer'

function isXmlPart(name: string): boolean {
  return name.endsWith('.xml') || name.endsWith('.rels')
}

/** Tokenize, sort each element's attributes, and re-emit compact — a stable canonical form. */
export function canonicalizeXml(xml: string): string {
  const w = new XmlWriter()
  for (const tok of tokenize(xml)) {
    if (tok.type === 'open') {
      const sorted: Record<string, string> = {}
      for (const key of Object.keys(tok.attributes).sort()) sorted[key] = tok.attributes[key]!
      if (tok.selfClosing) w.leaf(tok.name, sorted)
      else w.open(tok.name, sorted)
    } else if (tok.type === 'close') {
      w.close(tok.name)
    } else {
      w.text(tok.value)
    }
  }
  return w.toString()
}

/** Unzip xlsx bytes; canonicalize XML parts, mark binary parts by length. */
export async function unzipToParts(bytes: Uint8Array): Promise<Record<string, string>> {
  const entries = await readZip(bytes)
  const out: Record<string, string> = {}
  const decoder = new TextDecoder()
  for (const [name, data] of entries) {
    out[name] = isXmlPart(name) ? canonicalizeXml(decoder.decode(data)) : `bytes:${data.length}`
  }
  return out
}

/** Structural diff over canonical part maps. */
export function diffParts(
  a: Record<string, string>,
  b: Record<string, string>,
): { part: string; detail: string }[] {
  const diffs: { part: string; detail: string }[] = []
  const names = new Set([...Object.keys(a), ...Object.keys(b)])
  for (const name of [...names].sort()) {
    if (!(name in a)) diffs.push({ part: name, detail: 'missing on left' })
    else if (!(name in b)) diffs.push({ part: name, detail: 'missing on right' })
    else if (a[name] !== b[name]) diffs.push({ part: name, detail: 'content differs' })
  }
  return diffs
}
```

Keep `listFixtures`, `parseFixture`, and `withOracle` unchanged. If `NotImplemented` is no longer referenced, remove the class and its import.

- [ ] **Step 4: Update `test/conformance/harness.test.ts`**

Remove the test asserting `unzipToParts` throws `NotImplemented` (it now works). Keep the fixture-count, PK-bytes, and oracle-loads tests. If you removed `NotImplemented`, also remove its import from that test file.

- [ ] **Step 5: Run the harness tests + final full verification**

Run: `pnpm exec vitest run test/conformance/ && pnpm typecheck && pnpm lint && pnpm format && pnpm test && pnpm build && pnpm size`
Expected: all green. `unzipToParts` canonicalizes all 35 fixtures and each contains `[Content_Types].xml`. `opc/` modules tree-shake out of entry bundles (the harness imports them only in tests), so `dist` entry sizes are unchanged.

- [ ] **Step 6: Commit**

```bash
git add test/conformance/harness.ts test/conformance/harness.test.ts test/conformance/harness-fixtures.test.ts
git commit -m "feat(opc): wire conformance harness (unzip + canonicalize + diff) against fixtures"
```

---

## SP-1d Acceptance Criteria

1. `src/opc/` contains `crc32.ts`, `zip-reader.ts`, `zip-writer.ts`, `package.ts`, each with a colocated `*.test.ts`.
2. All modules are `node:`-free and dependency-free (core-purity guard passes).
3. `writeZip` → `readZip` round-trips (incl. empty + binary entries); `readZip` reads a real `.xlsx` fixture.
4. `OpcPackage.read` parses content types + root relationships from a real fixture; a built package round-trips.
5. The conformance harness's `unzipToParts`/`diffParts` are real and validate against **all 35** fixtures; the `NotImplemented` stubs are gone.
6. All gates green: typecheck, lint (0/0), format, test, build, size.

---

## Self-Review

**1. Spec coverage** (SP-1 "Unit 1 — OPC container"):
- ZIP container read/write (local headers, central directory, EOCD) → Tasks 2, 3. ✓
- Compression via the injectable codec (DEFLATE) → Tasks 2, 3 (consume `../io/codec`). ✓
- CRC-32 → Task 1. ✓
- `[Content_Types].xml` + relationships + part model → Task 4. ✓
- Capstone wiring the conformance harness (the SP-1 → SP-3 bridge) → Task 5. ✓
- **ZIP64**: explicitly deferred to SP-7 (documented in Global Constraints) — no fixture needs it; standard ZIP covers all real xlsx in scope.

**2. Placeholder scan:** No TBD/"handle later". Every step has complete code (including the full binary ZIP read/write), exact commands, and expected results.

**3. Type consistency:** `Codec`/`nativeCodec` (from SP-1b) used consistently in Tasks 2/3. `ZipEntry` defined in Task 3, consumed in Task 4. `readZip`/`writeZip` signatures match between definition and consumers. `tokenize`/`XmlToken`/`XmlWriter` (from SP-1c) used consistently in Task 4 and the Task 5 harness. `OpcRelationship` shape is consistent within Task 4. The harness's `unzipToParts`/`diffParts`/`canonicalizeXml` signatures match between Task 5's implementation and its tests.

Note for the executor: the binary ZIP code in Tasks 2-3 is the highest-risk part. The round-trip test (Task 3) and reading real fixtures (Tasks 2, 4, 5) are the correctness proof — if a real fixture fails to read, the bug is in the offset arithmetic; recheck against the standard ZIP layout (local header 30 bytes + name + extra; central dir 46 bytes + name + extra + comment; EOCD 22 bytes). All fields are little-endian.
