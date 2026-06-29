# SP-1b — `io/` (Web Streams, codec, FS adapter) Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Implement the universal IO layer of the `excelents` substrate: Web Streams helpers, a pluggable raw-DEFLATE compression codec over the native `CompressionStream` API, the `FileSystemAdapter` interface with an in-memory implementation, and the Node filesystem adapter (the `excelents/node` entry).

**Architecture:** Four focused modules. `src/io/streams.ts` and `src/io/codec.ts` and `src/io/fs.ts` are web-standard core (no `node:`), using global `ReadableStream`/`WritableStream`/`CompressionStream`/`TextEncoder`. `src/node.ts` (the only `node:`-using file) implements `FileSystemAdapter` over `node:fs`, bridging Node streams to Web Streams. The codec is an injectable interface so a runtime lacking `deflate-raw` can supply its own (e.g. `fflate`) without the core depending on it.

**Tech Stack:** TypeScript 7, Vitest 4, oxlint type-aware, oxfmt. Native Web Streams + Compression Streams API (Node 20+, browsers, edge).

## Global Constraints

Copied from the architecture + SP-1 foundation spec. Every task implicitly includes these:

- **Web-standard core only:** `src/io/*.ts` must import NO `node:` builtin (enforced by `test/core-purity.test.ts`). Only `src/node.ts` may use `node:`.
- **Zero runtime dependencies.** No third-party imports anywhere in `src/`. The default codec uses the platform `CompressionStream`/`DecompressionStream` globals.
- **`isolatedDeclarations`:** every exported binding needs an explicit return type / annotation.
- **Named exports only** (no default export). No barrel file.
- **`sideEffects: false`** stays true — these modules are pure (no top-level side effects).
- All gates green: `pnpm typecheck`, `pnpm lint` (0 errors/0 warnings), `pnpm format`, `pnpm test`, `pnpm build`, `pnpm size`.
- TDD: failing test first (RED), minimal implementation (GREEN), commit. Report RED/GREEN evidence.
- ZIP uses **raw DEFLATE** (no zlib/gzip header) — the codec uses `'deflate-raw'`, NOT `'deflate'` or `'gzip'`.
- Branch: `excelents-rewrite` (continues from SP-1a).

---

### Task 1: Web Streams helpers (`src/io/streams.ts`)

**Files:**
- Create: `src/io/streams.ts`
- Test: `src/io/streams.test.ts`

**Interfaces:**
- Produces:
  - `concatUint8(chunks: Uint8Array[]): Uint8Array` — concatenate byte chunks.
  - `bytesToReadable(bytes: Uint8Array): ReadableStream<Uint8Array>` — a one-chunk readable stream.
  - `readableToBytes(stream: ReadableStream<Uint8Array>): Promise<Uint8Array>` — collect a stream into one `Uint8Array`.

- [ ] **Step 1: Write the failing test** (`src/io/streams.test.ts`)

```ts
import { expect, test } from 'vitest'
import { bytesToReadable, concatUint8, readableToBytes } from './streams'

test('concatUint8 joins chunks in order', () => {
  const out = concatUint8([new Uint8Array([1, 2]), new Uint8Array([]), new Uint8Array([3])])
  expect([...out]).toEqual([1, 2, 3])
})

test('concatUint8 of nothing is empty', () => {
  expect(concatUint8([]).length).toBe(0)
})

test('bytesToReadable then readableToBytes round-trips', async () => {
  const input = new Uint8Array([10, 20, 30, 40])
  const out = await readableToBytes(bytesToReadable(input))
  expect([...out]).toEqual([10, 20, 30, 40])
})

test('readableToBytes collects a multi-chunk stream', async () => {
  const stream = new ReadableStream<Uint8Array>({
    start(c) {
      c.enqueue(new Uint8Array([1]))
      c.enqueue(new Uint8Array([2, 3]))
      c.close()
    },
  })
  expect([...(await readableToBytes(stream))]).toEqual([1, 2, 3])
})
```

- [ ] **Step 2: Run the test to verify it fails**

Run: `pnpm exec vitest run src/io/streams.test.ts`
Expected: FAIL — cannot resolve `./streams`.

- [ ] **Step 3: Write `src/io/streams.ts`**

```ts
/** Concatenate byte chunks into a single `Uint8Array`. */
export function concatUint8(chunks: Uint8Array[]): Uint8Array {
  let total = 0
  for (const c of chunks) total += c.length
  const out = new Uint8Array(total)
  let offset = 0
  for (const c of chunks) {
    out.set(c, offset)
    offset += c.length
  }
  return out
}

/** Wrap bytes in a single-chunk `ReadableStream`. */
export function bytesToReadable(bytes: Uint8Array): ReadableStream<Uint8Array> {
  return new ReadableStream<Uint8Array>({
    start(controller) {
      controller.enqueue(bytes)
      controller.close()
    },
  })
}

/** Collect a byte stream into one `Uint8Array`. */
export async function readableToBytes(stream: ReadableStream<Uint8Array>): Promise<Uint8Array> {
  const reader = stream.getReader()
  const chunks: Uint8Array[] = []
  for (;;) {
    const { done, value } = await reader.read()
    if (done) break
    if (value !== undefined) chunks.push(value)
  }
  return concatUint8(chunks)
}
```

- [ ] **Step 4: Run the test to verify it passes**

Run: `pnpm exec vitest run src/io/streams.test.ts`
Expected: PASS (4 tests).

- [ ] **Step 5: Commit**

```bash
git add src/io/streams.ts src/io/streams.test.ts
git commit -m "feat(io): Web Streams helpers (concat, bytes<->readable)"
```

---

### Task 2: Pluggable DEFLATE codec (`src/io/codec.ts`)

**Files:**
- Create: `src/io/codec.ts`
- Test: `src/io/codec.test.ts`

**Interfaces:**
- Consumes: `bytesToReadable`, `readableToBytes` from `./streams` (Task 1).
- Produces:
  - `interface Codec { deflateRaw(data: Uint8Array): Promise<Uint8Array>; inflateRaw(data: Uint8Array): Promise<Uint8Array> }`
  - `nativeCodec: Codec` — raw DEFLATE via the platform `CompressionStream('deflate-raw')` / `DecompressionStream('deflate-raw')`.

- [ ] **Step 1: Write the failing test** (`src/io/codec.test.ts`)

```ts
import { expect, test } from 'vitest'
import { nativeCodec } from './codec'

const enc = (s: string): Uint8Array => new TextEncoder().encode(s)
const dec = (b: Uint8Array): string => new TextDecoder().decode(b)

test('deflateRaw then inflateRaw recovers the original bytes', async () => {
  const input = enc('hello world '.repeat(50))
  const compressed = await nativeCodec.deflateRaw(input)
  const restored = await nativeCodec.inflateRaw(compressed)
  expect(dec(restored)).toBe(dec(input))
})

test('deflateRaw actually compresses redundant input', async () => {
  const input = enc('A'.repeat(10_000))
  const compressed = await nativeCodec.deflateRaw(input)
  expect(compressed.length).toBeLessThan(input.length)
})

test('round-trips empty input', async () => {
  const out = await nativeCodec.inflateRaw(await nativeCodec.deflateRaw(new Uint8Array(0)))
  expect(out.length).toBe(0)
})

test('round-trips binary (non-text) bytes', async () => {
  const input = new Uint8Array(256)
  for (let i = 0; i < 256; i++) input[i] = i
  const out = await nativeCodec.inflateRaw(await nativeCodec.deflateRaw(input))
  expect([...out]).toEqual([...input])
})
```

- [ ] **Step 2: Run the test to verify it fails**

Run: `pnpm exec vitest run src/io/codec.test.ts`
Expected: FAIL — cannot resolve `./codec`.

- [ ] **Step 3: Write `src/io/codec.ts`**

```ts
import { bytesToReadable, readableToBytes } from './streams'

/** A raw-DEFLATE compression codec (no zlib/gzip header — the format ZIP uses). */
export interface Codec {
  deflateRaw(data: Uint8Array): Promise<Uint8Array>
  inflateRaw(data: Uint8Array): Promise<Uint8Array>
}

async function transform(data: Uint8Array, stream: TransformStream<Uint8Array, Uint8Array>): Promise<Uint8Array> {
  return readableToBytes(bytesToReadable(data).pipeThrough(stream))
}

/** Default codec using the platform Compression Streams API (`deflate-raw`). */
export const nativeCodec: Codec = {
  deflateRaw(data: Uint8Array): Promise<Uint8Array> {
    return transform(data, new CompressionStream('deflate-raw'))
  },
  inflateRaw(data: Uint8Array): Promise<Uint8Array> {
    return transform(data, new DecompressionStream('deflate-raw'))
  },
}
```

- [ ] **Step 4: Run the test to verify it passes**

Run: `pnpm exec vitest run src/io/codec.test.ts`
Expected: PASS (4 tests). If `CompressionStream`/`DecompressionStream` is reported as undefined, the Node version is too old — confirm Node ≥ 20 (it is, per `engines`); do NOT add a polyfill dependency.

- [ ] **Step 5: Commit**

```bash
git add src/io/codec.ts src/io/codec.test.ts
git commit -m "feat(io): pluggable raw-DEFLATE codec over native CompressionStream"
```

---

### Task 3: FileSystemAdapter interface + in-memory implementation (`src/io/fs.ts`)

**Files:**
- Create: `src/io/fs.ts`
- Test: `src/io/fs.test.ts`

**Interfaces:**
- Consumes: `bytesToReadable`, `readableToBytes`, `concatUint8` from `./streams` (Task 1).
- Produces:
  - `interface FileSystemAdapter { readFile(path: string): Promise<Uint8Array>; writeFile(path: string, data: Uint8Array): Promise<void>; createReadable(path: string): ReadableStream<Uint8Array>; createWritable(path: string): WritableStream<Uint8Array> }`
  - `createMemoryFileSystem(initial?: Record<string, Uint8Array>): FileSystemAdapter` — a pure, in-memory adapter (no `node:`), useful for tests and edge/browser runtimes.

- [ ] **Step 1: Write the failing test** (`src/io/fs.test.ts`)

```ts
import { expect, test } from 'vitest'
import { readableToBytes } from './streams'
import { createMemoryFileSystem } from './fs'

test('writeFile then readFile round-trips', async () => {
  const fs = createMemoryFileSystem()
  await fs.writeFile('a.bin', new Uint8Array([1, 2, 3]))
  expect([...(await fs.readFile('a.bin'))]).toEqual([1, 2, 3])
})

test('readFile of a missing path rejects', async () => {
  const fs = createMemoryFileSystem()
  await expect(fs.readFile('nope')).rejects.toThrow()
})

test('seeds from initial contents', async () => {
  const fs = createMemoryFileSystem({ 'seed.txt': new Uint8Array([9]) })
  expect([...(await fs.readFile('seed.txt'))]).toEqual([9])
})

test('createReadable streams a written file', async () => {
  const fs = createMemoryFileSystem()
  await fs.writeFile('r.bin', new Uint8Array([5, 6, 7]))
  expect([...(await readableToBytes(fs.createReadable('r.bin')))]).toEqual([5, 6, 7])
})

test('createWritable collects chunks into a readable file', async () => {
  const fs = createMemoryFileSystem()
  const w = fs.createWritable('w.bin')
  const writer = w.getWriter()
  await writer.write(new Uint8Array([1, 2]))
  await writer.write(new Uint8Array([3]))
  await writer.close()
  expect([...(await fs.readFile('w.bin'))]).toEqual([1, 2, 3])
})
```

- [ ] **Step 2: Run the test to verify it fails**

Run: `pnpm exec vitest run src/io/fs.test.ts`
Expected: FAIL — cannot resolve `./fs`.

- [ ] **Step 3: Write `src/io/fs.ts`**

```ts
import { bytesToReadable, concatUint8, readableToBytes } from './streams'

/** Abstracts file access so the universal core never imports `node:fs` directly. */
export interface FileSystemAdapter {
  readFile(path: string): Promise<Uint8Array>
  writeFile(path: string, data: Uint8Array): Promise<void>
  createReadable(path: string): ReadableStream<Uint8Array>
  createWritable(path: string): WritableStream<Uint8Array>
}

/** A pure in-memory `FileSystemAdapter` (no `node:`), for tests and edge/browser use. */
export function createMemoryFileSystem(
  initial: Record<string, Uint8Array> = {},
): FileSystemAdapter {
  const files = new Map<string, Uint8Array>(Object.entries(initial))
  return {
    readFile(path: string): Promise<Uint8Array> {
      const data = files.get(path)
      if (data === undefined) return Promise.reject(new Error(`No such file: ${path}`))
      return Promise.resolve(data)
    },
    writeFile(path: string, data: Uint8Array): Promise<void> {
      files.set(path, data)
      return Promise.resolve()
    },
    createReadable(path: string): ReadableStream<Uint8Array> {
      const data = files.get(path)
      if (data === undefined) throw new Error(`No such file: ${path}`)
      return bytesToReadable(data)
    },
    createWritable(path: string): WritableStream<Uint8Array> {
      const chunks: Uint8Array[] = []
      return new WritableStream<Uint8Array>({
        write(chunk: Uint8Array): void {
          chunks.push(chunk)
        },
        close(): void {
          files.set(path, concatUint8(chunks))
        },
      })
    },
  }
}
```

- [ ] **Step 4: Run the test to verify it passes**

Run: `pnpm exec vitest run src/io/fs.test.ts`
Expected: PASS (5 tests).

- [ ] **Step 5: Commit**

```bash
git add src/io/fs.ts src/io/fs.test.ts
git commit -m "feat(io): FileSystemAdapter interface + in-memory implementation"
```

---

### Task 4: Node filesystem adapter (`src/node.ts`)

**Files:**
- Modify: `src/node.ts` (replace the `nodeAdapterPlaceholder` stub)
- Test: `src/node.test.ts`

**Interfaces:**
- Consumes: `FileSystemAdapter` from `./io/fs` (type-only).
- Produces (the `excelents/node` entry surface):
  - `nodeFileSystem: FileSystemAdapter` — a `FileSystemAdapter` backed by `node:fs`, bridging Node streams to Web Streams.
  - Removes the `nodeAdapterPlaceholder` export.

**Notes:** `Readable.toWeb()` / `Writable.toWeb()` (Node 17+) bridge Node streams ↔ Web Streams. This file is the ONE place `node:` imports are allowed.

- [ ] **Step 1: Write the failing test** (`src/node.test.ts`)

```ts
import { mkdtempSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterAll, beforeAll, expect, test } from 'vitest'
import { readableToBytes } from './io/streams'
import { nodeFileSystem } from './node'

let dir = ''
beforeAll(() => {
  dir = mkdtempSync(join(tmpdir(), 'excelents-node-'))
})
afterAll(() => {
  rmSync(dir, { recursive: true, force: true })
})

test('writeFile then readFile round-trips on disk', async () => {
  const p = join(dir, 'a.bin')
  await nodeFileSystem.writeFile(p, new Uint8Array([1, 2, 3]))
  expect([...(await nodeFileSystem.readFile(p))]).toEqual([1, 2, 3])
})

test('createReadable streams a file from disk', async () => {
  const p = join(dir, 'r.bin')
  await nodeFileSystem.writeFile(p, new Uint8Array([7, 8, 9]))
  expect([...(await readableToBytes(nodeFileSystem.createReadable(p)))]).toEqual([7, 8, 9])
})

test('createWritable writes a file to disk', async () => {
  const p = join(dir, 'w.bin')
  const writer = nodeFileSystem.createWritable(p).getWriter()
  await writer.write(new Uint8Array([4, 5]))
  await writer.write(new Uint8Array([6]))
  await writer.close()
  expect([...(await nodeFileSystem.readFile(p))]).toEqual([4, 5, 6])
})
```

- [ ] **Step 2: Run the test to verify it fails**

Run: `pnpm exec vitest run src/node.test.ts`
Expected: FAIL — `nodeFileSystem` is not exported (the file still has `nodeAdapterPlaceholder`).

- [ ] **Step 3: Rewrite `src/node.ts`**

```ts
import { createReadStream, createWriteStream } from 'node:fs'
import { readFile as nodeReadFile, writeFile as nodeWriteFile } from 'node:fs/promises'
import { Readable, Writable } from 'node:stream'
import type { FileSystemAdapter } from './io/fs'

/** A `FileSystemAdapter` backed by Node's filesystem, bridging Node streams to Web Streams. */
export const nodeFileSystem: FileSystemAdapter = {
  async readFile(path: string): Promise<Uint8Array> {
    return new Uint8Array(await nodeReadFile(path))
  },
  async writeFile(path: string, data: Uint8Array): Promise<void> {
    await nodeWriteFile(path, data)
  },
  createReadable(path: string): ReadableStream<Uint8Array> {
    return Readable.toWeb(createReadStream(path)) as ReadableStream<Uint8Array>
  },
  createWritable(path: string): WritableStream<Uint8Array> {
    return Writable.toWeb(createWriteStream(path)) as WritableStream<Uint8Array>
  },
}
```

- [ ] **Step 4: Run the test to verify it passes**

Run: `pnpm exec vitest run src/node.test.ts`
Expected: PASS (3 tests). If `Readable.toWeb`'s type doesn't line up with the global `ReadableStream<Uint8Array>` under TS strict, the `as` cast handles it (Node's web-stream types and the global lib types differ nominally); keep the cast minimal and oxlint-clean (a single documented assertion).

- [ ] **Step 5: Final full verification + commit**

Run: `pnpm typecheck && pnpm lint && pnpm format && pnpm test && pnpm build && pnpm size`
Expected: all green. The core-purity test still passes (the new `node:` imports are in `src/node.ts`, which it skips). `pnpm build` regenerates the `./node` entry (now exporting `nodeFileSystem`); confirm `pnpm size` passes (the node entry grew slightly but stays within budget — if it exceeds, raise `node.js` in `size-budget.json` to a sensible new value and note it).

```bash
git add src/node.ts src/node.test.ts
# size-budget.json only if you had to adjust it
git commit -m "feat(io): Node filesystem adapter (node:fs <-> Web Streams)"
```

---

## SP-1b Acceptance Criteria

1. `src/io/` contains `streams.ts`, `codec.ts`, `fs.ts`, each with a colocated `*.test.ts`; `src/node.ts` provides `nodeFileSystem`.
2. `src/io/*` are `node:`-free (core-purity guard passes); only `src/node.ts` uses `node:`.
3. The codec round-trips raw DEFLATE via the native API and is defined behind the injectable `Codec` interface (no `fflate` dependency).
4. Both an in-memory and a Node `FileSystemAdapter` exist and pass round-trip + streaming tests.
5. All gates green: typecheck, lint (0/0), format, test, build, size.

---

## Self-Review

**1. Spec coverage** (SP-1 "Unit 3 — IO & adapters"):
- Web Streams helpers (chunking, bytes↔stream) → Task 1. ✓
- Codec abstraction (pluggable, default native) → Task 2. ✓
- `FileSystemAdapter` interface → Task 3 (+ in-memory impl, a useful testable addition). ✓
- Node adapter implementation in `src/node.ts` (node:fs ↔ Web Streams) → Task 4. ✓
- base64 helpers: intentionally deferred to SP-5 (media), per YAGNI — not needed by the substrate spine.

**2. Placeholder scan:** No TBD/"add error handling"/"similar to". Every step has complete code, an exact command, and an expected result.

**3. Type consistency:** `bytesToReadable`/`readableToBytes`/`concatUint8` signatures match between Task 1 (defined) and Tasks 2/3 (consumed). `Codec` shape is consistent within Task 2. `FileSystemAdapter` shape matches between Task 3 (defined) and Task 4 (consumed as a type). `createMemoryFileSystem` and `nodeFileSystem` both satisfy the same `FileSystemAdapter` interface.

Note for the executor: removing `nodeAdapterPlaceholder` changes the `./node` entry surface — that's intended (it was an SP-0 stub). Confirm nothing references it (`grep -rn nodeAdapterPlaceholder src test`) before deleting; there should be no references.
