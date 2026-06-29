# SP-1c — `xml/` (owned streaming tokenizer + typed writer) Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Implement the owned XML layer of the `excelents` substrate: entity escape/unescape helpers, a typed XML writer (well-formed output with correct escaping and namespaced names), and a pull/SAX tokenizer for the OOXML subset — replacing the abandoned `saxes` dependency with zero-dependency owned code.

**Architecture:** Three focused web-standard modules under `src/xml/`. `entities.ts` handles the five predefined XML entities + numeric character references. `writer.ts` is a string-builder `XmlWriter` for generating OOXML. `tokenizer.ts` is a generator (`function* tokenize`) that lazily yields open/close/text tokens over a complete XML string (chunked-over-streams tokenizing is deferred to SP-7). The OOXML we parse is machine-generated and well-formed: no DTDs, no custom entity definitions, no exotic XML features — a focused tokenizer is sufficient and fully owned.

**Tech Stack:** TypeScript 7, Vitest 4, oxlint type-aware, oxfmt. Pure JS, zero deps.

## Global Constraints

- **Web-standard core only:** `src/xml/*.ts` must import NO `node:` builtin (enforced by `test/core-purity.test.ts`).
- **Zero runtime dependencies.** No third-party imports anywhere in `src/`.
- **`isolatedDeclarations`:** every exported binding needs an explicit return type / annotation.
- **Named exports only** (no default export). No barrel file.
- **`sideEffects: false`** stays true.
- All gates green: `pnpm typecheck`, `pnpm lint` (0 errors/0 warnings), `pnpm format`, `pnpm test`, `pnpm build`, `pnpm size`.
- TDD: failing test first (RED), minimal implementation (GREEN), commit. Report RED/GREEN evidence.
- The tokenizer handles the OOXML subset: elements, attributes (both `"` and `'` quotes), namespaced qualified names (e.g. `r:id`, `xmlns:r`), self-closing tags, text, CDATA, the five predefined entities, and numeric char refs. It SKIPS the `<?xml?>` declaration, comments (`<!-- -->`), and processing instructions. Names are treated lexically (no namespace resolution).
- Branch: `excelents-rewrite` (continues from SP-1b).

---

### Task 1: XML entity escape/unescape (`src/xml/entities.ts`)

**Files:**
- Create: `src/xml/entities.ts`
- Test: `src/xml/entities.test.ts`

**Interfaces:**
- Produces:
  - `escapeText(text: string): string` — escapes `&`, `<`, `>` for element text content.
  - `escapeAttr(value: string): string` — escapes `&`, `<`, `>`, `"` for attribute values (double-quoted).
  - `unescapeXml(text: string): string` — decodes the five predefined entities and numeric char refs (`&#NNN;`, `&#xHH;`).

- [ ] **Step 1: Write the failing test** (`src/xml/entities.test.ts`)

```ts
import { expect, test } from 'vitest'
import { escapeAttr, escapeText, unescapeXml } from './entities'

test('escapeText escapes &, <, > but leaves quotes', () => {
  expect(escapeText('a & b < c > d "e" \'f\'')).toBe('a &amp; b &lt; c &gt; d "e" \'f\'')
})

test('escapeAttr escapes &, <, >, and double-quote', () => {
  expect(escapeAttr('x "y" & <z>')).toBe('x &quot;y&quot; &amp; &lt;z&gt;')
})

test('escapeText escapes & before other entities (no double-escaping)', () => {
  expect(escapeText('&lt;')).toBe('&amp;lt;')
})

test('unescapeXml decodes the five predefined entities', () => {
  expect(unescapeXml('a &amp; b &lt; c &gt; d &quot;e&quot; &apos;f&apos;')).toBe('a & b < c > d "e" \'f\'')
})

test('unescapeXml decodes decimal and hex numeric char refs', () => {
  expect(unescapeXml('&#65;&#66;&#x43;')).toBe('ABC')
  expect(unescapeXml('tab&#9;end')).toBe('tab\tend')
})

test('escape then unescape round-trips arbitrary text', () => {
  const s = 'Tom & Jerry: <5 "quoted" & \'apos\'>'
  expect(unescapeXml(escapeAttr(s))).toBe(s)
  expect(unescapeXml(escapeText(s))).toBe(s)
})
```

- [ ] **Step 2: Run the test to verify it fails**

Run: `pnpm exec vitest run src/xml/entities.test.ts`
Expected: FAIL — cannot resolve `./entities`.

- [ ] **Step 3: Write `src/xml/entities.ts`**

```ts
/** Escape text content (`&`, `<`, `>`). */
export function escapeText(text: string): string {
  return text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
}

/** Escape a double-quoted attribute value (`&`, `<`, `>`, `"`). */
export function escapeAttr(value: string): string {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
}

const NAMED: Record<string, string> = {
  amp: '&',
  lt: '<',
  gt: '>',
  quot: '"',
  apos: "'",
}

/** Decode the five predefined entities and numeric char refs. */
export function unescapeXml(text: string): string {
  return text.replace(/&(#x?[0-9a-fA-F]+|[a-zA-Z]+);/g, (match, body: string): string => {
    if (body[0] === '#') {
      const code = body[1] === 'x' || body[1] === 'X' ? parseInt(body.slice(2), 16) : parseInt(body.slice(1), 10)
      return Number.isNaN(code) ? match : String.fromCodePoint(code)
    }
    return NAMED[body] ?? match
  })
}
```

- [ ] **Step 4: Run the test to verify it passes**

Run: `pnpm exec vitest run src/xml/entities.test.ts`
Expected: PASS (6 tests).

- [ ] **Step 5: Commit**

```bash
git add src/xml/entities.ts src/xml/entities.test.ts
git commit -m "feat(xml): entity escape/unescape helpers"
```

---

### Task 2: Typed XML writer (`src/xml/writer.ts`)

**Files:**
- Create: `src/xml/writer.ts`
- Test: `src/xml/writer.test.ts`

**Interfaces:**
- Consumes: `escapeText`, `escapeAttr` from `./entities` (Task 1).
- Produces:
  - `type XmlAttrs = Record<string, string | number | undefined>` — `undefined`-valued attributes are skipped.
  - `class XmlWriter` with chainable methods:
    - `declaration(): this` — emits `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`
    - `open(name: string, attrs?: XmlAttrs): this` — emits `<name a="b">`
    - `leaf(name: string, attrs?: XmlAttrs): this` — emits self-closing `<name a="b"/>`
    - `text(content: string): this` — emits escaped text
    - `close(name: string): this` — emits `</name>`
    - `toString(): string` — the accumulated XML

- [ ] **Step 1: Write the failing test** (`src/xml/writer.test.ts`)

```ts
import { expect, test } from 'vitest'
import { XmlWriter } from './writer'

test('writes a declaration', () => {
  expect(new XmlWriter().declaration().toString()).toBe(
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
  )
})

test('writes nested elements with attributes', () => {
  const xml = new XmlWriter()
    .open('worksheet', { 'xmlns:r': 'ns' })
    .leaf('dimension', { ref: 'A1:B2' })
    .close('worksheet')
    .toString()
  expect(xml).toBe('<worksheet xmlns:r="ns"><dimension ref="A1:B2"/></worksheet>')
})

test('skips undefined attributes and renders numbers', () => {
  expect(new XmlWriter().leaf('c', { r: 'A1', s: 3, t: undefined }).toString()).toBe(
    '<c r="A1" s="3"/>',
  )
})

test('escapes text content and attribute values', () => {
  const xml = new XmlWriter().open('t').text('a & b < c').close('t').toString()
  expect(xml).toBe('<t>a &amp; b &lt; c</t>')
  expect(new XmlWriter().leaf('a', { v: '"q" & <x>' }).toString()).toBe(
    '<a v="&quot;q&quot; &amp; &lt;x&gt;"/>',
  )
})

test('emits attributes in insertion order', () => {
  expect(new XmlWriter().leaf('c', { r: 'A1', s: 2, t: 's' }).toString()).toBe('<c r="A1" s="2" t="s"/>')
})
```

- [ ] **Step 2: Run the test to verify it fails**

Run: `pnpm exec vitest run src/xml/writer.test.ts`
Expected: FAIL — cannot resolve `./writer`.

- [ ] **Step 3: Write `src/xml/writer.ts`**

```ts
import { escapeAttr, escapeText } from './entities'

export type XmlAttrs = Record<string, string | number | undefined>

function renderAttrs(attrs: XmlAttrs | undefined): string {
  if (attrs === undefined) return ''
  let out = ''
  for (const key of Object.keys(attrs)) {
    const value = attrs[key]
    if (value === undefined) continue
    out += ` ${key}="${escapeAttr(String(value))}"`
  }
  return out
}

/** Builds well-formed XML as a string, with correct escaping. */
export class XmlWriter {
  private buf = ''

  declaration(): this {
    this.buf += '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    return this
  }

  open(name: string, attrs?: XmlAttrs): this {
    this.buf += `<${name}${renderAttrs(attrs)}>`
    return this
  }

  leaf(name: string, attrs?: XmlAttrs): this {
    this.buf += `<${name}${renderAttrs(attrs)}/>`
    return this
  }

  text(content: string): this {
    this.buf += escapeText(content)
    return this
  }

  close(name: string): this {
    this.buf += `</${name}>`
    return this
  }

  toString(): string {
    return this.buf
  }
}
```

- [ ] **Step 4: Run the test to verify it passes**

Run: `pnpm exec vitest run src/xml/writer.test.ts`
Expected: PASS (5 tests).

- [ ] **Step 5: Commit**

```bash
git add src/xml/writer.ts src/xml/writer.test.ts
git commit -m "feat(xml): typed XML writer (escaping, namespaced names)"
```

---

### Task 3: OOXML tokenizer (`src/xml/tokenizer.ts`)

**Files:**
- Create: `src/xml/tokenizer.ts`
- Test: `src/xml/tokenizer.test.ts`

**Interfaces:**
- Consumes: `unescapeXml` from `./entities` (Task 1); `XmlWriter` from `./writer` (Task 2, for the round-trip test only).
- Produces:
  - `type XmlToken = { type: 'open'; name: string; attributes: Record<string, string>; selfClosing: boolean } | { type: 'close'; name: string } | { type: 'text'; value: string }`
  - `tokenize(xml: string): Generator<XmlToken, void, unknown>` — lazily yields tokens; skips declaration/comments/PIs; emits CDATA as text; unescapes entities in text and attribute values.

- [ ] **Step 1: Write the failing test** (`src/xml/tokenizer.test.ts`)

```ts
import { expect, test } from 'vitest'
import { tokenize, type XmlToken } from './tokenizer'
import { XmlWriter } from './writer'

const toks = (xml: string): XmlToken[] => [...tokenize(xml)]

test('tokenizes nested elements and attributes', () => {
  expect(toks('<a x="1"><b/></a>')).toEqual([
    { type: 'open', name: 'a', attributes: { x: '1' }, selfClosing: false },
    { type: 'open', name: 'b', attributes: {}, selfClosing: true },
    { type: 'close', name: 'a' },
  ])
})

test('handles both quote styles and namespaced names', () => {
  expect(toks(`<c:e r:id='R1' n="2"/>`)).toEqual([
    { type: 'open', name: 'c:e', attributes: { 'r:id': 'R1', n: '2' }, selfClosing: true },
  ])
})

test('emits text and unescapes entities in text and attributes', () => {
  expect(toks('<t a="x &amp; y">a &lt; b</t>')).toEqual([
    { type: 'open', name: 't', attributes: { a: 'x & y' }, selfClosing: false },
    { type: 'text', value: 'a < b' },
    { type: 'close', name: 't' },
  ])
})

test('emits CDATA as raw text', () => {
  expect(toks('<t><![CDATA[a < b & c]]></t>')).toEqual([
    { type: 'open', name: 't', attributes: {}, selfClosing: false },
    { type: 'text', value: 'a < b & c' },
    { type: 'close', name: 't' },
  ])
})

test('skips the xml declaration and comments', () => {
  expect(toks('<?xml version="1.0"?><!-- hi --><a/>')).toEqual([
    { type: 'open', name: 'a', attributes: {}, selfClosing: true },
  ])
})

test('preserves whitespace-only text runs between elements', () => {
  expect(toks('<a>\n  <b/>\n</a>')).toEqual([
    { type: 'open', name: 'a', attributes: {}, selfClosing: false },
    { type: 'text', value: '\n  ' },
    { type: 'open', name: 'b', attributes: {}, selfClosing: true },
    { type: 'text', value: '\n' },
    { type: 'close', name: 'a' },
  ])
})

test('round-trips through the writer (open/leaf/text/close) to equivalent XML', () => {
  const input = '<root xmlns:r="ns"><a r:id="R1">hello &amp; bye</a><b/></root>'
  const w = new XmlWriter()
  for (const t of tokenize(input)) {
    if (t.type === 'open') {
      if (t.selfClosing) w.leaf(t.name, t.attributes)
      else w.open(t.name, t.attributes)
    } else if (t.type === 'close') w.close(t.name)
    else w.text(t.value)
  }
  expect(w.toString()).toBe(input)
})
```

- [ ] **Step 2: Run the test to verify it fails**

Run: `pnpm exec vitest run src/xml/tokenizer.test.ts`
Expected: FAIL — cannot resolve `./tokenizer`.

- [ ] **Step 3: Write `src/xml/tokenizer.ts`**

```ts
import { unescapeXml } from './entities'

export type XmlToken =
  | { type: 'open'; name: string; attributes: Record<string, string>; selfClosing: boolean }
  | { type: 'close'; name: string }
  | { type: 'text'; value: string }

function isSpace(ch: string): boolean {
  return ch === ' ' || ch === '\t' || ch === '\n' || ch === '\r'
}

/** Lazily tokenize an OOXML-subset XML string. */
export function* tokenize(xml: string): Generator<XmlToken, void, unknown> {
  let i = 0
  const n = xml.length
  while (i < n) {
    if (xml[i] !== '<') {
      const next = xml.indexOf('<', i)
      const end = next === -1 ? n : next
      yield { type: 'text', value: unescapeXml(xml.slice(i, end)) }
      i = end
      continue
    }
    if (xml.startsWith('<?', i)) {
      const end = xml.indexOf('?>', i + 2)
      i = end === -1 ? n : end + 2
    } else if (xml.startsWith('<!--', i)) {
      const end = xml.indexOf('-->', i + 4)
      i = end === -1 ? n : end + 3
    } else if (xml.startsWith('<![CDATA[', i)) {
      const end = xml.indexOf(']]>', i + 9)
      yield { type: 'text', value: xml.slice(i + 9, end === -1 ? n : end) }
      i = end === -1 ? n : end + 3
    } else if (xml.startsWith('</', i)) {
      const end = xml.indexOf('>', i + 2)
      yield { type: 'close', name: xml.slice(i + 2, end).trim() }
      i = end + 1
    } else {
      i += 1
      let j = i
      while (j < n && !isSpace(xml[j]!) && xml[j] !== '>' && xml[j] !== '/') j++
      const name = xml.slice(i, j)
      i = j
      const attributes: Record<string, string> = {}
      for (;;) {
        while (i < n && isSpace(xml[i]!)) i++
        if (i >= n) break
        if (xml[i] === '/' && xml[i + 1] === '>') {
          i += 2
          yield { type: 'open', name, attributes, selfClosing: true }
          break
        }
        if (xml[i] === '>') {
          i += 1
          yield { type: 'open', name, attributes, selfClosing: false }
          break
        }
        let k = i
        while (k < n && xml[k] !== '=' && !isSpace(xml[k]!) && xml[k] !== '>' && xml[k] !== '/') k++
        const attrName = xml.slice(i, k)
        i = k
        while (i < n && isSpace(xml[i]!)) i++
        if (xml[i] === '=') {
          i += 1
          while (i < n && isSpace(xml[i]!)) i++
          const quote = xml[i]
          if (quote === '"' || quote === "'") {
            i += 1
            const close = xml.indexOf(quote, i)
            attributes[attrName] = unescapeXml(xml.slice(i, close === -1 ? n : close))
            i = close === -1 ? n : close + 1
          }
        } else {
          attributes[attrName] = ''
        }
      }
    }
  }
}
```

- [ ] **Step 4: Run the test to verify it passes**

Run: `pnpm exec vitest run src/xml/tokenizer.test.ts`
Expected: PASS (7 tests). If an edge case fails, fix the tokenizer (not the tests) — the tests encode the required behavior.

- [ ] **Step 5: Final full verification + commit**

Run: `pnpm typecheck && pnpm lint && pnpm format && pnpm test && pnpm build && pnpm size`
Expected: all green; `xml/` modules tree-shake (unreferenced by entry points), so `dist` entry sizes are unchanged.

```bash
git add src/xml/tokenizer.ts src/xml/tokenizer.test.ts
git commit -m "feat(xml): OOXML-subset streaming tokenizer (zero-dep, replaces saxes)"
```

---

## SP-1c Acceptance Criteria

1. `src/xml/` contains `entities.ts`, `writer.ts`, `tokenizer.ts`, each with a colocated `*.test.ts`.
2. All modules are `node:`-free and dependency-free (core-purity guard passes).
3. The tokenizer handles the OOXML subset (elements, both quote styles, namespaced names, self-closing, text, CDATA, entities, numeric refs) and skips declaration/comments/PIs.
4. A round-trip test proves `tokenize` + `XmlWriter` reproduce equivalent XML.
5. All gates green: typecheck, lint (0/0), format, test, build, size.

---

## Self-Review

**1. Spec coverage** (SP-1 "Unit 2 — XML engine"):
- Streaming tokenizer (pull/SAX over OOXML subset) → Task 3. ✓
- Typed XML writer (escaping, namespaces) → Task 2. ✓
- Entity handling (five predefined + numeric refs) → Task 1 (shared by both). ✓
- Round-trip validation (feeds the conformance `diffParts` later) → Task 3 Step 1's final test. ✓
- Note: chunked tokenizing over a `ReadableStream` (true streaming for gigabyte files) is intentionally deferred to SP-7 — SP-1c tokenizes a complete string, which is what the OPC part reader (SP-1d) needs.

**2. Placeholder scan:** No TBD/"handle edge cases"/"similar to". Every step has complete code, an exact command, and an expected result. The tokenizer is fully written, not sketched.

**3. Type consistency:** `escapeText`/`escapeAttr`/`unescapeXml` signatures match between Task 1 (defined) and Tasks 2/3 (consumed). `XmlAttrs` and `XmlWriter`'s method signatures are consistent within Task 2 and reused in Task 3's round-trip test. `XmlToken` is consistent between Task 3's type definition, its tests, and the round-trip consumer. `tokenize`'s `Record<string,string>` attributes match `XmlWriter.open/leaf`'s `XmlAttrs` (string values are assignable to `string | number | undefined`).

Note for the executor: the tokenizer is the one piece with real edge-case risk. Make the prescribed tests pass first; if you discover an OOXML construct the tests don't cover (e.g. attribute values containing `>`), add a test and fix the tokenizer — but stay within the documented OOXML subset (no DTD/entity-definition/PI support).
