# ExcelJS → `excelents` Modernization — Architecture & Roadmap

- **Date:** 2026-06-29
- **Status:** Approved (brainstorming) — drives sub-project specs
- **Repo:** `mitresthen/excelents` (fork of `exceljs@4.4.0`)

## 1. Context & motivation

`exceljs` is effectively unmaintained upstream (last stable `4.4.0`, Oct 2023). The
current codebase is ~20,400 LOC of pure CommonJS (155 files `require`, 0 `import`,
151 files already using ES6 `class`), built with a legacy Grunt + Babel + Browserify +
Terser chain, tested with Mocha/Chai/Jasmine, and shipping a hand-maintained 48 KB
`index.d.ts` that is decoupled from the source (the root cause of weak type-safety).

We are doing a **ground-up rewrite** as a new, TypeScript-first, universal,
tree-shakeable library. The name remains `excelents`. The public API is **redesigned**
(we break compatibility with `exceljs`) and ships with a documented migration path.

## 2. Goals & non-goals

**Goals**

- TypeScript-first: all types **derived from source**, never hand-written.
- Universal runtime: Node 20+, browsers, and edge/Deno/Bun — web-standard core.
- Tree-shakeable: subpath exports, `sideEffects: false`, free-function serialization.
- **Zero runtime dependencies** by default.
- Nothing abandoned: every dependency is latest/maintained, or we own it.
- Correctness preserved via a conformance oracle (the old library + real fixtures).

**Non-goals**

- Backward API compatibility with `exceljs` (explicitly out; migration guide instead).
- Features we are dropping (see §4).
- A monorepo / multiple published packages (single package — see §8).

## 3. Runtime targets

- **Node 20+**, modern **browsers**, **edge/Deno/Bun**.
- Core uses only web-standard APIs: **Web Streams**, **`CompressionStream`/
  `DecompressionStream`** (`deflate-raw`), **`crypto.randomUUID`**, `TextEncoder`/
  `TextDecoder`.
- No `node:` built-ins in the core. Node-specific conveniences (filesystem, Node
  stream bridging) live behind an **injected adapter** exposed at `excelents/node`.

## 4. Scope

**In:**

- xlsx **writing** (cells, base styling: fonts/fills/borders/number formats/alignment,
  formulas, merged cells, column/row sizing).
- xlsx **reading** (parse `.xlsx` → object model).
- **CSV** read/write.
- **Streaming** engine (read/write large workbooks without full in-memory buffering,
  over Web Streams).
- Advanced features: **rich text + hyperlinks**, **images/drawings**, **data
  validation**, **tables**, **conditional formatting**, **defined names**.

**Out (dropped):**

- Cell **comments/notes**.
- **Pivot tables**.
- **Encryption** (reading/writing password-protected workbooks).

## 5. High-level architecture (layers)

```
public API  (treeshake-friendly: model + free-function codecs)
    │
object model  (Workbook / Worksheet / Row / Column / Cell / Range / styles)
    │
xlsx codecs  (write + read; feature modules layered on top)        CSV codec
    │                                                                  │
universal substrate  (OPC container · XML engine · IO/streams · utils)
    │
web-standard platform  (Web Streams · CompressionStream · crypto · Text*)
```

The **substrate** is web-standard and owns the ZIP/XML/stream plumbing. The **xlsx
codecs** translate between the object model and OOXML parts. The **public API** is
designed so an app that only writes never pulls the reader into its bundle.

## 6. Decomposition into sub-projects

A from-scratch universal rewrite is far too large for one spec. It is decomposed into
independently-buildable sub-projects, each with its own spec → plan → implementation
cycle.

| # | Sub-project | Purpose | Depends on |
|---|---|---|---|
| **SP-0** | Toolchain & scaffold | New repo layout, build/test/lint/format, `package.json` exports, CI. Empty but releasable skeleton + conformance harness. | — |
| **SP-1** | Universal substrate | OPC container (ZIP + content-types + relationships) over native `CompressionStream` + owned container; owned streaming XML tokenizer + typed XML writer; Web Streams plumbing; injected FS adapter; core utils. | SP-0 |
| **SP-2** | Object model + public API | In-memory model **and** the new TS-first public API. No serialization yet. | SP-1 |
| **SP-3** | xlsx write (baseline) | Object model → valid `.xlsx` with base styling; validated via the oracle. | SP-2 |
| **SP-4** | xlsx read (baseline) | Parse `.xlsx` → object model; round-trips with SP-3. | SP-3 |
| **SP-5** | Feature modules | Self-contained serialize+parse slices: rich text/hyperlinks, images/drawings, data validation, tables, conditional formatting, defined names. | SP-4 |
| **SP-6** | CSV | Owned Web-Streams CSV read/write. | SP-2 |
| **SP-7** | Streaming engine | Streaming write + read over Web Streams (no temp files in core). | SP-4 |
| **SP-8** | Packaging & release | Subpath exports finalization, treeshaking/bundle budgets, docs, migration guide from exceljs, v1 release. | all |

**Build order:** SP-0 → SP-1 → SP-2 → SP-3 → SP-4 → SP-5, with SP-6 (CSV) and SP-7
(streaming) branching after SP-2/SP-4, and SP-8 closing out.

This document is paired with the detailed **SP-0 + SP-1 foundation spec**
(`2026-06-29-excelents-sp0-sp1-foundation-design.md`).

## 7. Cross-cutting decision: toolchain (VoidZero / "Vite+" stack)

Researched against the live npm registry and primary sources on 2026-06-29.

| Concern | Decision | Notes |
|---|---|---|
| Typecheck | **TypeScript 7** native (`@typescript/native-preview`, pin a dated build; move to `typescript@rc`/GA when available), run as `tsc --noEmit` | Also *required* by oxlint type-aware |
| Shipped `.d.ts` | tsdown + **`isolatedDeclarations: true` → oxc-transform** | Decouples dts from the TS version; sidesteps tsgo's immature dts emit. Requires explicit return types on exports |
| Build | **tsdown** (Rolldown): `platform:'neutral'`, `format:['esm','cjs']`, `dts:true`, `exports:true` | `exports:true` auto-generates the `package.json` exports map |
| Test | **Vitest 4** (+ `@vitest/browser`) | Browser Mode (stable in v4) for the browser checks |
| Lint | **oxlint `--type-aware`** via **`oxlint-tsgolint`** (pinned) | The correct package is `oxlint-tsgolint`, **not** the stale standalone `tsgolint` |
| Format | **oxfmt** (pinned); **no Prettier** | Passes 100% of Prettier's JS/TS conformance, ~30× faster |
| Type backstop | `tsc --noEmit` (TS7) in CI; **attw** doubles as downstream TS-version compat check | Replaces a typescript-eslint net |
| Publish correctness | **publint** + **@arethetypeswrong/cli** (ci-only) | Catch dual ESM/CJS "masquerading types" bugs |
| Bundle budget | tsdown `--report` + a small **owned** CI gzip check (Node `zlib`) | No `size-limit` (redundant esbuild tree) |
| Pkg mgr / host | **pnpm 11**; build/CI host Node **≥ 22.18** | Output targets Node 20+/browser/edge |
| Umbrella CLI | **None** (no Vite+ `vp`) | `vp` is pre-1.0; `vp pack` is just tsdown; adopt later if this becomes a monorepo |

**Notes that shaped the above:**

- TS7 does **not** come free with the linter: `oxlint-tsgolint` statically embeds its
  own pinned `typescript-go`, independent of the installed `typescript`. TS7 is a
  separate, explicitly-pinned install — justified because type-aware lint needs it and
  native `tsc` is fast.
- All preview/0.x tools (TS7 native-preview, oxfmt, oxlint-tsgolint) are **pinned**;
  expect more frequent dependency bumps than a Prettier/ESLint stack — acceptable for a
  greenfield library.

## 8. Cross-cutting decision: dependency policy & packaging

- **Zero runtime dependencies.** Confirmed achievable: Rolldown inlines TS helpers (no
  hidden `tslib`), the dts plugin is build-time only, and nothing but our own code
  reaches `dist`. **`fflate` must never be added to `dependencies`** — it is an
  optional/vendored codec escape hatch only (single-maintainer, low-cadence project).
- **Owned subsystems** (no dependency, fully tested via the oracle): streaming XML
  tokenizer (OOXML subset), ZIP container (local headers, central directory, ZIP64),
  CSV codec, serial-date math.
- **Native platform** for: DEFLATE (`CompressionStream`/`DecompressionStream`,
  `deflate-raw`), UUID (`crypto.randomUUID`), streams (Web Streams), encoding
  (`TextEncoder`/`TextDecoder`).
- **Pluggable codec escape hatch:** the core accepts an optional injectable
  `deflate`/`inflate` codec, defaulting to native `CompressionStream`. Runtimes lacking
  `deflate-raw` can inject `fflate`.
- **Single published package** `excelents` with **subpath exports**:
  - `.` → universal core
  - `./node` → Node FS/stream adapter
  - `./csv` → CSV module
  - `type: module`, `sideEffects: false`, dual ESM/CJS + `.d.ts`. The `exports` map is
    auto-generated by tsdown (`exports:true`), not hand-maintained.

## 9. Cross-cutting decision: public API direction

Finalized in SP-2; committed here as a direction so SP-1 boundaries point correctly:

- **Model construction** via a lightweight builder/model.
- **Serialization as free functions** (`writeXlsx(workbook, sink)`,
  `readXlsx(source)`), so unused codecs tree-shake away.
- Named exports only; no default-export namespace object (it defeats treeshaking).

## 10. Cross-cutting decision: conformance oracle & testing

The rewrite stays safe only if every serialize/parse path is checked against
known-good output.

- **Oracle:** `exceljs@4.4.0` pinned as a **devDependency** (stale + CJS + Node-only;
  acceptable because it never ships and is used only in tests).
- **35 real `.xlsx` fixtures + 1 `.csv`** carried over from the current repo.
- **Write path:** build the same workbook with old `exceljs` and new `excelents` →
  unzip both → normalize → **diff XML parts**. Divergence fails the test.
- **Read path:** parse each fixture with both → compare resulting models.
- **Bundle budget:** per-entry gzipped size asserted in CI (owned `zlib` check).
- **Publish correctness:** `publint` + `attw` in CI.
- Test runner: **Vitest 4**; browser-mode subset for the browser build.

## 11. Risks & mitigations

| Risk | Mitigation |
|---|---|
| Re-deriving OOXML correctness from scratch (the long pole) | Conformance oracle (XML-diff + fixture round-trips) gates every codec path; build feature-by-feature |
| Preview/0.x tooling churn (TS7, oxfmt, oxlint-tsgolint) | Pin exact versions; `tsc --noEmit` backstop; tooling is dev-only and swappable |
| `CompressionStream` `deflate-raw` gaps on some edge runtimes | Pluggable codec; documented `fflate` fallback |
| `isolatedDeclarations` requires explicit export types | Accepted as authoring discipline; enforced by the compiler |
| Downstream consumers on older TypeScript | `attw` + a downstream type-compat test (TS 5.x/6) on the shipped `.d.ts` |
| Owned XML/ZIP/CSV bugs | Heavy fixture coverage via the oracle; these are well-specified, bounded formats |

## 12. Deferred to sub-project specs

- Exact public API shape and ergonomics (SP-2).
- Styles/number-format model details (SP-2/SP-3).
- Streaming engine backpressure model (SP-7).
- CSV dialect options (SP-6).
- Migration-guide mapping table from `exceljs` (SP-8).
