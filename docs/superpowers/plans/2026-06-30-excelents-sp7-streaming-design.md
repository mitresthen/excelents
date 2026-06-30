# SP-7: Streaming — Design Note (pre-plan)

> **Status:** DESIGN / pre-plan. SP-7's design is genuinely open (it adds new architecture, not a settled pattern like SP-3..SP-6), and the go-ahead vs. cutting a release is a pending user decision. This note captures the architecture, the web-standard API options, the hard constraints, and a proposed scope so the direction can be approved/redirected **before** a full bite-sized plan + build. No code is written yet.

## Why streaming

`writeXlsx`/`readXlsx`/`writeCsv`/`readCsv` all build the entire workbook in memory. For large spreadsheets (hundreds of MB, millions of rows) that is the dominant cost. Streaming lets a caller **produce** rows and have bytes flushed incrementally, and **consume** rows as they are parsed — bounded memory regardless of file size. This is the single feature most asked of `exceljs` (its `stream.xlsx.WorkbookWriter` / `WorkbookReader`).

## What already helps us

- `src/io/codec.ts` wraps the **web-standard** `CompressionStream('deflate-raw')` / `DecompressionStream('deflate-raw')`. These are real `TransformStream`s — the streaming deflate/inflate primitive already exists; today's `nativeCodec` just buffers around it. A streaming layer can use them as streams.
- The OOXML serializers (`worksheet-writer`, `styles-writer`, etc.) and the tokenizer are mostly pure string transforms — reusable per-row.
- `ReadableStream` / `WritableStream` / async iterables are web-standard and already the project's idiom (universal runtime goal).

## What's missing (the actual work)

The whole stack buffers: `writeZip(entries)` takes every entry as a full `Uint8Array` and returns one `Uint8Array`; `readZip(bytes)` consumes the whole archive; `OpcPackage` holds all parts in a `Map`. Streaming needs **new entry points**, not edits to these.

## Hard constraints to design around

1. **ZIP central directory is at the END.** Streaming *write* is fine: emit `local header → streamed deflated data → data descriptor` per entry (data descriptor carries crc32 + sizes that aren't known until the entry finishes), then the central directory last. Streaming *read* from a forward-only source is the hard direction — you cannot know entry offsets without the central directory at the tail. Practical answer: streaming **read** still needs random access to the bytes (a `Uint8Array` or a `Blob`/file with seek), but can stream the *parse* of `sheetData` (the one huge part) row-by-row. True forward-only network streaming read is out of scope.
2. **crc32 over streamed data.** We have `src/opc/crc32.ts`; it must support incremental updates (feed chunks) for the data-descriptor approach. Verify/extend it.
3. **Shared strings vs inline strings.** A streaming writer cannot buffer a shared-string table across the whole sheet. Use **inline strings** (`t="inlineStr"`, already supported by the writer) so each row is self-contained. (Trade-off: larger files; acceptable and standard for streaming writers.)
4. **Styles** must be registered up front (a streaming writer takes a known style set, or interns styles into a small in-memory table emitted at the end — the styles part is tiny even for huge sheets, so buffering it is fine).

## Proposed public API (web-standard)

Separate entry, e.g. `./stream` (tree-shakeable; pulls model + a streaming-zip layer, not the buffered codecs):

**Write** — a builder that yields bytes as rows are added:
```ts
const writer = createXlsxStreamWriter({ sheet: 'Data', columns? , styles? })
await writer.addRow(['a', 1, true])           // backpressure-aware
await writer.addRows(asyncIterableOfRows)
const bytes: ReadableStream<Uint8Array> = writer.readable  // pipe to a file/response
await writer.close()
```
or the functional form `writeXlsxStream(rows: AsyncIterable<Row>, opts): ReadableStream<Uint8Array>`.

**Read** — async-iterate rows without building the model:
```ts
for await (const { sheet, rowNumber, cells } of readXlsxRows(bytes /* Uint8Array */, opts)) { ... }
```

**CSV (easy companion)** — `writeCsvStream(rows): ReadableStream<Uint8Array>` and `readCsvRows(textStream): AsyncIterable<Row>`, reusing the SP-6 quoting/inference per row.

## Proposed scope

- **IN (high value, tractable):**
  - Streaming **xlsx write** — the killer feature; forward ZIP with data descriptors + incremental crc32 + inline strings + streamed deflate. Buffers only the small parts (workbook/styles/content-types/rels).
  - Streaming **xlsx read of `sheetData`** over in-memory bytes — async-iterable rows, SAX over the one big part (random access to the zip for the central directory; stream only the parse).
  - Streaming **CSV** read/write — cheap, reuses SP-6.
- **OUT (defer / later plan):** true forward-only network-streamed *read* (blocked by central-dir-at-end); streaming of drawings/images; multi-sheet streaming write in a single pass (start with one active sheet at a time).

## Rough task breakdown (for the eventual plan)

1. Incremental `crc32` (feed chunks) + a streaming-zip **write** primitive (local header / data descriptor / central dir) emitting a `ReadableStream<Uint8Array>`, deflating each entry through `CompressionStream`. TDD: the produced bytes unzip via the existing `readZip` and via Node's `unzip`.
2. `createXlsxStreamWriter` / `writeXlsxStream` — drive the streaming zip with a streamed worksheet part (inline strings) + buffered small parts. TDD: exceljs reads the streamed output; round-trips vs `readXlsx`.
3. `readXlsxRows` — locate `sheetData` via the zip central directory, SAX-parse rows incrementally into lightweight row objects. TDD: matches `readXlsx` cell-for-cell on fixtures; bounded memory.
4. Streaming CSV (`writeCsvStream` / `readCsvRows`) reusing SP-6. TDD: round-trip + matches `writeCsv`/`readCsv`.
5. `./stream` entry + size/tree-shake (must not pull the buffered xlsx codecs unnecessarily) + conformance.

## Open questions for the user (why this is a checkpoint, not a silent build)

1. **Do you want streaming now, or cut a release of SP-1..SP-6 first?** The library is already releasable.
2. **API shape:** builder (`createXlsxStreamWriter().addRow()`) vs functional (`writeXlsxStream(asyncIterable)`)? I lean toward offering **both** (the functional one wraps the builder).
3. **Read scope:** is async-iterable read **over bytes** (random-access) enough, or do you specifically need forward-only network-streamed read (much harder, separate effort)?
4. **Inline strings** for streamed writes (larger files, no shared-string table) — acceptable? (Standard for streaming writers; yes unless you object.)

Once direction is confirmed, this becomes a full bite-sized `writing-plans` plan and follows the same TDD + adversarial-review flow as SP-3..SP-6.
