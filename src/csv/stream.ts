import type { CellValue } from '../model/cell'
import type { CsvReadOptions, CsvWriteOptions } from './options'
import { infer } from './reader'
import { quoteField, renderValue } from './writer'

/** One streamed CSV row: cell values, rendered/inferred exactly like {@link writeCsv}/{@link readCsv}. */
export type CsvStreamRow = readonly CellValue[]

/**
 * Serialize rows to a streamed RFC 4180 CSV (`ReadableStream<Uint8Array>`), one row per pull so a
 * slow consumer applies backpressure. Byte-for-byte identical to `writeCsv` for the same grid:
 * the row delimiter is written *between* rows (no trailing newline).
 */
export function writeCsvStream(
  rows: AsyncIterable<CsvStreamRow> | Iterable<CsvStreamRow>,
  options: CsvWriteOptions = {},
): ReadableStream<Uint8Array> {
  const { delimiter = ',', rowDelimiter = '\n', quoteAll = false, bom = false } = options
  const encoder = new TextEncoder()
  const iterator: AsyncIterator<CsvStreamRow> | Iterator<CsvStreamRow> =
    Symbol.asyncIterator in rows ? rows[Symbol.asyncIterator]() : rows[Symbol.iterator]()
  let first = true
  let started = false
  return new ReadableStream<Uint8Array>({
    async pull(controller) {
      if (!started) {
        started = true
        if (bom) controller.enqueue(encoder.encode('﻿'))
      }
      const result = await iterator.next()
      if (result.done === true) {
        controller.close()
        return
      }
      const line = result.value
        .map((value) => quoteField(renderValue(value), delimiter, quoteAll))
        .join(delimiter)
      controller.enqueue(encoder.encode(first ? line : rowDelimiter + line))
      first = false
    },
    async cancel() {
      await iterator.return?.()
    },
  })
}

type DelimiterMatch = 'yes' | 'no' | 'defer'

/** Is there a delimiter at `buf[i]`? `defer` = a partial match runs off the end, await more input. */
function matchDelimiter(buf: string, i: number, delimiter: string, final: boolean): DelimiterMatch {
  if (delimiter.length === 1) return buf[i] === delimiter ? 'yes' : 'no'
  if (buf.startsWith(delimiter, i)) return 'yes'
  const remaining = buf.slice(i)
  if (remaining.length < delimiter.length && delimiter.startsWith(remaining)) {
    return final ? 'no' : 'defer'
  }
  return 'no'
}

/**
 * Stream CSV rows from bytes or a string (inverse of {@link writeCsvStream}), reusing the SP-6
 * inference. Parsing is single-pass with bounded memory: only the current row and a tiny
 * lookahead tail are held. Matches `readCsv` cell-for-cell, including quoted fields that span
 * chunk boundaries, `\r\n`/bare-`\r`/`\n` terminators, and multi-character delimiters.
 */
export async function* readCsvRows(
  src: ReadableStream<Uint8Array> | string,
  options: CsvReadOptions = {},
): AsyncGenerator<CsvStreamRow, void, unknown> {
  const { delimiter = ',', parseNumbers = true, parseBooleans = true } = options

  let buf = ''
  let field = ''
  let row: string[] = []
  let inQuotes = false
  let bomChecked = false

  const valuesOf = (raw: readonly string[]): CsvStreamRow =>
    raw.map((s) => infer(s, parseNumbers, parseBooleans))

  const stripBom = (): void => {
    if (bomChecked || buf.length === 0) return
    bomChecked = true
    if (buf.charCodeAt(0) === 0xfeff) buf = buf.slice(1)
  }

  // Process `buf` from the start, yielding completed rows; mutate the shared field/row/quote
  // state. In non-final mode, stop at an ambiguous tail (lone trailing `"`/`\r`, partial
  // delimiter) and leave it buffered for the next chunk.
  function* consume(final: boolean): Generator<CsvStreamRow> {
    let i = 0
    while (i < buf.length) {
      const ch = buf[i]!
      if (inQuotes) {
        if (ch === '"') {
          if (i + 1 >= buf.length) {
            if (!final) break
            inQuotes = false
            i++
          } else if (buf[i + 1] === '"') {
            field += '"'
            i += 2
          } else {
            inQuotes = false
            i++
          }
        } else {
          field += ch
          i++
        }
        continue
      }
      if (ch === '"' && field === '') {
        inQuotes = true
        i++
        continue
      }
      const dm = matchDelimiter(buf, i, delimiter, final)
      if (dm === 'defer') break
      if (dm === 'yes') {
        row.push(field)
        field = ''
        i += delimiter.length
        continue
      }
      if (ch === '\n') {
        row.push(field)
        field = ''
        yield valuesOf(row)
        row = []
        i++
        continue
      }
      if (ch === '\r') {
        if (i + 1 >= buf.length && !final) break
        row.push(field)
        field = ''
        yield valuesOf(row)
        row = []
        i += i + 1 < buf.length && buf[i + 1] === '\n' ? 2 : 1
        continue
      }
      field += ch
      i++
    }
    buf = buf.slice(i)
    if (final) {
      // Flush a trailing partial row, but not a spurious empty one from a final terminator.
      if (field !== '' || row.length > 0) {
        row.push(field)
        field = ''
        yield valuesOf(row)
        row = []
      }
      buf = ''
    }
  }

  if (typeof src === 'string') {
    buf = src
    stripBom()
    yield* consume(true)
    return
  }

  const decoder = new TextDecoder('utf-8', { ignoreBOM: true })
  const reader = src.getReader()
  try {
    for (;;) {
      const { done, value } = await reader.read()
      if (done) break
      if (value !== undefined && value.length > 0) {
        buf += decoder.decode(value, { stream: true })
        stripBom()
        yield* consume(false)
      }
    }
    buf += decoder.decode()
    stripBom()
    yield* consume(true)
  } finally {
    reader.releaseLock()
  }
}
