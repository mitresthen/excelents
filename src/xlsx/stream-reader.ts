import { readableToBytes } from '../io/streams'
import type { CellValue, FormulaValue } from '../model/cell'
import type { CellStyle } from '../model/style'
import { OpcPackage } from '../opc/package'
import { type ZipEntryLocation, inflateEntry, readZipEntries } from '../opc/zip-reader'
import { decodeAddress } from '../utils/address'
import { serialToDate } from '../utils/date'
import { isDateFormat } from '../utils/number-format'
import { tokenize } from '../xml/tokenizer'
import { type SharedStringValue, readSharedStrings } from './shared-strings-reader'
import { readStyles } from './styles-reader'
import { readWorkbookParts } from './workbook-reader'

/** One streamed worksheet row. */
export interface XlsxRow {
  /** Worksheet display name. */
  readonly sheet: string
  /** 1-based row number (from the row's `r` attribute, else its first cell's reference). */
  readonly rowNumber: number
  /** Values by column: index 0 = column A; absent cells are `null`; trailing nulls are omitted. */
  readonly cells: readonly CellValue[]
}

/** The supported random-access sources. A `ReadableStream`/`Blob` is buffered to bytes first. */
export type XlsxRowSource = Uint8Array | Blob | ReadableStream<Uint8Array>

const decode = (bytes: Uint8Array | undefined): string =>
  bytes === undefined ? '' : new TextDecoder().decode(bytes)

async function sourceToBytes(source: XlsxRowSource): Promise<Uint8Array> {
  if (source instanceof Uint8Array) return source
  if (source instanceof ReadableStream) return readableToBytes(source)
  return new Uint8Array(await source.arrayBuffer()) // Blob
}

function sharedToValue(value: SharedStringValue | undefined): CellValue {
  if (value === undefined) return ''
  if (value.kind === 'plain') return value.text
  return { richText: value.runs }
}

/** Inflate one worksheet entry and decode it to text incrementally (never the whole sheet at once). */
async function* inflateText(entry: ZipEntryLocation): AsyncGenerator<string> {
  if (entry.method === 0) {
    yield new TextDecoder().decode(entry.body)
    return
  }
  if (entry.method !== 8) throw new Error(`Unsupported ZIP compression method ${entry.method}`)
  const ds = new DecompressionStream('deflate-raw')
  const writer = ds.writable.getWriter()
  // Feed the compressed body and read the decompressed output concurrently (a sequential
  // write-then-read would deadlock once the writer's buffer fills).
  const pump = (async (): Promise<void> => {
    await writer.write(new Uint8Array(entry.body))
    await writer.close()
  })()
  const reader = ds.readable.getReader()
  const decoder = new TextDecoder()
  try {
    for (;;) {
      const { done, value } = await reader.read()
      if (done) break
      if (value !== undefined && value.length > 0) yield decoder.decode(value, { stream: true })
    }
    const tail = decoder.decode()
    if (tail.length > 0) yield tail
  } finally {
    await pump
  }
}

const ROW_DELIMS = new Set([' ', '\t', '\n', '\r', '>', '/'])

/** Index of the next `<row` element start at/after `from` (not `<rowBreaks` etc.), or -1. */
function findRowStart(s: string, from: number): number {
  let idx = s.indexOf('<row', from)
  while (idx !== -1) {
    const after = s[idx + 4]
    if (after === undefined) return -1 // a partial `<row` at the buffer end — wait for more
    if (ROW_DELIMS.has(after)) return idx
    idx = s.indexOf('<row', idx + 4)
  }
  return -1
}

/** Exclusive end index of the complete `<row>…</row>` (or self-closing `<row/>`) at `start`, or -1. */
function rowEnd(s: string, start: number): number {
  let inQuote = false
  let quote = ''
  let tagEnd = -1
  let selfClosing = false
  for (let i = start + 4; i < s.length; i++) {
    const c = s[i]!
    if (inQuote) {
      if (c === quote) inQuote = false
      continue
    }
    if (c === '"' || c === "'") {
      inQuote = true
      quote = c
    } else if (c === '>') {
      tagEnd = i
      selfClosing = s[i - 1] === '/'
      break
    }
  }
  if (tagEnd === -1) return -1
  if (selfClosing) return tagEnd + 1
  const close = s.indexOf('</row>', tagEnd + 1)
  return close === -1 ? -1 : close + 6
}

/** Yield complete `<row …>…</row>` substrings from a stream of decompressed text chunks. */
async function* extractRowXml(chunks: AsyncIterable<string>): AsyncGenerator<string> {
  let buf = ''
  for await (const chunk of chunks) {
    buf += chunk
    let pos = 0
    for (;;) {
      const start = findRowStart(buf, pos)
      if (start === -1) break
      const end = rowEnd(buf, start)
      if (end === -1) {
        pos = start // incomplete row — keep it buffered for the next chunk
        break
      }
      yield buf.slice(start, end)
      pos = end
    }
    buf = buf.slice(pos)
  }
}

/** Parse one row's XML into a row number and column-indexed cell values (matches readXlsx). */
function parseRow(
  xml: string,
  sharedStrings: readonly SharedStringValue[],
  cellStyles: readonly CellStyle[],
  date1904: boolean,
): { rowNumber: number; cells: CellValue[] } {
  let rowNumber = 0
  const cellByCol = new Map<number, CellValue>()
  let maxCol = 0

  let ref: string | undefined
  let type = 'n'
  let v: string | undefined
  let f: string | undefined
  let sIndex: number | undefined
  let inlineText: string | undefined
  let inV = false
  let inF = false
  let inIs = false
  let inT = false

  const finalize = (): void => {
    if (ref === undefined) return
    const { row: cellRow, col } = decodeAddress(ref)
    if (rowNumber === 0) rowNumber = cellRow
    const style = sIndex !== undefined ? cellStyles[sIndex] : undefined

    let value: CellValue | undefined
    if (f !== undefined) {
      let result: FormulaValue['result']
      if (type === 'str') result = v
      else if (type === 'b') result = v === '1'
      else if (type !== 'e' && v !== undefined) result = Number(v)
      value = result === undefined ? { formula: f } : { formula: f, result }
    } else if (type === 'e') {
      value = undefined // error cells have no model representation
    } else if (type === 's' && v !== undefined) {
      value = sharedToValue(sharedStrings[Number(v)])
    } else if (type === 'inlineStr' && inlineText !== undefined) {
      value = inlineText
    } else if (type === 'str' && v !== undefined) {
      value = v
    } else if (type === 'b' && v !== undefined) {
      value = v === '1'
    } else if (v !== undefined) {
      value =
        style?.numberFormat !== undefined && isDateFormat(style.numberFormat)
          ? serialToDate(Number(v), { date1904 })
          : Number(v)
    }
    if (value !== undefined) {
      cellByCol.set(col, value)
      if (col > maxCol) maxCol = col
    }
  }

  for (const tok of tokenize(xml)) {
    if (tok.type === 'open') {
      if (tok.name === 'row') {
        const r = tok.attributes['r']
        if (r !== undefined) rowNumber = Number(r)
      } else if (tok.name === 'c') {
        finalize()
        ref = tok.attributes['r']
        type = tok.attributes['t'] ?? 'n'
        const s = tok.attributes['s']
        sIndex = s !== undefined ? Number(s) : undefined
        v = undefined
        f = undefined
        inlineText = undefined
        inV = false
        inF = false
        inIs = false
        if (tok.selfClosing) ref = undefined
      } else if (tok.name === 'v') inV = !tok.selfClosing
      else if (tok.name === 'f') inF = !tok.selfClosing
      else if (tok.name === 'is') inIs = !tok.selfClosing
      else if (tok.name === 't') inT = !tok.selfClosing
    } else if (tok.type === 'close') {
      if (tok.name === 'v') inV = false
      else if (tok.name === 'f') inF = false
      else if (tok.name === 'is') inIs = false
      else if (tok.name === 't') inT = false
      else if (tok.name === 'c') {
        finalize()
        ref = undefined
      }
    } else if (tok.type === 'text') {
      if (inV) v = (v ?? '') + tok.value
      else if (inF) f = (f ?? '') + tok.value
      else if (inIs && inT) inlineText = (inlineText ?? '') + tok.value
    }
  }
  finalize()

  const cells: CellValue[] = []
  for (let c = 1; c <= maxCol; c++) cells.push(cellByCol.get(c) ?? null)
  return { rowNumber, cells }
}

/**
 * Stream the rows of an `.xlsx` without building the workbook model. Uses ONE random-access
 * engine: the source is buffered to bytes, the ZIP central directory locates parts, `sharedStrings`
 * and styles are resolved first (regardless of their physical order in the archive), then each
 * worksheet's `sheetData` is inflated and SAX-parsed row-by-row — bounded per-row memory, never
 * holding a whole decompressed sheet. Yields rows across every sheet in workbook order, and
 * matches `readXlsx` cell-for-cell (it does not apply the post-`sheetData` hyperlink rewrite).
 */
export async function* readXlsxRows(source: XlsxRowSource): AsyncGenerator<XlsxRow, void, unknown> {
  const bytes = await sourceToBytes(source)
  const entries = readZipEntries(bytes)
  const byName = new Map(entries.map((entry) => [entry.name, entry]))
  const decompress = async (name: string): Promise<Uint8Array | undefined> => {
    const entry = byName.get(name)
    return entry === undefined ? undefined : inflateEntry(entry)
  }

  // Assemble just the small structural parts (content-types + every *.rels) so the package can be
  // resolved without decompressing the large worksheet bodies.
  const structural = new Map<string, Uint8Array>()
  for (const entry of entries) {
    if (
      entry.name === '[Content_Types].xml' ||
      entry.name === '_rels/.rels' ||
      (entry.name.includes('/_rels/') && entry.name.endsWith('.rels'))
    ) {
      structural.set(entry.name, await inflateEntry(entry))
    }
  }
  const probe = OpcPackage.fromEntries(structural)
  const workbookRel = probe.rootRelationships().find((r) => r.type.endsWith('/officeDocument'))
  const workbookPath = (workbookRel?.target ?? 'xl/workbook.xml').replace(/^\//, '')
  const workbookBytes = await decompress(workbookPath)
  if (workbookBytes !== undefined) structural.set(workbookPath, workbookBytes)

  const parts = readWorkbookParts(OpcPackage.fromEntries(structural))

  const sharedStrings =
    parts.sharedStringsPath !== undefined
      ? readSharedStrings(decode(await decompress(parts.sharedStringsPath)))
      : []
  const cellStyles =
    parts.stylesPath !== undefined
      ? readStyles(decode(await decompress(parts.stylesPath))).cellStyles
      : []

  for (const sheet of parts.sheets) {
    const entry = byName.get(sheet.path)
    if (entry === undefined) continue
    for await (const rowXml of extractRowXml(inflateText(entry))) {
      const { rowNumber, cells } = parseRow(rowXml, sharedStrings, cellStyles, parts.date1904)
      yield { sheet: sheet.name, rowNumber, cells }
    }
  }
}
