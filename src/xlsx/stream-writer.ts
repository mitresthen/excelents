import { writeZipStream, type ZipStreamEntry } from '../opc/zip-stream'
import { encodeAddress } from '../utils/address'
import { dateToSerial } from '../utils/date'
import { XmlWriter } from '../xml/writer'
import { CT } from './content-types'
import { StyleRegistry, writeStylesXml } from './styles-writer'
import type { CellValue } from '../model/cell'

const MAIN_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
const REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const PKG_REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'
const OFFICE_DOC = `${REL_NS}/officeDocument`
const WORKSHEET_REL = `${REL_NS}/worksheet`
const STYLES_REL = `${REL_NS}/styles`
const DATE_FORMAT = 'mm-dd-yy'
const NUM_ERROR = '#NUM!'

/** One streamed row: a list of cell values (sparse cells render empty). */
export type StreamRow = readonly CellValue[]

export interface XlsxStreamWriteOptions {
  /** Worksheet name (default 'Sheet1'). */
  readonly sheet?: string
}

const enc = (s: string): Uint8Array => new TextEncoder().encode(s)

function contentTypesXml(): string {
  return new XmlWriter()
    .declaration()
    .open('Types', { xmlns: PKG_REL_NS.replace('relationships', 'content-types') })
    .leaf('Default', { Extension: 'rels', ContentType: CT.rels })
    .leaf('Default', { Extension: 'xml', ContentType: 'application/xml' })
    .leaf('Override', { PartName: '/xl/workbook.xml', ContentType: CT.workbook })
    .leaf('Override', { PartName: '/xl/worksheets/sheet1.xml', ContentType: CT.worksheet })
    .leaf('Override', { PartName: '/xl/styles.xml', ContentType: CT.styles })
    .close('Types')
    .toString()
}

function rootRelsXml(): string {
  return new XmlWriter()
    .declaration()
    .open('Relationships', { xmlns: PKG_REL_NS })
    .leaf('Relationship', { Id: 'rId1', Type: OFFICE_DOC, Target: 'xl/workbook.xml' })
    .close('Relationships')
    .toString()
}

function workbookXml(sheetName: string): string {
  return new XmlWriter()
    .declaration()
    .open('workbook', { xmlns: MAIN_NS, 'xmlns:r': REL_NS })
    .open('sheets')
    .leaf('sheet', { name: sheetName, sheetId: 1, 'r:id': 'rId1' })
    .close('sheets')
    .close('workbook')
    .toString()
}

function workbookRelsXml(): string {
  return new XmlWriter()
    .declaration()
    .open('Relationships', { xmlns: PKG_REL_NS })
    .leaf('Relationship', { Id: 'rId1', Type: WORKSHEET_REL, Target: 'worksheets/sheet1.xml' })
    .leaf('Relationship', { Id: 'rId2', Type: STYLES_REL, Target: 'styles.xml' })
    .close('Relationships')
    .toString()
}

/** Append one cell's XML to a row writer using inline strings (no shared-string table). */
function renderCell(w: XmlWriter, addr: string, value: CellValue, dateXf: number): void {
  if (value === null) return
  if (typeof value === 'string') {
    const space = value.trim() === value ? undefined : 'preserve'
    w.open('c', { r: addr, t: 'inlineStr' })
      .open('is')
      .open('t', space !== undefined ? { 'xml:space': space } : undefined)
      .text(value)
      .close('t')
      .close('is')
      .close('c')
    return
  }
  if (typeof value === 'number') {
    if (Number.isFinite(value)) {
      w.open('c', { r: addr }).open('v').text(String(value)).close('v').close('c')
    } else {
      w.open('c', { r: addr, t: 'e' }).open('v').text(NUM_ERROR).close('v').close('c')
    }
    return
  }
  if (typeof value === 'boolean') {
    w.open('c', { r: addr, t: 'b' })
      .open('v')
      .text(value ? '1' : '0')
      .close('v')
      .close('c')
    return
  }
  if (value instanceof Date) {
    w.open('c', { r: addr, s: dateXf })
      .open('v')
      .text(String(dateToSerial(value)))
      .close('v')
      .close('c')
    return
  }
  if ('richText' in value) {
    const text = value.richText.map((r) => r.text).join('')
    w.open('c', { r: addr, t: 'inlineStr' })
      .open('is')
      .open('t')
      .text(text)
      .close('t')
      .close('is')
      .close('c')
    return
  }
  if ('hyperlink' in value) {
    w.open('c', { r: addr, t: 'inlineStr' })
      .open('is')
      .open('t')
      .text(value.text)
      .close('t')
      .close('is')
      .close('c')
    return
  }
  // FormulaValue
  const result = value.result
  const t = typeof result === 'string' ? 'str' : typeof result === 'boolean' ? 'b' : undefined
  const rendered =
    typeof result === 'boolean'
      ? result
        ? '1'
        : '0'
      : result instanceof Date
        ? String(dateToSerial(result))
        : result === undefined || result === null
          ? undefined
          : String(result)
  w.open('c', { r: addr, t }).open('f').text(value.formula).close('f')
  if (rendered !== undefined) w.open('v').text(rendered).close('v')
  w.close('c')
}

async function* worksheetChunks(
  rows: AsyncIterable<StreamRow> | Iterable<StreamRow>,
  dateXf: number,
): AsyncGenerator<Uint8Array> {
  yield enc(
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<worksheet xmlns="${MAIN_NS}" xmlns:r="${REL_NS}"><sheetData>`,
  )
  let rowNum = 0
  for await (const row of rows) {
    rowNum++
    const w = new XmlWriter().open('row', { r: rowNum })
    row.forEach((value, i) => renderCell(w, encodeAddress(rowNum, i + 1), value, dateXf))
    w.close('row')
    yield enc(w.toString())
  }
  yield enc('</sheetData></worksheet>')
}

/** Serialize rows to a streamed `.xlsx` (inline strings, single sheet, bounded memory). */
export function writeXlsxStream(
  rows: AsyncIterable<StreamRow> | Iterable<StreamRow>,
  options: XlsxStreamWriteOptions = {},
): ReadableStream<Uint8Array> {
  const sheetName = options.sheet ?? 'Sheet1'
  const registry = new StyleRegistry()
  const dateXf = registry.xfIndexFor({ numberFormat: DATE_FORMAT })
  const stylesXml = writeStylesXml(registry)

  async function* entries(): AsyncGenerator<ZipStreamEntry> {
    yield { name: '[Content_Types].xml', data: [enc(contentTypesXml())] }
    yield { name: '_rels/.rels', data: [enc(rootRelsXml())] }
    yield { name: 'xl/workbook.xml', data: [enc(workbookXml(sheetName))] }
    yield { name: 'xl/_rels/workbook.xml.rels', data: [enc(workbookRelsXml())] }
    yield { name: 'xl/styles.xml', data: [enc(stylesXml)] }
    yield { name: 'xl/worksheets/sheet1.xml', data: worksheetChunks(rows, dateXf) }
  }
  return writeZipStream(entries())
}

/** A single-slot async hand-off: `put` blocks until the consumer takes the value (backpressure). */
class RowHandoff {
  private value: StreamRow | undefined
  private full = false
  private closed = false
  private wakeConsumer: (() => void) | null = null
  private wakeProducer: (() => void) | null = null

  async put(row: StreamRow): Promise<void> {
    while (this.full) await new Promise<void>((r) => (this.wakeProducer = r))
    this.value = row
    this.full = true
    this.wakeConsumer?.()
    this.wakeConsumer = null
    while (this.full) await new Promise<void>((r) => (this.wakeProducer = r))
  }

  close(): void {
    this.closed = true
    this.wakeConsumer?.()
    this.wakeConsumer = null
  }

  async *[Symbol.asyncIterator](): AsyncGenerator<StreamRow> {
    for (;;) {
      while (!this.full && !this.closed) await new Promise<void>((r) => (this.wakeConsumer = r))
      if (!this.full && this.closed) return
      const row = this.value!
      this.value = undefined
      this.full = false
      this.wakeProducer?.()
      this.wakeProducer = null
      yield row
    }
  }
}

/** A streamed `.xlsx` writer rows are pushed into imperatively; `readable` is the output. */
export interface XlsxStreamWriter {
  addRow(row: StreamRow): Promise<void>
  addRows(rows: Iterable<StreamRow>): Promise<void>
  close(): Promise<void>
  readonly readable: ReadableStream<Uint8Array>
}

/** Create a streamed `.xlsx` writer: push rows with `addRow`, pipe `readable`, then `close`. */
export function createXlsxStreamWriter(options: XlsxStreamWriteOptions = {}): XlsxStreamWriter {
  const channel = new RowHandoff()
  const readable = writeXlsxStream(channel, options)
  return {
    addRow: (row) => channel.put(row),
    async addRows(rows) {
      for (const row of rows) await channel.put(row)
    },
    close: async () => {
      channel.close()
    },
    readable,
  }
}
