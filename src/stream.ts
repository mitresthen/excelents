// Streaming entry: bounded-memory, web-standard read/write for large spreadsheets. Kept free of
// the buffered codecs (writeXlsx/readXlsx build the whole workbook in memory) so `./stream`
// tree-shakes to just the streaming surface.
export {
  createXlsxStreamWriter,
  writeXlsxStream,
  type StreamRow,
  type XlsxStreamWriteOptions,
  type XlsxStreamWriter,
} from './xlsx/stream-writer'
export { readXlsxRows, type XlsxRow, type XlsxRowSource } from './xlsx/stream-reader'
export { readCsvRows, writeCsvStream, type CsvStreamRow } from './csv/stream'
export type { CsvReadOptions, CsvWriteOptions } from './csv/options'
