import type { CellValue } from '../model/cell'
import { createWorkbook, type Workbook } from '../model/workbook'
import type { CsvReadOptions } from './options'

/**
 * Split CSV text into rows of raw string fields (RFC 4180).
 *
 * Parsing is lenient: an unterminated quoted field absorbs to end-of-text, and characters
 * after a field's closing quote but before the delimiter are appended (`"a"b` → `ab`) rather
 * than throwing. This favors recovering data from slightly-malformed input over strictness.
 */
function parseRows(text: string, delimiter: string): string[][] {
  const rows: string[][] = []
  let field = ''
  let row: string[] = []
  let inQuotes = false
  const pushField = (): void => {
    row.push(field)
    field = ''
  }
  const pushRow = (): void => {
    pushField()
    rows.push(row)
    row = []
  }
  for (let i = 0; i < text.length; i++) {
    const ch = text[i]!
    if (inQuotes) {
      if (ch === '"') {
        if (text[i + 1] === '"') {
          field += '"'
          i++
        } else inQuotes = false
      } else field += ch
      continue
    }
    if (ch === '"' && field === '') inQuotes = true
    else if (delimiter.length === 1 ? ch === delimiter : text.startsWith(delimiter, i)) {
      pushField()
      i += delimiter.length - 1 // no-op for a single-char delimiter
    } else if (ch === '\n') pushRow()
    else if (ch === '\r') {
      // \r\n: swallow here, the \n ends the row. Bare \r (old-Mac) ends the row itself.
      if (text[i + 1] !== '\n') pushRow()
    } else field += ch
  }
  // Flush a trailing partial row, but not a spurious empty one from a final newline.
  if (field !== '' || row.length > 0) pushRow()
  return rows
}

/** Conservative inference: round-trip-safe numbers and optional booleans; else keep the string. */
export function infer(s: string, parseNumbers: boolean, parseBooleans: boolean): CellValue {
  if (s === '') return null
  // String(Number(s)) === s keeps leading-zero ('00210') and date-like strings as strings.
  // Number.isFinite rejects 'NaN'/'Infinity'/'-Infinity' (which pass the round-trip test but
  // the writer cannot reproduce — it emits '' for non-finite numbers), so they stay strings.
  if (parseNumbers && s.trim() !== '' && Number.isFinite(Number(s)) && String(Number(s)) === s) {
    return Number(s)
  }
  if (parseBooleans && (s === 'true' || s === 'false')) return s === 'true'
  return s
}

/** Parse CSV text into a one-sheet Workbook (inverse of writeCsv). */
export function readCsv(text: string, options: CsvReadOptions = {}): Workbook {
  const {
    delimiter = ',',
    sheetName = 'Sheet1',
    parseNumbers = true,
    parseBooleans = true,
  } = options
  const clean = text.startsWith(String.fromCharCode(0xfeff)) ? text.slice(1) : text
  const rows = parseRows(clean, delimiter)

  const wb = createWorkbook()
  const ws = wb.addSheet(sheetName)
  rows.forEach((fields, r) => {
    fields.forEach((raw, c) => {
      const value = infer(raw, parseNumbers, parseBooleans)
      if (value !== null) ws.getCell(r + 1, c + 1).value = value
    })
  })
  return wb
}
