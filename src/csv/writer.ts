import type { CellValue } from '../model/cell'
import { Workbook } from '../model/workbook'
import type { Worksheet } from '../model/worksheet'
import type { CsvWriteOptions } from './options'

const BOM = String.fromCharCode(0xfeff)

function resolveSheet(
  source: Workbook | Worksheet,
  sheet: string | number | undefined,
): Worksheet | undefined {
  if (!(source instanceof Workbook)) return source
  if (typeof sheet === 'string') return source.getSheet(sheet)
  return source.sheets[typeof sheet === 'number' ? sheet : 0]
}

/** Render a cell value as its plain-text CSV representation (no quoting yet). */
function renderValue(v: CellValue): string {
  if (v === null) return ''
  if (typeof v === 'string') return v
  if (typeof v === 'number') return Number.isFinite(v) ? String(v) : ''
  if (typeof v === 'boolean') return v ? 'true' : 'false'
  if (v instanceof Date) return v.toISOString()
  if ('richText' in v) return v.richText.map((r) => r.text).join('')
  if ('hyperlink' in v) return v.text
  const result = v.result
  return result === undefined || result === null ? '' : renderValue(result)
}

function quoteField(field: string, delimiter: string, quoteAll: boolean): string {
  const mustQuote =
    quoteAll ||
    field.includes(delimiter) ||
    field.includes('"') ||
    field.includes('\n') ||
    field.includes('\r')
  return mustQuote ? `"${field.replaceAll('"', '""')}"` : field
}

/**
 * Serialize a worksheet (or a workbook's chosen sheet) to an RFC 4180 CSV string.
 * The grid spans `(1,1)` to the sheet's `(bottom, right)` dimension; absent cells are empty.
 *
 * Values are emitted verbatim — fields beginning with `=`, `+`, `-`, or `@` are NOT escaped
 * (this matches exceljs). If the CSV may be opened by a spreadsheet app and could contain
 * untrusted input, the caller is responsible for formula-injection sanitization.
 */
export function writeCsv(source: Workbook | Worksheet, options: CsvWriteOptions = {}): string {
  const { delimiter = ',', rowDelimiter = '\n', quoteAll = false, bom = false } = options
  const prefix = bom ? BOM : ''
  const ws = resolveSheet(source, options.sheet)
  const dims = ws?.dimensions
  if (ws === undefined || dims === undefined) return prefix

  const lines: string[] = []
  for (let r = 1; r <= dims.bottom; r++) {
    const fields: string[] = []
    for (let c = 1; c <= dims.right; c++) {
      fields.push(quoteField(renderValue(ws.getCell(r, c).value), delimiter, quoteAll))
    }
    lines.push(fields.join(delimiter))
  }
  return prefix + lines.join(rowDelimiter)
}
