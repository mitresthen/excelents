import type { Cell } from '../model/cell'
import type { Worksheet } from '../model/worksheet'
import type { SharedStrings } from '../utils/shared-strings'
import { encodeRange } from '../utils/range'
import { XmlWriter } from '../xml/writer'
import type { StyleRegistry } from './styles-writer'

const MAIN_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'

function writeCellValue(
  w: XmlWriter,
  cell: Cell,
  sst: SharedStrings,
  registry: StyleRegistry | undefined,
): void {
  const xfIndex = registry !== undefined ? registry.xfIndexFor(cell.style) : 0
  const s = xfIndex > 0 ? xfIndex : undefined
  const v = cell.value
  if (v === null) return
  if (typeof v === 'string') {
    const index = sst.add(v)
    w.open('c', { r: cell.address, t: 's', s }).open('v').text(String(index)).close('v').close('c')
    return
  }
  if (typeof v === 'number') {
    w.open('c', { r: cell.address, s }).open('v').text(String(v)).close('v').close('c')
    return
  }
  if (typeof v === 'boolean') {
    w.open('c', { r: cell.address, t: 'b', s })
      .open('v')
      .text(v ? '1' : '0')
      .close('v')
      .close('c')
    return
  }
  // Other types (date/formula/hyperlink/richText) land in Task 4; for the baseline,
  // serialize their string form so the file stays valid.
  const str = v instanceof Date ? v.toISOString() : JSON.stringify(v)
  const index = sst.add(str)
  w.open('c', { r: cell.address, t: 's', s }).open('v').text(String(index)).close('v').close('c')
}

/** Serialize one worksheet to `xl/worksheets/sheetN.xml`. */
export function writeWorksheetXml(
  ws: Worksheet,
  sst: SharedStrings,
  registry?: StyleRegistry,
): string {
  const w = new XmlWriter().declaration().open('worksheet', { xmlns: MAIN_NS })
  const dims = ws.dimensions
  if (dims !== undefined) w.leaf('dimension', { ref: encodeRange(dims) })
  w.open('sheetData')
  for (const row of ws.rows) {
    w.open('row', { r: row.number })
    for (const cell of row.cells) writeCellValue(w, cell, sst, registry)
    w.close('row')
  }
  w.close('sheetData').close('worksheet')
  return w.toString()
}
