import type { TableDefinition } from '../model/table'
import { XmlWriter } from '../xml/writer'

const MAIN_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'

/** Serialize one table to `xl/tables/tableN.xml`. `id` is the workbook-global table id. */
export function writeTableXml(table: TableDefinition, id: number): string {
  const w = new XmlWriter().declaration().open('table', {
    xmlns: MAIN_NS,
    id,
    name: table.name,
    displayName: table.name,
    ref: table.ref,
    // headerRowCount defaults to 1; emit 0 only when the caller suppresses the header row.
    headerRowCount: table.headerRow === false ? 0 : undefined,
  })
  w.leaf('autoFilter', { ref: table.ref })
  w.open('tableColumns', { count: table.columns.length })
  table.columns.forEach((name, i) => {
    w.leaf('tableColumn', { id: i + 1, name })
  })
  w.close('tableColumns')
  // tableStyleInfo is optional; emit it only for an explicitly-styled table so an
  // unstyled table round-trips without acquiring a phantom style name.
  if (table.styleName !== undefined) {
    w.leaf('tableStyleInfo', { name: table.styleName, showRowStripes: 1 })
  }
  return w.close('table').toString()
}
