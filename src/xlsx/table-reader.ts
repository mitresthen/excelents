import type { TableDefinition } from '../model/table'
import { tokenize } from '../xml/tokenizer'

/** Parse one `xl/tables/tableN.xml` into a TableDefinition (inverse of writeTableXml). */
export function readTableXml(xml: string): TableDefinition {
  let name = ''
  let ref = ''
  let headerRowCount: string | undefined
  let styleName: string | undefined
  const columns: string[] = []

  for (const tok of tokenize(xml)) {
    if (tok.type !== 'open') continue
    if (tok.name === 'table') {
      name = tok.attributes['name'] ?? tok.attributes['displayName'] ?? ''
      ref = tok.attributes['ref'] ?? ''
      headerRowCount = tok.attributes['headerRowCount']
    } else if (tok.name === 'tableColumn') {
      const colName = tok.attributes['name']
      if (colName !== undefined) columns.push(colName)
    } else if (tok.name === 'tableStyleInfo') {
      styleName = tok.attributes['name']
    }
  }

  return {
    name,
    ref,
    columns,
    ...(headerRowCount === '0' ? { headerRow: false } : {}),
    ...(styleName !== undefined ? { styleName } : {}),
  }
}
