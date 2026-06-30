import type { Workbook } from '../model/workbook'
import { OpcPackage } from '../opc/package'
import { SharedStrings } from '../utils/shared-strings'
import { CT } from './content-types'
import { writeSharedStringsXml } from './shared-strings-writer'
import { StyleRegistry, writeStylesXml } from './styles-writer'
import {
  writeRootRelsXml,
  writeWorkbookRelsXml,
  writeWorkbookXml,
  writeWorksheetRelsXml,
} from './workbook-writer'
import { writeTableXml } from './table-writer'
import { writeWorksheetXml } from './worksheet-writer'

const enc = (s: string): Uint8Array => new TextEncoder().encode(s)

/** Serialize a workbook to `.xlsx` bytes. */
export async function writeXlsx(wb: Workbook): Promise<Uint8Array> {
  const pkg = OpcPackage.empty()
  const sst = new SharedStrings()
  const registry = new StyleRegistry()

  // Serialize all worksheets first. This drives both interning passes: every cell
  // populates the SST and interns its style into the registry (xfIndexFor), so both
  // tables are complete before sharedStrings.xml and styles.xml are emitted below.
  // Each result also carries the sheet's external hyperlink relationships.
  const worksheetParts = wb.sheets.map((sheet, i) => {
    const result = writeWorksheetXml(sheet, sst, registry)
    return {
      index: i + 1,
      data: enc(result.xml),
      hyperlinks: result.hyperlinks,
      tables: result.tables,
    }
  })

  pkg.setPart('_rels/.rels', CT.rels, enc(writeRootRelsXml()))
  pkg.setPart('xl/workbook.xml', CT.workbook, enc(writeWorkbookXml(wb)))

  // Table parts are numbered workbook-globally (table1.xml, table2.xml, ...).
  let tableCounter = 0
  for (const { index, data, hyperlinks, tables } of worksheetParts) {
    pkg.setPart(`xl/worksheets/sheet${index}.xml`, CT.worksheet, data)

    const tableRels: Array<{ rid: string; target: string }> = []
    for (const { rid, table } of tables) {
      tableCounter++
      pkg.setPart(
        `xl/tables/table${tableCounter}.xml`,
        CT.table,
        enc(writeTableXml(table, tableCounter)),
      )
      tableRels.push({ rid, target: `../tables/table${tableCounter}.xml` })
    }

    if (hyperlinks.length > 0 || tableRels.length > 0) {
      pkg.setPart(
        `xl/worksheets/_rels/sheet${index}.xml.rels`,
        CT.rels,
        enc(writeWorksheetRelsXml(hyperlinks, tableRels)),
      )
    }
  }

  // styles.xml is always required; emit it before the rels so the part exists.
  pkg.setPart('xl/styles.xml', CT.styles, enc(writeStylesXml(registry)))

  // Emit sharedStrings.xml and its relationship only when there are strings.
  const hasStrings = sst.uniqueCount > 0
  pkg.setPart('xl/_rels/workbook.xml.rels', CT.rels, enc(writeWorkbookRelsXml(wb, hasStrings)))

  if (hasStrings) {
    pkg.setPart('xl/sharedStrings.xml', CT.sharedStrings, enc(writeSharedStringsXml(sst)))
  }

  return pkg.toBytes()
}
