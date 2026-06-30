import type { Workbook } from '../model/workbook'
import { OpcPackage } from '../opc/package'
import { SharedStrings } from '../utils/shared-strings'
import { CT } from './content-types'
import { writeSharedStringsXml } from './shared-strings-writer'
import { StyleRegistry, writeStylesXml } from './styles-writer'
import { writeRootRelsXml, writeWorkbookRelsXml, writeWorkbookXml } from './workbook-writer'
import { writeWorksheetXml } from './worksheet-writer'

const enc = (s: string): Uint8Array => new TextEncoder().encode(s)

/** Serialize a workbook to `.xlsx` bytes. */
export async function writeXlsx(wb: Workbook): Promise<Uint8Array> {
  const pkg = OpcPackage.empty()
  const sst = new SharedStrings()
  const registry = new StyleRegistry()

  // Pre-pass: intern every cell style so all xf indices are assigned before
  // worksheet serialization (which looks up indices from the same registry).
  for (const sheet of wb.sheets) {
    for (const row of sheet.rows) {
      for (const cell of row.cells) {
        registry.xfIndexFor(cell.style)
      }
    }
  }

  // Serialize all worksheets (with style references) so the SST is fully
  // populated before we emit it.
  const worksheetParts = wb.sheets.map((sheet, i) => ({
    name: `xl/worksheets/sheet${i + 1}.xml`,
    data: enc(writeWorksheetXml(sheet, sst, registry)),
  }))

  pkg.setPart('_rels/.rels', CT.rels, enc(writeRootRelsXml()))
  pkg.setPart('xl/workbook.xml', CT.workbook, enc(writeWorkbookXml(wb)))

  for (const { name, data } of worksheetParts) {
    pkg.setPart(name, CT.worksheet, data)
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
