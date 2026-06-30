import type { Workbook } from '../model/workbook'
import { OpcPackage } from '../opc/package'
import { CT } from './content-types'
import { writeRootRelsXml, writeWorkbookRelsXml, writeWorkbookXml } from './workbook-writer'
import { writeWorksheetXml } from './worksheet-writer'

const enc = (s: string): Uint8Array => new TextEncoder().encode(s)

/** Serialize a workbook to `.xlsx` bytes. */
export async function writeXlsx(wb: Workbook): Promise<Uint8Array> {
  const pkg = OpcPackage.empty()
  pkg.setPart('_rels/.rels', CT.rels, enc(writeRootRelsXml()))
  pkg.setPart('xl/workbook.xml', CT.workbook, enc(writeWorkbookXml(wb)))
  pkg.setPart('xl/_rels/workbook.xml.rels', CT.rels, enc(writeWorkbookRelsXml(wb)))
  wb.sheets.forEach((sheet, i) => {
    pkg.setPart(`xl/worksheets/sheet${i + 1}.xml`, CT.worksheet, enc(writeWorksheetXml(sheet)))
  })
  return pkg.toBytes()
}
