import type { Workbook } from '../model/workbook'
import { XmlWriter } from '../xml/writer'

const MAIN_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
const REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const PKG_REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'
const OFFICE_DOC =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument'
const WORKSHEET_REL =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet'

/** `xl/workbook.xml` — the sheet list. */
export function writeWorkbookXml(wb: Workbook): string {
  const w = new XmlWriter().declaration().open('workbook', { xmlns: MAIN_NS, 'xmlns:r': REL_NS })
  w.open('sheets')
  wb.sheets.forEach((sheet, i) => {
    w.leaf('sheet', { name: sheet.name, sheetId: i + 1, 'r:id': `rId${i + 1}` })
  })
  w.close('sheets').close('workbook')
  return w.toString()
}

/** `xl/_rels/workbook.xml.rels` — workbook → each worksheet. */
export function writeWorkbookRelsXml(wb: Workbook): string {
  const w = new XmlWriter().declaration().open('Relationships', { xmlns: PKG_REL_NS })
  wb.sheets.forEach((_sheet, i) => {
    w.leaf('Relationship', {
      Id: `rId${i + 1}`,
      Type: WORKSHEET_REL,
      Target: `worksheets/sheet${i + 1}.xml`,
    })
  })
  w.close('Relationships')
  return w.toString()
}

/** `_rels/.rels` — package → workbook. */
export function writeRootRelsXml(): string {
  return new XmlWriter()
    .declaration()
    .open('Relationships', { xmlns: PKG_REL_NS })
    .leaf('Relationship', { Id: 'rId1', Type: OFFICE_DOC, Target: 'xl/workbook.xml' })
    .close('Relationships')
    .toString()
}
