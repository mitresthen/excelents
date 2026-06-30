import type { Workbook } from '../model/workbook'
import { XmlWriter } from '../xml/writer'

const MAIN_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
const REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const PKG_REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'
const OFFICE_DOC =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument'
const WORKSHEET_REL =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet'
const SHARED_STRINGS_REL =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings'
const STYLES_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles'
const HYPERLINK_REL =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'

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

/**
 * `xl/_rels/workbook.xml.rels` — workbook → each worksheet, then sharedStrings (if present),
 * then styles (always included — Excel requires it).
 */
export function writeWorkbookRelsXml(wb: Workbook, includeSharedStrings = false): string {
  const w = new XmlWriter().declaration().open('Relationships', { xmlns: PKG_REL_NS })
  wb.sheets.forEach((_sheet, i) => {
    w.leaf('Relationship', {
      Id: `rId${i + 1}`,
      Type: WORKSHEET_REL,
      Target: `worksheets/sheet${i + 1}.xml`,
    })
  })
  let nextRid = wb.sheets.length + 1
  if (includeSharedStrings) {
    w.leaf('Relationship', {
      Id: `rId${nextRid}`,
      Type: SHARED_STRINGS_REL,
      Target: 'sharedStrings.xml',
    })
    nextRid++
  }
  w.leaf('Relationship', {
    Id: `rId${nextRid}`,
    Type: STYLES_REL,
    Target: 'styles.xml',
  })
  w.close('Relationships')
  return w.toString()
}

/**
 * `xl/worksheets/_rels/sheetN.xml.rels` — a worksheet's external hyperlink targets.
 * Each entry is `{ rid, target }`; targets are written with `TargetMode="External"`.
 */
export function writeWorksheetRelsXml(
  hyperlinks: ReadonlyArray<{ rid: string; target: string }>,
): string {
  const w = new XmlWriter().declaration().open('Relationships', { xmlns: PKG_REL_NS })
  for (const { rid, target } of hyperlinks) {
    w.leaf('Relationship', {
      Id: rid,
      Type: HYPERLINK_REL,
      Target: target,
      TargetMode: 'External',
    })
  }
  return w.close('Relationships').toString()
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
