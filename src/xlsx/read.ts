import type { Workbook } from '../model/workbook'
import { createWorkbook } from '../model/workbook'
import { OpcPackage } from '../opc/package'
import { readSharedStrings, type SharedStringValue } from './shared-strings-reader'
import { readStyles } from './styles-reader'
import { readTableXml } from './table-reader'
import { readWorkbookParts } from './workbook-reader'
import { readWorksheetInto } from './worksheet-reader'

const decode = (bytes: Uint8Array | undefined): string =>
  bytes === undefined ? '' : new TextDecoder().decode(bytes)

/** Parse `.xlsx` bytes into a Workbook. */
export async function readXlsx(bytes: Uint8Array): Promise<Workbook> {
  const pkg = await OpcPackage.read(bytes)
  const parts = readWorkbookParts(pkg)

  const sharedStrings: SharedStringValue[] =
    parts.sharedStringsPath !== undefined
      ? readSharedStrings(decode(pkg.getPart(parts.sharedStringsPath)))
      : []
  const cellStyles =
    parts.stylesPath !== undefined
      ? readStyles(decode(pkg.getPart(parts.stylesPath))).cellStyles
      : []

  const wb = createWorkbook()
  for (const sheet of parts.sheets) {
    const ws = wb.addSheet(sheet.name)
    // A worksheet's relationships (hyperlinks, tables) live in its own rels part.
    const sheetRels = pkg.relationshipsFor(sheet.path)
    const hyperlinkTargets = new Map<string, string>()
    for (const rel of sheetRels) {
      if (rel.targetMode === 'External' || rel.type.endsWith('/hyperlink')) {
        hyperlinkTargets.set(rel.id, rel.target)
      }
    }
    const tableRids = readWorksheetInto(ws, decode(pkg.getPart(sheet.path)), {
      sharedStrings,
      cellStyles,
      hyperlinkTargets,
    })
    // Resolve <tablePart> rIds to their table parts and reconstruct each table.
    const relById = new Map(sheetRels.map((rel) => [rel.id, rel]))
    for (const rid of tableRids) {
      const rel = relById.get(rid)
      if (rel === undefined || !rel.type.endsWith('/table')) continue
      const tableXml = decode(pkg.getPart(rel.target))
      if (tableXml !== '') ws.addTable(readTableXml(tableXml))
    }
  }
  for (const { name, formula, localSheetId } of parts.definedNames) {
    wb.defineName(name, formula, localSheetId)
  }
  return wb
}
