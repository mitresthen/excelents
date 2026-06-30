import type { Workbook } from '../model/workbook'
import { createWorkbook } from '../model/workbook'
import { OpcPackage } from '../opc/package'
import { readSharedStrings, type SharedStringValue } from './shared-strings-reader'
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
  const ctx = { sharedStrings }

  const wb = createWorkbook()
  for (const sheet of parts.sheets) {
    const ws = wb.addSheet(sheet.name)
    readWorksheetInto(ws, decode(pkg.getPart(sheet.path)), ctx)
  }
  return wb
}
