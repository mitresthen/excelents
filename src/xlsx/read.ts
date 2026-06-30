import type { Workbook } from '../model/workbook'
import { createWorkbook } from '../model/workbook'
import { OpcPackage } from '../opc/package'
import { readWorkbookParts } from './workbook-reader'

/** Parse `.xlsx` bytes into a Workbook. */
export async function readXlsx(bytes: Uint8Array): Promise<Workbook> {
  const pkg = await OpcPackage.read(bytes)
  const parts = readWorkbookParts(pkg)
  const wb = createWorkbook()
  for (const sheet of parts.sheets) wb.addSheet(sheet.name)
  return wb
}
