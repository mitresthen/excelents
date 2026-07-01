import { base64ToBytes } from '../utils/base64'
import type { WorkbookImage, WorkbookImageInput } from './image'
import { Worksheet } from './worksheet'

/** A workbook: an ordered collection of named worksheets. */
export class Workbook {
  private readonly worksheets: Worksheet[] = []
  private readonly names: Array<{ name: string; formula: string; localSheetId?: number }> = []
  private readonly mediaList: WorkbookImage[] = []

  addSheet(name: string): Worksheet {
    const ws = new Worksheet(name)
    this.worksheets.push(ws)
    return ws
  }

  /**
   * Register an image on the workbook and return its id (a stable index). Accepts raw bytes
   * or a base64 string. Place it on a sheet with `worksheet.placeImage(id, anchor)`.
   */
  addImage(image: WorkbookImageInput): number {
    const data = typeof image.data === 'string' ? base64ToBytes(image.data) : image.data
    this.mediaList.push({ data, extension: image.extension })
    return this.mediaList.length - 1
  }

  /** Registered images, indexed by the id returned from `addImage`. */
  get media(): readonly WorkbookImage[] {
    return this.mediaList
  }

  /**
   * Define a name mapping `name` to a formula/reference (e.g. `Sheet1!$A$1`). Omit
   * `localSheetId` for a workbook-global name; pass a 0-based sheet index to scope it to
   * that sheet (a scoped name may share its label with a global one without colliding).
   */
  defineName(name: string, formula: string, localSheetId?: number): void {
    this.names.push(
      localSheetId === undefined ? { name, formula } : { name, formula, localSheetId },
    )
  }

  get definedNames(): ReadonlyArray<{ name: string; formula: string; localSheetId?: number }> {
    return this.names
  }

  getSheet(name: string): Worksheet | undefined {
    return this.worksheets.find((ws) => ws.name === name)
  }

  removeSheet(name: string): void {
    const i = this.worksheets.findIndex((ws) => ws.name === name)
    if (i !== -1) this.worksheets.splice(i, 1)
  }

  get sheets(): readonly Worksheet[] {
    return this.worksheets
  }
}

/** Create an empty workbook. */
export function createWorkbook(): Workbook {
  return new Workbook()
}
