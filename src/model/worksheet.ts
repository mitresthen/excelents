import { decodeAddress } from '../utils/address'
import { decodeRange, type RangeBox } from '../utils/range'
import type { Cell, CellValue } from './cell'
import { Column } from './column'
import type { DataValidation } from './data-validation'
import type { ImageAnchor, ImagePlacement } from './image'
import { Row } from './row'
import type { TableDefinition } from './table'

/** A worksheet: a sparse grid of cells, plus columns, merges, and a name. */
export class Worksheet {
  private readonly rowsMap = new Map<number, Row>()
  private readonly cols = new Map<number, Column>()
  private readonly mergeRanges: string[] = []
  private readonly validations: DataValidation[] = []
  private readonly tableDefs: TableDefinition[] = []
  private autoFilterRef: string | undefined
  private frozenSplit: { rows: number; cols: number } | undefined
  private readonly imagePlacements: ImagePlacement[] = []

  constructor(public name: string) {}

  getRow(n: number): Row {
    let row = this.rowsMap.get(n)
    if (row === undefined) {
      row = new Row(n)
      this.rowsMap.set(n, row)
    }
    return row
  }

  getCell(row: number, col: number): Cell {
    return this.getRow(row).getCell(col)
  }

  cell(ref: string): Cell {
    const { row, col } = decodeAddress(ref)
    return this.getCell(row, col)
  }

  addRow(values: CellValue[]): Row {
    const row = this.getRow(this.rowCount + 1)
    values.forEach((value, i) => {
      row.getCell(i + 1).value = value
    })
    return row
  }

  column(n: number): Column {
    let col = this.cols.get(n)
    if (col === undefined) {
      col = new Column(n)
      this.cols.set(n, col)
    }
    return col
  }

  /** Defined columns (those touched via `column(n)`), in ascending column order. */
  get columns(): readonly Column[] {
    return [...this.cols.keys()].sort((a, b) => a - b).map((n) => this.cols.get(n)!)
  }

  /**
   * Merge a range, e.g. `'A1:F1'`. Every covered cell is materialized and
   * SHARES the master's (top-left) style object, so styling the master —
   * before or after merging — styles the whole range. This matches exceljs:
   * spreadsheet apps paint each underlying cell of a merged range, so a fill
   * carried only by the master would render over one grid position.
   */
  merge(range: string): void {
    this.mergeRanges.push(range)
    const box = decodeRange(range)
    const master = this.getCell(box.top, box.left)
    for (let r = box.top; r <= box.bottom; r += 1) {
      for (let c = box.left; c <= box.right; c += 1) {
        if (r === box.top && c === box.left) continue
        this.getCell(r, c).style = master.style
      }
    }
  }

  get merges(): readonly string[] {
    return this.mergeRanges
  }

  addDataValidation(rule: DataValidation): void {
    this.validations.push(rule)
  }

  get dataValidations(): readonly DataValidation[] {
    return this.validations
  }

  addTable(def: TableDefinition): void {
    this.tableDefs.push(def)
  }

  get tables(): readonly TableDefinition[] {
    return this.tableDefs
  }

  /** Set the sheet's AutoFilter range, e.g. `'A9:F123'`. */
  setAutoFilter(ref: string): void {
    this.autoFilterRef = ref
  }

  get autoFilter(): string | undefined {
    return this.autoFilterRef
  }

  /** Freeze `rows` header rows and/or `cols` leading columns (each defaults to 0). */
  freeze(opts: { rows?: number; cols?: number }): void {
    this.frozenSplit = { rows: opts.rows ?? 0, cols: opts.cols ?? 0 }
  }

  get frozen(): { rows: number; cols: number } | undefined {
    return this.frozenSplit
  }

  /** Place a workbook-registered image (by `addImage` id) at an anchor on this sheet. */
  placeImage(imageId: number, anchor: ImageAnchor): void {
    this.imagePlacements.push({ imageId, ...anchor })
  }

  get images(): readonly ImagePlacement[] {
    return this.imagePlacements
  }

  /** The highest row number that has been touched (created via getRow/getCell), used to place addRow. */
  get rowCount(): number {
    let max = 0
    for (const n of this.rowsMap.keys()) {
      if (n > max) max = n
    }
    return max
  }

  /** Populated rows (at least one cell) in ascending order. */
  get rows(): readonly Row[] {
    return [...this.rowsMap.keys()]
      .sort((a, b) => a - b)
      .map((n) => this.rowsMap.get(n)!)
      .filter((row) => row.cells.length > 0)
  }

  get dimensions(): RangeBox | undefined {
    let top = Infinity
    let left = Infinity
    let bottom = -Infinity
    let right = -Infinity
    let any = false
    for (const row of this.rowsMap.values()) {
      for (const cell of row.cells) {
        any = true
        if (cell.row < top) top = cell.row
        if (cell.row > bottom) bottom = cell.row
        if (cell.col < left) left = cell.col
        if (cell.col > right) right = cell.col
      }
    }
    return any ? { top, left, bottom, right } : undefined
  }
}
