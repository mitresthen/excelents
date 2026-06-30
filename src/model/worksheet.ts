import { decodeAddress } from '../utils/address'
import type { RangeBox } from '../utils/range'
import type { Cell, CellValue } from './cell'
import { Column } from './column'
import type { DataValidation } from './data-validation'
import { Row } from './row'

/** A worksheet: a sparse grid of cells, plus columns, merges, and a name. */
export class Worksheet {
  private readonly rowsMap = new Map<number, Row>()
  private readonly cols = new Map<number, Column>()
  private readonly mergeRanges: string[] = []
  private readonly validations: DataValidation[] = []

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

  merge(range: string): void {
    this.mergeRanges.push(range)
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
