import { decodeAddress } from '../utils/address'
import type { RangeBox } from '../utils/range'
import type { Cell, CellValue } from './cell'
import { Column } from './column'
import { Row } from './row'

/** A worksheet: a sparse grid of cells, plus columns, merges, and a name. */
export class Worksheet {
  private readonly rows = new Map<number, Row>()
  private readonly cols = new Map<number, Column>()
  private readonly mergeRanges: string[] = []

  constructor(public name: string) {}

  getRow(n: number): Row {
    let row = this.rows.get(n)
    if (row === undefined) {
      row = new Row(n)
      this.rows.set(n, row)
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

  merge(range: string): void {
    this.mergeRanges.push(range)
  }

  get merges(): readonly string[] {
    return this.mergeRanges
  }

  /** The highest row number that has been touched (created via getRow/getCell), used to place addRow. */
  get rowCount(): number {
    let max = 0
    for (const n of this.rows.keys()) {
      if (n > max) max = n
    }
    return max
  }

  get dimensions(): RangeBox | undefined {
    let top = Infinity
    let left = Infinity
    let bottom = -Infinity
    let right = -Infinity
    let any = false
    for (const row of this.rows.values()) {
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
