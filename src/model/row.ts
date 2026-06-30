import { Cell } from './cell'

/** A worksheet row (1-based) owning its cells sparsely. */
export class Row {
  height?: number
  private readonly cellsByCol = new Map<number, Cell>()

  constructor(readonly number: number) {}

  getCell(col: number): Cell {
    let cell = this.cellsByCol.get(col)
    if (cell === undefined) {
      cell = new Cell(this.number, col)
      this.cellsByCol.set(col, cell)
    }
    return cell
  }

  get cells(): Cell[] {
    return [...this.cellsByCol.keys()].sort((a, b) => a - b).map((col) => this.cellsByCol.get(col)!)
  }
}
