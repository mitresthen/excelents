import { Worksheet } from './worksheet'

/** A workbook: an ordered collection of named worksheets. */
export class Workbook {
  private readonly worksheets: Worksheet[] = []

  addSheet(name: string): Worksheet {
    const ws = new Worksheet(name)
    this.worksheets.push(ws)
    return ws
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
