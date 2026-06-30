import { numberToCol } from '../utils/address'

/** Column metadata (1-based). */
export class Column {
  key?: string
  width?: number

  constructor(readonly number: number) {}

  get letter(): string {
    return numberToCol(this.number)
  }
}
