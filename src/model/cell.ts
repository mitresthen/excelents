import { encodeAddress } from '../utils/address'
import type { CellStyle, Font } from './style'

export interface RichTextRun {
  text: string
  font?: Font
}

export interface FormulaValue {
  formula: string
  result?: string | number | boolean | Date | null
}

export interface RichTextValue {
  richText: RichTextRun[]
}

export interface HyperlinkValue {
  text: string
  hyperlink: string
}

export type CellValue =
  | null
  | string
  | number
  | boolean
  | Date
  | FormulaValue
  | RichTextValue
  | HyperlinkValue

export type CellType =
  | 'null'
  | 'string'
  | 'number'
  | 'boolean'
  | 'date'
  | 'formula'
  | 'richText'
  | 'hyperlink'

function detectType(value: CellValue): CellType {
  if (value === null) return 'null'
  if (typeof value === 'string') return 'string'
  if (typeof value === 'number') return 'number'
  if (typeof value === 'boolean') return 'boolean'
  if (value instanceof Date) return 'date'
  if ('formula' in value) return 'formula'
  if ('richText' in value) return 'richText'
  return 'hyperlink'
}

/** A single spreadsheet cell: a typed value plus an optional style. */
export class Cell {
  value: CellValue = null
  style: CellStyle = {}

  constructor(
    readonly row: number,
    readonly col: number,
  ) {}

  get address(): string {
    return encodeAddress(this.row, this.col)
  }

  get type(): CellType {
    return detectType(this.value)
  }
}
