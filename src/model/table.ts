/** A structured table over a worksheet range (header row + columns). */
export interface TableDefinition {
  /** Unique table name, e.g. `'Table1'` (used as both name and displayName). */
  readonly name: string
  /** The table's range including the header row, e.g. `'A1:C4'`. */
  readonly ref: string
  /** Column header names, left to right. */
  readonly columns: readonly string[]
  /** Whether the first row is a header row (default true). */
  readonly headerRow?: boolean
  /** A built-in table style, e.g. `'TableStyleMedium2'`. Omitted = workbook default. */
  readonly styleName?: string
}
