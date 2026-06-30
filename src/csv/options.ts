/** Options for {@link writeCsv}. */
export interface CsvWriteOptions {
  /** Field delimiter (default ','). */
  readonly delimiter?: string
  /** Row delimiter (default '\n'). */
  readonly rowDelimiter?: string
  /** Quote every field, not just those that require it (default false). */
  readonly quoteAll?: boolean
  /** Prepend a UTF-8 BOM (U+FEFF) for Excel compatibility (default false). */
  readonly bom?: boolean
  /** When given a Workbook, which sheet to emit (name or 0-based index; default first). */
  readonly sheet?: string | number
}

/** Options for {@link readCsv}. */
export interface CsvReadOptions {
  /** Field delimiter (default ','). */
  readonly delimiter?: string
  /** Name for the single produced worksheet (default 'Sheet1'). */
  readonly sheetName?: string
  /** Infer round-trip-safe numbers (default true). */
  readonly parseNumbers?: boolean
  /** Infer 'true'/'false' as booleans (default true). */
  readonly parseBooleans?: boolean
}
