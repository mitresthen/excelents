/** The package version, kept in sync with package.json (asserted by src/index.test.ts). */
export const version: string = '0.0.0'
export { createWorkbook, Workbook } from './model/workbook'
export { writeXlsx } from './xlsx/write'
export { readXlsx } from './xlsx/read'
export { Worksheet } from './model/worksheet'
export { Row } from './model/row'
export { Column } from './model/column'
export { Cell } from './model/cell'
export type {
  CellValue,
  CellType,
  RichTextRun,
  FormulaValue,
  RichTextValue,
  HyperlinkValue,
} from './model/cell'
export type { CellStyle, Font, Fill, Borders, BorderEdge, Alignment } from './model/style'
export type { DataValidation } from './model/data-validation'
export type { TableDefinition } from './model/table'
