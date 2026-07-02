/**
 * ExcelJS-usage scanner: the engine behind `npx @mitresthen/excelents` (the
 * `excelents-migrate` bin). Pure — no filesystem access — so it is unit-testable
 * and importable from `@mitresthen/excelents/migrate` by other tooling (including
 * coding agents driving a migration).
 *
 * It deliberately reports rather than rewrites: the mechanical renames are listed
 * with their excelents equivalent, and the judgment-heavy sites (restructures,
 * dropped features) point into docs/MIGRATING-FROM-EXCELJS.md.
 */

/** How much human judgment a finding needs. */
export type MigrationCategory = 'map' | 'restructure' | 'blocked'

export interface MigrationFinding {
  file: string
  /** 1-based line number. */
  line: number
  ruleId: string
  category: MigrationCategory
  /** What was matched, e.g. "wb.xlsx.writeBuffer()". */
  found: string
  /** The excelents equivalent or what to do about it. */
  advice: string
  /** Anchor into docs/MIGRATING-FROM-EXCELJS.md. */
  guideSection: string
}

interface Rule {
  id: string
  category: MigrationCategory
  pattern: RegExp
  found: string
  advice: string
  guideSection: string
  /**
   * Generic member names (`.font =`, `.hidden =`, …) only match in files that
   * visibly reference exceljs; distinctive ones (`.addWorksheet(`) match anywhere,
   * catching helper modules that receive a worksheet as an argument.
   */
  needsContext?: boolean
}

const EXCELJS_CONTEXT = /['"]exceljs['"]|\bExcelJS\b/

const RULES: readonly Rule[] = [
  // --- mechanical renames -------------------------------------------------
  {
    id: 'import',
    category: 'map',
    pattern:
      /(?:from\s*['"]exceljs['"]|require\s*\(\s*['"]exceljs['"]\s*\)|import\s*\(\s*['"]exceljs['"]\s*\))/,
    found: "import/require of 'exceljs'",
    advice:
      "import { createWorkbook, writeXlsx, readXlsx } from '@mitresthen/excelents' (ESM-only)",
    guideSection: 'package--imports',
  },
  {
    id: 'new-workbook',
    category: 'map',
    pattern: /new\s+(?:[A-Za-z_$][\w$]*\.)?Workbook\s*\(/,
    found: 'new Workbook()',
    advice: 'createWorkbook()',
    guideSection: 'workbook--worksheets',
    needsContext: true,
  },
  {
    id: 'add-worksheet',
    category: 'map',
    pattern: /\.addWorksheet\s*\(/,
    found: 'wb.addWorksheet(...)',
    advice:
      'wb.addSheet(name) — the options argument (views, pageSetup, tabColor) has no equivalent',
    guideSection: 'workbook--worksheets',
  },
  {
    id: 'get-worksheet',
    category: 'map',
    pattern: /\.getWorksheet\s*\(/,
    found: 'wb.getWorksheet(...)',
    advice: 'wb.getSheet(name)',
    guideSection: 'workbook--worksheets',
  },
  {
    id: 'remove-worksheet',
    category: 'map',
    pattern: /\.removeWorksheet\s*\(/,
    found: 'wb.removeWorksheet(...)',
    advice: 'wb.removeSheet(name)',
    guideSection: 'workbook--worksheets',
  },
  {
    id: 'each-sheet',
    category: 'map',
    pattern: /\.eachSheet\s*\(/,
    found: 'wb.eachSheet(cb)',
    advice: 'iterate wb.sheets (a readonly array)',
    guideSection: 'workbook--worksheets',
  },
  {
    id: 'get-cell-address',
    category: 'map',
    pattern: /\.getCell\s*\(\s*['"`]/,
    found: "ws.getCell('A1')",
    advice: "ws.cell('A1') — numeric getCell(row, col) keeps its name",
    guideSection: 'cells--values',
    needsContext: true,
  },
  {
    id: 'get-column',
    category: 'map',
    pattern: /\.getColumn\s*\(/,
    found: 'ws.getColumn(n)',
    advice: 'ws.column(n)',
    guideSection: 'rows-columns--layout',
  },
  {
    id: 'merge-cells',
    category: 'map',
    pattern: /\.mergeCells\s*\(/,
    found: 'ws.mergeCells(range)',
    advice: 'ws.merge(range)',
    guideSection: 'rows-columns--layout',
  },
  {
    id: 'auto-filter',
    category: 'map',
    pattern: /\.autoFilter\s*=/,
    found: 'ws.autoFilter = ref',
    advice: 'ws.setAutoFilter(ref) — write-side only (not parsed back on read yet)',
    guideSection: 'rows-columns--layout',
    needsContext: true,
  },
  {
    id: 'views-frozen',
    category: 'map',
    pattern: /state\s*:\s*['"]frozen['"]/,
    found: "ws.views = [{ state: 'frozen', ... }]",
    advice: 'ws.freeze({ rows: ySplit, cols: xSplit }) — write-side only',
    guideSection: 'rows-columns--layout',
  },
  {
    id: 'num-fmt',
    category: 'map',
    pattern: /\.numFmt\s*=/,
    found: 'cell.numFmt = ...',
    advice: 'cell.style.numberFormat = ...',
    guideSection: 'styles',
    needsContext: true,
  },
  {
    id: 'font-assign',
    category: 'map',
    pattern: /\.font\s*=/,
    found: 'cell.font = {...}',
    advice:
      "cell.style.font = {...} — colors become plain ARGB strings ('FFRRGGBB'); row/column-level fonts are not modeled",
    guideSection: 'styles',
    needsContext: true,
  },
  {
    id: 'fill-assign',
    category: 'map',
    pattern: /\.fill\s*=/,
    found: 'cell.fill = {...}',
    advice:
      "cell.style.fill = {...} — pattern fills only ('solid' | 'none'), fgColor as ARGB string",
    guideSection: 'styles',
    needsContext: true,
  },
  {
    id: 'border-assign',
    category: 'map',
    pattern: /\.border\s*=/,
    found: 'cell.border = {...}',
    advice:
      'cell.style.border = {...} — edge styles limited to thin|medium|thick|dashed|dotted|double',
    guideSection: 'styles',
    needsContext: true,
  },
  {
    id: 'alignment-assign',
    category: 'map',
    pattern: /\.alignment\s*=/,
    found: 'cell.alignment = {...}',
    advice:
      'cell.style.alignment = {...} — horizontal left|center|right, vertical top|middle|bottom, wrapText, indent',
    guideSection: 'styles',
    needsContext: true,
  },
  {
    id: 'xlsx-write-buffer',
    category: 'map',
    pattern: /\.xlsx\.writeBuffer\s*\(/,
    found: 'wb.xlsx.writeBuffer()',
    advice: 'await writeXlsx(wb) → Uint8Array',
    guideSection: 'reading--writing-files',
  },
  {
    id: 'xlsx-write-file',
    category: 'map',
    pattern: /\.xlsx\.writeFile\s*\(/,
    found: 'wb.xlsx.writeFile(path)',
    advice:
      "nodeFileSystem.writeFile(path, await writeXlsx(wb)) — from '@mitresthen/excelents/node'",
    guideSection: 'reading--writing-files',
  },
  {
    id: 'xlsx-write-stream',
    category: 'map',
    pattern: /\.xlsx\.write\s*\(/,
    found: 'wb.xlsx.write(stream)',
    advice: 'writeXlsxStream(rows).pipeTo(writable) — or buffered writeXlsx(wb)',
    guideSection: 'streaming',
  },
  {
    id: 'csv-write',
    category: 'map',
    pattern: /\.csv\.write\w*\s*\(/,
    found: 'wb.csv.write*(...)',
    advice: "writeCsv(wb) from '@mitresthen/excelents/csv' → string (you handle the IO)",
    guideSection: 'csv',
  },
  {
    id: 'csv-read',
    category: 'map',
    pattern: /\.csv\.read\w*\s*\(/,
    found: 'wb.csv.read*(...)',
    advice: 'readCsv(text, options) → Workbook',
    guideSection: 'csv',
  },
  {
    id: 'add-image',
    category: 'map',
    pattern: /\.addImage\s*\(/,
    found: 'addImage(...)',
    advice:
      'wb.addImage({ data, extension }) + ws.placeImage(id, { tl, size }) — write-side only; anchors are cell refs',
    guideSection: 'images',
  },
  {
    id: 'defined-names',
    category: 'map',
    pattern: /\.definedNames\b/,
    found: 'wb.definedNames...',
    advice: 'wb.defineName(name, formula)',
    guideSection: 'tables-data-validation-defined-names',
  },
  {
    id: 'value-type-enum',
    category: 'map',
    pattern: /\bValueType\b/,
    found: 'ExcelJS.ValueType enum',
    advice: "cell.type is a string union now: 'string' | 'number' | 'date' | 'formula' | ...",
    guideSection: 'cells--values',
    needsContext: true,
  },

  // --- needs restructuring ------------------------------------------------
  {
    id: 'xlsx-load',
    category: 'restructure',
    pattern: /\.xlsx\.load\s*\(/,
    found: 'wb.xlsx.load(buffer)',
    advice: 'const wb = await readXlsx(bytes) — returns a NEW workbook instead of mutating one',
    guideSection: 'reading--writing-files',
  },
  {
    id: 'xlsx-read-file',
    category: 'restructure',
    pattern: /\.xlsx\.readFile\s*\(/,
    found: 'wb.xlsx.readFile(path)',
    advice:
      'const wb = await readXlsx(await nodeFileSystem.readFile(path)) — returns a new workbook',
    guideSection: 'reading--writing-files',
  },
  {
    id: 'data-validation',
    category: 'restructure',
    pattern: /\.dataValidation\s*=/,
    found: 'cell.dataValidation = {...}',
    advice:
      'validations attach to ranges, not cells: ws.addDataValidation({ sqref, type, formula1, ... }) — aggregate per range; formulae: [a, b] becomes formula1/formula2',
    guideSection: 'tables-data-validation-defined-names',
    needsContext: true,
  },
  {
    id: 'add-table',
    category: 'restructure',
    pattern: /\.addTable\s*\(/,
    found: 'ws.addTable({ ..., rows })',
    advice:
      'tables no longer carry row data — write the cells with addRow, then ws.addTable({ name, ref, columns })',
    guideSection: 'tables-data-validation-defined-names',
  },
  {
    id: 'stream-writer',
    category: 'restructure',
    pattern: /stream\.xlsx\.WorkbookWriter/,
    found: 'new ExcelJS.stream.xlsx.WorkbookWriter(...)',
    advice:
      'writeXlsxStream(rows) or createXlsxStreamWriter() — Web Streams with backpressure via `await addRow`; plain values, single sheet (no useStyles)',
    guideSection: 'streaming',
  },
  {
    id: 'stream-reader',
    category: 'restructure',
    pattern: /stream\.xlsx\.WorkbookReader/,
    found: 'new ExcelJS.stream.xlsx.WorkbookReader(...)',
    advice: 'for await (const row of readXlsxRows(bytes | Blob | ReadableStream))',
    guideSection: 'streaming',
  },
  {
    id: 'row-splicing',
    category: 'restructure',
    pattern: /\.(?:spliceRows|spliceColumns|insertRows?|duplicateRow)\s*\(/,
    found: 'row splicing/insertion',
    advice: 'not supported — build rows in their final order',
    guideSection: 'rows-columns--layout',
  },
  {
    id: 'columns-def',
    category: 'restructure',
    pattern: /\.columns\s*=/,
    found: 'ws.columns = [{ header, key, width }]',
    advice: 'no column keys — write headers with addRow and set ws.column(n).width directly',
    guideSection: 'rows-columns--layout',
    needsContext: true,
  },
  {
    id: 'add-row-object',
    category: 'restructure',
    pattern: /\.addRow\s*\(\s*\{/,
    found: 'ws.addRow({ key: value })',
    advice: 'object rows (by column key) are not supported — pass arrays of cell values',
    guideSection: 'cells--values',
    needsContext: true,
  },
  {
    id: 'views-assign',
    category: 'restructure',
    pattern: /\.views\s*=/,
    found: 'ws.views = [...]',
    advice:
      'only frozen panes map (ws.freeze({ rows, cols })); other view settings have no equivalent',
    guideSection: 'rows-columns--layout',
    needsContext: true,
  },

  // --- no equivalent ------------------------------------------------------
  {
    id: 'cell-notes',
    category: 'blocked',
    pattern: /\.note\s*=|\.addComment\b|\.getComment\b/,
    found: 'cell comments / notes',
    advice: 'dropped — no equivalent; content is lost on a read-modify-write cycle',
    guideSection: 'dropped-entirely-no-equivalent',
    needsContext: true,
  },
  {
    id: 'pivot-tables',
    category: 'blocked',
    pattern: /addPivotTable|pivotTable/i,
    found: 'pivot tables',
    advice: 'dropped — no equivalent',
    guideSection: 'dropped-entirely-no-equivalent',
  },
  {
    id: 'protection',
    category: 'blocked',
    pattern: /\.protect\s*\(|\.unprotect\s*\(/,
    found: 'sheet/workbook protection',
    advice: 'dropped — no equivalent',
    guideSection: 'dropped-entirely-no-equivalent',
  },
  {
    id: 'page-setup',
    category: 'blocked',
    pattern: /\.pageSetup\b/,
    found: 'page setup / printing',
    advice: 'dropped — no equivalent (print areas, margins, orientation, breaks)',
    guideSection: 'dropped-entirely-no-equivalent',
    needsContext: true,
  },
  {
    id: 'header-footer',
    category: 'blocked',
    pattern: /\.headerFooter\b/,
    found: 'headers/footers',
    advice: 'dropped — no equivalent',
    guideSection: 'dropped-entirely-no-equivalent',
    needsContext: true,
  },
  {
    id: 'conditional-formatting',
    category: 'blocked',
    pattern: /ConditionalFormatting\s*\(/,
    found: 'conditional formatting',
    advice: 'not supported yet — on the excelents roadmap',
    guideSection: 'dropped-entirely-no-equivalent',
  },
  {
    id: 'shared-formula',
    category: 'blocked',
    pattern: /sharedFormula/,
    found: 'shared formulas',
    advice: 'not modeled — write the full formula in each cell',
    guideSection: 'cells--values',
  },
  {
    id: 'hidden',
    category: 'blocked',
    pattern: /\.hidden\s*=/,
    found: 'hidden rows/columns/sheets',
    advice: 'not modeled',
    guideSection: 'rows-columns--layout',
    needsContext: true,
  },
  {
    id: 'outline',
    category: 'blocked',
    pattern: /\.outlineLevel/,
    found: 'row/column outlines',
    advice: 'not modeled',
    guideSection: 'dropped-entirely-no-equivalent',
    needsContext: true,
  },
  {
    id: 'tab-color',
    category: 'blocked',
    pattern: /tabColor/,
    found: 'sheet tab color',
    advice: 'not modeled',
    guideSection: 'dropped-entirely-no-equivalent',
    needsContext: true,
  },
  {
    id: 'unmerge',
    category: 'blocked',
    pattern: /\.unMergeCells\s*\(/,
    found: 'ws.unMergeCells(...)',
    advice: 'not supported — build the sheet without the merge instead',
    guideSection: 'rows-columns--layout',
  },
  {
    id: 'doc-props',
    category: 'blocked',
    pattern: /\.(?:creator|lastModifiedBy|lastPrinted)\s*=/,
    found: 'document properties',
    advice: 'not written (creator, modified, …) — most consumers never notice',
    guideSection: 'dropped-entirely-no-equivalent',
    needsContext: true,
  },
]

/** Scan one file's source. `file` is only used to label findings. */
export function scanSource(source: string, file: string): MigrationFinding[] {
  const hasContext = EXCELJS_CONTEXT.test(source)
  const findings: MigrationFinding[] = []
  const lines = source.split('\n')
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i]
    if (line === undefined || line.length > 2000) continue // skip minified lines
    let frozenMatched = false
    for (const rule of RULES) {
      if (rule.needsContext === true && !hasContext) continue
      if (!rule.pattern.test(line)) continue
      if (rule.id === 'views-frozen') frozenMatched = true
      if (rule.id === 'views-assign' && frozenMatched) continue // freeze advice already given
      findings.push({
        file,
        line: i + 1,
        ruleId: rule.id,
        category: rule.category,
        found: rule.found,
        advice: rule.advice,
        guideSection: rule.guideSection,
      })
    }
  }
  return findings
}

/** True when the file references exceljs at all (import, require, or global). */
export function referencesExcelJs(source: string): boolean {
  return EXCELJS_CONTEXT.test(source)
}
