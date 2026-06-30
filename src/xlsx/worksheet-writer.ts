import type { Cell, FormulaValue } from '../model/cell'
import type { CellStyle } from '../model/style'
import type { Worksheet } from '../model/worksheet'
import { dateToSerial } from '../utils/date'
import { encodeRange } from '../utils/range'
import type { SharedStrings } from '../utils/shared-strings'
import { XmlWriter } from '../xml/writer'
import type { StyleRegistry } from './styles-writer'

const MAIN_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
const REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

/** Applied to Date cells that carry no explicit number-format, so exceljs reads them back as dates. */
const DEFAULT_DATE_FORMAT = 'mm-dd-yy'

/** Excel error literal used for values that cannot be represented numerically (NaN/±Infinity). */
const NUM_ERROR = '#NUM!'

/** A hyperlink discovered while serializing a sheet: its cell ref and external target. */
interface PendingHyperlink {
  ref: string
  target: string
}

/** A worksheet's serialized XML plus the external hyperlink relationships it references. */
export interface WorksheetWriteResult {
  readonly xml: string
  readonly hyperlinks: ReadonlyArray<{ rid: string; target: string }>
}

/** The style under which a value is rendered — Date cells force a date number-format. */
function effectiveStyle(value: Cell['value'], style: CellStyle): CellStyle {
  if (value instanceof Date && style.numberFormat === undefined) {
    return { ...style, numberFormat: DEFAULT_DATE_FORMAT }
  }
  return style
}

function writeCellValue(
  w: XmlWriter,
  cell: Cell,
  sst: SharedStrings,
  registry: StyleRegistry | undefined,
  hyperlinks: PendingHyperlink[],
): void {
  const v = cell.value
  if (v === null) return
  const xfIndex = registry !== undefined ? registry.xfIndexFor(effectiveStyle(v, cell.style)) : 0
  const s = xfIndex > 0 ? xfIndex : undefined
  if (typeof v === 'string') {
    const index = sst.add(v)
    w.open('c', { r: cell.address, t: 's', s }).open('v').text(String(index)).close('v').close('c')
    return
  }
  if (typeof v === 'number') {
    if (Number.isFinite(v)) {
      w.open('c', { r: cell.address, s }).open('v').text(String(v)).close('v').close('c')
    } else {
      // NaN/±Infinity are not valid OOXML numeric literals — emit a #NUM! error cell.
      w.open('c', { r: cell.address, t: 'e', s }).open('v').text(NUM_ERROR).close('v').close('c')
    }
    return
  }
  if (typeof v === 'boolean') {
    w.open('c', { r: cell.address, t: 'b', s })
      .open('v')
      .text(v ? '1' : '0')
      .close('v')
      .close('c')
    return
  }
  if (v instanceof Date) {
    w.open('c', { r: cell.address, s })
      .open('v')
      .text(String(dateToSerial(v)))
      .close('v')
      .close('c')
    return
  }
  if ('formula' in v) {
    writeFormulaCell(w, cell.address, s, v)
    return
  }
  if ('richText' in v) {
    const index = sst.addRich(v.richText)
    w.open('c', { r: cell.address, t: 's', s }).open('v').text(String(index)).close('v').close('c')
    return
  }
  // Hyperlink: display text is a shared string; the target becomes an external
  // relationship referenced from the <hyperlinks> block (emitted after sheetData).
  const index = sst.add(v.text)
  hyperlinks.push({ ref: cell.address, target: v.hyperlink })
  w.open('c', { r: cell.address, t: 's', s }).open('v').text(String(index)).close('v').close('c')
}

/** Serialize a formula cell `<c><f>EXPR</f><v>RESULT</v></c>`, tagging the result type. */
function writeFormulaCell(
  w: XmlWriter,
  address: string,
  s: number | undefined,
  v: FormulaValue,
): void {
  const result = v.result
  let t: string | undefined
  let rendered: string | undefined
  if (typeof result === 'string') {
    t = 'str'
    rendered = result
  } else if (typeof result === 'boolean') {
    t = 'b'
    rendered = result ? '1' : '0'
  } else if (typeof result === 'number') {
    if (Number.isFinite(result)) {
      rendered = String(result)
    } else {
      t = 'e'
      rendered = NUM_ERROR
    }
  } else if (result instanceof Date) {
    rendered = String(dateToSerial(result))
  }
  w.open('c', { r: address, t, s }).open('f').text(v.formula).close('f')
  if (rendered !== undefined) w.open('v').text(rendered).close('v')
  w.close('c')
}

/** Serialize one worksheet to `xl/worksheets/sheetN.xml`, returning its hyperlink relationships. */
export function writeWorksheetXml(
  ws: Worksheet,
  sst: SharedStrings,
  registry?: StyleRegistry,
): WorksheetWriteResult {
  const pending: PendingHyperlink[] = []
  const w = new XmlWriter().declaration().open('worksheet', { xmlns: MAIN_NS, 'xmlns:r': REL_NS })
  const dims = ws.dimensions
  if (dims !== undefined) w.leaf('dimension', { ref: encodeRange(dims) })

  // cols precedes sheetData per the CT_Worksheet schema sequence.
  const cols = ws.columns.filter((c) => c.width !== undefined)
  if (cols.length > 0) {
    w.open('cols')
    for (const col of cols) {
      w.leaf('col', { min: col.number, max: col.number, width: col.width, customWidth: 1 })
    }
    w.close('cols')
  }

  w.open('sheetData')
  for (const row of ws.rows) {
    const ht = row.height
    w.open('row', {
      r: row.number,
      ht,
      customHeight: ht !== undefined ? 1 : undefined,
    })
    for (const cell of row.cells) writeCellValue(w, cell, sst, registry, pending)
    w.close('row')
  }
  w.close('sheetData')

  // mergeCells follows sheetData per the CT_Worksheet schema sequence.
  const merges = ws.merges
  if (merges.length > 0) {
    w.open('mergeCells', { count: merges.length })
    for (const ref of merges) w.leaf('mergeCell', { ref })
    w.close('mergeCells')
  }

  // dataValidations follow mergeCells and precede hyperlinks (CT_Worksheet sequence).
  const validations = ws.dataValidations
  if (validations.length > 0) {
    w.open('dataValidations', { count: validations.length })
    for (const dv of validations) {
      w.open('dataValidation', {
        type: dv.type,
        operator: dv.operator,
        sqref: dv.sqref,
        allowBlank: dv.allowBlank === true ? 1 : undefined,
        showDropDown: dv.showDropDown === false ? 1 : undefined, // OOXML inverts this flag
        showErrorMessage: dv.showErrorMessage === true ? 1 : undefined,
      })
      w.open('formula1').text(dv.formula1).close('formula1')
      if (dv.formula2 !== undefined) w.open('formula2').text(dv.formula2).close('formula2')
      w.close('dataValidation')
    }
    w.close('dataValidations')
  }

  // hyperlinks follow dataValidations; each references an external relationship by r:id.
  const rels: Array<{ rid: string; target: string }> = []
  if (pending.length > 0) {
    w.open('hyperlinks')
    pending.forEach((h, i) => {
      const rid = `rId${i + 1}`
      w.leaf('hyperlink', { ref: h.ref, 'r:id': rid })
      rels.push({ rid, target: h.target })
    })
    w.close('hyperlinks')
  }

  w.close('worksheet')
  return { xml: w.toString(), hyperlinks: rels }
}
