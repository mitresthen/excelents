import type { CellValue, FormulaValue } from '../model/cell'
import type { DataValidation } from '../model/data-validation'
import type { CellStyle } from '../model/style'
import type { Worksheet } from '../model/worksheet'
import { serialToDate } from '../utils/date'
import { isDateFormat } from '../utils/number-format'
import { tokenize } from '../xml/tokenizer'
import type { SharedStringValue } from './shared-strings-reader'

export interface ReadContext {
  readonly sharedStrings: SharedStringValue[]
  readonly cellStyles: CellStyle[]
  /** rId -> external hyperlink URL, from the worksheet's own relationships part. */
  readonly hyperlinkTargets: Map<string, string>
}

/** Narrow a raw dataValidation type attribute to the model union (no cast). */
function asDvType(s: string | undefined): DataValidation['type'] | undefined {
  switch (s) {
    case 'list':
    case 'whole':
    case 'decimal':
    case 'textLength':
    case 'date':
      return s
    default:
      return undefined
  }
}

/** Narrow a raw dataValidation operator attribute to the model union (no cast). */
function asDvOperator(s: string | undefined): NonNullable<DataValidation['operator']> | undefined {
  switch (s) {
    case 'between':
    case 'notBetween':
    case 'equal':
    case 'notEqual':
    case 'greaterThan':
    case 'lessThan':
    case 'greaterThanOrEqual':
    case 'lessThanOrEqual':
      return s
    default:
      return undefined
  }
}

/** Assemble a DataValidation from parsed attributes + formula text, or undefined if invalid. */
function buildDataValidation(
  attrs: Record<string, string>,
  formula1: string | undefined,
  formula2: string | undefined,
): DataValidation | undefined {
  const type = asDvType(attrs['type'])
  const sqref = attrs['sqref']
  if (type === undefined || sqref === undefined || formula1 === undefined) return undefined
  const operator = asDvOperator(attrs['operator'])
  return {
    sqref,
    type,
    formula1,
    ...(operator !== undefined ? { operator } : {}),
    ...(formula2 !== undefined ? { formula2 } : {}),
    ...(attrs['allowBlank'] === '1' ? { allowBlank: true } : {}),
    // OOXML showDropDown="1" HIDES the dropdown -> model showDropDown:false.
    ...(attrs['showDropDown'] === '1' ? { showDropDown: false } : {}),
    ...(attrs['showErrorMessage'] === '1' ? { showErrorMessage: true } : {}),
  }
}

function sharedToValue(v: SharedStringValue | undefined): CellValue {
  if (v === undefined) return ''
  if (v.kind === 'plain') return v.text
  return { richText: v.runs }
}

/**
 * Populate a Worksheet from one `sheetN.xml` (inverse of writeWorksheetXml).
 * Returns the relationship ids of any `<tablePart>` elements, for the caller to resolve.
 */
export function readWorksheetInto(ws: Worksheet, xml: string, ctx: ReadContext): string[] {
  let ref: string | undefined
  let type = 'n'
  let v: string | undefined
  let f: string | undefined
  let sIndex: number | undefined
  let isV = false
  let inF = false
  let inlineText: string | undefined
  let inIs = false
  let inT = false
  const pendingHyperlinks: Array<{ ref: string; rid: string }> = []

  const parsedValidations: DataValidation[] = []
  let dvAttrs: Record<string, string> | undefined
  let dvF1: string | undefined
  let dvF2: string | undefined
  let inDvF1 = false
  let inDvF2 = false
  const tableRids: string[] = []

  const finalize = (): void => {
    if (ref === undefined) return
    const cell = ws.cell(ref)
    const style = sIndex !== undefined ? ctx.cellStyles[sIndex] : undefined
    if (style !== undefined) cell.style = style

    if (f !== undefined) {
      let result: FormulaValue['result']
      if (type === 'str') result = v
      else if (type === 'b') result = v === '1'
      else if (type !== 'e' && v !== undefined) result = Number(v)
      cell.value = result === undefined ? { formula: f } : { formula: f, result }
      return
    }
    if (type === 'e') return // error cells have no model representation; leave empty
    if (type === 's' && v !== undefined) {
      cell.value = sharedToValue(ctx.sharedStrings[Number(v)])
    } else if (type === 'inlineStr' && inlineText !== undefined) {
      cell.value = inlineText
    } else if (type === 'str' && v !== undefined) {
      cell.value = v
    } else if (type === 'b' && v !== undefined) {
      cell.value = v === '1'
    } else if (v !== undefined) {
      // A numeric cell whose format is a date renders back as a Date.
      cell.value =
        style?.numberFormat !== undefined && isDateFormat(style.numberFormat)
          ? serialToDate(Number(v))
          : Number(v)
    }
  }

  for (const tok of tokenize(xml)) {
    if (tok.type === 'open') {
      if (tok.name === 'c') {
        finalize()
        ref = tok.attributes['r']
        type = tok.attributes['t'] ?? 'n'
        const s = tok.attributes['s']
        sIndex = s !== undefined ? Number(s) : undefined
        v = undefined
        f = undefined
        inlineText = undefined
        isV = false
        inF = false
        inIs = false
        if (tok.selfClosing) {
          // A self-closing cell has no value, but when it carries a style
          // (`<c r="A2" s="3"/>` — fills across merges, spacer rows) the
          // style must survive the round-trip. Unstyled empties stay dropped.
          if (sIndex !== undefined) finalize()
          ref = undefined
        }
      } else if (tok.name === 'v') isV = !tok.selfClosing
      // A self-closing element emits no close event, so only "enter" it when it can hold text.
      // Self-closing <f/> (shared-formula member) must NOT leave the accumulator open.
      else if (tok.name === 'f') inF = !tok.selfClosing
      else if (tok.name === 'is') inIs = !tok.selfClosing
      else if (tok.name === 't') inT = !tok.selfClosing
      else if (tok.name === 'mergeCell') {
        const mref = tok.attributes['ref']
        // recordMerge, not merge(): parsed covered cells keep their own styles.
        if (mref !== undefined) ws.recordMerge(mref)
      } else if (tok.name === 'col') {
        const width = tok.attributes['width']
        const min = Number(tok.attributes['min'])
        const max = Number(tok.attributes['max'])
        if (width !== undefined && !Number.isNaN(min) && !Number.isNaN(max)) {
          for (let n = min; n <= max; n++) ws.column(n).width = Number(width)
        }
      } else if (tok.name === 'row') {
        const ht = tok.attributes['ht']
        const r = tok.attributes['r']
        if (ht !== undefined && r !== undefined) ws.getRow(Number(r)).height = Number(ht)
      } else if (tok.name === 'hyperlink') {
        const hr = tok.attributes['ref']
        const rid = tok.attributes['r:id']
        if (hr !== undefined && rid !== undefined) pendingHyperlinks.push({ ref: hr, rid })
      } else if (tok.name === 'dataValidation') {
        dvAttrs = tok.attributes
        dvF1 = undefined
        dvF2 = undefined
        if (tok.selfClosing) {
          const dv = buildDataValidation(dvAttrs, dvF1, dvF2)
          if (dv !== undefined) parsedValidations.push(dv)
          dvAttrs = undefined
        }
      } else if (tok.name === 'formula1') inDvF1 = !tok.selfClosing
      else if (tok.name === 'formula2') inDvF2 = !tok.selfClosing
      else if (tok.name === 'tablePart') {
        const rid = tok.attributes['r:id']
        if (rid !== undefined) tableRids.push(rid)
      }
    } else if (tok.type === 'close') {
      if (tok.name === 'v') isV = false
      else if (tok.name === 'f') inF = false
      else if (tok.name === 'is') inIs = false
      else if (tok.name === 't') inT = false
      else if (tok.name === 'formula1') inDvF1 = false
      else if (tok.name === 'formula2') inDvF2 = false
      else if (tok.name === 'dataValidation') {
        if (dvAttrs !== undefined) {
          const dv = buildDataValidation(dvAttrs, dvF1, dvF2)
          if (dv !== undefined) parsedValidations.push(dv)
        }
        dvAttrs = undefined
      } else if (tok.name === 'c') {
        finalize()
        ref = undefined
      }
    } else if (tok.type === 'text') {
      if (isV) v = (v ?? '') + tok.value
      else if (inF) f = (f ?? '') + tok.value
      else if (inIs && inT) inlineText = (inlineText ?? '') + tok.value
      else if (inDvF1) dvF1 = (dvF1 ?? '') + tok.value
      else if (inDvF2) dvF2 = (dvF2 ?? '') + tok.value
    }
  }
  finalize()
  for (const dv of parsedValidations) ws.addDataValidation(dv)

  // Resolve hyperlinks (declared after sheetData) over the already-populated cells.
  // Only a plain-string cell becomes a { text, hyperlink } value. A cell already holding a
  // richer value (formula/number/Date/richText) keeps it — the model cannot combine a
  // hyperlink with those, and clobbering would erase the cell's real content.
  for (const { ref: hRef, rid } of pendingHyperlinks) {
    const url = ctx.hyperlinkTargets.get(rid)
    if (url === undefined) continue
    const cell = ws.cell(hRef)
    if (typeof cell.value === 'string') cell.value = { text: cell.value, hyperlink: url }
  }

  return tableRids
}
