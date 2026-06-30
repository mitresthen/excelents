import type { CellValue, FormulaValue } from '../model/cell'
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

function sharedToValue(v: SharedStringValue | undefined): CellValue {
  if (v === undefined) return ''
  if (v.kind === 'plain') return v.text
  return { richText: v.runs }
}

/** Populate a Worksheet from one `sheetN.xml` (inverse of writeWorksheetXml). */
export function readWorksheetInto(ws: Worksheet, xml: string, ctx: ReadContext): void {
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
        if (tok.selfClosing) ref = undefined
      } else if (tok.name === 'v') isV = !tok.selfClosing
      // A self-closing element emits no close event, so only "enter" it when it can hold text.
      // Self-closing <f/> (shared-formula member) must NOT leave the accumulator open.
      else if (tok.name === 'f') inF = !tok.selfClosing
      else if (tok.name === 'is') inIs = !tok.selfClosing
      else if (tok.name === 't') inT = !tok.selfClosing
      else if (tok.name === 'mergeCell') {
        const mref = tok.attributes['ref']
        if (mref !== undefined) ws.merge(mref)
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
      }
    } else if (tok.type === 'close') {
      if (tok.name === 'v') isV = false
      else if (tok.name === 'f') inF = false
      else if (tok.name === 'is') inIs = false
      else if (tok.name === 't') inT = false
      else if (tok.name === 'c') {
        finalize()
        ref = undefined
      }
    } else if (tok.type === 'text') {
      if (isV) v = (v ?? '') + tok.value
      else if (inF) f = (f ?? '') + tok.value
      else if (inIs && inT) inlineText = (inlineText ?? '') + tok.value
    }
  }
  finalize()

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
}
