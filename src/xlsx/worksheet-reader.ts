import type { CellValue } from '../model/cell'
import type { CellStyle } from '../model/style'
import type { Worksheet } from '../model/worksheet'
import { serialToDate } from '../utils/date'
import { isDateFormat } from '../utils/number-format'
import { tokenize } from '../xml/tokenizer'
import type { SharedStringValue } from './shared-strings-reader'

export interface ReadContext {
  readonly sharedStrings: SharedStringValue[]
  readonly cellStyles: CellStyle[]
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
  let sIndex: number | undefined
  let isV = false
  let inlineText: string | undefined
  let inIs = false
  let inT = false

  const finalize = (): void => {
    if (ref === undefined) return
    const cell = ws.cell(ref)
    const style = sIndex !== undefined ? ctx.cellStyles[sIndex] : undefined
    if (style !== undefined) cell.style = style
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
        inlineText = undefined
        isV = false
        inIs = false
        if (tok.selfClosing) ref = undefined
      } else if (tok.name === 'v') isV = true
      else if (tok.name === 'is') inIs = true
      else if (tok.name === 't') inT = true
    } else if (tok.type === 'close') {
      if (tok.name === 'v') isV = false
      else if (tok.name === 'is') inIs = false
      else if (tok.name === 't') inT = false
      else if (tok.name === 'c') {
        finalize()
        ref = undefined
      }
    } else if (tok.type === 'text') {
      if (isV) v = (v ?? '') + tok.value
      else if (inIs && inT) inlineText = (inlineText ?? '') + tok.value
    }
  }
  finalize()
}
