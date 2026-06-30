import type { RichTextRun } from '../model/cell'
import type { Font } from '../model/style'
import { tokenize } from '../xml/tokenizer'

/** One entry of a parsed shared-string table: plain text or a sequence of formatted runs. */
export type SharedStringValue =
  | { readonly kind: 'plain'; readonly text: string }
  | { readonly kind: 'rich'; readonly runs: RichTextRun[] }

/** Parse `xl/sharedStrings.xml` into index-ordered values (inverse of writeSharedStringsXml). */
export function readSharedStrings(xml: string): SharedStringValue[] {
  const out: SharedStringValue[] = []
  let runs: RichTextRun[] = []
  let plain = ''
  let sawRun = false
  let inText = false
  let text = ''
  let font: Font | undefined
  let inFont = false

  for (const tok of tokenize(xml)) {
    if (tok.type === 'open') {
      if (tok.name === 'si') {
        runs = []
        plain = ''
        sawRun = false
      } else if (tok.name === 'r') {
        sawRun = true
        font = undefined
        text = ''
      } else if (tok.name === 'rPr') {
        inFont = true
        font = {}
      } else if (tok.name === 't') {
        inText = true
      } else if (inFont && font !== undefined) {
        if (tok.name === 'b') font.bold = true
        else if (tok.name === 'i') font.italic = true
        else if (tok.name === 'u') font.underline = true
        else if (tok.name === 'sz') font.size = Number(tok.attributes['val'])
        else if (tok.name === 'color') font.color = tok.attributes['rgb']
        else if (tok.name === 'rFont') font.name = tok.attributes['val']
      }
    } else if (tok.type === 'close') {
      if (tok.name === 't') inText = false
      else if (tok.name === 'rPr') inFont = false
      else if (tok.name === 'r') {
        runs.push(
          font !== undefined && Object.keys(font).length > 0 ? { text, font } : { text },
        )
      } else if (tok.name === 'si') {
        out.push(sawRun ? { kind: 'rich', runs } : { kind: 'plain', text: plain })
      }
    } else if (tok.type === 'text' && inText) {
      if (sawRun) text += tok.value
      else plain += tok.value
    }
  }
  return out
}
