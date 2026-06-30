import type { RichTextRun } from '../model/cell'
import type { SharedStrings } from '../utils/shared-strings'
import { XmlWriter } from '../xml/writer'

const MAIN_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'

/** Emit a `<t>`, preserving leading/trailing whitespace via `xml:space="preserve"` when significant. */
function writeTextEl(w: XmlWriter, text: string): void {
  const preserve = text.trim() !== text
  w.open('t', preserve ? { 'xml:space': 'preserve' } : undefined)
    .text(text)
    .close('t')
}

/** Emit one rich-text run `<r>[<rPr>…</rPr>]<t>…</t></r>`. */
function writeRun(w: XmlWriter, run: RichTextRun): void {
  w.open('r')
  const font = run.font
  if (font !== undefined) {
    w.open('rPr')
    if (font.bold === true) w.leaf('b')
    if (font.italic === true) w.leaf('i')
    if (font.underline === true) w.leaf('u')
    if (font.size !== undefined) w.leaf('sz', { val: font.size })
    if (font.color !== undefined) w.leaf('color', { rgb: font.color })
    // Inside <rPr> the font-name element is <rFont> (CT_RPrElt), not <name>.
    if (font.name !== undefined) w.leaf('rFont', { val: font.name })
    w.close('rPr')
  }
  writeTextEl(w, run.text)
  w.close('r')
}

/** Serialize the shared-string table to `xl/sharedStrings.xml`. */
export function writeSharedStringsXml(sst: SharedStrings): string {
  const w = new XmlWriter()
    .declaration()
    .open('sst', { xmlns: MAIN_NS, count: sst.count, uniqueCount: sst.uniqueCount })
  for (const entry of sst.entries) {
    w.open('si')
    if (entry.type === 'plain') {
      writeTextEl(w, entry.text)
    } else {
      for (const run of entry.runs) writeRun(w, run)
    }
    w.close('si')
  }
  return w.close('sst').toString()
}
