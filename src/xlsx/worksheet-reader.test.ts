import { expect, test } from 'vitest'
import { createWorkbook } from '../model/workbook'
import { readWorksheetInto, type ReadContext } from './worksheet-reader'

const EMPTY_CTX: ReadContext = {
  sharedStrings: [],
  cellStyles: [],
  hyperlinkTargets: new Map(),
}

test('a shared-formula member cell (self-closing <f/>) in whitespace-indented XML reads as its number', () => {
  // Pretty-printed OOXML: a newline+indent sits between the self-closing <f/> and <v>.
  // The self-closing <f/> must not leave the formula accumulator "open" and swallow it.
  const xml =
    '<worksheet><sheetData><row r="5">' +
    '<c r="A5" t="n"><f t="shared" si="0"/>\n      <v>5</v></c>' +
    '</row></sheetData></worksheet>'
  const ws = createWorkbook().addSheet('S')
  readWorksheetInto(ws, xml, EMPTY_CTX)
  expect(ws.cell('A5').value).toBe(5)
})

test('a shared-formula master cell (<f ...>EXPR</f>) keeps its formula', () => {
  const xml =
    '<worksheet><sheetData><row r="4">' +
    '<c r="A4" t="n"><f t="shared" ref="A4:A6" si="0">A3+A2</f><v>3</v></c>' +
    '</row></sheetData></worksheet>'
  const ws = createWorkbook().addSheet('S')
  readWorksheetInto(ws, xml, EMPTY_CTX)
  expect(ws.cell('A4').value).toEqual({ formula: 'A3+A2', result: 3 })
})
