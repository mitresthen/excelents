import { expect, test } from 'vitest'
import { Worksheet } from '../model/worksheet'
import { SharedStrings } from '../utils/shared-strings'
import { StyleRegistry } from './styles-writer'
import { writeWorksheetXml } from './worksheet-writer'

function render(ws: Worksheet): string {
  return writeWorksheetXml(ws, new SharedStrings(), new StyleRegistry()).xml
}

test('autoFilter is emitted after sheetData and before mergeCells', () => {
  const ws = new Worksheet('s')
  ws.cell('A1').value = 'h'
  ws.merge('A1:B1')
  ws.setAutoFilter('A1:B10')
  const xml = render(ws)
  expect(xml).toContain('<autoFilter ref="A1:B10"/>')
  const af = xml.indexOf('<autoFilter')
  expect(xml.indexOf('</sheetData>')).toBeLessThan(af)
  expect(af).toBeLessThan(xml.indexOf('<mergeCells'))
})

test('freeze rows emits sheetViews/pane between dimension and cols', () => {
  const ws = new Worksheet('s')
  ws.cell('A1').value = 'h'
  ws.column(1).width = 10
  ws.freeze({ rows: 9 })
  const xml = render(ws)
  expect(xml).toContain(
    '<pane ySplit="9" topLeftCell="A10" activePane="bottomLeft" state="frozen"/>',
  )
  expect(xml).toContain('<selection pane="bottomLeft" activeCell="A10" sqref="A10"/>')
  const sv = xml.indexOf('<sheetViews>')
  expect(xml.indexOf('<dimension')).toBeLessThan(sv)
  expect(sv).toBeLessThan(xml.indexOf('<cols>'))
})

test('freeze on rows and cols sets bottomRight active pane', () => {
  const ws = new Worksheet('s')
  ws.cell('A1').value = 'h'
  ws.freeze({ rows: 1, cols: 2 })
  const xml = render(ws)
  expect(xml).toContain(
    '<pane xSplit="2" ySplit="1" topLeftCell="C2" activePane="bottomRight" state="frozen"/>',
  )
})
