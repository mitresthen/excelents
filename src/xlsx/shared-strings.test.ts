import { expect, test } from 'vitest'
import { createWorkbook } from '../model/workbook'
import { SharedStrings } from '../utils/shared-strings'
import { writeSharedStringsXml } from './shared-strings-writer'
import { writeXlsx } from './write'

test('rich run properties are emitted in CT_RPrElt schema order', () => {
  const sst = new SharedStrings()
  sst.addRich([
    {
      text: 'x',
      font: {
        name: 'Arial',
        bold: true,
        italic: true,
        color: 'FF0000FF',
        size: 12,
        underline: true,
      },
    },
  ])
  const xml = writeSharedStringsXml(sst)
  const rpr = xml.slice(xml.indexOf('<rPr>'), xml.indexOf('</rPr>'))
  // CT_RPrElt (ECMA-376 §18.4.7): rFont, charset, family, b, i, ..., color, sz, u, ...
  const positions = ['rFont', 'b', 'i', 'color', 'sz', 'u'].map((tag) => rpr.indexOf(`<${tag}`))
  expect(positions.every((p) => p >= 0)).toBe(true)
  for (let i = 1; i < positions.length; i++) {
    expect(positions[i]).toBeGreaterThan(positions[i - 1]!)
  }
})

test('string cells round-trip via shared strings (exceljs reads them)', async () => {
  const wb = createWorkbook()
  const ws = wb.addSheet('S')
  ws.cell('A1').value = 'apple'
  ws.cell('A2').value = 'banana'
  ws.cell('A3').value = 'apple' // duplicate → same shared index
  const bytes = await writeXlsx(wb)

  const ExcelJS = (await import('exceljs')).default
  const oracle = new ExcelJS.Workbook()
  // @ts-expect-error -- @types/node 22 Buffer<ArrayBuffer> vs non-generic Buffer in exceljs decl
  await oracle.xlsx.load(Buffer.from(bytes))
  const sheet = oracle.getWorksheet('S')
  expect(sheet?.getCell('A1').value).toBe('apple')
  expect(sheet?.getCell('A2').value).toBe('banana')
  expect(sheet?.getCell('A3').value).toBe('apple')

  // the part exists and dedups (2 unique strings)
  const { OpcPackage } = await import('../opc/package')
  const pkg = await OpcPackage.read(bytes)
  const sst = new TextDecoder().decode(pkg.getPart('xl/sharedStrings.xml'))
  expect(sst).toContain('uniqueCount="2"')
})
