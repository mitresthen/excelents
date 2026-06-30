import { expect, test } from 'vitest'
import { listFixtures, parseFixture } from '../../test/conformance/harness'
import { readCsv } from './reader'

test('parses a grid with quoted fields, doubled quotes, and embedded newlines', () => {
  const wb = readCsv('name,42,true\n"has,comma","has""quote","line\nbreak"')
  const ws = wb.sheets[0]!
  expect(ws.cell('A1').value).toBe('name')
  expect(ws.cell('B1').value).toBe(42)
  expect(ws.cell('C1').value).toBe(true)
  expect(ws.cell('A2').value).toBe('has,comma')
  expect(ws.cell('B2').value).toBe('has"quote')
  expect(ws.cell('C2').value).toBe('line\nbreak')
})

test('keeps leading-zero and date-like strings as strings (no over-eager inference)', async () => {
  // test-issue-991.csv: none of these may become a Date or number.
  const path = listFixtures('.csv').find((p) => p.endsWith('test-issue-991.csv'))!
  const text = new TextDecoder().decode(await parseFixture(path))
  const ws = readCsv(text).sheets[0]!
  expect(ws.cell('A1').value).toBe('2019-11-04')
  expect(ws.cell('A2').value).toBe('11-04-2019')
  expect(ws.cell('A3').value).toBe('2019-11-04T10:17:55')
  expect(ws.cell('A4').value).toBe('00210PRG1')
  expect(ws.cell('A5').value).toBe('1234-5thisisnotadate')
})

test('strips a leading BOM and handles CRLF rows', () => {
  const text = `${String.fromCharCode(0xfeff)}a,b\r\nc,d`
  const ws = readCsv(text).sheets[0]!
  expect(ws.cell('A1').value).toBe('a')
  expect(ws.cell('B1').value).toBe('b')
  expect(ws.cell('A2').value).toBe('c')
  expect(ws.cell('B2').value).toBe('d')
})

test('does not emit a spurious empty row for a trailing newline', () => {
  const ws = readCsv('a\nb\n').sheets[0]!
  expect(ws.dimensions).toEqual({ top: 1, left: 1, bottom: 2, right: 1 })
})
