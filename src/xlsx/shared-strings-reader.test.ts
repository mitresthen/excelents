import { expect, test } from 'vitest'
import { SharedStrings } from '../utils/shared-strings'
import { readSharedStrings } from './shared-strings-reader'
import { writeSharedStringsXml } from './shared-strings-writer'

test('reads plain and rich shared strings round-tripped from the writer', () => {
  const sst = new SharedStrings()
  sst.add('apple')
  sst.addRich([{ text: 'Hello ' }, { text: 'World', font: { bold: true } }])
  const values = readSharedStrings(writeSharedStringsXml(sst))

  expect(values[0]).toEqual({ kind: 'plain', text: 'apple' })
  expect(values[1]?.kind).toBe('rich')
  if (values[1]?.kind === 'rich') {
    expect(values[1].runs.map((r) => r.text).join('')).toBe('Hello World')
    expect(values[1].runs[1]?.font?.bold).toBe(true)
  }
})

test('preserves significant whitespace in plain strings', () => {
  const sst = new SharedStrings()
  sst.add('  leading')
  const values = readSharedStrings(writeSharedStringsXml(sst))
  expect(values[0]).toEqual({ kind: 'plain', text: '  leading' })
})
