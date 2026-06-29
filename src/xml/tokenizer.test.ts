import { expect, test } from 'vitest'
import { tokenize, type XmlToken } from './tokenizer'
import { XmlWriter } from './writer'

const toks = (xml: string): XmlToken[] => [...tokenize(xml)]

test('tokenizes nested elements and attributes', () => {
  expect(toks('<a x="1"><b/></a>')).toEqual([
    { type: 'open', name: 'a', attributes: { x: '1' }, selfClosing: false },
    { type: 'open', name: 'b', attributes: {}, selfClosing: true },
    { type: 'close', name: 'a' },
  ])
})

test('handles both quote styles and namespaced names', () => {
  expect(toks(`<c:e r:id='R1' n="2"/>`)).toEqual([
    { type: 'open', name: 'c:e', attributes: { 'r:id': 'R1', n: '2' }, selfClosing: true },
  ])
})

test('emits text and unescapes entities in text and attributes', () => {
  expect(toks('<t a="x &amp; y">a &lt; b</t>')).toEqual([
    { type: 'open', name: 't', attributes: { a: 'x & y' }, selfClosing: false },
    { type: 'text', value: 'a < b' },
    { type: 'close', name: 't' },
  ])
})

test('emits CDATA as raw text', () => {
  expect(toks('<t><![CDATA[a < b & c]]></t>')).toEqual([
    { type: 'open', name: 't', attributes: {}, selfClosing: false },
    { type: 'text', value: 'a < b & c' },
    { type: 'close', name: 't' },
  ])
})

test('skips the xml declaration and comments', () => {
  expect(toks('<?xml version="1.0"?><!-- hi --><a/>')).toEqual([
    { type: 'open', name: 'a', attributes: {}, selfClosing: true },
  ])
})

test('preserves whitespace-only text runs between elements', () => {
  expect(toks('<a>\n  <b/>\n</a>')).toEqual([
    { type: 'open', name: 'a', attributes: {}, selfClosing: false },
    { type: 'text', value: '\n  ' },
    { type: 'open', name: 'b', attributes: {}, selfClosing: true },
    { type: 'text', value: '\n' },
    { type: 'close', name: 'a' },
  ])
})

test('round-trips through the writer (open/leaf/text/close) to equivalent XML', () => {
  const input = '<root xmlns:r="ns"><a r:id="R1">hello &amp; bye</a><b/></root>'
  const w = new XmlWriter()
  for (const t of tokenize(input)) {
    if (t.type === 'open') {
      if (t.selfClosing) w.leaf(t.name, t.attributes)
      else w.open(t.name, t.attributes)
    } else if (t.type === 'close') w.close(t.name)
    else w.text(t.value)
  }
  expect(w.toString()).toBe(input)
})

test('parses self-closing tag with attribute adjacent to />', () => {
  expect(toks('<c s="0"/>')).toEqual([
    { type: 'open', name: 'c', attributes: { s: '0' }, selfClosing: true },
  ])
})
