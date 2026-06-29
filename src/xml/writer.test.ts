import { expect, test } from 'vitest'
import { XmlWriter } from './writer'

test('writes a declaration', () => {
  expect(new XmlWriter().declaration().toString()).toBe(
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
  )
})

test('writes nested elements with attributes', () => {
  const xml = new XmlWriter()
    .open('worksheet', { 'xmlns:r': 'ns' })
    .leaf('dimension', { ref: 'A1:B2' })
    .close('worksheet')
    .toString()
  expect(xml).toBe('<worksheet xmlns:r="ns"><dimension ref="A1:B2"/></worksheet>')
})

test('skips undefined attributes and renders numbers', () => {
  expect(new XmlWriter().leaf('c', { r: 'A1', s: 3, t: undefined }).toString()).toBe(
    '<c r="A1" s="3"/>',
  )
})

test('escapes text content and attribute values', () => {
  const xml = new XmlWriter().open('t').text('a & b < c').close('t').toString()
  expect(xml).toBe('<t>a &amp; b &lt; c</t>')
  expect(new XmlWriter().leaf('a', { v: '"q" & <x>' }).toString()).toBe(
    '<a v="&quot;q&quot; &amp; &lt;x&gt;"/>',
  )
})

test('emits attributes in insertion order', () => {
  expect(new XmlWriter().leaf('c', { r: 'A1', s: 2, t: 's' }).toString()).toBe(
    '<c r="A1" s="2" t="s"/>',
  )
})
