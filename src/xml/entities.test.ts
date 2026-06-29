import { expect, test } from 'vitest'
import { escapeAttr, escapeText, unescapeXml } from './entities'

test('escapeText escapes &, <, > but leaves quotes', () => {
  expect(escapeText('a & b < c > d "e" \'f\'')).toBe('a &amp; b &lt; c &gt; d "e" \'f\'')
})

test('escapeAttr escapes &, <, >, and double-quote', () => {
  expect(escapeAttr('x "y" & <z>')).toBe('x &quot;y&quot; &amp; &lt;z&gt;')
})

test('escapeText escapes & before other entities (no double-escaping)', () => {
  expect(escapeText('&lt;')).toBe('&amp;lt;')
})

test('unescapeXml decodes the five predefined entities', () => {
  expect(unescapeXml('a &amp; b &lt; c &gt; d &quot;e&quot; &apos;f&apos;')).toBe(
    'a & b < c > d "e" \'f\'',
  )
})

test('unescapeXml decodes decimal and hex numeric char refs', () => {
  expect(unescapeXml('&#65;&#66;&#x43;')).toBe('ABC')
  expect(unescapeXml('tab&#9;end')).toBe('tab\tend')
})

test('escape then unescape round-trips arbitrary text', () => {
  const s = 'Tom & Jerry: <5 "quoted" & \'apos\'>'
  expect(unescapeXml(escapeAttr(s))).toBe(s)
  expect(unescapeXml(escapeText(s))).toBe(s)
})
